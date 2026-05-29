# VBA-MicrosoftGraph-2.1 — Technical Evaluation

**Date:** 2026-04-29  
**Scope:** `Src/Graph.bas`, `Src/AttachmentHelpers.bas`, `Src/TimeZoneHelpers.bas`, `authenticators/GraphAuthenticator.cls`, `authenticators/OAuth2Authenticator.cls`  
**References:** [Microsoft Graph best practices](https://learn.microsoft.com/graph/best-practices-concept), [Graph throttling guidance](https://learn.microsoft.com/graph/throttling), [OAuth2 auth code flow](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-auth-code-flow)

---

## Priority Fix List

Apply these in order — the top items cause active crashes or data loss in production.

1. `Scopes(1)` → `VBA.Join(Me.Scopes, " ")` — crashes with a single scope
2. `Replace(fillString, " ", "")` → `Split`/`Trim` — corrupts every attachment path with spaces
3. Add `= "Error"` sentinel to `GetContactFolderID` / `GetMailFolderID` — infinite loop in production
4. Add `MAX_RETRIES` cap to all 8 public Graph functions — infinite loop on persistent 401
5. Move `client_secret` from URL querystring to POST body (`OAuth2Authenticator`)
6. Fix `"Top"` → `"$top"` — `$top` is always silently ignored; lists are always page-capped
7. Add `@odata.nextLink` pagination loop — data silently truncated beyond first page
8. Add HTTP 429 + `Retry-After` handling

---

## 1. Security Issues

### 1.1 CRITICAL — `client_secret` in URL Querystring (`OAuth2Authenticator.cls`)

```vba
' Current (INSECURE):
auth_Request.AddQuerystringParam "client_secret", Me.ClientSecret
```

Query strings are captured in proxy logs, server access logs, browser history, and HTTP `Referer` headers. The `client_secret` must be sent in the POST body.

```vba
' Fix — use body parameter instead:
auth_Request.AddBodyParameter "client_secret", Me.ClientSecret
' Remove the AddQuerystringParam line above.
```

### 1.2 CRITICAL — `ClientSecret` Stored as Plaintext in `AdminTable` (`Graph.bas`)

```vba
pClientSecret = DLookup("ClientSecret", "AdminTable")
```

Any user who can open the `.accdb` file can read the secret from `AdminTable`. At minimum, encrypt the value using Windows DPAPI (`CryptProtectData` / `CryptUnprotectData`). Better: store it in the Windows Credential Manager and retrieve via `CredRead` (WinAPI `Declare PtrSafe`).

### 1.3 HIGH — ROPC Flow (`grant_type=password`) in `OAuth2Authenticator.cls`

```vba
auth_Request.AddQuerystringParam "grant_type", "password"
```

Microsoft [explicitly deprecates ROPC](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth-ropc) for production. It bypasses MFA, does not support Conditional Access policies, and routes raw user credentials through the application. Use `authorization_code` (interactive) or `client_credentials` (service-to-service) flows instead.

### 1.4 HIGH — `Public Token As String` Exposes Bearer Token (`GraphAuthenticator.cls`)

```vba
Public Token As String
```

Any VBA code in the same project can read or overwrite the live bearer token.

```vba
' Fix — encapsulate:
Private pToken As String

Public Property Get Token() As String
    Token = pToken
End Property
' Only set pToken internally via Friend/Private methods.
```

---

## 2. Correctness Bugs

### 2.1 CRITICAL — `Scopes(1)` Off-by-One — Runtime Crash or Dropped Scope (`GraphAuthenticator.cls`)

`AddScope` builds a **0-based** array. With a single scope (e.g., `.default` for `client_credentials`), `Scopes(1)` raises **runtime error 9 — Subscript out of range**. With multiple scopes, `Scopes(0)` (the first scope) is silently dropped.

```vba
' Current:
auth_Body.Add "scope", Me.Scopes(1)

' Fix — join all scopes space-separated per OAuth2 spec:
auth_Body.Add "scope", VBA.Join(Me.Scopes, " ")
```

### 2.2 HIGH — Infinite Retry Loop on Persistent 401 (`Graph.bas` — all 8 public functions)

```vba
' Current — hangs Access forever if token keeps failing:
While sStatus = "Retry"
    Set CreateDraftMessage = Client.Execute(Request)
    If CreateDraftMessage.StatusCode = WebStatusCode.Unauthorized _
       And InStr(CreateDraftMessage.Content, "expired") Then
        ClearAuthCodes
        sStatus = "Retry"
    Else
        sStatus = "Done"
    End If
Wend
```

```vba
' Fix — cap retries:
Const MAX_RETRIES As Integer = 2
Dim nRetry As Integer
nRetry = 0
Do
    Set CreateDraftMessage = Client.Execute(Request)
    If CreateDraftMessage.StatusCode = WebStatusCode.Unauthorized _
       And InStr(CreateDraftMessage.Content, "expired") _
       And nRetry < MAX_RETRIES Then
        ClearAuthCodes
        nRetry = nRetry + 1
    Else
        Exit Do
    End If
Loop
```

Apply the same pattern to `GraphSendMail`, `GetGroupID`, `GetCalendarGroupID`, `GetCalendarID`, `GetContactFolderID`, `GetMailFolderID`, `ListMessages`, and `CreateEvent`.

### 2.3 HIGH — `GetContactFolderID` / `GetMailFolderID` Infinite Loop When Folder Not Found (`Graph.bas`)

When the named folder does not exist, the loop completes the `For Each` without setting the return value away from `"Retry"`, so the `While` condition remains true and loops forever. Compare `GetGroupID`, which correctly assigns `= "Error"` after exhausting the list.

```vba
' Add immediately after the Next FolderInfo line:
If GetContactFolderID = "Retry" Then GetContactFolderID = "Error"
' (same pattern for GetMailFolderID)
```

### 2.4 HIGH — `FillAttachmentCollection` Strips Spaces from File Paths (`Graph.bas`)

```vba
' Current — destroys paths like "C:\My Documents\report.pdf":
fillString = Replace(fillString, " ", "")
```

The intent was to normalise spacing around the `;` delimiter. Fix: split on `";"` and `Trim` each token.

```vba
' Fix:
Dim aPaths() As String
aPaths = Split(fillString, ";")
Dim i As Long
For i = LBound(aPaths) To UBound(aPaths)
    Dim sPath As String
    sPath = Trim(aPaths(i))
    If sPath = "" Then GoTo NextPath
    Set attachment = New Dictionary
    attachment.Add "@odata.type", "#microsoft.graph.fileAttachment"
    attachment.Add "name", Mid(sPath, InStrRev(sPath, "\") + 1)
    attachment.Add "contentBytes", ConvertFileToBase64(sPath)
    fillCollection.Add attachment
NextPath:
Next i
```

The same fix applies to `FillEmailAddressCollection` and `FillAttendeeCollection` which use the identical `Replace(fillString, " ", "")` pattern.

### 2.5 HIGH — `CreateGUID` Without `Randomize` — Deterministic GUIDs (`Graph.bas`)

Without `Randomize`, VBA's `Rnd` produces the **same sequence on every session start**. GUIDs collide across sessions.

```vba
' Fix — minimum:
Public Function CreateGUID() As String
    Randomize
    ...

' Better — truly random:
Public Function CreateGUID() As String
    CreateGUID = CreateObject("Scriptlet.TypeLib").GUID
    CreateGUID = Mid(CreateGUID, 2, Len(CreateGUID) - 2) ' strip braces
End Function
```

### 2.6 MEDIUM — `"Top"` Instead of `"$top"` — Parameter Silently Ignored (`Graph.bas`)

OData system query options require the `$` prefix. `"Top"` (without `$`) is not a recognised parameter and is silently ignored — the server applies its own default page size (100 items).

```vba
' Current:
Request.AddQuerystringParam "Top", 1000

' Fix:
Request.AddQuerystringParam "$top", 1000
```

### 2.7 MEDIUM — `allowNewTimeProposals` Sent as String `"true"` (`Graph.bas`)

The Graph `event` schema defines `allowNewTimeProposals` as a **Boolean**. Sending `"true"` (a JSON string) is schema-invalid.

```vba
' Current:
.AddBodyParameter "allowNewTimeProposals", "true"

' Fix:
.AddBodyParameter "allowNewTimeProposals", True
```

### 2.8 LOW — `ClearAuthCodes` / `Logout` Use `Dim X As New ClassName` Before `Set X = ...` (`Graph.bas`)

```vba
' Current — misleading; creates a new instance then immediately discards it:
Public Sub ClearAuthCodes()
    Dim Auth As New GraphAuthenticator
    Set Auth = Client.Authenticator
    Call Auth.ClearCodes
End Sub
```

Functionally correct but allocates an unused object. Prefer:

```vba
Public Sub ClearAuthCodes()
    Dim Auth As GraphAuthenticator
    Set Auth = Client.Authenticator
    Auth.ClearCodes
End Sub
```

### 2.9 LOW — Dead `Exit Do` After `Err.Raise` (`GraphAuthenticator.Login`)

```vba
If Now > dtEnd Then
    Err.Raise 11042 + vbObjectError, "OAuthDialog", "Login error: Wait time exceeded."
    Exit Do   ' ← never reached; Err.Raise with active On Error GoTo transfers immediately
End If
```

The `Exit Do` is unreachable dead code. Remove it to avoid misleading readers.

### 2.10 LOW — `sSessionID` Never Set — Logout Navigates to Invalid URL (`GraphAuthenticator.cls`)

```vba
Private sSessionID As String   ' always ""

' Logout navigates to:
"https://m365.cloud.microsoft/estslogout?ru=%2F&sessionId="
```

`sSessionID` is declared but never assigned anywhere in the class. The logout URL is always missing the session ID. The session ID should be captured from the login response URL or the ID token claims.

---

## 3. Microsoft Graph API Best-Practice Violations

### 3.1 No `@odata.nextLink` Pagination — Data Silently Truncated

Graph API pages results. Default page sizes are 10 (messages), 100 (contacts, groups, calendars). When results exceed one page, the response includes `@odata.nextLink`. None of the list functions follow nextLink — only the first page is ever returned.

```vba
' Fix pattern — add after the OK branch:
Dim sNextLink As String
sNextLink = ""
If Response.Data.Exists("@odata.nextLink") Then
    sNextLink = Response.Data("@odata.nextLink")
End If
' Continue looping while sNextLink <> ""
```

### 3.2 No HTTP 429 / `Retry-After` Handling — Throttling Silently Fails

[Graph throttling guidance](https://learn.microsoft.com/graph/throttling) requires honouring the `Retry-After` response header on HTTP 429. All retry loops only handle 401 — a 429 falls through to `sStatus = "Done"` and the call silently fails.

```vba
' Add 429 handling to every retry loop:
ElseIf Response.StatusCode = 429 Then
    ' Retry-After header not directly accessible via VBA-Web without extension;
    ' fall back to a fixed wait or parse from Response.Headers collection.
    Application.Wait Now + TimeValue("0:00:30")  ' 30-second back-off
    ' Loop again (do not set sStatus = "Done")
```

### 3.3 No Server-Side `$filter` on `/groups` — Full Client-Side Scan

```vba
Request.Resource = "/groups"   ' fetches ALL groups; filters client-side
```

For large tenants this can be thousands of objects across multiple pages. Use server-side filtering:

```vba
Request.AddQuerystringParam "$filter", "displayName eq '" & sGroup & "'"
Request.AddQuerystringParam "$select", "id,displayName"
Request.AddQuerystringParam "$count", "true"
Request.SetHeader "ConsistencyLevel", "eventual"
```

### 3.4 No `$select` Projections — All Fields Fetched Unnecessarily

Every GET request omits `$select`, causing Graph to return all fields (including large body content for messages). Per Microsoft best practices: _"Choose only the properties your application really needs."_

Example for listing contacts:

```vba
Request.AddQuerystringParam "$select", "id,displayName,emailAddresses,businessPhones"
```

### 3.5 No `client-request-id` Header — Diagnostics Impossible

Microsoft support correlates server-side traces using the `client-request-id` request header. Without it, troubleshooting failed API calls is very difficult.

```vba
Request.SetHeader "client-request-id", CreateGUID()
```

### 3.6 `offline_access` Scope Commented Out — No Refresh Token

```vba
'        Auth.AddScope "offline_access"  'if using Refresh Token
```

Without `offline_access` the authorization server does not issue a refresh token. When the access token expires (~1 hour), the only recovery is a full interactive re-login. Uncomment this scope and implement the refresh token flow.

### 3.7 `transactionId` Commented Out — Non-Idempotent Event Creation

```vba
'.AddBodyParameter "transactionId", CreateGUID()   ' ← commented out
```

Without a `transactionId`, a network failure mid-request leaves the client unable to determine whether the event was created. Retrying creates duplicate calendar events. Uncomment this after fixing `CreateGUID` (§2.5).

### 3.8 API Base URL Hardcoded — No Path to `/beta` or Version Override

```vba
pClient.BaseUrl = "https://graph.microsoft.com/v1.0"
```

Extract to a named constant and expose as an overridable property:

```vba
Private Const GRAPH_BASE_URL As String = "https://graph.microsoft.com/v1.0"
```

---

## 4. Performance and Robustness Gaps

### 4.1 `ConvertFileToBase64` Leaks `ADODB.Stream` on Error (`AttachmentHelpers.bas`)

If `.LoadFromFile` fails (file not found, permission denied), the `With` block exits without calling `.Close`, leaking the COM reference.

```vba
' Fix:
Public Function ConvertFileToBase64(sPath As String) As String
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    On Error GoTo Cleanup
    oStream.Open
    oStream.Type = 1   ' adTypeBinary
    oStream.LoadFromFile sPath
    Dim bytes As Variant
    bytes = oStream.Read
    oStream.Close
    ConvertFileToBase64 = EncodeBase64(bytes)
    GoTo Done
Cleanup:
    On Error Resume Next
    oStream.Close
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
Done:
    Set oStream = Nothing
End Function
```

### 4.2 No 3 MB File Size Guard Before Inline Attachment (`AttachmentHelpers.bas`)

Graph API rejects `fileAttachment` objects larger than **3 MB**. Files larger than that require the [upload session API](https://learn.microsoft.com/en-us/graph/api/attachment-createuploadsession). No size check exists.

```vba
' Add before ConvertFileToBase64 call:
Const MAX_INLINE_BYTES As Long = 3 * 1024 * 1024   ' 3 MB
If FileLen(sPath) > MAX_INLINE_BYTES Then
    Err.Raise vbObjectError + 9001, "FillAttachmentCollection", _
        "File """ & sPath & """ exceeds 3 MB. Use upload session API for large attachments."
End If
```

### 4.3 No `Reset()` Method — Singleton `pClient` Never Cleared (`Graph.bas`)

If `AdminTable` credentials change at runtime, the cached `pClient` continues with stale credentials forever.

```vba
' Add to Graph.bas:
Public Sub Reset()
    Set pClient = Nothing
    pClientId = ""
    pTenantID = ""
    pClientSecret = ""
    pGrantType = ""
End Sub
```

### 4.4 `expires_in` Never Stored — Token Only Cleared Reactively (`GraphAuthenticator.cls`)

The token response includes `expires_in` (seconds until expiry), but it is never stored or used. Tokens are only cleared after a 401 — meaning in-flight requests can fail mid-execution when the token expires.

```vba
' In GraphAuthenticator.cls, add:
Private pTokenExpiry As Date

' In GetToken, after successful response:
pTokenExpiry = DateAdd("s", CLng(auth_Response.Data("expires_in")) - 60, Now)

' In IWebAuthenticator_BeforeExecute, replace:
If Me.Token = "" Then
' With:
If Me.Token = "" Or Now >= pTokenExpiry Then
    Me.Token = Me.GetToken(Client)
End If
```

### 4.5 Redundant `DLookup("GrantType", "AdminTable")` in 8+ Functions (`Graph.bas`)

```vba
If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable")
```

The `Client` property getter already sets `pGrantType` during lazy initialisation. This guard fires only if `pClient` was somehow reset without using `Reset()` (no such code path exists). These `DLookup` calls are unnecessary DAO overhead in every public function call. Trust the `Client` getter to initialise all module-level state; remove the per-function guards.

### 4.6 Fragile `InStr(..., "expired")` Token Expiry Detection (`Graph.bas`)

Used in 8+ retry loops. The AAD error body wording can change. Prefer checking the structured error code from the JSON response:

```vba
' More robust check using AAD error codes:
If Response.StatusCode = WebStatusCode.Unauthorized Then
    ' AAD returns error_codes like 700082 (refresh token expired), 70008, etc.
    If Response.Data.Exists("error") Then
        If InStr(Response.Data("error")("code"), "InvalidAuthenticationToken") > 0 _
           Or InStr(Response.Data("error")("code"), "TokenExpired") > 0 Then
            ' clear and retry
        End If
    End If
End If
```

### 4.7 `IntArrayToString` O(n²) Concatenation (`TimeZoneHelpers.bas`)

```vba
' Current — each & allocates a new string copy:
For N = LBound(V) To UBound(V)
    S = S & Chr(V(N))
Next N
```

For the 32-element timezone name array this is negligible, but the pattern is incorrect. Use `StrConv` on a byte array:

```vba
Function IntArrayToString(V As Variant) As String
    Dim buf() As Byte
    Dim N As Long
    ReDim buf(LBound(V) To UBound(V))
    For N = LBound(V) To UBound(V)
        buf(N) = V(N) And 255
    Next N
    ' Trim trailing null chars from fixed-length Win32 string buffer
    IntArrayToString = Left$(StrConv(buf, vbUnicode), _
                             InStr(StrConv(buf, vbUnicode), Chr$(0)) - 1)
End Function
```

---

## 5. Summary Table

| # | File | Issue | Severity |
| --- | ------ | ------- | ---------- |
| 1.1 | `OAuth2Authenticator.cls` | `client_secret` in URL querystring | **CRITICAL** |
| 1.2 | `Graph.bas` | `ClientSecret` plaintext in `AdminTable` | **CRITICAL** |
| 1.3 | `OAuth2Authenticator.cls` | ROPC (`grant_type=password`) flow deprecated | HIGH |
| 1.4 | `GraphAuthenticator.cls` | `Public Token As String` — no encapsulation | HIGH |
| 2.1 | `GraphAuthenticator.cls` | `Scopes(1)` off-by-one — crash or dropped scope | **CRITICAL** |
| 2.2 | `Graph.bas` | Infinite retry loop on persistent 401 | HIGH |
| 2.3 | `Graph.bas` | `GetContactFolderID`/`GetMailFolderID` infinite loop when not found | HIGH |
| 2.4 | `Graph.bas` | `Replace(fillString, " ", "")` corrupts paths with spaces | HIGH |
| 2.5 | `Graph.bas` | `Rnd` without `Randomize` — deterministic GUIDs | HIGH |
| 2.6 | `Graph.bas` | `"Top"` instead of `"$top"` — `$top` always ignored | MEDIUM |
| 2.7 | `Graph.bas` | `allowNewTimeProposals` sent as string not Boolean | MEDIUM |
| 2.8 | `Graph.bas` | `Dim Auth As New` before `Set Auth =` — misleading | LOW |
| 2.9 | `GraphAuthenticator.cls` | Dead `Exit Do` after `Err.Raise` | LOW |
| 2.10 | `GraphAuthenticator.cls` | `sSessionID` never set — Logout URL always invalid | LOW |
| 3.1 | `Graph.bas` | No `$filter` on `/groups` — full client-side scan | HIGH |
| 3.2 | `Graph.bas` | No HTTP 429 / `Retry-After` handling | HIGH |
| 3.3 | `Graph.bas` | No `@odata.nextLink` pagination — data truncated | HIGH |
| 3.4 | `Graph.bas` | No `$select` projections — all fields fetched | MEDIUM |
| 3.5 | `Graph.bas` | No `client-request-id` header — diagnostics blocked | MEDIUM |
| 3.6 | `GraphAuthenticator.cls` | `offline_access` commented out — no refresh token | MEDIUM |
| 3.7 | `Graph.bas` | `transactionId` commented out — duplicate events on retry | MEDIUM |
| 3.8 | `Graph.bas` | Base URL hardcoded with no abstraction | LOW |
| 4.1 | `AttachmentHelpers.bas` | `ADODB.Stream` COM leak on error | MEDIUM |
| 4.2 | `AttachmentHelpers.bas` | No 3 MB attachment size guard | MEDIUM |
| 4.3 | `Graph.bas` | No `Reset()` method for singleton `pClient` | LOW |
| 4.4 | `GraphAuthenticator.cls` | `expires_in` never stored — reactive-only token refresh | MEDIUM |
| 4.5 | `Graph.bas` | Redundant `DLookup` in 8+ functions | LOW |
| 4.6 | `Graph.bas` | Fragile `InStr(..., "expired")` for error detection | MEDIUM |
| 4.7 | `TimeZoneHelpers.bas` | O(n²) string concatenation + trailing null chars | LOW |
