# VBA-MicrosoftGraph

VBA-MicrosoftGraph makes working with the [Microsoft Graph REST API](https://learn.microsoft.com/graph/overview) from VBA easier. Built on [VBA-Web](https://github.com/VBA-tools/VBA-Web) (v4.1.6), it handles Microsoft Identity (OAuth2) authentication and provides a high-level API for mail, calendar, contacts, and groups — all from Microsoft Access VBA.

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif)](https://www.paypal.com/donate/?hosted_button_id=E4YQV7RBUKHAQ)

## Features

### Mail

| Function | Description |
| --- | --- |
| `CreateDraftMessage` | Create a draft message with To/CC/BCC recipients and file attachments |
| `SendDraft` | Send a previously created draft by message ID |
| `GraphSendMail` | Compose and send a message in a single call |
| `SendMailAs` | Send email as / on behalf of another address |
| `ListMessages` | List messages in a mail folder with optional `$select` projection |
| `DeleteMessage` | Delete a message by ID |
| `GetMailFolderID` | Resolve a mail folder name to its Graph ID |

### Calendar

| Function | Description |
| --- | --- |
| `CreateEvent` | Create a calendar event with attendees, location, and recurrence support |
| `ListEvents` | List events with optional date-range (`calendarView`) and `$select` |
| `UpdateEvent` | PATCH fields on an existing calendar event |
| `DeleteEvent` | Delete an event by ID |
| `GetCalendarGroupID` | Resolve a calendar group name to its Graph ID |
| `GetCalendarID` | Resolve a calendar name to its Graph ID |

### Contacts

| Function | Description |
| --- | --- |
| `CreateContact` | Create a contact in a specified folder |
| `ListContacts` | List contacts with optional `$select` projection |
| `UpdateContact` | PATCH fields on an existing contact |
| `GetContactFolderID` | Resolve a contact folder name to its Graph ID |

### Groups

| Function | Description |
| --- | --- |
| `GetGroupID` | Find a group by display name using server-side `$filter` |

### User Profile & Directory

| Function | Description |
| --- | --- |
| `GetCurrentUser` | Get the signed-in user's profile (`/me`) |
| `ListGroupMembership` | List groups and roles a user belongs to |
| `GetManager` | Get a user's manager |
| `ListDirectReports` | List a user's direct reports |

### Tasks & Planner

| Function | Description |
| --- | --- |
| `CreateTask` | Create a To-Do task with optional body and due date |
| `ListPlannerTasks` | List the user's Planner tasks |

### Teams & Online Meetings

| Function | Description |
| --- | --- |
| `ListJoinedTeams` | List teams the user has joined |
| `ListTeamsChannels` | List channels in a team |
| `CreateOnlineMeeting` | Create a Teams online meeting |
| `SendChannelMessage` | Send a message to a Teams channel |
| `ReplyToChannelMessage` | Reply to a message in a Teams channel |
| `SendChatMessage` | Send a direct/group chat message |
| `ListChats` | List the user's chats (1:1, group, meeting) |
| `ListChannelMessages` | List messages in a Teams channel |
| `CreateChat` | Create a new 1:1 or group chat |

### SharePoint & OneDrive

| Function | Description |
| --- | --- |
| `ListSharePointSites` | Search for SharePoint sites |
| `SearchOneDrive` | Search files in a user's OneDrive |
| `SearchSharePoint` | Search across SharePoint and OneDrive via `/search/query` |

### OneNote

| Function | Description |
| --- | --- |
| `ListOneNoteNotebooks` | List a user's OneNote notebooks |

### Configuration

| Function | Description |
| --- | --- |
| `UseGraphBeta` | Switch the base URL to the Microsoft Graph beta endpoint |
| `GraphReset` | Clear cached client/auth state for a fresh session |
| `ClearAuthCodes` | Clear stored authorization codes |
| `Logout` | Sign out and clear tokens |

## Authentication

Three OAuth2 flows are supported — configure via the `GrantType` field in `AdminTable`:

| Flow | `GrantType` Value | Use Case |
| --- | --- | --- |
| **Authorization Code** | `authorization_code` | Interactive sign-in (uses Edge WebView2 in Access) |
| **Client Credentials** | `client_credentials` | Unattended / service-to-service |
| **Device Code** | `device_code` | Headless / environments without a browser control |

Token lifecycle is handled automatically:

- **Refresh tokens** — when `offline_access` scope is enabled, expired tokens are silently refreshed
- **Proactive refresh** — tokens are refreshed 5 minutes before expiry to avoid mid-request failures

## Getting Started

### Prerequisites

- Microsoft Access (32-bit or 64-bit)
- An Azure AD app registration with Microsoft Graph API permissions
- VBA Trust Center → *Trust access to the VBA project object model* enabled

### Setup

1. Download or clone this repository
2. Import the modules and class modules into your Access database:
   - `Src/Graph.bas` — main API facade
   - `Src/AttachmentHelpers.bas` — file-to-Base64 conversion
   - `Src/TimeZoneHelpers.bas` — time zone mapping for calendar events
   - `Src/WebHelpers.bas`, `Src/WebClient.cls`, `Src/WebRequest.cls`, `Src/WebResponse.cls`, `Src/Dictionary.cls` — VBA-Web HTTP stack
   - `authenticators/GraphAuthenticator.cls` — Microsoft Graph OAuth2 authenticator
   - `authenticators/OAuth2Authenticator.cls` — generic OAuth2 authenticator
3. Create an `AdminTable` in your database with these fields:

   | Field | Description |
   | --- | --- |
   | `ClientID` | Application (client) ID from Azure AD |
   | `TenantID` | Directory (tenant) ID |
   | `ClientSecret` | Client secret value |
   | `GrantType` | `authorization_code`, `client_credentials`, or `device_code` |
   | `WaitForLogin` | Login timeout in seconds (default: 60) |

4. Import the sample forms from `Forms/` for a working demo

### Quick Example

```vba
' Send an email
Dim oResponse As WebResponse
Set oResponse = GraphSendMail( _
    "user@contoso.com", _
    "Subject line", _
    "HTML", _
    "<p>Hello from VBA!</p>", _
    "recipient@contoso.com", _
    "", "", "")

' List messages in Inbox
Dim sJson As String
sJson = ListMessages("user@contoso.com", "Inbox", "$select=subject,from,receivedDateTime")

' Create a calendar event
Set oResponse = CreateEvent( _
    "user@contoso.com", _
    "Team Meeting", "HTML", "<p>Agenda</p>", _
    Date, TimeValue("09:00"), Date, TimeValue("10:00"), _
    "Conference Room", "attendee@contoso.com", "", "", "")
```

## Project Structure

```text
├── Src/
│   ├── Graph.bas                   Main API — all Graph operations
│   ├── PKCE.bas                    Proof Key for Code Exchange (RFC 7636)
│   ├── AttachmentHelpers.bas       Base64 encoding for file attachments
│   ├── TimeZoneHelpers.bas         Windows ↔ IANA time zone mapping
│   ├── WebHelpers.bas              VBA-Web utilities
│   ├── WebClient.cls               HTTP client (WinHttp)
│   ├── WebRequest.cls              HTTP request builder
│   ├── WebResponse.cls             HTTP response wrapper
│   └── Dictionary.cls              VBA Dictionary polyfill
├── authenticators/
│   ├── GraphAuthenticator.cls      Graph-specific OAuth2 (auth code, client creds, device code)
│   └── OAuth2Authenticator.cls     Generic OAuth2 base authenticator
├── Forms/
│   ├── Form_SendEmail.cls          Send email UI
│   ├── Form_CreateEvent.cls        Create event UI
│   ├── Form_CreateContact.cls      Create contact UI
│   ├── Form_ListMessages.cls       List messages UI
│   └── Form_ListContacts.cls       List contacts UI
├── AccessAndGraph.accdb            Sample database with all modules + forms
├── EVALUATION.md                   Technical audit and change log
└── LICENSE                         MIT License
```

## What's New in v2.3

### Teams Messaging

- **Send messages** — `SendChannelMessage`, `ReplyToChannelMessage`, `SendChatMessage`
- **Read messages** — `ListChats`, `ListChannelMessages`
- **Create chats** — `CreateChat` with semicolon-delimited user IDs/UPNs, supports 1:1 and group chat types
- **Scopes** — `ChannelMessage.Send`, `ChatMessage.Send`, `Chat.ReadBasic`, `Chat.Create`, `ChannelMessage.Read.All` added to authorization_code flow
- **Delegated only** — Teams messaging requires interactive sign-in (authorization_code or device_code); client_credentials cannot send messages

---

## What's New in v2.2

### Expanded API Coverage

- **User & Directory** — `GetCurrentUser`, `ListGroupMembership`, `GetManager`, `ListDirectReports`
- **Tasks & Planner** — `CreateTask` (To-Do), `ListPlannerTasks`
- **Teams** — `ListJoinedTeams`, `ListTeamsChannels`, `CreateOnlineMeeting`
- **SharePoint & OneDrive** — `ListSharePointSites`, `SearchOneDrive`, `SearchSharePoint` (via `/search/query` POST)
- **OneNote** — `ListOneNoteNotebooks`
- **Mail** — `SendMailAs` (send as / on behalf of another address)
- **PATCH operations** — `UpdateEvent`, `UpdateContact` for partial updates

### PKCE Support (RFC 7636)

- New `PKCE.bas` module with SHA-256 via Windows CryptoAPI + Base64URL encoding
- `EnablePKCE()` / `DisablePKCE()` toggle PKCE for the authorization code flow
- When enabled, `code_challenge` (S256) is appended to the authorization URL
- Token exchange sends `code_verifier` instead of `client_secret`
- Eliminates the need for a client secret in public client apps

### Utility Helpers

- `EscapeJsonString()` — safe JSON string escaping (correct backslash-first order)
- `BuildResourcePath()` — auto-switches `/me` vs `/users/{UPN}` based on grant type

---

## What's New in v2.1

### Security Fixes

- **Client secret moved to POST body** — previously leaked in URL querystrings (proxy logs, Referer headers)
- **ROPC flow blocked** — `grant_type=password` now raises an error; use `authorization_code` or `client_credentials` instead
- **Bearer token encapsulated** — `Token` is now a read-only property (was `Public`)

### Correctness Fixes

- **Scopes join fix** — `VBA.Join(Scopes, " ")` replaces `Scopes(1)` which crashed with a single scope
- **Retry loops capped** — all API functions limited to 3 retries (previously could loop forever)
- **Folder lookup infinite loops fixed** — `GetContactFolderID` / `GetMailFolderID` return `""` on not-found
- **File path spaces preserved** — `Trim()` replaces `Replace(..., " ", "")` which corrupted paths
- **`$top` OData prefix fixed** — was sent as `Top` (silently ignored by Graph)
- **`allowNewTimeProposals`** — sent as Boolean `True` instead of string `"true"`
- **`CreateGUID` seeded** — `Randomize` added to prevent deterministic GUIDs across sessions
- **ADODB.Stream resource leak fixed** — proper cleanup in `ConvertFileToBase64`

### New API Functions

- `SendDraft` — send a previously saved draft message
- `DeleteMessage` — delete a message by ID
- `DeleteEvent` — delete a calendar event by ID
- `ListEvents` — list events with optional date-range filtering via `calendarView`

### Graph API Best Practices

- **Pagination** — `@odata.nextLink` automatically followed to retrieve all pages
- **Throttling** — HTTP 429 responses honoured with `Retry-After` backoff
- **`$select` projection** — `ListMessages`, `ListContacts`, and `ListEvents` accept an optional `$select` parameter to reduce payload size
- **Server-side `$filter`** — `GetGroupID` filters on the server with `ConsistencyLevel: eventual`
- **`client-request-id` header** — added to all API calls for request tracing
- **`transactionId`** — enabled on `CreateEvent` for idempotent event creation
- **3 MB attachment guard** — `ConvertFileToBase64` rejects files over 3 MB (Graph inline limit)

### Authentication Upgrades

- **Device Code Flow** — `LoginWithDeviceCode` for headless environments
- **Refresh token lifecycle** — automatic silent refresh when `offline_access` scope is enabled
- **Proactive token refresh** — tokens refreshed 5 minutes before expiry
- **Structured error parsing** — replaces fragile `InStr(Content, "expired")` string matching

### Developer Experience

- **Configurable base URL** — `GraphBaseUrl` Property Let/Get with `UseGraphBeta()` shortcut
- **Session reset** — `GraphReset()` clears all cached state
- **Centralized config helpers** — `GetClientID()`, `GetTenantID()`, `GetClientSecret()` for single-point credential management

See [EVALUATION.md](EVALUATION.md) for the full technical audit.

## About

- Original author: Maria Barnes
- v2.1 maintained by: John Colozzi and Github Copilot
- License: MIT
