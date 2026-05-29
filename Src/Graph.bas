Attribute VB_Name = "Graph"
Option Compare Database
Option Explicit

Private Const MAX_RETRIES As Long = 3

Private pClient As WebClient
Private pGraphBaseUrl As String
Private pClientId As String
Private pTenantID As String
Private pClientSecret As String
Private pWaitForLogin As Integer
Private pGrantType As String

Private Function GetClientID() As String
    GetClientID = Nz(DLookup("ClientID", "AdminTable"), "")
End Function

Private Function GetTenantID() As String
    GetTenantID = Nz(DLookup("TenantID", "AdminTable"), "")
End Function

Private Function GetClientSecret() As String
    ' TODO: Replace with Windows Credential Manager (CredRead API) for production
    ' For now, centralize the lookup so it can be swapped out in one place
    GetClientSecret = Nz(DLookup("ClientSecret", "AdminTable"), "")
End Function

Public Property Get GraphBaseUrl() As String
    If Len(pGraphBaseUrl) = 0 Then
        pGraphBaseUrl = "https://graph.microsoft.com/v1.0"
    End If
    GraphBaseUrl = pGraphBaseUrl
End Property

Public Property Let GraphBaseUrl(sUrl As String)
    pGraphBaseUrl = sUrl
End Property

Public Property Get Client() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = GraphBaseUrl
        pClientId = GetClientID() 'Application (client) ID
        pTenantID = GetTenantID() 'Directory (Tenant) ID
        pClientSecret = GetClientSecret() 'Client Secret
        pWaitForLogin = DLookup("WaitForLogin", "AdminTable") 'Login wait period defaults to 60 seconds
        pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
        
        Dim Auth As GraphAuthenticator
        Set Auth = New GraphAuthenticator
        Auth.Setup pClientId, pTenantID, pClientSecret, pWaitForLogin, pGrantType
'        Auth.AddScope "offline_access"  'if using Refresh Token
        If pGrantType = "authorization_code" Then
            Auth.AddScope "mail.readwrite"
            Auth.AddScope "mail.send"
            Auth.AddScope "calendars.readwrite"
            Auth.AddScope "calendars.readwrite.shared"
            Auth.AddScope "group.readwrite.all"
            Auth.AddScope "contacts.readwrite"
        Else
            Auth.AddScope ".default"
        End If
        Auth.AuthorizationUrl = "https://login.microsoftonline.com/" & pTenantID & "/oauth2/v2.0/authorize"
'        Call Auth.Login
        Set pClient.Authenticator = Auth
    End If
    
    Set Client = pClient
End Property

Public Sub ClearAuthCodes()
    Dim Auth As GraphAuthenticator
    Set Auth = Client.Authenticator
    Call Auth.ClearCodes
End Sub

Public Sub Logout()
    Dim Auth As GraphAuthenticator
    Set Auth = Client.Authenticator
    Call Auth.Logout
End Sub

Public Sub UseGraphBeta()
    GraphBaseUrl = "https://graph.microsoft.com/beta"
End Sub

Public Sub GraphReset()
    Set pClient = Nothing
    pGraphBaseUrl = ""
End Sub

Private Function GetRetryAfterSeconds(oResponse As WebResponse) As Long
    Dim sRetryAfter As String
    On Error Resume Next
    sRetryAfter = oResponse.Headers("Retry-After")
    On Error GoTo 0
    
    If Len(sRetryAfter) > 0 And IsNumeric(sRetryAfter) Then
        GetRetryAfterSeconds = CLng(sRetryAfter)
    Else
        GetRetryAfterSeconds = 10
    End If
End Function

Private Function IsTokenExpiredError(oResponse As WebResponse) As Boolean
    On Error Resume Next
    If oResponse.StatusCode = WebStatusCode.Unauthorized Then
        If oResponse.Data.Exists("error") Then
            Dim oErr As Object
            Set oErr = oResponse.Data("error")
            If oErr.Exists("code") Then
                IsTokenExpiredError = (oErr("code") = "InvalidAuthenticationToken")
            End If
        End If
    End If
    On Error GoTo 0
End Function

Private Function GraphPagedGet(oClient As WebClient, oInitialRequest As WebRequest, _
    Optional lMaxPages As Long = 0) As Collection
    
    Dim colResults As New Collection
    Dim oResponse As WebResponse
    Dim sNextLink As String
    Dim lPageCount As Long
    Dim oItem As Variant
    
    Set oResponse = oClient.Execute(oInitialRequest)
    
    If oResponse.StatusCode <> WebStatusCode.Ok Then
        Set GraphPagedGet = colResults
        Exit Function
    End If
    
    If oResponse.Data.Exists("value") Then
        For Each oItem In oResponse.Data("value")
            colResults.Add oItem
        Next oItem
    End If
    lPageCount = 1
    
    Do While oResponse.Data.Exists("@odata.nextLink")
        If lMaxPages > 0 And lPageCount >= lMaxPages Then Exit Do
        
        sNextLink = oResponse.Data("@odata.nextLink")
        
        Dim oNextRequest As New WebRequest
        oNextRequest.Resource = Replace(sNextLink, GraphBaseUrl, "")
        oNextRequest.Method = WebMethod.HttpGet
        oNextRequest.Format = WebFormat.JSON
        
        Set oResponse = oClient.Execute(oNextRequest)
        
        If oResponse.StatusCode = 429 Then
            Dim lRetryAfter As Long
            lRetryAfter = GetRetryAfterSeconds(oResponse)
            Application.Wait Now + TimeSerial(0, 0, lRetryAfter)
            Set oResponse = oClient.Execute(oNextRequest)
        End If
        
        If oResponse.StatusCode = WebStatusCode.Ok Then
            If oResponse.Data.Exists("value") Then
                For Each oItem In oResponse.Data("value")
                    colResults.Add oItem
                Next oItem
            End If
            lPageCount = lPageCount + 1
        Else
            Exit Do
        End If
    Loop
    
    Set GraphPagedGet = colResults
End Function


Public Function CreateDraftMessage(UserPrincipal As String, Subject As String, BodyType As String, BodyContent As String, toRecipients As String, ccRecipients As String, bccRecipients As String, AttachmentPath As String) As String
    Dim Request As New WebRequest
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    Request.Resource = Request.Resource & "/messages"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    
    Dim body As Dictionary
    Set body = New Dictionary
    body.Add "contentType", BodyType
    body.Add "content", BodyContent
    
    'Since toRecipients can be a list of email addresses
    Dim recipients As Collection
    Set recipients = New Collection
    FillEmailAddressCollection recipients, toRecipients
    
    Dim COLccRecipients As Collection
    If Trim(ccRecipients) <> "" Then
        Set COLccRecipients = New Collection
        FillEmailAddressCollection COLccRecipients, ccRecipients
    End If
    
    Dim COLbccRecipients As Collection
    If Trim(bccRecipients) <> "" Then
        Set COLbccRecipients = New Collection
        FillEmailAddressCollection COLbccRecipients, bccRecipients
    End If
    
    'Since attachments can be a list of attachments
    Dim attachments As Collection
    If Trim(AttachmentPath) <> "" Then
        Set attachments = New Collection
        FillAttachmentCollection attachments, AttachmentPath
    End If
    
    With Request
        .AddBodyParameter "subject", Subject
        .AddBodyParameter "body", body
        .AddBodyParameter "toRecipients", recipients
        If Trim(ccRecipients) <> "" Then .AddBodyParameter "ccRecipients", COLccRecipients
        If Trim(bccRecipients) <> "" Then .AddBodyParameter "bccRecipients", COLbccRecipients
        If Trim(AttachmentPath) <> "" Then .AddBodyParameter "attachments", attachments
    End With
    
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim oResponse As WebResponse
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set oResponse = Client.Execute(Request)
        If oResponse.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(oResponse))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(oResponse) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.CreateDraftMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
    
    If oResponse.StatusCode = WebStatusCode.Created Then
        CreateDraftMessage = oResponse.Data("id")
    End If

End Function

Private Sub FillAttachmentCollection(ByVal fillCollection As Collection, fillString As String)
    Dim sFileName As String
    Dim sFilePath As String
    Dim attachment As Dictionary
    
    fillString = Trim(fillString)
    If Mid(fillString, Len(fillString)) = ";" Then
        fillString = Left(fillString, Len(fillString) - 1)
    End If
    While InStr(fillString, ";") > 0
        Set attachment = New Dictionary
        attachment.Add "@odata.type", "#microsoft.graph.fileAttachment"
        sFilePath = Trim(Left(fillString, InStr(fillString, ";") - 1))
        fillString = Mid(fillString, InStr(fillString, ";") + 1)
        sFileName = Mid(sFilePath, InStrRev(sFilePath, "\") + 1)
        attachment.Add "name", sFileName
'        attachment.Add "contentType", "text/plain" 'Not mandatory so leave off for flexibility
        attachment.Add "contentBytes", ConvertFileToBase64(sFilePath)
        fillCollection.Add attachment
    Wend
    Set attachment = New Dictionary
    attachment.Add "@odata.type", "#microsoft.graph.fileAttachment"
    sFileName = Mid(fillString, InStrRev(fillString, "\") + 1)
    attachment.Add "name", sFileName
'   attachment.Add "contentType", "text/plain" 'Not mandatory so leave off for flexibility
    attachment.Add "contentBytes", ConvertFileToBase64(fillString)
    fillCollection.Add attachment
End Sub

Private Sub FillEmailAddressCollection(ByVal fillCollection As Collection, fillString As String)
    Dim sAddress As String
    Dim EmailAddress As Dictionary
    
    fillString = Replace(fillString, " ", "")
    If Mid(fillString, Len(fillString)) = ";" Then
        fillString = Left(fillString, Len(fillString) - 1)
    End If
    While InStr(fillString, ";") > 0
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        sAddress = Trim(Left(fillString, InStr(fillString, ";") - 1))
        fillString = Mid(fillString, InStr(fillString, ";") + 1)
        EmailAddress.Item("emailAddress").Add "address", sAddress
        fillCollection.Add EmailAddress
    Wend
    Set EmailAddress = New Dictionary
    EmailAddress.Add "emailAddress", New Dictionary
    EmailAddress.Item("emailAddress").Add "address", fillString
    fillCollection.Add EmailAddress
End Sub

Public Function GraphSendMail(UserPrincipal As String, Subject As String, BodyType As String, BodyContent As String, toRecipients As String, ccRecipients As String, bccRecipients As String, AttachmentPath As String) As WebResponse
    Dim Request As New WebRequest
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    Request.Resource = Request.Resource & "/sendMail"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    
    Dim message As Dictionary
    Set message = New Dictionary
    
    Dim body As Dictionary
    Set body = New Dictionary
    body.Add "contentType", BodyType
    body.Add "content", BodyContent
    
    'Since toRecipients can be a list of email addresses
    Dim recipients As Collection
    Set recipients = New Collection
    FillEmailAddressCollection recipients, toRecipients
    
    Dim COLccRecipients As Collection
    If Trim(ccRecipients) <> "" Then
        Set COLccRecipients = New Collection
        FillEmailAddressCollection COLccRecipients, ccRecipients
    End If
    
    Dim COLbccRecipients As Collection
    If Trim(bccRecipients) <> "" Then
        Set COLbccRecipients = New Collection
        FillEmailAddressCollection COLbccRecipients, bccRecipients
    End If
    
    'Since attachments can be a list of attachments
    Dim attachments As Collection
    If Trim(AttachmentPath) <> "" Then
        Set attachments = New Collection
        FillAttachmentCollection attachments, AttachmentPath
    End If
    
    With message
        .Add "subject", Subject
        .Add "body", body
        .Add "toRecipients", recipients
        If Trim(ccRecipients) <> "" Then .Add "ccRecipients", COLccRecipients
        If Trim(bccRecipients) <> "" Then .Add "bccRecipients", COLbccRecipients
        If Trim(AttachmentPath) <> "" Then .Add "attachments", attachments
    End With
    
    Request.AddBodyParameter "message", message
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set GraphSendMail = Client.Execute(Request)
        If GraphSendMail.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(GraphSendMail))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(GraphSendMail) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.GraphSendMail", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If

End Function

Public Function CreateGUID() As String
    Randomize  ' Seed the random number generator
    Do While Len(CreateGUID) < 32
        If Len(CreateGUID) = 16 Then
            '17th character holds version information
            CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
        End If
        CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
    Loop
    CreateGUID = Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12)
End Function

Private Sub FillAttendeeCollection(ByVal fillCollection As Collection, fillStringReq As String, fillStringOpt As String)
    Dim sAddress As String
    Dim EmailAddress As Dictionary
    
    fillStringReq = Replace(fillStringReq, " ", "")
    If Len(fillStringReq) > 0 Then
        If Mid(fillStringReq, Len(fillStringReq)) = ";" Then
            fillStringReq = Left(fillStringReq, Len(fillStringReq) - 1)
        End If
    End If
    fillStringOpt = Replace(fillStringOpt, " ", "")
    If Len(fillStringOpt) > 0 Then
        If Mid(fillStringOpt, Len(fillStringOpt)) = ";" Then
            fillStringOpt = Left(fillStringOpt, Len(fillStringOpt) - 1)
        End If
    End If
    While InStr(fillStringReq, ";") > 0
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        sAddress = Trim(Left(fillStringReq, InStr(fillStringReq, ";") - 1))
        fillStringReq = Mid(fillStringReq, InStr(fillStringReq, ";") + 1)
        EmailAddress.Item("emailAddress").Add "address", sAddress
        EmailAddress.Add "type", "required"
        fillCollection.Add EmailAddress
    Wend
    If fillStringReq <> "" Then
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        EmailAddress.Item("emailAddress").Add "address", fillStringReq
        EmailAddress.Add "type", "required"
        fillCollection.Add EmailAddress
    End If
    While InStr(fillStringOpt, ";") > 0
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        sAddress = Trim(Left(fillStringOpt, InStr(fillStringOpt, ";") - 1))
        fillStringOpt = Mid(fillStringOpt, InStr(fillStringOpt, ";") + 1)
        EmailAddress.Item("emailAddress").Add "address", sAddress
        EmailAddress.Add "type", "optional"
        fillCollection.Add EmailAddress
    Wend
    If fillStringOpt <> "" Then
        Set EmailAddress = New Dictionary
        EmailAddress.Add "emailAddress", New Dictionary
        EmailAddress.Item("emailAddress").Add "address", fillStringOpt
        EmailAddress.Add "type", "optional"
        fillCollection.Add EmailAddress
    End If
End Sub

Public Function GetGroupID(UserPrincipal As String, sGroup As String) As String
    Dim Request As New WebRequest
    Dim Response As New WebResponse
    Request.Resource = "/groups"
    Request.AddQuerystringParam "$filter", "displayName eq '" & sGroup & "'"
    Request.AddHeader "ConsistencyLevel", "eventual"
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim lRetryCount As Long
    GetGroupID = "Retry"
    lRetryCount = 0
    While GetGroupID = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set Response = Client.Execute(Request)
        If Response.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(Response))
            GetGroupID = "Retry"
        ElseIf Response.StatusCode = WebStatusCode.OK Then
            Dim GroupInfo As Dictionary
            For Each GroupInfo In Response.Data("value")
                GetGroupID = GroupInfo("id")
                Exit Function
            Next GroupInfo
            GetGroupID = "Error"
        ElseIf IsTokenExpiredError(Response) Then
            ClearAuthCodes
            GetGroupID = "Retry"
        Else
            MsgBox "Error " & Response.StatusCode & ": " & Response.Content
            GetGroupID = "Error"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.GetGroupID", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function GetCalendarGroupID(UserPrincipal As String, sCalendarGroup As String) As String
    Dim Request As New WebRequest
    Dim Response As New WebResponse
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    Request.Resource = Request.Resource & "/calendarGroups"
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim lRetryCount As Long
    GetCalendarGroupID = "Retry"
    lRetryCount = 0
    While GetCalendarGroupID = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set Response = Client.Execute(Request)
        If Response.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(Response))
            GetCalendarGroupID = "Retry"
        ElseIf Response.StatusCode = WebStatusCode.OK Then
            Dim CalendarInfo As Dictionary
            For Each CalendarInfo In Response.Data("value")
                If CalendarInfo("name") = sCalendarGroup Then
                    GetCalendarGroupID = CalendarInfo("id")
                    Exit Function
                End If
            Next CalendarInfo
            GetCalendarGroupID = "Error"
        ElseIf IsTokenExpiredError(Response) Then
            ClearAuthCodes
            GetCalendarGroupID = "Retry"
        Else
            MsgBox "Error " & Response.StatusCode & ": " & Response.Content
            GetCalendarGroupID = "Error"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.GetCalendarGroupID", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function GetCalendarID(UserPrincipal As String, sCalendarName As String, sCalendarGroup As String) As String
    Dim Request As New WebRequest
    Dim Response As New WebResponse
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    If sCalendarGroup <> "" Then
        If sCalendarGroup = "Groups" Then
            Request.Resource = "/groups/" & GetGroupID(UserPrincipal, sCalendarName) & "/calendar"
        Else
            Request.Resource = Request.Resource & "/calendarGroups/" & GetCalendarGroupID(UserPrincipal, sCalendarGroup) & "/calendars"
        End If
    Else
        Request.Resource = Request.Resource & "/calendars"
    End If
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim lRetryCount As Long
    GetCalendarID = "Retry"
    lRetryCount = 0
    While GetCalendarID = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set Response = Client.Execute(Request)
        If Response.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(Response))
            GetCalendarID = "Retry"
        ElseIf Response.StatusCode = WebStatusCode.OK Then
            Dim CalendarInfo As Dictionary
            If sCalendarGroup = "Groups" Then
                GetCalendarID = Response.Data("id")
            Else
                For Each CalendarInfo In Response.Data("value")
                    If CalendarInfo("name") = sCalendarName Then
                        GetCalendarID = CalendarInfo("id")
                        Exit Function
                    End If
                Next CalendarInfo
                GetCalendarID = "Error"
            End If
        ElseIf IsTokenExpiredError(Response) Then
            ClearAuthCodes
            GetCalendarID = "Retry"
        Else
            MsgBox "Error " & Response.StatusCode & ": " & Response.Content
            GetCalendarID = "Error"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.GetCalendarID", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function CreateEvent(UserPrincipal As String, Subject As String, BodyType As String, BodyContent As String, dStart As Date, tStart As Date, dEnd As Date, tEnd As Date, sLocation As String, sAttendees As String, sOptional As String, sCalendarGroup As String, sCalendarName As String) As WebResponse
    Dim Request As New WebRequest
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If sCalendarGroup = "Groups" Then
        Request.Resource = "groups/" & GetGroupID(UserPrincipal, sCalendarName) & "/calendar/events"
    Else
        If pGrantType = "authorization_code" Then
            Request.Resource = "/me"
        Else
            Request.Resource = "/users/" & UserPrincipal
        End If
        If sCalendarName <> "" Then
            'Get Calendar ID from sCalendarName
            Request.Resource = Request.Resource & "/calendars/" & GetCalendarID(UserPrincipal, sCalendarName, sCalendarGroup) & "/events"
        Else
            Request.Resource = Request.Resource & "/events"
        End If
    End If
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON

    Dim body As Dictionary
    Set body = New Dictionary
    body.Add "contentType", BodyType
    body.Add "content", BodyContent
    
    Dim start As Dictionary
    Set start = New Dictionary
    start.Add "dateTime", Format(dStart, "YYYY-MM-DD") + "T" + Format(tStart, "HH:MM:SS")
    start.Add "timeZone", Replace(CurrentTimeZone(), Chr(0), "")
    
    Dim enddic As Dictionary
    Set enddic = New Dictionary
    enddic.Add "dateTime", Format(dEnd, "YYYY-MM-DD") + "T" + Format(tEnd, "HH:MM:SS")
    enddic.Add "timeZone", Replace(CurrentTimeZone(), Chr(0), "")
    
    Dim location As Dictionary
    Set location = New Dictionary
    location.Add "displayName", sLocation
    
    'Since Attendees can be a list of email addresses
    Dim attendees As Collection
    Set attendees = New Collection
    FillAttendeeCollection attendees, sAttendees, sOptional
    
    With Request
        .AddBodyParameter "subject", Subject
        .AddBodyParameter "body", body
        .AddBodyParameter "start", start
        .AddBodyParameter "end", enddic
        If Trim(sLocation) <> "" Then .AddBodyParameter "location", location
        .AddBodyParameter "attendees", attendees
        .AddBodyParameter "allowNewTimeProposals", True
        .AddBodyParameter "transactionId", CreateGUID()
    End With
    
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateEvent = Client.Execute(Request)
        If CreateEvent.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateEvent))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateEvent) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.CreateEvent", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListContacts(UserPrincipal As String, sFolder As String, _
    Optional sSelectFields As String = "") As WebResponse
    Dim Request As New WebRequest
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    If sFolder = "" Then
        Request.Resource = Request.Resource & "/contacts"
    Else
        'Syntax for getting a specific folder uses the folder id so you have to get that first
        Request.Resource = Request.Resource & "/contactfolders/" & GetContactFolderID(UserPrincipal, sFolder) & "/contacts"
    End If
    Request.Method = WebMethod.HttpGET
    Request.Format = WebFormat.JSON
    Request.AddQuerystringParam "$top", 1000
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListContacts = Client.Execute(Request)
        If ListContacts.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListContacts))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListContacts) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.ListContacts", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function GetContactFolderID(UserPrincipal As String, sFolder As String) As String
    Dim Request As New WebRequest
    Dim Response As New WebResponse
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    Request.Resource = Request.Resource & "/contactfolders"
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim lRetryCount As Long
    GetContactFolderID = "Retry"
    lRetryCount = 0
    While GetContactFolderID = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set Response = Client.Execute(Request)
        If Response.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(Response))
            GetContactFolderID = "Retry"
        ElseIf Response.StatusCode = WebStatusCode.OK Then
            Dim FolderInfo As Dictionary
            For Each FolderInfo In Response.Data("value")
                If FolderInfo("displayName") = sFolder Then
                    GetContactFolderID = FolderInfo("id")
                    Exit Function
                End If
            Next FolderInfo
            GetContactFolderID = ""
        ElseIf IsTokenExpiredError(Response) Then
            ClearAuthCodes
            GetContactFolderID = "Retry"
        Else
            MsgBox "Error " & Response.StatusCode & ": " & Response.Content
            GetContactFolderID = "Error"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.GetContactFolderID", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function CreateContact(UserPrincipal As String, sFolder As String, givenName As String, surname As String, fileAs As String, jobTitle As String, companyName As String, sBusinessPhones As String, sEmailAddresses As String) As WebResponse
    Dim Request As New WebRequest
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    If sFolder = "" Then
        Request.Resource = Request.Resource & "/contacts"
    Else
        'Syntax for getting a specific folder uses the folder id so you have to get that first
        Request.Resource = Request.Resource & "/contactfolders/" & GetContactFolderID(UserPrincipal, sFolder) & "/contacts"
    End If
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    
    'Since emailAddresses can be a list of email addresses
    Dim emailAddresses As Collection
    Set emailAddresses = New Collection
    Dim sAddress As String
    Dim EmailAddress As Dictionary
    
    While InStr(sEmailAddresses, ";") > 0
        Set EmailAddress = New Dictionary
        sAddress = Trim(Left(sEmailAddresses, InStr(sEmailAddresses, ";") - 1))
        sEmailAddresses = Mid(sEmailAddresses, InStr(sEmailAddresses, ";") + 1)
        EmailAddress.Add "address", sAddress
        emailAddresses.Add EmailAddress
    Wend
    Set EmailAddress = New Dictionary
    EmailAddress.Add "address", sEmailAddresses
    emailAddresses.Add EmailAddress
    
    'Since businessPhones can be a list of phone numbers
    Dim businessPhones() As String
    Dim sPhone As String
    Dim iPhoneCount As Integer
    If Trim(sBusinessPhones) <> "" Then
        While InStr(sBusinessPhones, ";") > 0
            ReDim businessPhones(iPhoneCount)
            sPhone = Trim(Left(sBusinessPhones, InStr(sBusinessPhones, ";") - 1))
            sBusinessPhones = Mid(sBusinessPhones, InStr(sBusinessPhones, ";") + 1)
            businessPhones(iPhoneCount) = sPhone
            iPhoneCount = iPhoneCount + 1
        Wend
        ReDim businessPhones(iPhoneCount)
        businessPhones(iPhoneCount) = sBusinessPhones
        iPhoneCount = iPhoneCount + 1
    End If
    
    With Request
        .AddBodyParameter "givenName", givenName
        .AddBodyParameter "surname", surname
        .AddBodyParameter "fileAs", fileAs
        .AddBodyParameter "jobTitle", jobTitle
        .AddBodyParameter "companyName", companyName
        .AddBodyParameter "emailAddresses", emailAddresses
        If iPhoneCount > 0 Then .AddBodyParameter "businessPhones", businessPhones
    End With
    
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateContact = Client.Execute(Request)
        If CreateContact.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateContact))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateContact) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.CreateContact", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListMessages(UserPrincipal As String, sFolder As String, _
    Optional sSelectFields As String = "") As WebResponse
    Dim Request As New WebRequest
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    If sFolder = "" Then
        Request.Resource = Request.Resource & "/messages"
    Else
        'Syntax for getting a specific folder uses the folder id so you have to get that first
        Request.Resource = Request.Resource & "/mailfolders/" & GetMailFolderID(UserPrincipal, sFolder) & "/messages"
    End If
    Request.Method = WebMethod.HttpGET
    Request.Format = WebFormat.JSON
    Request.AddQuerystringParam "$top", 1000
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListMessages = Client.Execute(Request)
        If ListMessages.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListMessages))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListMessages) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.ListMessages", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function GetMailFolderID(UserPrincipal As String, sFolder As String) As String
    Dim Request As New WebRequest
    Dim Response As New WebResponse
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable") 'Authorization grant type
    If pGrantType = "authorization_code" Then
        Request.Resource = "/me"
    Else
        Request.Resource = "/users/" & UserPrincipal
    End If
    Request.Resource = Request.Resource & "/mailfolders"
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim lRetryCount As Long
    GetMailFolderID = "Retry"
    lRetryCount = 0
    While GetMailFolderID = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set Response = Client.Execute(Request)
        If Response.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(Response))
            GetMailFolderID = "Retry"
        ElseIf Response.StatusCode = WebStatusCode.OK Then
            Dim FolderInfo As Dictionary
            For Each FolderInfo In Response.Data("value")
                If FolderInfo("displayName") = sFolder Then
                    GetMailFolderID = FolderInfo("id")
                    Exit Function
                End If
            Next FolderInfo
            GetMailFolderID = ""
        ElseIf IsTokenExpiredError(Response) Then
            ClearAuthCodes
            GetMailFolderID = "Retry"
        Else
            MsgBox "Error " & Response.StatusCode & ": " & Response.Content
            GetMailFolderID = "Error"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.GetMailFolderID", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function SendDraft(sUserPrincipal As String, sMessageId As String) As WebResponse
    Dim Request As New WebRequest
    Request.Resource = "/users/{UserPrincipal}/messages/{MessageId}/send"
    Request.AddUrlSegment "UserPrincipal", sUserPrincipal
    Request.AddUrlSegment "MessageId", sMessageId
    Request.Method = WebMethod.HttpPost
    Request.AddHeader "client-request-id", CreateGUID()
    
    Set SendDraft = Client.Execute(Request)
End Function

Public Function DeleteMessage(sUserPrincipal As String, sMessageId As String) As WebResponse
    Dim Request As New WebRequest
    Request.Resource = "/users/{UserPrincipal}/messages/{MessageId}"
    Request.AddUrlSegment "UserPrincipal", sUserPrincipal
    Request.AddUrlSegment "MessageId", sMessageId
    Request.Method = WebMethod.HttpDelete
    Request.AddHeader "client-request-id", CreateGUID()
    
    Set DeleteMessage = Client.Execute(Request)
End Function

Public Function DeleteEvent(sUserPrincipal As String, sEventId As String) As WebResponse
    Dim Request As New WebRequest
    Request.Resource = "/users/{UserPrincipal}/events/{EventId}"
    Request.AddUrlSegment "UserPrincipal", sUserPrincipal
    Request.AddUrlSegment "EventId", sEventId
    Request.Method = WebMethod.HttpDelete
    Request.AddHeader "client-request-id", CreateGUID()
    
    Set DeleteEvent = Client.Execute(Request)
End Function

Public Function ListEvents(sUserPrincipal As String, _
    Optional sStartDateTime As String = "", _
    Optional sEndDateTime As String = "", _
    Optional sSelectFields As String = "", _
    Optional sCalendarName As String = "") As WebResponse
    
    Dim Request As New WebRequest
    
    If Len(sStartDateTime) > 0 And Len(sEndDateTime) > 0 Then
        ' Use calendarView for date-range queries
        If Len(sCalendarName) > 0 Then
            Dim sCalId As String
            sCalId = GetCalendarID(sUserPrincipal, sCalendarName, "")
            Request.Resource = "/users/{UserPrincipal}/calendars/{CalendarId}/calendarView"
            Request.AddUrlSegment "CalendarId", sCalId
        Else
            Request.Resource = "/users/{UserPrincipal}/calendarView"
        End If
        Request.AddQuerystringParam "startDateTime", sStartDateTime
        Request.AddQuerystringParam "endDateTime", sEndDateTime
    Else
        Request.Resource = "/users/{UserPrincipal}/events"
    End If
    
    Request.AddUrlSegment "UserPrincipal", sUserPrincipal
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set ListEvents = Client.Execute(Request)
End Function



