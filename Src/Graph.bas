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
            Auth.AddScope "channelmessage.send"
            Auth.AddScope "chat.readbasic"
            Auth.AddScope "chat.create"
            Auth.AddScope "chatmessage.send"
            Auth.AddScope "channelmessage.read.all"
            Auth.AddScope "sites.read.all"
            Auth.AddScope "sites.readwrite.all"
            Auth.AddScope "files.readwrite"
            Auth.AddScope "user.read.all"
            Auth.AddScope "people.read"
            Auth.AddScope "groupmember.readwrite.all"
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


' =============================================================================
' Utility Helpers
' =============================================================================

Private Function EscapeJsonString(sInput As String) As String
    ' Escapes special characters for JSON embedding.
    ' CRITICAL: Escape order matters — backslash FIRST to avoid double-escaping.
    If Len(sInput) = 0 Then Exit Function
    
    Dim sOutput As String
    sOutput = sInput
    sOutput = Replace(sOutput, "\", "\\")       ' 1. Backslash first
    sOutput = Replace(sOutput, """", "\""")      ' 2. Double quote
    sOutput = Replace(sOutput, vbCrLf, "\r\n")   ' 3. CRLF pair before individual
    sOutput = Replace(sOutput, vbCr, "\r")        ' 4. Carriage return
    sOutput = Replace(sOutput, vbLf, "\n")        ' 5. Line feed
    sOutput = Replace(sOutput, vbTab, "\t")       ' 6. Tab
    EscapeJsonString = sOutput
End Function

Private Function BuildResourcePath(sUserPrincipal As String) As String
    ' Builds user-scoped resource path: /me (auth_code) or /users/{UPN} (client_creds)
    If pGrantType = "" Then pGrantType = DLookup("GrantType", "AdminTable")
    If pGrantType = "authorization_code" Then
        BuildResourcePath = "/me"
    Else
        BuildResourcePath = "/users/" & sUserPrincipal
    End If
End Function


' =============================================================================
' User Profile & Directory
' =============================================================================

Public Function GetCurrentUser(Optional sSelectFields As String = "") As WebResponse
    ' GET /me — Returns the signed-in user's profile
    ' Scope: User.Read
    Dim Request As New WebRequest
    Request.Resource = "/me"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set GetCurrentUser = Client.Execute(Request)
End Function

Public Function ListGroupMembership(sUserPrincipal As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/memberOf — List groups/roles the user belongs to
    ' Scope: GroupMember.Read.All
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/memberOf"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set ListGroupMembership = Client.Execute(Request)
End Function

Public Function GetManager(sUserPrincipal As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/manager — Returns the user's manager
    ' Scope: User.Read.All
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/manager"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set GetManager = Client.Execute(Request)
End Function

Public Function ListDirectReports(sUserPrincipal As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/directReports — Lists the user's direct reports
    ' Scope: User.Read.All
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/directReports"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set ListDirectReports = Client.Execute(Request)
End Function


' =============================================================================
' Tasks & Planner
' =============================================================================

Public Function CreateTask(sTitle As String, Optional sBodyContent As String = "", _
    Optional sDueDate As String = "", Optional sImportance As String = "normal") As WebResponse
    ' POST /me/todo/lists/Tasks/tasks — Creates a To-Do task
    ' Scope: Tasks.ReadWrite
    Dim Request As New WebRequest
    Request.Resource = "/me/todo/lists/Tasks/tasks"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "title", sTitle
    Request.AddBodyParameter "importance", sImportance
    
    If Len(sBodyContent) > 0 Then
        Dim dictBody As New Dictionary
        dictBody.Add "contentType", "text"
        dictBody.Add "content", sBodyContent
        Request.AddBodyParameter "body", dictBody
    End If
    
    If Len(sDueDate) > 0 Then
        Dim dictDue As New Dictionary
        dictDue.Add "dateTime", sDueDate
        dictDue.Add "timeZone", "UTC"
        Request.AddBodyParameter "dueDateTime", dictDue
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateTask = Client.Execute(Request)
        If CreateTask.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateTask))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateTask) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.CreateTask", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListPlannerTasks(Optional sSelectFields As String = "") As WebResponse
    ' GET /me/planner/tasks — Lists the user's Planner tasks
    ' Scope: Tasks.Read
    Dim Request As New WebRequest
    Request.Resource = "/me/planner/tasks"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set ListPlannerTasks = Client.Execute(Request)
End Function


' =============================================================================
' Teams & Online Meetings
' =============================================================================

Public Function ListJoinedTeams(Optional sSelectFields As String = "") As WebResponse
    ' GET /me/joinedTeams — Lists teams the user has joined
    ' Scope: Team.ReadBasic.All
    Dim Request As New WebRequest
    Request.Resource = "/me/joinedTeams"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set ListJoinedTeams = Client.Execute(Request)
End Function

Public Function ListTeamsChannels(sTeamId As String) As WebResponse
    ' GET /teams/{TeamId}/channels — Lists channels in a team
    ' Scope: Channel.ReadBasic.All
    Dim Request As New WebRequest
    Request.Resource = "/teams/{TeamId}/channels"
    Request.AddUrlSegment "TeamId", sTeamId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Set ListTeamsChannels = Client.Execute(Request)
End Function

Public Function CreateOnlineMeeting(sSubject As String, dtStart As Date, dtEnd As Date, _
    Optional sTimeZone As String = "UTC") As WebResponse
    ' POST /me/onlineMeetings — Creates a Teams online meeting
    ' Scope: OnlineMeetings.ReadWrite
    ' Returns: WebResponse with joinWebUrl in response data
    Dim Request As New WebRequest
    Request.Resource = "/me/onlineMeetings"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "subject", sSubject
    Request.AddBodyParameter "startDateTime", Format(dtStart, "yyyy-mm-dd\Thh:nn:ss")
    Request.AddBodyParameter "endDateTime", Format(dtEnd, "yyyy-mm-dd\Thh:nn:ss")
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateOnlineMeeting = Client.Execute(Request)
        If CreateOnlineMeeting.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateOnlineMeeting))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateOnlineMeeting) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.CreateOnlineMeeting", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' Teams Messaging
' =============================================================================

Public Function SendChannelMessage(sTeamId As String, sChannelId As String, sContent As String, _
    Optional sContentType As String = "text") As WebResponse
    ' POST /teams/{TeamId}/channels/{ChannelId}/messages — Send a message to a channel
    ' Scope: ChannelMessage.Send (delegated only)
    Dim Request As New WebRequest
    Request.Resource = "/teams/{TeamId}/channels/{ChannelId}/messages"
    Request.AddUrlSegment "TeamId", sTeamId
    Request.AddUrlSegment "ChannelId", sChannelId
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim dictBody As New Dictionary
    dictBody.Add "contentType", sContentType
    dictBody.Add "content", sContent
    
    Request.AddBodyParameter "body", dictBody
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set SendChannelMessage = Client.Execute(Request)
        If SendChannelMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(SendChannelMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(SendChannelMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.SendChannelMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ReplyToChannelMessage(sTeamId As String, sChannelId As String, sMessageId As String, _
    sContent As String, Optional sContentType As String = "text") As WebResponse
    ' POST /teams/{TeamId}/channels/{ChannelId}/messages/{MessageId}/replies — Reply to a channel message
    ' Scope: ChannelMessage.Send (delegated only)
    Dim Request As New WebRequest
    Request.Resource = "/teams/{TeamId}/channels/{ChannelId}/messages/{MessageId}/replies"
    Request.AddUrlSegment "TeamId", sTeamId
    Request.AddUrlSegment "ChannelId", sChannelId
    Request.AddUrlSegment "MessageId", sMessageId
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim dictBody As New Dictionary
    dictBody.Add "contentType", sContentType
    dictBody.Add "content", sContent
    
    Request.AddBodyParameter "body", dictBody
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ReplyToChannelMessage = Client.Execute(Request)
        If ReplyToChannelMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ReplyToChannelMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ReplyToChannelMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.ReplyToChannelMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function SendChatMessage(sChatId As String, sContent As String, _
    Optional sContentType As String = "text") As WebResponse
    ' POST /chats/{ChatId}/messages — Send a message to a chat (DM or group)
    ' Scope: ChatMessage.Send (delegated only)
    Dim Request As New WebRequest
    Request.Resource = "/chats/{ChatId}/messages"
    Request.AddUrlSegment "ChatId", sChatId
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim dictBody As New Dictionary
    dictBody.Add "contentType", sContentType
    dictBody.Add "content", sContent
    
    Request.AddBodyParameter "body", dictBody
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set SendChatMessage = Client.Execute(Request)
        If SendChatMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(SendChatMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(SendChatMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.SendChatMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListChats(Optional sSelectFields As String = "", _
    Optional lTop As Long = 0) As WebResponse
    ' GET /me/chats — Lists the user's chats (1:1, group, meeting)
    ' Scope: Chat.ReadBasic (delegated)
    Dim Request As New WebRequest
    Request.Resource = "/me/chats"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Set ListChats = Client.Execute(Request)
End Function

Public Function ListChannelMessages(sTeamId As String, sChannelId As String, _
    Optional lTop As Long = 0) As WebResponse
    ' GET /teams/{TeamId}/channels/{ChannelId}/messages — Lists messages in a channel
    ' Scope: ChannelMessage.Read.All (delegated)
    Dim Request As New WebRequest
    Request.Resource = "/teams/{TeamId}/channels/{ChannelId}/messages"
    Request.AddUrlSegment "TeamId", sTeamId
    Request.AddUrlSegment "ChannelId", sChannelId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Set ListChannelMessages = Client.Execute(Request)
End Function

Public Function CreateChat(sChatType As String, sUserIds As String, _
    Optional sTopic As String = "") As WebResponse
    ' POST /chats — Creates a new 1:1 or group chat
    ' Scope: Chat.Create (delegated)
    Dim Request As New WebRequest
    Request.Resource = "/chats"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim colMembers As New Collection
    Dim vUsers As Variant
    Dim i As Long
    Dim dictMember As Dictionary
    
    vUsers = Split(sUserIds, ";")
    For i = LBound(vUsers) To UBound(vUsers)
        If Len(Trim(vUsers(i))) > 0 Then
            Set dictMember = New Dictionary
            dictMember.Add "@odata.type", "#microsoft.graph.aadUserConversationMember"
            dictMember.Add "roles", Array("owner")
            dictMember.Add "user@odata.bind", "https://graph.microsoft.com/v1.0/users('" & Trim(vUsers(i)) & "')"
            colMembers.Add dictMember
        End If
    Next i
    
    Request.AddBodyParameter "chatType", sChatType
    If Len(sTopic) > 0 Then
        Request.AddBodyParameter "topic", sTopic
    End If
    Request.AddBodyParameter "members", colMembers
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateChat = Client.Execute(Request)
        If CreateChat.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateChat))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateChat) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.CreateChat", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' SharePoint & OneDrive
' =============================================================================

Public Function ListSharePointSites(sSearchQuery As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /sites?search={query} — Search for SharePoint sites
    ' Scope: Sites.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/sites"
    Request.AddQuerystringParam "search", sSearchQuery
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set ListSharePointSites = Client.Execute(Request)
End Function

Public Function SearchOneDrive(sUserPrincipal As String, sQuery As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/drive/root/search(q='{query}') — Search files in OneDrive
    ' Scope: Files.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/root/search(q='" & sQuery & "')"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set SearchOneDrive = Client.Execute(Request)
End Function

Public Function SearchSharePoint(sQuery As String, _
    Optional sEntityType As String = "driveItem", _
    Optional lMaxResults As Long = 25) As WebResponse
    ' POST /search/query — Search across SharePoint and OneDrive
    ' Scope: Files.Read.All or Sites.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/search/query"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim dictQuery As New Dictionary
    dictQuery.Add "queryString", sQuery
    
    Dim colEntityTypes As New Collection
    colEntityTypes.Add sEntityType
    
    Dim dictRequest As New Dictionary
    dictRequest.Add "entityTypes", colEntityTypes
    dictRequest.Add "query", dictQuery
    dictRequest.Add "from", 0
    dictRequest.Add "size", lMaxResults
    
    Dim colRequests As New Collection
    colRequests.Add dictRequest
    
    Request.AddBodyParameter "requests", colRequests
    
    Set SearchSharePoint = Client.Execute(Request)
End Function


' =============================================================================
' OneNote
' =============================================================================

Public Function ListOneNoteNotebooks(sUserPrincipal As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/onenote/notebooks — Lists OneNote notebooks
    ' Scope: Notes.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/onenote/notebooks"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set ListOneNoteNotebooks = Client.Execute(Request)
End Function


' =============================================================================
' Send As / On Behalf Of
' =============================================================================

Public Function SendMailAs(sUserPrincipal As String, sSendAs As String, _
    sSubject As String, sBodyType As String, sBodyContent As String, _
    sToRecipients As String, _
    Optional sCcRecipients As String = "", _
    Optional sBccRecipients As String = "", _
    Optional sAttachmentPath As String = "") As WebResponse
    ' POST /me/sendMail — Send email as/on behalf of another address
    ' Scope: Mail.Send
    ' sSendAs: email address to send from (empty = send as self)
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/sendMail"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim dictMessage As New Dictionary
    
    ' Add "from" field for Send As
    If Len(sSendAs) > 0 Then
        Dim dictFromAddr As New Dictionary
        dictFromAddr.Add "address", sSendAs
        Dim dictFrom As New Dictionary
        dictFrom.Add "emailAddress", dictFromAddr
        dictMessage.Add "from", dictFrom
    End If
    
    dictMessage.Add "subject", sSubject
    
    Dim dictBody As New Dictionary
    dictBody.Add "contentType", sBodyType
    dictBody.Add "content", sBodyContent
    dictMessage.Add "body", dictBody
    
    Dim colToRecipients As New Collection
    FillEmailAddressCollection colToRecipients, sToRecipients
    dictMessage.Add "toRecipients", colToRecipients
    
    If Len(Trim(sCcRecipients)) > 0 Then
        Dim colCc As New Collection
        FillEmailAddressCollection colCc, sCcRecipients
        dictMessage.Add "ccRecipients", colCc
    End If
    
    If Len(Trim(sBccRecipients)) > 0 Then
        Dim colBcc As New Collection
        FillEmailAddressCollection colBcc, sBccRecipients
        dictMessage.Add "bccRecipients", colBcc
    End If
    
    If Len(Trim(sAttachmentPath)) > 0 Then
        Dim colAttachments As New Collection
        FillAttachmentCollection colAttachments, sAttachmentPath
        dictMessage.Add "attachments", colAttachments
    End If
    
    Request.AddBodyParameter "message", dictMessage
    Request.AddBodyParameter "saveToSentItems", True
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set SendMailAs = Client.Execute(Request)
        If SendMailAs.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(SendMailAs))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(SendMailAs) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11050, "Graph.SendMailAs", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' PATCH Operations (Update)
' =============================================================================

Public Function UpdateEvent(sUserPrincipal As String, sEventId As String, _
    dictUpdates As Dictionary) As WebResponse
    ' PATCH /me/events/{EventId} — Update fields on a calendar event
    ' Scope: Calendars.ReadWrite
    ' dictUpdates: Dictionary of field names → new values
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/events/" & sEventId
    Request.Method = WebMethod.HttpPatch
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim vKey As Variant
    For Each vKey In dictUpdates.Keys
        Request.AddBodyParameter CStr(vKey), dictUpdates(vKey)
    Next vKey
    
    Set UpdateEvent = Client.Execute(Request)
End Function

Public Function UpdateContact(sUserPrincipal As String, sContactId As String, _
    dictUpdates As Dictionary) As WebResponse
    ' PATCH /me/contacts/{ContactId} — Update fields on a contact
    ' Scope: Contacts.ReadWrite
    ' dictUpdates: Dictionary of field names → new values
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/contacts/" & sContactId
    Request.Method = WebMethod.HttpPatch
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim vKey As Variant
    For Each vKey In dictUpdates.Keys
        Request.AddBodyParameter CStr(vKey), dictUpdates(vKey)
    Next vKey
    
    Set UpdateContact = Client.Execute(Request)
End Function


' =============================================================================
' SharePoint Sites & Lists
' =============================================================================

Public Function GetSite(sSiteId As String, Optional sSelectFields As String = "") As WebResponse
    ' GET /sites/{SiteId}
    ' Scope: Sites.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/sites/{SiteId}"
    Request.AddUrlSegment "SiteId", sSiteId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Set GetSite = Client.Execute(Request)
End Function

Public Function ListSiteLists(sSiteId As String, Optional sSelectFields As String = "", Optional lTop As Long = 0) As WebResponse
    ' GET /sites/{SiteId}/lists
    ' Scope: Sites.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/sites/{SiteId}/lists"
    Request.AddUrlSegment "SiteId", sSiteId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Set ListSiteLists = Client.Execute(Request)
End Function

Public Function ListSiteListItems(sSiteId As String, sListId As String, Optional sSelectFields As String = "", Optional sFilter As String = "", Optional lTop As Long = 0) As WebResponse
    ' GET /sites/{SiteId}/lists/{ListId}/items
    ' Scope: Sites.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/sites/{SiteId}/lists/{ListId}/items"
    Request.AddUrlSegment "SiteId", sSiteId
    Request.AddUrlSegment "ListId", sListId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$expand", "fields(select=" & sSelectFields & ")"
    Else
        Request.AddQuerystringParam "$expand", "fields"
    End If
    If Len(sFilter) > 0 Then
        Request.AddQuerystringParam "$filter", sFilter
    End If
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Set ListSiteListItems = Client.Execute(Request)
End Function

Public Function CreateListItem(sSiteId As String, sListId As String, dictFields As Dictionary) As WebResponse
    ' POST /sites/{SiteId}/lists/{ListId}/items
    ' Scope: Sites.ReadWrite.All
    Dim Request As New WebRequest
    Request.Resource = "/sites/{SiteId}/lists/{ListId}/items"
    Request.AddUrlSegment "SiteId", sSiteId
    Request.AddUrlSegment "ListId", sListId
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "fields", dictFields
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateListItem = Client.Execute(Request)
        If CreateListItem.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateListItem))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateListItem) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11060, "Graph.CreateListItem", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function UpdateListItem(sSiteId As String, sListId As String, sItemId As String, dictFields As Dictionary) As WebResponse
    ' PATCH /sites/{SiteId}/lists/{ListId}/items/{ItemId}/fields
    ' Scope: Sites.ReadWrite.All
    Dim Request As New WebRequest
    Request.Resource = "/sites/{SiteId}/lists/{ListId}/items/{ItemId}/fields"
    Request.AddUrlSegment "SiteId", sSiteId
    Request.AddUrlSegment "ListId", sListId
    Request.AddUrlSegment "ItemId", sItemId
    Request.Method = WebMethod.HttpPatch
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim vKey As Variant
    For Each vKey In dictFields.Keys
        Request.AddBodyParameter CStr(vKey), dictFields(vKey)
    Next vKey
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set UpdateListItem = Client.Execute(Request)
        If UpdateListItem.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(UpdateListItem))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(UpdateListItem) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11070, "Graph.UpdateListItem", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function DeleteListItem(sSiteId As String, sListId As String, sItemId As String) As WebResponse
    ' DELETE /sites/{SiteId}/lists/{ListId}/items/{ItemId}
    ' Scope: Sites.ReadWrite.All
    Dim Request As New WebRequest
    Request.Resource = "/sites/{SiteId}/lists/{ListId}/items/{ItemId}"
    Request.AddUrlSegment "SiteId", sSiteId
    Request.AddUrlSegment "ListId", sListId
    Request.AddUrlSegment "ItemId", sItemId
    Request.Method = WebMethod.HttpDelete
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set DeleteListItem = Client.Execute(Request)
        If DeleteListItem.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(DeleteListItem))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(DeleteListItem) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11080, "Graph.DeleteListItem", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' OneDrive Files & Folders
' =============================================================================

Public Function ListDriveChildren(sUserPrincipal As String, Optional sItemId As String = "", _
    Optional sSelectFields As String = "", Optional lTop As Long = 0, _
    Optional sOrderBy As String = "") As WebResponse
    ' GET /me/drive/root/children  OR  /me/drive/items/{ItemId}/children
    ' Scope: Files.Read
    Dim Request As New WebRequest

    If Len(sItemId) > 0 Then
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/items/" & sItemId & "/children"
    Else
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/root/children"
    End If

    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()

    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    If Len(sOrderBy) > 0 Then
        Request.AddQuerystringParam "$orderby", sOrderBy
    End If

    Set ListDriveChildren = Client.Execute(Request)
End Function

Public Function DownloadDriveItem(sUserPrincipal As String, sItemId As String) As WebResponse
    ' GET /me/drive/items/{ItemId}?$select=id,name,size,@microsoft.graph.downloadUrl
    ' Scope: Files.Read
    Dim Request As New WebRequest

    Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/items/{ItemId}"
    Request.AddUrlSegment "ItemId", sItemId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    Request.AddQuerystringParam "$select", "id,name,size,@microsoft.graph.downloadUrl"

    Set DownloadDriveItem = Client.Execute(Request)
End Function

Public Function UploadSmallFile(sUserPrincipal As String, sParentPath As String, _
    sFileName As String, sContent As String) As WebResponse
    ' PUT /me/drive/root:/{ParentPath}/{FileName}:/content
    ' Scope: Files.ReadWrite
    Dim Request As New WebRequest

    If Len(sParentPath) > 0 Then
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/root:/" & sParentPath & "/" & sFileName & ":/content"
    Else
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/root:/" & sFileName & ":/content"
    End If

    Request.Method = WebMethod.HttpPut
    Request.Format = WebFormat.JSON
    Request.ContentType = "text/plain"
    Request.Body = sContent
    Request.AddHeader "client-request-id", CreateGUID()

    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set UploadSmallFile = Client.Execute(Request)
        If UploadSmallFile.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(UploadSmallFile))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(UploadSmallFile) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11090, "Graph.UploadSmallFile", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function CreateDriveFolder(sUserPrincipal As String, sFolderName As String, _
    Optional sParentItemId As String = "", _
    Optional sConflictBehavior As String = "rename") As WebResponse
    ' POST /me/drive/root/children  OR  /me/drive/items/{ParentItemId}/children
    ' Scope: Files.ReadWrite
    Dim Request As New WebRequest

    If Len(sParentItemId) > 0 Then
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/items/{ParentItemId}/children"
        Request.AddUrlSegment "ParentItemId", sParentItemId
    Else
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/root/children"
    End If

    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    Request.AddBodyParameter "name", sFolderName

    Dim dictFolder As New Dictionary
    Request.AddBodyParameter "folder", dictFolder
    Request.AddBodyParameter "@microsoft.graph.conflictBehavior", sConflictBehavior

    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateDriveFolder = Client.Execute(Request)
        If CreateDriveFolder.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateDriveFolder))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateDriveFolder) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11100, "Graph.CreateDriveFolder", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function CreateSharingLink(sUserPrincipal As String, sItemId As String, _
    sLinkType As String, Optional sScope As String = "organization") As WebResponse
    ' POST /me/drive/items/{ItemId}/createLink
    ' Scope: Files.ReadWrite
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/items/{ItemId}/createLink"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    Request.AddUrlSegment "ItemId", sItemId
    Request.AddBodyParameter "type", sLinkType
    Request.AddBodyParameter "scope", sScope

    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateSharingLink = Client.Execute(Request)
        If CreateSharingLink.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateSharingLink))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateSharingLink) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11110, "Graph.CreateSharingLink", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function CopyDriveItem(sUserPrincipal As String, sItemId As String, _
    sDestDriveId As String, sDestFolderId As String, _
    Optional sNewName As String = "", _
    Optional sConflictBehavior As String = "rename") As WebResponse
    ' POST /me/drive/items/{ItemId}/copy — Returns 202 Accepted with monitor URL
    ' Scope: Files.ReadWrite
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/drive/items/{ItemId}/copy"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    Request.AddUrlSegment "ItemId", sItemId

    Dim dictParent As New Dictionary
    dictParent.Add "driveId", sDestDriveId
    dictParent.Add "id", sDestFolderId
    Request.AddBodyParameter "parentReference", dictParent

    If Len(sNewName) > 0 Then Request.AddBodyParameter "name", sNewName
    If Len(sConflictBehavior) > 0 Then Request.AddQuerystringParam "@microsoft.graph.conflictBehavior", sConflictBehavior

    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CopyDriveItem = Client.Execute(Request)
        If CopyDriveItem.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CopyDriveItem))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CopyDriveItem) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11120, "Graph.CopyDriveItem", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' Calendar — Read & RSVP
' =============================================================================

Public Function GetEvent(sUserPrincipal As String, sEventId As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/events/{id} — Get a single calendar event by ID
    ' Scope: Calendars.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/events/" & sEventId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set GetEvent = Client.Execute(Request)
        If GetEvent.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(GetEvent))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(GetEvent) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11200, "Graph.GetEvent", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function AcceptEvent(sUserPrincipal As String, sEventId As String, _
    Optional sComment As String = "", Optional bSendResponse As Boolean = True) As WebResponse
    ' POST /me/events/{id}/accept — Accept a calendar event invitation
    ' Scope: Calendars.ReadWrite
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/events/" & sEventId & "/accept"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "comment", sComment
    Request.AddBodyParameter "sendResponse", bSendResponse
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set AcceptEvent = Client.Execute(Request)
        If AcceptEvent.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(AcceptEvent))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(AcceptEvent) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11210, "Graph.AcceptEvent", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function DeclineEvent(sUserPrincipal As String, sEventId As String, _
    Optional sComment As String = "", Optional bSendResponse As Boolean = True) As WebResponse
    ' POST /me/events/{id}/decline — Decline a calendar event invitation
    ' Scope: Calendars.ReadWrite
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/events/" & sEventId & "/decline"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "comment", sComment
    Request.AddBodyParameter "sendResponse", bSendResponse
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set DeclineEvent = Client.Execute(Request)
        If DeclineEvent.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(DeclineEvent))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(DeclineEvent) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11220, "Graph.DeclineEvent", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function TentativelyAcceptEvent(sUserPrincipal As String, sEventId As String, _
    Optional sComment As String = "", Optional bSendResponse As Boolean = True) As WebResponse
    ' POST /me/events/{id}/tentativelyAccept — Tentatively accept a calendar event
    ' Scope: Calendars.ReadWrite
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/events/" & sEventId & "/tentativelyAccept"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "comment", sComment
    Request.AddBodyParameter "sendResponse", bSendResponse
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set TentativelyAcceptEvent = Client.Execute(Request)
        If TentativelyAcceptEvent.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(TentativelyAcceptEvent))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(TentativelyAcceptEvent) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11230, "Graph.TentativelyAcceptEvent", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ForwardEvent(sUserPrincipal As String, sEventId As String, _
    sToRecipients As String, Optional sComment As String = "") As WebResponse
    ' POST /me/events/{id}/forward — Forward a calendar event to recipients
    ' Scope: Calendars.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/events/" & sEventId & "/forward"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim colRecipients As New Collection
    FillEmailAddressCollection colRecipients, sToRecipients
    
    Request.AddBodyParameter "ToRecipients", colRecipients
    Request.AddBodyParameter "Comment", sComment
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ForwardEvent = Client.Execute(Request)
        If ForwardEvent.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ForwardEvent))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ForwardEvent) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11240, "Graph.ForwardEvent", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function GetFreeBusySchedule(sUserPrincipal As String, sSchedules As String, _
    sStartDateTime As String, sEndDateTime As String, sTimeZone As String, _
    Optional lIntervalMinutes As Long = 30) As WebResponse
    ' POST /me/calendar/getSchedule — Get free/busy availability for users
    ' Scope: Calendars.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/calendar/getSchedule"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    ' Parse semicolon-delimited schedules into a collection of plain strings
    Dim colSchedules As New Collection
    Dim sTemp As String
    sTemp = Replace(sSchedules, " ", "")
    If Right(sTemp, 1) = ";" Then sTemp = Left(sTemp, Len(sTemp) - 1)
    While InStr(sTemp, ";") > 0
        colSchedules.Add Left(sTemp, InStr(sTemp, ";") - 1)
        sTemp = Mid(sTemp, InStr(sTemp, ";") + 1)
    Wend
    If Len(sTemp) > 0 Then colSchedules.Add sTemp
    
    Dim dictStartTime As New Dictionary
    dictStartTime.Add "dateTime", sStartDateTime
    dictStartTime.Add "timeZone", sTimeZone
    
    Dim dictEndTime As New Dictionary
    dictEndTime.Add "dateTime", sEndDateTime
    dictEndTime.Add "timeZone", sTimeZone
    
    Request.AddBodyParameter "schedules", colSchedules
    Request.AddBodyParameter "startTime", dictStartTime
    Request.AddBodyParameter "endTime", dictEndTime
    Request.AddBodyParameter "availabilityViewInterval", lIntervalMinutes
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set GetFreeBusySchedule = Client.Execute(Request)
        If GetFreeBusySchedule.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(GetFreeBusySchedule))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(GetFreeBusySchedule) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11250, "Graph.GetFreeBusySchedule", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' Mail — Read & Organize
' =============================================================================

Public Function GetMessage(sUserPrincipal As String, sMessageId As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/messages/{id} — Retrieve a single message by ID
    ' Scope: Mail.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/messages/" & sMessageId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set GetMessage = Client.Execute(Request)
        If GetMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(GetMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(GetMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11130, "Graph.GetMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function UpdateMessage(sUserPrincipal As String, sMessageId As String, _
    dictUpdates As Dictionary) As WebResponse
    ' PATCH /me/messages/{id} — Update fields on a message
    ' Scope: Mail.ReadWrite
    ' dictUpdates: Dictionary of field names -> new values
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/messages/" & sMessageId
    Request.Method = WebMethod.HttpPatch
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim vKey As Variant
    For Each vKey In dictUpdates.Keys
        Request.AddBodyParameter CStr(vKey), dictUpdates(vKey)
    Next vKey
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set UpdateMessage = Client.Execute(Request)
        If UpdateMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(UpdateMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(UpdateMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11140, "Graph.UpdateMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function MoveMessage(sUserPrincipal As String, sMessageId As String, _
    sDestinationId As String) As WebResponse
    ' POST /me/messages/{id}/move — Move a message to a folder
    ' Scope: Mail.ReadWrite
    ' sDestinationId: well-known folder name (e.g. "deleteditems") or folder ID
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/messages/" & sMessageId & "/move"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "destinationId", sDestinationId
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set MoveMessage = Client.Execute(Request)
        If MoveMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(MoveMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(MoveMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11150, "Graph.MoveMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListMailFolders(sUserPrincipal As String, _
    Optional sSelectFields As String = "", _
    Optional lTop As Long = 0) As WebResponse
    ' GET /me/mailFolders — List mail folders
    ' Scope: Mail.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/mailFolders"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListMailFolders = Client.Execute(Request)
        If ListMailFolders.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListMailFolders))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListMailFolders) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11160, "Graph.ListMailFolders", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' Mail — Reply & Forward
' =============================================================================

Public Function ReplyToMessage(sUserPrincipal As String, sMessageId As String, _
    sComment As String, Optional sAddRecipients As String = "") As WebResponse
    ' POST /me/messages/{id}/reply — Reply to a message with optional additional recipients
    ' Scope: Mail.Send
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/messages/" & sMessageId & "/reply"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "comment", sComment
    
    If Len(Trim(sAddRecipients)) > 0 Then
        Dim colToRecipients As New Collection
        FillEmailAddressCollection colToRecipients, sAddRecipients
        Dim dictMessage As New Dictionary
        dictMessage.Add "toRecipients", colToRecipients
        Request.AddBodyParameter "message", dictMessage
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ReplyToMessage = Client.Execute(Request)
        If ReplyToMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ReplyToMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ReplyToMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11170, "Graph.ReplyToMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ReplyAllToMessage(sUserPrincipal As String, sMessageId As String, _
    sComment As String) As WebResponse
    ' POST /me/messages/{id}/replyAll — Reply-all to a message
    ' Scope: Mail.Send
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/messages/" & sMessageId & "/replyAll"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "comment", sComment
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ReplyAllToMessage = Client.Execute(Request)
        If ReplyAllToMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ReplyAllToMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ReplyAllToMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11180, "Graph.ReplyAllToMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ForwardMessage(sUserPrincipal As String, sMessageId As String, _
    sToRecipients As String, Optional sComment As String = "") As WebResponse
    ' POST /me/messages/{id}/forward — Forward a message to recipients
    ' Scope: Mail.Send
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/messages/" & sMessageId & "/forward"
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim colToRecipients As New Collection
    FillEmailAddressCollection colToRecipients, sToRecipients
    Request.AddBodyParameter "toRecipients", colToRecipients
    Request.AddBodyParameter "comment", sComment
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ForwardMessage = Client.Execute(Request)
        If ForwardMessage.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ForwardMessage))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ForwardMessage) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11190, "Graph.ForwardMessage", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' Groups — Read & Manage Members
' =============================================================================

Public Function ListGroups(Optional sFilter As String = "", _
    Optional sSearch As String = "", _
    Optional sSelectFields As String = "", _
    Optional lTop As Long = 100) As WebResponse
    ' GET /groups — List all groups in the tenant
    ' Scope: Group.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/groups"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sFilter) > 0 Or Len(sSearch) > 0 Then
        Request.AddHeader "ConsistencyLevel", "eventual"
    End If
    
    If Len(sFilter) > 0 Then
        Request.AddQuerystringParam "$filter", sFilter
    End If
    If Len(sSearch) > 0 Then
        Request.AddQuerystringParam "$search", Chr$(34) & sSearch & Chr$(34)
        Request.AddQuerystringParam "$count", "true"
    End If
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListGroups = Client.Execute(Request)
        If ListGroups.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListGroups))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListGroups) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11300, "Graph.ListGroups", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function GetGroup(sGroupId As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /groups/{GroupId} — Retrieve a single group by ID
    ' Scope: Group.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/groups/{GroupId}"
    Request.AddUrlSegment "GroupId", sGroupId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set GetGroup = Client.Execute(Request)
        If GetGroup.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(GetGroup))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(GetGroup) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11310, "Graph.GetGroup", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListGroupMembers(sGroupId As String, _
    Optional sSelectFields As String = "", _
    Optional lTop As Long = 100) As WebResponse
    ' GET /groups/{GroupId}/members — List members of a group
    ' Scope: GroupMember.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/groups/{GroupId}/members"
    Request.AddUrlSegment "GroupId", sGroupId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListGroupMembers = Client.Execute(Request)
        If ListGroupMembers.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListGroupMembers))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListGroupMembers) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11320, "Graph.ListGroupMembers", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListGroupOwners(sGroupId As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /groups/{GroupId}/owners — List owners of a group
    ' Scope: GroupMember.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/groups/{GroupId}/owners"
    Request.AddUrlSegment "GroupId", sGroupId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListGroupOwners = Client.Execute(Request)
        If ListGroupOwners.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListGroupOwners))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListGroupOwners) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11330, "Graph.ListGroupOwners", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function AddGroupMember(sGroupId As String, sUserId As String) As WebResponse
    ' POST /groups/{GroupId}/members/$ref — Add a member to a group
    ' Scope: GroupMember.ReadWrite.All
    Dim Request As New WebRequest
    Request.Resource = "/groups/{GroupId}/members/$ref"
    Request.AddUrlSegment "GroupId", sGroupId
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Request.AddBodyParameter "@odata.id", "https://graph.microsoft.com/v1.0/directoryObjects/" & sUserId
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set AddGroupMember = Client.Execute(Request)
        If AddGroupMember.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(AddGroupMember))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(AddGroupMember) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11340, "Graph.AddGroupMember", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function RemoveGroupMember(sGroupId As String, sMemberId As String) As WebResponse
    ' DELETE /groups/{GroupId}/members/{MemberId}/$ref — Remove a member from a group
    ' Scope: GroupMember.ReadWrite.All
    Dim Request As New WebRequest
    Request.Resource = "/groups/{GroupId}/members/{MemberId}/$ref"
    Request.AddUrlSegment "GroupId", sGroupId
    Request.AddUrlSegment "MemberId", sMemberId
    Request.Method = WebMethod.HttpDelete
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set RemoveGroupMember = Client.Execute(Request)
        If RemoveGroupMember.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(RemoveGroupMember))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(RemoveGroupMember) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11350, "Graph.RemoveGroupMember", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' Contacts — Read & Manage
' =============================================================================

Public Function GetContact(sUserPrincipal As String, sContactId As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/contacts/{id} — Retrieve a single contact by ID
    ' Scope: Contacts.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/contacts/" & sContactId
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set GetContact = Client.Execute(Request)
        If GetContact.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(GetContact))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(GetContact) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11260, "Graph.GetContact", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function DeleteContact(sUserPrincipal As String, sContactId As String) As WebResponse
    ' DELETE /me/contacts/{id} — Delete a contact
    ' Scope: Contacts.ReadWrite
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/contacts/" & sContactId
    Request.Method = WebMethod.HttpDelete
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set DeleteContact = Client.Execute(Request)
        If DeleteContact.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(DeleteContact))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(DeleteContact) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11270, "Graph.DeleteContact", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListContactFolders(sUserPrincipal As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /me/contactFolders — List contact folders
    ' Scope: Contacts.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/contactFolders"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListContactFolders = Client.Execute(Request)
        If ListContactFolders.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListContactFolders))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListContactFolders) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11280, "Graph.ListContactFolders", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function CreateContactFolder(sUserPrincipal As String, sDisplayName As String, _
    Optional sParentFolderId As String = "") As WebResponse
    ' POST /me/contactFolders — Create a contact folder (or child folder)
    ' Scope: Contacts.ReadWrite
    Dim Request As New WebRequest
    If Len(sParentFolderId) > 0 Then
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/contactFolders/" & sParentFolderId & "/childFolders"
    Else
        Request.Resource = BuildResourcePath(sUserPrincipal) & "/contactFolders"
    End If
    Request.Method = WebMethod.HttpPOST
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    Request.AddBodyParameter "displayName", sDisplayName
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set CreateContactFolder = Client.Execute(Request)
        If CreateContactFolder.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(CreateContactFolder))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(CreateContactFolder) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11290, "Graph.CreateContactFolder", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function


' =============================================================================
' User Profile & Directory — Lookup
' =============================================================================

Public Function GetUser(sUserIdOrUpn As String, _
    Optional sSelectFields As String = "") As WebResponse
    ' GET /users/{sUserIdOrUpn} — Retrieve a single user profile
    ' Scope: User.Read.All
    Dim Request As New WebRequest
    Request.Resource = "/users/" & sUserIdOrUpn
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set GetUser = Client.Execute(Request)
        If GetUser.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(GetUser))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(GetUser) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11360, "Graph.GetUser", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListUsers(Optional sFilter As String = "", _
    Optional sSearch As String = "", _
    Optional sSelectFields As String = "", _
    Optional lTop As Long = 100) As WebResponse
    ' GET /users — List users in the directory
    ' Scope: User.ReadBasic.All
    Dim Request As New WebRequest
    Request.Resource = "/users"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sFilter) > 0 Or Len(sSearch) > 0 Then
        Request.AddHeader "ConsistencyLevel", "eventual"
    End If
    
    If Len(sFilter) > 0 Then
        Request.AddQuerystringParam "$filter", sFilter
    End If
    
    If Len(sSearch) > 0 Then
        Request.AddQuerystringParam "$search", Chr$(34) & sSearch & Chr$(34)
        Request.AddQuerystringParam "$count", "true"
    End If
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListUsers = Client.Execute(Request)
        If ListUsers.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListUsers))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListUsers) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11370, "Graph.ListUsers", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function

Public Function ListPeople(sUserPrincipal As String, _
    Optional sSelectFields As String = "", _
    Optional lTop As Long = 0) As WebResponse
    ' GET {BuildResourcePath}/people — List relevant people for a user
    ' Scope: People.Read
    Dim Request As New WebRequest
    Request.Resource = BuildResourcePath(sUserPrincipal) & "/people"
    Request.Method = WebMethod.HttpGet
    Request.Format = WebFormat.JSON
    Request.AddHeader "client-request-id", CreateGUID()
    
    If Len(sSelectFields) > 0 Then
        Request.AddQuerystringParam "$select", sSelectFields
    End If
    
    If lTop > 0 Then
        Request.AddQuerystringParam "$top", CStr(lTop)
    End If
    
    Dim sStatus As String
    Dim lRetryCount As Long
    sStatus = "Retry"
    lRetryCount = 0
    While sStatus = "Retry" And lRetryCount < MAX_RETRIES
        lRetryCount = lRetryCount + 1
        Set ListPeople = Client.Execute(Request)
        If ListPeople.StatusCode = 429 Then
            Application.Wait Now + TimeSerial(0, 0, GetRetryAfterSeconds(ListPeople))
            sStatus = "Retry"
        ElseIf IsTokenExpiredError(ListPeople) Then
            ClearAuthCodes
            sStatus = "Retry"
        Else
            sStatus = "Done"
        End If
    Wend
    If lRetryCount >= MAX_RETRIES Then
        Err.Raise vbObjectError + 11380, "Graph.ListPeople", "Max retries exceeded after " & MAX_RETRIES & " attempts"
    End If
End Function
