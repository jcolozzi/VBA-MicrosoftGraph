Attribute VB_Name = "AttachmentHelpers"
Option Compare Database
Option Explicit

Private Const MAX_INLINE_ATTACHMENT_SIZE As Long = 3145728  ' 3 MB = 3 * 1024 * 1024

Public Function ConvertFileToBase64(ByVal sPath As String) As String
    Dim objStream As Object
    Dim bytes As Variant
    
    On Error GoTo ErrHandler
    
    ' Validate file exists
    If Len(Dir(sPath)) = 0 Then
        Err.Raise vbObjectError + 11060, "ConvertFileToBase64", "File not found: " & sPath
    End If
    
    ' Validate file size (Graph API limit for inline attachments)
    If FileLen(sPath) > MAX_INLINE_ATTACHMENT_SIZE Then
        Err.Raise vbObjectError + 11061, "ConvertFileToBase64", _
            "File exceeds 3 MB limit for inline attachments. Use upload session for large files."
    End If
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1  'ADODB.adTypeBinary
    objStream.LoadFromFile sPath
    bytes = objStream.Read
    
    ' Close and cleanup
    If objStream.State <> 0 Then objStream.Close
    Set objStream = Nothing
    
    ConvertFileToBase64 = EncodeBase64(bytes)
    Exit Function

ErrHandler:
    If Not objStream Is Nothing Then
        If objStream.State <> 0 Then objStream.Close
        Set objStream = Nothing
    End If
    Err.Raise Err.Number, "ConvertFileToBase64", "Failed to convert file to Base64: " & Err.Description
End Function

Private Function EncodeBase64(ByRef bytes As Variant) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement

    On Error GoTo ErrHandler

    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = bytes
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
    Exit Function

ErrHandler:
    Set objNode = Nothing
    Set objXML = Nothing
    Err.Raise Err.Number, "EncodeBase64", "Failed to encode Base64: " & Err.Description
End Function

