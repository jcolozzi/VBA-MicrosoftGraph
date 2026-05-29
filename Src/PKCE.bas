Attribute VB_Name = "PKCE"
' =============================================================================
' Module: PKCE
' Purpose: Proof Key for Code Exchange (RFC 7636) support for OAuth 2.0
'          Enables authorization_code flow WITHOUT a client_secret
' =============================================================================
Option Compare Database
Option Explicit

' ===== Windows CryptoAPI Declarations =====
#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" _
        Alias "CryptAcquireContextW" ( _
        ByRef phProv As LongPtr, _
        ByVal szContainer As LongPtr, _
        ByVal szProvider As LongPtr, _
        ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, _
        ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, _
        ByVal algId As Long, _
        ByVal hKey As LongPtr, _
        ByVal dwFlags As Long, _
        ByRef phHash As LongPtr) As Long
    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, _
        ByVal pbData As LongPtr, _
        ByVal dwDataLen As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, _
        ByVal dwParam As Long, _
        ByVal pbData As LongPtr, _
        ByRef pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr) As Long
#Else
    Private Declare Function CryptAcquireContext Lib "advapi32.dll" _
        Alias "CryptAcquireContextW" ( _
        ByRef phProv As Long, _
        ByVal szContainer As Long, _
        ByVal szProvider As Long, _
        ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal algId As Long, _
        ByVal hKey As Long, _
        ByVal dwFlags As Long, _
        ByRef phHash As Long) As Long
    Private Declare Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As Long, _
        ByVal pbData As Long, _
        ByVal dwDataLen As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As Long, _
        ByVal dwParam As Long, _
        ByVal pbData As Long, _
        ByRef pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As Long) As Long
#End If

' CryptoAPI Constants
Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C&
Private Const HP_HASHVAL As Long = 2

' Module-level PKCE state
Private pCodeVerifier As String
Private pCodeChallenge As String
Private pUsePKCE As Boolean


' =============================================================================
' Public API
' =============================================================================

Public Sub EnablePKCE()
    ' Enable PKCE for the next authorization request.
    ' Call GeneratePKCEPair before redirecting to the authorization endpoint.
    pUsePKCE = True
End Sub

Public Sub DisablePKCE()
    ' Disable PKCE and clear stored verifier/challenge
    pUsePKCE = False
    pCodeVerifier = ""
    pCodeChallenge = ""
End Sub

Public Function IsPKCEEnabled() As Boolean
    IsPKCEEnabled = pUsePKCE
End Function

Public Sub GeneratePKCEPair()
    ' Generate a new code_verifier and compute its code_challenge.
    ' Call this before redirecting to the authorization endpoint.
    pCodeVerifier = GenerateCodeVerifier()
    pCodeChallenge = ComputeCodeChallenge(pCodeVerifier)
End Sub

Public Function GetCodeVerifier() As String
    ' Returns the current code_verifier (needed for the token exchange request)
    GetCodeVerifier = pCodeVerifier
End Function

Public Function GetCodeChallenge() As String
    ' Returns the current code_challenge (sent with the authorization URL)
    GetCodeChallenge = pCodeChallenge
End Function


' =============================================================================
' Private Implementation
' =============================================================================

Private Function GenerateCodeVerifier() As String
    ' RFC 7636 Section 4.1: Generate a 128-character random code_verifier
    ' using unreserved characters: [A-Z][a-z][0-9]-._~
    Const VERIFIER_LENGTH As Long = 128
    Const UNRESERVED_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~"
    
    Dim sVerifier As String
    Dim i As Long
    
    Randomize
    For i = 1 To VERIFIER_LENGTH
        sVerifier = sVerifier & Mid$(UNRESERVED_CHARS, Int(Rnd * Len(UNRESERVED_CHARS)) + 1, 1)
    Next i
    
    GenerateCodeVerifier = sVerifier
End Function

Private Function ComputeCodeChallenge(sVerifier As String) As String
    ' RFC 7636 Section 4.2: code_challenge = BASE64URL(SHA256(code_verifier))
    
    ' Step 1: SHA-256 hash
    Dim abHash() As Byte
    abHash = ComputeSHA256(sVerifier)
    
    ' Step 2: Base64 encode
    Dim sBase64 As String
    sBase64 = BytesToBase64(abHash)
    
    ' Step 3: Convert to Base64URL (RFC 4648 Section 5)
    sBase64 = Replace(sBase64, "+", "-")
    sBase64 = Replace(sBase64, "/", "_")
    Do While Right$(sBase64, 1) = "="
        sBase64 = Left$(sBase64, Len(sBase64) - 1)
    Loop
    
    ComputeCodeChallenge = sBase64
End Function

Private Function ComputeSHA256(sInput As String) As Byte()
    ' Compute SHA-256 hash of a string (as ANSI/UTF-8 bytes) via Windows CryptoAPI
#If VBA7 Then
    Dim hProv As LongPtr
    Dim hHash As LongPtr
#Else
    Dim hProv As Long
    Dim hHash As Long
#End If
    Dim lResult As Long
    Dim abData() As Byte
    Dim abHashValue(0 To 31) As Byte
    Dim lHashLen As Long
    
    abData = StrConv(sInput, vbFromUnicode)
    
    lResult = CryptAcquireContext(hProv, 0&, 0&, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)
    If lResult = 0 Then
        Err.Raise vbObjectError + 11060, "PKCE.ComputeSHA256", "CryptAcquireContext failed"
    End If
    
    lResult = CryptCreateHash(hProv, CALG_SHA_256, 0&, 0&, hHash)
    If lResult = 0 Then
        CryptReleaseContext hProv, 0
        Err.Raise vbObjectError + 11061, "PKCE.ComputeSHA256", "CryptCreateHash failed"
    End If
    
    lResult = CryptHashData(hHash, VarPtr(abData(0)), UBound(abData) + 1, 0)
    If lResult = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Err.Raise vbObjectError + 11062, "PKCE.ComputeSHA256", "CryptHashData failed"
    End If
    
    lHashLen = 32
    lResult = CryptGetHashParam(hHash, HP_HASHVAL, VarPtr(abHashValue(0)), lHashLen, 0)
    If lResult = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        Err.Raise vbObjectError + 11063, "PKCE.ComputeSHA256", "CryptGetHashParam failed"
    End If
    
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    
    ComputeSHA256 = abHashValue
End Function

Private Function BytesToBase64(abData() As Byte) As String
    ' Encode a byte array to Base64 using MSXML2.DOMDocument
    Dim oXML As Object
    Dim oNode As Object
    
    Set oXML = CreateObject("MSXML2.DOMDocument.6.0")
    Set oNode = oXML.createElement("b64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = abData
    
    BytesToBase64 = Replace(Replace(oNode.Text, vbCr, ""), vbLf, "")
    
    Set oNode = Nothing
    Set oXML = Nothing
End Function
