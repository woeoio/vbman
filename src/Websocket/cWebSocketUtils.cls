VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebSocketUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' mWebSocketUtils - WebSocket Utility Module for VB6
'
' Purpose: Provides common utilities for WebSocket implementation
'          - UTF-8 encoding/decoding
'          - Base64 encoding
'          - SHA1 hashing (via CryptoAPI)
'          - WebSocket key generation and validation
'
' Author: Claude (Anthropic)
' Date: 2026-01-10
'
'=========================================================================
Option Explicit

'=========================================================================
' Public Enums
'=========================================================================

' WebSocket Opcodes (RFC 6455)
Public Enum WsOpCode
    WS_OPCODE_CONTINUATION = 0
    WS_OPCODE_TEXT = 1
    WS_OPCODE_BINARY = 2
    WS_OPCODE_CLOSE = 8
    WS_OPCODE_PING = 9
    WS_OPCODE_PONG = 10
End Enum

' WebSocket Close Status Codes (RFC 6455)
Public Enum WsCloseCode
    WS_CLOSE_NORMAL = 1000
    WS_CLOSE_GOING_AWAY = 1001
    WS_CLOSE_PROTOCOL_ERROR = 1002
    WS_CLOSE_UNSUPPORTED_DATA = 1003
    WS_CLOSE_NO_STATUS = 1005
    WS_CLOSE_ABNORMAL = 1006
    WS_CLOSE_INVALID_DATA = 1007
    WS_CLOSE_POLICY_VIOLATION = 1008
    WS_CLOSE_MESSAGE_TOO_BIG = 1009
    WS_CLOSE_MANDATORY_EXT = 1010
    WS_CLOSE_INTERNAL_ERROR = 1011
    WS_CLOSE_SERVICE_RESTART = 1012
    WS_CLOSE_TRY_AGAIN = 1013
End Enum

' WebSocket Connection State
Public Enum WsState
    WS_STATE_CLOSED = 0
    WS_STATE_CONNECTING = 1
    WS_STATE_OPEN = 2
    WS_STATE_CLOSING = 3
End Enum

'=========================================================================
' API Declarations
'=========================================================================

' UTF-8 Conversion
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

' CryptoAPI for SHA1
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, ByRef pbData As Byte, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants
'=========================================================================

Private Const CP_UTF8 As Long = 65001
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA1 As Long = &H8004&
Private Const HP_HASHVAL As Long = 2
Private Const SHA1_HASH_SIZE As Long = 20

Const WS_MAGIC_GUID As String = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
Const WS_VERSION As String = "13"

'=========================================================================
' Module Variables
'=========================================================================

Private m_hCryptoProv As Long

'=========================================================================
' UTF-8 Functions
'=========================================================================

' Convert Unicode string to UTF-8 byte array
Public Function StringToUTF8(ByVal Text As String) As Byte()
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim baUTF8() As Byte

    lLength = Len(Text)
    If lLength = 0 Then
        StringToUTF8 = baUTF8
        Exit Function
    End If

    ' Calculate required buffer size
    lBufferSize = lLength * 3 + 1
    ReDim baUTF8(lBufferSize - 1)

    ' Convert
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), lLength, baUTF8(0), lBufferSize, vbNullString, 0)

    If lResult > 0 Then
        ReDim Preserve baUTF8(lResult - 1)
        StringToUTF8 = baUTF8
    End If
End Function

' Convert UTF-8 byte array to Unicode string
Public Function UTF8ToString(ByRef Utf8Data() As Byte) As String
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim sResult As String

    On Error Resume Next
    lLength = UBound(Utf8Data) + 1
    If ERR.Number <> 0 Or lLength <= 0 Then
        ERR.Clear
        UTF8ToString = ""
        Exit Function
    End If
    On Error GoTo 0

    ' Calculate required buffer size
    lBufferSize = lLength * 2
    sResult = String$(lBufferSize, Chr$(0))

    ' Convert
    lResult = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf8Data(0)), lLength, StrPtr(sResult), lBufferSize)

    If lResult > 0 Then
        UTF8ToString = Left$(sResult, lResult)
    Else
        UTF8ToString = ""
    End If
End Function

'=========================================================================
' Base64 Functions
'=========================================================================

' Base64 encode byte array
Public Function Base64Encode(ByRef Data() As Byte) As String
    Const B64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim lDataLen As Long
    Dim lGroups As Long
    Dim i As Long, j As Long
    Dim lTemp As Long
    Dim sResult As String

    On Error Resume Next
    lDataLen = UBound(Data) + 1
    If ERR.Number <> 0 Or lDataLen <= 0 Then
        ERR.Clear
        Base64Encode = ""
        Exit Function
    End If
    On Error GoTo 0

    ' Calculate output size
    lGroups = (lDataLen + 2) \ 3
    sResult = String$(lGroups * 4, "A")

    For i = 0 To lGroups - 1
        ' Combine 3 bytes into 24-bit value
        lTemp = 0
        For j = 0 To 2
            If i * 3 + j <= UBound(Data) Then
                lTemp = lTemp Or (CLng(Data(i * 3 + j)) * (256 ^ (2 - j)))
            End If
        Next j

        ' Extract 4 6-bit values
        For j = 0 To 3
            If i * 3 + (j * 3 \ 4) <= UBound(Data) Then
                Mid$(sResult, i * 4 + j + 1, 1) = Mid$(B64_CHARS, ((lTemp \ (64 ^ (3 - j))) And 63) + 1, 1)
            Else
                Mid$(sResult, i * 4 + j + 1, 1) = "="
            End If
        Next j
    Next i

    Base64Encode = sResult
End Function

'=========================================================================
' SHA1 Functions
'=========================================================================

' Compute SHA1 hash of byte array
Public Function SHA1Hash(ByRef Data() As Byte) As Byte()
    Dim hHash As Long
    Dim baHash(SHA1_HASH_SIZE - 1) As Byte
    Dim lHashLen As Long
    Dim lDataLen As Long

    On Error GoTo ErrorHandler

    ' Get data length
    On Error Resume Next
    lDataLen = UBound(Data) + 1
    If ERR.Number <> 0 Then
        ERR.Clear
        lDataLen = 0
    End If
    On Error GoTo ErrorHandler

    ' Initialize crypto provider if needed
    If m_hCryptoProv = 0 Then
        If CryptAcquireContext(m_hCryptoProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
            GoTo ErrorHandler
        End If
    End If

    ' Create hash object
    If CryptCreateHash(m_hCryptoProv, CALG_SHA1, 0, 0, hHash) = 0 Then
        GoTo ErrorHandler
    End If

    ' Hash data
    If lDataLen > 0 Then
        If CryptHashData(hHash, Data(0), lDataLen, 0) = 0 Then
            CryptDestroyHash hHash
            GoTo ErrorHandler
        End If
    End If

    ' Get hash value
    lHashLen = SHA1_HASH_SIZE
    If CryptGetHashParam(hHash, HP_HASHVAL, baHash(0), lHashLen, 0) = 0 Then
        CryptDestroyHash hHash
        GoTo ErrorHandler
    End If

    CryptDestroyHash hHash
    SHA1Hash = baHash
    Exit Function

ErrorHandler:
    If hHash <> 0 Then CryptDestroyHash hHash
    Dim baEmpty() As Byte
    SHA1Hash = baEmpty
End Function

'=========================================================================
' WebSocket Key Functions
'=========================================================================

' Generate random 16-byte key, Base64 encoded
Public Function GenerateWebSocketKey() As String
    Dim baKey(15) As Byte
    Dim i As Long

    Randomize Timer
    For i = 0 To 15
        baKey(i) = CByte(Int(Rnd * 256))
    Next i

    GenerateWebSocketKey = Base64Encode(baKey)
End Function

' Compute WebSocket accept key from client key
Public Function ComputeAcceptKey(ByVal ClientKey As String) As String
    Dim sCombined As String
    Dim baCombined() As Byte
    Dim baHash() As Byte

    sCombined = ClientKey & WS_MAGIC_GUID
    baCombined = StrConv(sCombined, vbFromUnicode)
    baHash = SHA1Hash(baCombined)

    On Error Resume Next
    If UBound(baHash) >= 0 Then
        ComputeAcceptKey = Base64Encode(baHash)
    Else
        ComputeAcceptKey = ""
    End If
End Function

'=========================================================================
' HTTP Header Parsing
'=========================================================================

' Get header value from HTTP request/response
Public Function GetHeaderValue(ByVal HttpText As String, ByVal HeaderName As String) As String
    Dim lPos As Long
    Dim lEnd As Long

    lPos = InStr(1, HttpText, HeaderName & ":", vbTextCompare)
    If lPos = 0 Then Exit Function

    lPos = lPos + Len(HeaderName) + 1

    ' Skip leading spaces
    Do While lPos <= Len(HttpText) And Mid$(HttpText, lPos, 1) = " "
        lPos = lPos + 1
    Loop

    lEnd = InStr(lPos, HttpText, vbCrLf)
    If lEnd = 0 Then lEnd = Len(HttpText) + 1

    GetHeaderValue = Trim$(Mid$(HttpText, lPos, lEnd - lPos))
End Function

'=========================================================================
' Close Code Description
'=========================================================================

' Get human-readable description for close code
Public Function GetCloseCodeDescription(ByVal Code As WsCloseCode) As String
    Select Case Code
    Case WS_CLOSE_NORMAL: GetCloseCodeDescription = "Normal closure"
    Case WS_CLOSE_GOING_AWAY: GetCloseCodeDescription = "Going away"
    Case WS_CLOSE_PROTOCOL_ERROR: GetCloseCodeDescription = "Protocol error"
    Case WS_CLOSE_UNSUPPORTED_DATA: GetCloseCodeDescription = "Unsupported data"
    Case WS_CLOSE_NO_STATUS: GetCloseCodeDescription = "No status received"
    Case WS_CLOSE_ABNORMAL: GetCloseCodeDescription = "Abnormal closure"
    Case WS_CLOSE_INVALID_DATA: GetCloseCodeDescription = "Invalid data"
    Case WS_CLOSE_POLICY_VIOLATION: GetCloseCodeDescription = "Policy violation"
    Case WS_CLOSE_MESSAGE_TOO_BIG: GetCloseCodeDescription = "Message too big"
    Case WS_CLOSE_MANDATORY_EXT: GetCloseCodeDescription = "Mandatory extension"
    Case WS_CLOSE_INTERNAL_ERROR: GetCloseCodeDescription = "Internal error"
    Case WS_CLOSE_SERVICE_RESTART: GetCloseCodeDescription = "Service restart"
    Case WS_CLOSE_TRY_AGAIN: GetCloseCodeDescription = "Try again later"
    Case Else: GetCloseCodeDescription = "Unknown (" & Code & ")"
    End Select
End Function

'=========================================================================
' Module Cleanup
'=========================================================================

Public Sub CleanupCryptoProvider()
    If m_hCryptoProv <> 0 Then
        CryptReleaseContext m_hCryptoProv, 0
        m_hCryptoProv = 0
    End If
End Sub

