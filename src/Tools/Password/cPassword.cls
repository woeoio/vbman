VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'=========================================================================
'
' cPassword - Password Hash Utility Class
'
' Purpose: Provides password creation and verification using hash algorithms
'          - Default algorithm: MD5
'          - Supports custom salt
'          - Returns detailed error information via LastError property
'
' Dependencies: cCryptoHash (src/Crypt/Hash)
'
'=========================================================================

Private Const MODULE_NAME As String = "cPassword"

Private m_Hash      As cCryptoHash
Private m_Algorithm As CryptoHashAlgorithm
Private m_Salt      As String
Private m_Encoding  As CryptoStringEncoding
Private m_LastError As String

Private Sub Class_Initialize()
    Set m_Hash = New cCryptoHash
    m_Algorithm = HASH_ALG_MD5
    m_Salt = vbNullString
    m_Encoding = ENCODING_UTF8
    m_LastError = vbNullString
End Sub

Private Sub Class_Terminate()
    Set m_Hash = Nothing
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Property Get Algorithm() As CryptoHashAlgorithm
    Algorithm = m_Algorithm
End Property

Public Property Let Algorithm(ByVal Value As CryptoHashAlgorithm)
    m_Algorithm = Value
End Property

Public Property Get Salt() As String
    Salt = m_Salt
End Property

Public Property Let Salt(ByVal Value As String)
    m_Salt = Value
End Property

Public Property Get Encoding() As CryptoStringEncoding
    Encoding = m_Encoding
End Property

Public Property Let Encoding(ByVal Value As CryptoStringEncoding)
    m_Encoding = Value
End Property

'=========================================================================
' Public Methods
'=========================================================================

'Create a password hash
'@param Password The plain-text password to hash
'@return Hexadecimal hash string, or empty string on error
Public Function Create(ByVal Password As String, Optional ByRef Salt As Variant) As String
    On Error GoTo ErrorHandler
    
    m_LastError = vbNullString
    
    If LenB(Password) = 0 Then
        m_LastError = "Password cannot be empty."
        Exit Function
    End If
    
    Dim sSalt As String
    If IsMissing(Salt) Then
        sSalt = m_Salt
    Else
        sSalt = CStr(Salt)
    End If
    
    Create = m_Hash.ComputeHash(sSalt & Password, m_Algorithm, m_Encoding)
    
    Exit Function
ErrorHandler:
    m_LastError = "[" & MODULE_NAME & ".Create] " & ERR.Description
End Function

'Verify a password against a stored hash
'@param Password The plain-text password to verify
'@param Hash The stored hash to compare against
'@return True if password matches the hash, False otherwise
'@remark On validation failure (empty input or mismatch), LastError is set
Public Function Verify(ByVal Password As String, ByVal Hash As String) As Boolean
    On Error GoTo ErrorHandler
    
    m_LastError = vbNullString
    
    If LenB(Password) = 0 Then
        m_LastError = "Password cannot be empty."
        Exit Function
    End If
    
    If LenB(Hash) = 0 Then
        m_LastError = "Hash cannot be empty."
        Exit Function
    End If
    
    Dim sComputed As String
    sComputed = m_Hash.ComputeHash(m_Salt & Password, m_Algorithm, m_Encoding)
    
    Verify = (StrComp(sComputed, Hash, vbBinaryCompare) = 0)
    
    If Not Verify Then
        m_LastError = "Password verification failed."
    End If
    
    Exit Function
ErrorHandler:
    m_LastError = "[" & MODULE_NAME & ".Verify] " & ERR.Description
End Function
