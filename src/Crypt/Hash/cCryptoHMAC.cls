VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCryptoHMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cCryptoHMAC - HMAC (Hash-based Message Authentication Code) Implementation
'
' Purpose: Provides HMAC computation using standard hash algorithms
'          - Support for HMAC-SHA1, HMAC-SHA256
'          - String and byte array input
'          - Hex string and Base64 output
'          - Chain call support (fluent API)
'
' Author: Auto
' Date: 2026-02-08
'
'=========================================================================

Option Explicit

'=========================================================================
' API Declarations
'=========================================================================
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
    (ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long

Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
    (ByVal hHash As Long) As Long

Private Declare Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As Long, _
    ByRef pbData As Byte, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
    (ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptDeriveKey Lib "advapi32.dll" _
    (ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hBaseData As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long

Private Declare Function CryptDestroyKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Private Declare Function CryptExportKey Lib "advapi32.dll" _
    (ByVal hKey As Long, _
    ByVal hExpKey As Long, _
    ByVal dwBlobType As Long, _
    ByVal dwFlags As Long, _
    ByRef pbData As Byte, _
    ByRef pdwDataLen As Long) As Long

    '=========================================================================
    ' Constants
    '=========================================================================
Private Const MODULE_NAME As String = "cCryptoHMAC"
Private Const PROV_RSA_FULL As Long = 1
Private Const PROV_RSA_AES As Long = 24
Private Const DEF_CRYPT_VERIFYCONTEXT As Long = &HF0000000

' Hash parameters
Private Const HP_HASHSIZE As Long = 4
Private Const HP_HASHVAL As Long = 2

' Hash algorithms
Private Const CALG_MD5 As Long = 32771
Private Const CALG_SHA1 As Long = 32772
Private Const CALG_SHA_256 As Long = 32780

' Block sizes (in bytes)
Private Const SHA1_BLOCK_SIZE As Long = 64
Private Const SHA256_BLOCK_SIZE As Long = 64

'=========================================================================
' Enum Definitions
'=========================================================================
Public Enum HMACAlgorithm
    HMAC_ALG_SHA1 = CALG_SHA1
    HMAC_ALG_SHA256 = CALG_SHA_256
End Enum

Public Enum HMACStringEncoding
    HMAC_ENCODING_ANSI = 0
    HMAC_ENCODING_UTF8 = 1
End Enum

Public Enum SecretKeyType
    SECRET_KEY_STRING = 0
    SECRET_KEY_HEX = 1
    SECRET_KEY_BASE64 = 2
End Enum

'=========================================================================
' Private Variables
'=========================================================================
Private m_Algorithm As HMACAlgorithm
Private m_Key() As Byte
Private m_KeyIsSet As Boolean
Private m_Utf8 As cToolsUtf8

' Chain call variables
Private m_ChainAlgorithm As HMACAlgorithm
Private m_ChainKey() As Byte
Private m_ChainData() As Byte
Private m_ChainDataHasBeenSet As Boolean
Private m_ChainKeyHasBeenSet As Boolean

' Base64 encoding helper objects (lazy initialization)
Private m_Base64XML As Object
Private m_Base64Node As Object

'=========================================================================
' Class Events
'=========================================================================
Private Sub Class_Initialize()
    ' Default algorithm is SHA256
    m_Algorithm = HMAC_ALG_SHA256
    m_KeyIsSet = False
    
    ' Initialize UTF8 converter
    Set m_Utf8 = New cToolsUtf8
    
    ' Initialize chain variables
    m_ChainAlgorithm = HMAC_ALG_SHA256
    m_ChainKeyHasBeenSet = False
    m_ChainDataHasBeenSet = False
End Sub

Private Sub Class_Terminate()
    ' Cleanup UTF8 converter
    Set m_Utf8 = Nothing
    
    ' Cleanup Base64 helper objects
    Set m_Base64Node = Nothing
    Set m_Base64XML = Nothing
End Sub

'=========================================================================
' Properties
'=========================================================================
Public Property Get Algorithm() As HMACAlgorithm
    Algorithm = m_Algorithm
End Property

Public Property Let Algorithm(ByVal Value As HMACAlgorithm)
    m_Algorithm = Value
End Property

'=========================================================================
' Public Methods - Key Setup
'=========================================================================

'Set HMAC key from string
'@param KeyString The secret key string
'@param Encoding String encoding (ENCODING_UTF8=1, ENCODING_ANSI=0), default UTF8
Public Sub SetKey(ByVal KeyString As String, _
                  Optional ByVal Encoding As HMACStringEncoding = HMAC_ENCODING_UTF8)
    If Encoding = HMAC_ENCODING_UTF8 Then
        m_Key = StringToUTF8Bytes(KeyString)
    Else
        m_Key = StrConv(KeyString, vbFromUnicode)
    End If
    m_KeyIsSet = True
End Sub

'Set HMAC key from byte array
'@param KeyBytes The secret key as byte array
Public Sub SetKeyBytes(ByRef KeyBytes() As Byte)
    m_Key = KeyBytes
    m_KeyIsSet = True
End Sub

'=========================================================================
' Public Methods - HMAC Computation (Traditional)
'=========================================================================

'Compute HMAC for string data (returns hex string)
'@param Data The data to sign
'@param KeyString The secret key string
'@param Algorithm HMAC algorithm, default uses class property
'@param Encoding String encoding, default UTF8
'@return Hexadecimal HMAC string
Public Function Compute(ByVal Data As String, _
                        ByVal KeyString As String, _
                        Optional ByVal Algorithm As HMACAlgorithm, _
                        Optional ByVal Encoding As HMACStringEncoding = HMAC_ENCODING_UTF8) As String
    On Error GoTo ErrorHandler
    
    Dim baData() As Byte
    Dim baKey() As Byte
    Dim baHMAC() As Byte
    
    ' Convert data and key to bytes
    If Encoding = HMAC_ENCODING_UTF8 Then
        baData = StringToUTF8Bytes(Data)
        baKey = StringToUTF8Bytes(KeyString)
    Else
        baData = StrConv(Data, vbFromUnicode)
        baKey = StrConv(KeyString, vbFromUnicode)
    End If
    
    ' Compute HMAC
    baHMAC = ComputeHMACBytes(baData, baKey, Algorithm)
    
    ' Return as hex
    Compute = BytesToHex(baHMAC)
    
    Exit Function
ErrorHandler:
    RaiseError "Compute failed: " & ERR.Description
End Function

'Compute HMAC for byte array (returns hex string)
'@param Data The data to sign
'@param Key The secret key as byte array
'@param Algorithm HMAC algorithm, default uses class property
'@return Hexadecimal HMAC string
Public Function ComputeBytesToHex(ByRef Data() As Byte, _
                                  ByRef Key() As Byte, _
                                  Optional ByVal Algorithm As HMACAlgorithm) As String
    On Error GoTo ErrorHandler
    
    Dim baHMAC() As Byte
    baHMAC = ComputeHMACBytes(Data, Key, Algorithm)
    ComputeBytesToHex = BytesToHex(baHMAC)
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeBytesToHex failed: " & ERR.Description
End Function

'Compute HMAC for byte array (returns byte array)
'@param Data The data to sign
'@param Key The secret key as byte array
'@param Algorithm HMAC algorithm, default uses class property
'@return HMAC as byte array
Public Function ComputeBytes(ByRef Data() As Byte, _
                             ByRef Key() As Byte, _
                             Optional ByVal Algorithm As HMACAlgorithm) As Byte()
    On Error GoTo ErrorHandler
    
    ComputeBytes = ComputeHMACBytes(Data, Key, Algorithm)
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeBytes failed: " & ERR.Description
End Function

'=========================================================================
' Chain Call Methods (Fluent API)
'=========================================================================

'Set HMAC algorithm (chain call)
'@param Algorithm HMAC algorithm type
'@return Returns self for chaining
Public Function Mode(ByVal Algorithm As HMACAlgorithm) As cCryptoHMAC
    m_ChainAlgorithm = Algorithm
    Set Mode = Me
End Function

'Set key from string (chain call)
'@param KeyString The secret key string
'@param KeyType Key input type (STRING/HEX/BASE64), default STRING
'@param Encoding String encoding, default UTF8
'@return Returns self for chaining
Public Function Secret(ByVal KeyString As String, _
                    Optional ByVal KeyType As SecretKeyType = SECRET_KEY_STRING, _
                    Optional ByVal Encoding As HMACStringEncoding = HMAC_ENCODING_UTF8) As cCryptoHMAC
    Select Case KeyType
        Case SECRET_KEY_HEX
            m_ChainKey = HexToBytes(KeyString)
        Case SECRET_KEY_BASE64
            m_ChainKey = Base64ToBytes(KeyString)
        Case Else 'SECRET_KEY_STRING
            If Encoding = HMAC_ENCODING_UTF8 Then
                m_ChainKey = StringToUTF8Bytes(KeyString)
            Else
                m_ChainKey = StrConv(KeyString, vbFromUnicode)
            End If
    End Select
    m_ChainKeyHasBeenSet = True
    Set Secret = Me
End Function

'Set key from byte array (chain call)
'@param KeyBytes The secret key as byte array
'@return Returns self for chaining
Public Function SecretBytes(ByRef KeyBytes() As Byte) As cCryptoHMAC
    m_ChainKey = KeyBytes
    m_ChainKeyHasBeenSet = True
    Set SecretBytes = Me
End Function

'Input string data (chain call)
'@param Text Data to sign
'@param Encoding String encoding, default UTF8
'@return Returns self for chaining
Public Function DataString(ByVal Text As String, _
                           Optional ByVal Encoding As HMACStringEncoding = HMAC_ENCODING_UTF8) As cCryptoHMAC
    If Encoding = HMAC_ENCODING_UTF8 Then
        m_ChainData = StringToUTF8Bytes(Text)
    Else
        m_ChainData = StrConv(Text, vbFromUnicode)
    End If
    m_ChainDataHasBeenSet = True
    Set DataString = Me
End Function

'Input byte array data (chain call)
'@param Data Data to sign as byte array
'@return Returns self for chaining
Public Function DataBytes(ByRef Data() As Byte) As cCryptoHMAC
    m_ChainData = Data
    m_ChainDataHasBeenSet = True
    Set DataBytes = Me
End Function

'Return HMAC as hexadecimal string (chain call)
'@param UpperCase Optional, use uppercase letters, default False (lowercase)
'@return Hexadecimal HMAC string
Public Function ReturnHex(Optional ByVal UpperCase As Boolean = False) As String
    Dim baHMAC() As Byte
    Dim sHex As String
    Dim i As Long

    ' Validate state
    If Not m_ChainKeyHasBeenSet Then
        ERR.Raise vbObjectError + 1, MODULE_NAME, "Key not set. Please call Secret or SecretBytes first."
    End If
    If Not m_ChainDataHasBeenSet Then
        ERR.Raise vbObjectError + 2, MODULE_NAME, "No data has been set. Please call DataString or DataBytes first."
    End If

    ' Compute HMAC
    baHMAC = ComputeHMACBytes(m_ChainData, m_ChainKey, m_ChainAlgorithm)

    ' Convert to hex
    sHex = ""
    For i = 0 To UBound(baHMAC)
        sHex = sHex & Right$("0" & Hex$(baHMAC(i)), 2)
    Next i

    ' Set case
    If UpperCase Then
        ReturnHex = UCase$(sHex)
    Else
        ReturnHex = LCase$(sHex)
    End If
End Function

'Return HMAC as Base64 string (chain call)
'@return Base64 encoded HMAC string
Public Function ReturnBase64() As String
    Dim baHMAC() As Byte
    Dim oBase64 As Object

    ' Validate state
    If Not m_ChainKeyHasBeenSet Then
        ERR.Raise vbObjectError + 1, MODULE_NAME, "Key not set. Please call Secret or SecretBytes first."
    End If
    If Not m_ChainDataHasBeenSet Then
        ERR.Raise vbObjectError + 2, MODULE_NAME, "No data has been set. Please call DataString or DataBytes first."
    End If

    ' Compute HMAC
    baHMAC = ComputeHMACBytes(m_ChainData, m_ChainKey, m_ChainAlgorithm)

    ' Convert to Base64 using cToolsBase64 if available, otherwise use simple Base64
    On Error Resume Next
    Set oBase64 = CreateObject("cToolsBase64")
    If ERR.Number = 0 Then
        ReturnBase64 = oBase64.BytesToBase64(baHMAC)
    Else
        ' Fallback to simple Base64 using MSXML
        ReturnBase64 = BytesToBase64(baHMAC)
    End If
    On Error GoTo 0
    
    Set oBase64 = Nothing
End Function

'Return HMAC as byte array (chain call)
'@return HMAC as byte array
Public Function ReturnBytes() As Byte()
    ' Validate state
    If Not m_ChainKeyHasBeenSet Then
        ERR.Raise vbObjectError + 1, MODULE_NAME, "Key not set. Please call Secret or SecretBytes first."
    End If
    If Not m_ChainDataHasBeenSet Then
        ERR.Raise vbObjectError + 2, MODULE_NAME, "No data has been set. Please call DataString or DataBytes first."
    End If

    ' Compute and return HMAC
    ReturnBytes = ComputeHMACBytes(m_ChainData, m_ChainKey, m_ChainAlgorithm)
End Function

'=========================================================================
' Private Methods - HMAC Implementation
'=========================================================================

'Core HMAC computation using RFC 2104 algorithm
'@param Data Data to sign
'@param Key Secret key
'@param Algorithm HMAC algorithm
'@return HMAC as byte array
Private Function ComputeHMACBytes(ByRef Data() As Byte, _
    ByRef Key() As Byte, _
    Optional ByVal Algorithm As HMACAlgorithm) As Byte()
    On Error GoTo ErrorHandler
    
    Dim lAlgorithm As Long
    Dim BlockSize As Long
    Dim i As Long
    
    ' Determine algorithm
    If Algorithm = 0 Then
        lAlgorithm = m_Algorithm
    Else
        lAlgorithm = Algorithm
    End If
    
    ' Determine block size
    If lAlgorithm = CALG_SHA1 Then
        BlockSize = SHA1_BLOCK_SIZE
    ElseIf lAlgorithm = CALG_SHA_256 Then
        BlockSize = SHA256_BLOCK_SIZE
    Else
        BlockSize = 64                                                          ' Default
    End If
    
    ' Prepare key (RFC 2104 Step 1)
    Dim K() As Byte
    Dim KeyLen As Long
    KeyLen = UBound(Key) + 1
    
    If KeyLen > BlockSize Then
        ' Key is longer than block size, hash it
        K = ComputeSimpleHash(Key, lAlgorithm)
        KeyLen = UBound(K) + 1
    Else
        ' Copy key
        ReDim K(KeyLen - 1)
        CopyMemory K, Key, KeyLen
    End If
    
    ' Pad key to block size
    If KeyLen < BlockSize Then
        ReDim Preserve K(BlockSize - 1)
        ' Remaining bytes are already 0
    End If
    
    ' Create inner and outer padded keys (RFC 2104 Step 2)
    Dim ipad() As Byte
    Dim opad() As Byte
    ReDim ipad(BlockSize - 1)
    ReDim opad(BlockSize - 1)
    
    For i = 0 To BlockSize - 1
        ipad(i) = K(i) Xor &H36
        opad(i) = K(i) Xor &H5C
    Next i
    
    ' Compute inner hash: H(K XOR ipad || text) (RFC 2104 Step 4)
    Dim InnerData() As Byte
    Dim InnerDataLen As Long
    Dim DataLen As Long
    
    DataLen = UBound(Data) + 1
    InnerDataLen = BlockSize + DataLen
    ReDim InnerData(InnerDataLen - 1)
    
    ' Copy ipad || data
    CopyMemory InnerData, ipad, BlockSize
    CopyMemory InnerData, Data, DataLen, BlockSize, 0
    
    Dim InnerHash() As Byte
    InnerHash = ComputeSimpleHash(InnerData, lAlgorithm)
    
    ' Compute outer hash: H(K XOR opad || inner_hash) (RFC 2104 Step 7)
    Dim OuterData() As Byte
    Dim OuterDataLen As Long
    Dim HashLen As Long
    
    HashLen = UBound(InnerHash) + 1
    OuterDataLen = BlockSize + HashLen
    ReDim OuterData(OuterDataLen - 1)
    
    ' Copy opad || inner_hash
    CopyMemory OuterData, opad, BlockSize
    CopyMemory OuterData, InnerHash, HashLen, BlockSize, 0
    
    ' Final HMAC result
    ComputeHMACBytes = ComputeSimpleHash(OuterData, lAlgorithm)
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeHMACBytes failed: " & ERR.Description
End Function

'Simple hash computation without HMAC
'@param Data Data to hash
'@param Algorithm Hash algorithm
'@return Hash as byte array
Private Function ComputeSimpleHash(ByRef Data() As Byte, ByVal Algorithm As Long) As Byte()
    On Error GoTo ErrorHandler
    
    Dim hProv As Long
    Dim hHash As Long
    Dim sProvider As String
    Dim dwProvType As Long
    Dim bSuccess As Boolean
    
    ' Choose provider based on algorithm
    If Algorithm = CALG_SHA_256 Then
        sProvider = "Microsoft Enhanced RSA and AES Cryptographic Provider"
        dwProvType = PROV_RSA_AES
    Else
        sProvider = vbNullString
        dwProvType = PROV_RSA_FULL
    End If
    
    ' Acquire context
    If CryptAcquireContext(hProv, vbNullString, sProvider, dwProvType, DEF_CRYPT_VERIFYCONTEXT) = 0 Then
        ' Try fallback
        If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, DEF_CRYPT_VERIFYCONTEXT) = 0 Then
            ERR.Raise vbObjectError + 10, MODULE_NAME, "Failed to acquire cryptographic context"
        End If
    End If
    
    ' Create hash
    If CryptCreateHash(hProv, Algorithm, 0, 0, hHash) = 0 Then
        CryptReleaseContext hProv, 0
        ERR.Raise vbObjectError + 11, MODULE_NAME, "Failed to create hash object"
    End If
    
    ' Hash data
    If CryptHashData(hHash, Data(0), UBound(Data) + 1, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        ERR.Raise vbObjectError + 12, MODULE_NAME, "Failed to hash data"
    End If
    
    ' Get hash size
    Dim dwHashSize As Long
    Dim dwHashLen As Long
    dwHashLen = 4
    
    If CryptGetHashParam(hHash, HP_HASHSIZE, dwHashSize, dwHashLen, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        ERR.Raise vbObjectError + 13, MODULE_NAME, "Failed to get hash size"
    End If
    
    ' Get hash value
    Dim baHash() As Byte
    ReDim baHash(dwHashSize - 1)
    
    dwHashLen = dwHashSize
    If CryptGetHashParam(hHash, HP_HASHVAL, baHash(0), dwHashLen, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        ERR.Raise vbObjectError + 14, MODULE_NAME, "Failed to get hash value"
    End If
    
    ' Cleanup
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    
    ComputeSimpleHash = baHash
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeSimpleHash failed: " & ERR.Description
End Function

'=========================================================================
' Helper Functions
'=========================================================================

'Convert string to UTF-8 bytes
Private Function StringToUTF8Bytes(ByVal Text As String) As Byte()
    StringToUTF8Bytes = m_Utf8.Encode(Text)
End Function

'Convert bytes to hex string
Private Function BytesToHex(ByRef Bytes() As Byte) As String
    Dim i As Long
    Dim sHex As String
    
    For i = 0 To UBound(Bytes)
        sHex = sHex & Right$("0" & Hex$(Bytes(i)), 2)
    Next i
    
    BytesToHex = sHex
End Function

'Convert hex string to bytes
Private Function HexToBytes(ByVal HexString As String) As Byte()
    Dim i As Long
    Dim ByteLen As Long
    Dim result() As Byte
    
    ' Remove spaces and validate
    HexString = Replace(HexString, " ", "")
    HexString = Replace(HexString, "-", "")
    
    If Len(HexString) = 0 Then
        HexToBytes = result
        Exit Function
    End If
    
    ' Handle odd length
    If Len(HexString) Mod 2 = 1 Then
        HexString = "0" & HexString
    End If
    
    ByteLen = Len(HexString) \ 2
    ReDim result(ByteLen - 1)
    
    On Error Resume Next
    For i = 0 To ByteLen - 1
        result(i) = CByte("&H" & Mid$(HexString, i * 2 + 1, 2))
    Next i
    On Error GoTo 0
    
    HexToBytes = result
End Function

'Convert Base64 string to bytes
Private Function Base64ToBytes(ByVal Base64String As String) As Byte()
    Dim oXML As Object
    Dim oNode As Object
    Dim result() As Byte
    
    On Error Resume Next
    Set oXML = CreateObject("MSXML2.DOMDocument")
    Set oNode = oXML.createElement("b64")
    oNode.DataType = "bin.base64"
    oNode.Text = Base64String
    result = oNode.nodeTypedValue
    On Error GoTo 0
    
    Base64ToBytes = result
End Function

'Simple memory copy with optional offset support
Private Sub CopyMemory(ByRef Dest() As Byte, ByRef Source() As Byte, ByVal Length As Long, _
                       Optional ByVal DestOffset As Long = 0, Optional ByVal SourceOffset As Long = 0)
    Dim i As Long
    For i = 0 To Length - 1
        Dest(DestOffset + i) = Source(SourceOffset + i)
    Next i
End Sub

'Simple Base64 encoding fallback (lazy initialization)
Private Function BytesToBase64(ByRef Data() As Byte) As String
    ' Lazy initialization of Base64 helper objects
    If m_Base64XML Is Nothing Then
        Set m_Base64XML = CreateObject("MSXML2.DOMDocument")
        Set m_Base64Node = m_Base64XML.createElement("b64")
    End If
    
    m_Base64Node.DataType = "bin.base64"
    m_Base64Node.nodeTypedValue = Data
    
    BytesToBase64 = m_Base64Node.Text
End Function

'=========================================================================
' Error Handling
'=========================================================================
Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[" & MODULE_NAME & "] Error: " & Description
    ERR.Raise vbObjectError + 1000, MODULE_NAME, Description
End Sub






