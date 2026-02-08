VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCryptoHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

    '=========================================================================
    '
    ' cCryptoHash - Windows CryptoAPI Hash Algorithm Implementation
    '
    ' Purpose: Provides hash computation functionality using Windows CryptoAPI
    '          - Support for MD5, SHA1, SHA256, SHA384, SHA512
    '          - String and byte array hash computation
    '          - File hash computation (binary mode, no encoding conversion)
    '          - Hex string and raw byte array output
    '
    ' Author: Auto
    ' Date: 2026-02-08
    '
    '=========================================================================

Option Explicit

'=========================================================================
' API 声明
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


    '=========================================================================
    ' 常量定义
    '=========================================================================
Private Const MODULE_NAME As String = "cCryptoHash"
Private Const PROV_RSA_FULL As Long = 1                                         ' 基础加密服务提供程序
Private Const PROV_RSA_AES As Long = 24                                         ' 支持 AES 的高级加密服务提供程序
Private Const DEF_CRYPT_VERIFYCONTEXT As Long = &HF0000000

' Hash parameters
Private Const HP_HASHSIZE As Long = 4
Private Const HP_HASHVAL As Long = 2

' Hash algorithms
Private Const CALG_MD5 As Long = 32771
Private Const CALG_SHA1 As Long = 32772
Private Const CALG_SHA_256 As Long = 32780
Private Const CALG_SHA_384 As Long = 32781
Private Const CALG_SHA_512 As Long = 32782

'=========================================================================
' 枚举定义
'=========================================================================
Public Enum CryptoHashAlgorithm
    HASH_ALG_DEFAULT = 0         ' 使用类属性默认算法
    HASH_ALG_MD5 = CALG_MD5
    HASH_ALG_SHA1 = CALG_SHA1
    HASH_ALG_SHA256 = CALG_SHA_256
    HASH_ALG_SHA384 = CALG_SHA_384
    HASH_ALG_SHA512 = CALG_SHA_512
End Enum

Public Enum CryptoStringEncoding
    ENCODING_ANSI = 0
    ENCODING_UTF8 = 1
End Enum

'=========================================================================
' 私有变量
'=========================================================================
Private m_Algorithm As CryptoHashAlgorithm
Private m_ProviderName As String
Private m_Utf8 As cToolsUtf8

'=========================================================================
' 类事件
'=========================================================================

Private Sub Class_Initialize()
    ' 初始化默认算法为 SHA256
    m_Algorithm = HASH_ALG_SHA256
    
    ' 初始化默认提供程序
    m_ProviderName = "Microsoft Enhanced RSA and AES Cryptographic Provider"
    
    ' 初始化 UTF8 转换对象
    Set m_Utf8 = New cToolsUtf8
End Sub

Private Sub Class_Terminate()
    ' 清理工作在计算完成后自动完成
    Set m_Utf8 = Nothing
End Sub

'=========================================================================
' 属性定义
'=========================================================================

Public Property Get Algorithm() As CryptoHashAlgorithm
    Algorithm = m_Algorithm
End Property

Public Property Let Algorithm(ByVal Value As CryptoHashAlgorithm)
    m_Algorithm = Value
End Property

Public Property Get ProviderName() As String
    ProviderName = m_ProviderName
End Property

Public Property Let ProviderName(ByVal Value As String)
    m_ProviderName = Value
End Property

'=========================================================================
' 公共方法 - Hash Computation
'=========================================================================

'计算字符串的哈希值（返回十六进制字符串）
'@param Text 要哈希的文本字符串
'@param Algorithm 可选参数，哈希算法，默认使用类属性值
'@param Encoding 可选参数，字符串编码（ENCODING_UTF8=1, ENCODING_ANSI=0），默认 UTF8
'@return 十六进制格式的哈希字符串
Public Function ComputeHash(ByVal Text As String, _
                            Optional ByVal Algorithm As CryptoHashAlgorithm, _
                            Optional ByVal Encoding As CryptoStringEncoding = ENCODING_UTF8) As String
    On Error GoTo ErrorHandler
    
    Dim baBytes() As Byte
    
    ' 根据编码类型转换字符串
    If Encoding = ENCODING_UTF8 Then
        baBytes = StringToUTF8Bytes(Text)
    Else
        baBytes = StrConv(Text, vbFromUnicode)
    End If
    
    ' 调用 ComputeHashBytesInternal 计算
    Dim baHash() As Byte
    baHash = ComputeHashBytesInternal(baBytes, Algorithm)
    
    ComputeHash = BytesToHex(baHash)
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeHash failed: " & ERR.Description
End Function

'计算字节数组的哈希值（返回十六进制字符串）
'@param Data 字节数组
'@param Algorithm 可选参数，哈希算法，默认使用类属性值
'@return 十六进制格式的哈希字符串
Public Function ComputeHashBytesToHex(ByRef Data() As Byte, _
                                      Optional ByVal Algorithm As CryptoHashAlgorithm) As String
    On Error GoTo ErrorHandler
    
    Dim baHash() As Byte
    baHash = ComputeHashBytesInternal(Data, Algorithm)
    
    ComputeHashBytesToHex = BytesToHex(baHash)
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeHashBytesToHex failed: " & ERR.Description
End Function

'计算字节数组的哈希值（返回原始字节数组）
'@param Data 字节数组
'@param Algorithm 可选参数，哈希算法，默认使用类属性值
'@return 哈希值的字节数组
Public Function ComputeHashBytes(ByRef Data() As Byte, _
                                 Optional ByVal Algorithm As CryptoHashAlgorithm) As Byte()
    On Error GoTo ErrorHandler
    
    ComputeHashBytes = ComputeHashBytesInternal(Data, Algorithm)
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeHashBytes failed: " & ERR.Description
End Function

'计算文件的哈希值（返回十六进制字符串）
'@param FilePath 文件路径
'@param Algorithm 可选参数，哈希算法，默认使用类属性值
'@return 十六进制格式的哈希字符串
'说明：直接二进制读取文件内容计算哈希，不进行字符编码处理
Public Function ComputeFileHash(ByVal FilePath As String, _
                                 Optional ByVal Algorithm As CryptoHashAlgorithm) As String
    On Error GoTo ErrorHandler
    
    Dim iFileNum As Integer
    Dim baData() As Byte
    Dim lFileSize As Long
    
    ' 检查文件是否存在
    If LenB(Dir$(FilePath)) = 0 Then
        ERR.Raise vbObjectError + 1, MODULE_NAME, "File not found: " & FilePath
    End If
    
    ' 获取文件大小
    lFileSize = FileLen(FilePath)
    
    If lFileSize = 0 Then
        ' 空文件也返回哈希值（空内容的哈希）
        ReDim baData(0 To 0) As Byte
    Else
        ' 读取文件（二进制模式）
        iFileNum = FreeFile
        Open FilePath For Binary Access Read As #iFileNum
        
        ReDim baData(lFileSize - 1) As Byte
        Get #iFileNum, , baData
        
        Close #iFileNum
    End If
    
    ' 计算哈希
    ComputeFileHash = ComputeHashBytesToHex(baData, Algorithm)
    
    Exit Function
ErrorHandler:
    ' 确保文件句柄被关闭
    On Error Resume Next
    If iFileNum > 0 Then Close #iFileNum
    On Error GoTo 0
    
    RaiseError "ComputeFileHash failed: " & ERR.Description
End Function

'=========================================================================
' 私有方法 - Helper Functions
'=========================================================================

'根据提供程序名称获取对应的 Provider Type
'@param sProviderName 提供程序名称
'@return Provider Type 常量
Private Function GetProviderType(ByVal sProviderName As String) As Long
    ' Microsoft Enhanced RSA and AES Cryptographic Provider 需要 PROV_RSA_AES
    If InStr(1, sProviderName, "RSA and AES", vbTextCompare) > 0 Then
        GetProviderType = PROV_RSA_AES
    Else
        ' 其他提供程序使用 PROV_RSA_FULL
        GetProviderType = PROV_RSA_FULL
    End If
End Function

'根据算法获取支持该算法的加密服务提供程序列表（按优先级排序）
'@param lAlgorithm 哈希算法标识符
'@return 提供程序名称数组
Private Function GetProvidersForAlgorithm(ByVal lAlgorithm As Long) As String()
    Dim aProviders(0 To 3) As String
    Dim iCount As Integer
    
    iCount = 0
    
    ' SHA256, SHA384, SHA512 需要支持高级加密的提供程序
    If lAlgorithm = CALG_SHA_256 Or lAlgorithm = CALG_SHA_384 Or lAlgorithm = CALG_SHA_512 Then
        ' 优先尝试支持 SHA2 的提供程序
        aProviders(iCount) = "Microsoft Enhanced RSA and AES Cryptographic Provider"
        iCount = iCount + 1
        aProviders(iCount) = "Microsoft Enhanced Cryptographic Provider v1.0"
        iCount = iCount + 1
        aProviders(iCount) = "Microsoft Strong Cryptographic Provider"
        iCount = iCount + 1
    Else
        ' MD5, SHA1 可以使用大多数提供程序
        aProviders(iCount) = "Microsoft Base Cryptographic Provider v1.0"
        iCount = iCount + 1
        aProviders(iCount) = "Microsoft Enhanced Cryptographic Provider v1.0"
        iCount = iCount + 1
        aProviders(iCount) = "Microsoft Strong Cryptographic Provider"
        iCount = iCount + 1
    End If
    
    ' 默认提供程序（使用空字符串表示系统默认）
    aProviders(iCount) = ""
    
    ' 返回数组
    GetProvidersForAlgorithm = aProviders
End Function

'内部哈希计算函数（带算法参数）
'@param Data 字节数组
'@param Algorithm 可选参数，哈希算法
'@return 哈希值的字节数组
Private Function ComputeHashBytesInternal(ByRef Data() As Byte, _
                                           Optional ByVal Algorithm As CryptoHashAlgorithm) As Byte()
    On Error GoTo ErrorHandler
    
    Dim hProv As Long
    Dim hHash As Long
    Dim lAlgorithm As Long
    
    ' 检查数据是否为空
    If (Not Data) = False Then
        ERR.Raise vbObjectError + 1, MODULE_NAME, "Input data is empty"
    End If
    
    ' 确定使用的算法
    If Algorithm = HASH_ALG_DEFAULT Then
        lAlgorithm = m_Algorithm
    Else
        lAlgorithm = Algorithm
    End If
    
    ' 获取加密服务提供程序句柄并创建哈希对象
    ' 根据算法选择合适的提供程序，按优先级尝试直到找到支持该算法的提供程序
    Dim sProvider As String
    Dim bSuccess As Boolean
    Dim i As Integer
    
    ' 获取支持该算法的提供程序列表
    Dim aProviders() As String
    aProviders = GetProvidersForAlgorithm(lAlgorithm)
    
    ' 按优先级尝试各个提供程序
    For i = LBound(aProviders) To UBound(aProviders)
        sProvider = aProviders(i)
        
        ' 根据提供程序名称获取对应的 Provider Type
        Dim dwProvType As Long
        dwProvType = GetProviderType(sProvider)
        
        ' 尝试获取加密上下文
        If CryptAcquireContext(hProv, _
                               vbNullString, _
                               sProvider, _
                               dwProvType, _
                               DEF_CRYPT_VERIFYCONTEXT) <> 0 Then
            ' 尝试创建哈希对象（检查 Provider 是否支持该算法）
            If CryptCreateHash(hProv, lAlgorithm, 0, 0, hHash) <> 0 Then
                bSuccess = True
                Exit For
            Else
                ' 该 Provider 不支持此算法，释放上下文，尝试下一个
                Call CryptReleaseContext(hProv, 0)
            End If
        End If
    Next i
    
    If Not bSuccess Then
        ERR.Raise vbObjectError + 2, MODULE_NAME, "Failed to acquire cryptographic context or create hash object. No suitable provider found for algorithm."
    End If
    
    ' 计算哈希值
    If CryptHashData(hHash, Data(0), UBound(Data) + 1, 0) = 0 Then
        Call CryptDestroyHash(hHash)
        Call CryptReleaseContext(hProv, 0)
        ERR.Raise vbObjectError + 4, MODULE_NAME, "Failed to hash data"
    End If
    
    ' 获取哈希值大小
    Dim dwHashSize As Long
    Dim dwHashLen As Long
    dwHashLen = 4
    
    If CryptGetHashParam(hHash, HP_HASHSIZE, dwHashSize, dwHashLen, 0) = 0 Then
        Call CryptDestroyHash(hHash)
        Call CryptReleaseContext(hProv, 0)
        ERR.Raise vbObjectError + 5, MODULE_NAME, "Failed to get hash size"
    End If
    
    ' 获取哈希值
    Dim baHash() As Byte
    ReDim baHash(dwHashSize - 1) As Byte
    
    dwHashLen = dwHashSize
    If CryptGetHashParam(hHash, HP_HASHVAL, baHash(0), dwHashLen, 0) = 0 Then
        Call CryptDestroyHash(hHash)
        Call CryptReleaseContext(hProv, 0)
        ERR.Raise vbObjectError + 6, MODULE_NAME, "Failed to get hash value"
    End If
    
    ' 清理资源
    Call CryptDestroyHash(hHash)
    Call CryptReleaseContext(hProv, 0)
    
    ' 返回哈希值
    ComputeHashBytesInternal = baHash
    
    Exit Function
ErrorHandler:
    RaiseError "ComputeHashBytesInternal failed: " & ERR.Description
End Function

'将字符串转换为 UTF8 编码的字节数组
'@param Text 字符串
'@return UTF8 编码的字节数组
Private Function StringToUTF8Bytes(ByVal Text As String) As Byte()
    ' 使用 cToolsUtf8 类进行 UTF-8 编码
    StringToUTF8Bytes = m_Utf8.Encode(Text)
End Function

'将字节数组转换为十六进制字符串
'@param Bytes 字节数组
'@return 十六进制字符串
Private Function BytesToHex(ByRef Bytes() As Byte) As String
    Dim i As Long
    Dim sHex As String
    
    For i = 0 To UBound(Bytes)
        sHex = sHex & Right$("0" & Hex$(Bytes(i)), 2)
    Next i
    
    BytesToHex = sHex
End Function

'=========================================================================
' 错误处理
'=========================================================================

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[" & MODULE_NAME & "] Error: " & Description
    ERR.Raise vbObjectError + 1000, MODULE_NAME, Description
End Sub










