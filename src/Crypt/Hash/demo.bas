Attribute VB_Name = "demoCryptoHash"

'=========================================================================
'
' demoCryptoHash - cCryptoHash Class Demo Module
'
' Purpose: Demonstrates usage of cCryptoHash class
'
' Author: Auto
' Date: 2026-02-08
'
'=========================================================================

Option Explicit

'=========================================================================
' 公共方法 - Demo Functions
'=========================================================================

'演示基本哈希计算
Public Sub DemoBasicHash()
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    
    Dim sText As String
    sText = "Hello VBMAN!"
    
    Debug.Print "=== Basic Hash Demo ==="
    Debug.Print "Text: " & sText
    Debug.Print ""
    
    ' MD5
    Hash.Algorithm = HASH_ALG_MD5
    Debug.Print "MD5: " & Hash.ComputeHash(sText)
    
    ' SHA1
    Hash.Algorithm = HASH_ALG_SHA1
    Debug.Print "SHA1: " & Hash.ComputeHash(sText)
    
    ' SHA256
    Hash.Algorithm = HASH_ALG_SHA256
    Debug.Print "SHA256: " & Hash.ComputeHash(sText)
    
    ' SHA384
    Hash.Algorithm = HASH_ALG_SHA384
    Debug.Print "SHA384: " & Hash.ComputeHash(sText)
    
    ' SHA512
    Hash.Algorithm = HASH_ALG_SHA512
    Debug.Print "SHA512: " & Hash.ComputeHash(sText)
    Debug.Print ""
    
    Set Hash = Nothing
End Sub

'演示字节数组哈希计算
Public Sub DemoByteArrayHash()
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    Hash.Algorithm = HASH_ALG_SHA256
    
    Dim baData() As Byte
    Dim i As Long
    
    ' 创建测试字节数组
    ReDim baData(15) As Byte
    For i = 0 To 15
        baData(i) = i * 10
    Next i
    
    Debug.Print "=== Byte Array Hash Demo ==="
    
    ' 方法1: 返回十六进制字符串
    Dim sHex As String
    sHex = Hash.ComputeHashBytesToHex(baData)
    Debug.Print "Hex: " & sHex
    
    ' 方法2: 返回原始字节数组
    Dim baHash() As Byte
    baHash = Hash.ComputeHashBytes(baData)
    Debug.Print "Hash bytes length: " & (UBound(baHash) + 1)
    Debug.Print ""
    
    Set Hash = Nothing
End Sub

'演示大字符串哈希
Public Sub DemoLargeStringHash()
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    Hash.Algorithm = HASH_ALG_SHA256
    
    ' 创建大字符串
    Dim sText As String
    Dim i As Long
    sText = String$(10000, "A")
    
    Debug.Print "=== Large String Hash Demo ==="
    Debug.Print "String length: " & Len(sText)
    
    Dim dStart As Double
    Dim dEnd As Double
    Dim sHash As String
    
    ' 测量时间
    dStart = Timer
    sHash = Hash.ComputeHash(sText)
    dEnd = Timer
    
    Debug.Print "Hash: " & Left$(sHash, 32) & "..."
    Debug.Print "Time: " & Format$((dEnd - dStart) * 1000, "0.00") & " ms"
    Debug.Print ""
    
    Set Hash = Nothing
End Sub

'演示文件哈希计算
Public Sub DemoFileHash()
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    
    Debug.Print "=== File Hash Demo ==="
    Debug.Print ""
    
    ' 使用当前程序文件作为示例
    Dim sThisFile As String
    sThisFile = App.Path & "\" & App.EXEName & ".exe"
    
    ' 检查文件是否存在
    If LenB(Dir$(sThisFile)) > 0 Then
        Debug.Print "File: " & sThisFile
        Debug.Print ""
        
        ' SHA256
        Debug.Print "SHA256: " & Hash.ComputeFileHash(sThisFile, HASH_ALG_SHA256)
        
        ' MD5
        Debug.Print "MD5: " & Hash.ComputeFileHash(sThisFile, HASH_ALG_MD5)
        
        ' SHA1
        Debug.Print "SHA1: " & Hash.ComputeFileHash(sThisFile, HASH_ALG_SHA1)
        
        Debug.Print ""
    Else
        Debug.Print "File not found, using current module as example"
        Debug.Print ""
        
        ' 使用当前模块作为示例
        sThisFile = App.Path & "\demo.bas"
        If LenB(Dir$(sThisFile)) > 0 Then
            Debug.Print "File: " & sThisFile
            Debug.Print ""
            Debug.Print "SHA256: " & Hash.ComputeFileHash(sThisFile, HASH_ALG_SHA256)
            Debug.Print ""
        End If
    End If
    
    Set Hash = Nothing
End Sub

'演示不同提供程序
Public Sub DemoDifferentProviders()
    Debug.Print "=== Different Providers Demo ==="
    
    Dim sText As String
    sText = "Test Provider"
    
    Dim sHash1 As String
    Dim sHash2 As String
    
    ' 默认提供程序
    Dim Hash1 As cCryptoHash
    Set Hash1 = New cCryptoHash
    Hash1.Algorithm = HASH_ALG_SHA256
    sHash1 = Hash1.ComputeHash(sText)
    Debug.Print "Default provider: " & sHash1
    
    ' 标准提供程序
    Dim Hash2 As cCryptoHash
    Set Hash2 = New cCryptoHash
    Hash2.ProviderName = "Microsoft Base Cryptographic Provider v1.0"
    Hash2.Algorithm = HASH_ALG_SHA256
    sHash2 = Hash2.ComputeHash(sText)
    Debug.Print "Base provider: " & sHash2
    
    Debug.Print "Are hashes equal: " & (sHash1 = sHash2)
    Debug.Print ""
    
    Set Hash1 = Nothing
    Set Hash2 = Nothing
End Sub

'演示哈希一致性
Public Sub DemoHashConsistency()
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    Hash.Algorithm = HASH_ALG_SHA256
    
    Dim sText As String
    sText = "Consistency Test"
    
    Dim sHash1 As String
    Dim sHash2 As String
    Dim sHash3 As String
    
    ' 计算三次
    sHash1 = Hash.ComputeHash(sText)
    sHash2 = Hash.ComputeHash(sText)
    sHash3 = Hash.ComputeHash(sText)
    
    Debug.Print "=== Hash Consistency Demo ==="
    Debug.Print "Hash 1: " & sHash1
    Debug.Print "Hash 2: " & sHash2
    Debug.Print "Hash 3: " & sHash3
    Debug.Print "All equal: " & (sHash1 = sHash2 And sHash2 = sHash3)
    Debug.Print ""
    
    Set Hash = Nothing
End Sub

'运行所有演示
Public Sub RunAllDemos()
    Debug.Print "=========================================="
    Debug.Print "  cCryptoHash - Complete Demo Suite"
    Debug.Print "=========================================="
    Debug.Print ""
    
    Call DemoBasicHash
    Call DemoByteArrayHash
    Call DemoLargeStringHash
    Call DemoFileHash
    Call DemoDifferentProviders
    Call DemoHashConsistency
    Call DemoNewFeatures
    
    Debug.Print "=========================================="
    Debug.Print "  All demos completed!"
    Debug.Print "=========================================="
End Sub

'演示新功能：一行代码完成哈希计算
Public Sub DemoNewFeatures()
    Debug.Print "=== New Features Demo ==="
    Debug.Print ""
    
    ' 1. 一行代码计算 SHA256 (使用默认算法)
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    
    Dim sText As String
    sText = "Hello World"
    
    Debug.Print "1. Default algorithm (SHA256):"
    Debug.Print "   " & Hash.ComputeHash(sText)
    Debug.Print ""
    
    ' 2. 一行代码计算 MD5 (指定算法)
    Debug.Print "2. One-line MD5:"
    Debug.Print "   " & Hash.ComputeHash(sText, HASH_ALG_MD5)
    Debug.Print ""
    
    ' 3. 一行代码计算 SHA512 (指定算法)
    Debug.Print "3. One-line SHA512:"
    Debug.Print "   " & Hash.ComputeHash(sText, HASH_ALG_SHA512)
    Debug.Print ""
    
    ' 4. UTF8 编码 (默认)
    Debug.Print "4. UTF8 encoding (default):"
    Debug.Print "   " & Hash.ComputeHash(sText, HASH_ALG_SHA256, ENCODING_UTF8)
    Debug.Print ""
    
    ' 5. ANSI 编码
    Debug.Print "5. ANSI encoding:"
    Debug.Print "   " & Hash.ComputeHash(sText, HASH_ALG_SHA256, ENCODING_ANSI)
    Debug.Print ""
    
    ' 6. 中文字符串测试
    Dim sChinese As String
    sChinese = "你好世界 Hello World"
    
    Debug.Print "6. Chinese text UTF8:"
    Debug.Print "   " & Hash.ComputeHash(sChinese, HASH_ALG_SHA256, ENCODING_UTF8)
    Debug.Print ""
    
    Debug.Print "7. Chinese text ANSI:"
    Debug.Print "   " & Hash.ComputeHash(sChinese, HASH_ALG_SHA256, ENCODING_ANSI)
    Debug.Print ""
    
    ' 7. 字节数组一行代码计算
    Dim baData() As Byte
    baData = StrConv(sText, vbFromUnicode)
    
    Debug.Print "8. Byte array one-line hex:"
    Debug.Print "   " & Hash.ComputeHashBytesToHex(baData, HASH_ALG_MD5)
    Debug.Print ""
    
    Debug.Print "9. Byte array one-line raw:"
    Dim baHash() As Byte
    baHash = Hash.ComputeHashBytes(baData, HASH_ALG_SHA256)
    Debug.Print "   Hash size: " & (UBound(baHash) + 1) & " bytes"
    Debug.Print ""
    
    Set Hash = Nothing
End Sub
