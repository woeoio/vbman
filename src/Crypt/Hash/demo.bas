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

'演示链式调用功能
Public Sub DemoChainedCall()
    Debug.Print "=== Chained Call Demo ==="
    Debug.Print ""
    
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    
    Dim sText As String
    sText = "Hello World"
    
    ' 1. 基本：字符串输入 + 十六进制输出（小写）
    Debug.Print "1. String input + Hex output (lowercase):"
    Dim sResult1 As String
    sResult1 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    Debug.Print "   " & sResult1
    Debug.Print ""
    
    ' 2. 字符串输入 + 十六进制输出（大写）
    Debug.Print "2. String input + Hex output (uppercase):"
    Dim sResult2 As String
    sResult2 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex(True)
    Debug.Print "   " & sResult2
    Debug.Print ""
    
    ' 3. 字符串输入 + Base64 输出
    Debug.Print "3. String input + Base64 output:"
    Dim sResult3 As String
    sResult3 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnBase64()
    Debug.Print "   " & sResult3
    Debug.Print ""
    
    ' 4. 字符串输入 + 字节数组输出
    Debug.Print "4. String input + Bytes output:"
    Dim baResult() As Byte
    baResult = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnBytes()
    Debug.Print "   Hash size: " & (UBound(baResult) + 1) & " bytes"
    Debug.Print ""
    
    ' 5. 字节数组输入 + 十六进制输出
    Debug.Print "5. Byte array input + Hex output:"
    Dim baData() As Byte
    baData = StrConv(sText, vbFromUnicode)
    Dim sResult4 As String
    sResult4 = Hash.Mode(HASH_ALG_MD5).DataBytes(baData).ReturnHex()
    Debug.Print "   " & sResult4
    Debug.Print ""
    
    ' 6. 字节数组输入 + Base64 输出
    Debug.Print "6. Byte array input + Base64 output:"
    Dim sResult5 As String
    sResult5 = Hash.Mode(HASH_ALG_SHA512).DataBytes(baData).ReturnBase64()
    Debug.Print "   " & sResult5
    Debug.Print ""
    
    ' 7. 不同算法比较
    Debug.Print "7. Different algorithms comparison:"
    Dim sMD5 As String, sSHA1 As String, sSHA256 As String
    sMD5 = Hash.Mode(HASH_ALG_MD5).DataString(sText).ReturnHex()
    sSHA1 = Hash.Mode(HASH_ALG_SHA1).DataString(sText).ReturnHex()
    sSHA256 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    Debug.Print "   MD5:    " & sMD5
    Debug.Print "   SHA1:   " & sSHA1
    Debug.Print "   SHA256: " & sSHA256
    Debug.Print ""
    
    ' 8. 中文字符测试
    Debug.Print "8. Chinese text with different encodings:"
    Dim sChinese As String
    sChinese = "你好世界"
    
    Debug.Print "   UTF8 (SHA256): " & Hash.Mode(HASH_ALG_SHA256).DataString(sChinese, ENCODING_UTF8).ReturnHex()
    Debug.Print "   ANSI (SHA256): " & Hash.Mode(HASH_ALG_SHA256).DataString(sChinese, ENCODING_ANSI).ReturnHex()
    Debug.Print ""
    
    ' 9. 链式调用验证一致性
    Debug.Print "9. Verify chain call consistency:"
    Dim sChain1 As String, sChain2 As String, sChain3 As String
    sChain1 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    sChain2 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    sChain3 = Hash.ComputeHash(sText, HASH_ALG_SHA256)
    Debug.Print "   Chain 1: " & sChain1
    Debug.Print "   Chain 2: " & sChain2
    Debug.Print "   Direct:  " & sChain3
    Debug.Print "   All equal: " & (sChain1 = sChain2 And sChain2 = sChain3)
    Debug.Print ""
    
    ' 10. 省略 Mode 的用法（默认使用 SHA256）
    Debug.Print "10. Omit Mode (default SHA256):"
    Dim sWithMode As String, sWithoutMode As String
    sWithMode = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    sWithoutMode = Hash.DataString(sText).ReturnHex()
    Debug.Print "   With Mode:    " & sWithMode
    Debug.Print "   Without Mode: " & sWithoutMode
    Debug.Print "   Equal: " & (sWithMode = sWithoutMode)
    Debug.Print ""
    
    ' 11. 简洁的链式调用（省略 Mode）
    Debug.Print "11. Simplified chain calls (no Mode):"
    Debug.Print "   Hex (default SHA256): " & Hash.DataString(sText).ReturnHex()
    Debug.Print "   Base64 (default SHA256): " & Hash.DataString(sText).ReturnBase64()
    Debug.Print "   Uppercase: " & Hash.DataString(sText).ReturnHex(True)
    Debug.Print ""
    
    ' 12. 重复使用 Hash 对象（自动重置）
    Debug.Print "12. Reuse Hash object (auto reset):"
    Dim sFirst As String, sSecond As String
    sFirst = Hash.DataString("Hello").ReturnHex()
    sSecond = Hash.DataString("World").ReturnHex()
    Debug.Print "   First hash (Hello):  " & sFirst
    Debug.Print "   Second hash (World): " & sSecond
    Debug.Print "   Are different: " & (sFirst <> sSecond)
    Debug.Print ""
    
    ' 13. 连续调用多个输出格式（基于同一数据）
    Debug.Print "13. Multiple output formats (same data):"
    Dim sHex As String, sBase64 As String, sUpper As String
    ' 重复设置同一数据
    sHex = Hash.DataString("Test").ReturnHex()
    sBase64 = Hash.DataString("Test").ReturnBase64()
    sUpper = Hash.DataString("Test").ReturnHex(True)
    Debug.Print "   Hex:      " & sHex
    Debug.Print "   Base64:   " & sBase64
    Debug.Print "   Uppercase: " & sUpper
    Debug.Print ""
    
    ' 14. 手动重置链式调用（使用 Mode）
    Debug.Print "14. Manual reset (using Mode):"
    Dim sBeforeReset As String, sAfterReset As String
    ' 设置数据但不调用 ReturnXxx
    Hash.Mode(HASH_ALG_MD5).DataString("Test1")
    ' 重新开始新的链式调用
    sAfterReset = Hash.Mode(HASH_ALG_SHA256).DataString("Test2").ReturnHex()
    Debug.Print "   After reset (SHA256 of Test2): " & sAfterReset
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
    Call DemoChainedCall
    
    Debug.Print "=========================================="
    Debug.Print "  All demos completed!"
    Debug.Print "=========================================="
End Sub
