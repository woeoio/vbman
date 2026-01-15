# cWinsock 编码指南

## 📖 目录

- [概述](#概述)
- [编码类型](#编码类型)
- [默认编码](#默认编码)
- [编码选择建议](#编码选择建议)
- [常见问题](#常见问题)

---

## 概述

`cWinsock` 支持多种文本编码方式，以适应不同的应用场景。正确使用编码对于确保数据传输的准确性和兼容性至关重要。

---

## 编码类型

### 🇨🇳 ScpAcp (ACP/GBK)

**值**: `0`

**说明**: 系统默认代码页（ANSI Code Page）

**特点**:
- 中文 Windows 上通常为 GBK 编码
- 与 VB6 内部字符串存储方式一致
- 单字节字符 1 字节，中文字符 2 字节
- 适合中文环境下的本地应用

**使用场景**:
- 传统 VB6 应用
- 中文字符为主的应用
- 与现有 VB6 系统兼容

**示例**:
```vb
' 默认使用 ACP/GBK 编码
m_oClient.SendData "中文测试"
m_oClient.GetData sData
```

---

### 🌐 ScpUtf8 (UTF-8)

**值**: `65001`

**说明**: UTF-8 编码

**特点**:
- 国际标准，支持所有 Unicode 字符
- 兼容 ASCII
- 单字节字符 1 字节，中文字符 3 字节
- 网络传输的首选编码

**使用场景**:
- Web 应用
- 国际化应用
- 与现代系统交互
- 需要支持多语言

**示例**:
```vb
' 使用 UTF-8 编码
m_oClient.SendData "中文测试", ScpUtf8
m_oClient.GetData sData, , , ScpUtf8
```

---

### 🌟 ScpUnicode (Unicode)

**值**: `-1`

**说明**: Unicode 编码（不进行转换）

**特点**:
- 保持字符串为宽字符（UTF-16）
- 不进行编码转换
- 每个字符 2 字节（大部分字符）

**使用场景**:
- 内部数据传输
- 需要保持原始字符串格式
- 不想进行编码转换

**示例**:
```vb
' 使用 Unicode（不转换）
m_oClient.SendData "中文测试", ScpUnicode
m_oClient.GetData sData, , , ScpUnicode
```

---

## 默认编码

### SendData 默认编码

```vb
Public Sub SendData(Data As Variant, Optional ByVal CodePage As EnumScpCodePage = ScpAcp)
```

**默认值**: `ScpAcp` (0) - ACP/GBK

### GetData/PeekData 默认编码

```vb
Public Sub GetData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal MaxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = ScpAcp)
Public Sub PeekData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal MaxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = ScpAcp)
```

**默认值**: `ScpAcp` (0) - ACP/GBK

### 设计原因

- 与 VB6 传统编码方式一致
- 与大多数 VB6 应用兼容
- 避免编码不一致导致的乱码问题

---

## 编码选择建议

### 🎯 场景 1: 新建网络应用

**建议**: 使用 UTF-8

```vb
' 发送
m_oClient.SendData "Hello 世界", ScpUtf8

' 接收
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData, , , ScpUtf8
    Debug.Print sData ' "Hello 世界"
End Sub
```

**优点**:
- 国际化支持
- 与现代系统兼容
- Web 标准编码

---

### 🏢 场景 2: 传统 VB6 应用

**建议**: 使用默认 ACP/GBK

```vb
' 发送（默认 ACP）
m_oClient.SendData "中文测试"

' 接收（默认 ACP）
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    Debug.Print sData ' "中文测试"
End Sub
```

**优点**:
- 与 VB6 内部编码一致
- 无需额外编码设置
- 与现有代码兼容

---

### 🌍 场景 3: 混合编码环境

**建议**: 根据对方编码动态选择

```vb
' 发送前检测对方编码
Private Sub SendDataAdaptive(ByVal sText As String)
    If m_oRemoteEncoding = "UTF8" Then
        m_oClient.SendData sText, ScpUtf8
    Else
        m_oClient.SendData sText, ScpAcp
    End If
End Sub

' 接收时检测编码
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    If m_oRemoteEncoding = "UTF8" Then
        Client.GetData sData, , , ScpUtf8
    Else
        Client.GetData sData
    End If
    
    ProcessData sData
End Sub
```

---

### 🔧 场景 4: 协议协商编码

**建议**: 在连接建立时协商编码

```vb
' 连接成功后发送编码协商
Private Sub m_oClient_Connect(Client As cWinsock)
    ' 发送支持的编码列表
    Client.SendData "ENCODING:SUPPORT:ACP,UTF8", ScpUtf8
End Sub

' 服务器响应
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData, , , ScpUtf8
    
    If Left$(sData, 21) = "ENCODING:SUPPORT:" Then
        ' 选择编码
        Dim sEncodings() As String
        sEncodings = Split(Mid$(sData, 22), ",")
        
        ' 优先选择 UTF-8
        Dim sSelected As String
        If InStr("UTF8", sEncodings(0)) > 0 Then
            sSelected = "UTF8"
        Else
            sSelected = "ACP"
        End If
        
        ' 响应选择的编码
        Client.SendData "ENCODING:SELECT:" & sSelected, ScpUtf8
        
        ' 保存客户端使用的编码
        Client.UserData = sSelected
    End If
End Sub
```

---

## 常见问题

### ❓ 问题 1: 中文乱码

**现象**: 接收到的中文显示为乱码

**原因**: 发送和接收使用了不同的编码

**解决方案**:
```vb
' ✅ 正确：发送和接收使用相同编码
m_oClient.SendData "中文测试", ScpUtf8
' 接收时
Client.GetData sData, , , ScpUtf8

' ❌ 错误：编码不一致
m_oClient.SendData "中文测试", ScpUtf8  ' 使用 UTF-8
' 接收时
Client.GetData sData  ' 使用默认 ACP → 乱码
```

---

### ❓ 问题 2: UTF-8 字节长度错误

**现象**: 统计字节数时与预期不符

**原因**: UTF-8 是变长编码，中文字符占用 3 字节

**解决方案**:
```vb
' 计算实际字节数
Function GetByteCount(ByVal sText As String, ByVal eCodePage As EnumScpCodePage) As Long
    Dim oSocket As New cAsyncSocket
    Dim baData() As Byte
    
    baData = oSocket.ToTextArray(sText, eCodePage)
    GetByteCount = UBound(baData) + 1
End Function

' 使用
Dim lLen As Long
lLen = GetByteCount("中文测试", ScpUtf8)
Debug.Print lLen ' 12 (每个中文字符 3 字节)
```

---

### ❓ 问题 3: 与 Web 服务器通信

**现象**: Web 服务器返回的内容显示不正确

**原因**: Web 服务器通常使用 UTF-8，但客户端使用了默认编码

**解决方案**:
```vb
' 发送 HTTP 请求（使用 UTF-8）
m_oClient.SendData "GET / HTTP/1.1" & vbCrLf & "Host: example.com" & vbCrLf & vbCrLf, ScpUtf8

' 接收响应（使用 UTF-8）
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData, , , ScpUtf8
    
    ' 解析响应
    Debug.Print sData
End Sub
```

---

### ❓ 问题 4: 数据库编码冲突

**现象**: 从数据库读取的字符串通过网络传输后显示异常

**原因**: 数据库编码与网络传输编码不一致

**解决方案**:
```vb
' 从数据库读取（假设数据库使用 UTF-8）
Dim sData As String
sData = GetFromDatabase()

' 直接发送（数据库已经是 UTF-8）
' 不需要转换
m_oClient.SendData sData, ScpUtf8

' 或者转换为 ACP 再发送
Dim baUtf8() As Byte
Dim sAcp As String
' 先转 UTF-8 字节数组
baUtf8 = ConvertToUtf8Bytes(sData)
' 再转为 ACP 字符串
sAcp = ConvertFromUtf8Bytes(baUtf8)
m_oClient.SendData sAcp
```

---

### ❓ 问题 5: 文件传输编码

**现象**: 传输文本文件后内容乱码

**原因**: 文件编码与网络传输编码不一致

**解决方案**:
```vb
' 读取文本文件
Private Function ReadFile(ByVal sFilePath As String, ByVal eCodePage As EnumScpCodePage) As String
    Dim iFileNum As Integer
    iFileNum = FreeFile
    
    Open sFilePath For Binary As #iFileNum
    Dim baData() As Byte
    ReDim baData(0 To LOF(iFileNum) - 1) As Byte
    Get #iFileNum, , baData
    Close #iFileNum
    
    Dim oSocket As New cAsyncSocket
    ReadFile = oSocket.FromTextArray(baData, eCodePage)
End Function

' 发送文件
Private Sub SendFile(ByVal sFilePath As String)
    Dim sContent As String
    
    ' 假设文件是 UTF-8 编码
    sContent = ReadFile(sFilePath, ScpUtf8)
    
    ' 使用 UTF-8 发送
    m_oClient.SendData sContent, ScpUtf8
End Sub
```

---

## 编码转换工具函数

### 编码检测

```vb
' 简单的 UTF-8 检测
Function IsLikelyUtf8(ByVal sText As String) As Boolean
    ' 检查是否包含高字节字符
    Dim i As Long
    For i = 1 To Len(sText)
        If AscW(Mid$(sText, i, 1)) > 255 Then
            IsLikelyUtf8 = True
            Exit Function
        End If
    Next
    IsLikelyUtf8 = False
End Function
```

### 编码转换

```vb
' ACP 转换为 UTF-8
Function AcpToUtf8(ByVal sText As String) As String
    Dim oSocket As New cAsyncSocket
    Dim baAcp() As Byte
    Dim baUtf8() As Byte
    
    ' ACP → 字节数组
    baAcp = oSocket.ToTextArray(sText, ScpAcp)
    
    ' 字节数组 → UTF-8 字符串（这里需要额外处理）
    ' VB6 中需要使用 Win32 API 进行转换
    ' 这里简化演示
    
    AcpToUtf8 = sText ' 实际实现需要调用 MultiByteToWideChar
End Function
```

---

## 最佳实践

### ✅ 推荐做法

1. **统一编码**: 发送和接收使用相同编码
2. **明确指定**: 始终显式指定编码参数，不依赖默认值
3. **文档记录**: 记录每个连接使用的编码
4. **编码协商**: 在协议层面协商编码
5. **错误处理**: 处理编码转换错误

```vb
' 示例：封装的网络类
Public Sub SendText(ByVal oSocket As cWinsock, ByVal sText As String)
    On Error GoTo EH
    
    ' 使用配置的编码
    Select Case m_eEncoding
        Case eEncoding.UTF8
            oSocket.SendData sText, ScpUtf8
        Case eEncoding.ACP
            oSocket.SendData sText, ScpAcp
        Case eEncoding.Unicode
            oSocket.SendData sText, ScpUnicode
    End Select
    
    Exit Sub
    
EH:
    Debug.Print "发送数据失败: " & Err.Description
End Sub
```

### ❌ 避免的做法

1. **混合编码**: 同一连接使用不同编码
2. **依赖默认**: 不指定编码参数，依赖默认值
3. **忽略检测**: 不检测对方使用的编码
4. **盲目转换**: 不验证就直接转换编码

---

**最后更新**: 2026-01-09
