# mWebSocketUtils 工具模块参考

## 📋 模块概述

`mWebSocketUtils` 是 WebSocket 类库的公共工具模块，提供以下功能：

- UTF-8 编码/解码
- Base64 编码
- SHA1 哈希计算
- WebSocket Key 生成和验证
- HTTP 头解析
- 关闭码描述获取

---

## 📊 公共枚举

### WsOpCode - WebSocket 操作码

```vb
Public Enum WsOpCode
    WS_OPCODE_CONTINUATION = 0   ' 分片帧的后续帧
    WS_OPCODE_TEXT = 1           ' 文本数据帧
    WS_OPCODE_BINARY = 2         ' 二进制数据帧
    WS_OPCODE_CLOSE = 8           ' 关闭连接帧
    WS_OPCODE_PING = 9           ' Ping 帧
    WS_OPCODE_PONG = 10          ' Pong 帧
End Enum
```

---

### WsCloseCode - WebSocket 关闭状态码

```vb
Public Enum WsCloseCode
    WS_CLOSE_NORMAL = 1000          ' 正常关闭
    WS_CLOSE_GOING_AWAY = 1001      ' 端点离开
    WS_CLOSE_PROTOCOL_ERROR = 1002    ' 协议错误
    WS_CLOSE_UNSUPPORTED_DATA = 1003  ' 不支持的数据类型
    WS_CLOSE_NO_STATUS = 1005        ' 无状态（本地）
    WS_CLOSE_ABNORMAL = 1006         ' 异常关闭（本地）
    WS_CLOSE_INVALID_DATA = 1007     ' 无效数据
    WS_CLOSE_POLICY_VIOLATION = 1008 ' 策略违规
    WS_CLOSE_MESSAGE_TOO_BIG = 1009   ' 消息过大
    WS_CLOSE_MANDATORY_EXT = 1010    ' 必需的扩展
    WS_CLOSE_INTERNAL_ERROR = 1011    ' 内部错误
    WS_CLOSE_SERVICE_RESTART = 1012   ' 服务重启
    WS_CLOSE_TRY_AGAIN = 1013        ' 稍后重试
End Enum
```

---

### WsState - WebSocket 连接状态

```vb
Public Enum WsState
    WS_STATE_CLOSED = 0      ' 已关闭
    WS_STATE_CONNECTING = 1  ' 正在连接
    WS_STATE_OPEN = 2         ' 已打开
    WS_STATE_CLOSING = 3      ' 正在关闭
End Enum
```

---

## 🔧 UTF-8 函数

### StringToUTF8 - 字符串转 UTF-8

**语法**:

```vb
Public Function StringToUTF8(ByVal Text As String) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Text` | String | 要转换的字符串 |

**返回值**: UTF-8 编码的字节数组

**示例**:

```vb
' 简单转换
Dim baUTF8() As Byte
baUTF8 = StringToUTF8("Hello WebSocket!")

' 中文转换
baUTF8 = StringToUTF8("你好世界")

' 空字符串
baUTF8 = StringToUTF8("")
' 返回空数组

' 用于 WebSocket 发送
Dim sMessage As String
sMessage = "测试消息"
m_Client.SendText sMessage  ' 内部会调用 StringToUTF8
```

---

### UTF8ToString - UTF-8 转字符串

**语法**:

```vb
Public Function UTF8ToString(ByRef Utf8Data() As Byte) As String
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Utf8Data()` | Byte() | UTF-8 编码的字节数组 |

**返回值**: 解码后的字符串

**示例**:

```vb
' 解码 UTF-8 数据
Dim baData() As Byte
baData = LoadFile("utf8.txt")
Dim sText As String
sText = UTF8ToString(baData)
Debug.Print sText

' WebSocket 消息解码
Private Sub m_Client_OnTextMessage(ByVal Message As String)
    ' Message 已经是解码后的字符串
    Debug.Print Message
End Sub

' 手动解码
Dim sDecoded As String
sDecoded = UTF8ToString(baPayload)
```

---

## 📦 Base64 函数

### Base64Encode - Base64 编码

**语法**:

```vb
Public Function Base64Encode(ByRef Data() As Byte) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | 要编码的字节数组 |

**返回值**: Base64 编码的字符串

**说明**: 
- 使用标准 Base64 字符集（A-Z, a-z, 0-9, +, /）
- 使用 `=` 作为填充符
- 用于 WebSocket 握手的 Key 和 Accept 编码

**示例**:

```vb
' 编码随机数据
Dim baData(15) As Byte
Randomize Timer
For i = 0 To 15
    baData(i) = CByte(Int(Rnd * 256))
Next i

Dim sBase64 As String
sBase64 = Base64Encode(baData)
Debug.Print "Base64: " & sBase64

' 生成 WebSocket Key
Dim sKey As String
sKey = GenerateWebSocketKey()
Debug.Print "WebSocket Key: " & sKey
```

---

## 🔐 SHA1 函数

### SHA1Hash - 计算 SHA1 哈希

**语法**:

```vb
Public Function SHA1Hash(ByRef Data() As Byte) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | 要哈希的数据 |

**返回值**: 20 字节的 SHA1 哈希值

**说明**: 
- 使用 Windows CryptoAPI
- 返回原始 20 字节哈希值

**示例**:

```vb
' 计算 SHA1 哈希
Dim baData() As Byte
baData = StringToUTF8("Hello World")

Dim baHash() As Byte
baHash = SHA1Hash(baData)

Dim sHashHex As String
sHashHex = BytesToHex(baHash)
Debug.Print "SHA1: " & sHashHex

' 用于 WebSocket Accept Key 计算
Dim sCombined As String
sCombined = sClientKey & WS_MAGIC_GUID
Dim baCombined() As Byte
baCombined = StrConv(sCombined, vbFromUnicode)
Dim baHash() As Byte
baHash = SHA1Hash(baCombined)
```

---

## 🔑 WebSocket Key 函数

### GenerateWebSocketKey - 生成 WebSocket Key

**语法**:

```vb
Public Function GenerateWebSocketKey() As String
```

**返回值**: Base64 编码的随机 16 字节 Key

**说明**: 
- 用于客户端握手
- 生成随机 16 字节
- Base64 编码后返回

**示例**:

```vb
' 客户端自动调用（内部使用）
Dim sKey As String
sKey = GenerateWebSocketKey()
Debug.Print "WebSocket Key: " & sKey

' 输出示例: dGhlIHNhbXBsZSBub25jZQ==
```

---

### ComputeAcceptKey - 计算 Accept Key

**语法**:

```vb
Public Function ComputeAcceptKey(ByVal ClientKey As String) As String
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientKey` | String | 客户端发送的 WebSocket Key |

**返回值**: 服务端应返回的 Accept Key

**说明**: 
- 计算 `ClientKey + WS_MAGIC_GUID` 的 SHA1 哈希
- Base64 编码结果
- 用于服务端握手验证

**示例**:

```vb
' 服务端计算 Accept Key（内部使用）
Dim sClientKey As String
sClientKey = GetHeaderValue(sHandshake, "Sec-WebSocket-Key")

Dim sAcceptKey As String
sAcceptKey = ComputeAcceptKey(sClientKey)
Debug.Print "Accept Key: " & sAcceptKey

' 输出 HTTP 响应头
Dim sResponse As String
sResponse = "HTTP/1.1 101 Switching Protocols" & vbCrLf
sResponse = sResponse & "Sec-WebSocket-Accept: " & sAcceptKey & vbCrLf
sResponse = sResponse & vbCrLf
```

---

### WS_MAGIC_GUID - WebSocket 魔数

**常量**:

```vb
Public Const WS_MAGIC_GUID As String = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
```

**说明**: WebSocket 握手使用的固定 GUID，用于计算 Accept Key。

---

### WS_VERSION - WebSocket 协议版本

**常量**:

```vb
Public Const WS_VERSION As String = "13"
```

**说明**: WebSocket 协议版本号（RFC 6455）。

---

## 📋 HTTP 头解析函数

### GetHeaderValue - 获取 HTTP 头值

**语法**:

```vb
Public Function GetHeaderValue(ByVal HttpText As String, ByVal HeaderName As String) As String
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `HttpText` | String | HTTP 请求/响应文本 |
| `HeaderName` | String | 头名称（不区分大小写） |

**返回值**: 头值，未找到返回空字符串

**示例**:

```vb
' 获取 Host 头
Dim sRequest As String
sRequest = "GET /chat HTTP/1.1" & vbCrLf
sRequest = sRequest & "Host: example.com:8080" & vbCrLf
sRequest = sRequest & "Upgrade: websocket" & vbCrLf
sRequest = sRequest & vbCrLf

Dim sHost As String
sHost = GetHeaderValue(sRequest, "Host")
Debug.Print "Host: " & sHost  ' 输出: example.com:8080

' 获取 WebSocket Key
Dim sKey As String
sKey = GetHeaderValue(sRequest, "Sec-WebSocket-Key")

' 获取 Connection 头
Dim sConnection As String
sConnection = GetHeaderValue(sRequest, "Connection")
```

---

## 📝 关闭码描述函数

### GetCloseCodeDescription - 获取关闭码描述

**语法**:

```vb
Public Function GetCloseCodeDescription(ByVal Code As WsCloseCode) As String
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Code` | WsCloseCode | 关闭状态码 |

**返回值**: 关闭码的文本描述

**示例**:

```vb
Dim sDesc As String
sDesc = GetCloseCodeDescription(WS_CLOSE_NORMAL)
Debug.Print sDesc  ' 输出: Normal closure

sDesc = GetCloseCodeDescription(WS_CLOSE_ABNORMAL)
Debug.Print sDesc  ' 输出: Abnormal closure

sDesc = GetCloseCodeDescription(1000)
Debug.Print sDesc  ' 输出: Normal closure

sDesc = GetCloseCodeDescription(9999)
Debug.Print sDesc  ' 输出: Unknown (9999)

' 在 OnClose 事件中使用
Private Sub m_Client_OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
    Debug.Print "关闭码: " & Code
    Debug.Print "描述: " & GetCloseCodeDescription(Code)
    If LenB(Reason) > 0 Then
        Debug.Print "原因: " & Reason
    End If
End Sub
```

---

## 🧹 清理函数

### CleanupCryptoProvider - 清理 CryptoAPI 资源

**语法**:

```vb
Public Sub CleanupCryptoProvider()
```

**说明**: 
- 释放 CryptoAPI 上下文
- 通常在程序结束时调用

**示例**:

```vb
' 程序退出时清理
Private Sub Form_Unload(Cancel As Integer)
    CleanupCryptoProvider
End Sub
```

---

## 📝 完整使用示例

### WebSocket 握手流程

```vb
' 客户端发送握手请求
Private Sub SendHandshake()
    Dim sKey As String
    sKey = GenerateWebSocketKey()
    
    Dim sHandshake As String
    sHandshake = "GET /chat HTTP/1.1" & vbCrLf
    sHandshake = sHandshake & "Host: example.com:8080" & vbCrLf
    sHandshake = sHandshake & "Upgrade: websocket" & vbCrLf
    sHandshake = sHandshake & "Connection: Upgrade" & vbCrLf
    sHandshake = sHandshake & "Sec-WebSocket-Key: " & sKey & vbCrLf
    sHandshake = sHandshake & "Sec-WebSocket-Version: " & WS_VERSION & vbCrLf
    sHandshake = sHandshake & vbCrLf
    
    m_Socket.SendData sHandshake, ScpUtf8
End Sub

' 服务端处理握手请求
Private Function HandleHandshake(ByVal sRequest As String) As Boolean
    ' 获取客户端 Key
    Dim sKey As String
    sKey = GetHeaderValue(sRequest, "Sec-WebSocket-Key")
    If LenB(sKey) = 0 Then
        HandleHandshake = False
        Exit Function
    End If
    
    ' 计算 Accept Key
    Dim sAccept As String
    sAccept = ComputeAcceptKey(sKey)
    
    ' 发送响应
    Dim sResponse As String
    sResponse = "HTTP/1.1 101 Switching Protocols" & vbCrLf
    sResponse = sResponse & "Upgrade: websocket" & vbCrLf
    sResponse = sResponse & "Connection: Upgrade" & vbCrLf
    sResponse = sResponse & "Sec-WebSocket-Accept: " & sAccept & vbCrLf
    sResponse = sResponse & vbCrLf
    
    m_Socket.SendData sResponse, ScpUtf8
    
    HandleHandshake = True
End Function

' 客户端验证握手响应
Private Function ValidateResponse(ByVal sResponse As String) As Boolean
    ' 检查状态码
    If InStr(sResponse, "101") = 0 Then
        ValidateResponse = False
        Exit Function
    End If
    
    ' 获取 Accept Key
    Dim sAccept As String
    sAccept = GetHeaderValue(sResponse, "Sec-WebSocket-Accept")
    
    ' 计算预期值
    Dim sExpected As String
    sExpected = ComputeAcceptKey(m_sClientKey)
    
    ' 验证
    ValidateResponse = (sAccept = sExpected)
End Function
```

### 文本消息处理

```vb
' 发送文本消息
Public Sub SendTextMessage(ByVal sText As String)
    Dim baPayload() As Byte
    baPayload = StringToUTF8(sText)
    
    Dim oFrame As New cWebSocketFrame
    Dim baFrame() As Byte
    baFrame = oFrame.BuildFrame(baPayload, WS_OPCODE_TEXT, True, True)
    
    m_Socket.SendData baFrame
End Sub

' 接收文本消息
Private Sub ProcessTextFrame(ByVal baPayload() As Byte)
    Dim sText As String
    sText = UTF8ToString(baPayload)
    
    Debug.Print "收到文本: " & sText
    ' 处理消息...
End Sub
```

### 自定义数据处理

```vb
' 序列化对象为 UTF-8 JSON
Public Function SerializeJSON(ByVal oObject As Object) As Byte()
    Dim sJSON As String
    sJSON = ToJSONString(oObject)
    
    Dim baData() As Byte
    baData = StringToUTF8(sJSON)
    
    SerializeJSON = baData
End Function

' 反序列化 UTF-8 JSON 为对象
Public Function DeserializeJSON(ByVal baData() As Byte) As Object
    Dim sJSON As String
    sJSON = UTF8ToString(baData)
    
    Set DeserializeJSON = FromJSONString(sJSON)
End Function
```

---

## ⚠️ 注意事项

1. **CryptoAPI 初始化** - SHA1Hash 会自动初始化 CryptoAPI
2. **资源清理** - 程序结束时调用 CleanupCryptoProvider
3. **编码一致性** - 发送和接收都使用 UTF-8 编码
4. **头名称不区分大小写** - GetHeaderValue 会忽略大小写

---

**最后更新**: 2026-01-10
