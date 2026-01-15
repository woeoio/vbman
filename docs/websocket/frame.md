# cWebSocketFrame 类参考

## 📋 类概述

`cWebSocketFrame` 是 WebSocket 帧解析和构建类，用于处理 WebSocket 协议的帧格式。

### 设计特点

- **只读解析** - `ParseHeader()` 不会修改输入数据
- **分离 unmask** - 解析头部后，再单独提取并 unmask 负载
- **支持构建** - 可以构建各种类型的 WebSocket 帧
- **分片支持** - 正确处理分片帧（CONTINUATION）

---

## 📊 属性参考

### FIN - 最终帧标志

**类型**: `Boolean`  
**读写**: 只读（解析后）

**说明**: 标识这是否是消息的最后一帧。

```vb
If oFrame.FIN Then
    Debug.Print "这是最后一帧"
Else
    Debug.Print "还有后续帧"
End If
```

---

### RSV1, RSV2, RSV3 - 保留位

**类型**: `Boolean`  
**读写**: 只读（解析后）

**说明**: WebSocket 协议的保留位。通常为 `False`，用于扩展。

```vb
If oFrame.RSV1 Or oFrame.RSV2 Or oFrame.RSV3 Then
    Debug.Print "使用了扩展"
End If
```

---

### OpCode - 操作码

**类型**: `WsOpCode` (枚举)  
**读写**: 只读（解析后）

**值**:

| OpCode | 常量 | 说明 |
|--------|-------|------|
| 0 | WS_OPCODE_CONTINUATION | 分片帧的后续帧 |
| 1 | WS_OPCODE_TEXT | 文本数据帧 |
| 2 | WS_OPCODE_BINARY | 二进制数据帧 |
| 8 | WS_OPCODE_CLOSE | 关闭连接帧 |
| 9 | WS_OPCODE_PING | Ping 帧 |
| 10 | WS_OPCODE_PONG | Pong 帧 |

```vb
Select Case oFrame.OpCode
    Case WS_OPCODE_TEXT
        Debug.Print "文本帧"
    Case WS_OPCODE_BINARY
        Debug.Print "二进制帧"
    Case WS_OPCODE_CLOSE
        Debug.Print "关闭帧"
    Case WS_OPCODE_PING
        Debug.Print "Ping 帧"
    Case WS_OPCODE_PONG
        Debug.Print "Pong 帧"
End Select
```

---

### HasMask - 是否有掩码

**类型**: `Boolean`  
**读写**: 只读（解析后）

**说明**: 标识负载是否使用了掩码。客户端发送的帧必须有掩码，服务端发送的帧不应有掩码。

```vb
If oFrame.HasMask Then
    Debug.Print "帧已掩码"
Else
    Debug.Print "帧未掩码"
End If
```

---

### PayloadLength - 负载长度

**类型**: `Long`  
**读写**: 只读（解析后）

**说明**: 帧负载数据的长度（字节数）。

```vb
Debug.Print "负载长度: " & oFrame.PayloadLength & " 字节"
```

---

### HeaderLength - 头部长度

**类型**: `Long`  
**读写**: 只读（解析后）

**说明**: WebSocket 帧头的长度（包括扩展长度和掩码键）。

```vb
Debug.Print "头部长度: " & oFrame.HeaderLength & " 字节"
```

---

### TotalFrameLength - 总帧长度

**类型**: `Long`  
**读写**: 只读（解析后）

**说明**: 完整帧的长度（头部 + 负载）。

```vb
Debug.Print "总帧长度: " & oFrame.TotalFrameLength & " 字节"
```

---

### IsValid - 是否有效

**类型**: `Boolean`  
**读写**: 只读（解析后）

**说明**: 帧头是否成功解析且有效。

```vb
If oFrame.IsValid Then
    Debug.Print "帧有效"
Else
    Debug.Print "帧无效: " & oFrame.ErrorMessage
End If
```

---

### ErrorMessage - 错误消息

**类型**: `String`  
**读写**: 只读（解析后）

**说明**: 当帧无效时，包含错误描述。

```vb
If Not oFrame.IsValid Then
    Debug.Print "错误: " & oFrame.ErrorMessage
End If
```

---

### IsControlFrame - 是否为控制帧

**类型**: `Boolean`  
**读写**: 只读

**说明**: 判断当前帧是否为控制帧（CLOSE、PING、PONG）。

```vb
If oFrame.IsControlFrame Then
    Debug.Print "这是控制帧"
    ' 控制帧不能分片
    If Not oFrame.FIN Then
        Debug.Print "警告：控制帧不应分片"
    End If
End If
```

---

### IsDataFrame - 是否为数据帧

**类型**: `Boolean`  
**读写**: 只读

**说明**: 判断当前帧是否为数据帧（TEXT、BINARY、CONTINUATION）。

```vb
If oFrame.IsDataFrame Then
    Debug.Print "这是数据帧"
    ' 可以分片
End If
```

---

## 🚀 方法参考

### ParseHeader - 解析帧头

**语法**:

```vb
Public Function ParseHeader(ByRef Buffer As cByteBuffer) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Buffer` | cByteBuffer | 字节缓冲区 |

**返回值**: `Boolean` - 解析成功返回 `True`，否则返回 `False`

**说明**: 
- **只读操作** - 不会修改或消费缓冲区数据
- 需要至少 2 字节数据才能解析
- 解析后可通过属性访问帧头信息

**示例**:

```vb
Dim oFrame As New cWebSocketFrame
Dim oBuffer As cByteBuffer

Set oBuffer = New cByteBuffer
oBuffer.Append baReceivedData

' 解析头部
If oFrame.ParseHeader(oBuffer) Then
    Debug.Print "帧类型: " & oFrame.OpCode
    Debug.Print "负载长度: " & oFrame.PayloadLength
    
    ' 检查是否完整
    If oFrame.IsCompleteFrame(oBuffer) Then
        ' 提取负载
        Dim baPayload() As Byte
        baPayload = oFrame.ExtractPayload(oBuffer)
    End If
Else
    Debug.Print "解析失败: " & oFrame.ErrorMessage
End If
```

---

### IsCompleteFrame - 检查帧是否完整

**语法**:

```vb
Public Function IsCompleteFrame(ByRef Buffer As cByteBuffer) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Buffer` | cByteBuffer | 字节缓冲区 |

**返回值**: `Boolean` - 帧完整返回 `True`，否则返回 `False`

**说明**: 
- 必须先调用 `ParseHeader()` 成功
- 检查缓冲区是否包含完整的帧

**示例**:

```vb
If oFrame.ParseHeader(oBuffer) Then
    ' 检查是否完整
    If oFrame.IsCompleteFrame(oBuffer) Then
        ' 可以提取
        baPayload = oFrame.ExtractPayload(oBuffer)
    Else
        Debug.Print "需要更多数据"
    End If
End If
```

---

### ExtractPayload - 提取并 unmask 负载

**语法**:

```vb
Public Function ExtractPayload(ByRef Buffer As cByteBuffer) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Buffer` | cByteBuffer | 字节缓冲区 |

**返回值**: `Byte()` - 提取的负载数据（已 unmask）

**说明**: 
- 必须先调用 `ParseHeader()` 成功
- 会**消费**整个帧（从缓冲区移除）
- 自动执行 unmask 操作（如果帧有掩码）

**示例**:

```vb
' 完整的帧处理流程
If oFrame.ParseHeader(oBuffer) Then
    If oFrame.IsCompleteFrame(oBuffer) Then
        ' 提取负载（消费帧）
        Dim baPayload() As Byte
        baPayload = oFrame.ExtractPayload(oBuffer)
        
        ' 处理负载
        ProcessPayload baPayload, oFrame.OpCode
    End If
End If
```

---

### SkipFrame - 跳过帧

**语法**:

```vb
Public Sub SkipFrame(ByRef Buffer As cByteBuffer)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Buffer` | cByteBuffer | 字节缓冲区 |

**说明**: 从缓冲区中移除整个帧，但不提取负载。用于处理不需要的帧。

**示例**:

```vb
' 跳过控制帧
If oFrame.ParseHeader(oBuffer) Then
    If oFrame.IsControlFrame Then
        ' 跳过控制帧
        oFrame.SkipFrame oBuffer
    Else
        ' 处理数据帧
        baPayload = oFrame.ExtractPayload(oBuffer)
    End If
End If
```

---

### BuildFrame - 构建帧

**语法**:

```vb
Public Function BuildFrame(ByRef Payload() As Byte, _
                         ByVal OpCode As WsOpCode, _
                         Optional ByVal UseMask As Boolean = False, _
                         Optional ByVal IsFinal As Boolean = True) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Payload` | Byte() | 负载数据 |
| `OpCode` | WsOpCode | 操作码 |
| `UseMask` | Boolean | 是否掩码 |
| `IsFinal` | Boolean | 是否为最终帧 |

**返回值**: `Byte()` - 完整的 WebSocket 帧字节数组

**说明**: 
- 客户端发送必须掩码 (`UseMask = True`)
- 服务端发送不应掩码 (`UseMask = False`)
- 掩码时自动生成随机掩码键

**示例**:

```vb
' 客户端发送（必须掩码）
Dim baFrame() As Byte
Dim baData() As Byte
baData = StringToUTF8("Hello")

baFrame = oFrame.BuildFrame(baData, WS_OPCODE_TEXT, True, True)
Socket.SendData baFrame

' 服务端发送（不掩码）
baFrame = oFrame.BuildFrame(baData, WS_OPCODE_TEXT, False, True)
Socket.SendData baFrame
```

---

### BuildTextFrame - 构建文本帧

**语法**:

```vb
Public Function BuildTextFrame(ByVal Text As String, _
                              Optional ByVal UseMask As Boolean = False) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Text` | String | 文本内容 |
| `UseMask` | Boolean | 是否掩码 |

**返回值**: `Byte()` - WebSocket 文本帧

**说明**: 文本会自动转换为 UTF-8 编码。

**示例**:

```vb
' 客户端发送文本
Dim baFrame() As Byte
baFrame = oFrame.BuildTextFrame("Hello WebSocket!", True)
Socket.SendData baFrame
```

---

### BuildBinaryFrame - 构建二进制帧

**语法**:

```vb
Public Function BuildBinaryFrame(ByRef Data() As Byte, _
                                Optional ByVal UseMask As Boolean = False) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | 二进制数据 |
| `UseMask` | Boolean | 是否掩码 |

**返回值**: `Byte()` - WebSocket 二进制帧

**示例**:

```vb
' 客户端发送二进制
Dim baData() As Byte
baData = LoadFile("image.png")

Dim baFrame() As Byte
baFrame = oFrame.BuildBinaryFrame(baData, True)
Socket.SendData baFrame
```

---

### BuildCloseFrame - 构建关闭帧

**语法**:

```vb
Public Function BuildCloseFrame(Optional ByVal StatusCode As WsCloseCode = WS_CLOSE_NORMAL, _
                               Optional ByVal Reason As String = "", _
                               Optional ByVal UseMask As Boolean = False) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `StatusCode` | WsCloseCode（可选） | 关闭状态码 |
| `Reason` | String（可选） | 关闭原因 |
| `UseMask` | Boolean（可选） | 是否掩码 |

**返回值**: `Byte()` - WebSocket 关闭帧

**示例**:

```vb
' 正常关闭
Dim baFrame() As Byte
baFrame = oFrame.BuildCloseFrame(WS_CLOSE_NORMAL, "正常关闭", True)
Socket.SendData baFrame

' 协议错误
baFrame = oFrame.BuildCloseFrame(WS_CLOSE_PROTOCOL_ERROR, "无效帧", True)
Socket.SendData baFrame
```

---

### BuildPingFrame - 构建 Ping 帧

**语法**:

```vb
Public Function BuildPingFrame(ByRef Payload() As Byte, _
                              Optional ByVal UseMask As Boolean = False) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Payload()` | Byte() | Ping 负载 |
| `UseMask` | Boolean（可选） | 是否掩码 |

**返回值**: `Byte()` - WebSocket Ping 帧

**示例**:

```vb
' 发送空 Ping
Dim baEmpty() As Byte
Dim baFrame() As Byte
baFrame = oFrame.BuildPingFrame(baEmpty, True)
Socket.SendData baFrame

' 发送带数据的 Ping
Dim baData() As Byte
baData = StringToUTF8("ping")
baFrame = oFrame.BuildPingFrame(baData, True)
Socket.SendData baFrame
```

---

### BuildPongFrame - 构建 Pong 帧

**语法**:

```vb
Public Function BuildPongFrame(ByRef Payload() As Byte, _
                              Optional ByVal UseMask As Boolean = False) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Payload()` | Byte() | Pong 负载 |
| `UseMask` | Boolean（可选） | 是否掩码 |

**返回值**: `Byte()` - WebSocket Pong 帧

**示例**:

```vb
' 回复 Pong（使用 Ping 的负载）
Dim baFrame() As Byte
baFrame = oFrame.BuildPongFrame(baPingPayload, True)
Socket.SendData baFrame
```

---

## 📝 使用示例

### 基本帧解析流程

```vb
Private Sub ProcessWebSocketFrame(oBuffer As cByteBuffer)
    Dim oFrame As New cWebSocketFrame
    
    Do While oBuffer.Size >= 2
        ' 1. 解析头部（只读）
        If Not oFrame.ParseHeader(oBuffer) Then
            Debug.Print "需要更多数据"
            Exit Do
        End If
        
        ' 2. 检查完整性
        If Not oFrame.IsCompleteFrame(oBuffer) Then
            Debug.Print "需要更多数据"
            Exit Do
        End If
        
        ' 3. 提取负载（消费帧）
        Dim baPayload() As Byte
        baPayload = oFrame.ExtractPayload(oBuffer)
        
        ' 4. 处理帧
        Select Case oFrame.OpCode
            Case WS_OPCODE_TEXT
                Dim sText As String
                sText = UTF8ToString(baPayload)
                Debug.Print "文本: " & sText
                
            Case WS_OPCODE_BINARY
                Debug.Print "二进制: " & (UBound(baPayload) + 1) & " 字节"
                
            Case WS_OPCODE_CLOSE
                Debug.Print "关闭帧"
                ProcessCloseFrame baPayload
                
            Case WS_OPCODE_PING
                Debug.Print "Ping 帧"
                ' 自动回复 Pong
                Dim baPong() As Byte
                baPong = oFrame.BuildPongFrame(baPayload, False)
                SendData baPong
                
            Case WS_OPCODE_PONG
                Debug.Print "Pong 帧"
        End Select
    Loop
End Sub
```

### 构建并发送帧

```vb
' 客户端发送文本
Private Sub SendTextMessage(ByVal sText As String)
    Dim oFrame As New cWebSocketFrame
    Dim baFrame() As Byte
    
    ' 构建文本帧（必须掩码）
    baFrame = oFrame.BuildTextFrame(sText, True)
    
    ' 发送
    m_Socket.SendData baFrame
End Sub

' 客户端发送二进制
Private Sub SendBinaryMessage(ByVal baData() As Byte)
    Dim oFrame As New cWebSocketFrame
    Dim baFrame() As Byte
    
    ' 构建二进制帧（必须掩码）
    baFrame = oFrame.BuildBinaryFrame(baData, True)
    
    ' 发送
    m_Socket.SendData baFrame
End Sub

' 发送关闭帧
Private Sub SendCloseFrame()
    Dim oFrame As New cWebSocketFrame
    Dim baFrame() As Byte
    
    ' 构建关闭帧（必须掩码）
    baFrame = oFrame.BuildCloseFrame(WS_CLOSE_NORMAL, "正常关闭", True)
    
    ' 发送
    m_Socket.SendData baFrame
End Sub
```

---

**最后更新**: 2026-01-10
