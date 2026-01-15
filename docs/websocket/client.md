# cWebSocketClient 类参考

## 📋 类概述

`cWebSocketClient` 是 WebSocket 客户端实现类，提供连接到 WebSocket 服务器、发送/接收消息、自动握手等功能。

---

## 📡 事件列表

| 事件名 | 触发时机 | 参数 |
|--------|----------|------|
| `OnOpen` | 连接成功建立 | 无 |
| `OnClose` | 连接已关闭 | `Code` (关闭码), `Reason` (关闭原因) |
| `OnTextMessage` | 收到文本消息 | `Message` (消息内容) |
| `OnBinaryMessage` | 收到二进制消息 | `Data()` (字节数组) |
| `OnError` | 发生错误 | `Description` (错误描述) |
| `OnPong` | 收到 Pong 响应 | `Data()` (Pong 负载) |

---

## 🔧 属性参考

### State - 连接状态

**类型**: `WsState` (枚举)  
**读写**: 只读

**值**:

| 常量 | 值 | 说明 |
|------|-----|------|
| `WS_STATE_CLOSED` | 0 | 已关闭 |
| `WS_STATE_CONNECTING` | 1 | 正在连接/握手 |
| `WS_STATE_OPEN` | 2 | 已连接，可以发送消息 |
| `WS_STATE_CLOSING` | 3 | 正在关闭 |

**示例**:

```vb
If m_Client.State = WS_STATE_OPEN Then
    m_Client.SendText "Hello"
Else
    Debug.Print "未连接"
End If
```

---

### URL - 连接地址

**类型**: `String`  
**读写**: 只读

**说明**: 当前或最后连接的 WebSocket URL。

**示例**:

```vb
Debug.Print "连接到: " & m_Client.URL
' 输出: 连接到: ws://127.0.0.1:8080
```

---

### Host - 服务器主机名

**类型**: `String`  
**读写**: 只读

**说明**: 从 URL 解析的服务器主机名或 IP 地址。

**示例**:

```vb
Debug.Print "服务器: " & m_Client.Host
```

---

### Port - 服务器端口

**类型**: `Long`  
**读写**: 只读

**说明**: 从 URL 解析的服务器端口号。

**示例**:

```vb
Debug.Print "端口: " & m_Client.Port
```

---

### AutoPing - 自动 Ping

**类型**: `Boolean`  
**读写**: 读写

**说明**: 是否启用自动 Ping（保活功能）。默认为 `False`。

**示例**:

```vb
' 启用自动 Ping
m_Client.AutoPing = True

' 禁用自动 Ping
m_Client.AutoPing = False
```

---

### PingInterval - Ping 间隔

**类型**: `Long`  
**读写**: 读写

**说明**: 自动 Ping 的间隔时间（毫秒）。默认为 30000 (30 秒)。

**示例**:

```vb
' 设置 Ping 间隔为 20 秒
m_Client.PingInterval = 20000
```

---

## 🚀 方法参考

### Connect - 连接服务器

**语法**:

```vb
Public Sub Connect(ByVal WebSocketURL As String, Optional ByVal SubProtocol As String = "")
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `WebSocketURL` | String | WebSocket 服务器 URL，格式：`ws[s]://host[:port][/path][?query]` |
| `SubProtocol` | String（可选） | WebSocket 子协议 |

**URL 格式**:

- 标准格式: `ws://example.com:8080/chat`
- 默认端口: `ws://example.com/chat` (默认 80)
- 查询参数: `ws://example.com/chat?token=abc123`
- 安全连接: `wss://example.com` (暂未实现)

**示例**:

```vb
' 基本连接
m_Client.Connect "ws://127.0.0.1:8080"

' 带路径的连接
m_Client.Connect "ws://example.com/chat"

' 带查询参数的连接
m_Client.Connect "ws://example.com/chat?token=abc123"

' 指定子协议
m_Client.Connect "ws://example.com/chat", "chat.v1"
```

**错误处理**:

```vb
Private Sub cmdConnect_Click()
    On Error GoTo EH

    m_Client.Connect "ws://example.com:8080"
    Exit Sub

EH:
    Debug.Print "连接失败: " & Err.Description
End Sub
```

---

### CloseConnection - 关闭连接

**语法**:

```vb
Public Sub CloseConnection(Optional ByVal Code As WsCloseCode = WS_CLOSE_NORMAL, _
                          Optional ByVal Reason As String = "")
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Code` | WsCloseCode（可选） | 关闭状态码，默认 `WS_CLOSE_NORMAL` |
| `Reason` | String（可选） | 关闭原因 |

**常用关闭码**:

```vb
m_Client.CloseConnection WS_CLOSE_NORMAL, "正常关闭"
m_Client.CloseConnection WS_CLOSE_GOING_AWAY, "用户退出"
m_Client.CloseConnection WS_CLOSE_PROTOCOL_ERROR, "协议错误"
```

**示例**:

```vb
' 正常关闭
m_Client.CloseConnection

' 指定关闭码和原因
m_Client.CloseConnection WS_CLOSE_GOING_AWAY, "用户注销"

' 窗体关闭时自动断开
Private Sub Form_Unload(Cancel As Integer)
    m_Client.CloseConnection WS_CLOSE_GOING_AWAY, "应用关闭"
End Sub
```

---

### SendText - 发送文本消息

**语法**:

```vb
Public Sub SendText(ByVal Message As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Message` | String | 要发送的文本消息 |

**说明**: 消息会自动转换为 UTF-8 编码并添加 WebSocket 帧头。客户端发送的帧会自动掩码。

**示例**:

```vb
' 发送简单文本
m_Client.SendText "Hello WebSocket!"

' 发送 JSON 数据
Dim sJSON As String
sJSON = "{""type"":""message"", ""content"":""Hello""}"
m_Client.SendText sJSON

' 发送多行文本
m_Client.SendText "第一行" & vbCrLf & "第二行"
```

**错误处理**:

```vb
On Error GoTo EH
m_Client.SendText "Hello"
Exit Sub
EH:
Debug.Print "发送失败: " & Err.Description
```

---

### SendBinary - 发送二进制消息

**语法**:

```vb
Public Sub SendBinary(Data() As Byte)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | 要发送的二进制数据 |

**示例**:

```vb
' 发送字节数组
Dim baData() As Byte
baData = StringToBytes("Hello")
m_Client.SendBinary baData

' 发送图片数据
Dim baImage() As Byte
baImage = LoadImageToByteArray()
m_Client.SendBinary baImage

' 发送序列化对象
Dim baObj() As Byte
baObj = SerializeObject(myObject)
m_Client.SendBinary baObj
```

---

### SendPing - 发送 Ping 帧

**语法**:

```vb
Public Sub SendPing(Optional Payload As Variant)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Payload` | Variant（可选） | Ping 负载，可以是字符串或字节数组 |

**说明**: 用于连接保活，服务器会自动回复 Pong 帧。

**示例**:

```vb
' 发送空 Ping（保活）
m_Client.SendPing

' 发送带数据的 Ping
m_Client.SendPing "ping"

' 发送带二进制数据的 Ping
Dim baData() As Byte
baData = StringToUTF8("ping")
m_Client.SendPing baData

' 测量延迟
Dim lStartTime As Long
lStartTime = GetTickCount()
m_Client.SendPing "ping"

Private Sub m_Client_OnPong(Data() As Byte)
    Dim lElapsed As Long
    lElapsed = GetTickCount() - lStartTime
    Debug.Print "延迟: " & lElapsed & " ms"
End Sub
```

---

## 📡 事件详解

### OnOpen - 连接成功

**语法**:

```vb
Event OnOpen()
```

**说明**: WebSocket 握手成功后触发，此时可以开始发送消息。

**示例**:

```vb
Private Sub m_Client_OnOpen()
    Debug.Print "已成功连接到 WebSocket 服务器"

    ' 发送欢迎消息
    m_Client.SendText "Hello Server!"

    ' 更新 UI
    lblStatus.Caption = "已连接"
    cmdSend.Enabled = True
End Sub
```

---

### OnClose - 连接关闭

**语法**:

```vb
Event OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Code` | WsCloseCode | 关闭状态码 |
| `Reason` | String | 关闭原因 |

**示例**:

```vb
Private Sub m_Client_OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
    Debug.Print "连接已关闭"
    Debug.Print "状态码: " & Code
    Debug.Print "原因: " & Reason

    ' 根据关闭码处理
    Select Case Code
        Case WS_CLOSE_NORMAL
            Debug.Print "正常关闭"
        Case WS_CLOSE_ABNORMAL
            Debug.Print "异常关闭"
            ' 尝试重连
            If m_bAutoReconnect Then
                tmrReconnect.Enabled = True
            End If
        Case WS_CLOSE_GOING_AWAY
            Debug.Print "服务器关闭"
        Case Else
            Debug.Print "其他原因: " & Reason
    End Select

    ' 更新 UI
    lblStatus.Caption = "已断开"
    cmdSend.Enabled = False
End Sub
```

---

### OnTextMessage - 收到文本消息

**语法**:

```vb
Event OnTextMessage(ByVal Message As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Message` | String | 收到的文本消息（UTF-8 解码后） |

**示例**:

```vb
Private Sub m_Client_OnTextMessage(ByVal Message As String)
    Debug.Print "收到消息: " & Message

    ' 处理 JSON 消息
    If Left$(Message, 1) = "{" Then
        Dim sType As String
        sType = GetJSONField(Message, "type")
        
        Select Case sType
            Case "chat"
                DisplayChatMessage Message
            Case "notification"
                DisplayNotification Message
        End Select
    Else
        ' 简单文本消息
        txtMessages.Text = txtMessages.Text & Message & vbCrLf
    End If
End Sub
```

---

### OnBinaryMessage - 收到二进制消息

**语法**:

```vb
Event OnBinaryMessage(Data() As Byte)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | 收到的二进制数据 |

**示例**:

```vb
Private Sub m_Client_OnBinaryMessage(Data() As Byte)
    Debug.Print "收到二进制消息: " & (UBound(Data) + 1) & " 字节"

    ' 检查数据类型（假设前 4 字节是类型标识）
    If UBound(Data) >= 3 Then
        Dim lType As Long
        lType = CLng(Data(0)) * 256& ^ 3 + CLng(Data(1)) * 256& ^ 2 + _
                CLng(Data(2)) * 256& + CLng(Data(3))

        Select Case lType
            Case 1 ' 文本
                Dim sText As String
                sText = UTF8ToString(ExtractData(Data, 4))
                Debug.Print "文本数据: " & sText

            Case 2 ' 图片
                DisplayImage ExtractData(Data, 4)

            Case 3 ' 自定义
                ProcessCustomData ExtractData(Data, 4)
        End Select
    End If
End Sub
```

---

### OnError - 发生错误

**语法**:

```vb
Event OnError(ByVal Description As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Description` | String | 错误描述 |

**示例**:

```vb
Private Sub m_Client_OnError(ByVal Description As String)
    Debug.Print "错误: " & Description

    ' 显示错误提示
    MsgBox "发生错误: " & Description, vbExclamation

    ' 记录错误日志
    LogError Description

    ' 更新 UI
    lblStatus.Caption = "错误"
End Sub
```

---

### OnPong - 收到 Pong 响应

**语法**:

```vb
Event OnPong(Data() As Byte)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | Pong 负载（对应 Ping 的数据） |

**说明**: 通常用于测量网络延迟或确认连接活性。

**示例**:

```vb
Private m_lPingTimes() As Long
Private m_lPingIndex As Long

Private Sub SendPingForLatency()
    ReDim m_lPingTimes(1) As Long
    m_lPingIndex = 0
    m_lPingTimes(0) = GetTickCount()
    m_Client.SendPing "ping"
End Sub

Private Sub m_Client_OnPong(Data() As Byte)
    m_lPingTimes(1) = GetTickCount()
    Dim lLatency As Long
    lLatency = m_lPingTimes(1) - m_lPingTimes(0)
    Debug.Print "网络延迟: " & lLatency & " ms"

    ' 更新 UI
    lblLatency.Caption = lLatency & " ms"
End Sub
```

---

## 📝 完整示例

### 基本客户端示例

```vb
Private WithEvents m_Client As cWebSocketClient

Private Sub Form_Load()
    Set m_Client = New cWebSocketClient
End Sub

Private Sub cmdConnect_Click()
    m_Client.Connect "ws://127.0.0.1:8080"
End Sub

Private Sub cmdSend_Click()
    m_Client.SendText txtMessage.Text
End Sub

Private Sub cmdDisconnect_Click()
    m_Client.CloseConnection
End Sub

Private Sub m_Client_OnOpen()
    Debug.Print "已连接"
    cmdSend.Enabled = True
End Sub

Private Sub m_Client_OnTextMessage(ByVal Message As String)
    txtLog.Text = txtLog.Text & Message & vbCrLf
End Sub

Private Sub m_Client_OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
    Debug.Print "已断开: " & Reason
    cmdSend.Enabled = False
End Sub
```

### 带重连的客户端

```vb
Private WithEvents m_Client As cWebSocketClient
Private m_bAutoReconnect As Boolean
Private m_sServerURL As String

Private Sub Form_Load()
    Set m_Client = New cWebSocketClient
    m_bAutoReconnect = True
    m_sServerURL = "ws://127.0.0.1:8080"
    ConnectToServer
End Sub

Private Sub ConnectToServer()
    If m_Client.State = WS_STATE_CLOSED Then
        Debug.Print "正在连接..."
        m_Client.Connect m_sServerURL
    End If
End Sub

Private Sub m_Client_OnOpen()
    Debug.Print "已连接"
    tmrReconnect.Enabled = False
End Sub

Private Sub m_Client_OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
    Debug.Print "连接关闭: " & Reason
    
    If m_bAutoReconnect Then
        Debug.Print "3 秒后重连..."
        tmrReconnect.Interval = 3000
        tmrReconnect.Enabled = True
    End If
End Sub

Private Sub tmrReconnect_Timer()
    tmrReconnect.Enabled = False
    ConnectToServer
End Sub
```

---

**最后更新**: 2026-01-10
