# 进阶应用

本指南介绍 WebSocket 类库的高级功能和最佳实践。

---

## 📡 消息分片传输

### 概念

WebSocket 允许将大消息分成多个帧发送：

```
第一帧: FIN=0, OpCode=TEXT/BINARY  (消息开始)
后续帧: FIN=0, OpCode=CONTINUATION
最后一帧: FIN=1, OpCode=CONTINUATION
```

### 服务端处理分片

```vb
Private Sub ProcessDataFrame(ByVal Client As cWebSocketServerClient, _
                            ByRef Payload() As Byte, _
                            ByVal OpCode As WsOpCode, _
                            ByVal IsFinal As Boolean)
    If IsFinal And Not Client.IsFragmented Then
        ' 完整的单帧消息
        DeliverClientMessage Client, Payload, OpCode
    Else
        ' 开始或继续分片消息
        Client.IsFragmented = True
        Client.FragmentOpCode = OpCode
        Client.FragmentBuffer.Clear
        
        On Error Resume Next
        If UBound(Payload) >= 0 Then
            Client.FragmentBuffer.Append Payload
        End If
        On Error GoTo 0
        
        If IsFinal Then
            DeliverFragmentedMessage Client
        End If
    End If
End Sub

Private Sub DeliverFragmentedMessage(ByVal Client As cWebSocketServerClient)
    Dim baData() As Byte
    baData = Client.GetFragmentedData
    
    DeliverClientMessage Client, baData, Client.FragmentOpCode
    
    Client.ClearFragmentBuffer
End Sub
```

### 客户端发送分片

```vb
' 分片发送大消息
Public Sub SendLargeMessage(ByVal sMessage As String)
    Dim baPayload() As Byte
    baPayload = StringToUTF8(sMessage)
    
    Dim lChunkSize As Long
    lChunkSize = 4096  ' 每帧 4KB
    
    Dim lTotal As Long
    lTotal = UBound(baPayload) + 1
    
    Dim oFrame As New cWebSocketFrame
    Dim i As Long
    Dim lOffset As Long
    
    Do While lOffset < lTotal
        Dim lSize As Long
        lSize = lChunkSize
        If lOffset + lSize > lTotal Then
            lSize = lTotal - lOffset
        End If
        
        Dim baChunk() As Byte
        ReDim baChunk(lSize - 1) As Byte
        CopyMemory baChunk(0), baPayload(lOffset), lSize
        
        Dim baFrame() As Byte
        Dim bIsFinal As Boolean
        bIsFinal = (lOffset + lSize >= lTotal)
        
        If lOffset = 0 Then
            ' 第一帧
            baFrame = oFrame.BuildFrame(baChunk, WS_OPCODE_TEXT, True, bIsFinal)
        Else
            ' 后续帧
            baFrame = oFrame.BuildFrame(baChunk, WS_OPCODE_CONTINUATION, True, bIsFinal)
        End If
        
        m_Socket.SendData baFrame
        lOffset = lOffset + lSize
    Loop
End Sub
```

---

## 🔄 自动重连机制

### 客户端自动重连

```vb
Option Explicit

Private WithEvents m_Client As cWebSocketClient
Private WithEvents tmrReconnect As Timer
Private m_bAutoReconnect As Boolean
Private m_sServerURL As String
Private m_lMaxRetries As Long
Private m_lRetryCount As Long

Private Sub Form_Load()
    Set m_Client = New cWebSocketClient
    Set tmrReconnect = New Timer
    
    m_bAutoReconnect = True
    m_sServerURL = "ws://127.0.0.1:8080"
    m_lMaxRetries = 5
    m_lRetryCount = 0
    tmrReconnect.Interval = 5000  ' 5 秒
    
    ConnectToServer
End Sub

Private Sub ConnectToServer()
    If m_Client.State = WS_STATE_CLOSED Then
        Debug.Print "正在连接... (" & (m_lRetryCount + 1) & "/" & m_lMaxRetries & ")"
        On Error Resume Next
        m_Client.Connect m_sServerURL
        On Error GoTo 0
    End If
End Sub

Private Sub m_Client_OnOpen()
    Debug.Print "已连接"
    m_lRetryCount = 0
    tmrReconnect.Enabled = False
End Sub

Private Sub m_Client_OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
    Debug.Print "连接关闭: " & Reason
    
    If m_bAutoReconnect And Code <> WS_CLOSE_NORMAL Then
        m_lRetryCount = m_lRetryCount + 1
        
        If m_lRetryCount < m_lMaxRetries Then
            Debug.Print "5 秒后重连..."
            tmrReconnect.Enabled = True
        Else
            Debug.Print "已达到最大重试次数"
            MsgBox "无法连接到服务器，请稍后重试", vbExclamation
        End If
    End If
End Sub

Private Sub tmrReconnect_Timer()
    tmrReconnect.Enabled = False
    ConnectToServer
End Sub
```

---

## 📡 心跳保活

### 客户端自动 Ping

```vb
Option Explicit

Private WithEvents m_Client As cWebSocketClient
Private WithEvents tmrPing As Timer
Private m_bAutoPing As Boolean
Private m_lPingInterval As Long

Private Sub Form_Load()
    Set m_Client = New cWebSocketClient
    Set tmrPing = New Timer
    
    m_bAutoPing = True
    m_lPingInterval = 30000  ' 30 秒
    tmrPing.Interval = m_lPingInterval
End Sub

Private Sub m_Client_OnOpen()
    If m_bAutoPing Then
        tmrPing.Enabled = True
        Debug.Print "心跳已启用，间隔: " & m_lPingInterval & " ms"
    End If
End Sub

Private Sub tmrPing_Timer()
    If m_Client.State = WS_STATE_OPEN Then
        ' 发送 Ping（带时间戳用于测量延迟）
        Dim lTimestamp As Long
        lTimestamp = GetTickCount()
        
        Dim baData(3) As Byte
        baData(0) = (lTimestamp And &HFF000000) \ &H1000000
        baData(1) = (lTimestamp And &HFF0000) \ &H10000
        baData(2) = (lTimestamp And &HFF00&) \ &H100&
        baData(3) = lTimestamp And &HFF&
        
        m_Client.SendPing baData
        Debug.Print "Ping 已发送"
    End If
End Sub

Private Sub m_Client_OnPong(Data() As Byte)
    If UBound(Data) >= 3 Then
        Dim lSendTime As Long
        lSendTime = CLng(Data(0)) * 256& ^ 3 + CLng(Data(1)) * 256& ^ 2 + _
                  CLng(Data(2)) * 256& + CLng(Data(3))
        
        Dim lLatency As Long
        lLatency = GetTickCount() - lSendTime
        Debug.Print "Pong 收到，延迟: " & lLatency & " ms"
    End If
End Sub
```

---

## 🔐 认证与授权

### 客户端 Token 认证

```vb
Public Sub ConnectWithToken(ByVal ServerURL As String, ByVal Token As String)
    ' 在 URL 中添加 Token
    Dim sURL As String
    sURL = ServerURL & "?token=" & Token
    
    m_Client.Connect sURL
End Sub

' 或者通过握手后发送
Private Sub m_Client_OnOpen()
    ' 发送认证信息
    Dim sAuth As String
    sAuth = "{""type"":""auth"", ""token"":""abc123""}"
    
    m_Client.SendText sAuth
End Sub
```

### 服务端认证验证

```vb
Private Sub m_Server_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)
    ' 获取 Token（假设在 URL 查询参数中）
    ' 注意：需要扩展握手逻辑以获取查询参数
    
    Dim sToken As String
    sToken = GetTokenFromHandshake(m_Server, ClientID)
    
    If Not ValidateToken(sToken) Then
        Debug.Print "认证失败: " & ClientID
        m_Server.DisconnectClient ClientID, WS_CLOSE_POLICY_VIOLATION, "无效的 Token"
        Exit Sub
    End If
    
    ' 认证成功
    Debug.Print "认证成功: " & ClientID
End Sub

Private Function ValidateToken(ByVal Token As String) As Boolean
    ' 验证 Token（示例）
    If LenB(Token) = 0 Then
        ValidateToken = False
        Exit Function
    End If
    
    ' 检查数据库或配置
    ' ...
    
    ValidateToken = True
End Function
```

---

## 📦 自定义协议

### 协议定义

```vb
' 自定义消息类型
Private Const MSG_TYPE_CHAT As Long = 1
Private Const MSG_TYPE_JOIN As Long = 2
Private Const MSG_TYPE_LEAVE As Long = 3
Private Const MSG_TYPE_SYSTEM As Long = 4

' 消息头结构
Private Type tMessageHeader
    Type As Long       ' 消息类型
    Length As Long     ' 消息长度
    SenderID As String ' 发送者 ID
End Type
```

### 构建自定义消息

```vb
Public Function BuildCustomMessage(ByVal MsgType As Long, ByVal SenderID As String, ByVal Content As String) As Byte()
    Dim baContent() As Byte
    baContent = StringToUTF8(Content)
    
    ' 构建头
    Dim oBuffer As New cByteBuffer
    oBuffer.AppendByte (MsgType And &HFF000000) \ &H1000000
    oBuffer.AppendByte (MsgType And &HFF0000) \ &H10000
    oBuffer.AppendByte (MsgType And &HFF00&) \ &H100&
    oBuffer.AppendByte (MsgType And &HFF&)
    
    Dim lLen As Long
    lLen = UBound(baContent) + 1
    oBuffer.AppendByte (lLen And &HFF000000) \ &H1000000
    oBuffer.AppendByte (lLen And &HFF0000) \ &H10000
    oBuffer.AppendByte (lLen And &HFF00&) \ &H100&
    oBuffer.AppendByte (lLen And &HFF&)
    
    ' 添加 SenderID 长度和内容
    Dim baSenderID() As Byte
    baSenderID = StringToUTF8(SenderID)
    oBuffer.AppendByte (UBound(baSenderID) + 1)
    If UBound(baSenderID) >= 0 Then
        oBuffer.Append baSenderID
    End If
    
    ' 添加内容
    If UBound(baContent) >= 0 Then
        oBuffer.Append baContent
    End If
    
    BuildCustomMessage = oBuffer.ToArray
End Function
```

### 解析自定义消息

```vb
Public Sub ParseCustomMessage(ByVal Data() As Byte)
    Dim oBuffer As New cByteBuffer
    oBuffer.Append Data
    
    ' 读取类型
    Dim lType As Long
    lType = CLng(oBuffer.PeekByte(0)) * 256& ^ 3 + _
             CLng(oBuffer.PeekByte(1)) * 256& ^ 2 + _
             CLng(oBuffer.PeekByte(2)) * 256& + _
             CLng(oBuffer.PeekByte(3))
    oBuffer.Consume 4
    
    ' 读取长度
    Dim lLength As Long
    lLength = CLng(oBuffer.PeekByte(0)) * 256& ^ 3 + _
               CLng(oBuffer.PeekByte(1)) * 256& ^ 2 + _
               CLng(oBuffer.PeekByte(2)) * 256& + _
               CLng(oBuffer.PeekByte(3))
    oBuffer.Consume 4
    
    ' 读取 SenderID
    Dim lSenderLen As Byte
    lSenderLen = oBuffer.PeekByte(0)
    oBuffer.Consume 1
    
    Dim baSenderID() As Byte
    ReDim baSenderID(lSenderLen - 1) As Byte
    If lSenderLen > 0 Then
        Dim i As Long
        For i = 0 To lSenderLen - 1
            baSenderID(i) = oBuffer.PeekByte(i)
        Next i
        oBuffer.Consume lSenderLen
    End If
    Dim sSenderID As String
    sSenderID = UTF8ToString(baSenderID)
    
    ' 读取内容
    Dim baContent() As Byte
    If lLength > 0 Then
        ReDim baContent(lLength - 1) As Byte
        For i = 0 To lLength - 1
            baContent(i) = oBuffer.PeekByte(i)
        Next i
    End If
    Dim sContent As String
    sContent = UTF8ToString(baContent)
    
    ' 处理消息
    Select Case lType
        Case MSG_TYPE_CHAT
            HandleChatMessage sSenderID, sContent
        Case MSG_TYPE_JOIN
            HandleJoinMessage sSenderID
        Case MSG_TYPE_LEAVE
            HandleLeaveMessage sSenderID
        Case MSG_TYPE_SYSTEM
            HandleSystemMessage sContent
    End Select
End Sub
```

---

## 📊 性能优化

### 1. 批量发送

```vb
' ❌ 不好：多次调用 SendText
For i = 0 To 100
    m_Client.SendText "Message " & i
Next i

' ✅ 好：拼接后一次发送
Dim sMessages As String
For i = 0 To 100
    sMessages = sMessages & "Message " & i & vbLf
Next i
m_Client.SendText sMessages
```

### 2. 使用事件而非轮询

```vb
' ✅ 好：使用事件
Private Sub m_Client_OnTextMessage(ByVal Message As String)
    ProcessMessage Message
End Sub

' ❌ 不好：轮询检查
Private Sub Timer1_Timer()
    If m_Client.State = WS_STATE_OPEN Then
        ' 轮询数据（不推荐）
    End If
End Sub
```

### 3. 限制广播频率

```vb
Private WithEvents tmrBroadcast As Timer
Private m_sBroadcastQueue As String

Private Sub QueueBroadcast(ByVal Message As String)
    m_sBroadcastQueue = m_sBroadcastQueue & Message & vbLf
End Sub

Private Sub tmrBroadcast_Timer()
    If LenB(m_sBroadcastQueue) > 0 Then
        m_Server.BroadcastText m_sBroadcastQueue
        m_sBroadcastQueue = ""
    End If
End Sub
```

---

## 🐛 错误处理最佳实践

### 统一错误处理

```vb
' 日志模块
Public Sub LogError(ByVal ModuleName As String, ByVal Procedure As String, ByVal Description As String)
    Dim sLog As String
    sLog = "[" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "] "
    sLog = sLog & ModuleName & "." & Procedure & ": " & Description
    
    Debug.Print sLog
    
    ' 写入文件
    Dim iFile As Integer
    iFile = FreeFile
    Open "error.log" For Append As #iFile
    Print #iFile, sLog
    Close #iFile
End Sub

' 使用示例
Private Sub m_Client_OnError(ByVal Description As String)
    LogError "frmClient", "OnError", Description
End Sub
```

### 连接状态检查

```vb
Public Sub SendMessageSafe(ByVal Message As String)
    If m_Client Is Nothing Then
        Debug.Print "客户端未初始化"
        Exit Sub
    End If
    
    Select Case m_Client.State
        Case WS_STATE_OPEN
            ' 可以发送
            m_Client.SendText Message
            
        Case WS_STATE_CONNECTING
            Debug.Print "正在连接，请稍后"
            
        Case WS_STATE_CLOSING
            Debug.Print "连接正在关闭"
            
        Case WS_STATE_CLOSED
            Debug.Print "连接已关闭"
            
    End Select
End Sub
```

---

## 🔍 调试技巧

### 日志输出

```vb
Private Sub DebugFrame(oFrame As cWebSocketFrame)
    Debug.Print "=== WebSocket 帧 ==="
    Debug.Print "FIN: " & oFrame.FIN
    Debug.Print "OpCode: " & oFrame.OpCode
    Debug.Print "HasMask: " & oFrame.HasMask
    Debug.Print "PayloadLength: " & oFrame.PayloadLength
    Debug.Print "HeaderLength: " & oFrame.HeaderLength
    Debug.Print "TotalFrameLength: " & oFrame.TotalFrameLength
    Debug.Print "IsValid: " & oFrame.IsValid
    Debug.Print "==================="
End Sub
```

### 消息跟踪

```vb
Private Sub LogMessage(ByVal ClientID As String, ByVal Direction As String, ByVal Message As String)
    Dim sLog As String
    sLog = Format$(Now, "hh:nn:ss") & " [" & Direction & "] " & ClientID & ": " & Message
    
    txtLog.Text = txtLog.Text & sLog & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
    
    ' 保存到文件
    Dim iFile As Integer
    iFile = FreeFile
    Open "messages.log" For Append As #iFile
    Print #iFile, sLog
    Close #iFile
End Sub

' 使用
LogMessage ClientID, "OUT", Message
LogMessage ClientID, "IN", Message
```

---

## 📚 参考资料

- [RFC 6455 - WebSocket Protocol](https://tools.ietf.org/html/rfc6455)
- [MDN WebSockets API](https://developer.mozilla.org/zh-CN/docs/Web/API/WebSocket)
- [WebSocket 在线测试工具](https://www.piesocket.com/websocket-tester)

---

**最后更新**: 2026-01-10
