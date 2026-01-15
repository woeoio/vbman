# 快速开始

本指南将帮助您快速上手 WebSocket 类库，创建基本的服务端和客户端应用。

---

## 📦 前置准备

### 必需文件

确保以下文件已添加到项目中：

| 文件 | 位置 | 说明 |
|------|------|------|
| `cWinsock.cls` | `add/` | 底层 Socket 封装 |
| `cWebSocketClient.cls` | `newWebsocket/` | 客户端类 |
| `cWebSocketServer.cls` | `newWebsocket/` | 服务端类 |
| `cWebSocketFrame.cls` | `newWebsocket/` | 帧解析类 |
| `cByteBuffer.cls` | `newWebsocket/` | 缓冲区类 |
| `cWebSocketServerClient.cls` | `newWebsocket/` | 服务端客户端类 |
| `mWebSocketUtils.bas` | `newWebsocket/` | 工具模块 |

### 添加到项目

1. 打开 VB6 项目
2. 菜单：项目 → 添加类模块 / 添加模块
3. 浏览到相应文件并添加

---

## 🚀 客户端快速入门

### 步骤 1：创建窗体

创建一个新窗体，添加以下控件：

- 1 个 TextBox (`txtURL`) - 用于输入服务器地址
- 2 个 CommandButton (`cmdConnect`, `cmdDisconnect`) - 连接/断开
- 1 个 TextBox (`txtMessage`) - 输入消息
- 1 个 CommandButton (`cmdSend`) - 发送消息
- 1 个 TextBox (`txtLog`) - 显示日志（MultiLine = True）

### 步骤 2：编写代码

```vb
Option Explicit

Private WithEvents m_Client As cWebSocketClient

Private Sub Form_Load()
    Set m_Client = New cWebSocketClient
    
    txtURL.Text = "ws://127.0.0.1:8080"
    txtMessage.Text = "Hello WebSocket!"
    
    UpdateUI False
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo EH
    
    m_Client.Connect txtURL.Text
    LogMessage "正在连接到: " & txtURL.Text
    Exit Sub
    
EH:
    LogMessage "连接失败: " & Err.Description
End Sub

Private Sub cmdDisconnect_Click()
    If Not m_Client Is Nothing Then
        m_Client.CloseConnection
        LogMessage "已断开连接"
    End If
    UpdateUI False
End Sub

Private Sub cmdSend_Click()
    On Error GoTo EH
    
    If m_Client.State = WS_STATE_OPEN Then
        m_Client.SendText txtMessage.Text
        LogMessage "已发送: " & txtMessage.Text
    End If
    Exit Sub
    
EH:
    LogMessage "发送失败: " & Err.Description
End Sub

' ====== WebSocket 事件处理 ======

Private Sub m_Client_OnOpen()
    LogMessage "已成功连接到 WebSocket 服务器"
    UpdateUI True
End Sub

Private Sub m_Client_OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
    LogMessage "连接已关闭: " & Reason & " (状态码: " & Code & ")"
    UpdateUI False
End Sub

Private Sub m_Client_OnTextMessage(ByVal Message As String)
    LogMessage "收到消息: " & Message
End Sub

Private Sub m_Client_OnError(ByVal Description As String)
    LogMessage "错误: " & Description
End Sub

' ====== 辅助函数 ======

Private Sub LogMessage(sMessage As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:mm:ss") & " - " & sMessage & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub UpdateUI(bConnected As Boolean)
    cmdConnect.Enabled = Not bConnected
    cmdDisconnect.Enabled = bConnected
    cmdSend.Enabled = bConnected
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_Client Is Nothing Then
        m_Client.CloseConnection
    End If
End Sub
```

### 步骤 3：运行测试

1. 按 F5 运行程序
2. 输入服务器地址（如 `ws://127.0.0.1:8080`）
3. 点击"连接"
4. 连接成功后，输入消息并点击"发送"

---

## 🌐 服务端快速入门

### 步骤 1：创建窗体

创建一个新窗体，添加以下控件：

- 1 个 TextBox (`txtPort`) - 端口号
- 2 个 CommandButton (`cmdStart`, `cmdStop`) - 启动/停止
- 1 ListBox (`lstClients`) - 客户端列表
- 1 个 TextBox (`txtLog`) - 显示日志（MultiLine = True）

### 步骤 2：编写代码

```vb
Option Explicit

Private WithEvents m_Server As cWebSocketServer

Private Sub Form_Load()
    Set m_Server = New cWebSocketServer
    txtPort.Text = "8080"
End Sub

Private Sub cmdStart_Click()
    On Error GoTo EH
    
    m_Server.Listen CLng(txtPort.Text)
    LogMessage "服务器已启动，监听端口: " & txtPort.Text
    Exit Sub
    
EH:
    LogMessage "启动失败: " & Err.Description
End Sub

Private Sub cmdStop_Click()
    If Not m_Server Is Nothing Then
        m_Server.StopServer
        LogMessage "服务器已停止"
    End If
    lstClients.Clear
End Sub

' ====== WebSocket 服务端事件处理 ======

Private Sub m_Server_OnStart(ByVal Port As Long)
    LogMessage "服务已启动，监听端口: " & Port
End Sub

Private Sub m_Server_OnStop()
    LogMessage "服务已停止"
End Sub

Private Sub m_Server_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)
    LogMessage "客户端连接: " & ClientID & " (" & RemoteAddress & ":" & RemotePort & ")"
    lstClients.AddItem ClientID & " - " & RemoteAddress
    
    ' 发送欢迎消息
    m_Server.SendText ClientID, "欢迎连接到 WebSocket 服务器！"
End Sub

Private Sub m_Server_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    LogMessage "客户端断开: " & ClientID & " - " & Reason
    
    ' 从列表中移除
    Dim i As Long
    For i = 0 To lstClients.ListCount - 1
        If InStr(lstClients.List(i), ClientID) > 0 Then
            lstClients.RemoveItem i
            Exit For
        End If
    Next
End Sub

Private Sub m_Server_OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
    LogMessage "收到来自 " & ClientID & " 的消息: " & Message
    
    ' 回显消息
    m_Server.SendText ClientID, "服务器收到: " & Message
    
    ' 广播给所有其他客户端（聊天模式）
    m_Server.BroadcastText ClientID & ": " & Message, ClientID
End Sub

Private Sub m_Server_OnClientBinaryMessage(ByVal ClientID As String, Data() As Byte)
    LogMessage "收到来自 " & ClientID & " 的二进制消息: " & (UBound(Data) + 1) & " 字节"
End Sub

Private Sub m_Server_OnError(ByVal Description As String)
    LogMessage "错误: " & Description
End Sub

' ====== 辅助函数 ======

Private Sub LogMessage(sMessage As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:mm:ss") & " - " & sMessage & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_Server Is Nothing Then
        m_Server.StopServer
    End If
End Sub
```

### 步骤 3：运行测试

1. 按 F5 运行服务端程序
2. 点击"启动服务"
3. 运行上面创建的客户端程序
4. 点击"连接"
5. 发送消息测试

---

## 💬 聊天室示例

### 服务端代码

```vb
Option Explicit

Private WithEvents m_Server As cWebSocketServer

Private Sub Form_Load()
    Set m_Server = New cWebSocketServer
    m_Server.Listen 8080
End Sub

Private Sub m_Server_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)
    Debug.Print ClientID & " 加入了聊天室"
    
    ' 通知其他用户
    m_Server.BroadcastText "[系统] " & ClientID & " 加入了聊天室", ClientID
    
    ' 发送欢迎消息
    m_Server.SendText ClientID, "欢迎来到聊天室！当前在线: " & m_Server.ClientCount & " 人"
End Sub

Private Sub m_Server_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    Debug.Print ClientID & " 离开了聊天室"
    m_Server.BroadcastText "[系统] " & ClientID & " 离开了聊天室"
End Sub

Private Sub m_Server_OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
    ' 广播给所有用户
    m_Server.BroadcastText ClientID & ": " & Message, ClientID
    
    Debug.Print ClientID & ": " & Message
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_Server.StopServer
End Sub
```

### 客户端代码

```vb
Option Explicit

Private WithEvents m_Client As cWebSocketClient
Private m_sUsername As String

Private Sub Form_Load()
    Set m_Client = New cWebSocketClient
    m_sUsername = "用户" & CStr(Int(Rnd * 1000))
    
    txtURL.Text = "ws://127.0.0.1:8080"
End Sub

Private Sub cmdConnect_Click()
    m_Client.Connect txtURL.Text
End Sub

Private Sub cmdSend_Click()
    If m_Client.State = WS_STATE_OPEN Then
        m_Client.SendText txtMessage.Text
        txtMessage.Text = ""
    End If
End Sub

Private Sub m_Client_OnOpen()
    Debug.Print "已连接到聊天室"
    UpdateUI True
End Sub

Private Sub m_Client_OnTextMessage(ByVal Message As String)
    txtChat.Text = txtChat.Text & Message & vbCrLf
    txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub UpdateUI(bConnected As Boolean)
    cmdConnect.Enabled = Not bConnected
    cmdSend.Enabled = bConnected
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_Client.CloseConnection
End Sub
```

---

## 🔍 常见问题

### Q1: 编译错误"用户定义类型未定义"

**原因**: 未添加 `mWebSocketUtils.bas` 模块。

**解决**: 
1. 菜单：项目 → 添加模块
2. 浏览到 `newWebsocket/mWebSocketUtils.bas`
3. 添加到项目

---

### Q2: 连接失败"无法解析主机名"

**原因**: URL 格式错误或网络问题。

**解决**: 
- 检查 URL 格式：`ws://host:port`
- 确保服务端已启动
- 使用 `127.0.0.1` 而非 `localhost`

---

### Q3: 握手失败"Handshake failed"

**原因**: 
- 服务端未实现 WebSocket 握手
- 端口被其他程序占用
- 防火墙阻止

**解决**: 
- 确保使用 WebSocket 类库的服务端
- 更换端口
- 检查防火墙设置

---

### Q4: 收到乱码

**原因**: 编码问题。

**解决**: WebSocket 类库自动处理 UTF-8 编码，不需要手动转换。

---

### Q5: 如何发送二进制数据

```vb
' 发送二进制数据
Dim baData() As Byte
baData = LoadFile("image.png")

m_Client.SendBinary baData

' 接收二进制数据
Private Sub m_Client_OnBinaryMessage(Data() As Byte)
    Debug.Print "收到 " & (UBound(Data) + 1) & " 字节"
    SaveFile Data, "received.png"
End Sub
```

---

### Q6: 如何实现自动重连

```vb
Private WithEvents m_Client As cWebSocketClient
Private WithEvents tmrReconnect As Timer
Private m_bAutoReconnect As Boolean
Private m_sServerURL As String

Private Sub Form_Load()
    Set m_Client = New cWebSocketClient
    Set tmrReconnect = New Timer
    tmrReconnect.Interval = 5000  ' 5 秒后重连
    
    m_bAutoReconnect = True
    m_sServerURL = "ws://127.0.0.1:8080"
    
    ConnectToServer
End Sub

Private Sub ConnectToServer()
    If m_Client.State = WS_STATE_CLOSED Then
        m_Client.Connect m_sServerURL
    End If
End Sub

Private Sub m_Client_OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
    Debug.Print "连接关闭: " & Reason
    
    If m_bAutoReconnect Then
        Debug.Print "5 秒后重连..."
        tmrReconnect.Enabled = True
    End If
End Sub

Private Sub tmrReconnect_Timer()
    tmrReconnect.Enabled = False
    ConnectToServer
End Sub
```

---

## 📚 下一步

- 查看 [client.md](./client.md) 了解客户端详细 API
- 查看 [server.md](./server.md) 了解服务端详细 API
- 查看 [advanced.md](./advanced.md) 了解高级功能

---

**最后更新**: 2026-01-10
