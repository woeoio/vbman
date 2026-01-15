# cWinsock 属性参考

## 📋 属性列表

| 属性名 | 类型 | 读写 | 说明 |
|--------|------|------|------|
| `State` | WinsockState | 只读 | 当前 Socket 状态 |
| `Protocol` | WinsockProtocol | 读写 | 协议类型（TCP/UDP） |
| `RecvBuffer` | Byte() | 读写 | 自定义接收缓冲区 |
| `LocalPort` | Long | 读写 | 本地端口 |
| `RemoteHost` | String | 读写 | 远程主机名 |
| `RemotePort` | Long | 读写 | 远程端口 |
| `RemoteHostIP` | String | 只读 | 解析后的远程 IP 地址 |
| `LocalHostName` | String | 只读 | 本地主机名 |
| `LocalIP` | String | 只读 | 本地 IP 地址 |
| `Tag` | String | 读写 | 用户自定义标签 |
| `UserData` | Variant | 读写 | 用户自定义数据 |
| `SocketHandle` | Long | 只读 | Socket 句柄 |
| `BytesReceived` | Long | 只读 | 可用数据字节数 |
| `IsServer` | Boolean | 只读 | 是否为服务器模式 |
| `IsAcceptedClient` | Boolean | 只读 | 是否为服务器接受的客户端 |
| `ParentServer` | cWinsock | 只读 | 父服务器对象（仅客户端） |
| `Clients` | Collection | 只读 | 所有连接的客户端集合（仅服务器） |
| `ClientCount` | Long | 只读 | 客户端连接数（仅服务器） |

---

## 🔄 State 属性

### 说明

返回当前 Socket 的状态。

### 语法

```vb
Property Get State() As WinsockState
```

### 返回值

| 常量 | 值 | 说明 |
|------|-----|------|
| `sckClosed` | 0 | 已关闭 |
| `sckOpen` | 1 | 已打开（UDP 绑定后） |
| `sckListening` | 2 | 监听中（TCP 服务器） |
| `sckConnectionPending` | 3 | 连接挂起 |
| `sckResolvingHost` | 4 | 正在解析主机名 |
| `sckHostResolved` | 5 | 主机名已解析 |
| `sckConnecting` | 6 | 正在连接 |
| `sckConnected` | 7 | 已连接 |
| `sckClosing` | 8 | 正在关闭 |
| `sckError` | 9 | 发生错误 |

### 使用示例

```vb
Private Sub cmdConnect_Click()
    If m_oClient.State = sckClosed Then
        m_oClient.Connect "127.0.0.1", 8080
    Else
        MsgBox "Socket 未关闭，当前状态: " & GetStateName(m_oClient.State)
    End If
End Sub

Private Function GetStateName(ByVal eState As WinsockState) As String
    Select Case eState
        Case sckClosed:   GetStateName = "已关闭"
        Case sckOpen:     GetStateName = "已打开"
        Case sckListening: GetStateName = "监听中"
        Case sckConnected: GetStateName = "已连接"
        Case sckClosing:  GetStateName = "关闭中"
        Case sckError:    GetStateName = "错误"
        Case Else:        GetStateName = "未知"
    End Select
End Function
```

---

## 🌐 Protocol 属性

### 说明

获取或设置 Socket 使用的协议类型。

### 语法

```vb
Property Get Protocol() As WinsockProtocol
Property Let Protocol(ByVal Value As WinsockProtocol)
```

### 值

| 常量 | 值 | 说明 |
|------|-----|------|
| `sckTCPProtocol` | 1 | TCP 协议（可靠，面向连接） |
| `sckUDPProtocol` | 2 | UDP 协议（不可靠，无连接） |

### 使用示例

```vb
' 设置为 TCP 协议
m_oSocket.Protocol = sckTCPProtocol

' 设置为 UDP 协议
m_oSocket.Protocol = sckUDPProtocol

' 检查当前协议
If m_oSocket.Protocol = sckTCPProtocol Then
    Debug.Print "使用 TCP 协议"
Else
    Debug.Print "使用 UDP 协议"
End If
```

### ⚠️ 注意事项

- 只能在 `State = sckClosed` 时修改
- 修改后需要重新调用 `Connect()`、`Listen()` 或 `Bind()`

---

## 📦 RecvBuffer 属性

### 说明

设置或获取自定义接收缓冲区。通常用于高级场景。

### 语法

```vb
Property Let RecvBuffer(ByRef Value() As Byte)
```

### 使用示例

```vb
' 设置自定义缓冲区
Dim baCustomBuffer() As Byte
ReDim baCustomBuffer(0 To 8191) ' 8KB 缓冲区
m_oSocket.RecvBuffer = baCustomBuffer
```

---

## 🔌 LocalPort 属性

### 说明

获取或设置本地端口号。

### 语法

```vb
Property Get LocalPort() As Long
Property Let LocalPort(ByVal Value As Long)
```

### 使用示例

```vb
' 设置本地端口（必须在调用 Connect/Listen/Bind 之前）
m_oServer.LocalPort = 8080
m_oServer.Listen

' 获取实际绑定的端口
Debug.Print "本地端口: " & m_oSocket.LocalPort
```

### ⚠️ 注意事项

- 只能在 `State = sckClosed` 时设置
- 范围：0-65535
- 0 表示由系统自动分配

---

## 🌍 RemoteHost 属性

### 说明

获取或设置远程主机名（域名或 IP）。

### 语法

```vb
Property Get RemoteHost() As String
Property Let RemoteHost(ByVal Value As String)
```

### 使用示例

```vb
' 设置远程主机（可以使用域名）
m_oClient.RemoteHost = "example.com"
m_oClient.RemotePort = 80
m_oClient.Connect

' 使用 IP 地址
m_oClient.RemoteHost = "192.168.1.100"
m_oClient.RemotePort = 8080
m_oClient.Connect

' 获取远程主机名
Debug.Print "远程主机: " & m_oClient.RemoteHost
```

---

## 🔢 RemotePort 属性

### 说明

获取或设置远程端口号。

### 语法

```vb
Property Get RemotePort() As Long
Property Let RemotePort(ByVal Value As Long)
```

### 使用示例

```vb
' 设置远程端口
m_oClient.RemotePort = 8080

' 获取远程端口
Debug.Print "远程端口: " & m_oClient.RemotePort
```

---

## 🖥️ RemoteHostIP 属性

### 说明

获取解析后的远程 IP 地址（只读）。

### 语法

```vb
Property Get RemoteHostIP() As String
```

### 使用示例

```vb
Private Sub m_oClient_Connect(Client As cWinsock)
    Debug.Print "连接成功!"
    Debug.Print "主机名: " & Client.RemoteHost
    Debug.Print "IP 地址: " & Client.RemoteHostIP
    Debug.Print "端口: " & Client.RemotePort
End Sub
```

### 特殊情况：UDP 服务器虚拟客户端

```vb
Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    ' UDP 服务器模式下，虚拟客户端的 RemoteHostIP 返回发送方 IP
    Debug.Print "收到来自 " & Client.RemoteHostIP & ":" & Client.RemotePort & " 的数据"
End Sub
```

---

## 💻 LocalHostName 属性

### 说明

获取本地主机名。

### 语法

```vb
Property Get LocalHostName() As String
```

### 使用示例

```vb
Debug.Print "本机名: " & m_oSocket.LocalHostName
```

---

## 🌐 LocalIP 属性

### 说明

获取本地 IP 地址。

### 语法

```vb
Property Get LocalIP() As String
```

### 使用示例

```vb
Debug.Print "本机 IP: " & m_oSocket.LocalIP
```

---

## 🏷️ Tag 属性

### 说明

用户自定义标签，用于标识对象。

### 语法

```vb
Property Get Tag() As String
Property Let Tag(ByVal Value As String)
```

### 使用示例

```vb
' 为每个客户端设置标签
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 服务器会自动设置 Tag 为 "#1", "#2", "#3"...
    ' 也可以自定义
    Client.Tag = "客户端-" & Client.RemoteHostIP
    
    Debug.Print "新客户端 Tag: " & Client.Tag
End Sub

' 通过 Tag 查找客户端
Private Function FindClientByTag(ByVal sTag As String) As cWinsock
    Dim oClient As cWinsock
    For Each oClient In m_oServer.Clients
        If oClient.Tag = sTag Then
            Set FindClientByTag = oClient
            Exit Function
        End If
    Next
    Set FindClientByTag = Nothing
End Function
```

---

## 💾 UserData 属性

### 说明

用户自定义数据存储，可以存储任意类型的数据。

### 语法

```vb
Property Get UserData() As Variant
Property Let UserData(ByVal Value As Variant)
Property Set UserData(ByVal Value As Variant)
```

### 使用示例

```vb
' 存储字符串
m_oClient.UserData = "用户信息: 张三"

' 存储数字
m_oClient.UserData = 12345

' 存储对象
Dim oUserInfo As New CUserInfo
oUserInfo.Name = "张三"
oUserInfo.Age = 25
Set m_oClient.UserData = oUserInfo

' 读取数据
Dim sInfo As String
sInfo = m_oClient.UserData
Debug.Print sInfo

' 读取对象
Dim oUserData As CUserInfo
Set oUserData = m_oClient.UserData
Debug.Print oUserInfo.Name & ", " & oUserData.Age
```

### 高级用法：客户端会话数据

```vb
Private Type tSessionData
    LoginTime As Date
    LastActivity As Date
    LoginAttempts As Long
    Authenticated As Boolean
End Type

Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    Dim tSession As tSessionData
    tSession.LoginTime = Now
    tSession.LastActivity = Now
    tSession.LoginAttempts = 0
    tSession.Authenticated = False
    
    Client.UserData = tSession
End Sub

Private Sub CheckSessionTimeout()
    Dim oClient As cWinsock
    Dim tSession As tSessionData
    
    For Each oClient In m_oServer.Clients
        tSession = oClient.UserData
        If DateDiff("s", tSession.LastActivity, Now) > 300 Then ' 5 分钟无活动
            Debug.Print "会话超时，断开: " & oClient.Tag
            oClient.Close_
        End If
    Next
End Sub
```

---

## 🔑 SocketHandle 属性

### 说明

获取底层的 Socket 句柄（只读）。

### 语法

```vb
Property Get SocketHandle() As Long
```

### 使用示例

```vb
' 获取 Socket 句柄
Debug.Print "Socket 句柄: " & m_oSocket.SocketHandle

' 用于高级操作（如与 Win32 API 交互）
If m_oSocket.SocketHandle <> 0 Then
    Call SomeWin32Function(m_oSocket.SocketHandle)
End If
```

---

## 📊 BytesReceived 属性

### 说明

获取接收缓冲区中可用的字节数（只读）。

### 语法

```vb
Property Get BytesReceived() As Long
```

### 使用示例

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Debug.Print "事件通知: " & bytesTotal & " 字节"
    Debug.Print "缓冲区总计: " & Client.BytesReceived & " 字节"
    
    ' 只读取部分数据
    If Client.BytesReceived > 100 Then
        Dim sData As String
        Client.GetData sData, vbString, 100 ' 只读取前 100 字节
        Debug.Print "读取了部分数据: " & sData
    End If
End Sub
```

---

## 🏢 IsServer 属性

### 说明

判断当前对象是否为服务器模式（只读）。

### 语法

```vb
Property Get IsServer() As Boolean
```

### 使用示例

```vb
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    If Client.IsServer Then
        Debug.Print "来自服务器的数据"
    Else
        Debug.Print "来自客户端的数据"
    End If
End Sub
```

---

## 🔗 IsAcceptedClient 属性

### 说明

判断当前对象是否为服务器接受的客户端（只读）。

### 语法

```vb
Property Get IsAcceptedClient() As Boolean
```

### 使用示例

```vb
Private Sub SomeFunction(oSocket As cWinsock)
    If oSocket.IsAcceptedClient Then
        Debug.Print "这是服务器接受的客户端"
        Debug.Print "父服务器: " & oSocket.ParentServer.Tag
    Else
        Debug.Print "这是独立客户端或服务器对象"
    End If
End Sub
```

---

## 👆 ParentServer 属性

### 说明

获取父服务器对象（仅对服务器接受的客户端有效）。

### 语法

```vb
Property Get ParentServer() As cWinsock
```

### 使用示例

```vb
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 服务器设置 ParentServer
    ' 客户端可以访问父服务器
    
    Debug.Print "新客户端的父服务器: " & Client.ParentServer.Tag
End Sub
```

### 高级用法：客户端广播消息

```vb
' 在某个客户端的事件中，通过父服务器向其他客户端广播
Private Sub ClientBroadcastToOthers(ByVal oSender As cWinsock, ByVal sMessage As String)
    Dim oClient As cWinsock
    For Each oClient In oSender.ParentServer.Clients
        If Not oClient Is oSender Then ' 不发送给自己
            oClient.SendData sMessage
        End If
    Next
End Sub
```

---

## 👥 Clients 属性

### 说明

获取所有连接的客户端集合（仅对服务器对象有效）。

### 语法

```vb
Property Get Clients() As Collection
```

### 使用示例

```vb
' 遍历所有客户端
Private Sub ListAllClients()
    Debug.Print "当前连接数: " & m_oServer.ClientCount
    
    Dim oClient As cWinsock
    For Each oClient In m_oServer.Clients
        Debug.Print oClient.Tag & ": " & oClient.RemoteHostIP & ":" & oClient.RemotePort
    Next
End Sub

' 查找特定客户端
Private Function FindClientByIP(ByVal sIP As String) As cWinsock
    Dim oClient As cWinsock
    For Each oClient In m_oServer.Clients
        If oClient.RemoteHostIP = sIP Then
            Set FindClientByIP = oClient
            Exit Function
        End If
    Next
    Set FindClientByIP = Nothing
End Function

' 向所有客户端广播
Private Sub BroadcastToAll(ByVal sMessage As String)
    Dim oClient As cWinsock
    For Each oClient In m_oServer.Clients
        On Error Resume Next
        oClient.SendData sMessage
        On Error GoTo 0
    Next
End Sub
```

---

## 🔢 ClientCount 属性

### 说明

获取当前连接的客户端数量（只读，仅对服务器对象有效）。

### 语法

```vb
Property Get ClientCount() As Long
```

### 使用示例

```vb
' 显示连接数
lblClientCount.Caption = "当前连接: " & m_oServer.ClientCount

' 限制最大连接数
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    If m_oServer.ClientCount >= m_lMaxClients Then
        Debug.Print "达到最大连接数限制: " & m_lMaxClients
        DisConnect = True
    End If
End Sub
```

---

## 📌 属性使用场景总结

### 客户端常用属性

```vb
' 连接前设置
m_oClient.Protocol = sckTCPProtocol
m_oClient.RemoteHost = "192.168.1.100"
m_oClient.RemotePort = 8080
m_oClient.Connect

' 连接后获取
Debug.Print "IP: " & m_oClient.RemoteHostIP
Debug.Print "端口: " & m_oClient.RemotePort
Debug.Print "状态: " & m_oClient.State

' 自定义标签
m_oClient.Tag = "客户端-001"
m_oClient.UserData = "用户信息"
```

### 服务器常用属性

```vb
' 启动服务器
m_oServer.Protocol = sckTCPProtocol
m_oServer.LocalPort = 8080
m_oServer.Listen

' 管理客户端
Debug.Print "连接数: " & m_oServer.ClientCount

Dim oClient As cWinsock
For Each oClient In m_oServer.Clients
    Debug.Print oClient.Tag & ": " & oClient.RemoteHostIP
    oClient.SendData "广播消息"
Next
```

---

**最后更新**: 2026-01-09
