# cWinsock 方法参考

## 📋 方法列表

| 方法名 | 返回类型 | 说明 |
|--------|----------|------|
| `Connect` | Sub | 连接到远程服务器 |
| `Listen` | Sub | 开始监听端口 |
| `Bind` | Sub | 绑定本地端口（UDP） |
| `SendData` | Sub | 发送数据 |
| `GetData` | Sub | 接收数据 |
| `PeekData` | Sub | 查看数据但不移除 |
| `Close_` | Sub | 关闭连接 |
| `GetErrorDescription` | String | 获取错误描述 |
| `AcceptFrom` | Sub | 接受连接（内部方法） |
| `SetUdpClientInfo` | Sub | 设置 UDP 客户端信息（内部方法） |
| `RemoveClient` | Sub | 移除客户端（内部方法） |
| `RaiseDataArrivalEvent` | Sub | 触发数据到达事件（内部方法） |

---

## 🔗 Connect 方法

### 说明

连接到指定的远程服务器（TCP 客户端模式）。

### 语法

```vb
Public Sub Connect(Optional RemoteHost As String, Optional ByVal RemotePort As Long)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `RemoteHost` | String（可选） | 远程主机名或 IP 地址。如果不提供，使用 `RemoteHost` 属性的值 |
| `RemotePort` | Long（可选） | 远程端口号。如果不提供，使用 `RemotePort` 属性的值 |

### 使用示例

```vb
' 使用参数连接
m_oClient.Connect "127.0.0.1", 8080

' 使用属性连接
m_oClient.RemoteHost = "example.com"
m_oClient.RemotePort = 80
m_oClient.Connect

' 连接到特定主机
m_oClient.RemoteHost = "192.168.1.100"
m_oClient.Connect , 8080  ' 只指定端口，使用已设置的 RemoteHost
```

### 连接流程

```
1. 调用 Connect()
2. 关闭现有连接（如果有）
3. 解析主机名 → sckResolvingHost
4. 主机名解析完成 → sckHostResolved
5. 开始连接 → sckConnecting
6. 连接成功 → sckConnected
7. 触发 Connect 事件
```

### 错误处理

```vb
Private Sub cmdConnect_Click()
    On Error GoTo EH
    
    m_oClient.Connect "example.com", 8080
    Exit Sub
    
EH:
    Debug.Print "连接错误: " & Err.Description
    Select Case Err.Number
        Case 10060
            MsgBox "连接超时，请检查网络"
        Case 10061
            MsgBox "服务器拒绝连接，请检查端口"
        Case Else
            MsgBox "连接失败: " & Err.Description
    End Select
End Sub
```

---

## 🎧 Listen 方法

### 说明

开始监听指定端口，等待客户端连接（TCP 服务器模式）。

### 语法

```vb
Public Sub Listen(Optional ByVal Port As Long)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Port` | Long（可选） | 要监听的端口号。如果不提供，使用 `LocalPort` 属性的值 |

### 使用示例

```vb
' 使用参数监听
m_oServer.Listen 8080

' 使用属性监听
m_oServer.LocalPort = 8080
m_oServer.Listen

' 监听多个端口（需要多个 cWinsock 对象）
Dim oServer1 As New cWinsock
Dim oServer2 As New cWinsock
oServer1.Listen 8080
oServer2.Listen 8081
```

### 服务器启动流程

```vb
Private Sub StartServer()
    On Error GoTo EH
    
    ' 设置协议
    m_oServer.Protocol = sckTCPProtocol
    
    ' 开始监听
    m_oServer.Listen 8080
    
    Debug.Print "服务器已启动，监听端口: " & m_oServer.LocalPort
    
    ' 更新 UI
    btnStart.Enabled = False
    btnStop.Enabled = True
    lblStatus.Caption = "监听中..."
    
    Exit Sub
    
EH:
    Debug.Print "启动服务器失败: " & Err.Description
    MsgBox "无法启动服务器: " & Err.Description, vbExclamation
End Sub
```

### ⚠️ 注意事项

- 调用 `Listen()` 前必须设置 `Protocol = sckTCPProtocol`
- 端口必须未被占用
- `State` 将变为 `sckListening`

---

## 📌 Bind 方法

### 说明

绑定本地端口（UDP 服务器模式）。

### 语法

```vb
Public Sub Bind(Optional ByVal LocalPort As Long, Optional LocalIP As String)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `LocalPort` | Long（可选） | 要绑定的本地端口号 |
| `LocalIP` | String（可选） | 要绑定的本地 IP 地址（可选） |

### 使用示例

```vb
' 绑定端口
m_oUdp.Protocol = sckUDPProtocol
m_oUdp.Bind 8888

' 绑定到特定 IP
m_oUdp.Bind 8888, "192.168.1.100"
```

### UDP 服务器启动

```vb
Private Sub StartUdpServer()
    On Error GoTo EH
    
    ' 设置协议
    m_oUdp.Protocol = sckUDPProtocol
    
    ' 绑定端口
    m_oUdp.Bind 8888
    
    Debug.Print "UDP 服务器已启动，绑定端口: " & m_oUdp.LocalPort
    
    Exit Sub
    
EH:
    Debug.Print "UDP 绑定失败: " & Err.Description
    MsgBox "无法绑定 UDP 端口: " & Err.Description, vbExclamation
End Sub
```

---

## 📤 SendData 方法

### 说明

发送数据到远程主机。

### 语法

```vb
Public Sub SendData(Data As Variant, Optional ByVal CodePage As EnumScpCodePage = ScpAcp)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data` | Variant | 要发送的数据，可以是字符串或字节数组 |
| `CodePage` | EnumScpCodePage（可选） | 文本编码，默认 `ScpAcp`（GBK/ACP） |

### 编码选项

| 常量 | 值 | 说明 |
|------|-----|------|
| `ScpAcp` | 0 | 系统默认代码页（中文 Windows 上为 GBK） |
| `ScpOem` | 1 | OEM 代码页 |
| `ScpUtf8` | 65001 | UTF-8 编码 |
| `ScpUnicode` | -1 | Unicode，不进行编码转换 |

### 发送字符串

```vb
' 默认使用 ACP/GBK 编码
m_oClient.SendData "中文测试"

' 使用 UTF-8 编码
m_oClient.SendData "中文测试", ScpUtf8

' 使用 Unicode（不转换）
m_oClient.SendData "中文测试", ScpUnicode
```

### 发送字节数组

```vb
' 发送字节数组
Dim baData() As Byte
baData = GetBinaryData()
m_oClient.SendData baData
```

### UDP 服务器发送

```vb
' UDP 服务器模式下，需要指定远程地址
Private Sub cmdUdpSend_Click()
    ' 设置目标
    m_oUdp.RemoteHost = "127.0.0.1"
    m_oUdp.RemotePort = 9999
    
    ' 发送数据
    m_oUdp.SendData "Hello, UDP!"
End Sub

' 向特定客户端回复（虚拟客户端）
Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    ' 直接通过 Client 对象回复
    ' cWinsock 会自动使用正确的远程地址:端口
    Client.SendData "Reply: " & sData
End Sub
```

### 大数据发送

```vb
' 分块发送大数据
Private Sub SendLargeFile(ByVal sFilePath As String)
    Dim baChunk() As Byte
    Dim lChunkSize As Long
    lChunkSize = 8192 ' 8KB 每块
    
    ' 打开文件...
    ' 循环读取并发送
    Do While Not EOF
        ' 读取数据块
        ReadChunk baChunk, lChunkSize
        
        ' 发送
        m_oClient.SendData baChunk
        
        ' 等待发送完成（通过 SendComplete 事件）
        Do While m_bSending
            DoEvents
        Loop
    Loop
End Sub
```

---

## 📥 GetData 方法

### 说明

从接收缓冲区读取数据。

### 语法

```vb
Public Sub GetData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal MaxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = ScpAcp)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data` | Variant | 用于接收数据的变量 |
| `VarType_` | Long（可选） | 期望的数据类型（如 `vbString`, `vbByte + vbArray`） |
| `MaxLen` | Long（可选） | 最大读取字节数，-1 表示读取全部 |
| `CodePage` | EnumScpCodePage（可选） | 文本编码，默认 `ScpAcp` |

### 读取字符串

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 读取所有数据
    Client.GetData sData
    
    Debug.Print "收到: " & sData
End Sub
```

### 读取字节数组

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    
    ' 读取字节数组
    Client.GetData baData
    
    Debug.Print "收到 " & bytesTotal & " 字节"
End Sub
```

### 部分读取

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sHeader As String
    Dim sBody As String
    
    ' 读取前 10 字节作为头部
    Client.GetData sHeader, vbString, 10
    Debug.Print "头部: " & sHeader
    
    ' 读取剩余数据
    Client.GetData sBody
    Debug.Print "正文: " & sBody
End Sub
```

### 指定编码读取

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 使用 UTF-8 编码读取
    Client.GetData sData, vbString, -1, ScpUtf8
    
    Debug.Print "UTF-8 数据: " & sData
End Sub
```

### 协议解析

```vb
Private Type tPacketHeader
    Magic As Long ' 魔数
    Length As Long ' 数据长度
    Type As Long ' 数据类型
End Type

Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim tHeader As tPacketHeader
    Dim baBody() As Byte
    
    ' 读取头部
    Client.GetData tHeader
    
    ' 验证魔数
    If tHeader.Magic = &H12345678 Then
        ' 读取数据体
        ReDim baBody(0 To tHeader.Length - 1) As Byte
        Client.GetData baBody
        
        Debug.Print "数据类型: " & tHeader.Type
        Debug.Print "数据长度: " & tHeader.Length
    End If
End Sub
```

---

## 👁️ PeekData 方法

### 说明

查看数据但不从缓冲区移除。

### 语法

```vb
Public Sub PeekData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal MaxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = ScpAcp)
```

### 参数

与 `GetData` 相同。

### 使用示例

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sPeek As String
    Dim sActual As String
    
    ' 先查看数据
    Client.PeekData sPeek
    Debug.Print "查看数据: " & sPeek
    
    ' 然后读取数据
    Client.GetData sActual
    Debug.Print "实际数据: " & sActual
End Sub
```

### 协议检测

```vb
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sPeek As String
    
    ' 先查看前几个字符以检测协议
    Client.PeekData sPeek, vbString, 4
    
    If Left$(sPeek, 4) = "HTTP" Then
        Debug.Print "HTTP 请求"
        HandleHttpRequest Client
    ElseIf Left$(sPeek, 4) = "CHAT" Then
        Debug.Print "聊天协议"
        HandleChatMessage Client
    Else
        Debug.Print "未知协议"
    End If
End Sub
```

---

## 🔒 Close_ 方法

### 说明

关闭连接或停止监听。

### 语法

```vb
Public Sub Close_()
```

### 使用示例

```vb
' 关闭客户端连接
Private Sub cmdDisconnect_Click()
    m_oClient.Close_
    Debug.Print "已断开连接"
End Sub

' 停止服务器
Private Sub cmdStopServer_Click()
    m_oServer.Close_
    Debug.Print "服务器已停止"
End Sub

' 关闭特定客户端
Private Sub DisconnectClient(ByVal oClient As cWinsock)
    oClient.Close_
    m_oServer.RemoveClient oClient
End Sub
```

### 自动关闭

```vb
' 窗体卸载时自动关闭
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    m_oClient.Close_
    m_oServer.Close_
    m_oUdp.Close_
End Sub
```

---

## 📝 GetErrorDescription 方法

### 说明

获取错误代码的描述信息。

### 语法

```vb
Public Function GetErrorDescription(ByVal ErrorCode As Long) As String
```

### 使用示例

```vb
Private Sub m_oClient_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    ' 使用参数中的描述
    Debug.Print "错误: " & Description
    
    ' 或使用 GetErrorDescription 获取
    Debug.Print "错误描述: " & Client.GetErrorDescription(Number)
End Sub

' 独立使用
Dim sDesc As String
sDesc = m_oClient.GetErrorDescription(10060)
Debug.Print sDesc ' "连接超时"
```

---

## 🤝 Friend 方法

以下方法是内部使用的方法，通常不需要直接调用：

### AcceptFrom

接受新的连接（由 `OnAccept` 事件调用）。

### SetUdpClientInfo

设置 UDP 虚拟客户端的信息（由 `OnReceive` 事件调用）。

### RemoveClient

移除客户端（由 `CloseEvent` 或手动调用）。

### RaiseDataArrivalEvent

触发数据到达事件（由客户端对象调用，通过父服务器触发）。

---

## 📌 方法使用场景总结

### TCP 客户端流程

```vb
1. m_oClient.Connect("127.0.0.1", 8080)
2. 等待 m_oClient_Connect 事件
3. m_oClient.SendData("Hello")
4. 等待 m_oClient_DataArrival 事件
5. Client.GetData sData
6. m_oClient.Close_()
```

### TCP 服务器流程

```vb
1. m_oServer.Listen(8080)
2. 等待 m_oServer_ConnectionRequest 事件
3. 设置 DisConnect = False 接受连接
4. 等待 m_oServer_DataArrival 事件
5. Client.GetData sData
6. Client.SendData("Reply")
7. 等待 m_oServer_CloseEvent 事件
```

### UDP 流程

```vb
1. m_oUdp.Bind(8888)
2. 等待 m_oUdp_DataArrival 事件
3. Client.GetData sData
4. Client.SendData("Reply")
5. m_oUdp.Close_()
```

---

**最后更新**: 2026-01-09
