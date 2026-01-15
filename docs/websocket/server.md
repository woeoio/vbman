# cWebSocketServer 类参考

## 📋 类概述

`cWebSocketServer` 是 WebSocket 服务端实现类，提供监听端口、管理多个客户端连接、广播消息等功能。

---

## 📡 事件列表

| 事件名 | 触发时机 | 参数 |
|--------|----------|------|
| `OnStart` | 服务器启动 | `Port` (监听端口) |
| `OnStop` | 服务器停止 | 无 |
| `OnClientConnect` | 新客户端连接 | `ClientID`, `RemoteAddress`, `RemotePort` |
| `OnClientDisconnect` | 客户端断开 | `ClientID`, `Reason` |
| `OnClientTextMessage` | 收到客户端文本消息 | `ClientID`, `Message` |
| `OnClientBinaryMessage` | 收到客户端二进制消息 | `ClientID`, `Data()` |
| `OnError` | 发生错误 | `Description` |

---

## 🔧 属性参考

### Port - 监听端口

**类型**: `Long`  
**读写**: 只读

**说明**: 当前监听的端口号。

**示例**:

```vb
Debug.Print "服务器监听端口: " & m_Server.Port
```

---

### IsListening - 是否监听中

**类型**: `Boolean`  
**读写**: 只读

**说明**: 服务器是否正在监听。

**示例**:

```vb
If m_Server.IsListening Then
    Debug.Print "服务器正在监听"
End If
```

---

### ClientCount - 客户端连接数

**类型**: `Long`  
**读写**: 只读

**说明**: 当前连接的客户端数量。

**示例**:

```vb
Debug.Print "当前连接数: " & m_Server.ClientCount

' 更新 UI
lblClientCount.Caption = "连接数: " & m_Server.ClientCount
```

---

### ClientIDs - 客户端 ID 数组

**类型**: `Variant` (String 数组)  
**读写**: 只读

**说明**: 所有连接客户端的 ID 数组。

**示例**:

```vb
' 获取所有客户端 ID
Dim vIDs() As Variant
vIDs = m_Server.ClientIDs

' 遍历所有客户端 ID
Dim i As Long
For i = LBound(vIDs) To UBound(vIDs)
    Debug.Print "客户端: " & vIDs(i)
Next i
```

---

## 🚀 方法参考

### Listen - 启动监听

**语法**:

```vb
Public Sub Listen(Optional ByVal Port As Long = 8080)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Port` | Long（可选） | 监听端口号，默认 8080 |

**说明**: 
- 如果服务器已在监听，会先停止再启动
- 启动后会触发 `OnStart` 事件
- 会自动初始化客户端集合

**示例**:

```vb
' 使用默认端口 8080
m_Server.Listen

' 指定端口
m_Server.Listen 9000

' 从配置文件读取端口
m_Server.Listen CLng(GetConfig("ServerPort"))
```

**错误处理**:

```vb
Private Sub cmdStart_Click()
    On Error GoTo EH

    m_Server.Listen CLng(txtPort.Text)
    Debug.Print "服务器已启动"
    Exit Sub

EH:
    Debug.Print "启动失败: " & Err.Description
    MsgBox "无法启动服务器: " & Err.Description, vbExclamation
End Sub
```

---

### StopServer - 停止服务器

**语法**:

```vb
Public Sub StopServer()
```

**说明**:
- 会向所有客户端发送关闭帧
- 关闭所有客户端连接
- 清空客户端集合
- 关闭监听 Socket
- 触发 `OnStop` 事件

**示例**:

```vb
' 停止服务器
m_Server.StopServer
Debug.Print "服务器已停止"

' 窗体关闭时自动停止
Private Sub Form_Unload(Cancel As Integer)
    m_Server.StopServer
End Sub
```

---

### SendText - 发送文本消息

**语法**:

```vb
Public Sub SendText(ByVal ClientID As String, ByVal Message As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 目标客户端 ID |
| `Message` | String | 要发送的文本消息 |

**说明**: 向指定客户端发送文本消息。消息会自动添加 WebSocket 帧头，服务端发送不进行掩码。

**示例**:

```vb
' 发送欢迎消息
m_Server.SendText ClientID, "欢迎连接到 WebSocket 服务器！"

' 回显消息
m_Server.SendText ClientID, "服务器收到: " & Message

' 发送 JSON 数据
Dim sJSON As String
sJSON = "{""type"":""notification"", ""message"":""Hello""}"
m_Server.SendText ClientID, sJSON

' 发送系统消息
m_Server.SendText ClientID, "[系统] 服务器将于 5 分钟后维护"
```

---

### SendBinary - 发送二进制消息

**语法**:

```vb
Public Sub SendBinary(ByVal ClientID As String, Data() As Byte)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 目标客户端 ID |
| `Data()` | Byte() | 要发送的二进制数据 |

**示例**:

```vb
' 发送图片数据
Dim baImage() As Byte
baImage = LoadImageAsByteArray()
m_Server.SendBinary ClientID, baImage

' 发送文件数据
Dim baFile() As Byte
baFile = LoadFile("document.pdf")
m_Server.SendBinary ClientID, baFile

' 发送序列化对象
Dim baObj() As Byte
baObj = SerializeObject(myObject)
m_Server.SendBinary ClientID, baObj
```

---

### BroadcastText - 广播文本消息

**语法**:

```vb
Public Sub BroadcastText(ByVal Message As String, Optional ByVal ExcludeClientID As String = "")
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Message` | String | 要广播的文本消息 |
| `ExcludeClientID` | String（可选） | 排除的客户端 ID（不向该客户端发送） |

**说明**: 向所有连接的客户端发送文本消息。帧只构建一次，然后发送给所有客户端，提高性能。

**示例**:

```vb
' 向所有客户端广播
m_Server.BroadcastText "欢迎来到聊天室！"

' 广播但不包括发送者
m_Server.BroadcastText Message, SenderClientID

' 系统公告
m_Server.BroadcastText "[系统] 服务器将在 5 分钟后重启"

' 聊天消息广播
Private Sub m_Server_OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
    ' 将消息广播给所有客户端，但不包括发送者
    m_Server.BroadcastText ClientID & ": " & Message, ClientID
End Sub

' 定时广播
Private Sub Timer1_Timer()
    Dim sTime As String
    sTime = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    m_Server.BroadcastText "[时间] " & sTime
End Sub
```

---

### BroadcastBinary - 广播二进制消息

**语法**:

```vb
Public Sub BroadcastBinary(Data() As Byte, Optional ByVal ExcludeClientID As String = "")
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | 要广播的二进制数据 |
| `ExcludeClientID` | String（可选） | 排除的客户端 ID |

**示例**:

```vb
' 广播图片更新
Dim baImage() As Byte
baImage = GetUpdatedImage()
m_Server.BroadcastBinary baImage

' 广播配置文件
Dim baConfig() As Byte
baConfig = SerializeConfig()
m_Server.BroadcastBinary baConfig

' 排除发送者
m_Server.BroadcastBinary baData, SenderClientID
```

---

### DisconnectClient - 断开客户端连接

**语法**:

```vb
Public Sub DisconnectClient(ByVal ClientID As String, _
                           Optional ByVal Code As WsCloseCode = WS_CLOSE_NORMAL, _
                           Optional ByVal Reason As String = "")
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 要断开的客户端 ID |
| `Code` | WsCloseCode（可选） | 关闭状态码，默认 `WS_CLOSE_NORMAL` |
| `Reason` | String（可选） | 关闭原因 |

**示例**:

```vb
' 正常断开
m_Server.DisconnectClient ClientID

' 指定关闭原因
m_Server.DisconnectClient ClientID, WS_CLOSE_GOING_AWAY, "管理员断开"

' 违规用户断开
If IsViolation(ClientID) Then
    m_Server.DisconnectClient ClientID, WS_CLOSE_POLICY_VIOLATION, "违反聊天规则"
End If

' 服务器维护时断开所有客户端
Private Sub PrepareForMaintenance()
    Dim vIDs() As Variant
    Dim i As Long
    vIDs = m_Server.ClientIDs
    
    For i = LBound(vIDs) To UBound(vIDs)
        m_Server.DisconnectClient vIDs(i), WS_CLOSE_GOING_AWAY, "服务器维护中"
    Next i
End Sub
```

---

## 📡 事件详解

### OnStart - 服务器启动

**语法**:

```vb
Event OnStart(ByVal Port As Long)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Port` | Long | 监听的端口号 |

**示例**:

```vb
Private Sub m_Server_OnStart(ByVal Port As Long)
    Debug.Print "服务器已启动，监听端口: " & Port
    
    ' 更新 UI
    lblStatus.Caption = "运行中"
    lblPort.Caption = Port
    
    ' 记录日志
    LogEvent "Server started on port " & Port
    
    ' 启动定时任务
    Timer1.Enabled = True
End Sub
```

---

### OnStop - 服务器停止

**语法**:

```vb
Event OnStop()
```

**示例**:

```vb
Private Sub m_Server_OnStop()
    Debug.Print "服务器已停止"
    
    ' 更新 UI
    lblStatus.Caption = "已停止"
    
    ' 记录日志
    LogEvent "Server stopped"
    
    ' 停止定时任务
    Timer1.Enabled = False
End Sub
```

---

### OnClientConnect - 客户端连接

**语法**:

```vb
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 客户端唯一标识（格式：`Client#N`） |
| `RemoteAddress` | String | 客户端 IP 地址 |
| `RemotePort` | Long | 客户端端口 |

**示例**:

```vb
Private Sub m_Server_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)
    Debug.Print "客户端连接: " & ClientID & " (" & RemoteAddress & ":" & RemotePort & ")"
    
    ' 添加到客户端列表
    lstClients.AddItem ClientID & " - " & RemoteAddress
    
    ' 发送欢迎消息
    m_Server.SendText ClientID, "欢迎连接到 WebSocket 服务器！"
    
    ' 发送当前在线人数
    m_Server.SendText ClientID, "当前在线人数: " & m_Server.ClientCount
    
    ' 广播新用户上线
    m_Server.BroadcastText "[系统] " & ClientID & " 已上线", ClientID
    
    ' 记录连接日志
    LogConnection ClientID, RemoteAddress, RemotePort, "Connected"
    
    ' IP 黑白名单检查（实际应该在握手阶段）
    If IsBlacklisted(RemoteAddress) Then
        m_Server.DisconnectClient ClientID, WS_CLOSE_POLICY_VIOLATION, "IP 被拒绝"
        LogConnection ClientID, RemoteAddress, RemotePort, "Blocked (Blacklist)"
    End If
End Sub
```

---

### OnClientDisconnect - 客户端断开

**语法**:

```vb
Event OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 客户端 ID |
| `Reason` | String | 断开原因 |

**示例**:

```vb
Private Sub m_Server_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    Debug.Print "客户端断开: " & ClientID & " - " & Reason
    
    ' 从列表中移除
    Dim i As Long
    For i = 0 To lstClients.ListCount - 1
        If InStr(lstClients.List(i), ClientID) > 0 Then
            lstClients.RemoveItem i
            Exit For
        End If
    Next
    
    ' 更新连接数
    UpdateClientCount
    
    ' 广播用户下线
    m_Server.BroadcastText "[系统] " & ClientID & " 已离线"
    
    ' 记录断开日志
    LogDisconnect ClientID, Reason
    
    ' 如果是 VIP 用户，发送通知
    If IsVIPClient(ClientID) Then
        NotifyAdmin "VIP user " & ClientID & " disconnected: " & Reason
    End If
End Sub
```

---

### OnClientTextMessage - 收到客户端文本消息

**语法**:

```vb
Event OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 发送消息的客户端 ID |
| `Message` | String | 消息内容 |

**示例**:

```vb
Private Sub m_Server_OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
    Debug.Print "收到来自 " & ClientID & " 的消息: " & Message
    
    ' 记录消息日志
    LogMessage ClientID, Message
    
    ' 处理命令
    If Left$(Message, 1) = "/" Then
        ProcessCommand ClientID, Message
        Exit Sub
    End If
    
    ' 回显消息
    m_Server.SendText ClientID, "服务器收到: " & Message
    
    ' 广播给其他客户端（聊天模式）
    If m_bChatMode Then
        m_Server.BroadcastText ClientID & ": " & Message, ClientID
    End If
    
    ' 特殊命令：broadcast
    If LCase$(Message) = "broadcast" Then
        m_Server.BroadcastText "这是一条广播消息，来自客户端 " & ClientID, ClientID
    End If
End Sub

Private Sub ProcessCommand(ByVal ClientID As String, ByVal Command As String)
    Dim sCmd As String
    Dim sArgs() As String
    Dim sArgsList As String
    
    ' 解析命令
    sCmd = LCase$(Mid$(Command, 2))
    sArgsList = Mid$(Command, 2)
    sArgs = Split(sArgsList, " ")
    
    Select Case sCmd
        Case "users"
            ' 列出所有用户
            Dim sUserList As String
            sUserList = "在线用户: "
            Dim vIDs() As Variant
            vIDs = m_Server.ClientIDs
            Dim i As Long
            For i = LBound(vIDs) To UBound(vIDs)
                sUserList = sUserList & vIDs(i) & " "
            Next i
            m_Server.SendText ClientID, sUserList
            
        Case "time"
            ' 发送服务器时间
            m_Server.SendText ClientID, "服务器时间: " & Now
            
        Case "ping"
            ' 回复 Pong
            m_Server.SendText ClientID, "pong"
            
        Case Else
            m_Server.SendText ClientID, "未知命令: " & sCmd
    End Select
End Sub
```

---

### OnClientBinaryMessage - 收到客户端二进制消息

**语法**:

```vb
Event OnClientBinaryMessage(ByVal ClientID As String, Data() As Byte)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 发送消息的客户端 ID |
| `Data()` | Byte() | 二进制数据 |

**示例**:

```vb
Private Sub m_Server_OnClientBinaryMessage(ByVal ClientID As String, Data() As Byte)
    On Error Resume Next
    
    Debug.Print "收到来自 " & ClientID & " 的二进制消息: " & (UBound(Data) + 1) & " 字节"
    
    ' 记录二进制消息日志
    LogBinaryMessage ClientID, UBound(Data) + 1
    
    ' 检查数据类型（假设前 4 字节是类型标识）
    If UBound(Data) >= 3 Then
        Dim lType As Long
        lType = CLng(Data(0)) * 256& ^ 3 + CLng(Data(1)) * 256& ^ 2 + _
                CLng(Data(2)) * 256& + CLng(Data(3))
        
        Select Case lType
            Case 1 ' 图片上传
                SaveUploadedPicture ClientID, ExtractData(Data, 4)
                m_Server.SendText ClientID, "图片已保存"
                
            Case 2 ' 文件上传
                SaveUploadedFile ClientID, ExtractData(Data, 4)
                m_Server.SendText ClientID, "文件已保存"
                
            Case 3 ' 自定义数据
                ProcessCustomData ClientID, ExtractData(Data, 4)
                
            Case Else
                Debug.Print "未知数据类型: " & lType
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
Private Sub m_Server_OnError(ByVal Description As String)
    Debug.Print "服务器错误: " & Description
    
    ' 记录错误日志
    LogError Description
    
    ' 显示错误提示
    If m_bShowErrors Then
        MsgBox "服务器错误: " & Description, vbExclamation
    End If
    
    ' 严重错误时停止服务器
    If InStr(Description, "严重") > 0 Then
        m_Server.StopServer
    End If
End Sub
```

---

## 📝 完整示例

### 基本聊天服务器

```vb
Private WithEvents m_Server As cWebSocketServer
Private m_bChatMode As Boolean

Private Sub Form_Load()
    Set m_Server = New cWebSocketServer
    m_bChatMode = True
End Sub

Private Sub cmdStart_Click()
    m_Server.Listen CLng(txtPort.Text)
End Sub

Private Sub cmdStop_Click()
    m_Server.StopServer
End Sub

Private Sub m_Server_OnStart(ByVal Port As Long)
    Debug.Print "服务器已启动: " & Port
    lblStatus.Caption = "运行中"
End Sub

Private Sub m_Server_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)
    Debug.Print "客户端连接: " & ClientID
    lstClients.AddItem ClientID
    m_Server.SendText ClientID, "欢迎来到聊天室！"
    m_Server.BroadcastText ClientID & " 加入了聊天室", ClientID
End Sub

Private Sub m_Server_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    Debug.Print "客户端断开: " & ClientID
    Dim i As Long
    For i = 0 To lstClients.ListCount - 1
        If lstClients.List(i) = ClientID Then
            lstClients.RemoveItem i
            Exit For
        End If
    Next
    m_Server.BroadcastText ClientID & " 离开了聊天室"
End Sub

Private Sub m_Server_OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
    Debug.Print ClientID & ": " & Message
    txtLog.Text = txtLog.Text & ClientID & ": " & Message & vbCrLf
    
    ' 广播给所有其他客户端
    m_Server.BroadcastText ClientID & ": " & Message, ClientID
End Sub

Private Sub m_Server_OnError(ByVal Description As String)
    Debug.Print "错误: " & Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_Server.StopServer
End Sub
```

---

**最后更新**: 2026-01-10
