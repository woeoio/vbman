# cWinsock TCP 编程指南

## 📖 目录

- [概述](#概述)
- [TCP 客户端编程](#tcp-客户端编程)
- [TCP 服务器编程](#tcp-服务器编程)
- [高级功能](#高级功能)
- [常见问题](#常见问题)

---

## 概述

TCP（Transmission Control Protocol）是一种面向连接的、可靠的传输协议，适合需要保证数据完整性和顺序的应用场景。

### TCP 特点

- ✅ **面向连接**: 需要先建立连接
- ✅ **可靠传输**: 保证数据到达、顺序和完整性
- ✅ **流量控制**: 避免网络拥塞
- ✅ **拥塞控制**: 自动调整传输速率
- ❌ **开销较大**: 相比 UDP 有更多协议开销

### 适用场景

- 文件传输
- 聊天应用
- 远程控制
- 数据库连接
- Web 服务

---

## TCP 客户端编程

### 基本流程

```
1. 创建 cWinsock 对象
2. 设置协议为 TCP
3. 连接到服务器
4. 等待 Connect 事件
5. 发送/接收数据
6. 关闭连接
```

### 完整示例

```vb
' 声明客户端对象
Private WithEvents m_oClient As cWinsock

' 连接按钮
Private Sub cmdConnect_Click()
    On Error GoTo EH
    
    If m_oClient Is Nothing Then
        Set m_oClient = New cWinsock
    End If
    
    ' 设置协议
    m_oClient.Protocol = sckTCPProtocol
    
    ' 连接服务器
    m_oClient.Connect txtHost.Text, CLng(txtPort.Text)
    
    ' 更新 UI
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = False
    lblStatus.Caption = "连接中..."
    
    Exit Sub
    
EH:
    Debug.Print "连接错误: " & Err.Description
    lblStatus.Caption = "连接失败"
End Sub

' 连接成功事件
Private Sub m_oClient_Connect(Client As cWinsock)
    Debug.Print "已连接到 " & Client.RemoteHostIP & ":" & Client.RemotePort
    
    ' 更新 UI
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    lblStatus.Caption = "已连接"
End Sub

' 断开连接按钮
Private Sub cmdDisconnect_Click()
    If Not m_oClient Is Nothing Then
        m_oClient.Close_
    End If
    
    ' 更新 UI
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    lblStatus.Caption = "已断开"
End Sub

' 发送数据按钮
Private Sub cmdSend_Click()
    On Error GoTo EH
    
    If Not m_oClient Is Nothing And m_oClient.State = sckConnected Then
        m_oClient.SendData txtSend.Text
        Debug.Print "已发送: " & txtSend.Text
        
        ' 清空输入框
        txtSend.Text = ""
    Else
        MsgBox "未连接", vbExclamation
    End If
    
    Exit Sub
    
EH:
    Debug.Print "发送错误: " & Err.Description
End Sub

' 数据到达事件
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 接收数据
    Client.GetData sData
    
    Debug.Print "收到数据 (" & bytesTotal & " 字节): " & sData
    
    ' 显示在界面上
    txtReceive.SelStart = Len(txtReceive.Text)
    txtReceive.SelText = sData & vbCrLf
    txtReceive.SelStart = Len(txtReceive.Text)
End Sub

' 连接关闭事件
Private Sub m_oClient_CloseEvent(Client As cWinsock)
    Debug.Print "连接已关闭"
    
    ' 更新 UI
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    lblStatus.Caption = "连接已关闭"
End Sub

' 错误事件
Private Sub m_oClient_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Debug.Print "错误 [" & Number & "]: " & Description
    
    ' 更新 UI
    lblStatus.Caption = "错误: " & Description
End Sub

' 窗体卸载
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    m_oClient.Close_
End Sub
```

---

## TCP 服务器编程

### 基本流程

```
1. 创建 cWinsock 对象
2. 设置协议为 TCP
3. 监听端口
4. 等待 ConnectionRequest 事件
5. 接受或拒绝连接
6. 通过 DataArrival 处理客户端数据
7. 客户端断开时清理
```

### 完整示例

```vb
' 声明服务器对象
Private WithEvents m_oServer As cWinsock

' 启动服务器按钮
Private Sub cmdStart_Click()
    On Error GoTo EH
    
    If m_oServer Is Nothing Then
        Set m_oServer = New cWinsock
    End If
    
    ' 设置协议
    m_oServer.Protocol = sckTCPProtocol
    
    ' 开始监听
    m_oServer.Listen CLng(txtPort.Text)
    
    Debug.Print "服务器已启动，监听端口: " & m_oServer.LocalPort
    
    ' 更新 UI
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    lblStatus.Caption = "监听中..."
    lblClientCount.Caption = "0"
    
    Exit Sub
    
EH:
    Debug.Print "启动服务器失败: " & Err.Description
    MsgBox "无法启动服务器: " & Err.Description, vbExclamation
End Sub

' 停止服务器按钮
Private Sub cmdStop_Click()
    On Error Resume Next
    
    If Not m_oServer Is Nothing Then
        m_oServer.Close_
    End If
    
    Debug.Print "服务器已停止"
    
    ' 更新 UI
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    lblStatus.Caption = "已停止"
    lstClients.Clear
End Sub

' 新连接请求事件
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    Debug.Print "新客户端连接: " & Client.RemoteHostIP & ":" & Client.RemotePort
    
    ' 检查连接数限制
    If m_oServer.ClientCount >= 100 Then
        Debug.Print "达到最大连接数，拒绝连接"
        DisConnect = True
        Exit Sub
    End If
    
    ' IP 黑名单检查
    If IsInBlacklist(Client.RemoteHostIP) Then
        Debug.Print "IP 在黑名单中，拒绝连接: " & Client.RemoteHostIP
        DisConnect = True
        Exit Sub
    End If
    
    ' 接受连接（DisConnect = False）
    Debug.Print "接受连接: " & Client.Tag
    
    ' 更新客户端列表
    UpdateClientList
End Sub

' 数据到达事件（所有客户端数据通过此事件触发）
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 接收数据
    Client.GetData sData
    
    Debug.Print "来自 " & Client.Tag & " (" & Client.RemoteHostIP & ") 的数据: " & sData
    
    ' 显示在日志中
    LogMessage Client.Tag & ": " & sData
    
    ' 回显给客户端
    Client.SendData "Echo: " & sData
End Sub

' 连接关闭事件
Private Sub m_oServer_CloseEvent(Client As cWinsock)
    Debug.Print "客户端 " & Client.Tag & " 已断开"
    
    ' 更新客户端列表
    UpdateClientList
End Sub

' 错误事件
Private Sub m_oServer_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Debug.Print "错误 [" & Number & "]: " & Description
    
    If Client.IsServer Then
        ' 服务器错误
        LogMessage "服务器错误: " & Description
    Else
        ' 客户端错误
        LogMessage "客户端 " & Client.Tag & " 错误: " & Description
    End If
End Sub

' 更新客户端列表
Private Sub UpdateClientList()
    lstClients.Clear
    lblClientCount.Caption = m_oServer.ClientCount
    
    Dim oClient As cWinsock
    For Each oClient In m_oServer.Clients
        lstClients.AddItem oClient.Tag & " - " & oClient.RemoteHostIP & ":" & oClient.RemotePort
    Next
End Sub

' 添加日志
Private Sub LogMessage(ByVal sMsg As String)
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelText = Format$(Now, "hh:mm:ss") & " - " & sMsg & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

' 窗体卸载
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    m_oServer.Close_
End Sub

' 黑名单检查
Private Function IsInBlacklist(ByVal sIP As String) As Boolean
    ' 从配置文件或数据库加载黑名单
    ' 这里简化演示
    IsInBlacklist = False
End Function
```

---

## 高级功能

### 🔄 心跳检测

```vb
' 服务器端心跳检测
Private Const HEARTBEAT_INTERVAL As Long = 30 ' 30 秒

Private Sub tmrHeartbeat_Timer()
    Dim oClient As cWinsock
    Dim tSession As tSessionData
    
    For Each oClient In m_oServer.Clients
        tSession = oClient.UserData
        
        ' 检查是否超时
        If DateDiff("s", tSession.LastActivity, Now) > HEARTBEAT_INTERVAL Then
            Debug.Print "客户端 " & oClient.Tag & " 心跳超时"
            oClient.Close_
        Else
            ' 发送心跳请求
            oClient.SendData "PING"
        End If
    Next
End Sub

' 客户端响应心跳
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    If sData = "PING" Then
        ' 响应心跳
        Client.SendData "PONG"
    ElseIf sData = "PONG" Then
        Debug.Print "收到心跳响应"
    Else
        ' 处理其他数据
        ProcessData sData
    End If
End Sub
```

---

### 📦 协议封装

```vb
' 定义协议头部
Private Type tPacketHeader
    Magic As Long       ' 魔数：&H12345678
    Length As Long      ' 数据长度
    Type As Long       ' 数据类型：1=文本，2=二进制
    Checksum As Long   ' 校验和
End Type

' 发送结构化数据
Private Sub SendPacket(ByVal oSocket As cWinsock, ByVal eType As Long, ByVal baData() As Byte)
    Dim tHeader As tPacketHeader
    Dim baPacket() As Byte
    Dim lOffset As Long
    
    ' 填充头部
    tHeader.Magic = &H12345678
    tHeader.Length = UBound(baData) + 1
    tHeader.Type = eType
    tHeader.Checksum = CalculateChecksum(baData)
    
    ' 构造完整数据包
    ReDim baPacket(0 To Len(tHeader) + tHeader.Length - 1) As Byte
    
    ' 复制头部
    CopyMemory baPacket(0), tHeader, Len(tHeader)
    
    ' 复制数据体
    lOffset = Len(tHeader)
    CopyMemory baPacket(lOffset), baData(0), tHeader.Length
    
    ' 发送
    oSocket.SendData baPacket
End Sub

' 接收结构化数据
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim tHeader As tPacketHeader
    Dim baData() As Byte
    
    ' 先读取头部
    Client.GetData tHeader
    
    ' 验证魔数
    If tHeader.Magic <> &H12345678 Then
        Debug.Print "无效的数据包"
        Exit Sub
    End If
    
    ' 验证校验和
    ReDim baData(0 To tHeader.Length - 1) As Byte
    Client.GetData baData
    
    If CalculateChecksum(baData) <> tHeader.Checksum Then
        Debug.Print "校验和错误"
        Exit Sub
    End If
    
    ' 根据类型处理数据
    Select Case tHeader.Type
        Case 1 ' 文本
            Dim sText As String
            sText = BytesToString(baData)
            Debug.Print "文本数据: " & sText
            
        Case 2 ' 二进制
            Debug.Print "二进制数据: " & tHeader.Length & " 字节"
            
    End Select
End Sub
```

---

### 🔄 自动重连

```vb
' 自动重连客户端
Private WithEvents m_oClient As cWinsock
Private m_bAutoReconnect As Boolean
Private m_lReconnectInterval As Long

Private Sub StartClient()
    Set m_oClient = New cWinsock
    m_oClient.Protocol = sckTCPProtocol
    m_bAutoReconnect = True
    m_lReconnectInterval = 5 ' 5 秒
    
    ConnectToServer
End Sub

Private Sub ConnectToServer()
    On Error GoTo EH
    
    m_oClient.Connect "127.0.0.1", 8080
    Debug.Print "正在连接..."
    
    Exit Sub
    
EH:
    Debug.Print "连接失败: " & Err.Description
    
    If m_bAutoReconnect Then
        Debug.Print m_lReconnectInterval & " 秒后重连..."
        tmrReconnect.Interval = m_lReconnectInterval * 1000
        tmrReconnect.Enabled = True
    End If
End Sub

Private Sub m_oClient_CloseEvent(Client As cWinsock)
    Debug.Print "连接已关闭"
    
    If m_bAutoReconnect Then
        Debug.Print m_lReconnectInterval & " 秒后重连..."
        tmrReconnect.Interval = m_lReconnectInterval * 1000
        tmrReconnect.Enabled = True
    End If
End Sub

Private Sub tmrReconnect_Timer()
    tmrReconnect.Enabled = False
    ConnectToServer
End Sub
```

---

### 🚦 流量控制

```vb
' 带流量控制的文件传输
Private m_oClient As cWinsock
Private m_lChunkSize As Long
Private m_bSending As Boolean

Private Sub SendFile(ByVal sFilePath As String)
    Dim iFileNum As Integer
    Dim baChunk() As Byte
    Dim lFileSize As Long
    Dim lSent As Long
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
    lFileSize = LOF(iFileNum)
    
    m_lChunkSize = 8192 ' 8KB 每块
    lSent = 0
    m_bSending = True
    
    Do While lSent < lFileSize And m_bSending
        ' 读取数据块
        ReDim baChunk(0 To m_lChunkSize - 1) As Byte
        Get #iFileNum, , baChunk
        
        ' 发送
        m_oClient.SendData baChunk
        lSent = lSent + m_lChunkSize
        
        ' 更新进度
        UpdateProgress lSent, lFileSize
        
        ' 等待发送完成
        Do While m_bSending
            DoEvents
            If Not m_oClient.State = sckConnected Then Exit Do
        Loop
        
        If Not m_oClient.State = sckConnected Then Exit Do
    Loop
    
    Close #iFileNum
    
    Debug.Print "文件传输完成"
End Sub

Private Sub m_oClient_SendComplete(Client As cWinsock)
    m_bSending = False
End Sub
```

---

## 常见问题

### ❓ 问题 1: 连接被拒绝

**现象**: `Error 10061` - 连接被拒绝

**原因**: 
- 服务器未启动
- 端口被防火墙阻止
- IP 地址或端口错误

**解决方案**:
```vb
' 检查服务器是否启动
If m_oServer.State <> sckListening Then
    MsgBox "服务器未启动", vbExclamation
    Exit Sub
End If

' 检查端口
If Not IsPortOpen(txtPort.Text) Then
    MsgBox "端口 " & txtPort.Text & " 未开放", vbExclamation
End If
```

---

### ❓ 问题 2: 连接超时

**现象**: `Error 10060` - 连接超时

**原因**: 
- 网络不通
- 服务器响应慢
- 防火墙阻止

**解决方案**:
```vb
' 增加重试机制
Private Function ConnectWithRetry(ByVal sHost As String, ByVal lPort As Long, ByVal lRetries As Long) As Boolean
    Dim i As Long
    
    For i = 1 To lRetries
        On Error Resume Next
        m_oClient.Connect sHost, lPort
        
        If Err.Number = 0 Then
            ConnectWithRetry = True
            Exit Function
        End If
        
        Debug.Print "重连 " & i & "/" & lRetries & " 失败: " & Err.Description
        
        ' 等待后重试
        Sleep 2000
    Next
    
    ConnectWithRetry = False
End Function
```

---

### ❓ 问题 3: 数据丢失

**现象**: 发送的数据对方没有收到

**原因**: 
- 网络问题
- 缓冲区溢出
- 对方未正确读取

**解决方案**:
```vb
' 确认对方收到
Private Sub SendWithAck(ByVal sData As String)
    Dim sAck As String
    
    ' 发送数据
    m_oClient.SendData "DATA:" & sData
    
    ' 等待确认
    sAck = WaitForAck(5000) ' 5 秒超时
    
    If sAck = "ACK" Then
        Debug.Print "对方已确认收到"
    Else
        Debug.Print "未收到确认，重发"
        m_oClient.SendData "DATA:" & sData
    End If
End Sub
```

---

**最后更新**: 2026-01-09
