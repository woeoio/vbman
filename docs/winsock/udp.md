# cWinsock UDP 编程指南

## 📖 目录

- [概述](#概述)
- [UDP 基本编程](#udp-基本编程)
- [UDP 服务器虚拟客户端](#udp-服务器虚拟客户端)
- [UDP 广播与多播](#udp-广播与多播)
- [高级功能](#高级功能)
- [常见问题](#常见问题)

---

## 概述

UDP（User Datagram Protocol）是一种无连接的、不可靠的传输协议，适合对实时性要求高、可以容忍少量数据丢失的应用场景。

### UDP 特点

- ✅ **无连接**: 不需要建立连接
- ✅ **低延迟**: 无连接开销
- ✅ **简单高效**: 协议开销小
- ✅ **支持广播**: 可以同时发送给多个接收者
- ❌ **不可靠**: 不保证数据到达、顺序和完整性
- ❌ **无流量控制**: 可能导致网络拥塞

### 适用场景

- 实时游戏
- 视频直播
- 音频通话
- DNS 查询
- 网络发现
- IoT 设备通信

---

## UDP 基本编程

### UDP 客户端

```vb
' 声明 UDP 对象
Private WithEvents m_oUdp As cWinsock

' 初始化
Private Sub Form_Load()
    Set m_oUdp = New cWinsock
    m_oUdp.Protocol = sckUDPProtocol
    m_oUdp.Bind 0 ' 0 表示系统自动分配端口
    
    Debug.Print "UDP 客户端已绑定到端口: " & m_oUdp.LocalPort
End Sub

' 发送数据
Private Sub cmdSend_Click()
    ' 设置远程地址和端口
    m_oUdp.RemoteHost = "127.0.0.1"
    m_oUdp.RemotePort = 8888
    
    ' 发送数据
    m_oUdp.SendData "Hello, UDP!"
    
    Debug.Print "已发送到 " & m_oUdp.RemoteHost & ":" & m_oUdp.RemotePort
End Sub

' 接收数据
Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 接收数据
    Client.GetData sData
    
    Debug.Print "收到数据 (" & bytesTotal & " 字节): " & sData
    Debug.Print "来自: " & Client.RemoteHostIP & ":" & Client.RemotePort
End Sub

' 关闭
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    m_oUdp.Close_
End Sub
```

### UDP 服务器

```vb
' 声明 UDP 服务器对象
Private WithEvents m_oUdpServer As cWinsock

' 启动服务器
Private Sub cmdStart_Click()
    Set m_oUdpServer = New cWinsock
    m_oUdpServer.Protocol = sckUDPProtocol
    m_oUdpServer.Bind 8888
    
    Debug.Print "UDP 服务器已绑定到端口: 8888"
End Sub

' 数据到达
Private Sub m_oUdpServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 接收数据
    Client.GetData sData
    
    Debug.Print "收到来自 " & Client.RemoteHostIP & ":" & Client.RemotePort & " 的数据"
    Debug.Print "内容: " & sData
    
    ' 回复
    Client.SendData "Reply: " & sData
    
    Debug.Print "已回复"
End Sub

' 关闭
Private Sub cmdStop_Click()
    m_oUdpServer.Close_
    Debug.Print "UDP 服务器已停止"
End Sub
```

---

## UDP 服务器虚拟客户端

### 概述

UDP 是无连接协议，但 `cWinsock` 为每个不同的远程地址:端口组合创建虚拟客户端对象，模拟连接行为。

### 工作原理

```
首次收到来自 192.168.1.100:5000 的数据
    ↓
创建虚拟客户端对象，Tag = "192.168.1.100:5000"
    ↓
触发 ConnectionRequest 事件
    ↓
触发 DataArrival 事件
    ↓
可以向该虚拟客户端回复数据
```

### 示例代码

```vb
Private WithEvents m_oUdpServer As cWinsock

' 启动服务器
Private Sub cmdStart_Click()
    Set m_oUdpServer = New cWinsock
    m_oUdpServer.Protocol = sckUDPProtocol
    m_oUdpServer.Bind 8888
    
    Debug.Print "UDP 服务器已启动"
End Sub

' 新连接请求（首次收到某个地址:端口的数据）
Private Sub m_oUdpServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    Debug.Print "新 UDP 客户端: " & Client.RemoteHostIP & ":" & Client.RemotePort
    Debug.Print "Tag: " & Client.Tag
    
    ' 可以在这里进行拦截
    If IsInBlacklist(Client.RemoteHostIP) Then
        Debug.Print "拒绝黑名单 IP: " & Client.RemoteHostIP
        DisConnect = True
    End If
End Sub

' 数据到达
Private Sub m_oUdpServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 接收数据
    Client.GetData sData
    
    Debug.Print "来自 " & Client.Tag & " 的数据: " & sData
    
    ' 直接向该虚拟客户端回复
    Client.SendData "Echo: " & sData
End Sub

' 连接关闭（虚拟客户端超时）
Private Sub m_oUdpServer_CloseEvent(Client As cWinsock)
    Debug.Print "UDP 客户端 " & Client.Tag & " 已断开"
    
    ' 更新客户端列表
    UpdateClientList
End Sub

' 更新客户端列表
Private Sub UpdateClientList()
    lstClients.Clear
    lblCount.Caption = m_oUdpServer.ClientCount
    
    Dim oClient As cWinsock
    For Each oClient In m_oUdpServer.Clients
        lstClients.AddItem oClient.Tag & " - " & oClient.RemoteHostIP & ":" & oClient.RemotePort
    Next
End Sub
```

### 虚拟客户端超时处理

```vb
Private Sub tmrCleanup_Timer()
    Dim oClient As cWinsock
    Dim tSession As tSessionData
    
    For Each oClient In m_oUdpServer.Clients
        tSession = oClient.UserData
        
        ' 检查超时（5 分钟无活动）
        If DateDiff("s", tSession.LastActivity, Now) > 300 Then
            Debug.Print "清理超时客户端: " & oClient.Tag
            oClient.Close_
        End If
    Next
End Sub

Private Sub m_oUdpServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    ' 更新活动时间
    Dim tSession As tSessionData
    If IsEmpty(Client.UserData) Then
        tSession.FirstSeen = Now
    Else
        tSession = Client.UserData
    End If
    tSession.LastActivity = Now
    Client.UserData = tSession
    
    ' 处理数据...
    ProcessData Client, sData
End Sub
```

---

## UDP 广播与多播

### 广播

```vb
' 向局域网广播
Private Sub cmdBroadcast_Click()
    ' 绑定到任意端口
    m_oUdp.Bind 0
    
    ' 设置广播地址
    m_oUdp.RemoteHost = "255.255.255.255"
    m_oUdp.RemotePort = 9999
    
    ' 发送广播消息
    m_oUdp.SendData "Broadcast message"
    
    Debug.Print "已发送广播"
End Sub

' 接收广播
Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    Debug.Print "收到广播: " & sData
    Debug.Print "来自: " & Client.RemoteHostIP & ":" & Client.RemotePort
End Sub
```

### 局域网设备发现

```vb
' 发送发现请求
Private Sub cmdDiscover_Click()
    m_oUdp.RemoteHost = "255.255.255.255"
    m_oUdp.RemotePort = 9999
    m_oUdp.SendData "DISCOVER_SERVERS"
End Sub

' 服务器响应发现
Private Sub m_oUdpServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    If sData = "DISCOVER_SERVERS" Then
        ' 响应自己的信息
        Client.SendData "SERVER_INFO:" & GetLocalIP() & ":" & m_oUdpServer.LocalPort
    End If
End Sub

' 客户端收集响应
Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    If Left$(sData, 12) = "SERVER_INFO:" Then
        Dim sInfo As String
        sInfo = Mid$(sData, 13)
        
        Debug.Print "发现服务器: " & sInfo
        
        ' 添加到列表
        lstServers.AddItem sInfo
    End If
End Sub
```

### 多播（组播）

```vb
' 加入多播组
Private Sub JoinMulticastGroup(ByVal sGroupIP As String, ByVal lPort As Long)
    ' 注意：cWinsock 本身不直接支持多播
    ' 需要使用底层 socket API 或 cAsyncSocket
    
    ' 这里演示基本用法
    m_oUdp.Bind lPort
    m_oUdp.RemoteHost = sGroupIP
    m_oUdp.RemotePort = lPort
    
    Debug.Print "已加入多播组: " & sGroupIP
End Sub

' 发送到多播组
Private Sub cmdSendMulticast_Click()
    m_oUdp.SendData "Multicast message"
    
    Debug.Print "已发送到多播组"
End Sub
```

---

## 高级功能

### ✅ 可靠 UDP（带确认）

```vb
' 可靠 UDP 发送
Private Type tReliablePacket
    Sequence As Long      ' 序列号
    Total As Long         ' 总包数
    Index As Long         ' 当前包索引
    Data As String       ' 数据
    Acked As Boolean     ' 已确认
    Timestamp As Double  ' 发送时间
End Type

Private m_lSequence As Long
Private m_lWindowSize As Long

' 发送数据
Private Sub SendReliable(ByVal sData As String)
    Dim lChunkSize As Long
    lChunkSize = 1000 ' 每包 1KB
    
    Dim lTotalChunks As Long
    lTotalChunks = (Len(sData) \ lChunkSize) + 1
    
    Dim i As Long
    For i = 0 To lTotalChunks - 1
        Dim lStart As Long
        lStart = i * lChunkSize + 1
        
        Dim lEnd As Long
        lEnd = Min(lStart + lChunkSize - 1, Len(sData))
        
        Dim sChunk As String
        sChunk = Mid$(sData, lStart, lEnd - lStart + 1)
        
        ' 发送数据包
        m_oUdp.SendData "PKT:" & m_lSequence & ":" & lTotalChunks & ":" & i & ":" & sChunk
        
        Debug.Print "发送包 " & i + 1 & "/" & lTotalChunks
        
        m_lSequence = m_lSequence + 1
        
        ' 控制发送速率
        If (i + 1) Mod m_lWindowSize = 0 Then
            Sleep 50 ' 窗口满后等待
        End If
    Next
End Sub

' 接收并确认
Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    If Left$(sData, 4) = "PKT:" Then
        ' 解析数据包
        Dim sParts() As String
        sParts = Split(Mid$(sData, 5), ":")
        
        Dim lSeq As Long
        Dim lTotal As Long
        Dim lIndex As Long
        Dim sPayload As String
        
        lSeq = CLng(sParts(0))
        lTotal = CLng(sParts(1))
        lIndex = CLng(sParts(2))
        sPayload = sParts(3)
        
        Debug.Print "收到包 " & lIndex + 1 & "/" & lTotal & " (序列号: " & lSeq & ")"
        
        ' 发送确认
        Client.SendData "ACK:" & lSeq & ":" & lIndex
        
        ' 处理数据...
        ProcessPacket lSeq, lIndex, lTotal, sPayload
    ElseIf Left$(sData, 4) = "ACK:" Then
        ' 收到确认
        Dim sAckParts() As String
        sAckParts = Split(Mid$(sData, 5), ":")
        
        Dim lAckSeq As Long
        Dim lAckIndex As Long
        
        lAckSeq = CLng(sAckParts(0))
        lAckIndex = CLng(sAckParts(1))
        
        ' 标记为已确认
        MarkPacketAcked lAckSeq, lAckIndex
        
        Debug.Print "收到确认: " & lAckSeq & ":" & lAckIndex
    End If
End Sub
```

---

### 📊 数据包顺序重组

```vb
' 数据包重组
Private Type tReassemblyBuffer
    Packets() As String
    ReceivedCount As Long
    TotalCount As Long
End Type

Private m_oReassembly As Collection

' 初始化重组缓冲区
Private Sub InitReassembly(ByVal lSequence As Long, ByVal lTotal As Long)
    Dim tBuffer As tReassemblyBuffer
    
    ReDim tBuffer.Packets(0 To lTotal - 1) As String
    tBuffer.ReceivedCount = 0
    tBuffer.TotalCount = lTotal
    
    m_oReassembly.Add tBuffer, CStr(lSequence)
End Sub

' 添加数据包
Private Sub AddPacket(ByVal lSequence As Long, ByVal lIndex As Long, ByVal sData As String)
    Dim tBuffer As tReassemblyBuffer
    On Error Resume Next
    tBuffer = m_oReassembly(CStr(lSequence))
    On Error GoTo 0
    
    If tBuffer.TotalCount = 0 Then
        ' 新序列
        Debug.Print "新序列: " & lSequence
    End If
    
    ' 存储数据包
    tBuffer.Packets(lIndex) = sData
    tBuffer.ReceivedCount = tBuffer.ReceivedCount + 1
    
    ' 更新缓冲区
    If tBuffer.ReceivedCount = tBuffer.TotalCount Then
        ' 所有包都收到，重组数据
        Dim sComplete As String
        Dim i As Long
        
        For i = 0 To tBuffer.TotalCount - 1
            sComplete = sComplete & tBuffer.Packets(i)
        Next
        
        Debug.Print "序列 " & lSequence & " 重组完成，总大小: " & Len(sComplete)
        
        ' 处理完整数据
        ProcessCompleteData sComplete
        
        ' 删除缓冲区
        m_oReassembly.Remove CStr(lSequence)
    Else
        ' 更新缓冲区
        m_oReassembly.Remove CStr(lSequence)
        m_oReassembly.Add tBuffer, CStr(lSequence)
    End If
End Sub
```

---

### 🕐 超时重传

```vb
' 超时重传机制
Private Type tPendingPacket
    Sequence As Long
    Index As Long
    Data As String
    Timestamp As Double
    RetryCount As Long
End Type

Private m_oPendingPackets As Collection
Private Const PACKET_TIMEOUT As Double = 5 ' 5 秒超时
Private Const MAX_RETRIES As Long = 3

' 发送带超时的数据包
Private Sub SendWithTimeout(ByVal lSeq As Long, ByVal lIdx As Long, ByVal sData As String)
    ' 发送数据
    m_oUdp.SendData "PKT:" & lSeq & ":" & lIdx & ":" & sData
    
    ' 添加到待确认列表
    Dim tPending As tPendingPacket
    tPending.Sequence = lSeq
    tPending.Index = lIdx
    tPending.Data = sData
    tPending.Timestamp = Timer
    tPending.RetryCount = 0
    
    m_oPendingPackets.Add tPending, CStr(lSeq) & ":" & CStr(lIdx)
    
    Debug.Print "发送包 " & lSeq & ":" & lIdx & "，等待确认..."
End Sub

' 检查超时
Private Sub tmrTimeout_Timer()
    Dim i As Long
    For i = m_oPendingPackets.Count To 1 Step -1
        Dim tPending As tPendingPacket
        tPending = m_oPendingPackets(i)
        
        ' 检查是否超时
        If Timer - tPending.Timestamp > PACKET_TIMEOUT Then
            If tPending.RetryCount < MAX_RETRIES Then
                ' 重传
                Debug.Print "包 " & tPending.Sequence & ":" & tPending.Index & " 超时，重传..."
                
                ' 重发
                m_oUdp.SendData "PKT:" & tPending.Sequence & ":" & tPending.Index & ":" & tPending.Data
                
                ' 更新重试计数和时间
                tPending.RetryCount = tPending.RetryCount + 1
                tPending.Timestamp = Timer
                
                m_oPendingPackets.Remove i
                m_oPendingPackets.Add tPending, CStr(tPending.Sequence) & ":" & CStr(tPending.Index)
            Else
                ' 达到最大重试次数，放弃
                Debug.Print "包 " & tPending.Sequence & ":" & tPending.Index & " 超过最大重试次数，放弃"
                
                m_oPendingPackets.Remove i
            End If
        End If
    Next
End Sub

' 确认数据包
Private Sub AckPacket(ByVal lSeq As Long, ByVal lIdx As Long)
    Dim sKey As String
    sKey = CStr(lSeq) & ":" & CStr(lIdx)
    
    On Error Resume Next
    m_oPendingPackets.Remove sKey
    On Error GoTo 0
    
    Debug.Print "确认包 " & lSeq & ":" & lIdx
End Sub
```

---

## 常见问题

### ❓ 问题 1: UDP 数据包丢失

**现象**: 发送的数据对方没有收到

**原因**: UDP 是不可靠传输，数据包可能丢失

**解决方案**: 实现确认重传机制（如上所示）

---

### ❓ 问题 2: 数据包顺序错乱

**现象**: 接收到的数据包顺序与发送顺序不一致

**原因**: UDP 不保证顺序

**解决方案**: 实现序列号和重组机制

---

### ❓ 问题 3: 数据包重复

**现象**: 收到重复的数据包

**原因**: 重传导致

**解决方案**: 根据序列号去重

```vb
Private Function IsDuplicate(ByVal lSequence As Long) As Boolean
    Static lLastSequence As Long
    
    If lSequence <= lLastSequence Then
        IsDuplicate = True
    Else
        IsDuplicate = False
        lLastSequence = lSequence
    End If
End Function
```

---

### ❓ 问题 4: 广播失败

**现象**: 发送广播后没有收到响应

**原因**: 防火墙阻止广播

**解决方案**: 配置防火墙或使用特定端口

```vb
' 尝试不同的广播地址
m_oUdp.RemoteHost = "192.168.1.255" ' 子网广播
m_oUdp.SendData "Broadcast"
```

---

### ❓ 问题 5: 虚拟客户端未清理

**现象**: UDP 服务器客户端列表持续增长

**原因**: UDP 无连接，客户端断开时无法自动检测

**解决方案**: 实现超时清理机制

```vb
Private Sub tmrCleanup_Timer()
    Dim oClient As cWinsock
    Dim tSession As tSessionData
    
    For Each oClient In m_oUdpServer.Clients
        tSession = oClient.UserData
        
        ' 检查超时
        If DateDiff("s", tSession.LastActivity, Now) > 300 Then
            Debug.Print "清理超时客户端: " & oClient.Tag
            oClient.Close_
        End If
    Next
End Sub
```

---

**最后更新**: 2026-01-09
