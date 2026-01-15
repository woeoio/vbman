# cWinsock 事件详解

## 📋 事件列表

| 事件名 | 说明 | 触发时机 |
|--------|------|----------|
| `Connect` | 客户端连接成功 | TCP 客户端成功连接到服务器 |
| `CloseEvent` | 连接关闭 | TCP 连接被关闭 |
| `ConnectionRequest` | 新连接请求 | 服务器收到新的连接请求 |
| `DataArrival` | 数据到达 | 接收到新数据 |
| `SendProgress` | 发送进度 | 数据发送过程中触发 |
| `SendComplete` | 发送完成 | 数据发送完成 |
| `Error` | 发生错误 | 发生 Socket 错误 |

---

## 🔗 Connect 事件

### 说明

当 TCP 客户端成功连接到服务器时触发。

### 语法

```vb
Private Sub object_Connect(Client As cWinsock)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Client` | cWinsock | 触发事件的客户端对象（即 Me） |

### 使用示例

```vb
Private WithEvents m_oClient As cWinsock

Private Sub m_oClient_Connect(Client As cWinsock)
    Debug.Print "已连接到服务器"
    Debug.Print "远程地址: " & Client.RemoteHostIP
    Debug.Print "远程端口: " & Client.RemotePort
    
    ' 连接成功后发送登录请求
    Client.SendData "LOGIN|user|password"
End Sub
```

---

## 🚪 CloseEvent 事件

### 说明

当 TCP 连接被关闭时触发。

### 语法

```vb
Private Sub object_CloseEvent(Client As cWinsock)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Client` | cWinsock | 触发事件的客户端对象 |

### 重要说明

对于服务器接受的客户端，`CloseEvent` 触发时会自动：
1. 从服务器的 `Clients` 集合中移除该客户端
2. 清理相关资源

### 使用示例

```vb
Private Sub m_oClient_CloseEvent(Client As cWinsock)
    Debug.Print "连接已关闭"
    
    ' 可以在这里尝试重连
    If m_bAutoReconnect Then
        Debug.Print "3 秒后尝试重连..."
        tmrReconnect.Enabled = True
    End If
End Sub
```

**服务器端示例：**
```vb
Private Sub m_oServer_CloseEvent(Client As cWinsock)
    Debug.Print "客户端 " & Client.Tag & " 已断开连接"
    
    ' 更新 UI
    Dim i As Long
    For i = 0 To lstClients.ListCount - 1
        If lstClients.List(i) = Client.Tag Then
            lstClients.RemoveItem i
            Exit For
        End If
    Next
    
    ' 更新统计
    UpdateClientCount
End Sub
```

---

## 🔔 ConnectionRequest 事件

### 说明

服务器收到新的连接请求时触发。支持通过 `DisConnect` 参数拦截连接。

### 语法

```vb
Private Sub object_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Client` | cWinsock | 新的客户端对象 |
| `DisConnect` | Boolean | 设置为 `True` 可拒绝连接并清理资源 |

### 连接拦截机制

```vb
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 默认接受连接（DisConnect = False）
    
    ' 1. IP 黑名单检查
    If IsInBlacklist(Client.RemoteHostIP) Then
        Debug.Print "拒绝黑名单 IP: " & Client.RemoteHostIP
        DisConnect = True
        Exit Sub
    End If
    
    ' 2. 端口范围限制
    If Client.RemotePort < 1024 Then
        Debug.Print "拒绝特权端口: " & Client.RemotePort
        DisConnect = True
        Exit Sub
    End If
    
    ' 3. 白名单模式
    If m_bWhitelistMode Then
        If Not IsInWhitelist(Client.RemoteHostIP) Then
            Debug.Print "IP 不在白名单中: " & Client.RemoteHostIP
            DisConnect = True
            Exit Sub
        End If
    End If
    
    ' 4. 连接数限制
    If m_oServer.ClientCount >= m_lMaxClients Then
        Debug.Print "达到最大连接数限制"
        DisConnect = True
        Exit Sub
    End If
    
    ' 接受连接
    Debug.Print "接受新客户端: " & Client.RemoteHostIP & ":" & Client.RemotePort
    DisConnect = False
End Sub
```

### 高级示例：动态白名单

```vb
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 从数据库或配置文件加载白名单
    Dim sWhitelist() As String
    sWhitelist = LoadWhitelistFromDatabase()
    
    Dim bAllowed As Boolean
    bAllowed = False
    
    Dim i As Long
    For i = LBound(sWhitelist) To UBound(sWhitelist)
        If sWhitelist(i) = Client.RemoteHostIP Then
            bAllowed = True
            Exit For
        End If
    Next
    
    If Not bAllowed Then
        Debug.Print "拒绝未授权 IP: " & Client.RemoteHostIP
        DisConnect = True
    End If
End Sub
```

### TCP vs UDP 中的触发

| 协议 | 触发时机 |
|------|----------|
| TCP | 收到新的连接请求（`accept` 系统调用） |
| UDP | 首次收到来自新地址:端口的数据包 |

---

## 📨 DataArrival 事件

### 说明

当接收到新数据时触发。**这是最常用的事件之一**。

### 语法

```vb
Private Sub object_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Client` | cWinsock | 接收数据的客户端对象 |
| `bytesTotal` | Long | 可用数据字节数 |

### 基本使用

```vb
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 读取字符串数据
    Client.GetData sData
    
    Debug.Print "收到 " & bytesTotal & " 字节: " & sData
    
    ' 处理数据...
    ProcessData Client, sData
End Sub
```

### 读取字节数组

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    
    ' 读取字节数组
    Client.GetData baData
    
    Debug.Print "收到 " & bytesTotal & " 字节数据"
    
    ' 处理二进制数据...
    ProcessBinaryData baData
End Sub
```

### 部分读取数据

```vb
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sHeader As String
    Dim sBody As String
    
    ' 先读取前 10 字节作为头部
    Client.GetData sHeader, vbString, 10
    Debug.Print "头部: " & sHeader
    
    ' 读取剩余数据（仍在缓冲区中）
    Client.GetData sBody
    Debug.Print "正文: " & sBody
End Sub
```

### 指定编码读取

```vb
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    
    ' 使用 UTF-8 编码读取
    Client.GetData sData, vbString, -1, ucsScpUtf8
    
    Debug.Print "UTF-8 数据: " & sData
End Sub
```

### 事件代理机制

**重要：** 服务器接受的客户端对象，其 `DataArrival` 事件会通过父服务器对象触发。

```vb
' 只需订阅服务器的事件，即可处理所有客户端的数据
Private WithEvents m_oServer As cWinsock

Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    ' Client 参数就是具体的客户端对象
    Debug.Print "来自 " & Client.Tag & " 的数据"
    
    Dim sData As String
    Client.GetData sData
    
    ' 可以直接向该客户端回复
    Client.SendData "Echo: " & sData
End Sub
```

---

## 📊 SendProgress 事件

### 说明

数据发送过程中定期触发，用于显示发送进度。

### 语法

```vb
Private Sub object_SendProgress(Client As cWinsock, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Client` | cWinsock | 发送数据的客户端对象 |
| `bytesSent` | Long | 已发送的字节数 |
| `bytesRemaining` | Long | 剩余待发送的字节数 |

### 使用示例

```vb
Private Sub m_oClient_SendProgress(Client As cWinsock, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Dim lTotal As Long
    lTotal = bytesSent + bytesRemaining
    
    Dim dPercent As Double
    dPercent = (bytesSent / lTotal) * 100
    
    Debug.Print "发送进度: " & Format$(dPercent, "0.00") & "% (" & bytesSent & "/" & lTotal & ")"
    
    ' 更新进度条
    If Not prgProgress Is Nothing Then
        prgProgress.Value = CInt(dPercent)
    End If
End Sub
```

### 实际应用：文件传输进度

```vb
Private Sub m_oClient_SendProgress(Client As cWinsock, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Static lStartTime As Long
    Static lLastUpdate As Long
    
    If lStartTime = 0 Then lStartTime = Timer
    If lLastUpdate = 0 Then lLastUpdate = lStartTime
    
    ' 每 0.5 秒更新一次 UI
    If Timer - lLastUpdate >= 0.5 Then
        Dim lTotal As Long
        lTotal = bytesSent + bytesRemaining
        
        Dim dElapsed As Double
        dElapsed = Timer - lStartTime
        
        Dim dSpeed As Double
        dSpeed = bytesSent / dElapsed ' 字节/秒
        
        ' 更新 UI
        lblStatus.Caption = "发送中: " & FormatSize(bytesSent) & " / " & FormatSize(lTotal)
        lblSpeed.Caption = "速度: " & FormatSize(dSpeed) & "/s"
        
        lLastUpdate = Timer
    End If
End Sub
```

---

## ✅ SendComplete 事件

### 说明

数据发送完成时触发。

### 语法

```vb
Private Sub object_SendComplete(Client As cWinsock)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Client` | cWinsock | 发送完成的客户端对象 |

### 使用示例

```vb
Private Sub m_oClient_SendComplete(Client As cWinsock)
    Debug.Print "数据发送完成"
    
    ' 重置发送状态
    m_bSending = False
    
    ' 更新 UI
    cmdSend.Enabled = True
    lblStatus.Caption = "就绪"
End Sub
```

### 实际应用：命令队列

```vb
Private m_lCommandQueue() As String
Private m_lQueueIndex As Long

Private Sub SendNextCommand()
    If m_lQueueIndex <= UBound(m_lCommandQueue) Then
        m_oClient.SendData m_lCommandQueue(m_lQueueIndex)
        m_lQueueIndex = m_lQueueIndex + 1
    End If
End Sub

Private Sub m_oClient_SendComplete(Client As cWinsock)
    Debug.Print "命令发送完成，发送下一个..."
    SendNextCommand
End Sub
```

---

## ❌ Error 事件

### 说明

发生 Socket 错误时触发。

### 语法

```vb
Private Sub object_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
```

### 参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `Client` | cWinsock | 发生错误的客户端对象 |
| `Number` | Long | 错误代码 |
| `Description` | String | 错误描述 |
| `Scode` | Long | SCODE（通常与 Number 相同） |

### 常见错误代码

| 错误代码 | 说明 |
|----------|------|
| 10053 | 连接被远程主机强制关闭 |
| 10054 | 远程主机关闭了连接 |
| 10060 | 连接超时 |
| 10061 | 连接被拒绝 |
| 10065 | 无法到达目标主机 |
| 10048 | 地址已被使用 |

### 使用示例

```vb
Private Sub m_oClient_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Debug.Print "Socket 错误 [" & Number & "]: " & Description
    
    Select Case Number
        Case 10053, 10054
            ' 连接被关闭
            Debug.Print "远程主机已断开连接"
            
        Case 10060
            ' 连接超时
            Debug.Print "连接超时，请检查网络"
            
        Case 10061
            ' 连接被拒绝
            Debug.Print "服务器拒绝连接，请检查端口和防火墙"
            
        Case Else
            ' 其他错误
            Debug.Print "未知错误: " & Description
    End Select
End Sub
```

### 错误恢复

```vb
Private Sub m_oServer_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Debug.Print "服务器错误 [" & Number & "]: " & Description
    
    ' 移除出错的客户端
    If Not Client Is Nothing Then
        m_oServer.RemoveClient Client
    End If
    
    ' 如果是严重错误，重启服务器
    If Number >= 10000 Then
        Debug.Print "严重错误，重启服务器..."
        m_oServer.Close_
        m_oServer.Listen m_lServerPort
    End If
End Sub
```

---

## 🎯 事件触发顺序

### TCP 客户端连接流程

```
1. Connect() 调用
2. 内部解析主机名
3. Connect 事件触发
4. 可以开始发送数据
5. SendProgress 事件（多次）
6. SendComplete 事件触发
7. DataArrival 事件（接收数据）
8. CloseEvent 事件（连接关闭）
```

### TCP 服务器接受连接流程

```
1. Listen() 调用
2. 收到新连接请求
3. ConnectionRequest 事件触发（可拦截）
4. 如果接受，创建客户端对象
5. 客户端 DataArrival 事件（通过服务器触发）
6. 客户端 CloseEvent 事件（断开时触发）
```

### UDP 通信流程

```
1. Bind() 调用（UDP 服务器）
2. 收到数据包
3. 首次收到该地址:端口 → ConnectionRequest 事件
4. DataArrival 事件触发
5. 可以通过 SendData 回复
```

---

## 📌 注意事项

1. **事件处理时间**
   - 避免在事件处理中执行耗时操作
   - 使用 `DoEvents` 释放控制权
   - 或将耗时操作放入队列异步处理

2. **对象生命周期**
   - 不要在事件处理中 `Set Client = Nothing`
   - 客户端对象由服务器管理，自动清理

3. **线程安全**
   - 事件在主线程触发，可直接访问 UI
   - 但避免重入问题

4. **错误处理**
   - 总是使用 `On Error GoTo` 处理事件中的错误
   - 防止一个客户端的错误影响其他客户端

---

**最后更新**: 2026-01-09
