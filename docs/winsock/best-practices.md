# cWinsock 最佳实践

## 📖 目录

- [性能优化](#性能优化)
- [错误处理](#错误处理)
- [安全建议](#安全建议)
- [调试技巧](#调试技巧)
- [常见陷阱](#常见陷阱)

---

## 性能优化

### 1️⃣ 事件处理优化

避免在事件处理中执行耗时操作

```vb
' ❌ 错误：在事件中处理大量数据
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    ' 在事件中进行耗时操作
    ProcessLargeData sData ' 可能耗时很久
    SaveToDatabase sData     ' 可能超时
End Sub

' ✅ 正确：将耗时操作放入队列
Private m_oWorkQueue As Collection

Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    ' 添加到工作队列
    m_oWorkQueue.Add Array(Client.Tag, sData)
    
    ' 定时器处理队列
    tmrWorkQueue_Timer
End Sub

Private Sub tmrWorkQueue_Timer()
    Dim vItem As Variant
    Dim sTag As String
    Dim sData As String
    
    If m_oWorkQueue.Count > 0 Then
        vItem = m_oWorkQueue(1)
        sTag = vItem(0)
        sData = vItem(1)
        
        ' 处理数据
        ProcessData sTag, sData
        
        ' 从队列移除
        m_oWorkQueue.Remove 1
    End If
End Sub
```

---

### 2️⃣ 缓冲区大小优化

根据应用场景调整缓冲区大小

```vb
' 小数据频繁传输
Private Const SMALL_BUFFER_SIZE As Long = 1024 ' 1KB

' 大数据块传输
Private Const LARGE_BUFFER_SIZE As Long = 65536 ' 64KB

' 文件传输
Private Const FILE_CHUNK_SIZE As Long = 8192 ' 8KB

' 使用示例
Private Sub SendOptimal(ByVal sData As String)
    Dim lSize As Long
    lSize = Len(sData)
    
    If lSize < SMALL_BUFFER_SIZE Then
        ' 小数据，直接发送
        m_oClient.SendData sData
    Else
        ' 大数据，分块发送
        Dim lOffset As Long
        lOffset = 1
        
        Do While lOffset <= lSize
            Dim sChunk As String
            sChunk = Mid$(sData, lOffset, FILE_CHUNK_SIZE)
            
            m_oClient.SendData sChunk
            lOffset = lOffset + FILE_CHUNK_SIZE
            
            ' 等待发送完成
            Do While m_bSending
                DoEvents
            Loop
        Loop
    End If
End Sub
```

---

### 3️⃣ 连接池管理

对于需要频繁建立连接的场景，使用连接池

```vb
' 连接池类
Private Type tConnection
    Socket As cWinsock
    InUse As Boolean
    LastUsed As Date
End Type

Private m_oConnections() As tConnection
Private m_lPoolSize As Long

' 初始化连接池
Private Sub InitConnectionPool(ByVal lSize As Long)
    ReDim m_oConnections(0 To lSize - 1) As tConnection
    m_lPoolSize = lSize
    
    Dim i As Long
    For i = 0 To lSize - 1
        Set m_oConnections(i).Socket = New cWinsock
        m_oConnections(i).InUse = False
        m_oConnections(i).LastUsed = Now
    Next
End Sub

' 获取连接
Private Function GetConnection() As cWinsock
    Dim i As Long
    
    ' 查找可用连接
    For i = 0 To m_lPoolSize - 1
        If Not m_oConnections(i).InUse Then
            If m_oConnections(i).Socket.State = sckConnected Then
                m_oConnections(i).InUse = True
                Set GetConnection = m_oConnections(i).Socket
                Exit Function
            End If
        End If
    Next
    
    ' 没有可用连接，返回 Nothing
    Set GetConnection = Nothing
End Function

' 释放连接
Private Sub ReleaseConnection(ByVal oSocket As cWinsock)
    Dim i As Long
    For i = 0 To m_lPoolSize - 1
        If m_oConnections(i).Socket Is oSocket Then
            m_oConnections(i).InUse = False
            m_oConnections(i).LastUsed = Now
            Exit For
        End If
    Next
End Sub
```

---

### 4️⃣ 批量发送优化

```vb
' 批量发送
Private Sub SendBatch(ByVal vData() As Variant)
    Const BATCH_SIZE As Long = 100
    
    Dim lStart As Long
    lStart = LBound(vData)
    
    Do While lStart <= UBound(vData)
        Dim lEnd As Long
        lEnd = Min(lStart + BATCH_SIZE - 1, UBound(vData))
        
        Dim lBatchCount As Long
        lBatchCount = lEnd - lStart + 1
        
        ' 一次性发送多个数据包（使用分隔符）
        Dim i As Long
        Dim sBatch As String
        
        For i = lStart To lEnd
            sBatch = sBatch & vData(i) & vbCrLf
        Next
        
        m_oClient.SendData sBatch
        
        ' 等待发送完成
        Do While m_bSending
            DoEvents
        Loop
        
        lStart = lEnd + 1
    Loop
End Sub
```

---

## 错误处理

### 1️⃣ 统一错误处理

```vb
' 错误处理模块
Public Enum ErrorLevel
    elInfo = 0
    elWarning = 1
    elError = 2
    elCritical = 3
End Enum

' 统一错误日志
Public Sub LogError(ByVal eLevel As ErrorLevel, ByVal sSource As String, ByVal sMessage As String, ByVal lErrNum As Long)
    Dim sPrefix As String
    
    Select Case eLevel
        Case elInfo:      sPrefix = "[INFO]"
        Case elWarning:   sPrefix = "[WARN]"
        Case elError:     sPrefix = "[ERROR]"
        Case elCritical:  sPrefix = "[CRIT]"
    End Select
    
    Dim sLog As String
    sLog = Format$(Now, "yyyy-mm-dd hh:mm:ss") & " " & sPrefix & " [" & sSource & "] " & sMessage & " (Error " & lErrNum & ")"
    
    Debug.Print sLog
    
    ' 写入文件
    WriteToLogFile sLog
End Sub

' 使用示例
Private Sub m_oClient_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    LogError elError, "Client", Description, Number
    
    Select Case Number
        Case 10053, 10054
            ' 连接关闭，正常
            LogError elInfo, "Client", "连接被远程关闭", Number
            
        Case 10060
            ' 连接超时
            LogError elWarning, "Client", "连接超时", Number
            
        Case Else
            ' 其他错误
            LogError elError, "Client", Description, Number
    End Select
End Sub
```

---

### 2️⃣ 重试机制

```vb
' 带重试的操作
Private Function DoWithRetry(ByVal sFuncName As String, ByVal lMaxRetries As Long, ByVal vFunc As Variant) As Boolean
    Dim lRetry As Long
    Dim bSuccess As Boolean
    
    For lRetry = 1 To lMaxRetries
        On Error Resume Next
        bSuccess = CallByName(vFunc, sFuncName, VbMethod)
        
        If bSuccess And Err.Number = 0 Then
            LogError elInfo, "Retry", sFuncName & " 成功 (尝试 " & lRetry & "/" & lMaxRetries & ")", 0
            DoWithRetry = True
            Exit Function
        End If
        
        LogError elWarning, "Retry", sFuncName & " 失败 (尝试 " & lRetry & "/" & lMaxRetries & ")", Err.Number
        
        ' 等待后重试
        Sleep 1000 * lRetry
    Next
    
    LogError elError, "Retry", sFuncName & " 失败，超过最大重试次数", 0
    DoWithRetry = False
End Function

' 使用示例
Private Function SendDataWithRetry(ByVal sData As String) As Boolean
    On Error Resume Next
    m_oClient.SendData sData
    SendDataWithRetry = (Err.Number = 0)
End Function

Private Sub SendImportantData(ByVal sData As String)
    If Not DoWithRetry("SendDataWithRetry", 3, Me) Then
        LogError elCritical, "Send", "无法发送重要数据", 0
    End If
End Sub
```

---

### 3️⃣ 资源清理

```vb
' 确保资源清理
Private Sub SafeCloseSocket(ByRef oSocket As cWinsock)
    On Error Resume Next
    
    If Not oSocket Is Nothing Then
        If oSocket.State <> sckClosed Then
            oSocket.Close_
            Debug.Print "Socket 已关闭"
        End If
        Set oSocket = Nothing
    End If
End Sub

' 窗体卸载时清理所有资源
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' 关闭所有 socket
    SafeCloseSocket m_oClient
    SafeCloseSocket m_oServer
    SafeCloseSocket m_oUdp
    
    ' 停止所有定时器
    tmrHeartbeat.Enabled = False
    tmrCleanup.Enabled = False
    
    ' 清理集合
    Set m_oWorkQueue = Nothing
    Set m_oClients = Nothing
    
    Debug.Print "所有资源已清理"
End Sub
```

---

## 安全建议

### 1️⃣ 连接验证

```vb
' 连接前验证
Private Function ValidateConnection(ByVal sHost As String, ByVal lPort As Long) As Boolean
    ' 检查白名单
    If Not IsWhitelisted(sHost) Then
        LogError elWarning, "Security", sHost & " 不在白名单中", 0
        ValidateConnection = False
        Exit Function
    End If
    
    ' 检查端口范围
    If lPort < 1024 Or lPort > 65535 Then
        LogError elWarning, "Security", "端口 " & lPort & " 超出允许范围", 0
        ValidateConnection = False
        Exit Function
    End If
    
    ' 检查连接数限制
    If m_oServer.ClientCount >= MAX_CONNECTIONS Then
        LogError elWarning, "Security", "达到最大连接数", 0
        ValidateConnection = False
        Exit Function
    End If
    
    ValidateConnection = True
End Function

' 使用
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    If Not ValidateConnection(Client.RemoteHostIP, Client.RemotePort) Then
        DisConnect = True
    End If
End Sub
```

---

### 2️⃣ 数据验证

```vb
' 验证接收的数据
Private Function ValidateData(ByVal sData As String) As Boolean
    ' 检查长度
    If Len(sData) > MAX_DATA_SIZE Then
        LogError elWarning, "Security", "数据大小超过限制", 0
        ValidateData = False
        Exit Function
    End If
    
    ' 检查危险字符
    If InStr(sData, "<script") > 0 Or InStr(sData, "javascript:") > 0 Then
        LogError elWarning, "Security", "检测到危险内容", 0
        ValidateData = False
        Exit Function
    End If
    
    ' 自定义验证
    If Not CustomValidation(sData) Then
        ValidateData = False
        Exit Function
    End If
    
    ValidateData = True
End Function

' 使用
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    If Not ValidateData(sData) Then
        LogError elError, "Security", "拒绝无效数据", 0
        Client.Close_
        Exit Sub
    End If
    
    ' 处理数据
    ProcessData sData
End Sub
```

---

### 3️⃣ 防止缓冲区溢出

```vb
' 限制缓冲区大小
Private Const MAX_BUFFER_SIZE As Long = 1048576 ' 1MB

Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    ' 检查缓冲区大小
    If Client.BytesReceived > MAX_BUFFER_SIZE Then
        LogError elCritical, "Security", "缓冲区溢出，关闭连接", 0
        Client.Close_
        Exit Sub
    End If
    
    ' 读取数据
    Dim sData As String
    Client.GetData sData
End Sub
```

---

### 4️⃣ 速率限制

```vb
' 速率限制
Private Type tRateLimit
    Window As Date
    RequestCount As Long
End Type

Private m_oRateLimits As Collection
Private Const MAX_REQUESTS_PER_MINUTE As Long = 60

Private Function CheckRateLimit(ByVal sIP As String) As Boolean
    Dim tLimit As tRateLimit
    On Error Resume Next
    tLimit = m_oRateLimits(sIP)
    
    ' 如果是新 IP，创建记录
    If Err.Number <> 0 Then
        tLimit.Window = Now
        tLimit.RequestCount = 0
        m_oRateLimits.Add tLimit, sIP
    End If
    
    ' 检查时间窗口
    If DateDiff("s", tLimit.Window, Now) > 60 Then
        ' 超过 1 分钟，重置
        tLimit.Window = Now
        tLimit.RequestCount = 0
    End If
    
    ' 检查请求数
    If tLimit.RequestCount >= MAX_REQUESTS_PER_MINUTE Then
        LogError elWarning, "Security", sIP & " 超过速率限制", 0
        CheckRateLimit = False
    Else
        tLimit.RequestCount = tLimit.RequestCount + 1
        m_oRateLimits.Remove sIP
        m_oRateLimits.Add tLimit, sIP
        CheckRateLimit = True
    End If
End Function

' 使用
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    If Not CheckRateLimit(Client.RemoteHostIP) Then
        DisConnect = True
    End If
End Sub
```

---

## 调试技巧

### 1️⃣ 详细的日志记录

```vb
' 日志级别
Public Enum LogLevel
    llDebug = 0
    llInfo = 1
    llWarning = 2
    llError = 3
End Enum

Public m_eLogLevel As LogLevel

' 带级别的日志
Public Sub Log(ByVal eLevel As LogLevel, ByVal sSource As String, ByVal sMessage As String)
    If eLevel < m_eLogLevel Then Exit Sub
    
    Dim sPrefix As String
    Select Case eLevel
        Case llDebug:   sPrefix = "[DEBUG]"
        Case llInfo:    sPrefix = "[INFO]"
        Case llWarning:  sPrefix = "[WARN]"
        Case llError:    sPrefix = "[ERROR]"
    End Select
    
    Dim sLog As String
    sLog = Format$(Now, "hh:mm:ss") & " " & sPrefix & " [" & sSource & "] " & sMessage
    
    Debug.Print sLog
    
    ' 写入日志文件
    WriteLogToFile sLog
End Sub

' 使用
Private Sub m_oClient_Connect(Client As cWinsock)
    Log llInfo, "Client", "已连接到 " & Client.RemoteHostIP & ":" & Client.RemotePort
End Sub

Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Log llDebug, "Client", "收到 " & bytesTotal & " 字节"
    
    Dim sData As String
    Client.GetData sData
    
    Log llDebug, "Client", "数据内容: " & Left$(sData, 100) ' 只记录前 100 字符
End Sub
```

---

### 2️⃣ 数据包捕获

```vb
' 数据包捕获
Private Type tPacketCapture
    Timestamp As Date
    Direction As String ' "IN" or "OUT"
    Data As String
    Size As Long
End Type

Private m_oPackets As Collection

Private Sub CapturePacket(ByVal sDir As String, ByVal sData As String)
    Dim tPacket As tPacketCapture
    
    tPacket.Timestamp = Now
    tPacket.Direction = sDir
    tPacket.Data = Left$(sData, 200) ' 限制长度
    tPacket.Size = Len(sData)
    
    m_oPackets.Add tPacket
    
    Debug.Print "[" & sDir & "] " & Format$(tPacket.Timestamp, "hh:mm:ss") & " " & Len(sData) & " bytes"
End Sub

' 使用
Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    CapturePacket "IN", sData
End Sub

Private Sub cmdSend_Click()
    Dim sData As String
    sData = txtSend.Text
    
    m_oClient.SendData sData
    CapturePacket "OUT", sData
End Sub
```

---

### 3️⃣ 性能监控

```vb
' 性能统计
Private Type tPerformanceStats
    TotalPackets As Long
    TotalBytes As Long
    StartTime As Date
    PacketsPerSecond As Double
    BytesPerSecond As Double
End Type

Private m_oStats As tPerformanceStats

' 初始化统计
Private Sub InitStats()
    m_oStats.TotalPackets = 0
    m_oStats.TotalBytes = 0
    m_oStats.StartTime = Now
End Sub

' 更新统计
Private Sub UpdateStats(ByVal lBytes As Long)
    m_oStats.TotalPackets = m_oStats.TotalPackets + 1
    m_oStats.TotalBytes = m_oStats.TotalBytes + lBytes
    
    Dim lElapsed As Double
    lElapsed = DateDiff("s", m_oStats.StartTime, Now)
    
    If lElapsed > 0 Then
        m_oStats.PacketsPerSecond = m_oStats.TotalPackets / lElapsed
        m_oStats.BytesPerSecond = m_oStats.TotalBytes / lElapsed
    End If
End Sub

' 显示统计
Private Sub ShowStats()
    Debug.Print "===== 性能统计 ====="
    Debug.Print "运行时间: " & DateDiff("s", m_oStats.StartTime, Now) & " 秒"
    Debug.Print "总包数: " & m_oStats.TotalPackets
    Debug.Print "总字节数: " & m_oStats.TotalBytes
    Debug.Print "包/秒: " & Format$(m_oStats.PacketsPerSecond, "0.00")
    Debug.Print "字节/秒: " & Format$(m_oStats.BytesPerSecond, "0.00")
    Debug.Print "=================="
End Sub
```

---

## 常见陷阱

### 1️⃣ 忘记 `DoEvents`

```vb
' ❌ 错误：长时间处理会阻塞 UI
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    ' 长时间操作，UI 会冻结
    ProcessLargeData sData
End Sub

' ✅ 正确：定期释放控制权
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    Dim i As Long
    For i = 1 To 1000
        ProcessDataChunk sData, i
        
        ' 定期释放控制权
        If i Mod 10 = 0 Then
            DoEvents
        End If
    Next
End Sub
```

---

### 2️⃣ 内存泄漏

```vb
' ❌ 错误：不及时释放对象
Private Sub ProcessClients()
    Dim oClient As cWinsock
    For Each oClient In m_oServer.Clients
        Dim oData As New CDataObject
        oData.Data = "xxx"
        
        ' oData 没有释放
    Next
End Sub

' ✅ 正确：及时释放
Private Sub ProcessClients()
    Dim oClient As cWinsock
    For Each oClient In m_oServer.Clients
        Dim oData As New CDataObject
        oData.Data = "xxx"
        
        ' 使用后立即释放
        Set oData = Nothing
    Next
End Sub
```

---

### 3️⃣ 忽略状态检查

```vb
' ❌ 错误：不检查状态就操作
Private Sub cmdSend_Click()
    m_oClient.SendData "Hello" ' 可能失败
End Sub

' ✅ 正确：先检查状态
Private Sub cmdSend_Click()
    If m_oClient.State = sckConnected Then
        m_oClient.SendData "Hello"
    Else
        MsgBox "未连接", vbExclamation
    End If
End Sub
```

---

### 4️⃣ 错误的编码使用

```vb
' ❌ 错误：编码不一致
m_oClient.SendData "中文", ScpUtf8  ' UTF-8
' 接收时
Client.GetData sData  ' 默认 ACP → 乱码

' ✅ 正确：保持一致
m_oClient.SendData "中文", ScpUtf8  ' UTF-8
' 接收时
Client.GetData sData, , , ScpUtf8  ' UTF-8
```

---

**最后更新**: 2026-01-09
