# Modbus 进阶应用

本文档介绍 Modbus 类库的高级功能、最佳实践和常见应用场景。

---

## ? 目录

- [高级主题](#高级主题)
- [性能优化](#性能优化)
- [错误处理](#错误处理)
- [多从站管理](#多从站管理)
- [数据缓存策略](#数据缓存策略)
- [日志与调试](#日志与调试)
- [安全考虑](#安全考虑)
- [实际应用场景](#实际应用场景)
- [常见问题](#常见问题)

---

## 高级主题

### 1. 主从站合一模式

某些应用场景需要同一设备既作为主站又作为从站:

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents mbSlave As cModbusSlave

Private Sub Form_Load()
    Set mbMaster = New cModbusMaster
    Set mbSlave = New cModbusSlave
    
    ' 初始化主站 - 向上位机请求数据
    mbMaster.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mbMaster.TCPHost = "192.168.1.100"
    mbMaster.TCPPort = 502
    mbMaster.SlaveID = 1
    
    ' 初始化从站 - 向下位机提供数据
    mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_TCP
    mbSlave.SlaveID = 2
    mbSlave.BindAddress = "0.0.0.0"  ' 监听所有接口
    mbSlave.Start 1502
End Sub

' 从主站接收数据后,更新从站提供给下位机
Private Sub tmrSync_Timer()
    On Error Resume Next
    
    ' 从上位机读取
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(0, 10)
    
    ' 写入本地从站,供下位机读取
    If UBound(iRegs) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(iRegs)
            mbSlave.SetHoldingRegister i, iRegs(i)
        Next i
    End If
End Sub
```

---

### 2. 事务管理

实现带重试的事务机制:

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private Const MAX_RETRIES As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

Private Function ReadWithRetry(StartAddr As Long, Quantity As Long) As Integer()
    Dim iRetry As Long
    Dim bSuccess As Boolean
    Dim iRegs() As Integer
    
    For iRetry = 1 To MAX_RETRIES
        On Error Resume Next
        
        If mbMaster.State = MB_MASTER_STATE_DISCONNECTED Then
            mbMaster.Connect
        End If
        
        iRegs = mbMaster.ReadHoldingRegisters(StartAddr, Quantity)
        
        If UBound(iRegs) >= 0 Then
            bSuccess = True
            Exit For
        End If
        
        ' 失败后等待
        If iRetry < MAX_RETRIES Then
            Sleep RETRY_DELAY_MS
        End If
    Next iRetry
    
    If bSuccess Then
        ReadWithRetry = iRegs
    Else
        ' 所有重试都失败
        RaiseError "ReadWithRetry failed after " & MAX_RETRIES & " attempts"
    End If
End Function
```

---

### 3. 异步操作模式

使用 Timer 实现非阻塞的异步操作:

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents tmrAsync As Timer
Private m_lPendingAddr As Long
Private m_lPendingCount As Long
Private m_bAsyncBusy As Boolean

Private Sub Form_Load()
    Set mbMaster = New cModbusMaster
    Set tmrAsync = New Timer
    
    tmrAsync.Interval = 100  ' 100ms 检查间隔
End Sub

' 发起异步读取请求
Public Sub AsyncReadHoldingRegisters(StartAddr As Long, Quantity As Long)
    If m_bAsyncBusy Then
        Debug.Print "Async operation in progress"
        Exit Sub
    End If
    
    m_bAsyncBusy = True
    m_lPendingAddr = StartAddr
    m_lPendingCount = Quantity
    
    tmrAsync.Enabled = True
End Sub

' Timer 回调执行读取
Private Sub tmrAsync_Timer()
    On Error Resume Next
    
    tmrAsync.Enabled = False
    
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(m_lPendingAddr, m_lPendingCount)
    
    If UBound(iRegs) >= 0 Then
        ' 读取成功,触发回调
        OnAsyncReadComplete m_lPendingAddr, m_lPendingCount, iRegs
    Else
        ' 读取失败
        OnAsyncError "Read failed"
    End If
    
    m_bAsyncBusy = False
End Sub

' 读取完成回调
Private Sub OnAsyncReadComplete(StartAddr As Long, Quantity As Long, ByRef Values() As Integer)
    ' 处理数据
    ProcessData Values
    
    RaiseEvent AsyncReadComplete(StartAddr, Quantity, Values)
End Sub

' 错误回调
Private Sub OnAsyncError(Description As String)
    Debug.Print "Async error: " & Description
    RaiseEvent AsyncError(Description)
End Sub
```

---

## 性能优化

### 1. 批量读取优化

尽量使用单次批量读取,减少通信次数:

```vb
' 差的做法 - 多次读取
Private Sub BadReadApproach()
    Dim i As Long
    Dim iValue As Integer
    
    For i = 0 To 99
        iValue = mbMaster.ReadHoldingRegisters(i, 1)(0)
        ' 处理数据
    Next i
End Sub

' 好的做法 - 批量读取
Private Sub GoodReadApproach()
    Dim iRegs() As Integer
    Dim i As Long
    
    ' 一次读取100个寄存器
    iRegs = mbMaster.ReadHoldingRegisters(0, 100)
    
    For i = 0 To UBound(iRegs)
        ' 处理数据
    Next i
End Sub
```

---

### 2. 读取频率控制

避免过于频繁的读取请求:

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents tmrPoll As Timer
Private m_dLastReadTime As Double
Private Const MIN_READ_INTERVAL_SEC As Double = 0.5  ' 最小间隔0.5秒

Private Sub tmrPoll_Timer()
    Dim dNow As Double
    dNow = Timer
    
    ' 检查是否到达最小间隔
    If dNow - m_dLastReadTime < MIN_READ_INTERVAL_SEC Then
        Exit Sub
    End If
    
    ' 执行读取
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(0, 10)
    
    m_dLastReadTime = dNow
End Sub
```

---

### 3. 数据预缓存

启动时预加载常用数据:

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private m_iRegCache() As Integer
Private Const CACHE_SIZE As Long = 100

Private Sub mbMaster_OnConnect()
    ' 连接成功后立即加载缓存
    LoadCache
End Sub

Private Sub LoadCache()
    On Error Resume Next
    
    ReDim m_iRegCache(CACHE_SIZE - 1) As Integer
    
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(0, CACHE_SIZE)
    
    If UBound(iRegs) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(iRegs)
            If i < CACHE_SIZE Then
                m_iRegCache(i) = iRegs(i)
            End If
        Next i
        
        Debug.Print "Cache loaded: " & (UBound(iRegs) + 1) & " registers"
    End If
End Sub

' 从缓存快速读取
Public Function GetCachedRegister(Addr As Long) As Integer
    If Addr >= 0 And Addr < CACHE_SIZE Then
        GetCachedRegister = m_iRegCache(Addr)
    Else
        GetCachedRegister = -1  ' 缓存未命中
    End If
End Function

' 定期更新缓存
Private Sub tmrCacheUpdate_Timer()
    On Error Resume Next
    
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(0, CACHE_SIZE)
    
    If UBound(iRegs) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(iRegs)
            If i < CACHE_SIZE Then
                m_iRegCache(i) = iRegs(i)
            End If
        Next i
    End If
End Sub
```

---

### 4. 连接池管理

对于多从站场景,实现连接池:

```vb
Option Explicit

Private Type ModbusDevice
    ID As Long
    Host As String
    Port As Long
    SlaveID As Byte
    LastUsed As Double
    Connected As Boolean
End Type

Private m_Devices() As ModbusDevice
Private m_MaxDevices As Long
Private m_ConnectionTimeout As Double  ' 秒

Private Sub InitializeConnectionPool(MaxDevices As Long, TimeoutSeconds As Double)
    m_MaxDevices = MaxDevices
    m_ConnectionTimeout = TimeoutSeconds
    
    ReDim m_Devices(MaxDevices - 1) As ModbusDevice
    
    Dim i As Long
    For i = 0 To MaxDevices - 1
        m_Devices(i).Connected = False
        m_Devices(i).LastUsed = Timer - m_ConnectionTimeout - 1
    Next i
End Sub

' 获取设备连接
Public Function GetDevice(DeviceID As Long) As cModbusMaster
    Dim mbDev As cModbusMaster
    Dim i As Long
    
    ' 查找设备
    For i = 0 To m_MaxDevices - 1
        If m_Devices(i).ID = DeviceID Then
            ' 更新使用时间
            m_Devices(i).LastUsed = Timer
            
            Set mbDev = New cModbusMaster
            mbDev.ProtocolType = MB_MASTER_PROTOCOL_TCP
            mbDev.TCPHost = m_Devices(i).Host
            mbDev.TCPPort = m_Devices(i).Port
            mbDev.SlaveID = m_Devices(i).SlaveID
            
            Set GetDevice = mbDev
            Exit Function
        End If
    Next i
    
    Set GetDevice = Nothing
End Function

' 清理过期连接
Public Sub CleanupConnections()
    Dim i As Long
    Dim dNow As Double
    dNow = Timer
    
    For i = 0 To m_MaxDevices - 1
        If m_Devices(i).Connected Then
            ' 超过超时时间
            If dNow - m_Devices(i).LastUsed > m_ConnectionTimeout Then
                ' 标记为断开
                m_Devices(i).Connected = False
                Debug.Print "Device " & m_Devices(i).ID & " connection timed out"
            End If
        End If
    Next i
End Sub
```

---

## 错误处理

### 1. 综合错误处理

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster

' 读取并处理错误
Public Function SafeReadRegisters(StartAddr As Long, Quantity As Long) As Integer()
    On Error GoTo ErrorHandler
    
    ' 检查连接状态
    If mbMaster.State <> MB_MASTER_STATE_CONNECTED Then
        Err.Raise vbObjectError + 1, "SafeReadRegisters", "Not connected"
    End If
    
    ' 检查参数
    If StartAddr < 0 Then
        Err.Raise vbObjectError + 2, "SafeReadRegisters", "Invalid start address"
    End If
    
    If Quantity < 1 Or Quantity > mbMaster.Defaults.MAX_REGISTERS Then
        Err.Raise vbObjectError + 3, "SafeReadRegisters", "Invalid quantity"
    End If
    
    ' 执行读取
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(StartAddr, Quantity)
    
    SafeReadRegisters = iRegs
    Exit Function
    
ErrorHandler:
    Dim sError As String
    sError = "Error " & Err.Number & ": " & Err.Description
    
    Debug.Print sError
    LogError sError
    
    ' 返回空数组
    SafeReadRegisters = Array()
End Function

' 错误日志
Private Sub LogError(sMessage As String)
    Dim iFile As Integer
    iFile = FreeFile
    
    Open App.Path & "\modbus_error.log" For Append As #iFile
    Print #iFile, Format$(Now, "yyyy-mm-dd hh:mm:ss") & " - " & sMessage
    Close #iFile
End Sub
```

---

### 2. 异常码处理

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster

Private Sub mbMaster_OnError(ByVal Description As String)
    Select Case True
        Case InStr(Description, "Illegal Function") > 0
            HandleIllegalFunction
            
        Case InStr(Description, "Illegal Data Address") > 0
            HandleIllegalDataAddress
            
        Case InStr(Description, "Illegal Data Value") > 0
            HandleIllegalDataValue
            
        Case InStr(Description, "Slave Device Failure") > 0
            HandleSlaveDeviceFailure
            
        Case InStr(Description, "Slave Device Busy") > 0
            HandleSlaveDeviceBusy
            
        Case InStr(Description, "Response timeout") > 0
            HandleTimeout
            
        Case Else
            HandleGenericError Description
    End Select
End Sub

Private Sub HandleIllegalFunction()
    Debug.Print "功能码不支持"
    ' 可能需要使用不同的功能码
End Sub

Private Sub HandleIllegalDataAddress()
    Debug.Print "地址超出范围"
    ' 检查地址是否在有效范围内
End Sub

Private Sub HandleIllegalDataValue()
    Debug.Print "数据值非法"
    ' 检查写入的数据值
End Sub

Private Sub HandleSlaveDeviceFailure()
    Debug.Print "从站设备故障"
    ' 可能需要报警或切换到备用设备
End Sub

Private Sub HandleSlaveDeviceBusy()
    Debug.Print "从站忙"
    ' 等待后重试
    Sleep 1000
End Sub

Private Sub HandleTimeout()
    Debug.Print "响应超时"
    ' 检查网络或串口连接
    ' 可能需要重连
End Sub

Private Sub HandleGenericError(sDescription As String)
    Debug.Print "通用错误: " & sDescription
End Sub
```

---

## 多从站管理

### 1. 设备配置管理

```vb
Option Explicit

Private Type DeviceConfig
    DeviceID As Long
    Name As String
    ProtocolType As ModbusMasterProtocolType
    Host As String
    Port As Long
    SerialPort As String
    BaudRate As Long
    SlaveID As Byte
    PollInterval As Long  ' 轮询间隔(毫秒)
    Enabled As Boolean
End Type

Private m_Devices() As DeviceConfig
Private m_MasterConnections() As cModbusMaster

' 添加设备配置
Public Sub AddDevice(DevID As Long, sName As String, Protocol As ModbusProtocolType, _
                      sHost As String, lPort As Long, sSerialPort As String, _
                      lBaudRate As Long, bSlaveID As Byte, lPollInterval As Long)
    
    Dim iCount As Long
    iCount = UBound(m_Devices) + 1
    
    ReDim Preserve m_Devices(iCount) As DeviceConfig
    ReDim Preserve m_MasterConnections(iCount) As cModbusMaster
    
    With m_Devices(iCount)
        .DeviceID = DevID
        .Name = sName
        .ProtocolType = Protocol
        .Host = sHost
        .Port = lPort
        .SerialPort = sSerialPort
        .BaudRate = lBaudRate
        .SlaveID = bSlaveID
        .PollInterval = lPollInterval
        .Enabled = True
    End With
    
    Set m_MasterConnections(iCount) = New cModbusMaster
    
    Debug.Print "Device added: " & sName & " (ID: " & DevID & ")"
End Sub

' 连接所有设备
Public Sub ConnectAllDevices()
    Dim i As Long
    
    For i = 0 To UBound(m_Devices)
        If m_Devices(i).Enabled Then
            ConnectDevice i
        End If
    Next i
End Sub

' 连接单个设备
Private Sub ConnectDevice(Index As Long)
    On Error Resume Next
    
    Dim mbDev As cModbusMaster
    Set mbDev = m_MasterConnections(Index)
    
    With m_Devices(Index)
        mbDev.ProtocolType = .ProtocolType
        mbDev.SlaveID = .SlaveID
        
        If .ProtocolType = MB_MASTER_PROTOCOL_TCP Then
            mbDev.TCPHost = .Host
            mbDev.TCPPort = .Port
            mbDev.Connect
        Else
            mbDev.SerialPort = .SerialPort
            mbDev.BaudRate = .BaudRate
            mbDev.DataBits = 8
            mbDev.Parity = "N"
            mbDev.StopBits = 1
            mbDev.Connect .SerialPort
        End If
    End With
    
    If Err.Number = 0 Then
        Debug.Print "Device " & m_Devices(Index).Name & " connected"
    Else
        Debug.Print "Device " & m_Devices(Index).Name & " connect failed: " & Err.Description
    End If
End Sub
```

---

### 2. 统一轮询管理

```vb
Option Explicit

Private WithEvents tmrPoll As Timer
Private m_CurrentDeviceIndex As Long

Private Sub StartPolling()
    Set tmrPoll = New Timer
    tmrPoll.Interval = 100  ' 100ms 轮询间隔
    m_CurrentDeviceIndex = 0
    tmrPoll.Enabled = True
End Sub

Private Sub tmrPoll_Timer()
    tmrPoll.Enabled = False
    
    ' 轮询当前设备
    PollDevice m_CurrentDeviceIndex
    
    ' 移动到下一个设备
    m_CurrentDeviceIndex = m_CurrentDeviceIndex + 1
    If m_CurrentDeviceIndex > UBound(m_Devices) Then
        m_CurrentDeviceIndex = 0
    End If
    
    tmrPoll.Enabled = True
End Sub

Private Sub PollDevice(Index As Long)
    On Error Resume Next
    
    Dim mbDev As cModbusMaster
    Set mbDev = m_MasterConnections(Index)
    
    ' 检查设备是否启用
    If Not m_Devices(Index).Enabled Then
        Exit Sub
    End If
    
    ' 检查连接状态
    If mbDev.State <> MB_MASTER_STATE_CONNECTED Then
        ConnectDevice Index
        Exit Sub
    End If
    
    ' 读取设备数据
    Dim iRegs() As Integer
    iRegs = mbDev.ReadHoldingRegisters(0, 10)
    
    If UBound(iRegs) >= 0 Then
        ' 处理数据
        ProcessDeviceData m_Devices(Index).DeviceID, iRegs
        
        ' 更新最后通信时间
        m_Devices(Index).LastUsed = Timer
    End If
End Sub

Private Sub ProcessDeviceData(DeviceID As Long, ByRef Values() As Integer)
    ' 根据设备ID处理数据
    Select Case DeviceID
        Case 1
            ProcessDevice1Data Values
        Case 2
            ProcessDevice2Data Values
        ' ...
    End Select
End Sub
```

---

## 数据缓存策略

### 1. 多级缓存

```vb
Option Explicit

Private Enum CacheLevel
    CACHE_LEVEL_NONE = 0
    CACHE_LEVEL_L1 = 1      ' 内存缓存
    CACHE_LEVEL_L2 = 2      ' 文件缓存
    CACHE_LEVEL_L3 = 3      ' 数据库缓存
End Enum

Private m_L1Cache() As Integer      ' 内存缓存
Private m_L1CacheValid() As Boolean ' 缓存有效性标记
Private m_CacheLevel As CacheLevel

' 初始化缓存
Public Sub InitializeCache(Size As Long, Level As CacheLevel)
    ReDim m_L1Cache(Size - 1) As Integer
    ReDim m_L1CacheValid(Size - 1) As Boolean
    m_CacheLevel = Level
    
    Dim i As Long
    For i = 0 To Size - 1
        m_L1CacheValid(i) = False
    Next i
End Sub

' 读取寄存器(带缓存)
Public Function ReadWithCache(Addr As Long) As Integer
    ' 先检查L1缓存
    If Addr >= 0 And Addr <= UBound(m_L1Cache) Then
        If m_L1CacheValid(Addr) And m_CacheLevel >= CACHE_LEVEL_L1 Then
            ReadWithCache = m_L1Cache(Addr)
            Debug.Print "Cache L1 hit: Reg[" & Addr & "] = " & m_L1Cache(Addr)
            Exit Function
        End If
    End If
    
    ' L1缓存未命中,从设备读取
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(Addr, 1)
    
    If UBound(iRegs) >= 0 Then
        ' 更新L1缓存
        If Addr >= 0 And Addr <= UBound(m_L1Cache) Then
            m_L1Cache(Addr) = iRegs(0)
            m_L1CacheValid(Addr) = True
        End If
        
        ReadWithCache = iRegs(0)
        Debug.Print "Cache miss: Reg[" & Addr & "] = " & iRegs(0)
    End If
End Function

' 无效化缓存
Public Sub InvalidateCache(Addr As Long)
    On Error Resume Next
    
    If Addr >= 0 And Addr <= UBound(m_L1Cache) Then
        m_L1CacheValid(Addr) = False
    End If
End Sub

' 无效化所有缓存
Public Sub InvalidateAllCache()
    Dim i As Long
    For i = 0 To UBound(m_L1Cache)
        m_L1CacheValid(i) = False
    Next i
    
    Debug.Print "All cache invalidated"
End Sub
```

---

### 2. 写入同步策略

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents mbSlave As cModbusSlave

' 主站读取后同步到从站
Private Sub SyncMasterToSlave()
    On Error Resume Next
    
    ' 从上位机读取
    Dim iMasterRegs() As Integer
    iMasterRegs = mbMaster.ReadHoldingRegisters(0, 10)
    
    If UBound(iMasterRegs) >= 0 Then
        ' 比较数据是否变化
        Dim bChanged As Boolean
        bChanged = CheckDataChanged(iMasterRegs)
        
        ' 如果有变化,同步到从站
        If bChanged Then
            Dim i As Long
            For i = 0 To UBound(iMasterRegs)
                mbSlave.SetHoldingRegister i, iMasterRegs(i)
            Next i
            
            Debug.Print "Synced " & (UBound(iMasterRegs) + 1) & " registers"
        End If
    End If
End Sub

' 检查数据是否变化
Private Function CheckDataChanged(ByRef NewData() As Integer) As Boolean
    Static iLastData() As Integer
    Static bInitialized As Boolean
    
    If Not bInitialized Then
        ' 首次初始化
        iLastData = NewData
        bInitialized = True
        CheckDataChanged = True
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To UBound(NewData)
        If iLastData(i) <> NewData(i) Then
            iLastData(i) = NewData(i)
            CheckDataChanged = True
            Exit Function
        End If
    Next i
    
    CheckDataChanged = False
End Function
```

---

## 日志与调试

### 1. 详细日志记录

```vb
Option Explicit

Private Enum LogLevel
    LOG_LEVEL_DEBUG = 0
    LOG_LEVEL_INFO = 1
    LOG_LEVEL_WARNING = 2
    LOG_LEVEL_ERROR = 3
    LOG_LEVEL_CRITICAL = 4
End Enum

Private m_LogLevel As LogLevel
Private m_LogFile As String

' 初始化日志
Public Sub InitLog(sFilePath As String, Level As LogLevel)
    m_LogFile = sFilePath
    m_LogLevel = Level
    
    ' 创建日志目录
    Dim sDir As String
    sDir = Left$(sFilePath, InStrRev(sFilePath, "\") - 1)
    
    On Error Resume Next
    MkDir sDir
    On Error GoTo 0
End Sub

' 写入日志
Public Sub WriteLog(Level As LogLevel, sMessage As String)
    ' 只记录指定级别及以上的日志
    If Level < m_LogLevel Then Exit Sub
    
    Dim sLevel As String
    Select Case Level
        Case LOG_LEVEL_DEBUG:    sLevel = "DEBUG"
        Case LOG_LEVEL_INFO:     sLevel = "INFO "
        Case LOG_LEVEL_WARNING:  sLevel = "WARN "
        Case LOG_LEVEL_ERROR:    sLevel = "ERROR"
        Case LOG_LEVEL_CRITICAL: sLevel = "CRIT "
    End Select
    
    Dim sLogLine As String
    sLogLine = Format$(Now, "yyyy-mm-dd hh:mm:ss") & " [" & sLevel & "] " & sMessage
    
    Debug.Print sLogLine
    
    On Error Resume Next
    Dim iFile As Integer
    iFile = FreeFile
    Open m_LogFile For Append As #iFile
    Print #iFile, sLogLine
    Close #iFile
    On Error GoTo 0
End Sub

' 封装的日志函数
Public Sub LogDebug(sMessage As String)
    WriteLog LOG_LEVEL_DEBUG, sMessage
End Sub

Public Sub LogInfo(sMessage As String)
    WriteLog LOG_LEVEL_INFO, sMessage
End Sub

Public Sub LogWarning(sMessage As String)
    WriteLog LOG_LEVEL_WARNING, sMessage
End Sub

Public Sub LogError(sMessage As String)
    WriteLog LOG_LEVEL_ERROR, sMessage
End Sub

Public Sub LogCritical(sMessage As String)
    WriteLog LOG_LEVEL_CRITICAL, sMessage
End Sub
```

---

### 2. 数据包调试

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster

' 数据包调试开关
Private m_bDebugPackets As Boolean

Public Sub EnablePacketDebug(bEnable As Boolean)
    m_bDebugPackets = bEnable
End Sub

' 调试事件处理
Private Sub mbMaster_OnDataReceived(Data() As Byte)
    If Not m_bDebugPackets Then Exit Sub
    
    Dim sHex As String
    Dim i As Long
    
    For i = 0 To UBound(Data)
        sHex = sHex & Format$(Data(i), "00") & " "
    Next i
    
    Debug.Print "RX [" & Format$(Now, "hh:mm:ss") & "] " & sHex
    LogInfo "RX: " & sHex
End Sub

' 发送数据包调试
Private Sub DebugSendPacket(ByRef Data() As Byte)
    If Not m_bDebugPackets Then Exit Sub
    
    Dim sHex As String
    Dim i As Long
    
    For i = 0 To UBound(Data)
        sHex = sHex & Format$(Data(i), "00") & " "
    Next i
    
    Debug.Print "TX [" & Format$(Now, "hh:mm:ss") & "] " & sHex
    LogInfo "TX: " & sHex
End Sub

' 解析Modbus数据包
Public Sub ParseModbusPacket(ByRef Data() As Byte)
    If UBound(Data) < 2 Then Exit Sub
    
    Dim bFC As Byte
    bFC = Data(1)
    
    Debug.Print "=== Modbus Packet ==="
    Debug.Print "Slave ID: " & Data(0)
    Debug.Print "Function Code: 0x" & Hex$(bFC) & " (" & GetFunctionName(bFC) & ")"
    
    Select Case bFC
        Case &H1  ' Read Coils
            If UBound(Data) >= 3 Then
                Debug.Print "Byte Count: " & Data(2)
            End If
            
        Case &H3  ' Read Holding Registers
            If UBound(Data) >= 3 Then
                Dim iByteCount As Long
                iByteCount = Data(2)
                Debug.Print "Byte Count: " & iByteCount
                
                Dim i As Long
                For i = 0 To (iByteCount \ 2) - 1
                    Dim iReg As Integer
                    iReg = Data(3 + i * 2) * 256 + Data(4 + i * 2)
                    Debug.Print "  Reg[" & i & "] = " & iReg & " (0x" & Hex$(iReg) & ")"
                Next i
            End If
    End Select
    
    Debug.Print "====================="
End Sub

Private Function GetFunctionName(bFC As Byte) As String
    Select Case bFC
        Case &H1:  GetFunctionName = "Read Coils"
        Case &H2:  GetFunctionName = "Read Discrete Inputs"
        Case &H3:  GetFunctionName = "Read Holding Registers"
        Case &H4:  GetFunctionName = "Read Input Registers"
        Case &H5:  GetFunctionName = "Write Single Coil"
        Case &H6:  GetFunctionName = "Write Single Register"
        Case &HF:  GetFunctionName = "Write Multiple Coils"
        Case &H10: GetFunctionName = "Write Multiple Registers"
        Case Else: GetFunctionName = "Unknown"
    End Select
End Function
```

---

## 安全考虑

### 1. 连接认证

```vb
Option Explicit

Private Type AuthConfig
    Enabled As Boolean
    Username As String
    Password As String
    Token As String
    TokenExpiry As Date
End Type

Private m_Auth As AuthConfig

' 初始化认证
Public Sub SetAuthentication(bEnabled As Boolean, sUser As String, sPass As String)
    m_Auth.Enabled = bEnabled
    m_Auth.Username = sUser
    m_Auth.Password = sPass
    m_Auth.Token = ""
    m_Auth.TokenExpiry = #1/1/1900#
End Sub

' 检查认证
Public Function CheckAuthentication() As Boolean
    If Not m_Auth.Enabled Then
        CheckAuthentication = True  ' 未启用认证,直接通过
        Exit Function
    End If
    
    ' 检查Token是否过期
    If m_Auth.Token <> "" And Now < m_Auth.TokenExpiry Then
        CheckAuthentication = True
        Exit Function
    End If
    
    ' 执行认证
    CheckAuthentication = Authenticate()
End Function

' 执行认证
Private Function Authenticate() As Boolean
    ' 这里实现实际的认证逻辑
    ' 可以是通过Modbus发送认证请求
    ' 或者通过其他渠道认证
    
    ' 示例: 通过Modbus特定寄存器认证
    Dim bSuccess As Boolean
    
    On Error Resume Next
    
    ' 发送认证信息
    Dim iUserHash As Integer
    Dim iPassHash As Integer
    
    iUserHash = SimpleHash(m_Auth.Username)
    iPassHash = SimpleHash(m_Auth.Password)
    
    ' 写入认证寄存器
    mbMaster.WriteSingleRegister 998, iUserHash
    mbMaster.WriteSingleRegister 999, iPassHash
    
    ' 读取认证结果
    Dim iResult() As Integer
    iResult = mbMaster.ReadHoldingRegisters(1000, 1)
    
    If UBound(iResult) >= 0 Then
        If iResult(0) = 1 Then
            ' 认证成功
            m_Auth.Token = GenerateToken()
            m_Auth.TokenExpiry = DateAdd("h", 1, Now)  ' 1小时过期
            
            Authenticate = True
            Debug.Print "Authentication successful"
            Exit Function
        End If
    End If
    
    Authenticate = False
    Debug.Print "Authentication failed"
End Function

' 简单哈希(仅示例,实际应用应使用更安全的哈希)
Private Function SimpleHash(sInput As String) As Integer
    Dim i As Long
    Dim lHash As Long
    
    For i = 1 To Len(sInput)
        lHash = lHash + Asc(Mid$(sInput, i, 1)) * (i Mod 7 + 1)
    Next i
    
    SimpleHash = lHash And &HFFFF
End Function

' 生成Token(仅示例)
Private Function GenerateToken() As String
    GenerateToken = "TOKEN_" & Format$(Now, "yyyymmddhhmmss")
End Function
```

---

### 2. 数据加密(示例)

```vb
Option Explicit

Private m_bEncryptData As Boolean
Private m_EncryptionKey As String

' 启用数据加密
Public Sub EnableEncryption(bEnable As Boolean, sKey As String)
    m_bEncryptData = bEnable
    m_EncryptionKey = sKey
End Sub

' 加密数据(简单XOR加密,仅示例)
Public Function EncryptData(ByRef Data() As Byte) As Byte()
    If Not m_bEncryptData Then
        EncryptData = Data
        Exit Function
    End If
    
    Dim baEncrypted() As Byte
    ReDim baEncrypted(UBound(Data)) As Byte
    
    Dim i As Long
    Dim lKeyLen As Long
    lKeyLen = Len(m_EncryptionKey)
    
    For i = 0 To UBound(Data)
        Dim bKeyByte As Byte
        bKeyByte = Asc(Mid$(m_EncryptionKey, (i Mod lKeyLen) + 1, 1))
        baEncrypted(i) = Data(i) Xor bKeyByte
    Next i
    
    EncryptData = baEncrypted
End Function

' 解密数据
Public Function DecryptData(ByRef Data() As Byte) As Byte()
    ' XOR加密是对称的,解密和加密相同
    DecryptData = EncryptData(Data)
End Function
```

---

## 实际应用场景

### 1. 工业数据采集

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents tmrCollect As Timer

' 数据采集配置
Private Type DataPoint
    RegisterAddr As Long
    Name As String
    Unit As String
    MinValue As Single
    MaxValue As Single
    AlarmHigh As Single
    AlarmLow As Single
    LastValue As Single
    LastUpdate As Date
End Type

Private m_DataPoints() As DataPoint

' 初始化数据点
Public Sub InitializeDataPoints()
    ReDim m_DataPoints(9) As DataPoint
    
    ' 温度
    m_DataPoints(0).RegisterAddr = 0
    m_DataPoints(0).Name = "温度"
    m_DataPoints(0).Unit = "°C"
    m_DataPoints(0).MinValue = -50
    m_DataPoints(0).MaxValue = 150
    m_DataPoints(0).AlarmHigh = 120
    m_DataPoints(0).AlarmLow = -20
    
    ' 压力
    m_DataPoints(1).RegisterAddr = 1
    m_DataPoints(1).Name = "压力"
    m_DataPoints(1).Unit = "kPa"
    m_DataPoints(1).MinValue = 0
    m_DataPoints(1).MaxValue = 10000
    m_DataPoints(1).AlarmHigh = 8000
    m_DataPoints(1).AlarmLow = 0
    
    ' 流量
    m_DataPoints(2).RegisterAddr = 2
    m_DataPoints(2).Name = "流量"
    m_DataPoints(2).Unit = "L/min"
    m_DataPoints(2).MinValue = 0
    m_DataPoints(2).MaxValue = 1000
    m_DataPoints(2).AlarmHigh = 900
    m_DataPoints(2).AlarmLow = 0
    
    ' ... 更多数据点
End Sub

' 数据采集循环
Private Sub tmrCollect_Timer()
    On Error Resume Next
    
    Dim i As Long
    Dim iRegs() As Integer
    
    ' 批量读取所有寄存器
    iRegs = mbMaster.ReadHoldingRegisters(0, 10)
    
    If UBound(iRegs) >= 0 Then
        For i = 0 To UBound(m_DataPoints)
            If i <= UBound(iRegs) Then
                Dim fValue As Single
                
                ' 转换原始值
                fValue = ConvertRawValue(iRegs(i), m_DataPoints(i))
                
                ' 更新数据点
                m_DataPoints(i).LastValue = fValue
                m_DataPoints(i).LastUpdate = Now
                
                ' 检查报警
                CheckAlarm i, fValue
                
                ' 记录数据
                LogDataPoint i, fValue
                
                ' 更新UI显示
                UpdateDataPointUI i, fValue
            End If
        Next i
    End If
End Sub

' 转换原始值
Private Function ConvertRawValue(iRaw As Integer, pt As DataPoint) As Single
    ' 根据单位转换原始值
    Select Case pt.Unit
        Case "°C"
            ' 温度: 直接使用
            ConvertRawValue = CSng(iRaw)
            
        Case "kPa"
            ' 压力: 原始值 * 10
            ConvertRawValue = CSng(iRaw) * 10
            
        Case "L/min"
            ' 流量: 原始值 * 0.1
            ConvertRawValue = CSng(iRaw) * 0.1
            
        Case Else
            ConvertRawValue = CSng(iRaw)
    End Select
End Function

' 检查报警
Private Sub CheckAlarm(Index As Long, fValue As Single)
    With m_DataPoints(Index)
        If fValue >= .AlarmHigh Then
            TriggerAlarm .Name, "高报警", fValue, .AlarmHigh
        ElseIf fValue <= .AlarmLow Then
            TriggerAlarm .Name, "低报警", fValue, .AlarmLow
        End If
    End With
End Sub

' 触发报警
Private Sub TriggerAlarm(sPointName As String, sAlarmType As String, fValue As Single, fThreshold As Single)
    Dim sMessage As String
    sMessage = sPointName & " " & sAlarmType & ": " & Format$(fValue, "0.00") & _
              " (阈值: " & Format$(fThreshold, "0.00") & ")"
    
    Debug.Print "[ALARM] " & Format$(Now, "hh:mm:ss") & " - " & sMessage
    
    ' 发送邮件通知
    ' SendAlarmEmail sMessage
    
    ' 记录报警日志
    ' LogAlarm sMessage
    
    ' 更新报警UI
    ' UpdateAlarmUI sPointName, sAlarmType, sMessage
End Sub

' 记录数据点
Private Sub LogDataPoint(Index As Long, fValue As Single)
    With m_DataPoints(Index)
        ' 保存到数据库或文件
        Dim sLine As String
        sLine = Format$(Now, "yyyy-mm-dd hh:mm:ss") & "," & _
                 .Name & "," & _
                 Format$(fValue, "0.00") & "," & _
                 .Unit
        
        ' WriteToFile sLine
    End With
End Sub
```

---

### 2. 设备控制

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster

' 设备控制命令
Public Enum DeviceCommand
    CMD_START = 1
    CMD_STOP = 2
    CMD_RESET = 3
    CMD_EMERGENCY_STOP = 99
End Enum

' 控制线圈映射
Private Type CoilMapping
    Command As DeviceCommand
    Address As Long
    Value As Boolean
End Type

Private m_CoilMappings() As CoilMapping

' 初始化控制映射
Public Sub InitializeControlMappings()
    ReDim m_CoilMappings(3) As CoilMapping
    
    m_CoilMappings(0).Command = CMD_START
    m_CoilMappings(0).Address = 0
    m_CoilMappings(0).Value = True
    
    m_CoilMappings(1).Command = CMD_STOP
    m_CoilMappings(1).Address = 1
    m_CoilMappings(1).Value = True
    
    m_CoilMappings(2).Command = CMD_RESET
    m_CoilMappings(2).Address = 2
    m_CoilMappings(2).Value = True
    
    m_CoilMappings(3).Command = CMD_EMERGENCY_STOP
    m_CoilMappings(3).Address = 10
    m_CoilMappings(3).Value = True
End Sub

' 执行设备控制命令
Public Function ExecuteCommand(Cmd As DeviceCommand) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim bFound As Boolean
    
    ' 查找命令映射
    For i = 0 To UBound(m_CoilMappings)
        If m_CoilMappings(i).Command = Cmd Then
            bFound = True
            Exit For
        End If
    Next i
    
    If Not bFound Then
        Debug.Print "Command not found: " & Cmd
        ExecuteCommand = False
        Exit Function
    End If
    
    ' 写入控制线圈
    Dim bSuccess As Boolean
    bSuccess = mbMaster.WriteSingleCoil(m_CoilMappings(i).Address, m_CoilMappings(i).Value)
    
    If bSuccess Then
        Debug.Print "Command executed: " & GetCommandName(Cmd)
        ExecuteCommand = True
    Else
        Debug.Print "Command failed: " & GetCommandName(Cmd)
        ExecuteCommand = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "ExecuteCommand error: " & Err.Description
    ExecuteCommand = False
End Function

' 获取命令名称
Private Function GetCommandName(Cmd As DeviceCommand) As String
    Select Case Cmd
        Case CMD_START:          GetCommandName = "启动"
        Case CMD_STOP:           GetCommandName = "停止"
        Case CMD_RESET:          GetCommandName = "复位"
        Case CMD_EMERGENCY_STOP: GetCommandName = "紧急停止"
        Case Else:                GetCommandName = "未知"
    End Select
End Function

' 安全控制:带确认的控制
Public Function SafeExecuteCommand(Cmd As DeviceCommand, bRequireConfirm As Boolean) As Boolean
    ' 如果需要确认
    If bRequireConfirm Then
        Dim iResponse As Integer
        iResponse = MsgBox("确定要执行命令: " & GetCommandName(Cmd) & "?", _
                           vbQuestion + vbYesNo + vbDefaultButton2, _
                           "确认操作")
        
        If iResponse <> vbYes Then
            Debug.Print "Command cancelled by user"
            SafeExecuteCommand = False
            Exit Function
        End If
    End If
    
    ' 执行命令
    Dim bResult As Boolean
    bResult = ExecuteCommand(Cmd)
    
    ' 记录操作日志
    LogCommand Cmd, bResult
    
    SafeExecuteCommand = bResult
End Function

' 记录命令日志
Private Sub LogCommand(Cmd As DeviceCommand, bSuccess As Boolean)
    Dim sStatus As String
    sStatus = IIf(bSuccess, "成功", "失败")
    
    Dim sLine As String
    sLine = Format$(Now, "yyyy-mm-dd hh:mm:ss") & "," & _
             GetCommandName(Cmd) & "," & _
             sStatus
    
    ' 写入日志文件
    On Error Resume Next
    Dim iFile As Integer
    iFile = FreeFile
    Open App.Path & "\command_log.csv" For Append As #iFile
    Print #iFile, sLine
    Close #iFile
    On Error GoTo 0
End Sub
```

---

### 3. 数据网关

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster     ' 连接上位机
Private WithEvents mbSlave As cModbusSlave       ' 连接下位机

' 数据映射配置
Private Type DataMapping
    MasterAddr As Long      ' 上位机地址
    SlaveAddr As Long       ' 下位机地址
    Direction As Integer    ' 0=Master到Slave, 1=Slave到Master
    ScaleFactor As Single   ' 缩放因子
    Offset As Single        ' 偏移量
    Enabled As Boolean
End Type

Private m_Mappings() As DataMapping

' 初始化数据映射
Public Sub InitializeDataMappings()
    ReDim m_Mappings(4) As DataMapping
    
    ' 映射1: 下位机温度 -> 上位机寄存器
    m_Mappings(0).MasterAddr = 0
    m_Mappings(0).SlaveAddr = 100
    m_Mappings(0).Direction = 1  ' Slave到Master
    m_Mappings(0).ScaleFactor = 1.0
    m_Mappings(0).Offset = 0
    m_Mappings(0).Enabled = True
    
    ' 映射2: 上位机设定值 -> 下位机寄存器
    m_Mappings(1).MasterAddr = 10
    m_Mappings(1).SlaveAddr = 200
    m_Mappings(1).Direction = 0  ' Master到Slave
    m_Mappings(1).ScaleFactor = 10.0
    m_Mappings(1).Offset = 0
    m_Mappings(1).Enabled = True
    
    ' ... 更多映射
End Sub

' 网关同步
Public Sub SyncGateway()
    On Error Resume Next
    
    Dim i As Long
    
    For i = 0 To UBound(m_Mappings)
        If Not m_Mappings(i).Enabled Then GoTo NextMapping
        
        Select Case m_Mappings(i).Direction
            Case 0  ' Master到Slave
                SyncMasterToSlave i
                
            Case 1  ' Slave到Master
                SyncSlaveToMaster i
        End Select
        
NextMapping:
    Next i
End Sub

' 同步Master到Slave
Private Sub SyncMasterToSlave(Index As Long)
    Dim iMasterReg As Integer
    iMasterReg = mbMaster.ReadHoldingRegisters(m_Mappings(Index).MasterAddr, 1)(0)
    
    Dim fScaled As Single
    fScaled = (CSng(iMasterReg) * m_Mappings(Index).ScaleFactor) + m_Mappings(Index).Offset
    
    ' 写入Slave
    mbSlave.SetHoldingRegister m_Mappings(Index).SlaveAddr, CInt(fScaled)
    
    Debug.Print "Sync M[" & m_Mappings(Index).MasterAddr & "] -> S[" & _
              m_Mappings(Index).SlaveAddr & "]: " & fScaled
End Sub

' 同步Slave到Master
Private Sub SyncSlaveToMaster(Index As Long)
    Dim iSlaveReg As Integer
    iSlaveReg = mbSlave.GetHoldingRegister(m_Mappings(Index).SlaveAddr)
    
    Dim fScaled As Single
    fScaled = (CSng(iSlaveReg) * m_Mappings(Index).ScaleFactor) + m_Mappings(Index).Offset
    
    ' 写入Master
    mbMaster.WriteSingleRegister m_Mappings(Index).MasterAddr, CInt(fScaled)
    
    Debug.Print "Sync S[" & m_Mappings(Index).SlaveAddr & "] -> M[" & _
              m_Mappings(Index).MasterAddr & "]: " & fScaled
End Sub
```

---

## 常见问题

### Q1: 如何处理大数据量?

**方案**:
1. 分批读取,每次不超过最大寄存器数(125)
2. 使用定时器分时段读取
3. 实现数据缓存机制

```vb
' 分批读取大数据量
Public Sub ReadLargeData(StartAddr As Long, TotalCount As Long)
    Dim lOffset As Long
    Dim lRemaining As Long
    Dim lBatchSize As Long
    Dim iBatch As Long
    
    lBatchSize = mbMaster.Defaults.MAX_REGISTERS
    lOffset = 0
    lRemaining = TotalCount
    
    Do While lRemaining > 0
        Dim lThisBatch As Long
        lThisBatch = IIf(lRemaining > lBatchSize, lBatchSize, lRemaining)
        
        Dim iRegs() As Integer
        iRegs = mbMaster.ReadHoldingRegisters(StartAddr + lOffset, lThisBatch)
        
        ' 处理数据
        ProcessBatch iRegs, lOffset
        
        lOffset = lOffset + lThisBatch
        lRemaining = lRemaining - lThisBatch
        
        iBatch = iBatch + 1
        Debug.Print "Batch " & iBatch & " completed"
        
        ' 避免过快请求
        Sleep 10
    Loop
End Sub
```

---

### Q2: 如何实现热备份?

**方案**:
1. 同时连接主从两台设备
2. 主设备正常时读取主设备
3. 主设备故障时自动切换到备用设备

```vb
Option Explicit

Private WithEvents mbMasterPrimary As cModbusMaster
Private WithEvents mbMasterBackup As cModbusMaster
Private m_UsePrimary As Boolean

' 切换到备用设备
Private Sub SwitchToBackup()
    On Error Resume Next
    
    mbMasterPrimary.Disconnect
    
    Dim iRegs() As Integer
    iRegs = mbMasterBackup.ReadHoldingRegisters(0, 1)
    
    If UBound(iRegs) >= 0 Then
        m_UsePrimary = False
        Debug.Print "Switched to backup device"
        RaiseEvent DeviceSwitched("Backup")
    Else
        Debug.Print "Backup device also failed"
        RaiseEvent AllDevicesFailed
    End If
End Sub

' 尝试切换回主设备
Private Sub TrySwitchBack()
    On Error Resume Next
    
    mbMasterPrimary.Connect
    
    Dim iRegs() As Integer
    iRegs = mbMasterPrimary.ReadHoldingRegisters(0, 1)
    
    If UBound(iRegs) >= 0 Then
        m_UsePrimary = True
        Debug.Print "Switched back to primary device"
        RaiseEvent DeviceSwitched("Primary")
    End If
End Sub

' 智能读取(自动选择设备)
Public Function SmartRead(StartAddr As Long, Quantity As Long) As Integer()
    Dim iRegs() As Integer
    
    If m_UsePrimary Then
        ' 尝试从主设备读取
        On Error Resume Next
        iRegs = mbMasterPrimary.ReadHoldingRegisters(StartAddr, Quantity)
        
        If UBound(iRegs) < 0 Then
            ' 主设备失败,切换到备用
            SwitchToBackup
            If Not m_UsePrimary Then
                iRegs = mbMasterBackup.ReadHoldingRegisters(StartAddr, Quantity)
            End If
        End If
    Else
        ' 从备用设备读取
        iRegs = mbMasterBackup.ReadHoldingRegisters(StartAddr, Quantity)
        
        ' 定期尝试切换回主设备
        Static lSwitchCount As Long
        lSwitchCount = lSwitchCount + 1
        If lSwitchCount > 10 Then
            TrySwitchBack
            lSwitchCount = 0
        End If
    End If
    
    SmartRead = iRegs
End Function
```

---

### Q3: 如何实现断线重连?

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents tmrReconnect As Timer
Private m_bAutoReconnect As Boolean
Private m_lReconnectInterval As Long  ' 毫秒

' 启用自动重连
Public Sub EnableAutoReconnect(bEnable As Boolean, lInterval As Long)
    m_bAutoReconnect = bEnable
    m_lReconnectInterval = lInterval
    
    Set tmrReconnect = New Timer
    tmrReconnect.Interval = lInterval
End Sub

Private Sub mbMaster_OnDisconnect()
    If m_bAutoReconnect Then
        Debug.Print "Disconnected, will reconnect in " & (m_lReconnectInterval \ 1000) & "s"
        tmrReconnect.Enabled = True
    End If
End Sub

Private Sub tmrReconnect_Timer()
    tmrReconnect.Enabled = False
    
    On Error Resume Next
    mbMaster.Connect
    
    If Err.Number = 0 And mbMaster.State = MB_MASTER_STATE_CONNECTED Then
        Debug.Print "Reconnected successfully"
    Else
        ' 重连失败,继续尝试
        tmrReconnect.Enabled = True
    End If
End Sub
```

---

## 下一步

- 查看 [master.md](./master.md) 了解主站详细 API
- 查看 [slave.md](./slave.md) 了解从站详细 API  
- 查看 [quickstart.md](./quickstart.md) 快速入门
- 查看 [overview.md](./overview.md) 类库总览

---

**最后更新**: 2026-01-16

### 更新日志

#### 2026-01-16 (v1.1.0)
- 更新设备配置结构体，使用 `ModbusMasterProtocolType`
- 添加 `BindAddress` 配置示例（从站）
- 更新枚举引用（适配 v1.1.0 命名规范）
