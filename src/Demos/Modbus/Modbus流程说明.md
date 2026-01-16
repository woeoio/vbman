# Modbus 读写流程详解

## 一、Modbus 通信模式说明

**重要概念：Modbus 是主从（Master-Slave）通信模式**

- **主站（Master）**：主动发起请求，控制通信
- **从站（Slave）**：被动响应请求，不能主动发起通信

**重要说明：**
- **在当前代码库的实现中**：PC 作为主站，PLC 作为从站
- **Modbus 协议限制**：从站不能主动发送数据，必须由主站主动请求
- **角色可以互换**：PLC 也可以作为主站，PC 也可以作为从站（需要实现 Modbus Server 功能，详见 `Modbus角色互换讨论.md`）

**对于当前使用场景（PC 主站，PLC 从站）**：
- PLC 不会主动发送数据给你
- 你必须主动调用读取函数来获取 PLC 的数据

## 二、读取流程（接收 PLC 数据）

### 流程步骤：

```
1. 你的程序调用读取函数
   ↓
2. 构建 Modbus 请求帧（包含：从站ID、功能码、起始地址、数量）
   ↓
3. 通过 TCP/RTU 发送请求到 PLC
   ↓
4. PLC 收到请求后，从内部寄存器读取数据
   ↓
5. PLC 构建响应帧并发送回来
   ↓
6. 你的程序接收响应数据
   ↓
7. 解析响应，提取数据（寄存器值、线圈状态等）
```

### 代码示例：

```vb
' 1. 创建 Modbus 对象并连接
Dim mb As New cModbus
mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
mb.SlaveID = 1
mb.TCPHost = "192.168.1.100"
mb.TCPPort = 502
mb.Connect

' 2. 读取保持寄存器（PLC 的数据）
Dim regs() As Integer
regs = mb.ReadHoldingRegisters(0, 10)  ' 从地址0读取10个寄存器

' 3. 处理接收到的数据
Dim i As Long
For i = 0 To UBound(regs)
    Debug.Print "寄存器 " & i & " 的值: " & regs(i)
Next i

' 4. 断开连接
mb.Disconnect
```

### 内部实现细节（参考 cModbus.cls）：

1. **构建请求帧** (`BuildReadHoldingRegistersRequest`)
   - 从站ID（SlaveID）
   - 功能码（03 = 读保持寄存器）
   - 起始地址（高字节 + 低字节）
   - 数量（高字节 + 低字节）

2. **发送请求** (`SendRequest` → `SendRequestTCP` 或 `SendRequestRTU`)
   - TCP 模式：添加 MBAP 头，通过 Socket 发送
   - RTU 模式：添加 CRC 校验，通过串口发送

3. **接收响应** (`SendRequestTCP` 或 `SendRequestRTU`)
   - 等待 PLC 响应（带超时控制）
   - 接收数据包
   - 验证响应格式

4. **解析响应** (`ReadHoldingRegisters`)
   - 提取字节数
   - 将字节数据转换为整数数组
   - 返回结果

## 三、写入流程（发送数据到 PLC）

### 流程步骤：

```
1. 你的程序调用写入函数
   ↓
2. 构建 Modbus 写入请求帧（包含：从站ID、功能码、地址、数据）
   ↓
3. 通过 TCP/RTU 发送请求到 PLC
   ↓
4. PLC 收到请求后，将数据写入内部寄存器
   ↓
5. PLC 构建确认响应并发送回来
   ↓
6. 你的程序接收确认响应
   ↓
7. 验证写入是否成功
```

### 代码示例：

```vb
' 1. 连接（同上）
Dim mb As New cModbus
mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
mb.SlaveID = 1
mb.TCPHost = "192.168.1.100"
mb.TCPPort = 502
mb.Connect

' 2. 写入单个寄存器
Dim bSuccess As Boolean
bSuccess = mb.WriteSingleRegister(0, 1234)  ' 将值 1234 写入地址 0

If bSuccess Then
    Debug.Print "写入成功"
Else
    Debug.Print "写入失败"
End If

' 3. 写入多个寄存器
Dim values(4) As Integer
values(0) = 100
values(1) = 200
values(2) = 300
values(3) = 400
values(4) = 500
bSuccess = mb.WriteMultipleRegisters(10, values)  ' 从地址10开始写入5个寄存器

mb.Disconnect
```

## 四、发送当前时间到 PLC 的完整示例

### 方法1：将时间转换为整数写入单个寄存器

```vb
Public Sub SendCurrentTimeToPLC()
    Dim mb As New cModbus
    Dim bSuccess As Boolean
    Dim lTimeValue As Long
    
    On Error GoTo ErrorHandler
    
    ' 配置连接
    mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    ' 连接
    mb.Connect
    
    ' 获取当前时间（转换为秒数，从某个基准时间开始）
    ' 例如：从 1970-01-01 开始的秒数，或从今天 00:00:00 开始的秒数
    Dim dtNow As Date
    dtNow = Now
    
    ' 方法1：将时间转换为从今天 00:00:00 开始的秒数
    Dim dtToday As Date
    dtToday = DateValue(dtNow)
    lTimeValue = CLng((dtNow - dtToday) * 86400)  ' 86400 = 24*60*60 秒
    
    ' 写入到寄存器（注意：单个寄存器只能存储 0-65535，所以秒数不能超过 65535）
    ' 如果超过，需要拆分成多个寄存器
    If lTimeValue <= 65535 Then
        bSuccess = mb.WriteSingleRegister(0, CInt(lTimeValue))
        Debug.Print "当前时间（秒数）: " & lTimeValue & " 写入 " & IIf(bSuccess, "成功", "失败")
    Else
        Debug.Print "时间值超出单个寄存器范围，需要使用多个寄存器"
    End If
    
    mb.Disconnect
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub
```

### 方法2：将时间拆分为多个寄存器（年、月、日、时、分、秒）

```vb
Public Sub SendDateTimeToPLC()
    Dim mb As New cModbus
    Dim values(5) As Integer
    Dim dtNow As Date
    Dim bSuccess As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 配置连接
    mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    ' 连接
    mb.Connect
    
    ' 获取当前时间
    dtNow = Now
    
    ' 将时间拆分为多个寄存器
    values(0) = Year(dtNow)      ' 年（如 2026）
    values(1) = Month(dtNow)     ' 月（1-12）
    values(2) = Day(dtNow)       ' 日（1-31）
    values(3) = Hour(dtNow)      ' 时（0-23）
    values(4) = Minute(dtNow)    ' 分（0-59）
    values(5) = Second(dtNow)    ' 秒（0-59）
    
    ' 写入多个寄存器（从地址 0 开始）
    bSuccess = mb.WriteMultipleRegisters(0, values)
    
    If bSuccess Then
        Debug.Print "时间写入成功: " & Format$(dtNow, "yyyy-mm-dd hh:nn:ss")
        Debug.Print "  年: " & values(0)
        Debug.Print "  月: " & values(1)
        Debug.Print "  日: " & values(2)
        Debug.Print "  时: " & values(3)
        Debug.Print "  分: " & values(4)
        Debug.Print "  秒: " & values(5)
    Else
        Debug.Print "时间写入失败"
    End If
    
    mb.Disconnect
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub
```

### 方法3：将时间戳（Unix 时间戳）写入两个寄存器

```vb
Public Sub SendUnixTimestampToPLC()
    Dim mb As New cModbus
    Dim values(1) As Integer
    Dim dtNow As Date
    Dim lTimestamp As Long
    Dim bSuccess As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 配置连接
    mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    ' 连接
    mb.Connect
    
    ' 计算 Unix 时间戳（从 1970-01-01 00:00:00 开始的秒数）
    dtNow = Now
    Dim dtEpoch As Date
    dtEpoch = #1/1/1970#
    lTimestamp = CLng((dtNow - dtEpoch) * 86400)
    
    ' 将 32 位时间戳拆分为两个 16 位寄存器
    ' 高 16 位
    values(0) = CInt((lTimestamp \ 65536) And &HFFFF)
    ' 低 16 位
    values(1) = CInt(lTimestamp And &HFFFF)
    
    ' 写入两个寄存器（从地址 0 开始）
    bSuccess = mb.WriteMultipleRegisters(0, values)
    
    If bSuccess Then
        Debug.Print "Unix 时间戳写入成功: " & lTimestamp
        Debug.Print "  高16位（寄存器0）: " & values(0)
        Debug.Print "  低16位（寄存器1）: " & values(1)
    Else
        Debug.Print "时间戳写入失败"
    End If
    
    mb.Disconnect
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub
```

## 五、完整的数据交互示例

### 示例：读取 PLC 数据，然后发送当前时间

```vb
Public Sub ReadAndWriteExample()
    Dim mb As New cModbus
    Dim regs() As Integer
    Dim values(5) As Integer
    Dim dtNow As Date
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' 配置连接
    mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    ' 连接
    Debug.Print "正在连接..."
    mb.Connect
    Debug.Print "连接成功"
    
    ' ========== 读取 PLC 数据 ==========
    Debug.Print vbCrLf & "=== 读取 PLC 数据 ==="
    regs = mb.ReadHoldingRegisters(0, 10)
    Debug.Print "读取到 " & (UBound(regs) + 1) & " 个寄存器的值:"
    For i = 0 To UBound(regs)
        Debug.Print "  寄存器 " & i & ": " & regs(i)
    Next i
    
    ' ========== 发送当前时间到 PLC ==========
    Debug.Print vbCrLf & "=== 发送当前时间到 PLC ==="
    dtNow = Now
    Debug.Print "当前时间: " & Format$(dtNow, "yyyy-mm-dd hh:nn:ss")
    
    ' 将时间拆分为多个寄存器
    values(0) = Year(dtNow)
    values(1) = Month(dtNow)
    values(2) = Day(dtNow)
    values(3) = Hour(dtNow)
    values(4) = Minute(dtNow)
    values(5) = Second(dtNow)
    
    ' 写入到寄存器地址 100（避免覆盖之前的数据）
    If mb.WriteMultipleRegisters(100, values) Then
        Debug.Print "时间写入成功"
    Else
        Debug.Print "时间写入失败"
    End If
    
    ' 验证写入：重新读取
    Debug.Print vbCrLf & "=== 验证写入 ==="
    regs = mb.ReadHoldingRegisters(100, 6)
    Debug.Print "从地址 100 读取到的值:"
    For i = 0 To UBound(regs)
        Debug.Print "  寄存器 " & (100 + i) & ": " & regs(i)
    Next i
    
    ' 断开连接
    mb.Disconnect
    Debug.Print vbCrLf & "已断开连接"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub
```

## 六、关键点总结

### 1. Modbus 是请求-响应模式
- **PLC 不会主动发送数据**
- 你必须主动调用读取函数来获取 PLC 的数据
- 读取函数内部会：发送请求 → 等待响应 → 解析数据 → 返回结果

### 2. 读取数据（接收 PLC 数据）
```vb
' 调用读取函数即可，内部自动完成请求-响应流程
Dim regs() As Integer
regs = mb.ReadHoldingRegisters(起始地址, 数量)
' regs 数组就是接收到的 PLC 数据
```

### 3. 写入数据（发送数据到 PLC）
```vb
' 调用写入函数即可，内部自动完成请求-响应流程
Dim bSuccess As Boolean
bSuccess = mb.WriteSingleRegister(地址, 值)
' 或
bSuccess = mb.WriteMultipleRegisters(起始地址, 值数组)
```

### 4. 数据类型限制
- **寄存器（Register）**：16 位整数，范围 -32768 到 32767（或 0 到 65535，取决于 PLC）
- **线圈（Coil）**：布尔值，True/False
- 如果要发送更大的数据，需要拆分成多个寄存器

### 5. 地址说明
- Modbus 地址从 0 开始
- 不同 PLC 厂商可能使用不同的地址映射
- 需要查阅 PLC 的 Modbus 地址表

## 七、常见问题

### Q1: PLC 主动发送数据给我，我如何接收？
**A:** 这取决于通信角色：
- **如果 PC 是主站，PLC 是从站**（当前代码库的场景）：
  - Modbus 协议不支持从站主动发送数据
  - 你需要使用定时器定期调用读取函数来获取 PLC 数据
  - 或者使用其他协议（如 MQTT、OPC）实现主动推送
- **如果 PLC 是主站，PC 是从站**：
  - PLC 可以主动读取 PC 的数据（需要 PC 实现 Modbus Server 功能）
  - 详见 `Modbus角色互换讨论.md`

### Q2: 如何发送浮点数？
**A:** 需要将浮点数转换为整数：
- 方法1：乘以倍数后转为整数（如 123.45 → 12345，保留2位小数）
- 方法2：使用 IEEE 754 格式，拆分为两个寄存器（高16位 + 低16位）

### Q3: 如何发送字符串？
**A:** 将字符串的每个字符转换为 ASCII 码，每个字符占用一个寄存器

### Q4: 读取/写入失败怎么办？
**A:** 检查：
- 连接是否正常
- 从站ID是否正确
- 地址是否在有效范围内
- 寄存器是否可读/可写
- 超时时间是否足够
