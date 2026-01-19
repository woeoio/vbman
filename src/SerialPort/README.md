# cSerialPort - 高级串口通信类

## 简介

`cSerialPort` 是一个使用 Windows 系统 API 实现的高级串口通信类，支持异步操作和事件通知机制。

## 特性

- ✅ **纯 API 实现**：不依赖 MSComm 控件或其他第三方组件
- ✅ **异步 IO**：支持异步读写操作，不阻塞主线程
- ✅ **事件驱动**：提供完整的事件通知机制
- ✅ **高波特率**：最高支持 921600 波特率
- ✅ **流控制**：支持无流控、软件流控(XOn/XOff)、硬件流控(RTS/CTS)
- ✅ **信号控制**：支持 DTR/RTS/Break 信号控制
- ✅ **缓冲区管理**：可自定义输入/输出缓冲区大小
- ✅ **运行时重配置**：支持在打开状态下重新配置参数

## 支持的波特率

- 110, 300, 600, 1200, 2400, 4800, 9600
- 14400, 19200, 38400, 56000, 57600
- 115200, 128000, 230400, 256000, 460800, 921600

## 基本使用

### 1. 创建和打开串口

```vb
Dim sp As New cSerialPort

' 打开串口
If sp.Open("COM1", 115200, 8, peNone, sb1, fcNone) Then
    Debug.Print "串口打开成功"
    
    ' 启动事件监听
    sp.StartListening
    
    ' 发送数据
    sp.WriteString "Hello World"
    
    ' 关闭串口
    sp.Close
End If
```

### 2. 使用事件 (需在窗体模块中使用 WithEvents)

```vb
Private WithEvents m_SerialPort As cSerialPort

Private Sub Form_Load()
    Set m_SerialPort = New cSerialPort
    m_SerialPort.Open "COM1", 115200
    m_SerialPort.StartListening
End Sub

' 数据接收事件
Private Sub m_SerialPort_OnDataReceived(ByVal Data As String, ByVal Length As Long)
    Debug.Print "收到数据: " & Data
End Sub

' 错误事件
Private Sub m_SerialPort_OnError(ByVal ErrorCode As Long, ByVal ErrorMessage As String)
    Debug.Print "错误: " & ErrorMessage
End Sub
```

## API 参考

### 主要方法

#### `Open` - 打开串口

```vb
Function Open(PortName As String, _
              Optional BaudRate As Long = 9600, _
              Optional DataBits As Integer = 8, _
              Optional Parity As ParityEnum = peNone, _
              Optional StopBits As StopBitsEnum = sb1, _
              Optional FlowControl As FlowControlEnum = fcNone) As Boolean
```

**参数：**
- `PortName`: 串口名称，如 "COM1"
- `BaudRate`: 波特率，支持 110 到 921600
- `DataBits`: 数据位 (4-8)
- `Parity`: 校验位 (peNone, peOdd, peEven, peMark, peSpace)
- `StopBits`: 停止位 (sb1, sb1_5, sb2)
- `FlowControl`: 流控制 (fcNone, fcXOnXOff, fcRTSCTS, fcRTSCTSXOnXOff)

**返回值：** 成功返回 True，失败返回 False

#### `Close` - 关闭串口

```vb
Sub Close()
```

#### `WriteString` - 发送字符串

```vb
Function WriteString(Data As String) As Long
```

**返回值：** 实际发送的字节数

#### `WriteBytes` - 发送字节数组

```vb
Function WriteBytes(Data() As Byte, Optional Length As Long = -1) As Long
```

**参数：**
- `Data`: 字节数组
- `Length`: 要发送的字节数，-1 表示发送整个数组

**返回值：** 实际发送的字节数

#### `ReadString` - 读取字符串

```vb
Function ReadString(Optional MaxBytes As Long = 4096) As String
```

**返回值：** 读取的字符串

#### `ReadBytes` - 读取字节数组

```vb
Function ReadBytes(Optional MaxBytes As Long = 4096) As Byte()
```

**返回值：** 字节数组

#### `WriteHex` - 发送 HEX 字符串

```vb
Function WriteHex(HexString As String) As Long
```

**参数：**
- `HexString`: HEX 字符串，支持各种分隔符（如 "A1B2C3", "A1 B2 C3", "A1-B2-C3"）

**返回值：** 实际发送的字节数

**说明：** 自动清理分隔符，将 HEX 字符串转换为字节数组后发送

#### `ReadHex` - 读取 HEX 字符串

```vb
Function ReadHex(Optional MaxBytes As Long = 4096, Optional Separator As String = " ") As String
```

**参数：**
- `MaxBytes`: 最大读取字节数，默认 4096
- `Separator`: HEX 分隔符，默认空格 " "（可设为 "" 表示无分隔符）

**返回值：** HEX 字符串，如 "A1 B2 C3 D4"

#### `StartListening` - 启动事件监听

```vb
Sub StartListening()
```

#### `StopListening` - 停止事件监听

```vb
Sub StopListening()
```

#### `ClearBuffer` - 清空所有缓冲区

```vb
Sub ClearBuffer()
```

#### `ClearSendBuffer` - 清空发送缓冲区

```vb
Sub ClearSendBuffer()
```

#### `ClearReceiveBuffer` - 清空接收缓冲区

```vb
Sub ClearReceiveBuffer()
```

#### `Flush` - 刷新缓冲区

```vb
Function Flush() As Boolean
```

#### `ReConfigure` - 重新配置串口

```vb
Function ReConfigure(Optional BaudRate As Long = -1, _
                     Optional DataBits As Integer = -1, _
                     Optional Parity As ParityEnum = -1, _
                     Optional StopBits As StopBitsEnum = -1, _
                     Optional FlowControl As FlowControlEnum = -1) As Boolean
```

### 信号控制方法

#### `SetDTR` / `ClearDTR`

```vb
Sub SetDTR()
Sub ClearDTR()
```

#### `SetRTS` / `ClearRTS`

```vb
Sub SetRTS()
Sub ClearRTS()
```

#### `SetBreak` / `ClearBreak`

```vb
Sub SetBreak()
Sub ClearBreak()
```

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `IsOpen` | Boolean | 串口是否打开 (只读) |
| `PortName` | String | 串口名称 (只读) |
| `BaudRate` | Long | 波特率 (只读) |
| `DataBits` | Integer | 数据位 (只读) |
| `StopBits` | Integer | 停止位 (只读) |
| `Parity` | Integer | 校验位 (只读) |
| `FlowControl` | Integer | 流控制 (只读) |
| `UseAsyncIO` | Boolean | 是否使用异步IO (可读写) |
| `InQueueSize` | Long | 输入缓冲区大小 (可读写) |
| `OutQueueSize` | Long | 输出缓冲区大小 (可读写) |
| `Listening` | Boolean | 是否正在监听事件 (只读) |

### 缓冲区状态查询

#### `BytesInQueue` - 获取接收缓冲区字节数

```vb
Function BytesInQueue() As Long
```

#### `BytesOutQueue` - 获取发送缓冲区字节数

```vb
Function BytesOutQueue() As Long
```

## 事件

### `OnDataReceived` - 数据接收事件 (字符串格式)

```vb
Event OnDataReceived(Data As String, Length As Long)
```

### `OnDataReceivedBytes` - 数据接收事件 (字节数组格式)

```vb
Event OnDataReceivedBytes(Data() As Byte, Length As Long)
```

### `OnError` - 错误事件

```vb
Event OnError(ErrorCode As Long, ErrorMessage As String)
```

### `OnCTSChanged` - CTS 信号变化

```vb
Event OnCTSChanged(State As Boolean)
```

### `OnDSRChanged` - DSR 信号变化

```vb
Event OnDSRChanged(State As Boolean)
```

### `OnRing` - 振铃事件

```vb
Event OnRing(State As Boolean)
```

### `OnRLSDChanged` - RLSD (DCD) 信号变化

```vb
Event OnRLSDChanged(State As Boolean)
```

### `OnBreak` - 中断事件

```vb
Event OnBreak()
```

### `OnTxEmpty` - 发送缓冲区空事件

```vb
Event OnTxEmpty()
```

## 错误码

| 错误码 | 说明 |
|--------|------|
| `CE_RXOVER` | 接收溢出 |
| `CE_OVERRUN` | 超时错误 |
| `CE_RXPARITY` | 校验错误 |
| `CE_FRAME` | 帧错误 |
| `CE_BREAK` | 中断错误 |
| `CE_TXFULL` | 发送缓冲区满 |

## 使用建议

### 1. 高波特率设置

当使用高波特率 (如 460800, 921600) 时，建议增大缓冲区：

```vb
sp.InQueueSize = 16384
sp.OutQueueSize = 16384
sp.Open "COM1", 921600
```

### 2. 流控制选择

- **无流控 (fcNone)**: 短距离通信，数据量不大
- **软件流控 (fcXOnXOff)**: 不使用硬件握手时
- **硬件流控 (fcRTSCTS)**: 高速、大数据量通信，推荐使用

### 3. 异步 IO

默认启用异步 IO，适合需要同时处理其他任务的场景。如果需要同步操作：

```vb
sp.UseAsyncIO = False
sp.Open "COM1", 9600
```

### 4. 事件监听

如果要使用事件通知，必须在打开串口后调用 `StartListening()`：

```vb
sp.Open "COM1", 9600
sp.StartListening  ' 启动事件监听
```

## 常见问题

### Q: 为什么收不到数据？

A: 检查以下几点：
1. 是否调用了 `StartListening()` 启动事件监听
2. 串口参数是否匹配（波特率、数据位、校验位、停止位）
3. 是否有其他程序占用了该串口

### Q: 高波特率通信不稳定？

A: 建议使用硬件流控 (fcRTSCTS) 并增大缓冲区：

```vb
sp.InQueueSize = 32768
sp.OutQueueSize = 32768
sp.Open "COM1", 921600, 8, peNone, sb1, fcRTSCTS
```

### Q: 如何检测可用的串口？

A: 可以枚举串口，参考示例代码中的 `Example_ListAvailablePorts` 方法。

### Q: 可以同时打开多个串口吗？

A: 可以，创建多个 `cSerialPort` 实例即可：

```vb
Dim sp1 As New cSerialPort
Dim sp2 As New cSerialPort

sp1.Open "COM1", 115200
sp2.Open "COM2", 9600
```

## 完整示例

参考 `cSerialPort_示例.bas` 文件中的详细示例代码，包括：
- 基本使用
- 高波特率使用
- 事件处理
- 二进制数据收发
- HEX 字符串收发
- 流控制设置
- 动态重配置
- 缓冲区管理
- 信号控制
- Modbus RTU 通信示例

## 技术细节

### API 使用的核心函数

- `CreateFile` - 打开串口设备
- `SetCommState` / `GetCommState` - 配置/获取串口参数
- `SetCommTimeouts` - 配置超时参数
- `WriteFile` / `ReadFile` - 读写数据
- `WaitCommEvent` - 等待串口事件
- `SetCommMask` - 设置事件掩码
- `EscapeCommFunction` - 控制信号线
- `PurgeComm` - 清空缓冲区
- `ClearCommError` - 清除错误并获取状态

### 异步 IO 实现原理

使用 `OVERLAPPED` 结构体和 Windows 事件对象实现异步 IO：
1. 读写操作时传入 `OVERLAPPED` 结构
2. 操作返回 `ERROR_IO_PENDING` 表示异步操作进行中
3. 通过 `WaitForSingleObject` 等待操作完成
4. 通过 `GetOverlappedResult` 获取实际传输的字节数

## 许可证

本项目为开源项目，可自由使用和修改。

## 更新日志

### v1.0.0
- 初始版本
- 支持基本串口功能
- 支持异步 IO
- 支持事件通知
- 最高支持 921600 波特率
