# cModbusMaster 类参考

## ? 类概述

`cModbusMaster` 是 Modbus 主站（客户端）实现类，提供连接到 Modbus 从站、发送/接收数据、自动处理协议帧等功能。

---

## ? 事件列表

| 事件名 | 触发时机 | 参数 |
|--------|----------|------|
| `OnConnect` | 连接成功建立 | 无 |
| `OnDisconnect` | 连接已关闭 | 无 |
| `OnError` | 发生错误 | `Description` (错误描述) |
| `OnDataReceived` | 收到数据（调试用） | `Data()` (字节数组) |

---

## ? 属性参考

### ProtocolType - 协议类型

**类型**: `ModbusMasterProtocolType` (枚举)  
**读写**: 读写

**值**:

| 常量 | 值 | 说明 |
|------|-----|------|
| `MB_PROTOCOL_RTU` | 1 | RTU 模式（串口通信） |
| `MB_PROTOCOL_TCP` | 2 | TCP 模式（网络通信） |

**示例**:

```vb
' 设置为 TCP 模式
mbMaster.ProtocolType = MB_PROTOCOL_TCP

' 设置为 RTU 模式
mbMaster.ProtocolType = MB_PROTOCOL_RTU
```

---

### State - 连接状态

**类型**: `ModbusMasterState` (枚举)  
**读写**: 只读

**值**:

| 常量 | 值 | 说明 |
|------|-----|------|
| `MB_STATE_DISCONNECTED` | 0 | 已断开 |
| `MB_STATE_CONNECTING` | 1 | 正在连接 |
| `MB_STATE_CONNECTED` | 2 | 已连接 |
| `MB_STATE_ERROR` | 3 | 错误状态 |

**示例**:

```vb
If mbMaster.State = MB_STATE_CONNECTED Then
    Debug.Print "已连接"
Else
    Debug.Print "未连接"
End If
```

---

### SlaveID - 从站 ID

**类型**: `Byte`  
**读写**: 读写

**说明**: 目标从站的设备地址（1-247）。0 表示广播地址。

**示例**:

```vb
' 设置从站 ID 为 1
mbMaster.SlaveID = 1

' 广播到所有从站
mbMaster.SlaveID = 0
```

---

### ResponseTimeout - 响应超时

**类型**: `Long`  
**读写**: 读写

**说明**: 等待从站响应的超时时间（毫秒）。默认为 1000ms。

**示例**:

```vb
' 设置超时为 3 秒
mbMaster.ResponseTimeout = 3000

' 恢复默认值
mbMaster.ResponseTimeout = 1000
```

---

### Defaults - 默认常量

**类型**: `ModbusMasterDefaults` (结构体)  
**读写**: 只读

**字段**:

| 字段 | 类型 | 说明 |
|------|------|------|
| `TCP_PORT` | Long | 默认 TCP 端口（502） |
| `RTU_DEFAULT_BAUDRATE` | Long | 默认波特率（9600） |
| `RTU_DEFAULT_DATABITS` | Long | 默认数据位（8） |
| `RTU_DEFAULT_PARITY` | String | 默认校验位（"N"） |
| `RTU_DEFAULT_STOPBITS` | Long | 默认停止位（1） |
| `RTU_DEFAULT_TIMEOUT` | Long | 默认超时（1000ms） |
| `MAX_PDU_SIZE` | Long | 最大 PDU 大小（253） |
| `MAX_REGISTERS` | Long | 最大寄存器数量（125） |
| `MAX_COILS` | Long | 最大线圈数量（2000） |
| `TCP_MBAP_SIZE` | Long | MBAP 头大小（7） |

**示例**:

```vb
Debug.Print "默认端口: " & mbMaster.Defaults.TCP_PORT
Debug.Print "最大寄存器: " & mbMaster.Defaults.MAX_REGISTERS
```

---

### RTU 模式属性

#### SerialPort - 串口名称

**类型**: `String`  
**读写**: 读写

**示例**:

```vb
mbMaster.SerialPort = "COM1"
```

#### BaudRate - 波特率

**类型**: `Long`  
**读写**: 读写

**示例**:

```vb
mbMaster.BaudRate = 9600
```

#### DataBits - 数据位

**类型**: `Long`  
**读写**: 读写

**示例**:

```vb
mbMaster.DataBits = 8
```

#### Parity - 校验位

**类型**: `String`  
**读写**: 读写

**值**: "N" (无), "E" (偶校验), "O" (奇校验)

**示例**:

```vb
mbMaster.Parity = "N"  ' 无校验
mbMaster.Parity = "E"  ' 偶校验
mbMaster.Parity = "O"  ' 奇校验
```

#### StopBits - 停止位

**类型**: `Long`  
**读写**: 读写

**示例**:

```vb
mbMaster.StopBits = 1  ' 1 个停止位
mbMaster.StopBits = 2  ' 2 个停止位
```

---

### TCP 模式属性

#### TCPHost - TCP 主机地址

**类型**: `String`  
**读写**: 读写

**示例**:

```vb
mbMaster.TCPHost = "192.168.1.100"
mbMaster.TCPHost = "127.0.0.1"
```

#### TCPPort - TCP 端口

**类型**: `Long`  
**读写**: 读写

**示例**:

```vb
mbMaster.TCPPort = 502  ' Modbus 默认端口
```

---

## ? 方法参考

### Connect - 连接从站

**语法**:

```vb
Public Sub Connect(Optional ByVal SerialPort As String = "", _
                   Optional ByVal TCPHost As String = "", _
                   Optional ByVal TCPPort As Long = 0)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `SerialPort` | String (可选） | RTU 模式：串口名称（COM1, COM2...） |
| `TCPHost` | String (可选） | TCP 模式：主机地址或 IP |
| `TCPPort` | Long (可选） | TCP 模式：端口号 |

**示例**:

```vb
' TCP 连接
mbMaster.ProtocolType = MB_PROTOCOL_TCP
mbMaster.TCPHost = "192.168.1.100"
mbMaster.TCPPort = 502
mbMaster.SlaveID = 1
mbMaster.Connect

' RTU 连接
mbMaster.ProtocolType = MB_PROTOCOL_RTU
mbMaster.SerialPort = "COM1"
mbMaster.BaudRate = 9600
mbMaster.DataBits = 8
mbMaster.Parity = "N"
mbMaster.StopBits = 1
mbMaster.SlaveID = 1
mbMaster.Connect "COM1"
```

---

### Disconnect - 断开连接

**语法**:

```vb
Public Sub Disconnect()
```

**示例**:

```vb
mbMaster.Disconnect
```

---

### ReadCoils - 读取线圈 (0x01)

**语法**:

```vb
Public Function ReadCoils(ByVal StartAddress As Long, ByVal Quantity As Long) As Boolean()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `StartAddress` | Long | 起始线圈地址（0-based） |
| `Quantity` | Long | 要读取的线圈数量（1-2000） |

**返回值**: `Boolean()` - 线圈值数组

**示例**:

```vb
Dim baCoils() As Boolean
baCoils = mbMaster.ReadCoils(0, 10)

Dim i As Long
For i = 0 To UBound(baCoils)
    Debug.Print "Coil[" & i & "] = " & baCoils(i)
Next i
```

---

### ReadDiscreteInputs - 读取离散输入 (0x02)

**语法**:

```vb
Public Function ReadDiscreteInputs(ByVal StartAddress As Long, ByVal Quantity As Long) As Boolean()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `StartAddress` | Long | 起始离散输入地址（0-based） |
| `Quantity` | Long | 要读取的离散输入数量（1-2000） |

**返回值**: `Boolean()` - 离散输入值数组

**示例**:

```vb
Dim baInputs() As Boolean
baInputs = mbMaster.ReadDiscreteInputs(0, 10)

Dim i As Long
For i = 0 To UBound(baInputs)
    Debug.Print "Input[" & i & "] = " & baInputs(i)
Next i
```

---

### ReadHoldingRegisters - 读取保持寄存器 (0x03)

**语法**:

```vb
Public Function ReadHoldingRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `StartAddress` | Long | 起始寄存器地址（0-based） |
| `Quantity` | Long | 要读取的寄存器数量（1-125） |

**返回值**: `Integer()` - 16位寄存器值数组

**示例**:

```vb
Dim iRegs() As Integer
iRegs = mbMaster.ReadHoldingRegisters(0, 10)

Dim i As Long
For i = 0 To UBound(iRegs)
    Debug.Print "Reg[" & i & "] = " & iRegs(i)
Next i
```

---

### ReadInputRegisters - 读取输入寄存器 (0x04)

**语法**:

```vb
Public Function ReadInputRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `StartAddress` | Long | 起始寄存器地址（0-based） |
| `Quantity` | Long | 要读取的寄存器数量（1-125） |

**返回值**: `Integer()` - 16位寄存器值数组

**示例**:

```vb
Dim iRegs() As Integer
iRegs = mbMaster.ReadInputRegisters(0, 10)

Dim i As Long
For i = 0 To UBound(iRegs)
    Debug.Print "InputReg[" & i & "] = " & iRegs(i)
Next i
```

---

### WriteSingleCoil - 写入单个线圈 (0x05)

**语法**:

```vb
Public Function WriteSingleCoil(ByVal Address As Long, ByVal Value As Boolean) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 线圈地址 |
| `Value` | Boolean | 要写入的值（True/False） |

**返回值**: `Boolean` - 成功返回 True，失败返回 False

**示例**:

```vb
Dim bSuccess As Boolean
bSuccess = mbMaster.WriteSingleCoil(0, True)

If bSuccess Then
    Debug.Print "写入成功"
Else
    Debug.Print "写入失败"
End If
```

---

### WriteMultipleCoils - 写入多个线圈 (0x0F)

**语法**:

```vb
Public Function WriteMultipleCoils(ByVal StartAddress As Long, ByRef Values() As Boolean) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `StartAddress` | Long | 起始线圈地址 |
| `Values()` | Boolean() | 要写入的线圈值数组 |

**返回值**: `Boolean` - 成功返回 True，失败返回 False

**示例**:

```vb
Dim baCoils(4) As Boolean
baCoils(0) = True
baCoils(1) = False
baCoils(2) = True
baCoils(3) = False
baCoils(4) = True

Dim bSuccess As Boolean
bSuccess = mbMaster.WriteMultipleCoils(0, baCoils)
```

---

### WriteSingleRegister - 写入单个寄存器 (0x06)

**语法**:

```vb
Public Function WriteSingleRegister(ByVal Address As Long, ByVal Value As Integer) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 寄存器地址 |
| `Value` | Integer | 要写入的值（16位） |

**返回值**: `Boolean` - 成功返回 True，失败返回 False

**示例**:

```vb
Dim bSuccess As Boolean
bSuccess = mbMaster.WriteSingleRegister(0, 1234)

If bSuccess Then
    Debug.Print "写入成功"
Else
    Debug.Print "写入失败"
End If
```

---

### WriteMultipleRegisters - 写入多个寄存器 (0x10)

**语法**:

```vb
Public Function WriteMultipleRegisters(ByVal StartAddress As Long, ByRef Values() As Integer) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `StartAddress` | Long | 起始寄存器地址 |
| `Values()` | Integer() | 要写入的寄存器值数组 |

**返回值**: `Boolean` - 成功返回 True，失败返回 False

**示例**:

```vb
Dim iRegs(4) As Integer
iRegs(0) = 100
iRegs(1) = 200
iRegs(2) = 300
iRegs(3) = 400
iRegs(4) = 500

Dim bSuccess As Boolean
bSuccess = mbMaster.WriteMultipleRegisters(0, iRegs)
```

---

## ? 事件详解

### OnConnect - 连接成功

**语法**:

```vb
Event OnConnect()
```

**示例**:

```vb
Private Sub mbMaster_OnConnect()
    Debug.Print "已成功连接到 Modbus 从站"
    lblStatus.Caption = "已连接"
    cmdRead.Enabled = True
    cmdWrite.Enabled = True
End Sub
```

---

### OnDisconnect - 连接断开

**语法**:

```vb
Event OnDisconnect()
```

**示例**:

```vb
Private Sub mbMaster_OnDisconnect()
    Debug.Print "连接已断开"
    lblStatus.Caption = "已断开"
    cmdRead.Enabled = False
    cmdWrite.Enabled = False
End Sub
```

---

### OnError - 发生错误

**语法**:

```vb
Event OnError(ByVal Description As String)
```

**示例**:

```vb
Private Sub mbMaster_OnError(ByVal Description As String)
    Debug.Print "错误: " & Description
    
    ' 显示错误提示
    MsgBox "发生错误: " & Description, vbExclamation
    
    ' 记录错误日志
    LogError Description
    
    ' 更新 UI
    lblStatus.Caption = "错误"
End Sub
```

---

### OnDataReceived - 收到数据

**语法**:

```vb
Event OnDataReceived(Data() As Byte)
```

**说明**: 调试事件，用于查看原始收到的数据。

**示例**:

```vb
Private Sub mbMaster_OnDataReceived(Data() As Byte)
    Debug.Print "收到 " & (UBound(Data) + 1) & " 字节数据"
    
    ' 显示十六进制数据
    Dim sHex As String
    Dim i As Long
    For i = 0 To UBound(Data)
        sHex = sHex & Hex$(Data(i)) & " "
    Next i
    Debug.Print "数据: " & sHex
End Sub
```

---

## ? 完整示例

### 基本主站示例

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster

Private Sub Form_Load()
    Set mbMaster = New cModbusMaster
    mbMaster.ProtocolType = MB_PROTOCOL_TCP
    mbMaster.TCPHost = "127.0.0.1"
    mbMaster.TCPPort = 502
    mbMaster.SlaveID = 1
End Sub

Private Sub cmdConnect_Click()
    mbMaster.Connect
End Sub

Private Sub cmdRead_Click()
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(0, 10)
    
    Dim i As Long
    For i = 0 To UBound(iRegs)
        lstRegisters.AddItem "Reg[" & i & "] = " & iRegs(i)
    Next i
End Sub

Private Sub mbMaster_OnConnect()
    Debug.Print "已连接"
End Sub

Private Sub mbMaster_OnDisconnect()
    Debug.Print "已断开"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbMaster.Disconnect
End Sub
```

### 带重连的主站

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents tmrReconnect As Timer
Private m_bAutoReconnect As Boolean

Private Sub Form_Load()
    Set mbMaster = New cModbusMaster
    Set tmrReconnect = New Timer
    
    tmrReconnect.Interval = 5000  ' 5 秒后重连
    m_bAutoReconnect = True
    
    mbMaster.ProtocolType = MB_PROTOCOL_TCP
    mbMaster.TCPHost = "127.0.0.1"
    mbMaster.TCPPort = 502
    mbMaster.SlaveID = 1
    
    ConnectToServer
End Sub

Private Sub ConnectToServer()
    If mbMaster.State = MB_STATE_DISCONNECTED Then
        Debug.Print "正在连接..."
        mbMaster.Connect
    End If
End Sub

Private Sub mbMaster_OnConnect()
    Debug.Print "已连接"
    tmrReconnect.Enabled = False
End Sub

Private Sub mbMaster_OnDisconnect()
    Debug.Print "连接断开"
    
    If m_bAutoReconnect Then
        Debug.Print "5 秒后重连..."
        tmrReconnect.Enabled = True
    End If
End Sub

Private Sub tmrReconnect_Timer()
    tmrReconnect.Enabled = False
    ConnectToServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrReconnect.Enabled = False
    mbMaster.Disconnect
End Sub
```

### 数据采集示例

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster
Private WithEvents tmrPoll As Timer

Private Sub Form_Load()
    Set mbMaster = New cModbusMaster
    Set tmrPoll = New Timer
    
    tmrPoll.Interval = 1000  ' 每秒采集一次
    
    mbMaster.ProtocolType = MB_PROTOCOL_TCP
    mbMaster.TCPHost = "127.0.0.1"
    mbMaster.TCPPort = 502
    mbMaster.SlaveID = 1
    
    mbMaster.Connect
End Sub

Private Sub mbMaster_OnConnect()
    Debug.Print "已连接，开始数据采集"
    tmrPoll.Enabled = True
End Sub

Private Sub tmrPoll_Timer()
    On Error Resume Next
    
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(0, 10)
    
    If UBound(iRegs) >= 0 Then
        ' 更新显示
        UpdateDisplay iRegs
        
        ' 保存到数据库或文件
        SaveToDatabase iRegs
    End If
End Sub

Private Sub UpdateDisplay(ByRef iRegs() As Integer)
    Dim i As Long
    For i = 0 To UBound(iRegs)
        Dim sKey As String
        sKey = "txtReg" & i
        
        On Error Resume Next
        Dim txtBox As Control
        Set txtBox = Me.Controls(sKey)
        If Not txtBox Is Nothing Then
            txtBox.Text = iRegs(i)
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrPoll.Enabled = False
    mbMaster.Disconnect
End Sub
```

---

**最后更新**: 2026-01-16

### 更新日志

#### 2026-01-16 (v1.1.0)
- 更新枚举命名：`ModbusProtocolType` → `ModbusMasterProtocolType`
- 更新状态枚举：`ModbusState` → `ModbusMasterState`
- 更新结构体命名：`ModbusDefaults` → `ModbusMasterDefaults`
- 所有示例代码已同步更新
