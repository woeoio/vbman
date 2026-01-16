# cModbusSlave 类参考

## ? 类概述

`cModbusSlave` 是 Modbus 从站（服务器）实现类，提供监听连接、处理主站请求、维护数据存储等功能。

---

## ? 事件列表

| 事件名 | 触发时机 | 参数 |
|--------|----------|------|
| `OnStarted` | 服务器启动 | 无 |
| `OnStopped` | 服务器停止 | 无 |
| `OnClientConnect` | 客户端连接（TCP 模式） | `ClientID` (客户端 ID), `RemoteAddress` (远程地址) |
| `OnClientDisconnect` | 客户端断开（TCP 模式） | `ClientID` (客户端 ID), `Reason` (断开原因) |
| `OnReadRequest` | 收到读取请求 | `ClientID`, `FunctionCode`, `Address`, `Quantity` |
| `OnWriteRequest` | 收到写入请求 | `ClientID`, `FunctionCode`, `Data` |
| `OnError` | 发生错误 | `Description` (错误描述) |
| `OnDataReceived` | 收到数据（调试用） | `ClientID`, `Data()` (字节数组) |

---

## ? 属性参考

### ProtocolType - 协议类型

**类型**: `ModbusSlaveProtocolType` (枚举)  
**读写**: 读写

**值**:

| 常量 | 值 | 说明 |
|------|-----|------|
| `MB_SLAVE_PROTOCOL_RTU` | 1 | RTU 模式（串口通信） |
| `MB_SLAVE_PROTOCOL_TCP` | 2 | TCP 模式（网络通信） |

**示例**:

```vb
' 设置为 TCP 模式
mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_TCP

' 设置为 RTU 模式
mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_RTU
```

---

### State - 服务器状态

**类型**: `ModbusSlaveState` (枚举)  
**读写**: 只读

**值**:

| 常量 | 值 | 说明 |
|------|-----|------|
| `MB_SLAVE_STATE_STOPPED` | 0 | 已停止 |
| `MB_SLAVE_STATE_STARTING` | 1 | 正在启动 |
| `MB_SLAVE_STATE_RUNNING` | 2 | 运行中 |
| `MB_SLAVE_STATE_ERROR` | 3 | 错误状态 |

**示例**:

```vb
If mbSlave.State = MB_SLAVE_STATE_RUNNING Then
    Debug.Print "服务器运行中"
End If
```

---

### SlaveID - 从站 ID

**类型**: `Byte`  
**读写**: 读写

**说明**: 从站设备地址（1-247）。0 表示广播地址。

**示例**:

```vb
mbSlave.SlaveID = 1
```

---

### Defaults - 默认常量

**类型**: `ModbusSlaveDefaults` (结构体)  
**读写**: 只读

**示例**:

```vb
Debug.Print "默认端口: " & mbSlave.Defaults.TCP_PORT
```

---

### RTU 模式属性

#### SerialPort - 串口名称

**类型**: `String`  
**读写**: 读写

**示例**:

```vb
mbSlave.SerialPort = "COM1"
```

#### BaudRate - 波特率

**类型**: `Long`  
**读写**: 读写

**示例**:

```vb
mbSlave.BaudRate = 9600
```

#### DataBits - 数据位

**类型**: `Long`  
**读写**: 读写

**示例**:

```vb
mbSlave.DataBits = 8
```

#### Parity - 校验位

**类型**: `String`  
**读写**: 读写

**值**: "N" (无), "E" (偶校验), "O" (奇校验)

**示例**:

```vb
mbSlave.Parity = "N"  ' 无校验
```

#### StopBits - 停止位

**类型**: `Long`  
**读写**: 读写

**示例**:

```vb
mbSlave.StopBits = 1
```

---

### TCP 模式属性

#### Port - 监听端口

**类型**: `Long` (只读)  
**读写**: 只读

**说明**: 获取当前监听的端口。

---

#### BindAddress - 监听地址

**类型**: `String`  
**读写**: 读写

**说明**: TCP 监听地址配置（v1.1.0+）。支持以下取值：
- `"0.0.0.0"` 或空字符串 - 监听所有网络接口（默认）
- `"127.0.0.1"` - 仅监听本地回环地址
- `"192.168.1.100"` - 监听指定网络接口

**示例**:

```vb
' 仅监听本地连接（更安全）
mbSlave.BindAddress = "127.0.0.1"
mbSlave.Start 502

' 监听所有网络接口（默认）
mbSlave.BindAddress = "0.0.0.0"
mbSlave.Start 502

' 在 Start 方法中指定监听地址
mbSlave.Start 502, "192.168.1.100"
```

---

### ClientCount - 连接的客户端数

**类型**: `Long` (只读)  
**读写**: 只读

**说明**: TCP 模式下已连接的客户端数量。

**示例**:

```vb
Debug.Print "当前连接数: " & mbSlave.ClientCount
```

---

## ? 方法参考

### Start - 启动服务器

**语法**:

```vb
Public Sub Start(Optional ByVal PortOrSerial As String = "", Optional ByVal BindAddress As String = "")
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `PortOrSerial` | String (可选) | TCP 模式：端口号；RTU 模式：串口名称 |
| `BindAddress` | String (可选) | TCP 模式：监听地址（如 "127.0.0.1"），RTU 模式忽略此参数（v1.1.0+） |

**示例**:

```vb
' TCP 模式 - 启动服务器
mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_TCP
mbSlave.SlaveID = 1
mbSlave.Start 502

' TCP 模式 - 指定监听地址
mbSlave.Start 502, "127.0.0.1"  ' 仅监听本地
mbSlave.Start 502, "0.0.0.0"    ' 监听所有接口

' RTU 模式 - 启动服务器
mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_RTU
mbSlave.SerialPort = "COM1"
mbSlave.BaudRate = 9600
mbSlave.DataBits = 8
mbSlave.Parity = "N"
mbSlave.StopBits = 1
mbSlave.SlaveID = 1
mbSlave.Start "COM1"
```

---

### StopMe - 停止服务器

**语法**:

```vb
Public Sub StopMe()
```

**说明**: 从 v1.1.0 开始，`Stop` 方法更名为 `StopMe`，避免与 VB 关键字冲突。

**示例**:

```vb
mbSlave.StopMe
```

---

### SetCoil - 设置线圈值

**语法**:

```vb
Public Sub SetCoil(ByVal Address As Long, ByVal Value As Boolean)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 线圈地址 (0-1990) |
| `Value` | Boolean | 线圈值 |

**示例**:

```vb
mbSlave.SetCoil 0, True
mbSlave.SetCoil 1, False
```

---

### GetCoil - 获取线圈值

**语法**:

```vb
Public Function GetCoil(ByVal Address As Long) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 线圈地址 |

**返回值**: `Boolean` - 线圈值

**示例**:

```vb
Dim bValue As Boolean
bValue = mbSlave.GetCoil(0)
Debug.Print "Coil[0] = " & bValue
```

---

### SetDiscreteInput - 设置离散输入值

**语法**:

```vb
Public Sub SetDiscreteInput(ByVal Address As Long, ByVal Value As Boolean)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 离散输入地址 (0-2000) |
| `Value` | Boolean | 离散输入值 |

**示例**:

```vb
mbSlave.SetDiscreteInput 0, True
```

---

### GetDiscreteInput - 获取离散输入值

**语法**:

```vb
Public Function GetDiscreteInput(ByVal Address As Long) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 离散输入地址 |

**返回值**: `Boolean` - 离散输入值

**示例**:

```vb
Dim bValue As Boolean
bValue = mbSlave.GetDiscreteInput(0)
```

---

### SetHoldingRegister - 设置保持寄存器值

**语法**:

```vb
Public Sub SetHoldingRegister(ByVal Address As Long, ByVal Value As Integer)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 寄存器地址 (0-125) |
| `Value` | Integer | 寄存器值 (16位) |

**示例**:

```vb
mbSlave.SetHoldingRegister 0, 1234
mbSlave.SetHoldingRegister 1, 5678
```

---

### GetHoldingRegister - 获取保持寄存器值

**语法**:

```vb
Public Function GetHoldingRegister(ByVal Address As Long) As Integer
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 寄存器地址 |

**返回值**: `Integer` - 寄存器值

**示例**:

```vb
Dim iValue As Integer
iValue = mbSlave.GetHoldingRegister(0)
Debug.Print "Reg[0] = " & iValue
```

---

### SetInputRegister - 设置输入寄存器值

**语法**:

```vb
Public Sub SetInputRegister(ByVal Address As Long, ByVal Value As Integer)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 寄存器地址 (0-125) |
| `Value` | Integer | 寄存器值 (16位) |

**示例**:

```vb
mbSlave.SetInputRegister 0, 1234
```

---

### GetInputRegister - 获取输入寄存器值

**语法**:

```vb
Public Function GetInputRegister(ByVal Address As Long) As Integer
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Address` | Long | 寄存器地址 |

**返回值**: `Integer` - 寄存器值

**示例**:

```vb
Dim iValue As Integer
iValue = mbSlave.GetInputRegister(0)
```

---

### ClearAllData - 清空所有数据

**语法**:

```vb
Public Sub ClearAllData()
```

**示例**:

```vb
mbSlave.ClearAllData
```

---

## ? 事件详解

### OnStarted - 服务器启动

**语法**:

```vb
Event OnStarted()
```

**示例**:

```vb
Private Sub mbSlave_OnStarted()
    Debug.Print "服务器已启动"
    lblStatus.Caption = "运行中"
    cmdStart.Enabled = False
    cmdStop.Enabled = True
End Sub
```

---

### OnStopped - 服务器停止

**语法**:

```vb
Event OnStopped()
```

**示例**:

```vb
Private Sub mbSlave_OnStopped()
    Debug.Print "服务器已停止"
    lblStatus.Caption = "已停止"
    cmdStart.Enabled = True
    cmdStop.Enabled = False
End Sub
```

---

### OnClientConnect - 客户端连接

**语法**:

```vb
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `ClientID` | String | 客户端 ID |
| `RemoteAddress` | String | 远程地址 |

**示例**:

```vb
Private Sub mbSlave_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
    Debug.Print "客户端连接: " & ClientID & " (" & RemoteAddress & ")"
    lstClients.AddItem ClientID & " - " & RemoteAddress
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
Private Sub mbSlave_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    Debug.Print "客户端断开: " & ClientID & " - " & Reason
    
    Dim i As Long
    For i = 0 To lstClients.ListCount - 1
        If InStr(lstClients.List(i), ClientID) > 0 Then
            lstClients.RemoveItem i
            Exit For
        End If
    Next i
End Sub
```

---

### OnReadRequest - 读取请求

**语法**:

```vb
Event OnReadRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByVal Quantity As Long)
```

**说明**: 当主站发送读取请求时触发。

**示例**:

```vb
Private Sub mbSlave_OnReadRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByVal Quantity As Long)
    Debug.Print "读取请求: " & FunctionCode & ", 地址=" & Address & ", 数量=" & Quantity
    
    ' 根据功能码处理
    Select Case FunctionCode
        Case MB_SLAVE_FC_READ_COILS
            Debug.Print "读取线圈"
        Case MB_SLAVE_FC_READ_HOLDING_REGISTERS
            Debug.Print "读取保持寄存器"
    End Select
End Sub
```

---

### OnWriteRequest - 写入请求

**语法**:

```vb
Event OnWriteRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByRef Data As Variant)
```

**说明**: 当主站发送写入请求时触发。

**示例**:

```vb
Private Sub mbSlave_OnWriteRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByRef Data As Variant)
    Debug.Print "写入请求: " & FunctionCode & ", 地址=" & Address
    
    ' 根据功能码处理
    Select Case FunctionCode
        Case MB_SLAVE_FC_WRITE_SINGLE_REGISTER
            Debug.Print "写入单个寄存器"
        Case MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
            Debug.Print "写入多个寄存器"
    End Select
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
Private Sub mbSlave_OnError(ByVal Description As String)
    Debug.Print "错误: " & Description
    MsgBox "发生错误: " & Description, vbExclamation
    LogError Description
End Sub
```

---

### OnDataReceived - 收到数据

**语法**:

```vb
Event OnDataReceived(ByVal ClientID As String, Data() As Byte)
```

**说明**: 调试事件，用于查看原始收到的数据。

**示例**:

```vb
Private Sub mbSlave_OnDataReceived(ByVal ClientID As String, Data() As Byte)
    Debug.Print ClientID & " 收到 " & (UBound(Data) + 1) & " 字节"
    
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

### 基本从站示例

```vb
Option Explicit

Private WithEvents mbSlave As cModbusSlave

Private Sub Form_Load()
    Set mbSlave = New cModbusSlave
End Sub

Private Sub cmdStart_Click()
    mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_TCP
    mbSlave.SlaveID = 1
    mbSlave.Start 502
End Sub

Private Sub cmdStop_Click()
    mbSlave.StopMe
End Sub

Private Sub mbSlave_OnStarted()
    Debug.Print "已启动"
    cmdStart.Enabled = False
    cmdStop.Enabled = True
End Sub

Private Sub mbSlave_OnStopped()
    Debug.Print "已停止"
    cmdStart.Enabled = True
    cmdStop.Enabled = False
End Sub
```

---

### 动态数据更新示例

```vb
Option Explicit

Private WithEvents mbSlave As cModbusSlave
Private WithEvents tmrUpdate As Timer

Private Sub Form_Load()
    Set mbSlave = New cModbusSlave
    Set tmrUpdate = New Timer
    
    tmrUpdate.Interval = 1000  ' 每秒更新
    
    mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_TCP
    mbSlave.SlaveID = 1
    mbSlave.Start 502
End Sub

Private Sub mbSlave_OnStarted()
    tmrUpdate.Enabled = True
    Debug.Print "开始动态更新数据"
End Sub

Private Sub tmrUpdate_Timer()
    Dim i As Long
    
    ' 更新时间数据到寄存器 0-5
    Dim dtNow As Date
    dtNow = Now
    
    mbSlave.SetHoldingRegister 0, Year(dtNow)
    mbSlave.SetHoldingRegister 1, Month(dtNow)
    mbSlave.SetHoldingRegister 2, Day(dtNow)
    mbSlave.SetHoldingRegister 3, Hour(dtNow)
    mbSlave.SetHoldingRegister 4, Minute(dtNow)
    mbSlave.SetHoldingRegister 5, Second(dtNow)
    
    ' 更新传感器数据到寄存器 10-20
    For i = 10 To 20
        Dim iValue As Integer
        iValue = ReadSensor(i - 10)  ' 从传感器读取
        mbSlave.SetHoldingRegister i, iValue
    Next i
End Sub

Private Function ReadSensor(iSensorID As Long) As Integer
    ' 模拟传感器数据
    ReadSensor = Rnd * 10000
End Function

Private Sub Form_Unload(Cancel As Integer)
    tmrUpdate.Enabled = False
    mbSlave.StopMe
End Sub
```

---

### 多客户端处理示例

```vb
Option Explicit

Private WithEvents mbSlave As cModbusSlave
Private m_Clients As Collection

Private Sub Form_Load()
    Set mbSlave = New cModbusSlave
    Set m_Clients = New Collection
    
    mbSlave.ProtocolType = MB_SLAVE_PROTOCOL_TCP
    mbSlave.SlaveID = 1
    mbSlave.Start 502
End Sub

Private Sub mbSlave_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
    Debug.Print "客户端连接: " & ClientID & " (" & RemoteAddress & ")"
    
    ' 记录客户端信息
    Dim clientInfo As New Collection
    clientInfo.Add ClientID, "ID"
    clientInfo.Add RemoteAddress, "Address"
    clientInfo.Add Now, "ConnectTime"
    
    m_Clients.Add clientInfo, ClientID
    
    ' 发送欢迎数据
    mbSlave.SetHoldingRegister 100, Year(Now)
    mbSlave.SetHoldingRegister 101, Month(Now)
End Sub

Private Sub mbSlave_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    Debug.Print "客户端断开: " & ClientID
    
    On Error Resume Next
    m_Clients.Remove ClientID
    On Error GoTo 0
End Sub

Private Sub GetClientInfo(ByVal ClientID As String) As String
    On Error Resume Next
    Dim clientInfo As Collection
    Set clientInfo = m_Clients(ClientID)
    
    If Not clientInfo Is Nothing Then
        GetClientInfo = "ID: " & clientInfo("ID") & _
                        ", Address: " & clientInfo("Address") & _
                        ", ConnectTime: " & Format$(clientInfo("ConnectTime"), "hh:mm:ss")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbSlave.StopMe
    Set m_Clients = Nothing
End Sub
```

---

**最后更新**: 2026-01-16

### 更新日志

#### 2026-01-16 (v1.1.0)
- 新增 `BindAddress` 属性 - 支持配置 TCP 监听地址
- 新增 `Start` 方法重载 - 支持在启动时指定监听地址
- 更新方法命名：`Stop()` → `StopMe()`
- 更新枚举命名：添加 `Slave` 后缀区分
  - `ModbusFunctionCode` → `ModbusSlaveFunctionCode`
  - `ModbusExceptionCode` → `ModbusSlaveExceptionCode`
  - `ModbusState` → `ModbusSlaveState`
  - `ModbusProtocolType` → `ModbusSlaveProtocolType`
  - `ModbusDefaults` → `ModbusSlaveDefaults`
- 所有示例代码已同步更新
