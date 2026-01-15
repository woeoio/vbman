# Modbus 通信流程图解

## 一、Modbus 通信模式

```
┌─────────────┐                    ┌─────────────┐
│   主站      │                    │   从站      │
│  (Master)   │                    │  (Slave)    │
│  你的程序   │                    │   PLC       │
└──────┬──────┘                    └──────┬──────┘
       │                                   │
       │         TCP/IP 或 串口            │
       │?─────────────────────────────────┤
       │                                   │
       │                                   │
```

**重要：Modbus 是主从模式，只有主站可以主动发起通信！**

## 二、读取流程（接收 PLC 数据）

### 流程图：

```
你的程序调用读取函数
    │
    ├─> ReadHoldingRegisters(地址, 数量)
    │
    │   构建请求帧
    │   ├─> 从站ID (SlaveID)
    │   ├─> 功能码 (03 = 读保持寄存器)
    │   ├─> 起始地址 (高字节 + 低字节)
    │   └─> 数量 (高字节 + 低字节)
    │
    │   添加协议头
    │   ├─> TCP: 添加 MBAP 头
    │   └─> RTU: 添加 CRC 校验
    │
    │   发送请求
    │   ├─> TCP: 通过 Socket 发送
    │   └─> RTU: 通过串口发送
    │
    │   ┌─────────────────────────────┐
    │   │  等待 PLC 响应（带超时）    │
    │   └─────────────────────────────┘
    │
    │   PLC 收到请求
    │   ├─> 验证从站ID
    │   ├─> 验证功能码
    │   ├─> 从内部寄存器读取数据
    │   └─> 构建响应帧
    │
    │   PLC 发送响应
    │   ├─> 从站ID
    │   ├─> 功能码
    │   ├─> 字节数
    │   └─> 数据（寄存器值）
    │
    │   你的程序接收响应
    │   ├─> 验证响应格式
    │   ├─> 提取数据
    │   └─> 转换为整数数组
    │
    └─> 返回数据数组
```

### 时序图：

```
主站（你的程序）             从站（PLC）
     │                          │
     │── 请求帧 ───────────────>│
     │  [从站ID][功能码][地址][数量]│
     │                          │
     │                          │ 读取内部寄存器
     │                          │
     │<── 响应帧 ────────────────│
     │  [从站ID][功能码][字节数][数据]│
     │                          │
     │  解析数据                │
     │  返回数组                │
```

### 代码示例：

```vb
' 1. 连接
Dim mb As New cModbus
mb.ProtocolType = MB_PROTOCOL_TCP
mb.SlaveID = 1
mb.TCPHost = "192.168.1.100"
mb.TCPPort = 502
mb.Connect

' 2. 读取（内部自动完成：发送请求 → 等待响应 → 解析数据）
Dim regs() As Integer
regs = mb.ReadHoldingRegisters(0, 10)

' 3. 使用数据
Dim i As Long
For i = 0 To UBound(regs)
    Debug.Print "寄存器 " & i & " = " & regs(i)
Next i

' 4. 断开
mb.Disconnect
```

## 三、写入流程（发送数据到 PLC）

### 流程图：

```
你的程序调用写入函数
    │
    ├─> WriteSingleRegister(地址, 值)
    │   或
    └─> WriteMultipleRegisters(起始地址, 值数组)
    │
    │   构建请求帧
    │   ├─> 从站ID (SlaveID)
    │   ├─> 功能码 (06 = 写单个寄存器 或 16 = 写多个寄存器)
    │   ├─> 地址 (高字节 + 低字节)
    │   ├─> 值/数量 (高字节 + 低字节)
    │   └─> 数据（如果是多个寄存器）
    │
    │   添加协议头
    │   ├─> TCP: 添加 MBAP 头
    │   └─> RTU: 添加 CRC 校验
    │
    │   发送请求
    │   ├─> TCP: 通过 Socket 发送
    │   └─> RTU: 通过串口发送
    │
    │   ┌─────────────────────────────┐
    │   │  等待 PLC 响应（带超时）    │
    │   └─────────────────────────────┘
    │
    │   PLC 收到请求
    │   ├─> 验证从站ID
    │   ├─> 验证功能码
    │   ├─> 验证地址范围
    │   ├─> 将数据写入内部寄存器
    │   └─> 构建确认响应帧
    │
    │   PLC 发送确认响应
    │   ├─> 从站ID
    │   ├─> 功能码
    │   ├─> 地址（回显）
    │   └─> 值（回显）
    │
    │   你的程序接收响应
    │   ├─> 验证响应格式
    │   ├─> 比较请求和响应
    │   └─> 确认写入成功
    │
    └─> 返回 True/False
```

### 时序图：

```
主站（你的程序）             从站（PLC）
     │                          │
     │── 写入请求 ──────────────>│
     │  [从站ID][功能码][地址][值]│
     │                          │
     │                          │ 写入内部寄存器
     │                          │
     │<── 确认响应 ──────────────│
     │  [从站ID][功能码][地址][值]│
     │                          │
     │  验证响应                │
     │  返回成功/失败            │
```

### 代码示例：

```vb
' 1. 连接
Dim mb As New cModbus
mb.ProtocolType = MB_PROTOCOL_TCP
mb.SlaveID = 1
mb.TCPHost = "192.168.1.100"
mb.TCPPort = 502
mb.Connect

' 2. 写入单个寄存器（内部自动完成：发送请求 → 等待响应 → 验证）
Dim bSuccess As Boolean
bSuccess = mb.WriteSingleRegister(0, 1234)

' 3. 写入多个寄存器
Dim values(4) As Integer
values(0) = 100
values(1) = 200
values(2) = 300
values(3) = 400
values(4) = 500
bSuccess = mb.WriteMultipleRegisters(10, values)

' 4. 断开
mb.Disconnect
```

## 四、发送当前时间的完整流程

### 流程图：

```
获取当前时间
    │
    ├─> Now() → 2026-01-21 14:30:45
    │
    │   拆分时间
    │   ├─> Year()  → 2026
    │   ├─> Month() → 1
    │   ├─> Day()   → 21
    │   ├─> Hour()  → 14
    │   ├─> Minute()→ 30
    │   └─> Second()→ 45
    │
    │   转换为整数数组
    │   values(0) = 2026
    │   values(1) = 1
    │   values(2) = 21
    │   values(3) = 14
    │   values(4) = 30
    │   values(5) = 45
    │
    │   调用写入函数
    │   └─> WriteMultipleRegisters(100, values)
    │
    │   ┌─────────────────────────────┐
    │   │  内部执行写入流程           │
    │   │  （见上面的写入流程图）     │
    │   └─────────────────────────────┘
    │
    └─> 返回成功/失败
```

### 代码示例：

```vb
' 1. 连接
Dim mb As New cModbus
mb.ProtocolType = MB_PROTOCOL_TCP
mb.SlaveID = 1
mb.TCPHost = "192.168.1.100"
mb.TCPPort = 502
mb.Connect

' 2. 获取当前时间
Dim dtNow As Date
dtNow = Now  ' 例如：2026-01-21 14:30:45

' 3. 拆分时间
Dim values(5) As Integer
values(0) = Year(dtNow)    ' 2026
values(1) = Month(dtNow)   ' 1
values(2) = Day(dtNow)     ' 21
values(3) = Hour(dtNow)    ' 14
values(4) = Minute(dtNow) ' 30
values(5) = Second(dtNow)  ' 45

' 4. 写入到 PLC（从地址 100 开始）
Dim bSuccess As Boolean
bSuccess = mb.WriteMultipleRegisters(100, values)

If bSuccess Then
    Debug.Print "时间写入成功"
Else
    Debug.Print "时间写入失败"
End If

' 5. 断开
mb.Disconnect
```

## 五、数据包结构

### TCP 模式数据包结构：

```
┌─────────────────────────────────────────┐
│         MBAP 头（7 字节）               │
├─────────────────────────────────────────┤
│ 字节0-1: 事务ID (Transaction ID)       │
│ 字节2-3: 协议ID (Protocol ID = 0)      │
│ 字节4-5: 长度 (Length)                  │
│ 字节6:   单元ID (Unit ID = SlaveID)     │
└─────────────────────────────────────────┘
┌─────────────────────────────────────────┐
│         PDU（协议数据单元）             │
├─────────────────────────────────────────┤
│ 字节0:   功能码 (Function Code)         │
│ 字节1-2: 起始地址 (Start Address)       │
│ 字节3-4: 数量 (Quantity)                │
│ 字节5+:  数据 (Data)                    │
└─────────────────────────────────────────┘
```

### RTU 模式数据包结构：

```
┌─────────────────────────────────────────┐
│         ADU（应用数据单元）             │
├─────────────────────────────────────────┤
│ 字节0:   从站ID (Slave ID)              │
│ 字节1:   功能码 (Function Code)         │
│ 字节2-3: 起始地址 (Start Address)       │
│ 字节4-5: 数量 (Quantity)                │
│ 字节6+:  数据 (Data)                    │
│ 最后2字节: CRC16 校验                   │
└─────────────────────────────────────────┘
```

## 六、关键理解点

### 1. Modbus 是请求-响应模式

```
? 错误理解：
   PLC 主动发送数据 → 你的程序接收

? 正确理解：
   你的程序主动请求 → PLC 响应 → 你的程序接收
```

### 2. 读取数据 = 主动请求

```vb
' 当你调用这个函数时：
regs = mb.ReadHoldingRegisters(0, 10)

' 内部发生的事情：
' 1. 构建请求帧
' 2. 发送到 PLC
' 3. 等待 PLC 响应
' 4. 接收响应数据
' 5. 解析数据
' 6. 返回结果
```

### 3. 写入数据 = 主动发送

```vb
' 当你调用这个函数时：
bSuccess = mb.WriteSingleRegister(0, 1234)

' 内部发生的事情：
' 1. 构建写入请求帧
' 2. 发送到 PLC
' 3. PLC 写入数据
' 4. PLC 发送确认响应
' 5. 接收确认响应
' 6. 验证写入成功
' 7. 返回 True/False
```

### 4. 发送时间 = 写入操作

```vb
' 发送时间就是写入操作
' 1. 获取当前时间
' 2. 转换为整数数组
' 3. 调用写入函数
' 4. 完成！
```

## 七、常见误解澄清

### 误解1：PLC 会主动发送数据给我
**澄清：** Modbus 不支持从站主动发送。如果需要实时数据，你需要：
- 使用定时器定期调用读取函数
- 或者使用其他协议（如 OPC、MQTT 等）

### 误解2：读取函数是"被动接收"
**澄清：** 读取函数是"主动请求"，内部会：
1. 主动发送读取请求
2. 等待 PLC 响应
3. 接收并解析响应

### 误解3：写入函数只是"发送数据"
**澄清：** 写入函数会：
1. 发送写入请求
2. 等待 PLC 确认
3. 验证写入是否成功

## 八、实际应用场景

### 场景1：定期读取 PLC 数据

```vb
' 在定时器中调用
Private Sub Timer1_Timer()
    Dim regs() As Integer
    regs = mb.ReadHoldingRegisters(0, 10)
    ' 处理数据...
End Sub
```

### 场景2：定期发送时间到 PLC

```vb
' 在定时器中调用
Private Sub Timer1_Timer()
    Dim dtNow As Date
    Dim values(5) As Integer
    dtNow = Now
    values(0) = Year(dtNow)
    values(1) = Month(dtNow)
    values(2) = Day(dtNow)
    values(3) = Hour(dtNow)
    values(4) = Minute(dtNow)
    values(5) = Second(dtNow)
    mb.WriteMultipleRegisters(100, values)
End Sub
```

### 场景3：读取后写入

```vb
' 读取 PLC 数据
Dim regs() As Integer
regs = mb.ReadHoldingRegisters(0, 10)

' 处理数据...

' 写入结果
mb.WriteMultipleRegisters(20, regs)
```
