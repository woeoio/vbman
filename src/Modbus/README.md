# Modbus 库 - 完整项目概览

完整的 VB6 Modbus TCP/RTU 库实现，包含主站、从站、演示程序。

## 📁 项目结构

```
vbman/src/Modbus/
├── mModbus.bas                     # 公共模块（常数、工具函数、CRC计算）
├── cModbusMaster.cls               # 主站类（支持 TCP/RTU）
├── cModbusSlave.cls                # 从站类（支持 TCP/RTU）
├── cModbusTransportRTU.cls         # RTU 传输层（串口通信）
├── cModbusTransportTCP.cls         # TCP 传输层（网络通信）
├── _bak/                           # 备份目录
│   ├── cModbus.cls
│   ├── cModbusComm.cls
│   └── ...
└── README.md
```

```
vbman/src/Demos/
├── Modbus2/                        # 第一代演示（基础版本）
│   ├── Master/
│   │   ├── FMaster.frm
│   │   └── MasterDemo2.vbp
│   └── Slave/
│       ├── FSlave.frm
│       └── SlaveDemo2.vbp
└── Modbus3/                        # 第二代演示（完整TCP/RTU）
    ├── Master/
    │   ├── FMasterEx.frm           # 主站演示（TCP/RTU 双模式）
    │   └── MasterDemo3.vbp
    ├── Slave/
    │   ├── FSlaveEx.frm            # 从站演示（TCP/RTU 双模式）
    │   └── SlaveDemo3.vbp
    └── README.md
```

## 🎯 核心功能

### cModbusMaster.cls - 主站实现

**连接管理：**
- TCP 模式：IP + Port 连接
- RTU 模式：COM Port + 波特率 + 校验位等

**读操作（Function Code）：**
- FC 01: Read Coils
- FC 02: Read Discrete Inputs
- FC 03: Read Holding Registers
- FC 04: Read Input Registers

**写操作（Function Code）：**
- FC 05: Write Single Coil
- FC 06: Write Single Register
- FC 15: Write Multiple Coils
- FC 16: Write Multiple Registers

**事件：**
- OnConnect: 连接成功
- OnDisconnect: 连接断开
- OnDataReceived: 接收数据
- OnError: 发生错误

### cModbusSlave.cls - 从站实现

**服务器管理：**
- TCP 模式：监听指定端口
- RTU 模式：监听串口

**数据区：**
- Coils (0x00): 线圈（读写）
- Discrete Inputs (0x10000): 离散输入（只读）
- Holding Registers (0x20000): 保存寄存器（读写）
- Input Registers (0x30000): 输入寄存器（只读）

**事件：**
- OnStart: 服务器启动
- OnStop: 服务器停止
- OnCoilsChanged: 线圈被改变
- OnRegistersChanged: 寄存器被改变
- OnError: 发生错误

## 📊 技术参数

### 支持的数据类型

| 数据区 | 地址范围 | 类型 | 数量 |
|--------|----------|------|------|
| Coils | 0-1023 | Boolean | 1024 |
| Discrete Inputs | 0-1023 | Boolean | 1024 |
| Holding Registers | 0-1023 | Integer | 1024 |
| Input Registers | 0-1023 | Integer | 1024 |

### 通信参数

**TCP 模式：**
- 协议：TCP/IP
- 默认端口：502（Modbus 标准端口）
- 超时：可配置（默认 3000ms）

**RTU 模式：**
- 协议：Modbus RTU over Serial
- 波特率：支持 9600、19200、38400、57600
- 数据位：8（固定）
- 停止位：支持 1 或 2
- 校验位：支持 N、E、O
- 超时：可配置（默认 3000ms）

### 限制

- 单次读取最多 100 条数据
- 单次写入最多 100 条数据
- 最大 PDU 大小：252 字节
- 最大请求数据：251 字节

## 🚀 使用示例

### TCP 主站示例

```vb
Dim master As New cModbusMaster

' 配置 TCP 连接
master.ProtocolType = MB_MASTER_PROTOCOL_TCP
master.TCPHost = "127.0.0.1"
master.TCPPort = 502
master.SlaveID = 1
master.ResponseTimeout = 3000

' 连接
master.Connect

' 读取线圈
Dim coils() As Boolean
coils = master.ReadCoils(0, 10)

' 写入单个线圈
master.WriteSingleCoil(0, True)

' 断开连接
master.DisConnect
```

### RTU 主站示例

```vb
Dim master As New cModbusMaster

' 配置 RTU 连接
master.ProtocolType = MB_MASTER_PROTOCOL_RTU
master.SerialPort = "COM1"
master.BaudRate = 9600
master.DataBits = 8
master.StopBits = 1
master.Parity = "N"
master.SlaveID = 1
master.ResponseTimeout = 3000

' 连接
master.Connect

' 读取寄存器
Dim regs() As Integer
regs = master.ReadHoldingRegisters(0, 10)

' 断开连接
master.DisConnect
```

### TCP 从站示例

```vb
Dim slave As New cModbusSlave

' 配置 TCP 服务器
slave.ProtocolType = MB_SLAVE_PROTOCOL_TCP
slave.SlaveID = 1
slave.HoldingRegistersSize = 100

' 初始化数据
slave.SetHoldingRegister 0, 123

' 启动服务器
slave.Start "502"  ' 端口号

' 处理数据变化
Private Sub slave_OnRegistersChanged(ByVal StartAddress As Long, ByVal Quantity As Long, Values() As Integer)
    ' 处理寄存器改变事件
    LogMessage "Registers changed at " & StartAddress
End Sub

' 停止服务器
slave.Stop
```

## 📝 Modbus3 演示程序

### 主站演示 (FMasterEx)

**功能：**
- ✅ TCP/RTU 协议选择
- ✅ 灵活的连接设置
- ✅ 8 种读写操作
- ✅ 实时通信日志
- ✅ 操作结果显示

**使用流程：**
1. 选择协议类型（TCP 或 RTU）
2. 输入连接参数
3. 点击 Connect
4. 执行读写操作
5. 查看结果和日志

### 从站演示 (FSlaveEx)

**功能：**
- ✅ TCP/RTU 协议选择
- ✅ 启动/停止服务器
- ✅ 手动设置数据值
- ✅ 实时数据显示
- ✅ 通信事件日志
- ✅ 自动数据初始化

**初始化数据：**
```
Coils [0-15]:           10101010 10101010
Discrete Inputs [0-15]: 10010010 10010010
Holding Registers [0-99]: 0, 10, 20, 30, ...
Input Registers [0-99]: 1000, 1001, 1002, ...
```

## 🔧 编译和部署

### 编译项目

```bash
# 编译主站库
vb6.exe /make "src/Modbus/mModbus.bas"

# 编译主站演示
vb6.exe /make "src/Demos/Modbus3/Master/MasterDemo3.vbp"

# 编译从站演示
vb6.exe /make "src/Demos/Modbus3/Slave/SlaveDemo3.vbp"
```

### 生成 EXE

```bash
# 生成可执行文件
vb6.exe /make "src/Demos/Modbus3/Master/MasterDemo3.vbp" /OutEXE "MasterDemo3.exe"
vb6.exe /make "src/Demos/Modbus3/Slave/SlaveDemo3.vbp" /OutEXE "SlaveDemo3.exe"
```

## 🧪 测试场景

### 场景1：本地 TCP 通信

```
┌─────────────────┐         TCP          ┌─────────────────┐
│ MasterDemo3     │ ◄──────────────────► │ SlaveDemo3      │
│ (主站)          │   127.0.0.1:502      │ (从站)          │
│ TCP 客户端      │                       │ TCP 服务器      │
└─────────────────┘                       └─────────────────┘
```

**步骤：**
1. 从站：选择 TCP，端口 502，点击 Start Server
2. 主站：选择 TCP，Host 127.0.0.1，Port 502，点击 Connect
3. 主站：执行读写操作

### 场景2：串口 RTU 通信

```
┌─────────────────┐    RS485/RS232    ┌─────────────────┐
│ MasterDemo3     │ ◄───────────────► │ SlaveDemo3      │
│ (主站)          │     COM1/COM2      │ (从站)          │
│ RTU 客户端      │                    │ RTU 服务器      │
└─────────────────┘                    └─────────────────┘
```

**步骤：**
1. 从站：选择 RTU，COM2，9600，点击 Start Server
2. 主站：选择 RTU，COM1，9600，点击 Connect
3. 主站：执行读写操作

## 📚 相关文档

- [README.md](src/Modbus/README.md) - Modbus 库说明
- [Modbus3 演示文档](src/Demos/Modbus3/README.md) - 演示程序详细说明
- [Modbus 协议标准](https://en.wikipedia.org/wiki/Modbus) - 外部参考

## ⚠️ 注意事项

1. **TCP 端口权限**：Linux 下需要 root 权限使用端口 502
2. **串口驱动**：RTU 模式需要正确安装串口驱动
3. **数据类型**：线圈为 Boolean，寄存器为 Integer (-32768 到 32767)
4. **地址范围**：所有地址从 0 开始，最大 1023
5. **超时处理**：网络延迟可能导致超时，需要合理设置超时时间

## 📊 版本历史

- **v1.0.0** (2026-01-16): 完整的 Modbus TCP/RTU 实现
  - 支持 8 种 Modbus 功能码
  - TCP 和 RTU 双通信模式
  - 主站和从站完整实现
  - 丰富的演示程序

## 🤝 许可

该项目为教学和开发用途。

---

**版本**：1.0
**日期**：2026-01-16
**VB6 兼容性**：完全支持
