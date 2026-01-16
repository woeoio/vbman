# Modbus 类库开发文档

> ? **Modbus 类库** - 基于 cWinsock 封装的 VB6 Modbus 实现库，由 woeoio@qq.com 使用 claude ai 辅助开发

## ? 目录

- [概述](#概述)
- [核心亮点](#核心亮点)
- [架构设计](#架构设计)
- [文档索引](#文档索引)

---

## 概述

Modbus 类库是一个为 VB6 设计的轻量级 Modbus 通信库，完全符合 Modbus 协议规范（RTU 和 TCP）。它基于 cWinsock 类实现，提供了简洁易用的 API 和完整的功能支持。

### ? 主要特性

- ? **纯类实现** - 无需控件，直接使用对象编程
- ? **分离式设计** - 主站（Master）和从站（Slave）独立类库，职责清晰
- ? **完整协议支持** - 支持 Modbus RTU 和 TCP 两种模式
- ? **完整功能码** - 支持所有标准 Modbus 功能码（0x01-0x10, 0x16, 0x17）
- ? **高效缓冲区** - 使用 cByteBuffer 预分配字节缓冲区，减少内存分配操作
- ?? **自动处理** - 自动处理 MBAP 头（TCP）和 CRC 校验（RTU）
- ? **数据存储** - 从站内置数据存储区，支持动态扩展
- ? **事件驱动** - 完整的事件机制，轻松处理连接、数据和错误

---

## 核心亮点

### 1?? 清晰的职责分离 ?

类库采用模块化设计，主站和从站完全独立：

```vb
' cModbusMaster - Modbus 主站（客户端）
Set mbMaster = New cModbusMaster
mbMaster.ProtocolType = MB_PROTOCOL_TCP
mbMaster.TCPHost = "192.168.1.100"
mbMaster.TCPPort = 502
mbMaster.Connect

' cModbusSlave - Modbus 从站（服务器）
Set mbSlave = New cModbusSlave
mbSlave.ProtocolType = MB_PROTOCOL_TCP
mbSlave.SlaveID = 1
mbSlave.Start 502

' cByteBuffer - 高效字节缓冲区（内部使用）
' cWinsock - 底层 Socket 封装
```

---

### 2?? 双协议支持 ?

类库同时支持 Modbus RTU 和 TCP 两种模式：

#### TCP 模式
```vb
mbMaster.ProtocolType = MB_PROTOCOL_TCP
mbMaster.TCPHost = "192.168.1.100"
mbMaster.TCPPort = 502
mbMaster.Connect
```

#### RTU 模式
```vb
mbMaster.ProtocolType = MB_PROTOCOL_RTU
mbMaster.SerialPort = "COM1"
mbMaster.BaudRate = 9600
mbMaster.DataBits = 8
mbMaster.Parity = "N"
mbMaster.StopBits = 1
mbMaster.Connect "COM1"
```

---

### 3?? 完整的功能码支持 ?

支持所有标准 Modbus 功能码：

| 功能码 | 名称 | 说明 |
|--------|------|------|
| 0x01 | Read Coils | 读取线圈 |
| 0x02 | Read Discrete Inputs | 读取离散输入 |
| 0x03 | Read Holding Registers | 读取保持寄存器 |
| 0x04 | Read Input Registers | 读取输入寄存器 |
| 0x05 | Write Single Coil | 写入单个线圈 |
| 0x06 | Write Single Register | 写入单个寄存器 |
| 0x0F | Write Multiple Coils | 写入多个线圈 |
| 0x10 | Write Multiple Registers | 写入多个寄存器 |
| 0x16 | Mask Write Register | 掩码写寄存器 |
| 0x17 | Read/Write Multiple Registers | 读写多个寄存器 |

---

### 4?? 从站数据存储 ?

从站类内置数据存储区，支持动态扩展：

```vb
' 设置线圈
mbSlave.SetCoil 0, True
mbSlave.SetCoil 1, False

' 设置寄存器
mbSlave.SetHoldingRegister 0, 1234
mbSlave.SetHoldingRegister 1, 5678

' 读取数据
Dim bCoil As Boolean
bCoil = mbSlave.GetCoil(0)

Dim iReg As Integer
iReg = mbSlave.GetHoldingRegister(0)
```

---

### 5?? 自动协议处理 ??

#### TCP 模式 - MBAP 头自动处理

```vb
' 主站自动添加 MBAP 头
' 从站自动解析 MBAP 头
' 无需手动处理 Transaction ID、Protocol ID、Length、Unit ID
```

#### RTU 模式 - CRC 校验自动处理

```vb
' 主站自动计算并添加 CRC16
' 从站自动验证 CRC16
' 无需手动处理校验
```

---

### 6?? 事件驱动模型 ?

#### 主站事件
```vb
Event OnConnect()                    ' 连接成功
Event OnDisconnect()                 ' 连接断开
Event OnError(ByVal Description As String)  ' 发生错误
Event OnDataReceived(Data() As Byte)  ' 收到数据
```

#### 从站事件
```vb
Event OnStarted()                              ' 服务器启动
Event OnStopped()                               ' 服务器停止
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)  ' 客户端连接
Event OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)  ' 客户端断开
Event OnReadRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusFunctionCode, ...)  ' 读取请求
Event OnWriteRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusFunctionCode, ...)  ' 写入请求
Event OnError(ByVal Description As String)  ' 发生错误
Event OnDataReceived(ByVal ClientID As String, Data() As Byte)  ' 收到数据
```

---

### 7?? 异常码支持 ??

完整的 Modbus 异常码支持：

| 异常码 | 名称 | 说明 |
|--------|------|------|
| 0x01 | Illegal Function | 非法功能码 |
| 0x02 | Illegal Data Address | 非法数据地址 |
| 0x03 | Illegal Data Value | 非法数据值 |
| 0x04 | Slave Device Failure | 从站设备故障 |
| 0x05 | Acknowledge | 确认 |
| 0x06 | Slave Device Busy | 从站设备忙 |
| 0x08 | Memory Parity Error | 内存奇偶校验错误 |
| 0x0A | Gateway Path Unavailable | 网关路径不可用 |
| 0x0B | Gateway Target Device Failed | 网关目标设备失败 |

---

## 架构设计

### 类层次结构

```
Modbus 类库
├── cModbusMaster (主站/客户端)
│   ├── m_Socket: cWinsock (TCP 连接 Socket)
│   ├── m_hSerialPort: Long (RTU 串口句柄)
│   ├── m_RTUBuffer: cByteBuffer (RTU 接收缓冲区)
│   └── 请求构建/响应解析
│
├── cModbusSlave (从站/服务器)
│   ├── m_ListenSocket: cWinsock (TCP 监听 Socket)
│   ├── m_Clients: Collection (TCP 客户端集合)
│   ├── m_hSerialPort: Long (RTU 串口句柄)
│   ├── m_RTUBuffer: cByteBuffer (RTU 接收缓冲区)
│   ├── m_Coils: Boolean() (线圈数组)
│   ├── m_DiscreteInputs: Boolean() (离散输入数组)
│   ├── m_HoldingRegisters: Integer() (保持寄存器数组)
│   └── m_InputRegisters: Integer() (输入寄存器数组)
│
├── cByteBuffer (字节缓冲区)
│   └── 预分配、自动增长、Peek/Consume/Extract
│
└── cWinsock (底层 Socket 封装)
    └── TCP 连接和数据收发
```

---

### 主站对象关系图

```
cModbusMaster (主站)
├── TCP 模式
│   ├── Socket (连接 Socket: cWinsock)
│   └── TransactionID (事务 ID)
│
└── RTU 模式
    ├── SerialPort (串口句柄)
    ├── RecvBuffer (cByteBuffer)
    └── SerialConfig (波特率、数据位、校验位、停止位)
```

---

### 从站对象关系图

```
cModbusSlave (从站)
├── TCP 模式
│   ├── ListenSocket (监听 Socket: cWinsock)
│   ├── Clients 集合
│   │   ├── 客户端 1 (cWinsock)
│   │   ├── 客户端 2 (cWinsock)
│   │   └── ...
│   └── 数据存储
│       ├── Coils (Boolean 数组)
│       ├── DiscreteInputs (Boolean 数组)
│       ├── HoldingRegisters (Integer 数组)
│       └── InputRegisters (Integer 数组)
│
└── RTU 模式
    ├── SerialPort (串口句柄)
    ├── RecvBuffer (cByteBuffer)
    └── 数据存储
        ├── Coils (Boolean 数组)
        ├── DiscreteInputs (Boolean 数组)
        ├── HoldingRegisters (Integer 数组)
        └── InputRegisters (Integer 数组)
```

---

### 通信流程

#### 主站读取流程

```
1. 调用读取函数（如 ReadHoldingRegisters）
   ↓
2. 构建请求帧
   - RTU: SlaveID + FC + Addr(2) + Quantity(2) + CRC(2)
   - TCP: MBAP(7) + FC + Addr(2) + Quantity(2)
   ↓
3. 发送请求
   - RTU: 通过串口发送
   - TCP: 通过 Socket 发送
   ↓
4. 等待响应（带超时）
   ↓
5. 接收响应数据
   - RTU: 验证 CRC
   - TCP: 验证 Transaction ID
   ↓
6. 解析响应
   - 检查异常码
   - 提取数据
   ↓
7. 返回结果给调用者
```

#### 从站处理流程

```
1. 启动服务器（监听 TCP 或打开串口）
   ↓
2. 等待请求
   ↓
3. 接收到请求
   - TCP: 客户端发送数据
   - RTU: 串口接收数据
   ↓
4. 解析请求帧
   - RTU: 验证 CRC
   - TCP: 解析 MBAP 头
   ↓
5. 检查 Slave ID
   - 是否发给我
   ↓
6. 根据功能码执行操作
   - 读取数据：从内部存储区读取
   - 写入数据：更新内部存储区
   ↓
7. 构建响应帧
   - RTU: PDU + CRC(2)
   - TCP: MBAP(7) + PDU
   ↓
8. 发送响应
   - RTU: 通过串口发送
   - TCP: 通过 Socket 发送
```

---

## 文档索引

| 文档 | 描述 |
|------|------|
| [总览文档](./overview.md) | Modbus 类库的整体介绍和设计理念（当前文档） |
| [主站类详细文档](./master.md) | cModbusMaster 类的详细说明 |
| [从站类详细文档](./slave.md) | cModbusSlave 类的详细说明 |
| [快速开始](./quickstart.md) | 快速入门示例 |
| [进阶应用](./advanced.md) | 高级功能和最佳实践 |

---

## 依赖关系

| 组件 | 描述 |
|------|------|
| **cWinsock.cls** | 位于 `add/` 目录下的底层 Socket 封装，提供 TCP 连接功能 |
| **cByteBuffer.cls** | 位于 `src/` 目录下的字节缓冲区类，用于高效处理字节数据 |
| **cModbusMaster.cls** | 位于 `src/Modbus/` 目录下的 Modbus 主站类 |
| **cModbusSlave.cls** | 位于 `src/Modbus/` 目录下的 Modbus 从站类 |

---

## 兼容性

- **VB6/VBA** - 完全兼容
- **Windows** - Windows XP 及以上版本
- **Modbus 协议** - Modbus RTU 和 TCP（完全兼容）
- **串口** - 标准 COM 端口（RTU 模式）
- **网络** - 标准 TCP/IP（TCP 模式）

---

## 许可证

基于 VbAsyncSocket (wqweto@gmail.com) 开发

---

## 作者

**Modbus 类库**: woeoio@qq.com  
**基础 Socket 库**: woeoio@qq.com  
**原始 Socket 库**: wqweto@gmail.com

---

**最后更新**: 2026-01-16
