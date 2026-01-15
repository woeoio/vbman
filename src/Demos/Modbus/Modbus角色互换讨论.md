# Modbus 角色互换讨论

## 一、Modbus 协议的基本设计

### 1. 标准 Modbus 协议

**Modbus 是主从（Master-Slave）协议**，这是协议的基本设计：

- **主站（Master）**：主动发起请求，控制通信
- **从站（Slave）**：被动响应请求，不能主动发起通信

### 2. 当前代码库的实现

从 `cModbus.cls` 的注释可以看到：

```vb
' Purpose: Provides Modbus RTU and TCP client functionality
```

**当前实现是 Modbus Client（客户端/主站）**，只能作为主站使用。

## 二、角色互换的可能性

### ? 理论上完全可行

**PLC 和 PC 的角色是可以互换的！**

### 1. PLC 作为主站

**很多 PLC 都支持 Modbus 主站功能**：

- **西门子 PLC**：支持 Modbus Master 功能块
- **三菱 PLC**：支持 Modbus RTU/TCP 主站通信
- **施耐德 PLC**：支持 Modbus 主站功能
- **欧姆龙 PLC**：支持 Modbus 主站通信

**应用场景**：
- PLC 需要读取其他设备的数据（如传感器、仪表）
- PLC 需要控制其他 Modbus 从站设备
- PLC 作为数据采集中心，主动读取多个从站

### 2. PC 作为从站

**PC 完全可以实现 Modbus 从站功能**：

- 实现 Modbus TCP Server（监听 502 端口）
- 实现 Modbus RTU Slave（通过串口）
- 响应 PLC 的读取/写入请求

**应用场景**：
- PC 作为数据服务器，让 PLC 读取数据
- PC 提供计算服务，PLC 读取计算结果
- PC 作为网关，连接其他系统

## 三、实际应用场景

### 场景1：PLC 作为主站，PC 作为从站

```
┌─────────────┐                    ┌─────────────┐
│    PLC      │                    │     PC      │
│  (Master)   │                    │  (Slave)    │
│  主动请求   │                    │  被动响应   │
└──────┬──────┘                    └──────┬──────┘
       │                                   │
       │── 读取请求 ──────────────────────>│
       │  [从站ID][功能码][地址][数量]      │
       │                                   │
       │                                   │ 读取 PC 内存中的数据
       │                                   │
       │<── 响应数据 ──────────────────────│
       │  [从站ID][功能码][字节数][数据]   │
       │                                   │
```

**实现方式**：
- PC 运行 Modbus 从站程序（需要实现 Server 功能）
- PLC 配置 Modbus 主站功能
- PLC 主动读取 PC 的数据

### 场景2：PC 作为主站，PLC 作为从站（常见场景）

```
┌─────────────┐                    ┌─────────────┐
│     PC      │                    │    PLC       │
│  (Master)   │                    │  (Slave)     │
│  主动请求   │                    │  被动响应   │
└──────┬──────┘                    └──────┬──────┘
       │                                   │
       │── 读取请求 ──────────────────────>│
       │                                   │
       │<── 响应数据 ──────────────────────│
       │                                   │
```

**当前代码库就是这种实现**。

### 场景3：双向通信（需要两个连接）

```
┌─────────────┐                    ┌─────────────┐
│     PC      │                    │    PLC       │
└──────┬──────┘                    └──────┬──────┘
       │                                   │
       │  连接1: PC(Master) → PLC(Slave)  │
       │  连接2: PLC(Master) → PC(Slave)  │
       │                                   │
```

**注意**：需要两个独立的连接，因为 Modbus 是主从协议，不能在同一连接上双向主动通信。

## 四、如何实现角色互换

### 1. 让 PC 作为从站（需要实现 Modbus Server）

**当前代码库缺少的功能**：

需要实现 `cModbusServer` 类，功能包括：

```vb
' Modbus Server (从站) 需要实现的功能
Public Class cModbusServer
    ' 启动服务器
    Public Sub Start(Optional Port As Long = 502)
    
    ' 停止服务器
    Public Sub Stop
    
    ' 设置从站ID
    Public Property SlaveID As Byte
    
    ' 注册数据回调（当 PLC 读取数据时）
    Event OnReadHoldingRegisters(ByVal Address As Long, ByVal Quantity As Long, ByRef Values() As Integer)
    
    ' 注册写入回调（当 PLC 写入数据时）
    Event OnWriteHoldingRegisters(ByVal Address As Long, ByRef Values() As Integer)
    
    ' 设置内部寄存器值（供 PLC 读取）
    Public Sub SetHoldingRegister(ByVal Address As Long, ByVal Value As Integer)
    
    ' 获取内部寄存器值
    Public Function GetHoldingRegister(ByVal Address As Long) As Integer
End Class
```

**实现要点**：
- TCP 模式：监听 502 端口，接受连接
- RTU 模式：通过串口接收数据
- 解析请求帧
- 根据功能码执行相应操作
- 构建响应帧并发送

### 2. 让 PLC 作为主站

**PLC 端配置**（以西门子为例）：

```pascal
// 西门子 PLC Modbus Master 示例
VAR
    MB_MASTER_DB : TMB_MASTER_DB;
    Error : BOOL;
    Status : WORD;
END_VAR

// 调用 Modbus Master 功能块
MB_MASTER_DB(
    REQ := TRUE,
    DONE => ,
    ERROR => Error,
    STATUS => Status,
    MB_ADDR := 1,        // 从站地址（PC 的从站ID）
    MB_MODE := 3,        // 功能码 03（读保持寄存器）
    MB_DATA_ADDR := 0,   // 起始地址
    MB_DATA_LEN := 10,   // 数量
    MB_DATA_PTR => ...   // 数据指针
);
```

## 五、技术限制和注意事项

### 1. Modbus 协议本身的限制

**标准 Modbus 不支持从站主动发送**：
- 从站只能响应主站的请求
- 不能主动推送数据
- 不能主动发起通信

**如果需要从站主动发送，需要**：
- 使用 Modbus Plus（需要特殊硬件）
- 使用扩展协议（非标准）
- 使用其他协议（如 OPC、MQTT 等）

### 2. 双向通信的实现方式

**方式1：两个独立连接**
```
连接1: PC(Master) → PLC(Slave)  // PC 读取 PLC
连接2: PLC(Master) → PC(Slave)   // PLC 读取 PC
```

**方式2：轮询方式**
```
PC 定期读取 PLC（PC 作为主站）
PLC 定期读取 PC（PLC 作为主站，需要 PC 实现从站）
```

**方式3：使用其他协议**
```
Modbus 用于 PC → PLC
MQTT/OPC 用于双向通信
```

### 3. 实际应用中的考虑

**性能考虑**：
- Modbus 是同步协议，每次请求都需要等待响应
- 双向通信会增加延迟
- 需要考虑超时和错误处理

**可靠性考虑**：
- 网络中断时的处理
- 数据一致性问题
- 冲突避免（如果两个主站同时操作）

## 六、代码实现示例

### 示例1：PC 作为从站（伪代码）

```vb
' PC 端实现 Modbus Server
Public Class cModbusServer
    Private m_Socket As cTcpServer
    Private m_Registers(65535) As Integer
    Private m_SlaveID As Byte
    
    Public Sub Start()
        Set m_Socket = New cTcpServer
        m_Socket.Listen 502  ' 监听 502 端口
    End Sub
    
    Private Sub m_Socket_OnClientConnect(Client As cTcpClient)
        ' 新客户端连接
    End Sub
    
    Private Sub m_Socket_OnDataReceived(Client As cTcpClient, Data() As Byte)
        ' 解析 Modbus 请求
        Dim bSlaveID As Byte
        Dim bFunctionCode As Byte
        Dim lAddress As Long
        Dim lQuantity As Long
        
        ' 解析请求帧
        bSlaveID = Data(6)  ' TCP 模式的 Unit ID
        bFunctionCode = Data(7)
        
        If bSlaveID <> m_SlaveID Then Exit Sub  ' 不是发给我的
        
        Select Case bFunctionCode
            Case MB_FC_READ_HOLDING_REGISTERS
                ' 读取保持寄存器
                lAddress = Data(8) * 256 + Data(9)
                lQuantity = Data(10) * 256 + Data(11)
                
                ' 构建响应
                Dim baResponse() As Byte
                baResponse = BuildReadResponse(lAddress, lQuantity)
                Client.SendData baResponse
                
            Case MB_FC_WRITE_SINGLE_REGISTER
                ' 写入单个寄存器
                lAddress = Data(8) * 256 + Data(9)
                Dim iValue As Integer
                iValue = Data(10) * 256 + Data(11)
                m_Registers(lAddress) = iValue
                
                ' 构建响应（回显请求）
                Dim baResponse() As Byte
                baResponse = BuildWriteResponse(lAddress, iValue)
                Client.SendData baResponse
        End Select
    End Sub
    
    Public Sub SetRegister(ByVal Address As Long, ByVal Value As Integer)
        m_Registers(Address) = Value
    End Sub
    
    Public Function GetRegister(ByVal Address As Long) As Integer
        GetRegister = m_Registers(Address)
    End Function
End Class
```

### 示例2：使用场景

```vb
' PC 作为从站，让 PLC 读取当前时间
Public Sub StartTimeServer()
    Dim mbServer As New cModbusServer
    mbServer.SlaveID = 1
    mbServer.Start
    
    ' 定时更新寄存器值（当前时间）
    Dim tmr As Timer
    Set tmr = New Timer
    tmr.Interval = 1000  ' 每秒更新一次
    
    ' 在定时器中更新寄存器
    ' mbServer.SetRegister(0, Year(Now))
    ' mbServer.SetRegister(1, Month(Now))
    ' ...
End Sub
```

## 七、总结

### 关键点

1. **PLC 不一定只能作为从站**
   - 很多 PLC 支持 Modbus 主站功能
   - PLC 可以作为主站读取其他设备

2. **PC 可以作为从站**
   - 需要实现 Modbus Server 功能
   - 当前代码库只有 Client 功能，需要扩展

3. **角色可以互换**
   - PC 作为主站，PLC 作为从站（当前实现）
   - PLC 作为主站，PC 作为从站（需要实现 Server）
   - 双向通信需要两个独立连接

4. **Modbus 协议限制**
   - 从站不能主动发送数据
   - 需要主站主动请求
   - 双向通信需要两个连接或轮询

### 实际建议

- **如果只需要 PC 读取 PLC**：使用当前实现（PC 主站，PLC 从站）
- **如果需要 PLC 读取 PC**：需要实现 Modbus Server 功能
- **如果需要双向实时通信**：考虑使用 MQTT、OPC 等协议
- **如果需要从站主动推送**：考虑使用其他协议或扩展 Modbus

### 下一步

如果需要实现 PC 作为从站的功能，可以：
1. 创建 `cModbusServer` 类
2. 实现 TCP Server 监听
3. 实现请求解析和响应构建
4. 提供数据存储和回调接口
