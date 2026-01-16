# Modbus3 Demo - Complete Master/Slave Implementation

完整的 Modbus Master 和 Slave 演示程序，支持 **TCP** 和 **RTU** 两种通信模式。

## 目录结构

```
Modbus3/
├── Master/
│   ├── FMasterEx.frm          # 主站演示窗体
│   └── MasterDemo3.vbp        # 主站项目文件
└── Slave/
    ├── FSlaveEx.frm           # 从站演示窗体
    └── SlaveDemo3.vbp         # 从站项目文件
```

## 功能特性

### 主站 (FMasterEx)

**协议支持：**
- TCP 模式：通过网络连接到从站（IP + Port）
- RTU 模式：通过串口连接到从站（COM Port）

**读操作：**
- Read Coils (FC 01)
- Read Discrete Inputs (FC 02)
- Read Holding Registers (FC 03)
- Read Input Registers (FC 04)

**写操作：**
- Write Single Coil (FC 05)
- Write Multiple Coils (FC 15)
- Write Single Register (FC 06)
- Write Multiple Registers (FC 16)

### 从站 (FSlaveEx)

**协议支持：**
- TCP 模式：监听指定端口等待主站连接
- RTU 模式：监听串口等待主站连接

**数据区：**
- Coils（线圈）：0-1023（1024 个）
- Discrete Inputs（离散输入）：0-1023（1024 个）
- Holding Registers（保存寄存器）：0-1023（1024 个）
- Input Registers（输入寄存器）：0-1023（1024 个）

**功能：**
- 启动/停止服务器
- 手动设置数据值
- 查看当前数据状态
- 记录通信日志

## 使用方法

### TCP 模式演示

#### 1. 启动从站服务器

1. 打开 `SlaveDemo3.vbp`
2. 确保 Protocol 选择 **TCP**
3. 设置 Port: **502**（或其他可用端口）
4. Slave ID: **1**
5. 点击 **Start Server** 按钮

预期结果：
- Status 显示 "Running"
- 日志显示 "Server started and listening"

#### 2. 启动主站并连接

1. 打开 `MasterDemo3.vbp`
2. 确保 Protocol 选择 **TCP**
3. 设置：
   - Host: **127.0.0.1**（本地测试）或从站 IP
   - Port: **502**（需要与从站一致）
   - Slave ID: **1**
4. 点击 **Connect** 按钮

预期结果：
- Status 显示 "Connected"
- 日志显示 "Connected successfully"

#### 3. 执行读写操作

**读操作示例：**
- Address: 0
- Qty: 10
- 点击 **Read Coils** 获取线圈值

**写操作示例：**
- Address: 0
- Value: 1
- 点击 **Write Single Coil** 写入单个线圈

**观察从站：**
- 从站日志显示接收到的请求
- 从站数据显示相应的改变

### RTU 模式演示

#### 1. 配置串口

需要准备：
- 两个 COM 口（物理或虚拟）
- USB 转 RS485 模块或串口连接线

#### 2. 启动从站服务器 (RTU)

1. 打开 `SlaveDemo3.vbp`
2. Protocol 选择 **RTU**
3. 设置：
   - Port: **COM2**（从站使用）
   - BaudRate: **9600**
   - StopBits: **1**
   - Parity: **N**
   - Slave ID: **1**
4. 点击 **Start Server**

#### 3. 启动主站并连接 (RTU)

1. 打开 `MasterDemo3.vbp`
2. Protocol 选择 **RTU**
3. 设置：
   - Port: **COM1**（主站使用）
   - BaudRate: **9600**（需要与从站一致）
   - StopBits: **1**
   - Parity: **N**
   - Slave ID: **1**
4. 点击 **Connect**

#### 4. 执行读写操作

同 TCP 模式，读写操作完全相同。

## 初始化数据

从站启动时自动初始化的测试数据：

```
Coils [0-15]:           10101010 10101010  (交替模式)
Discrete Inputs [0-15]: 10010010 10010010  (每 3 个中 1 个)
Holding Registers [0-99]: 0, 10, 20, 30, ... (value = address * 10)
Input Registers [0-99]:   1000, 1001, 1002, ... (value = 1000 + address)
```

## 命令行编译

```bash
# 编译主站演示
vb6.exe /make "MasterDemo3.vbp"

# 编译从站演示
vb6.exe /make "SlaveDemo3.vbp"
```

## 故障排除

### TCP 连接失败

- 确保从站服务器已启动
- 检查防火墙是否允许端口访问
- 尝试 localhost (127.0.0.1) 进行本地测试

### RTU 连接失败

- 确保串口驱动已安装
- 检查波特率、数据位、校验位、停止位是否匹配
- 使用虚拟串口工具 (com0com) 进行测试

### 读写操作失败

- 检查地址是否在有效范围内 (0-1023)
- 检查数量是否超过限制 (最多 100)
- 查看日志确认请求是否正确发送

## 相关文件

- `cModbusMaster.cls` - 主站实现
- `cModbusSlave.cls` - 从站实现
- `cModbusTransportRTU.cls` - RTU 传输层
- `cModbusTransportTCP.cls` - TCP 传输层
- `mModbus.bas` - 公共模块（常量、工具函数）

## 技术参数

### 支持的 Modbus 功能码

| 功能码 | 功能 | 类型 |
|--------|------|------|
| 01 | Read Coils | Read |
| 02 | Read Discrete Inputs | Read |
| 03 | Read Holding Registers | Read |
| 04 | Read Input Registers | Read |
| 05 | Write Single Coil | Write |
| 06 | Write Single Register | Write |
| 15 | Write Multiple Coils | Write |
| 16 | Write Multiple Registers | Write |

### 限制

- 单次读取最多 100 个线圈或寄存器
- 单次写入最多 100 个线圈或寄存器
- 最大 PDU 大小: 252 字节
- 响应超时: 可配置（默认 3000ms）

## 注意事项

1. **地址范围**：所有地址从 0 开始，最大 1023
2. **线圈值**：True (1) 或 False (0)
3. **寄存器值**：整数 (Integer) 类型，范围 -32768 到 32767
4. **数据初始化**：从站启动时自动初始化，可手动修改

## 版本信息

- 版本: 1.0
- 日期: 2026-01-16
- VB6 兼容性: 完全支持
