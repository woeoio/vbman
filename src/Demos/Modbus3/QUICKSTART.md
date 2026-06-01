# Modbus3 Demo - 快速开始指南

## 5分钟快速上手

### 方案1：本地 TCP 通信测试（推荐）

#### 第一步：启动从站服务器（SlaveDemo3）

1. 用 VB6 打开 `src/Demos/Modbus3/Slave/SlaveDemo3.vbp`
2. 按 F5 运行程序
3. 确认 **Protocol** 选择 **TCP**
4. 点击 **Start Server** 按钮

✅ 预期结果：
- 日志显示：`[hh:mm:ss] Server started and listening`
- Status 显示：`Status: Running` (绿色)

#### 第二步：启动主站客户端（MasterDemo3）

1. 用 VB6 打开 `src/Demos/Modbus3/Master/MasterDemo3.vbp`
2. 按 F5 运行程序
3. 确认 **Protocol** 选择 **TCP**
4. 确认 **Host** 是 `127.0.0.1`，**Port** 是 `502`
5. 点击 **Connect** 按钮

✅ 预期结果：
- 日志显示：`[hh:mm:ss] Connected successfully`
- Status 显示：`Status: Connected` (绿色)

#### 第三步：执行读操作

主站程序：
1. **Address**: `0`
2. **Qty**: `10`
3. 点击 **Read Coils** 按钮

✅ 预期结果：
- Result 显示：`Coils [0-9]: 1010101010`
- 日志显示：`[hh:mm:ss] Read successful`

#### 第四步：执行写操作

主站程序：
1. **Address**: `0`
2. **Value**: `1`
3. 点击 **Write Single Coil** 按钮

✅ 预期结果：
- 主站日志显示：`[hh:mm:ss] Write successful`
- 从站日志显示接收到写入请求
- 从站的 "Current Data Values" 会更新

### 方案2：使用编译好的 EXE

如果已经有编译好的可执行文件：

1. 运行 `SlaveDemo3.exe`
2. 点击 **Start Server**
3. 运行 `MasterDemo3.exe`
4. 点击 **Connect**
5. 执行读写操作

## 常见操作速查

### TCP 配置（同一台电脑）

**从站（SlaveDemo3）：**
```
Protocol:  TCP
Port:      502
Slave ID:  1
```

**主站（MasterDemo3）：**
```
Protocol:  TCP
Host:      127.0.0.1
Port:      502
Slave ID:  1
```

### RTU 配置（需要虚拟串口）

**从站（SlaveDemo3）：**
```
Protocol:   RTU
Port:       COM2
BaudRate:   9600
StopBits:   1
Parity:     N
Slave ID:   1
```

**主站（MasterDemo3）：**
```
Protocol:   RTU
Port:       COM1
BaudRate:   9600
StopBits:   1
Parity:     N
Slave ID:   1
```

## 8种读写操作说明

### 读操作（4种）

| 按钮 | 功能 | 说明 |
|------|------|------|
| Read Coils | FC 01 | 读取线圈（0/1） |
| Read Discrete Inputs | FC 02 | 读取离散输入（只读） |
| Read Holding Regs | FC 03 | 读取保存寄存器 |
| Read Input Regs | FC 04 | 读取输入寄存器（只读） |

### 写操作（4种）

| 按钮 | 功能 | 说明 |
|------|------|------|
| Write Coil | FC 05 | 写入单个线圈 |
| Write Coils | FC 15 | 写入多个线圈 |
| Write Reg | FC 06 | 写入单个寄存器 |
| Write Regs | FC 16 | 写入多个寄存器 |

## 操作示例

### 示例1：读取并显示值

```
从站当前数据：
Holding Registers [0-9]: 0, 10, 20, 30, 40, 50, 60, 70, 80, 90

主站操作：
  Address: 0
  Qty:     10
  点击：Read Holding Regs

结果显示：
  Holding Registers [0-9]:
  [0]=0, [1]=10, [2]=20, [3]=30, [4]=40,
  [5]=50, [6]=60, [7]=70, [8]=80, [9]=90
```

### 示例2：写入单个值

```
主站操作：
  Address: 5
  Value:   999
  点击：Write Single Reg

结果：
  从站 Holding Register [5] 改为 999
  从站日志显示写入请求
```

### 示例3：写入多个值

```
主站操作：
  Address: 0
  Qty:     5（表示写5个）
  点击：Write Regs

结果：
  从站 Holding Registers [0-4] 被设置为：100, 101, 102, 103, 104
```

## 初始化数据参考

从站启动时的初始数据：

```
=== Current Data Status ===

Coils [0-15]: 10101010 10101010
Discrete Inputs [0-15]: 10010010 10010010
Holding Registers [0-9]: 0, 10, 20, 30, 40, 50, 60, 70, 80, 90
Input Registers [0-9]: 1000, 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009
```

## 日志解读

### 主站日志

```
[14:23:45] Modbus Master Demo (TCP + RTU) started
[14:23:47] Connecting to 127.0.0.1:502...
[14:23:48] Connected successfully
[14:23:50] Reading Coils: Address=0, Qty=10
[14:23:50] Read successful
```

### 从站日志

```
[14:23:45] Modbus Slave Demo (TCP + RTU) started
[14:23:47] Starting TCP server on port 502...
[14:23:47] Server started and listening
[14:23:50] Coils changed: Address=0, Qty=10, Values: 1010101010
[14:23:52] Registers changed: Address=0, Qty=5
```

## 故障排除速查

### 问题：连接失败

**症状：** 主站点击 Connect 后，日志显示 "Connect failed"

**检查清单：**
- [ ] 从站服务器已启动？
- [ ] 从站 Status 是绿色 Running？
- [ ] Protocol 都选了 TCP？
- [ ] Host 和 Port 一致？（都是 127.0.0.1:502）

### 问题：读写失败

**症状：** 读写操作后，日志显示 "Read/Write failed"

**检查清单：**
- [ ] 连接已建立？(Status 显示 Connected)
- [ ] Address 在 0-1023 范围内？
- [ ] Qty 不超过 100？
- [ ] Value 在 -32768 到 32767 范围内？

### 问题：RTU 连接失败

**症状：** RTU 模式无法连接

**原因和解决：**
1. 虚拟串口未建立
   - 安装虚拟串口工具 (com0com)
   - 创建 COM1<->COM2 的虚拟连接

2. 波特率不匹配
   - 确保主从站波特率相同
   - 从站：COM2，主站：COM1

3. 串口占用
   - 关闭其他占用串口的程序
   - 重新启动演示程序

## 更多信息

- **完整文档**：查看 `src/Demos/Modbus3/README.md`
- **库文档**：查看 `src/Modbus/README.md`
- **源码**：查看 `src/Modbus/cModbusMaster.cls` 和 `cModbusSlave.cls`

## 快速技巧

### 💡 同一屏幕看两个程序

使用 VB6 的分窗口运行：
1. 打开从站 VBP，F5 运行
2. 新建一个 VB6 实例，打开主站 VBP，F5 运行
3. 将两个窗口排列在屏幕两侧

### 💡 快速清除日志

点击对应程序窗口底部的 **Clear** 按钮，日志立即清空。

### 💡 手动修改从站数据

在从站程序中：
1. 输入 Address（0-1023）
2. 输入 Value
3. 点击对应的按钮（Set Coil/Reg）
4. 点击 Refresh 查看更新后的值

### 💡 查看原始数据

在从站程序的 "Current Data Values" 部分可以实时看到所有数据的当前值。

---

**祝你使用愉快！如有问题，请查看完整文档。**
