# 快速开始

本指南将帮助您快速上手 Modbus 类库，创建基本的主站和从站应用。

---

## ? 前置准备

### 必需文件

确保以下文件已添加到项目中：

| 文件 | 位置 | 说明 |
|------|------|------|
| `cWinsock.cls` | `add/` | 底层 Socket 封装 |
| `cByteBuffer.cls` | `src/` | 字节缓冲区类 |
| `cModbusMaster.cls` | `src/Modbus/` | 主站类 |
| `cModbusSlave.cls` | `src/Modbus/` | 从站类 |

### 添加到项目

1. 打开 VB6 项目
2. 菜单：项目 → 添加类模块
3. 浏览到相应文件并添加

---

## ? 主站快速入门（TCP 模式）

### 步骤 1：创建窗体

创建一个新窗体，添加以下控件：

- 1 个 TextBox (`txtHost`) - 服务器地址
- 1 个 TextBox (`txtPort`) - 端口号
- 1 个 TextBox (`txtSlaveID`) - 从站 ID
- 2 个 CommandButton (`cmdConnect`, `cmdDisconnect`) - 连接/断开
- 1 个 TextBox (`txtAddress`) - 寄存器地址
- 1 个 TextBox (`txtCount`) - 读取数量
- 1 个 CommandButton (`cmdRead`) - 读取数据
- 1 个 TextBox (`txtLog`) - 显示日志（MultiLine = True）

### 步骤 2：编写代码

```vb
Option Explicit

Private WithEvents mbMaster As cModbusMaster

Private Sub Form_Load()
    Set mbMaster = New cModbusMaster
    
    txtHost.Text = "127.0.0.1"
    txtPort.Text = "502"
    txtSlaveID.Text = "1"
    txtAddress.Text = "0"
    txtCount.Text = "10"
    
    UpdateUI False
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo EH
    
    mbMaster.ProtocolType = MB_PROTOCOL_TCP
    mbMaster.TCPHost = txtHost.Text
    mbMaster.TCPPort = CLng(txtPort.Text)
    mbMaster.SlaveID = CByte(txtSlaveID.Text)
    mbMaster.Connect
    
    LogMessage "正在连接到: " & txtHost.Text & ":" & txtPort.Text
    Exit Sub
    
EH:
    LogMessage "连接失败: " & Err.Description
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next
    
    mbMaster.Disconnect
    LogMessage "已断开连接"
    UpdateUI False
End Sub

Private Sub cmdRead_Click()
    On Error GoTo EH
    
    If mbMaster.State <> MB_STATE_CONNECTED Then
        LogMessage "未连接"
        Exit Sub
    End If
    
    Dim lAddr As Long
    Dim lCount As Long
    Dim iRegs() As Integer
    Dim i As Long
    
    lAddr = CLng(txtAddress.Text)
    lCount = CLng(txtCount.Text)
    
    LogMessage "读取寄存器: 地址=" & lAddr & ", 数量=" & lCount
    
    iRegs = mbMaster.ReadHoldingRegisters(lAddr, lCount)
    
    Dim sResult As String
    For i = 0 To UBound(iRegs)
        sResult = sResult & "Reg[" & (lAddr + i) & "]=" & iRegs(i) & "  "
    Next i
    
    LogMessage sResult
    Exit Sub
    
EH:
    LogMessage "读取失败: " & Err.Description
End Sub

' ====== Modbus Master 事件处理 ======

Private Sub mbMaster_OnConnect()
    LogMessage "*** 连接成功! ***"
    UpdateUI True
End Sub

Private Sub mbMaster_OnDisconnect()
    LogMessage "*** 连接断开 ***"
    UpdateUI False
End Sub

Private Sub mbMaster_OnError(ByVal Description As String)
    LogMessage "*** 错误: " & Description & " ***"
End Sub

' ====== 辅助函数 ======

Private Sub LogMessage(sMessage As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:mm:ss") & " - " & sMessage & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub UpdateUI(bConnected As Boolean)
    txtHost.Enabled = Not bConnected
    txtPort.Enabled = Not bConnected
    txtSlaveID.Enabled = Not bConnected
    cmdConnect.Enabled = Not bConnected
    cmdDisconnect.Enabled = bConnected
    cmdRead.Enabled = bConnected
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mbMaster.Disconnect
End Sub
```

### 步骤 3：运行测试

1. 按 F5 运行程序
2. 输入服务器地址（如 `127.0.0.1`）
3. 点击"连接"
4. 连接成功后，输入地址和数量
5. 点击"读取数据"

---

## ? 从站快速入门（TCP 模式）

### 步骤 1：创建窗体

创建一个新窗体，添加以下控件：

- 1 个 TextBox (`txtPort`) - 监听端口
- 1 个 TextBox (`txtSlaveID`) - 从站 ID
- 2 个 CommandButton (`cmdStart`, `cmdStop`) - 启动/停止
- 1 个 ListBox (`lstRegisters`) - 寄存器列表
- 1 个 TextBox (`txtRegAddr`) - 寄存器地址
- 1 个 TextBox (`txtRegValue`) - 寄存器值
- 1 个 CommandButton (`cmdSetReg`) - 设置寄存器
- 1 个 TextBox (`txtLog`) - 显示日志（MultiLine = True）

### 步骤 2：编写代码

```vb
Option Explicit

Private WithEvents mbSlave As cModbusSlave

Private Sub Form_Load()
    Set mbSlave = New cModbusSlave
    
    txtPort.Text = "502"
    txtSlaveID.Text = "1"
    txtRegAddr.Text = "0"
    txtRegValue.Text = "0"
    
    UpdateServerUI False
End Sub

Private Sub cmdStart_Click()
    On Error GoTo EH
    
    mbSlave.ProtocolType = MB_PROTOCOL_TCP
    mbSlave.SlaveID = CByte(txtSlaveID.Text)
    mbSlave.Start CLng(txtPort.Text)
    
    LogMessage "正在启动服务器, 端口: " & txtPort.Text
    Exit Sub
    
EH:
    LogMessage "启动失败: " & Err.Description
End Sub

Private Sub cmdStop_Click()
    On Error Resume Next
    
    mbSlave.StopMe
    LogMessage "服务器已停止"
    UpdateServerUI False
End Sub

Private Sub cmdSetReg_Click()
    On Error GoTo EH
    
    If mbSlave.State <> MB_SLAVE_STATE_RUNNING Then
        LogMessage "服务器未运行"
        Exit Sub
    End If
    
    Dim lAddr As Long
    Dim iValue As Integer
    
    lAddr = CLng(txtRegAddr.Text)
    iValue = CInt(txtRegValue.Text)
    
    mbSlave.SetHoldingRegister lAddr, iValue
    LogMessage "设置寄存器: 地址=" & lAddr & ", 值=" & iValue
    
    RefreshRegisterList
    Exit Sub
    
EH:
    LogMessage "设置失败: " & Err.Description
End Sub

Private Sub RefreshRegisterList()
    Dim i As Long
    
    lstRegisters.Clear
    For i = 0 To 10
        Dim iValue As Integer
        iValue = mbSlave.GetHoldingRegister(i)
        lstRegisters.AddItem "Reg[" & i & "] = " & iValue
    Next i
End Sub

' ====== Modbus Slave 事件处理 ======

Private Sub mbSlave_OnStarted()
    LogMessage "*** 服务器已启动! ***"
    UpdateServerUI True
End Sub

Private Sub mbSlave_OnStopped()
    LogMessage "*** 服务器已停止 ***"
    UpdateServerUI False
End Sub

Private Sub mbSlave_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
    LogMessage "客户端连接: " & ClientID & " (" & RemoteAddress & ")"
End Sub

Private Sub mbSlave_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    LogMessage "客户端断开: " & ClientID & " - " & Reason
End Sub

Private Sub mbSlave_OnError(ByVal Description As String)
    LogMessage "*** 错误: " & Description & " ***"
End Sub

' ====== 辅助函数 ======

Private Sub LogMessage(sMessage As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:mm:ss") & " - " & sMessage & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub UpdateServerUI(bRunning As Boolean)
    txtPort.Enabled = Not bRunning
    txtSlaveID.Enabled = Not bRunning
    cmdStart.Enabled = Not bRunning
    cmdStop.Enabled = bRunning
    cmdSetReg.Enabled = bRunning
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mbSlave.StopMe
End Sub
```

### 步骤 3：运行测试

1. 按 F5 运行从站程序
2. 点击"启动服务"
3. 运行上面创建的主站程序
4. 点击"连接"
5. 发送读取请求测试

---

## ? 主从通信示例

### 测试场景：主站读取从站寄存器

#### 从站代码（提供数据）

```vb
Private Sub Form_Load()
    Set mbSlave = New cModbusSlave
    
    ' 设置一些初始数据
    mbSlave.SetHoldingRegister 0, 100
    mbSlave.SetHoldingRegister 1, 200
    mbSlave.SetHoldingRegister 2, 300
    mbSlave.SetHoldingRegister 3, 400
    mbSlave.SetHoldingRegister 4, 500
    
    ' 启动服务器
    mbSlave.ProtocolType = MB_PROTOCOL_TCP
    mbSlave.SlaveID = 1
    mbSlave.Start 502
End Sub
```

#### 主站代码（读取数据）

```vb
Private Sub cmdRead_Click()
    ' 连接从站
    mbMaster.ProtocolType = MB_PROTOCOL_TCP
    mbMaster.TCPHost = "127.0.0.1"
    mbMaster.TCPPort = 502
    mbMaster.SlaveID = 1
    mbMaster.Connect
    
    ' 读取5个寄存器
    Dim iRegs() As Integer
    iRegs = mbMaster.ReadHoldingRegisters(0, 5)
    
    ' 显示结果
    Dim i As Long
    For i = 0 To UBound(iRegs)
        Debug.Print "Reg[" & i & "] = " & iRegs(i)
    Next i
    ' 输出:
    ' Reg[0] = 100
    ' Reg[1] = 200
    ' Reg[2] = 300
    ' Reg[3] = 400
    ' Reg[4] = 500
End Sub
```

---

## ? RTU 模式快速开始

### 主站 RTU 模式

```vb
' 配置 RTU 模式
mbMaster.ProtocolType = MB_PROTOCOL_RTU
mbMaster.SerialPort = "COM1"
mbMaster.BaudRate = 9600
mbMaster.DataBits = 8
mbMaster.Parity = "N"
mbMaster.StopBits = 1
mbMaster.SlaveID = 1

' 连接
mbMaster.Connect "COM1"

' 读取寄存器
Dim iRegs() As Integer
iRegs = mbMaster.ReadHoldingRegisters(0, 10)
```

### 从站 RTU 模式

```vb
' 配置 RTU 模式
mbSlave.ProtocolType = MB_PROTOCOL_RTU
mbSlave.SerialPort = "COM1"
mbSlave.BaudRate = 9600
mbSlave.DataBits = 8
mbSlave.Parity = "N"
mbSlave.StopBits = 1
mbSlave.SlaveID = 1

' 启动服务器
mbSlave.Start "COM1"

' 设置数据供主站读取
mbSlave.SetHoldingRegister 0, 1234
mbSlave.SetHoldingRegister 1, 5678
```

---

## ? 完整的功能示例

### 主站 - 读取所有类型的数据

```vb
Private Sub ReadAllTypes()
    ' 读取线圈
    Dim baCoils() As Boolean
    baCoils = mbMaster.ReadCoils(0, 10)
    
    ' 读取离散输入
    Dim baInputs() As Boolean
    baInputs = mbMaster.ReadDiscreteInputs(0, 10)
    
    ' 读取保持寄存器
    Dim iHoldingRegs() As Integer
    iHoldingRegs = mbMaster.ReadHoldingRegisters(0, 10)
    
    ' 读取输入寄存器
    Dim iInputRegs() As Integer
    iInputRegs = mbMaster.ReadInputRegisters(0, 10)
End Sub
```

### 主站 - 写入所有类型的数据

```vb
Private Sub WriteAllTypes()
    ' 写入单个线圈
    mbMaster.WriteSingleCoil 0, True
    
    ' 写入多个线圈
    Dim baCoils(4) As Boolean
    baCoils(0) = True
    baCoils(1) = False
    baCoils(2) = True
    baCoils(3) = False
    baCoils(4) = True
    mbMaster.WriteMultipleCoils 0, baCoils
    
    ' 写入单个寄存器
    mbMaster.WriteSingleRegister 0, 1234
    
    ' 写入多个寄存器
    Dim iRegs(4) As Integer
    iRegs(0) = 100
    iRegs(1) = 200
    iRegs(2) = 300
    iRegs(3) = 400
    iRegs(4) = 500
    mbMaster.WriteMultipleRegisters 0, iRegs
End Sub
```

### 从站 - 动态更新数据

```vb
Private Sub UpdateDataRealtime()
    Dim i As Long
    Dim iValue As Integer
    
    ' 持续更新寄存器
    Do While mbSlave.State = MB_SLAVE_STATE_RUNNING
        For i = 0 To 10
            iValue = GetSensorValue(i)  ' 从传感器获取值
            mbSlave.SetHoldingRegister i, iValue
        Next i
        
        DoEvents  ' 让出 CPU 时间
        Sleep 1000  ' 等待1秒
    Loop
End Sub

Private Function GetSensorValue(iSensorID As Long) As Integer
    ' 模拟传感器数据
    GetSensorValue = Rnd * 10000
End Function
```

---

## ? 常见问题

### Q1: 编译错误"用户定义类型未定义"

**原因**: 未添加 `cByteBuffer.cls` 类。

**解决**: 
1. 菜单：项目 → 添加类模块
2. 浏览到 `src/cByteBuffer.cls`
3. 添加到项目

---

### Q2: 连接失败"无法解析主机名"

**原因**: 
- TCP 模式：URL 格式错误或网络问题
- RTU 模式：串口不存在或被占用

**解决**: 
- TCP: 检查主机地址和端口，确保从站已启动
- RTU: 检查串口名称（COM1, COM2 等），确保串口未被占用

---

### Q3: 读取超时

**原因**: 
- 从站未启动
- 网络连接问题
- 从站 ID 不匹配
- 超时时间设置过短

**解决**: 
- 确保从站已启动
- 检查 Slave ID 是否匹配
- 增加 ResponseTimeout 值（默认 1000ms）

```vb
mbMaster.ResponseTimeout = 3000  ' 3秒超时
```

---

### Q4: 收到异常响应

**原因**: 
- 功能码不支持
- 地址超出范围
- 数据值非法

**解决**: 
- 检查从站支持的功能码
- 确保地址在有效范围内
- 检查数据值是否合法

```vb
Private Sub mbMaster_OnError(ByVal Description As String)
    If InStr(Description, "Modbus Exception") > 0 Then
        MsgBox "Modbus 异常: " & Description
    End If
End Sub
```

---

### Q5: RTU 模式 CRC 校验失败

**原因**: 
- 串口配置不正确（波特率、数据位、校验位、停止位）
- 通信线路干扰

**解决**: 
- 确保主从站串口配置完全一致
- 检查波特率、数据位、校验位、停止位
- 检查通信线路质量

---

### Q6: 如何处理多个从站

```vb
' 连接多个从站
mbMaster.SlaveID = 1
mbMaster.Connect
Dim iRegs1() As Integer
iRegs1 = mbMaster.ReadHoldingRegisters(0, 10)

mbMaster.Disconnect

mbMaster.SlaveID = 2
mbMaster.Connect
Dim iRegs2() As Integer
iRegs2 = mbMaster.ReadHoldingRegisters(0, 10)

mbMaster.Disconnect
```

---

### Q7: 如何实现数据缓存

```vb
' 主站 - 数据缓存
Private m_iRegisterCache(99) As Integer

Private Sub ReadWithCache(lAddr As Long, lCount As Long) As Integer()
    Dim iRegs() As Integer
    
    ' 先从缓存读取
    If lAddr + lCount <= UBound(m_iRegisterCache) + 1 Then
        ReDim iRegs(lCount - 1) As Integer
        Dim i As Long
        For i = 0 To lCount - 1
            iRegs(i) = m_iRegisterCache(lAddr + i)
        Next i
    Else
        ' 缓存未命中，从从站读取
        iRegs = mbMaster.ReadHoldingRegisters(lAddr, lCount)
        
        ' 更新缓存
        If lAddr + lCount <= UBound(m_iRegisterCache) + 1 Then
            For i = 0 To lCount - 1
                m_iRegisterCache(lAddr + i) = iRegs(i)
            Next i
        End If
    End If
    
    ReadWithCache = iRegs
End Sub
```

---

## ? 下一步

- 查看 [master.md](./master.md) 了解主站详细 API
- 查看 [slave.md](./slave.md) 了解从站详细 API
- 查看 [advanced.md](./advanced.md) 了解高级功能

---

**最后更新**: 2026-01-16

### 更新日志

#### 2026-01-16 (v1.1.0)
- 更新从站示例代码，使用 `StopMe()` 替代 `Stop()`
- 添加 `BindAddress` 使用示例（从站）
- 更新枚举引用（适配 v1.1.0 命名规范）
