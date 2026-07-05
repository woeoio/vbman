Attribute VB_Name = "modSerialPortTest"
Option Explicit

'===============================================================================
' 模块: modSerialPortTest
' 说明: cSerialPort 串口通信类测试用例
'       包含配置测试、端口操作测试、读写测试、信号线测试、事件测试等
'       部分测试需要硬件串口或虚拟串口对才能执行
'===============================================================================

Private m_TestCount As Long
Private m_PassCount As Long
Private m_FailCount As Long

'===============================================================================
' 测试框架 - 简易断言
'===============================================================================

Private Sub AssertEqual(ByVal Actual As Variant, ByVal Expected As Variant, ByVal TestName As String)
    m_TestCount = m_TestCount + 1
    If Actual = Expected Then
        m_PassCount = m_PassCount + 1
        Debug.Print "  [PASS] " & TestName
    Else
        m_FailCount = m_FailCount + 1
        Debug.Print "  [FAIL] " & TestName & " (期望=" & CStr(Expected) & ", 实际=" & CStr(Actual) & ")"
    End If
End Sub

Private Sub AssertTrue(ByVal Condition As Boolean, ByVal TestName As String)
    m_TestCount = m_TestCount + 1
    If Condition Then
        m_PassCount = m_PassCount + 1
        Debug.Print "  [PASS] " & TestName
    Else
        m_FailCount = m_FailCount + 1
        Debug.Print "  [FAIL] " & TestName
    End If
End Sub

Private Sub AssertFalse(ByVal Condition As Boolean, ByVal TestName As String)
    AssertTrue Not Condition, TestName
End Sub

Private Sub BeginTestGroup(ByVal GroupName As String)
    Debug.Print vbCrLf & "--- " & GroupName & " ---"
End Sub

'===============================================================================
' 运行所有测试（入口）
'===============================================================================

Public Sub RunAllTests()
    m_TestCount = 0
    m_PassCount = 0
    m_FailCount = 0

    Debug.Print "========================================"
    Debug.Print " cSerialPort 测试套件"
    Debug.Print " " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "========================================"

    ' 不依赖硬件的测试
    TestConfigDefaults
    TestConfigBuildDCB
    TestConfigApplyDCB
    TestConfigBuildTimeouts
    TestConfigApplyTimeouts
    TestConfigClone
    TestConfigModeString
    TestConfigPresets
    TestConfigValidation
    TestPortCreateDestroy
    TestPortOpenInvalid
    TestBitOperations
    TestErrorStrings
    TestEventStrings
    TestEnumPorts

    ' 需要串口硬件的测试
    TestPortOpenClose
    TestPortConfigApply
    TestPortReadWrite
    TestSignalLines
    TestBreakSignal
    TestBufferPurge
    TestMonitoring
    TestPortStatus
    TestLoopback

    ' 输出汇总
    Debug.Print vbCrLf & "========================================"
    Debug.Print " 测试完成: 共 " & m_TestCount & " 项"
    Debug.Print " 通过: " & m_PassCount & "  失败: " & m_FailCount
    Debug.Print "========================================"
End Sub

'===============================================================================
' 测试组1: cSerialConfig 默认值
'===============================================================================

Private Sub TestConfigDefaults()
    BeginTestGroup "cSerialConfig 默认值"

    Dim cfg As New cSerialConfig
    AssertEqual cfg.BaudRate, br9600, "默认波特率=9600"
    AssertEqual cfg.DataBits, 8, "默认数据位=8"
    AssertEqual cfg.Parity, ptNone, "默认校验=None"
    AssertEqual cfg.StopBits, sb1, "默认停止位=1"
    AssertEqual cfg.FlowControl, fcNone, "默认流控制=None"
    AssertEqual cfg.DtrControl, dcEnable, "默认DTR=Enable"
    AssertEqual cfg.RtsControl, rcEnable, "默认RTS=Enable"
    AssertEqual cfg.InBufferSize, 4096, "默认输入缓冲区=4096"
    AssertEqual cfg.OutBufferSize, 4096, "默认输出缓冲区=4096"
    AssertEqual cfg.XonChar, &H11, "默认XON=0x11"
    AssertEqual cfg.XoffChar, &H13, "默认XOFF=0x13"
    AssertFalse cfg.NullDiscard, "默认不丢弃NULL"
    AssertFalse cfg.AbortOnError, "默认错误时不终止"
    AssertFalse cfg.TXContinueOnXoff, "默认XOFF时不继续发送"
    AssertTrue cfg.ParityReplace, "默认启用校验替换"
End Sub

'===============================================================================
' 测试组2: cSerialConfig BuildDCB
'===============================================================================

Private Sub TestConfigBuildDCB()
    BeginTestGroup "cSerialConfig BuildDCB"

    Dim cfg As New cSerialConfig
    Dim d As DCB
    d = cfg.BuildDCB

    AssertEqual d.DCBlength, LenB(d), "DCBlength=LenB(DCB)"
    AssertEqual d.BaudRate, 9600, "DCB BaudRate=9600"
    AssertEqual d.ByteSize, 8, "DCB ByteSize=8"
    AssertEqual d.Parity, 0, "DCB Parity=None"
    AssertEqual d.StopBits, 0, "DCB StopBits=1"

    ' fBitFields必须包含fBinary位
    AssertTrue GetBit(d.fBitFields, DCB_fBinary), "DCB fBinary=1"

    ' 测试修改配置后构建DCB
    cfg.BaudRate = br115200
    cfg.Parity = ptEven
    cfg.DataBits = 7
    cfg.StopBits = sb2
    cfg.FlowControl = fcRtsCts

    d = cfg.BuildDCB
    AssertEqual d.BaudRate, 115200, "DCB 修改后BaudRate=115200"
    AssertEqual d.ByteSize, 7, "DCB 修改后ByteSize=7"
    AssertEqual d.Parity, 2, "DCB 修改后Parity=Even"
    AssertEqual d.StopBits, 2, "DCB 修改后StopBits=2"
    AssertTrue GetBit(d.fBitFields, DCB_fOutxCtsFlow), "DCB RTS/CTS流控制=启用"
End Sub

'===============================================================================
' 测试组3: cSerialConfig ApplyDCB（DCB → Config 反序列化）
'===============================================================================

Private Sub TestConfigApplyDCB()
    BeginTestGroup "cSerialConfig ApplyDCB"

    Dim cfg As New cSerialConfig
    Dim d As DCB

    ' 手动构建DCB
    d.DCBlength = LenB(d)
    d.BaudRate = 57600
    d.ByteSize = 7
    d.Parity = 1    ' Odd
    d.StopBits = 2  ' 2
    d.XonChar = &H11
    d.XoffChar = &H13
    d.ErrorChar = 0
    d.EofChar = 0
    d.EvtChar = &H42

    ' 设置fBitFields
    Dim f As Long
    SetBit f, DCB_fBinary, True
    SetBit f, DCB_fParity, True
    SetBits f, DCB_fDtrControl, 2, 2  ' Handshake
    SetBits f, DCB_fRtsControl, 2, 2  ' Handshake
    SetBit f, DCB_fOutxCtsFlow, True
    d.fBitFields = f

    cfg.ApplyDCB d
    AssertEqual cfg.BaudRate, 57600, "ApplyDCB BaudRate=57600"
    AssertEqual cfg.DataBits, 7, "ApplyDCB DataBits=7"
    AssertEqual cfg.Parity, ptOdd, "ApplyDCB Parity=Odd"
    AssertEqual cfg.StopBits, sb2, "ApplyDCB StopBits=2"
    AssertEqual cfg.DtrControl, dcHandshake, "ApplyDCB DTR=Handshake"
    AssertEqual cfg.RtsControl, rcHandshake, "ApplyDCB RTS=Handshake"
    AssertEqual cfg.EvtChar, &H42, "ApplyDCB EvtChar=0x42"
    AssertEqual cfg.FlowControl, fcRtsCts, "ApplyDCB FlowControl=RtsCts"
End Sub

'===============================================================================
' 测试组4: cSerialConfig BuildTimeouts
'===============================================================================

Private Sub TestConfigBuildTimeouts()
    BeginTestGroup "cSerialConfig BuildTimeouts"

    Dim cfg As New cSerialConfig
    Dim ct As COMMTIMEOUTS
    ct = cfg.BuildTimeouts

    AssertEqual ct.ReadIntervalTimeout, 50, "默认ReadInterval=50"
    AssertEqual ct.ReadTotalTimeoutConstant, 100, "默认ReadTotalConstant=100"
    AssertEqual ct.WriteTotalTimeoutConstant, 100, "默认WriteTotalConstant=100"

    ' 修改后验证
    cfg.ReadIntervalTimeout = 0
    cfg.ReadTotalTimeoutConstant = 500
    ct = cfg.BuildTimeouts
    AssertEqual ct.ReadIntervalTimeout, 0, "修改后ReadInterval=0"
    AssertEqual ct.ReadTotalTimeoutConstant, 500, "修改后ReadTotalConstant=500"
End Sub

'===============================================================================
' 测试组5: cSerialConfig ApplyTimeouts
'===============================================================================

Private Sub TestConfigApplyTimeouts()
    BeginTestGroup "cSerialConfig ApplyTimeouts"

    Dim cfg As New cSerialConfig
    Dim ct As COMMTIMEOUTS

    ' 手动设置超时
    ct.ReadIntervalTimeout = 100
    ct.ReadTotalTimeoutMultiplier = 10
    ct.ReadTotalTimeoutConstant = 1000
    ct.WriteTotalTimeoutMultiplier = 5
    ct.WriteTotalTimeoutConstant = 2000

    cfg.ApplyTimeouts ct
    AssertEqual cfg.ReadIntervalTimeout, 100, "ApplyTimeouts ReadInterval=100"
    AssertEqual cfg.ReadTotalTimeoutMultiplier, 10, "ApplyTimeouts ReadMultiplier=10"
    AssertEqual cfg.ReadTotalTimeoutConstant, 1000, "ApplyTimeouts ReadConstant=1000"
    AssertEqual cfg.WriteTotalTimeoutMultiplier, 5, "ApplyTimeouts WriteMultiplier=5"
    AssertEqual cfg.WriteTotalTimeoutConstant, 2000, "ApplyTimeouts WriteConstant=2000"
End Sub

'===============================================================================
' 测试组6: cSerialConfig Clone
'===============================================================================

Private Sub TestConfigClone()
    BeginTestGroup "cSerialConfig Clone"

    Dim cfg As New cSerialConfig
    cfg.BaudRate = br115200
    cfg.DataBits = 7
    cfg.Parity = ptEven
    cfg.FlowControl = fcXonXoff

    Dim cfg2 As cSerialConfig
    Set cfg2 = cfg.Clone()

    AssertEqual cfg2.BaudRate, br115200, "Clone BaudRate=115200"
    AssertEqual cfg2.DataBits, 7, "Clone DataBits=7"
    AssertEqual cfg2.Parity, ptEven, "Clone Parity=Even"
    AssertEqual cfg2.FlowControl, fcXonXoff, "Clone FlowControl=XonXoff"

    ' 修改克隆不影响原件
    cfg2.BaudRate = br9600
    AssertEqual cfg.BaudRate, br115200, "Clone独立性: 修改克隆不影响原件"
End Sub

'===============================================================================
' 测试组7: cSerialConfig 模式字符串
'===============================================================================

Private Sub TestConfigModeString()
    BeginTestGroup "cSerialConfig 模式字符串"

    Dim cfg As New cSerialConfig
    cfg.BaudRate = br115200
    cfg.Parity = ptNone
    cfg.DataBits = 8
    cfg.StopBits = sb1

    Dim modeStr As String
    modeStr = cfg.ToModeString()
    AssertTrue InStr(modeStr, "115200") > 0, "ToModeString包含115200"
    AssertTrue InStr(modeStr, "parity=N") > 0, "ToModeString包含parity=N"
    AssertTrue InStr(modeStr, "data=8") > 0, "ToModeString包含data=8"
    AssertTrue InStr(modeStr, "stop=1") > 0, "ToModeString包含stop=1"
End Sub

'===============================================================================
' 测试组8: cSerialConfig 预设模式
'===============================================================================

Private Sub TestConfigPresets()
    BeginTestGroup "cSerialConfig 预设模式"

    Dim cfg As New cSerialConfig

    ' 非阻塞读取
    cfg.SetNonBlockingRead
    AssertEqual cfg.ReadIntervalTimeout, &HFFFFFFFF, "非阻塞: ReadInterval=MAXDWORD"

    ' 阻塞读取
    cfg.SetBlockingRead
    AssertEqual cfg.ReadIntervalTimeout, 0, "阻塞: ReadInterval=0"
    AssertEqual cfg.ReadTotalTimeoutConstant, 0, "阻塞: ReadConstant=0"

    ' 带超时读取
    cfg.SetReadTimeout 2000
    AssertEqual cfg.ReadTotalTimeoutConstant, 2000, "超时读取: ReadConstant=2000"

    ' 写入超时
    cfg.SetWriteTimeout 3000
    AssertEqual cfg.WriteTotalTimeoutConstant, 3000, "写入超时: WriteConstant=3000"
End Sub

'===============================================================================
' 测试组9: cSerialConfig 数据验证
'===============================================================================

Private Sub TestConfigValidation()
    BeginTestGroup "cSerialConfig 数据验证"

    Dim cfg As New cSerialConfig

    ' DataBits 范围检查
    On Error Resume Next
    ERR.Clear
    cfg.DataBits = 3
    AssertTrue ERR.Number <> 0, "DataBits=3应报错"
    ERR.Clear
    cfg.DataBits = 9
    AssertTrue ERR.Number <> 0, "DataBits=9应报错"
    On Error GoTo 0

    ' 合法值
    cfg.DataBits = 5
    AssertEqual cfg.DataBits, 5, "DataBits=5合法"

    ' InBufferSize 检查
    On Error Resume Next
    ERR.Clear
    cfg.InBufferSize = 0
    AssertTrue ERR.Number <> 0, "InBufferSize=0应报错"
    On Error GoTo 0

    cfg.InBufferSize = 8192
    AssertEqual cfg.InBufferSize, 8192, "InBufferSize=8192合法"
End Sub

'===============================================================================
' 测试组10: cSerialPort 创建和销毁
'===============================================================================

Private Sub TestPortCreateDestroy()
    BeginTestGroup "cSerialPort 创建和销毁"
    
    Dim sp As New cSerialPort
    AssertFalse sp.IsOpen, "新建对象: 未打开"
    AssertEqual sp.portName, "COM1", "新建对象: 端口名=COM1"
    AssertFalse sp.Monitoring, "新建对象: 未监控"
    AssertEqual sp.ReceivedCount, 0, "新建对象: 接收计数=0"
    AssertEqual sp.SentCount, 0, "新建对象: 发送计数=0"
    
    ' 修改端口名（未打开时可修改）
    sp.portName = "COM3"
    AssertEqual sp.portName, "COM3", "修改端口名=COM3"
End Sub

'===============================================================================
' 测试组11: cSerialPort 打开不存在的端口
'===============================================================================

Private Sub TestPortOpenInvalid()
    BeginTestGroup "cSerialPort 打开无效端口"

    Dim sp As New cSerialPort
    Dim result As Boolean

    ' 尝试打开一个几乎不可能存在的端口
    result = sp.OpenPort("COM999")
    AssertFalse result, "打开COM999应失败"
    AssertFalse sp.IsOpen, "打开失败后IsOpen=False"
    AssertTrue sp.LastError <> 0, "打开失败后LastError非零"
End Sub

'===============================================================================
' 测试组12: 位操作辅助函数
'===============================================================================

Private Sub TestBitOperations()
    BeginTestGroup "位操作辅助函数"

    ' GetBit / SetBit
    Dim v As Long
    v = 0
    SetBit v, 0, True
    AssertTrue GetBit(v, 0), "SetBit Bit0=True"
    AssertFalse GetBit(v, 1), "Bit1仍为False"

    SetBit v, 1, True
    AssertTrue GetBit(v, 1), "SetBit Bit1=True"
    AssertEqual v, 3, "Bit0+Bit1 = 3"

    SetBit v, 0, False
    AssertFalse GetBit(v, 0), "SetBit Bit0=False"
    AssertEqual v, 2, "清Bit0后=2"

    ' GetBits / SetBits (2位字段)
    v = 0
    SetBits v, 4, 2, 2  ' fDtrControl=Handshake(2)
    AssertEqual GetBits(v, 4, 2), 2, "GetBits(4,2)=2"
    AssertTrue GetBit(v, 4), "Bit4=True (Handshake低位)"
    AssertFalse GetBit(v, 5), "Bit5=False"

    SetBits v, 4, 2, 3
    AssertEqual GetBits(v, 4, 2), 3, "GetBits(4,2)=3"
    AssertTrue GetBit(v, 4), "Bit4=True"
    AssertTrue GetBit(v, 5), "Bit5=True"

    ' SetBits不破坏其他位
    v = &HFF00
    SetBits v, 4, 2, 0
    AssertEqual GetBits(v, 4, 2), 0, "SetBits清零DTR字段"
    AssertTrue GetBit(v, 8), "Bit8不受影响"
End Sub

'===============================================================================
' 测试组13: 错误字符串辅助
'===============================================================================

Private Sub TestErrorStrings()
    BeginTestGroup "错误字符串辅助"

    ' 通信错误描述
    Dim s As String
    s = GetCommErrorString(CE_RXOVER Or CE_RXPARITY)
    AssertTrue InStr(s, "溢出") > 0, "CE_RXOVER描述包含'溢出'"
    AssertTrue InStr(s, "校验") > 0, "CE_RXPARITY描述包含'校验'"

    ' 空错误
    s = GetCommErrorString(0)
    AssertEqual Len(s), 0, "CE=0无描述"
End Sub

'===============================================================================
' 测试组14: 事件字符串辅助
'===============================================================================

Private Sub TestEventStrings()
    BeginTestGroup "事件字符串辅助"

    Dim s As String
    s = GetCommEventString(EV_RXCHAR Or EV_CTS)
    AssertTrue InStr(s, "字符接收") > 0, "EV_RXCHAR包含'字符接收'"
    AssertTrue InStr(s, "CTS") > 0, "EV_CTS包含'CTS'"

    ' COMSTAT描述
    s = GetCommStatString(1)  ' fCtsHold
    AssertTrue InStr(s, "CTS") > 0, "fCtsHold包含'CTS'"
End Sub

'===============================================================================
' 测试组15: 串口枚举
'===============================================================================

Private Sub TestEnumPorts()
    BeginTestGroup "串口枚举"

    Dim ports As Collection
    Set ports = EnumSerialPorts()
    ' 不断言具体数量（因硬件而异），只验证返回类型
    AssertTrue Not ports Is Nothing, "EnumSerialPorts返回非Nothing"
    Debug.Print "  [INFO] 发现 " & CStr(ports.count) & " 个可用串口"
    Dim i As Long
    For i = 1 To ports.count
        Debug.Print "  [INFO]   " & ports(i)
    Next i
End Sub

'===============================================================================
' 以下测试需要串口硬件或虚拟串口对（如com0com/Virtual Serial Port Driver）
' 如果没有可用的串口，这些测试将跳过
'===============================================================================

' 获取第一个可用的串口号
Private Function GetFirstAvailablePort() As String
    Dim ports As Collection
    Set ports = EnumSerialPorts()
    If ports.count > 0 Then
        GetFirstAvailablePort = ports(1)
    Else
        GetFirstAvailablePort = ""
    End If
End Function

'===============================================================================
' 测试组16: 打开和关闭实际串口
'===============================================================================

Private Sub TestPortOpenClose()
    BeginTestGroup "打开和关闭实际串口"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    Dim result As Boolean

    ' 打开
    result = sp.OpenPort(portName)
    AssertTrue result, "打开" & portName & "成功"
    AssertTrue sp.IsOpen, "IsOpen=True"
    Debug.Print "  [INFO] 已打开 " & portName

    ' 再次打开应报错
    On Error Resume Next
    ERR.Clear
    Dim result2 As Boolean
    result2 = sp.OpenPort(portName)
    AssertTrue ERR.Number <> 0, "重复打开应报错"
    On Error GoTo 0

    ' 关闭
    sp.ClosePort
    AssertFalse sp.IsOpen, "关闭后IsOpen=False"

    ' 关闭后可重新打开
    result = sp.OpenPort(portName)
    AssertTrue result, "重新打开成功"
    sp.ClosePort
End Sub

'===============================================================================
' 测试组17: 动态配置应用
'===============================================================================

Private Sub TestPortConfigApply()
    BeginTestGroup "动态配置应用"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    sp.portName = portName

    ' 打开时指定配置
    Dim cfg As New cSerialConfig
    cfg.BaudRate = br115200
    cfg.Parity = ptEven

    If sp.OpenPort(, cfg) Then
        ' 验证配置已应用
        AssertEqual sp.Config.BaudRate, br115200, "配置BaudRate=115200"
        AssertEqual sp.Config.Parity, ptEven, "配置Parity=Even"

        ' 动态修改配置
        sp.Config.BaudRate = br9600
        sp.ApplyConfig
        AssertEqual sp.Config.BaudRate, br9600, "动态修改BaudRate=9600"

        ' 从端口读取当前DCB
        sp.RefreshConfig
        AssertEqual sp.Config.BaudRate, br9600, "RefreshConfig确认BaudRate=9600"

        sp.ClosePort
    Else
        Debug.Print "  [SKIP] 打开" & portName & "失败"
    End If
End Sub

'===============================================================================
' 测试组18: 读写测试
'===============================================================================

Private Sub TestPortReadWrite()
    BeginTestGroup "读写测试"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    If Not sp.OpenPort(portName) Then
        Debug.Print "  [SKIP] 打开" & portName & "失败"
        Exit Sub
    End If

    ' 写入文本
    Dim bytesWritten As Long
    bytesWritten = sp.WriteText("Hello Serial!")
    AssertTrue bytesWritten > 0, "WriteText返回>0"
    AssertTrue sp.SentCount > 0, "SentCount>0"

    ' 写入字节数组
    Dim buf(2) As Byte
    buf(0) = &H41: buf(1) = &H42: buf(2) = &H43
    bytesWritten = sp.WriteData(buf)
    AssertEqual bytesWritten, 3, "WriteData 3字节"

    ' 写入单字节
    sp.WriteByte &HD
    AssertTrue sp.SentCount > 3, "WriteByte后SentCount递增"

    ' 写入一行
    bytesWritten = sp.WriteLine("Test Line")
    AssertTrue bytesWritten > 0, "WriteLine返回>0"

    ' 优先传输字符
    Dim txResult As Boolean
    txResult = sp.TransmitImmediate(&H55)
    AssertTrue txResult, "TransmitImmediate成功"

    ' 检查输入缓冲区
    Dim inCount As Long
    inCount = sp.InBufferCount
    Debug.Print "  [INFO] InBufferCount=" & CStr(inCount)
    ' (没有回环时buffer可能为0，不断言)

    ' 读取（如果有数据）
    If inCount > 0 Then
        Dim readBuf() As Byte
        Dim bytesRead As Long
        bytesRead = sp.ReadData(readBuf)
        AssertTrue bytesRead > 0, "ReadData返回>0"
        Debug.Print "  [INFO] 读取到 " & CStr(bytesRead) & " 字节"
    End If

    sp.ClosePort
End Sub

'===============================================================================
' 测试组19: 信号线状态
'===============================================================================

Private Sub TestSignalLines()
    BeginTestGroup "信号线状态"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    If Not sp.OpenPort(portName) Then
        Debug.Print "  [SKIP] 打开端口失败"
        Exit Sub
    End If

    ' 读取信号线状态（因硬件而异，只验证不报错）
    Dim cts As Boolean: cts = sp.CtsHolding
    Dim dsr As Boolean: dsr = sp.DsrHolding
    Dim ring As Boolean: ring = sp.RingHolding
    Dim cd As Boolean: cd = sp.CdHolding
    Debug.Print "  [INFO] CTS=" & CStr(cts) & " DSR=" & CStr(dsr) & " RING=" & CStr(ring) & " CD=" & CStr(cd)

    ' 获取完整Modem状态字
    Dim modemStat As Long
    modemStat = sp.ReadModemStatus()
    AssertTrue modemStat >= 0, "ReadModemStatus成功"

    ' 设置DTR/RTS（不应报错）
    sp.SetDTR True
    sp.SetRTS True
    AssertTrue True, "SetDTR/SetRTS不报错"

    sp.SetDTR False
    sp.SetRTS False
    AssertTrue True, "清除DTR/RTS不报错"

    ' XON/XOFF
    sp.SendXOn
    sp.SendXOff
    AssertTrue True, "SendXOn/SendXOff不报错"

    sp.ClosePort
End Sub

'===============================================================================
' 测试组20: Break信号
'===============================================================================

Private Sub TestBreakSignal()
    BeginTestGroup "Break信号"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    If Not sp.OpenPort(portName) Then
        Debug.Print "  [SKIP] 打开端口失败"
        Exit Sub
    End If

    ' 设置Break
    sp.SetBreak True
    AssertTrue True, "SetBreak(True)不报错"

    ' 短暂延时
    Dim t As Single: t = Timer
    Do While Timer - t < 0.1: DoEvents: Loop

    ' 清除Break
    sp.SetBreak False
    AssertTrue True, "SetBreak(False)不报错"

    sp.ClosePort
End Sub

'===============================================================================
' 测试组21: 缓冲区清空
'===============================================================================

Private Sub TestBufferPurge()
    BeginTestGroup "缓冲区清空"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    If Not sp.OpenPort(portName) Then
        Debug.Print "  [SKIP] 打开端口失败"
        Exit Sub
    End If

    ' 写入一些数据
    sp.WriteText "PurgeTest"

    ' 清空发送缓冲区
    sp.PurgeTx
    AssertEqual sp.OutBufferCount, 0, "PurgeTx后OutBuffer=0"

    ' 清空接收缓冲区
    sp.PurgeRx
    AssertEqual sp.InBufferCount, 0, "PurgeRx后InBuffer=0"

    ' 写入后清空全部
    sp.WriteText "PurgeAll"
    sp.PurgeAll
    AssertEqual sp.InBufferCount, 0, "PurgeAll后InBuffer=0"
    AssertEqual sp.OutBufferCount, 0, "PurgeAll后OutBuffer=0"

    sp.ClosePort
End Sub

'===============================================================================
' 测试组22: 异步监控
'===============================================================================

Private Sub TestMonitoring()
    BeginTestGroup "异步监控"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    If Not sp.OpenPort(portName) Then
        Debug.Print "  [SKIP] 打开端口失败"
        Exit Sub
    End If

    ' 启动监控
    Dim result As Boolean
    result = sp.StartMonitoring(100)
    AssertTrue result, "StartMonitoring成功"
    AssertTrue sp.Monitoring, "Monitoring=True"

    ' 等待一小段时间
    Dim t As Single: t = Timer
    Do While Timer - t < 0.5: DoEvents: Loop

    ' 停止监控
    sp.StopMonitoring
    AssertFalse sp.Monitoring, "StopMonitoring后Monitoring=False"

    sp.ClosePort
End Sub

'===============================================================================
' 测试组23: 状态字符串
'===============================================================================

Private Sub TestPortStatus()
    BeginTestGroup "状态字符串"

    Dim portName As String
    portName = GetFirstAvailablePort()
    If Len(portName) = 0 Then
        Debug.Print "  [SKIP] 无可用串口"
        Exit Sub
    End If

    Dim sp As New cSerialPort
    If Not sp.OpenPort(portName) Then
        Debug.Print "  [SKIP] 打开端口失败"
        Exit Sub
    End If

    Dim status As String
    status = sp.GetStatusString()
    AssertTrue Len(status) > 0, "GetStatusString非空"
    AssertTrue InStr(status, portName) > 0, "状态包含端口名"
    AssertTrue InStr(status, "BaudRate") > 0, "状态包含BaudRate"

    Debug.Print "  [INFO] 状态信息:" & vbCrLf & status

    sp.ClosePort
End Sub

'===============================================================================
' 测试组24: 自收自发(回环)测试
' 需要虚拟串口对（如com0com）将COMx的TX连接到COMx的RX
' 或者短路物理串口的TX/RX引脚
'===============================================================================

Private Sub TestLoopback()
    BeginTestGroup "自收自发(回环)测试"

    Dim ports As Collection
    Set ports = EnumSerialPorts()

    ' 需要至少2个串口（虚拟串口对）
    If ports.count < 2 Then
        Debug.Print "  [SKIP] 需要至少2个串口进行回环测试"
        Exit Sub
    End If

    Dim spTx As New cSerialPort
    Dim spRx As New cSerialPort

    ' 打开两个串口
    If Not spTx.OpenPort(ports(1)) Then
        Debug.Print "  [SKIP] 打开" & ports(1) & "失败"
        Exit Sub
    End If

    If Not spRx.OpenPort(ports(2)) Then
        Debug.Print "  [SKIP] 打开" & ports(2) & "失败"
        spTx.ClosePort
        Exit Sub
    End If

    ' 确保双方配置一致
    Dim cfg As New cSerialConfig
    cfg.BaudRate = br9600
    cfg.DataBits = 8
    cfg.Parity = ptNone
    cfg.StopBits = sb1
    Set spTx.Config = cfg
    Set spRx.Config = cfg.Clone()

    ' 清空缓冲区
    spRx.PurgeAll
    spTx.PurgeAll

    ' 发送
    Dim testStr As String
    testStr = "LOOPBACK_TEST_12345"
    Dim nWritten As Long
    nWritten = spTx.WriteText(testStr)
    Debug.Print "  [INFO] 发送 " & CStr(nWritten) & " 字节到 " & ports(1)

    ' 等待数据到达
    Dim t As Single: t = Timer
    Do While spRx.InBufferCount < Len(testStr) And (Timer - t) < 2
        DoEvents
    Loop

    ' 接收
    Dim received As String
    received = spRx.ReadExisting()
    Debug.Print "  [INFO] 从 " & ports(2) & " 接收: " & received

    If Len(received) > 0 Then
        AssertEqual received, testStr, "回环数据匹配"
    Else
        Debug.Print "  [INFO] 未收到数据(可能不是虚拟串口对)"
    End If

    spTx.ClosePort
    spRx.ClosePort
End Sub

'===============================================================================
' COM3 收发测试 - 发送 hello，期望接收 world（10秒超时）
' 使用方式: 在立即窗口输入 TestCOM3HelloWorld 回车执行
' 前提: COM3 已连接对端设备，对端收到 hello 后会回应 world
'===============================================================================

Public Sub TestCOM3HelloWorld()
    Dim sp As New cSerialPort
    sp.PortName = "COM3"
    sp.Config.BaudRate = br9600
    sp.Config.DataBits = 8
    sp.Config.Parity = ptNone
    sp.Config.StopBits = sb1
    sp.Config.SetReadTimeout 10000   ' 读取超时10秒

    If Not sp.OpenPort() Then
        Debug.Print "[FAIL] 无法打开 COM3: " & sp.LastErrorMsg
        Exit Sub
    End If

    Debug.Print "[INFO] COM3 已打开，正在发送 hello ..."

    ' 发送
    Dim nSent As Long
    nSent = sp.WriteText("hello")
    Debug.Print "[INFO] 已发送 " & nSent & " 字节"

    ' 等待接收（带10秒超时）
    Dim received As String
    Dim startTime As Single
    startTime = Timer

    Do
        DoEvents
        If sp.InBufferCount > 0 Then
            received = received & sp.ReadExisting()
            ' 检查是否已收到完整响应
            If InStr(received, "world") > 0 Then Exit Do
        End If
    Loop While Timer - startTime < 10

    ' 结果判定
    If InStr(received, "world") > 0 Then
        Debug.Print "[PASS] 收到预期响应: " & received
    ElseIf Len(received) > 0 Then
        Debug.Print "[FAIL] 收到非预期数据: " & received
    Else
        Debug.Print "[FAIL] 10秒内未收到任何数据"
    End If

    sp.ClosePort
    Debug.Print "[INFO] COM3 已关闭"
End Sub
