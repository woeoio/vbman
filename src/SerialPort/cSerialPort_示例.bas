Attribute VB_Name = "mSerialPortExample"
Option Explicit

'================================================================================
' cSerialPort 使用示例
'================================================================================

' 示例1: 基本使用
Public Sub Example_BasicUsage()
    Dim sp As New cSerialPort
    
    ' 打开串口 (波特率 115200)
    If sp.Open("COM1", 115200, 8, peNone, sb1, fcNone) Then
        Debug.Print "串口打开成功!"
        
        ' 启动事件监听
        sp.StartListening
        
        ' 发送数据
        Dim lSent As Long
        lSent = sp.WriteString("Hello, Serial Port!")
        Debug.Print "发送字节数: " & lSent
        
        ' 关闭串口
        sp.Close
    Else
        Debug.Print "串口打开失败!"
    End If
    
    Set sp = Nothing
End Sub

'================================================================================
' 示例2: 高波特率 (921600) 使用
'================================================================================
Public Sub Example_HighBaudRate()
    Dim sp As New cSerialPort
    
    ' 设置更大的缓冲区以适应高波特率
    sp.InQueueSize = 16384
    sp.OutQueueSize = 16384
    
    ' 打开串口 (波特率 921600)
    If sp.Open("COM1", br921600, 8, peNone, sb1, fcRTSCTS) Then
        Debug.Print "高波特率串口打开成功!"
        
        ' 启动事件监听
        sp.StartListening
        
        ' 发送大量数据
        Dim lSent As Long
        Dim i As Long
        Dim sData As String
        
        For i = 1 To 100
            sData = sData & "Data " & i & vbCrLf
        Next i
        
            lSent = sp.WriteString(sData)
            Debug.Print "发送字节数: " & lSent
        
        ' 等待一段时间
        Dim lStart As Long
        lStart = Timer
        Do While Timer - lStart < 2
            DoEvents
        Loop
        
        ' 关闭串口
        sp.Close
    End If
    
    Set sp = Nothing
End Sub

'================================================================================
' 示例3: 使用事件 (需要在窗体模块中使用 WithEvents)
'================================================================================
' 在窗体模块中声明:
' Private WithEvents m_SerialPort As cSerialPort
'
' Private Sub Form_Load()
'     Set m_SerialPort = New cSerialPort
'     If m_SerialPort.Open("COM1", 115200, 8, peNone, sb1, fcNone) Then
'         m_SerialPort.StartListening
'     End If
' End Sub
'
' Private Sub Form_Unload(Cancel As Integer)
'     m_SerialPort.StopListening
'     m_SerialPort.Close
'     Set m_SerialPort = Nothing
' End Sub
'
' ' 数据接收事件 (字符串格式)
' Private Sub m_SerialPort_OnDataReceived(ByVal Data As String, ByVal Length As Long)
'     Debug.Print "收到数据: " & Data
'     Debug.Print "数据长度: " & Length
' End Sub
'
' ' 数据接收事件 (字节数组格式)
' Private Sub m_SerialPort_OnDataReceivedBytes(ByRef Data() As Byte, ByVal Length As Long)
'     Debug.Print "收到字节数据, 长度: " & Length
'     ' 处理二进制数据...
' End Sub
'
' ' 错误事件
' Private Sub m_SerialPort_OnError(ByVal ErrorCode As Long, ByVal ErrorMessage As String)
'     Debug.Print "错误: " & ErrorMessage & " (代码: " & ErrorCode & ")"
' End Sub
'
' ' CTS 信号变化
' Private Sub m_SerialPort_OnCTSChanged(ByVal State As Boolean)
'     Debug.Print "CTS 信号: " & IIf(State, "高", "低")
' End Sub
'
' ' DSR 信号变化
' Private Sub m_SerialPort_OnDSRChanged(ByVal State As Boolean)
'     Debug.Print "DSR 信号: " & IIf(State, "高", "低")
' End Sub

'================================================================================
' 示例4: 发送字节数组 (二进制数据)
'================================================================================
Public Sub Example_SendBinaryData()
    Dim sp As New cSerialPort
    Dim bytData() As Byte
    Dim i As Long
    
    ' 准备二进制数据
    ReDim bytData(255)
    For i = 0 To 255
        bytData(i) = CByte(i)
    Next i
    
    ' 打开串口
    If sp.Open("COM1", 115200) Then
        sp.StartListening
        
        ' 发送字节数组
        Dim lSent As Long
        lSent = sp.WriteBytes(bytData)
        Debug.Print "发送二进制数据: " & lSent & " 字节"
        
        sp.Close
    End If
    
    Set sp = Nothing
End Sub

'================================================================================
' 示例4.5: 发送/读取 HEX 字符串
'================================================================================
Public Sub Example_HexString()
    Dim sp As New cSerialPort
    
    ' 打开串口
    If sp.Open("COM1", 115200) Then
        sp.StartListening
        
        ' 发送 HEX 字符串 (支持各种分隔符)
        ' 输入格式可以是: "A1B2C3", "A1 B2 C3", "A1-B2-C3" 等
        Dim lSent As Long
        lSent = sp.WriteHex("A1 B2 C3 D4 E5 F6")
        Debug.Print "发送 HEX: " & lSent & " 字节"
        
        ' 等待一段时间
        Dim lStart As Long
        lStart = Timer
        Do While Timer - lStart < 1
            DoEvents
        Loop
        
        ' 读取 HEX 字符串 (默认用空格分隔)
        Dim sHex As String
        sHex = sp.ReadHex(1024)
        Debug.Print "读取 HEX (空格分隔): " & sHex
        
        ' 读取 HEX 字符串 (使用横线分隔)
        sHex = sp.ReadHex(1024, "-")
        Debug.Print "读取 HEX (横线分隔): " & sHex
        
        ' 读取 HEX 字符串 (无分隔符)
        sHex = sp.ReadHex(1024, "")
        Debug.Print "读取 HEX (无分隔符): " & sHex
        
        ' 等待一段时间
        lStart = Timer
        Do While Timer - lStart < 1
            DoEvents
        Loop
        
        ' 另一个示例: 发送无分隔符的 HEX
        lSent = sp.WriteHex("0102030405060708")
        Debug.Print "发送无分隔符 HEX: " & lSent & " 字节"
        
        sp.Close
    End If
    
    Set sp = Nothing
End Sub

'================================================================================
' 示例5: 不同流控制方式
'================================================================================
Public Sub Example_FlowControl()
    Dim sp As New cSerialPort
    
    ' 无流控
    sp.Open "COM1", 115200, 8, peNone, sb1, fcNone
    Debug.Print "无流控模式"
    sp.Close
    
    ' 软件流控 (XOn/XOff)
    sp.Open "COM1", 115200, 8, peNone, sb1, fcXOnXOff
    Debug.Print "软件流控模式 (XOn/XOff)"
    sp.Close
    
    ' 硬件流控 (RTS/CTS)
    sp.Open "COM1", 115200, 8, peNone, sb1, fcRTSCTS
    Debug.Print "硬件流控模式 (RTS/CTS)"
    sp.Close
    
    Set sp = Nothing
End Sub

'================================================================================
' 示例6: 动态重新配置
'================================================================================
Public Sub Example_ReConfigure()
    Dim sp As New cSerialPort
    
    ' 初始配置: 9600 波特率
    sp.Open "COM1", 9600, 8, peNone, sb1, fcNone
    Debug.Print "初始波特率: " & sp.BaudRate
    
    ' 运行时重新配置: 115200 波特率
    If sp.ReConfigure(115200) Then
        Debug.Print "新波特率: " & sp.BaudRate
    End If
    
    ' 修改数据位和校验位
    If sp.ReConfigure(-1, 7, peOdd) Then
        Debug.Print "新配置: " & sp.DataBits & " 数据位, 奇校验"
    End If
    
    sp.Close
    Set sp = Nothing
End Sub

'================================================================================
' 示例7: 缓冲区管理
'================================================================================
Public Sub Example_BufferManagement()
    Dim sp As New cSerialPort
    
    ' 打开串口
    sp.Open "COM1", 115200
    sp.StartListening
    
    ' 发送一些数据
    sp.WriteString "Test Data"
    
    ' 检查队列状态
    Debug.Print "接收队列: " & sp.BytesInQueue & " 字节"
    Debug.Print "发送队列: " & sp.BytesOutQueue & " 字节"
    sp.ClearSendBuffer
    
    ' 清空接收缓冲区
    sp.ClearReceiveBuffer
    
    ' 或一次性清空所有缓冲区
    sp.ClearBuffer
    
    ' 刷新缓冲区
    sp.Flush
    
    sp.Close
    Set sp = Nothing
End Sub

'================================================================================
' 示例8: 信号控制 (DTR/RTS)
'================================================================================
Public Sub Example_SignalControl()
    Dim sp As New cSerialPort
    
    sp.Open "COM1", 115200
    
    ' 设置 DTR
    sp.SetDTR
    Debug.Print "DTR 已设置"
    
    ' 清除 DTR
    sp.ClearDTR
    Debug.Print "DTR 已清除"
    
    ' 设置 RTS
    sp.SetRTS
    Debug.Print "RTS 已设置"
    
    ' 清除 RTS
    sp.ClearRTS
    Debug.Print "RTS 已清除"
    
    ' 设置中断
    sp.SetBreak
    Debug.Print "中断已设置"
    
    ' 清除中断
    sp.ClearBreak
    Debug.Print "中断已清除"
    
    sp.Close
    Set sp = Nothing
End Sub

'================================================================================
' 示例9: 读取数据
'================================================================================
Public Sub Example_ReadData()
    Dim sp As New cSerialPort
    
    sp.Open "COM1", 115200
    sp.StartListening
    
    ' 发送数据 (假设有回环)
    sp.WriteString "AT"
    
    ' 等待一段时间
    Dim lStart As Long
    lStart = Timer
    Do While Timer - lStart < 1 And sp.BytesInQueue = 0
        DoEvents
    Loop
    
    ' 读取数据 (字符串)
    Dim sData As String
    sData = sp.ReadString(1024)
    Debug.Print "读取数据: " & sData
    
    ' 读取数据 (字节数组)
    Dim bytData() As Byte
    bytData = sp.ReadBytes(1024)
    Debug.Print "读取字节: " & UBound(bytData) + 1
    
    ' 读取数据 (HEX 字符串)
    Dim sHex As String
    sHex = sp.ReadHex(1024)
    Debug.Print "读取 HEX: " & sHex
    
    sp.Close
    Set sp = Nothing
End Sub

'================================================================================
' 示例10: 获取所有可用串口 (需要枚举)
'================================================================================
Public Sub Example_ListAvailablePorts()
    ' 这是一个简单的枚举方法
    ' 实际应用中可以使用 Windows API 或 WMI 来获取真实可用的串口列表
    
    Dim i As Integer
    Debug.Print "尝试检测可用串口:"
    
    For i = 1 To 20
        Dim sp As New cSerialPort
        On Error Resume Next
        sp.Open "COM" & i, 9600
        If Err.Number = 0 And sp.IsOpen Then
            Debug.Print "COM" & i & " - 可用"
            sp.Close
        End If
        On Error GoTo 0
        Set sp = Nothing
    Next i
End Sub

'================================================================================
' 示例11: Modbus RTU 通信
'================================================================================
Public Sub Example_ModbusRTU()
    Dim sp As New cSerialPort
    
    ' Modbus RTU 典型配置
    If sp.Open("COM1", 9600, 8, peNone, sb1, fcNone) Then
        sp.StartListening
        
        ' 构建 Modbus 读取请求帧
        ' 示例: 读取从机地址 1, 功能码 03, 起始地址 0, 读取 2 个寄存器
        Dim bytFrame() As Byte
        
        ' 从机地址
        ReDim bytFrame(5)
        bytFrame(0) = &H1       ' 从机地址
        bytFrame(1) = &H3       ' 功能码 (读取保持寄存器)
        bytFrame(2) = &H0       ' 起始地址高位
        bytFrame(3) = &H0       ' 起始地址低位
        bytFrame(4) = &H0       ' 寄存器数量高位
        bytFrame(5) = &H2       ' 寄存器数量低位
        
        ' 计算 CRC16 (需要另外实现 CRC16 函数)
        ' Dim wCRC As Integer
        ' wCRC = CRC16(bytFrame)
        ' ReDim Preserve bytFrame(6)
        ' bytFrame(6) = CByte(wCRC And &HFF)
        ' bytFrame(7) = CByte((wCRC \ 256) And &HFF)
        
        ' 发送帧
        sp.WriteBytes bytFrame
        
        ' 等待响应...
        
        sp.Close
    End If
    
    Set sp = Nothing
End Sub

'================================================================================
' 完整的窗体示例
'================================================================================
' 复制以下代码到一个窗体模块中:
'
' Option Explicit
'
' Private WithEvents m_SerialPort As cSerialPort
'
' Private Sub Form_Load()
'     Set m_SerialPort = New cSerialPort
'
'     ' 配置串口
'     m_SerialPort.InQueueSize = 8192
'     m_SerialPort.OutQueueSize = 8192
'
'     ' 打开串口
'     If m_SerialPort.Open("COM1", 115200, 8, peNone, sb1, fcNone) Then
'         m_SerialPort.StartListening
'         Me.Caption = "串口已连接"
'     Else
'         Me.Caption = "串口连接失败"
'     End If
' End Sub
'
' Private Sub Form_Unload(Cancel As Integer)
'     If Not m_SerialPort Is Nothing Then
'         m_SerialPort.StopListening
'         m_SerialPort.Close
'         Set m_SerialPort = Nothing
'     End If
' End Sub
'
' Private Sub cmdSend_Click()
'     Dim lSent As Long
'     lSent = m_SerialPort.Write(txtSend.Text)
'     Debug.Print "已发送: " & lSent & " 字节"
' End Sub
'
' Private Sub m_SerialPort_OnDataReceived(ByVal Data As String, ByVal Length As Long)
'     ' 在文本框中显示接收到的数据
'     txtReceive.Text = txtReceive.Text & Data
'     ' 自动滚动到底部
'     txtReceive.SelStart = Len(txtReceive.Text)
' End Sub
'
' Private Sub m_SerialPort_OnError(ByVal ErrorCode As Long, ByVal ErrorMessage As String)
'     MsgBox "串口错误: " & ErrorMessage & vbCrLf & "错误代码: " & ErrorCode, vbExclamation
' End Sub
'
' Private Sub m_SerialPort_OnCTSChanged(ByVal State As Boolean)
'     lblCTS.Caption = IIf(State, "CTS: ON", "CTS: OFF")
' End Sub
'
' Private Sub m_SerialPort_OnDSRChanged(ByVal State As Boolean)
'     lblDSR.Caption = IIf(State, "DSR: ON", "DSR: OFF")
' End Sub
'
' Private Sub cmdClear_Click()
'     txtReceive.Text = ""
'     m_SerialPort.ClearBuffer
' End Sub
'
' Private Sub cmdSetDTR_Click()
'     If chkDTR.Value = 1 Then
'         m_SerialPort.SetDTR
'     Else
'         m_SerialPort.ClearDTR
'     End If
' End Sub
'
' Private Sub cmdSetRTS_Click()
'     If chkRTS.Value = 1 Then
'         m_SerialPort.SetRTS
'     Else
'         m_SerialPort.ClearRTS
'     End If
' End Sub
