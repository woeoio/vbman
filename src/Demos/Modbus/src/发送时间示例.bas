Attribute VB_Name = "发送时间示例"
'=========================================================================
'
' 发送时间示例 - 演示如何将当前时间发送到 PLC
'
' Purpose: 提供完整的示例，展示如何读取 PLC 数据和发送时间到 PLC
'
' Author: Auto
' Date: 2026-01-21
'
'=========================================================================
Option Explicit

'=========================================================================
' 示例1：发送当前时间（拆分为年、月、日、时、分、秒）
'=========================================================================

Public Sub SendCurrentTimeToPLC()
    Dim mb As New cModbus
    Dim values(5) As Integer
    Dim dtNow As Date
    Dim bSuccess As Boolean
    
    On Error GoTo ErrorHandler
    
    ' ========== 配置连接 ==========
    mb.ProtocolType = MB_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"  ' 修改为你的 PLC IP 地址
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    Debug.Print "=== 发送当前时间到 PLC ==="
    Debug.Print "正在连接到: " & mb.TCPHost & ":" & mb.TCPPort
    
    ' ========== 连接 ==========
    mb.Connect
    Debug.Print "连接成功"
    
    ' ========== 获取当前时间 ==========
    dtNow = Now
    Debug.Print "当前时间: " & Format$(dtNow, "yyyy-mm-dd hh:nn:ss")
    
    ' ========== 将时间拆分为多个寄存器 ==========
    values(0) = Year(dtNow)      ' 年（如 2026）
    values(1) = Month(dtNow)     ' 月（1-12）
    values(2) = Day(dtNow)       ' 日（1-31）
    values(3) = Hour(dtNow)     ' 时（0-23）
    values(4) = Minute(dtNow)   ' 分（0-59）
    values(5) = Second(dtNow)    ' 秒（0-59）
    
    Debug.Print "时间数据:"
    Debug.Print "  年: " & values(0)
    Debug.Print "  月: " & values(1)
    Debug.Print "  日: " & values(2)
    Debug.Print "  时: " & values(3)
    Debug.Print "  分: " & values(4)
    Debug.Print "  秒: " & values(5)
    
    ' ========== 写入到 PLC（从地址 100 开始） ==========
    Debug.Print vbCrLf & "正在写入到寄存器地址 100-105..."
    bSuccess = mb.WriteMultipleRegisters(100, values)
    
    If bSuccess Then
        Debug.Print "? 时间写入成功"
    Else
        Debug.Print "? 时间写入失败"
    End If
    
    ' ========== 验证写入：重新读取 ==========
    Debug.Print vbCrLf & "验证写入，重新读取寄存器 100-105..."
    Dim regs() As Integer
    regs = mb.ReadHoldingRegisters(100, 6)
    
    Debug.Print "读取到的值:"
    Dim i As Long
    For i = 0 To UBound(regs)
        Debug.Print "  寄存器 " & (100 + i) & ": " & regs(i)
    Next i
    
    ' ========== 断开连接 ==========
    mb.Disconnect
    Debug.Print vbCrLf & "已断开连接"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub

'=========================================================================
' 示例2：发送 Unix 时间戳（32位，需要2个寄存器）
'=========================================================================

Public Sub SendUnixTimestampToPLC()
    Dim mb As New cModbus
    Dim values(1) As Integer
    Dim dtNow As Date
    Dim lTimestamp As Long
    Dim bSuccess As Boolean
    
    On Error GoTo ErrorHandler
    
    ' ========== 配置连接 ==========
    mb.ProtocolType = MB_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"  ' 修改为你的 PLC IP 地址
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    Debug.Print "=== 发送 Unix 时间戳到 PLC ==="
    Debug.Print "正在连接到: " & mb.TCPHost & ":" & mb.TCPPort
    
    ' ========== 连接 ==========
    mb.Connect
    Debug.Print "连接成功"
    
    ' ========== 计算 Unix 时间戳 ==========
    dtNow = Now
    Dim dtEpoch As Date
    dtEpoch = #1/1/1970#
    lTimestamp = CLng((dtNow - dtEpoch) * 86400)
    
    Debug.Print "当前时间: " & Format$(dtNow, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Unix 时间戳: " & lTimestamp
    
    ' ========== 将 32 位时间戳拆分为两个 16 位寄存器 ==========
    ' 高 16 位
    values(0) = CInt((lTimestamp \ 65536) And &HFFFF)
    ' 低 16 位
    values(1) = CInt(lTimestamp And &HFFFF)
    
    Debug.Print "时间戳拆分:"
    Debug.Print "  高16位（寄存器0）: " & values(0)
    Debug.Print "  低16位（寄存器1）: " & values(1)
    
    ' ========== 写入到 PLC（从地址 200 开始） ==========
    Debug.Print vbCrLf & "正在写入到寄存器地址 200-201..."
    bSuccess = mb.WriteMultipleRegisters(200, values)
    
    If bSuccess Then
        Debug.Print "? 时间戳写入成功"
    Else
        Debug.Print "? 时间戳写入失败"
    End If
    
    ' ========== 断开连接 ==========
    mb.Disconnect
    Debug.Print vbCrLf & "已断开连接"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub

'=========================================================================
' 示例3：完整的读写示例（读取 PLC 数据 + 发送时间）
'=========================================================================

Public Sub CompleteReadWriteExample()
    Dim mb As New cModbus
    Dim regs() As Integer
    Dim values(5) As Integer
    Dim dtNow As Date
    Dim i As Long
    Dim bSuccess As Boolean
    
    On Error GoTo ErrorHandler
    
    ' ========== 配置连接 ==========
    mb.ProtocolType = MB_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"  ' 修改为你的 PLC IP 地址
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    Debug.Print "========================================"
    Debug.Print "Modbus 完整读写示例"
    Debug.Print "========================================" & vbCrLf
    
    Debug.Print "正在连接到: " & mb.TCPHost & ":" & mb.TCPPort
    
    ' ========== 连接 ==========
    mb.Connect
    Debug.Print "? 连接成功" & vbCrLf
    
    ' ========== 步骤1：读取 PLC 数据 ==========
    Debug.Print "=== 步骤1：读取 PLC 数据 ==="
    Debug.Print "从地址 0 读取 10 个保持寄存器..."
    
    regs = mb.ReadHoldingRegisters(0, 10)
    
    Debug.Print "读取成功，共 " & (UBound(regs) + 1) & " 个寄存器:"
    For i = 0 To UBound(regs)
        Debug.Print "  寄存器 " & i & ": " & regs(i)
    Next i
    Debug.Print ""
    
    ' ========== 步骤2：发送当前时间到 PLC ==========
    Debug.Print "=== 步骤2：发送当前时间到 PLC ==="
    dtNow = Now
    Debug.Print "当前时间: " & Format$(dtNow, "yyyy-mm-dd hh:nn:ss")
    
    ' 将时间拆分为多个寄存器
    values(0) = Year(dtNow)
    values(1) = Month(dtNow)
    values(2) = Day(dtNow)
    values(3) = Hour(dtNow)
    values(4) = Minute(dtNow)
    values(5) = Second(dtNow)
    
    Debug.Print "写入到寄存器地址 100-105..."
    bSuccess = mb.WriteMultipleRegisters(100, values)
    
    If bSuccess Then
        Debug.Print "? 时间写入成功"
    Else
        Debug.Print "? 时间写入失败"
    End If
    Debug.Print ""
    
    ' ========== 步骤3：验证写入 ==========
    Debug.Print "=== 步骤3：验证写入 ==="
    Debug.Print "重新读取寄存器 100-105..."
    
    regs = mb.ReadHoldingRegisters(100, 6)
    
    Debug.Print "读取到的值:"
    For i = 0 To UBound(regs)
        Debug.Print "  寄存器 " & (100 + i) & ": " & regs(i)
    Next i
    
    ' 验证数据是否正确
    Dim bMatch As Boolean
    bMatch = True
    For i = 0 To 5
        If regs(i) <> values(i) Then
            bMatch = False
            Exit For
        End If
    Next i
    
    If bMatch Then
        Debug.Print "? 验证通过：写入的数据与读取的数据一致"
    Else
        Debug.Print "? 验证失败：写入的数据与读取的数据不一致"
    End If
    Debug.Print ""
    
    ' ========== 断开连接 ==========
    mb.Disconnect
    Debug.Print "已断开连接"
    Debug.Print "========================================"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub

'=========================================================================
' 示例4：定时发送时间（需要在定时器中调用）
'=========================================================================

Public Sub SendTimePeriodically()
    ' 这个函数可以在定时器中调用，定期发送当前时间到 PLC
    ' 注意：每次调用都会重新连接和断开，实际应用中应该保持连接
    
    Static lLastSecond As Long
    Dim lCurrentSecond As Long
    
    ' 获取当前秒数
    lCurrentSecond = CLng(Second(Now))
    
    ' 每秒执行一次（避免频繁发送）
    If lCurrentSecond <> lLastSecond Then
        lLastSecond = lCurrentSecond
        
        ' 发送时间
        SendCurrentTimeToPLC
    End If
End Sub

'=========================================================================
' 示例5：RTU 模式发送时间
'=========================================================================

Public Sub SendTimeRTU()
    Dim mb As New cModbus
    Dim values(5) As Integer
    Dim dtNow As Date
    Dim bSuccess As Boolean
    
    On Error GoTo ErrorHandler
    
    ' ========== 配置 RTU 连接 ==========
    mb.ProtocolType = MB_PROTOCOL_RTU
    mb.SlaveID = 1
    mb.SerialPort = "COM3"        ' 修改为你的串口
    mb.BaudRate = 9600
    mb.DataBits = 8
    mb.Parity = "N"
    mb.StopBits = 1
    mb.ResponseTimeout = 2000
    
    Debug.Print "=== RTU 模式：发送当前时间到 PLC ==="
    Debug.Print "串口: " & mb.SerialPort
    Debug.Print "波特率: " & mb.BaudRate
    
    ' ========== 连接 ==========
    mb.Connect
    Debug.Print "连接成功"
    
    ' ========== 获取当前时间 ==========
    dtNow = Now
    Debug.Print "当前时间: " & Format$(dtNow, "yyyy-mm-dd hh:nn:ss")
    
    ' ========== 将时间拆分为多个寄存器 ==========
    values(0) = Year(dtNow)
    values(1) = Month(dtNow)
    values(2) = Day(dtNow)
    values(3) = Hour(dtNow)
    values(4) = Minute(dtNow)
    values(5) = Second(dtNow)
    
    ' ========== 写入到 PLC ==========
    Debug.Print "正在写入到寄存器地址 100-105..."
    bSuccess = mb.WriteMultipleRegisters(100, values)
    
    If bSuccess Then
        Debug.Print "? 时间写入成功"
    Else
        Debug.Print "? 时间写入失败"
    End If
    
    ' ========== 断开连接 ==========
    mb.Disconnect
    Debug.Print "已断开连接"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub
