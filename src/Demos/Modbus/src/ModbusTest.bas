Attribute VB_Name = "ModbusTest"
'=========================================================================
'
' ModbusTest - Modbus 测试模块
'
' Purpose: 提供 Modbus 功能的命令行测试示例
'
' Author: Auto
' Date: 2026-01-21
'
'=========================================================================
Option Explicit

'=========================================================================
' TCP 模式测试示例
'=========================================================================

Public Sub TestModbusTCP()
    Dim mb As New cModbus
    Dim aiRegs() As Integer
    Dim baCoils() As Boolean
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' 配置 TCP 连接
    mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    Debug.Print "=== Modbus TCP 测试 ==="
    Debug.Print "连接到: " & mb.TCPHost & ":" & mb.TCPPort
    
    ' 连接
    mb.Connect
    Debug.Print "连接成功"
    
    ' 测试读取保持寄存器
    Debug.Print vbCrLf & "读取保持寄存器 (地址 0, 数量 10)..."
    aiRegs = mb.ReadHoldingRegisters(0, 10)
    Debug.Print "读取成功，值: "
    For i = 0 To UBound(aiRegs)
        Debug.Print "  寄存器 " & i & ": " & aiRegs(i)
    Next i
    
    ' 测试写入单个寄存器
    Debug.Print vbCrLf & "写入单个寄存器 (地址 0, 值 1234)..."
    If mb.WriteSingleRegister(0, 1234) Then
        Debug.Print "写入成功"
    Else
        Debug.Print "写入失败"
    End If
    
    ' 测试读取线圈
    Debug.Print vbCrLf & "读取线圈 (地址 0, 数量 16)..."
    baCoils = mb.ReadCoils(0, 16)
    Debug.Print "读取成功，值: "
    For i = 0 To UBound(baCoils)
        Debug.Print "  线圈 " & i & ": " & IIf(baCoils(i), "ON", "OFF")
    Next i
    
    ' 测试写入单个线圈
    Debug.Print vbCrLf & "写入单个线圈 (地址 0, 值 ON)..."
    If mb.WriteSingleCoil(0, True) Then
        Debug.Print "写入成功"
    Else
        Debug.Print "写入失败"
    End If
    
    ' 测试写入多个寄存器
    Debug.Print vbCrLf & "写入多个寄存器 (地址 10, 数量 5)..."
    Dim aiValues(4) As Integer
    For i = 0 To 4
        aiValues(i) = 1000 + i
    Next i
    If mb.WriteMultipleRegisters(10, aiValues) Then
        Debug.Print "写入成功"
    Else
        Debug.Print "写入失败"
    End If
    
    ' 断开连接
    mb.Disconnect
    Debug.Print vbCrLf & "已断开连接"
    
    Exit Sub
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub

'=========================================================================
' RTU 模式测试示例
'=========================================================================

Public Sub TestModbusRTU()
    Dim mb As New cModbus
    Dim aiRegs() As Integer
    Dim baCoils() As Boolean
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' 配置 RTU 连接
    mb.ProtocolType = MB_MASTER_PROTOCOL_RTU
    mb.SlaveID = 1
    mb.SerialPort = "COM3"
    mb.BaudRate = 9600
    mb.DataBits = 8
    mb.Parity = "N"
    mb.StopBits = 1
    mb.ResponseTimeout = 2000
    
    Debug.Print "=== Modbus RTU 测试 ==="
    Debug.Print "串口: " & mb.SerialPort
    Debug.Print "波特率: " & mb.BaudRate & ", 数据位: " & mb.DataBits & ", 校验: " & mb.Parity & ", 停止位: " & mb.StopBits
    
    ' 连接
    mb.Connect
    Debug.Print "连接成功"
    
    ' 测试读取保持寄存器
    Debug.Print vbCrLf & "读取保持寄存器 (地址 0, 数量 10)..."
    aiRegs = mb.ReadHoldingRegisters(0, 10)
    Debug.Print "读取成功，值: "
    For i = 0 To UBound(aiRegs)
        Debug.Print "  寄存器 " & i & ": " & aiRegs(i)
    Next i
    
    ' 测试读取输入寄存器
    Debug.Print vbCrLf & "读取输入寄存器 (地址 0, 数量 10)..."
    aiRegs = mb.ReadInputRegisters(0, 10)
    Debug.Print "读取成功，值: "
    For i = 0 To UBound(aiRegs)
        Debug.Print "  寄存器 " & i & ": " & aiRegs(i)
    Next i
    
    ' 测试读取线圈
    Debug.Print vbCrLf & "读取线圈 (地址 0, 数量 16)..."
    baCoils = mb.ReadCoils(0, 16)
    Debug.Print "读取成功，值: "
    For i = 0 To UBound(baCoils)
        Debug.Print "  线圈 " & i & ": " & IIf(baCoils(i), "ON", "OFF")
    Next i
    
    ' 测试读取离散输入
    Debug.Print vbCrLf & "读取离散输入 (地址 0, 数量 16)..."
    baCoils = mb.ReadDiscreteInputs(0, 16)
    Debug.Print "读取成功，值: "
    For i = 0 To UBound(baCoils)
        Debug.Print "  输入 " & i & ": " & IIf(baCoils(i), "ON", "OFF")
    Next i
    
    ' 测试写入单个寄存器
    Debug.Print vbCrLf & "写入单个寄存器 (地址 0, 值 5678)..."
    If mb.WriteSingleRegister(0, 5678) Then
        Debug.Print "写入成功"
    Else
        Debug.Print "写入失败"
    End If
    
    ' 测试写入单个线圈
    Debug.Print vbCrLf & "写入单个线圈 (地址 0, 值 ON)..."
    If mb.WriteSingleCoil(0, True) Then
        Debug.Print "写入成功"
    Else
        Debug.Print "写入失败"
    End If
    
    ' 测试写入多个线圈
    Debug.Print vbCrLf & "写入多个线圈 (地址 0, 数量 8)..."
    Dim baValues(7) As Boolean
    For i = 0 To 7
        baValues(i) = ((i Mod 2) = 0)  ' 交替 0/1
    Next i
    If mb.WriteMultipleCoils(0, baValues) Then
        Debug.Print "写入成功"
    Else
        Debug.Print "写入失败"
    End If
    
    ' 断开连接
    mb.Disconnect
    Debug.Print vbCrLf & "已断开连接"
    
    Exit Sub
ErrorHandler:
    Debug.Print "错误: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub

'=========================================================================
' 综合测试示例
'=========================================================================

Public Sub TestModbusAll()
    Debug.Print "========================================"
    Debug.Print "Modbus 完整功能测试"
    Debug.Print "========================================" & vbCrLf
    
    ' 测试 TCP 模式
    Debug.Print ">>> 开始 TCP 模式测试 <<<"
    TestModbusTCP
    Debug.Print vbCrLf
    
    ' 测试 RTU 模式（需要实际串口设备）
    ' Debug.Print ">>> 开始 RTU 模式测试 <<<"
    ' TestModbusRTU
    ' Debug.Print vbCrLf
    
    Debug.Print "========================================"
    Debug.Print "测试完成"
    Debug.Print "========================================"
End Sub

'=========================================================================
' 快速测试示例
'=========================================================================

Public Sub QuickTest()
    Dim mb As New cModbus
    
    On Error GoTo ErrorHandler
    
    ' 快速 TCP 测试
    mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"
    mb.TCPPort = 502
    
    mb.Connect
    
    ' 读取 10 个保持寄存器
    Dim regs() As Integer
    regs = mb.ReadHoldingRegisters(0, 10)
    
    ' 写入一个寄存器
    mb.WriteSingleRegister 0, 1234
    
    mb.Disconnect
    
    Debug.Print "快速测试完成"
    Exit Sub
ErrorHandler:
    Debug.Print "快速测试失败: " & Err.Description
    On Error Resume Next
    mb.Disconnect
End Sub
