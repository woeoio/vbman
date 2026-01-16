Attribute VB_Name = "demo"
'=========================================================================
'
' Modbus Demo - 快速示例代码
'
' Purpose: 提供 Modbus RTU 和 TCP 模式的简单使用示例
'
' Author: Auto
' Date: 2026-01-21
'
'=========================================================================
Option Explicit

'=========================================================================
' TCP 模式示例
'=========================================================================

Public Sub DemoTCP()
    Dim mb As New cModbus
    
    ' 配置 TCP 连接
    mb.ProtocolType = MB_MASTER_PROTOCOL_TCP
    mb.SlaveID = 1
    mb.TCPHost = "192.168.1.100"
    mb.TCPPort = 502
    mb.ResponseTimeout = 2000
    
    ' 连接
    mb.Connect
    
    ' 读取保持寄存器
    Dim regs() As Integer
    regs = mb.ReadHoldingRegisters(0, 10)                                       ' 从地址0读取10个寄存器
    
    ' 写入单个寄存器
    mb.WriteSingleRegister 0, 1234
    
    ' 断开连接
    mb.Disconnect
End Sub

'=========================================================================
' RTU 模式示例
'=========================================================================

Public Sub DemoRTU()
    Dim mbRTU As New cModbus
    
    ' 配置 RTU 连接
    mbRTU.ProtocolType = MB_MASTER_PROTOCOL_RTU
    mbRTU.SlaveID = 1
    mbRTU.SerialPort = "COM1"
    mbRTU.BaudRate = 9600
    mbRTU.DataBits = 8
    mbRTU.Parity = "N"
    mbRTU.StopBits = 1
    
    ' 连接
    mbRTU.Connect
    
    ' 读取线圈
    Dim coils() As Boolean
    coils = mbRTU.ReadCoils(0, 16)  ' 从地址0读取16个线圈
    
    ' 断开连接
    mbRTU.Disconnect
End Sub
