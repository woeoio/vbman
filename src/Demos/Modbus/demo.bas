' TCP 模式示例
Dim mb As New cModbus
mb.ProtocolType = MB_PROTOCOL_TCP
mb.SlaveID = 1
mb.TCPHost = "192.168.1.100"
mb.TCPPort = 502
mb.ResponseTimeout = 2000

mb.Connect

' 读取保持寄存器
Dim regs() As Integer
regs = mb.ReadHoldingRegisters(0, 10)  ' 从地址0读取10个寄存器

' 写入单个寄存器
mb.WriteSingleRegister(0, 1234)

mb.Disconnect

' RTU 模式示例
Dim mbRTU As New cModbus
mbRTU.ProtocolType = MB_PROTOCOL_RTU
mbRTU.SlaveID = 1
mbRTU.SerialPort = "COM1"
mbRTU.BaudRate = 9600
mbRTU.DataBits = 8
mbRTU.Parity = "N"
mbRTU.StopBits = 1

mbRTU.Connect

' 读取线圈
Dim coils() As Boolean
coils = mbRTU.ReadCoils(0, 16)  ' 从地址0读取16个线圈

mbRTU.Disconnect