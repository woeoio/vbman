VERSION 5.00
Begin VB.Form FMaster
   Caption         =   "Modbus 主站示例 (Master)"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  '屏幕中央
   Begin VB.Frame fraConnection
      Caption         =   "连接设置"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtSlaveID
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Text            =   "1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtPort
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Text            =   "502"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtHost
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "127.0.0.1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdDisconnect
         Caption         =   "断开"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdConnect
         Caption         =   "连接"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblSlaveID
         Caption         =   "从站ID:"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblPort
         Caption         =   "端口:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblHost
         Caption         =   "主机:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblStatus
         Caption         =   "状态: 未连接"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraRead
      Caption         =   "读取操作"
      Height          =   1695
      Left            =   4800
      TabIndex        =   10
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtReadQuantity
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Text            =   "10"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtReadAddress
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdReadInputRegs
         Caption         =   "读输入寄存器"
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadHoldingRegs
         Caption         =   "读保持寄存器"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadDiscreteInputs
         Caption         =   "读离散输入"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadCoils
         Caption         =   "读线圈"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblReadQuantity
         Caption         =   "数量:"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblReadAddress
         Caption         =   "地址:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraWrite
      Caption         =   "写入操作"
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   4575
      Begin VB.TextBox txtWriteValue
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         Text            =   "1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtWriteAddress
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdWriteMultipleRegs
         Caption         =   "写多寄存器"
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdWriteSingleReg
         Caption         =   "写单寄存器"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdWriteMultipleCoils
         Caption         =   "写多线圈"
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdWriteSingleCoil
         Caption         =   "写单线圈"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblWriteValue
         Caption         =   "值:"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblWriteAddress
         Caption         =   "地址:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraResult
      Caption         =   "读取结果"
      Height          =   1695
      Left            =   4800
      TabIndex        =   28
      Top             =   1920
      Width           =   4695
      Begin VB.TextBox txtResult
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   29
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame fraLog
      Caption         =   "通信日志"
      Height          =   3375
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   9375
      Begin VB.CommandButton cmdClearLog
         Caption         =   "清空"
         Height          =   375
         Left            =   8400
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtLog
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   31
         Top             =   360
         Width           =   8175
      End
   End
End
Attribute VB_Name = "FMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' FMaster - Modbus 主站示例程序
'
' 功能: 演示 cModbusMaster 类的基本用法
'       - TCP 模式连接到从站
'       - 读取线圈、离散输入、保持寄存器、输入寄存器
'       - 写入单个/多个线圈和寄存器
'
' 使用方法:
'   1. 先启动从站程序 (SlaveDemo)
'   2. 输入从站地址和端口
'   3. 点击"连接"按钮
'   4. 进行读写操作
'
' 作者: Auto
' 日期: 2026-01-16
'
'=========================================================================
Option Explicit

' Modbus 主站对象
Private WithEvents m_Master As VBMANLIB.cModbusMaster
Attribute m_Master.VB_VarHelpID = -1

'=========================================================================
' 窗体事件
'=========================================================================

Private Sub Form_Load()
    ' 创建 Modbus 主站对象
    Set m_Master = New VBMANLIB.cModbusMaster

    ' 设置为 TCP 模式
    m_Master.ProtocolType = MB_MASTER_PROTOCOL_TCP

    LogMessage "Modbus 主站示例程序已启动"
    LogMessage "请先启动从站程序，然后点击连接"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    ' 断开连接
    If Not m_Master Is Nothing Then
        If m_Master.State = MB_MASTER_STATE_CONNECTED Then
            m_Master.DisConnect
        End If
        Set m_Master = Nothing
    End If
End Sub

'=========================================================================
' 连接管理
'=========================================================================

Private Sub cmdConnect_Click()
    On Error GoTo ErrorHandler

    ' 设置连接参数
    m_Master.TCPHost = txtHost.Text
    m_Master.TCPPort = CLng(Val(txtPort.Text))
    m_Master.SlaveID = CByte(Val(txtSlaveID.Text))
    m_Master.ResponseTimeout = 3000

    LogMessage "正在连接到 " & m_Master.TCPHost & ":" & m_Master.TCPPort & "..."

    ' 连接
    m_Master.Connect

    Exit Sub
ErrorHandler:
    LogMessage "连接失败: " & Err.Description
    MsgBox "连接失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next
    m_Master.DisConnect
    LogMessage "已断开连接"
End Sub

'=========================================================================
' 读取操作
'=========================================================================

Private Sub cmdReadCoils_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim lQuantity As Long
    Dim baCoils() As Boolean
    Dim i As Long
    Dim sResult As String

    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))

    LogMessage "读线圈: 地址=" & lAddress & ", 数量=" & lQuantity

    ' 读取线圈
    baCoils = m_Master.ReadCoils(lAddress, lQuantity)

    ' 显示结果
    sResult = "线圈 [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(baCoils)
        sResult = sResult & IIf(baCoils(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sResult = sResult & " "
    Next i

    txtResult.Text = sResult
    LogMessage "读取成功"
    Exit Sub

ErrorHandler:
    LogMessage "读线圈失败: " & Err.Description
End Sub

Private Sub cmdReadDiscreteInputs_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim lQuantity As Long
    Dim baInputs() As Boolean
    Dim i As Long
    Dim sResult As String

    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))

    LogMessage "读离散输入: 地址=" & lAddress & ", 数量=" & lQuantity

    ' 读取离散输入
    baInputs = m_Master.ReadDiscreteInputs(lAddress, lQuantity)

    ' 显示结果
    sResult = "离散输入 [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(baInputs)
        sResult = sResult & IIf(baInputs(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sResult = sResult & " "
    Next i

    txtResult.Text = sResult
    LogMessage "读取成功"
    Exit Sub

ErrorHandler:
    LogMessage "读离散输入失败: " & Err.Description
End Sub

Private Sub cmdReadHoldingRegs_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim lQuantity As Long
    Dim aiRegs() As Integer
    Dim i As Long
    Dim sResult As String

    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))

    LogMessage "读保持寄存器: 地址=" & lAddress & ", 数量=" & lQuantity

    ' 读取保持寄存器
    aiRegs = m_Master.ReadHoldingRegisters(lAddress, lQuantity)

    ' 显示结果
    sResult = "保持寄存器 [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(aiRegs)
        sResult = sResult & "[" & (lAddress + i) & "]=" & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
        If (i + 1) Mod 5 = 0 Then sResult = sResult & vbCrLf
    Next i

    txtResult.Text = sResult
    LogMessage "读取成功"
    Exit Sub

ErrorHandler:
    LogMessage "读保持寄存器失败: " & Err.Description
End Sub

Private Sub cmdReadInputRegs_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim lQuantity As Long
    Dim aiRegs() As Integer
    Dim i As Long
    Dim sResult As String

    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))

    LogMessage "读输入寄存器: 地址=" & lAddress & ", 数量=" & lQuantity

    ' 读取输入寄存器
    aiRegs = m_Master.ReadInputRegisters(lAddress, lQuantity)

    ' 显示结果
    sResult = "输入寄存器 [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(aiRegs)
        sResult = sResult & "[" & (lAddress + i) & "]=" & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
        If (i + 1) Mod 5 = 0 Then sResult = sResult & vbCrLf
    Next i

    txtResult.Text = sResult
    LogMessage "读取成功"
    Exit Sub

ErrorHandler:
    LogMessage "读输入寄存器失败: " & Err.Description
End Sub

'=========================================================================
' 写入操作
'=========================================================================

Private Sub cmdWriteSingleCoil_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim bValue As Boolean
    Dim bResult As Boolean

    lAddress = CLng(Val(txtWriteAddress.Text))
    bValue = (Val(txtWriteValue.Text) <> 0)

    LogMessage "写单线圈: 地址=" & lAddress & ", 值=" & IIf(bValue, "ON", "OFF")

    ' 写入单个线圈
    bResult = m_Master.WriteSingleCoil(lAddress, bValue)

    If bResult Then
        LogMessage "写入成功"
        txtResult.Text = "写单线圈成功" & vbCrLf & "地址: " & lAddress & vbCrLf & "值: " & IIf(bValue, "ON (1)", "OFF (0)")
    Else
        LogMessage "写入失败"
    End If

    Exit Sub
ErrorHandler:
    LogMessage "写单线圈失败: " & Err.Description
End Sub

Private Sub cmdWriteMultipleCoils_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim lQuantity As Long
    Dim baValues() As Boolean
    Dim i As Long
    Dim bResult As Boolean

    lAddress = CLng(Val(txtWriteAddress.Text))
    lQuantity = CLng(Val(txtWriteValue.Text))

    If lQuantity < 1 Or lQuantity > 100 Then
        MsgBox "数量应在 1-100 之间", vbExclamation
        Exit Sub
    End If

    ' 生成交替模式的值
    ReDim baValues(lQuantity - 1) As Boolean
    For i = 0 To lQuantity - 1
        baValues(i) = ((i Mod 2) = 0)
    Next i

    LogMessage "写多线圈: 地址=" & lAddress & ", 数量=" & lQuantity

    ' 写入多个线圈
    bResult = m_Master.WriteMultipleCoils(lAddress, baValues)

    If bResult Then
        LogMessage "写入成功 (交替模式: 1,0,1,0...)"
        txtResult.Text = "写多线圈成功" & vbCrLf & "地址: " & lAddress & vbCrLf & "数量: " & lQuantity & vbCrLf & "模式: 交替 (1,0,1,0...)"
    Else
        LogMessage "写入失败"
    End If

    Exit Sub
ErrorHandler:
    LogMessage "写多线圈失败: " & Err.Description
End Sub

Private Sub cmdWriteSingleReg_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim iValue As Integer
    Dim bResult As Boolean

    lAddress = CLng(Val(txtWriteAddress.Text))
    iValue = CInt(Val(txtWriteValue.Text))

    LogMessage "写单寄存器: 地址=" & lAddress & ", 值=" & iValue

    ' 写入单个寄存器
    bResult = m_Master.WriteSingleRegister(lAddress, iValue)

    If bResult Then
        LogMessage "写入成功"
        txtResult.Text = "写单寄存器成功" & vbCrLf & "地址: " & lAddress & vbCrLf & "值: " & iValue
    Else
        LogMessage "写入失败"
    End If

    Exit Sub
ErrorHandler:
    LogMessage "写单寄存器失败: " & Err.Description
End Sub

Private Sub cmdWriteMultipleRegs_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim lQuantity As Long
    Dim aiValues() As Integer
    Dim i As Long
    Dim bResult As Boolean

    lAddress = CLng(Val(txtWriteAddress.Text))
    lQuantity = CLng(Val(txtWriteValue.Text))

    If lQuantity < 1 Or lQuantity > 100 Then
        MsgBox "数量应在 1-100 之间", vbExclamation
        Exit Sub
    End If

    ' 生成递增值
    ReDim aiValues(lQuantity - 1) As Integer
    For i = 0 To lQuantity - 1
        aiValues(i) = 100 + i
    Next i

    LogMessage "写多寄存器: 地址=" & lAddress & ", 数量=" & lQuantity

    ' 写入多个寄存器
    bResult = m_Master.WriteMultipleRegisters(lAddress, aiValues)

    If bResult Then
        LogMessage "写入成功 (值: 100, 101, 102...)"
        txtResult.Text = "写多寄存器成功" & vbCrLf & "地址: " & lAddress & vbCrLf & "数量: " & lQuantity & vbCrLf & "值: 100, 101, 102..."
    Else
        LogMessage "写入失败"
    End If

    Exit Sub
ErrorHandler:
    LogMessage "写多寄存器失败: " & Err.Description
End Sub

'=========================================================================
' 事件处理
'=========================================================================

Private Sub m_Master_OnConnect()
    LogMessage "连接成功"
    lblStatus.Caption = "状态: 已连接"
    lblStatus.ForeColor = &HC000&
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
End Sub

Private Sub m_Master_OnDisconnect()
    LogMessage "连接已断开"
    lblStatus.Caption = "状态: 未连接"
    lblStatus.ForeColor = &HFF&
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
End Sub

Private Sub m_Master_OnError(ByVal Description As String)
    LogMessage "错误: " & Description
End Sub

Private Sub m_Master_OnDataReceived(Data() As Byte)
    ' 可选: 显示原始数据
    ' Dim sHex As String
    ' Dim i As Long
    ' sHex = "RX: "
    ' For i = 0 To UBound(Data)
    '     sHex = sHex & Right$("0" & Hex$(Data(i)), 2) & " "
    ' Next i
    ' LogMessage sHex
End Sub

'=========================================================================
' 辅助函数
'=========================================================================

Private Sub cmdClearLog_Click()
    txtLog.Text = ""
End Sub

Private Sub LogMessage(ByVal Message As String)
    Dim sTime As String
    sTime = Format$(Now, "hh:mm:ss")
    txtLog.Text = txtLog.Text & "[" & sTime & "] " & Message & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
