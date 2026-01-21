VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNTPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cNTPServer - NTP 服务端
'
' Purpose: 提供标准 NTP 时间服务器
'
' Features:
'   - 标准 NTP 协议实现 (RFC 5905)
'   - 支持多客户端并发
'   - 可配置 Stratum 层级
'   - 事件驱动架构
'
' Usage:
'   Dim oServer As New cNTPServer
'
'   oServer.Start 123
'   ' Server is now running...
'
'   oServer.Stop
'
'=========================================================================

Option Explicit
    DefObj A-Z
    Private Const MODULE_NAME As String = "cNTPServer"

'=========================================================================
' 私有成员
'=========================================================================

Private WithEvents m_oSocket As cWinsock
Attribute m_oSocket.VB_VarHelpID = -1
Private m_lPort As Long
Private m_bRunning As Boolean
Private m_bStratum As Byte
Private m_sReferenceID As String
Private m_dRootDelay As Double
Private m_dRootDispersion As Double
Private m_lRequestCount As Long
Private m_oTimeProvider As cNTPTimeProvider

'=========================================================================
' 公共事件
'=========================================================================

' 服务器启动事件
Public Event ServerStarted(ByVal Port As Long)

' 服务器停止事件
Public Event ServerStopped()

' 接收到请求事件
Public Event RequestReceived(ByVal RemoteIP As String, ByVal RemotePort As Long)

' 发送响应事件
Public Event ResponseSent(ByVal RemoteIP As String, ByVal RemotePort As Long)

' 错误事件
Public Event ServerError(ByVal Number As Long, ByVal Description As String)

'=========================================================================
' 属性
'=========================================================================

' 服务器端口
Public Property Get Port() As Long
    Port = m_lPort
End Property

' 服务器是否运行中
Public Property Get IsRunning() As Boolean
    IsRunning = m_bRunning
End Property

' Stratum 层级 (1-15)
Public Property Get Stratum() As Byte
    Stratum = m_bStratum
End Property

Public Property Let Stratum(ByVal Value As Byte)
    If Value = 0 Or Value > 15 Then
        Err.Raise vbObjectError, MODULE_NAME, "Stratum must be 1-15"
    End If
    m_bStratum = Value
End Property

' Reference ID
Public Property Get ReferenceID() As String
    ReferenceID = m_sReferenceID
End Property

Public Property Let ReferenceID(ByVal Value As String)
    If Len(Value) > 4 Then
        m_sReferenceID = Left(Value, 4)
    Else
        m_sReferenceID = Value
    End If
End Property

' Root Delay (秒)
Public Property Get RootDelay() As Double
    RootDelay = m_dRootDelay
End Property

Public Property Let RootDelay(ByVal Value As Double)
    m_dRootDelay = Value
End Property

' Root Dispersion (秒)
Public Property Get RootDispersion() As Double
    RootDispersion = m_dRootDispersion
End Property

Public Property Let RootDispersion(ByVal Value As Double)
    m_dRootDispersion = Value
End Property

' 请求计数器
Public Property Get RequestCount() As Long
    RequestCount = m_lRequestCount
End Property

' 时间提供者对象 (可外部设置自定义时间源)
Public Property Get TimeProvider() As cNTPTimeProvider
    Set TimeProvider = m_oTimeProvider
End Property

Public Property Set TimeProvider(ByVal Value As cNTPTimeProvider)
    Set m_oTimeProvider = Value
End Property

' Socket 对象 (用于高级用户)
Public Property Get Socket() As cWinsock
    Set Socket = m_oSocket
End Property

'=========================================================================
' 公共方法
'=========================================================================

' 启动 NTP 服务器
Public Sub Start(Optional ByVal Port As Long = 123)
    If m_bRunning Then
        Err.Raise vbObjectError, MODULE_NAME, "Server is already running"
    End If
    
    m_lPort = Port
    
    ' 创建 Socket
    Set m_oSocket = New cWinsock
    m_oSocket.Protocol = wskUDPProtocol
    m_oSocket.LocalPort = m_lPort
    
    ' 绑定端口
    m_oSocket.Bind
    
    m_bRunning = True
    m_lRequestCount = 0
    
    RaiseEvent ServerStarted(m_lPort)
End Sub

' 停止 NTP 服务器
Public Sub Stop()
    If Not m_bRunning Then
        Exit Sub
    End If
    
    m_bRunning = False
    
    On Error Resume Next
    m_oSocket.CloseSck
    Set m_oSocket = Nothing
    On Error GoTo 0
    
    RaiseEvent ServerStopped
End Sub

' 重置请求计数器
Public Sub ResetRequestCount()
    m_lRequestCount = 0
End Sub

'=========================================================================
' 私有方法
'=========================================================================

' 处理 NTP 请求
Private Sub ProcessRequest(ByRef baData() As Byte, ByVal RemoteIP As String, ByVal RemotePort As Long)
    Dim oRequestPacket As cNTPPacket
    Dim oResponsePacket As cNTPPacket
    Dim baResponse() As Byte
    Dim dtNow As Date
    Dim oSocket As cWinsock
    
    On Error GoTo ErrorHandler
    
    ' 解析请求包
    Set oRequestPacket = New cNTPPacket
    If Not oRequestPacket.FromBytes(baData) Then
        Exit Sub
    End If
    
    ' 检查是否为客户端请求
    If oRequestPacket.Mode <> NTPMode_Client Then
        Exit Sub
    End If
    
    ' 触发请求接收事件
    m_lRequestCount = m_lRequestCount + 1
    RaiseEvent RequestReceived(RemoteIP, RemotePort)
    
    ' 获取当前时间
    If m_oTimeProvider Is Nothing Then
        dtNow = Now
    Else
        dtNow = m_oTimeProvider.GetCurrentTime()
    End If
    
    ' 创建响应包
    Set oResponsePacket = New cNTPPacket
    
    ' 设置响应包字段
    oResponsePacket.LI = 0
    oResponsePacket.VN = oRequestPacket.VN
    oResponsePacket.Mode = NTPMode_Server
    oResponsePacket.Stratum = m_bStratum
    oResponsePacket.Poll = oRequestPacket.Poll
    oResponsePacket.Precision = &HE8  ' 默认精度
    oResponsePacket.RootDelay = m_dRootDelay
    oResponsePacket.RootDispersion = m_dRootDispersion
    oResponsePacket.ReferenceID = m_sReferenceID
    
    ' 设置时间戳
    oResponsePacket.ReferenceTimestamp = dtNow
    oResponsePacket.OriginTimestamp = oRequestPacket.TransmitTimestamp
    oResponsePacket.ReceiveTimestamp = dtNow
    oResponsePacket.TransmitTimestamp = dtNow
    
    ' 编码并发送响应
    baResponse = oResponsePacket.ToBytes()
    
    ' 创建临时 Socket 发送响应
    Set oSocket = New cWinsock
    oSocket.Protocol = wskUDPProtocol
    oSocket.LocalPort = 0
    oSocket.RemoteHost = RemoteIP
    oSocket.RemotePort = RemotePort
    oSocket.Bind
    oSocket.SendData baResponse
    
    ' 清理临时 Socket
    On Error Resume Next
    oSocket.CloseSck
    Set oSocket = Nothing
    On Error GoTo 0
    
    ' 触发响应发送事件
    RaiseEvent ResponseSent(RemoteIP, RemotePort)
    
    Exit Sub
    
ErrorHandler:
    RaiseEvent ServerError(Err.Number, Err.Description)
End Sub

'=========================================================================
' Socket 事件处理
'=========================================================================

Private Sub m_oSocket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    Dim RemoteIP As String
    Dim RemotePort As Long
    
    On Error GoTo ErrorHandler
    
    ' 接收数据
    ReDim baData(bytesTotal - 1)
    Client.GetData baData
    
    ' 获取远程地址和端口
    RemoteIP = Client.RemoteHostIP
    RemotePort = Client.RemotePort
    
    ' 处理请求
    ProcessRequest baData, RemoteIP, RemotePort
    
    Exit Sub
    
ErrorHandler:
    RaiseEvent ServerError(Err.Number, Err.Description)
End Sub

Private Sub m_oSocket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    RaiseEvent ServerError(Number, Description)
End Sub

'=========================================================================
' 类初始化/终止
'=========================================================================

Private Sub Class_Initialize()
    m_lPort = 123
    m_bRunning = False
    m_bStratum = 2           ' 默认为二级服务器
    m_sReferenceID = "LOCL"  ' 本地时间源
    m_dRootDelay = 0
    m_dRootDispersion = 0.01 ' 默认10ms
    m_lRequestCount = 0
    
    ' 创建默认时间提供者
    Set m_oTimeProvider = New cNTPTimeProvider
End Sub

Private Sub Class_Terminate()
    Stop
End Sub
