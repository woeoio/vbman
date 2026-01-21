VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNTPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'=========================================================================
'
' cNTPClient - NTP 客户端
'
' Purpose: 提供简单易用的 NTP 时间同步客户端
'
' Features:
'   - 同步/异步时间查询
'   - 自动超时处理
'   - 支持多服务器查询
'   - 提供时间偏移和往返延迟计算
'
' Usage:
'   Dim oClient As New cNTPClient
'   Dim dtServerTime As Date
'   Dim dOffset As Double
'
'   dtServerTime = oClient.GetServerTime("pool.ntp.org", 123)
'   dOffset = oClient.GetTimeOffset("pool.ntp.org", 123)
'
'=========================================================================

Option Explicit
    DefObj A-Z
    Private Const MODULE_NAME As String = "cNTPClient"

'=========================================================================
' 私有成员
'=========================================================================

Private WithEvents m_oSocket As cWinsock
Attribute m_oSocket.VB_VarHelpID = -1
Private m_sServerHost As String
Private m_lServerPort As Long
Private m_dTimeout As Double
Private m_dSendTime As Double
Private m_baResponseData() As Byte
Private m_oResponsePacket As cNTPPacket
Private m_bAsyncMode As Boolean
Private m_vAsyncTag As Variant

'=========================================================================
' 公共事件
'=========================================================================

' 时间同步完成事件 (异步模式)
Public Event TimeSyncComplete(ByVal ServerHost As String, ByVal ServerTime As Date, _
                              ByVal Offset As Double, ByVal RoundTripDelay As Double, _
                              ByVal Success As Boolean, ByVal ErrorMsg As String)

'=========================================================================
' 属性
'=========================================================================

' 服务器地址
Public Property Get ServerHost() As String
    ServerHost = m_sServerHost
End Property

Public Property Let ServerHost(ByVal Value As String)
    If LenB(Value) = 0 Then
        Err.Raise vbObjectError, MODULE_NAME, "ServerHost cannot be empty"
    End If
    m_sServerHost = Value
End Property

' 服务器端口
Public Property Get ServerPort() As Long
    ServerPort = m_lServerPort
End Property

Public Property Let ServerPort(ByVal Value As Long)
    If Value < 1 Or Value > 65535 Then
        Err.Raise vbObjectError, MODULE_NAME, "Invalid port number"
    End If
    m_lServerPort = Value
End Property

' 超时时间(秒)
Public Property Get Timeout() As Double
    Timeout = m_dTimeout
End Property

Public Property Let Timeout(ByVal Value As Double)
    If Value < 0 Then
        Err.Raise vbObjectError, MODULE_NAME, "Timeout cannot be negative"
    End If
    m_dTimeout = Value
End Property

' 是否异步模式
Public Property Get AsyncMode() As Boolean
    AsyncMode = m_bAsyncMode
End Property

' Socket 对象 (用于高级用户)
Public Property Get Socket() As cWinsock
    Set Socket = m_oSocket
End Property

'=========================================================================
' 公共方法
'=========================================================================

' 获取服务器时间 (同步模式)
Public Function GetServerTime(Optional ByVal Server As String = "", _
                              Optional ByVal Port As Long = 123) As Date
    Dim oPacket As cNTPPacket
    Dim baData() As Byte
    Dim dtTime As Date
    
    ' 设置参数
    If LenB(Server) <> 0 Then
        m_sServerHost = Server
    End If
    If Port <> 123 Then
        m_lServerPort = Port
    End If
    
    ' 初始化 Socket
    InitializeSocket
    
    ' 创建并发送请求
    Set oPacket = New cNTPPacket
    Set oPacket = oPacket.CreateClientRequest()
    baData = oPacket.ToBytes()
    
    ' 记录发送时间
    m_dSendTime = Timer
    
    ' 发送并等待响应
    If SendRequestAndWait(baData) Then
        Set m_oResponsePacket = New cNTPPacket
        If m_oResponsePacket.FromBytes(m_baResponseData) Then
            dtTime = m_oResponsePacket.TransmitTimestamp
            GetServerTime = dtTime
        Else
            Err.Raise vbObjectError, MODULE_NAME, "Failed to parse response packet"
        End If
    Else
        Err.Raise vbObjectError, MODULE_NAME, "Timeout or error receiving response"
    End If
    
    ' 清理
    CleanupSocket
End Function

' 获取时间偏移量 (同步模式)
' 返回值: 正数表示本地时间比服务器快，负数表示本地时间比服务器慢
Public Function GetTimeOffset(Optional ByVal Server As String = "", _
                              Optional ByVal Port As Long = 123) As Double
    Dim oPacket As cNTPPacket
    Dim baData() As Byte
    Dim dT1 As Double      ' 发送时间
    Dim dT2 As Double      ' 接收时间
    Dim dT3 As Double      ' 源时间戳
    Dim dT4 As Double      ' 目标时间戳
    Dim dOffset As Double
    Dim dRoundTrip As Double
    
    ' 设置参数
    If LenB(Server) <> 0 Then
        m_sServerHost = Server
    End If
    If Port <> 123 Then
        m_lServerPort = Port
    End If
    
    ' 初始化 Socket
    InitializeSocket
    
    ' 创建并发送请求
    Set oPacket = New cNTPPacket
    Set oPacket = oPacket.CreateClientRequest()
    
    ' 记录 T1: 本地发送时间
    dT1 = DateToNtpTime(Now)
    oPacket.TransmitTimestamp = Now
    
    baData = oPacket.ToBytes()
    m_dSendTime = Timer
    
    ' 发送并等待响应
    If SendRequestAndWait(baData) Then
        Set m_oResponsePacket = New cNTPPacket
        If m_oResponsePacket.FromBytes(m_baResponseData) Then
            
            ' T3: 服务器接收时间
            dT3 = DateToNtpTime(m_oResponsePacket.ReceiveTimestamp)
            
            ' T4: 服务器发送时间
            dT4 = DateToNtpTime(m_oResponsePacket.TransmitTimestamp)
            
            ' T2: 本地接收时间
            dT2 = DateToNtpTime(Now)
            
            ' 计算往返延迟
            dRoundTrip = (dT2 - dT1) - (dT4 - dT3)
            
            ' 计算时间偏移
            dOffset = ((dT2 - dT1) + (dT3 - dT4)) / 2
            
            GetTimeOffset = dOffset
        Else
            Err.Raise vbObjectError, MODULE_NAME, "Failed to parse response packet"
        End If
    Else
        Err.Raise vbObjectError, MODULE_NAME, "Timeout or error receiving response"
    End If
    
    ' 清理
    CleanupSocket
End Function

' 异步请求服务器时间
Public Sub RequestTimeAsync(Optional ByVal Server As String = "", _
                            Optional ByVal Port As Long = 123, _
                            Optional ByVal Tag As Variant)
    Dim oPacket As cNTPPacket
    Dim baData() As Byte
    
    ' 设置参数
    If LenB(Server) <> 0 Then
        m_sServerHost = Server
    End If
    If Port <> 123 Then
        m_lServerPort = Port
    End If
    
    ' 保存 Tag
    m_vAsyncTag = Tag
    m_bAsyncMode = True
    
    ' 初始化 Socket
    InitializeSocket
    
    ' 创建并发送请求
    Set oPacket = New cNTPPacket
    Set oPacket = oPacket.CreateClientRequest()
    baData = oPacket.ToBytes()
    
    ' 记录发送时间
    m_dSendTime = Timer
    
    ' 发送请求
    m_oSocket.SendData baData
End Sub

' 同步时间到本地系统 (需要管理员权限)
Public Function SyncSystemTime(Optional ByVal Server As String = "", _
                               Optional ByVal Port As Long = 123) As Boolean
    Dim dtServerTime As Date
    Dim dOffset As Double
    
    On Error GoTo ErrorHandler
    
    ' 获取服务器时间
    dtServerTime = GetServerTime(Server, Port)
    
    ' 设置系统时间 (需要管理员权限)
    SetSystemTime dtServerTime
    
    SyncSystemTime = True
    Exit Function
    
ErrorHandler:
    SyncSystemTime = False
End Function

'=========================================================================
' 私有方法
'=========================================================================

' 初始化 Socket
Private Sub InitializeSocket()
    Set m_oSocket = New cWinsock
    m_oSocket.Protocol = wskUDPProtocol
    m_oSocket.LocalPort = 0
    m_oSocket.RemoteHost = m_sServerHost
    m_oSocket.RemotePort = m_lServerPort
    m_oSocket.Bind
End Sub

' 清理 Socket
Private Sub CleanupSocket()
    On Error Resume Next
    If Not m_oSocket Is Nothing Then
        m_oSocket.CloseSck
        Set m_oSocket = Nothing
    End If
    On Error GoTo 0
End Sub

' 发送请求并等待响应
Private Function SendRequestAndWait(ByRef baData() As Byte) As Boolean
    Dim lStartTime As Double
    Dim lTimeout As Long
    Dim bReceived As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 清空接收缓冲区
    Erase m_baResponseData
    bReceived = False
    lStartTime = Timer
    lTimeout = CLng(m_dTimeout * 1000)
    
    ' 发送数据
    m_oSocket.SendData baData
    
    ' 等待响应 (循环检查)
    Do While Not bReceived And (Timer - lStartTime) * 1000 < lTimeout
        DoEvents
        ' 检查是否有数据到达
        If m_oSocket.BytesReceived > 0 Then
            ReDim m_baResponseData(m_oSocket.BytesReceived - 1)
            m_oSocket.GetData m_baResponseData
            If UBound(m_baResponseData) + 1 >= 48 Then
                bReceived = True
            End If
        End If
    Loop
    
    SendRequestAndWait = bReceived
    Exit Function
    
ErrorHandler:
    SendRequestAndWait = False
End Function

' 设置系统时间 (Windows API)
Private Sub SetSystemTime(ByVal dtTime As Date)
    Dim st As SYSTEMTIME
    
    With st
        .wYear = Year(dtTime)
        .wMonth = Month(dtTime)
        .wDayOfWeek = Weekday(dtTime) - 1
        .wDay = Day(dtTime)
        .wHour = Hour(dtTime)
        .wMinute = Minute(dtTime)
        .wSecond = Second(dtTime)
        .wMilliseconds = 0
    End With
    
    If SetLocalSystemTime(st) = 0 Then
        Err.Raise vbObjectError, MODULE_NAME, "Failed to set system time (Admin rights required)"
    End If
End Sub

' 将 Date 转换为 NTP 时间戳
Private Function DateToNtpTime(ByVal Dt As Date) As Double
    Dim dSeconds As Double
    dSeconds = DateDiff("s", #1/1/1970#, Dt)
    DateToNtpTime = dSeconds + 2208988800#
End Function

'=========================================================================
' Socket 事件处理
'=========================================================================

Private Sub m_oSocket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    If m_bAsyncMode Then
        ' 异步模式处理
        ReDim m_baResponseData(bytesTotal - 1)
        Client.GetData m_baResponseData
        
        ' 解析响应
        Dim dOffset As Double
        Dim dRoundTrip As Double
        Dim dtServerTime As Date
        Dim bSuccess As Boolean
        Dim sErrorMsg As String
        
        On Error Resume Next
        
        Set m_oResponsePacket = New cNTPPacket
        If m_oResponsePacket.FromBytes(m_baResponseData) Then
            dtServerTime = m_oResponsePacket.TransmitTimestamp
            
            ' 计算偏移
            Dim dT1 As Double, dT2 As Double, dT3 As Double, dT4 As Double
            dT1 = DateToNtpTime(DateAdd("s", -CDbl(Timer - m_dSendTime), Now))
            dT2 = DateToNtpTime(Now)
            dT3 = DateToNtpTime(m_oResponsePacket.ReceiveTimestamp)
            dT4 = DateToNtpTime(m_oResponsePacket.TransmitTimestamp)
            
            dRoundTrip = (dT2 - dT1) - (dT4 - dT3)
            dOffset = ((dT2 - dT1) + (dT3 - dT4)) / 2
            
            bSuccess = True
        Else
            sErrorMsg = "Failed to parse response packet"
            bSuccess = False
        End If
        
        On Error GoTo 0
        
        ' 触发完成事件
        RaiseEvent TimeSyncComplete(m_sServerHost, dtServerTime, dOffset, dRoundTrip, bSuccess, sErrorMsg)
        
        ' 清理
        CleanupSocket
        m_bAsyncMode = False
    End If
End Sub

Private Sub m_oSocket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    If m_bAsyncMode Then
        RaiseEvent TimeSyncComplete(m_sServerHost, DateSerial(2000, 1, 1), 0, 0, False, Description)
        CleanupSocket
        m_bAsyncMode = False
    End If
End Sub

'=========================================================================
' Windows API 声明
'=========================================================================

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function SetLocalSystemTime Lib "kernel32" Alias "SetLocalTime" (lpSystemTime As SYSTEMTIME) As Long

'=========================================================================
' 类初始化/终止
'=========================================================================

Private Sub Class_Initialize()
    m_lServerPort = 123
    m_dTimeout = 5      ' 默认5秒超时
    m_bAsyncMode = False
End Sub

Private Sub Class_Terminate()
    CleanupSocket
End Sub
