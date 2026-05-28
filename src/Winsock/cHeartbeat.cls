VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHeartbeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cHeartbeat - TCP Heartbeat Manager for cWinsock
' by woeoio@qq.com
'
' Purpose: Provides automatic heartbeat keep-alive for TCP connections.
'          Server-side: timeout detection, auto-disconnect idle clients.
'          Client-side: periodic heartbeat sending with smart skip.
'          Internally uses cTimer for automatic timing - no external
'          polling required.
'
' Usage: This class is used internally by cWinsock.
'        Users configure heartbeat through cWinsock properties.
'
'=========================================================================
Option Explicit
    DefObj A-Z
Private Const MODULE_NAME As String = "cHeartbeat"

Private Declare Function GetTickCount Lib "kernel32" () As Long

'--- Constants ---
Private Const DEF_SERVER_TIMEOUT As Long = 120    ' 2 minutes default timeout
Private Const DEF_CLIENT_INTERVAL As Long = 50     ' 50 seconds default interval
Private Const DEF_CHECK_INTERVAL_MS As Long = 10000 ' 10 seconds check interval (ms)

'--- Events ---
' Fires every tick (10s by default) - cWinsock handles business logic
Public Event Tick()

'--- Member Variables ---
Private m_bEnabled As Boolean
Private m_lServerTimeout As Long        ' Seconds - server-side idle timeout
Private m_lClientInterval As Long       ' Seconds - client-side heartbeat interval
Private m_lLastActivityTick As Long     ' GetTickCount of last activity
Private m_lLastHeartbeatTick As Long    ' GetTickCount of last heartbeat sent
Private m_lHeartbeatCount As Long       ' Number of heartbeats sent
Private m_baData() As Byte              ' Heartbeat packet content

' Internal cTimer for automatic timing
Private WithEvents m_oTimer As cTimer
Attribute m_oTimer.VB_VarHelpID = -1

'=========================================================================
' Properties
'=========================================================================

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    If Value = m_bEnabled Then Exit Property
    m_bEnabled = Value
    If Value Then
        ResetActivity
        StartTimer
    Else
        StopTimer
    End If
End Property

' Server: idle timeout in seconds (0 = disabled)
Public Property Get ServerTimeout() As Long
    ServerTimeout = m_lServerTimeout
End Property

Public Property Let ServerTimeout(ByVal Value As Long)
    m_lServerTimeout = Value
End Property

' Client: heartbeat interval in seconds (0 = disabled)
Public Property Get ClientInterval() As Long
    ClientInterval = m_lClientInterval
End Property

Public Property Let ClientInterval(ByVal Value As Long)
    m_lClientInterval = Value
End Property

' Heartbeat packet content (byte array)
Public Property Get HeartbeatData() As Byte()
    HeartbeatData = m_baData
End Property

Public Property Let HeartbeatData(ByRef Value() As Byte)
    m_baData = Value
End Property

' Last activity timestamp (GetTickCount)
Public Property Get LastActivityTick() As Long
    LastActivityTick = m_lLastActivityTick
End Property

' Last heartbeat timestamp (GetTickCount)
Public Property Get LastHeartbeatTick() As Long
    LastHeartbeatTick = m_lLastHeartbeatTick
End Property

' Heartbeat count
Public Property Get HeartbeatCount() As Long
    HeartbeatCount = m_lHeartbeatCount
End Property

' Seconds since last activity
Public Property Get IdleSeconds() As Long
    IdleSeconds = (GetTickCount() - m_lLastActivityTick) \ 1000
End Property

' Is connection alive (not timed out)
Public Property Get IsAlive() As Boolean
    If m_lServerTimeout <= 0 Then
        IsAlive = True
    Else
        IsAlive = (IdleSeconds < m_lServerTimeout)
    End If
End Property

' Timer check interval in milliseconds (default 10000 = 10 seconds)
Public Property Get TimerInterval() As Long
    If m_oTimer Is Nothing Then
        TimerInterval = DEF_CHECK_INTERVAL_MS
    Else
        TimerInterval = m_oTimer.Interval
    End If
End Property

Public Property Let TimerInterval(ByVal Value As Long)
    If Value < 1000 Then Value = 1000  ' Minimum 1 second
    If Not m_oTimer Is Nothing Then
        m_oTimer.Interval = Value
    End If
End Property

'=========================================================================
' Public Methods
'=========================================================================

' Reset activity timestamp (call on any data send/receive)
Public Sub ResetActivity()
    m_lLastActivityTick = GetTickCount()
End Sub

' Check all clients for timeout (called by server's tick handler)
' Returns Collection of timed-out cWinsock clients
Public Function CheckClientTimeouts(ByVal oServer As Object) As Collection
    Dim colResult As New Collection
    Dim oClient As Object
    Dim lNow As Long
    
    On Error GoTo EH
    
    lNow = GetTickCount()
    
    If m_lServerTimeout <= 0 Then
        Set CheckClientTimeouts = Nothing
        Exit Function
    End If
    
    ' Iterate server's Clients collection
    For Each oClient In oServer.Clients
        If Not oClient.Heartbeat Is Nothing Then
            If oClient.Heartbeat.IdleSeconds >= m_lServerTimeout Then
                colResult.Add oClient
            End If
        End If
    Next
    
    If colResult.count > 0 Then
        Set CheckClientTimeouts = colResult
    Else
        Set CheckClientTimeouts = Nothing
    End If
    
    Exit Function
EH:
    Set CheckClientTimeouts = Nothing
End Function

' Check if client should send heartbeat now
Public Function ShouldSendHeartbeat() As Boolean
    Dim lNow As Long
    Dim lElapsed As Long
    
    ShouldSendHeartbeat = False
    
    If Not m_bEnabled Then Exit Function
    If m_lClientInterval <= 0 Then Exit Function
    
    lNow = GetTickCount()
    lElapsed = (lNow - m_lLastActivityTick) \ 1000
    
    ' Smart skip: if there was recent data activity, skip this heartbeat cycle
    If lElapsed < m_lClientInterval Then Exit Function
    
    ' Enough time has passed since last activity, send heartbeat
    ShouldSendHeartbeat = True
End Function

' Record that a heartbeat was sent
Public Sub RecordHeartbeatSent()
    m_lLastHeartbeatTick = GetTickCount()
    m_lHeartbeatCount = m_lHeartbeatCount + 1
End Sub

' Stop heartbeat
Public Sub StopMe()
    m_bEnabled = False
    StopTimer
End Sub

'=========================================================================
' Internal Timer (cTimer-based, fully automatic)
'=========================================================================

Private Sub StartTimer()
    If m_oTimer Is Nothing Then
        Set m_oTimer = New cTimer
    End If
    m_oTimer.Interval = DEF_CHECK_INTERVAL_MS
    m_oTimer.Enabled = True
End Sub

Private Sub StopTimer()
    If Not m_oTimer Is Nothing Then
        m_oTimer.Enabled = False
    End If
End Sub

' cTimer fires Tick event every 10 seconds
Private Sub m_oTimer_Timer()
    If m_bEnabled Then
        RaiseEvent Tick
    End If
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_lServerTimeout = DEF_SERVER_TIMEOUT
    m_lClientInterval = DEF_CLIENT_INTERVAL
    m_lLastActivityTick = GetTickCount()
    m_lLastHeartbeatTick = 0
    m_lHeartbeatCount = 0
    m_bEnabled = False
    m_baData = vbNullString
End Sub

Private Sub Class_Terminate()
    Stop
    Set m_oTimer = Nothing
End Sub
