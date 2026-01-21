VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNTPTimeProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cNTPTimeProvider - NTP 时间提供者
'
' Purpose: 抽象时间源接口，支持多种时间源
'
' Design Pattern: Strategy Pattern
'   - 可以使用系统时间、硬件时钟、外部 NTP 服务器等
'   - 便于单元测试和自定义时间源
'
' Usage:
'   Dim oProvider As New cNTPTimeProvider
'   Dim dtTime As Date
'
'   dtTime = oProvider.GetCurrentTime()
'
'=========================================================================

Option Explicit
    DefObj A-Z
    Private Const MODULE_NAME As String = "cNTPTimeProvider"

'=========================================================================
' 私有成员
'=========================================================================

Private m_eTimeSource As TimeSourceType
Private m_sExternalServer As String
Private m_lExternalPort As Long
Private m_dCustomOffset As Double
Private m_bUseCustomTime As Boolean
Private m_dtCustomTime As Date

'=========================================================================
' 时间源类型枚举
'=========================================================================

Public Enum TimeSourceType
    TimeSource_System = 0      ' 使用系统时间
    TimeSource_External = 1    ' 使用外部 NTP 服务器
    TimeSource_Custom = 2      ' 使用自定义时间（测试用）
End Enum

'=========================================================================
' 属性
'=========================================================================

' 时间源类型
Public Property Get TimeSource() As TimeSourceType
    TimeSource = m_eTimeSource
End Property

Public Property Let TimeSource(ByVal Value As TimeSourceType)
    m_eTimeSource = Value
End Property

' 外部服务器地址 (当 TimeSource = TimeSource_External 时使用)
Public Property Get ExternalServer() As String
    ExternalServer = m_sExternalServer
End Property

Public Property Let ExternalServer(ByVal Value As String)
    m_sExternalServer = Value
End Property

' 外部服务器端口
Public Property Get ExternalPort() As Long
    ExternalPort = m_lExternalPort
End Property

Public Property Let ExternalPort(ByVal Value As Long)
    m_lExternalPort = Value
End Property

' 自定义时间偏移 (秒)
Public Property Get CustomOffset() As Double
    CustomOffset = m_dCustomOffset
End Property

Public Property Let CustomOffset(ByVal Value As Double)
    m_dCustomOffset = Value
End Property

' 是否使用自定义时间 (测试模式)
Public Property Get UseCustomTime() As Boolean
    UseCustomTime = m_bUseCustomTime
End Property

Public Property Let UseCustomTime(ByVal Value As Boolean)
    m_bUseCustomTime = Value
End Property

' 自定义时间 (测试模式)
Public Property Get CustomTime() As Date
    CustomTime = m_dtCustomTime
End Property

Public Property Let CustomTime(ByVal Value As Date)
    m_dtCustomTime = Value
End Property

'=========================================================================
' 公共方法
'=========================================================================

' 获取当前时间
Public Function GetCurrentTime() As Date
    Select Case m_eTimeSource
        Case TimeSource_System
            GetCurrentTime = GetSystemTime()
            
        Case TimeSource_External
            GetCurrentTime = GetExternalTime()
            
        Case TimeSource_Custom
            If m_bUseCustomTime Then
                GetCurrentTime = m_dtCustomTime
            Else
                GetCurrentTime = GetSystemTime()
            End If
            
    End Select
    
    ' 应用自定义偏移
    If Abs(m_dCustomOffset) > 0.0001 Then
        GetCurrentTime = DateAdd("s", m_dCustomOffset, GetCurrentTime)
    End If
End Function

' 同步外部时间源 (如果使用外部服务器)
Public Function SyncExternalSource() As Boolean
    If m_eTimeSource <> TimeSource_External Then
        Exit Function
    End If
    
    If LenB(m_sExternalServer) = 0 Then
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Dim oClient As cNTPClient
    Set oClient = New cNTPClient
    oClient.ServerHost = m_sExternalServer
    oClient.ServerPort = m_lExternalPort
    oClient.Timeout = 5
    
    Dim dtTime As Date
    dtTime = oClient.GetServerTime()
    
    ' 更新自定义偏移
    m_dCustomOffset = DateDiff("s", Now, dtTime)
    
    SyncExternalSource = True
    Exit Function
    
ErrorHandler:
    SyncExternalSource = False
End Function

'=========================================================================
' 私有方法
'=========================================================================

' 获取系统时间
Private Function GetSystemTime() As Date
    GetSystemTime = Now
End Function

' 从外部服务器获取时间
Private Function GetExternalTime() As Date
    Dim oClient As cNTPClient
    Set oClient = New cNTPClient
    
    oClient.ServerHost = m_sExternalServer
    oClient.ServerPort = m_lExternalPort
    oClient.Timeout = 5
    
    On Error GoTo ErrorHandler
    
    GetExternalTime = oClient.GetServerTime()
    Exit Function
    
ErrorHandler:
    ' 如果外部服务器失败，回退到系统时间
    GetExternalTime = Now
End Function

'=========================================================================
' 类初始化/终止
'=========================================================================

Private Sub Class_Initialize()
    m_eTimeSource = TimeSource_System
    m_sExternalServer = "pool.ntp.org"
    m_lExternalPort = 123
    m_dCustomOffset = 0
    m_bUseCustomTime = False
End Sub

Private Sub Class_Terminate()
End Sub
