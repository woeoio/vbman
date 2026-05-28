VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMSMQAsyncReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
' 类名:    cMSMQAsyncReceiver
' 描述:    MSMQ 异步消息接收器
' 作者:    [作者]
' 日期:    [日期]
' 版本:    1.0.0
' 说明:    使用事件驱动方式异步接收消息，避免阻塞UI线程
'===========================================================================

'----------------------------------
' MSMQ 事件对象声明
'----------------------------------
Private WithEvents m_Events As MSMQEvent
Attribute m_Events.VB_VarHelpID = -1

'----------------------------------
' 成员变量
'----------------------------------
Private m_Queue As Object           ' MSMQQueue
Private m_QueueInfo As Object       ' MSMQQueueInfo
Private m_IsListening As Boolean
Private m_PathName As String
Private m_LastError As String
Private m_AutoReconnect As Boolean
Private m_ReceiveTimeout As Long

'----------------------------------
' 事件声明
'----------------------------------
Public Event OnMessageArrived(ByRef Message As cMSMQMessage)
Public Event OnReceiveError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Public Event OnListenerStarted()
Public Event OnListenerStopped()

'----------------------------------
' 属性
'----------------------------------
Public Property Get IsListening() As Boolean
    IsListening = m_IsListening
End Property

Public Property Get PathName() As String
    PathName = m_PathName
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Property Get AutoReconnect() As Boolean
    AutoReconnect = m_AutoReconnect
End Property

Public Property Let AutoReconnect(ByVal Value As Boolean)
    m_AutoReconnect = Value
End Property

Public Property Get ReceiveTimeout() As Long
    ReceiveTimeout = m_ReceiveTimeout
End Property

Public Property Let ReceiveTimeout(ByVal Value As Long)
    m_ReceiveTimeout = Value
End Property

'----------------------------------
' 方法: 开始监听
'----------------------------------
Public Function StartListening( _
    ByVal QueueName As String, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate, _
    Optional ByVal CreateIfNotExist As Boolean = True _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 如果已在监听，先停止
    If m_IsListening Then StopListening
    
    ' 构建路径
    If QueueType = qtPrivate Then
        m_PathName = ".\private$\" & QueueName
    Else
        m_PathName = ".\" & QueueName
    End If
    
    ' 创建队列信息对象
    Set m_QueueInfo = CreateObject("MSMQ.MSMQQueueInfo")
    m_QueueInfo.PathName = m_PathName
    
    ' 如果需要创建
    If CreateIfNotExist Then
        On Error Resume Next
        m_QueueInfo.Create
        On Error GoTo ErrorHandler
    End If
    
    ' 以接收模式打开队列
    Set m_Queue = m_QueueInfo.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    
    ' 创建事件对象
    Set m_Events = New MSMQEvent
    
    ' 启用通知
    m_Queue.EnableNotification m_Events, , m_ReceiveTimeout
    
    m_IsListening = True
    RaiseEvent OnListenerStarted
    StartListening = True
    Exit Function
    
ErrorHandler:
    m_LastError = "开始监听失败: " & Err.Description
    RaiseEvent OnReceiveError(Err.Number, m_LastError)
    StartListening = False
End Function

'----------------------------------
' 方法: 通过格式名监听（远程队列）
'----------------------------------
Public Function StartListeningByFormatName( _
    ByVal FormatName As String _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    If m_IsListening Then StopListening
    
    m_PathName = FormatName
    
    Set m_QueueInfo = CreateObject("MSMQ.MSMQQueueInfo")
    m_QueueInfo.FormatName = FormatName
    
    Set m_Queue = m_QueueInfo.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set m_Events = New MSMQEvent
    
    m_Queue.EnableNotification m_Events, , m_ReceiveTimeout
    
    m_IsListening = True
    RaiseEvent OnListenerStarted
    StartListeningByFormatName = True
    Exit Function
    
ErrorHandler:
    m_LastError = "监听远程队列失败: " & Err.Description
    RaiseEvent OnReceiveError(Err.Number, m_LastError)
    StartListeningByFormatName = False
End Function

'----------------------------------
' 方法: 停止监听
'----------------------------------
Public Sub StopListening()
    On Error Resume Next
    
    m_IsListening = False
    
    ' 禁用通知
    If Not m_Queue Is Nothing Then
        m_Queue.EnableNotification Nothing
        m_Queue.Close
    End If
    
    Set m_Events = Nothing
    Set m_Queue = Nothing
    Set m_QueueInfo = Nothing
    
    RaiseEvent OnListenerStopped
End Sub

'----------------------------------
' 事件处理: 消息到达
'----------------------------------
Private Sub m_Events_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    Dim msg As Object
    Dim result As New cMSMQMessage
    
    On Error GoTo ErrorHandler
    
    ' 接收消息
    Set msg = Queue.Receive(MQ_NO_TRANSACTION, False, True, 0)
    
    If Not msg Is Nothing Then
        result.LoadFromMSMQMessage msg
        RaiseEvent OnMessageArrived(result)
    End If
    
    ' 继续监听（如果仍在监听状态）
    If m_IsListening Then
        Queue.EnableNotification m_Events, , m_ReceiveTimeout
    End If
    Exit Sub
    
ErrorHandler:
    m_LastError = "处理消息到达事件失败: " & Err.Description
    RaiseEvent OnReceiveError(Err.Number, m_LastError)
    
    ' 错误处理：继续监听或重连
    If m_IsListening Then
        If m_AutoReconnect Then
            On Error Resume Next
            Queue.EnableNotification m_Events, , m_ReceiveTimeout
        End If
    End If
End Sub

'----------------------------------
' 事件处理: 接收错误
'----------------------------------
Private Sub m_Events_ArrivedError(ByVal Queue As Object, ByVal ErrorCode As Long, ByVal Cursor As Long)
    m_LastError = "接收消息错误，错误码: " & ErrorCode
    RaiseEvent OnReceiveError(ErrorCode, m_LastError)
    
    ' 继续监听
    If m_IsListening And m_AutoReconnect Then
        On Error Resume Next
        Queue.EnableNotification m_Events, , m_ReceiveTimeout
    End If
End Sub

'----------------------------------
' 清理
'----------------------------------
Private Sub Class_Initialize()
    m_AutoReconnect = True
    m_ReceiveTimeout = 0  ' 默认无超时
End Sub

Private Sub Class_Terminate()
    StopListening
End Sub

'----------------------------------
' 常量
'----------------------------------
Private Const MQ_RECEIVE_ACCESS = 1
Private Const MQ_DENY_NONE = 0
Private Const MQ_NO_TRANSACTION = 0

Private Enum MSMQQueueType
    qtPrivate = 0
    qtPublic = 1
End Enum
