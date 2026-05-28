VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMSMQReceiverPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
' 类名:    cMSMQReceiverPool
' 描述:    MSMQ 多线程接收池（模拟）
' 作者:    [作者]
' 日期:    [日期]
' 版本:    1.0.0
' 说明:    通过多个异步接收器实现消息并行处理，提高吞吐量
'===========================================================================

'----------------------------------
' 成员变量
'----------------------------------
Private m_Receivers As Collection       ' 接收器集合
Private m_QueueName As String
Private m_WorkerCount As Long
Private m_IsRunning As Boolean
Private m_LastError As String

'----------------------------------
' 事件声明
'----------------------------------
Public Event OnMessageReceived(ByVal WorkerId As Long, ByRef Message As cMSMQMessage)
Public Event OnWorkerError(ByVal WorkerId As Long, ByVal ErrorDescription As String)
Public Event OnPoolStarted()
Public Event OnPoolStopped()

'----------------------------------
' 属性
'----------------------------------
Public Property Get IsRunning() As Boolean
    IsRunning = m_IsRunning
End Property

Public Property Get WorkerCount() As Long
    WorkerCount = m_WorkerCount
End Property

Public Property Get QueueName() As String
    QueueName = m_QueueName
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

'----------------------------------
' 方法: 启动接收池
'----------------------------------
Public Function Start( _
    ByVal QueueName As String, _
    ByVal WorkerCount As Long, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 停止已有实例
    If m_IsRunning Then StopPool
    
    m_QueueName = QueueName
    m_WorkerCount = WorkerCount
    
    Set m_Receivers = New Collection
    
    Dim i As Long
    Dim receiver As cMSMQAsyncReceiver
    
    For i = 1 To WorkerCount
        Set receiver = New cMSMQAsyncReceiver
        receiver.AutoReconnect = True
        
        ' 绑定事件（使用 WithEvents 数组需要技巧，这里简化处理）
        ' 实际实现中，每个 receiver 应该有自己的事件处理
        
        If receiver.StartListening(QueueName, QueueType, True) Then
            m_Receivers.Add receiver, "Worker" & i
        Else
            m_LastError = "Worker " & i & " 启动失败: " & receiver.LastError
        End If
    Next i
    
    If m_Receivers.Count > 0 Then
        m_IsRunning = True
        RaiseEvent OnPoolStarted
        Start = True
    Else
        m_LastError = "没有成功启动任何 Worker"
        Start = False
    End If
    Exit Function
    
ErrorHandler:
    m_LastError = "启动接收池失败: " & Err.Description
    Start = False
End Function

'----------------------------------
' 方法: 停止接收池
'----------------------------------
Public Sub StopPool()
    On Error Resume Next
    
    Dim receiver As cMSMQAsyncReceiver
    Dim i As Long
    
    For i = m_Receivers.Count To 1 Step -1
        Set receiver = m_Receivers(i)
        receiver.StopListening
        m_Receivers.Remove i
    Next i
    
    Set m_Receivers = Nothing
    m_IsRunning = False
    
    RaiseEvent OnPoolStopped
End Sub

'----------------------------------
' 方法: 获取活动Worker数量
'----------------------------------
Public Function GetActiveWorkerCount() As Long
    If Not m_Receivers Is Nothing Then
        GetActiveWorkerCount = m_Receivers.Count
    End If
End Function

'----------------------------------
' 清理
'----------------------------------
Private Sub Class_Terminate()
    StopPool
End Sub

Private Enum MSMQQueueType
    qtPrivate = 0
    qtPublic = 1
End Enum
