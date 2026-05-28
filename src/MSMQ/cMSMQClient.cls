VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMSMQClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
' 类名:    cMSMQClient
' 描述:    MSMQ 消息队列客户端封装类
' 作者:    [作者]
' 日期:    [日期]
' 版本:    1.0.0
' 说明:    封装 MSMQ 常用操作，简化消息队列使用
' 引用:    Microsoft Message Queue Object Library (mqoa.dll)
'===========================================================================

'----------------------------------
' Windows API 声明
'----------------------------------
' 如需超时处理可添加 Sleep API
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'----------------------------------
' 枚举类型
'----------------------------------
Public Enum MSMQAccessMode
    maSendOnly = 0      ' 仅发送
    maReceiveOnly = 1   ' 仅接收
    maSendReceive = 2   ' 发送和接收
    maPeekOnly = 3      ' 仅查看（不删除）
End Enum

Public Enum MSMQQueueType
    qtPrivate = 0       ' 专用队列（本地）
    qtPublic = 1        ' 公共队列（需要域环境）
End Enum

Public Enum MSMQDeliveryMode
    dmExpress = 0       ' 快速（内存中，可能丢失）
    dmRecoverable = 1   ' 可恢复（持久化到磁盘）
End Enum

'----------------------------------
' 事件声明
'----------------------------------
Public Event OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Public Event OnConnected()
Public Event OnDisconnected()

'----------------------------------
' 成员变量
'----------------------------------
Private m_QueueInfo As Object      ' MSMQQueueInfo
Private m_Queue As Object          ' MSMQQueue
Private m_IsOpen As Boolean
Private m_PathName As String
Private m_FormatName As String
Private m_AccessMode As MSMQAccessMode
Private m_LastError As String

'----------------------------------
' 属性
'----------------------------------
Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property

Public Property Get PathName() As String
    PathName = m_PathName
End Property

Public Property Get FormatName() As String
    FormatName = m_FormatName
End Property

Public Property Get AccessMode() As MSMQAccessMode
    AccessMode = m_AccessMode
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

'----------------------------------
' 方法: 连接到队列
'----------------------------------
Public Function Connect( _
    ByVal QueueName As String, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate, _
    Optional ByVal AccessMode As MSMQAccessMode = maSendReceive, _
    Optional ByVal CreateIfNotExist As Boolean = True _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 关闭已有连接
    If m_IsOpen Then CloseConnection
    
    ' 保存参数
    m_AccessMode = AccessMode
    
    ' 构建队列路径
    If QueueType = qtPrivate Then
        m_PathName = ".\private$\" & QueueName
    Else
        m_PathName = ".\" & QueueName
    End If
    
    ' 创建队列信息对象
    Set m_QueueInfo = CreateObject("MSMQ.MSMQQueueInfo")
    m_QueueInfo.PathName = m_PathName
    
    ' 如果队列不存在且允许创建
    If CreateIfNotExist Then
        On Error Resume Next
        m_QueueInfo.Create
        On Error GoTo ErrorHandler
    End If
    
    ' 获取格式名
    m_FormatName = m_QueueInfo.FormatName
    
    ' 根据访问模式打开队列
    Dim lAccess As Long
    Select Case AccessMode
        Case maSendOnly
            lAccess = MQ_SEND_ACCESS
        Case maReceiveOnly
            lAccess = MQ_RECEIVE_ACCESS
        Case maSendReceive
            lAccess = MQ_SEND_ACCESS Or MQ_RECEIVE_ACCESS
        Case maPeekOnly
            lAccess = MQ_PEEK_ACCESS
        Case Else
            lAccess = MQ_SEND_ACCESS
    End Select
    
    Set m_Queue = m_QueueInfo.Open(lAccess, MQ_DENY_NONE)
    m_IsOpen = True
    
    RaiseEvent OnConnected
    Connect = True
    Exit Function
    
ErrorHandler:
    m_LastError = "连接队列失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    Connect = False
End Function

'----------------------------------
' 方法: 通过格式名连接（用于远程队列）
'----------------------------------
Public Function ConnectByFormatName( _
    ByVal FormatName As String, _
    Optional ByVal AccessMode As MSMQAccessMode = maSendReceive _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    If m_IsOpen Then CloseConnection
    
    m_AccessMode = AccessMode
    m_FormatName = FormatName
    
    Set m_QueueInfo = CreateObject("MSMQ.MSMQQueueInfo")
    m_QueueInfo.FormatName = FormatName
    
    Dim lAccess As Long
    Select Case AccessMode
        Case maSendOnly: lAccess = MQ_SEND_ACCESS
        Case maReceiveOnly: lAccess = MQ_RECEIVE_ACCESS
        Case maSendReceive: lAccess = MQ_SEND_ACCESS Or MQ_RECEIVE_ACCESS
        Case maPeekOnly: lAccess = MQ_PEEK_ACCESS
        Case Else: lAccess = MQ_SEND_ACCESS
    End Select
    
    Set m_Queue = m_QueueInfo.Open(lAccess, MQ_DENY_NONE)
    m_IsOpen = True
    
    RaiseEvent OnConnected
    ConnectByFormatName = True
    Exit Function
    
ErrorHandler:
    m_LastError = "连接远程队列失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    ConnectByFormatName = False
End Function

'----------------------------------
' 方法: 发送消息
'----------------------------------
Public Function SendMessage( _
    ByVal MessageBody As Variant, _
    Optional ByVal Label As String = "", _
    Optional ByVal Priority As Long = 3, _
    Optional ByVal Delivery As MSMQDeliveryMode = dmRecoverable, _
    Optional ByVal AppSpecific As Long = 0 _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    If Not m_IsOpen Then
        m_LastError = "队列未连接"
        SendMessage = False
        Exit Function
    End If
    
    If m_AccessMode = maReceiveOnly Or m_AccessMode = maPeekOnly Then
        m_LastError = "当前模式不允许发送"
        SendMessage = False
        Exit Function
    End If
    
    Dim msg As Object  ' MSMQMessage
    Set msg = CreateObject("MSMQ.MSMQMessage")
    
    msg.Body = MessageBody
    msg.Label = Label
    msg.Priority = Priority
    msg.Delivery = Delivery
    msg.AppSpecific = AppSpecific
    
    m_Queue.Send msg
    
    SendMessage = True
    Exit Function
    
ErrorHandler:
    m_LastError = "发送消息失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    SendMessage = False
End Function

'----------------------------------
' 方法: 接收消息（同步）
'----------------------------------
Public Function ReceiveMessage( _
    Optional ByVal Timeout As Long = 0, _
    Optional ByVal WantBody As Boolean = True _
) As cMSMQMessage
    
    On Error GoTo ErrorHandler
    
    If Not m_IsOpen Then
        m_LastError = "队列未连接"
        Set ReceiveMessage = Nothing
        Exit Function
    End If
    
    If m_AccessMode = maSendOnly Then
        m_LastError = "当前模式不允许接收"
        Set ReceiveMessage = Nothing
        Exit Function
    End If
    
    Dim msg As Object  ' MSMQMessage
    Dim result As New cMSMQMessage
    
    ' 接收消息
    If Timeout > 0 Then
        Set msg = m_Queue.Receive(Timeout)
    Else
        Set msg = m_Queue.Receive()
    End If
    
    If msg Is Nothing Then
        Set ReceiveMessage = Nothing
        Exit Function
    End If
    
    result.LoadFromMSMQMessage msg
    Set ReceiveMessage = result
    Exit Function
    
ErrorHandler:
    m_LastError = "接收消息失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    Set ReceiveMessage = Nothing
End Function

'----------------------------------
' 方法: 查看消息（不删除）
'----------------------------------
Public Function PeekMessage( _
    Optional ByVal Timeout As Long = 0 _
) As cMSMQMessage
    
    On Error GoTo ErrorHandler
    
    If Not m_IsOpen Then
        m_LastError = "队列未连接"
        Set PeekMessage = Nothing
        Exit Function
    End If
    
    Dim msg As Object
    Dim result As New cMSMQMessage
    
    If Timeout > 0 Then
        Set msg = m_Queue.Peek(Timeout)
    Else
        Set msg = m_Queue.Peek()
    End If
    
    If msg Is Nothing Then
        Set PeekMessage = Nothing
        Exit Function
    End If
    
    result.LoadFromMSMQMessage msg
    Set PeekMessage = result
    Exit Function
    
ErrorHandler:
    m_LastError = "查看消息失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    Set PeekMessage = Nothing
End Function

'----------------------------------
' 方法: 清空队列
'----------------------------------
Public Function Purge() As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsOpen Then
        m_LastError = "队列未连接"
        Purge = False
        Exit Function
    End If
    
    m_Queue.Purge
    Purge = True
    Exit Function
    
ErrorHandler:
    m_LastError = "清空队列失败: " & Err.Description
    Purge = False
End Function

'----------------------------------
' 方法: 关闭连接
'----------------------------------
Public Sub CloseConnection()
    On Error Resume Next
    
    If Not m_Queue Is Nothing Then
        m_Queue.Close
        Set m_Queue = Nothing
    End If
    
    Set m_QueueInfo = Nothing
    m_IsOpen = False
    m_PathName = ""
    m_FormatName = ""
    
    RaiseEvent OnDisconnected
End Sub

'----------------------------------
' 清理
'----------------------------------
Private Sub Class_Terminate()
    CloseConnection
End Sub

'----------------------------------
' 常量定义（本类内部使用）
'----------------------------------
Private Const MQ_RECEIVE_ACCESS = 1
Private Const MQ_SEND_ACCESS = 2
Private Const MQ_PEEK_ACCESS = 32
Private Const MQ_DENY_NONE = 0
