VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMSMQManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
' 类名:    cMSMQManager
' 描述:    MSMQ 队列管理类
' 作者:    [作者]
' 日期:    [日期]
' 版本:    1.0.0
' 说明:    提供队列的创建、删除、查询等管理功能
'===========================================================================

'----------------------------------
' 事件声明
'----------------------------------
Public Event OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Public Event OnOperationComplete(ByVal Operation As String, ByVal Success As Boolean)

'----------------------------------
' 成员变量
'----------------------------------
Private m_LastError As String

'----------------------------------
' 属性: 最后错误信息
'----------------------------------
Public Property Get LastError() As String
    LastError = m_LastError
End Property

'----------------------------------
' 方法: 创建队列
'----------------------------------
Public Function CreateQueue( _
    ByVal QueueName As String, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate, _
    Optional ByVal Label As String = "", _
    Optional ByVal IsTransactional As Boolean = False _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim qinfo As Object
    Dim pathName As String
    
    Set qinfo = CreateObject("MSMQ.MSMQQueueInfo")
    
    ' 构建路径
    If QueueType = qtPrivate Then
        pathName = ".\private$\" & QueueName
    Else
        pathName = ".\" & QueueName
    End If
    
    qinfo.PathName = pathName
    
    ' 检查是否已存在
    If QueueExists(QueueName, QueueType) Then
        m_LastError = "队列已存在: " & QueueName
        CreateQueue = False
        Exit Function
    End If
    
    ' 设置标签
    If Len(Label) > 0 Then
        qinfo.Label = Label
    Else
        qinfo.Label = QueueName
    End If
    
    ' 创建队列
    qinfo.Create IsTransactional
    
    RaiseEvent OnOperationComplete("CreateQueue", True)
    CreateQueue = True
    Exit Function
    
ErrorHandler:
    m_LastError = "创建队列失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    RaiseEvent OnOperationComplete("CreateQueue", False)
    CreateQueue = False
End Function

'----------------------------------
' 方法: 删除队列
'----------------------------------
Public Function DeleteQueue( _
    ByVal QueueName As String, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim qinfo As Object
    Dim pathName As String
    
    Set qinfo = CreateObject("MSMQ.MSMQQueueInfo")
    
    ' 构建路径
    If QueueType = qtPrivate Then
        pathName = ".\private$\" & QueueName
    Else
        pathName = ".\" & QueueName
    End If
    
    qinfo.PathName = pathName
    
    ' 检查是否存在
    If Not QueueExists(QueueName, QueueType) Then
        m_LastError = "队列不存在: " & QueueName
        DeleteQueue = False
        Exit Function
    End If
    
    ' 删除队列
    qinfo.Delete
    
    RaiseEvent OnOperationComplete("DeleteQueue", True)
    DeleteQueue = True
    Exit Function
    
ErrorHandler:
    m_LastError = "删除队列失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    RaiseEvent OnOperationComplete("DeleteQueue", False)
    DeleteQueue = False
End Function

'----------------------------------
' 方法: 检查队列是否存在
'----------------------------------
Public Function QueueExists( _
    ByVal QueueName As String, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim qinfo As Object
    Dim pathName As String
    
    Set qinfo = CreateObject("MSMQ.MSMQQueueInfo")
    
    ' 构建路径
    If QueueType = qtPrivate Then
        pathName = ".\private$\" & QueueName
    Else
        pathName = ".\" & QueueName
    End If
    
    qinfo.PathName = pathName
    
    ' 尝试刷新，如果失败则队列不存在
    qinfo.Refresh
    QueueExists = True
    Exit Function
    
ErrorHandler:
    QueueExists = False
End Function

'----------------------------------
' 方法: 获取队列信息
'----------------------------------
Public Function GetQueueInfo( _
    ByVal QueueName As String, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate _
) As Object  ' 返回 Dictionary 对象
    
    On Error GoTo ErrorHandler
    
    Dim qinfo As Object
    Dim pathName As String
    Dim result As Object
    
    ' 创建 Dictionary 对象（需要引用 Scripting Runtime 或后期绑定）
    Set result = CreateObject("Scripting.Dictionary")
    
    Set qinfo = CreateObject("MSMQ.MSMQQueueInfo")
    
    If QueueType = qtPrivate Then
        pathName = ".\private$\" & QueueName
    Else
        pathName = ".\" & QueueName
    End If
    
    qinfo.PathName = pathName
    qinfo.Refresh
    
    result.Add "PathName", qinfo.PathName
    result.Add "FormatName", qinfo.FormatName
    result.Add "Label", qinfo.Label
    result.Add "IsTransactional", qinfo.IsTransactional
    result.Add "MessageCount", qinfo.MessageCount
    
    Set GetQueueInfo = result
    Exit Function
    
ErrorHandler:
    m_LastError = "获取队列信息失败: " & Err.Description
    Set GetQueueInfo = Nothing
End Function

'----------------------------------
' 方法: 清空队列（删除所有消息）
'----------------------------------
Public Function PurgeQueue( _
    ByVal QueueName As String, _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim qinfo As Object
    Dim queue As Object
    Dim pathName As String
    
    Set qinfo = CreateObject("MSMQ.MSMQQueueInfo")
    
    If QueueType = qtPrivate Then
        pathName = ".\private$\" & QueueName
    Else
        pathName = ".\" & QueueName
    End If
    
    qinfo.PathName = pathName
    
    ' 以接收模式打开并清空
    Set queue = qinfo.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    queue.Purge
    queue.Close
    
    RaiseEvent OnOperationComplete("PurgeQueue", True)
    PurgeQueue = True
    Exit Function
    
ErrorHandler:
    m_LastError = "清空队列失败: " & Err.Description
    RaiseEvent OnError(Err.Number, m_LastError)
    RaiseEvent OnOperationComplete("PurgeQueue", False)
    PurgeQueue = False
End Function

'----------------------------------
' 方法: 获取本地队列列表
'----------------------------------
Public Function GetQueueList( _
    Optional ByVal QueueType As MSMQQueueType = qtPrivate _
) As Collection
    
    Dim result As New Collection
    Dim query As Object
    Dim queues As Object
    Dim qinfo As Object
    
    On Error GoTo ErrorHandler
    
    Set query = CreateObject("MSMQ.MSMQQuery")
    
    ' 查询队列
    Set queues = query.LookupQueue
    
    ' 遍历队列
    Set qinfo = queues.Next
    Do While Not qinfo Is Nothing
        result.Add qinfo.PathName
        Set qinfo = queues.Next
    Loop
    
    Set GetQueueList = result
    Exit Function
    
ErrorHandler:
    m_LastError = "获取队列列表失败: " & Err.Description
    Set GetQueueList = result
End Function

'----------------------------------
' 私有常量
'----------------------------------
Private Const MQ_RECEIVE_ACCESS = 1
Private Const MQ_DENY_NONE = 0

'----------------------------------
' 枚举（本类内部使用，与 cMSMQClient 保持一致）
'----------------------------------
Private Enum MSMQQueueType
    qtPrivate = 0
    qtPublic = 1
End Enum
