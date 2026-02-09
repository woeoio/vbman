'=========================================================================
' cMQTTSession - MQTT Client Session
' MQTT 客户端会话管理
'=========================================================================
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMQTTSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const MODULE_NAME As String = "cMQTTSession"

' 会话属性
Public ClientId As String
Public UserName As String
Public IsConnected As Boolean
Public CleanSession As Boolean
Public KeepAlive As Long
Public LastActivity As Date
Public Socket As cWinsock          ' 关联的网络连接

' 订阅列表 (Key: TopicFilter, Value: QoS)
Private m_oSubscriptions As Collection

' 待发送的消息队列 (QoS 1/2)
Private m_oOutgoingQueue As Collection

' 待确认的消息 (QoS 2, Step 1: PUBREC received, waiting PUBREL)
Private m_oPendingPubRel As Collection

' 下一个报文标识符
Private m_lNextPacketId As Long

'=========================================================================
' 初始化/清理
'=========================================================================

Private Sub Class_Initialize()
    Set m_oSubscriptions = New Collection
    Set m_oOutgoingQueue = New Collection
    Set m_oPendingPubRel = New Collection
    m_lNextPacketId = 1
    LastActivity = Now
End Sub

Public Sub Reset()
    ' 保留会话状态
    If CleanSession Then
        Set m_oSubscriptions = New Collection
    End If
    
    Set m_oOutgoingQueue = New Collection
    Set m_oPendingPubRel = New Collection
    IsConnected = False
End Sub

'=========================================================================
' 订阅管理
'=========================================================================

Public Sub AddSubscription(ByVal sTopicFilter As String, ByVal eQoS As MqttQoS)
    On Error Resume Next
    m_oSubscriptions.Remove sTopicFilter
    On Error GoTo 0
    m_oSubscriptions.Add eQoS, sTopicFilter
End Sub

Public Sub RemoveSubscription(ByVal sTopicFilter As String)
    On Error Resume Next
    m_oSubscriptions.Remove sTopicFilter
End Sub

Public Function GetSubscriptionQoS(ByVal sTopicFilter As String) As MqttQoS
    On Error Resume Next
    GetSubscriptionQoS = m_oSubscriptions(sTopicFilter)
    If Err.Number <> 0 Then
        GetSubscriptionQoS = -1  ' Not subscribed
        Err.Clear
    End If
End Function

Public Function HasSubscription(ByVal sTopicFilter As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = m_oSubscriptions(sTopicFilter)
    HasSubscription = (Err.Number = 0)
    Err.Clear
End Function

' 检查主题是否匹配订阅
Public Function TopicMatches(ByVal sTopic As String) As Boolean
    Dim sFilter As Variant
    
    For Each sFilter In m_oSubscriptions
        If MatchTopicFilter(sTopic, CStr(sFilter)) Then
            TopicMatches = True
            Exit Function
        End If
    Next
End Function

' 获取匹配的最大 QoS
Public Function GetMatchingQoS(ByVal sTopic As String) As MqttQoS
    Dim sFilter As Variant
    Dim eQoS As MqttQoS
    Dim eMaxQoS As MqttQoS
    
    eMaxQoS = -1
    
    For Each sFilter In m_oSubscriptions
        If MatchTopicFilter(sTopic, CStr(sFilter)) Then
            eQoS = m_oSubscriptions(sFilter)
            If eQoS > eMaxQoS Then
                eMaxQoS = eQoS
            End If
        End If
    Next
    
    GetMatchingQoS = eMaxQoS
End Function

'=========================================================================
' 报文标识符管理
'=========================================================================

Public Function GetNextPacketId() As Long
    m_lNextPacketId = m_lNextPacketId + 1
    If m_lNextPacketId > 65535 Then
        m_lNextPacketId = 1
    End If
    GetNextPacketId = m_lNextPacketId
End Function

'=========================================================================
' QoS 消息队列管理
'=========================================================================

Public Sub AddOutgoingMessage(ByVal lPacketId As Long, oMessage As cMQTTMessage)
    Dim oEntry As New cMQTTPendingMessage
    oEntry.PacketId = lPacketId
    Set oEntry.Message = oMessage
    oEntry.Timestamp = Now
    oEntry.RetryCount = 0
    
    On Error Resume Next
    m_oOutgoingQueue.Remove "P" & lPacketId
    On Error GoTo 0
    m_oOutgoingQueue.Add oEntry, "P" & lPacketId
End Sub

Public Function GetOutgoingMessage(ByVal lPacketId As Long) As cMQTTPendingMessage
    On Error Resume Next
    Set GetOutgoingMessage = m_oOutgoingQueue("P" & lPacketId)
    Err.Clear
End Function

Public Sub RemoveOutgoingMessage(ByVal lPacketId As Long)
    On Error Resume Next
    m_oOutgoingQueue.Remove "P" & lPacketId
End Sub

' QoS 2: 等待 PUBREL
Public Sub AddPendingPubRel(ByVal lPacketId As Long)
    Dim oEntry As New cMQTTPendingMessage
    oEntry.PacketId = lPacketId
    oEntry.Timestamp = Now
    
    On Error Resume Next
    m_oPendingPubRel.Remove "P" & lPacketId
    On Error GoTo 0
    m_oPendingPubRel.Add oEntry, "P" & lPacketId
End Sub

Public Sub RemovePendingPubRel(ByVal lPacketId As Long)
    On Error Resume Next
    m_oPendingPubRel.Remove "P" & lPacketId
End Sub

Public Function IsPendingPubRel(ByVal lPacketId As Long) As Boolean
    On Error Resume Next
    Dim v As Variant
    Set v = m_oPendingPubRel("P" & lPacketId)
    IsPendingPubRel = (Err.Number = 0)
    Err.Clear
End Function

'=========================================================================
' 保活检查
'=========================================================================

Public Function IsExpired(ByVal lTimeout As Long) As Boolean
    If KeepAlive = 0 Then Exit Function
    
    Dim lElapsed As Long
    lElapsed = DateDiff("s", LastActivity, Now)
    IsExpired = (lElapsed > KeepAlive * 1.5 + lTimeout)
End Function

Public Sub UpdateActivity()
    LastActivity = Now
End Sub

'=========================================================================
' 属性
'=========================================================================

Public Property Get SubscriptionCount() As Long
    SubscriptionCount = m_oSubscriptions.Count
End Property

Public Property Get Subscriptions() As Collection
    Set Subscriptions = m_oSubscriptions
End Property

Public Property Get OutgoingMessageCount() As Long
    OutgoingMessageCount = m_oOutgoingQueue.Count
End Property

'=========================================================================
' 私有方法: 主题匹配
'=========================================================================

Private Function MatchTopicFilter(ByVal sTopic As String, ByVal sFilter As String) As Boolean
    Dim vTopicLevels As Variant
    Dim vFilterLevels As Variant
    Dim i As Long
    
    vTopicLevels = Split(sTopic, "/")
    vFilterLevels = Split(sFilter, "/")
    
    i = 0
    Do While i < UBound(vFilterLevels) + 1
        ' 多级通配符 # 必须位于最后
        If vFilterLevels(i) = "#" Then
            MatchTopicFilter = True
            Exit Function
        End If
        
        ' 检查是否还有主题层级
        If i >= UBound(vTopicLevels) + 1 Then
            MatchTopicFilter = False
            Exit Function
        End If
        
        ' 单级通配符 +
        If vFilterLevels(i) <> "+" Then
            If vFilterLevels(i) <> vTopicLevels(i) Then
                MatchTopicFilter = False
                Exit Function
            End If
        End If
        
        i = i + 1
    Loop
    
    ' 检查主题是否还有额外层级
    MatchTopicFilter = (i = UBound(vTopicLevels) + 1)
End Function
