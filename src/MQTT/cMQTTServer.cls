'=========================================================================
' cMQTTServer - MQTT Broker/Server
' MQTT 服务端实现
'=========================================================================
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMQTTServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const MODULE_NAME As String = "cMQTTServer"

' 事件
Event ClientConnected(ClientId As String, Session As cMQTTSession)
Event ClientDisconnected(ClientId As String)
Event MessageReceived(ClientId As String, Topic As String, Payload As Variant, QoS As MqttQoS)
Event Subscribe(ClientId As String, TopicFilter As String, QoS As MqttQoS)
Event Unsubscribe(ClientId As String, TopicFilter As String)
Event ErrorEvent(Number As Long, Description As String)

' 网络层
Private WithEvents m_oServerSocket As cWinsock
Attribute m_oServerSocket.VB_VarHelpID = -1

' 客户端会话 (Key: ClientId, Value: cMQTTSession)
Private m_oSessions As Collection

' 保留消息 (Key: Topic, Value: cMQTTMessage)
Private m_oRetainedMessages As Collection

' 报文编解码器
Private m_oPacket As cMQTTPacket

' 服务端配置
Private m_lPort As Long
Private m_bRunning As Boolean
Private m_lKeepAliveCheckInterval As Long

' 保活检查定时器
Private m_lKeepAliveTimer As Long

'=========================================================================
' 初始化
'=========================================================================

Private Sub Class_Initialize()
    Set m_oServerSocket = New cWinsock
    Set m_oSessions = New Collection
    Set m_oRetainedMessages = New Collection
    Set m_oPacket = New cMQTTPacket
    
    m_lPort = MQTT_DEFAULT_PORT
    m_lKeepAliveCheckInterval = 5  ' 5 seconds
End Sub

Private Sub Class_Terminate()
    StopServer
End Sub

'=========================================================================
' 公共方法
'=========================================================================

Public Function StartServer(Optional ByVal Port As Long = 0) As Boolean
    On Error GoTo EH
    
    If Port > 0 Then
        m_lPort = Port
    End If
    
    ' 创建监听套接字
    m_oServerSocket.Protocol = wskTCPProtocol
    m_oServerSocket.Listen m_lPort
    
    m_bRunning = True
    
    Debug.Print "MQTT Server started on port " & m_lPort
    StartServer = True
    Exit Function
EH:
    RaiseEvent ErrorEvent(Err.Number, "Failed to start server: " & Err.Description)
End Function

Public Sub StopServer()
    Dim oSession As cMQTTSession
    
    On Error Resume Next
    
    m_bRunning = False
    
    ' 断开所有客户端
    For Each oSession In m_oSessions
        If Not oSession.Socket Is Nothing Then
            DisconnectClient oSession, True
        End If
    Next
    
    ' 清空会话
    Set m_oSessions = New Collection
    Set m_oRetainedMessages = New Collection
    
    ' 关闭监听
    m_oServerSocket.Close_
    
    Debug.Print "MQTT Server stopped"
End Sub

' 发布消息到指定主题
Public Sub PublishMessage(ByVal sTopic As String, _n                          ByVal vPayload As Variant, _n                          Optional ByVal eQoS As MqttQoS = mqttQoS0, _n                          Optional ByVal bRetain As Boolean = False)
    Dim oMessage As New cMQTTMessage
    
    oMessage.Topic = sTopic
    If VarType(vPayload) = vbByte + vbArray Then
        oMessage.PayloadBytes = vPayload
    Else
        oMessage.Payload = CStr(vPayload)
    End If
    oMessage.QoS = eQoS
    oMessage.Retain = bRetain
    
    ' 保存保留消息
    If bRetain Then
        SaveRetainedMessage oMessage
    End If
    
    ' 路由消息到匹配的订阅者
    RouteMessage oMessage
End Sub

' 强制断开客户端
Public Sub KickClient(ByVal sClientId As String)
    Dim oSession As cMQTTSession
    
    On Error Resume Next
    Set oSession = m_oSessions(sClientId)
    If Not oSession Is Nothing Then
        DisconnectClient oSession, False
    End If
End Sub

'=========================================================================
' 属性
'=========================================================================

Public Property Get IsRunning() As Boolean
    IsRunning = m_bRunning
End Property

Public Property Get Port() As Long
    Port = m_lPort
End Property

Public Property Get ClientCount() As Long
    ClientCount = m_oServerSocket.ClientCount
End Property

'=========================================================================
' 套接字事件处理
'=========================================================================

Private Sub m_oServerSocket_ConnectionRequest(Client As cWinsock, DisConnect As Boolean)
    Debug.Print "MQTT Client connecting from " & Client.RemoteHostIP
End Sub

Private Sub m_oServerSocket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    Dim lRemainingOffset As Long
    Dim lRemainingLength As Long
    Dim oSession As cMQTTSession
    
    On Error GoTo EH
    
    ' 读取数据
    Client.GetData baData, vbByte + vbArray
    
    ' 查找或创建会话
    Set oSession = FindSessionBySocket(Client)
    
    ' 解析固定报头
    If Not m_oPacket.DecodeFixedHeader(baData, lRemainingOffset, lRemainingLength) Then
        Debug.Print "Invalid MQTT packet from " & Client.RemoteHostIP
        GoTo DisconnectClient_
    End If
    
    ' 检查报文完整性
    If UBound(baData) + 1 < lRemainingOffset + lRemainingLength Then
        Debug.Print "Incomplete MQTT packet"
        Exit Sub
    End If
    
    ' 处理报文
    Select Case m_oPacket.PacketType
        Case mqttConnect
            HandleConnect Client, baData, lRemainingOffset
        Case mqttPublish
            HandlePublish oSession, baData, lRemainingOffset
        Case mqttPubAck
            HandlePubAck oSession, baData, lRemainingOffset
        Case mqttPubRec
            HandlePubRec oSession, baData, lRemainingOffset
        Case mqttPubRel
            HandlePubRel oSession, baData, lRemainingOffset
        Case mqttPubComp
            HandlePubComp oSession, baData, lRemainingOffset
        Case mqttSubscribe
            HandleSubscribe oSession, baData, lRemainingOffset
        Case mqttUnsubscribe
            HandleUnsubscribe oSession, baData, lRemainingOffset
        Case mqttPingReq
            HandlePingReq oSession
        Case mqttDisconnect
            HandleDisconnect oSession
        Case Else
            Debug.Print "Unknown packet type: " & m_oPacket.PacketType
    End Select
    
    Exit Sub
EH:
    Debug.Print "DataArrival Error: " & Err.Description
    Exit Sub
DisconnectClient_:
    Client.Close_
End Sub

Private Sub m_oServerSocket_CloseEvent(Client As cWinsock)
    Dim oSession As cMQTTSession
    
    Set oSession = FindSessionBySocket(Client)
    If Not oSession Is Nothing Then
        Debug.Print "Client " & oSession.ClientId & " disconnected"
        If oSession.IsConnected Then
            DisconnectClient oSession, False
        End If
    End If
End Sub

Private Sub m_oServerSocket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Debug.Print "Socket Error " & Number & ": " & Description
End Sub

'=========================================================================
' 报文处理
'=========================================================================

Private Sub HandleConnect(oClient As cWinsock, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim oSession As cMQTTSession
    Dim baResponse() As Byte
    Dim eReturnCode As MqttConnectReturnCode
    Dim bSessionPresent As Boolean
    
    On Error GoTo EH
    
    ' 解析 CONNECT 报文
    Set oPacket.PacketType = mqttConnect
    If Not oPacket.DecodePayload(baData, lOffset) Then
        eReturnCode = mqttUnacceptableVersion
        GoTo SendConnAck
    End If
    
    ' 验证协议版本
    If oPacket.ProtocolName <> "MQTT" Then
        eReturnCode = mqttUnacceptableVersion
        GoTo SendConnAck
    End If
    
    If oPacket.ProtocolLevel <> MQTT_PROTOCOL_LEVEL_311 Then
        eReturnCode = mqttUnacceptableVersion
        GoTo SendConnAck
    End If
    
    ' 验证 Client ID
    If Len(oPacket.ClientId) = 0 Or Len(oPacket.ClientId) > MQTT_MAX_CLIENT_ID_LEN Then
        eReturnCode = mqttIdentifierRejected
        GoTo SendConnAck
    End If
    
    ' 检查是否已有同名会话
    On Error Resume Next
    Set oSession = m_oSessions(oPacket.ClientId)
    On Error GoTo EH
    
    If Not oSession Is Nothing Then
        ' 断开旧连接
        If Not oSession.Socket Is Nothing Then
            oSession.Socket.Close_
        End If
        
        ' 清理旧会话（如果是 Clean Session）
        If oPacket.CleanSession Then
            m_oSessions.Remove oPacket.ClientId
            Set oSession = Nothing
        Else
            ' 恢复会话
            oSession.Reset
            bSessionPresent = True
        End If
    End If
    
    ' 创建新会话
    If oSession Is Nothing Then
        Set oSession = New cMQTTSession
        oSession.ClientId = oPacket.ClientId
        m_oSessions.Add oSession, oPacket.ClientId
    End If
    
    ' 设置会话属性
    Set oSession.Socket = oClient
    oSession.UserName = oPacket.UserName
    oSession.CleanSession = oPacket.CleanSession
    oSession.KeepAlive = oPacket.KeepAlive
    oSession.IsConnected = True
    oSession.UpdateActivity
    
    ' 设置遗嘱消息
    If Len(oPacket.WillTopic) > 0 Then
        ' 遗嘱消息会在连接异常断开时发送
        ' 存储在会话中，待实现
    End If
    
    eReturnCode = mqttAccepted
    
    Debug.Print "Client connected: " & oPacket.ClientId & " from " & oClient.RemoteHostIP
    RaiseEvent ClientConnected(oPacket.ClientId, oSession)
    
SendConnAck:
    ' 发送 CONNACK
    baResponse = oPacket.EncodeConnAck(eReturnCode, bSessionPresent)
    oClient.SendData baResponse
    
    ' 如果连接被拒绝，断开连接
    If eReturnCode <> mqttAccepted Then
        oClient.Close_
    End If
    
    Exit Sub
EH:
    Debug.Print "HandleConnect Error: " & Err.Description
End Sub

Private Sub HandlePublish(oSession As cMQTTSession, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim baResponse() As Byte
    Dim oMessage As New cMQTTMessage
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    
    oSession.UpdateActivity
    
    ' 解析 PUBLISH 报文
    Set oPacket.PacketType = mqttPublish
    oPacket.DecodeFixedHeader baData, lOffset, 0
    If Not oPacket.DecodePayload(baData, lOffset) Then Exit Sub
    
    ' 构建消息对象
    oMessage.Topic = oPacket.TopicName
    oMessage.PayloadBytes = oPacket.PayloadBytes
    oMessage.QoS = oPacket.QoS
    oMessage.Retain = oPacket.Retain
    oMessage.Duplicate = oPacket.Duplicate
    
    ' 保存保留消息
    If oMessage.Retain Then
        SaveRetainedMessage oMessage
    End If
    
    ' 路由消息
    RouteMessage oMessage
    
    ' QoS 1: 发送 PUBACK
    If oPacket.QoS = mqttQoS1 Then
        baResponse = oPacket.EncodePubAck(oPacket.PacketId)
        oSession.Socket.SendData baResponse
    End If
    
    ' QoS 2: 发送 PUBREC
    If oPacket.QoS = mqttQoS2 Then
        oSession.AddPendingPubRel oPacket.PacketId
        baResponse = oPacket.EncodePubRec(oPacket.PacketId)
        oSession.Socket.SendData baResponse
    End If
    
    RaiseEvent MessageReceived(oSession.ClientId, oMessage.Topic, oMessage.PayloadBytes, oMessage.QoS)
    
    Exit Sub
EH:
    Debug.Print "HandlePublish Error: " & Err.Description
End Sub

Private Sub HandlePubAck(oSession As cMQTTSession, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim lPacketId As Long
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    oSession.UpdateActivity
    
    oPacket.DecodeFixedHeader baData, lOffset, 0
    If Not oPacket.DecodePayload(baData, lOffset) Then Exit Sub
    
    lPacketId = oPacket.PacketId
    
    ' 从待确认队列中移除
    oSession.RemoveOutgoingMessage lPacketId
    
    Debug.Print "PUBACK received for packet " & lPacketId
    
    Exit Sub
EH:
    Debug.Print "HandlePubAck Error: " & Err.Description
End Sub

Private Sub HandlePubRec(oSession As cMQTTSession, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim baResponse() As Byte
    Dim lPacketId As Long
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    oSession.UpdateActivity
    
    oPacket.DecodeFixedHeader baData, lOffset, 0
    If Not oPacket.DecodePayload(baData, lOffset) Then Exit Sub
    
    lPacketId = oPacket.PacketId
    
    ' 发送 PUBREL
    baResponse = oPacket.EncodePubRel(lPacketId)
    oSession.Socket.SendData baResponse
    
    Debug.Print "PUBREC received for packet " & lPacketId
    
    Exit Sub
EH:
    Debug.Print "HandlePubRec Error: " & Err.Description
End Sub

Private Sub HandlePubRel(oSession As cMQTTSession, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim baResponse() As Byte
    Dim lPacketId As Long
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    oSession.UpdateActivity
    
    oPacket.DecodeFixedHeader baData, lOffset, 0
    If Not oPacket.DecodePayload(baData, lOffset) Then Exit Sub
    
    lPacketId = oPacket.PacketId
    
    ' 发送 PUBCOMP
    baResponse = oPacket.EncodePubComp(lPacketId)
    oSession.Socket.SendData baResponse
    
    oSession.RemovePendingPubRel lPacketId
    
    Debug.Print "PUBREL received for packet " & lPacketId
    
    Exit Sub
EH:
    Debug.Print "HandlePubRel Error: " & Err.Description
End Sub

Private Sub HandlePubComp(oSession As cMQTTSession, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim lPacketId As Long
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    oSession.UpdateActivity
    
    oPacket.DecodeFixedHeader baData, lOffset, 0
    If Not oPacket.DecodePayload(baData, lOffset) Then Exit Sub
    
    lPacketId = oPacket.PacketId
    
    ' QoS 2 完成
    oSession.RemoveOutgoingMessage lPacketId
    
    Debug.Print "PUBCOMP received for packet " & lPacketId
    
    Exit Sub
EH:
    Debug.Print "HandlePubComp Error: " & Err.Description
End Sub

Private Sub HandleSubscribe(oSession As cMQTTSession, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim oResponse As New cMQTTPacket
    Dim oFilter As cMQTTTopicFilter
    Dim vCode As Variant
    Dim baResponse() As Byte
    Dim i As Long
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    
    oSession.UpdateActivity
    
    ' 解析 SUBSCRIBE 报文
    Set oPacket.PacketType = mqttSubscribe
    oPacket.DecodeFixedHeader baData, lOffset, 0
    If Not oPacket.DecodePayload(baData, lOffset) Then Exit Sub
    
    ' 构建 SUBACK 响应
    Set oResponse.PacketType = mqttSubAck
    oResponse.PacketId = oPacket.PacketId
    Set oResponse.ReturnCodes = New Collection
    
    ' 处理每个订阅
    For Each oFilter In oPacket.TopicFilters
        ' 添加订阅
        oSession.AddSubscription oFilter.Topic, oFilter.QoS
        
        ' 返回成功代码（限制为服务端支持的最大 QoS）
        oResponse.ReturnCodes.Add IIf(oFilter.QoS > mqttQoS2, mqttSubSuccessQoS2, oFilter.QoS)
        
        RaiseEvent Subscribe(oSession.ClientId, oFilter.Topic, oFilter.QoS)
        Debug.Print "Client " & oSession.ClientId & " subscribed to " & oFilter.Topic & " (QoS " & oFilter.QoS & ")"
        
        ' 发送保留消息
        SendRetainedMessage oSession, oFilter.Topic
    Next
    
    ' 发送 SUBACK
    baResponse = oResponse.EncodeSubAck
    oSession.Socket.SendData baResponse
    
    Exit Sub
EH:
    Debug.Print "HandleSubscribe Error: " & Err.Description
End Sub

Private Sub HandleUnsubscribe(oSession As cMQTTSession, baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim oFilter As cMQTTTopicFilter
    Dim baResponse() As Byte
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    
    oSession.UpdateActivity
    
    ' 解析 UNSUBSCRIBE 报文
    Set oPacket.PacketType = mqttUnsubscribe
    oPacket.DecodeFixedHeader baData, lOffset, 0
    If Not oPacket.DecodePayload(baData, lOffset) Then Exit Sub
    
    ' 处理每个取消订阅
    For Each oFilter In oPacket.TopicFilters
        oSession.RemoveSubscription oFilter.Topic
        RaiseEvent Unsubscribe(oSession.ClientId, oFilter.Topic)
        Debug.Print "Client " & oSession.ClientId & " unsubscribed from " & oFilter.Topic
    Next
    
    ' 发送 UNSUBACK
    baResponse = oPacket.EncodeUnsubAck(oPacket.PacketId)
    oSession.Socket.SendData baResponse
    
    Exit Sub
EH:
    Debug.Print "HandleUnsubscribe Error: " & Err.Description
End Sub

Private Sub HandlePingReq(oSession As cMQTTSession)
    Dim oPacket As New cMQTTPacket
    Dim baResponse() As Byte
    
    On Error GoTo EH
    
    If oSession Is Nothing Then Exit Sub
    
    oSession.UpdateActivity
    
    ' 发送 PINGRESP
    baResponse = oPacket.EncodePingResp
    oSession.Socket.SendData baResponse
    
    Exit Sub
EH:
    Debug.Print "HandlePingReq Error: " & Err.Description
End Sub

Private Sub HandleDisconnect(oSession As cMQTTSession)
    If oSession Is Nothing Then Exit Sub
    
    Debug.Print "Client " & oSession.ClientId & " disconnected gracefully"
    DisconnectClient oSession, True
End Sub

'=========================================================================
' 消息路由
'=========================================================================

Private Sub RouteMessage(oMessage As cMQTTMessage)
    Dim oSession As cMQTTSession
    Dim lPacketId As Long
    Dim baPacket() As Byte
    Dim oPubPacket As New cMQTTPacket
    
    On Error Resume Next
    
    ' 遍历所有会话
    For Each oSession In m_oSessions
        If Not oSession.Socket Is Nothing And oSession.IsConnected Then
            ' 检查是否匹配订阅
            If oSession.TopicMatches(oMessage.Topic) Then
                ' 确定有效的 QoS (min(message QoS, subscription QoS))
                Dim eEffectiveQoS As MqttQoS
                Dim eSubQoS As MqttQoS
                
                eSubQoS = oSession.GetMatchingQoS(oMessage.Topic)
                If eSubQoS >= 0 Then
                    eEffectiveQoS = IIf(oMessage.QoS < eSubQoS, oMessage.QoS, eSubQoS)
                    
                    ' 构建 PUBLISH 报文
                    Set oPubPacket = New cMQTTPacket
                    oPubPacket.PacketType = mqttPublish
                    oPubPacket.TopicName = oMessage.Topic
                    oPubPacket.PayloadBytes = oMessage.PayloadBytes
                    oPubPacket.QoS = eEffectiveQoS
                    oPubPacket.Retain = False  ' 路由的消息不保留
                    oPubPacket.Duplicate = oMessage.Duplicate
                    
                    If eEffectiveQoS > 0 Then
                        lPacketId = oSession.GetNextPacketId
                        oPubPacket.PacketId = lPacketId
                        oSession.AddOutgoingMessage lPacketId, oMessage
                    End If
                    
                    baPacket = oPubPacket.EncodePublish
                    oSession.Socket.SendData baPacket
                    
                    Debug.Print "Message routed to " & oSession.ClientId & " on topic " & oMessage.Topic
                End If
            End If
        End If
    Next
End Sub

'=========================================================================
' 保留消息管理
'=========================================================================

Private Sub SaveRetainedMessage(oMessage As cMQTTMessage)
    On Error Resume Next
    
    ' 如果消息为空，删除保留消息
    If (Not oMessage.PayloadBytes) = -1 Or UBound(oMessage.PayloadBytes) < 0 Then
        m_oRetainedMessages.Remove oMessage.Topic
        Exit Sub
    End If
    
    ' 保存或更新保留消息
    m_oRetainedMessages.Remove oMessage.Topic
    m_oRetainedMessages.Add oMessage, oMessage.Topic
End Sub

Private Sub SendRetainedMessage(oSession As cMQTTSession, ByVal sTopicFilter As String)
    Dim oMessage As cMQTTMessage
    Dim oPubPacket As New cMQTTPacket
    Dim baPacket() As Byte
    Dim lPacketId As Long
    
    On Error Resume Next
    
    ' 遍历所有保留消息，匹配主题过滤器
    For Each oMessage In m_oRetainedMessages
        If MatchTopicFilter(oMessage.Topic, sTopicFilter) Then
            ' 构建 PUBLISH 报文
            Set oPubPacket = New cMQTTPacket
            oPubPacket.PacketType = mqttPublish
            oPubPacket.TopicName = oMessage.Topic
            oPubPacket.PayloadBytes = oMessage.PayloadBytes
            oPubPacket.QoS = oMessage.QoS
            oPubPacket.Retain = True
            oPubPacket.Duplicate = False
            
            If oMessage.QoS > 0 Then
                lPacketId = oSession.GetNextPacketId
                oPubPacket.PacketId = lPacketId
                oSession.AddOutgoingMessage lPacketId, oMessage
            End If
            
            baPacket = oPubPacket.EncodePublish
            oSession.Socket.SendData baPacket
            
            Debug.Print "Retained message sent to " & oSession.ClientId & " on topic " & oMessage.Topic
        End If
    Next
End Sub

'=========================================================================
' 辅助方法
'=========================================================================

Private Function FindSessionBySocket(oSocket As cWinsock) As cMQTTSession
    Dim oSession As cMQTTSession
    
    For Each oSession In m_oSessions
        If oSession.Socket Is oSocket Then
            Set FindSessionBySocket = oSession
            Exit Function
        End If
    Next
End Function

Private Sub DisconnectClient(oSession As cMQTTSession, ByVal bGraceful As Boolean)
    On Error Resume Next
    
    If Not oSocket Is Nothing Then
        oSession.Socket.Close_
    End If
    
    oSession.IsConnected = False
    
    If oSession.CleanSession Then
        m_oSessions.Remove oSession.ClientId
    End If
    
    RaiseEvent ClientDisconnected(oSession.ClientId)
End Sub

Private Function MatchTopicFilter(ByVal sTopic As String, ByVal sFilter As String) As Boolean
    Dim vTopicLevels As Variant
    Dim vFilterLevels As Variant
    Dim i As Long
    
    vTopicLevels = Split(sTopic, "/")
    vFilterLevels = Split(sFilter, "/")
    
    i = 0
    Do While i < UBound(vFilterLevels) + 1
        If vFilterLevels(i) = "#" Then
            MatchTopicFilter = True
            Exit Function
        End If
        
        If i >= UBound(vTopicLevels) + 1 Then
            MatchTopicFilter = False
            Exit Function
        End If
        
        If vFilterLevels(i) <> "+" Then
            If vFilterLevels(i) <> vTopicLevels(i) Then
                MatchTopicFilter = False
                Exit Function
            End If
        End If
        
        i = i + 1
    Loop
    
    MatchTopicFilter = (i = UBound(vTopicLevels) + 1)
End Function
