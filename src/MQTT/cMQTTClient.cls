'=========================================================================
' cMQTTClient - MQTT Client
' MQTT 客户端实现
'=========================================================================
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMQTTClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const MODULE_NAME As String = "cMQTTClient"

' 事件
Event Connected()
Event Disconnected()
Event MessageArrived(Topic As String, Payload As Variant, QoS As MqttQoS, Retain As Boolean)
Event Subscribed(TopicFilter As String, QoS As MqttQoS)
Event Unsubscribed(TopicFilter As String)
Event ErrorEvent(Number As Long, Description As String)

' 网络层
Private WithEvents m_oSocket As cWinsock
Attribute m_oSocket.VB_VarHelpID = -1

' 客户端状态
Public ClientId As String
Public UserName As String
Public Password As String
Public CleanSession As Boolean
Public KeepAlive As Long
Public WillTopic As String
Public WillMessage As String
Public WillQoS As MqttQoS
Public WillRetain As Boolean

Private m_sServerHost As String
Private m_lServerPort As Long
Private m_bConnected As Boolean
Private m_bConnecting As Boolean

' 报文编解码器
Private m_oPacket As cMQTTPacket

' 订阅列表
Private m_oSubscriptions As Collection

' 待确认消息 (QoS 1/2)
Private m_oPendingPub As Collection
Private m_oPendingPubRel As Collection

' 下一个报文标识符
Private m_lNextPacketId As Long

' 接收缓冲区
Private m_baRecvBuffer() As Byte

'=========================================================================
' 初始化
'=========================================================================

Private Sub Class_Initialize()
    Set m_oSocket = New cWinsock
    Set m_oPacket = New cMQTTPacket
    Set m_oSubscriptions = New Collection
    Set m_oPendingPub = New Collection
    Set m_oPendingPubRel = New Collection
    
    m_lServerPort = MQTT_DEFAULT_PORT
    KeepAlive = 60
    CleanSession = True
    m_lNextPacketId = 1
    m_baRecvBuffer = vbNullString
End Sub

Private Sub Class_Terminate()
    Disconnect
End Sub

'=========================================================================
' 公共方法
'=========================================================================

' 连接到 MQTT 服务器
Public Function Connect(ByVal ServerHost As String, _
                        Optional ByVal ServerPort As Long = 0, _
                        Optional ByVal Timeout As Long = 5000) As Boolean
    On Error GoTo EH
    
    If m_bConnected Or m_bConnecting Then
        Connect = True
        Exit Function
    End If
    
    m_sServerHost = ServerHost
    If ServerPort > 0 Then
        m_lServerPort = ServerPort
    End If
    
    ' 生成默认 Client ID
    If Len(ClientId) = 0 Then
        ClientId = "VBClient_" & Timer & "_" & Rnd * 10000
    End If
    
    m_bConnecting = True
    
    ' 建立 TCP 连接
    m_oSocket.Protocol = wskTCPProtocol
    If Not m_oSocket.SyncConnect(m_sServerHost, m_lServerPort, Timeout) Then
        GoTo CleanUp
    End If
    
    ' 发送 CONNECT 报文
    If Not SendConnectPacket Then
        GoTo CleanUp
    End If
    
    ' 等待 CONNACK
    If Not WaitForConnAck(Timeout) Then
        GoTo CleanUp
    End If
    
    m_bConnected = True
    m_bConnecting = False
    RaiseEvent Connected
    Connect = True
    Exit Function
    
CleanUp:
    m_oSocket.Close_
    m_bConnecting = False
    Exit Function
EH:
    RaiseEvent ErrorEvent(Err.Number, "Connect failed: " & Err.Description)
    m_bConnecting = False
End Function

' 断开连接
Public Sub Disconnect()
    Dim oPacket As New cMQTTPacket
    Dim baPacket() As Byte
    
    On Error Resume Next
    
    If m_bConnected Then
        ' 发送 DISCONNECT
        baPacket = oPacket.EncodeDisconnect
        m_oSocket.SendData baPacket
        
        m_bConnected = False
    End If
    
    m_oSocket.Close_
    Set m_oSubscriptions = New Collection
    Set m_oPendingPub = New Collection
    Set m_oPendingPubRel = New Collection
    m_baRecvBuffer = vbNullString
    
    RaiseEvent Disconnected
End Sub

' 订阅主题
Public Sub Subscribe(ByVal TopicFilter As String, Optional ByVal QoS As MqttQoS = mqttQoS0)
    Dim oPacket As New cMQTTPacket
    Dim oFilter As New cMQTTTopicFilter
    Dim baPacket() As Byte
    Dim lPacketId As Long
    
    On Error GoTo EH
    
    If Not m_bConnected Then
        RaiseEvent ErrorEvent(mqttErrNotConnected, "Not connected")
        Exit Sub
    End If
    
    lPacketId = GetNextPacketId
    
    ' 构建 SUBSCRIBE 报文
    oPacket.PacketType = mqttSubscribe
    oPacket.PacketId = lPacketId
    Set oPacket.TopicFilters = New Collection
    
    oFilter.Topic = TopicFilter
    oFilter.QoS = QoS
    oPacket.TopicFilters.Add oFilter
    
    baPacket = oPacket.EncodeSubscribe
    m_oSocket.SendData baPacket
    
    ' 等待确认
    ' (简化处理，实际应该异步等待 SUBACK)
    
    ' 保存订阅
    On Error Resume Next
    m_oSubscriptions.Remove TopicFilter
    On Error GoTo EH
    m_oSubscriptions.Add QoS, TopicFilter
    
    RaiseEvent Subscribed(TopicFilter, QoS)
    
    Exit Sub
EH:
    RaiseEvent ErrorEvent(Err.Number, "Subscribe failed: " & Err.Description)
End Sub

' 取消订阅
Public Sub Unsubscribe(ByVal TopicFilter As String)
    Dim oPacket As New cMQTTPacket
    Dim oFilter As New cMQTTTopicFilter
    Dim baPacket() As Byte
    Dim lPacketId As Long
    
    On Error GoTo EH
    
    If Not m_bConnected Then
        RaiseEvent ErrorEvent(mqttErrNotConnected, "Not connected")
        Exit Sub
    End If
    
    lPacketId = GetNextPacketId
    
    ' 构建 UNSUBSCRIBE 报文
    oPacket.PacketType = mqttUnsubscribe
    oPacket.PacketId = lPacketId
    Set oPacket.TopicFilters = New Collection
    
    oFilter.Topic = TopicFilter
    oPacket.TopicFilters.Add oFilter
    
    baPacket = oPacket.EncodeUnsubscribe
    m_oSocket.SendData baPacket
    
    ' 移除订阅
    On Error Resume Next
    m_oSubscriptions.Remove TopicFilter
    
    RaiseEvent Unsubscribed(TopicFilter)
    
    Exit Sub
EH:
    RaiseEvent ErrorEvent(Err.Number, "Unsubscribe failed: " & Err.Description)
End Sub

' 发布消息
Public Sub Publish(ByVal Topic As String, _
                     ByVal Payload As Variant, _
                     Optional ByVal QoS As MqttQoS = mqttQoS0, _
                     Optional ByVal Retain As Boolean = False)
    Dim oPacket As New cMQTTPacket
    Dim oMessage As New cMQTTMessage
    Dim baPacket() As Byte
    Dim lPacketId As Long
    
    On Error GoTo EH
    
    If Not m_bConnected Then
        RaiseEvent ErrorEvent(mqttErrNotConnected, "Not connected")
        Exit Sub
    End If
    
    ' 构建消息
    oMessage.Topic = Topic
    If VarType(Payload) = vbByte + vbArray Then
        oMessage.PayloadBytes = Payload
    Else
        oMessage.Payload = CStr(Payload)
    End If
    oMessage.QoS = QoS
    oMessage.Retain = Retain
    
    ' 构建 PUBLISH 报文
    oPacket.PacketType = mqttPublish
    oPacket.TopicName = Topic
    If VarType(Payload) = vbByte + vbArray Then
        oPacket.PayloadBytes = Payload
    Else
        oPacket.Payload = CStr(Payload)
    End If
    oPacket.QoS = QoS
    oPacket.Retain = Retain
    oPacket.Duplicate = False
    
    If QoS > 0 Then
        lPacketId = GetNextPacketId
        oPacket.PacketId = lPacketId
        
        ' 保存到待确认队列
        Dim oPending As New cMQTTPendingMessage
        oPending.PacketId = lPacketId
        Set oPending.Message = oMessage
        oPending.Timestamp = Now
        m_oPendingPub.Add oPending, "P" & lPacketId
    End If
    
    baPacket = oPacket.EncodePublish
    m_oSocket.SendData baPacket
    
    Exit Sub
EH:
    RaiseEvent ErrorEvent(Err.Number, "Publish failed: " & Err.Description)
End Sub

' 发送 PING 保活
Public Sub Ping()
    Dim oPacket As New cMQTTPacket
    Dim baPacket() As Byte
    
    On Error GoTo EH
    
    If Not m_bConnected Then Exit Sub
    
    baPacket = oPacket.EncodePingReq
    m_oSocket.SendData baPacket
    
    Exit Sub
EH:
    RaiseEvent ErrorEvent(Err.Number, "Ping failed: " & Err.Description)
End Sub

'=========================================================================
' 属性
'=========================================================================

Public Property Get IsConnected() As Boolean
    IsConnected = m_bConnected
End Property

Public Property Get ServerHost() As String
    ServerHost = m_sServerHost
End Property

Public Property Get ServerPort() As Long
    ServerPort = m_lServerPort
End Property

'=========================================================================
' 套接字事件处理
'=========================================================================

Private Sub m_oSocket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    Dim lRemainingOffset As Long
    Dim lRemainingLength As Long
    Dim oPacket As New cMQTTPacket
    
    On Error GoTo EH
    
    ' 读取数据
    Client.GetData baData, vbByte + vbArray
    
    ' 追加到缓冲区
    If UBound(m_baRecvBuffer) >= 0 Then
        Dim lOldLen As Long
        lOldLen = UBound(m_baRecvBuffer) + 1
        ReDim Preserve m_baRecvBuffer(0 To lOldLen + UBound(baData))
        CopyMemory m_baRecvBuffer(lOldLen), baData(0), UBound(baData) + 1
    Else
        m_baRecvBuffer = baData
    End If
    
    ' 处理缓冲区中的完整报文
    Do While UBound(m_baRecvBuffer) >= 1
        If Not oPacket.DecodeFixedHeader(m_baRecvBuffer, lRemainingOffset, lRemainingLength) Then Exit Do
        
        ' 检查是否收到完整报文
        If UBound(m_baRecvBuffer) + 1 < lRemainingOffset + lRemainingLength Then Exit Do
        
        ' 提取完整报文
        Dim baFullPacket() As Byte
        Dim lPacketLen As Long
        lPacketLen = lRemainingOffset + lRemainingLength
        ReDim baFullPacket(0 To lPacketLen - 1)
        CopyMemory baFullPacket(0), m_baRecvBuffer(0), lPacketLen
        
        ' 从缓冲区移除已处理的报文
        If UBound(m_baRecvBuffer) + 1 > lPacketLen Then
            Dim lRemaining As Long
            lRemaining = UBound(m_baRecvBuffer) + 1 - lPacketLen
            Dim baNewBuffer() As Byte
            ReDim baNewBuffer(0 To lRemaining - 1)
            CopyMemory baNewBuffer(0), m_baRecvBuffer(lPacketLen), lRemaining
            m_baRecvBuffer = baNewBuffer
        Else
            m_baRecvBuffer = vbNullString
        End If
        
        ' 处理报文
        HandleIncomingPacket baFullPacket, lRemainingOffset
    Loop
    
    Exit Sub
EH:
    Debug.Print "DataArrival Error: " & Err.Description
End Sub

Private Sub m_oSocket_CloseEvent(Client As cWinsock)
    If m_bConnected Then
        m_bConnected = False
        RaiseEvent Disconnected
    End If
End Sub

Private Sub m_oSocket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    RaiseEvent ErrorEvent(Number, "Socket error: " & Description)
    If m_bConnected Then
        m_bConnected = False
        RaiseEvent Disconnected
    End If
End Sub

'=========================================================================
' 私有方法
'=========================================================================

Private Function SendConnectPacket() As Boolean
    Dim oPacket As New cMQTTPacket
    Dim baPacket() As Byte
    
    On Error GoTo EH
    
    oPacket.ClientId = ClientId
    oPacket.CleanSession = CleanSession
    oPacket.KeepAlive = KeepAlive
    oPacket.UserName = UserName
    oPacket.Password = Password
    oPacket.WillTopic = WillTopic
    oPacket.WillMessage = WillMessage
    oPacket.WillQoS = WillQoS
    oPacket.WillRetain = WillRetain
    
    baPacket = oPacket.EncodeConnect
    m_oSocket.SendData baPacket
    
    SendConnectPacket = True
    Exit Function
EH:
    SendConnectPacket = False
End Function

Private Function WaitForConnAck(ByVal Timeout As Long) As Boolean
    Dim baBuffer(0 To 3) As Byte
    Dim oPacket As New cMQTTPacket
    Dim lReceived As Long
    Dim lStart As Long
    
    lStart = Timer * 1000
    
    Do While (Timer * 1000 - lStart) < Timeout
        If m_oSocket.BytesReceived >= 4 Then
            m_oSocket.GetData baBuffer, vbByte + vbArray, 4
            
            ' 解析 CONNACK
            If baBuffer(0) = &H20 And baBuffer(1) = &H2 Then
                If baBuffer(3) = 0 Then
                    WaitForConnAck = True
                    Exit Function
                Else
                    Debug.Print "CONNACK return code: " & baBuffer(3)
                    Exit Function
                End If
            End If
        End If
        
        DoEvents
        Sleep 10
    Loop
    
    Exit Function
End Function

Private Sub HandleIncomingPacket(baData() As Byte, ByVal lOffset As Long)
    Dim oPacket As New cMQTTPacket
    Dim lRemainingOffset As Long
    Dim lRemainingLength As Long
    
    On Error GoTo EH
    
    oPacket.DecodeFixedHeader baData, lRemainingOffset, lRemainingLength
    oPacket.DecodePayload baData, lRemainingOffset
    
    Select Case oPacket.PacketType
        Case mqttConnAck
            ' Already handled in WaitForConnAck
        Case mqttPublish
            HandlePublishPacket oPacket
        Case mqttPubAck
            m_oPendingPub.Remove "P" & oPacket.PacketId
        Case mqttPubRec
            ' QoS 2: Send PUBREL
            Dim oPubRel As New cMQTTPacket
            Dim baPubRel() As Byte
            baPubRel = oPubRel.EncodePubRel(oPacket.PacketId)
            m_oSocket.SendData baPubRel
        Case mqttPubRel
            ' QoS 2: Send PUBCOMP
            Dim oPubComp As New cMQTTPacket
            Dim baPubComp() As Byte
            baPubComp = oPubComp.EncodePubComp(oPacket.PacketId)
            m_oSocket.SendData baPubComp
        Case mqttPubComp
            m_oPendingPub.Remove "P" & oPacket.PacketId
        Case mqttSubAck
            ' Handle subscription acknowledgment
        Case mqttUnsubAck
            ' Handle unsubscribe acknowledgment
        Case mqttPingResp
            ' Ping response received
        Case Else
            Debug.Print "Unhandled packet type: " & oPacket.PacketType
    End Select
    
    Exit Sub
EH:
    Debug.Print "HandleIncomingPacket Error: " & Err.Description
End Sub

Private Sub HandlePublishPacket(oPacket As cMQTTPacket)
    Dim oAck As New cMQTTPacket
    Dim baAck() As Byte
    Dim vPayload As Variant
    
    On Error GoTo EH
    
    ' QoS 1: Send PUBACK
    If oPacket.QoS = mqttQoS1 Then
        baAck = oAck.EncodePubAck(oPacket.PacketId)
        m_oSocket.SendData baAck
    End If
    
    ' QoS 2: Send PUBREC
    If oPacket.QoS = mqttQoS2 Then
        baAck = oAck.EncodePubRec(oPacket.PacketId)
        m_oSocket.SendData baAck
    End If
    
    ' 触发消息到达事件
    If (Not oPacket.PayloadBytes) <> -1 Then
        vPayload = oPacket.PayloadBytes
    Else
        vPayload = oPacket.Payload
    End If
    
    RaiseEvent MessageArrived(oPacket.TopicName, vPayload, oPacket.QoS, oPacket.Retain)
    
    Exit Sub
EH:
    Debug.Print "HandlePublishPacket Error: " & Err.Description
End Sub

Private Function GetNextPacketId() As Long
    m_lNextPacketId = m_lNextPacketId + 1
    If m_lNextPacketId > 65535 Then
        m_lNextPacketId = 1
    End If
    GetNextPacketId = m_lNextPacketId
End Function

Private Sub CopyMemory(Destination As Any, Source As Any, ByVal Length As Long)
    RtlMoveMemory Destination, Source, Length
End Sub

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
