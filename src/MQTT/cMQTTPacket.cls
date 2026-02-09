'=========================================================================
' cMQTTPacket - MQTT Packet Encoder/Decoder
' 负责 MQTT 报文的编码和解码
'=========================================================================
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMQTTPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const MODULE_NAME As String = "cMQTTPacket"

'=========================================================================
' 报文结构
'=========================================================================
Public PacketType     As MqttPacketType
Public Flags          As Long
Public RemainingLen   As Long

' 连接报文字段
Public ProtocolName   As String
Public ProtocolLevel  As Long
Public ConnectFlags   As Long
Public KeepAlive      As Long
Public ClientId       As String
Public WillTopic      As String
Public WillMessage    As String
Public WillQoS        As MqttQoS
Public WillRetain     As Boolean
Public UserName       As String
Public Password       As String
Public CleanSession   As Boolean

' 发布报文字段
Public TopicName      As String
Public PacketId       As Long
Public Payload        As String
Public PayloadBytes() As Byte
Public QoS            As MqttQoS
Public Duplicate      As Boolean
Public Retain         As Boolean

' 订阅报文字段
Public TopicFilters   As Collection    ' of cMQTTTopicFilter

' 确认报文字段
Public ReturnCode     As Long
Public ReturnCodes    As Collection    ' for SUBACK

'=========================================================================
' 编码方法
'=========================================================================

' 编码 CONNECT 报文
Public Function EncodeConnect() As Byte()
    Dim baPayload() As Byte
    Dim baVariable() As Byte
    Dim baFixed() As Byte
    Dim baResult() As Byte
    Dim lLen As Long
    Dim lPos As Long
    
    ' 构建可变报头 (10 bytes)
    ReDim baVariable(0 To 9)
    
    ' Protocol Name: "MQTT" (4 bytes + 2 length)
    lPos = 0
    baVariable(lPos) = 0: lPos = lPos + 1
    baVariable(lPos) = 4: lPos = lPos + 1
    CopyMemory baVariable(lPos), ByVal "MQTT", 4: lPos = lPos + 4
    
    ' Protocol Level: 4 (MQTT 3.1.1)
    baVariable(lPos) = MQTT_PROTOCOL_LEVEL_311: lPos = lPos + 1
    
    ' Connect Flags
    baVariable(lPos) = BuildConnectFlags(): lPos = lPos + 1
    
    ' Keep Alive (2 bytes, MSB first)
    baVariable(lPos) = (KeepAlive \\ 256) And &HFF: lPos = lPos + 1
    baVariable(lPos) = KeepAlive And &HFF: lPos = lPos + 1
    
    ' 构建有效载荷
    baPayload = BuildConnectPayload()
    
    ' 计算剩余长度
    lLen = UBound(baVariable) + 1 + UBound(baPayload) + 1
    
    ' 构建固定报头
    baFixed = BuildFixedHeader(mqttConnect, 0, lLen)
    
    ' 合并所有部分
    ReDim baResult(0 To UBound(baFixed) + UBound(baVariable) + UBound(baPayload) + 2)
    CopyMemory baResult(0), baFixed(0), UBound(baFixed) + 1
    CopyMemory baResult(UBound(baFixed) + 1), baVariable(0), UBound(baVariable) + 1
    CopyMemory baResult(UBound(baFixed) + UBound(baVariable) + 2), baPayload(0), UBound(baPayload) + 1
    
    EncodeConnect = baResult
End Function

' 编码 CONNACK 报文
Public Function EncodeConnAck(ByVal ReturnCode_ As MqttConnectReturnCode, _n                              Optional ByVal SessionPresent As Boolean = False) As Byte()
    Dim baResult(0 To 3) As Byte
    
    ' Fixed header: 0010 0000 = 0x20, remaining length = 2
    baResult(0) = &H20
    baResult(1) = 2
    
    ' Variable header: Session Present (1 bit) + reserved (7 bits)
    baResult(2) = IIf(SessionPresent, 1, 0)
    
    ' Return code
    baResult(3) = ReturnCode_
    
    EncodeConnAck = baResult
End Function

' 编码 PUBLISH 报文
Public Function EncodePublish() As Byte()
    Dim baTopic() As Byte
    Dim baPayloadData() As Byte
    Dim baVariable() As Byte
    Dim baFixed() As Byte
    Dim baResult() As Byte
    Dim lLen As Long
    Dim lPos As Long
    
    ' 主题编码为 UTF-8
    baTopic = EncodeUTF8String(TopicName)
    
    ' 计算可变报头长度
    lLen = UBound(baTopic) + 1
    If QoS > 0 Then
        lLen = lLen + 2  ' Packet ID
    End If
    
    ' 构建可变报头
    ReDim baVariable(0 To lLen - 1)
    CopyMemory baVariable(0), baTopic(0), UBound(baTopic) + 1
    lPos = UBound(baTopic) + 1
    
    If QoS > 0 Then
        baVariable(lPos) = (PacketId \\ 256) And &HFF: lPos = lPos + 1
        baVariable(lPos) = PacketId And &HFF
    End If
    
    ' 有效载荷
    If (Not PayloadBytes) <> -1 Then
        baPayloadData = PayloadBytes
    Else
        baPayloadData = EncodeUTF8String(Payload)
    End If
    
    ' 计算总长度
    lLen = UBound(baVariable) + 1 + UBound(baPayloadData) + 1
    
    ' 构建固定报头
    Dim lFlags As Long
    lFlags = BuildPublishFlags()
    baFixed = BuildFixedHeader(mqttPublish, lFlags, lLen)
    
    ' 合并
    ReDim baResult(0 To UBound(baFixed) + UBound(baVariable) + UBound(baPayloadData) + 2)
    CopyMemory baResult(0), baFixed(0), UBound(baFixed) + 1
    CopyMemory baResult(UBound(baFixed) + 1), baVariable(0), UBound(baVariable) + 1
    CopyMemory baResult(UBound(baFixed) + UBound(baVariable) + 2), baPayloadData(0), UBound(baPayloadData) + 1
    
    EncodePublish = baResult
End Function

' 编码 PUBACK 报文
Public Function EncodePubAck(ByVal PacketId_ As Long) As Byte()
    Dim baResult(0 To 3) As Byte
    baResult(0) = &H40  ' PUBACK type (0100 0000)
    baResult(1) = 2     ' Remaining length
    baResult(2) = (PacketId_ \\ 256) And &HFF
    baResult(3) = PacketId_ And &HFF
    EncodePubAck = baResult
End Function

' 编码 PUBREC 报文
Public Function EncodePubRec(ByVal PacketId_ As Long) As Byte()
    Dim baResult(0 To 3) As Byte
    baResult(0) = &H50  ' PUBREC type (0101 0000)
    baResult(1) = 2
    baResult(2) = (PacketId_ \\ 256) And &HFF
    baResult(3) = PacketId_ And &HFF
    EncodePubRec = baResult
End Function

' 编码 PUBREL 报文
Public Function EncodePubRel(ByVal PacketId_ As Long) As Byte()
    Dim baResult(0 To 3) As Byte
    baResult(0) = &H62  ' PUBREL type (0110 0010) - QoS 1
    baResult(1) = 2
    baResult(2) = (PacketId_ \\ 256) And &HFF
    baResult(3) = PacketId_ And &HFF
    EncodePubRel = baResult
End Function

' 编码 PUBCOMP 报文
Public Function EncodePubComp(ByVal PacketId_ As Long) As Byte()
    Dim baResult(0 To 3) As Byte
    baResult(0) = &H70  ' PUBCOMP type (0111 0000)
    baResult(1) = 2
    baResult(2) = (PacketId_ \\ 256) And &HFF
    baResult(3) = PacketId_ And &HFF
    EncodePubComp = baResult
End Function

' 编码 SUBSCRIBE 报文
Public Function EncodeSubscribe() As Byte()
    Dim baVariable() As Byte
    Dim baPayload() As Byte
    Dim baFixed() As Byte
    Dim baResult() As Byte
    Dim oFilter As cMQTTTopicFilter
    Dim lLen As Long
    Dim lPos As Long
    
    ' 可变报头: Packet ID
    ReDim baVariable(0 To 1)
    baVariable(0) = (PacketId \\ 256) And &HFF
    baVariable(1) = PacketId And &HFF
    
    ' 有效载荷: 主题过滤器列表
    lLen = 0
    For Each oFilter In TopicFilters
        lLen = lLen + 2 + Len(oFilter.Topic) + 1  ' length(2) + topic + qos(1)
    Next
    ReDim baPayload(0 To lLen - 1)
    
    lPos = 0
    For Each oFilter In TopicFilters
        ' Topic
        Dim baTopic() As Byte
        baTopic = EncodeUTF8String(oFilter.Topic)
        CopyMemory baPayload(lPos), baTopic(0), UBound(baTopic) + 1
        lPos = lPos + UBound(baTopic) + 1
        ' QoS
        baPayload(lPos) = oFilter.QoS
        lPos = lPos + 1
    Next
    
    ' 总长度
    lLen = UBound(baVariable) + 1 + UBound(baPayload) + 1
    
    ' 固定报头 (1000 0010 = 0x82, reserved = 0010)
    baFixed = BuildFixedHeader(mqttSubscribe, 2, lLen)
    
    ' 合并
    ReDim baResult(0 To UBound(baFixed) + UBound(baVariable) + UBound(baPayload) + 2)
    CopyMemory baResult(0), baFixed(0), UBound(baFixed) + 1
    CopyMemory baResult(UBound(baFixed) + 1), baVariable(0), UBound(baVariable) + 1
    CopyMemory baResult(UBound(baFixed) + UBound(baVariable) + 2), baPayload(0), UBound(baPayload) + 1
    
    EncodeSubscribe = baResult
End Function

' 编码 SUBACK 报文
Public Function EncodeSubAck() As Byte()
    Dim baVariable() As Byte
    Dim baPayload() As Byte
    Dim baFixed() As Byte
    Dim baResult() As Byte
    Dim vCode As Variant
    Dim lLen As Long
    Dim lPos As Long
    
    ' 可变报头: Packet ID
    ReDim baVariable(0 To 1)
    baVariable(0) = (PacketId \\ 256) And &HFF
    baVariable(1) = PacketId And &HFF
    
    ' 有效载荷: Return codes
    ReDim baPayload(0 To ReturnCodes.Count - 1)
    lPos = 0
    For Each vCode In ReturnCodes
        baPayload(lPos) = vCode
        lPos = lPos + 1
    Next
    
    ' 总长度
    lLen = UBound(baVariable) + 1 + UBound(baPayload) + 1
    
    ' 固定报头
    baFixed = BuildFixedHeader(mqttSubAck, 0, lLen)
    
    ' 合并
    ReDim baResult(0 To UBound(baFixed) + UBound(baVariable) + UBound(baPayload) + 2)
    CopyMemory baResult(0), baFixed(0), UBound(baFixed) + 1
    CopyMemory baResult(UBound(baFixed) + 1), baVariable(0), UBound(baVariable) + 1
    CopyMemory baResult(UBound(baFixed) + UBound(baVariable) + 2), baPayload(0), UBound(baPayload) + 1
    
    EncodeSubAck = baResult
End Function

' 编码 UNSUBSCRIBE 报文
Public Function EncodeUnsubscribe() As Byte()
    Dim baVariable() As Byte
    Dim baPayload() As Byte
    Dim baFixed() As Byte
    Dim baResult() As Byte
    Dim oFilter As cMQTTTopicFilter
    Dim lLen As Long
    Dim lPos As Long
    
    ' 可变报头: Packet ID
    ReDim baVariable(0 To 1)
    baVariable(0) = (PacketId \\ 256) And &HFF
    baVariable(1) = PacketId And &HFF
    
    ' 有效载荷: 主题过滤器列表 (只有主题，没有 QoS)
    lLen = 0
    For Each oFilter In TopicFilters
        lLen = lLen + 2 + Len(oFilter.Topic)
    Next
    ReDim baPayload(0 To lLen - 1)
    
    lPos = 0
    For Each oFilter In TopicFilters
        Dim baTopic() As Byte
        baTopic = EncodeUTF8String(oFilter.Topic)
        CopyMemory baPayload(lPos), baTopic(0), UBound(baTopic) + 1
        lPos = lPos + UBound(baTopic) + 1
    Next
    
    ' 总长度
    lLen = UBound(baVariable) + 1 + UBound(baPayload) + 1
    
    ' 固定报头 (1010 0010 = 0xA2, reserved = 0010)
    baFixed = BuildFixedHeader(mqttUnsubscribe, 2, lLen)
    
    ' 合并
    ReDim baResult(0 To UBound(baFixed) + UBound(baVariable) + UBound(baPayload) + 2)
    CopyMemory baResult(0), baFixed(0), UBound(baFixed) + 1
    CopyMemory baResult(UBound(baFixed) + 1), baVariable(0), UBound(baVariable) + 1
    CopyMemory baResult(UBound(baFixed) + UBound(baVariable) + 2), baPayload(0), UBound(baPayload) + 1
    
    EncodeUnsubscribe = baResult
End Function

' 编码 UNSUBACK 报文
Public Function EncodeUnsubAck(ByVal PacketId_ As Long) As Byte()
    Dim baResult(0 To 3) As Byte
    baResult(0) = &HB0  ' UNSUBACK type (1011 0000)
    baResult(1) = 2
    baResult(2) = (PacketId_ \\ 256) And &HFF
    baResult(3) = PacketId_ And &HFF
    EncodeUnsubAck = baResult
End Function

' 编码 PINGREQ 报文
Public Function EncodePingReq() As Byte()
    Dim baResult(0 To 1) As Byte
    baResult(0) = &HC0  ' PINGREQ type (1100 0000)
    baResult(1) = 0     ' No remaining length
    EncodePingReq = baResult
End Function

' 编码 PINGRESP 报文
Public Function EncodePingResp() As Byte()
    Dim baResult(0 To 1) As Byte
    baResult(0) = &HD0  ' PINGRESP type (1101 0000)
    baResult(1) = 0
    EncodePingResp = baResult
End Function

' 编码 DISCONNECT 报文
Public Function EncodeDisconnect() As Byte()
    Dim baResult(0 To 1) As Byte
    baResult(0) = &HE0  ' DISCONNECT type (1110 0000)
    baResult(1) = 0
    EncodeDisconnect = baResult
End Function

'=========================================================================
' 解码方法
'=========================================================================

' 解析固定报头，返回剩余长度在数据包中的偏移量
Public Function DecodeFixedHeader(baData() As Byte, ByRef OutRemainingOffset As Long, _n                                   ByRef OutRemainingLength As Long) As Boolean
    Dim lByte1 As Long
    Dim lMultiplier As Long
    Dim lValue As Long
    Dim lPos As Long
    
    If UBound(baData) < 0 Then Exit Function
    
    ' 第一个字节: 报文类型 + 标志
    lByte1 = baData(0)
    PacketType = (lByte1 \\ 16) And &HF
    Flags = lByte1 And &HF
    
    ' 解码剩余长度 (Variable Byte Integer)
    lMultiplier = 1
    lValue = 0
    lPos = 1
    Do
        If lPos > UBound(baData) Then Exit Function
        Dim lEncodedByte As Long
        lEncodedByte = baData(lPos)
        lValue = lValue + (lEncodedByte And 127) * lMultiplier
        lMultiplier = lMultiplier * 128
        lPos = lPos + 1
        If lMultiplier > 128 * 128 * 128 Then Exit Function  ' Malformed
    Loop While (lEncodedByte And 128) <> 0
    
    RemainingLen = lValue
    OutRemainingOffset = lPos
    OutRemainingLength = lValue
    DecodeFixedHeader = True
End Function

' 解码可变报头和有效载荷
Public Function DecodePayload(baData() As Byte, ByVal RemainingOffset As Long) As Boolean
    On Error GoTo EH
    
    Select Case PacketType
        Case mqttConnect
            DecodePayload = DecodeConnectPayload(baData, RemainingOffset)
        Case mqttPublish
            DecodePayload = DecodePublishPayload(baData, RemainingOffset)
        Case mqttSubscribe
            DecodePayload = DecodeSubscribePayload(baData, RemainingOffset)
        Case mqttUnsubscribe
            DecodePayload = DecodeUnsubscribePayload(baData, RemainingOffset)
        Case mqttPubAck, mqttPubRec, mqttPubRel, mqttPubComp
            DecodePayload = DecodePacketIdPayload(baData, RemainingOffset)
    End Select
    
    Exit Function
EH:
    DecodePayload = False
End Function

'=========================================================================
' 私有辅助方法
'=========================================================================

Private Function BuildConnectFlags() As Long
    Dim lFlags As Long
    
    lFlags = lFlags Or IIf(CleanSession, &H2, 0)
    lFlags = lFlags Or IIf(Len(WillTopic) > 0, &H4, 0)
    lFlags = lFlags Or ((WillQoS And &H3) * &H8)
    lFlags = lFlags Or IIf(WillRetain, &H20, 0)
    lFlags = lFlags Or IIf(Len(Password) > 0, &H40, 0)
    lFlags = lFlags Or IIf(Len(UserName) > 0, &H80, 0)
    
    BuildConnectFlags = lFlags
End Function

Private Function BuildConnectPayload() As Byte()
    Dim baResult() As Byte
    Dim lLen As Long
    Dim lPos As Long
    Dim baTemp() As Byte
    
    ' 计算总长度
    lLen = 2 + Len(ClientId)
    If Len(WillTopic) > 0 Then
        lLen = lLen + 2 + Len(WillTopic) + 2 + Len(WillMessage)
    End If
    If Len(UserName) > 0 Then
        lLen = lLen + 2 + Len(UserName)
    End If
    If Len(Password) > 0 Then
        lLen = lLen + 2 + Len(Password)
    End If
    
    ReDim baResult(0 To lLen - 1)
    
    ' Client ID
    baTemp = EncodeUTF8String(ClientId)
    CopyMemory baResult(lPos), baTemp(0), UBound(baTemp) + 1
    lPos = lPos + UBound(baTemp) + 1
    
    ' Will Topic & Message
    If Len(WillTopic) > 0 Then
        baTemp = EncodeUTF8String(WillTopic)
        CopyMemory baResult(lPos), baTemp(0), UBound(baTemp) + 1
        lPos = lPos + UBound(baTemp) + 1
        
        baTemp = EncodeUTF8String(WillMessage)
        CopyMemory baResult(lPos), baTemp(0), UBound(baTemp) + 1
        lPos = lPos + UBound(baTemp) + 1
    End If
    
    ' User Name
    If Len(UserName) > 0 Then
        baTemp = EncodeUTF8String(UserName)
        CopyMemory baResult(lPos), baTemp(0), UBound(baTemp) + 1
        lPos = lPos + UBound(baTemp) + 1
    End If
    
    ' Password
    If Len(Password) > 0 Then
        baTemp = EncodeUTF8String(Password)
        CopyMemory baResult(lPos), baTemp(0), UBound(baTemp) + 1
    End If
    
    BuildConnectPayload = baResult
End Function

Private Function BuildPublishFlags() As Long
    Dim lFlags As Long
    lFlags = lFlags Or IIf(Duplicate, MQTT_FLAG_DUP, 0)
    lFlags = lFlags Or ((QoS And &H3) * 2)
    lFlags = lFlags Or IIf(Retain, MQTT_FLAG_RETAIN, 0)
    BuildPublishFlags = lFlags
End Function

Private Function BuildFixedHeader(ByVal eType As MqttPacketType, _n                                   ByVal lFlags As Long, _n                                   ByVal lRemainingLen As Long) As Byte()
    Dim baResult() As Byte
    Dim baVarLen() As Byte
    
    ' 编码剩余长度
    baVarLen = EncodeVariableByte(lRemainingLen)
    
    ' 构建固定报头
    ReDim baResult(0 To UBound(baVarLen) + 1)
    baResult(0) = (eType * 16) Or (lFlags And &HF)
    CopyMemory baResult(1), baVarLen(0), UBound(baVarLen) + 1
    
    BuildFixedHeader = baResult
End Function

Private Function EncodeVariableByte(ByVal lValue As Long) As Byte()
    Dim baResult() As Byte
    Dim lPos As Long
    
    ReDim baResult(0 To 3)
    
    Do
        Dim lEncodedByte As Long
        lEncodedByte = lValue Mod 128
        lValue = lValue \\ 128
        If lValue > 0 Then
            lEncodedByte = lEncodedByte Or 128
        End If
        baResult(lPos) = lEncodedByte
        lPos = lPos + 1
    Loop While lValue > 0 And lPos < 4
    
    ReDim Preserve baResult(0 To lPos - 1)
    EncodeVariableByte = baResult
End Function

Private Function EncodeUTF8String(ByVal sText As String) As Byte()
    Dim baResult() As Byte
    Dim baText() As Byte
    Dim lLen As Long
    
    baText = StrConv(sText, vbFromUnicode)
    lLen = UBound(baText) + 1
    
    ReDim baResult(0 To lLen + 1)
    baResult(0) = (lLen \\ 256) And &HFF
    baResult(1) = lLen And &HFF
    If lLen > 0 Then
        CopyMemory baResult(2), baText(0), lLen
    End If
    
    EncodeUTF8String = baResult
End Function

' 解码方法实现...
Private Function DecodeConnectPayload(baData() As Byte, ByVal lOffset As Long) As Boolean
    Dim lPos As Long
    lPos = lOffset
    
    On Error GoTo EH
    
    ' Protocol Name
    ProtocolName = DecodeUTF8String(baData, lPos)
    If ProtocolName <> "MQTT" Then
        ProtocolName = DecodeUTF8String(baData, lPos)  ' Try read again
    End If
    
    ' Protocol Level
    ProtocolLevel = baData(lPos): lPos = lPos + 1
    
    ' Connect Flags
    ConnectFlags = baData(lPos): lPos = lPos + 1
    CleanSession = (ConnectFlags And &H2) <> 0
    WillRetain = (ConnectFlags And &H20) <> 0
    WillQoS = (ConnectFlags \\ 3) And &H3
    
    ' Keep Alive
    KeepAlive = baData(lPos) * 256 + baData(lPos + 1): lPos = lPos + 2
    
    ' Client ID
    ClientId = DecodeUTF8String(baData, lPos)
    
    ' Will Topic & Message
    If (ConnectFlags And &H4) <> 0 Then
        WillTopic = DecodeUTF8String(baData, lPos)
        WillMessage = DecodeUTF8String(baData, lPos)
    End If
    
    ' User Name
    If (ConnectFlags And &H80) <> 0 Then
        UserName = DecodeUTF8String(baData, lPos)
    End If
    
    ' Password
    If (ConnectFlags And &H40) <> 0 Then
        Password = DecodeUTF8String(baData, lPos)
    End If
    
    DecodeConnectPayload = True
    Exit Function
EH:
    DecodeConnectPayload = False
End Function

Private Function DecodePublishPayload(baData() As Byte, ByVal lOffset As Long) As Boolean
    Dim lPos As Long
    lPos = lOffset
    
    On Error GoTo EH
    
    ' Parse flags
    Duplicate = (Flags And MQTT_FLAG_DUP) <> 0
    QoS = (Flags \\ 2) And &H3
    Retain = (Flags And MQTT_FLAG_RETAIN) <> 0
    
    ' Topic Name
    TopicName = DecodeUTF8String(baData, lPos)
    
    ' Packet ID (if QoS > 0)
    If QoS > 0 Then
        PacketId = baData(lPos) * 256 + baData(lPos + 1): lPos = lPos + 2
    End If
    
    ' Payload (remaining bytes)
    Dim lPayloadLen As Long
    lPayloadLen = RemainingLen - (lPos - lOffset)
    If lPayloadLen > 0 Then
        ReDim PayloadBytes(0 To lPayloadLen - 1)
        CopyMemory PayloadBytes(0), baData(lPos), lPayloadLen
    End If
    
    DecodePublishPayload = True
    Exit Function
EH:
    DecodePublishPayload = False
End Function

Private Function DecodeSubscribePayload(baData() As Byte, ByVal lOffset As Long) As Boolean
    Dim lPos As Long
    Dim oFilter As cMQTTTopicFilter
    
    lPos = lOffset
    Set TopicFilters = New Collection
    
    On Error GoTo EH
    
    ' Packet ID
    PacketId = baData(lPos) * 256 + baData(lPos + 1): lPos = lPos + 2
    
    ' Topic Filters
    Do While lPos < UBound(baData) + 1
        Set oFilter = New cMQTTTopicFilter
        oFilter.Topic = DecodeUTF8String(baData, lPos)
        oFilter.QoS = baData(lPos): lPos = lPos + 1
        TopicFilters.Add oFilter
    Loop
    
    DecodeSubscribePayload = True
    Exit Function
EH:
    DecodeSubscribePayload = False
End Function

Private Function DecodeUnsubscribePayload(baData() As Byte, ByVal lOffset As Long) As Boolean
    Dim lPos As Long
    Dim oFilter As cMQTTTopicFilter
    
    lPos = lOffset
    Set TopicFilters = New Collection
    
    On Error GoTo EH
    
    ' Packet ID
    PacketId = baData(lPos) * 256 + baData(lPos + 1): lPos = lPos + 2
    
    ' Topic Filters
    Do While lPos < UBound(baData) + 1
        Set oFilter = New cMQTTTopicFilter
        oFilter.Topic = DecodeUTF8String(baData, lPos)
        TopicFilters.Add oFilter
    Loop
    
    DecodeUnsubscribePayload = True
    Exit Function
EH:
    DecodeUnsubscribePayload = False
End Function

Private Function DecodePacketIdPayload(baData() As Byte, ByVal lOffset As Long) As Boolean
    On Error GoTo EH
    PacketId = baData(lOffset) * 256 + baData(lOffset + 1)
    DecodePacketIdPayload = True
    Exit Function
EH:
    DecodePacketIdPayload = False
End Function

Private Function DecodeUTF8String(baData() As Byte, ByRef lPos As Long) As String
    Dim lLen As Long
    Dim baText() As Byte
    
    lLen = baData(lPos) * 256 + baData(lPos + 1)
    lPos = lPos + 2
    
    If lLen > 0 Then
        ReDim baText(0 To lLen - 1)
        CopyMemory baText(0), baData(lPos), lLen
        lPos = lPos + lLen
        DecodeUTF8String = StrConv(baText, vbUnicode)
    Else
        DecodeUTF8String = vbNullString
    End If
End Function

Private Sub CopyMemory(Destination As Any, Source As Any, ByVal Length As Long)
    Dim lRet As Long
    lRet = RtlMoveMemory(Destination, Source, Length)
End Sub

Private Declare Function RtlMoveMemory Lib "kernel32" (ByVal Dest As Long, ByVal Src As Long, ByVal Length As Long) As Long
