'=========================================================================
' mMQTTConst - MQTT Protocol Constants
' MQTT Version 3.1.1 (Standard)
'=========================================================================
Option Explicit

'-----------------------------------------------------------------------------
' MQTT Packet Types (Fixed Header Byte 1, bits 7-4)
'-----------------------------------------------------------------------------
Public Enum MqttPacketType
    mqttReserved    = 0     ' Reserved
    mqttConnect     = 1     ' Client request to connect to Server
    mqttConnAck     = 2     ' Connect acknowledgment
    mqttPublish     = 3     ' Publish message
    mqttPubAck      = 4     ' Publish acknowledgment (QoS 1)
    mqttPubRec      = 5     ' Publish received (QoS 2 delivery part 1)
    mqttPubRel      = 6     ' Publish release (QoS 2 delivery part 2)
    mqttPubComp     = 7     ' Publish complete (QoS 2 delivery part 3)
    mqttSubscribe   = 8     ' Client subscribe request
    mqttSubAck      = 9     ' Subscribe acknowledgment
    mqttUnsubscribe = 10    ' Unsubscribe request
    mqttUnsubAck    = 11    ' Unsubscribe acknowledgment
    mqttPingReq     = 12    ' PING request (keep alive)
    mqttPingResp    = 13    ' PING response
    mqttDisconnect  = 14    ' Client is disconnecting
    mqttReserved2   = 15    ' Reserved
End Enum

'-----------------------------------------------------------------------------
' MQTT Packet Flags (Fixed Header Byte 1, bits 3-0)
'-----------------------------------------------------------------------------
' PUBLISH flags: DUP(3) | QoS(2-1) | RETAIN(0)
Public Const MQTT_FLAG_DUP     As Long = &H8   ' Duplicate delivery
Public Const MQTT_FLAG_QOS0    As Long = &H0   ' At most once
Public Const MQTT_FLAG_QOS1    As Long = &H2   ' At least once
Public Const MQTT_FLAG_QOS2    As Long = &H4   ' Exactly once
Public Const MQTT_FLAG_RETAIN  As Long = &H1   ' Retain message

'-----------------------------------------------------------------------------
' MQTT QoS Levels
'-----------------------------------------------------------------------------
Public Enum MqttQoS
    mqttQoS0 = 0    ' At most once (Fire and Forget)
    mqttQoS1 = 1    ' At least once (Acknowledged delivery)
    mqttQoS2 = 2    ' Exactly once (Assured delivery)
End Enum

'-----------------------------------------------------------------------------
' Connect Return Codes (CONNACK)
'-----------------------------------------------------------------------------
Public Enum MqttConnectReturnCode
    mqttAccepted            = 0     ' Connection accepted
    mqttUnacceptableVersion = 1     ' Unacceptable protocol version
    mqttIdentifierRejected  = 2     ' Identifier rejected
    mqttServerUnavailable   = 3     ' Server unavailable
    mqttBadCredentials      = 4     ' Bad user name or password
    mqttNotAuthorized       = 5     ' Not authorized
End Enum

'-----------------------------------------------------------------------------
' Subscribe Return Codes (SUBACK)
'-----------------------------------------------------------------------------
Public Enum MqttSubscribeReturnCode
    mqttSubSuccessQoS0 = &H0    ' Success - Maximum QoS 0
    mqttSubSuccessQoS1 = &H1    ' Success - Maximum QoS 1
    mqttSubSuccessQoS2 = &H2    ' Success - Maximum QoS 2
    mqttSubFailure     = &H80   ' Failure
End Enum

'-----------------------------------------------------------------------------
' Protocol Constants
'-----------------------------------------------------------------------------
Public Const MQTT_PROTOCOL_NAME       As String = "MQTT"
Public Const MQTT_PROTOCOL_LEVEL_311  As Long = 4      ' MQTT 3.1.1
Public Const MQTT_PROTOCOL_LEVEL_31   As Long = 3      ' MQTT 3.1

Public Const MQTT_DEFAULT_PORT        As Long = 1883
Public Const MQTT_DEFAULT_SSL_PORT    As Long = 8883

' Maximum values
Public Const MQTT_MAX_CLIENT_ID_LEN   As Long = 23
Public Const MQTT_MAX_PACKET_SIZE     As Long = 268435455   ' 256 MB (Variable byte max)
Public Const MQTT_MAX_TOPIC_LEN       As Long = 65535
Public Const MQTT_MAX_KEEP_ALIVE      As Long = 65535

'-----------------------------------------------------------------------------
' Topic Wildcards
'-----------------------------------------------------------------------------
Public Const MQTT_WILDCARD_SINGLE     As String = "+"   ' Single level
Public Const MQTT_WILDCARD_MULTI      As String = "#"   ' Multi level
Public Const MQTT_TOPIC_SEPARATOR     As String = "/"   ' Topic level separator

'-----------------------------------------------------------------------------
' Error Codes
'-----------------------------------------------------------------------------
Public Enum MqttErrorCode
    mqttErrSuccess = 0
    mqttErrInvalidPacket = vbObjectError + 1001
    mqttErrInvalidProtocol
    mqttErrInvalidClientId
    mqttErrInvalidTopic
    mqttErrInvalidQoS
    mqttErrPacketTooLarge
    mqttErrNotConnected
    mqttErrTimeout
    mqttErrSocketError
End Enum
