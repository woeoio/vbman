'=========================================================================
' cMQTTMessage - MQTT Message
' MQTT 消息对象
'=========================================================================
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMQTTMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Topic As String
Public Payload As String
Public PayloadBytes() As Byte
Public QoS As MqttQoS
Public Retain As Boolean
Public Duplicate As Boolean
Public Timestamp As Date

Private Sub Class_Initialize()
    Timestamp = Now
    QoS = mqttQoS0
End Sub
