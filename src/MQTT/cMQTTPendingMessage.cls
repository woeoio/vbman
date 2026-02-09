'=========================================================================
' cMQTTPendingMessage - 待确认消息
' 用于 QoS 1/2 的消息跟踪
'=========================================================================
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMQTTPendingMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public PacketId As Long
Public Message As cMQTTMessage
Public Timestamp As Date
Public RetryCount As Long
