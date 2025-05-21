VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSSEContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim SSEInst As cSSE

Friend Property Set flsSSEInst(o As cSSE)
    Set SSEInst = o
End Property

Public Function SendPack(Action As String, Content As Variant, Optional hSocket As Long) As Boolean
    SendPack = SSEInst.SendPack(Action, Content, hSocket)
End Function
