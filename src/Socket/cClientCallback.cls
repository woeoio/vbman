VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClientCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Index                As Long
Public hSocket                As Long
Public dataTotal                As Long
Public ClientInfo As cHttpServerClientInfo
Public Context As cHttpServerContext
Public Parent               As cHttpServer
Public WithEvents Socket    As cTlsReMaster
Attribute Socket.VB_VarHelpID = -1

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    dataTotal = bytesTotal
    Parent.OnDataArrival Index
End Sub

Private Sub Socket_CloseSck()
    Parent.OnCloseSck Index
End Sub

