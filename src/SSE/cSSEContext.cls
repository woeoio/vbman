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

Public Property Get ClientCount() As Long
    ClientCount = SSEInst.ClientCount
End Property

Public Function BindUser(ByVal hSocket As Long, ByVal User As String) As Boolean
    BindUser = SSEInst.BindUser(hSocket, User)
End Function

Public Function SendPack(Action As String, Content As Variant, Optional hSocket As Long) As Boolean
    SendPack = SSEInst.SendPack(Action, Content, hSocket)
End Function

Public Function CloseClient(Optional ByVal hSocket As Long) As Boolean
    CloseClient = SSEInst.CloseClient(hSocket)
End Function

Public Function ToUser(ByVal User As String) As cSSEContext
    Set ToUser = Me
    Call SSEInst.ToUser(User)
End Function
Public Function ToGroup(ByVal Name As String) As cSSEContext
    Set ToGroup = Me
    Call SSEInst.ToGroup(Name)
End Function

Public Function GetClientByUser(ByVal User As String) As Long
    GetClientByUser = SSEInst.GetClientByUser(User)
End Function
Public Function GetUser(ByVal hSocket As Long) As String
    GetUser = SSEInst.GetUser(hSocket)
End Function

Private Sub Class_Terminate()
    Set SSEInst = Nothing
End Sub
