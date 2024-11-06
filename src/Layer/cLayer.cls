VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Msg(ByVal Content As String, Optional ByVal Seconds As Long = 3) As cLayer
    Set Msg = Me
    Call FLayer.ShowTo(Content)
    If Seconds > 0 Then FLayer.CloseAt (Seconds)
End Function

Public Function CloseAt(Optional ByVal Seconds As Long = 0) As cLayer
    Set CloseAt = Me
    Call FLayer.CloseAt(Seconds)
End Function

Public Function Alert(ByVal Content As String) As cLayer
    Set Alert = Me
    MsgBox Content
End Function

Public Function Confirm(ByVal Content As String) As cLayer
    Set Confirm = Me
    MsgBox Content
End Function
