VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerRouteBefore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Dim Dic As New Scripting.Dictionary

Public Property Get Handles() As Scripting.Dictionary
    Set Handles = Dic
End Property

Friend Function Entry(ctx As cHttpServerContext) As Boolean
    On Error GoTo EH
    Dim k
    For Each k In Dic.Keys()
        Call Dic(k).Entry(ctx)
        If ctx.fIsAbort = True Then Exit For
        If ctx.fIsSkipNextMiddleWare = True Then Exit For
    Next
    Entry = True
    Exit Function
EH:
End Function

Public Function Add(HandleName As String, Handle As Object) As Boolean
    Set Dic(HandleName) = Handle
End Function

