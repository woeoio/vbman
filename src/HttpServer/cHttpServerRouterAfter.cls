VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerRouterAfter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'后置路由采用的是异步回调

Dim Dic As New Scripting.Dictionary

Public Property Get Handles() As Scripting.Dictionary
    Set Handles = Dic
End Property

Friend Function Entry(ctx As cHttpServerContext) As Boolean
    
    Entry = True
End Function

Public Function Add(HandleName As String, Handle As Object) As Boolean
    Set Dic(HandleName) = Handle
End Function

