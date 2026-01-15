VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim Reg As New cRegedit
Const K As String = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"

Public LastError As String

Public Function Toggle(ByVal Name As String, ByRef Path As Variant, ParamArray StartArgs() As Variant) As Boolean
    '    App.Path , App.EXEName
    Dim P As String
    If IsObject(Path) = True Then
        P = Path.Path & "\" & Path.EXEName & ".exe"
    Else
        P = Path
    End If
    Dim R() As TypeRegData: R = Reg.FindItem(K, Name)
    If R(0).HasName = False Then
        Dim A As Variant: A = StartArgs
        Toggle = Reg.SaveItem(K, Name, P, A)
    Else
        Toggle = Reg.DeleteItem(K, Name)
    End If
End Function

Public Function Has(ByVal Name As String) As Boolean
    Dim R() As TypeRegData: R = Reg.FindItem(K, Name)
    Has = R(0).HasName
End Function
