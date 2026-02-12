VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim Cmd As New cSTDIO

Public Type TypeRegData
    HasName As Boolean
    RegName As String
    RegType As String
    RegValue As String
End Type

Public LastError As String

Public Function FindItem(ByVal Path As String, Optional ByVal Name As String) As TypeRegData()
    Const Q As String = "reg query "
    Dim N As String
    If Name <> "" Then N = " /v " & Name
    Dim R As String: R = Cmd.ExecuteCommand(Q & """" & Path & """" & N)
    Dim H As Variant: H = Split(R, vbCrLf)
    Dim i As Long, A As Variant
    Dim Arr() As TypeRegData
    Dim L As Long: L = UBound(H)
    If L < 3 Then
        ReDim Arr(0)
        Arr(0).RegType = "ERROR"
        Arr(0).RegValue = R
        FindItem = Arr
        LastError = R
        Exit Function
    End If
    For i = 2 To L - 1
        If H(i) = "" Then GoTo CT
        A = Split(H(i), "    ")
        If Trim(A(0)) = "" Then
            ReDim Preserve Arr(i - 2)
            With Arr(i - 2)
                .HasName = True
                .RegName = A(1)
                .RegType = A(2)
                .RegValue = A(3)
            End With
        Else
            ReDim Arr(0)
            Arr(0).RegType = "ERROR"
            Arr(0).RegValue = A(0)
            LastError = R
            Exit For
        End If
CT:
    Next
    FindItem = Arr
    LastError = ""
End Function

Public Function SaveItem( _
    ByVal Path As String, _
    ByVal Name As String, _
    ByVal Data As Variant, _
    Optional StartArgs As Variant, _
    Optional IsOverWrite As Boolean = True) As Boolean
    'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Notepad" /d "C:\Windows\System32\notepad.exe" /f
    '    Const IsOverWrite As Boolean = True
    Const Q As String = "reg add "
    Dim N As String: N = " /v """ & Name
    Dim D As String
    If IsMissing(StartArgs) = True Then
        D = " /d """ & Data & """"
    Else
        If IsArray(StartArgs) = True Then
            StartArgs = Join(StartArgs, " ")
        End If
        D = " /d ""\""" & Data & "\"" " & StartArgs & """"
    End If
    Dim F As String: If IsOverWrite = True Then F = " /f"
    Dim C As String: C = Q & """" & Path & """" & N & """" & D & F
    Dim R As String: R = Cmd.ExecuteCommand(C)
    Dim H As Variant: H = Split(R, vbCrLf)
    Dim L As Long: L = UBound(H)
    If L > 1 Then
        LastError = R
        Exit Function
    End If
    SaveItem = True
    LastError = ""
End Function

Public Function DeleteItem(ByVal Path As String, ByVal Name As String) As Boolean
    Const Q As String = "reg delete "
    Dim N As String: N = " /v " & Name
    Dim F As String:  F = " /f"
    Dim R As String: R = Cmd.ExecuteCommand(Q & """" & Path & """" & N & F)
    Dim H As Variant: H = Split(R, vbCrLf)
    Dim L As Long: L = UBound(H)
    If L > 1 Then
        LastError = R
        Exit Function
    End If
    DeleteItem = True
    LastError = ""
End Function

Public Function FindFirst(ByVal Path As String, Optional ByVal Name As String) As Variant
    Dim Arr() As TypeRegData
    Arr = FindItem(Path, Name)
    If Arr(0).HasName = False Then
        FindFirst = vbNullString
    Else
        FindFirst = Arr(0).RegValue
    End If
End Function

Public Function FindLast(ByVal Path As String, Optional ByVal Name As String) As Variant
    Dim Arr() As TypeRegData
    Arr = FindItem(Path, Name)
    If Arr(0).HasName = False Then
        FindLast = vbNullString
    Else
        FindLast = Arr(UBound(Arr)).RegValue
    End If
End Function

