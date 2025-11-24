VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TimeOut As Long
'Public PliPath As String
Dim StdIO As New cSTDIO
Dim PliPath As String

Const PLI_NAME As String = "VBMAN.PLI"
Const PSC_NAME As String = "VBMAN.PSC"

Public Function Request(ParamArray Data() As Variant) As String
    PliPath = App.Path & "\"
    If Dir(PliPath & PLI_NAME) = "" Then
        PliPath = App.Path & "\bin\"
    End If
    If Dir(PliPath & PLI_NAME) = "" Then
        PliPath = App.Path & "\..\bin\"
    End If
    If Dir(PliPath & PLI_NAME) = "" Then
        ERR.Raise 404, "PLI", "ÕÒ²»µ½" & PLI_NAME
    End If
    '
    '    Dim Data As String: Data = Join(Array( _
    '    ToolsBase64.Utf8Encode(ConfigFile), _
    '    ToolsBase64.Utf8Encode(Method), _
    '    ToolsBase64.Utf8Encode(FormID), _
    '    ToolsBase64.Utf8Encode(FormData) _
    '    ), "|")
    Dim Dat As String, i As Long, l As Long, NewArr() As String
    l = UBound(Data)
    If l > -1 Then
        ReDim NewArr(l)
        For i = 0 To l
            NewArr(i) = ToolsBase64.Utf8Encode(Data(i))
        Next
        Dat = Join(NewArr, "|")
    End If
    '
    Dim Cmd As String: Cmd = Join(Array( _
    PliPath & PLI_NAME, _
    ".\" & PSC_NAME, _
    Dat _
    ), " ")
    '    MsgBox Cmd
    Request = StdIO.ExecuteCommand(Cmd, TimeOut, PliPath)
End Function


Private Sub Class_Initialize()
    TimeOut = 30& * 1000&
End Sub
