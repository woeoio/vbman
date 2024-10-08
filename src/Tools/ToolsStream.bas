Attribute VB_Name = "ToolsStream"
Option Explicit

'    Library ADODB
'C:  \Program Files (x86)\Common Files\System\ado\msado28.tlb
'    Microsoft ActiveX Data Objects 2.8 Library

Public Inst As New ADODB.Stream
Public LastError As String

Public Function LoadFileAsBinary(ByVal Path As String, OutData() As Byte) As Boolean
    On Error GoTo EH
    LastError = ""
    With Inst
        If .State <> adStateClosed Then .Close
        .Type = adTypeBinary
        .Open
        .LoadFromFile Path
        OutData = .Read()
        .Close
    End With
    LoadFileAsBinary = True
    Exit Function
EH:
    LastError = Err.Description & "#" & Err.Description
End Function

