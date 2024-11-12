Attribute VB_Name = "ToolsStream"
Option Explicit

'    Library ADODB
'C:  \Program Files (x86)\Common Files\System\ado\msado28.tlb
'    Microsoft ActiveX Data Objects 2.8 Library

Public Inst As New ADODB.Stream
Public LastError As String

Public Function LoadFileAsText(ByVal FileName As String, Optional CharSet As String = "UTF-8") As String
    With Inst
        .Open
        .CharSet = CharSet
        .LoadFromFile FileName
        LoadFileAsText = .ReadText()
        .Close
    End With
End Function
Public Function SaveFileAsText(ByVal FileName As String, Data As String, Optional CharSet As String = "UTF-8") As Boolean
    With Inst
        .Open
        .CharSet = CharSet
        .WriteText Data
        .SaveToFile FileName, adSaveCreateNotExist + adSaveCreateOverWrite
        .Close
    End With
End Function

Public Function LoadFileAsBinary(ByVal Path As String, OutData() As Byte) As Boolean
    On Error GoTo Eh
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
Eh:
    LastError = ERR.Description & "#" & ERR.Description
End Function

Public Function SaveFileAsBinary(ByVal Path As String, OutData() As Byte) As Boolean
    On Error GoTo Eh
    LastError = ""
    With Inst
        If .State <> adStateClosed Then .Close
        .Type = adTypeBinary
        .Open
        .Write OutData
        .SaveToFile Path, adSaveCreateNotExist + adSaveCreateOverWrite
        .Close
    End With
    SaveFileAsBinary = True
    Exit Function
Eh:
    LastError = ERR.Description & "#" & ERR.Description
End Function

