VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function LoadFileAsText(ByVal Path As String, Optional CharSet As String = "UTF-8") As String
    LoadFileAsText = ToolsStream.LoadFileAsText(Path, CharSet)
End Function

Public Function SaveFileAsText(ByVal Path As String, Data As String, Optional CharSet As String = "UTF-8") As Boolean
    SaveFileAsText = ToolsStream.SaveFileAsText(Path, Data, CharSet)
End Function


Public Function LoadFileAsBinary(ByVal Path As String, OutData() As Byte) As Boolean
    LoadFileAsBinary = ToolsStream.LoadFileAsBinary(Path, OutData)
End Function


Public Function SaveFileAsBinary(ByVal Path As String, OutData() As Byte) As Boolean
    SaveFileAsBinary = ToolsStream.SaveFileAsBinary(Path, OutData)
End Function
