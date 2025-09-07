VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVBMAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Json As New cJson
Public Layer As New cLayer
Public Db As New cDataBase
Public HttpClient As New cHttpClient
Public Csv As New cCsv
Public Ini As New cIni
Public Logs As New cLogs
Public Cmd As New cCmd
Public TimeUse As New cTimeUse

'µÚÈý·½
Public Baidu As New cBaidu

'tools
Public ToolsStream As New cToolsStream
Public ToolsWindow As New cToolsWindow
Public ToolsStr As New cToolsStr
Public ToolsMath As New cToolsMath
Public ToolsHttp As New cToolsHttp
Public ToolsBase64 As New cToolsBase64
Public ToolsUtf8 As New cToolsUtf8

Public Function Version(Optional HostApp As Object) As String
    Version = Common.Version(HostApp)
End Function

