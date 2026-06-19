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
Public Db As New cDataBase
Public HttpClient As New cHttpClient
Public Csv As New cCsv
Public Ini As New cIni
Public Logs As New cLogs
Public Cmd As New cCmd
Public TimeUse As New cTimeUse
Public PLI As New cPLI
Public QRcode As New cQRcode
Public Collection As New cCollection
Public Regedit As New cRegedit
Public StartUp As New cStartUp
Public Toast As New cToast
Public Password As New cPassword
Public Delay As New cDelay
Public FileEx As New cFileEx
Public Dialog As New cDialog
Public Formater As New cFormater
Public StdIO As New cSTDIO
Public Zip As New cZipArchive

Public Hash As New cCryptoHash
Public HashMAC As New cCryptoHMAC

'第三方
Public Baidu As New cBaidu

'AI
''Public AI As New cAI

'tools
Public ToolsWindow As New cToolsWindow
Public ToolsStr As New cToolsStr
Public ToolsMath As New cToolsMath
Public ToolsHttp As New cToolsHttp
Public ToolsBase64 As New cToolsBase64
Public ToolsUtf8 As New cToolsUtf8
Public ToolsStream As New cToolsStream
Public ToolsArray As New cToolsArray
Public ToolsDic As New cToolsDic
Public ToolsFso As New cToolsFso
Public ToolsList As New cToolsList
Public ToolsSystem As New cToolsSystem
Public ToolsCrc As New cToolsCrc

'实验证明，dll编译后的app对象就是dll所属的
Public Function Version(Optional HostApp As Object) As String
    Version = Common.Version(HostApp)
End Function

Public Function Path(Optional HostApp As Object) As String
    Path = Common.Path(HostApp)
End Function


