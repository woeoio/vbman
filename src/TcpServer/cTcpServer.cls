VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTcpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim IsRunning  As Boolean

Private WithEvents m_oServer    As cTlsReMaster
Attribute m_oServer.VB_VarHelpID = -1
Private m_cConnPool             As New Collection

Dim WithEvents Svr As cTlsSocket
Attribute Svr.VB_VarHelpID = -1

Dim InstHandler As iTcpServer

Public LastError As String

Private Sub Class_Initialize()
    ''
End Sub

Public Function StartMe(Handler As Object, Optional ByVal Port As Long = 1024, Optional IP As String = "0.0.0.0") As Boolean
    If IsRunning = True Then LastError = "服务器正在运行中": Exit Function
    '    Set m_oServer = New cTlsReMaster
    '    Set InstHandler = Handle
    '    With m_oServer
    '        .Protocol = RemasterProtocolConstants.sckTCPProtocol
    '        .Bind Port, IP
    '        .Listen
    '    End With
End Function

Public Function StopMe() As Boolean
    '
End Function


