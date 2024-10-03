VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents m_oServer    As cTlsReMaster
Attribute m_oServer.VB_VarHelpID = -1
Private m_cConnPool             As New Collection

Public Event OnAccept(ClientInfo As cHttpServerClientInfo, Disconnect As Boolean)

Public Router As New cHttpServerRouter

Public Function Start(Optional Port As Long = 80, Optional WebRoot As String = "\www", Optional IP As String = "0.0.0.0") As Boolean
    Set m_oServer = New cTlsReMaster
    m_oServer.Protocol = RemasterProtocolConstants.sckTCPProtocol
    m_oServer.Bind Port, IP
    m_oServer.Listen
    '    Shell "cmd /c start http://localhost:8088/"
    Start = True
End Function


Private Sub m_oServer_ConnectionRequest(ByVal requestID As Long, Disconnect As Boolean)
    
    ' requestID 就是客户端的句柄 hSocket
    '    Debug.Print "m_oServer_ConnectionRequest, requestID=" & requestID & ", RemoteHostIP=" & m_oServer.RemoteHostIP & ", RemotePort=" & m_oServer.RemotePort, Timer
    'todo 实现 Disconnect,
    Dim CI As New cHttpServerClientInfo, IsDiss As Boolean
    With CI
        .IP = m_oServer.RemoteHostIP
        .Port = m_oServer.RemotePort
        .ConnectAt = Now()
    End With
    RaiseEvent OnAccept(CI, IsDiss)
    If IsDiss = True Then
        Disconnect = True
        Exit Sub
    End If
    Dim oCallback           As New cClientCallback
    m_cConnPool.Add oCallback
    With oCallback
        .Index = m_cConnPool.Count
        Set .Parent = Me
        Set .ClientInfo = CI
        Set .Socket = New cTlsReMaster
        Set .Context = New cHttpServerContext
        Set .Context.ClientInfo = CI
        .Socket.Accept requestID
    End With
    '    Debug.Print "oCallback.Socket.Protocol=" & oCallback.Socket.Protocol
End Sub

Friend Sub OnDataArrival(Index As Long)
    Dim sRequest            As String
    Dim oCallback           As cClientCallback
    Set oCallback = m_cConnPool.Item(Index)
    '    Debug.Print "OnDataArrival, Index=" & Index & ", bytesTotal=" & oCallback.dataTotal, Timer
    oCallback.Socket.GetData sRequest
    
    Test oCallback, sRequest
End Sub

Friend Sub OnCloseSck(Index As Long)
    m_cConnPool.Remove Index
    If m_cConnPool.Count >= Index Then
        m_cConnPool.Add Nothing, Before:=Index
    End If
    Do While m_cConnPool.Count > 0
        If Not m_cConnPool.Item(m_cConnPool.Count) Is Nothing Then
            Exit Do
        End If
        m_cConnPool.Remove m_cConnPool.Count
    Loop
End Sub


Private Sub Test(oCallback As cClientCallback, sRequest As String)
    Dim Index As Long: Index = oCallback.Index
    Dim vSplit              As Variant
    Dim sBody               As String
    vSplit = Split(sRequest, vbCrLf)
    If UBound(vSplit) >= 0 Then
        '        Debug.Print vSplit(0)
        sBody = "<html><body><p>" & Join(vSplit, "</p>" & vbCrLf & "<p>" & Index & ": ") & "</p>" & vbCrLf & _
        "<p>" & Index & ": Current time is " & Now & "</p>" & _
        "<p>" & Index & ": RemoteHostIP is " & oCallback.Socket.RemoteHostIP & "</p>" & vbCrLf & _
        "<p>" & Index & ": RemotePort is " & oCallback.Socket.RemotePort & "</p>" & vbCrLf & _
        "</body></html>" & vbCrLf
        oCallback.Socket.SendData "HTTP/1.1 200 OK" & vbCrLf & _
"Content-Type: text/html" & vbCrLf & _
"Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
        sBody
    End If
    '    Debug.Print "OnDataArrival, Index=" & Index & ", done", Timer
End Sub

