VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim Json As New cJson
Dim IsStart As Boolean
Dim Path As String
Dim AcceptConnect As Boolean
Dim Clients As New Scripting.Dictionary
Dim UserHandler As cSSEHandle
Public Data As New Scripting.Dictionary
'sse通用对象
Public Function Start(Optional MatchPath As String, Optional OnConnect As Object) As Boolean
    Path = MatchPath
    IsStart = True
    Start = True
End Function

Public Function Remove(hSocket As Long) As Boolean
    If Clients.Exists(hSocket) = False Then Exit Function
    Dim Client As cClientCallback
    Set Client = Clients(hSocket)
    Set Clients(hSocket) = Nothing
    Clients.Remove hSocket
    Client.Socket.CloseSck
    Set Client = Nothing
    Remove = True
End Function

Public Property Get ClientCount() As Long
    ClientCount = Clients.Count
End Property

'处理入口
Friend Function Entry(Client As cClientCallback) As Boolean
    If IsStart = False Then GoTo QH
    If Path <> "" Then
        If Client.Context.Request.PathInfo <> Path Then GoTo QH
    End If
    Clients.Add Client.hSocket, Client
    '必须立即返回一个数据，否则链接无效
    AcceptConnect = True
    Data.Add "hello", "init"
    Send Client.hSocket
    '处理用户函数
    
    '
    '    RaiseEvent HandleRecvMsg
    '
    Entry = True
    Exit Function
QH:
    Client.Socket.CloseSck
End Function

Public Function ToClient(hSocket As Long) As cSSE
    '
End Function
Public Function ToUser(User As String) As cSSE
    '
End Function
Public Function ToGroup(Name As String) As cSSE
    '
End Function

Public Function SendPack(Action As String, Content As Variant, Optional hSocket As Long) As Boolean
    Data.RemoveAll
    Data.Add "action", Action
    Data.Add "content", Content
    SendPack = Send(hSocket)
End Function
Public Function Send(Optional hSocket As Long) As Boolean
    If Data.Count = 0 Then GoTo QH
    If hSocket <> 0 Then
        If Clients.Exists(hSocket) = False Then GoTo QH
    End If
    If Clients.Count = 0 Then GoTo QH
    Dim Header As String, Client As cClientCallback, x As Variant
    If AcceptConnect = True Then Header = MakeHeaders()
    Dim SseDataText As String: SseDataText = Header & "event: message" & vbCrLf & "data: " & Json.Encode(Data) & "" & vbCrLf & vbCrLf
    Dim d() As Byte: d = ToolsUtf8.Encode(SseDataText)
    If hSocket = 0 Then
        For Each x In Clients.Keys()
            Set Client = Clients(x)
            Call Client.Socket.SendData(d)
        Next
    Else
        Set Client = Clients(hSocket)
        Call Client.Socket.SendData(d)
    End If
    Send = True
QH:
    Data.RemoveAll
End Function
Private Function MakeHeaders() As String
    Dim ResponseLine$, Tstr$
    ResponseLine = "HTTP/1.1 200 OK"
    Tstr = ResponseLine
    Tstr = Tstr & vbCrLf & "Server: VBMAN-SSE Server 1.0.0 By vb6.pro/vbman"
    Tstr = Tstr & vbCrLf & "Access-Control-Allow-Origin: *"
    Tstr = Tstr & vbCrLf & "Access-Control-Allow-Methods:*"
    Tstr = Tstr & vbCrLf & "Access-Control-Allow-Headers:*"
    Tstr = Tstr & vbCrLf & "Date: " & Format(Now, "ddd, yyyy-mm-dd hh:mm:ss GMT")
    Tstr = Tstr & vbCrLf & "Content-Type: text/event-stream"
    MakeHeaders = Tstr & vbCrLf & vbCrLf
End Function
