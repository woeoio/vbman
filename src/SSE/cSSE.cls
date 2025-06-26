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

Public MapClients As New cJson
Public MapUsersToHSocket As New cJson
Public MapHSocketToUsers As New cJson
Public MapGroups As New cJson

Public LastError As String

'To 对象
Dim hSocketByUser As Long
Dim hSocketsByGroup As New Scripting.Dictionary

Dim EH As Object
'sse通用对象
Public Function Start(Optional MatchPath As String, Optional EventsHandler As Object) As Boolean
    Path = MatchPath
    If Not EventsHandler Is Nothing Then
        Set EH = EventsHandler
    End If
    IsStart = True
    Start = True
End Function

Public Function BindUser(ByVal hSocket As Long, ByVal User As String) As Boolean
    MapUsersToHSocket.RootItem.Add User, hSocket
    MapHSocketToUsers.RootItem.Add hSocket, User
    '
    Dim Client As cClientCallback
    Set Client = Clients(hSocket)
    Client.ClientInfo.User = User
End Function

Public Function CloseClient(Optional ByVal hSocket As Long) As Boolean
    If hSocket = 0 Then
        hSocket = hSocketByUser
        GoSub Start
    Else
        If Clients.Exists(hSocket) = False Then GoTo QH
    End If
    Dim Client As cClientCallback, x As Variant
    If hSocket = 0 Then
        For Each x In Clients.Keys()
            Set Client = Clients(x)
            Call Client.Socket.SendData(d)
        Next
    Else
        Set Client = Clients(hSocket)
        Call Client.Socket.SendData(d)
    End If
    GoTo QH
Start:
    'SSE客户端关闭的时候会触发这个函数
    '    If Clients.Exists(hSocket) = False Then Exit Function
    '    Dim Client As cClientCallback
    Set Client = Clients(hSocket)
    '处理用户事件
    If Not EH Is Nothing Then
        On Error Resume Next
        CallByName EH, "OnClose", VbMethod, Client.Context
        On Error GoTo 0
    End If
    '清理 sse 上下文的实例
    Set Client.Context.SSE = Nothing
    '
    Set Clients(hSocket) = Nothing
    Clients.Remove hSocket
    Client.Socket.CloseSck
    'clear
    Dim User As String
    If MapHSocketToUsers.RootItem.Exists(hSocket) = True Then
        User = MapHSocketToUsers.RootItem(hSocket)
        MapUsersToHSocket.RootItem.Remove User
        MapHSocketToUsers.RootItem.Remove hSocket
    End If
    MapClients.RootItem.Remove hSocket
    Set Client = Nothing
    CloseClient = True
    Return
QH:
End Function

Public Property Get ClientCount() As Long
    ClientCount = Clients.Count
End Property

Public Function GetUser(ByVal hSocket As Long) As String
    If MapHSocketToUsers.RootItem.Exists(hSocket) = True Then
        GetUser = MapHSocketToUsers.RootItem(hSocket)
    End If
End Function
Public Function GetClientByUser(ByVal User As String) As Long
    If MapUsersToHSocket.RootItem.Exists(User) = True Then
        GetClientByUser = MapUsersToHSocket.RootItem(User)
    End If
End Function

'处理入口
Friend Function Entry(Client As cClientCallback) As Boolean
    On Error GoTo QH
    'SSE客户端连接都时候会触发这个函数
    If IsStart = False Then GoTo QH
    If Path <> "" Then
        If Client.Context.Request.PathInfo <> Path Then GoTo QH
    End If
    '设置 sse 上下文的实例
    Set Client.Context.SSE.flsSSEInst = Me
    '
    Clients.Add Client.hSocket, Client
    '必须立即返回一个数据，否则链接无效
    AcceptConnect = True
    Data.Add "hello", "init"
    Send Client.hSocket
    '先处理事件，如果用户拒绝连接就断开
    If Not EH Is Nothing Then
        CallByName EH, "OnConnect", VbMethod, Client.Context
    End If
    '映射
    MapClients.RootItem.Add Client.hSocket, Now()
    '
    '    RaiseEvent HandleRecvMsg
    '
    Entry = True
    Exit Function
QH:
    LastError = ERR.Number & "#" & ERR.Description
    Debug.Print "SSE Entry Error:", LastError
    Client.Socket.CloseSck
End Function

'Public Function ToClient(hSocket As Long) As cSSE
'    '
'End Function
Public Function ToUser(User As String) As cSSE
    Set ToUser = Me
    If MapUsersToHSocket.RootItem.Exists(User) = True Then
        hSocketByUser = MapUsersToHSocket.RootItem(User)
    End If
End Function
Public Function ToGroup(Name As String) As cSSE
    Set ToGroup = Me
End Function

Public Function SendPack(Action As String, Content As Variant, Optional hSocket As Long) As Boolean
    Data.RemoveAll
    Data.Add "action", Action
    Data.Add "content", Content
    SendPack = Send(hSocket)
End Function
Public Function Send(Optional ByVal hSocket As Long) As Boolean
    If Data.Count = 0 Then GoTo QH
    If Clients.Count = 0 Then GoTo QH
    If hSocket = 0 Then
        hSocket = hSocketByUser
    Else
        If Clients.Exists(hSocket) = False Then GoTo QH
    End If
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
    hSocketByUser = 0
    hSocketsByGroup.RemoveAll
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
