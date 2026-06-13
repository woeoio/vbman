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

' As New→显式：所有字段改为显式创建，确保 Set = Nothing 真正释放
Dim Json As cJson
Dim IsStart As Boolean
Dim Path As String
Dim AcceptConnect As Boolean
Dim Clients As Scripting.Dictionary
Public Data As Scripting.Dictionary

Public MapClients As cJson
Public MapUsersToHSocket As cJson
Public MapHSocketToUsers As cJson
Public MapGroups As cJson

Public LastError As String

'To 对象
Dim hSocketByUser As Long
Dim hSocketsByGroup As Scripting.Dictionary

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

' BUG-15 修复：重构 CloseClient，统一清理路径并关闭 Socket
' 原实现问题：通过 hSocket 参数调用时，只检查是否存在但不做任何清理
' 修复：无论哪种调用方式，都执行完整的 SSE 清理 + 关闭 Socket
Public Function CloseClient(Optional ByVal hSocket As Long) As Boolean
    On Error GoTo QH
    
    If hSocket = 0 Then hSocket = hSocketByUser
    If hSocket = 0 Then GoTo QH
    If Clients.Exists(hSocket) = False Then GoTo QH
    
    Dim Client As cClientCallback
    Set Client = Clients(hSocket)
    
    ' 处理用户事件
    If Not EH Is Nothing Then
        On Error Resume Next
        CallByName EH, "OnClose", VbMethod, Client.Context
        On Error GoTo QH
    End If
    
    ' 清理 SSE 上下文的实例
    Set Client.Context.SSE = Nothing
    
    ' 从 Clients 字典移除（防止 Socket_Close 重入时重复清理）
    Set Clients(hSocket) = Nothing
    Clients.Remove hSocket
    
    ' 清理映射
    Dim User As String
    If MapHSocketToUsers.RootItem.Exists(hSocket) = True Then
        User = MapHSocketToUsers.RootItem(hSocket)
        MapUsersToHSocket.RootItem.Remove User
        MapHSocketToUsers.RootItem.Remove hSocket
    End If
    On Error Resume Next
    MapClients.RootItem.Remove hSocket
    On Error GoTo QH
    
    ' 关闭 Socket，触发 cClientCallback 的完整清理流程
    ' CloseSck 会触发 Socket_Close 事件，但此时 Client 已从 Clients 移除、
    ' Context.SSE 已清空，cClientCallback 的 Close 处理不会重复清理 SSE
    On Error Resume Next
    Client.Socket.CloseSck
    On Error GoTo QH
    
    Set Client = Nothing
    CloseClient = True
    
QH:
    hSocketByUser = 0
End Function

Public Property Get ClientCount() As Long
    If Not Clients Is Nothing Then
        ClientCount = Clients.Count
    End If
End Property

Public Function GetUser(ByVal hSocket As Long) As String
    If Not MapHSocketToUsers Is Nothing Then
        If MapHSocketToUsers.RootItem.Exists(hSocket) = True Then
            GetUser = MapHSocketToUsers.RootItem(hSocket)
        End If
    End If
End Function
Public Function GetClientByUser(ByVal User As String) As Long
    If Not MapUsersToHSocket Is Nothing Then
        If MapUsersToHSocket.RootItem.Exists(User) = True Then
            GetClientByUser = MapUsersToHSocket.RootItem(User)
        End If
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
    ' BUG-9 修复：Entry 失败时清理已设置的引用，防止循环引用和内存泄漏
    On Error Resume Next
    If Not Client Is Nothing Then
        If Not Client.Context Is Nothing Then
            Set Client.Context.SSE = Nothing
        End If
        ' 如果 Client 已添加到 Clients，需要移除
        If Not Clients Is Nothing Then
            If Clients.Exists(Client.hSocket) Then
                Clients.Remove Client.hSocket
            End If
        End If
        Client.Socket.CloseSck
    End If
    On Error GoTo 0
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

' BUG-10 修复：Send 方法在发送前先捕获 Data 快照到局部变量
' 原实现问题：共享 Data 成员在 SendData 过程中可能被重入调用修改
' 修复：先将 Data 复制到局部 SendData，再清空 Data，用 SendData 编码发送
Public Function Send(Optional ByVal hSocket As Long) As Boolean
    If Data.Count = 0 Then GoTo QH
    If Clients.Count = 0 Then GoTo QH
    
    ' 先捕获 Data 快照到局部变量，防止 SendData 触发重入时 Data 被污染
    Dim SendData As Scripting.Dictionary
    Set SendData = New Scripting.Dictionary
    Dim k As Variant
    For Each k In Data.Keys
        SendData.Add k, Data(k)
    Next
    Data.RemoveAll  ' 立即清空共享 Data，后续重入调用不会读到脏数据
    
    If hSocket = 0 Then
        hSocket = hSocketByUser
    Else
        If Clients.Exists(hSocket) = False Then GoTo QH
    End If
    
    Dim Header As String, Client As cClientCallback, x As Variant
    If AcceptConnect = True Then Header = MakeHeaders()
    Dim SseDataText As String: SseDataText = Header & "event: message" & vbCrLf & "data: " & Json.Encode(SendData) & "" & vbCrLf & vbCrLf
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
    If Not hSocketsByGroup Is Nothing Then hSocketsByGroup.RemoveAll
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

Private Sub Class_Initialize()
    ' 显式创建所有子对象（替代 As New 隐式创建）
    Set Json = New cJson
    Set Clients = New Scripting.Dictionary
    Set Data = New Scripting.Dictionary
    Set MapClients = New cJson
    Set MapUsersToHSocket = New cJson
    Set MapHSocketToUsers = New cJson
    Set MapGroups = New cJson
    Set hSocketsByGroup = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    ' 释放所有 SSE 客户端引用
    If Not Clients Is Nothing Then
        Dim k As Variant
        For Each k In Clients.Keys
            Set Clients(k) = Nothing
        Next
        Clients.RemoveAll
        Set Clients = Nothing
    End If
    Set Json = Nothing
    If Not Data Is Nothing Then
        Data.RemoveAll
        Set Data = Nothing
    End If
    Set MapClients = Nothing
    Set MapUsersToHSocket = Nothing
    Set MapHSocketToUsers = Nothing
    Set MapGroups = Nothing
    If Not hSocketsByGroup Is Nothing Then
        hSocketsByGroup.RemoveAll
        Set hSocketsByGroup = Nothing
    End If
    Set EH = Nothing
End Sub
