VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim IsAbort As Boolean
Dim IsSkipNextMiddleWare As Boolean



Public App As cHttpServerApp
Public Server As cHttpServerSvr                                                 'LastError
Public ClientInfo As cHttpServerClientInfo
Public Request As cHttpServerRequest
Public Response As cHttpServerResponse
Public Session As cHttpServerSession
Public Cookies As cHttpServerCookies
Public TimeUse As New cTimeUse
Public Db As cDataBase

Public SSE As New cSSEContext
'用户数据，用于整个请求生命周期传递数据
Public UserData As New Scripting.Dictionary

Friend Property Get fIsAbort() As Boolean
    fIsAbort = IsAbort
End Property
Friend Property Get fIsSkipNextMiddleWare() As Boolean
    fIsSkipNextMiddleWare = IsSkipNextMiddleWare
End Property


Public Sub Abort()
    IsAbort = True
End Sub

Public Sub SkipNextMiddleWare()
    IsSkipNextMiddleWare = True
End Sub

' 请求处理完毕后显式释放本次请求相关对象，切断循环引用链
' 请求级短生命周期对象：Request/Response/Cookies/Session
' 连接级长生命周期对象：SSE/UserData/TimeUse/Server/App/Db/ClientInfo 保留不释放
Friend Sub ReleaseRequest()
    On Error Resume Next
    
    ' 1. 先断开 Response 对 ClientCallback 的反向引用（最关键的循环引用）
    If Not Response Is Nothing Then
        Set Response.fClient = Nothing
    End If
    
    ' 2. 释放请求级对象
    If Not Request Is Nothing Then
        ' 清理 Request 内部大对象
        Request.ClearRequest
        Set Request = Nothing
    End If
    Set Response = Nothing
    Set Session = Nothing
    If Not Cookies Is Nothing Then
        Cookies.Clear
        Set Cookies = Nothing
    End If
    
    ' 3. 重置标志
    IsAbort = False
    IsSkipNextMiddleWare = False
    
    ' 4. 清理 UserData（每次请求独立）
    If Not UserData Is Nothing Then
        UserData.RemoveAll
    End If
End Sub

' 连接关闭时释放所有对象（包括长生命周期对象）
Friend Sub Release()
    Set Request = Nothing
    Set Response = Nothing
    Set Session = Nothing
    Set Cookies = Nothing
    Set Db = Nothing
    If Not SSE Is Nothing Then
        Set SSE = Nothing
    End If
    Set ClientInfo = Nothing
    Set Server = Nothing
    Set App = Nothing
    If Not UserData Is Nothing Then
        UserData.RemoveAll
        Set UserData = Nothing
    End If
    Set TimeUse = Nothing
End Sub

Private Sub Class_Terminate()
    ' 保险清理：当对象引用计数归零时释放内部引用
    Set Request = Nothing
    Set Response = Nothing
    Set Session = Nothing
    Set Cookies = Nothing
    Set Db = Nothing
    Set SSE = Nothing
    Set ClientInfo = Nothing
    Set Server = Nothing
    Set App = Nothing
    If Not UserData Is Nothing Then
        UserData.RemoveAll
        Set UserData = Nothing
    End If
    Set TimeUse = Nothing
End Sub
