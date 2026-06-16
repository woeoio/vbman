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
Public Server As cHttpServerSvr
Public ClientInfo As cHttpServerClientInfo
Public Request As cHttpServerRequest
Public Response As cHttpServerResponse
Public Session As cHttpServerSession
Public Cookies As cHttpServerCookies
Public TimeUse As cTimeUse
Public Db As cDataBase

Public SSE As cSSEContext

' 服务器级统计对象引用（来自 Server.Statistics，仅引用不持有）
Public Statistics As cHttpServerStatistics
'用户数据，用于整个请求生命周期传递数据
Public UserData As Scripting.Dictionary

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
' 请求级短生命周期对象：Request/Response/Cookies/Session/UserData
' 连接级长生命周期对象：SSE/TimeUse/Server/App/Db/ClientInfo 保留不释放
Friend Sub ReleaseRequest()
    On Error Resume Next
    
    ' 1. 先断开 Response 对 ClientCallback 的反向引用（最关键的循环引用）
    If Not Response Is Nothing Then
        Set Response.fClient = Nothing
    End If
    
    ' 2. 释放请求级对象
    If Not Request Is Nothing Then
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
    ' Statistics 是服务器级单例引用，不在此释放
End Sub

' 连接关闭时释放所有对象（包括长生命周期对象）
Friend Sub Release()
    Set Request = Nothing
    Set Response = Nothing
    Set Session = Nothing
    Set Cookies = Nothing
    Set Db = Nothing
    Set SSE = Nothing
    Set ClientInfo = Nothing
    Set Server = Nothing
    Set App = Nothing
    ' Statistics 是服务器级单例引用，不在此释放
    If Not UserData Is Nothing Then
        UserData.RemoveAll
        Set UserData = Nothing
    End If
    Set TimeUse = Nothing
End Sub

Private Sub Class_Initialize()
    ' 显式创建子对象（替代 As New 隐式创建）
    Set TimeUse = New cTimeUse
    Set SSE = New cSSEContext
    Set UserData = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set Request = Nothing
    Set Response = Nothing
    Set Session = Nothing
    Set Cookies = Nothing
    Set Db = Nothing
    Set SSE = Nothing
    Set ClientInfo = Nothing
    Set Server = Nothing
    Set App = Nothing
    ' Statistics 是服务器级单例引用，不在此释放
    If Not UserData Is Nothing Then
        UserData.RemoveAll
        Set UserData = Nothing
    End If
    Set TimeUse = Nothing
End Sub
