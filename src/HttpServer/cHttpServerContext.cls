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
