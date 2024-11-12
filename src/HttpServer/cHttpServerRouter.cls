VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'路由才是开发者和框架之间的桥梁
'路由流程：
'1，如果没有自动路由，则先从方法字典匹配
'2，如果开了自动路由，则尝试执行路由方法
'3，如果没有路由成功，则判读是否静态文件


Dim RouteBefore As cHttpServerRouteBefore
Dim RouteAfter As cHttpServerRouterAfter

Dim LastError As String

Public AutoRoute As Boolean

Public Enum EnumRouteMethod
    Any_ = 0
    OnlyPost = 1
    OnlyGet = 2
    OnlyPut = 3
    OnlyDelete = 4
End Enum

Private Const C_METHOD_ANY As String = "ANY"
Private Const C_METHOD_GET As String = "GET"
Private Const C_METHOD_POST As String = "POST"
Private Const C_METHOD_PUT As String = "PUT"
Private Const C_METHOD_DELETE As String = "DELTET"

'控制器字典
Dim Controllers As New Scripting.Dictionary
'方法字典
Dim Dic As New Scripting.Dictionary
Dim DicAny As New Scripting.Dictionary
Dim DicGet As New Scripting.Dictionary
Dim DicPost As New Scripting.Dictionary
Dim DicPut As New Scripting.Dictionary
Dim DicDelete As New Scripting.Dictionary
'路由组
Dim DicGroup As New Scripting.Dictionary

''开发者定义的控制器字典注册函数
Public Function Reg(ControllerName As String, Controller As Object) As Boolean
    '    If Controllers.Exists(ControllerName) = True Then Err.Raise 1, , "ControllerName Exists：" & ControllerName
    Set Controllers(ControllerName) = Controller
    Reg = True
End Function
'开发者定义的路由处理，如果没有用户路由，则执行自动路由
Public Function Add(RouteName As String, Handler As String, Optional MethodLimit As EnumRouteMethod = Any_) As Boolean
    Dim n  As String: n = ToolsHttp.MapMethodName(MethodLimit)
    '    If Dic(n).Exists(RouteName) = True Then Err.Raise 1, , "RouteName Exists：" & RouteName
    Dic(n)(RouteName) = Handler
    '    Dic(n).Add RouteName, Handler
    Add = True
End Function

'路由处理入口
Friend Function Entry(Client As cClientCallback) As Boolean
    '    Debug.Print "OnDataArrival, Index=" & Index & ", bytesTotal=" & oCallback.dataTotal, Timer
    If App.LogMode <> 0 Then On Error GoTo EH
    Dim Request As cHttpServerRequest
    Dim Response As cHttpServerResponse
    '    Dim Session As cHttpServerSession
    '    Dim Cookies As cHttpServerCookies
    With Client.Context
        Set Request = .Request
        Set Response = .Response
    End With
    Set RouteBefore = Client.Parent.RouteBefore
    Set RouteAfter = Client.Parent.RouteAfter
    '判断是否静态文件
    If Client.Context.Server.WebRoot <> "" Then
        If Client.Context.Server.IsStaticFile(Request.PathInfo) = True Then
            Response.File Request.PathInfo
            GoTo QH
        End If
    End If
    '执行路由前置
    If RouteBefore.Entry(Client.Context) = False Then GoTo EH
    If Client.Context.fIsAbort = True Then GoTo QH
    '开始路由
    If Request.PathInfo = "/_dev" Then Response.Test: Exit Function
    '匹配用户路由
    If Dic.Exists(Request.MethodName) = True Then
        Dim RD As Scripting.Dictionary, Handler As String
        Set RD = Dic(Request.MethodName)
        If RD.Exists(Request.PathInfo) = True Then Handler = RD(Request.PathInfo): GoTo MatchController
        If DicAny.Exists(Request.PathInfo) = True Then Handler = DicAny(Request.PathInfo): GoTo MatchController
    End If
    '自动路由（不如手动路由性能高）
    If AutoRoute = True Then
        ' 匹配脚本文件 asp/vbs
        ' 匹配控制器对象
        If Request.PathInfoController <> "" And Request.PathInfoAction <> "" Then
            If RunHandlerController(Request.PathInfoController, Request.PathInfoAction, Client.Context) = True Then GoTo QH
        End If
    End If
    '输出
    Response.State404 "404 Not Found"
    Exit Function
MatchController:
    ' 匹配脚本文件 asp/vbs
    ' 匹配控制器对象
    If ToolsStr.HasStr("@", Handler) > 0 Then
        Dim CN As String, AN As String
        ToolsArray.DeArray Split(Handler & "@@", "@"), CN, AN
        If CN <> "" And AN <> "" Then
            If RunHandlerController(CN, AN, Client.Context) = True Then GoTo QH
        Else
            LastError = "Handler not good: " & Handler
        End If
    Else
        LastError = "No Handler for route: " & Request.PathInfo
    End If
    GoTo EH
QH:
    If LastError = "" Then
        '执行后置路由
        Call RouteAfter.Entry(Client.Context)
    Else
        GoTo EH
    End If
    Entry = True
    Exit Function
EH:
    If LastError = "" Then LastError = ERR.Number & "#" & ERR.Description
    Call Client.Context.Response.State500(LastError)
    LastError = ""
    'log
End Function
Private Function RunHandlerController(ControllerName As String, ActionName As String, ctx As cHttpServerContext) As Boolean
    On Error GoTo EH
    If Controllers.Exists(ControllerName) = True Then
        Dim c As Object: Set c = Controllers(ControllerName)
        CallByName c, ActionName, VbMethod, ctx
        RunHandlerController = True
    Else
        '控制器不存在
        LastError = "ControllerName not registered: " & ControllerName
    End If
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
End Function
Private Function RunHandlerAsp(Request As cHttpServerRequest) As Boolean
    '
End Function
Private Function RunHandlerVbs(Request As cHttpServerRequest) As Boolean
    '
End Function

Private Sub Class_Initialize()
    With Dic
        .Add C_METHOD_ANY, DicAny
        .Add C_METHOD_GET, DicGet
        .Add C_METHOD_POST, DicPost
        .Add C_METHOD_PUT, DicPut
        .Add C_METHOD_DELETE, DicDelete
    End With
    Controllers.CompareMode = TextCompare
End Sub
