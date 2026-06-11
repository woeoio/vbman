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
'路由流程（按扩展名分流）：
'1，路径最后一段含点号（如 /css/style.css）→ 静态文件快速通道，跳过路由
'2，路径最后一段无点号（如 /api/users）→ 路由优先通道：
'   a. 前置中间件
'   b. 匹配用户路由
'   c. 自动路由（CallByName）
'   d. 静态文件兜底（目录重定向 + 默认文档）
'   e. 404


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
Private Const C_METHOD_DELETE As String = "DELETE"

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
    Dim item As New cHttpServerRouteItem
    Call item.Init(RouteName, Handler)
    Set Dic(n)(RouteName) = item
    Add = True
End Function

'路由处理入口
Friend Function Entry(Client As cClientCallback) As Boolean
    '    Debug.Print "OnDataArrival, Index=" & Index & ", bytesTotal=" & oCallback.dataTotal, Timer
    If App.LogMode <> 0 Then On Error GoTo EH
    Dim Request As cHttpServerRequest
    Dim Response As cHttpServerResponse
    With Client.Context
        Set Request = .Request
        Set Response = .Response
    End With
    Set RouteBefore = Client.Parent.RouteBefore
    Set RouteAfter = Client.Parent.RouteAfter
    
    ' ========== 按扩展名分流 ==========
    ' 路径最后一段含点号 → 静态文件快速通道（跳过所有路由）
    ' 路径最后一段无点号 → 路由优先通道（路由不匹配才走静态文件）
    Dim lastSlash As Long
    lastSlash = InStrRev(Request.PathInfo, "/")
    Dim lastSegment As String
    lastSegment = Mid(Request.PathInfo, lastSlash + 1)
    
    If InStr(lastSegment, ".") > 0 Then
        ' ===== 快速通道：含点号，视为静态文件 =====
        If Client.Context.Server.WebRoot <> "" Then
            If Client.Context.Server.IsStaticFile(Request.PathInfo) = True Then
                Response.File Request.PathInfo
                GoTo QH
            End If
        End If
        ' 不是已知静态文件，也尝试直接发送（可能是深层子目录的文件）
        If Client.Context.Server.WebRoot <> "" Then
            Response.File Request.PathInfo
            GoTo QH
        End If
        ' 无 WebRoot 则走路由
    End If
    
    ' ===== 路由优先通道：无点号，先路由后静态文件 =====
    '执行路由前置
    If RouteBefore.Entry(Client.Context) = False Then GoTo EH
    If Client.Context.fIsAbort = True Then GoTo QH
    '开始路由
    If Request.PathInfo = "/_dev" Then Response.test: Exit Function
    '匹配用户路由（支持参数路由 {param}）
    If Dic.Exists(Request.MethodName) = True Then
        Dim RD As Scripting.Dictionary, Handler As String
        Set RD = Dic(Request.MethodName)
        
        '1.精确匹配优先（O(1)，性能最高）
        If RD.Exists(Request.PathInfo) = True Then
            Dim exactItem As cHttpServerRouteItem
            Set exactItem = RD(Request.PathInfo)
            If Not exactItem.HasParams Then
                Handler = exactItem.Handler
                GoTo MatchController
            End If
        End If
        'ANY 字典精确匹配
        If DicAny.Exists(Request.PathInfo) = True Then
            Set exactItem = DicAny(Request.PathInfo)
            If Not exactItem.HasParams Then
                Handler = exactItem.Handler
                GoTo MatchController
            End If
        End If
        
        '2.参数路由模式匹配（遍历含参数的路由项）
        Dim RouteKey As Variant
        Dim routeItem As cHttpServerRouteItem
        Dim routeParams As Scripting.Dictionary
        
        '在当前方法字典中匹配
        For Each RouteKey In RD.Keys
            Set routeItem = RD(RouteKey)
            If routeItem.HasParams Then
                Set routeParams = New Scripting.Dictionary
                routeParams.CompareMode = TextCompare
                If routeItem.Match(Request.PathInfo, routeParams) = True Then
                    Set Request.RouteParams = routeParams
                    Handler = routeItem.Handler
                    GoTo MatchController
                End If
            End If
        Next RouteKey
        
        '在 ANY 字典中匹配
        For Each RouteKey In DicAny.Keys
            Set routeItem = DicAny(RouteKey)
            If routeItem.HasParams Then
                Set routeParams = New Scripting.Dictionary
                routeParams.CompareMode = TextCompare
                If routeItem.Match(Request.PathInfo, routeParams) = True Then
                    Set Request.RouteParams = routeParams
                    Handler = routeItem.Handler
                    GoTo MatchController
                End If
            End If
        Next RouteKey
    End If
    '自动路由（不如手动路由性能高）
    If AutoRoute = True Then
        ' 匹配脚本文件 asp/vbs
        ' 匹配控制器对象
        If Request.PathInfoController <> "" And Request.PathInfoAction <> "" Then
            If RunHandlerController(Request.PathInfoController, Request.PathInfoAction, Client.Context) = True Then GoTo QH
        End If
    End If
    
    ' 所有路由均未匹配，尝试静态文件兜底（目录重定向 + 默认文档）
    If Client.Context.Server.WebRoot <> "" Then
        ' 目录重定向：路径不以 / 结尾且对应物理目录，302 到带 / 的路径
        If Right(Request.PathInfo, 1) <> "/" And Request.PathInfo <> "/" Then
            Dim checkPath As String
            checkPath = ToolsFso.AutoCompleteFullPath(Client.Context.Server.WebRoot & Request.PathInfo, False)
            If ToolsFso.Inst.FolderExists(checkPath) = True Then
                Response.Redirect Request.PathInfo & "/"
                GoTo QH
            End If
        End If
        ' 尝试作为静态文件/目录处理（含默认文档查找）
        If Client.Context.Server.IsStaticFile(Request.PathInfo) = True Then
            Response.File Request.PathInfo
            GoTo QH
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
    '忽略大小写
    '    Dic.CompareMode = TextCompare
    DicAny.CompareMode = TextCompare
    DicGet.CompareMode = TextCompare
    DicPost.CompareMode = TextCompare
    DicPut.CompareMode = TextCompare
    DicDelete.CompareMode = TextCompare
    DicGroup.CompareMode = TextCompare
    
End Sub
