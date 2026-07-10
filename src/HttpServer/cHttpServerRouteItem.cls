VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerRouteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'路由项：支持精确匹配和参数路由（如 /api/user/{id}）

'路由模式，如 "/api/user/{id}"
Public Pattern As String
'处理器，如 "UserController@Show"
Public Handler As String
'参数名数组，如 Array("id")，无参数时为空数组
Public ParamNames As Variant
'是否含参数（用于区分精确匹配和模式匹配）
Public HasParams As Boolean
'预编译的路径段数组（提升匹配性能）
Private Segments As Variant
'路径段数
Private SegCount As Long

''初始化路由项
Public Sub Init(ByVal RoutePattern As String, ByVal RouteHandler As String)
    Pattern = RoutePattern
    Handler = RouteHandler
    
    '去掉前导 /
    Dim raw As String
    If Left(RoutePattern, 1) = "/" Then raw = Mid(RoutePattern, 2) Else raw = RoutePattern
    
    '根路径特殊处理
    If raw = "" Then
        Segments = Array()
        SegCount = 0
        ParamNames = Array()
        HasParams = False
        Exit Sub
    End If
    
    '拆分为路径段
    Segments = Split(raw, "/")
    SegCount = UBound(Segments) + 1
    
    '提取参数名 {xxx} 和 {xxx?}
    Dim i As Long, cnt As Long
    Dim names() As String
    ReDim names(SegCount - 1)
    cnt = 0
    For i = 0 To SegCount - 1
        Dim seg As String: seg = CStr(Segments(i))
        If Left(seg, 1) = "{" And Right(seg, 1) = "}" Then
            Dim pName As String
            pName = Mid(seg, 2, Len(seg) - 2)
            '可选参数 {id?}
            If Right(pName, 1) = "?" Then pName = Left(pName, Len(pName) - 1)
            names(cnt) = pName
            cnt = cnt + 1
        End If
    Next i
    
    If cnt > 0 Then
        ReDim Preserve names(cnt - 1)
        ParamNames = names
        HasParams = True
    Else
        ParamNames = Array()
        HasParams = False
    End If
End Sub

''匹配路径，返回 True 表示匹配成功，params 输出提取的参数
''调用方需负责创建 params（New Dictionary）并传入
Public Function Match(ByVal PathInfo As String, params As Dictionary) As Boolean
    '去掉前导 /
    Dim raw As String
    If Left(PathInfo, 1) = "/" Then raw = Mid(PathInfo, 2) Else raw = PathInfo
    
    '根路径特殊处理
    If raw = "" And Pattern = "/" Then
        Match = True
        Exit Function
    End If
    
    Dim pathSegs As Variant
    pathSegs = Split(raw, "/")
    Dim pathCount As Long: pathCount = UBound(pathSegs) + 1
    
    '段数必须一致
    If pathCount <> SegCount Then Exit Function
    
    '逐段匹配
    Dim i As Long, paramIdx As Long
    paramIdx = 0
    For i = 0 To SegCount - 1
        Dim patternSeg As String: patternSeg = CStr(Segments(i))
        Dim valueSeg As String: valueSeg = CStr(pathSegs(i))
        
        If Left(patternSeg, 1) = "{" And Right(patternSeg, 1) = "}" Then
            '参数段 - 匹配任意非空值
            If valueSeg = "" Then Exit Function
            If paramIdx <= UBound(ParamNames) Then
                params(CStr(ParamNames(paramIdx))) = valueSeg
                paramIdx = paramIdx + 1
            End If
        Else
            '固定段 - 必须精确匹配（忽略大小写）
            If StrComp(patternSeg, valueSeg, vbTextCompare) <> 0 Then Exit Function
        End If
    Next i
    
    Match = True
End Function
