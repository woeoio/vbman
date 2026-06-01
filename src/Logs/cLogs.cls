VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'作者：邓伟
'联系：QQ 215879458
'网站：http://vb6.pro
'说明：原创模块，转载请保留本注释块
'早年写的一个日志函数，有时间优化以下，比如支持内存缓冲，定时异步写入等。

'20240706 加入了 OutputDebugString 函数，暂时先简陋的实现一下远程实时调试。
'日志输出到 DbgView 控制台参考：
'vb6使用 OutputDebugString 配合 debugview 工具进行调试输出
'https://vb6.pro/vi/1233

'20240706 加入了网页版远程调试，并支持日志级别控制，完善了代码逻辑和命名规范。
'用法：
'先设置一个 WebUserCode 的值，可以是你喜欢的任意字符串
'仅支持 字母和数字 组合，然后到 http://log.vb6.pro 去填写这个值
'也可以在网站得到一个随机字符串再回来这里设置 WebUserCode
'设置好之后，就可以在网站实时看到本模块的 WebView 函数提交的日志了

'日志级别
Public Enum EnumLogLevel
    LvInfo = 0
    LvWarn = 1
    LvDanger = 2
    LvError = 3
    LvDebugger = 4
    LvCustom = 5
End Enum

'日志文件名规则
Public Enum EnumLogFileNameRule
    None = 0
    ByMonth = 1
    ByDay = 2
    byUser = 3
End Enum

'当前设置的日志控制级别，级别值越高，日志输出越少
Public LogLevelControl  As EnumLogLevel
'设置全局的通道开关，可代替 save 函数的开关参数
Public SendToFormView As Boolean
Public SendToDbgView As Boolean
Public SendToWebView As Boolean

'用于回显到任意窗体，任何 Form 窗体把 listbox set 给这个对象就可以输出
'比如：Set LogTools.FormListBox = List1
Public FormListBox As VB.ListBox
Public FormListBoxMaxItem As Long
'日志文件根目录，为空则使用当前程序的目录
Public LogDir As String
'日志文件子目录，一般用于区分业务模块的日志
Public LogSubDir As String
'日志文件名称规则，默认按月份细分，也可以按日，或者用户定义
Public LogFileNameRule As EnumLogFileNameRule
'日志文件名称用户定义，支持任意细分深层目录和文件名
Public LogFileNameByUer As String
'设置网页调试使用的用户识别码
Public WebUserCode As String
Private Const WEB_URL As String = "http://log.vb6.pro/submit"


Public DebugPring As Boolean

Dim pvHostApp As Object

Dim FileNumber As Long
Dim LogData As String
Dim LevelData As Variant

Public CacheDatas As New Collection

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
'MakeSureDirectoryPathExists 支持长路径自动创建 多级文件夹
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long

Public Property Get ShowLogsViewer() As Boolean
    ShowLogsViewer = FLogs.Visible
End Property
Public Property Let ShowLogsViewer(Value As Boolean)
    If Value = True Then
        If FLogs.Visible = False Then FLogs.Show
    Else
        Unload FLogs
    End If
End Property

Public Function HostApp(o As Object) As cLogs
    Set HostApp = Me
    Set pvHostApp = o
End Function

Public Function View(Value As Boolean) As cLogs
    Set View = Me
    ShowLogsViewer = Value
End Function

Public Function DataLine(ByVal LogContent As String, Optional ByVal LogTitle As String, Optional LogLevel As EnumLogLevel = LvInfo) As cLogs
    Set DataLine = Data(LogContent & vbCrLf, LogTitle, LogLevel)
End Function
Public Function Data(ByVal LogContent As String, Optional ByVal LogTitle As String, Optional LogLevel As EnumLogLevel = LvInfo) As cLogs
Attribute Data.VB_UserMemId = 0
    Set Data = Me
    Dim Title As String: Title = MakeLogTitle(LogTitle, LogLevel)
    If Title = "" Then Exit Function
    CacheDatas.Add Array(Title, LogContent)
    If FLogs.Visible = True Then
        FLogs.AddLog Title & vbCrLf & LogContent
    End If
End Function

Private Function CacheLogData(LogContent As String, LogTitle As String, LogLevel As EnumLogLevel) As Boolean
    
    CacheLogData = True
End Function

Public Function ToFormView(LogStr As String, Optional LogLevel As EnumLogLevel) As String
    '用途：日志输出到窗体的列表控件
    If FormListBox Is Nothing Then Exit Function
    '    If MakeLogData(LogStr, LogLevel) = False Then Exit Function
    FormListBox.AddItem LogData, 0
    '移除超出最大条目数量
    If FormListBoxMaxItem < 9 Then FormListBoxMaxItem = 999
    If FormListBox.ListCount > FormListBoxMaxItem Then FormListBox.RemoveItem FormListBoxMaxItem
    ToFormView = LogData
End Function

Public Function ToWebView(LogStr As String, Optional LogLevel As EnumLogLevel) As String
    '用途：日志输出到网页控制台
    If WebUserCode = "" Then Exit Function
    '    If MakeLogData(LogStr, LogLevel) = False Then Exit Function
    Static HttpClient As Object
    ' 创建 WinHttpRequest 对象
    Set HttpClient = CreateObject("WinHttp.WinHttpRequest.5.1")
    '采用异步提交，不管成功与否。避免请求等待浪费时间
    With HttpClient
        'Sub SetTimeouts(ResolveTimeout As Long, ConnectTimeout As Long, SendTimeout As Long, ReceiveTimeout As Long)
        .Open "POST", WEB_URL & "?code=" & WebUserCode, True
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send "log=" & LogData
        '        Debug.Print .ResponseText
    End With
    ToWebView = LogData
End Function

Public Function ToDbgView(LogStr As String, Optional LogLevel As EnumLogLevel) As String
    '用途：日志输出到 DbgView 控制台
    '加个前缀，方便 debugview 工具进行筛选过滤本类消息
    '    If MakeLogData(LogStr, LogLevel) = False Then Exit Function
    OutputDebugString "VBMAN LOG DATA：" & LogData
    ToDbgView = LogData
End Function

Public Function Save() As cLogs
    Set Save = Me
    If CacheDatas.count = 0 Then Exit Function
    '用途：日志输出到文件，支持其他通道开关
    Dim LogPath As String
    Dim LogFile As String
    '默认值处理
    If LogSubDir = "" Then LogSubDir = "system"
    If LogFileNameByUer = "" Then LogFileNameByUer = "vb.txt"
    If LogFileNameRule = None Then LogFileNameRule = ByDay
    
    '    If MakeLogData(LogStr, LogLevel) = False Then Exit Function
    
    If LogFileNameRule = byUser Then
        LogPath = LogSubDir & "\"
        LogFile = LogFileNameByUer
    Else
        Select Case LogFileNameRule
        Case EnumLogFileNameRule.ByMonth
            LogPath = Format$(Now, "yyyy") & "\"
            LogFile = "\" & Format$(Now, "yyyyMM") & ".txt"
        Case EnumLogFileNameRule.ByDay
            LogPath = Format$(Now, "yyyy") & "\" & Format$(Now, "MM") & "\"
            LogFile = Format$(Now, "yyyyMMdd") & ".txt"
        Case Else
            LogPath = Format$(Now, "yyyy") & "\" & Format$(Now, "MM") & "\"
            LogFile = LogFileNameByUer & ".txt"
        End Select
        LogPath = "\logs\" + LogSubDir + "\" + LogPath
    End If
    If LogDir = "" Then
        LogPath = pvHostApp.Path & "\" & LogPath
    Else
        LogPath = LogDir & "\" & LogPath
    End If
    If Dir(LogPath, vbDirectory) = "" Then                                      '判断文件夹是否存在
        MakeSureDirectoryPathExists LogPath                                     '创建文件夹
    End If
    FileNumber = FreeFile
    Open LogPath & "\" & LogFile For Append As #FileNumber
    Print #FileNumber, MakeLogContent()
    Close #FileNumber
End Function

Private Function MakeLogContent(Optional Span As String = vbCrLf) As String
    Dim i As Long, Txt As String
    For i = 1 To CacheDatas.count
        Txt = Txt & Join(CacheDatas.Item(i), Span)                              '& vbCrLf
    Next
    Set CacheDatas = Nothing
    MakeLogContent = Txt
End Function

Private Function MakeLogTitle(LogTitle As String, LogLevel As EnumLogLevel) As String
    '    If IsEmpty(LogStr) = True Then Exit Function
    '判断日志控制级别
    If LogLevel < LogLevelControl Then Exit Function
    'ok
    If IsEmpty(LevelData) = True Then
        LevelData = Array("INFO", "WARN", "DANGER", "ERROR")
    End If
    MakeLogTitle = "::[" & Now & "] [" & LevelData(LogLevel) & "] " & LogTitle
    '    If pvhostApp.LogMode = 0 And DebugPring = True Then Debug.Print LogData
End Function


Private Sub Class_Initialize()
    Set pvHostApp = App
End Sub
