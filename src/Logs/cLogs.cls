VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mListBox As ListBox
Public mListMax As Long                                                         '= 999
    Rem 配置保存到日志文件还是日志数据库，严格按标准日志格式
Public mLogSystem As String
    Rem 保存完整日志信息，和mListBox同步
Public mLogList As cCollection
Public mRootPath As String
Public mLogType As String
Public mLogSpan As String
    Rem 配置实时调试服务器
Public mSocket As String

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal dirPath As String) As Long

Public Sub Bind(LB As Object)
    Rem 绑定外部列表框对象
    Set mListBox = LB
End Sub

Public Function ShowAndSave(ByVal LogStr As String) As String
Attribute ShowAndSave.VB_UserMemId = 0
    Rem 默认方法，显示并保存日志
    Call Show(LogStr): ShowAndSave = Save(LogStr)
End Function


Public Function Show(ByVal LogStr As String) As String
    Rem 只显示
    If mListBox Is Nothing Then Exit Function
    If mListBox.ListCount > mListMax Then mListBox.RemoveItem mListMax
    Show = "【" & Now & "】" & LogStr
    mListBox.AddItem Show, 0
End Function

Public Function Save(ByVal LogStr As String) As String
    Rem 只保存到文件
    Rem 可优化为内存缓冲，并持续打开磁盘文件，来提高性能
    Dim LogPath As String
    Dim LogFile As String
    LogStr = "【" & Now & "】" & LogStr
    Select Case mLogSpan
    Case "m"
        LogPath = Format$(Now, "yyyy") & "\"
        LogFile = "\" & Format$(Now, "yyyyMM") & ".txt"
    Case "d"
        LogPath = Format$(Now, "yyyy") & "\" & Format$(Now, "MM") & "\"
        LogFile = Format$(Now, "yyyyMMdd") & ".txt"
    Case Else
        LogPath = Format$(Now, "yyyy") & "\" & Format$(Now, "MM") & "\"
        LogFile = mLogSpan & ".txt"
    End Select
    LogPath = "\logs\" + mLogType + "\" + LogPath
    If mRootPath = "" Then
        LogPath = App.Path & LogPath
    Else
        LogPath = mRootPath & LogPath
    End If
    If Dir(LogPath, vbDirectory) = "" Then                                      '判断文件夹是否存在
        MakeSureDirectoryPathExists LogPath                                     '创建文件夹
    End If
    Open LogPath & LogFile For Append As #1
    Print #1, LogStr
    Close #1
    Save = LogStr
End Function

Private Sub Class_Initialize()
    mLogType = "system"
    mLogSpan = "d"
    mListMax = 999
    Set mLogList = vbRichClient5.New_c.JSONArray
End Sub

'Public Function OnStartPage(MyScriptingContext As ScriptingContext)
    Rem HTTP生命周期控制：开始
    '    Call Common.SetAspContext(MyScriptingContext)
'End Function

Public Function OnEndPage()
    Rem HTTP生命周期控制：结束
    '    Call Common.DelAspContext
End Function

