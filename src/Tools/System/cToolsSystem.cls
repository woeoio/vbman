VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'API声明：用于以管理员身份运行程序
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

'以管理员身份运行指定的程序
'参数:
'   Path - 程序的完整路径，或 App 对象
'   ParamArray StartArgs - 可选，传递给程序的命令行参数
'返回值:
'   返回 ShellExecute 的执行结果
Public Function RunAsAdmin(ByRef Path As Variant, ParamArray StartArgs() As Variant) As Long
    Const SW_SHOWNORMAL = 1
    Const SW_SHOW = 5
    
    Dim sFilePath As String
    Dim sParameters As String
    
    '处理 Path 参数：支持 App 对象或直接传入路径字符串
    If IsObject(Path) = True Then
        sFilePath = Path.Path & "\" & Path.EXEName & ".exe"
    Else
        sFilePath = Path
    End If
    
    '处理 StartArgs 参数
    Dim A As Variant: A = StartArgs
    If Not IsEmpty(A) Then
        sParameters = Join(A, " ")
    End If
    
    '检查文件路径是否为空
    If Trim(sFilePath) = "" Then
        Exit Function
    End If
    
    '调用 ShellExecute，使用 "runas" 操作以管理员身份运行
    '参数说明：
    '   hWnd - 0 (桌面窗口句柄)
    '   lpOperation - "runas" (以管理员身份运行)
    '   lpFile - 程序完整路径
    '   lpParameters - 命令行参数
    '   lpDirectory - 0 (使用程序所在目录)
    '   nShowCmd - SW_SHOWNORMAL (正常显示窗口)
    RunAsAdmin = ShellExecute(0, "runas", sFilePath, sParameters, 0, SW_SHOWNORMAL)
End Function



