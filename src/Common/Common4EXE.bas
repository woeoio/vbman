Attribute VB_Name = "Common"
Option Explicit

Public Lang As New cLang

' 在标准模块中添加
Public g_bRunningAsEXE As Boolean

Sub Main()
    ' 检测是否作为独立EXE启动
    If App.StartMode = vbSModeStandalone Then
        g_bRunningAsEXE = True
        ' 显示主窗体
        Fmain.Show
    Else
        g_bRunningAsEXE = False
        ' 作为COM服务器运行
    End If
End Sub

Public Function Version(Optional HostApp As Object) As String
    If HostApp Is Nothing Then Set HostApp = App
    Version = "v" & HostApp.Major & "." & HostApp.Minor & "." & HostApp.Revision
End Function

