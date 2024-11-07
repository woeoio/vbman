Attribute VB_Name = "ToolsFso"
Option Explicit

'Library Scripting
'    C:\Windows\SysWOW64\scrrun.dll
'    Microsoft Scripting Runtime

Public Inst As New Scripting.FileSystemObject
'申明
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long

Public Function AutoCompleteFullPath(Path As String, Optional IsFile As Boolean) As String
    '支持路径自动补全
    AutoCompleteFullPath = Path
    If ToolsStr.HasStr(":\", Path) = 0 Then AutoCompleteFullPath = App.Path & "\" & Path
    AutoCompleteFullPath = ClearSpan(AutoCompleteFullPath, IsFile)
End Function

Public Function ClearSpan(Path As String, Optional IsFile As Boolean) As String
    Dim DirEnd As String: If IsFile = False Then DirEnd = "\"
    ClearSpan = Replace(Path & DirEnd, "\\", "\")
    ClearSpan = Replace(ClearSpan, "//", "\")
    ClearSpan = Replace(ClearSpan, "\/", "\")
    ClearSpan = Replace(ClearSpan, "/", "\")
End Function
Public Function AutoMakeDir(ByVal Path As String, Optional IsFile As Boolean) As String
    If IsFile = True Then Path = Inst.GetParentFolderName(Path)
    Dim DirPart As String: DirPart = ClearSpan(Path)
    '使用：
    MakeSureDirectoryPathExists DirPart
    AutoMakeDir = DirPart
End Function


Public Function MakeNewFileFulPath(FileSrc As String, AppendFix As String, Optional JoinStr As String = "_") As String
    Dim fd As String: fd = Inst.GetParentFolderName(FileSrc)
    Dim Fn As String: Fn = Inst.GetBaseName(FileSrc)
    Dim Fe As String: Fe = Inst.GetExtensionName(FileSrc)
    MakeNewFileFulPath = Fn & JoinStr & AppendFix & "." & Fe
End Function


Public Sub Test()
    Dim r As Scripting.Folder
    Dim rDirs As Scripting.Folders
    Dim rFiles As Scripting.Files
    Set r = Inst.GetFolder("D:\code\_html\easyweb\spa-3.1.8")
    Set rDirs = r.SubFolders
    Set rFiles = r.Files
    Dim x
    For Each x In rDirs
        Debug.Print x.Name
    Next
    For Each x In rFiles
        Debug.Print x.Name
    Next
End Sub
