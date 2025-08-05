VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim pWebRoot As String
Dim rootFiles As New Scripting.Dictionary
Dim rootDirs As New Scripting.Dictionary

Dim Fso As New Scripting.FileSystemObject

Public IP As String
Public Port As Long


Public Property Get WebRoot() As String
    WebRoot = pWebRoot
End Property

Public Property Let WebRoot(v As String)
    If v = "" Then Exit Property
    pWebRoot = ToolsFso.AutoCompleteFullPath(v)
    ToolsFso.AutoMakeDir pWebRoot
    '如果设置了web目录，则把根目录的文件夹名称和文件名放到字典，给路由判断是否静态文件
    Dim r As Scripting.Folder
    Dim rDirs As Scripting.Folders
    Dim rFiles As Scripting.Files
    Set r = Fso.GetFolder(pWebRoot)
    Set rDirs = r.SubFolders
    Set rFiles = r.Files
    Dim x
    For Each x In rDirs
        rootDirs.Add "/" & x.Name, vbNull
    Next
    For Each x In rFiles
        rootFiles.Add "/" & x.Name, vbNull
    Next
End Property

Friend Function IsStaticFile(Path As String) As Boolean
    Dim AppendFix As String: AppendFix = LCase(Right(Path, 4))
    If AppendFix = ".asp" Then Exit Function
    If AppendFix = ".vbs" Then Exit Function
    If rootFiles.Exists(Path) = True Then IsStaticFile = True: Exit Function
    Dim x, p, l As Long
    For Each x In rootDirs.Keys()
        l = Len(x)
        p = Left(Path, l)
        If p = x Then IsStaticFile = True: Exit Function
    Next
End Function

Private Sub Class_Initialize()
    '设置大小写不敏感
    rootFiles.CompareMode = TextCompare
    rootDirs.CompareMode = TextCompare
End Sub
