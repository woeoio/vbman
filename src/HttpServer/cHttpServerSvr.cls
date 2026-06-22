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
' As New→显式：改 As New 为显式创建
Dim rootFiles As Scripting.Dictionary
Dim rootDirs As Scripting.Dictionary

Dim Fso As Scripting.FileSystemObject

Public IP As String
Public Port As Long

' 默认文档列表
Dim m_DefaultDocuments As Collection

' 静态文件内容缓存（懒缓存：首次访问且未超限才加入）
Dim m_FileCache As Scripting.Dictionary
' 缓存上限：单文件最大缓存字节数（默认 1MB，0=禁用缓存）
Public MaxCacheFileSize As Long
' 缓存 TTL：秒数（默认 300=5分钟，0=永不过期）
Public CacheTTLSeconds As Long


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
    ' 根路径视为静态文件（用于默认文档查找）
    If Path = "/" Then IsStaticFile = True: Exit Function
    If rootFiles.Exists(Path) = True Then IsStaticFile = True: Exit Function
    Dim x, p, l As Long
    For Each x In rootDirs.Keys()
        l = Len(x)
        p = Left(Path, l)
        If p = x Then IsStaticFile = True: Exit Function
    Next
End Function

Private Sub Class_Initialize()
    ' 显式创建子对象（替代 As New 隐式创建）
    Set rootFiles = New Scripting.Dictionary
    Set rootDirs = New Scripting.Dictionary
    Set m_FileCache = New Scripting.Dictionary
    m_FileCache.CompareMode = TextCompare
    Set Fso = New Scripting.FileSystemObject
    '设置大小写不敏感
    rootFiles.CompareMode = TextCompare
    rootDirs.CompareMode = TextCompare
    ' 默认文档列表
    Set m_DefaultDocuments = New Collection
    With m_DefaultDocuments
        .Add "index.html"
        .Add "index.htm"
        .Add "default.html"
        .Add "default.htm"
    End With
    ' 缓存默认配置
    MaxCacheFileSize = 1048576  ' 1MB
    CacheTTLSeconds = 300       ' 5分钟
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not rootFiles Is Nothing Then
        rootFiles.RemoveAll
        Set rootFiles = Nothing
    End If
    If Not rootDirs Is Nothing Then
        rootDirs.RemoveAll
        Set rootDirs = Nothing
    End If
    If Not m_FileCache Is Nothing Then
        ClearFileCache
        Set m_FileCache = Nothing
    End If
    Set Fso = Nothing
    Set m_DefaultDocuments = Nothing
End Sub

' 添加自定义默认文档
Public Sub AddDefaultDocument(FileName As String)
    Dim i As Long
    For i = 1 To m_DefaultDocuments.Count
        If LCase(m_DefaultDocuments(i)) = LCase(FileName) Then Exit Sub
    Next
    m_DefaultDocuments.Add FileName
End Sub

' 获取默认文档列表
Friend Function GetDefaultDocuments() As Collection
    Set GetDefaultDocuments = m_DefaultDocuments
End Function

'=========================================================================
' 目录结构缓存刷新
'=========================================================================

' 刷新根目录结构缓存（手动调用，部署新文件后使用）
Friend Sub RefreshDirCache()
    On Error Resume Next
    rootFiles.RemoveAll
    rootDirs.RemoveAll
    If pWebRoot = "" Then Exit Sub
    Dim r As Scripting.Folder
    Set r = Fso.GetFolder(pWebRoot)
    Dim x
    For Each x In r.SubFolders
        rootDirs.Add "/" & x.Name, vbNull
    Next
    For Each x In r.Files
        rootFiles.Add "/" & x.Name, vbNull
    Next
End Sub

'=========================================================================
' 文件内容缓存（懒缓存机制）
'=========================================================================

' 获取缓存条目（返回 cHttpServerFileCacheEntry 或 Nothing）
Friend Function GetCachedFile(sPath As String) As cHttpServerFileCacheEntry
    If m_FileCache Is Nothing Then Exit Function
    If m_FileCache.Exists(sPath) = False Then Exit Function
    Set GetCachedFile = m_FileCache(sPath)
End Function

' 添加文件到缓存（仅当未超出大小上限）
Friend Sub AddFileToCache(sPath As String, Data() As Byte, ETagValue As String, LastModTime As Date)
    If MaxCacheFileSize <= 0 Then Exit Sub  ' 缓存已禁用
    Dim fileSize As Long: fileSize = UBound(Data) + 1
    If fileSize > MaxCacheFileSize Then Exit Sub  ' 超出单文件上限
    Dim entry As cHttpServerFileCacheEntry
    Set entry = New cHttpServerFileCacheEntry
    entry.Content = Data
    entry.ETag = ETagValue
    entry.LastModified = LastModTime
    entry.FileSize = fileSize
    entry.CachedAt = Now
    If m_FileCache.Exists(sPath) Then
        m_FileCache.Remove sPath
    End If
    m_FileCache.Add sPath, entry
End Sub

' 移除指定文件的缓存条目
Friend Sub InvalidateFileCache(sPath As String)
    On Error Resume Next
    If m_FileCache.Exists(sPath) Then m_FileCache.Remove sPath
End Sub

' 清空所有文件内容缓存
Friend Sub ClearFileCache()
    On Error Resume Next
    If m_FileCache Is Nothing Then Exit Sub
    Dim key As Variant
    For Each key In m_FileCache.Keys
        Set m_FileCache(key) = Nothing
    Next
    m_FileCache.RemoveAll
End Sub

' 计算 ETag：基于文件修改时间+大小（locale无关）
Friend Function ComputeETag(lastModified As Date, fileSize As Long) As String
    ComputeETag = """" & _
        Right(Year(lastModified), 4) & _
        Right("0" & Month(lastModified), 2) & _
        Right("0" & Day(lastModified), 2) & _
        Right("0" & Hour(lastModified), 2) & _
        Right("0" & Minute(lastModified), 2) & _
        Right("0" & Second(lastModified), 2) & _
        "-" & fileSize & """"
End Function

' 格式化 HTTP 日期（RFC 2616: "Day, DD Mon YYYY HH:MM:SS GMT"）
' 使用硬编码英文星期/月份名，确保 locale 无关
Friend Function FormatHttpDate(d As Date) As String
    Dim dayNames As Variant
    dayNames = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    Dim monthNames As Variant
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    Dim wday As Long: wday = Weekday(d) - 1  ' VB6 Weekday: 1=Sun
    Dim mon As Long: mon = Month(d) - 1
    FormatHttpDate = dayNames(wday) & ", " & _
        Right("0" & Day(d), 2) & " " & monthNames(mon) & " " & _
        Year(d) & " " & _
        Right("0" & Hour(d), 2) & ":" & _
        Right("0" & Minute(d), 2) & ":" & _
        Right("0" & Second(d), 2) & " GMT"
End Function
