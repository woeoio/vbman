VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=====================================================================
' cDialog.twin - Windows API 对话框对象类
' 支持打开/保存文件对话框、文件夹选择对话框、后缀名过滤

' 作者：215879458@qq.com
' 版本：v1.0.2.0
'=====================================================================

Option Explicit

'=====================================================================
' Windows API 声明
'=====================================================================
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
'=====================================================================
' 方法：打开文件（使用 ShellExecute）
'=====================================================================
#If VBA7 Then
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'=====================================================================
' 常量定义
'=====================================================================
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_ENABLEINCLUDENOTIFY = &H400000
Private Const OFN_ENABLESIZING = &H800000

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_BROWSEINCLUDEURLS = &H80
Private Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Const BIF_NOTRANSLATETARGETS = &H400
Private Const BIF_SHAREABLE = &H8000

Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const MAX_PATH = 260
Private Enum LongPtr
    [_]
End Enum

'=====================================================================
' 用户定义类型
'=====================================================================
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As LongPtr
    nMaxFile As Long
    lpstrFileTitle As LongPtr
    nMaxFileTitle As Long
    lpstrInitialDir As LongPtr
    lpstrTitle As LongPtr
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'=====================================================================
' cDialog 类定义
'=====================================================================
Private Type cDialogData
    mstrDialogTitle As String
    mstrInitialDir As String
    mstrDefaultExt As String
    mstrFileName As String
    mstrFilter As String
    mblnMultiSelect As Boolean
    mblnOverwritePrompt As Boolean
    mblnPathMustExist As Boolean
    mblnFileMustExist As Boolean
    mblnHideReadOnly As Boolean
    mblnReadOnly As Boolean
    mblnNoChangeDir As Boolean
    mblnCreatePrompt As Boolean
    mblnNoReadOnlyReturn As Boolean
    mblnLongNames As Boolean
    mblnExplorer As Boolean
    mblnEnableSizing As Boolean
    mblnNewDialogStyle As Boolean
End Type

Private mData As cDialogData

    '=====================================================================
    ' 类初始化/终止
    '=====================================================================
    Private Sub Class_Initialize()
        With mData
            .mstrDialogTitle = ""
            .mstrInitialDir = ""
            .mstrDefaultExt = ""
            .mstrFileName = ""
            .mstrFilter = ""
            .mblnMultiSelect = False
            .mblnOverwritePrompt = True
            .mblnPathMustExist = True
            .mblnFileMustExist = True
            .mblnHideReadOnly = True
            .mblnReadOnly = False
            .mblnNoChangeDir = False
            .mblnCreatePrompt = False
            .mblnNoReadOnlyReturn = False
            .mblnLongNames = True
            .mblnExplorer = True
            .mblnEnableSizing = True
            .mblnNewDialogStyle = True
        End With
    End Sub

    Private Sub Class_Terminate()
        ' 清理代码（如果需要）
    End Sub

    '=====================================================================
    ' 属性：对话框标题
    '=====================================================================
    Public Property Let DialogTitle(ByVal vValue As String)
        mData.mstrDialogTitle = vValue
    End Property

    Public Property Get DialogTitle() As String
        DialogTitle = mData.mstrDialogTitle
    End Property

    '=====================================================================
    ' 属性：初始目录
    '=====================================================================
    Public Property Let InitialDir(ByVal vValue As String)
        mData.mstrInitialDir = vValue
    End Property

    Public Property Get InitialDir() As String
        InitialDir = mData.mstrInitialDir
    End Property

    '=====================================================================
    ' 属性：默认扩展名
    '=====================================================================
    Public Property Let DefaultExt(ByVal vValue As String)
        mData.mstrDefaultExt = vValue
    End Property

    Public Property Get DefaultExt() As String
        DefaultExt = mData.mstrDefaultExt
    End Property

    '=====================================================================
    ' 属性：文件名
    '=====================================================================
    Public Property Let FileName(ByVal vValue As String)
        mData.mstrFileName = vValue
    End Property

    Public Property Get FileName() As String
        FileName = mData.mstrFileName
    End Property

    '=====================================================================
    ' 属性：过滤器
    '=====================================================================
    Public Property Let Filter(ByVal vValue As String)
        mData.mstrFilter = vValue
    End Property

    Public Property Get Filter() As String
        Filter = mData.mstrFilter
    End Property

    '=====================================================================
    ' 属性：是否允许多选
    '=====================================================================
    Public Property Let MultiSelect(ByVal vValue As Boolean)
        mData.mblnMultiSelect = vValue
    End Property

    Public Property Get MultiSelect() As Boolean
        MultiSelect = mData.mblnMultiSelect
    End Property

    '=====================================================================
    ' 属性：是否覆盖提示
    '=====================================================================
    Public Property Let OverwritePrompt(ByVal vValue As Boolean)
        mData.mblnOverwritePrompt = vValue
    End Property

    Public Property Get OverwritePrompt() As Boolean
        OverwritePrompt = mData.mblnOverwritePrompt
    End Property

    '=====================================================================
    ' 属性：路径必须存在
    '=====================================================================
    Public Property Let PathMustExist(ByVal vValue As Boolean)
        mData.mblnPathMustExist = vValue
    End Property

    Public Property Get PathMustExist() As Boolean
        PathMustExist = mData.mblnPathMustExist
    End Property

    '=====================================================================
    ' 属性：文件必须存在
    '=====================================================================
    Public Property Let FileMustExist(ByVal vValue As Boolean)
        mData.mblnFileMustExist = vValue
    End Property

    Public Property Get FileMustExist() As Boolean
        FileMustExist = mData.mblnFileMustExist
    End Property

    '=====================================================================
    ' 属性：隐藏只读复选框
    '=====================================================================
    Public Property Let HideReadOnly(ByVal vValue As Boolean)
        mData.mblnHideReadOnly = vValue
    End Property

    Public Property Get HideReadOnly() As Boolean
        HideReadOnly = mData.mblnHideReadOnly
    End Property

    '=====================================================================
    ' 属性：只读
    '=====================================================================
    Public Property Let ReadOnly(ByVal vValue As Boolean)
        mData.mblnReadOnly = vValue
    End Property

    Public Property Get ReadOnly() As Boolean
        ReadOnly = mData.mblnReadOnly
    End Property

    '=====================================================================
    ' 属性：不改变当前目录
    '=====================================================================
    Public Property Let NoChangeDir(ByVal vValue As Boolean)
        mData.mblnNoChangeDir = vValue
    End Property

    Public Property Get NoChangeDir() As Boolean
        NoChangeDir = mData.mblnNoChangeDir
    End Property

    '=====================================================================
    ' 属性：创建提示
    '=====================================================================
    Public Property Let CreatePrompt(ByVal vValue As Boolean)
        mData.mblnCreatePrompt = vValue
    End Property

    Public Property Get CreatePrompt() As Boolean
        CreatePrompt = mData.mblnCreatePrompt
    End Property

    '=====================================================================
    ' 属性：启用新对话框样式
    '=====================================================================
    Public Property Let NewDialogStyle(ByVal vValue As Boolean)
        mData.mblnNewDialogStyle = vValue
    End Property

    Public Property Get NewDialogStyle() As Boolean
        NewDialogStyle = mData.mblnNewDialogStyle
    End Property

    '=====================================================================
    ' 方法：添加过滤器
    '=====================================================================
    Public Sub AddFilter(ByVal strDescription As String, ByVal strExtension As String)
        ' 格式：描述|扩展名
        If mData.mstrFilter <> "" Then
            mData.mstrFilter = mData.mstrFilter & "|"
        End If
        mData.mstrFilter = mData.mstrFilter & strDescription & "|" & strExtension
    End Sub

    '=====================================================================
    ' 方法：清除过滤器
    '=====================================================================
    Public Sub ClearFilter()
        mData.mstrFilter = ""
    End Sub

    '=====================================================================
    ' 方法：显示打开文件对话框
    '=====================================================================
    Public Function ShowOpen() As Variant
        Dim ofn As OPENFILENAME
        Dim strFile As String
        Dim strFilter As String
        Dim lngFlags As Long
        Dim lngResult As Long
        Dim lngFileHandle As Long
        Dim strResult As String

        ' 准备文件名缓冲区
        Dim nBufferSize As Long
        nBufferSize = 32768 ' 为多选分配32KB缓冲区
        If mData.mblnMultiSelect Then
            strFile = String$(nBufferSize, 0)
        Else
            strFile = String$(MAX_PATH, 0)
        End If
        If mData.mstrFileName <> "" Then
            If Len(mData.mstrFileName) < nBufferSize Then
                strFile = mData.mstrFileName & String$(nBufferSize - Len(mData.mstrFileName), 0)
            End If
        End If

        ' 准备过滤器
        strFilter = Replace(mData.mstrFilter, "|", vbNullChar)
        If strFilter <> "" Then
            strFilter = strFilter & vbNullChar
        End If

        ' 设置标志
        lngFlags = 0
        If mData.mblnMultiSelect Then lngFlags = lngFlags Or OFN_ALLOWMULTISELECT
        If mData.mblnOverwritePrompt Then lngFlags = lngFlags Or OFN_OVERWRITEPROMPT
        If mData.mblnPathMustExist Then lngFlags = lngFlags Or OFN_PATHMUSTEXIST
        If mData.mblnFileMustExist Then lngFlags = lngFlags Or OFN_FILEMUSTEXIST
        If mData.mblnHideReadOnly Then lngFlags = lngFlags Or OFN_HIDEREADONLY
        If mData.mblnReadOnly Then lngFlags = lngFlags Or OFN_READONLY
        If mData.mblnNoChangeDir Then lngFlags = lngFlags Or OFN_NOCHANGEDIR
        If mData.mblnCreatePrompt Then lngFlags = lngFlags Or OFN_CREATEPROMPT
        If mData.mblnLongNames Then lngFlags = lngFlags Or OFN_LONGNAMES
        If mData.mblnExplorer Then lngFlags = lngFlags Or OFN_EXPLORER
        If mData.mblnEnableSizing Then lngFlags = lngFlags Or OFN_ENABLESIZING

        ' 填充 OPENFILENAME 结构
        With ofn
            .lStructSize = LenB(ofn)
            .hwndOwner = 0
            .hInstance = App.hInstance
            .lpstrFilter = StrPtr(strFilter)
            .lpstrCustomFilter = 0
            .nMaxCustFilter = 0
            .nFilterIndex = 1
            .lpstrFile = StrPtr(strFile)
            .nMaxFile = Len(strFile)
            .lpstrFileTitle = StrPtr(Space$(MAX_PATH))
            .nMaxFileTitle = MAX_PATH
            .lpstrInitialDir = StrPtr(mData.mstrInitialDir)
            .lpstrTitle = StrPtr(mData.mstrDialogTitle)
            .flags = lngFlags
            .lpstrDefExt = StrPtr(mData.mstrDefaultExt)
            .lCustData = 0
            .lpfnHook = 0
            .lpTemplateName = 0
        End With

        ' 显示对话框
        lngResult = GetOpenFileName(ofn)

        If lngResult <> 0 Then
            ' 用户点击了"打开"
            Dim nNullPos As Long
            nNullPos = InStr(strFile, vbNullChar)
            If nNullPos > 0 Then
                strFile = Left$(strFile, nNullPos - 1)
            Else
                strFile = strFile
            End If

            If mData.mblnMultiSelect Then
                ' 处理多选
                If InStr(strFile, vbNullChar) > 0 Then
                    ' 多选模式：第一个空格前的部分是路径，后面是文件名
                    Dim strPath As String
                    Dim strFiles() As String
                    Dim i As Long
                    Dim strTemp As String

                    strPath = Mid$(strFile, 1, InStr(strFile, vbNullChar) - 1)
                    strTemp = Mid$(strFile, InStr(strFile, vbNullChar) + 1)

                    ' 分割文件名
                    Dim arrFiles() As String
                    Dim nCount As Long
                    Dim nPos As Long
                    Dim nStart As Long

                    ReDim arrFiles(0 To 100)
                    nCount = 0
                    nStart = 1

                    Do
                        nPos = InStr(nStart, strTemp, vbNullChar)
                        If nPos = 0 Then nPos = Len(strTemp) + 1

                        If nPos > nStart Then
                            arrFiles(nCount) = strPath & "\" & Mid$(strTemp, nStart, nPos - nStart)
                            nCount = nCount + 1
                        End If

                        nStart = nPos + 1
                    Loop Until nPos >= Len(strTemp) + 1

                    If nCount = 1 Then
                        ShowOpen = arrFiles(0)
                    Else
                        ReDim Preserve arrFiles(0 To nCount - 1)
                        ShowOpen = arrFiles
                    End If
                Else
                    ShowOpen = strFile
                End If
            Else
                ShowOpen = strFile
            End If

            ' 更新 FileName 属性
            If varType(ShowOpen) = vbString Then
                mData.mstrFileName = ShowOpen
            Else
                mData.mstrFileName = ""
            End If
        Else
            ShowOpen = ""
        End If
    End Function

    '=====================================================================
    ' 方法：选择文件（返回集合对象）
    '=====================================================================
    Public Function SelectFiles() As Collection
        Dim colFiles As Collection
        Dim varResult As Variant
        Dim i As Long

        Set colFiles = New Collection

        ' 调用 ShowOpen 获取结果
        varResult = Me.ShowOpen()

        If varType(varResult) = vbString Then
            ' 单选文件
            If varResult <> "" Then
                colFiles.Add varResult
            End If
        ElseIf IsArray(varResult) Then
            ' 多选文件
            For i = LBound(varResult) To UBound(varResult)
                If varResult(i) <> "" Then
                    colFiles.Add varResult(i)
                End If
            Next i
        End If

        Set SelectFiles = colFiles
    End Function

    '=====================================================================
    ' 方法：显示保存文件对话框
    '=====================================================================
    Public Function ShowSave() As String
        Dim ofn As OPENFILENAME
        Dim strFile As String
        Dim strFilter As String
        Dim lngFlags As Long
        Dim lngResult As Long
        
        ' 准备文件名缓冲区
        strFile = Space$(MAX_PATH)
        If mData.mstrFileName <> "" Then
            strFile = mData.mstrFileName & Space$(MAX_PATH - Len(mData.mstrFileName))
        End If
        
        ' 准备过滤器
        strFilter = Replace(mData.mstrFilter, "|", vbNullChar)
        If strFilter <> "" Then
            strFilter = strFilter & vbNullChar
        End If

        ' 设置标志
        lngFlags = 0
        If mData.mblnOverwritePrompt Then lngFlags = lngFlags Or OFN_OVERWRITEPROMPT
        If mData.mblnPathMustExist Then lngFlags = lngFlags Or OFN_PATHMUSTEXIST
        If mData.mblnHideReadOnly Then lngFlags = lngFlags Or OFN_HIDEREADONLY
        If mData.mblnNoChangeDir Then lngFlags = lngFlags Or OFN_NOCHANGEDIR
        If mData.mblnLongNames Then lngFlags = lngFlags Or OFN_LONGNAMES
        If mData.mblnExplorer Then lngFlags = lngFlags Or OFN_EXPLORER
        If mData.mblnEnableSizing Then lngFlags = lngFlags Or OFN_ENABLESIZING

        ' 填充 OPENFILENAME 结构
        With ofn
            .lStructSize = LenB(ofn)
            .hwndOwner = 0
            .hInstance = App.hInstance
            .lpstrFilter = StrPtr(strFilter)
            .lpstrCustomFilter = 0
            .nMaxCustFilter = 0
            .nFilterIndex = 1
            .lpstrFile = StrPtr(strFile)
            .nMaxFile = Len(strFile)
            .lpstrFileTitle = StrPtr(Space$(MAX_PATH))
            .nMaxFileTitle = MAX_PATH
            .lpstrInitialDir = StrPtr(mData.mstrInitialDir)
            .lpstrTitle = StrPtr(mData.mstrDialogTitle)
            .flags = lngFlags
            .lpstrDefExt = StrPtr(mData.mstrDefaultExt)
            .lCustData = 0
            .lpfnHook = 0
            .lpTemplateName = 0
        End With

        ' 显示对话框
        lngResult = GetSaveFileName(ofn)

        If lngResult <> 0 Then
            ' 用户点击了"保存"
            ShowSave = Left$(strFile, InStr(strFile, vbNullChar) - 1)
            mData.mstrFileName = ShowSave
        Else
            ShowSave = ""
        End If
    End Function

    '=====================================================================
    ' 方法：显示选择文件夹对话框
    '=====================================================================
    Public Function ShowBrowseForFolder() As String
        Dim bi As BROWSEINFO
        Dim pidl As Long
        Dim strPath As String
        Dim lngFlags As Long
        
        ' 初始化路径缓冲区
        strPath = Space$(MAX_PATH)
        
        ' 设置标志
        lngFlags = BIF_RETURNONLYFSDIRS
        If mData.mblnNewDialogStyle Then lngFlags = lngFlags Or BIF_USENEWUI
        
        ' 填充 BROWSEINFO 结构
        With bi
            .hwndOwner = 0
            .pidlRoot = 0
            .pszDisplayName = Space$(MAX_PATH)
            .lpszTitle = mData.mstrDialogTitle
            .ulFlags = lngFlags
            .lpfnCallback = 0
            .lParam = 0
            .iImage = 0
        End With
        
        ' 设置初始目录（需要通过回调函数实现，此处简化处理）
        
        ' 显示对话框
        pidl = SHBrowseForFolder(bi)
        
        If pidl <> 0 Then
            ' 获取选择的路径
            If SHGetPathFromIDList(pidl, strPath) <> 0 Then
                ShowBrowseForFolder = Left$(strPath, InStr(strPath, vbNullChar) - 1)
                mData.mstrInitialDir = ShowBrowseForFolder
            Else
                ShowBrowseForFolder = ""
            End If
            
            ' 释放内存
            CoTaskMemFree pidl
        Else
            ShowBrowseForFolder = ""
        End If
    End Function



Public Function OpenFile(ByVal strFilePath As String, Optional ByVal strOperation As String = "open") As Long
    OpenFile = ShellExecute(0, strOperation, strFilePath, vbNullString, vbNullString, 1)
End Function

    '=====================================================================
    ' 方法：选择文件并打开
    '=====================================================================
    Public Function SelectAndOpenFile() As Long
        Dim strFile As Variant
        Dim i As Long
        
        strFile = Me.ShowOpen()
        
        If varType(strFile) = vbString Then
            If strFile <> "" Then
                SelectAndOpenFile = Me.OpenFile(strFile)
            Else
                SelectAndOpenFile = 0
            End If
        ElseIf IsArray(strFile) Then
            ' 多选文件
            For i = LBound(strFile) To UBound(strFile)
                If strFile(i) <> "" Then
                    Me.OpenFile strFile(i)
                End If
            Next i
            SelectAndOpenFile = 1
        Else
            SelectAndOpenFile = 0
        End If
    End Function

    '=====================================================================
    ' 方法：获取文件路径（不带文件名）
    '=====================================================================
    Public Function GetFilePath(ByVal strFullPath As String) As String
        Dim nPos As Long
        
        nPos = InStrRev(strFullPath, "\")
        If nPos > 0 Then
            GetFilePath = Left$(strFullPath, nPos - 1)
        Else
            GetFilePath = strFullPath
        End If
    End Function

    '=====================================================================
    ' 方法：获取文件名（不带路径）
    '=====================================================================
    Public Function GetFileName(ByVal strFullPath As String) As String
        Dim nPos As Long
        
        nPos = InStrRev(strFullPath, "\")
        If nPos > 0 Then
            GetFileName = Mid$(strFullPath, nPos + 1)
        Else
            GetFileName = strFullPath
        End If
    End Function

    '=====================================================================
    ' 方法：获取文件扩展名（包括点号）
    '=====================================================================
    Public Function GetFileExtension(ByVal strFullPath As String) As String
        Dim nPos As Long
        
        nPos = InStrRev(strFullPath, ".")
        If nPos > 0 Then
            GetFileExtension = Mid$(strFullPath, nPos)
        Else
            GetFileExtension = ""
        End If
    End Function

    '=====================================================================
    ' 方法：获取文件名（不带扩展名）
    '=====================================================================
    Public Function GetFileNameWithoutExt(ByVal strFullPath As String) As String
        Dim strFileName As String
        Dim nPos As Long
        
        strFileName = Me.GetFileName(strFullPath)
        nPos = InStrRev(strFileName, ".")
        If nPos > 0 Then
            GetFileNameWithoutExt = Left$(strFileName, nPos - 1)
        Else
            GetFileNameWithoutExt = strFileName
        End If
    End Function

    '=====================================================================
    ' Method: Set Common Filters
    '=====================================================================
    Public Sub SetCommonFilters()
        Me.ClearFilter
        Me.AddFilter "Text Files (*.txt)", "*.txt"
        Me.AddFilter "All Files (*.*)", "*.*"
    End Sub

    Public Sub SetImageFilters()
        Me.ClearFilter
        Me.AddFilter "Bitmap Files (*.bmp)", "*.bmp"
        Me.AddFilter "JPEG Files (*.jpg;*.jpeg)", "*.jpg;*.jpeg"
        Me.AddFilter "PNG Files (*.png)", "*.png"
        Me.AddFilter "GIF Files (*.gif)", "*.gif"
        Me.AddFilter "All Image Files (*.bmp;*.jpg;*.jpeg;*.png;*.gif)", "*.bmp;*.jpg;*.jpeg;*.png;*.gif"
        Me.AddFilter "All Files (*.*)", "*.*"
    End Sub

    Public Sub SetDocumentFilters()
        Me.ClearFilter
        Me.AddFilter "Word Documents (*.doc;*.docx)", "*.doc;*.docx"
        Me.AddFilter "Excel Workbooks (*.xls;*.xlsx)", "*.xls;*.xlsx"
        Me.AddFilter "PowerPoint Presentations (*.ppt;*.pptx)", "*.ppt;*.pptx"
        Me.AddFilter "PDF Documents (*.pdf)", "*.pdf"
        Me.AddFilter "Text Files (*.txt)", "*.txt"
        Me.AddFilter "All Files (*.*)", "*.*"
    End Sub

    Public Sub SetCodeFilters()
        Me.ClearFilter
        Me.AddFilter "Basic Files (*.bas;*.vba)", "*.bas;*.vba"
        Me.AddFilter "C/C++ Files (*.c;*.cpp;*.h;*.hpp)", "*.c;*.cpp;*.h;*.hpp"
        Me.AddFilter "C# Files (*.cs)", "*.cs"
        Me.AddFilter "Java Files (*.java)", "*.java"
        Me.AddFilter "Python Files (*.py)", "*.py"
        Me.AddFilter "JavaScript Files (*.js)", "*.js"
        Me.AddFilter "HTML Files (*.html;*.htm)", "*.html;*.htm"
        Me.AddFilter "XML Files (*.xml)", "*.xml"
        Me.AddFilter "All Files (*.*)", "*.*"
    End Sub

'=====================================================================
' 方法：重置为默认设置
'=====================================================================
Public Sub Reset()
    With mData
        .mstrDialogTitle = ""
        .mstrInitialDir = ""
        .mstrDefaultExt = ""
        .mstrFileName = ""
        .mstrFilter = ""
        .mblnMultiSelect = False
        .mblnOverwritePrompt = True
        .mblnPathMustExist = True
        .mblnFileMustExist = True
        .mblnHideReadOnly = True
        .mblnReadOnly = False
        .mblnNoChangeDir = False
        .mblnCreatePrompt = False
        .mblnNoReadOnlyReturn = False
        .mblnLongNames = True
        .mblnExplorer = True
        .mblnEnableSizing = True
        .mblnNewDialogStyle = True
    End With
End Sub
