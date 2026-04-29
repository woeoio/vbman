VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ============================================================
' 高级文件操作类 (CreateFile API 实现)
'   - 支持大文件(>2GB)和后续扩展异步IO
'   - 一个实例同一时间只能操作一个文件
'   - 读取内容先进入内部缓冲区，再通过 ReturnXXX 提取
' ============================================================

' --- Win32 API 声明 ---
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, liDistanceToMove As Currency, lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

' --- 常量 ---
Private Const GENERIC_READ              As Long = &H80000000
Private Const GENERIC_WRITE             As Long = &H40000000
Private Const GENERIC_READWRITE         As Long = &HC0000000
Private Const FILE_SHARE_READ           As Long = &H1
Private Const FILE_SHARE_WRITE          As Long = &H2
Private Const OPEN_EXISTING             As Long = 3
Private Const OPEN_ALWAYS               As Long = 4
Private Const CREATE_ALWAYS             As Long = 2
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_FLAG_OVERLAPPED      As Long = &H40000000
Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const FILE_BEGIN                As Long = 0
Private Const FILE_CURRENT              As Long = 1
Private Const FILE_END                  As Long = 2
Private Const CP_UTF8                   As Long = 65001
Private Const CP_ACP                    As Long = 0

' --- 成员变量 ---
Private m_hFile         As Long         ' 文件句柄
Private m_Path          As String       ' 打开的文件路径
Private m_Buffer()      As Byte         ' 内部数据缓冲区
Private m_IsOpen        As Boolean      ' 是否已打开
Private m_LastError     As String       ' 最后一次错误信息

' ============================================================
' 属性
' ============================================================

Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property

Public Property Get FilePath() As String
    FilePath = m_Path
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

' 文件大小 (支持大文件，返回 Currency 类型，单位字节)
Public Property Get FileSize() As Currency
    If Not m_IsOpen Then ERR.Raise 53, "cFileIO.FileSize", "文件未打开"
    Dim sz As Currency
    If GetFileSizeEx(m_hFile, sz) = 0 Then
        ERR.Raise 5, "cFileIO.FileSize", "获取文件大小失败"
    End If
    FileSize = sz * 10000@  ' Currency 内部存储为 1/10000
End Property

' 当前缓冲区字节数
Public Property Get BufferSize() As Long
    BufferSize = UBound(m_Buffer) + 1
End Property

' ============================================================
' 文件打开/关闭
' ============================================================

' 打开文件，返回 Me 支持链式调用
' AccessMode: "R"=只读, "W"=只写, "RW"=读写(默认)
Public Function OpenFile(ByVal FilePath As String, Optional ByVal AccessMode As String = "RW") As cFileIO
    CloseFile
    
    Dim dwAccess As Long
    Select Case UCase$(Trim$(AccessMode))
        Case "R":  dwAccess = GENERIC_READ
        Case "W":  dwAccess = GENERIC_WRITE
        Case Else: dwAccess = GENERIC_READWRITE
    End Select
    
    m_hFile = CreateFileW(StrPtr(FilePath), dwAccess, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If m_hFile = INVALID_HANDLE_VALUE Then
        ERR.Raise 53, "cFileIO.OpenFile", "无法打开文件: " & FilePath
    End If
    
    m_Path = FilePath
    m_IsOpen = True
    m_LastError = ""
    Set OpenFile = Me
End Function

Public Sub CloseFile()
    If m_IsOpen And m_hFile <> INVALID_HANDLE_VALUE Then
        FlushFileBuffers m_hFile
        CloseHandle m_hFile
        m_hFile = INVALID_HANDLE_VALUE
        m_IsOpen = False
        m_Path = ""
    End If
End Sub

' ============================================================
' 缓冲区操作
' ============================================================

' 清空内部缓冲区
Public Function ClearBuffer() As cFileIO
    Erase m_Buffer
    Set ClearBuffer = Me
End Function

' 用外部字节数组填充缓冲区，返回 Me
Public Function SetBuffer(Data() As Byte) As cFileIO
    m_Buffer = Data
    Set SetBuffer = Me
End Function

' 追加字节到缓冲区
Public Function AppendBuffer(Data() As Byte) As cFileIO
    Dim oldLen As Long, newLen As Long, addLen As Long
    oldLen = UBound(m_Buffer) + 1
    addLen = UBound(Data) + 1
    If addLen <= 0 Then
        Set AppendBuffer = Me
        Exit Function
    End If
    newLen = oldLen + addLen
    ReDim Preserve m_Buffer(newLen - 1)
    CopyMemory m_Buffer(oldLen), Data(0), addLen
    Set AppendBuffer = Me
End Function

' 用字符串填充缓冲区（默认 UTF-8），返回 Me
Public Function SetBufferText(ByVal Text As String, Optional ByVal CharSet As String = "UTF-8") As cFileIO
    m_Buffer = StringToBytes(Text, CharSet)
    Set SetBufferText = Me
End Function

' ============================================================
' 数据读取（读入内部缓冲区）
' ============================================================

' 将文件指定范围读入内部缓冲区，返回 Me 支持链式调用。
'   StartPos / EndPos 为 0-based 字节偏移；传 -1 表示使用默认值。
'   StartPos = -1 : 从文件开头（第 0 字节）开始
'   EndPos   = -1 : 读到文件最后一个字节
Public Function ReadData(Optional ByVal StartPos As Long = -1, Optional ByVal EndPos As Long = -1) As cFileIO
    If Not m_IsOpen Then ERR.Raise 53, "cFileIO.ReadData", "文件未打开"
    
    Dim fSize As Currency
    If GetFileSizeEx(m_hFile, fSize) = 0 Then
        m_LastError = "获取文件大小失败"
        ERR.Raise 5, "cFileIO.ReadData", m_LastError
    End If
    Dim fLen As Currency
    fLen = fSize * 10000@
    If fLen <= 0 Then
        Erase m_Buffer
        Set ReadData = Me
        Exit Function
    End If
    
    Dim sPos As Currency, ePos As Currency
    sPos = IIf(StartPos < 0, 0@, CCur(StartPos))
    If sPos < 0@ Then sPos = 0@
    ePos = IIf(EndPos < 0, fLen - 1@, CCur(EndPos))
    If ePos >= fLen Then ePos = fLen - 1@
    If ePos < sPos Then ePos = sPos
    
    Dim rLen As Currency
    rLen = ePos - sPos + 1@
    If rLen > 2147483647@ Then
        m_LastError = "单次读取不能超过 2GB"
        ERR.Raise 6, "cFileIO.ReadData", m_LastError
    End If
    
    ' 定位到开始位置
    Dim posCur As Currency
    If SetFilePointerEx(m_hFile, CCur(sPos / 10000@), posCur, FILE_BEGIN) = 0 Then
        m_LastError = "文件定位失败"
        ERR.Raise 5, "cFileIO.ReadData", m_LastError
    End If
    
    ' 分块读取（每次最多 2MB）
    Const CHUNK As Long = 2097152
    Dim totalRead As Currency
    totalRead = 0@
    ReDim m_Buffer(0 To CInt(rLen) - 1)
    
    Do While totalRead < rLen
        Dim toRead As Long
        toRead = IIf(rLen - totalRead > CCur(CHUNK), CHUNK, CInt(rLen - totalRead))
        Dim readLen As Long
        If ReadFile(m_hFile, m_Buffer(CInt(totalRead)), toRead, readLen, 0) = 0 Then
            m_LastError = "读取文件失败"
            ERR.Raise 5, "cFileIO.ReadData", m_LastError
        End If
        If readLen = 0 Then Exit Do
        totalRead = totalRead + CCur(readLen)
    Loop
    
    If totalRead < rLen Then
        ReDim Preserve m_Buffer(0 To CInt(totalRead) - 1)
    End If
    
    m_LastError = ""
    Set ReadData = Me
End Function

' ============================================================
' 数据返回（从内部缓冲区提取）
' ============================================================

' 返回文本，CharSet 默认 UTF-8
Public Function ReturnText(Optional ByVal CharSet As String = "UTF-8") As String
    If UBound(m_Buffer) < 0 Then
        ReturnText = ""
        Exit Function
    End If
    ReturnText = BytesToString(m_Buffer, CharSet)
End Function

' 返回字节数组（副本）
Public Function ReturnBytes() As Byte()
    ReturnBytes = m_Buffer
End Function

' 返回 Base64 编码字符串
Public Function ReturnBase64() As String
    ReturnBase64 = ToolsBase64.BytesToBase64(m_Buffer)
End Function

' ============================================================
' 数据写入（从内部缓冲区写入文件）
' ============================================================

' 将内部缓冲区写入文件。
'   FilePath : 目标路径，未提供则使用 OpenFile 缓存的路径
'   IsAppend : False=覆盖(默认), True=追加到文件末尾
Public Function SaveData(Optional ByVal FilePath As String, Optional ByVal IsAppend As Boolean = False) As cFileIO
    Dim targetPath As String
    targetPath = IIf(Len(FilePath) > 0, FilePath, m_Path)
    If Len(targetPath) = 0 Then
        m_LastError = "未提供文件路径"
        ERR.Raise 5, "cFileIO.SaveData", m_LastError
    End If
    
    Dim hOut As Long
    Dim needClose As Boolean
    
    If targetPath = m_Path And m_IsOpen Then
        ' 写入当前打开的文件
        hOut = m_hFile
        needClose = False
    Else
        ' 打开目标文件
        hOut = CreateFileW(StrPtr(targetPath), GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        If hOut = INVALID_HANDLE_VALUE Then
            m_LastError = "无法创建/打开目标文件: " & targetPath
            ERR.Raise 53, "cFileIO.SaveData", m_LastError
        End If
        needClose = True
    End If
    
    ' 追加模式：定位到文件末尾
    If IsAppend Then
        Dim posEnd As Currency
        If SetFilePointerEx(hOut, 0@, posEnd, FILE_END) = 0 Then
            If needClose Then CloseHandle hOut
            m_LastError = "定位到文件末尾失败"
            ERR.Raise 5, "cFileIO.SaveData", m_LastError
        End If
    ElseIf needClose Then
        ' 新建文件，从头开始写
        Dim posZero As Currency
        SetFilePointerEx hOut, 0@, posZero, FILE_BEGIN
    End If
    
    ' 分块写入
    Dim bufLen As Long
    bufLen = UBound(m_Buffer) + 1
    If bufLen > 0 Then
        Const CHUNK As Long = 2097152
        Dim totalWritten As Long
        totalWritten = 0
        Do While totalWritten < bufLen
            Dim toWrite As Long
            toWrite = IIf(bufLen - totalWritten > CHUNK, CHUNK, bufLen - totalWritten)
            Dim written As Long
            If WriteFile(hOut, m_Buffer(totalWritten), toWrite, written, 0) = 0 Then
                If needClose Then CloseHandle hOut
                m_LastError = "写入文件失败"
                ERR.Raise 5, "cFileIO.SaveData", m_LastError
            End If
            totalWritten = totalWritten + written
        Loop
    End If
    
    FlushFileBuffers hOut
    If needClose Then CloseHandle hOut
    
    m_LastError = ""
    Set SaveData = Me
End Function

' ============================================================
' 辅助函数
' ============================================================

' 字节数组 -> 字符串（支持 UTF-8 / ANSI / UTF-16LE）
Private Function BytesToString(ByRef b() As Byte, ByVal CharSet As String) As String
    Dim lb As Long: lb = LBound(b)
    Dim cnt As Long: cnt = UBound(b) - lb + 1
    If cnt <= 0 Then Exit Function
    
    Select Case LCase$(Trim$(CharSet))
        Case "utf-8", "utf8"
            ' 检测并跳过 BOM
            Dim startIdx As Long: startIdx = lb
            If cnt >= 3 Then
                If b(lb) = &HEF And b(lb + 1) = &HBB And b(lb + 2) = &HBF Then
                    startIdx = lb + 3
                    cnt = cnt - 3
                End If
            End If
            If cnt > 0 Then
                Dim wlen As Long
                wlen = MultiByteToWideChar(CP_UTF8, 0, b(startIdx), cnt, 0, 0)
                If wlen > 0 Then
                    BytesToString = String$(wlen, vbNullChar)
                    MultiByteToWideChar CP_UTF8, 0, b(startIdx), cnt, StrPtr(BytesToString), wlen
                End If
            End If
            
        Case "utf-16", "utf-16le", "utf16", "unicode"
            ' 检测并跳过 BOM
            startIdx = lb
            If cnt >= 2 Then
                If b(lb) = &HFF And b(lb + 1) = &HFE Then
                    startIdx = lb + 2
                    cnt = cnt - 2
                End If
            End If
            If cnt > 0 Then
                BytesToString = b
            End If
            
        Case Else ' ANSI / 系统默认编码
            wlen = MultiByteToWideChar(CP_ACP, 0, b(lb), cnt, 0, 0)
            If wlen > 0 Then
                BytesToString = String$(wlen, vbNullChar)
                MultiByteToWideChar CP_ACP, 0, b(lb), cnt, StrPtr(BytesToString), wlen
            End If
    End Select
End Function

' 字符串 -> 字节数组（支持 UTF-8 / ANSI / UTF-16LE）
Private Function StringToBytes(ByVal s As String, ByVal CharSet As String) As Byte()
    Dim b() As Byte
    If Len(s) = 0 Then
        StringToBytes = b
        Exit Function
    End If
    
    Select Case LCase$(Trim$(CharSet))
        Case "utf-8", "utf8"
            Dim blen As Long
            blen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), ByVal vbNullString, 0, vbNullString, 0)
            If blen > 0 Then
                ReDim b(blen - 1)
                WideCharToMultiByte CP_UTF8, 0, StrPtr(s), Len(s), b(0), blen, vbNullString, 0
            End If
            
        Case "utf-16", "utf-16le", "utf16", "unicode"
            b = s
            
        Case Else ' ANSI
            blen = WideCharToMultiByte(CP_ACP, 0, StrPtr(s), Len(s), ByVal vbNullString, 0, vbNullString, 0)
            If blen > 0 Then
                ReDim b(blen - 1)
                WideCharToMultiByte CP_ACP, 0, StrPtr(s), Len(s), b(0), blen, vbNullString, 0
            End If
    End Select
    StringToBytes = b
End Function

Private Sub Class_Terminate()
    CloseFile
End Sub
