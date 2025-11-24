VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSTDIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' API 常量声明
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const STARTF_USESTDHANDLES = &H100
Private Const STARTF_USESHOWWINDOW = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2
Private Const INFINITE = &HFFFF
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_TIMEOUT = &H102
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_QUERY_INFORMATION = &H400
' 声明 timeGetTime 函数
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
' 结构体声明
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

' API 函数声明
Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

    'Private Declare Function CreateProcessA Lib "kernel32" ( _
    '    ByVal lpApplicationName As Long, _
    '    ByVal lpCommandLine As String, _
    '    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    '    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    '    ByVal bInheritHandles As Long, _
    '    ByVal dwCreationFlags As Long, _
    '    ByVal lpEnvironment As Long, _
    '    ByVal lpCurrentDirectory As Long, _
    '    lpStartupInfo As STARTUPINFO, _
    '    lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Declare Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As Long, _
    ByVal lpBuffer As Long, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function DuplicateHandle Lib "kernel32" ( _
    ByVal hSourceProcessHandle As Long, _
    ByVal hSourceHandle As Long, _
    ByVal hTargetProcessHandle As Long, _
    lpTargetHandle As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwOptions As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

' 错误枚举
Public Enum ProcessExecutorError
    ERROR_OK = 0
    ERROR_CREATE_PIPE = 1
    ERROR_DUPLICATE_HANDLE = 2
    ERROR_CREATE_PROCESS = 3
    ERROR_READ_OUTPUT = 4
    ERROR_PROCESS_TIMEOUT = 5
End Enum

' 私有变量
Private m_hChildStdInRd As Long
Private m_hChildStdInWr As Long
Private m_hChildStdOutRd As Long
Private m_hChildStdOutWr As Long
Private m_hProcess As Long
Private m_hThread As Long
Private m_dwProcessId As Long
Private m_bProcessRunning As Boolean

' 清理资源
Private Sub CleanUp()
    If m_hChildStdInRd <> 0 Then CloseHandle m_hChildStdInRd
    If m_hChildStdInWr <> 0 Then CloseHandle m_hChildStdInWr
    If m_hChildStdOutRd <> 0 Then CloseHandle m_hChildStdOutRd
    If m_hChildStdOutWr <> 0 Then CloseHandle m_hChildStdOutWr
    If m_hThread <> 0 Then CloseHandle m_hThread
    If m_hProcess <> 0 Then CloseHandle m_hProcess
    
    m_hChildStdInRd = 0
    m_hChildStdInWr = 0
    m_hChildStdOutRd = 0
    m_hChildStdOutWr = 0
    m_hThread = 0
    m_hProcess = 0
    m_dwProcessId = 0
    m_bProcessRunning = False
End Sub

' 类初始化
Private Sub Class_Initialize()
    CleanUp
End Sub

' 类终止
Private Sub Class_Terminate()
    CleanUp
End Sub

' 执行程序并获取输出
Public Function ExecuteCommand( _
    ByVal CommandLine As String, _
    Optional ByVal TimeoutMs As Long = 10000, _
    Optional ByVal CurrentDirectory As String = "", _
    Optional ByRef exitCode As Long) As String
    
    On Error GoTo ErrorHandler
    
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim hOutputRead As Long
    Dim hOutputWrite As Long
    Dim hInputRead As Long
    Dim hInputWrite As Long
    Dim hProcess As Long
    Dim dwRead As Long
    Dim dwAvail As Long
    Dim dwExitCode As Long
    Dim bSuccess As Boolean
    Dim sOutput As String
    Dim sBuffer As String * 4096
    Dim StartTime As Long
    Dim CurrentTime As Long
    
    ' 初始化安全属性
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    sa.lpSecurityDescriptor = 0
    
    ' 创建输出管道
    If CreatePipe(hOutputRead, hOutputWrite, sa, 0) = 0 Then
        ExecuteCommand = "Error creating output pipe"
        exitCode = ProcessExecutorError.ERROR_CREATE_PIPE
        Exit Function
    End If
    
    ' 创建输入管道
    If CreatePipe(hInputRead, hInputWrite, sa, 0) = 0 Then
        CloseHandle hOutputRead
        CloseHandle hOutputWrite
        ExecuteCommand = "Error creating input pipe"
        exitCode = ProcessExecutorError.ERROR_CREATE_PIPE
        Exit Function
    End If
    
    ' 复制输出读取句柄，使其不可继承
    Dim hOutputReadTmp As Long
    bSuccess = DuplicateHandle(GetCurrentProcess(), hOutputRead, GetCurrentProcess(), _
    hOutputReadTmp, 0, False, DUPLICATE_SAME_ACCESS)
    
    If bSuccess = 0 Then
        CloseHandle hOutputRead
        CloseHandle hOutputWrite
        CloseHandle hInputRead
        CloseHandle hInputWrite
        ExecuteCommand = "Error duplicating output handle"
        exitCode = ProcessExecutorError.ERROR_DUPLICATE_HANDLE
        Exit Function
    End If
    
    ' 关闭原始输出读取句柄
    CloseHandle hOutputRead
    hOutputRead = hOutputReadTmp
    
    ' 设置启动信息
    si.cb = Len(si)
    si.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    si.wShowWindow = 0                                                          ' 隐藏窗口
    si.hStdOutput = hOutputWrite
    si.hStdError = hOutputWrite
    si.hStdInput = hInputRead
    
    ' 创建进程
    bSuccess = CreateProcessA(0, CommandLine, sa, sa, True, NORMAL_PRIORITY_CLASS, 0, CurrentDirectory, si, pi)
    
    If bSuccess = 0 Then
        CloseHandle hOutputRead
        CloseHandle hOutputWrite
        CloseHandle hInputRead
        CloseHandle hInputWrite
        ExecuteCommand = "Error creating process @ " & App.Path & " # " & CurrentDirectory & " | " & CommandLine
        exitCode = ProcessExecutorError.ERROR_CREATE_PROCESS
        Exit Function
    End If
    
    ' 关闭不需要的句柄
    CloseHandle hOutputWrite
    CloseHandle hInputRead
    CloseHandle pi.hThread
    
    ' 记录进程信息
    m_hProcess = pi.hProcess
    m_hThread = pi.hThread
    m_dwProcessId = pi.dwProcessId
    m_bProcessRunning = True
    
    ' 记录开始时间
    StartTime = GetTickCount()
    
    ' 读取输出
    Do
        ' 检查进程是否已经退出
        GetExitCodeProcess m_hProcess, dwExitCode
        If dwExitCode <> 259 Then                                               ' STILL_ACTIVE = 259
            m_bProcessRunning = False
            Exit Do
        End If
        
        ' 检查是否有数据可读
        If PeekNamedPipe(hOutputRead, 0, 0, 0, dwAvail, 0) = 0 Then
            Exit Do
        End If
        
        If dwAvail > 0 Then
            ' 读取数据
            sBuffer = Space$(4096)
            If ReadFile(hOutputRead, sBuffer, Len(sBuffer), dwRead, 0&) = 0 Then
                Exit Do
            End If
            
            If dwRead > 0 Then
                sOutput = sOutput & Left$(sBuffer, dwRead)
            End If
        Else
            ' 检查超时
            CurrentTime = GetTickCount()
            If CurrentTime - StartTime > TimeoutMs Then
                exitCode = ProcessExecutorError.ERROR_PROCESS_TIMEOUT
                ERR.Raise "执行超时：" & TimeoutMs
                Exit Do
            End If
            
            ' 短暂等待
            WaitForSingleObject m_hProcess, 100
        End If
        
        DoEvents                                                                ' 允许处理其他事件
    Loop
    
    ' 确保读取所有剩余输出
    Do While True
        If PeekNamedPipe(hOutputRead, 0, 0, 0, dwAvail, 0) = 0 Then Exit Do
        If dwAvail = 0 Then Exit Do
        
        sBuffer = Space$(4096)
        If ReadFile(hOutputRead, sBuffer, Len(sBuffer), dwRead, 0&) = 0 Then Exit Do
        If dwRead = 0 Then Exit Do
        
        sOutput = sOutput & Left$(sBuffer, dwRead)
    Loop
    
    ' 获取最终退出代码
    GetExitCodeProcess m_hProcess, exitCode
    
    ' 返回输出
    ExecuteCommand = sOutput
    
    ' 清理资源
    CleanUp
    Exit Function
    
ErrorHandler:
    ExecuteCommand = "STDIO ERROR: " & ERR.Description
    exitCode = ERR.Number
    CleanUp
End Function

' 获取当前时间戳（毫秒）
Private Function GetTickCount() As Long
    GetTickCount = timeGetTime()
End Function






