Attribute VB_Name = "ToolsDelay"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Function DelayMicroseconds(us As Double) As Long
    Dim freq As Currency
    Dim startCount As Currency
    Dim endCount As Currency
    
    ' 获取计时器的频率
    QueryPerformanceFrequency freq
    ' 获取开始时间
    QueryPerformanceCounter startCount
    
    ' 计算结束时间
    endCount = startCount + (freq * us / 1000000#)
    
    ' 循环等待直到达到结束时间
    Do
        QueryPerformanceCounter startCount
    Loop While startCount < endCount
    
    DelayMicroseconds = 0
End Function

' 用于线程的函数
Private Function ThreadDelay(lpParam As Long) As Long
    Sleep lpParam                                                               ' 延时，单位：毫秒
    ThreadDelay = 0
End Function

Sub AsyncDelayWithThread(ms As Long)
    Dim threadID As Long
    Dim threadHandle As Long
    
    ' 创建线程
    threadHandle = CreateThread(0, 0, AddressOf DelayMicroseconds, ByVal ms, 0, threadID)
    If threadHandle Then
        ' 等待线程结束
        WaitForSingleObject threadHandle, &HFFFFFFFF
        ' 关闭线程句柄
        CloseHandle threadHandle
    End If
End Sub
'
'Sub TestAsyncDelayWithThread()
'    Debug.Print "开始: "; Timer
'    AsyncDelayWithThread 500                                                    ' 延时 500 毫秒
'    Debug.Print "结束: "; Timer
'End Sub

