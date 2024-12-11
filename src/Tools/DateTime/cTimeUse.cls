VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimeUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'todo 胎神说GetTickCount的精度太差，最好还是用 Query 系列来测速。
'用 NtQueryPerformanceCounter 更方便，这一个函数相当于同时调用了 QueryPerformanceCounter 和 QueryPerformanceFrequency

'https://learn.microsoft.com/zh-cn/windows/win32/devnotes/ntqueryperformancecounter

'Private Declare Function NtQueryPerformanceCounter Lib "ntdll" (ByRef PerformanceCounter As Currency, ByRef PerformanceFrequency As Currency) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim startTime As Long

Public Function Show(Optional Txt As String = "耗时：{*} 毫秒") As Variant
    Show = Replace(Txt, "{*}", GetTickCount - startTime)
End Function

Public Sub Start()
    ' 记录开始时间
    startTime = GetTickCount
End Sub

