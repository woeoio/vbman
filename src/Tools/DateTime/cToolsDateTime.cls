VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' VB6 获取当前时间戳 (13 位)
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Function GetUnixTimestamp() As Currency
    Dim sysTime As SYSTEMTIME
    Dim timestamp As Currency
    Dim dt As Date
    
    ' 获取系统时间
    GetSystemTime sysTime
    
    ' 将 SYSTEMTIME 转换为 VB6 日期类型
    dt = DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay) + _
    TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
    
    ' 获取 Unix 时间戳（秒）
    timestamp = (dt - #1/1/1970#) * 86400
    
    ' 返回 13 位时间戳（即秒数乘以1000）
    GetUnixTimestamp = timestamp * 1000# + sysTime.wMilliseconds
End Function

' 生成ISO8601格式时间戳 (UTC)
Public Function GetIso8601Timestamp() As String
    ' 格式: 2023-03-05T12:00:00Z (UTC时间)
    Dim st As SYSTEMTIME
    Call GetSystemTime(st)
    
    GetIso8601Timestamp = Format$(st.wYear, "0000") & "-" & _
                          Format$(st.wMonth, "00") & "-" & _
                          Format$(st.wDay, "00") & "T" & _
                          Format$(st.wHour, "00") & ":" & _
                          Format$(st.wMinute, "00") & ":" & _
                          Format$(st.wSecond, "00") & "Z"
End Function

' 生成UTC时间戳（ISO8601格式）
Public Function GetUtcTimestamp() As String
    Dim dt As Date
    dt = DateAdd("h", -8, Now) ' 北京时间转UTC
    GetUtcTimestamp = Format$(dt, "yyyy-mm-dd\Thh:mm:ss\Z")
End Function

Public Function IsDatePast(targetDate As Variant) As Boolean
    Dim parsedDate As Date
    Dim currentDate As Date
    If IsEmpty(targetDate) = True Then Exit Function
    If targetDate = "" Then Exit Function
    ' 尝试将字符串转换为日期
    On Error GoTo ErrorHandler
    parsedDate = CDate(targetDate)
    
    ' 获取当前日期
    currentDate = Date
    
    ' 比较当前日期与目标日期
    If DateDiff("d", parsedDate, currentDate) > 0 Then
        IsDatePast = True                                                       ' 如果当前日期大于目标日期，返回 True
    Else
        IsDatePast = False                                                      ' 如果当前日期小于或等于目标日期，返回 False
    End If
    Exit Function
    
ErrorHandler:
    ' 如果转换失败，返回错误信息
    IsDatePast = "Invalid Date Format"
End Function

Public Function FormatDateTime(Optional ByVal strFormat As String = "yyyy-mm-dd hh:nn:ss", Optional ByVal varDate As Variant) As String
    Dim dt As Date
    
    ' 如果未提供第二个参数，使用当前时间
    If IsMissing(varDate) Then
        dt = Now
    ElseIf IsDate(varDate) Then
        dt = CDate(varDate)
    ElseIf varType(varDate) = vbString Then
        dt = CDate(varDate)
    Else
        ' 如果第二个参数既不是日期也不是字符串，则返回空字符串
        FormatDateTime = ""
        Exit Function
    End If
    
    ' 使用提供的格式格式化日期
    On Error Resume Next
    FormatDateTime = Format(dt, strFormat)
    On Error GoTo 0
End Function

'=========================================================================
' HTTP 日期与缓存相关
'=========================================================================

' 格式化 HTTP 日期（RFC 1123 / RFC 2616 格式）
' 输出格式: "Day, DD Mon YYYY HH:MM:SS GMT"（如 "Mon, 22 Jun 2026 09:42:00 GMT"）
' 使用硬编码英文星期/月份名，确保 locale 无关
' 用途: Last-Modified、Cookie Expires、Cache-Control Expires 等 HTTP 头
Public Function FormatHttpDate(d As Date) As String
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

' 计算 ETag 标识（基于文件修改时间 + 大小，locale 无关）
' 输出格式: "YYYYMMDDHHNNSS-FileSize"（如 "20260622094200-1048576"）
' 用途: HTTP 协商缓存 ETag 头，比对文件是否更新
' 修复：fileSize 参数从 Long 改为 Variant，防止 fileInfo.Size > 2GB 时隐式转 Long 溢出崩溃
' Scripting.File.Size 返回 Variant，超过 Long 范围时会导致运行时错误 6: Overflow
Public Function ComputeETag(lastModified As Date, fileSize As Variant) As String
    ComputeETag = """" & _
        Right(Year(lastModified), 4) & _
        Right("0" & Month(lastModified), 2) & _
        Right("0" & Day(lastModified), 2) & _
        Right("0" & Hour(lastModified), 2) & _
        Right("0" & Minute(lastModified), 2) & _
        Right("0" & Second(lastModified), 2) & _
        "-" & CStr(fileSize) & """"
End Function