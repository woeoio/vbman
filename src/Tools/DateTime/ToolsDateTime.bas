Attribute VB_Name = "ToolsDateTime"
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
    ElseIf VarType(varDate) = vbString Then
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
    
    'Dim formattedDate As String
    
    'Rem  使用当前时间并按指定格式返回
    'formattedDate = FormatDateTimeCustom("yyyy-mm-dd hh:nn:ss")
    'MsgBox formattedDate
    '
    'Rem  使用给定的日期时间字符串
    'formattedDate = FormatDateTimeCustom("dd/mm/yyyy", "2024-08-23 14:30:00")
    'MsgBox formattedDate
    '
    'Rem  使用给定的日期类型
    'formattedDate = FormatDateTimeCustom("hh:nn:ss AM/PM", #8/23/2024 2:30:00 PM#)
    'MsgBox formattedDate
    
End Function
