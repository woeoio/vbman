Attribute VB_Name = "ToolsLogs"
Option Explicit

Public Const LEVEL_DEBUG As String = "DEBUGGER"
Public Const LEVEL_INFO As String = "INFO"
Public Const LEVEL_WARN As String = "WARN"
Public Const LEVEL_DANGER As String = "DANGER"
Public Const LEVEL_ERROR As String = "ERRORS"
Public Const LEVEL_CUSTOM As String = "CUSTOM"
    
'Public Enum EnumLevel
'    Debugger = 0
'    Info = 1
'    Warn = 2
'    Danger = 3
'    Errors = 4
'    CUSTOM = 9
'End Enum
Public LogLevel As EnumLogLevel
Dim EnumLevelNames As Variant, IsEnumLevelNamesInit As Boolean

Public Property Get EnumLevelName(Level As EnumLogLevel) As String
    If IsEnumLevelNamesInit = False Then
        EnumLevelNames = Array(LEVEL_DEBUG, LEVEL_INFO, LEVEL_WARN, LEVEL_DANGER, LEVEL_ERROR, "", "", "", "", LEVEL_CUSTOM)
    End If
    If Level <= LvCustom Then
        EnumLevelName = EnumLevelNames(Level)
    Else
        EnumLevelName = "CUSTOM+" & (Level - EnumLogLevel.LvCustom)
    End If
End Property
