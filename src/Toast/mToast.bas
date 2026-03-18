Attribute VB_Name = "mToast"
Option Explicit

Public Enum EnumPos
    LeftTop = 10
    LeftCenter = 12
    LeftBottom = 14
    CenterTop = 20
    Center = 22
    CenterBottom = 24
    RightTop = 30
    RightCenter = 32
    RightBottom = 34
End Enum

Public Enum EnumTheme
    Light = 1
    Dark = 2
End Enum

Public Enum EnumState
    Info = 1
    Success = 2
    Warning = 3
    Danger = 4
End Enum

Public Function IsLeft(ByVal p As Long) As Boolean
    IsLeft = p < 20
End Function
Public Function IsCenter(ByVal p As Long) As Boolean
    IsCenter = p >= 20 And p < 30
End Function
Public Function IsRight(ByVal p As Long) As Boolean
    IsRight = p >= 30
End Function

