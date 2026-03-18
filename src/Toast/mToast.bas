Attribute VB_Name = "mToast"
Option Explicit



Public Function IsLeft(ByVal p As Long) As Boolean
    IsLeft = p < 20
End Function
Public Function IsCenter(ByVal p As Long) As Boolean
    IsCenter = p >= 20 And p < 30
End Function
Public Function IsRight(ByVal p As Long) As Boolean
    IsRight = p >= 30
End Function

