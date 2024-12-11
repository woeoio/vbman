Attribute VB_Name = "ToolsMath"
Option Explicit

Public Function Ceil(ByVal Num as Variant, Optional ByVal Dot As Long) As Double
    If Dot <> 0 Then Num = Num * (10 ^ Dot)
    Ceil = 0 - Int(0 - Num)
    If Dot <> 0 Then Ceil = Ceil / (10 ^ Dot)
End Function

Public Function GetRandRange(a As Long, b As Long) As Long
    '°üº¬ab
    Randomize
    GetRandRange = Int((b - a + 1) * Rnd() + a)
End Function

