VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EnumRoundingMode
    Rounding = 0
    Upper = 1
    Downer = 2
End Enum


'截断取整
Public Function Ceil(ByVal Num As Variant, Optional ByVal Dot As Long) As Currency
    If Dot <> 0 Then Num = Num * (10 ^ Dot)
    Ceil = 0 - Int(0 - Num)
    If Dot <> 0 Then Ceil = Ceil / (10 ^ Dot)
End Function

Public Function GetRandRange(a As Long, b As Long) As Long
    '包含ab
    Randomize
    GetRandRange = Int((b - a + 1) * Rnd() + a)
End Function

'四舍五入，Currency类型的范围：-922,337,203,685,477.5808 到 922,337,203,685,477.5807
Public Function RoundCurrency(ByVal Value As Variant, Optional ByVal DecimalPlaces As Integer = 0, Optional ByVal RoundingMode As EnumRoundingMode = 0) As Currency
    Dim Factor As Currency
    Dim TempValue As Currency
    
    ' 处理特殊情况：值为0
    If Value = 0 Then
        RoundCurrency = 0
        Exit Function
    End If
    ' 处理超出Currency范围的情况
    If DecimalPlaces > 4 Then DecimalPlaces = 4
    If DecimalPlaces < -4 Then DecimalPlaces = -4
    ' 计算放大因子
    Factor = 10 ^ DecimalPlaces
    ' 放大数值，进行四舍五入，然后缩小
    TempValue = Value * Factor
    ' 检查是否会溢出
    If Abs(Value) > (2 ^ 63 - 1) / Abs(Factor) Then
        ' 如果可能溢出，使用更安全的方法
'        RoundCurrencySafe = Value
'        Exit Function
    End If
    
    Select Case RoundingMode
        Case Rounding ' 传统四舍五入
            TempValue = Sgn(Value) * Fix(Abs(TempValue) + 0.5)
        Case Upper ' 向上取整
            TempValue = Sgn(Value) * Fix(Abs(TempValue) + 0.9999)
        Case Downer ' 向下取整
            TempValue = Sgn(Value) * Fix(Abs(TempValue))
        Case Else ' 默认传统四舍五入
            TempValue = Sgn(Value) * Fix(Abs(TempValue) + 0.5)
    End Select
    
    RoundCurrency = TempValue / Factor
End Function
