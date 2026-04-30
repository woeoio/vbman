VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Base64编码表
Private Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

' Base64编码函数
Public Function Encode(ByVal Data As String) As String
    Dim i As Long
    Dim j As Long
    Dim ByteArray() As Byte
    Dim Result As String
    Dim Temp As Long
    Dim GroupCount As Long
    Dim Padding As Long
    
    ' 将字符串转换为字节数组
    ByteArray = StrConv(Data, vbFromUnicode)
    
    ' 计算需要多少个4字节组
    GroupCount = (UBound(ByteArray) + 1 + 2) \ 3                                ' 向上取整
    
    ' 初始化结果字符串
    Result = String(GroupCount * 4, "A")                                        ' 先用A占位
    
    ' 处理每个3字节组
    For i = 0 To GroupCount - 1
        ' 获取3个字节(不足的补0)
        Temp = 0
        For j = 0 To 2
            If i * 3 + j <= UBound(ByteArray) Then
                Temp = Temp Or (ByteArray(i * 3 + j) * (256 ^ (2 - j)))
            End If
        Next j
        
        ' 转换为4个Base64字符
        For j = 0 To 3
            If i * 4 + j < Len(Result) Then
                If i * 3 + j * 3 \ 4 <= UBound(ByteArray) Then
                    Mid$(Result, i * 4 + j + 1, 1) = Mid$(BASE64_TABLE, ((Temp \ (64 ^ (3 - j))) And 63) + 1, 1)
                Else
                    Mid$(Result, i * 4 + j + 1, 1) = "="                        ' 填充=
                End If
            End If
        Next j
    Next i
    
    Encode = Result
End Function

' Base64解码函数(保持不变)
Public Function Decode(ByVal Data As String) As String
    Dim i As Long
    Dim j As Long
    Dim ByteArray() As Byte
    Dim Result As String
    Dim Temp As Long
    Dim GroupCount As Long
    Dim CharValue As Long
    
    ' 去除无效字符
    Data = Replace(Data, vbCr, "")
    Data = Replace(Data, vbLf, "")
    Data = Replace(Data, vbTab, "")
    Data = Replace(Data, " ", "")
    
    ' 计算需要多少个3字节组
    GroupCount = Len(Data) \ 4
    
    ' 初始化字节数组
    ReDim ByteArray(GroupCount * 3 - 1)
    
    ' 处理每个4字符组
    For i = 0 To GroupCount - 1
        Temp = 0
        
        ' 获取4个Base64字符的值
        For j = 0 To 3
            CharValue = InStr(1, BASE64_TABLE, Mid$(Data, i * 4 + j + 1, 1)) - 1
            If CharValue < 0 Then CharValue = 0                                 ' 处理填充字符'='
            Temp = Temp Or (CharValue * (64 ^ (3 - j)))
        Next j
        
        ' 转换为3个字节
        For j = 0 To 2
            If i * 3 + j <= UBound(ByteArray) Then
                ByteArray(i * 3 + j) = (Temp \ (256 ^ (2 - j))) And 255
            End If
        Next j
    Next i
    
    ' 将字节数组转换为字符串
    Result = StrConv(ByteArray, vbUnicode)
    
    ' 去除可能的空字符
    If InStr(Result, vbNullChar) > 0 Then
        Result = Left$(Result, InStr(Result, vbNullChar) - 1)
    End If
    
    Decode = Result
End Function

' ==============================================
' UTF-8编码的Base64函数
' ==============================================

' UTF-8 Base64编码函数
Public Function Utf8Encode(ByVal Data As String) As String
    Dim utf8Bytes() As Byte
    utf8Bytes = StrToUtf8Bytes(Data)
    Utf8Encode = BytesToBase64(utf8Bytes)
End Function

' UTF-8 Base64解码函数
Public Function Utf8Decode(ByVal Data As String) As String
    Dim utf8Bytes() As Byte
    utf8Bytes = Base64ToBytes(Data)
    Utf8Decode = Utf8BytesToStr(utf8Bytes)
End Function

' ==============================================
' 辅助函数
' ==============================================

' 将字符串转换为UTF-8字节数组
Private Function StrToUtf8Bytes(ByVal UCS As String) As Byte()
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    Dim Encode As String
    lLength = Len(UCS)
    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, abUTF8(0), lBufferSize, vbNullString, 0)
    If lResult <> 0 Then
        lResult = lResult - 1
        ReDim Preserve abUTF8(lResult)
        StrToUtf8Bytes = abUTF8
    End If
End Function

' 将UTF-8字节数组转换为字符串
Private Function Utf8BytesToStr(ByRef Utf() As Byte) As String
    Dim lret As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim Decode As String
    On Error GoTo errline:
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    Decode = String$(lBufferSize, Chr(0))
    lret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(Decode), lBufferSize)
    If lret <> 0 Then
        Decode = Left(Decode, lret)
    End If
    Utf8BytesToStr = Decode
    Exit Function
errline:
    Utf8BytesToStr = ""
End Function

' 字节数组转Base64
Public Function BytesToBase64(ByRef Bytes() As Byte) As String
    Dim i As Long, j As Long
    Dim Result As String
    Dim Temp As Long
    Dim GroupCount As Long
    
    ' 计算需要多少个4字节组
    GroupCount = (UBound(Bytes) + 1 + 2) \ 3                                    ' 向上取整
    
    ' 初始化结果字符串
    Result = String(GroupCount * 4, "A")                                        ' 先用A占位
    
    ' 处理每个3字节组
    For i = 0 To GroupCount - 1
        ' 获取3个字节(不足的补0)
        Temp = 0
        For j = 0 To 2
            If i * 3 + j <= UBound(Bytes) Then
                Temp = Temp Or (Bytes(i * 3 + j) * (256 ^ (2 - j)))
            End If
        Next j
        
        ' 转换为4个Base64字符
        For j = 0 To 3
            If i * 4 + j < Len(Result) Then
                If i * 3 + j * 3 \ 4 <= UBound(Bytes) Then
                    Mid$(Result, i * 4 + j + 1, 1) = Mid$(BASE64_TABLE, ((Temp \ (64 ^ (3 - j))) And 63) + 1, 1)
                Else
                    Mid$(Result, i * 4 + j + 1, 1) = "="                        ' 填充=
                End If
            End If
        Next j
    Next i
    
    BytesToBase64 = Result
End Function

' Base64转字节数组
Public Function Base64ToBytes(ByVal Data As String) As Byte()
    Dim i As Long, j As Long
    Dim Bytes() As Byte
    Dim Temp As Long
    Dim GroupCount As Long
    Dim CharValue As Long
    Dim PaddingCount As Long
    
    ' 去除无效字符
    Data = Replace(Data, vbCr, "")
    Data = Replace(Data, vbLf, "")
    Data = Replace(Data, vbTab, "")
    Data = Replace(Data, " ", "")
    
    ' 计算需要多少个3字节组
    GroupCount = Len(Data) \ 4
    
    ' 计算填充字符数量（末尾的=号）
    If Len(Data) > 0 Then
        If Right$(Data, 1) = "=" Then PaddingCount = 1
        If Len(Data) > 1 Then
            If Mid$(Data, Len(Data) - 1, 1) = "=" Then PaddingCount = 2
        End If
    End If
    
    ' 初始化字节数组（根据填充字符调整实际大小）
    ReDim Bytes(GroupCount * 3 - PaddingCount - 1)
    
    ' 处理每个4字符组
    For i = 0 To GroupCount - 1
        Temp = 0
        
        ' 获取4个Base64字符的值
        For j = 0 To 3
            CharValue = InStr(1, BASE64_TABLE, Mid$(Data, i * 4 + j + 1, 1)) - 1
            If CharValue < 0 Then CharValue = 0                                 ' 处理填充字符'='
            Temp = Temp Or (CharValue * (64 ^ (3 - j)))
        Next j
        
        ' 转换为3个字节（最后一个组根据填充数量减少输出）
        Dim BytesToWrite As Long
        If i = GroupCount - 1 Then
            BytesToWrite = 3 - PaddingCount
        Else
            BytesToWrite = 3
        End If
        
        For j = 0 To BytesToWrite - 1
            If i * 3 + j <= UBound(Bytes) Then
                Bytes(i * 3 + j) = (Temp \ (256 ^ (2 - j))) And 255
            End If
        Next j
    Next i
    
    Base64ToBytes = Bytes
End Function

Public Function FileToBase64(ByVal Path As String, Optional ImageType As String) As String
    Dim d() As Byte
    ToolsStream.LoadFileAsBinary Path, d
    FileToBase64 = BytesToBase64(d)
    If ImageType <> "" Then
        FileToBase64 = "data:image/" & ImageType & ";base64," & FileToBase64
    End If
End Function


Public Function Base64ToFile(ByRef Data As String, Optional ByVal Path As String) As Byte()
    Dim Pos As Long: Pos = ToolsStr.HasStr(";base64,", Data, 11)
    Dim Dat As String
    If Pos = 0 Then
        Dat = Data
    Else
        Dat = Mid$(Data, Pos + 8)
    End If
    Base64ToFile = Base64ToBytes(Dat)
    If Path <> "" Then
        ToolsStream.SaveFileAsBinary Path, Base64ToFile
    End If
End Function
