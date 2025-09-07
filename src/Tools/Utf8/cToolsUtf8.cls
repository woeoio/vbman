VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsUtf8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'* ************************************** *

'* 模块名称：modCharset.bas

'* 模块功能：GB2312与UTF8相互转换函数

'* 作者：lyserver

'* ************************************** *
 
Option Explicit


Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

Public Function Decode(ByRef Utf() As Byte) As String
    Dim lret As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    On Error GoTo errline:
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    Decode = String$(lBufferSize, Chr(0))
    lret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(Decode), lBufferSize)
    If lret <> 0 Then
        Decode = Left(Decode, lret)
    End If
    Exit Function
errline:
    Decode = ""
End Function

Public Function DecodeToByteArray(ByRef Utf() As Byte) As Byte()
    Dim lret As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim BT() As Byte
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2 - 1
    ReDim BT(lBufferSize)
    lret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, VarPtr(BT(0)), lBufferSize + 1)
    If lret <> 0 Then
        ReDim Preserve BT(lret - 1)
        DecodeToByteArray = BT
    End If
End Function


Public Function Encode(ByVal UCS As String) As Byte()
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    lLength = Len(UCS)
    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, abUTF8(0), lBufferSize, vbNullString, 0)
    If lResult <> 0 Then
        lResult = lResult - 1
        ReDim Preserve abUTF8(lResult)
        Encode = abUTF8
    End If
End Function


'Dim adoStream As New ADODB.Stream

Rem 作者和来源不详，如有请告知
Rem 工程要引用 Microsoft ActiveX Data Objects 2.8，下面的通用方法建议放在模块中
'Public Sub SaveTo(ByVal FileName As String, ByVal Text As String)
'    With adoStream
'        .Open
'        .Charset = "UTF-8"
'        .Type = adTypeText
'        .WriteText Text
'        .SaveToFile FileName, adSaveCreateOverWrite
'        .Close
'    End With
'End Sub

'Public Function LoadFrom(ByVal FileName As String) As String
'    With adoStream
'        .Open
'        .Charset = "UTF-8"
'        .LoadFromFile FileName
'        LoadFrom = .ReadText()
'        .Close
'    End With
'End Function

'- ------------------------------------------- -

'  函数说明：GB2312转换为UTF8

'- ------------------------------------------- -

Public Function EncodeWithBom(strIn As String) As Byte()
    Dim adoStream  As New ADODB.Stream
    '    Static adoStream As Object
    Set adoStream = CreateObject("ADODB.Stream")
    With adoStream
        .CharSet = "utf-8"
        .Type = 2                                                               'adTypeText
        .Open
        .WriteText strIn
        .position = 0
        .Type = 1                                                               'adTypeBinary
        Rem  这个办法可以的     239 187 191 = utf8bom
        '        Dim A() As Byte, b() As Byte, C() As Byte
        '        A = .Read(1): .Position = 1
        '        b = .Read(1): .Position = 2
        '        C = .Read(1): .Position = 3
        '        If A(0) = 239 And b(0) = 187 And C(0) = 191 Then
        '            '如果不带bom就从第四个字节开始读取
        '        Else
        '            .Position = 0
        '        End If
        EncodeWithBom = .Read()
        .Close
    End With
End Function

'- ------------------------------------------- -

'  函数说明：UTF8转换为GB2312

'- ------------------------------------------- -

Public Function DecodeWithBom(ByVal varIn As Variant) As String
    'Dim adoStream  As New ADODB.Stream
    Static adoStream As Object
    Set adoStream = CreateObject("ADODB.Stream")
    Dim bytesData() As Byte
    bytesData = varIn
    With adoStream
        .CharSet = "utf-8"
        .Type = 1                                                               'adTypeBinary
        .Open
        .Write bytesData
        .position = 0
        .Type = 2                                                               'adTypeText
        DecodeWithBom = .ReadText()
        .Close
    End With
End Function

