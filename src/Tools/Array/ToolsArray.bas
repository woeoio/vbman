Attribute VB_Name = "ToolsArray"
Option Explicit

' 声明 CopyMemory API
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function SplitLongToBytes Lib "msvbvm60" Alias "#644" (ByVal lngNum As Long) As longByteType
Private Type longType
    v As Long
End Type
Private Type longByteType
    a1 As Byte
    a2 As Byte
    a3 As Byte
    a4 As Byte
End Type

Private Sub Command1_Click()
    
    Dim lngNum As longType
    Dim lngByte As longByteType
    
    lngNum.v = &H11223344
    
    LSet lngByte = lngNum
    
    Debug.Print Hex(lngByte.a1), Hex(lngByte.a2), Hex(lngByte.a3), Hex(lngByte.a4)
    
End Sub



Public Property Let Extend(Vars, Value As Variant)
    '未完待续
    Dim Min As Long: Min = LBound(Value)
    Dim Max As Long: Max = UBound(Value)
    Dim All As Long: All = UBound(Vars)
    Dim i As Long
    For i = Min To Max
        If All < i Then Exit For
        If IsMissing(Vars(i)) = False Then Vars(i) = Value(i)
        '        Debug.Print Arr(i)
    Next
End Property



Sub Test()
    Dim a As String, c As String
    DeArray Split("a/b/c", "/"), a, , c
    Debug.Print a
    Extend(Array(a, c)) = Split("fdsdsf/dsfsdf/fdsfsd", "/")
    Debug.Print c
End Sub

'解构数组到变量
'Sub test()
'    Dim a As String, c As String
'    DeArray Array("", "", 12), a
'    Debug.Print a
'    DeArray Split("a/b/c", "/"), a,,c
'    Debug.Print a
'End Sub
Public Sub DeArray(Arr, ParamArray OutVars())
    '未处理 下标 不为零 的情况
    Dim Min As Long: Min = LBound(Arr)
    Dim Max As Long: Max = UBound(Arr)
    Dim All As Long: All = UBound(OutVars)
    Dim i As Long
    For i = Min To Max
        If All < i Then Exit For
        If IsMissing(OutVars(i)) = False Then OutVars(i) = Arr(i)
        '        Debug.Print Arr(i)
    Next
End Sub

Public Function IsArrayEmpty(Arr) As Boolean
    On Error Resume Next                                                        ' 忽略错误
    IsArrayEmpty = (LBound(Arr) > UBound(Arr))                                  ' 如果数组为空，会引发错误，此时LBound(arr) > UBound(arr)的值为True
    If ERR.Number <> 0 Then
        ' 如果触发了错误，说明数组为空
        IsArrayEmpty = True
        ERR.Clear                                                               ' 清除错误
    End If
    On Error GoTo 0                                                             ' 恢复正常的错误处理
End Function


Public Function GetIndexByValue(Arr As Variant, Value As String) As Long
    Dim i As Long
    GetIndexByValue = -1                                                        ' 默认返回 -1，表示未找到
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) = Value Then
            GetIndexByValue = i                                                 ' 找到值，返回索引
            Exit Function
        End If
    Next i
End Function


'=======================vb数组切片函数, 来自 chatgpt========================
' 处理 Byte 数组的切片
Function SliceByteArray(ByRef Arr() As Byte, ByVal StartPos As Long, Optional ByVal EndPos As Long = -1) As Byte()
    Dim slicedArray() As Byte
    Dim length As Long
    
    If EndPos = -1 Then
        length = UBound(Arr) - StartPos + 1                                     ' 从 startPos 到数组末尾
    ElseIf StartPos < EndPos Then
        length = EndPos - StartPos + 1                                          ' 从 startPos 到 endPos
    Else
        ' 如果 startPos > endPos，返回一个空数组
        ReDim slicedArray(0)
        SliceByteArray = slicedArray
        Exit Function
    End If
    
    ' ReDim 保证目标数组有足够的空间
    ReDim slicedArray(length - 1)
    ' 使用 CopyMemory 将字节数组切片到新的数组
    CopyMemory slicedArray(0), Arr(StartPos), length
    SliceByteArray = slicedArray
End Function

'' 处理 String 类型数组的切片
'Function SliceString(ByVal Arr As String, ByVal StartPos As Long, Optional ByVal EndPos As Long = -1) As String
'    Dim slicedStr As String
'
'    If EndPos = -1 Then
'        slicedStr = Mid$(Arr, StartPos + 1)                                     ' 从 startPos 到字符串末尾
'    Else
'        slicedStr = Mid$(Arr, StartPos + 1, EndPos - StartPos)                  ' 从 startPos 到 endPos
'    End If
'
'    SliceString = slicedStr
'End Function

' 处理 Long 类型数组的切片
Function SliceLongArray(ByRef Arr() As Long, ByVal StartPos As Long, Optional ByVal EndPos As Long = -1) As Long()
    Dim slicedLong() As Long
    Dim length As Long
    
    If EndPos = -1 Then
        length = (UBound(Arr) - StartPos + 1) * 4                               ' 每个 Long 占 4 字节
    ElseIf StartPos <= EndPos Then
        length = (EndPos - StartPos + 1) * 4                                    ' 每个 Long 占 4 字节
    Else
        ' 如果 startPos > endPos，返回一个空数组
        ReDim slicedLong(0)
        SliceLongArray = slicedLong
        Exit Function
    End If
    
    ' ReDim 保证目标数组有足够的空间
    ReDim slicedLong((length / 4) - 1)
    ' 使用 CopyMemory 将 Long 数组切片到新的数组
    CopyMemory slicedLong(0), Arr(StartPos), length
    SliceLongArray = slicedLong
End Function

' 测试代码
Sub TestSliceArray()
    Dim ByteArray() As Byte
    Dim slicedByte() As Byte
    Dim slicedStr As String
    Dim longArray() As Long
    Dim slicedLong() As Long
    Dim i As Long
    
    ' 创建一个示例字节数组（Byte 类型）
    ByteArray = StrConv("Hello-World", vbFromUnicode)
    ' 使用切片函数提取 Byte 数组（从 0 到 5）
    slicedByte = SliceByteArray(ByteArray, 0, 5)
    Debug.Print "Byte Slice 0 to 5:"
    For i = LBound(slicedByte) To UBound(slicedByte)
        Debug.Print Chr(slicedByte(i))
    Next i
    
    ' 创建一个示例字符串（String 类型）
    Dim str As String
    str = "Hello World"
    ' 使用切片函数提取 String（从 0 到 5）
    slicedStr = SliceString(str, 0, 5)
    Debug.Print "String Slice 0 to 5: " & slicedStr
    
    ' 创建一个示例 Long 数组（Long 类型）
    longArray = Array(10, 20, 30, 40, 50, 60, 70)
    ' 使用切片函数提取 Long 数组（从 2 到 5）
    slicedLong = SliceLongArray(longArray, 2, 5)
    Debug.Print "Long Slice 2 to 5:"
    For i = LBound(slicedLong) To UBound(slicedLong)
        Debug.Print slicedLong(i)
    Next i
End Sub

' 测试代码
Public Sub TestSliceArray2()
    Dim ByteArray() As Byte
    Dim sliced() As Byte
    Dim i As Long
    
    ' 创建一个示例字节数组
    ByteArray = StrConv("Hello-World, This is a test!", vbFromUnicode)
    
    ' 从 0 到 n（比如 5）
    sliced = SliceByteArray(ByteArray, 0, 5)
    Debug.Print "Slice 0 to 5:"
    For i = LBound(sliced) To UBound(sliced)
        Debug.Print Chr(sliced(i))
    Next i
    
    ' 从 n 到 max（比如从 6 开始到结尾）
    sliced = SliceByteArray(ByteArray, 6, -1)
    Debug.Print "Slice 6 to max:"
    For i = LBound(sliced) To UBound(sliced)
        Debug.Print Chr(sliced(i))
    Next i
    
    ' 从 n 到 m（比如从 6 到 10）
    sliced = SliceByteArray(ByteArray, 6, 10)
    Debug.Print "Slice 6 to 10:"
    For i = LBound(sliced) To UBound(sliced)
        Debug.Print Chr(sliced(i))
    Next i
End Sub


'============================以下是通过关键词分割字节数组的函数=============

' 将字符串转为字节数组
Public Function StringToByteArray(ByVal St As String) As Byte()
    StringToByteArray = StrConv(St, vbFromUnicode)
End Function

Public Function GetArrayLength(Arr()) As Long
    GetArrayLength = UBound(Arr) - LBound(Arr) + 1
End Function

' 在字节数组中查找关键词字节数组的位置
Public Function FindByteArray(ByRef Arr() As Byte, ByRef Keyword() As Byte) As Long
    Dim i As Long
    Dim keywordLen As Long
    keywordLen = UBound(Keyword) - LBound(Keyword) + 1
    
    ' 在字节数组中查找关键词的起始位置
    For i = LBound(Arr) To UBound(Arr) - keywordLen + 1
        '        If Arr(i) = Keyword(0) Then
        Dim j As Long
        For j = LBound(Keyword) To UBound(Keyword)
            If Arr(i + j) <> Keyword(j) Then
                Exit For
            End If
        Next j
        If j > UBound(Keyword) Then
            ' 找到匹配的位置，返回索引
            FindByteArray = i
            Exit Function
        End If
        '    End If
    Next i
    
    ' 如果没有找到，返回 -1
    FindByteArray = -1
End Function

' 将字节数组分割成两个数组
Public Function SplitByteArray(ByRef Arr() As Byte, ByVal position As Long, Optional Offset As Long) As Variant
    Dim part1() As Byte
    Dim part2() As Byte
    
    Dim ArrLen As Long
    ArrLen = UBound(Arr) - LBound(Arr) + 1
    
    If ArrLen <= position Then
        part1 = Arr
        ReDim part2(0)
        GoTo RET
    End If
    ' 前半部分 (0 到 position-1)
    ReDim part1(position - 1)
    CopyMemory part1(0), Arr(0), position
    
    position = position + Offset
    
    If ArrLen <= position Then
        ReDim part2(0)
        GoTo RET
    End If
    ' 后半部分 (position 到末尾)
    ReDim part2(UBound(Arr) - position)
    CopyMemory part2(0), Arr(position), UBound(Arr) - position + 1
RET:
    ' 返回一个包含两个数组的 Variant
    Dim result(1) As Variant
    result(0) = part1
    result(1) = part2
    SplitByteArray = result
End Function

' 主函数：根据关键词分割字节数组，支持多次分割 , 关键词支持字节数组和字符串
Function SplitByteArrayByKeyword(ByRef ByteArray() As Byte, ByVal Keyword As Variant, Optional ByVal SplitCount As Long) As Variant
    Dim keywordBytes() As Byte
    Dim splitResult() As Variant
    Dim currentArray() As Byte
    Dim keywordPos As Long
    Dim i As Long
    Dim keywordLen As Long
    
    ' 将关键词转换为字节数组
    keywordBytes = StringToByteArray(Keyword)
    keywordLen = UBound(keywordBytes) - LBound(keywordBytes) + 1
    ' 初始化当前字节数组为输入字节数组
    currentArray = ByteArray
    
    ' 存储分割结果
    Dim resultList() As Variant
    Dim resultIndex As Long
    resultIndex = 0
    
    ' 循环分割指定次数
    Do
        ' 在当前字节数组中查找关键词位置
        keywordPos = FindByteArray(currentArray, keywordBytes)
        
        If keywordPos = -1 Then
            ' 如果找不到关键词，停止分割，剩下的数组作为最后一部分
            Exit Do
        End If
        
        ' 将字节数组分割为两部分
        splitResult = SplitByteArray(currentArray, keywordPos, keywordLen)
        
        ' 将前半部分加入结果列表
        ReDim Preserve resultList(resultIndex)
        resultList(resultIndex) = splitResult(0)
        resultIndex = resultIndex + 1
        
        ' 更新当前字节数组为后半部分（从关键词后开始）
        currentArray = splitResult(1)
        
        i = i + 1
        
        '判断次数
        If SplitCount > 0 Then
            If i = SplitCount Then Exit Do
        End If
        
    Loop
    
    ' 将剩余部分加入到结果列表
    ReDim Preserve resultList(resultIndex)
    resultList(resultIndex) = currentArray
    
    ' 返回分割后的所有部分
    SplitByteArrayByKeyword = resultList
End Function

' 测试代码
Sub TestSplitArrayByKeyword()
    Dim ByteArray() As Byte
    Dim Keyword As String
    Dim result As Variant
    Dim i As Long, j As Long
    Dim part() As Byte
    
    ' 创建一个示例字节数组
    ByteArray = StrConv("Hello World, This is a test. Hello World again!", vbFromUnicode)
    
    ' 设置关键词
    Keyword = "World"
    
    ' 调用分割函数，分割两次
    result = SplitByteArrayByKeyword(ByteArray, Keyword)
    
    ' 打印分割后的结果
    For i = LBound(result) To UBound(result)
        Debug.Print "Part " & (i + 1) & ":"
        part = result(i)
        Debug.Print StrConv(part, vbUnicode)
        '        For j = LBound(part) To UBound(part)
        '            Debug.Print Chr(part(j))
        '        Next j
    Next i
End Sub
