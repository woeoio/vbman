把这里的 tlb 利用起来
C:\Pro\VB6Mini3\lib\TypeLib

关于Public Type结构体作为参数使用的总结

1，在工程内部时，type可以放到模块中，多个类对象的friend函数可以使用
2，编译为dll的时候，type必须放到类内部，多个public函数可以使用

```vb
    ' 安全地将 ObjPtr 还原为对象引用
    ' 调用方必须确保 ptr 指向的对象仍然存活
    Public Function ObjectFromObjPtr(ByVal ptr As LongPtr) As Object
        If ptr = 0 Then Exit Function
        Dim obj As Object
        CopyMemory obj, ptr, LenB(ptr)
        Set ObjectFromObjPtr = obj
        ' 不要 Set obj = Nothing —— 会导致引用计数减1
    End Function
```
