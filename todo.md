### 2026-01-20

今天看到一个牛逼的函数，来自维托的：

```vb
Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property
```

### 2026-01-19

- 给所有字典对象改为大小写不敏感

### 2026-01-18

- 给cWinsock添加一个函数，用于绑定业务 userID，方便针对用户发送，模仿 webman
- 给cHttpServer增加 view 函数，用于渲染html

### 2025-06-22

- CGI模块

### 2025-01-03

- 系统服务模块，支持一键安装为服务，可获取运行状态， 可控制服务
- sqlite模块
- 动态路由参数支持
- webdav客户端
- ctx.request 默认成员
- ctx.request 未处理 utf8
