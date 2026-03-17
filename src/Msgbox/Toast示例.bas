Attribute VB_Name = "Toast示例"
Option Explicit

'===========================================================================
' cToast 使用示例
'===========================================================================

Public Sub Demo_BasicUsage()
    ' ========== 基础用法 ==========
    
    ' 1. 最简单的调用 - 使用便捷函数
    ToastSuccess "操作成功！"
    ToastDanger "发生错误！"
    ToastWarning "请注意！"
    ToastInfo "提示信息"
    
    ' 2. 链式调用 - 创建并显示
    Toast().Text("Hello World").Show
    
    ' 3. 指定位置显示
    Toast().Text("左上角提示").TopLeft.Show
    Toast().Text("顶部居中").TopCenter.Show
    Toast().Text("右上角提示").TopRight.Show
    Toast().Text("左侧居中").MiddleLeft.Show
    Toast().Text("屏幕中央").MiddleCenter.Show
    Toast().Text("右侧居中").MiddleRight.Show
    Toast().Text("左下角提示").BottomLeft.Show
    Toast().Text("底部居中").BottomCenter.Show
    Toast().Text("右下角提示").BottomRight.Show
    
    ' 4. 自定义坐标
    Toast().Text("自定义位置").Pos(100, 100).Show
End Sub

Public Sub Demo_StyleCustomization()
    ' ========== 样式定制 ==========
    
    ' 1. 使用预置样式
    Toast().Text("成功！").Style(ToastStyle_Success).Show
    Toast().Text("错误！").Style(ToastStyle_Error).Show
    Toast().Text("警告！").Style(ToastStyle_Warning).Show
    Toast().Text("信息！").Style(ToastStyle_Info).Show
    
    ' 2. 自定义颜色
    Toast().Text("紫色背景"). _
        BackColor(RGB(128, 0, 128)). _
        ForeColor(RGB(255, 255, 255)). _
        Show
    
    ' 3. 渐变背景
    Toast().Text("渐变效果"). _
        Gradient(RGB(255, 0, 0), RGB(0, 0, 255)). _
        Show
    
    ' 4. 带边框
    Toast().Text("带边框的提示"). _
        BackColor(RGB(50, 50, 50)). _
        Border(RGB(255, 255, 255), 2). _
        Show
End Sub

Public Sub Demo_AdvancedFeatures()
    ' ========== 高级功能 ==========
    
    ' 1. 自定义字体
    Toast().Text("微软雅黑 14号"). _
        Font("微软雅黑", 14, True). _
        Show
    
    ' 2. 圆角半径
    Toast().Text("大圆角"). _
        Radius(20). _
        Show
    
    ' 3. 阴影效果
    Toast().Text("带阴影"). _
        Shadow(True, RGB(0, 0, 0), 6, 6, 15). _
        Show
    
    ' 4. 无阴影
    Toast().Text("无阴影"). _
        Shadow(False). _
        Show
    
    ' 5. 图标（使用 Emoji）
    Toast().Text("带图标"). _
        Icon("success"). _
        Style(ToastStyle_Success). _
        Show
    
    ' 6. 内边距调整
    Toast().Text("大内边距"). _
        Padding(40). _
        Show
End Sub

Public Sub Demo_AnimationTiming()
    ' ========== 动画和时间 ==========
    
    ' 1. 自定义显示时长
    Toast().Text("显示5秒"). _
        Duration(5000). _
        Show
    
    ' 2. 慢速淡入淡出
    Toast().Text("慢速动画"). _
        FadeInTime(500). _
        FadeOutTime(500). _
        Show
    
    ' 3. 快速显示
    Toast().Text("快速动画"). _
        FadeInTime(50). _
        FadeOutTime(50). _
        Show
End Sub

Public Sub Demo_ChainedCalls()
    ' ========== 完整链式调用示例 ==========
    
    ' 最完整的链式调用
    Toast().Text("保存成功！"). _
        Style(ToastStyle_Success). _
        TopRight(30, 30). _
        Font("微软雅黑", 12, True). _
        Padding(25). _
        Radius(10). _
        Shadow(True, RGB(0, 0, 0), 4, 4, 10). _
        Icon("success", 28). _
        Duration(3000). _
        FadeInTime(200). _
        FadeOutTime(200). _
        Show
    
    ' 深色主题
    Toast().Text("夜间模式提示"). _
        BackColor(RGB(30, 30, 30)). _
        ForeColor(RGB(200, 200, 200)). _
        Border(RGB(80, 80, 80), 1). _
        BottomCenter(50). _
        Show
    
    ' 彩色渐变
    Toast().Text("彩虹渐变"). _
        Gradient(RGB(255, 0, 128), RGB(128, 0, 255)). _
        ForeColor(RGB(255, 255, 255)). _
        MiddleCenter. _
        Radius(15). _
        Shadow(True, RGB(128, 0, 128), 0, 4, 20). _
        Show
End Sub

Public Sub Demo_MultipleToasts()
    ' ========== 多个 Toast 依次显示 ==========
    
    Dim t As cToast
    
    ' 第一个
    Set t = Toast().Text("第一步完成").Style(ToastStyle_Info).TopLeft
    t.Show
    Sleep 1000
    
    ' 第二个
    Set t = Toast().Text("第二步完成").Style(ToastStyle_Warning).TopCenter
    t.Show
    Sleep 1000
    
    ' 第三个
    Set t = Toast().Text("全部完成！").Style(ToastStyle_Success).TopRight
    t.Show
End Sub

' 辅助函数 - 延时
Private Sub Sleep(ByVal Milliseconds As Long)
    Dim StartTime As Long
    StartTime = Timer * 1000
    Do While Timer * 1000 - StartTime < Milliseconds
        DoEvents
    Loop
End Sub
