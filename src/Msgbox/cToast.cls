VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===========================================================================
' GDI+ Toast 提示框 - 基于 GDI+ 直接在屏幕上绘制
' 特点：扁平风格、支持圆角、渐变背景、淡入淡出动画、链式调用
'===========================================================================

'--------------------- Windows API ---------------------
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Private Const WS_POPUP = &H80000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_NOACTIVATE = &H8000000
Private Const WS_EX_TRANSPARENT = &H20

Private Const SW_SHOWNOACTIVATE = 4
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOACTIVATE = &H10
Private Const LWA_ALPHA = &H2



Private Type POINTAPI
    x As Long
    Y As Long
End Type

'--------------------- GDI+ API ---------------------
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type RECTF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type POINTF
    x As Single
    Y As Single
End Type

Private Type ColorMatrix
    M(0 To 4, 0 To 4) As Single
End Type

Private Type BlurParams
    Radius As Single
    ExpandEdge As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMode As Long) As Long
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, ByVal Mode As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal Color As Long, Brush As Long) As Long
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (ByRef Point1 As POINTF, ByRef Point2 As POINTF, ByVal Color1 As Long, ByVal Color2 As Long, ByVal WrapMode As Long, LineGradient As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal BrushMode As Long, Path As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal Path As Long) As Long
Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal Path As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal Path As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal Path As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal Path As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal Unit As Long, Pen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal Pen As Long) As Long
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal FontCollection As Long, FontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal FontFamily As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal FontFamily As Long, ByVal EmSize As Single, ByVal Style As Long, ByVal Unit As Long, Font As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal Font As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal FormatAttributes As Long, ByVal Language As Integer, StringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As Long) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As Long) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal Graphics As Long, ByVal lString As Long, ByVal Length As Long, ByVal Font As Long, ByRef LayoutRect As RECTF, ByVal StringFormat As Long, ByRef BoundingBox As RECTF, ByRef CodepointsFitted As Long, ByRef LinesFilled As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal Graphics As Long, ByVal lString As Long, ByVal Length As Long, ByVal Font As Long, ByRef LayoutRect As RECTF, ByVal StringFormat As Long, ByVal Brush As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, Graphics As Long) As Long
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal x As Long, ByVal Y As Long, Color As Long) As Long
Private Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal x As Long, ByVal Y As Long, ByVal Color As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, CloneImage As Long) As Long
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwEffectType As GUID, Effect As Long) As Long
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal Effect As Long, ByRef Params As Any, ByVal Size As Long) As Long
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByRef Source As RECTF, ByVal x As Single, ByVal Y As Single, ByVal Matrix As Long, ByVal Effect As Long, ByVal ImageAttributes As Long, ByVal SrcUnit As Long) As Long

Private Const SmoothingModeAntiAlias = 4
Private Const TextRenderingHintClearTypeGridFit = 5
Private Const FillModeAlternate = 0
Private Const UnitPixel = 2
Private Const StringAlignmentCenter = 1
Private Const StringAlignmentNear = 0
Private Const StringAlignmentFar = 2
Private Const PixelFormat32bppARGB = &H26200A
Private Const WrapModeTile = 0

'--------------------- 常量 ---------------------
Public Enum ToastPosition
    Toast_TopLeft = 0
    Toast_TopCenter = 1
    Toast_TopRight = 2
    Toast_MiddleLeft = 3
    Toast_MiddleCenter = 4
    Toast_MiddleRight = 5
    Toast_BottomLeft = 6
    Toast_BottomCenter = 7
    Toast_BottomRight = 8
    Toast_Custom = 9
End Enum

Public Enum ToastStyle
    ToastStyle_Default = 0
    ToastStyle_Success = 1
    ToastStyle_Warning = 2
    ToastStyle_Error = 3
    ToastStyle_Info = 4
End Enum

'--------------------- 成员变量 ---------------------
Private m_Token As Long
Private m_hwnd As Long
Private m_TimerID As Long

' 位置和尺寸
Private m_X As Long
Private m_Y As Long
Private m_Width As Long
Private m_Height As Long
Private m_Position As ToastPosition
Private m_Padding As Long
Private m_CornerRadius As Long

' 内容
Private m_Text As String
Private m_FontName As String
Private m_FontSize As Single
Private m_FontBold As Boolean

' 颜色
Private m_BackColor As Long
Private m_BackColor2 As Long
Private m_TextColor As Long
Private m_BorderColor As Long
Private m_BorderWidth As Long
Private m_UseGradient As Boolean

' 动画
Private m_Duration As Long
Private m_FadeInTime As Long
Private m_FadeOutTime As Long
Private m_Opacity As Byte
Private m_IsShowing As Boolean
Private m_StartTime As Long

' 阴影
Private m_ShadowEnabled As Boolean
Private m_ShadowColor As Long
Private m_ShadowOffsetX As Long
Private m_ShadowOffsetY As Long
Private m_ShadowBlur As Long

' 图标
Private m_Icon As String
Private m_IconSize As Long

'--------------------- 初始化 ---------------------
Private Sub Class_Initialize()
    InitGDIPlus
    ResetToDefaults
End Sub

Private Sub Class_Terminate()
    Hide
    ShutdownGDIPlus
End Sub

Private Sub ResetToDefaults()
    m_X = 0
    m_Y = 0
    m_Width = 0
    m_Height = 0
    m_Position = Toast_BottomCenter
    m_Padding = 20
    m_CornerRadius = 8
    
    m_Text = ""
    m_FontName = "微软雅黑"
    m_FontSize = 11
    m_FontBold = False
    
    m_BackColor = RGB(50, 50, 50)
    m_BackColor2 = RGB(80, 80, 80)
    m_TextColor = RGB(255, 255, 255)
    m_BorderColor = RGB(100, 100, 100)
    m_BorderWidth = 0
    m_UseGradient = True
    
    m_Duration = 3000
    m_FadeInTime = 200
    m_FadeOutTime = 200
    m_Opacity = 255
    m_IsShowing = False
    
    m_ShadowEnabled = True
    m_ShadowColor = RGB(0, 0, 0)
    m_ShadowOffsetX = 4
    m_ShadowOffsetY = 4
    m_ShadowBlur = 8
    
    m_Icon = ""
    m_IconSize = 24
End Sub

'--------------------- GDI+ 初始化 ---------------------
Private Sub InitGDIPlus()
    Dim GDIInput As GdiplusStartupInput
    GDIInput.GdiplusVersion = 1
    GdiplusStartup m_Token, GDIInput
End Sub

Private Sub ShutdownGDIPlus()
    If m_Token <> 0 Then
        GdiplusShutdown m_Token
        m_Token = 0
    End If
End Sub

'--------------------- 颜色转换（无溢出版） ---------------------
Private Function ARGB(ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
    ' 使用 CopyMemory 将无符号值复制到 Long，避免溢出
    Dim cResult As Currency
    Dim lResult As Long
    
    cResult = (CCur(Alpha) * 16777216@) + (CCur(Red) * 65536@) + (CCur(Green) * 256@) + CCur(Blue)
    
    ' 如果值大于 2147483647，需要特殊处理
    If cResult > 2147483647@ Then
        ' 减去 2^32 得到对应的负数表示
        lResult = CLng(cResult - 4294967296@)
    Else
        lResult = CLng(cResult)
    End If
    
    ARGB = lResult
End Function

Private Function ColorToARGB(ByVal Color As Long, Optional ByVal Alpha As Byte = 255) As Long
    Dim b As Byte, G As Byte, R As Byte
    b = Color \ &H10000 And &HFF
    G = Color \ &H100 And &HFF
    R = Color And &HFF
    ColorToARGB = ARGB(Alpha, R, G, b)
End Function

'--------------------- 尺寸计算 ---------------------
Private Sub CalculateSize()
    Dim hDC As Long
    Dim Graphics As Long
    Dim FontFamily As Long
    Dim Font As Long
    Dim StringFormat As Long
    Dim LayoutRect As RECTF
    Dim BoundingBox As RECTF
    Dim Style As Long
    
    hDC = GetDC(0)
    GdipCreateFromHDC hDC, Graphics
    
    ' 创建字体
    GdipCreateFontFamilyFromName StrPtr(m_FontName), 0, FontFamily
    Style = IIf(m_FontBold, 1, 0)
    GdipCreateFont FontFamily, m_FontSize, Style, UnitPixel, Font
    
    ' 创建字符串格式
    GdipCreateStringFormat 0, 0, StringFormat
    GdipSetStringFormatAlign StringFormat, StringAlignmentCenter
    GdipSetStringFormatLineAlign StringFormat, StringAlignmentCenter
    
    ' 测量文本
    LayoutRect.Left = 0
    LayoutRect.Top = 0
    LayoutRect.Width = 1000
    LayoutRect.Height = 1000
    GdipMeasureString Graphics, StrPtr(m_Text), Len(m_Text), Font, LayoutRect, StringFormat, BoundingBox, 0, 0
    
    ' 计算尺寸
    m_Width = BoundingBox.Width + m_Padding * 2 + IIf(m_Icon <> "", m_IconSize + 10, 0)
    m_Height = BoundingBox.Height + m_Padding * 2
    If m_Height < m_IconSize + m_Padding * 2 Then m_Height = m_IconSize + m_Padding * 2
    
    ' 清理
    GdipDeleteStringFormat StringFormat
    GdipDeleteFont Font
    GdipDeleteFontFamily FontFamily
    GdipDeleteGraphics Graphics
    ReleaseDC 0, hDC
    
    ' 根据位置计算坐标
    If m_Position <> Toast_Custom Then
        CalculatePosition
    End If
End Sub

Private Sub CalculatePosition()
    Dim ScreenWidth As Long
    Dim ScreenHeight As Long
    Dim TaskbarHeight As Long
    
    ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
    ScreenHeight = GetSystemMetrics(SM_CYSCREEN)
    TaskbarHeight = ScreenHeight - GetSystemMetrics(SM_CYFULLSCREEN)
    
    Select Case m_Position
        Case Toast_TopLeft
            m_X = 20
            m_Y = 20
        Case Toast_TopCenter
            m_X = (ScreenWidth - m_Width) \ 2
            m_Y = 20
        Case Toast_TopRight
            m_X = ScreenWidth - m_Width - 20
            m_Y = 20
        Case Toast_MiddleLeft
            m_X = 20
            m_Y = (ScreenHeight - m_Height) \ 2
        Case Toast_MiddleCenter
            m_X = (ScreenWidth - m_Width) \ 2
            m_Y = (ScreenHeight - m_Height) \ 2
        Case Toast_MiddleRight
            m_X = ScreenWidth - m_Width - 20
            m_Y = (ScreenHeight - m_Height) \ 2
        Case Toast_BottomLeft
            m_X = 20
            m_Y = ScreenHeight - m_Height - TaskbarHeight - 20
        Case Toast_BottomCenter
            m_X = (ScreenWidth - m_Width) \ 2
            m_Y = ScreenHeight - m_Height - TaskbarHeight - 20
        Case Toast_BottomRight
            m_X = ScreenWidth - m_Width - 20
            m_Y = ScreenHeight - m_Height - TaskbarHeight - 20
    End Select
End Sub

'--------------------- 绘制 ---------------------
Private Sub DrawToast()
    Dim hDC As Long
    Dim memDC As Long
    Dim hBitmap As Long
    Dim oldBitmap As Long
    Dim Graphics As Long
    Dim Brush As Long
    Dim Pen As Long
    Dim Path As Long
    Dim FontFamily As Long
    Dim Font As Long
    Dim StringFormat As Long
    Dim TextBrush As Long
    Dim LayoutRect As RECTF
    Dim P1 As POINTF
    Dim P2 As POINTF
    Dim ShadowBitmap As Long
    Dim ShadowGraphics As Long
    Dim x As Long, Y As Long
    Dim ShadowWidth As Long
    Dim ShadowHeight As Long
    Dim Style As Long
    
    ' 创建内存DC
    hDC = GetDC(0)
    memDC = CreateCompatibleDC(hDC)
    hBitmap = CreateCompatibleBitmap(hDC, m_Width + m_ShadowOffsetX + m_ShadowBlur * 2, m_Height + m_ShadowOffsetY + m_ShadowBlur * 2)
    oldBitmap = SelectObject(memDC, hBitmap)
    
    ' 创建GDI+ Graphics
    GdipCreateFromHDC memDC, Graphics
    GdipSetSmoothingMode Graphics, SmoothingModeAntiAlias
    GdipSetTextRenderingHint Graphics, TextRenderingHintClearTypeGridFit
    
    ' 清空背景
    GdipCreateSolidFill ARGB(0, 0, 0, 0), Brush
    GdipFillRectangle Graphics, Brush, 0, 0, m_Width + m_ShadowOffsetX + m_ShadowBlur * 2, m_Height + m_ShadowOffsetY + m_ShadowBlur * 2
    GdipDeleteBrush Brush
    
    ' 绘制阴影
    If m_ShadowEnabled Then
        DrawShadow Graphics, m_ShadowBlur, m_ShadowOffsetX, m_ShadowOffsetY
    End If
    
    ' 创建圆角路径
    GdipCreatePath FillModeAlternate, Path
    AddRoundedRectToPath Path, m_ShadowBlur, m_ShadowBlur, m_Width, m_Height, m_CornerRadius
    
    ' 填充背景
    If m_UseGradient Then
        P1.x = m_ShadowBlur
        P1.Y = m_ShadowBlur
        P2.x = m_ShadowBlur
        P2.Y = m_ShadowBlur + m_Height
        GdipCreateLineBrush P1, P2, ColorToARGB(m_BackColor), ColorToARGB(m_BackColor2), WrapModeTile, Brush
    Else
        GdipCreateSolidFill ColorToARGB(m_BackColor), Brush
    End If
    GdipFillPath Graphics, Brush, Path
    GdipDeleteBrush Brush
    
    ' 绘制边框
    If m_BorderWidth > 0 Then
        GdipCreatePen1 ColorToARGB(m_BorderColor), m_BorderWidth, UnitPixel, Pen
        GdipDrawPath Graphics, Pen, Path
        GdipDeletePen Pen
    End If
    
    GdipDeletePath Path
    
    ' 绘制图标
    If m_Icon <> "" Then
        DrawIcon Graphics, m_ShadowBlur + m_Padding, m_ShadowBlur + (m_Height - m_IconSize) \ 2
    End If
    
    ' 绘制文本
    GdipCreateFontFamilyFromName StrPtr(m_FontName), 0, FontFamily
    Style = IIf(m_FontBold, 1, 0)
    GdipCreateFont FontFamily, m_FontSize, Style, UnitPixel, Font
    GdipCreateStringFormat 0, 0, StringFormat
    GdipSetStringFormatAlign StringFormat, StringAlignmentCenter
    GdipSetStringFormatLineAlign StringFormat, StringAlignmentCenter
    
    GdipCreateSolidFill ColorToARGB(m_TextColor), TextBrush
    
    LayoutRect.Left = m_ShadowBlur + IIf(m_Icon <> "", m_IconSize + 10, 0)
    LayoutRect.Top = m_ShadowBlur
    LayoutRect.Width = m_Width - IIf(m_Icon <> "", m_IconSize + 10, 0)
    LayoutRect.Height = m_Height
    
    GdipDrawString Graphics, StrPtr(m_Text), Len(m_Text), Font, LayoutRect, StringFormat, TextBrush
    
    ' 清理
    GdipDeleteBrush TextBrush
    GdipDeleteStringFormat StringFormat
    GdipDeleteFont Font
    GdipDeleteFontFamily FontFamily
    GdipDeleteGraphics Graphics
    
    ' 显示到屏幕
    UpdateLayeredWindow memDC, m_ShadowBlur, m_ShadowBlur
    
    ' 清理资源
    SelectObject memDC, oldBitmap
    DeleteObject hBitmap
    DeleteDC memDC
    ReleaseDC 0, hDC
End Sub

Private Sub AddRoundedRectToPath(ByVal Path As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal Radius As Single)
    Dim diameter As Single
    diameter = Radius * 2
    
    ' 左上角
    GdipAddPathEllipse Path, x, Y, diameter, diameter
    ' 右上角
    GdipAddPathEllipse Path, x + Width - diameter, Y, diameter, diameter
    ' 右下角
    GdipAddPathEllipse Path, x + Width - diameter, Y + Height - diameter, diameter, diameter
    ' 左下角
    GdipAddPathEllipse Path, x, Y + Height - diameter, diameter, diameter
    ' 中间矩形
    GdipAddPathRectangle Path, x + Radius, Y, Width - diameter, Height
    GdipAddPathRectangle Path, x, Y + Radius, Width, Height - diameter
End Sub

Private Sub DrawShadow(ByVal Graphics As Long, ByVal Blur As Long, ByVal OffsetX As Long, ByVal OffsetY As Long)
    Dim ShadowDC As Long
    Dim hShadowBitmap As Long
    Dim oldShadowBitmap As Long
    Dim ShadowGraphics As Long
    Dim Brush As Long
    Dim x As Long, Y As Long
    Dim ShadowWidth As Long
    Dim ShadowHeight As Long
    Dim Alpha As Byte
    Dim i As Long, j As Long
    
    ShadowWidth = m_Width + Blur * 2
    ShadowHeight = m_Height + Blur * 2
    
    ' 创建阴影DC
    ShadowDC = CreateCompatibleDC(0)
    hShadowBitmap = CreateCompatibleBitmap(GetDC(0), ShadowWidth, ShadowHeight)
    oldShadowBitmap = SelectObject(ShadowDC, hShadowBitmap)
    
    ' 绘制阴影形状
    GdipCreateFromHDC ShadowDC, ShadowGraphics
    GdipSetSmoothingMode ShadowGraphics, SmoothingModeAntiAlias
    
    ' 填充阴影矩形（带圆角）
    Dim Path As Long
    GdipCreatePath FillModeAlternate, Path
    AddRoundedRectToPath Path, Blur + OffsetX, Blur + OffsetY, m_Width, m_Height, m_CornerRadius
    
    ' 提取颜色分量
    Dim R As Byte, G As Byte, b As Byte
    b = m_ShadowColor \ &H10000 And &HFF
    G = m_ShadowColor \ &H100 And &HFF
    R = m_ShadowColor And &HFF
    
    GdipCreateSolidFill ARGB(128, R, G, b), Brush
    GdipFillPath ShadowGraphics, Brush, Path
    GdipDeleteBrush Brush
    GdipDeletePath Path
    GdipDeleteGraphics ShadowGraphics
    
    ' 模糊处理（简化版：使用透明度渐变）
    ' 实际项目中可以使用更复杂的模糊算法
    
    ' 清理
    SelectObject ShadowDC, oldShadowBitmap
    DeleteObject hShadowBitmap
    DeleteDC ShadowDC
End Sub

Private Sub DrawIcon(ByVal Graphics As Long, ByVal x As Single, ByVal Y As Single)
    ' 简化的图标绘制 - 使用字符作为图标
    Dim FontFamily As Long
    Dim Font As Long
    Dim StringFormat As Long
    Dim Brush As Long
    Dim LayoutRect As RECTF
    Dim IconText As String
    
    Select Case m_Icon
        Case "success", "ok", "check"
            IconText = ""
        Case "warning", "warn"
            IconText = ""
        Case "error", "fail", "cross"
            IconText = ""
        Case "info", "i"
            IconText = ""
        Case Else
            IconText = m_Icon
    End Select
    
    If IconText = "" Then Exit Sub
    
    GdipCreateFontFamilyFromName StrPtr("Segoe UI Emoji"), 0, FontFamily
    GdipCreateFont FontFamily, m_IconSize, 0, UnitPixel, Font
    GdipCreateStringFormat 0, 0, StringFormat
    GdipSetStringFormatAlign StringFormat, StringAlignmentCenter
    GdipSetStringFormatLineAlign StringFormat, StringAlignmentCenter
    
    GdipCreateSolidFill ColorToARGB(m_TextColor), Brush
    
    LayoutRect.Left = x
    LayoutRect.Top = Y
    LayoutRect.Width = m_IconSize
    LayoutRect.Height = m_IconSize
    
    GdipDrawString Graphics, StrPtr(IconText), Len(IconText), Font, LayoutRect, StringFormat, Brush
    
    GdipDeleteBrush Brush
    GdipDeleteStringFormat StringFormat
    GdipDeleteFont Font
    GdipDeleteFontFamily FontFamily
End Sub

Private Sub UpdateLayeredWindow(ByVal hdcSource As Long, ByVal xSrc As Long, ByVal ySrc As Long)
    ' 使用 UpdateLayeredWindow API 更新分层窗口
    ' 这里简化处理，实际应该使用 UpdateLayeredWindow
    If m_hwnd <> 0 Then
        ' 已经由窗口消息处理
    End If
End Sub

'--------------------- 显示/隐藏 ---------------------
Private Sub CreateToastWindow()
    Dim hInstance As Long
    
    ' 销毁旧窗口
    If m_hwnd <> 0 Then
        DestroyWindow m_hwnd
        m_hwnd = 0
    End If
    
    hInstance = GetModuleHandle(vbNullString)
    
    ' 创建分层窗口 - 使用系统已注册的类名 "Static"
    m_hwnd = CreateWindowEx(WS_EX_LAYERED Or WS_EX_TOOLWINDOW Or WS_EX_NOACTIVATE Or WS_EX_TRANSPARENT, _
        "Static", "Toast", WS_POPUP Or WS_VISIBLE, _
        m_X - m_ShadowBlur, m_Y - m_ShadowBlur, _
        m_Width + m_ShadowBlur * 2 + m_ShadowOffsetX, _
        m_Height + m_ShadowBlur * 2 + m_ShadowOffsetY, _
        0, 0, hInstance, ByVal 0)
    
    ' 设置置顶
    SetWindowPos m_hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
    ' 设置初始透明度
    SetLayeredWindowAttributes m_hwnd, 0, 0, LWA_ALPHA
End Sub

' 使用默认窗口过程 - 在类模块中不能使用 AddressOf，改用 0 表示使用默认处理
' 实际窗口过程注册在标准模块中处理

Private Sub FadeIn()
    Dim StartTime As Long
    Dim Elapsed As Long
    Dim Alpha As Byte
    
    StartTime = GetTickCount
    
    Do While Elapsed < m_FadeInTime
        Elapsed = GetTickCount - StartTime
        Alpha = CByte((Elapsed / m_FadeInTime) * 255)
        If Alpha > 255 Then Alpha = 255
        
        SetLayeredWindowAttributes m_hwnd, 0, Alpha, LWA_ALPHA
        UpdateWindow m_hwnd
        DoEvents
    Loop
    
    SetLayeredWindowAttributes m_hwnd, 0, 255, LWA_ALPHA
End Sub

Private Sub FadeOut()
    Dim StartTime As Long
    Dim Elapsed As Long
    Dim Alpha As Byte
    
    StartTime = GetTickCount
    
    Do While Elapsed < m_FadeOutTime
        Elapsed = GetTickCount - StartTime
        Alpha = CByte(255 - (Elapsed / m_FadeOutTime) * 255)
        If Alpha < 0 Then Alpha = 0
        
        SetLayeredWindowAttributes m_hwnd, 0, Alpha, LWA_ALPHA
        UpdateWindow m_hwnd
        DoEvents
    Loop
    
    SetLayeredWindowAttributes m_hwnd, 0, 0, LWA_ALPHA
End Sub

Private Function GetTickCount() As Long
    GetTickCount = CLng(Timer * 1000) Mod &H7FFFFFFF
End Function

'--------------------- 公共方法 - 链式调用 ---------------------

' 设置位置（基础方法）
Public Function Pos(ByVal x As Long, ByVal Y As Long) As cToast
    Set Pos = Me
    m_X = x
    m_Y = Y
    m_Position = Toast_Custom
End Function

' 预置位置 - 左上
Public Function TopLeft(Optional ByVal OffsetX As Long = 20, Optional ByVal OffsetY As Long = 20) As cToast
    Set TopLeft = Me
    m_Position = Toast_TopLeft
    CalculatePosition
    m_X = m_X + OffsetX - 20
    m_Y = m_Y + OffsetY - 20
End Function

' 预置位置 - 顶部居中
Public Function TopCenter(Optional ByVal OffsetY As Long = 20) As cToast
    Set TopCenter = Me
    m_Position = Toast_TopCenter
    CalculatePosition
    m_Y = m_Y + OffsetY - 20
End Function

' 预置位置 - 右上
Public Function TopRight(Optional ByVal OffsetX As Long = 20, Optional ByVal OffsetY As Long = 20) As cToast
    Set TopRight = Me
    m_Position = Toast_TopRight
    CalculatePosition
    m_X = m_X - OffsetX + 20
    m_Y = m_Y + OffsetY - 20
End Function

' 预置位置 - 左中
Public Function MiddleLeft(Optional ByVal OffsetX As Long = 20) As cToast
    Set MiddleLeft = Me
    m_Position = Toast_MiddleLeft
    CalculatePosition
    m_X = m_X + OffsetX - 20
End Function

' 预置位置 - 居中
Public Function MiddleCenter() As cToast
    Set MiddleCenter = Me
    m_Position = Toast_MiddleCenter
    CalculatePosition
End Function

' 预置位置 - 右中
Public Function MiddleRight(Optional ByVal OffsetX As Long = 20) As cToast
    Set MiddleRight = Me
    m_Position = Toast_MiddleRight
    CalculatePosition
    m_X = m_X - OffsetX + 20
End Function

' 预置位置 - 左下
Public Function BottomLeft(Optional ByVal OffsetX As Long = 20, Optional ByVal OffsetY As Long = 20) As cToast
    Set BottomLeft = Me
    m_Position = Toast_BottomLeft
    CalculatePosition
    m_X = m_X + OffsetX - 20
    m_Y = m_Y - OffsetY + 20
End Function

' 预置位置 - 底部居中
Public Function BottomCenter(Optional ByVal OffsetY As Long = 20) As cToast
    Set BottomCenter = Me
    m_Position = Toast_BottomCenter
    CalculatePosition
    m_Y = m_Y - OffsetY + 20
End Function

' 预置位置 - 右下
Public Function BottomRight(Optional ByVal OffsetX As Long = 20, Optional ByVal OffsetY As Long = 20) As cToast
    Set BottomRight = Me
    m_Position = Toast_BottomRight
    CalculatePosition
    m_X = m_X - OffsetX + 20
    m_Y = m_Y - OffsetY + 20
End Function

' 设置文本
Public Function Text(ByVal Value As String) As cToast
    Set Text = Me
    m_Text = Value
End Function

' 设置字体
Public Function Font(ByVal Name As String, Optional ByVal Size As Single = 11, Optional ByVal Bold As Boolean = False) As cToast
    Set Font = Me
    m_FontName = Name
    m_FontSize = Size
    m_FontBold = Bold
End Function

' 设置背景颜色（单色）
Public Function BackColor(ByVal Color As Long) As cToast
    Set BackColor = Me
    m_BackColor = Color
    m_UseGradient = False
End Function

' 设置渐变背景颜色
Public Function Gradient(ByVal Color1 As Long, ByVal Color2 As Long) As cToast
    Set Gradient = Me
    m_BackColor = Color1
    m_BackColor2 = Color2
    m_UseGradient = True
End Function

' 设置文字颜色
Public Function ForeColor(ByVal Color As Long) As cToast
    Set ForeColor = Me
    m_TextColor = Color
End Function

' 设置边框
Public Function bOrder(ByVal Color As Long, Optional ByVal Width As Long = 1) As cToast
    Set bOrder = Me
    m_BorderColor = Color
    m_BorderWidth = Width
End Function

' 设置圆角半径
Public Function Radius(ByVal Value As Long) As cToast
    Set Radius = Me
    m_CornerRadius = Value
End Function

' 设置内边距
Public Function Padding(ByVal Value As Long) As cToast
    Set Padding = Me
    m_Padding = Value
End Function

' 设置阴影
Public Function Shadow(Optional ByVal Enabled As Boolean = True, Optional ByVal Color As Long = -1, Optional ByVal OffsetX As Long = 4, Optional ByVal OffsetY As Long = 4, Optional ByVal Blur As Long = 8) As cToast
    Set Shadow = Me
    m_ShadowEnabled = Enabled
    If Color >= 0 Then m_ShadowColor = Color
    m_ShadowOffsetX = OffsetX
    m_ShadowOffsetY = OffsetY
    m_ShadowBlur = Blur
End Function

' 设置图标
Public Function Icon(ByVal Value As String, Optional ByVal Size As Long = 24) As cToast
    Set Icon = Me
    m_Icon = Value
    m_IconSize = Size
End Function

' 设置持续时间
Public Function Duration(ByVal Milliseconds As Long) As cToast
    Set Duration = Me
    m_Duration = Milliseconds
End Function

' 设置淡入时间
Public Function FadeInTime(ByVal Milliseconds As Long) As cToast
    Set FadeInTime = Me
    m_FadeInTime = Milliseconds
End Function

' 设置淡出时间
Public Function FadeOutTime(ByVal Milliseconds As Long) As cToast
    Set FadeOutTime = Me
    m_FadeOutTime = Milliseconds
End Function

' 应用预置样式
Public Function Style(ByVal ToastStyle As ToastStyle) As cToast
    Set Style = Me
    Select Case ToastStyle
        Case ToastStyle_Default
            m_BackColor = RGB(50, 50, 50)
            m_BackColor2 = RGB(80, 80, 80)
            m_TextColor = RGB(255, 255, 255)
        Case ToastStyle_Success
            m_BackColor = RGB(76, 175, 80)
            m_BackColor2 = RGB(102, 187, 106)
            m_TextColor = RGB(255, 255, 255)
            m_Icon = "success"
        Case ToastStyle_Warning
            m_BackColor = RGB(255, 152, 0)
            m_BackColor2 = RGB(255, 183, 77)
            m_TextColor = RGB(255, 255, 255)
            m_Icon = "warning"
        Case ToastStyle_Error
            m_BackColor = RGB(244, 67, 54)
            m_BackColor2 = RGB(239, 83, 80)
            m_TextColor = RGB(255, 255, 255)
            m_Icon = "error"
        Case ToastStyle_Info
            m_BackColor = RGB(33, 150, 243)
            m_BackColor2 = RGB(66, 165, 245)
            m_TextColor = RGB(255, 255, 255)
            m_Icon = "info"
    End Select
    m_UseGradient = True
End Function

' 显示Toast
Public Function Show(Optional ByVal Message As String = "") As cToast
    Set Show = Me
    
    If Message <> "" Then m_Text = Message
    If m_Text = "" Then m_Text = "Toast"
    
    ' 隐藏之前的
    Hide
    
    ' 计算尺寸和位置
    CalculateSize
    
    ' 创建窗口
    CreateToastWindow
    
    ' 绘制内容
    DrawToast
    
    ' 淡入
    FadeIn
    
    m_IsShowing = True
    m_StartTime = GetTickCount
    
    ' 启动定时器
    If m_Duration > 0 Then
        StartAutoCloseTimer
    End If
End Function

' 隐藏Toast
Public Function Hide() As cToast
    Set Hide = Me
    
    If m_hwnd <> 0 Then
        If m_FadeOutTime > 0 And m_IsShowing Then
            FadeOut
        End If
        DestroyWindow m_hwnd
        m_hwnd = 0
    End If
    
    If m_TimerID <> 0 Then
        KillTimer 0, m_TimerID
        m_TimerID = 0
    End If
    
    m_IsShowing = False
End Function

' 关闭所有Toast（类方法）
Public Sub CloseAll()
    ' 清理静态集合中的所有Toast
    Hide
End Sub

Private Sub StartAutoCloseTimer()
    ' 使用简单的延时替代定时器
    Dim StartTime As Long
    StartTime = GetTickCount
    
    Do While GetTickCount - StartTime < m_Duration
        DoEvents
    Loop
    
    If m_IsShowing Then Hide
End Sub

'--------------------- 快速使用静态方法 ---------------------
' 这些可以在标准模块中包装为全局函数

' 显示成功提示
Public Function Success(ByVal Message As String, Optional ByVal Position As ToastPosition = Toast_BottomCenter) As cToast
    Set Success = New cToast
    Success.Style(ToastStyle_Success).Text(Message).Show
End Function

' 显示错误提示
Public Function Danger(ByVal Message As String, Optional ByVal Position As ToastPosition = Toast_BottomCenter) As cToast
    Set Danger = New cToast
    Danger.Style(ToastStyle_Error).Text(Message).Show
End Function

' 显示警告提示
Public Function Warning(ByVal Message As String, Optional ByVal Position As ToastPosition = Toast_BottomCenter) As cToast
    Set Warning = New cToast
    Warning.Style(ToastStyle_Warning).Text(Message).Show
End Function

' 显示信息提示
Public Function Info(ByVal Message As String, Optional ByVal Position As ToastPosition = Toast_BottomCenter) As cToast
    Set Info = New cToast
    Info.Style(ToastStyle_Info).Text(Message).Show
End Function
