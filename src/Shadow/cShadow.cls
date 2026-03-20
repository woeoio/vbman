VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-----------------------------------------------
'Name: ClsShadow
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 20/06/2021
'-----------------------------------------------
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As Rect) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "GDI32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTL, ByVal mFlags As ImageLockMode, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mpath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipAddPathRectangleI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetClipHrgn Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mHRgn As Long, ByVal mCombineMode As Long) As Long

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type BitmapData
    Width                       As Long
    Height                      As Long
    stride                      As Long
    PixelFormat                 As Long
    Scan0Ptr                    As Long
    ReservedPtr                 As Long
End Type

Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum
 
Private Type Size
    cx                       As Long
    cy                       As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                   As Long
    biWidth                  As Long
    biHeight                 As Long
    biPlanes                 As Integer
    biBitCount               As Integer
    biCompression            As Long
    biSizeImage              As Long
    biXPelsPerMeter          As Long
    biYPelsPerMeter          As Long
    biClrUsed                As Long
    biClrImportant           As Long
End Type

Private Type BITMAPINFO
    bmiHeader                As BITMAPINFOHEADER
    bmiColors(3)             As Byte
End Type

Private Type BLENDFUNCTION
    BlendOp                  As Byte
    BlendFlags               As Byte
    SourceConstantAlpha      As Byte
    AlphaFormat              As Byte
End Type

Private Const WS_EX_LAYERED     As Long = &H80000
Private Const WS_POPUP          As Long = &H80000000
Private Const WS_EX_TOOLWINDOW  As Long = &H80&
Private Const WS_CHILD          As Long = &H40000000
Private Const WM_PAINT          As Long = &HF&
Private Const WM_MOVE           As Long = &H3
Private Const WM_DESTROY        As Long = &H2
Private Const GWL_STYLE         As Long = -16
Private Const GWL_HWNDPARENT    As Long = -8
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_SHOWWINDOW    As Long = &H40
Private Const SWP_NOSIZE        As Long = &H1
Private Const LOGPIXELSX        As Long = 88
Private Const ULW_ALPHA         As Long = &H2
Private Const DIB_RGB_COLORS    As Long = 0&
Private Const AC_SRC_ALPHA      As Long = &H1
Private Const CombineModeXor    As Long = &H3
Private Const CombineModeIntersect      As Long = &H1
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const CLR_NONE          As Long = &HFFFFFFFF
Dim m_Margin            As Integer
Dim m_BackColor         As OLE_COLOR
Dim m_BackColorOpacity  As Integer
Dim m_BorderColor       As OLE_COLOR
Dim m_BorderOpacity     As Integer
Dim m_BorderWidth       As Integer
Dim m_ShadowColor       As OLE_COLOR
Dim m_ShadowOpacity     As Integer
Dim m_ShadowSize        As Integer
Dim m_ShadowOffsetX     As Integer
Dim m_ShadowOffsetY     As Integer
Dim m_BorderRadius      As Long
Dim m_ContainerBkColor  As Long
Dim nSize               As Long
Dim hImgShadow          As Long
Dim lWidth              As Long
Dim lHeight             As Long
Dim GdipToken           As Long
Dim m_hwnd              As Long
Dim hParent             As Long
Dim c_lhDC              As Long
Dim c_lDIB              As Long
Dim m_Pointer           As Long
Dim m_OldBmp            As Long
Dim DpiFactor           As Single
Dim tWinSize            As Size
Dim tBLENDFUNCTION      As BLENDFUNCTION
Dim cSubClass           As pvSubClass
Dim bRecreateShadow     As Boolean
Dim m_isWinChild        As Boolean
Dim hWinShadow          As Long


Public Property Let ContainerBkColor(ByVal New_Value As Long)
    m_ContainerBkColor = New_Value
End Property

Public Property Get hwnd() As Long
    hwnd = hWinShadow
End Property

Public Property Get Margin() As Integer
    Margin = m_Margin
End Property

Public Property Let Margin(ByVal New_Value As Integer)
    m_Margin = New_Value
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    m_BackColor = New_Value
End Property

Public Property Get BackColorOpacity() As Integer
    BackColorOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorOpacity(ByVal New_Value As Integer)
    m_BackColorOpacity = New_Value
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_Value As OLE_COLOR)
    m_BorderColor = New_Value
End Property

Public Property Get BorderOpacity() As Integer
    BorderOpacity = m_BorderOpacity
End Property

Public Property Let BorderOpacity(ByVal New_Value As Integer)
    m_BorderOpacity = New_Value
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_Value As OLE_COLOR)
    m_ShadowColor = New_Value
    bRecreateShadow = True
End Property

Public Property Get ShadowOpacity() As Integer
    ShadowOpacity = m_ShadowOpacity
End Property

Public Property Let ShadowOpacity(ByVal New_Value As Integer)
    m_ShadowOpacity = New_Value
    bRecreateShadow = True
End Property

Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_Value As Integer)
    m_BorderWidth = New_Value
End Property

Public Property Get ShadowSize() As Integer
    ShadowSize = m_ShadowSize
End Property

Public Property Let ShadowSize(ByVal New_Value As Integer)
    m_ShadowSize = New_Value
    bRecreateShadow = True
End Property

Public Property Get ShadowOffsetX() As Integer
    ShadowOffsetX = m_ShadowOffsetX
End Property

Public Property Let ShadowOffsetX(ByVal New_Value As Integer)
    m_ShadowOffsetX = New_Value
End Property

Public Property Get ShadowOffsetY() As Integer
    ShadowOffsetY = m_ShadowOffsetY
End Property

Public Property Let ShadowOffsetY(ByVal New_Value As Integer)
    m_ShadowOffsetY = New_Value
End Property

Public Property Get BorderRadius() As Long
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal New_Value As Long)
    m_BorderRadius = New_Value
    bRecreateShadow = True
End Property

Public Sub DrawBitmap(hBmp As Long, hwnd As Long, Width As Long, Height As Long)
    Dim tSize As Size
    Dim PT As POINTAPI
    Dim hGraphics As Long
    Dim hImage As Long
    Dim x As Long, y As Long
    Dim hRgn As Long
    
    x = (m_ShadowSize * 2 + m_BorderWidth) * DpiFactor
    y = (m_ShadowSize * 2 + m_BorderWidth) * DpiFactor
    
    If m_ShadowOffsetX < 0 Then x = x - m_ShadowOffsetX * DpiFactor
    If m_ShadowOffsetY < 0 Then y = y - m_ShadowOffsetY * DpiFactor
    
    If x <= 0 Then x = 0
    If y < 0 Then y = 0

    hRgn = CreateRectRgn(0, 0, Width, Height)
    GetWindowRgn hwnd, hRgn
    OffsetRgn hRgn, x, y
  
    GdipCreateFromHDC c_lhDC, hGraphics

    Call GdipSetClipHrgn(hGraphics, hRgn, CombineModeIntersect)

    GdipCreateBitmapFromHBITMAP hBmp, 0&, hImage
    GdipDrawImageRectRectI hGraphics, hImage, x, y, Width, Height, 0, 0, Width, Height, &H2, 0&, 0&, 0& 'TOP_LEFT
    GdipDisposeImage hImage
    GdipDeleteGraphics hGraphics
    
    DeleteObject hRgn
    
    tSize.cx = lWidth
    tSize.cy = lHeight
    
    Call UpdateLayeredWindow(hWinShadow, 0&, ByVal 0&, tSize, c_lhDC, PT, 0&, tBLENDFUNCTION, ULW_ALPHA)
End Sub

Public Sub Refresh()
    If m_hwnd <> 0 Then ShowBorders m_hwnd, m_isWinChild
End Sub

Public Sub ShowBorders(ByVal hwnd As Long, isWinChild As Boolean)
    Dim R As Rect, PT As POINTAPI
    Dim Left As Long, Top As Long
    Dim TotMargin As Long
    
    TotMargin = (m_ShadowSize * 2 + m_BorderWidth + m_Margin) * DpiFactor
    
    GetWindowRect hwnd, R
    tWinSize.cx = R.Right - R.Left
    tWinSize.cy = R.Bottom - R.Top
    lWidth = tWinSize.cx + TotMargin * 2 + Abs(m_ShadowOffsetX) * DpiFactor     '+ m_BorderWidth
    lHeight = tWinSize.cy + TotMargin * 2 + Abs(m_ShadowOffsetY) * DpiFactor    '+ m_BorderWidth
    
    If m_hwnd = 0 Then
        m_hwnd = hwnd
        hParent = GetParent(hwnd)
        Set cSubClass = New pvSubClass
        If isWinChild Then
            m_isWinChild = True
            hWinShadow = CreateWindowEx(0, "static", vbNullString, 0, 0&, 0&, 0&, 0&, hParent, 0&, App.hInstance, 0&)
            Call SetWindowLong(hWinShadow, GWL_STYLE, WS_CHILD)
            SetParent hWinShadow, hParent
            
            cSubClass.ssc_Subclass hWinShadow, , , Me
            cSubClass.ssc_AddMsg hWinShadow, WM_PAINT, MSG_AFTER
            cSubClass.ssc_AddMsg hWinShadow, WM_DESTROY, MSG_BEFORE
            
        Else
            hWinShadow = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_LAYERED, "static", vbNullString, WS_POPUP, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, 0&)
            SetWindowLong hwnd, GWL_HWNDPARENT, hWinShadow
            cSubClass.ssc_Subclass m_hwnd, , , Me
            cSubClass.ssc_AddMsg m_hwnd, WM_MOVE, MSG_AFTER
        End If
    End If
    
    If isWinChild Then
        ScreenToClient hParent, PT
        ClientToScreen hwnd, PT
        R.Left = PT.x
        R.Top = PT.y
    End If
    
    If bRecreateShadow Then
        Call CreateShadow
        bRecreateShadow = False
    End If
    
    Left = R.Left - TotMargin
    Top = R.Top - TotMargin
    If m_ShadowOffsetX < 0 Then Left = Left + m_ShadowOffsetX * DpiFactor
    If m_ShadowOffsetY < 0 Then Top = Top + m_ShadowOffsetY * DpiFactor
    
    SetWindowPos hWinShadow, hwnd, Left, Top, lWidth, lHeight, SWP_NOACTIVATE Or SWP_SHOWWINDOW 'Or SWP_FRAMECHANGED
    If Not isWinChild Then SetWindowLong hwnd, GWL_HWNDPARENT, hWinShadow
    
    Update
End Sub

Public Sub HideBorders()
    ShowWindow hWinShadow, vbHide
    pvCleanUp
End Sub

Private Function pvCreateDIB() As Boolean
    Dim tBITMAPINFO     As BITMAPINFO

    pvCleanUp
    
    With tBITMAPINFO.bmiHeader
        .biSize = Len(tBITMAPINFO.bmiHeader)
        .biBitCount = 32
        .biHeight = lHeight
        .biWidth = lWidth
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * 4
    End With

    c_lhDC = CreateCompatibleDC(0)
    c_lDIB = CreateDIBSection(c_lhDC, tBITMAPINFO, DIB_RGB_COLORS, m_Pointer, 0&, 0&)
    m_OldBmp = SelectObject(c_lhDC, c_lDIB)
End Function

Private Sub pvCleanUp()
    If c_lDIB Then Call DeleteObject(SelectObject(c_lhDC, m_OldBmp)): c_lDIB = 0
    If c_lhDC Then Call DeleteDC(c_lhDC): c_lhDC = 0
End Sub
'*2
Public Sub Update()
    Dim hGraphics As Long
    Dim tSize As Size
    Dim PT As Size
    Dim x As Long, y As Long
    Dim hRgn As Long
    Dim hPath As Long, hBrush As Long, hPen As Long
    Dim Rect As Rect
    Dim lDC As Long
    Dim BkColor As Long
    
    pvCreateDIB
    
    If m_isWinChild Then
        If m_ContainerBkColor = CLR_NONE Then
            lDC = GetDC(hParent)
            BkColor = GetBkColor(lDC)
            ReleaseDC hParent, lDC
        Else
            BkColor = m_ContainerBkColor
            If (BkColor And &H80000000) Then BkColor = GetSysColor(BkColor And &HFF&)
        End If
        
        lDC = GetDC(hWinShadow)
        hBrush = CreateSolidBrush(BkColor)
        SetRect Rect, 0, 0, lWidth, lHeight
        FillRect c_lhDC, Rect, hBrush
        DeleteObject hBrush
    End If

    GdipCreateFromHDC c_lhDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

    x = (m_ShadowSize * 2 + m_BorderWidth / 2) * DpiFactor
    y = (m_ShadowSize * 2 + m_BorderWidth / 2) * DpiFactor
    
    If m_ShadowOffsetX < 0 Then x = x - m_ShadowOffsetX * DpiFactor
    If m_ShadowOffsetY < 0 Then y = y - m_ShadowOffsetY * DpiFactor
    
    If x <= 0 Then x = 0
    If y < 0 Then y = 0
    
    hPath = CreateRoundPath(x, y, _
        tWinSize.cx + (m_BorderWidth + m_Margin * 2) * DpiFactor, _
        tWinSize.cy + (m_BorderWidth + m_Margin * 2) * DpiFactor, m_BorderRadius * DpiFactor)
    
    GdipSetClipPath hGraphics, hPath, CombineModeXor

    x = 0: y = 0
    If m_ShadowOffsetX > 0 Then x = m_ShadowOffsetX * DpiFactor
    If m_ShadowOffsetY > 0 Then y = m_ShadowOffsetY * DpiFactor
    
    If m_ShadowSize > 0 And m_ShadowOpacity > 0 Then
        RenderStretchPlus hGraphics, x, y, lWidth - Abs(m_ShadowOffsetX) * DpiFactor, lHeight - Abs(m_ShadowOffsetY) * DpiFactor, hImgShadow, 0, 0, nSize + m_ShadowSize * 2 * DpiFactor, nSize + m_ShadowSize * 2 * DpiFactor, (m_ShadowSize * 2 + m_BorderRadius) * DpiFactor
    End If
    GdipResetClip hGraphics
    
    x = (m_ShadowSize * 2 + m_BorderWidth + m_Margin) * DpiFactor
    y = (m_ShadowSize * 2 + m_BorderWidth + m_Margin) * DpiFactor

    hRgn = CreateRectRgn(0, 0, tWinSize.cx, tWinSize.cy)
    Call GetWindowRgn(m_hwnd, hRgn)
    OffsetRgn hRgn, x, y
    GdipSetClipHrgn hGraphics, hRgn, CombineModeXor
    DeleteObject hRgn

    GdipCreateSolidFill RGBtoARGB(m_BackColor, m_BackColorOpacity), hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    
    If m_BorderWidth > 0 And m_BorderOpacity > 0 Then
        GdipCreatePen1 RGBtoARGB(m_BorderColor, m_BorderOpacity), m_BorderWidth * DpiFactor, &H2, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If

    GdipDeletePath hPath
    GdipDeleteGraphics hGraphics
    If m_isWinChild Then
        BitBlt lDC, 0, 0, lWidth, lHeight, c_lhDC, 0, 0, vbSrcCopy
        ReleaseDC hWinShadow, lDC
        pvCleanUp
    Else
        tSize.cx = lWidth: tSize.cy = lHeight
        Call UpdateLayeredWindow(hWinShadow, 0&, ByVal 0&, tSize, c_lhDC, PT, 0&, tBLENDFUNCTION, ULW_ALPHA)
    End If
End Sub

Private Sub CreateShadow()
    Dim hImage As Long
    Dim hGraphics As Long
    Dim hPath As Long, hBrush As Long
    nSize = (m_ShadowSize * 4 + m_BorderRadius * 2) * DpiFactor
    If hImgShadow Then GdipDisposeImage hImgShadow
    GdipCreateBitmapFromScan0 nSize, nSize, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics

    hPath = CreateRoundPath(0, 0, nSize, nSize, m_BorderRadius * 1.5 * DpiFactor)
    GdipCreateSolidFill RGBtoARGB(m_ShadowColor, m_ShadowOpacity), hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
    
    hImgShadow = CreateBlurShadowImage(hImage, m_ShadowColor, m_ShadowSize * DpiFactor, 0, 0, nSize, nSize)
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
End Sub

Private Sub Class_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    
    DpiFactor = GetWindowsDPI
    
    With tBLENDFUNCTION
        .AlphaFormat = AC_SRC_ALPHA
        .SourceConstantAlpha = 255
    End With
    
    m_BorderColor = vbActiveBorder
    m_BackColor = vbWindowBackground
    m_BorderWidth = 1
    m_BorderRadius = 10
    m_ShadowOpacity = 8
    m_BackColorOpacity = 100
    m_BorderOpacity = 100
    m_ShadowSize = 3
    m_ContainerBkColor = CLR_NONE
    bRecreateShadow = True
End Sub

Private Sub Class_Terminate()
    
    If Not m_isWinChild Then
        If Not cSubClass Is Nothing Then
            cSubClass.ssc_UnSubclass m_hwnd
        End If
    End If
    If hWinShadow Then DestroyWindow hWinShadow
    pvCleanUp
    If hImgShadow Then GdipDisposeImage hImgShadow: hImgShadow = 0
    Call GdiplusShutdown(GdipToken)
End Sub
'*1
Private Function CreateRoundPath(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal Radius As Long) As Long
    Dim mpath As Long
    If GdipCreatePath(&H0, mpath) = 0 Then
        If Radius = 0 Then
            GdipAddPathRectangleI mpath, Left, Top, Width, Height
            CreateRoundPath = mpath
            Exit Function
        End If
        GdipAddPathArcI mpath, Left, Top, Radius, Radius, 180, 90
        GdipAddPathArcI mpath, Left + Width - Radius, Top, Radius, Radius, 270, 90
        GdipAddPathArcI mpath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
        GdipAddPathArcI mpath, Left, Top + Height - Radius, Radius, Radius, 90, 90
        GdipClosePathFigure mpath
    End If
    CreateRoundPath = mpath
End Function

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long 'By LaVople
    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
End Function

Private Function CreateBlurShadowImage(ByVal hImage As Long, ByVal Color As Long, blurDepth As Integer, _
                                        Optional ByVal Left As Long, Optional ByVal Top As Long, _
                                        Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
                                        
    Dim REC As RECTL
    Dim x As Long, y As Long
    Dim hImgShadow As Long
    Dim bmpData1 As BitmapData
    Dim bmpData2 As BitmapData
    Dim t2xBlur As Long
    Dim R As Long, G As Long, B As Long
    Dim dBytes() As Byte
    Dim srcBytes() As Byte
    Dim vTally() As Long
    Dim tAlpha As Long, tColumn As Long, tAvg As Long
    Dim initY As Long, initYstop As Long, initYstart As Long
    
    If hImage = 0& Then Exit Function
 
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    t2xBlur = blurDepth * 2
 
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)

    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
 
    SetRect REC, Left, Top, Width, Height
 
    ReDim srcBytes(REC.Width * 4 - 1&, REC.Height - 1&)
  
    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
   
    Call GdipBitmapLockBits(hImage, REC, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppPARGB, bmpData1)
 
    SetRect REC, Left, Top, Width + t2xBlur, Height + t2xBlur
    
    Call GdipCreateBitmapFromScan0(REC.Width, REC.Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImgShadow)

    ReDim dBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    
    With bmpData2
        .Scan0Ptr = VarPtr(dBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
    
    Call GdipBitmapLockBits(hImgShadow, REC, ImageLockModeUserInputBuf Or ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppPARGB, bmpData2)
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For y = 0 To Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = y - t2xBlur ' start n blur rows above y
        End If
        ' how may source rows can we use for blurring?
        If y < Height Then initYstop = y Else initYstop = Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + srcBytes(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For x = 3 To (Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            
            ' set shadow alpha using blur average
            dBytes(x, y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case dBytes(x, y)
            Case 255
                dBytes(x - 1, y) = R
                dBytes(x - 2, y) = G
                dBytes(x - 3, y) = B
            Case 0
            Case Else
                dBytes(x - 1, y) = R * dBytes(x, y) \ 255
                dBytes(x - 2, y) = G * dBytes(x, y) \ 255
                dBytes(x - 3, y) = B * dBytes(x, y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + srcBytes(x + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For x = x To (Width + t2xBlur - 1) * 4 - 1 Step 4
            dBytes(x, y) = tAlpha \ tAvg
            Select Case dBytes(x, y)
            Case 255
                dBytes(x - 1, y) = R
                dBytes(x - 2, y) = G
                dBytes(x - 3, y) = B
            Case 0
            Case Else
                dBytes(x - 1, y) = R * dBytes(x, y) \ 255
                dBytes(x - 2, y) = G * dBytes(x, y) \ 255
                dBytes(x - 3, y) = B * dBytes(x, y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
 
    Call GdipBitmapUnlockBits(hImage, bmpData1)
    Call GdipBitmapUnlockBits(hImgShadow, bmpData2)
    
    CreateBlurShadowImage = hImgShadow
End Function

Private Sub RenderStretchPlus(ByVal hGraphics As Long, _
            ByVal DestX As Long, ByVal DestY As Long, ByVal DestW As Long, ByVal DestH As Long, ByVal hImage As Long, _
            ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Size As Long)

    Dim Sx2 As Long
  
    Sx2 = Size * 2

    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY, Size, Size, x, y, Size, Size, &H2, 0&, 0&, 0& 'TOP_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY, DestW - Sx2, Size, x + Size, y, Width - Sx2, Size, &H2, 0&, 0&, 0& 'TOP_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY, Size, Size, x + Width - Size, y, Size, Size, &H2, 0&, 0&, 0& 'TOP_RIGHT
    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY + Size, Size, DestH - Sx2, x, y + Size, Size, Height - Sx2, &H2, 0&, 0&, 0& 'MID_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY + Size, DestW - Sx2, DestH - Sx2, x + Size, y + Size, Width - Sx2, Height - Sx2, &H2, 0&, 0&, 0& 'MID_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY + Size, Size, DestH - Sx2, x + Width - Size, y + Size, Size, Height - Sx2, &H2, 0&, 0&, 0& 'MID_RIGHT
    GdipDrawImageRectRectI hGraphics, hImage, DestX, DestY + DestH - Size, Size, Size, x, y + Height - Size, Size, Size, &H2, 0&, 0&, 0& 'BOTTOM_LEFT
    GdipDrawImageRectRectI hGraphics, hImage, DestX + Size, DestY + DestH - Size, DestW - Sx2, Size, x + Size, y + Height - Size, Width - Sx2, Size, &H2, 0&, 0&, 0& 'BOTTOM_CENTER
    GdipDrawImageRectRectI hGraphics, hImage, DestX + DestW - Size, DestY + DestH - Size, Size, Size, x + Width - Size, y + Height - Size, Size, Size, &H2, 0&, 0&, 0& 'BOTTOM_RIGHT
End Sub

Private Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
 Else
    HiWord = dw \ 65535
 End If
End Function

Private Function LoWord(dw As Long) As Integer
  If dw And &H8000& Then
      LoWord = &H8000 Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function

Private Sub WndProc(ByVal bBefore As Boolean, _
   ByRef bHandled As Boolean, _
   ByRef lReturn As Long, _
   ByVal hwnd As Long, _
   ByVal uMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long, _
   ByRef lParamUser As Long)

    If uMsg = WM_PAINT Then ' CHILD WINDOWS
        Update
    ElseIf uMsg = WM_DESTROY Then
        cSubClass.ssc_UnSubclass hWinShadow
    ElseIf uMsg = WM_MOVE Then 'PARENT WINDOWS
        Dim x As Long, y As Long

        x = LoWord(lParam) - (m_ShadowSize * 2 + m_BorderWidth + m_Margin) * DpiFactor
        y = HiWord(lParam) - (m_ShadowSize * 2 + m_BorderWidth + m_Margin) * DpiFactor
        If m_ShadowOffsetX < 0 Then x = x + m_ShadowOffsetX * DpiFactor
        If m_ShadowOffsetY < 0 Then y = y + m_ShadowOffsetY * DpiFactor

        SetWindowPos hWinShadow, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE
    End If
End Sub


