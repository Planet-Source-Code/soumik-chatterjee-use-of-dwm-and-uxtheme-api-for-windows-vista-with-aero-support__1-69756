Attribute VB_Name = "Mod_Declare"
Option Explicit

' DWM API Declarations ##################################
'#################################################

Public Type TRect
    M_Left      As Long
    M_Right     As Long
    M_Top       As Long
    M_Buttom    As Long
End Type

Public Type DWM_BlurBehind
        dwFlags As Long
        fEnable As Boolean
        RGNBlur As Long
        tMAX As Boolean
End Type
 
Public Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Public Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, TheRect As TRect) As Long
Public Declare Function DwmEnableBlurBehindWindow Lib "dwmapi.dll" (ByVal hWnd As Long, BB As DWM_BlurBehind) As Long

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

' GDI+ API Declarations ############################################
'###########################################################

Public Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type

Public Type PICTDESC
   size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type

Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Public Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Public Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

Public Type PAINTSTRUCT
        hDC As Long
        fErase As Boolean
        rcArea As RECT
        fRestore As Boolean
        fIncUpdate As Boolean
        rgbReserve(32) As Byte
End Type

Public Type FA_Type_ARGB
        Alpha As Single
        Red As Single
        Green As Single
        Blue As Single
End Type

Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, LPPaintStruct As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, LPPaintStruct As PAINTSTRUCT) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal DCState As Integer) As Long

Public Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Public Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Public Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Public Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Public Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Public Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Public Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Public Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Public Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Public Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Public Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const PLANES = 14
Public Const BITSPIXEL = 12
Public Const PATCOPY = &HF00021
Public Const PICTYPE_BITMAP = 1
Public Const InterpolationModeHighQualityBicubic = 7
Public Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Public Const UnitPixel = 2

' Theme text with glow effect #######################################
'#########################################################

Public Const DTT_COMPOSITED    As Long = &H2000
Public Const DTT_GLOWSIZE      As Long = &H800
Public Const DTT_TEXTCOLOR As Long = 1


Public Const DT_SINGLELINE     As Long = &H20
Public Const DT_CENTER         As Long = &H1
Public Const DT_VCENTER        As Long = &H4
Public Const DT_NOPREFIX       As Long = &H800
Public Const DT_TEXTFORMAT     As Long = DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX
Public Const SRCCOPY           As Long = &HCC0020

Public Type BITMAPINFOHEADER
    biSize                  As Long
    biWidth                 As Long
    biHeight                As Long
    biPlanes                As Integer
    biBitCount              As Integer
    biCompression           As Long
    biSizeImage             As Long
    biXPelsPerMeter         As Long
    biYPelsPerMeter         As Long
    biClrUsed               As Long
    biClrImportant          As Long
End Type

Public Type RGBQUAD
    rgbBlue                 As Byte
    rgbGreen                As Byte
    rgbRed                  As Byte
    rgbReserved             As Byte
End Type

Public Type BITMAPINFO
    bmiHeader               As BITMAPINFOHEADER
    bmiColors               As RGBQUAD
End Type

Public Type POINTAPI
    x                       As Long
    Y                       As Long
End Type

Public Type DTTOPTS
    dwSize                  As Long
    dwFlags                 As Long
    crText                  As Long
    crBorder                As Long
    crShadow                As Long
    iTextShadowType         As Long
    ptShadowOffset          As POINTAPI
    iBorderSize             As Long
    iFontPropId             As Long
    iColorPropId            As Long
    iStateId                As Long
    fApplyOverlay           As Long
    iGlowSize               As Long
    pfnDrawTextCallback     As Long
    LParam                  As Long
End Type

Public Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal szClases As Long) As Long
Public Declare Function CloseThemeData Lib "UxTheme" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeTextEx Lib "UxTheme" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal text As Long, ByVal iCharCount As Long, ByVal dwFlags As Long, pRect As RECT, pOptions As DTTOPTS) As Long

' Font functions ################################################
'##########################################################

Public Const LF_FACESIZE = 32
Public Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FF_DONTCARE = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" ( _
   lpLogFont As LOGFONT) As Long
Public Declare Function MulDiv Lib "kernel32" ( _
   ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public Const LOGPIXELSY = 90

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" ( _
   ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
   lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_CALCRECT = &H400

Public Declare Function OffsetRect Lib "user32" ( _
   lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long

' Alpha Blending ##############################################
'#########################################################

Public Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Public Declare Sub RtlMoveMemory Lib "Kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type AlphaOptions
        AlphaOption As Byte
        AlphaFlags As Byte
        SourceConstantAlpha As Byte
        AlphaFormat As Byte
End Type

Public Const AC_Src_Over As Long = &H0&
Public Const AC_Src_Alpha As Long = &H1&

' SubClasing #################################################
'##########################################################

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal WParam As Long, ByVal LParam As Long) As Long
Public Declare Function TRACKMOUSEEVENT Lib "comctl32.dll" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT) As Long

Public Const SWP_FrameChanged = 24

Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_CTLCOLOR = &H19
Public Const WM_WININICHANGE = &H1A
Public Const WM_SETTINGCHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_GETOBJECT = &H3D
Public Const WM_COMPACTING = &H41
Public Const WM_COMMNOTIFY = &H44
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_POWER = &H48
Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_NOTIFY = &H4E
Public Const WM_INPUTLANGCHANGEREQUEST = &H50
Public Const WM_INPUTLANGCHANGE = &H51
Public Const WM_TCARD = &H52
Public Const WM_HELP = &H53
Public Const WM_USERCHANGED = &H54
Public Const WM_NOTIFYFORMAT = &H55
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_STYLECHANGING = &H7C
Public Const WM_STYLECHANGED = &H7D
Public Const WM_DISPLAYCHANGE = &H7E
Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_SYNCPAINT = &H88
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121
Public Const WM_MENURBUTTONUP = &H122
Public Const WM_MENUDRAG = &H123
Public Const WM_MENUGETOBJECT = &H124
Public Const WM_UNINITMENUPOPUP = &H125
Public Const WM_MENUCOMMAND = &H126
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_NEXTMENU = &H213
Public Const WM_SIZING = &H214
Public Const WM_CAPTURECHANGED = &H215
Public Const WM_MOVING = &H216
Public Const WM_DEVICECHANGE = &H219
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_ENTERSIZEMOVE = &H231
Public Const WM_EXITSIZEMOVE = &H232
Public Const WM_DROPFILES = &H233
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_REQUEST = &H288
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291
Public Const WM_MOUSEHOVER = &H2A1
Public Const WM_MOUSELEAVE = &H2A3
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_HOTKEY = &H312
Public Const WM_PRINT = &H317
Public Const WM_PRINTCLIENT = &H318
Public Const WM_HANDHELDFIRST = &H358
Public Const WM_HANDHELDLAST = &H35F
Public Const WM_AFXFIRST = &H360
Public Const WM_AFXLAST = &H37F
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_APP = &H8000
Public Const WM_USER = &H400
Public Const WM_REFLECT = WM_USER + &H1C00

Public Type TRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As TrackMouseEventFlags
    dwHoverTime As Long
End Type

Public Enum TrackMouseEventFlags
    TME_HOVER = 1&
    TME_LEAVE = 2&
    TME_NONCLIENT = &H10&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Function GetARGBVal(ByVal LnColor As Long, ByRef ARGBStruct As FA_Type_ARGB) As Long

ARGBStruct.Red = LnColor And &HFF&
ARGBStruct.Green = (LnColor And &HFF00&) \ &H100&
ARGBStruct.Blue = (LnColor And &HFF0000) \ &H10000

End Function

