Attribute VB_Name = "mDeclares"
Public Const MSM_NCACTIVATE As Long = &H86

Public Enum eMsg
    ALL_MESSAGES = -1&
    WM_NULL = &H0&
    WM_CREATE = &H1&
    WM_DESTROY = &H2&
    WM_MOVE = &H3&
    WM_SIZE = &H5&
    WM_ACTIVATE = &H6&
    WM_SETFOCUS = &H7&
    WM_KILLFOCUS = &H8&
    WM_ENABLE = &HA&
    WM_SETREDRAW = &HB&
    WM_SETTEXT = &HC&
    WM_GETTEXT = &HD&
    WM_GETTEXTLENGTH = &HE&
    WM_PAINT = &HF&
    WM_CLOSE = &H10&
    WM_QUERYENDSESSION = &H11&
    WM_QUIT = &H12&
    WM_QUERYOPEN = &H13&
    WM_ERASEBKGND = &H14&
    WM_SYSCOLORCHANGE = &H15&
    WM_ENDSESSION = &H16&
    WM_SHOWWINDOW = &H18&
    WM_WININICHANGE = &H1A&
    WM_SETTINGCHANGE = &H1A&
    WM_DEVMODECHANGE = &H1B&
    WM_ACTIVATEAPP = &H1C&
    WM_FONTCHANGE = &H1D&
    WM_TIMECHANGE = &H1E&
    WM_CANCELMODE = &H1F&
    WM_SETCURSOR = &H20&
    WM_MOUSEACTIVATE = &H21&
    WM_CHILDACTIVATE = &H22&
    WM_QUEUESYNC = &H23&
    WM_GETMINMAXINFO = &H24&
    WM_PAINTICON = &H26&
    WM_ICONERASEBKGND = &H27&
    WM_NEXTDLGCTL = &H28&
    WM_SPOOLERSTATUS = &H2A&
    WM_DRAWITEM = &H2B&
    WM_MEASUREITEM = &H2C&
    WM_DELETEITEM = &H2D&
    WM_VKEYTOITEM = &H2E&
    WM_CHARTOITEM = &H2F&
    WM_SETFONT = &H30&
    WM_GETFONT = &H31&
    WM_SETHOTKEY = &H32&
    WM_GETHOTKEY = &H33&
    WM_QUERYDRAGICON = &H37&
    WM_COMPAREITEM = &H39&
    WM_GETOBJECT = &H3D&
    WM_COMPACTING = &H41&
    WM_WINDOWPOSCHANGING = &H46&
    WM_WINDOWPOSCHANGED = &H47&
    WM_POWER = &H48&
    WM_COPYDATA = &H4A&
    WM_CANCELJOURNAL = &H4B&
    WM_NOTIFY = &H4E&
    WM_INPUTLANGCHANGEREQUEST = &H50&
    WM_INPUTLANGCHANGE = &H51&
    WM_TCARD = &H52&
    WM_HELP = &H53&
    WM_USERCHANGED = &H54&
    WM_NOTIFYFORMAT = &H55&
    WM_CONTEXTMENU = &H7B&
    WM_STYLECHANGING = &H7C&
    WM_STYLECHANGED = &H7D&
    WM_DISPLAYCHANGE = &H7E&
    WM_GETICON = &H7F&
    WM_SETICON = &H80&
    WM_NCCREATE = &H81&
    WM_NCDESTROY = &H82&
    WM_NCCALCSIZE = &H83&
    WM_NCHITTEST = &H84&
    WM_NCPAINT = &H85&
    WM_NCACTIVATE = &H86&
    WM_GETDLGCODE = &H87&
    WM_SYNCPAINT = &H88&
    WM_NCMOUSEMOVE = &HA0&
    WM_NCLBUTTONDOWN = &HA1&
    WM_NCLBUTTONUP = &HA2&
    WM_NCLBUTTONDBLCLK = &HA3&
    WM_NCRBUTTONDOWN = &HA4&
    WM_NCRBUTTONUP = &HA5&
    WM_NCRBUTTONDBLCLK = &HA6&
    WM_NCMBUTTONDOWN = &HA7&
    WM_NCMBUTTONUP = &HA8&
    WM_NCMBUTTONDBLCLK = &HA9&
    WM_KEYFIRST = &H100&
    WM_KEYDOWN = &H100&
    WM_KEYUP = &H101&
    WM_CHAR = &H102&
    WM_DEADCHAR = &H103&
    WM_SYSKEYDOWN = &H104&
    WM_SYSKEYUP = &H105&
    WM_SYSCHAR = &H106&
    WM_SYSDEADCHAR = &H107&
    WM_KEYLAST = &H108&
    WM_IME_STARTCOMPOSITION = &H10D&
    WM_IME_ENDCOMPOSITION = &H10E&
    WM_IME_COMPOSITION = &H10F&
    WM_IME_KEYLAST = &H10F&
    WM_INITDIALOG = &H110&
    WM_COMMAND = &H111&
    WM_SYSCOMMAND = &H112&
    WM_TIMER = &H113&
    WM_HSCROLL = &H114&
    WM_VSCROLL = &H115&
    WM_INITMENU = &H116&
    WM_INITMENUPOPUP = &H117&
    WM_MENUSELECT = &H11F&
    WM_MENUCHAR = &H120&
    WM_ENTERIDLE = &H121&
    WM_MENURBUTTONUP = &H122&
    WM_MENUDRAG = &H123&
    WM_MENUGETOBJECT = &H124&
    WM_UNINITMENUPOPUP = &H125&
    WM_MENUCOMMAND = &H126&
    WM_CTLCOLORMSGBOX = &H132&
    WM_CTLCOLOREDIT = &H133&
    WM_CTLCOLORLISTBOX = &H134&
    WM_CTLCOLORBTN = &H135&
    WM_CTLCOLORDLG = &H136&
    WM_CTLCOLORSCROLLBAR = &H137&
    WM_CTLCOLORSTATIC = &H138&
    WM_MOUSEFIRST = &H200&
    WM_MOUSEMOVE = &H200&
    WM_LBUTTONDOWN = &H201&
    WM_LBUTTONUP = &H202&
    WM_LBUTTONDBLCLK = &H203&
    WM_RBUTTONDOWN = &H204&
    WM_RBUTTONUP = &H205&
    WM_RBUTTONDBLCLK = &H206&
    WM_MBUTTONDOWN = &H207&
    WM_MBUTTONUP = &H208&
    WM_MBUTTONDBLCLK = &H209&
    WM_MOUSEWHEEL = &H20A&
    WM_PARENTNOTIFY = &H210&
    WM_ENTERMENULOOP = &H211&
    WM_EXITMENULOOP = &H212&
    WM_NEXTMENU = &H213&
    WM_SIZING = &H214&
    WM_CAPTURECHANGED = &H215&
    WM_MOVING = &H216&
    WM_DEVICECHANGE = &H219&
    WM_MDICREATE = &H220&
    WM_MDIDESTROY = &H221&
    WM_MDIACTIVATE = &H222&
    WM_MDIRESTORE = &H223&
    WM_MDINEXT = &H224&
    WM_MDIMAXIMIZE = &H225&
    WM_MDITILE = &H226&
    WM_MDICASCADE = &H227&
    WM_MDIICONARRANGE = &H228&
    WM_MDIGETACTIVE = &H229&
    WM_MDISETMENU = &H230&
    WM_ENTERSIZEMOVE = &H231&
    WM_EXITSIZEMOVE = &H232&
    WM_DROPFILES = &H233&
    WM_MDIREFRESHMENU = &H234&
    WM_IME_SETCONTEXT = &H281&
    WM_IME_NOTIFY = &H282&
    WM_IME_CONTROL = &H283&
    WM_IME_COMPOSITIONFULL = &H284&
    WM_IME_SELECT = &H285&
    WM_IME_CHAR = &H286&
    WM_IME_REQUEST = &H288&
    WM_IME_KEYDOWN = &H290&
    WM_IME_KEYUP = &H291&
    WM_MOUSEHOVER = &H2A1&
    WM_MOUSELEAVE = &H2A3&
    WM_CUT = &H300&
    WM_COPY = &H301&
    WM_PASTE = &H302&
    WM_CLEAR = &H303&
    WM_UNDO = &H304&
    WM_RENDERFORMAT = &H305&
    WM_RENDERALLFORMATS = &H306&
    WM_DESTROYCLIPBOARD = &H307&
    WM_DRAWCLIPBOARD = &H308&
    WM_PAINTCLIPBOARD = &H309&
    WM_VSCROLLCLIPBOARD = &H30A&
    WM_SIZECLIPBOARD = &H30B&
    WM_ASKCBFORMATNAME = &H30C&
    WM_CHANGECBCHAIN = &H30D&
    WM_HSCROLLCLIPBOARD = &H30E&
    WM_QUERYNEWPALETTE = &H30F&
    WM_PALETTEISCHANGING = &H310&
    WM_PALETTECHANGED = &H311&
    WM_HOTKEY = &H312&
    WM_PRINT = &H317&
    WM_PRINTCLIENT = &H318&
    WM_THEMECHANGED = &H31A&
    WM_HANDHELDFIRST = &H358&
    WM_HANDHELDLAST = &H35F&
    WM_AFXFIRST = &H360&
    WM_AFXLAST = &H37F&
    WM_PENWINFIRST = &H380&
    WM_PENWINLAST = &H38F&
    WM_USER = &H400&
    WM_APP = &H8000&
    WM_UPDATEUISTATE = &H128&
    WM_CHANGEUISTATE = &H127&
End Enum

Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    hwnd_notopmost = -2
End Enum

Public Const SC_CLOSE           As Long = &HF060
Public Const SC_MAXIMIZE        As Long = &HF030&
Public Const SC_MINIMIZE        As Long = &HF020&
Public Const SC_RESTORE         As Long = &HF120&
Public Const SC_MOVE            As Long = &HF010&
Public Const HTCAPTION          As Integer = 2
Public Const HTBOTTOM           As Integer = 15
Public Const HTBOTTOMLEFT       As Integer = 16
Public Const HTBOTTOMRIGHT      As Integer = 17
Public Const HTLEFT             As Integer = 10
Public Const HTRIGHT            As Integer = 11
Public Const HTTOP              As Integer = 12
Public Const HTTOPLEFT          As Integer = 13
Public Const HTTOPRIGHT         As Integer = 14

Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYSCREEN = 1

Public Const WA_INACTIVE As Long = 0
Public Const WA_ACTIVE As Long = 1
Public Const WA_CLICKACTIVE As Long = 2

Public Const GWL_STYLE As Long = -16
Public Const GWL_EXSTYLE As Long = -20
Public Const GWL_HWNDPARENT = (-8)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TOOLWINDOW As Long = &H80&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const SS_OWNERDRAW = &HD&
Public Const WS_POPUP = &H80000000

Public Const ULW_OPAQUE = &H4
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const BI_RGB As Long = 0&
Public Const DIB_RGB_COLORS As Long = 0
Public Const AC_SRC_ALPHA As Long = &H1
Public Const AC_SRC_OVER = &H0

Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_PATH_ELLIPSIS = &H4000

Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_DOT = 2
Public Const PS_NULL = 5
Public Const PS_SOLID = 0
Public Const PS_USERSTYLE = 7

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type POINTAPI
    X                               As Long
    Y                               As Long
End Type

Public Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&

Public Const TPM_NONOTIFY = &H80&           '/* Don't send any notification msgs */
Public Const TPM_RETURNCMD = &H100
Public Const TPM_HORIZONTAL = &H0          '/* Horz alignment matters more */
Public Const TPM_VERTICAL = &H40           '/* Vert alignment matters more */

   ' Win98/2000 menu animation and menu within menu options:
Public Const TPM_RECURSE = &H1&
Public Const TPM_HORPOSANIMATION = &H400&
Public Const TPM_HORNEGANIMATION = &H800&
Public Const TPM_VERPOSANIMATION = &H1000&
Public Const TPM_VERNEGANIMATION = &H2000&
   ' Win2000 only:
Public Const TPM_NOANIMATION = &H4000&

Public Type MINMAXINFO
    ptReserved                      As POINTAPI
    ptMaxSize                       As POINTAPI
    ptMaxPosition                   As POINTAPI
    ptMinTrackSize                  As POINTAPI
    ptMaxTrackSize                  As POINTAPI
End Type

Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Public Declare Function TrackPopupMenuByLong Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Long) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal Handle As Long, ByVal dw As Long) As Long
Public Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hDCDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, X As Long, Y As Long) As Long

Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

'******************************************************
' GDI Plus
'******************************************************
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As String, Image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, BITMAP As Long) As GpStatus

Public Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
                                       ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
                                       ' its internal image codecs.
End Type


Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)

Public Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Public Function TranslateColor(ByVal theColor As Long) As Long
  Call OleTranslateColor(theColor, 0, TranslateColor)
End Function

Public Sub RGBToHLS( _
     ByVal R As Long, ByVal G As Long, ByVal B As Long, _
     H As Single, S As Single, l As Single _
     )
 Dim Max As Single
 Dim Min As Single
 Dim Delta As Single
 Dim rR As Single, Rg As Single, rB As Single

     rR = R / 255: Rg = G / 255: rB = B / 255

 '{Given: rgb each in [0,1].
 ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
         Max = Maximum(rR, Rg, rB)
         Min = Minimum(rR, Rg, rB)
             l = (Max + Min) / 2 '{This is the lightness}
         '{Next calculate saturation}
         If Max = Min Then
             'begin {Acrhomatic case}
             S = 0
             H = 0
             'end {Acrhomatic case}
         Else
             'begin {Chromatic case}
                 '{First calculate the saturation.}
             If l <= 0.5 Then
                 S = (Max - Min) / (Max + Min)
             Else
                 S = (Max - Min) / (2 - Max - Min)
             End If
             '{Next calculate the hue.}
             Delta = Max - Min
             If rR = Max Then
                     H = (Rg - rB) / Delta '{Resulting color is between yellow and magenta}
             ElseIf Rg = Max Then
                 H = 2 + (rB - rR) / Delta '{Resulting color is between cyan and yellow}
             ElseIf rB = Max Then
                 H = 4 + (rR - Rg) / Delta '{Resulting color is between magenta and cyan}
             End If
         'end {Chromatic Case}
     End If
 End Sub

Public Sub HLSToRGB( _
     ByVal H As Single, ByVal S As Single, ByVal l As Single, _
     R As Long, G As Long, B As Long _
     )
Dim rR As Single, Rg As Single, rB As Single
Dim Min As Single, Max As Single

     If S = 0 Then
     ' Achromatic case:
     rR = l: Rg = l: rB = l
     Else
     ' Chromatic case:
     ' delta = Max-Min
     If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - S)
     Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - S * (1 - l)
     End If
     ' Get the Max value:
     Max = 2 * l - Min
     
     ' Now depending on sector we can evaluate the h,l,s:
     If (H < 1) Then
         rR = Max
         If (H < 0) Then
             Rg = Min
             rB = Rg - H * (Max - Min)
         Else
             rB = Min
             Rg = H * (Max - Min) + rB
         End If
     ElseIf (H < 3) Then
         Rg = Max
         If (H < 2) Then
             rB = Min
             rR = rB - (H - 2) * (Max - Min)
         Else
             rR = Min
             rB = (H - 2) * (Max - Min) + rR
         End If
     Else
         rB = Max
         If (H < 4) Then
             rR = Min
             Rg = rR - (H - 4) * (Max - Min)
         Else
             Rg = Min
             rR = (H - 4) * (Max - Min) + Rg
         End If
         
     End If
             
     End If
     R = rR * 255: G = Rg * 255: B = rB * 255
 End Sub

 Private Function Maximum(rR As Single, Rg As Single, rB As Single) As Single
   If (rR > Rg) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > Rg) Then
         Maximum = rB
      Else
         Maximum = Rg
      End If
   End If
End Function
Private Function Minimum(rR As Single, Rg As Single, rB As Single) As Single
   If (rR < Rg) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < Rg) Then
         Minimum = rB
      Else
         Minimum = Rg
      End If
   End If
End Function

Public Function BlendColor(lColor1 As Long, lColor2 As Long, Level As Byte) As Long
  Dim lRed    As Long
  Dim lGrn    As Long
  Dim lBlu    As Long
  Dim lRed2   As Long
  Dim lGrn2   As Long
  Dim lBlu2   As Long
  Dim fRedStp As Single
  Dim fGrnStp As Single
  Dim fBluStp As Single
  If Level = 0 Then
    BlendColor = lColor1
  ElseIf Level = 255 Then
    BlendColor = lColor2
  Else
    'Extract Red, Blue and Green values from the start and end colors.
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000
    lRed2 = (lColor2 And &HFF&)
    lGrn2 = (lColor2 And &HFF00&) / &H100
    lBlu2 = (lColor2 And &HFF0000) / &H10000
    
    If lRed2 >= lRed Then
      fRedStp = (lRed2 - lRed) / 255
    Else
      fRedStp = (lRed - lRed2) / 255
      fRedStp = -fRedStp
    End If
    If lGrn2 >= lGrn Then
      fGrnStp = (lGrn2 - lGrn) / 255
    Else
      fGrnStp = (lGrn - lGrn2) / 255
      fGrnStp = -fGrnStp
    End If
    If lBlu2 >= lBlu Then
      fBluStp = (lBlu2 - lBlu) / 255
    Else
      fBluStp = (lBlu - lBlu2) / 255
      fBluStp = -fBluStp
    End If

    BlendColor = RGB(lRed + (fRedStp * (255 - Level)), lGrn + (fGrnStp * (255 - Level)), lBlu + (fBluStp * (255 - Level)))
  End If
End Function

Private Function HookAddress(ByVal lPtr As Long) As Long
   HookAddress = lPtr
End Function

Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   If Not (lPtr = 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory objT, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set ObjectFromPtr = objT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory objT, 0&, 4
   End If
End Property

Public Function IsMouseOver(m_hWnd) As Boolean

    Dim pt As POINTAPI

    GetCursorPos pt
    IsMouseOver = (WindowFromPoint(pt.X, pt.Y) = m_hWnd)

End Function

