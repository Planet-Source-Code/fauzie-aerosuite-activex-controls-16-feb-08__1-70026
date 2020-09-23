VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrameControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Read comments in clsCustomWindow for general information

' APIs used for debugging & can be removed when finalized
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function IsMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_NCRBUTTONDOWN As Long = &HA4
Private Const WM_RBUTTONDOWN As Long = &H204
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long

' API Declarations used in this class
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function PtInRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32.dll" () As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function EqualRect Lib "user32.dll" (lpRect1 As RECT, lpRect2 As RECT) As Long
Private Declare Function DrawMenuBar Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DrawFrameControl Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

' API Constants used in this class
Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
Private Const COLOR_GRADIENTACTIVECAPTION As Long = 27
Private Const DFC_BUTTON As Long = 4
Private Const DFC_CAPTION As Long = 1
Private Const DFCS_BUTTONPUSH As Long = &H10
Private Const DFCS_CAPTIONCLOSE As Long = &H0
Private Const DFCS_CAPTIONHELP As Long = &H4
Private Const DFCS_CAPTIONMAX As Long = &H2
Private Const DFCS_CAPTIONMIN As Long = &H1
Private Const DFCS_HOT As Long = &H1000
Private Const DFCS_CAPTIONRESTORE As Long = &H3
Private Const DFCS_INACTIVE As Long = &H100
Private Const DFCS_PUSHED As Long = &H200
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const IMAGE_ICON As Long = 1
Private Const LR_COPYFROMRESOURCE As Long = &H4000
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16
Private Const GWL_WNDPROC As Long = -4
Private Const HTNOWHERE As Long = 0
Private Const HTCLIENT As Long = 1
Private Const HTCAPTION As Long = 2
Private Const HTSYSMENU As Long = 3
Private Const HTGROWBOX As Long = 4
Private Const HTMENU As Long = 5
Private Const HTMenuPlus As Long = 30 ' custom
Private Const HTMINBUTTON As Long = 8
Private Const HTMAXBUTTON As Long = 9
Private Const HTLEFT As Long = 10
Private Const HTRIGHT As Long = 11
Private Const HTTOP As Long = 12
Private Const HTTOPLEFT As Long = 13
Private Const HTTOPRIGHT As Long = 14
Private Const HTBOTTOM As Long = 15
Private Const HTBOTTOMLEFT As Long = 16
Private Const HTBOTTOMRIGHT As Long = 17
Private Const HTBORDER As Long = 18
Private Const HTCLOSE As Long = 20
Private Const HTHELP As Long = 21
Private Const MIIM_STATE     As Long = &H1&
Private Const MIIM_ID        As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const SC_CLOSE = &HF060&
Private Const SC_MOVE As Long = &HF010&
Private Const SC_KEYMENU As Long = &HF100&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_MOUSEMENU As Long = &HF090
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE As Long = &HF000&
Private Const SC_MOVE_Clone As Long = &H7AA ' custom
Private Const SC_SIZE_Clone As Long = &H7AB ' custom
Private Const SC_CLOSE_Clone = &H7AC ' custom
Private Const SC_MINIMIZE_Clone = &H7AD ' custom
Private Const SC_RESTORE_Clone = &H7AE ' custom
Private Const SC_MAXIMIZE_Clone As Long = &H7AF ' custom
Private Const SM_CYICON As Long = 12
Private Const SM_CXICON As Long = 11
Private Const SM_CXDLGFRAME As Long = 7
Private Const SM_CYDLGFRAME As Long = 8
Private Const SM_CXBORDER As Long = 5
Private Const SM_CYBORDER As Long = 6
Private Const SM_CYSMCAPTION As Long = 51
Private Const SM_CYCAPTION As Long = 4
Private Const SM_CXFRAME As Long = 32
Private Const SM_CYFRAME As Long = 33
Private Const SM_CXSIZE As Long = 30
Private Const SM_CYSIZE As Long = 31
Private Const SM_CXSMICON As Long = 49
Private Const SM_CYSMICON As Long = 50
Private Const SM_CXMINTRACK As Long = 34
Private Const SM_CYMINTRACK As Long = 35
Private Const SPI_GETWORKAREA As Long = 48
Private Const WM_APPACTIVATE As Long = &H1C
Private Const WM_CANCELMODE As Long = &H1F
Private Const WM_CLOSE As Long = &H10
Private Const WM_COMMAND As Long = &H111
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_DESTROY As Long = &H2
Private Const WM_GETICON As Long = &H7F
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const WM_GETSYSMENU = &H313
Private Const WM_GETTEXT As Long = &HD
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_NCACTIVATE As Long = &H86
Private Const WM_NCCALCSIZE As Long = &H83
Private Const WM_NCHITTEST As Long = &H84
Private Const WM_NCLBUTTONDBLCLK As Long = &HA3
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const WM_NCLBUTTONUP As Long = &HA2
Private Const WM_NCMOUSEMOVE As Long = &HA0
Private Const WM_NCPAINT As Long = &H85
Private Const WM_NCRBUTTONUP As Long = &HA5
Private Const WM_SETCURSOR As Long = &H20
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_SETICON As Long = &H80
Private Const WM_SETTEXT As Long = &HC
Private Const WM_SIZE = &H5
Private Const WM_STYLECHANGED = &H7D
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private Const WM_XBUTTONDOWN As Long = &H20B
Private Const WS_BORDER As Long = &H800000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_MAXIMIZE As Long = &H1000000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZE As Long = &H20000000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WVR_VALIDRECTS As Long = &H400


' API UDTs and Custom UDTs
Private Type LOGFONT        ' used to create fonts
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
    lfFaceName As String * 32
End Type
Private Type NONCLIENTMETRICS     ' used to retrieve/set system settings
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Private Type StyleStructure  ' used when window style changing or changed
    StyleOld As Long
    StyleNew As Long
End Type
Private Type NCCALCSIZE_PARAMS  ' used when client rect needs to be calculated
   rgrc(0 To 2) As RECT
   lppos As Long 'pointer to WINDOWPOS
End Type
Private Type MINMAXINFO         ' used when sizing a window
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Private Type WINDOWPOS          ' used when sizing or moving
   hWnd As Long
   hWndInsertAfter As Long
   X As Long
   Y As Long
   CX As Long
   CY As Long
   Flags As Long
End Type
Private Type MENUITEMINFO        ' used to retrieve/store menu items
     cbSize As Long              ' this structure is used with all O/S
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long 'String
     cch As Long
End Type
Private Type WindowIcon         ' Custom, used to track creation of 16x icons
    Handle As Long
    ToBeDeleted As Boolean
    hOriginalIcon As Long
    ToBeReplaced As Boolean
End Type
Private Type tbarIconsData      ' Custom, user-defined titlebar buttons
    tRgn As Long
    tDefault As Boolean
    tState As Integer   '0=up, 1=down, 2=disabled
    tPosition As Integer '0=default, 1=lockX, 2=lockY, 128=no frame
    CurXY As POINTAPI   ' most current X,Y for a non-static, no-default custom button
    Size As POINTAPI    ' height & width of a custom button
    toolTip As String * 255
    ID As String * 10
    ' more later
End Type
Private Type MOVESIZE       ' Custom, used to track core window data
    wRgn As Long
    wRect As RECT
    mPts As POINTAPI
    HitTest As Long
    Action As Long
    Tag As Long
End Type

' *** class if created must be properly terminated to destroy memory objects it creates
Private WithEvents c_MBar As clsMenuBarControl
Attribute c_MBar.VB_VarHelpID = -1

' *** These 3 must be destroyed to prevent memory leaks
Private m_DCuser As Long        ' DC passed to user for custom drawing
Private m_bmpTitlebar As Long   ' workarea passed to user for drawing entire titlebar region
Private m_bmpMenuBar As Long    ' workarea passed to user for drawing background menubar

' *** Handle in next UDT if .ToBeDeleted=TRue must be destroyed to prevent memory leaks
Private m_SmallIcon As WindowIcon   ' tracks handle for small icon & delete state
' *** Region in next UDT must be destroyed to prevent memory leaks
Private m_Position As MOVESIZE      ' tracks window size, hittest & current action

' *** Each region in UDT array on next line must be destroyed to prevent memory leaks
Private tBarIcon() As tbarIconsData ' collection of titlebar buttons (min/max/close, etc)
Private m_OriginalStyle As Long     ' used to restore original GWL_STYLE
Private m_OriginalStyleEX As Long   ' used to restore original GWL_EXSTYLE
Private m_CaptionSize As Integer    ' height of titlebar (not including menubar)
Private m_IconCX As Long            ' height of titlebar buttons
Private m_IconCY As Long            ' height of titlebar buttons
Private m_bLockWindow As Boolean    ' used to prevent sending style changes to subclassed window
Private m_InTransition As Boolean   ' used when minimizing/maximizing/restoring
Private m_Terminating As Boolean    ' flag to indicate subclassing terminating
Private m_bIsActive As Boolean      ' track if window is active per O/S
Private m_fColor(0 To 1) As Long    ' active/inactive font color
Private m_tbarFarExtent As Integer  ' caption right offset = left edge of nearest button
Private m_tbarNearExtent As Integer ' caption left offset = right edge of window icon
Private hOldWndProc As Long         ' subclassed window's previous GWL_WNDPROC
Private mainHwnd As Long            ' subclassed window handle

' ------------- PROPERTY VARIABLES -----------
Private m_Font As Long  ' title bar font (LOGFONT)
Private m_SizingBorder As Byte  ' width/height of invisible grab handles
Private m_BorderXY As RECT      ' width/height of 4 border edges (all can be different)
Private m_BorderStyle As Integer ' similar to VB's BorderStyle
Private m_tbarAlign As Integer  '0=horizonal,1=vertical,2=caption centered
Private m_bHideInactiveBtns As Boolean  ' hide disabled buttons vs showing as disabled
Private m_sysMenu As Integer    ' used to track user-opted disabling of system menu items
Private m_KeepActive As Integer     ' option to show as active vs inactive (same thread)
    '0=current thread;1=other thread; 2=option is active,4=thread only,8=all threads
' ------------- REGIONS/RECTS USED FOR HITTEST & PAINTING ------------
' *** These next 3 must be destroyed to prevent memory leaks
Private hRgn_Client As Long
Private hRgn_Title As Long
Private hClipRgn As Long
Private hRect_Menu As RECT

Friend Property Get MenuBar() As clsMenuBarControl
If c_MBar Is Nothing Then   ' create new menubar class if needed
    Set c_MBar = New clsMenuBarControl
End If
Set MenuBar = c_MBar
End Property

' -------------------------- CLASS ADJUSTBLE PROPERTIES

Friend Property Get sysMenuEnabled(mnuIndex As SysMenuItemConstants) As Boolean
If mnuIndex < 0 Or mnuIndex > smSysIcon Then Exit Property
' function returns whether or not a single system menu item is enabled
sysMenuEnabled = Not ((m_sysMenu Or mnuIndex) = m_sysMenu)
End Property
Friend Property Let sysMenuEnabled(mnuIndex As SysMenuItemConstants, isEnabled As Boolean)
' function sets multiple system menu items at once using the OR operand
Dim tIndex As Long
Dim lIndex As Long
Dim exStyle As Long, bToggleVisibility As Boolean

lIndex = mnuIndex ' use as a guide
Do Until lIndex = 0
    ' test to see which were sent to this routine
    If (lIndex Or smClose) = lIndex Then
        tIndex = smClose    ' add/remove close function
    ElseIf (lIndex Or smMaximize) = lIndex Then
        tIndex = smMaximize ' add/remove maximize function
    ElseIf (lIndex Or smMinimize) = lIndex Then
        tIndex = smMinimize ' add/remove minimize function
    ElseIf (lIndex Or smMove) = lIndex Then
        tIndex = smMove ' add/remove move function
    ElseIf (lIndex Or smSize) = lIndex Then
        tIndex = smSize ' add/remove sizing function
    ElseIf (lIndex Or smSysIcon) = lIndex Then
        tIndex = smSysIcon ' add/remove sizing function
    Else    ' stop when all have been done or invalid entry passed
        lIndex = 0
    End If
    If lIndex Then
        If isEnabled Then   ' enable items
            ' when enabled, we remove their flag from the friend variable
            If (m_sysMenu Or tIndex) = m_sysMenu Then
                m_sysMenu = m_sysMenu And Not tIndex
            End If
        Else    ' when disabled we add the flag
            m_sysMenu = m_sysMenu Or tIndex
        End If
        ' remove the tested item from the guide & continue on
        lIndex = lIndex And Not tIndex
    End If
Loop
' settings can be cached before window sublcassed
If mainHwnd Then    ' subclassed & apply now; otherwise apply later
    TweakSysMenu False      ' disable items on the system menu
    CalculateButtonRect     ' adjust button regions as needed
    DoDrawTitleBar m_bIsActive, False, "sysmenuenabled" ' redraw window as needed
End If
End Property

Friend Function SetHideDisabledBtns(bSet As Boolean, isHidden As Boolean) As Boolean
' returns whether disabled buttons are hidden or displayed
If bSet Then
    m_bHideInactiveBtns = isHidden
    If mainHwnd Then
        CalculateButtonRect   ' disable/hide button regions
        DoDrawTitleBar m_bIsActive, False, "hidedisabledbtns" ' redraw as needed
    End If
Else
    SetHideDisabledBtns = m_bHideInactiveBtns
End If
End Function

Friend Property Set Font(newFont As StdFont)
' creates a logical font from a passed standard font
Dim nFont As LOGFONT
If m_Font Then DeleteObject m_Font  ' delete previous font if needed
If newFont Is Nothing Then          ' use default menubar font
    Dim ncm As NONCLIENTMETRICS
    ncm.cbSize = Len(ncm)
    ' this will return the system menu font info
    SystemParametersInfo 41, 0, ncm, 0
    nFont = ncm.lfCaptionFont
    nFont.lfCharSet = 1             ' kinda needed for vertical fonts
    m_Font = CreateFontIndirect(nFont)  ' create the test font
    ' test the font to ensure it cnan be rotated; required for this class
    GetGDIObject m_Font, Len(nFont), nFont
    If nFont.lfCharSet = 0 Then ' nope, can't be rotated select one that can
        nFont.lfFaceName = "Tahoma" & String$(32, 0)
    End If
    DeleteObject m_Font             ' delete the test font
Else
    ' build logical font attributes from standard font
    With newFont
        nFont.lfFaceName = .Name & String$(32, 0)
        nFont.lfHeight = (.Size * -20) / Screen.TwipsPerPixelY
        nFont.lfItalic = Abs(.Italic)
        nFont.lfStrikeOut = Abs(.Strikethrough)
        nFont.lfUnderline = Abs(.Underline)
        nFont.lfWeight = Abs(.Bold) * 300 + 400
    End With
End If
nFont.lfCharSet = 1 ' needed for vertical fonts
If (m_tbarAlign Or 1) = m_tbarAlign Then
    ' vertical title bar; make 90 degrees
    nFont.lfOrientation = 900
    nFont.lfEscapement = 900
End If
m_Font = CreateFontIndirect(nFont)  ' now create the actual logical font
' redraw title bar
If mainHwnd Then
    MeasureCaption
    'DoDrawTitleBar m_bIsActive, False, "font"
End If
End Property

Friend Property Get Font() As StdFont
' routine returns a standard font from a logical font
Dim tFont As StdFont, nFont As LOGFONT

' if we don't have a font yet, pass the standard menu font
If m_Font = 0 Then Set Font = Nothing

' begin building the font attributes
GetGDIObject m_Font, Len(nFont), nFont
Set tFont = New StdFont
With tFont
    If InStr(nFont.lfFaceName, Chr$(0)) Then
        .Name = Left$(nFont.lfFaceName, InStr(nFont.lfFaceName, Chr$(0)) - 1)
    Else
        .Name = nFont.lfFaceName
    End If
    .Bold = nFont.lfWeight > 400
    .Italic = nFont.lfItalic <> 0
    .Underline = nFont.lfUnderline <> 0
    .Strikethrough = nFont.lfStrikeOut <> 0
    If nFont.lfHeight < 0 Then
        .Size = (nFont.lfHeight * Screen.TwipsPerPixelY) / -20
    Else
        .Size = 8.25
    End If
End With
On Error Resume Next ' just in case user call property without Set keyword
Set Font = tFont
End Property

Friend Function SetCenterCaption(bSet As Boolean, isCentered As Boolean) As Boolean
' optional centering of titlebar caption
If bSet Then
    m_tbarAlign = m_tbarAlign Or 2
    If Not isCentered Then m_tbarAlign = m_tbarAlign And Not 2
    ' redraw titlebar as needed
    If mainHwnd Then DoDrawTitleBar m_bIsActive, False, "titlecentered"
Else
    SetCenterCaption = ((m_tbarAlign Or 2) = m_tbarAlign)
End If
End Function
Friend Function SetVerticalCaption(bSet As Boolean, isVertical As Boolean) As Boolean
' optional vertical title bar (basic window frame only; not custom)
If bSet Then
    m_tbarAlign = m_tbarAlign Or 1
    If Not isVertical Then m_tbarAlign = m_tbarAlign And Not 1
    m_bLockWindow = True    ' prevent next line from triggering repaint
    Set Font = Font         ' toggle vertical/horizontal
    m_bLockWindow = False   ' allow repainting/resizing
    If mainHwnd Then
        SetRect m_Position.wRect, 0, 0, -1, -1      ' force complete recalc
        AdjustImplementSize False
        'UpdateWindowStyles False, False, -3, 0  ' trigger resize/repaint
    End If
Else
    SetVerticalCaption = ((m_tbarAlign Or 1) = m_tbarAlign)
End If
End Function

Friend Function SetInTaskBar(bSet As Boolean, AddToTaskBar As Boolean) As Boolean
' place window in taskbar or remove from taskbar
Dim exStyle As Long
If bSet Then
    ' place window in taskbar or remove from taskbar
    Dim bToggleVisibility
    If mainHwnd Then    ' not subclassed yet
        exStyle = GetWindowLong(mainHwnd, GWL_EXSTYLE)
        bToggleVisibility = ((GetWindowLong(mainHwnd, GWL_STYLE) And WS_VISIBLE) = WS_VISIBLE)
    End If
    If AddToTaskBar Then 'put in taskbar
        If mainHwnd Then    ' subclassed, so let's do it
            If (exStyle Or WS_EX_APPWINDOW) <> exStyle Then ' not already there
                If bToggleVisibility Then ShowWindow mainHwnd, 0
                SetWindowLong mainHwnd, GWL_EXSTYLE, exStyle Or WS_EX_APPWINDOW
                SetSmallIcon
                If bToggleVisibility Then SetWindowPos mainHwnd, 0, 0, 0, 0, 0, &H40 Or 551
            End If
        Else    ' not subclassed yet, let's flag a variable to do this later
            m_OriginalStyleEX = 2
        End If
    Else
        If mainHwnd Then    ' subclassed & doable
            If (exStyle Or WS_EX_APPWINDOW) = exStyle Then
                If bToggleVisibility Then ShowWindow mainHwnd, 0
                SetWindowLong mainHwnd, GWL_EXSTYLE, exStyle Or WS_EX_APPWINDOW
                SetWindowLong mainHwnd, GWL_EXSTYLE, exStyle And Not WS_EX_APPWINDOW
                If bToggleVisibility Then SetWindowPos mainHwnd, 0, 0, 0, 0, 0, &H40 Or 551
            End If
        Else    ' not subclassed yet, let's flag a variable to do this later
            m_OriginalStyleEX = 1
        End If
    End If
Else
    ' quickie -- see if window is a taskbar item or not
    If mainHwnd Then
        SetInTaskBar = ((GetWindowLong(mainHwnd, GWL_EXSTYLE) And WS_EX_APPWINDOW) = WS_EX_APPWINDOW)
    End If
End If
End Function

Public Function SetSizingHandle(bSet As Boolean, newSize As Byte) As Byte
' set optional size. Making this larger allows a greater chunk of the
' window edges to be marked for sizing windows
If bSet Then
    If newSize > 0 Then m_SizingBorder = newSize
Else
    SetSizingHandle = m_SizingBorder
End If
End Function


Public Function SetBorderStyle(bSet As Boolean, newStyle As WindowBorderStyleConstants) As WindowBorderStyleConstants
If bSet Then
    If newStyle < wbBlackEdge Or newStyle > wbCustom Or m_BorderStyle = newStyle + 0 Then Exit Function
    ' simply apply a standard-type border around a basic window
    Dim edgeThickness As Long
    Select Case newStyle
    Case 1: ' border only
        edgeThickness = 1
    Case 2: ' thin
        edgeThickness = GetSystemMetrics(SM_CXDLGFRAME) - 1
    Case 3: ' dialog
        edgeThickness = GetSystemMetrics(SM_CXDLGFRAME)
    Case 4: ' thick frame
        edgeThickness = GetSystemMetrics(SM_CXDLGFRAME) + 1
    Case Else   ' custom
        edgeThickness = 8
        ' todo: used for skinning forms
'        Exit Function
    End Select
    m_BorderStyle = newStyle
    SetRect m_BorderXY, edgeThickness, edgeThickness, edgeThickness, edgeThickness
    SetRect m_Position.wRect, 0, 0, -1, -1
    UpdateWindowStyles False, False, -3, 0
Else
    SetBorderStyle = m_BorderStyle + 0
End If
End Function

Public Function SetFontColor(bSet As Boolean, bActive As Boolean, newColor As Long) As Long
If bSet Then
    m_fColor(Abs(bActive = False)) = newColor
    DoDrawTitleBar True, False, "setfontcolor"
Else
    SetFontColor = m_fColor(Abs(bActive = False))
End If
End Function

Public Function SetKeepActive(bSet As Boolean, AlwaysActive As Boolean, AllThreads As Boolean) As Boolean
' property allows a window to not be shown as inactive in various scenarios
' if alwaysActive is false, then window only shown when it has the focus
If bSet Then
    ' when m_KeepActiveAlways is true, it never is displayed as inactive;
    ' however when set to false, it will remain active when any window in
    ' the same thread gets the focus. This option is kinda neat for message
    ' boxes that make the calling form inactive
    If AlwaysActive Then
        m_KeepActive = 2
        m_KeepActive = m_KeepActive Or (4 + Abs(4 * AllThreads))
    Else
        m_KeepActive = 0
    End If
    ' just in case it was changed from a timer or outside source....
    If mainHwnd Then DoDrawTitleBar m_bIsActive, True, "keepactive"
Else
    If AllThreads Then
        SetKeepActive = (m_KeepActive Or 8) = m_KeepActive
    Else
        SetKeepActive = (m_KeepActive Or 4) = m_KeepActive
    End If
End If
End Function

Friend Sub SetTitle(sCaption As String, ByVal hWnd As Long)
If mainHwnd Then hWnd = mainHwnd
If hWnd <> 0 Then SendMessage hWnd, WM_SETTEXT, 0&, ByVal sCaption
End Sub

Friend Sub AddToolBarButton(ByVal btnID As String, ByVal newPos As Long, _
        Optional PosType As TitlelBarBtnPosition, _
        Optional X As Long, Optional Y As Long, _
        Optional Width As Long, Optional Height As Long)

' NewPos has these meanings
' = 0 : add button before minimize button
'       - X/Y/Width/Height not used if PosType = tbPosDefault
' > 0 : add button # of items left of minimize button.
'       - X/Y/Width/Height not used if PosType = tbPosDefault
' < 0 : add button at far left from minimize button
'       - X/Y/Width/Height not used if PosType = tbPosDefault
' PosType is used this way & is for custom placement of buttons....
'   PosType = tbPosDefault : button positioned before minimize button using NewPos
'   PosType = tbPosStatic : the location does not change ever regardless of size/movement
'       - this setting overrides tbPosLockX and/or tbPosLockY if used
'       - X,Y are required, Width/Height are optional
'   PosType = tbPosLockX : location changes with reference to right edge of window
'       - X,Y are required, Width/Height are optional
'   PosType = tbPosLockY : location changes with reference to bottom edge of window
'       - X,Y are required, Width/Height are optional
'   PosType = tbNoFrame : if not used, a button will be drawn
' X & Y : are the top left coordinates of the custom placement
' Width & Height are optional dimensions of the custom button.
'       - If either one is zero, then dimension is calcualted by the project
' ID is a 10 character-max user-defined string value to help reference the
'   button in callback routines and to delete/edit the button

If btnID = "" Then Exit Sub

Dim I As Integer, J As Integer
Dim inID As String * 11
    
inID = Left$(btnID, 10)
ReDim Preserve tBarIcon(0 To UBound(tBarIcon) + 1)
If PosType = tbPosDefault Then
    If newPos = 0 Then  ' add as 1st button left of minimize
        newPos = 4
    ElseIf newPos > UBound(tBarIcon) - 5 Or newPos < 0 Then
        newPos = UBound(tBarIcon)   ' add as furthest left of minimize
    Else    ' insert somewhere between
        newPos = newPos + 4
        For I = UBound(tBarIcon) To newPos + 1 Step -1
            tBarIcon(I) = tBarIcon(I - 1)
        Next
    End If
Else    ' custom placement, the actual position doesn't matter
    newPos = UBound(tBarIcon)
End If
With tBarIcon(newPos)
    .ID = inID
    .tDefault = False
    .tPosition = PosType
    .CurXY.X = X
    .CurXY.Y = Y
    .Size.X = Width
    .Size.Y = Height
    .tRgn = 0
End With
CalculateButtonRect
DoDrawTitleBar m_bIsActive, False, "addtoolbarbutton"
End Sub

Friend Sub DeleteToolBarButton(ByVal btnID As String)

If btnID = "" Then Exit Sub
Dim I As Integer, J As Integer, inID As String * 10
inID = btnID
For I = 4 To UBound(tBarIcon)
    If tBarIcon(I).ID = inID Then ' found it
        For J = I + 1 To UBound(tBarIcon) - 1
            tBarIcon(J) = tBarIcon(J + 1)
        Next
        ReDim Preserve tBarIcon(0 To UBound(tBarIcon) - 1)
    End If
    Exit For
Next
End Sub


Private Function ConvertHitTest(wParam As Long, inMenuLoop As Boolean) As Long

' function returns the hittest for our custom window
' wParam is the overriding return value if we do not get a valid hittest
Debug.Assert m_Position.wRgn <> 0

Dim mPts As POINTAPI, cRect As RECT
Dim I As Integer, rtnVal As Long, bGotHit As Boolean

GetCursorPos mPts   ' always use current X,Y coords
rtnVal = HTNOWHERE  ' default return value


' test to see if pt is even in our window....
If PtInRegion(m_Position.wRgn, mPts.X, mPts.Y) Then
    ' test the menubar
    If hRect_Menu.Bottom > hRect_Menu.Top Then
        ' check the menu
        If PtInRect(hRect_Menu, mPts.X, mPts.Y) Then
            rtnVal = c_MBar.GetHitTest(mPts.X - hRect_Menu.Left, mPts.Y - hRect_Menu.Top)
            If rtnVal Then
                m_Position.Tag = rtnVal
                rtnVal = HTMenuPlus
            Else
                rtnVal = HTCAPTION
            End If
        End If
    End If
    If rtnVal = HTNOWHERE Then
        ' not on menubar, let's look at the titlebar
        If PtInRegion(hRgn_Title, mPts.X, mPts.Y) Then
            rtnVal = HTCAPTION
            ' in the title bar area, check for hits on the title bar buttons
            For I = 0 To UBound(tBarIcon)
                With tBarIcon(I)
                    If .tRgn Then
                        If PtInRegion(.tRgn, mPts.X, mPts.Y) Then
                            If I < 4 Then ' min/max/close/sysmenu
                                If sysMenuEnabled(Choose(I + 1, smSysIcon, smClose, smMaximize, smMinimize)) Then
                                    rtnVal = Choose(I + 1, HTSYSMENU, HTCLOSE, HTMAXBUTTON, HTMINBUTTON)
                                Else
                                    rtnVal = -1
                                End If
                            Else    ' custom buttons
                                ''Debug.Print "hit test on custom button "; I - 4
                                rtnVal = HTMenuPlus + I
                            End If
                            Exit For
                        End If
                    End If
                End With
            Next
        End If
    End If
    If rtnVal = HTNOWHERE Then ' not on titlebar, see if in client area
' added functionality. Since windows with thin borders don't have much of a
' sizing edge, the optional property SizingHandle can be adjusted to give
' the user a much larger, user-defined, edge for sizing purposes...
' When this option is used and the window is sizable, we increase the
' actual sizing edge.
        GetRgnBox hRgn_Client, cRect
        If sysMenuEnabled(smSize) Then   'sizable
            ' if window is not sizable, no need tweaking hotspots
            With cRect
                If m_SizingBorder > m_BorderXY.Right Then .Right = .Right - m_SizingBorder - m_BorderXY.Right
                If m_SizingBorder > m_BorderXY.Bottom Then .Bottom = .Bottom - m_SizingBorder - m_BorderXY.Bottom
                If (m_tbarAlign Or 1) = m_tbarAlign Then ' vertical titlebar
                    If m_SizingBorder > m_BorderXY.Top Then .Top = .Top + m_SizingBorder - m_BorderXY.Top
                Else
                    If m_SizingBorder > m_BorderXY.Left Then .Left = .Left + m_SizingBorder - m_BorderXY.Left
                End If
            End With
        End If
        If PtInRect(cRect, mPts.X, mPts.Y) Then ' test for client region
            rtnVal = HTCLIENT
        ' not inside client region; let's look at borders
        ElseIf sysMenuEnabled(smSize) And IsZoomed(mainHwnd) = 0 Then ' if sizable, then continue
            ' check for hits on the borders
            If PtInRegion(m_Position.wRgn, mPts.X - m_SizingBorder, mPts.Y) = 0 Then
                ' on the left side of the window
                rtnVal = HTLEFT
            ElseIf PtInRegion(m_Position.wRgn, mPts.X + m_SizingBorder, mPts.Y) = 0 Then
                ' on the right side of the window
                rtnVal = HTRIGHT
            End If
            If PtInRegion(m_Position.wRgn, mPts.X, mPts.Y - m_SizingBorder) = 0 Then
                ' on the top side of the window, include topright & topleft
                rtnVal = HTTOP + ((rtnVal > 0) * (HTMAXBUTTON - rtnVal))
            ElseIf PtInRegion(m_Position.wRgn, mPts.X, mPts.Y + m_SizingBorder) = 0 Then
                ' on the bottom side of window, include botright & botleft
                rtnVal = HTBOTTOM + ((rtnVal > 0) * (HTMAXBUTTON - rtnVal))
            End If
        End If
    Else
        If rtnVal < 0 Then rtnVal = HTNOWHERE
    End If
End If
''Debug.Print "converted hittest to "; wParam; rtnVal
' return overriding value if no return value was calculated
If inMenuLoop Then
    If rtnVal = HTMenuPlus Then
        m_Position.HitTest = rtnVal
    Else
        m_Position.HitTest = 0
    End If
    rtnVal = HTNOWHERE
End If
ConvertHitTest = rtnVal
End Function

Private Function SetHitTestAction(ByVal wMsg As Long, ByVal wParam As Long) As Long
' sets the custom movesize structure
' TODO: need to code for ht_buttons & sysmenu & HTBORDER (unsizable)
' TOOD: if used for MDI Child, the next function isn't screen orientated

GetCursorPos m_Position.mPts
Dim btnID As Integer

Select Case wParam
    Case HTNOWHERE, HTCLIENT, HTBORDER
        m_Position.Action = 0 ' no action
    Case HTLEFT To HTBOTTOMRIGHT    ' borders
        If sysMenuEnabled(smSize) Then
            m_Position.Action = 2
            UpdateUserResize True
        Else
            m_Position.Action = -1
        End If
    Case HTCAPTION
        If IsZoomed(mainHwnd) = 0 And sysMenuEnabled(smMove) Then
            m_Position.Action = 1
        Else
            m_Position.Action = -1
        End If
    Case HTMenuPlus ' menu
        c_MBar.TrackMenuBar m_Position.Tag - 1, True, (tBarIcon(0).tRgn <> 0), True, , , , -2
        m_Position.Action = 0
    Case HTMINBUTTON, HTMAXBUTTON, HTCLOSE
        btnID = GetButtonRef(wParam, True)
        If sysMenuEnabled(Choose(btnID, smClose, smMaximize, smMinimize)) Then
            m_Position.Action = wParam
            tBarIcon(btnID).tState = IIf(wMsg = WM_NCMOUSEMOVE, 1, 2)
            DrawBarIcons 0, btnID, IIf(wMsg = WM_NCMOUSEMOVE, bsHover, bsDown)
'            tBarIcon(btnID).tState = 2
'            DrawBarIcons 0, btnID, bsDown
        Else
            m_Position.Action = -1
            tBarIcon(btnID).tState = 3
            DrawBarIcons 0, btnID, bsDisabled
        End If
    Case Is > HTMenuPlus ' custom button
        m_Position.Action = wParam
        tBarIcon(wParam - HTMenuPlus).tState = 1
        DrawBarIcons 0, wParam - HTMenuPlus, bsDown
    Case HTSYSMENU
        If sysMenuEnabled(smSysIcon) Then
            m_Position.Tag = -GetTickCount()
            m_Position.Action = 0
        Else
            m_Position.Action = -1
        End If
    Case Else
        m_Position.Action = -1
End Select
If m_Position.Action = 1 Or m_Position.Action = 2 Then ' either moving or resizing
    Dim wRect As RECT
    ' determine clipping area for the desktop excluding toolbars
    SystemParametersInfo SPI_GETWORKAREA, 0, wRect, 0
    ClipCursor wRect    ' set clipping area for cursor
    SetCapture mainHwnd     ' capture the cursor
End If
m_Position.HitTest = wParam               ' cache the hittest code & action
If m_Position.Action < 0 Then
    m_Position.Action = HTNOWHERE
    SetHitTestAction = m_Position.Action
Else
    SetHitTestAction = wParam
End If
End Function

Private Function SetNewWindowPos(wParam As Long) As Long

Dim mPt As POINTAPI, newRect As RECT, bRecAdjusted As Boolean, bValidSize As Boolean

GetCursorPos mPt            ' get current mouse points

Select Case m_Position.Action

    Case 1  ' moving
    
    ' simply offset the top/left coords vs recalculating all edges & regions
    Dim xOffset As Long, yOffset As Long
    xOffset = (mPt.X - m_Position.mPts.X)
    yOffset = (mPt.Y - m_Position.mPts.Y)
    m_Position.mPts = mPt
        OffsetRect m_Position.wRect, xOffset, yOffset
        OffsetRgn m_Position.wRgn, xOffset, yOffset
        OffsetRgn hRgn_Client, xOffset, yOffset
        OffsetRgn hRgn_Title, xOffset, yOffset
        OffsetRect hRect_Menu, xOffset, yOffset
        CalculateButtonRect
    bRecAdjusted = True
    
    Case 2  ' resizing
    
    newRect = m_Position.wRect  ' start window rect with last window rect
    ' left & right adjustments
    bValidSize = True
    Select Case m_Position.HitTest
    Case HTRIGHT, HTTOPRIGHT, HTBOTTOMRIGHT ' all require right edge adjusting
        newRect.Right = m_Position.wRect.Right + (mPt.X - m_Position.mPts.X)
    Case HTLEFT, HTTOPLEFT, HTBOTTOMLEFT    ' all require left edge adjusting
        newRect.Left = m_Position.wRect.Left + (mPt.X - m_Position.mPts.X)
    Case Else
        bValidSize = False
    End Select
    If bValidSize Then
        ' test to see if resized window within minimum sizes per sys settings (X)
        If (m_tbarAlign Or 1) = m_tbarAlign Then
            bValidSize = GetSystemMetrics(SM_CYMINTRACK) <= newRect.Right - newRect.Left
        Else
            bValidSize = GetSystemMetrics(SM_CXMINTRACK) <= newRect.Right - newRect.Left
        End If
        If bValidSize Then
            m_Position.wRect.Right = newRect.Right
            m_Position.wRect.Left = newRect.Left
            m_Position.mPts.X = mPt.X
            bRecAdjusted = True
        End If
    End If
    ' top & bottom adjustments
    bValidSize = True
        Select Case m_Position.HitTest
        Case HTTOP, HTTOPLEFT, HTTOPRIGHT       ' all require top edge adjusting
            newRect.Top = m_Position.wRect.Top + (mPt.Y - m_Position.mPts.Y)
        Case HTBOTTOM, HTBOTTOMLEFT, HTBOTTOMRIGHT  ' all require bottom edge adjusting
            newRect.Bottom = m_Position.wRect.Bottom + (mPt.Y - m_Position.mPts.Y)
        Case Else
            bValidSize = False
        End Select
        If bValidSize Then
            ' test to see if resized window within minimum sizes per sys settings (Y)
            If (m_tbarAlign Or 1) = m_tbarAlign Then
                bValidSize = GetSystemMetrics(SM_CXMINTRACK) <= newRect.Bottom - newRect.Top
            Else
                bValidSize = GetSystemMetrics(SM_CYMINTRACK) <= newRect.Bottom - newRect.Top
            End If
            If bValidSize Then
                m_Position.wRect.Bottom = newRect.Bottom
                m_Position.wRect.Top = newRect.Top
                m_Position.mPts.Y = mPt.Y
                bRecAdjusted = True
            End If
        End If
            
    Case Else
        
'        'Debug.Print m_Position.Action; m_Position.hitTest
End Select
' set the new window position
If bRecAdjusted Then
    ' inflate region, but for now
'    If m_Position.wRgn Then DeleteObject m_Position.wRgn
'    m_Position.wRgn = CreateRectRgnIndirect(wRect)
'    If hRgn_Client Then DeleteObject hRgn_Client
'    hRgn_Client = CreateRectRgn(wRect.Left + m_SizingBorder, _
'        wRect.Right - m_SizingBorder, wRect.Top + m_SizingBorder + 16, _
'        wRect.Bottom - m_SizingBorder)
    
    
    ' apply new region
    
    'With m_Position.wRect
    SetWindowPos mainHwnd, 0, m_Position.wRect.Left, m_Position.wRect.Top, m_Position.wRect.Right - m_Position.wRect.Left, m_Position.wRect.Bottom - m_Position.wRect.Top, 0 '&H400
    'End With
    SetNewWindowPos = 1
End If
End Function

Friend Sub BeginSubclass(hWnd As Long)
If mainHwnd Then
    EndSubclass
    Set Font = Nothing
End If
' stil playing with this & may go to the wayside later on...

' now start preparing for subclassing
mainHwnd = hWnd   ' get hwnd reference & set ref to this class
SetSmallIcon
' determine if window is active/inactive: todo: disabled status
m_bIsActive = (GetForegroundWindow() = mainHwnd)
MenuBar.ParentWIndow = mainHwnd ' create a menubar as needed
ReDim tBarIcon(0 To 3) ' min/max/close/sysicon
SetRect m_Position.wRect, 0, 0, -1, -1 ' reset to invalid value used in next line
CalculateInitialStyles  ' calculate new window areas & modify passed window

' begin subclassing & drawing
m_Terminating = False

hOldWndProc = SetWindowLong(mainHwnd, GWL_WNDPROC, GetSubClassAddr())
TweakSysMenu False   ' disable menu items per user-defined options
c_MBar.WindowMenu("beginclass") = 0
UpdateWindowStyles False, False, -3, 0  ' force a complete redraw
End Sub

Friend Sub EndSubclass()
Dim I As Integer, bToggleVisibility As Boolean

'Debug.Print ">>>>>>>>> received END SUBCLASS <<<<<<<<<<"; mainHwnd

If mainHwnd Then    ' clean up here
    ' remove pointer to user-defined drawing routines
    RemoveProp mainHwnd, "lvImpCB_Ptr"
    ' unsubclass the window
    SetWindowLong mainHwnd, GWL_WNDPROC, hOldWndProc
    ' remove pointer to this class for subclassing
    RemoveProp mainHwnd, "lvCFrame_Optr"
    ' return the menubar if needed
    If Not c_MBar Is Nothing Then Set c_MBar = Nothing
    ' add these flags if window is now maximized or minimized
    If IsZoomed(mainHwnd) Then
        m_OriginalStyle = m_OriginalStyle Or WS_MAXIMIZE
    ElseIf IsIconic(mainHwnd) Then
        m_OriginalStyle = m_OriginalStyle Or WS_MINIMIZE
    End If
    ' remove item from taskbar if originally not on taskbar
    If (GetWindowLong(mainHwnd, GWL_STYLE) And WS_VISIBLE) = WS_VISIBLE Then
        If GetWindowLong(mainHwnd, GWL_EXSTYLE) <> m_OriginalStyleEX Then
             ShowWindow mainHwnd, 0 ' to remove item from taskbar if needed
        End If
    End If
    ' update/replace appropriate Style & ExStyle properties
    SetWindowLong mainHwnd, GWL_EXSTYLE, m_OriginalStyleEX
    UpdateWindowStyles False, False, -2, m_OriginalStyle
    ' return the original small icon if replaced
    If m_SmallIcon.ToBeReplaced Then SendMessage mainHwnd, WM_SETICON, 0, ByVal m_SmallIcon.hOriginalIcon
    DrawMenuBar mainHwnd    ' update the menubar
    AdjustImplementSize True    ' resize window to account for removal of custom settings
    TweakSysMenu True       ' re-enable all system menu items
    For I = 0 To UBound(tBarIcon)   ' delete regions
        If tBarIcon(I).tRgn Then DeleteObject tBarIcon(I).tRgn
    Next
End If
' continue with clean up
If m_SmallIcon.ToBeDeleted Then DestroyIcon m_SmallIcon.Handle
If m_DCuser Then DeleteDC m_DCuser
If m_bmpMenuBar Then DeleteObject m_bmpMenuBar
If m_bmpTitlebar Then DeleteObject m_bmpTitlebar
If hRgn_Client Then DeleteObject hRgn_Client
If hRgn_Title Then DeleteObject hRgn_Title
If m_Position.wRgn Then DeleteObject m_Position.wRgn
If hClipRgn Then DeleteObject hClipRgn
If m_Font Then DeleteObject m_Font
m_Font = 0
m_DCuser = 0
m_bmpMenuBar = 0
m_bmpTitlebar = 0
hRgn_Client = 0
hRgn_Title = 0
hClipRgn = 0
m_Position.wRgn = 0
m_SmallIcon.ToBeDeleted = False
m_SmallIcon.ToBeReplaced = False
m_SmallIcon.Handle = 0
m_KeepActive = 0
m_OriginalStyle = 0
m_OriginalStyleEX = 0
mainHwnd = 0
m_sysMenu = 0
SetRect hRect_Menu, 0, 0, 0, 0
Erase tBarIcon()
m_fColor(0) = vbActiveTitleBarText
m_fColor(1) = vbInactiveTitleBarText
End Sub

Friend Function NewWndProc(wMsg As Long, wParam As Long, lParam As Long) As Long

Dim SS As StyleStructure        ' used to test applied window styles
Dim bWasIntercepted As Boolean  ' flag prevents forwarding message on
Dim lRtnVal As Long             ' general use variable
Dim lSysMenuMsg As Long         ' used to identify which msg is calling sysmenu
Dim IsMinimized As Boolean      ' is window iconic
Static p_LockUpdate As Boolean  ' internal for SetCursor hack
Dim callBack As CustomWindowCalls

''Debug.Print mainHwnd, wMsg, wParam; lParam

If Not m_Terminating Then

    IsMinimized = (IsIconic(mainHwnd) <> 0)
    
    Select Case wMsg
    ' these messages to be processed whether maximized/minimized or normal
    
    Case WM_DESTROY ' stop subclassing if parent is being destroyed
        m_sysMenu = -1  ' flag to prevent restoring form's original styles
        RemoveProp mainHwnd, "lvImpCB_Ptr"
        NewWndProc = CallWindowProc(hOldWndProc, mainHwnd, wMsg, wParam, lParam)
        m_Terminating = True
        bWasIntercepted = True  ' don't forward this message
        
    Case WM_APPACTIVATE
        If wParam = 0 Then
            m_KeepActive = m_KeepActive Or 1
            m_bIsActive = False
        Else
            m_KeepActive = m_KeepActive And Not 1
        End If
        If m_KeepActive > 1 Then
            DoDrawTitleBar m_bIsActive, True, "wm_appactivate"
        End If
        
    Case WM_GETSYSMENU  ' either called from outside application or click on sysIcon
        lSysMenuMsg = wMsg  ' identify which message is displaying the sysmenu
    
    Case WM_SYSCOMMAND, WM_COMMAND
        ' the system menu comes via wm_syscommand but if shown
        ' with wm_getsysmenu, then it comes via wm_command
            'Debug.Print wParam; lParam; c_MBar.WindowMenu(""); GetSystemMenu(mainHwnd, 0); GetMenu(mainHwnd); GetWindowLong(mainHwnd, GWL_STYLE); GetWindowLong(mainHwnd, GWL_EXSTYLE)
        Select Case wParam
        Case SC_KEYMENU ' check for Alt+Space
            
            
            If lParam = vbKeySpace Then
                If IsMinimized Then
                    ' when minimized without being in taskbar, this is fired when user
                    ' clicks sysIcon or iconic titlebar
                    lSysMenuMsg = wMsg
                Else    ' alt+space, need to track separately 'cause if the
                    ' sysIcon is not in top left corner, menu won't show!
                    PostMessage mainHwnd, WM_GETSYSMENU, SC_KEYMENU, ByVal 0&
                    bWasIntercepted = True
                End If
            ElseIf IsMinimized = False Then
                If lParam < 1 Then
                    ' highlight first menubar item if any
                    SendMessage mainHwnd, WM_SETCURSOR, HTCAPTION, 0&
                    c_MBar.TrackMenuBar Abs(lParam), False, (tBarIcon(0).tRgn <> 0), True, True
                Else
                    c_MBar.TraceHotKey lParam, (tBarIcon(0).tRgn <> 0), True
                End If
                bWasIntercepted = True  ' don't forward this message
            End If
        Case SC_CLOSE, SC_MOVE, SC_SIZE
            ' key system menu items. Call routine to ensure a command isn't
            ' being exectuted that is disabled :: outside code
            If ProcessSysMenuItem(wParam, False) = False Then
                ' ok, send message along & modify system menu if needed
                wMsg = WM_SYSCOMMAND    ' change any wm_command to wm_syscommand
                CallWindowProc hOldWndProc, mainHwnd, wMsg, wParam, lParam
                c_MBar.WindowMenu("sc_close,move,size") = 0
            End If
            bWasIntercepted = True  ' don't forward this message
        Case SC_MINIMIZE, SC_RESTORE
            If ProcessSysMenuItem(wParam, False) = False Then
                UpdateWindowStyles False, False, 0, WS_THICKFRAME, WS_DLGFRAME
                m_InTransition = True   ' see wm_gettext,wm_geticon
                wMsg = WM_SYSCOMMAND    ' change any wm_command to wm_syscommand
                CallWindowProc hOldWndProc, mainHwnd, wMsg, wParam, lParam
                TweakSysMenu False
                c_MBar.WindowMenu("sc_minimize,restore") = 0
            End If
            bWasIntercepted = True  ' don't forward this message
        Case SC_MAXIMIZE
            If ProcessSysMenuItem(wParam, False) = False Then
                UpdateWindowStyles False, False, -2, AddRqdStyles(GWL_STYLE, 0, True)
                m_InTransition = True
'                If m_BorderStyle = -1 Then UpdateWindowStyles True, False, 0, WS_CAPTION
                wMsg = WM_SYSCOMMAND
                CallWindowProc hOldWndProc, mainHwnd, wMsg, wParam, lParam
                TweakSysMenu False
                c_MBar.WindowMenu("sc_maximize") = 0
'                If m_BorderStyle = -1 Then UpdateWindowStyles True, True, 0, WS_CAPTION
            End If
            bWasIntercepted = True
        Case SC_MOUSEMENU, SC_MOUSEMENU + 3 ' << a valid variation of sc_mousemenu
            If IsMinimized Then
                lSysMenuMsg = wMsg  ' identify message that is triggering sysmenu
            Else    ' track separately, see SC_KEYMENU remarks above
                PostMessage mainHwnd, WM_GETSYSMENU, SC_KEYMENU, ByVal 0&
                bWasIntercepted = True
            End If
        Case Else   ' todo: finish up; i.e., SC_NEXT for MDI children
            If wMsg = WM_COMMAND And HiWord(wParam) = 0 And lParam = 0 Then ' menu item
                'Debug.Print "got a menu item "; wParam
'                bWasIntercepted = True
            Else
                'If wParam = WM_SYSCOMMAND Then 'Debug.Print "Unknown wparam "; wParam; lParam
                'Debug.Print "Unknown wparam "; wMsg; wParam; lParam, WM_COMMAND
            End If
        End Select
        
    Case WM_GETICON, WM_GETTEXT
        ' when maximizing/minimizing or restoring from these states,
        ' VB draws the titlebar via USER+xxx messages. Not worth the
        ' hassle to trap for those since they may change. Instead,
        ' we will try to prevent the text & icon from being drawn
        ' while window is in transition & just allow VB to draw
        ' a blank, blue titlebar
        bWasIntercepted = m_InTransition
        
    Case WM_SETICON, WM_SETTEXT ' VB does not send SetText when Me.Caption is used
        ' but other apps can... remove this style to prevent redraw
        UpdateWindowStyles True, True, 0, WS_CAPTION
        CallWindowProc hOldWndProc, mainHwnd, wMsg, wParam, lParam
        ' add style back & redraw caption
        UpdateWindowStyles True, False, 0, WS_CAPTION
        bWasIntercepted = True
        If wMsg = WM_SETICON Then
            If wParam = 1 Then SetSmallIcon
        End If
        DoDrawTitleBar m_bIsActive, False, "wmseticon/wmsettext"
    
    Case WM_NCCALCSIZE
        ' see if we are calculating the non-client area
        If Not IsMinimized Then
            If p_LockUpdate = False And m_bLockWindow = False Then
            If wParam Then
                DefWindowProc mainHwnd, wMsg, wParam, ByVal lParam
                ConfigureNCandClient lParam
                NewWndProc = WVR_VALIDRECTS  ' special meaning
                bWasIntercepted = True  ' don't forward this message
            End If
            End If
        End If
        m_InTransition = False
    
    Case &H1A   ' still testing "what-ifs" here; i.e., system menu font changes
        'Debug.Print "got overall system change"
        
       
    Case Else
    
        ' messages to be processed only if not minimized
    
        If Not IsMinimized Then
        
        Select Case wMsg
        
        Case WM_WINDOWPOSCHANGED
            ' take into credit desktop toolbars moving around while this is maximized
'            SetRect m_Position.wRect, -1, -1, 0, 0
'            CalculateBarRects True
            c_MBar.DrawMenuBarItems 1, 0, hRect_Menu, m_Position.wRect, -2, 0 ' refresh
            
        Case WM_STYLECHANGED
    '        'Debug.Print "got style change? "; wParam, GetWindowLong(mainHwnd, wParam)
            If m_bLockWindow Or p_LockUpdate Then
                ' self-imposed style changes
                If p_LockUpdate Then p_LockUpdate = False
                If m_bLockWindow Then m_bLockWindow = False
                bWasIntercepted = True
            Else    ' not expected or fired from elsewhere
                If wParam = GWL_STYLE Then
                    ' class requires the GWL_STYLE to remain fairly static
                    ' VB screws up when the original window was borderless and
                    ' additional styles were applied. It tends to revert back
                    ' to a style of borderless or near borderless if a caption
                    ' or sysIcon were modified... Fix that here
                    CopyMemory SS, ByVal lParam, Len(SS)
                    lRtnVal = SS.StyleNew
                    AddRqdStyles wParam, SS.StyleNew, IsZoomed(mainHwnd)
                    If lRtnVal = SS.StyleNew Then
                        CallWindowProc hOldWndProc, mainHwnd, wMsg, wParam, lParam
                    Else    ' force window to accept our style + the new style
                        UpdateWindowStyles False, False, -2, SS.StyleNew
                    End If
                Else    ' let it go for GWL_EXSTYLE; no known conflicts yet
                    CallWindowProc hOldWndProc, mainHwnd, wMsg, wParam, lParam
                End If
                c_MBar.WindowMenu("stylechanged") = 0
                bWasIntercepted = True
            End If
        
        Case WM_NCACTIVATE  ' nonclient actiating/deactivating
            UpdateWindowStyles True, True, 0, WS_CAPTION
            If wParam Then  ' active
                ' need to pass this to defwndproc per msdn examples
                DefWindowProc mainHwnd, wMsg, wParam, lParam
'                NewWndProc = CallWindowProc(hOldWndProc, mainHwnd, wMsg, wParam, lParam)
            Else    ' trial & error, don't pass this & bad things happen
                ReleaseMouse        ' do if needed
                'MenuBar.ResetMenu
'                DefWindowProc mainHwnd, wMsg, wParam, lParam
                NewWndProc = CallWindowProc(hOldWndProc, mainHwnd, wMsg, wParam, lParam)
            End If
'            'Debug.Print "got ncactivate "; wParam
            ' redraw the titlebar
            UpdateWindowStyles True, False, 0, WS_CAPTION
            c_MBar.WindowMenu("wm_ncactivate") = 0
            DoDrawTitleBar (wParam <> 0), True, "wmncactivate"
            bWasIntercepted = True  ' don't forward this message
        
        Case WM_NCPAINT ' painting non client area
'            'Debug.Print "got ncpaint"; wParam
            DoDrawTitleBar (GetForegroundWindow() = mainHwnd), True, "wm_ncpaint"
            bWasIntercepted = True  ' don't forward this message
            
    
        Case WM_NCHITTEST
'            'Debug.Print "got nc hittest"
            ' important to trick windows
            ' into believing the hittest we want is valid
            NewWndProc = ConvertHitTest(0, False)  ' c_MBar.TrackingState <> 0)
            bWasIntercepted = True  ' don't forward this message
        
        Case WM_NCLBUTTONDOWN
    '         'Debug.Print "got nclbuttondown"
            ' can't have it touch down on Non-client, determine where clicked
            If Not m_bIsActive Then ' could happen some how; check
                SetForegroundWindow mainHwnd
                'DoDrawTitleBar True,true
            End If
            ' figure out what action is expected from x,y location on titlebar
            Select Case SetHitTestAction(wMsg, wParam)
            Case Is < 1 ' ate the action, tell windows NO
                wParam = HTNOWHERE
            Case HTSYSMENU
                ' track sysmenu separately; see wm_sycommand/sc_keymenu remarks above
                PostMessage mainHwnd, WM_GETSYSMENU, SC_KEYMENU, ByVal 0&
                bWasIntercepted = True
            Case HTMINBUTTON, HTMAXBUTTON, HTCAPTION, HTCLOSE
                ' handled already
                bWasIntercepted = True ' don't forward this message
            Case HTLEFT To HTBOTTOMRIGHT    ' borders
                ' remove this style as it prevents resizing vertical captions
                ' to their absoulte minimum width. Add it back after resizing
                UpdateWindowStyles True, True, 0, WS_DLGFRAME
                bWasIntercepted = True
            Case Else ' other actions - don't care
                bWasIntercepted = True ' don't forward this message
            End Select
            
        Case &H211 ' enter menu loop
            'Debug.Print "entering menu loop"
            If wParam = 0 Then UpdateWindowStyles True, True, 0, WS_CAPTION

        Case &H212 ' exit meu loop
            'Debug.Print "exiting menu loop"
            UpdateWindowStyles True, False, 0, WS_CAPTION
    
        Case WM_NCLBUTTONDBLCLK
            Select Case wParam
            Case HTCAPTION  ' maximize or restore
                If IsZoomed(mainHwnd) Then
                    ProcessSysMenuItem SC_RESTORE, True
                Else
                    ProcessSysMenuItem SC_MAXIMIZE, True
                End If
            Case HTSYSMENU ' default sysicon double click
                ' only triggers when window iconic on desktop & nowhere else due
                ' to how sysMenu is handled in this class. See wm_nclbuttonup too
                ProcessSysMenuItem SC_CLOSE, True
            End Select
            ReleaseMouse
            m_Position.Tag = 0
            m_Position.Action = 0
            bWasIntercepted = True  ' don't forward this message
               
        Case WM_MOUSEMOVE, WM_NCMOUSEMOVE ' only used when moving/resizing
    '            'Debug.Print "getting mouse move "; wMsg = WM_NCMOUSEMOVE
            ' nc mouse messages are never forwarded
            If m_Position.Action > 0 Then
                Select Case m_Position.Action
                Case Is < 3 'only if tracking for moving/sizing window or rolling over min/max/close
                    bWasIntercepted = SetNewWindowPos(wParam)
                Case Is > HTMenuPlus, HTMAXBUTTON, HTMINBUTTON, HTCLOSE
                    ' probably dragging over titlebar buttons
                    DoButtonClick wParam, True
                    bWasIntercepted = True
                End Select
            Else
                If wMsg = WM_NCMOUSEMOVE Then
                    Select Case wParam
                    Case HTMenuPlus
                        c_MBar.DrawMenuBarItems m_DCuser, 0, hRect_Menu, m_Position.wRect, m_Position.Tag - 1, 0
                    End Select
                    bWasIntercepted = True
                End If
            End If
            If wParam <> HTMenuPlus And m_Position.Tag > 0 Then
                If c_MBar.TrackingState = False Then
                    c_MBar.DrawMenuBarItems m_DCuser, 0, hRect_Menu, m_Position.wRect, -2, -3
                    m_Position.Tag = 0
                    'Debug.Print "release menu hilighting"
                End If
            End If
            
        Case WM_MOUSEACTIVATE   ' still testing; may not need this message
    '         If Not m_bIsActive Then DoDrawTitleBar True,true
        
        Case WM_NCRBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_XBUTTONDOWN
           'bWasIntercepted = c_MBar.SetMouseAction(WM_XBUTTONDOWN, 0, 0, 0, wMsg)
           
        Case WM_NCRBUTTONUP, WM_CONTEXTMENU
            Select Case ConvertHitTest(0, False)
            Case HTCAPTION, HTSYSMENU ' user right clicked or user
                ' right clicked client area & dragged click on nonclient area
                ' Show system menu at x,y coords
                PostMessage mainHwnd, WM_GETSYSMENU, 0, ByVal lParam
            Case Is > 100: wParam = HTNOWHERE
            End Select
            bWasIntercepted = True  ' don't forward this message
            
        Case WM_LBUTTONDOWN ' failsafe to ensure we are drawing active
            'Debug.Print "got lbutton down"
            If Not m_bIsActive Then DoDrawTitleBar True, True, "wm_nlbuttondown"
'            bWasIntercepted = c_MBar.SetMouseAction(WM_LBUTTONDOWN, m_Position.hitTest, 0, 0, wMsg)
            m_Position.HitTest = 0
            
        Case WM_LBUTTONUP, WM_NCLBUTTONUP
            If c_MBar.TrackingState = 0 Then
                If m_Position.Action > 0 Then
                    If m_Position.Action > 2 Then
                        ' we were resizing or moving... terminate now
                        DoButtonClick wParam, False
                        m_Position.Tag = 0
                    Else
                        ' finished moving/sizing; replace window style as needed
                        If m_Position.Action = 2 Then UpdateUserResize False
                        UpdateWindowStyles True, False, 0, WS_DLGFRAME
                    End If
                    bWasIntercepted = True  ' don't forward this message
                    ReleaseMouse
                ElseIf wMsg = WM_NCLBUTTONUP Then ' test as double click on sysmenu
                    ' when sysmenu icon double clicked, the wm_getsysmenu prevents
                    ' the second click from registering as a doubleclick. Therefore,
                    ' we test for this message as it only occurs once in this scenario
                    ReleaseMouse
                    If m_Position.HitTest = HTSYSMENU And m_Position.Tag < 0 Then
                        ' see if double click was within the system doubleclick time
                        If GetTickCount() + m_Position.Tag < GetDoubleClickTime() Then
                            ' yep, get ready to close
                            m_Position.Tag = 0
                            SendMessage mainHwnd, WM_CANCELMODE, 0, 0
                            ProcessSysMenuItem SC_CLOSE, True
                            bWasIntercepted = True  ' don't forward this message
                        End If
                    End If
                End If
                m_Position.Action = 0
            Else
                bWasIntercepted = True
            End If
            
        Case WM_SETCURSOR
            If c_MBar.TrackingState Then
                wParam = HTCAPTION
                lParam = 0
            Else
                Select Case LoWord(lParam)
                Case HTLEFT To HTBOTTOMRIGHT    ' edges/borders
                    ' credit here goes 100% to Steve McMahon with this neat trick...
                    ' Hovering cursor over borders can trigger the min/max to redraw.
                    ' By making the window not visible & forwarding message then
                    ' making it visible again resolves that annoying side affect
                    p_LockUpdate = True
                    SetWindowLong mainHwnd, GWL_STYLE, UpdateWindowStyles(False, True, -1, WS_VISIBLE)
                    CallWindowProc hOldWndProc, mainHwnd, wMsg, wParam, lParam
                    p_LockUpdate = True
                    SetWindowLong mainHwnd, GWL_STYLE, UpdateWindowStyles(False, False, -1, WS_VISIBLE)
                    p_LockUpdate = False
                    bWasIntercepted = True  ' don't forward this message
                End Select
            End If
        Case WM_CANCELMODE
            c_MBar.TrackingState = False
            
        Case WM_SIZE
            Dim hRgn As Long, eRect As RECT
            GetWindowRect mainHwnd, eRect
            If GetObjectFromPointer(GetProp(mainHwnd, "lvImpCB_Ptr"), callBack) Then
                On Error Resume Next
                callBack.OnCreateRegion hRgn, 0, 0, eRect.Right - eRect.Left, eRect.Bottom - eRect.Top
                On Error GoTo 0
            End If
            If hRgn <> 0 Then SetWindowRgn mainHwnd, hRgn, True: DeleteObject hRgn
        
        Case &H120   ' wm_menuchar
            NewWndProc = c_MBar.TraceHotKey(LoWord(wParam), (tBarIcon(0).tRgn <> 0), False)
            bWasIntercepted = True
        
        Case &H231 ' enter size/move
            'Debug.Print "entering sizemove"
            UpdateUserResize True
        Case &H232 ' exit size/move
            'Debug.Print "exiting sizemove"
            UpdateUserResize False
        
        Case Else
            ' all other unTrapped windows messages
        End Select
    
        End If  ' end of non-minimized messages
        
    End Select
    
End If

If lSysMenuMsg Then ' system menu is to be displayed...
    ' remove styles that can make windows redraw min/max/close
    ' style will be replaced on wm_exitmenuloop
    UpdateWindowStyles True, True, 0, WS_CAPTION
    ' ok, either manually track system menu or let windows do it
    If (wParam = SC_KEYMENU And wMsg = WM_GETSYSMENU) Then
        c_MBar.TrackSystemMenu tBarIcon(0).tRgn, ((m_tbarAlign Or 1) = m_tbarAlign)
    Else
        CallWindowProc hOldWndProc, mainHwnd, lSysMenuMsg, wParam, lParam
    End If

Else    ' all other messages, window minimized or not
    
    If Not bWasIntercepted Then
'        'Debug.Print "trapped "; mainHwnd, wMsg, wParam; lParam
        NewWndProc = CallWindowProc(hOldWndProc, mainHwnd, wMsg, wParam, lParam)
    Else
'        'Debug.Print "not trapped "; mainHwnd, wMsg, wParam; lParam
    End If

End If
End Function

Private Sub ReleaseMouse()
' when mouse is clicked on titlebar/borders, the mouse is captured so that
' moving & sizing can be done via code. Gotta release the capture and
' any associated cursor clipping that is active.
If GetCapture() = mainHwnd Then
    ReleaseCapture
    ClipCursor ByVal 0
    ''Debug.Print "mouse released"
End If
End Sub

Private Sub DoDrawTitleBar(ByVal bIsActive As Boolean, bDrawAll As Boolean, bDebugWhere As String)

If m_Position.wRgn = 0 Or m_bLockWindow = True Then Exit Sub
If IsIconic(mainHwnd) Then Exit Sub

'If m_Position.wRgn = 0 Then Exit Sub

Dim hDC As Long

Dim textRect As RECT
Dim dcTBarRect As RECT, tBarRect As RECT

Dim hBrush As Long, hFont As Long
Dim fontFlags As Long, Looper As Integer
Dim tCaption As String
Dim capLen As Long, bOverridden As Boolean
Dim hOldBmp As Long, fColor(0 To 1) As Long
Dim callBack As CustomWindowCalls

m_bIsActive = bIsActive

' test for optional "keep active" flags
If m_KeepActive > 2 Then    ' flags are set
    If (m_KeepActive Or 8) = m_KeepActive Then  ' active in all threads
        bIsActive = True
    Else            ' check if in current thread & make active if so
        If (m_KeepActive Or 4) = m_KeepActive And (m_KeepActive And 1) <> 1 Then bIsActive = True
    End If
End If

hDC = GetWindowDC(mainHwnd)

With m_Position.wRect
    SetRect dcTBarRect, .Left + m_BorderXY.Left, .Top + m_BorderXY.Top, .Right - m_BorderXY.Right, .Bottom - m_BorderXY.Bottom
    OffsetRect dcTBarRect, -m_Position.wRect.Left, -m_Position.wRect.Top
'    Debug.Print "Left"; .Left
End With
If hClipRgn Then DeleteObject hClipRgn
hClipRgn = CreateRectRgnIndirect(dcTBarRect)
If bDrawAll Then
'    'Debug.Print ">>>drawing borders", bDebugWhere, bIsActive
    Dim edgeRgn As Long, eRect As RECT
    SetRect eRect, 0, 0, m_Position.wRect.Right - m_Position.wRect.Left, m_Position.wRect.Bottom - m_Position.wRect.Top
    edgeRgn = CreateRectRgnIndirect(eRect)
    CombineRgn edgeRgn, edgeRgn, hClipRgn, 4   ' exclude the client region
    SelectClipRgn hDC, edgeRgn
    If GetObjectFromPointer(GetProp(mainHwnd, "lvImpCB_Ptr"), callBack) Then
        On Error Resume Next
        callBack.UserDrawnBorder hDC, 0, 0, eRect.Right - eRect.Left, eRect.Bottom - eRect.Top, bIsActive, bOverridden
        On Error GoTo 0
    End If
    SelectClipRgn hDC, 0
    If Not bOverridden Then TempDrawBorders hDC, m_Position.wRgn, hClipRgn
Else
'    'Debug.Print ">>>drawing titlebar only", bDebugWhere, bIsActive
End If
SelectClipRgn hDC, hClipRgn

bOverridden = False

If bDrawAll Then
    If hRect_Menu.Bottom > 0 Then
    'Debug.Print ">>>drawing menubar"
        c_MBar.DrawMenuBarBkg m_DCuser, hDC, hRect_Menu.Right - hRect_Menu.Left, hRect_Menu.Bottom - hRect_Menu.Top, bIsActive
        c_MBar.DrawMenuBarItems m_DCuser, hDC, hRect_Menu, m_Position.wRect, -1, 0
    End If
End If
    
If hRgn_Title Then
'    'Debug.Print ">>>drawing titlebar", bDebugWhere
    
    ' use the titlebar region for a starting point
    GetRgnBox hRgn_Title, tBarRect  ' and offset to nonclient coords
    OffsetRect tBarRect, -m_Position.wRect.Left, -m_Position.wRect.Top
    dcTBarRect = tBarRect   ' copy & offset to 0,0 for DC printing
    OffsetRect dcTBarRect, -tBarRect.Left, -tBarRect.Top
    ' ensure correct size bitmap
    hOldBmp = ResizeBitmap(hDC, m_bmpTitlebar, dcTBarRect.Right, dcTBarRect.Bottom, m_DCuser, False)
    
    
    ' offer user opportunity to customize
    If GetObjectFromPointer(GetProp(mainHwnd, "lvImpCB_Ptr"), callBack) Then
        On Error Resume Next
        Call callBack.UserDrawnTitleBar(m_DCuser, dcTBarRect.Right, dcTBarRect.Bottom, bIsActive, bOverridden)
        On Error GoTo 0
    End If
    If Not bOverridden Then ' not customized
        If bIsActive Then
            fColor(0) = vbActiveTitleBar
            fColor(1) = MakeDWord(COLOR_GRADIENTACTIVECAPTION, -255)
        Else
            fColor(0) = vbInactiveTitleBar
            fColor(1) = MakeDWord(COLOR_GRADIENTINACTIVECAPTION, -255)
        End If
        If (m_tbarAlign Or 1) = m_tbarAlign Then
            GradientFill fColor(0), fColor(1), m_DCuser, 0, 0, dcTBarRect.Right, -dcTBarRect.Bottom, Abs(5 * (m_Position.Action = 2))
        Else
            GradientFill fColor(0), fColor(1), m_DCuser, 0, 0, dcTBarRect.Right, dcTBarRect.Bottom, Abs(5 * (m_Position.Action = 2))
        End If
    End If
    
    ' get the window title if any
    tCaption = String$(255, 0)
    capLen = GetWindowText(mainHwnd, tCaption, 255)
    If capLen Then
        ' extract window title
        tCaption = Left$(tCaption, capLen)
        If (m_tbarAlign Or 1) = m_tbarAlign Then    ' vertical fonts
            ' cached font is vertical, APIs can't measure vertical fonts
            ' So we need to create a horizontal one for measuring
            hFont = SelectObject(m_DCuser, GetUnrotatedFont())    ' into window DC
            ' calculate a visable title bar rect minus space for min/max/close
            SetRect textRect, 0, 0, dcTBarRect.Bottom - m_tbarFarExtent - m_tbarNearExtent, 1
            ' now have API calculate if caption fits & if not, it will add the ellipses ...
            DrawText m_DCuser, tCaption, -1, textRect, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT Or &H8000& Or &H10000
            ' done with font, need to destroy or memory leak occurs
            DeleteObject SelectObject(m_DCuser, hFont)
            fontFlags = DT_LEFT Or &H100 ' always left; APIs don't do centered vertical fonts
            ' if centering, then calculate the DT_VCENTER value
            textRect.Left = (dcTBarRect.Right - textRect.Bottom) \ 2 - 1
            If (m_tbarAlign Or 2) = m_tbarAlign Then    ' centered
                ' calculate the DT_CENTER value 'cause APIs don't do vertical fonts
                textRect.Bottom = (dcTBarRect.Bottom - textRect.Right - m_tbarFarExtent - m_tbarNearExtent) \ 2
                If textRect.Bottom < 0 Then textRect.Bottom = 0
            Else
                textRect.Bottom = 0
            End If
            SetRect textRect, textRect.Left, dcTBarRect.Bottom - textRect.Bottom - 3 - m_tbarNearExtent, dcTBarRect.Right - m_tbarFarExtent, 0
        Else    ' horizontal fonts, much easier
            textRect = dcTBarRect
            fontFlags = DT_LEFT Or DT_VCENTER Or &H8000&  ' includes ellipses
            ' modify to subtract space for min/max/close, icon & border edges
            textRect.Right = textRect.Right - m_tbarFarExtent - 3
            textRect.Left = textRect.Left + 3 + m_tbarNearExtent
            If (m_tbarAlign Or 2) = m_tbarAlign Then fontFlags = fontFlags Or DT_CENTER ' centered
        End If
        SetTextColor m_DCuser, ConvertVBSysColor(m_fColor(Abs(bIsActive = False)))
        If GetObjectFromPointer(GetProp(mainHwnd, "lvImpCB_Ptr"), callBack) Then
            On Error Resume Next
            callBack.BeforeDrawCaption m_DCuser, tCaption, textRect.Left, textRect.Top
            On Error GoTo 0
        End If
        ' select our vertical/horizontal font & draw the window title
        hFont = SelectObject(m_DCuser, m_Font)
        DrawText m_DCuser, tCaption, -1, textRect, fontFlags Or DT_SINGLELINE
        ' remove the vertical/horizontal font
        SelectObject m_DCuser, hFont
    End If
    ' call function to draw icon & buttons
    DrawBarIcons m_DCuser, -1, bsNormal
    ' copy to the window from buffer & remove buffer bitmap
    BitBlt hDC, tBarRect.Left, tBarRect.Top, dcTBarRect.Right, dcTBarRect.Bottom, m_DCuser, 0, 0, vbSrcCopy
    SelectObject m_DCuser, hOldBmp
    
End If
' remove any clipping region & release window's DC
SelectClipRgn hDC, 0
ReleaseDC mainHwnd, hDC

End Sub

Private Sub DoButtonClick(wParam As Long, bMouseMove As Boolean)
' routine determines what actions to take when mouse click is over the titlebar

Dim bOverButton As Boolean, IsDown As Boolean, btnID As Integer
' determine if mouse is over min/max/close

bOverButton = (m_Position.HitTest = wParam)

    Select Case m_Position.HitTest
    Case Is > HTMenuPlus, HTCLOSE, HTMINBUTTON, HTMAXBUTTON
        btnID = GetButtonRef(m_Position.HitTest, True)
        If tBarIcon(btnID).tState <> 2 * Abs(bOverButton) Then
            tBarIcon(btnID).tState = 2 * Abs(bOverButton)
            DrawBarIcons 0, btnID, 2 * Abs(bOverButton) ' toggle button
        End If
        If bOverButton And Not bMouseMove Then
            Select Case btnID
            Case 0: ' error do nothing? hasn't occured--just being safe
            Case 1: ProcessSysMenuItem SC_CLOSE, True
            Case 3: ProcessSysMenuItem SC_MINIMIZE, True
            Case 2
                If IsZoomed(mainHwnd) Then
                    ProcessSysMenuItem SC_RESTORE, True
                Else
                    ProcessSysMenuItem SC_MAXIMIZE, True
                End If
            Case Else   ' custom buttons
                Dim callBack As CustomWindowCalls
                If GetObjectFromPointer(GetProp(mainHwnd, "lvImpCB_Ptr"), callBack) And btnID < UBound(tBarIcon) + 1 Then
                    On Error Resume Next
                    callBack.UserButtonClick (Trim$(tBarIcon(btnID).ID))
                End If
            End Select
            tBarIcon(btnID).tState = bsNormal
            DrawBarIcons 0, btnID, bsNormal
        End If
   'Case HTSYSMENU  ' not handled here
    Case HTMENU     ' work in progress
    
    Case Else
    
    End Select
End Sub

Private Sub DrawBarIcons(hDC As Long, itemNr As Integer, _
    ByVal lState As ButtonStateConstants)

' Routine draws the min/max/close & sysIcon as needed
Dim dfc_RestoreMax As Long, lStyle As Long, bOverridden As Boolean
Dim bDisabled As Boolean, NewState As Long, tRect As RECT
Dim Looper As Integer, LoopStart As Integer, LoopEnd As Integer
Dim tBarRect As RECT, bReturnDC As Boolean, userRgn As Long
Dim callBack As CustomWindowCalls

' use the titlebar as a reference for location
GetRgnBox hRgn_Title, tBarRect

If itemNr < 0 Then  ' draw all
    LoopStart = 1
    LoopEnd = UBound(tBarIcon)
ElseIf itemNr = 0 Then
    LoopEnd = -1
Else    ' draw one item
    LoopStart = itemNr
    LoopEnd = LoopStart
End If
If hDC = 0 Then ' may or may not be passed; if itemnr>-1 then probably not
    hDC = GetWindowDC(mainHwnd)
    bReturnDC = True
    SetBkMode hDC, 3
    OffsetRect tBarRect, (m_Position.wRect.Left - tBarRect.Left), m_Position.wRect.Top - tBarRect.Top
End If
    
NewState = lState   ' cache desired draw state

For Looper = LoopStart To LoopEnd
    With tBarIcon(Looper)
        If .tRgn Then           ' 0 means hidden
            lState = NewState   ' cache state & may be changed
            If .tDefault = True Then
                Select Case lState
                Case bsDisabled
                    lState = DFCS_INACTIVE
                Case bsHover
                    lState = DFCS_HOT
                Case bsDown
                    lState = DFCS_PUSHED Or DFCS_HOT    ' hot may not show on all OS
                Case Else   ' normal state; whatever that may be
                    ' see if we have it optionally disabled, otherwise lState will = 0
'                    lState = bsNormal
                    lState = (sysMenuEnabled(Choose(Looper, smClose, smMaximize, smMinimize)) + 1) * DFCS_INACTIVE
                End Select
            Else
               lState = DFCS_BUTTONPUSH * Abs(.tState - 1)
            End If
            GetRgnBox .tRgn, tRect
            ' offset the Rect as needed & draw the button
            OffsetRect tRect, -tBarRect.Left, -tBarRect.Top
            If .tDefault = True Then ' default close/max/min vs custom
                If Looper = 2 Then ' max/restore depends on window state
                    If IsZoomed(mainHwnd) Then  ' draw restore
                        dfc_RestoreMax = DFCS_CAPTIONRESTORE
                    Else                        ' otherwise draw max
                        dfc_RestoreMax = DFCS_CAPTIONMAX
                    End If
                End If
                ' either hide or draw disabled if button is disabled
                If lState < 0 And m_bHideInactiveBtns = False Then lState = DFCS_INACTIVE
'                    ' get the button's current region location & convert to Rect
'                    DrawFrameControl hDC, tRect, DFC_CAPTION, _
                        lState Or Choose(Looper, DFCS_CAPTIONCLOSE, dfc_RestoreMax, DFCS_CAPTIONMIN)
'            Else
'                If .tPosition < 128 Then
'                    DrawFrameControl hdc, tRect, DFC_BUTTON, lState
'                End If
                If GetObjectFromPointer(GetProp(mainHwnd, "lvImpCB_Ptr"), callBack) Then
                    On Error Resume Next
                    userRgn = CreateRectRgnIndirect(tRect)
                    SelectClipRgn hDC, userRgn
                    Call callBack.UserDrawnButton(Looper, IIf(lState <> DFCS_INACTIVE, .tState, 3), hDC, tRect.Left, tRect.Top, tRect.Right - tRect.Left, tRect.Bottom - tRect.Top, m_bIsActive, bOverridden)
'                    Call callBack.UserDrawnButton(Trim$(.ID), .tState, hdc, tRect.Left, tRect.Top, tRect.Right - tRect.Left, tRect.Bottom - tRect.Top)
                    SelectClipRgn hDC, 0
                    DeleteObject userRgn
                    On Error GoTo 0
                End If
                If Not bOverridden Then DrawFrameControl hDC, tRect, DFC_CAPTION, _
                        lState Or Choose(Looper, DFCS_CAPTIONCLOSE, dfc_RestoreMax, DFCS_CAPTIONMIN)
            End If
        End If
    End With
Next
If itemNr < 1 Then ' do the system icon
    With tBarIcon(0)
        Dim hIcon As Long, bIsCopied As Boolean
        If .tRgn <> 0 Then 'sysmenu
            GetRgnBox .tRgn, tRect
            Select Case tRect.Right - tRect.Left
            Case Is < 2: NewState = 0
            Case Is < 24: NewState = 16
                hIcon = m_SmallIcon.Handle
            Case Is < 32: NewState = 24
                hIcon = SendMessage(mainHwnd, WM_GETICON, 1, ByVal 0)
                lStyle = CopyImage(hIcon, IMAGE_ICON, 24, 24, LR_COPYFROMRESOURCE)
                If lStyle Then
                    hIcon = lStyle
                    bIsCopied = True
                End If
            Case Else: NewState = 32
                hIcon = SendMessage(mainHwnd, WM_GETICON, 1, ByVal 0)
            End Select
            
            If hIcon Then
                OffsetRect tRect, -tBarRect.Left, -tBarRect.Top
                DrawIconEx hDC, tRect.Left + 1, tRect.Top + 1, hIcon, NewState, NewState, 0, 0, &H3
                If bIsCopied Then DestroyIcon hIcon
            End If
        Else
        
        End If
    End With
End If

If bReturnDC Then ReleaseDC mainHwnd, hDC
End Sub

Private Sub MeasureCaption()

If m_Font = 0 Or mainHwnd = 0 Then Exit Sub

Dim hDC As Long, cRect As RECT, tRect As RECT
Dim hOldFont As Long, hFont As Long, rtnVal As Boolean
Dim minHt As Long

hDC = GetWindowDC(mainHwnd)
hOldFont = SelectObject(hDC, GetUnrotatedFont())
DrawText hDC, "Testing", -1, cRect, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE
m_CaptionSize = 21 'cRect.Bottom + 6
DeleteObject SelectObject(hDC, hOldFont)
ReleaseDC mainHwnd, hDC
'Select Case m_CaptionSize
'    Case Is > 30: m_IconCX = 24
'    Case Is > 16: m_IconCX = m_CaptionSize - 6
'    Case Is > 3: m_IconCX = m_CaptionSize - 2
'    Case Else: m_IconCX = 0
'End Select
m_IconCX = 28: m_IconCY = 15
If mainHwnd Then UpdateWindowStyles False, False, -3, 0
End Sub

Private Function GetUnrotatedFont() As Long
If m_Font = 0 Then Exit Function

Dim tLogFont As LOGFONT
GetGDIObject m_Font, Len(tLogFont), tLogFont
tLogFont.lfOrientation = 0
tLogFont.lfEscapement = 0
GetUnrotatedFont = CreateFontIndirect(tLogFont)
End Function

Private Sub CalculateInitialStyles()

If m_OriginalStyle <> 0 Then Exit Sub

Dim CX As Long, CY As Long, hDC As Long
Dim styleEX As Long, styleL As Long
Dim hasCaption As Boolean, bExpandSizingBorder As Boolean
Dim lTaskBarOption As Long

styleEX = GetWindowLong(mainHwnd, GWL_EXSTYLE)
styleL = GetWindowLong(mainHwnd, GWL_STYLE)

lTaskBarOption = m_OriginalStyleEX
m_OriginalStyle = UpdateWindowStyles(True, True, -1, WS_MAXIMIZE, WS_MINIMIZE)
m_OriginalStyleEX = styleEX

bExpandSizingBorder = (m_SizingBorder = 0)
If m_BorderStyle = 0 Then

    If (styleL And WS_THICKFRAME) = WS_THICKFRAME Then
        If m_SizingBorder < 1 Then
            m_SizingBorder = GetSystemMetrics(SM_CXDLGFRAME)
            If (styleL And WS_BORDER) = WS_BORDER Then
                m_SizingBorder = m_SizingBorder + 1
                hasCaption = True
            End If
        End If
    ElseIf (styleL And WS_DLGFRAME) = WS_DLGFRAME Then
        If m_SizingBorder < 1 Then m_SizingBorder = GetSystemMetrics(SM_CXDLGFRAME)
        If (styleL And WS_BORDER) = WS_BORDER Then hasCaption = True
    Else
        If m_SizingBorder < 1 Then m_SizingBorder = 1
    End If
    m_BorderStyle = m_SizingBorder
    SetRect m_BorderXY, m_SizingBorder, m_SizingBorder, m_SizingBorder, m_SizingBorder
    
Else
    hasCaption = (m_BorderStyle > 1)
    If m_SizingBorder < 1 Then m_SizingBorder = m_BorderStyle
End If
If bExpandSizingBorder Then m_SizingBorder = m_SizingBorder + 5

If hasCaption Then
    If (styleEX And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW Then
        m_CaptionSize = GetSystemMetrics(SM_CYSMCAPTION)
    Else
        m_CaptionSize = GetSystemMetrics(SM_CYCAPTION) * Abs(m_SizingBorder > 0)
    End If
    m_IconCX = GetSystemMetrics(SM_CYSMICON)
    If m_IconCX > m_CaptionSize - 5 Then
        m_IconCX = m_CaptionSize - 5
    End If
Else
    m_CaptionSize = 0
    m_IconCX = 0
End If

If lTaskBarOption > 0 Then
    SetInTaskBar True, (lTaskBarOption = 2)
End If
UpdateWindowStyles False, False, -2, AddRqdStyles(GWL_STYLE, 0, ((styleL Or WS_MAXIMIZE) = styleL))

Set Font = Font

hDC = GetWindowDC(GetDesktopWindow())
m_DCuser = CreateCompatibleDC(hDC)
ReleaseDC GetDesktopWindow(), hDC
SetBkMode m_DCuser, 3

AdjustImplementSize False
If m_SmallIcon.Handle = 0 Then SetSmallIcon


End Sub

Private Sub CalculateBarRects(Optional bGetWindowRect As Boolean = True)
' calculates the dimensions of the titlebar and menubar due to
' initialization, style changes, and/or sizing

If m_bLockWindow Then Exit Sub

Dim clientRect As RECT, lMBarWidth As Long, tRect As RECT

If bGetWindowRect Then
    
    GetWindowRect mainHwnd, tRect
    If EqualRect(tRect, m_Position.wRect) Then
'        'Debug.Print "quick abort calcrects"
        Exit Sub
    End If
    m_Position.wRect = tRect
    
End If

' remove existing regions as necessary
If hRgn_Title Then DeleteObject hRgn_Title
If m_Position.wRgn Then DeleteObject m_Position.wRgn
If hRgn_Client Then DeleteObject hRgn_Client

' create new overall window region (usable for non-rectangular windows)
m_Position.wRgn = CreateRectRgnIndirect(m_Position.wRect)

' calculate the title bar region
With m_Position.wRect
    If m_CaptionSize Then
        If (m_tbarAlign Or 1) = m_tbarAlign Then
            hRgn_Title = CreateRectRgn(.Left + m_BorderXY.Left, _
                .Top + m_BorderXY.Top, .Left + m_BorderXY.Left + m_CaptionSize, _
                .Bottom - m_BorderXY.Bottom)
            clientRect.Left = .Left + m_BorderXY.Left + m_CaptionSize
            clientRect.Top = .Top + m_BorderXY.Top
            lMBarWidth = .Right - m_BorderXY.Right - (.Left + m_BorderXY.Left + m_CaptionSize)
        Else
            hRgn_Title = CreateRectRgn(.Left + m_BorderXY.Left, _
                .Top + m_BorderXY.Top, .Right - m_BorderXY.Right, _
                .Top + m_BorderXY.Top + m_CaptionSize)
            clientRect.Top = .Top + m_BorderXY.Top + m_CaptionSize
            clientRect.Left = .Left + m_BorderXY.Left
            lMBarWidth = .Right - m_BorderXY.Right - .Left - m_BorderXY.Left
        End If
    Else
        hRgn_Title = 0
        lMBarWidth = .Right - .Left - m_BorderXY.Right - m_BorderXY.Left
        clientRect.Left = .Left + m_BorderXY.Left
        clientRect.Top = .Top + m_BorderXY.Top
    End If

' calculate the menubar region
' if a menu exists, the lClientTop will probably be higher value unless
' the menubar exists within the title bar region
OffsetRect hRect_Menu, 0, -hRect_Menu.Top
MenuBar.GetBarDimensions m_DCuser, lMBarWidth, hRect_Menu.Bottom, clientRect.Left, clientRect.Top

If hRect_Menu.Bottom = 0 Then
    SetRect hRect_Menu, 0, 0, 0, 0
Else
    If (m_tbarAlign Or 1) = m_tbarAlign Then
        SetRect hRect_Menu, clientRect.Left, clientRect.Top, _
            .Right - m_BorderXY.Right, hRect_Menu.Bottom + m_BorderXY.Top + .Top
    Else
        SetRect hRect_Menu, clientRect.Left, _
            clientRect.Top, .Right - m_BorderXY.Right, _
            hRect_Menu.Bottom + clientRect.Top
    End If
    clientRect.Top = hRect_Menu.Bottom
End If

' calculate the client region
    If (m_tbarAlign Or 1) = m_tbarAlign Then
        SetRect clientRect, clientRect.Left, clientRect.Top, _
             .Right - m_BorderXY.Right, .Bottom - m_BorderXY.Bottom
'    'Debug.Print "new client rect is "; lClientLeft, lClientTop, _
             .Right - m_BorderXY.Right, .Bottom - m_BorderXY.Bottom
    Else
        SetRect clientRect, .Left + m_BorderXY.Left, clientRect.Top, _
            .Right - m_BorderXY.Right, .Bottom - m_BorderXY.Bottom
    End If
    If clientRect.Top > clientRect.Bottom Or clientRect.Left > clientRect.Right Then
        hRgn_Client = CreateRectRgn(.Left, .Top, .Left, .Top)
    Else
        hRgn_Client = CreateRectRgnIndirect(clientRect)
    End If
End With


If (m_sysMenu And 1024) <> 1024 Then CalculateButtonRect

'GetRgnBox hRgn_Client, tRect
'With tRect
'    'Debug.Print ">client "; .Left; .Right, .Top; .Bottom
'End With
'GetRgnBox hRgn_Title, tRect
'With tRect
'    'Debug.Print ">tBar "; .Left; .Right, .Top; .Bottom
'End With
'With hRect_Menu
'    'Debug.Print "menubar "; .Left; .Right, .Top; .Bottom
'End With
'With m_Position.wRect
'    'Debug.Print "window "; .Left; .Right, .Top; .Bottom
'End With

End Sub

Private Sub CalculateButtonRect()
' function calculates the titlebar button regions used
' for drawing and hittests

Dim xRect As RECT, wRect As RECT
Dim AdjustmentX As Long, AdjustmentY As Long, bDoButton As Boolean
Dim LoopStart As Integer, LoopEnd As Integer, Looper As Integer

m_tbarFarExtent = 0
m_tbarNearExtent = 0

If m_IconCX < 1 Or hRgn_Title = 0 Then  ' no titlebar

    For Looper = 0 To UBound(tBarIcon)
        If tBarIcon(Looper).tRgn Then DeleteObject tBarIcon(Looper).tRgn
        tBarIcon(Looper).tRgn = 0
    Next

Else

    GetRgnBox hRgn_Title, wRect

    If (m_tbarAlign Or 1) = m_tbarAlign Then
        AdjustmentY = m_IconCX + 1
'        SetRect xRect, wRect.Left, wRect.Top + 3, wRect.Left + m_IconCX, wRect.Top + m_IconCX + 3
        SetRect xRect, wRect.Left, wRect.Top + 3, wRect.Left + m_IconCX, wRect.Top + m_IconCY + 3
        OffsetRect xRect, (wRect.Right - wRect.Left - m_IconCX) \ 2, 0
    Else ' top aligned
        ' right aligned (default)
        AdjustmentX = -m_IconCX - 1
'        SetRect xRect, wRect.Right - m_IconCX - 3, wRect.Top, wRect.Right - 3, wRect.Top + m_IconCX
        SetRect xRect, wRect.Right - m_IconCX - 3, wRect.Top, wRect.Right - 3, wRect.Top + m_IconCY
        OffsetRect xRect, 0, (wRect.Bottom - wRect.Top - m_IconCY) \ 2
    End If
    For Looper = 1 To UBound(tBarIcon) ' close/max/min in that order
        With tBarIcon(Looper)
            If .tRgn Then DeleteObject .tRgn
            If Looper < 4 Then
                bDoButton = sysMenuEnabled(Choose(Looper, smClose, smMaximize, smMinimize))
                .tDefault = True
            Else
                bDoButton = (.tState < 2)
            End If
            If bDoButton Or m_bHideInactiveBtns = False Then
                If (.tPosition And Not 128) = tbPosDefault Then
                    .tRgn = CreateRectRgnIndirect(xRect)
                    If (m_tbarAlign Or 1) = m_tbarAlign Then
                        m_tbarFarExtent = xRect.Bottom + 3 - wRect.Top
                    Else
                        m_tbarFarExtent = wRect.Right - xRect.Left - 3
                    End If
                    'Debug.Print "calc btn "; Looper, xRect.Left; xRect.Top; xRect.Right; xRect.Bottom
                    OffsetRect xRect, AdjustmentX, AdjustmentY
                Else    ' custom position....
                    
                End If
            Else
                .tRgn = 0
            End If
        End With
    Next
    ' system icon
    If tBarIcon(0).tRgn Then DeleteObject tBarIcon(0).tRgn
    If sysMenuEnabled(smSysIcon) Then 'either show or not show icon
        If m_CaptionSize > 34 Then
            AdjustmentX = 32
        Else
            AdjustmentX = m_CaptionSize - 2
        End If
        If (m_tbarAlign Or 1) = m_tbarAlign Then
            SetRect xRect, wRect.Left, wRect.Bottom - AdjustmentX - 1, _
                wRect.Left + AdjustmentX, wRect.Bottom - 1
            m_tbarNearExtent = wRect.Bottom - xRect.Top
            OffsetRect xRect, (wRect.Right - wRect.Left - AdjustmentX) \ 2, 0
        Else ' top aligned
            ' right aligned (default)
            SetRect xRect, wRect.Left + 1, wRect.Top, _
                wRect.Left + AdjustmentX + 1, wRect.Top + AdjustmentX
            m_tbarNearExtent = xRect.Right - wRect.Left
            OffsetRect xRect, 0, (wRect.Bottom - wRect.Top - AdjustmentX) \ 2
        End If
        tBarIcon(0).tRgn = CreateRectRgnIndirect(xRect)
        tBarIcon(0).tDefault = True
    Else
        tBarIcon(0).tRgn = 0
    End If
End If
End Sub

Private Function UpdateWindowStyles(bLockRefresh As Boolean, bRemove As Boolean, ParamArray StyleChange() As Variant) As Long

Dim StyleOld As Long, styleL As Long, I As Integer, tgtStyle As Long

If UBound(StyleChange) < 1 Then Exit Function
If m_bLockWindow Then
    m_bLockWindow = False
    Exit Function
End If

styleL = GetWindowLong(mainHwnd, GWL_STYLE)

''Debug.Print "setting setwindowstyle "; styleL

m_bLockWindow = bLockRefresh
Select Case StyleChange(0)
Case -2 ' REPLACE WINDOWSTYLE
    SetWindowLong mainHwnd, GWL_STYLE, StyleChange(1)
    If Not m_bLockWindow Then
        m_bLockWindow = bLockRefresh
        SetWindowPos mainHwnd, 0, 0, 0, 0, 0, 551
    End If
Case -3
    SetWindowPos mainHwnd, 0, 0, 0, 0, 0, 551
Case Else   ' add/remove styles
    StyleOld = styleL
    For I = 1 To UBound(StyleChange)
        tgtStyle = CLng(StyleChange(I))
        If bRemove Then ' removing window style(s)
            If (styleL Or tgtStyle) = styleL Then styleL = styleL And Not tgtStyle
        Else    ' adding styles
            styleL = styleL Or tgtStyle
        End If
    Next
    If StyleChange(0) < 0 Then  ' just return new style
        UpdateWindowStyles = styleL
    Else
        If StyleChange(0) <> 0 Or styleL <> StyleOld Then
'            'Debug.Print "sending setwindowlong"
            SetWindowLong mainHwnd, GWL_STYLE, styleL
            If Not m_bLockWindow Then
'                'Debug.Print "sending setwindowpos"
                SetWindowPos mainHwnd, 0, 0, 0, 0, 0, 551 Or 1032
            End If
            'If bRepaint Then ShowWindow mainHwnd, 1
            UpdateWindowStyles = styleL
        End If
    End If
End Select
m_bLockWindow = False
''Debug.Print "adjusted style is "; styleL
End Function

Private Function AddRqdStyles(styleID As Long, curStyle As Long, bAsMaximized As Boolean) As Long

Const ws_requiredStyles As Long = WS_DLGFRAME Or WS_THICKFRAME _
    Or WS_SYSMENU Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_CAPTION

If curStyle = 0 Then curStyle = GetWindowLong(mainHwnd, styleID)
If styleID = GWL_STYLE Then
    curStyle = curStyle Or ws_requiredStyles
    If bAsMaximized Then
'        Select Case curStyle
'        Case wbBlackEdge
'            curStyle = curStyle And Not WS_DLGFRAME
            curStyle = curStyle And Not WS_THICKFRAME
'            curStyle = curStyle Or WS_BORDER
'        Case wbThin, wbDialog
'            curStyle = curStyle And Not WS_THICKFRAME
'        Case Else
'        End Select
    End If
    'If m_BorderStyle = -1 Then curStyle = curStyle And Not WS_CAPTION
Else
    ' todo: test other extended window styles to see if it screws things up
End If
AddRqdStyles = curStyle
End Function

Private Sub ConfigureNCandClient(lParam As Long)
            
'Exit Function
            
Dim tNCR As NCCALCSIZE_PARAMS
Dim tWP As WINDOWPOS

' Get the new window position & recommended client rect
    CopyMemory tNCR, ByVal lParam, Len(tNCR)
    CopyMemory tWP, ByVal tNCR.lppos, Len(tWP)

' set the new window rect here
    SetRect m_Position.wRect, tWP.X, tWP.Y, tWP.X + tWP.CX, tWP.Y + tWP.CY
    
' calculate the client, titlebar, menubar, & nonclient button regions
    CalculateBarRects False
       
' use the calculated client region here (screen coords vs client coords)
    GetRgnBox hRgn_Client, tNCR.rgrc(0)
        
' Return the new client area
    tNCR.rgrc(1) = tNCR.rgrc(0)
    CopyMemory ByVal lParam, tNCR, Len(tNCR)
                     
End Sub

Private Sub c_MBar_GetWindowRegion(lRegion As Long)
lRegion = hClipRgn
End Sub

Private Sub c_MBar_RefreshMenuBar()
DoDrawTitleBar m_bIsActive, True, "refresh"
End Sub

Private Function ProcessSysMenuItem(mnuID As Long, bPost As Boolean) As Boolean
Dim EatMsg As Boolean

EatMsg = True
Select Case mnuID
Case SC_CLOSE
    If sysMenuEnabled(smClose) Then PostMessage mainHwnd, WM_CLOSE, 0, ByVal 0&
Case SC_MAXIMIZE
    If sysMenuEnabled(smMaximize) Then
        If bPost Then PostMessage mainHwnd, WM_SYSCOMMAND, SC_MAXIMIZE, ByVal 0&
        EatMsg = False
    End If
Case SC_MINIMIZE
    If sysMenuEnabled(smMinimize) Then
        If bPost Then PostMessage mainHwnd, WM_SYSCOMMAND, SC_MINIMIZE, ByVal 0&
        EatMsg = False
    End If
Case SC_RESTORE
    If bPost Then PostMessage mainHwnd, WM_SYSCOMMAND, SC_RESTORE, ByVal 0&
    EatMsg = False
Case SC_MOVE
    If sysMenuEnabled(smMove) And IsIconic(mainHwnd) = 0 Then
        If bPost Then PostMessage mainHwnd, WM_SYSCOMMAND, SC_MOVE, ByVal 0&
        EatMsg = False
    End If
Case SC_SIZE
    If sysMenuEnabled(smSize) Then
        If bPost Then PostMessage mainHwnd, WM_SYSCOMMAND, SC_SIZE, ByVal 0&
        EatMsg = False
    End If
Case Else
    EatMsg = False
End Select
ProcessSysMenuItem = EatMsg

End Function

Private Sub Class_Initialize()
Set Font = Nothing
m_fColor(0) = vbActiveTitleBarText
m_fColor(1) = vbInactiveTitleBarText
End Sub

Private Sub TweakSysMenu(bRestore As Boolean)

' Function toggles enabled/default status of system menu items &
' is called just before the system menu is displayed...
' bRestore is one-time & resets system menu IDs back to defaults
' bToShow is one-time & called when custom window first being called
' bUpdate is to toggle status before displayed
' bSysKeyMenu is used to indicate shown via click on sysIcon or Alt+Space

Dim MII As MENUITEMINFO, hSysMenu As Long, I As Integer
Dim mnuID As Long   ' actual system menu id: i.e., SC_CLOSE
Dim cloneID As Long ' must be integer size. Custom ID for system menu
Dim curID As Long   ' ID to be used to ref menu structure
Dim curState As Long
Dim isMaxed As Boolean, isMinD As Boolean   ' zoomed/iconic states

isMaxed = (IsZoomed(mainHwnd) <> 0) ' get states
isMinD = (IsIconic(mainHwnd) <> 0)

MII.cbSize = Len(MII)
MII.fMask = MIIM_ID Or MIIM_STATE   ' only need to tweak ID & States
hSysMenu = GetSystemMenu(mainHwnd, 0)   ' get current system menu
If hSysMenu = 0 Then                    ' if none, get standard default
    GetSystemMenu mainHwnd, 1
    hSysMenu = GetSystemMenu(mainHwnd, 0)
End If
For I = 1 To 5  ' loop thru the 6 system menu items
    Select Case I
    Case 1  ' close
        mnuID = SC_CLOSE: cloneID = SC_CLOSE_Clone
    Case 2  ' max
        mnuID = SC_MAXIMIZE: cloneID = SC_MAXIMIZE_Clone
    Case 3  ' min
        mnuID = SC_MINIMIZE: cloneID = SC_MINIMIZE_Clone
    Case 4  ' move
        mnuID = SC_MOVE: cloneID = SC_MOVE_Clone
    Case 5  ' size
        mnuID = SC_SIZE: cloneID = SC_SIZE_Clone
    Case 6  ' restore
        mnuID = SC_RESTORE: cloneID = SC_RESTORE_Clone
    End Select
    ' now get the structure. Non-system items, I don't care about
    curID = cloneID
    If bRestore = True Then
        If GetMenuItemInfo(hSysMenu, curID, False, MII) Then
            MII.fState = 0
            MII.wID = mnuID
            SetMenuItemInfo hSysMenu, curID, False, MII
        End If
    Else                ' start or conntinue subclass - set ID
        ' different parts of the program can enable/disable the menu items
        ' We need to test for both unfortunately...
        If GetMenuItemInfo(hSysMenu, curID, False, MII) = 0 Then
            curID = mnuID
            If GetMenuItemInfo(hSysMenu, curID, False, MII) = 0 Then cloneID = mnuID
        End If
        If cloneID <> mnuID Then
            MII.fState = 0
            MII.wID = mnuID
            Select Case I
            Case 1  ' close
                If Not sysMenuEnabled(smClose) Then ' Close disabled: optional
                    MII.wID = cloneID
                End If
            Case 2: ' maximized
                ' if disabled option or already zoomed, disable
                If sysMenuEnabled(smMaximize) = False Then
                    MII.wID = cloneID
                End If
            Case 3: ' minimized
                ' if disabled option or already minimized, disable
                If sysMenuEnabled(smMinimize) = False And isMinD = False Then
                    MII.wID = cloneID
                End If
            Case 4: ' move
                ' if disabled option or zoomed/iconic, disable
                If sysMenuEnabled(smMove) = False And isMinD = False Then
                    MII.wID = cloneID
                End If
            Case 5: ' size
                ' if disabled option or zoomed/iconic, disable
                If sysMenuEnabled(smSize) = False And isMinD = False Then
                    MII.wID = cloneID
                End If
            Case 6: ' restore
                If Not bRestore Then MII.wID = cloneID
            End Select
            ' update as needed
            If cloneID <> mnuID Then
                If MII.wID = cloneID Then MII.fState = MFS_GRAYED
                SetMenuItemInfo hSysMenu, curID, False, MII
            End If
        End If
    End If
Next
End Sub

Friend Sub SetApplicationIcon(hIcon As Long)
Dim tHwnd As Long, cParent As Long
On Error Resume Next

tHwnd = GetWindowLong(mainHwnd, GWL_HWNDPARENT)
Do While tHwnd
    cParent = tHwnd
    tHwnd = GetWindowLong(cParent, GWL_HWNDPARENT)
Loop
PostMessage cParent, WM_SETICON, 1, ByVal hIcon
End Sub

Private Sub SetSmallIcon()

' routine sets the small icon displayed on the titlebar, if any
' VB won't always create a small icon that can be returned by a
' call to SendMessage WM_GETICON, ICON_SMALL.
' So we will create one and set it ourselves. Other routines in
' this project call the small icon for painting on the titlebar

Dim smallIcon As Long, bigIcon As Long

' see if a big icon exists. If not, the user has set Me.Icon=Nothing

If m_SmallIcon.ToBeReplaced = False Then
    m_SmallIcon.ToBeReplaced = True
    m_SmallIcon.hOriginalIcon = SendMessage(mainHwnd, WM_GETICON, 0, ByVal 0&)
End If
bigIcon = SendMessage(mainHwnd, WM_GETICON, 1, ByVal 0&)
If bigIcon Then ' got a big icon
    ' testing seems to show that creating our own may be better
    ' quality than the one VB gives us; at worse it is the same
    ' quality. So we will always create a new one
    smallIcon = CopyImage(bigIcon, IMAGE_ICON, 16, 16, LR_COPYFROMRESOURCE)
    If smallIcon Then   ' small icon created properly
        ' set the icon now
        PostMessage mainHwnd, WM_SETICON, 0, ByVal smallIcon
        ' destroy previous icon & cache handle & delete status
        If m_SmallIcon.ToBeDeleted Then DestroyIcon m_SmallIcon.Handle
        m_SmallIcon.Handle = smallIcon
        m_SmallIcon.ToBeDeleted = True
    Else    ' couldn't create one, but if VB gave us one, use it
        smallIcon = SendMessage(mainHwnd, WM_GETICON, 0, ByVal 0&)
        m_SmallIcon.ToBeDeleted = False
        If m_SmallIcon.ToBeDeleted Then DestroyIcon m_SmallIcon.Handle
        If smallIcon Then
            m_SmallIcon.Handle = smallIcon
        Else
            ' still more checking. What if there was no small icon?
            ' Here we will use the bigicon which will be shrunk while drawing
            m_SmallIcon.Handle = bigIcon
        End If
    End If
Else    ' user set Me.Icon=Nothing, what to do?
    ' VB displays the Flying Window icon in these cases if the form is
    ' in the taskbar. So we will simply set the icon to that
    If m_SmallIcon.ToBeDeleted Then DestroyIcon m_SmallIcon.Handle
    m_SmallIcon.Handle = 0
    m_SmallIcon.ToBeDeleted = False
    ' Assumption: All Win O/S have a user.exe file?
    ExtractIconEx "user.exe", 0, ByVal 0&, smallIcon, 1
    If smallIcon Then
        m_SmallIcon.Handle = smallIcon
        m_SmallIcon.ToBeDeleted = True
        PostMessage mainHwnd, WM_SETICON, 0, ByVal smallIcon
    Else
        ' could this happen? only if there was no user.exe
    End If
End If
End Sub

Private Function GetButtonRef(inRef As Long, byHitTest As Boolean) As Integer
Dim btnID As Integer
If byHitTest Then   ' hit test code
    Select Case inRef
    Case HTMINBUTTON: btnID = 3
    Case HTMAXBUTTON: btnID = 2
    Case HTCLOSE: btnID = 1
    Case HTSYSMENU: btnID = 0
    Case Is > HTMenuPlus: btnID = inRef - HTMenuPlus
    End Select
Else    ' local enum smClose, smMaximize, etc
    Select Case inRef
    Case smClose: btnID = 1
    Case smMaximize: btnID = 2
    Case smMinimize: btnID = 3
    Case Else: btnID = 0
    End Select
End If
GetButtonRef = btnID
End Function

Private Sub UpdateUserResize(bEnterSizeMode As Boolean)
Dim callBack As CustomWindowCalls, wasRedrawn As Boolean
If GetObjectFromPointer(GetProp(mainHwnd, "lvImpCB_Ptr"), callBack) Then
    callBack.EnterExitSizing bEnterSizeMode, wasRedrawn
End If
If Not wasRedrawn And Not bEnterSizeMode Then
    m_Position.Action = 0
    c_MBar.ReleaseRefreshLock = True
    DoDrawTitleBar m_bIsActive, True, "updateuserresize"
End If
End Sub

Private Sub AdjustImplementSize(isTerminating As Boolean)
' routine attempts to adjust the window's right & bottom
' margins to accomodate the new custom window style whether
' caption is vertical or not.

If IsZoomed(mainHwnd) = 0 And IsIconic(mainHwnd) = 0 Then
    Dim cRectNow As RECT, cRectNew As RECT
    If isTerminating Then   ' subclassing has ended
        GetRgnBox hRgn_Client, cRectNow
        GetClientRect mainHwnd, cRectNew
        OffsetRect cRectNow, -cRectNow.Left, -cRectNow.Top
        cRectNew.Left = (cRectNow.Right - cRectNew.Right)
        cRectNew.Top = (cRectNow.Bottom - cRectNew.Bottom)
    Else
        m_sysMenu = m_sysMenu Or 1024
        GetClientRect mainHwnd, cRectNow
        CalculateBarRects True
        GetRgnBox hRgn_Client, cRectNew
        OffsetRect cRectNew, -cRectNew.Left, -cRectNew.Top
        cRectNew.Left = (cRectNow.Right - cRectNew.Right)
        cRectNew.Top = (cRectNow.Bottom - cRectNew.Bottom)
        m_sysMenu = m_sysMenu And Not 1024
        c_MBar.WindowMenu("adjustImplementSize") = 0
    End If
    SetWindowPos mainHwnd, 0, 0, 0, _
        m_Position.wRect.Right - m_Position.wRect.Left + cRectNew.Left, _
        m_Position.wRect.Bottom - m_Position.wRect.Top + cRectNew.Top, 550
End If
End Sub

Private Sub Class_Terminate()
Erase m_fColor()
End Sub
