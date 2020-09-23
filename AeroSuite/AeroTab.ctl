VERSION 5.00
Begin VB.UserControl AeroTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   194
   ToolboxBitmap   =   "AeroTab.ctx":0000
End
Attribute VB_Name = "AeroTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------
' AeroTab ActiveX Control
' Based on ucXTab v1.0.152 by Paul R. Territo, Ph.D [CodeId=66998]
'--------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'--------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'--------------------------------------------------------------------------

Option Explicit

'Bitmap type used to store Bitmap Data
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type TabInfo
    ControlBoxRect As RECT              'Coordinates for the ControlBox rectangle
    ControlID As String                 'Unique ID for the Item
    Caption As String                   'Caption for the tab
    ClickableRect As RECT               'Coordinates of the clickable rectangle
    Count() As Long                     'Tab Count
    Enabled As Boolean                  'Enabled?
    AccessKey As Long                   'Accelerator Key
    TabPicture As StdPicture            'Tab Picture
    TabStop As Long                     'Tab Stop Number
End Type

Private Enum xStateEnum                 'Style for ControlBox State
    xStateNormal = &H0                  '-->Normal Colors
    xStateHover = &H1                   '-->Hover Colors
    xStateDown = &H2                    '-->Down Colors
End Enum
#If False Then
    Const xStateNormal = &H0
    Const xStateHover = &H1
    Const xStateDown = &H2
#End If

Public Enum Style                       'Style for the Tabs
    xStyleTabbedDialog = &H0            '-->Tabbed Dialog
    xStylePropertyPages = &H1           '-->Property Pages
End Enum
#If False Then
    Const xStyleTabbedDialog = &H0
    Const xStylePropertyPages = &H1
#End If

Public Enum PictureAlign
    xAlignLeftEdge = &H0                '-->Left edge of the Tab
    xAlignRightEdge = &H1               '-->Right Edge of the Tab
    xAlignLeftOfCaption = &H2           '-->Left of the caption
    xAlignRightOfCaption = &H3          '-->Right of the caption
End Enum
#If False Then
    Const xAlignLeftEdge = &H0
    Const xAlignRightEdge = &H1
    Const xAlignLeftOfCaption = &H2
    Const xAlignRightOfCaption = &H3
#End If

Public Enum PictureSize                 'Determines Picture size on tabs
    xSizeSmall = &H0
    xSizeLarge = &H1
End Enum
#If False Then
    Const xSizeSmall = &H0
    Const xSizeLarge = &H1
#End If
'===Constants=========================================================================================================

'   Draw Text Constants
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_LEFT = &H0
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4

'   Window Position Constants
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

'   DrawIcon Related Constants
Private Const DI_NORMAL As Long = &H3

'   Windows Versioning Constants
Private Const VER_PLATFORM_WIN32_NT = 2

'   GetSystemMetrics Related Condtants
Private Const SM_CXICON As Long = 11
Private Const SM_CXSMICON As Long = 49

'=====================================================================================================================
'   The distance between the text (caption) of the tab and the focus Rect
Private Const iFOCUS_RECT_AND_TEXT_DISTANCE As Integer = 2
'   The distance between the text and the border in a Property Pages style tab
Private Const iPROP_PAGE_BORDER_AND_TEXT_DISTANCE As Integer = 7
'   The top for the property page (inactive property page)
Private Const iPROP_PAGE_INACTIVE_TOP As Integer = 2
'   The width of the control box [x]
Private Const iPROP_CONTROLBOX As Integer = 12
'===Declarations======================================================================================================

'   Drawing/Painting Declarations
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetObjectA Lib "gdi32.dll" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

'   Subclassing Related Declararions
Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'===Constants=========================================================================================================

'   Default property values
Private Const m_def_bUseControlBox As Boolean = False                   'Init Control Box on Each Tab
Private Const m_def_bShowFocusRect As Boolean = True                    'Init Focus Rect
Private Const m_def_bEnabled As Boolean = True                          'Init Usercontrol Enabled State
Private Const m_def_bTabEnabled As Boolean = True                       'Init Tab Enabled
Private Const m_def_bUseFocusedColor As Boolean = True                  'Init UseFocusColor Flag
Private Const m_def_bUseMaskColor As Boolean = True                     'Init UseMaskColor Flag
Private Const m_def_lActiveTab As Long = 0                              'Init Active Tab
Private Const m_def_lActiveTabBackEndColor As Long = vbButtonFace       'Init TabBack End Color
Private Const m_def_lActiveTabBackStartColor As Long = vbButtonFace     'Init TabBack Start Color
Private Const m_def_lActiveTabForeColor As Long = vbButtonText          'Init Active Tab ForeColor
Private Const m_def_lActiveTabHeight As Long = 20                       'Init Height for Active Tab
Private Const m_def_lBottomRightInnerBorderColor As Long = vb3DShadow   'Init Bottom-Right Border Color
Private Const m_def_lDisabledTabBackColor As Long = vb3DShadow          'Init Disable Tab Backcolor
Private Const m_def_lDisabledTabForeColor As Long = vb3DHighlight       'Init Disabled Tab ForeColor
Private Const m_def_lFocusedColor As Long = &HEE8269                    'Init Focused Color
Private Const m_def_lHoverColor As Long = &H3BC7FF                      'Init Hover Color
Private Const m_def_lInActiveTabBackEndColor As Long = vbButtonFace     'Init Inactive TabBack End Color
Private Const m_def_lInActiveTabBackStartColor As Long = vbButtonFace   'Init Inactive TabBack Start Color
Private Const m_def_lInActiveTabForeColor As Long = vbButtonText        'Init Inactive Tab ForColor
Private Const m_def_lInActiveTabHeight As Long = 18                     'Init Height for Inactive Tab
Private Const m_def_lLastActiveTab As Long = m_def_lActiveTab           'Init Last Active Tab
Private Const m_def_lOuterBorderColor As Long = vb3DDKShadow            'Init Outer Border Coor
Private Const m_def_lPictureAlign As Long = xAlignLeftEdge              'Init Picture Align
Private Const m_def_lPictureMaskColor As Long = &HC0C0C0                'Init Transparany Color
Private Const m_def_lPictureSize As Long = xSizeSmall                   'Init Picture Size
Private Const m_def_lStyle As Long = xStyleTabbedDialog                 'Init Tab Style
Private Const m_def_lTabCount As Long = 3                               'Init Tab Count
Private Const m_def_lTopLeftInnerBorderColor As Long = vb3DHighlight    'Init Top-Left Inner Border Color
Private Const m_def_lXRadius As Long = 10                               'Init Radius used in the RoundTabs Theme to draw the rounded tab
Private Const m_def_lYRadius As Long = 10                               'Init Radius used in the RoundTabs Theme to draw the rounded tab
Private Const m_def_sCaption As String = "Tab"                          'Default caption that is appended to form default name "Tab 0", "Tab 1" etc
Private Const m_def_UseMouseWheelScroll As Boolean = True               'Init MouseWheel Supoort for Scrolling
'=====================================================================================================================

'===Property Variables================================================================================================
Private m_aryTabs() As TabInfo                      'Array of Tabs
Private m_bAreControlsAdded As Boolean              'Controls Loaded Flag
Private m_bCancelFlag As Boolean                    'Used to pass as a cancel flag with the events to the container.
Private m_bEnabled As Boolean                       'Usercontrol Enabled State
Private m_bIsBackgroundPaintDelayed As Boolean      'See lTheme_DrawBackground() to get a description of this flag
Private m_bIsMouseOver As Boolean                   'Used to Track the MouseMovement
Private m_bIsRecursive As Boolean                   'IsRecursive Flag - Prevents Recursion on Load and Redraws
Private m_bShowFocusRect As Boolean                 'Focus Rectangle Flag
Private m_bUseControlBox As Boolean                 'UseControl Box on each Tab
Private m_bUseFocusedColor As Boolean               'UseFocused Color Flag
Private m_bUseMaskColor As Boolean                  'Use MaskColor for Pictures - Not needed for Icon Image Types
Private m_bUseMouseWheelScroll As Boolean           'Use Mouse
Private m_enmPictureAlign As PictureAlign           'Tab Picture Alignment
Private m_enmPictureSize As PictureSize             'Tab Picture Size (Small/Large)
Private m_enmStyle As Style                         'Tab Style for the UserControl
Private m_IsFocused As Boolean                      'Determines Focused State
Private m_lActiveTab As Long                        'Stores the Active Tab index
Private m_lActiveTabBackEndColor As OLE_COLOR       'Active Tab's Back Start Color
Private m_lActiveTabBackStartColor As OLE_COLOR     'Active Tab's Back End Color
Private m_lActiveTabForeColor As OLE_COLOR          'Active Tab ForeColor
Private m_lActiveTabHeight As Long                  'Active Tab Height
Private m_lBottomRightInnerBorderColor As OLE_COLOR 'Tab's Bottom Right Inner Border Color
Private m_lDisabledTabBackColor As OLE_COLOR        'Disabled TabBackColor
Private m_lDisabledTabForeColor As OLE_COLOR        'Disabled TabForeColor
Private m_lFocusedColor As OLE_COLOR                'Focused Color - Used with WinXP Style
Private m_lhDC As Long                              'Device Context for the UserControl Window
Private m_lHoverColor As OLE_COLOR                  'Tab's MouseOver Color
Private m_lhWnd As Long                             'Handle for the UserControl Window
Private m_lIconSize As Long                         'Icon Size
Private m_lInActiveTabBackEndColor As OLE_COLOR     'InActive Tab's Back End Color
Private m_lInActiveTabBackStartColor As OLE_COLOR   'InActive Tab's Back Start Color
Private m_lInActiveTabForeColor As OLE_COLOR        'Inactive Tab ForeColor
Private m_lInActiveTabHeight As Long                'InActive Tab Height
Private m_lLastActiveTab As Long                    'Stores the Last Active Tab Index
Private m_lMouseOverTabIndex As Long                'Stores Index of Hover Tab
Private m_lMoveOffset As Long                       'Offset Move Controls when a Tab is Clicked
Private m_lOuterBorderColor As OLE_COLOR            'Outer Border Color
Private m_lTabCount As Long                         'The Number of Tabs
Private m_lTabStripBackColor As OLE_COLOR           'TabStrip Back Color
Private m_lTopLeftInnerBorderColor As OLE_COLOR     'Tabs Top Left Inner Border Color
Private m_lXRadius As Long                          'Corner XRadius for Rounded Tabs
Private m_lYRadius As Long                          'Corner YRadius for Rounded Tabs
Private m_oActiveTabFont As StdFont                 'Tab's Font
Private m_oInActiveTabFont As StdFont               'InActive Tab's Font
Private m_Pnt As POINTAPI                           'Used to Store the X,Y Position in Subclassing uMsg Section
Private m_lScaleHeight As Long                      'Scale Height of UserControl
Private m_lScaleWidth As Long                       'Scale Width of UserControl
Private m_utRect As RECT                            'Stores a Copy of RECT
'=====================================================================================================================

Private WithEvents SDIHost  As Form
Attribute SDIHost.VB_VarHelpID = -1
Private WithEvents MDIHost  As MDIForm
Attribute MDIHost.VB_VarHelpID = -1

'Public Events=======================================================================================================


'   Note that bCancel is passed by Reference in below event. This event is called just before a
'   tab is being switched, we can prevent tab switch by making bCancel as true
Public Event BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
                                                                                  
'   If we Set bCancel in the BeforeTabSwitch following event will not occur.
Public Event TabSwitch(ByVal lLastActiveTab As Integer)

'   Public Events
Public Event Click()
Public Event ControlBoxEnter()
Public Event ControlBoxExit()
Public Event ControlBoxHover(X As Single, Y As Single)
Public Event ControlBoxMouseDown(X As Single, Y As Single)
Public Event ControlBoxMouseUp(X As Single, Y As Single)
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseHover(ActiveTab As Long, X As Single, Y As Single)
Public Event MouseScrollUp()
Public Event MouseScrollDown()
Public Event TabInsert(AfterTabIndex As Long)
Public Event TabRemove(TabIndex As Long)

'IMPORTANT EVENT :  used to solve a bug with original ssTab....
'used to tell container when the tab is completely initialised.
Public Event AfterCompleteInit()
'==================================================================================================
' ucSubclass - A sample UserControl demonstrating self-subclassing
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........
'==================================================================================================
'
'   SelfSubClasser Events
Public Event MouseEnter()
Public Event MouseLeave()
Public Event Status(ByVal sStatus As String)

'   Message Constants
Private Const WM_ACTIVATE               As Long = &H6
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_MOVING                 As Long = &H216
Private Const WM_SIZING                 As Long = &H214
Private Const WM_MOUSEWHEEL             As Long = &H20A
Private Const WM_EXITSIZEMOVE           As Long = &H232
Private Const WM_TIMER                  As Long = &H113
Private Const WM_LBUTTONDBLCLK          As Long = &H203
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_LBUTTONUP              As Long = &H202

'   Mouse Tracking Constants
Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

'   Mouse Tracking Structure
Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

'   SelfSubclasser Local Properties
Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean
Private bInCtrlBox                   As Boolean
Private m_lTimerID                   As Long
Private bSubClass                    As Boolean

'   SelfSubclasser API Declares
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'==================================================================================================

'   Subclasser Declarations
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hWnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'======================================================================================================
'   UserControl Private Routines

'   Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hmod        As Long
    Dim bLibLoaded  As Boolean

    hmod = GetModuleHandleA(sModule)

    If hmod = 0 Then
        hmod = LoadLibraryA(sModule)
        If hmod Then
            bLibLoaded = True
        End If
    End If

    If hmod Then
        If GetProcAddress(hmod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    If bLibLoaded Then
        Call FreeLibrary(hmod)
    End If
End Function


Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
'   Track the mouse leaving the indicated window
    Dim tme As TRACKMOUSEEVENT_STRUCT
  
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    'Parameters:
        'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
        'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
        'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
        'hWnd     - The window handle
        'uMsg     - The message number
        'wParam   - Message related data
        'lParam   - Message related data
    'Notes:
        'If you really know what you're doing, it's possible to change the values of the
        'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
        'values get passed to the default handler.. and optionaly, the 'after' callback
    Static bMoving As Boolean
    Dim ActiveRect As RECT, MouseOverRect As RECT, ControlBoxRect As RECT
    Dim iCnt As Long

    Select Case uMsg
        Case WM_ACTIVATE
            '   Store the current tab mouseover data so when the
            '   control regains focus we can repaint the header caps
            '   correctly....this mainly is required for XP Style
            '   when using a Focused Color
            m_lMouseOverTabIndex = m_lActiveTab
            Refresh
            
        Case WM_MOUSEMOVE
            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                RaiseEvent MouseEnter
            End If
            '   Get our position
            Call GetCursorPos(m_Pnt)
            '   Convert coordinates
            Call ScreenToClient(m_lhWnd, m_Pnt)
            '   Copy the active Rect to temp variable
            ActiveRect = AryTabs(m_lActiveTab).ClickableRect
            MouseOverRect = AryTabs(m_lMouseOverTabIndex).ClickableRect
            '   Check if the mouse is in the clickable region of the tab...
            If PtInRect(ActiveRect, m_Pnt.X, m_Pnt.Y) Then
                RaiseEvent MouseHover(m_lActiveTab, CSng(m_Pnt.X), CSng(m_Pnt.Y))
                If m_bUseControlBox Then
                    '   Copy the ControlBox Rect to temp variable
                    ControlBoxRect = AryTabs(m_lActiveTab).ControlBoxRect
                    '   See if we are in the ControlBox
                    If PtInRect(ControlBoxRect, m_Pnt.X, m_Pnt.Y) Then
                        If AryTabs(iCnt).Enabled Then
                            If Not bInCtrlBox Then
                                bInCtrlBox = True
                                RaiseEvent ControlBoxEnter
                            Else
                                RaiseEvent ControlBoxHover(CSng(m_Pnt.X), CSng(m_Pnt.Y))
                            End If
                            Call DrawControlBox(xStateHover, m_lActiveTab)
                        Else
                            Call DrawControlBox(xStateDown, m_lActiveTab)
                        End If
                    Else
                        If m_bEnabled Then
                            If AryTabs(m_lActiveTab).Enabled Then
                                If bInCtrlBox Then
                                    RaiseEvent ControlBoxExit
                                    bInCtrlBox = False
                                End If
                                Call DrawControlBox(xStateNormal, m_lActiveTab)
                            Else
                                Call DrawControlBox(xStateDown, m_lActiveTab)
                            End If
                        Else
                            Call DrawControlBox(xStateDown, m_lActiveTab)
                        End If
                    End If
                End If
            ElseIf PtInRect(MouseOverRect, m_Pnt.X, m_Pnt.Y) Then
                RaiseEvent MouseHover(m_lMouseOverTabIndex, CSng(m_Pnt.X), CSng(m_Pnt.Y))
                '   Copy the ControlBox Rect to temp variable
                If m_bUseControlBox Then
                    ControlBoxRect = AryTabs(m_lMouseOverTabIndex).ControlBoxRect
                    '   See if we are in the ControlBox
                    If PtInRect(ControlBoxRect, m_Pnt.X, m_Pnt.Y) Then
                        If AryTabs(iCnt).Enabled Then
                            If Not bInCtrlBox Then
                                bInCtrlBox = True
                                RaiseEvent ControlBoxEnter
                            Else
                                RaiseEvent ControlBoxHover(CSng(m_Pnt.X), CSng(m_Pnt.Y))
                            End If
                            Call DrawControlBox(xStateHover, m_lMouseOverTabIndex)
                        Else
                            Call DrawControlBox(xStateDown, m_lMouseOverTabIndex)
                        End If
                    Else
                        If AryTabs(iCnt).Enabled Then
                            If bInCtrlBox Then
                                RaiseEvent ControlBoxExit
                                bInCtrlBox = False
                            End If
                            Call DrawControlBox(xStateNormal, m_lMouseOverTabIndex)
                        Else
                            Call DrawControlBox(xStateDown, m_lMouseOverTabIndex)
                        End If
                    End If
                End If
            End If

        Case WM_MOUSELEAVE
            bInCtrl = False
            RaiseEvent MouseLeave
        
        Case WM_MOUSEWHEEL
            If m_aryTabs(m_lActiveTab).Enabled Then
                If m_bUseMouseWheelScroll = True Then
                    If wParam < 0 Then
                        '   Scrolling Up
                        RaiseEvent MouseScrollUp
                    Else
                        '   Scrolling Down
                        RaiseEvent MouseScrollDown
                    End If
                    Call ScrollTabs(wParam)
                End If
            End If

        Case WM_MOVING
            bMoving = True
            RaiseEvent Status("Form is moving...")

        Case WM_SIZING
            bMoving = False
            RaiseEvent Status("Form is sizing...")

        Case WM_EXITSIZEMOVE
            RaiseEvent Status("Finished " & IIf(bMoving, "Moving.", "Sizing."))

        Case WM_LBUTTONDOWN
            '   Handle ControlBox Closure Events
            If m_bUseControlBox Then
                '   Get our position
                Call GetCursorPos(m_Pnt)
                '   Convert coordinates
                Call ScreenToClient(m_lhWnd, m_Pnt)
                For iCnt = 0 To m_lTabCount - 1
                    '   Copy the ControlBox Rect to temp variable
                    ControlBoxRect = AryTabs(iCnt).ControlBoxRect
                    '   See if we are in the ControlBox
                    If PtInRect(ControlBoxRect, m_Pnt.X, m_Pnt.Y) Then
                        If AryTabs(iCnt).Enabled Then
                            bInCtrlBox = True
                            '   Found it, so paint it as a MouseDown for the ControlBox
                            RaiseEvent ControlBoxMouseDown(CSng(m_Pnt.X), CSng(m_Pnt.Y))
                            Call DrawControlBox(xStateDown, iCnt)
                        Else
                            Call DrawControlBox(xStateDown, iCnt)
                        End If
                        Exit For
                    End If
                Next
            End If
            '   The Remaining Events Handled by Normal UserControl Events
        
        Case WM_LBUTTONUP
            '   Only remove the tab....if we are still over it
            If m_bUseControlBox Then
                '   Get our position
                Call GetCursorPos(m_Pnt)
                '   Convert coordinates
                Call ScreenToClient(m_lhWnd, m_Pnt)
                For iCnt = 0 To m_lTabCount - 1
                    '   Copy the ControlBox Rect to temp variable
                    ControlBoxRect = AryTabs(iCnt).ControlBoxRect
                    '   See if we are in the ControlBox
                    If PtInRect(ControlBoxRect, m_Pnt.X, m_Pnt.Y) Then
                        If AryTabs(iCnt).Enabled Then
                            bInCtrlBox = True
                            '   Yep, so we must want to remove it, so call our routine
                            RaiseEvent ControlBoxMouseUp(CSng(m_Pnt.X), CSng(m_Pnt.Y))
                            Call RemoveTab(iCnt)
                        Else
                            Call DrawControlBox(xStateDown, iCnt)
                        End If
                        Exit For
                    End If
                Next
            End If
        
        Case WM_LBUTTONDBLCLK
            '   Nothing
            '   Handled by Normal UserControl Events
            
        Case WM_RBUTTONDOWN
            '   Nothing
            '   Handled by Normal UserControl Events
            
        Case WM_RBUTTONDBLCLK
            '   Nothing
            '   Handled by Normal UserControl Events
            
        Case WM_TIMER
            '   This calls the private TimerEvents but uses the SelfSubclasser
            '   to handle the processing of these events....
            Call TimerEvent

    End Select
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
    Const CODE_LEN              As Long = 204   'Use 204 Bytes to prevent Win9X GPF                                             'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
    Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
    Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                                             'Address of the owner object
    Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
    Static pCWP                 As Long                                                   'Address of the CallWindowsProc
    Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
    Dim i                       As Long                                                   'Loop index
    Dim j                       As Long                                                   'Loop index
    Dim nSubIdx                 As Long                                                   'Subclass data index
    Dim sHex                    As String                                                 'Hex code string
  
    'If it's the first time through here..
    If aBuf(1) = 0 Then
  
        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                                'Next pair of hex characters
        
        'Get API function addresses
        If Subclass_InIDE Then                                                              'If we're running in the VB IDE
            aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                               'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
            End If
        End If
        
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
        End If
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                                                    'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long
    
    i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
    Do While i >= 0                                                                       'Iterate through each element
        With sc_aSubData(i)
            If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
            End If
        End With
        i = i - 1                                                                           'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
        .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                                       'Clear the before table
        .nMsgCntA = 0                                                                       'Clear the after table
        Erase .aMsgTblB                                                                     'Erase the before table
        Erase .aMsgTblA                                                                     'Erase the after table
    End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long                                                                   'Message table entry index
    Dim nOff1   As Long                                                                   'Machine code buffer offset 1
    Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
    If uMsg = ALL_MESSAGES Then                                                           'If all messages
        nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
    Else                                                                                  'Else a specific message number
        Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
                Exit Sub                                                                        'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
                Exit Sub                                                                        'Bail
            End If
        Loop                                                                                'Next entry
        nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
    End If
    If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
        nOff1 = PATCH_04                                                                    'Offset to the Before table
        nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
    Else                                                                                  'Else after
        nOff1 = PATCH_08                                                                    'Offset to the After table
        nOff2 = PATCH_09                                                                    'Offset to the After table entry count
    End If
    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
        nMsgCnt = 0                                                                         'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
            nEntry = PATCH_05                                                                 'Patch the before table message count location
        Else                                                                                'Else after
            nEntry = PATCH_09                                                                 'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
    Else                                                                                  'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                           'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
                Exit Do                                                                         'Bail
            End If
        Loop                                                                                'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                                'If we're searching not adding
                    Exit Function                                                                 'Found
                End If
            ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
                If bAdd Then                                                                    'If we're adding
                    Exit Function                                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                                     'Decrement the index
    Loop
    
    If Not bAdd Then
        Debug.Assert False                                                                  'hWnd not found, programmer error
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
'*******************************************************************************
'   End Subclasser Section - Start Usercontrol Sections
'*******************************************************************************

Public Sub About()
Attribute About.VB_UserMemId = -552
  fAbout.Show vbModal
End Sub

Public Property Get ActiveTabBackEndColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ActiveTabBackEndColor = m_lActiveTabBackEndColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabBackEndColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ActiveTabBackEndColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lActiveTabBackEndColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "ActiveTabBackEndColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabBackEndColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ActiveTabBackStartColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ActiveTabBackStartColor = m_lActiveTabBackStartColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabBackStartColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ActiveTabBackStartColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lActiveTabBackStartColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "ActiveTabBackStartColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabBackStartColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ActiveTabFont() As StdFont

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set ActiveTabFont = m_oActiveTabFont
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabFont", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Set ActiveTabFont(oNewFont As StdFont)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set m_oActiveTabFont = oNewFont
    '   Redraw
    Refresh
    PropertyChanged "ActiveTabFont"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabFont", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ActiveTabForeColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ActiveTabForeColor = m_lActiveTabForeColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ActiveTabForeColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lActiveTabForeColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "ActiveTabForeColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ActiveTab() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    '   Active tab index
    ActiveTab = m_lActiveTab

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTab", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ActiveTabHeight() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ActiveTabHeight = m_lActiveTabHeight
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabHeight", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ActiveTabHeight(ByVal lNewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lActiveTabHeight = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "ActiveTabHeight"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTabHeight", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ActiveTab(ByVal lNewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If (lNewValue < 0) Or (lNewValue >= m_lTabCount) Then
        '   Handle Under/Over Ranged Values by wrapping
        If (lNewValue < 0) Then lNewValue = m_lTabCount
        If (lNewValue >= m_lTabCount) Then lNewValue = 0
    End If
    '   If already we are on the same tab (this is important or else all
    '   the contained controls for active tab will be moved to -75000 and so...
    If lNewValue = m_lActiveTab Then Exit Property
    m_bCancelFlag = False
    '   Raise event and confirm that the user want to allow the tab switch
    RaiseEvent BeforeTabSwitch(lNewValue, m_bCancelFlag)
    '   If user set the cancel flag in the BeforeTabSwitch event
    If m_bCancelFlag Then Exit Property
    '   Show/Hide Controls for active tab
    Call HandleContainedControls(lNewValue)
    '   Store current tab in last active tab
    m_lLastActiveTab = m_lActiveTab
    '   Now ste the New Current Tab
    m_lActiveTab = lNewValue
    '   Set the MouseOver TabIndex to the New Tab
    m_lMouseOverTabIndex = m_lActiveTab
    '   Now draw the tabs with changed state
    Call DrawBackground
    Call DrawTabs
    PropertyChanged "ActiveTab"
    '   Redraw
    Refresh
    UserControl.Refresh
    RaiseEvent TabSwitch(m_lLastActiveTab)
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ActiveTab", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Sub AddControl(Optional ByRef Ctrl As Control, Optional lhWnd As Long = -1, _
    Optional ByVal lLeft As Long = -1, Optional ByVal lTop As Long = -1, _
    Optional ByVal lWidth As Long = -1, Optional ByVal lHeight As Long = -1, _
    Optional ByVal lTabIndex As Long = -1)
    Dim lpRect As RECT
    Dim mPt As POINTAPI
    
    With UserControl
        If (lTabIndex <> -1) And (lTabIndex <> m_lActiveTab) Then
            ActiveTab = lTabIndex
        End If
        On Error Resume Next
        If ((Ctrl Is Nothing) Or (Not Ctrl.hWnd)) And (lhWnd <> -1) Then
            If Not Ctrl Is Nothing Then
                SetParent Ctrl.hWnd, UserControl.hWnd
                If lLeft <> -1 Then
                    Ctrl.Left = lLeft
                End If
                If lTop <> -1 Then
                    Ctrl.Top = lTop
                End If
                If lWidth <> -1 Then
                    Ctrl.Width = lWidth
                End If
                If lHeight <> -1 Then
                    Ctrl.Height = lHeight
                End If
                Ctrl.Visible = True
            Else
                SetParent lhWnd, UserControl.hWnd
                SetWindowPos lhWnd, 0, lLeft, lTop, lWidth, lHeight, SWP_SHOWWINDOW
            End If
        Else
            SetParent Ctrl.hWnd, UserControl.hWnd
            If lLeft <> -1 Then
                Ctrl.Left = lLeft
            End If
            If lTop <> -1 Then
                Ctrl.Top = lTop
            End If
            If lWidth <> -1 Then
                Ctrl.Width = lWidth
            End If
            If lHeight <> -1 Then
                Ctrl.Height = lHeight
            End If
            Ctrl.Visible = True
        End If
    End With
End Sub

Private Property Get Ambient() As Object

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set Ambient = UserControl.Ambient
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.Ambient", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Sub APIFillRectByCoords(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, lColor As Long)
    
    Dim NewBrush As Long
    Dim tmpRect As RECT
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Convert the Color to RGB
    lColor = GetRGBFromOLE(lColor)
    '   Create a New Brush with the Color Passed
    NewBrush = CreateSolidBrush(lColor)
    '   Set the Coords into a RECT Strcture
    SetRect tmpRect, X, Y, X + W, Y + H
    '   Draw a Filled Rect
    Call FillRect(UserControl.hdc, tmpRect, NewBrush)
    '   Delete the New Brush
    Call DeleteObject(NewBrush)
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.APIFillRectByCoords", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub APILine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)
        
    Dim pt As POINTAPI
    Dim hPen As Long, hPenOld As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Convert the Color to RGB
    lColor = GetRGBFromOLE(lColor)
    '   Create a New Pen
    hPen = CreatePen(0, 1, lColor)
    '   Select the New Pen into the DC
    hPenOld = SelectObject(UserControl.hdc, hPen)
    '   Move to the Fist Point
    MoveToEx UserControl.hdc, X1, Y1, pt
    '   Draw the Line Segment
    LineTo UserControl.hdc, X2, Y2
    '   Select the Old Pen back into DC
    SelectObject UserControl.hdc, hPenOld
    '   Delete the New Pen
    DeleteObject hPen
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.APILine", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Function APIRectangle(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional lColor As OLE_COLOR = -1) As Long
    Dim hPen As Long, hPenOld As Long
    Dim pt As POINTAPI
    
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler
    
    '   Convert the Color to RGB from System Colors
    lColor = GetRGBFromOLE(lColor)
    '   Now Create a new Pen
    hPen = CreatePen(0, 1, lColor)
    '   Select it in to the DC
    hPenOld = SelectObject(UserControl.hdc, hPen)
    '   Move to the Starting Position
    MoveToEx UserControl.hdc, X, Y, pt
    '   Draw the Line Segments
    LineTo UserControl.hdc, X + W, Y
    LineTo UserControl.hdc, X + W, Y + H
    LineTo UserControl.hdc, X, Y + H
    LineTo UserControl.hdc, X, Y
    '   Restore the Old Pen
    SelectObject UserControl.hdc, hPenOld
    '   Delete the New Pen
    DeleteObject hPen

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "AeroTab.APIRectangle", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Property Get AryTabs(iIndex As Long) As TabInfo
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If iIndex < LBound(m_aryTabs) Then iIndex = LBound(m_aryTabs)
    If iIndex > UBound(m_aryTabs) Then iIndex = UBound(m_aryTabs)
    '   Private Structure Properties for the Tabs
    AryTabs = m_aryTabs(iIndex)

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.AryTabs", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Property Let AryTabs(iIndex As Long, utNewValue As TabInfo)
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    '   Private Structure Properties for the Tabs
    m_aryTabs(iIndex) = utNewValue

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.AryTabs", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Property Get BackColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    BackColor = UserControl.BackColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.BackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Property Let BackColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UserControl.BackColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "BackColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.BackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get BottomRightInnerBorderColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    BottomRightInnerBorderColor = m_lBottomRightInnerBorderColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.BottomRightInnerBorderColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let BottomRightInnerBorderColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lBottomRightInnerBorderColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "BottomRightInnerBorderColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.BottomRightInnerBorderColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Property Get bUserMode() As Boolean
    On Error Resume Next
    '   Used to prevent an error which occurs when a
    '   form with this control gets unloaded
    '   This is strange but the control gets a "GotFocus" event
    '   sometimes when the container form is unloaded
    bUserMode = Ambient.UserMode
End Property

Public Sub Cls()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    UserControl.Cls
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.Cls", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub CopyTabImagesFromImageList(ByRef oIml As Object)
    '   Copy From a standard Image List to current tabs images
    Dim iTmp As Long
    
    '   If the number of images is less than number of tabs error may occur
    '   provided this, the control will only paint upto the number of Images
    '   and exit the sub
    On Error GoTo Finally
    For iTmp = 0 To UBound(m_aryTabs)
        '   Free Existing Picture
        Set m_aryTabs(iTmp).TabPicture = Nothing
        Set m_aryTabs(iTmp).TabPicture = oIml.ListImages(iTmp + 1).Picture
    Next
Finally:
    '   Redraw
    Refresh
End Sub

Public Property Get DisabledTabBackColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    DisabledTabBackColor = m_lDisabledTabBackColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DisabledTabBackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let DisabledTabBackColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lDisabledTabBackColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "DisabledTabBackColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DisabledTabBackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get DisabledTabForeColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    DisabledTabForeColor = m_lDisabledTabForeColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DisabledTabForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let DisabledTabForeColor(ByVal lNewValue As OLE_COLOR)
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    m_lDisabledTabForeColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "DisabledTabForeColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DisabledTabForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Sub DrawBackground()
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Erases the Previous Background and draws a new one
    '   (does not draw the tabs), Called from Refresh Event
    '   of the UserControl
    '   Get the current cached properties
    Call pGetCachedProperties
    Select Case m_enmStyle
        Case xStyleTabbedDialog:
            Call DrawBackgroundTabbedDialog
        Case xStylePropertyPages:
            Call DrawBackgroundPropertyPages
    End Select
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawBackground", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub DrawBackgroundPropertyPages()
    '   This is called from the DrawBackground function if the
    '   tab style is set to "Property Page"
    Dim iTmp As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Call DrawBackgroundTabbedDialog
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawBackgroundPropertyPages", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub DrawBackgroundTabbedDialog()
    '   This is called from the DrawBackground function if the
    '   tab style is set to "Tabbed Dialog"
    Dim iTmp As Long
    Dim utRect As RECT
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Get the larger of the active tab height and inactive tab height
    iTmp = IIf(m_lActiveTabHeight > m_lInActiveTabHeight, m_lActiveTabHeight, m_lInActiveTabHeight)
    UserControl.Cls   'clear control
    '   Fill background color based on tab's enabled property
    If m_bEnabled Then
        If AryTabs(m_lActiveTab).Enabled Then
            BackColor = m_lActiveTabBackEndColor
        Else
            BackColor = m_lDisabledTabBackColor
        End If
    Else
        BackColor = m_lDisabledTabBackColor
    End If
    '   Draw inner shadow (left)
    APILine 0, iTmp, 0, m_lScaleHeight - 1, m_lOuterBorderColor
    '   Draw inner shadow (right)
    APILine m_lScaleWidth - 1, iTmp, m_lScaleWidth - 1, m_lScaleHeight - 1, m_lOuterBorderColor
     '   Draw inner shadow (bottom)
    APILine 0, m_lScaleHeight - 1, m_lScaleWidth, m_lScaleHeight - 1, m_lOuterBorderColor
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawBackgroundTabbedDialog", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub DrawControlBox(Optional ByVal xState As xStateEnum, Optional ByVal lIndex As Long = -1)
    Dim iCnt As Long
    Dim utTabInfo As TabInfo
    Dim lStartColor As Long
    Dim lEndColor As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
        
    If lIndex = -1 Then
        ForeColor = m_lOuterBorderColor
        For iCnt = 0 To m_lTabCount - 1
            '   Fetch local copy
            utTabInfo = AryTabs(iCnt)
            If (m_bEnabled) Then
                If AryTabs(iCnt).Enabled Then
                    If iCnt = m_lActiveTab Then
                        lStartColor = m_lActiveTabBackStartColor
                        lEndColor = m_lActiveTabBackEndColor
                    Else
                        lStartColor = m_lInActiveTabBackStartColor
                        lEndColor = m_lInActiveTabBackEndColor
                    End If
                Else
                    If iCnt = m_lActiveTab Then
                        lStartColor = OffsetColor(m_lActiveTabBackEndColor, -&H30)
                        lEndColor = OffsetColor(m_lActiveTabBackEndColor, -&H30)
                    Else
                        lStartColor = OffsetColor(m_lInActiveTabBackEndColor, -&H30)
                        lEndColor = OffsetColor(m_lInActiveTabBackEndColor, -&H30)
                    End If
                End If
            Else
                If iCnt = m_lActiveTab Then
                    lStartColor = OffsetColor(m_lActiveTabBackEndColor, -&H30)
                    lEndColor = OffsetColor(m_lActiveTabBackEndColor, -&H30)
                Else
                    lStartColor = OffsetColor(m_lInActiveTabBackEndColor, -&H30)
                    lEndColor = OffsetColor(m_lInActiveTabBackEndColor, -&H30)
                End If
            End If
            With utTabInfo.ControlBoxRect
                '   Draw the controlbox frame
                RoundRect UserControl.hdc, .Left, .Top, .Right, .Bottom, 2, 2
                '   Fill the gradient inside
                pFillCurvedGradient .Left + 1, .Top + 1, .Right - 1, .Bottom - 2, lStartColor, lEndColor, , True, True
                '   Now draw the "/" of the X
                APILine .Left + 3, .Bottom - 4, .Right - 3, .Top + 2, OffsetColor(m_lOuterBorderColor, -&HC0)
                '   Now draw the "\" of the X
                APILine .Left + 3, .Top + 3, .Right - 3, .Bottom - 3, OffsetColor(m_lOuterBorderColor, -&HC0)
            End With
        Next
    Else
        Select Case xState
            Case xStateNormal
                    If lIndex = m_lActiveTab Then
                        lStartColor = m_lActiveTabBackStartColor
                        lEndColor = m_lActiveTabBackEndColor
                    Else
                        lStartColor = m_lInActiveTabBackStartColor
                        lEndColor = m_lInActiveTabBackEndColor
                    End If
            Case xStateHover
                If lIndex = m_lActiveTab Then
                    lStartColor = m_lActiveTabBackStartColor
                    lEndColor = OffsetColor(m_lActiveTabBackEndColor, -&H15)
                Else
                    lStartColor = m_lInActiveTabBackStartColor
                    lEndColor = OffsetColor(m_lInActiveTabBackEndColor, -&H15)
                End If
            Case xStateDown
                If lIndex = m_lActiveTab Then
                    lStartColor = OffsetColor(m_lActiveTabBackEndColor, -&H30)
                    lEndColor = OffsetColor(m_lActiveTabBackEndColor, -&H30)
                Else
                    lStartColor = OffsetColor(m_lInActiveTabBackEndColor, -&H30)
                    lEndColor = OffsetColor(m_lInActiveTabBackEndColor, -&H30)
                End If
        End Select
        '   Fetch local copy
        utTabInfo = AryTabs(lIndex)
        With utTabInfo.ControlBoxRect
            '   Draw the controlbox frame
            RoundRect UserControl.hdc, .Left, .Top, .Right, .Bottom, 2, 2
            '   Fill the gradient inside
            pFillCurvedGradient .Left + 1, .Top + 1, .Right - 1, .Bottom - 2, lStartColor, lEndColor, , True, True
            '   Now draw the "/" of the X
            APILine .Left + 3, .Bottom - 4, .Right - 3, .Top + 2, OffsetColor(m_lOuterBorderColor, -&HC0)
            '   Now draw the "\" of the X
            APILine .Left + 3, .Top + 3, .Right - 3, .Bottom - 3, OffsetColor(m_lOuterBorderColor, -&HC0)
        End With
    End If

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawControlBox", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub DrawImage(ByVal lDestHDC As Long, ByVal lhBmp As Long, ByVal lTransColor As Long, ByVal iLeft As Integer, ByVal iTop As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer)
    Dim lhDC As Long
    Dim lhBmpOld As Long
    Dim utBmp As BITMAP
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Create a Compatible DC
    lhDC = CreateCompatibleDC(lDestHDC)
    '   Select the Bitmap into the New Compatible DC using its handle
    lhBmpOld = SelectObject(lhDC, lhBmp)
    '   Get the Objects Properties...in this case the Bitmaps Props
    Call GetObjectA(lhBmp, Len(utBmp), utBmp)
    '   Blit this onto the DC (Tab)
    Call TransparentBlt(lDestHDC, iLeft, iTop, iWidth, iHeight, lhDC, 0, 0, utBmp.bmWidth, utBmp.bmHeight, lTransColor)
    '   Select the Old Bitmap
    Call SelectObject(lhDC, lhBmpOld)
    '   Delete the New DC
    DeleteDC (lhDC)
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawImage", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub DrawTabs()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Erases the previous tabs and draws each one by one
    'Get the current cached properties
    Call pGetCachedProperties
    With Me
        Select Case .TabStyle
        Case xStyleTabbedDialog:
            Call DrawTabsTabbedDialog
        Case xStylePropertyPages:
            Call DrawTabsPropertyPages
        End Select
    End With
    If m_bUseControlBox Then
        Call DrawControlBox
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawTabs", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub DrawTabsPropertyPages()
    '   This is called from the DrawTabs function if the tab style
    '   is Property Page
    Dim iCnt As Long
    Dim iTabWidth As Long
    Dim utFontRect As RECT
    Dim sTmp As String
    Dim utTabInfo As TabInfo
    Dim iTmpW As Long
    Dim iTmpH As Long
    Dim iAdjustedIconSize As Long
    Dim iTmpX As Long
    Dim iTmpY As Long
    Dim iTmpHeight As Long
    Dim iOrigLeft As Long
    Dim iOrigRight As Long
    Dim lpRect As RECT
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Set the Inactive tab's font as current font (since the TextWidth function
    '   Will use the current font's size)
    Set Font = InActiveTabFont
    '   Store the larger height in tmp var
    iTmpHeight = IIf(m_lActiveTabHeight > m_lInActiveTabHeight, m_lActiveTabHeight, m_lInActiveTabHeight)
    '   Initialize the clickable items
    For iCnt = 0 To m_lTabCount - 1
        utTabInfo = AryTabs(iCnt)
        sTmp = Replace$(utTabInfo.Caption, "&&", "&")
        If InStr(1, sTmp, "&") Then
            'if still there is one '&' in the string then reduce the width by one more character (since the '&' will be conveted into an underline when painted)
            sTmp = Mid$(sTmp, 1, Len(sTmp) - 1)
        End If
        If utTabInfo.TabPicture Is Nothing Then
            '   Get tab width acc to the text size and border
            iTabWidth = TextWidth(sTmp) + iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2
        Else
            If iTmpHeight - 2 < m_lIconSize Then    '-2 for borders
                '   Here we adjust the size of the icon if it does
                '   not fit into current tab
                iAdjustedIconSize = iTmpHeight - 2
            Else
                iAdjustedIconSize = m_lIconSize
            End If
            '   Get tab width based on the text size, border and Image
            iTabWidth = TextWidth(sTmp) + (iPROP_PAGE_BORDER_AND_TEXT_DISTANCE * 2) + iAdjustedIconSize + 1
        End If
        '   Following adjustments are used in case of property pages only.
        '   We must shift the left (+2) or (-2) to make it look like
        '   standard XP property pages
        With utTabInfo.ClickableRect
            If iCnt = 0 And iCnt <> m_lActiveTab Then
                .Left = iPROP_PAGE_INACTIVE_TOP
                .Right = .Left + iTabWidth - iPROP_PAGE_INACTIVE_TOP + 1
            Else
                If iCnt = 0 Then
                  .Left = 0
                Else
                    If iCnt = m_lActiveTab Or iCnt = m_lActiveTab + 1 Then
                        .Left = AryTabs(iCnt - 1).ClickableRect.Right
                    Else
                        '   1 pixel distance between property pages (in XP)
                        .Left = AryTabs(iCnt - 1).ClickableRect.Right + 1
                    End If
                End If
                .Right = .Left + iTabWidth
            End If
            If m_bUseControlBox Then
                If iCnt = m_lActiveTab Then
                    .Right = .Right + iPROP_CONTROLBOX * 1.75
                Else
                    .Right = .Right + iPROP_CONTROLBOX + 3
                End If
            End If
            If iCnt = m_lActiveTab Then
                If m_lActiveTabHeight > m_lInActiveTabHeight Then
                    .Top = 0
                Else
                    .Top = m_lInActiveTabHeight - m_lActiveTabHeight
                End If
                .Bottom = .Top + m_lActiveTabHeight
            Else
                If m_lInActiveTabHeight > m_lActiveTabHeight Then
                    .Top = 0
                    .Bottom = .Top + m_lInActiveTabHeight
                Else
                    .Top = m_lActiveTabHeight - m_lInActiveTabHeight
                    .Bottom = .Top + m_lInActiveTabHeight
                End If
            End If
        End With
        '   Store the ControlBox values for hit testing
        With utTabInfo.ControlBoxRect
            .Left = utTabInfo.ClickableRect.Right - 13
            .Top = utTabInfo.ClickableRect.Top + 6
            .Right = utTabInfo.ClickableRect.Right - 3
            .Bottom = utTabInfo.ClickableRect.Top + 16
        End With
        '   Assign the new tab info to the existing one
        AryTabs(iCnt) = utTabInfo
    Next
    '   Fill the tab strip with TabStripBackColor (customizable... so that tab's can easily blend with the background)
    APIFillRectByCoords 0, 0, m_lScaleWidth, iTmpHeight, TabStripBackColor
    '   Now Draw Each Tab
    For iCnt = 0 To m_lTabCount - 1
        '   Get Local copy
        utTabInfo = AryTabs(iCnt)
        With utTabInfo.ClickableRect
            '   If we are drawing the active tab
            If iCnt = m_lActiveTab Then
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        Call pFillCurvedGradientR(utTabInfo.ClickableRect, m_lActiveTabBackStartColor, m_lActiveTabBackEndColor)
                    Else
                        Call pFillCurvedGradientR(utTabInfo.ClickableRect, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                    End If
                Else
                    Call pFillCurvedGradientR(utTabInfo.ClickableRect, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                End If
                '   Top line
                APILine .Left, .Top, .Right, .Top, m_lOuterBorderColor
                '   Right line
                APILine .Right - 1, .Top, .Right - 1, .Bottom + 1, m_lOuterBorderColor
                '   Left line
                APILine .Left, .Top, .Left, .Bottom + 1, m_lOuterBorderColor
                Set Font = ActiveTabFont 'set the font
                '   Set fore ground color
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        ForeColor = m_lActiveTabForeColor
                    Else
                        ForeColor = m_lDisabledTabForeColor
                    End If
                Else
                    ForeColor = m_lDisabledTabForeColor
                End If
            ElseIf iCnt = m_lActiveTab - 1 Then
                '   If we are drawing tab just b4 active tab, then
                If (m_bEnabled) And utTabInfo.Enabled Then
                    Call pFillCurvedGradient(.Left, .Top + 1, .Right, .Top + ((.Bottom - .Top) \ 2) + 1, ShiftColorOXP(m_lInActiveTabBackStartColor, 100), m_lInActiveTabBackStartColor)
                    Call pFillCurvedGradient(.Left, .Top + ((.Bottom - .Top) \ 2) + 1, .Right, .Bottom, ShiftColorOXP(m_lInActiveTabBackEndColor, 70), m_lInActiveTabBackEndColor)
                Else
                    Call pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                End If
                '   Following adjustments are needed if the inactive tab's
                '   height is more than active tab's height (the tab's corner
                '   should be properly rounded)
                If m_lInActiveTabHeight > m_lActiveTabHeight Then
                    '   Right line
                    APILine .Right - 1, .Top, .Right - 1, .Bottom + 1, m_lOuterBorderColor
                End If
                '   Top line
                APILine .Left, .Top, .Right, .Top, m_lOuterBorderColor
                APILine .Left + 1, .Top + 1, .Right - 1, .Top + 1, vbWhite
                '   Bottom line
                APILine .Left, .Bottom, .Right + 1, .Bottom, m_lOuterBorderColor
                '   Left Line
                APILine .Left, .Top, .Left, .Bottom + 1, m_lOuterBorderColor
                APILine .Left + 1, .Top + 1, .Left + 1, .Bottom, vbWhite
                '   Set the font
                Set Font = InActiveTabFont
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        ForeColor = m_lInActiveTabForeColor
                    Else
                        ForeColor = m_lDisabledTabForeColor
                    End If
                Else
                    ForeColor = m_lDisabledTabForeColor
                End If
            ElseIf iCnt = m_lActiveTab + 1 Then
                '   If we are drawing tab just after the active tab, then
                If (m_bEnabled) And utTabInfo.Enabled Then
                    Call pFillCurvedGradient(.Left, .Top + 1, .Right, .Top + ((.Bottom - .Top) \ 2) + 1, ShiftColorOXP(m_lInActiveTabBackStartColor, 100), m_lInActiveTabBackStartColor)
                    Call pFillCurvedGradient(.Left, .Top + ((.Bottom - .Top) \ 2) + 1, .Right, .Bottom, ShiftColorOXP(m_lInActiveTabBackEndColor, 70), m_lInActiveTabBackEndColor)
                Else
                    Call pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                End If
                '   Following adjustments are needed if the inactive tab's
                '   height is more than active tab's height (the tab's corner
                '   should be properly rounded)
                If m_lInActiveTabHeight > m_lActiveTabHeight Then
                    '   Left Line
                    APILine .Left, .Top, .Left, .Bottom + 1, m_lOuterBorderColor
                End If
                '   Top line
                APILine .Left, .Top, .Right, .Top, m_lOuterBorderColor
                APILine .Left + 1, .Top + 1, .Right - 1, .Top + 1, vbWhite
                '   Right line
                APILine .Right, .Top, .Right, .Bottom, m_lOuterBorderColor
                APILine .Right - 1, .Top + 1, .Right - 1, .Bottom - 1, vbWhite
                '   Bottom line
                APILine .Left, .Bottom, .Right + 1, .Bottom, m_lOuterBorderColor
                '   Set the font
                Set Font = InActiveTabFont
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        ForeColor = m_lInActiveTabForeColor
                    Else
                        ForeColor = m_lDisabledTabForeColor
                    End If
                Else
                    ForeColor = m_lDisabledTabForeColor
                End If
            Else
                '   Other non active tab (must draw full curves on both
                '   the sides)
                If (m_bEnabled) And utTabInfo.Enabled Then
                    Call pFillCurvedGradient(.Left, .Top + 1, .Right, .Top + ((.Bottom - .Top) \ 2) + 1, ShiftColorOXP(m_lInActiveTabBackStartColor, 100), m_lInActiveTabBackStartColor)
                    Call pFillCurvedGradient(.Left, .Top + ((.Bottom - .Top) \ 2) + 1, .Right, .Bottom, ShiftColorOXP(m_lInActiveTabBackEndColor, 70), m_lInActiveTabBackEndColor)
                Else
                    Call pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                End If
                '   Top line
                APILine .Left, .Top, .Right, .Top, m_lOuterBorderColor
                APILine .Left + 1, .Top + 1, .Right - 1, .Top + 1, vbWhite
                '   Right line
                APILine .Right, .Top, .Right, .Bottom, m_lOuterBorderColor
                APILine .Right - 1, .Top + 1, .Right - 1, .Bottom - 1, vbWhite
                '   Bottom line
                APILine .Left, .Bottom, .Right + 1, .Bottom, m_lOuterBorderColor
                '   Left Line
                APILine .Left, .Top, .Left, .Bottom + 1, m_lOuterBorderColor
                APILine .Left + 1, .Top + 1, .Left + 1, .Bottom, vbWhite
                Set Font = InActiveTabFont 'set the font
                '   set font color
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        ForeColor = m_lInActiveTabForeColor
                    Else
                        ForeColor = m_lDisabledTabForeColor
                    End If
                Else
                    ForeColor = m_lDisabledTabForeColor
                End If
            End If
            '   If its left most tab then adjust the bottom line
            If iCnt = 0 Then
              ' Bottom line
              APILine .Left - 2, .Bottom, .Left, .Bottom, m_lOuterBorderColor
            End If
            '   Do the adjustments for the border
            utFontRect.Left = .Left + 2
            utFontRect.Top = .Top + 4
            utFontRect.Bottom = .Bottom
            utFontRect.Right = .Right - 2
            If m_bUseControlBox Then
                OffsetRect utFontRect, -2, 0
            End If
            sTmp = utTabInfo.Caption
            If Not utTabInfo.TabPicture Is Nothing Then
                If iTmpHeight - 2 < m_lIconSize Then    '-2 for borders
                  ' Here we adjust the size of the icon if it does not
                  ' fit into current tab
                  iAdjustedIconSize = iTmpHeight - 2
                Else
                  iAdjustedIconSize = m_lIconSize
                End If
                iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
                Select Case PictureAlign
                    Case xAlignLeftEdge, xAlignLeftOfCaption:
                        If utTabInfo.TabPicture.Type = vbPicTypeBitmap And UseMaskColor Then
                            Call DrawImage(m_lhDC, utTabInfo.TabPicture.Handle, GetRGBFromOLE(PictureMaskColor), utFontRect.Left + 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        Else
                            Call pPaintPicture(utTabInfo.TabPicture, utFontRect.Left + 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        End If
                        '   Shift the text to be drawn after the picture
                        utFontRect.Left = (utFontRect.Left + iAdjustedIconSize + 6) - iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
                        '   Call the API for the text drawing
                        DrawText m_lhDC, sTmp, -1, utFontRect, DT_VCENTER Or DT_SINGLELINE Or DT_CENTER
                        '   Revert the changes so that the focus rectangle
                        '   can be drawn for the whole tab's clickable area
                        utFontRect.Left = (utFontRect.Left - iAdjustedIconSize - 6) + iPROP_PAGE_BORDER_AND_TEXT_DISTANCE
                    Case xAlignRightEdge, xAlignRightOfCaption:
                        If utTabInfo.TabPicture.Type = vbPicTypeBitmap And UseMaskColor Then
                            Call DrawImage(m_lhDC, utTabInfo.TabPicture.Handle, GetRGBFromOLE(PictureMaskColor), utFontRect.Right - iAdjustedIconSize - 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        Else
                            Call pPaintPicture(utTabInfo.TabPicture, utFontRect.Right - iAdjustedIconSize - 2, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        End If
                        '   Shift the text to be drawn after the picture
                        utFontRect.Right = (utFontRect.Right + 1) - iAdjustedIconSize - 6
                        '   Call the API for the text drawing
                        DrawText m_lhDC, sTmp, -1, utFontRect, DT_VCENTER Or DT_SINGLELINE Or DT_CENTER
                        '   Revert the changes so that the focus rectangle
                        '   can be drawn for the whole tab's clickable area
                        utFontRect.Right = (utFontRect.Right - 1) + iAdjustedIconSize + 6
                End Select
    
            Else
                '   Call the API for the text drawing
                DrawText m_lhDC, sTmp, -1, utFontRect, DT_VCENTER Or DT_SINGLELINE Or DT_CENTER
            End If
            '   Only if in the run mode
            If bUserMode Then
                If iCnt = m_lActiveTab And m_IsFocused And ShowFocusRect Then
                    If m_bUseControlBox Then
                        OffsetRect utFontRect, 2, 0
                    End If
                    '   Draw focus rectangle
                    Call DrawFocusRect(m_lhDC, utFontRect)
                End If
            End If
        End With
    Next
    '   Draw the line in the empty area after all the property pages heads
    '   are drawn
    APILine AryTabs(m_lTabCount - 1).ClickableRect.Right, AryTabs(m_lTabCount - 1).ClickableRect.Bottom, m_lScaleWidth, AryTabs(m_lTabCount - 1).ClickableRect.Bottom, m_lOuterBorderColor
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawTabsPropertyPages", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub DrawTabsTabbedDialog()
    ' This is called from the DrawTabs function if the tab style
    ' is Tabbed Dialog
    Dim iCnt As Long
    Dim iTabWidth As Long
    Dim utFontRect As RECT
    Dim sTmp As String
    Dim utTabInfo As TabInfo
    Dim iTmpW As Long
    Dim iTmpH As Long
    Dim iTmpX As Long
    Dim iTmpY As Long
    Dim iOrigLeft As Long
    Dim iOrigRight As Long
    Dim iAdjustedIconSize As Long
    Dim lColor As Long
    Dim lpRect As RECT

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Remember iTabWidth is an Long ...
    '   So the result is automatically rounded
    iTabWidth = m_lScaleWidth / m_lTabCount
    '   Initialize the clickable items
    For iCnt = 0 To m_lTabCount - 1
        utTabInfo = AryTabs(iCnt)
        '   No need to calculate the text size(like in property pages)....
        '   since this is a tabbed dialog style
        With utTabInfo.ClickableRect
            .Left = iCnt * iTabWidth
            .Right = .Left + iTabWidth
            If iCnt = m_lActiveTab Then
                If m_lActiveTabHeight > m_lInActiveTabHeight Then
                    .Top = 0
                Else
                    .Top = m_lInActiveTabHeight - m_lActiveTabHeight
                End If
                .Bottom = .Top + m_lActiveTabHeight
            Else
                If m_lInActiveTabHeight > m_lActiveTabHeight Then
                    .Top = 0
                Else
                    .Top = m_lActiveTabHeight - m_lInActiveTabHeight
                End If
                .Bottom = .Top + m_lInActiveTabHeight
            End If
        End With
        If iCnt = m_lTabCount - 1 Then
            '   If the last tab is shorter or longer than the usual size..
            '   then adjust it to perfect size
            utTabInfo.ClickableRect.Right = m_lScaleWidth - 1
        End If
        '   Store the ControlBox values for hit testing
        With utTabInfo.ControlBoxRect
            .Left = utTabInfo.ClickableRect.Right - 13
            .Top = utTabInfo.ClickableRect.Top + 6
            .Right = utTabInfo.ClickableRect.Right - 3
            .Bottom = utTabInfo.ClickableRect.Top + 16
        End With
        AryTabs(iCnt) = utTabInfo
    Next
    '   Added to prevent lines etc (we are filling the tab strip with
    '   the tab strip color)
    APIFillRectByCoords 0, 0, m_lScaleWidth, IIf(m_lActiveTabHeight > m_lInActiveTabHeight, m_lActiveTabHeight, m_lInActiveTabHeight), TabStripBackColor
    '   Now Draw Each Tab
    For iCnt = 0 To m_lTabCount - 1
        '   Fetch local copy
        utTabInfo = AryTabs(iCnt)
        With utTabInfo.ClickableRect
            '   If we are drawing active tab then
            If iCnt = m_lActiveTab Then
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        Call pFillCurvedGradientR(utTabInfo.ClickableRect, m_lActiveTabBackStartColor, m_lActiveTabBackEndColor)
                    Else
                        Call pFillCurvedGradientR(utTabInfo.ClickableRect, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                    End If
                Else
                    Call pFillCurvedGradientR(utTabInfo.ClickableRect, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                End If
                '   Top line
                APILine .Left, .Top, .Right, .Top, m_lOuterBorderColor
                '   Right line
                APILine .Right, .Top, .Right, .Bottom + 1, m_lOuterBorderColor
                '   Left line
                APILine .Left, .Top, .Left, .Bottom + 1, m_lOuterBorderColor
                Set Font = ActiveTabFont       'set the font
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        ForeColor = m_lActiveTabForeColor
                    Else
                        ForeColor = m_lDisabledTabForeColor
                    End If
                Else
                    ForeColor = m_lDisabledTabForeColor
                End If
            Else      ' We are drawing inactive tab
                If (m_bEnabled) And utTabInfo.Enabled Then
                    Call pFillCurvedGradient(.Left, .Top + 1, .Right, .Top + ((.Bottom - .Top) \ 2) + 1, ShiftColorOXP(m_lInActiveTabBackStartColor, 100), m_lInActiveTabBackStartColor)
                    Call pFillCurvedGradient(.Left, .Top + ((.Bottom - .Top) \ 2) + 1, .Right, .Bottom, ShiftColorOXP(m_lInActiveTabBackEndColor, 70), m_lInActiveTabBackEndColor)
                Else
                    Call pFillCurvedGradient(.Left, .Top, .Right, .Bottom, m_lDisabledTabBackColor, m_lDisabledTabBackColor)
                End If
                '   Top line
                APILine .Left, .Top, .Right, .Top, m_lOuterBorderColor
                APILine .Left + 1, .Top + 1, .Right - 1, .Top + 1, vbWhite
                '   Right line
                APILine .Right, .Top, .Right, .Bottom, m_lOuterBorderColor
                APILine .Right - 1, .Top + 1, .Right - 1, .Bottom, vbWhite
                '   Bottom line
                APILine .Left, .Bottom, .Right + 1, .Bottom, m_lOuterBorderColor
                '   Left Line
                APILine .Left, .Top, .Left, .Bottom + 1, m_lOuterBorderColor
                APILine .Left + 1, .Top + 1, .Left + 1, .Bottom, vbWhite
                '   Set the font
                Set Font = InActiveTabFont
                '   Set foreground color
                If (m_bEnabled) Then
                    If utTabInfo.Enabled Then
                        ForeColor = m_lInActiveTabForeColor
                    Else
                        ForeColor = m_lDisabledTabForeColor
                    End If
                Else
                    ForeColor = m_lDisabledTabForeColor
                End If
            End If
            '   Do the adjustments for the border
            utFontRect.Left = .Left + 3
            utFontRect.Top = .Top + 3
            utFontRect.Bottom = .Bottom
            utFontRect.Right = .Right - 2
            If Not utTabInfo.TabPicture Is Nothing Then
                If utFontRect.Top + m_lIconSize > utFontRect.Bottom + 1 Then '+1 for minor adjustments
                    '   Adjust if going out of current tab's bottom
                    iAdjustedIconSize = (utFontRect.Bottom + 1) - utFontRect.Top
                Else
                    iAdjustedIconSize = m_lIconSize
                End If
                iTmpY = utFontRect.Top + Round((utFontRect.Bottom - utFontRect.Top - iAdjustedIconSize) / 2)
                Select Case PictureAlign
                    Case xAlignLeftEdge:
                        iTmpX = utFontRect.Left
                        '   If active tab then give a popup effect
                        If iCnt = m_lActiveTab Then
                            iTmpX = iTmpX + 1
                            iTmpY = iTmpY - 1
                            '   Make sure our adjustment dosen't make it
                            '   out of the font area
                            If iTmpY < utFontRect.Top Then iTmpY = utFontRect.Top
                        End If
                        If utTabInfo.TabPicture.Type = vbPicTypeBitmap And UseMaskColor Then
                            Call DrawImage(m_lhDC, utTabInfo.TabPicture.Handle, GetRGBFromOLE(PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        Else
                            Call pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        End If
                    Case xAlignRightEdge:
                        iTmpX = utFontRect.Right - iAdjustedIconSize
                        '   If active tab then give a popup effect
                        If iCnt = m_lActiveTab Then
                            iTmpX = iTmpX - 1
                            iTmpY = iTmpY - 1
                            '   Make sure our adjustment dosen't make it
                            '   out of the font area
                            If iTmpY < utFontRect.Top Then iTmpY = utFontRect.Top
                        End If
                        If utTabInfo.TabPicture.Type = vbPicTypeBitmap And UseMaskColor Then
                            Call DrawImage(m_lhDC, utTabInfo.TabPicture.Handle, GetRGBFromOLE(PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        Else
                            Call pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        End If
                    Case xAlignLeftOfCaption:
                        iOrigLeft = utFontRect.Left
                    Case xAlignRightOfCaption:
                        iOrigRight = utFontRect.Right
                End Select
            End If
            sTmp = utTabInfo.Caption
            '   Calculate the rect to draw the text, also modify the string
            '   to get ellipsis etc
            DrawText m_lhDC, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
            iTmpW = utFontRect.Right - utFontRect.Left + iFOCUS_RECT_AND_TEXT_DISTANCE
            iTmpH = utFontRect.Bottom - utFontRect.Top + iFOCUS_RECT_AND_TEXT_DISTANCE / 2
            '   Do the adjustments to center the text (both vertically and
            '   horizontally)
            utFontRect.Left = (utFontRect.Left - (iFOCUS_RECT_AND_TEXT_DISTANCE / 2)) + .Right / 2 - utFontRect.Right / 2
            utFontRect.Right = utFontRect.Left + iTmpW
            utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
            utFontRect.Bottom = utFontRect.Top + iTmpH
            If Not utTabInfo.TabPicture Is Nothing Then
                Select Case PictureAlign
                    Case xAlignLeftOfCaption:
                        iTmpX = utFontRect.Left - iAdjustedIconSize - 1
                        '   Make sure our adjustment dosen't make it out of
                        '   the font area
                        If iTmpX < iOrigLeft Then iTmpX = iOrigLeft
                        If utTabInfo.TabPicture.Type = vbPicTypeBitmap And UseMaskColor Then
                            Call DrawImage(m_lhDC, utTabInfo.TabPicture.Handle, GetRGBFromOLE(PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        Else
                            Call pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        End If
                    Case xAlignRightOfCaption:
                        iTmpX = utFontRect.Right + 1
                        '   Make sure our adjustment dosen't make it out of
                        '   the font area
                        If iTmpX + iAdjustedIconSize > iOrigRight Then iTmpX = iOrigRight - iAdjustedIconSize
                        If utTabInfo.TabPicture.Type = vbPicTypeBitmap And UseMaskColor Then
                            Call DrawImage(m_lhDC, utTabInfo.TabPicture.Handle, GetRGBFromOLE(PictureMaskColor), iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        Else
                            Call pPaintPicture(utTabInfo.TabPicture, iTmpX, iTmpY, iAdjustedIconSize, iAdjustedIconSize)
                        End If
                End Select
            End If
            '   Now draw the text
            DrawText m_lhDC, sTmp, -1, utFontRect, DT_SINGLELINE Or DT_CENTER
            If bUserMode Then    'only if in the run mode
                If iCnt = m_lActiveTab And m_IsFocused And ShowFocusRect Then
                    '   Draw focus rectangle
                    Call DrawFocusRect(m_lhDC, utFontRect)
                End If
            End If
        End With
    Next
    
    '   Store the larger tab height
    iCnt = IIf(m_lActiveTabHeight > m_lInActiveTabHeight, m_lActiveTabHeight, m_lInActiveTabHeight)
    
    'Adjust the corners (whole tab control's corners)
    APILine 0, iCnt + 1, 0, iCnt + 4, m_lOuterBorderColor
    APILine m_lScaleWidth - 1, iCnt + 1, m_lScaleWidth - 1, iCnt + 4, m_lOuterBorderColor
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.DrawTabsTabbedDialog", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get Enabled() As Boolean

    '   TabEnable Poperty is for individual tab's
    '   and this is for the whole control
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Enabled = UserControl.Enabled
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.Enabled", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let Enabled(ByVal bNewValue As Boolean)

    '   TabEnable Poperty is for individual tab's
    '   and this is for the whole control
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UserControl.Enabled() = bNewValue
    m_bEnabled = bNewValue
    '   Redraw
    Refresh
    PropertyChanged "Enabled"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.Enabled", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get FocusedColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    FocusedColor = m_lFocusedColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.FocusedColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let FocusedColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lFocusedColor = lNewValue
    '   Redraw
    Refresh
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.FocusedColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Set Font(ByVal oNewFont As StdFont)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set UserControl.Font = oNewFont
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.Font", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ForeColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ForeColor = UserControl.ForeColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ForeColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UserControl.ForeColor = lNewValue
    Refresh
    PropertyChanged "ForeColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Function GetRGBFromOLE(ByVal lOleColor As Long) As Long
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    ' Convert the OLE color into equivalent RGB Combination
    ' i.e. Convert vbButtonFace into ==> Light Grey
    Dim lRGBColor As Long
    Call TranslateColor(lOleColor, 0, lRGBColor)
    GetRGBFromOLE = lRGBColor
    
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "AeroTab.GetRGBFromOLE", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Sub HandleContainedControls(ByVal New_ActiveTab As Long)
    ' VERY IMPORTANT FUNCTION:
    '   Handles the appearing and disappearing of controls for the current
    '   tab and last active tab
    '
    '   This routine replaces the original routine implemented with Collections
    '   and is based on the PCS article by Evan Todder:
    '   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=57642&lngWId=1
    '
    '   NOTE:
    '   Unfortunatly, the above article was removed by the Author from PCS, so the link
    '   is not active; however I felt it was important to give credit for the cleaver
    '   idea none the less.
    '
    Dim Ctl As Control
    Dim MoveVal As Long
 
    On Error Resume Next
    '   The difference between what was the active
    '   Tab and the newly set activetab
    MoveVal = (New_ActiveTab - m_lActiveTab)
    '   Move the controls by a Factor which is
    '   tied to the Tab Diff....the default value
    '   is set to 10K, but for Objects greater than
    '   this, size the Width + 1000 will be used.
    MoveVal = (MoveVal * m_lMoveOffset)
    '   This is what creates the illusion of
    '   Changing the Tab of a tab control
    For Each Ctl In UserControl.ContainedControls
         Ctl.Left = (Ctl.Left + MoveVal)
    Next Ctl

End Sub

Public Property Get hdc() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    hdc = UserControl.hdc
    m_lhDC = UserControl.hdc
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.hDC", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get HoverColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    HoverColor = m_lHoverColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.HoverColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let HoverColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lHoverColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "HoverColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.HoverColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get hWnd() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    hWnd = UserControl.hWnd
    m_lhWnd = UserControl.hWnd
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.hwnd", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get InActiveTabBackEndColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    InActiveTabBackEndColor = m_lInActiveTabBackEndColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabBackEndColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let InActiveTabBackEndColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lInActiveTabBackEndColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "InActiveTabBackEndColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabBackEndColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get InActiveTabBackStartColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    InActiveTabBackStartColor = m_lInActiveTabBackStartColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabBackStartColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let InActiveTabBackStartColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lInActiveTabBackStartColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "InActiveTabBackStartColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabBackStartColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get InActiveTabFont() As StdFont

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set InActiveTabFont = m_oInActiveTabFont
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabFont", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Set InActiveTabFont(ByVal oNewFnt As StdFont)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Set m_oInActiveTabFont = oNewFnt
    '   Redraw
    Refresh
    PropertyChanged "InActiveTabFont"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabFont", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get InActiveTabForeColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    InActiveTabForeColor = m_lInActiveTabForeColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let InActiveTabForeColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lInActiveTabForeColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "InActiveTabForeColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabForeColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get InActiveTabHeight() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    InActiveTabHeight = m_lInActiveTabHeight
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabHeight", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let InActiveTabHeight(ByVal lNewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lInActiveTabHeight = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "InActiveTabHeight"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InActiveTabHeight", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Sub InsertTab(ByVal lAfterIndex As Long, Optional sCaption As String = "NewTab")
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    Dim Ctl As Control          'Control Default Type
    Dim i As Long               'Loop Counter
    Dim lTab As Long            'The Computed Tab Number for Each Control
    Dim MoveVal As Long         'Computed value to move the controls
    Dim lActiveTab              'ActiveTab
    
    '   Lock the window to improve speed and reduce flicker
    LockWindowUpdate UserControl.hWnd
    '   Save the active tab
    lActiveTab = m_lActiveTab
    '   Set the inital tab to be "0" so that the
    '   Left values are ordinal to the tab position
    ActiveTab = 0
    '   Loop over all controls
    For Each Ctl In UserControl.ContainedControls
        '   Now find out which Tab they are on...
        If (Ctl.Left > 0) Then
            '   Tab(0), since the Left values are all positive
            lTab = (Abs(Ctl.Left) \ m_lMoveOffset)
        Else
            '   Must be Tab(1)....Tab(n)
            lTab = (Abs(Ctl.Left) \ m_lMoveOffset) + 1
        End If
        '   Check ths index against the lTab value
        If lTab > lAfterIndex Then
            '   These are the target....we need to move
            '   these up by one increment of the MoveOffset
            MoveVal = m_lMoveOffset
            Ctl.Left = (Ctl.Left - MoveVal)
        Else
            '   Do nothing, these are below the tab we want to
            '   remove so leave them as they are...
        End If
    Next Ctl
    
    '   Change the Tab Count
    TabCount = m_lTabCount + 1
    '   Now loop over the tabs and add the captions back offset
    '   by the new tab
    For i = m_lTabCount - 1 To 0 Step -1
        '   Loop backwards or we copy the same name over and over ;-)
        If i > lAfterIndex + 1 Then
            TabCaption(i) = m_aryTabs(i - 1).Caption
        End If
    Next i
    TabCaption(lAfterIndex + 1) = sCaption
    '   Set back the tab
    If lActiveTab > lAfterIndex + 1 Then
        ActiveTab = lActiveTab + 1
    Else
        ActiveTab = lActiveTab
    End If
    '   Unlock the window
    LockWindowUpdate 0&
    RaiseEvent TabInsert(lAfterIndex)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.InsertTab", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Property Get IsFocused() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    IsFocused = m_IsFocused
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.IsFocused", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Property Let IsFocused(ByVal bNewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_IsFocused = bNewValue
    Refresh
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.IsFocused", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get LastActiveTab() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    'Note: Read Only Property: returns the last active tab index
    LastActiveTab = m_lLastActiveTab
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.LastActiveTab", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Function OffsetColor(lColor As OLE_COLOR, lOffset As Long) As OLE_COLOR
    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lR As OLE_COLOR, lG As OLE_COLOR, LB As OLE_COLOR
    
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    '   Translate the color first to make sure it is on our palette
    lColor = GetRGBFromOLE(lColor)
    '   Now split the colors into RGB
    lR = (lColor And &HFF)
    lG = ((lColor And 65280) \ 256)
    LB = ((lColor) And 16711680) \ 65536
    lRed = (lOffset + lR)
    lGreen = (lOffset + lG)
    lBlue = (lOffset + LB)
    
    If lRed > 255 Then lRed = 255
    If lRed < 0 Then lRed = 0
    If lGreen > 255 Then lGreen = 255
    If lGreen < 0 Then lGreen = 0
    If lBlue > 255 Then lBlue = 255
    If lBlue < 0 Then lBlue = 0
    OffsetColor = RGB(lRed, lGreen, lBlue)
    
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "AeroTab.OffsetColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Public Property Get OuterBorderColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    OuterBorderColor = m_lOuterBorderColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.OuterBorderColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let OuterBorderColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lOuterBorderColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "OuterBorderColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.OuterBorderColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Function pAlphaBlend(ByVal lColor1 As Long, ByVal lColor2 As Long, Optional ByVal bAlpha As Byte = 128) As Long
    Dim R1 As Long
    Dim R2 As Long
    Dim RA As Long
    Dim G1 As Long
    Dim G2 As Long
    Dim GA As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim BA As Long
    
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler
    
    '   Split the colors to get RGB
    R1 = pGetRValue(lColor1)
    R2 = pGetRValue(lColor2)
    G1 = pGetGValue(lColor1)
    G2 = pGetGValue(lColor2)
    b1 = pGetBValue(lColor1)
    b2 = pGetBValue(lColor2)
    
    RA = (R1 * (bAlpha / 255)) + (R2 * (1# - (bAlpha / 255)))
    GA = G1 * ((bAlpha / 255)) + (G2 * (1# - (bAlpha / 255)))
    BA = b1 * ((bAlpha / 255)) + (b2 * (1# - (bAlpha / 255)))
    
    pAlphaBlend = RGB(RA, BA, GA)
    
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    '   On errors pass back the original color unchanged
    pAlphaBlend = lColor1
    'Err.Raise Err.Number, "AeroTab.pAlphaBlend", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Sub pAssignAccessKeys()
    '   Function used to extact the access keys from tabs and
    '   reassign them to AccessKeys property
    
    Dim iCnt As Long
    Dim sTmp As String
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    For iCnt = 0 To m_lTabCount - 1
        If m_aryTabs(iCnt).AccessKey <> 0 Then
            sTmp = sTmp & Chr$(m_aryTabs(iCnt).AccessKey)
        End If
    Next
    AccessKeys = sTmp
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pAssignAccessKeys", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pDestroyResources()
    '   Function deletes the pictures etc and frees up the res
    On Error Resume Next
    
    Dim iCnt As Long
    
    '   Free up the memory
    For iCnt = 0 To m_lTabCount - 1
        Set m_aryTabs(iCnt).TabPicture = Nothing
    Next
End Sub

Private Sub pDrawMe()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Start by Clearing things
    UserControl.Cls
    ' Function draws both Background and the Tabs
    Call DrawBackground
    Call DrawTabs
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pDrawMe", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pDrawOnMouseOverProperty()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Function called when mouse hovers over the tab (Property Style)
    m_bIsMouseOver = True
    m_utRect = AryTabs(m_lMouseOverTabIndex).ClickableRect
    '   Since the mouse is over the tab.. track the mouse going out of
    '   the clickable region
    Call pSetTimer(10)
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pDrawOnMouseOverProperty", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pDrawOnMouseOverTabbed()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Function called when mouse hovers over the tab (Tabbed Style)
    m_bIsMouseOver = True
    m_utRect = AryTabs(m_lMouseOverTabIndex).ClickableRect
    Call pSetTimer(10)
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pDrawOnMouseOverTabbed", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pFillCurvedGradient(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '   Over-ridden function for pFillCurvedGradientR, performs same job,
    '   but takes integers instead of Rect as parameter
    Dim utRect As RECT
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    utRect.Left = lLeft
    utRect.Top = lTop
    utRect.Right = lRight
    utRect.Bottom = lBottom
    Call pFillCurvedGradientR(utRect, lStartColor, lEndColor, iCurveValue, bCurveLeft, bCurveRight)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pFillCurvedGradient", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pFillCurvedGradientR(utRect As RECT, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '   Function used to Fill a rectangular area by Gradient
    '   This function can draw using the curve value to generate a rounded
    '   RECT kind of effect
    
    Dim sngRedInc As Single, sngGreenInc As Single, sngBlueInc As Single
    Dim sngRed As Single, sngGreen As Single, sngBlue As Single
    Dim intCnt As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    lStartColor = GetRGBFromOLE(lStartColor)
    lEndColor = GetRGBFromOLE(lEndColor)
    sngRedInc = (pGetRValue(lEndColor) - pGetRValue(lStartColor)) / (utRect.Bottom - utRect.Top)
    sngGreenInc = (pGetGValue(lEndColor) - pGetGValue(lStartColor)) / (utRect.Bottom - utRect.Top)
    sngBlueInc = (pGetBValue(lEndColor) - pGetBValue(lStartColor)) / (utRect.Bottom - utRect.Top)
    sngRed = pGetRValue(lStartColor)
    sngGreen = pGetGValue(lStartColor)
    sngBlue = pGetBValue(lStartColor)
    
    If iCurveValue = -1 Then
        For intCnt = utRect.Top To utRect.Bottom
            Call APILine(utRect.Left, intCnt, utRect.Right, intCnt, RGB(sngRed, sngGreen, sngBlue))
            sngRed = sngRed + sngRedInc
            sngGreen = sngGreen + sngGreenInc
            sngBlue = sngBlue + sngBlueInc
        Next
    Else
        If bCurveLeft And bCurveRight Then
            For intCnt = utRect.Top To utRect.Bottom
                Call APILine(utRect.Left + iCurveValue + 1, intCnt, utRect.Right - iCurveValue, intCnt, RGB(sngRed, sngGreen, sngBlue))
                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        ElseIf bCurveLeft Then
            For intCnt = utRect.Top To utRect.Bottom
                Call APILine(utRect.Left + iCurveValue + 1, intCnt, utRect.Right, intCnt, RGB(sngRed, sngGreen, sngBlue))
                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        Else    'curve right
            For intCnt = utRect.Top To utRect.Bottom
                Call APILine(utRect.Left, intCnt, utRect.Right - iCurveValue, intCnt, RGB(sngRed, sngGreen, sngBlue))
                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        End If
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pFillCurvedGradientR", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pFillCurvedSolid(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal lColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '   Over-ridden function for pFillCurveSolid, performs same job,
    '   but takes integers instead of Rect as parameter
    Dim utRect As RECT
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    utRect.Left = lLeft
    utRect.Top = lTop
    utRect.Right = lRight
    utRect.Bottom = lBottom
    Call pFillCurvedSolidR(utRect, lColor, iCurveValue, bCurveLeft, bCurveRight)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pFillCurvedSolid", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pFillCurvedSolidR(utRect As RECT, ByVal lColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    '   Function used to Fill a rectangular area by Solid color.
    '   This function can draw using the curve value to generate
    '   a rounded rect kind of effect
    
    Dim intCnt As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If iCurveValue = -1 Then
        For intCnt = utRect.Top To utRect.Bottom
            Call APILine(utRect.Left, intCnt, utRect.Right, intCnt, lColor)
        Next
    Else
        If bCurveLeft And bCurveRight Then
            For intCnt = utRect.Top To utRect.Bottom
                Call APILine(utRect.Left + iCurveValue + 1, intCnt, utRect.Right - iCurveValue, intCnt, lColor)
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        ElseIf bCurveLeft Then
            For intCnt = utRect.Top To utRect.Bottom
                Call APILine(utRect.Left + iCurveValue + 1, intCnt, utRect.Right, intCnt, lColor)
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        Else    '   Curve right
            For intCnt = utRect.Top To utRect.Bottom
                Call APILine(utRect.Left, intCnt, utRect.Right - iCurveValue, intCnt, lColor)
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        End If
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pFillCurvedSolidR", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Function pGetBValue(ByVal RGBValue As Long) As Long

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    '   Extract Blue component from a color
    pGetBValue = ((RGBValue And &HFF0000) / &H10000) And &HFF
    
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pGetBValue", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Sub pGetCachedProperties()
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Get the Current Cached properties from the control
    '   (this prevents trips to again and again fetch properties
    '   from the user control)
    m_lhWnd = hWnd
    m_lhDC = hdc
    m_lActiveTab = ActiveTab
    m_lActiveTabHeight = ActiveTabHeight
    m_lInActiveTabHeight = InActiveTabHeight
    m_lScaleWidth = ScaleWidth
    m_lScaleHeight = ScaleHeight
    m_lTabCount = TabCount
    m_IsFocused = IsFocused
    m_lOuterBorderColor = OuterBorderColor

    m_lActiveTabForeColor = ActiveTabForeColor
    m_lActiveTabBackStartColor = ActiveTabBackStartColor
    m_lActiveTabBackEndColor = ActiveTabBackEndColor

    m_lInActiveTabForeColor = InActiveTabForeColor
    m_lInActiveTabBackStartColor = InActiveTabBackStartColor
    m_lInActiveTabBackEndColor = InActiveTabBackEndColor
    m_lHoverColor = HoverColor
    m_lDisabledTabBackColor = DisabledTabBackColor
    m_lDisabledTabForeColor = DisabledTabForeColor

    '   Get System's default size for a Icon.
    If PictureSize = xSizeSmall Then
        m_lIconSize = GetSystemMetrics(SM_CXSMICON)
    Else
        m_lIconSize = GetSystemMetrics(SM_CXICON)
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pGetCachedProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Function pGetControlId(ByRef oCtl As Control) As String
    On Error Resume Next
    '   Function returns control's name & control's index combination
    Static sCtlName As String
    Static iCtlIndex As Long
    
    iCtlIndex = -1
    sCtlName = oCtl.Name
    iCtlIndex = oCtl.Index
    pGetControlId = sCtlName & IIf(iCtlIndex <> -1, iCtlIndex, "")
End Function

Private Function pGetGValue(ByVal RGBValue As Long) As Long
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    '   Extract Green component from a color
    pGetGValue = ((RGBValue And &HFF00) / &H100) And &HFF

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pGetGValue", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Function pGetRValue(ByVal RGBValue As Long) As Long
    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    '   Extract Red component from a color
    pGetRValue = RGBValue And &HFF

Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pGetRValue", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Sub pHandleMouseDown(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    '   Function calls MouseDownHandler of current theme
    '   Called when the user presses mouse on the clickable area of the user control
    Dim iCnt As Long
    Dim iX As Long
    Dim iY As Long
    Dim utTabInfo As TabInfo
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   If the mouse is already over then stop the timer and reset the
    '   over flag
    If m_bIsMouseOver Then
        m_bIsMouseOver = False
        Call pSetTimer(0)
    End If
    iX = CInt(sngX)
    iY = CInt(sngY)
    If iY > IIf(m_lActiveTabHeight > m_lInActiveTabHeight, m_lActiveTabHeight, m_lInActiveTabHeight) Then
        '   If lower than the larger tab height then exit sub since
        '   anything lower than active tab's height will not result
        '   in a tab switch
        Exit Sub
    End If
    '   Now go through each tab's rect to determine if the mouse was
    '   clicked within its boundaries
    For iCnt = 0 To m_lTabCount - 1
        utTabInfo = AryTabs(iCnt)
        If iX >= utTabInfo.ClickableRect.Left And iX <= utTabInfo.ClickableRect.Right And _
           iY >= utTabInfo.ClickableRect.Top And iY <= utTabInfo.ClickableRect.Bottom And utTabInfo.Enabled Then
            '   If its the active tab then no need to switch
            If m_lActiveTab <> iCnt Then
                ActiveTab = iCnt
            End If
            '   Our work is finished....so exit
            Exit Sub
        End If
    Next
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pHandleMouseDown", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pHandleMouseMove(iButton As Integer, iShift As Integer, sngX As Single, sngY As Single)
    ' Function calls MouseMove of current theme
    'Called as the user moves the mouse on the control
    Dim iCnt As Long
    Dim iX As Long
    Dim iY As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    iX = CInt(sngX)
    iY = CInt(sngY)
    '   No use going in if already mouse is over
    If m_bIsMouseOver Then
        Exit Sub
    End If
    If iY > IIf(m_lInActiveTabHeight > m_lActiveTabHeight, m_lInActiveTabHeight, m_lActiveTabHeight) Then
        '   If lower than the larger tab height then exit sub since
        '   anything lower than larger tab's height will not result
        '   in a tab switch
        Exit Sub
    End If
    '   Get the current cached properties
    Call pGetCachedProperties
    For iCnt = 0 To m_lTabCount - 1
        If iX >= AryTabs(iCnt).ClickableRect.Left And iX <= AryTabs(iCnt).ClickableRect.Right And AryTabs(iCnt).Enabled Then
            '   No need to draw for the active tab
            If iCnt = m_lActiveTab Then Exit Sub
            '   Store the index of the tab on which the mouse is over
            m_lMouseOverTabIndex = iCnt
            Select Case TabStyle
                Case xStyleTabbedDialog:
                    Call pDrawOnMouseOverTabbed
                Case xStylePropertyPages:
                    Call pDrawOnMouseOverProperty
            End Select
            Exit Sub
        End If
    Next
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pHandleMouseMove", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pHandleTabCount()
    '   Function Called to handle the addition or deletion of tabs
    Dim iCnt As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If (m_lTabCount - 1) > UBound(m_aryTabs) Then
        '   Tabs added
        iCnt = UBound(m_aryTabs) + 1
        '   Redim the tabs array
        ReDim Preserve m_aryTabs(m_lTabCount - 1)
        '   Initialize the added tabs
        For iCnt = iCnt To m_lTabCount - 1
            m_aryTabs(iCnt).Caption = m_def_sCaption & " " & iCnt
            m_aryTabs(iCnt).Enabled = m_def_bTabEnabled
        Next
    ElseIf (m_lTabCount - 1) <= UBound(m_aryTabs) Then
        '   Tabs removed
        iCnt = UBound(m_aryTabs) + 1
        '   Redim the tabs array
        ReDim Preserve m_aryTabs(m_lTabCount - 1)
        '   Make sure if the active tab is within the tab count
        If (m_lActiveTab >= m_lTabCount) Or (m_lTabCount = 1) Then
            m_lActiveTab = m_lTabCount - 1
        End If
        'Else ' :No need since this means that the number of tabs has not changed
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pHandleTabCount", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get PictureAlign() As PictureAlign

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    PictureAlign = m_enmPictureAlign
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.PictureAlign", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let PictureAlign(ByVal iNewValue As PictureAlign)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_enmPictureAlign = iNewValue
    '   Redraw
    Refresh
    PropertyChanged "PictureAlign"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.PictureAlign", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get PictureMaskColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    PictureMaskColor = UserControl.MaskColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.PictureMaskColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let PictureMaskColor(ByVal lNewColor As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UserControl.MaskColor = lNewColor
    '   Redraw
    Refresh
    PropertyChanged "PictureMaskColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.PictureMaskColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get PictureSize() As PictureSize

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    PictureSize = m_enmPictureSize
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.PictureSize", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let PictureSize(ByVal iNewSize As PictureSize)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_enmPictureSize = iNewSize
    '   Redraw
    Refresh
    PropertyChanged "PictureSize"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.PictureSize", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Function pIsControlAdded(ByVal lTabIndex As Long, ByVal sCtrlName As String) As Boolean
    '   Determines if a control is added in a specific tab or not
    Dim oCtl As Control
    
    On Error GoTo Err_Handler
    
    '   If no error occured while accessing the value that means
    '   the control is already added
    For Each oCtl In UserControl.ContainedControls
        If oCtl.Name = sCtrlName Then
            pIsControlAdded = True
            Exit Function
        End If
    Next
    pIsControlAdded = False
    
Err_Handler:
    '   Error occured while accessing the control.....
    '   that means it is not added
    pIsControlAdded = False
End Function

Private Sub pPaintPicture(oStdPic As StdPicture, ByVal iX1 As Integer, ByVal iY1 As Integer, Optional ByVal iWidth As Integer = -1, Optional ByVal iHeight As Integer = -1)
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Wrapper function for Form's PaintPicture Method
    If iWidth = -1 And iHeight = -1 Then
        Call PaintPicture(oStdPic, iX1, iY1)
    ElseIf iWidth = -1 Then
        Call PaintPicture(oStdPic, iX1, iY1, , iHeight)
    ElseIf iHeight = -1 Then
        Call PaintPicture(oStdPic, iX1, iY1, iWidth)
    Else
        Call PaintPicture(oStdPic, iX1, iY1, iWidth, iHeight)
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pPaintPicture", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pSetTimer(ByVal lInterval As Long)
    '   Function used to set the timer to ON/OFF
    '   based on the parameter value....this routine is
    '   basically the same as the one publish in the
    '   original article on PCS, but uses Sublassing
    '   Timers in lieu of vbTimer.
    Dim lRet    As Long
    
    On Error Resume Next
    If lInterval = 0 Then
        lRet = KillTimer(UserControl.hWnd, m_lTimerID)
    Else
        lRet = SetTimer(UserControl.hWnd, m_lTimerID, lInterval, 0)
    End If
End Sub

Private Sub pShowHideFocus()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Function called as a result of show/hide focus
    If m_bIsRecursive Then Exit Sub
    '   Set Recursive flag
    m_bIsRecursive = True
    '   This is done to allow the control to draw properly first time
    If m_bAreControlsAdded Then
        AutoRedraw = True
        ShowHideFocus
        '   Refresh the control
        Refresh
        AutoRedraw = False
    Else
        '   This also, will only draw the focus rectangle and
        '   prevent complete repaint
        ShowHideFocus
    End If
    m_bIsRecursive = False
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.pShowHideFocus", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pStoreOriginalTabStopValues()
   'TODO: Update this routine to take advantage of the new UDT
   '      and wire this to the main Tab functionality

    Dim oCtl As Control
    Dim iCnt As Long
    Dim sControlId As String
    'Dim oControlDetails As ControlDetails
    ' called only once to initialize the tab stop values
    ' that is store the original tab stop values for the controls
    ' and set there tab stop to false

    On Error Resume Next    'used to prevent errors (

    For Each oCtl In ContainedControls
        sControlId = pGetControlId(oCtl)  'Get Control's Id (i.e. Control Name & Control Index)
        For iCnt = 0 To m_lTabCount - 1
            'Now see if the control is already added (it should be)
            If pIsControlAdded(iCnt, sControlId) Then
                'oControlDetails.ControlID = sControlId
                'oControlDetails.TabStop = oCtl.TabStop
                With m_aryTabs(iCnt)
                    .ControlID = sControlId
                    .TabStop = oCtl.TabStop
                End With
                If iCnt <> m_lActiveTab Then    'if not the active tab then set the control's tabstop to false
                    oCtl.TabStop = False
                End If
                Exit For
            End If
        Next
    Next
End Sub

Public Sub Refresh()
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Prevent recursion. (flag set from calling function)
    If m_bIsRecursive Then Exit Sub
    m_bIsRecursive = True
    AutoRedraw = True
    '   If the controls have not been added
    If Not m_bAreControlsAdded Then
        '   This is a big problem with Container ActiveX controls. :-(
        '   Information about contained controls is not available until Paint/Show
        '   The "AfterCompleteInit" event fires to notify the host application that
        '   the UserControl finished loading.
        If Ambient.UserMode Then
            '   Indicate control is initialised
            m_bAreControlsAdded = True
            '   Tell container that we've finished loading
            RaiseEvent AfterCompleteInit
        Else
            '   Indicate that control is initialised
            m_bAreControlsAdded = True
        End If
    End If
    '   Now draw ourselves
    Call pDrawMe
    If Not Ambient.UserMode Then
        AutoRedraw = False
    End If
    m_bIsRecursive = False

    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.Refresh", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub RemoveTab(ByVal lIndex As Long)
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    Dim Ctl As Control          'Control Default Type
    Dim i As Long               'Loop Counter
    Dim lTab As Long            'The Computed Tab Number for Each Control
    Dim MoveVal As Long         'Computed value to move the controls
    Dim lActiveTab              'ActiveTab
    
    '   Must have at least one tab ;-)
    If m_lTabCount = 1 Then Exit Sub
    '   Lock the window to improve speed and reduce flicker
    LockWindowUpdate UserControl.hWnd
    '   Save the current tab
    lActiveTab = m_lActiveTab
    '   Set the inital tab to be "0" so that the
    '   Left values are ordinal to the tab position
    ActiveTab = 0
    '   Loop over all controls
    For Each Ctl In UserControl.ContainedControls
        '   Now find out which Tab they are on...
        If (Ctl.Left > 0) Then
            '   Tab(0), since the Left values are all positive
            lTab = (Abs(Ctl.Left) \ m_lMoveOffset)
        Else
            '   Must be Tab(1)....Tab(n)
            lTab = (Abs(Ctl.Left) \ m_lMoveOffset) + 1
        End If
        '   Check ths index against the lTab value
        If lTab = lIndex Then
            '   This is the tab we are removing, so we need to set
            '   these to an offset which will never be seen
            MoveVal = m_lMoveOffset * m_lTabCount
            Ctl.Left = (Ctl.Left + MoveVal)
        ElseIf lTab > lIndex Then
            '   These are the target....we need to move
            '   these down by one increment of the MoveOffset
            MoveVal = m_lMoveOffset
            Ctl.Left = (Ctl.Left + MoveVal)
        Else
            '   Do nothing, these are below the tab we want to
            '   remove so leave them as they are...
        End If
    Next Ctl
    
    '   Now loop over the tabs and add the captions back
    For i = 0 To m_lTabCount - 1
        If i > lIndex Then
            TabCaption(i - 1) = m_aryTabs(i).Caption
        End If
    Next i
    '   Change the Tab Count
    TabCount = m_lTabCount - 1
    '   Now set the tab
    If lActiveTab = lIndex Then
        If lIndex = m_lTabCount - 1 Then
            ActiveTab = lActiveTab - 1
        Else
            ActiveTab = lActiveTab + 1
        End If
    Else
        ActiveTab = lActiveTab - 1
    End If
    '   Unlock the window
    LockWindowUpdate 0&
    RaiseEvent TabRemove(lIndex)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.RemoveTab", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub RemoveTabImages(Optional Index As Long = -1, Optional bRemoveAll As Boolean = True)
    '   Copy From a standard Image List to current tabs images
    Dim iTmp As Long
    
    If Index = -1 And bRemoveAll = True Then
        '   If the number of images is less than number of tabs error may occur
        On Error GoTo Finally
        For iTmp = 0 To UBound(m_aryTabs)
            Set m_aryTabs(iTmp).TabPicture = Nothing  'Free Existing Picture
        Next
    Else
        '   If the number of images is less than number of tabs error may occur
        On Error GoTo Finally
        Set m_aryTabs(Index).TabPicture = Nothing  'Free Existing Picture
    End If
Finally:
    '   Redraw
    Refresh
    '   Do nothing as this is possibly because the num of images is less
End Sub

Public Sub ResetAllColors()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Reset all the colors to the default colors
    '
    '   Prevent Redrawing until all the properties are set
    '   (in the ResetColorsToDefault function)
    m_bIsRecursive = True
    '   Call the Theme specific function for reseting the colors
    Call ResetColorsToDefault
    '   Prevent Redrawing untill all the properties are set
    '   (in the ResetColorsToDefault function)
    m_bIsRecursive = False
    '   Now Force Redraw
    Refresh
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ResetAllColors", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Sub ResetColorsToDefault()
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    AutoRedraw = False
    '   Function used to replace all the colors with the system colors
    With Me
        m_lActiveTabBackStartColor = &HFFFFFF
        m_lActiveTabBackEndColor = &HFFFFFF
        m_lInActiveTabBackStartColor = RGB(235, 235, 235)
        m_lInActiveTabBackEndColor = RGB(207, 207, 207)
        '   Color drawn in XOR mode to achive the hover effect
        m_lHoverColor = m_def_lHoverColor
        m_lFocusedColor = m_def_lFocusedColor
        m_lActiveTabForeColor = vbBlack
        m_lInActiveTabForeColor = vbBlack
        On Error Resume Next
        'm_lTabStripBackColor = UserControl.Ambient.BackColor
        'm_lTabStripBackColor = IIf(TabStripBackColor <> Ambient.BackColor, TabStripBackColor, Ambient.BackColor)
        m_lDisabledTabBackColor = RGB(201, 202, 203)
        m_lDisabledTabForeColor = &HA0A0A0
        m_lOuterBorderColor = RGB(137, 140, 149)
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ResetColorsToDefault", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get ScaleHeight() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ScaleHeight = UserControl.ScaleHeight
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ScaleHeight", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get ScaleWidth() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ScaleWidth = UserControl.ScaleWidth
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ScaleWidth", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Sub ScrollTabs(lDirection As Long)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    If lDirection < 0 Then
        '   Scroll Up the Tab Order
        If m_lActiveTab < m_lTabCount Then
            ActiveTab = m_lActiveTab + 1
        End If
    Else
        '   Scroll Down the Tab Order
        If m_lActiveTab > 0 Then
            ActiveTab = m_lActiveTab - 1
        End If
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ScrollTabs", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get ShowFocusRect() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ShowFocusRect = m_bShowFocusRect
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ShowFocusRect", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ShowFocusRect(ByVal bNewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_bShowFocusRect = bNewValue
    Refresh
    PropertyChanged ShowFocusRect
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ShowFocusRect", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Sub ShowHideFocus()
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Called when User control Get's or looses the focus
    '   This way we are able to prevent complete control redrawing
    '   only the focus rect is drawn or erased. Thus preventing flicker
    '   Get local copy for control properties
    Call pGetCachedProperties
    Select Case TabStyle
        Case xStylePropertyPages
            ShowHideFocusPropertyPages
        Case xStyleTabbedDialog
            ShowHideFocusTabbedDialog
    End Select
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ShowHideFocus", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub ShowHideFocusPropertyPages()
    '   Called from ShowHideFocus for more speific processing
    Dim utFontRect As RECT
    Dim utTabInfo As TabInfo
    Dim iTmpW As Long
    Dim iTmpH As Long
    Dim sTmp As String
    Dim iAdjustedIconSize As Long
    Dim iOrigLeft As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Only if in the run mode
    If Not bUserMode Then
        Exit Sub
    End If
    '   Only if Show Focus Rect is true for the control
    If Not ShowFocusRect Then
        Exit Sub
    End If
    utTabInfo = AryTabs(m_lActiveTab)
    With utTabInfo.ClickableRect
        '   Do the adjustments for the border
        utFontRect.Left = .Left + 2
        utFontRect.Top = .Top + 4
        utFontRect.Bottom = .Bottom
        utFontRect.Right = .Right - 2
        '   Done to allow proper drawing of focus rect
        If (m_bEnabled) Then
            If utTabInfo.Enabled Then
                ForeColor = m_lActiveTabForeColor
            Else
                ForeColor = m_lDisabledTabForeColor
            End If
        Else
            ForeColor = m_lDisabledTabForeColor
        End If
        '   Show/hide the focus rectangle (drawn in XOR mode,
        '   so calling it again with same coords will erase it)
        Call DrawFocusRect(m_lhDC, utFontRect)
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ShowHideFocusPropertyPages", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub ShowHideFocusTabbedDialog()
    '   Called from ShowHideFocus for more speific processing
    Dim utFontRect As RECT
    Dim sTmp As String
    Dim utTabInfo As TabInfo
    Dim iTmpW As Long
    Dim iTmpH As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Only if in the run mode
    If Not bUserMode Then
        Exit Sub
    End If
    '   Only if Show Focus Rect is true for the control
    If Not ShowFocusRect Then
        Exit Sub
    End If
    utTabInfo = AryTabs(m_lActiveTab)
    With utTabInfo.ClickableRect
        '   Do the adjustments for the border
        utFontRect.Left = .Left + 3
        utFontRect.Top = .Top + 3
        utFontRect.Bottom = .Bottom
        utFontRect.Right = .Right - 2
        sTmp = utTabInfo.Caption
        Set Font = ActiveTabFont
        '   Calculate the rect to draw the text, and get proper
        '   string (including ellipsis etc)
        DrawText m_lhDC, sTmp, -1, utFontRect, DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
        iTmpW = utFontRect.Right - utFontRect.Left + iFOCUS_RECT_AND_TEXT_DISTANCE
        iTmpH = utFontRect.Bottom - utFontRect.Top + iFOCUS_RECT_AND_TEXT_DISTANCE / 2
        '   Do the adjustments to center the text (both vertically
        '   and horizontally)
        utFontRect.Left = (utFontRect.Left - (iFOCUS_RECT_AND_TEXT_DISTANCE / 2)) + .Right / 2 - utFontRect.Right / 2
        utFontRect.Right = utFontRect.Left + iTmpW
        utFontRect.Top = utFontRect.Top + .Bottom / 2 - utFontRect.Bottom / 2
        utFontRect.Bottom = utFontRect.Top + iTmpH
        '   Done to allow proper drawing of focus rect
        If (m_bEnabled) Then
            If utTabInfo.Enabled Then
                ForeColor = m_lActiveTabForeColor
            Else
                ForeColor = m_lDisabledTabForeColor
            End If
        Else
            ForeColor = m_lDisabledTabForeColor
        End If
        '   Show/hide the focus rectangle (drawn in XOR mode, so calling
        '   it again with same coords will erase it)
        Call DrawFocusRect(m_lhDC, utFontRect)
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.ShowHideFocusTabbedDialog", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get TabCaption(ByVal iTabIndex As Long) As String

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If iTabIndex > -1 And iTabIndex < m_lTabCount Then
        TabCaption = m_aryTabs(iTabIndex).Caption
    End If
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabCaption", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TabCaption(ByVal iTabIndex As Long, sTabCaption As String)
    Dim sTmp As String
    Dim lPos As Long
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If iTabIndex > -1 And iTabIndex < m_lTabCount Then
        '   First get the existing caption's access key and remove it from the
        '   current "AccessKeys" property for the control
        sTmp = Replace$(m_aryTabs(iTabIndex).Caption, "&&", Chr$(1))
        lPos = InStrRev(sTmp, "&")
        If m_aryTabs(iTabIndex).AccessKey <> 0 Then
            '   Remove from AccessKey
            AccessKeys = Replace$(AccessKeys, UCase$(Mid$(sTmp, lPos + 1, 1)), "")
        End If
        m_aryTabs(iTabIndex).Caption = sTabCaption
        '   Now get the new caption's access key and append it to the "AccessKeys" property
        sTmp = Replace$(m_aryTabs(iTabIndex).Caption, "&&", Chr$(1))
        lPos = InStrRev(sTmp, "&")
        If lPos Then
            '   Note: we are using Ucase$... since we store all the access keys in uppercase only
            m_aryTabs(iTabIndex).AccessKey = Asc(UCase$(Mid$(sTmp, lPos + 1, 1)))
            AccessKeys = AccessKeys & Chr$(m_aryTabs(iTabIndex).AccessKey)
        Else
            '   Reset the access key
            m_aryTabs(iTabIndex).AccessKey = 0
        End If
        PropertyChanged "TabCaption"
        '   Redraw
        Refresh
    Else
        '   Subscript Out of Range
        Err.Raise 9
        Exit Property
    End If
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabCaption", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get TabCount() As Long
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    '   The number of tabs
    TabCount = m_lTabCount

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabCount", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TabCount(ByVal lNewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If lNewValue < 1 Then
        '   Invalid property value
        Err.Raise 380
        Exit Property
    End If
    m_lTabCount = lNewValue
    '   Handle the change in tabcount
    '   (i.e. resize/initialize the array of tabs)
    Call pHandleTabCount
    '   Redraw
    Refresh
    PropertyChanged "TabCount"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabCount", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get TabEnabled(ByVal lTabIndex As Long) As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If (lTabIndex > -1) And (lTabIndex < m_lTabCount) Then
        TabEnabled = m_aryTabs(lTabIndex).Enabled
    End If
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabEnabled", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TabEnabled(ByVal lTabIndex As Long, bNewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If (lTabIndex > -1) And (lTabIndex < m_lTabCount) Then
        m_aryTabs(lTabIndex).Enabled = bNewValue
        PropertyChanged "TabEnabled"
        '   Redraw
        Refresh
        If m_bUseControlBox Then
            If bNewValue Then
                Call DrawControlBox(xStateNormal, lTabIndex)
            Else
                Call DrawControlBox(xStateDown, lTabIndex)
            End If
        End If
    Else
        '   Subscript Out of Range
        Err.Raise 9
        Exit Property
    End If
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabEnabled", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get TabOffset() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    TabOffset = m_lMoveOffset
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabOffset", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TabOffset(ByVal lOffset As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lMoveOffset = lOffset
    PropertyChanged "TabOffset"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabOffset", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get TabPicture(ByVal lTabIndex As Long) As StdPicture

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If (lTabIndex > -1) And (lTabIndex < m_lTabCount) Then
        Set TabPicture = m_aryTabs(lTabIndex).TabPicture
    Else
        '   Subscript Out of Range
        Err.Raise 9
        Exit Property
    End If
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabPicture", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Set TabPicture(ByVal lTabIndex As Long, oTabPicture As StdPicture)
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If (lTabIndex > -1) And (lTabIndex < m_lTabCount) Then
        Set m_aryTabs(lTabIndex).TabPicture = oTabPicture
        '   Redraw
        Refresh
        PropertyChanged "TabPicture"
    Else
        '   Subscript Out of Range
        Err.Raise 9
        Exit Property
    End If
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabPicture", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get TabStripBackColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    TabStripBackColor = m_lTabStripBackColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabStripBackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TabStripBackColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lTabStripBackColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "TabStripBackColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabStripBackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get TabStyle() As Style

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    '   Style for the tab
    TabStyle = m_enmStyle

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabStyle", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TabStyle(ByVal enmNewStyle As Style)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_enmStyle = enmNewStyle
    '   Redraw
    Refresh
    PropertyChanged "TabStyle"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TabStyle", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Sub TimerEvent()
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Called when the user control's timer event occurs
    '   Get cusror position
    Call GetCursorPos(m_Pnt)
    '   Convert coordinates
    Call ScreenToClient(m_lhWnd, m_Pnt)
    '   Check if the mouse is out of the clickable region
    If (m_Pnt.X < m_utRect.Left Or m_Pnt.X > m_utRect.Right) Or _
        (m_Pnt.Y < m_utRect.Top Or m_Pnt.Y > m_utRect.Bottom) Then
        '   Disable the timer
        Call pSetTimer(0)
        '   Indicate mouse up
        m_bIsMouseOver = False
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TimerEvent", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get TopLeftInnerBorderColor() As OLE_COLOR

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    TopLeftInnerBorderColor = m_lTopLeftInnerBorderColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TopLeftInnerBorderColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TopLeftInnerBorderColor(ByVal lNewValue As OLE_COLOR)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lTopLeftInnerBorderColor = lNewValue
    '   Redraw
    Refresh
    PropertyChanged "TopLeftInnerBorderColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.TopLeftInnerBorderColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get UseControlBox() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UseControlBox = m_bUseControlBox

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseControlBox", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let UseControlBox(ByVal NewValue As Boolean)
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    m_bUseControlBox = NewValue
    Refresh
    PropertyChanged "UseControlBox"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseControlBox", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get UseFocusedColor() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UseFocusedColor = m_bUseFocusedColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseFocusedColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let UseFocusedColor(ByVal bNewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_bUseFocusedColor = bNewValue
    '   Redraw
    Refresh
    PropertyChanged "UseFocusedColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseFocusedColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get UseMaskColor() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UseMaskColor = m_bUseMaskColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseMaskColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let UseMaskColor(bNewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_bUseMaskColor = bNewValue
    '   Redraw
    Refresh
    PropertyChanged "UseMaskColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseMaskColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get UseMouseWheelScroll() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    UseMouseWheelScroll = m_bUseMouseWheelScroll
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseMouseWheelScroll", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let UseMouseWheelScroll(bValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_bUseMouseWheelScroll = bValue
    PropertyChanged "UseMouseWheelScroll"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UseMouseWheelScroll", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Sub MDIHost_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    '   We need to be sure and unsubclass on the Query else
    '   the host object may be destroyed before the AeroTab
    '   which results in a GPF ;-)
    If bSubClass And (Not Cancel) Then
        Call UserControl_Terminate
    End If
End Sub

Private Sub SDIHost_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    '   We need to be sure and unsubclass on the Query else
    '   the host object may be destroyed before the AeroTab
    '   which results in a GPF ;-)
    If bSubClass And (Not Cancel) Then
        Call UserControl_Terminate
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Dim iCnt As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Since we are using the access keys in uppercase, convert
    '   any lowercase keys to uppercase before comparision
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32  'convert to uppercase
    End If
    
    'compare each with the stored access keys
    For iCnt = 0 To m_lTabCount - 1
        If m_aryTabs(iCnt).AccessKey = KeyAscii And iCnt <> m_lActiveTab And m_aryTabs(iCnt).Enabled Then
            'if we find the pressed key as access key for any tab,
            ' simply make that tab active
            If m_bIsRecursive Then Exit Sub
            m_bIsRecursive = True
            AutoRedraw = True
            ActiveTab = iCnt
            AutoRedraw = False
            m_bIsRecursive = False
            Exit For
        End If
    Next
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_AccessKeyPress", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Click()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent Click
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_Click", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_DblClick()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent DblClick
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_DblClick", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_GotFocus()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    m_lMouseOverTabIndex = m_lActiveTab
    m_lMouseOverTabIndex = ActiveTab
    m_IsFocused = True
    Call pShowHideFocus
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_GotFocus", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub


Private Sub UserControl_InitProperties()
    Dim iCnt As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   Initialize all properties
    m_bEnabled = True
    m_bShowFocusRect = m_def_bShowFocusRect
    m_bUseFocusedColor = m_def_bUseFocusedColor
    m_bUseMaskColor = m_def_bUseMaskColor
    m_bUseMouseWheelScroll = m_def_UseMouseWheelScroll
    m_enmPictureAlign = m_def_lPictureAlign
    m_enmPictureSize = m_def_lPictureSize
    m_enmStyle = m_def_lStyle
    m_lActiveTab = m_def_lActiveTab
    m_lActiveTabBackEndColor = m_def_lActiveTabBackEndColor
    m_lActiveTabBackStartColor = m_def_lActiveTabBackStartColor
    m_lActiveTabForeColor = m_def_lActiveTabForeColor
    m_lActiveTabHeight = m_def_lActiveTabHeight
    m_lBottomRightInnerBorderColor = m_def_lBottomRightInnerBorderColor
    m_lDisabledTabBackColor = m_def_lDisabledTabBackColor
    m_lDisabledTabForeColor = m_def_lDisabledTabForeColor
    m_lFocusedColor = m_def_lFocusedColor
    m_lHoverColor = m_def_lHoverColor
    m_lInActiveTabBackEndColor = m_def_lInActiveTabBackEndColor
    m_lInActiveTabBackStartColor = m_def_lInActiveTabBackStartColor
    m_lInActiveTabForeColor = m_def_lInActiveTabForeColor
    m_lInActiveTabHeight = m_def_lInActiveTabHeight
    m_lLastActiveTab = m_def_lLastActiveTab
    m_lOuterBorderColor = m_def_lOuterBorderColor
    m_lTabCount = m_def_lTabCount
    m_lTabStripBackColor = Ambient.BackColor
    m_lTopLeftInnerBorderColor = m_def_lTopLeftInnerBorderColor
    m_lXRadius = m_def_lXRadius
    m_lYRadius = m_def_lYRadius
    '   Default the active tab's font is bold
    Set m_oActiveTabFont = Ambient.Font
    Set m_oInActiveTabFont = Ambient.Font
    m_oActiveTabFont.Bold = True
    UserControl.MaskColor = m_def_lPictureMaskColor
    If UserControl.Parent.Width > 10000 Then
        m_lMoveOffset = UserControl.Parent.Width + 1000
    Else
        m_lMoveOffset = 10000
    End If
    '   Redim the tabs array
    ReDim m_aryTabs(m_lTabCount - 1)
    '   Initialize the tabs
    For iCnt = 0 To m_lTabCount - 1
        m_aryTabs(iCnt).Caption = m_def_sCaption & " " & iCnt
        m_aryTabs(iCnt).Enabled = m_def_bTabEnabled
    Next
    '   Set initial theme defaults based on m_enmTheme
    Call ResetColorsToDefault
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_InitProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    m_bIsRecursive = True
    AutoRedraw = True
    '   Continue if proper key pressed
    If (KeyCode = vbKeyPageDown And ((Shift And vbCtrlMask) > 0)) Or (KeyCode = vbKeyRight) Then
        '   If right key or the Ctrl+PageDown Key is pressed
        '
        '   If on some middle tab
        If m_lActiveTab < m_lTabCount - 1 Then
            If m_aryTabs(m_lActiveTab + 1).Enabled Then
                '   Increment tab by 1
                ActiveTab = m_lActiveTab + 1
            End If
        Else
            '   We are on the last tab
            If m_aryTabs(0).Enabled Then
                '   Set it to 0
                ActiveTab = 0
            End If
        End If
    ElseIf KeyCode = vbKeyPageUp And ((Shift And vbCtrlMask) > 0) Or (KeyCode = vbKeyLeft) Then
        '   If left key or the Ctrl+PageUp Key is pressed
        '
        '   If on some middle tab
        If m_lActiveTab > 0 Then
            If m_lTabCount > 1 Then
                If m_aryTabs(m_lActiveTab - 1).Enabled Then
                    '   Decrement tab by 1
                    ActiveTab = ActiveTab - 1
                End If
            End If
        Else
            '   We are on the first tab
            If m_lTabCount > 1 Then
                If m_aryTabs(m_lTabCount - 1).Enabled Then
                    '   Then set it to last tab
                    ActiveTab = m_lTabCount - 1
                End If
            End If
        End If
    End If
    AutoRedraw = False
    m_bIsRecursive = False
    '   Raise event, note: Byref arguments user can change there value to
    '   Control how tabs behave on key down
    RaiseEvent KeyDown(KeyCode, Shift)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_KeyDown", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyPress(KeyAscii)
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_KeyPress", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    RaiseEvent KeyUp(KeyCode, Shift)
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_KeyUp", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_LostFocus()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    m_lMouseOverTabIndex = m_lActiveTab
    m_IsFocused = False
    Call pShowHideFocus
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_LostFocus", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   This routine delegates the Handling of the Mouse Events to a Private Sub
    '
    '   Prevent recursion
    If m_bIsRecursive Then Exit Sub

    '   Raise event, for MouseDown
    RaiseEvent MouseDown(Button, Shift, X, Y)
    '   Only if left mouse down
    If Button = vbLeftButton Then
        m_bIsRecursive = True
        AutoRedraw = True
        '   Call theme specific HandleMouseDown
        Call pHandleMouseDown(Button, Shift, X, Y)
        AutoRedraw = False
        m_bIsRecursive = False
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_MouseDown", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   This routine delegates the Handling of the Mouse Events to a Private Sub
    '   raise event, for MouseMove
    RaiseEvent MouseMove(Button, Shift, X, Y)
    '   Call theme specific HandleMouseDown
    Call pHandleMouseMove(Button, Shift, X, Y)
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_MouseMove", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    '   This routine delegates the Handling of the Mouse Events to a Private Sub
    If m_bIsRecursive Then Exit Sub
    '   Raise event, for Mouse Up
    RaiseEvent MouseUp(Button, Shift, X, Y)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_MouseUp", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Paint()

    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    Refresh
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_Paint", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim iCnt As Long
    Dim iCCCount As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Read previously saved the property values
    With PropBag
        m_lActiveTab = .ReadProperty("ActiveTab", m_def_lActiveTab)
        m_bUseControlBox = .ReadProperty("UseControlBox", m_def_bUseControlBox)
        m_bShowFocusRect = .ReadProperty("ShowFocusRect", m_def_bShowFocusRect)
        m_bUseFocusedColor = .ReadProperty("UseFocusedColor", m_def_bUseFocusedColor)
        m_bEnabled = .ReadProperty("Enabled", m_def_bEnabled)
        m_bUseMaskColor = .ReadProperty("UseMaskColor", m_def_bUseMaskColor)
        m_bUseMouseWheelScroll = .ReadProperty("UseMouseWheelScroll", m_def_UseMouseWheelScroll)
        m_enmPictureAlign = .ReadProperty("PictureAlign", m_def_lPictureAlign)
        m_enmPictureSize = .ReadProperty("PictureSize", m_def_lPictureSize)
        m_enmStyle = .ReadProperty("TabStyle", m_def_lStyle)
        m_lActiveTabBackEndColor = .ReadProperty("ActiveTabBackEndColor", m_def_lActiveTabBackEndColor)
        m_lActiveTabBackStartColor = .ReadProperty("ActiveTabBackStartColor", m_def_lActiveTabBackStartColor)
        m_lActiveTabForeColor = .ReadProperty("ActiveTabForeColor", m_def_lActiveTabForeColor)
        m_lActiveTabHeight = .ReadProperty("ActiveTabHeight", m_def_lActiveTabHeight)
        m_lBottomRightInnerBorderColor = .ReadProperty("BottomRightInnerBorderColor", m_def_lBottomRightInnerBorderColor)
        m_lDisabledTabBackColor = .ReadProperty("DisabledTabBackColor", m_def_lDisabledTabBackColor)
        m_lDisabledTabForeColor = .ReadProperty("DisabledTabForeColor", m_def_lDisabledTabForeColor)
        m_lFocusedColor = .ReadProperty("FocusedColor", m_def_lFocusedColor)
        m_lHoverColor = .ReadProperty("HoverColor", m_def_lHoverColor)
        m_lInActiveTabBackEndColor = .ReadProperty("InActiveTabBackEndColor", m_def_lInActiveTabBackEndColor)
        m_lInActiveTabBackStartColor = .ReadProperty("InActiveTabBackStartColor", m_def_lInActiveTabBackStartColor)
        m_lInActiveTabForeColor = .ReadProperty("InActiveTabForeColor", m_def_lInActiveTabForeColor)
        m_lInActiveTabHeight = .ReadProperty("InActiveTabHeight", m_def_lInActiveTabHeight)
        m_lOuterBorderColor = .ReadProperty("OuterBorderColor", m_def_lOuterBorderColor)
        m_lTabCount = .ReadProperty("TabCount", m_def_lTabCount)
        m_lTabStripBackColor = .ReadProperty("TabStripBackColor", Ambient.BackColor)
        m_lTopLeftInnerBorderColor = .ReadProperty("TopLeftInnerBorderColor", m_def_lTopLeftInnerBorderColor)
        m_lXRadius = .ReadProperty("XRadius", m_def_lXRadius)
        m_lYRadius = .ReadProperty("YRadius", m_def_lYRadius)
        Set m_oActiveTabFont = .ReadProperty("ActiveTabFont", Ambient.Font)
        Set m_oInActiveTabFont = .ReadProperty("InActiveTabFont", Ambient.Font)
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        UserControl.ForeColor = .ReadProperty("ForeColor", &H0)
        UserControl.MaskColor = .ReadProperty("PictureMaskColor", m_def_lPictureMaskColor)
        m_lMoveOffset = .ReadProperty("TabOffset", 10000)
        '   Handle the Contained Controls
        Call HandleContainedControls(m_lActiveTab)
        '   Redim the tabs array
        ReDim m_aryTabs(m_lTabCount - 1)
        For iCnt = 0 To m_lTabCount - 1
            m_aryTabs(iCnt).Caption = .ReadProperty("TabCaption(" & iCnt & ")", m_def_sCaption & iCnt)
            m_aryTabs(iCnt).Enabled = .ReadProperty("TabEnabled(" & iCnt & ")", m_def_bTabEnabled)
            m_aryTabs(iCnt).AccessKey = .ReadProperty("TabAccessKey(" & iCnt & ")", 0)
            Set m_aryTabs(iCnt).TabPicture = .ReadProperty("TabPicture(" & iCnt & ")", Nothing)
            iCCCount = .ReadProperty("TabContCtrlCnt(" & iCnt & ")", 0)
        Next
        '   Get the controls tabstop info
        'pStoreOriginalTabStopValues
        Call ResetColorsToDefault
    End With
    '   Extract and set the acess keys for control
    Call pAssignAccessKeys
    '   If its not user mode then call the code to start subclassing
    If Ambient.UserMode Then                                                              'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        If bTrack Then
            '   OS supports mouse leave so subclass for it
            With UserControl
                '   Start subclassing the UserControl
                Call Subclass_Start(.hWnd)
                Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_LBUTTONDOWN, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_LBUTTONUP, MSG_AFTER)
                Call Subclass_AddMsg(.hWnd, WM_TIMER, MSG_BEFORE)
                
                '   Start subclassing the Parent form
                With .Parent
                    Call Subclass_Start(.hWnd)
                    Call Subclass_AddMsg(.hWnd, WM_ACTIVATE, MSG_AFTER)
                    Call Subclass_AddMsg(.hWnd, WM_MOVING, MSG_AFTER)
                    Call Subclass_AddMsg(.hWnd, WM_SIZING, MSG_AFTER)
                    Call Subclass_AddMsg(.hWnd, WM_EXITSIZEMOVE, MSG_AFTER)
                End With
                '   Now Store a Reference to the Host Objects
                If TypeOf .Parent Is Form Then
                    Set SDIHost = .Parent
                ElseIf TypeOf .Parent Is MDIForm Then
                    Set MDIHost = .Parent
                End If
                '   Set our Subclassing Flag
                bSubClass = True
            End With
        End If
    End If
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_ReadProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Show()

    '   Repaint the UserControl before it
    '   show in either thr IDE or at RunTime
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    

    Refresh
'    UserControl.Refresh
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_Show", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Terminate()
    '   The control is terminating - a good place to stop the subclasser
    On Error GoTo Catch
    ActiveTab = 0
    '   Call to free up system resources by deleting pictures etc
    Call pDestroyResources
    '   Stop all subclassing
    Call Subclass_StopAll
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim iCnt As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        Call .WriteProperty("TabCount", m_lTabCount, m_def_lTabCount)
        For iCnt = 0 To m_lTabCount - 1
            Call .WriteProperty("TabCaption(" & iCnt & ")", m_aryTabs(iCnt).Caption, m_def_sCaption & iCnt)
            Call .WriteProperty("TabEnabled(" & iCnt & ")", m_aryTabs(iCnt).Enabled, True)
            Call .WriteProperty("TabAccessKey(" & iCnt & ")", m_aryTabs(iCnt).AccessKey, 0)
            Call .WriteProperty("TabPicture(" & iCnt & ")", m_aryTabs(iCnt).TabPicture, Nothing)
        Next
        Call .WriteProperty("ActiveTab", m_lActiveTab, m_def_lActiveTab)
        Call .WriteProperty("ActiveTabBackEndColor", m_lActiveTabBackEndColor, m_def_lActiveTabBackEndColor)
        Call .WriteProperty("ActiveTabBackStartColor", m_lActiveTabBackStartColor, m_def_lActiveTabBackStartColor)
        Call .WriteProperty("ActiveTabFont", m_oActiveTabFont, Ambient.Font)
        Call .WriteProperty("ActiveTabForeColor", m_lActiveTabForeColor, m_def_lActiveTabForeColor)
        Call .WriteProperty("ActiveTabHeight", m_lActiveTabHeight, m_def_lActiveTabHeight)
        Call .WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
        Call .WriteProperty("BottomRightInnerBorderColor", m_lBottomRightInnerBorderColor, m_def_lBottomRightInnerBorderColor)
        Call .WriteProperty("DisabledTabBackColor", m_lDisabledTabBackColor, m_def_lDisabledTabBackColor)
        Call .WriteProperty("DisabledTabForeColor", m_lDisabledTabForeColor, m_def_lDisabledTabForeColor)
        Call .WriteProperty("Enabled", m_bEnabled, m_def_bEnabled)
        Call .WriteProperty("FocusedColor", m_lFocusedColor, m_def_lFocusedColor)
        Call .WriteProperty("ForeColor", UserControl.ForeColor, &H0)
        Call .WriteProperty("HoverColor", m_lHoverColor, m_def_lHoverColor)
        Call .WriteProperty("InActiveTabBackEndColor", m_lInActiveTabBackEndColor, m_def_lInActiveTabBackEndColor)
        Call .WriteProperty("InActiveTabBackStartColor", m_lInActiveTabBackStartColor, m_def_lInActiveTabBackStartColor)
        Call .WriteProperty("InActiveTabFont", m_oInActiveTabFont, Ambient.Font)
        Call .WriteProperty("InActiveTabForeColor", m_lInActiveTabForeColor, m_def_lInActiveTabForeColor)
        Call .WriteProperty("InActiveTabHeight", m_lInActiveTabHeight, m_def_lInActiveTabHeight)
        Call .WriteProperty("OuterBorderColor", m_lOuterBorderColor, m_def_lOuterBorderColor)
        Call .WriteProperty("PictureAlign", m_enmPictureAlign, m_def_lPictureAlign)
        Call .WriteProperty("PictureMaskColor", UserControl.MaskColor, m_def_lPictureMaskColor)
        Call .WriteProperty("PictureSize", m_enmPictureSize, m_def_lPictureSize)
        Call .WriteProperty("ShowFocusRect", m_bShowFocusRect, m_def_bShowFocusRect)
        Call .WriteProperty("TabStripBackColor", m_lTabStripBackColor, Ambient.BackColor)
        Call .WriteProperty("TabStyle", m_enmStyle, m_def_lStyle)
        Call .WriteProperty("TopLeftInnerBorderColor", m_lTopLeftInnerBorderColor, m_def_lTopLeftInnerBorderColor)
        Call .WriteProperty("UseControlBox", m_bUseControlBox, m_def_bUseControlBox)
        Call .WriteProperty("UseFocusedColor", m_bUseFocusedColor, m_def_bUseFocusedColor)
        Call .WriteProperty("UseMaskColor", m_bUseMaskColor, m_def_bUseMaskColor)
        Call .WriteProperty("UseMouseWheelScroll", m_bUseMouseWheelScroll, m_def_UseMouseWheelScroll)
        Call .WriteProperty("XRadius", m_lXRadius, m_def_lXRadius)
        Call .WriteProperty("YRadius", m_lYRadius, m_def_lYRadius)
        Call .WriteProperty("TabOffset", m_lMoveOffset, 10000)
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "AeroTab.UserControl_WriteProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Property Get XRadius() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    XRadius = m_lXRadius
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.XRadius", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let XRadius(iNewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_lXRadius = iNewValue
    '   Redraw
    Refresh
    PropertyChanged "XRadius"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.XRadius", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get YRadius() As Long
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    YRadius = m_lYRadius

Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.YRadius", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let YRadius(iNewValue As Long)
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    m_lYRadius = iNewValue
    '   Redraw
    Refresh
    PropertyChanged "YRadius"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "AeroTab.YRadius", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
  Dim Red As Long, Blue As Long, Green As Long
  Dim Delta As Long
  
  Blue = ((theColor \ &H10000) Mod &H100)
  Green = ((theColor \ &H100) Mod &H100)
  Red = (theColor And &HFF)
  Delta = &HFF - Base
  
  Blue = Base + Blue * Delta \ &HFF
  Green = Base + Green * Delta \ &HFF
  Red = Base + Red * Delta \ &HFF
  
  If Red > 255 Then Red = 255
  If Green > 255 Then Green = 255
  If Blue > 255 Then Blue = 255
  
  ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function
