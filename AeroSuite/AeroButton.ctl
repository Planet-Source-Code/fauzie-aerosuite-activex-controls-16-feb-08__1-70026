VERSION 5.00
Begin VB.UserControl AeroButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   73
   ToolboxBitmap   =   "AeroButton.ctx":0000
End
Attribute VB_Name = "AeroButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
' AeroButton ActiveX Control
' Based on dcButton by Noel A. Dacara [CodeId=65941]
'-----------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'-----------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'-----------------------------------------------------------------------------

Option Explicit

' Create transparent areas on the control
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Private Const RGN_OR As Long = 2
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Cursor tracking APIs
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win98 or later
Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win95 w/ IE 3.0
    Private Const TME_LEAVE     As Long = &H2
    Private Const WM_ACTIVATE   As Long = &H6
    Private Const WM_MOUSELEAVE As Long = &H2A3
    Private Const WM_MOUSEMOVE  As Long = &H200
    Private Const WM_NCACTIVATE As Long = &H86
    Private Type TRACKMOUSEEVENTTYPE
        cbSize      As Long
        dwFlags     As Long
        hwndTrack   As Long
        dwHoverTime As Long
    End Type
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Determines if a function is supported by a library
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

' Determines if the control's parent form/window is an MDI child window
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Const GWL_EXSTYLE    As Long = -20
    Private Const WS_EX_MDICHILD As Long = &H40&

' Determine if the system is in the NT platform (for unicode support)
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
    Private Const VER_PLATFORM_WIN32_NT As Long = 2
    Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion      As Long
        dwMinorVersion      As Long
        dwBuildNumber       As Long
        dwPlatformId        As Long
        szCSDVersion        As String * 128 ' Maintenance string for PSS usage
    End Type

' Drawing APIs (GDI32 library)
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Const SRCCOPY As Long = &HCC0020
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
#If USE_STANDARD Then
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
#End If
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Private Const PS_SOLID As Long = 0
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Const BI_RGB As Long = 0&
    Private Type BITMAPINFOHEADER
        biSize          As Long
        biWidth         As Long
        biHeight        As Long
        biPlanes        As Integer
        biBitCount      As Integer
        biCompression   As Long
        biSizeImage     As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed       As Long
        biClrImportant  As Long
    End Type
    Private Type RGBQUAD
        rgbBlue     As Byte
        rgbGreen    As Byte
        rgbRed      As Byte
        'rgbReserved As Byte ' Removed so that we can use RGBQUAD as
                             ' datatype for the bitmap data (lpBits)
    End Type
    Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
    End Type
Private Declare Function GetNearestColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Any) As Long ' Modified
Private Declare Function PatBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Private Const PATCOPY As Long = &HF00021
Private Declare Function Polyline Lib "gdi32.dll" (ByVal hDC As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' Drawing APIs (User32 library)
Private Declare Function CopyRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Long
    Private Type RECT
        Left    As Long
        Top     As Long
        Right   As Long
        Bottom  As Long
    End Type
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
    Private Const DI_NORMAL As Long = &H3
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' Drawing text in ansi/unicode
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long ' Modified
    Private Const DT_CALCRECT   As Long = &H400
    Private Const DT_CENTER     As Long = &H1
    Private Const DT_NOCLIP     As Long = &H100 ' Allow text to exceed specified drawing area (necessary for the vertical text effect)
    Private Const DT_WORDBREAK  As Long = &H10
    Private Const DT_CALCFLAG   As Long = DT_WORDBREAK Or DT_CALCRECT Or DT_NOCLIP Or DT_CENTER
    Private Const DT_DRAWFLAG   As Long = DT_WORDBREAK Or DT_NOCLIP Or DT_CENTER

' Load hand pointer as the control's cursor
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
    Private Const IDC_HAND As Long = 32649

' Restrict user from selecting other controls while spacebar is held down
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long

' SelfSub APIs and declarations
Private Declare Function CallWindowProcA Lib "user32.dll" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Const IDX_SHUTDOWN  As Long = 1
    Private Const IDX_HWND      As Long = 2
    Private Const IDX_WNDPROC   As Long = 9
    Private Const IDX_BTABLE    As Long = 11
    Private Const IDX_ATABLE    As Long = 12
    Private Const IDX_PARM_USER As Long = 13
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32.dll" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32.dll" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Const GWL_WNDPROC   As Long = -4
    Private Const WNDPROC_OFF   As Long = &H38
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long) ' Modified
    Private Const ALL_MESSAGES  As Long = -1
    Private Const MSG_ENTRIES   As Long = 32
    Private Enum eMsgWhen
        MSG_BEFORE = 1
        MSG_AFTER = 2
        MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
    End Enum
    Private z_ScMem  As Long
    Private z_Sc(64) As Long
    Private z_Funk   As Collection

' Events
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over the button."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
'Occurs when the user presses and then releases a mouse button over the button.
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user clicks over the button twice."
Attribute DblClick.VB_UserMemId = -601
'Occurs when the user clicks over the button twice.
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while the button has the focus."
Attribute KeyDown.VB_UserMemId = -602
'Occurs when the user presses a key while the button has the focus.
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
'Occurs when the user presses and releases an ANSI key.
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while the button has the focus."
Attribute KeyUp.VB_UserMemId = -604
'Occurs when the user releases a key while the button has the focus.
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the button has the focus."
Attribute MouseDown.VB_UserMemId = -605
'Occurs when the user presses the mouse button while the button has the focus.
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occrus when the cursor moves around the button for the first time."
'Occrus when the cursor moves around the button for the first time.
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the cursor leaves/moves outside the button."
'Occurs when the cursor leaves/moves outside the button.
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the cursor moves over the button."
Attribute MouseMove.VB_UserMemId = -606
'Occurs when the cursor moves over the button.
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the button has the focus."
Attribute MouseUp.VB_UserMemId = -607
'Occurs when the user releases the mouse button while the button has the focus.

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private KeyCode, Shift, KeyAscii, Button, X, Y
#End If

Public Enum eButtonShapes
    elsCutNone  ' Normal
    elsCutLeft  '
    elsCutRight '
    elsCutSides ' Both left & right
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private elsCutNone, elsCutLeft, elsCutRight, elsCutSides
#End If

Private Enum eButtonStates
    elsNormal   ' Normal state
    elsHot      ' Cursor is over the control
    elsDown     ' Mouse/key is down
    elsDisabled ' Disabled state
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private elsNormal, elsHot, elsDown, elsDisabled
#End If

Public Enum ePicAlignments
    elaBehindText
    elaBottomEdge
    elaBottomOfCaption
    elaLeftEdge
    elaLeftOfCaption
    elaRightEdge
    elaRightOfCaption
    elaTopEdge
    elaTopOfCaption
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private elaBehindText, elaBottomEdge, elaBottomOfCaption
    Private elaLeftEdge, elaLeftOfCaption, elaRightEdge
    Private elaRightOfCaption, elaTopEdge, elaTopOfCaption
#End If

Public Enum ePictureSizes
    elpNormal ' Use original size of main picture
    elp16x16  ' Small icon size
    elp24x24  ' Standard toolbar icon size
    elp32x32  ' Standard icon size
    elp48x48  ' Explorer thumbnail size
    elpCustom ' Use picture size defined in property
End Enum

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private elpNormal, elp16x16, elp24x24, elp32x32, elp48x48, elpCustom
#End If

Private Type tButtonColors  ' Safe button colors (Translated)
    BackColor   As Long     ' Normal
    DownColor   As Long     ' Down state
    FocusBorder As Long     ' Focus state
    ForeColor   As Long     ' Normal text color
    GrayColor   As Long     ' Disabled background color
    GrayText    As Long     ' Disabled text and border color
    HoverColor  As Long     ' Hot state
    MaskColor   As Long     ' Mask color for the picture
    StartColor  As Long     ' Start color for gradient effect
End Type

Private Type tButtonProperties ' Cached button properties
    BackColor   As Long
    Caption     As String
    CheckBox    As Boolean
    Enabled     As Boolean
    ForeColor   As Long
    HandPointer As Boolean
    MaskColor   As Long
    PicAlign    As ePicAlignments
    PicDown     As StdPicture
    PicHot      As StdPicture
    PicNormal   As StdPicture
    PicOpacity  As Single
    PicSize     As ePictureSizes
    PicSizeH    As Long
    PicSizeW    As Long
    Shape       As eButtonShapes
    UseMask     As Boolean
    Value       As Boolean
End Type

Private Type tButtonSettings        ' Runtime/designtime settings
    Button      As Integer          ' Last button clicked
    Caption     As RECT             ' Area where to draw caption
    Cursor      As Long             ' Handle to the sytem hand pointer
    Default     As Boolean          ' Is control set as DEFAULT button of a form?
    Focus       As RECT             ' Area where to draw FocusRect/ Shine object
    HasFocus    As Boolean          ' Is the control currently in focus
    Height      As Long             ' ScaleHeight (pixels)
    Picture     As RECT             ' Area where to draw icon/picture
    State       As eButtonStates    ' Current drawing state of the button
    Width       As Long             ' ScaleWidth (pixels)
End Type

Private Type tRGB
    R As Long
    G As Long
    b As Long
End Type

' Variables
Private m_bButtonHasFocus   As Boolean
Private m_bButtonIsDown     As Boolean
Private m_bCalculateRects   As Boolean
Private m_bControlHidden    As Boolean
Private m_bIsTracking       As Boolean
Private m_bIsPlatformNT     As Boolean
Private m_bMouseIsDown      As Boolean
Private m_bMouseOnButton    As Boolean
Private m_bParentActive     As Boolean
Private m_bRedrawOnResize   As Boolean
Private m_bSpacebarIsDown   As Boolean
Private m_bTrackHandler32   As Boolean

Private m_tButtonProperty   As tButtonProperties
Private m_tButtonColors     As tButtonColors
Private m_tButtonSettings   As tButtonSettings

' //-- Properties --//

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used for the button."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
'   Returns/sets the background color used for the button.
    BackColor = m_tButtonProperty.BackColor
    
End Property

Public Property Let BackColor(Value As OLE_COLOR)
    m_tButtonProperty.BackColor = Value
    SetButtonColors ' Update color changes
    DrawButton Force:=True
    PropertyChanged "BackColor"
    
End Property

Public Property Get ButtonShape() As eButtonShapes
Attribute ButtonShape.VB_Description = "Returns/sets a value to determine the shape used to draw the button."
Attribute ButtonShape.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets a value to determine the shape used to draw the button.
    ButtonShape = m_tButtonProperty.Shape
    
End Property

Public Property Let ButtonShape(Value As eButtonShapes)
    m_tButtonProperty.Shape = Value
    Me.Refresh
    PropertyChanged "ButtonShape"
    
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the button."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
'   Returns/sets the text displayed in the button.
    Caption = m_tButtonProperty.Caption
    
End Property

Public Property Let Caption(Value As String)
    m_tButtonProperty.Caption = Value
    UserControl.AccessKeys = GetAccessKey(Value)
    Me.Refresh
    PropertyChanged "Caption"
    
End Property

Public Property Get CheckBoxMode() As Boolean
Attribute CheckBoxMode.VB_Description = "Returns/sets the type of control the button will observe."
Attribute CheckBoxMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
'   Returns/sets the type of control the button will observe.
    CheckBoxMode = m_tButtonProperty.CheckBox
    
End Property

Public Property Let CheckBoxMode(Value As Boolean)
    m_tButtonProperty.CheckBox = Value
    
    If (Not Value) And (m_tButtonProperty.Value) Then
        m_tButtonProperty.Value = False ' Normal state
        PropertyChanged "Value"
    End If
    
    DrawButton Force:=True
    PropertyChanged "CheckBox"
    
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value to determine whether the button can respond to events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
'   Returns/sets a value to determine whether the button can respond to events.
    Enabled = m_tButtonProperty.Enabled
    
End Property

Public Property Let Enabled(Value As Boolean)
    m_tButtonProperty.Enabled = Value
    UserControl.Enabled = Value
    
    If (Not Value) Then ' Disabled
        DrawButton elsDisabled
    ElseIf (m_bMouseOnButton) Then
        DrawButton elsHot
    Else
        DrawButton elsNormal
    End If
    
    PropertyChanged "Enabled"
    
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/sets the Font used to display text on the button."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
'   Returns/sets the Font used to display text on the button.
    Set Font = UserControl.Font
    
End Property

Public Property Set Font(Value As StdFont)
    Set UserControl.Font = Value
    Me.Refresh
    PropertyChanged "Font"
    
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font style."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontBold.VB_MemberFlags = "400"
'   Returns/sets bold font style.
    FontBold = UserControl.FontBold
    
End Property

Public Property Let FontBold(Value As Boolean)
    UserControl.FontBold = Value
    Me.Refresh
    
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font style."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontItalic.VB_MemberFlags = "400"
'   Returns/sets italic font style.
    FontItalic = UserControl.FontItalic
    
End Property

Public Property Let FontItalic(Value As Boolean)
    UserControl.FontItalic = Value
    Me.Refresh
    
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font used for the button caption."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
'   Specifies the name of the font used for the button caption.
    FontName = UserControl.FontName
    
End Property

Public Property Let FontName(Value As String)
    UserControl.FontName = Value
    Me.Refresh
    
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font used for the button caption."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontSize.VB_MemberFlags = "400"
'   Specifies the size (in points) of the font used for the button caption.
    FontSize = UserControl.FontSize
    
End Property

Public Property Let FontSize(Value As Single)
    UserControl.FontSize = Value
    Me.Refresh
    
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font style."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontStrikethru.VB_MemberFlags = "400"
'   Returns/sets strikethrough font style.
    FontStrikethru = UserControl.FontStrikethru
    
End Property

Public Property Let FontStrikethru(Value As Boolean)
    UserControl.FontStrikethru = Value
    Me.Refresh
    
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font style."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontUnderline.VB_MemberFlags = "400"
'   Returns/sets underline font style.
    FontUnderline = UserControl.FontUnderline
    
End Property

Public Property Let FontUnderline(Value As Boolean)
    UserControl.FontUnderline = Value
    Me.Refresh
    
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the text color of the button caption."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
'   Returns/sets the text color of the button caption.
    ForeColor = m_tButtonProperty.ForeColor
    
End Property

Public Property Let ForeColor(Value As OLE_COLOR)
    m_tButtonProperty.ForeColor = Value
    
    If (m_tButtonProperty.Enabled) Then ' Disabled text uses its own color
        SetButtonColors ' Update color changes
        DrawButton Force:=True
    End If
    
    PropertyChanged "ForeColor"
    
End Property

Public Property Get HandPointer() As Boolean
Attribute HandPointer.VB_Description = "Returns/sets a value to determine whether the control uses the system's hand pointer as its cursor."
Attribute HandPointer.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets a value to determine whether the control uses the system's hand pointer as its cursor.
    HandPointer = m_tButtonProperty.HandPointer
    
End Property

Public Property Let HandPointer(Value As Boolean)
    m_tButtonProperty.HandPointer = Value
    PropertyChanged "HandPointer"
    
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle that uniquely identifies the control."
Attribute hWnd.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
'   Returns a handle that uniquely identifies the control.
    hWnd = UserControl.hWnd
    
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in a button's picture to be transparent."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a color in a button's picture to be transparent.
    MaskColor = m_tButtonProperty.MaskColor
    
End Property

Public Property Let MaskColor(Value As OLE_COLOR)
    m_tButtonProperty.MaskColor = Value
    m_tButtonColors.MaskColor = TranslateColor(Value)
    DrawButton Force:=True
    PropertyChanged "MaskColor"
    
End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon for the button."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Sets a custom mouse icon for the button.
    Set MouseIcon = UserControl.MouseIcon
    
End Property

Public Property Set MouseIcon(Value As IPictureDisp)
    Set UserControl.MouseIcon = Value       ' Set new cursor
                                            '
    If (Value Is Nothing) Then              '
        Me.MousePointer = 0 ' vbDefault     ' Apply appropriate
    Else                                    ' mouse pointer setting
        Me.MousePointer = 99 ' vbCustom     ' automatically
    End If                                  '
    
    PropertyChanged "MouseIcon"
    
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when cursor over the button."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets the type of mouse pointer displayed when cursor over the button.
    MousePointer = UserControl.MousePointer
    
End Property

Public Property Let MousePointer(Value As MousePointerConstants)
    If (Not Value = vbCustom) And (m_tButtonProperty.HandPointer) Then
        Me.HandPointer = False          ' Unset hand pointer option
    End If                              ' on demand
                                        '
    UserControl.MousePointer = Value    ' Set new mouse pointer
    PropertyChanged "MousePointer"
    
End Property

Public Property Get PictureAlignment() As ePicAlignments
Attribute PictureAlignment.VB_Description = "Returns/sets a value to determine where to draw the picture in the button."
Attribute PictureAlignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value to determine where to draw the picture in the button.
    PictureAlignment = m_tButtonProperty.PicAlign
    
End Property

Public Property Let PictureAlignment(Value As ePicAlignments)
    m_tButtonProperty.PicAlign = Value
    Me.Refresh
    PropertyChanged "PicAlign"
    
End Property

Public Property Get PictureDown() As StdPicture
Attribute PictureDown.VB_Description = "Returns/sets the picture displayed when the control is pressed down or in checked state."
Attribute PictureDown.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the picture displayed when the control is pressed down or in checked state.
    Set PictureDown = m_tButtonProperty.PicDown
    
End Property

Public Property Set PictureDown(Value As StdPicture)
'   Main picture (PictureNormal) must be set first before this property (PictureDown)
'   Else the specified value will be set as the main picture instead.
    
    If (Value Is Nothing) Then
        Set m_tButtonProperty.PicDown = Nothing
        GoTo Jmp_Skip
    End If
    
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Set Me.PictureNormal = Value
        Exit Property
    Else
        If (m_tButtonProperty.PicSize = elpNormal) Then
            If (Not m_tButtonProperty.PicNormal.Width = Value.Width) Or _
               (Not m_tButtonProperty.PicNormal.Height = Value.Height) Then
                ' If pictures do not have the same sizes (width or height)
                ' Use main picture's size as the standard for all pictures
                Me.PictureSize = elpCustom
            End If
        End If
        Set m_tButtonProperty.PicDown = Value
    End If
    
Jmp_Skip:
    If (m_bButtonIsDown) Then
        DrawButton Force:=True
    End If
    PropertyChanged "PicDown"
    
End Property

Public Property Get PictureHot() As StdPicture
Attribute PictureHot.VB_Description = "Returns/sets the picture displayed when the cursor is over the control."
Attribute PictureHot.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the picture displayed when the cursor is over the control.
    Set PictureHot = m_tButtonProperty.PicHot
    
End Property

Public Property Set PictureHot(Value As StdPicture)
'   Main picture (PictureNormal) must be set first before this property (PictureHot)
'   Else the specified value will be set as the main picture instead.
    
    If (Value Is Nothing) Then
        Set m_tButtonProperty.PicHot = Nothing
        GoTo Jmp_Skip
    End If
    
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Set Me.PictureNormal = Value
        Exit Property
    Else
        If (m_tButtonProperty.PicSize = elpNormal) Then
            If (Not m_tButtonProperty.PicNormal.Width = Value.Width) Or _
               (Not m_tButtonProperty.PicNormal.Height = Value.Height) Then
                ' If pictures do not have the same sizes (width or height)
                ' Use main picture's size as the standard for all pictures
                Me.PictureSize = elpCustom
            End If
        End If
        Set m_tButtonProperty.PicHot = Value
    End If
    
Jmp_Skip:
    If (m_bMouseOnButton) Then
        DrawButton Force:=True
    End If
    PropertyChanged "PicHot"
    
End Property

Public Property Get PictureNormal() As StdPicture
Attribute PictureNormal.VB_Description = "Returns/sets the picture displayed on a normal state button."
Attribute PictureNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the picture displayed on a normal state button.
    Set PictureNormal = m_tButtonProperty.PicNormal
    
End Property

Public Property Set PictureNormal(Value As StdPicture)
    
    If (Value Is Nothing) Then
        ' Cannot work without the main picture
        Set Me.PictureDown = Nothing
        Set Me.PictureHot = Nothing
            m_tButtonProperty.PicOpacity = 0
    ElseIf (m_tButtonProperty.PicOpacity = 0) Then
        m_tButtonProperty.PicOpacity = 1
    End If
    
    Set m_tButtonProperty.PicNormal = Value
        Me.PictureSize = m_tButtonProperty.PicSize ' Update picture sizes
        
    Me.Refresh
    PropertyChanged "PicNormal"
    PropertyChanged "PicOpacity"
    
End Property

Public Property Get PictureOpacity() As Long
Attribute PictureOpacity.VB_Description = "Returns/sets a value in percent how the pictures will be blended to the button."
Attribute PictureOpacity.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value in percent how the pictures will be blended to the button.
    PictureOpacity = m_tButtonProperty.PicOpacity * 100
    
End Property

Public Property Let PictureOpacity(Value As Long)
'   Below 10% means the picture will not be visible at all (or almost)
'   So why blend it this way if you could just remove the picture instead?
    m_tButtonProperty.PicOpacity = TranslateNumber(Value, 10, 100)
    m_tButtonProperty.PicOpacity = m_tButtonProperty.PicOpacity / 100
    DrawButton Force:=True
    PropertyChanged "PicOpacity"
    
End Property

Public Property Get PictureSize() As ePictureSizes
Attribute PictureSize.VB_Description = "Returns/sets a value to determine the size of the picture to draw."
Attribute PictureSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets a value to determine the size of the picture to draw.
    PictureSize = m_tButtonProperty.PicSize
    
End Property

Public Property Let PictureSize(Value As ePictureSizes)
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        m_tButtonProperty.PicSize = elpNormal
        m_tButtonProperty.PicSizeH = 0
        m_tButtonProperty.PicSizeW = 0
        GoTo Jmp_Skip
    End If
    
    m_tButtonProperty.PicSize = Value
    
    With m_tButtonProperty
    Select Case Value
        Case elpNormal
            .PicSizeH = ScaleY(m_tButtonProperty.PicNormal.Height, 8, 3) ' 8 = vbHimetric; 3 = vbPixels
            .PicSizeW = ScaleY(m_tButtonProperty.PicNormal.Width, 8, 3)
        Case elp16x16
            .PicSizeH = 16
            .PicSizeW = 16
        Case elp24x24
            .PicSizeH = 24
            .PicSizeW = 24
        Case elp32x32
            .PicSizeH = 32
            .PicSizeW = 32
        Case elp48x48
            .PicSizeH = 48
            .PicSizeW = 48
    End Select
    End With
    
Jmp_Skip:
    Me.Refresh
    
    PropertyChanged "PicSize"
    PropertyChanged "PicSizeH"
    PropertyChanged "PicSizeW"
    
End Property

Public Property Get PictureSizeH() As Long
Attribute PictureSizeH.VB_Description = "Returns/sets the standard/custom height of the defined pictures in pixels."
Attribute PictureSizeH.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the standard/custom height of the defined pictures in pixels.
    PictureSizeH = m_tButtonProperty.PicSizeH
    
End Property

Public Property Let PictureSizeH(Value As Long)
    m_tButtonProperty.PicSize = elpCustom   ' If modified then set size to custom
    m_tButtonProperty.PicSizeH = Value      '
    Me.Refresh                              '
    PropertyChanged "PicSize"               ' Notify the container that
    PropertyChanged "PicSizeH"              ' properties has been changed
    
End Property

Public Property Get PictureSizeW() As Long
Attribute PictureSizeW.VB_Description = "Returns/sets the standard/custom width of the defined pictures in pixels."
Attribute PictureSizeW.VB_ProcData.VB_Invoke_Property = ";Appearance"
'   Returns/sets the standard/custom width of the defined pictures in pixels.
    PictureSizeW = m_tButtonProperty.PicSizeW
    
End Property

Public Property Let PictureSizeW(Value As Long)
    m_tButtonProperty.PicSize = elpCustom   ' If modified then set size to custom
    m_tButtonProperty.PicSizeW = Value      '
    Me.Refresh                              '
    PropertyChanged "PicSize"               ' Notify the container that
    PropertyChanged "PicSizeW"              ' properties has been changed
    
End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value to determine whether to use MaskColor to create transparent areas of the picture."
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets a value to determine whether to use MaskColor to create transparent areas of the picture.
    UseMaskColor = m_tButtonProperty.UseMask
    
End Property

Public Property Let UseMaskColor(Value As Boolean)
    m_tButtonProperty.UseMask = Value
    DrawButton Force:=True
    PropertyChanged "UseMask"
    
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value or state of the button."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
'   Returns/sets the value or state of the button.
    Value = m_tButtonProperty.Value
    
End Property

Public Property Let Value(Value As Boolean)
    m_tButtonProperty.Value = Value
    
    If (Value) And (Not m_tButtonProperty.CheckBox) Then
        If (Ambient.UserMode) Then
            m_tButtonSettings.Button = 1
            UserControl_Click ' Trigger click event
        Else
            m_tButtonProperty.Value = False
        End If
    Else ' Value is False or CheckBoxMode is True
        DrawButton Force:=True
    End If
    
    PropertyChanged "Value"
    
End Property

' //-- Public Procedures --//

Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Shows information about the control and its author."
Attribute AboutBox.VB_UserMemId = -552
    fAbout.Show vbModal
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of the button."
Attribute Refresh.VB_UserMemId = -550
'   Forces a complete repaint of the button.
    m_bCalculateRects = True
    DrawButton Force:=True
    
End Sub

' //-- Private Procedures --//

Private Function BlendColors( _
        Color1 As Long, _
        Color2 As Long, _
        Optional PercentInDecimal As Single = 0.5) As Long
'   Combines two colors together by how many percent.
    
    Dim Color1RGB As tRGB
    Dim Color2RGB As tRGB
    Dim Color3RGB As tRGB
    
    Color1RGB = GetRGB(Color1)
    Color2RGB = GetRGB(Color2)
    
    Color3RGB.R = Color1RGB.R + (Color2RGB.R - Color1RGB.R) * PercentInDecimal ' Percent should already
    Color3RGB.G = Color1RGB.G + (Color2RGB.G - Color1RGB.G) * PercentInDecimal ' be translated.
    Color3RGB.b = Color1RGB.b + (Color2RGB.b - Color1RGB.b) * PercentInDecimal ' Ex. 50% -> 50 / 100 = 0.5
    
    BlendColors = RGBEx(Color3RGB.R, Color3RGB.G, Color3RGB.b)
    
End Function

Private Function BlendRGBQUAD( _
        RGB1 As RGBQUAD, _
        RGB2 As RGBQUAD, _
        Optional PercentInDecimal As Single = 0.5) As RGBQUAD
'   Blend two colors particularly RGBQUAD structures together by how many percent.
    
    Dim RGB3 As tRGB
        RGB3.R = RGB2.rgbRed            ' An overflow will cause to occur
        RGB3.G = RGB2.rgbGreen          ' if we directly subtract RGBQUAD
        RGB3.b = RGB2.rgbBlue           ' values from each other.
                                        '
        RGB3.R = RGB3.R - RGB1.rgbRed   ' So instead, we store the first RGBQUAD
        RGB3.G = RGB3.G - RGB1.rgbGreen ' structure as the result then after
        RGB3.b = RGB3.b - RGB1.rgbBlue  ' we can safely subtract it the other.
        
        RGB3.R = RGB1.rgbRed + RGB3.R * PercentInDecimal    ' Percent should already
        RGB3.G = RGB1.rgbGreen + RGB3.G * PercentInDecimal  ' be translated.
        RGB3.b = RGB1.rgbBlue + RGB3.b * PercentInDecimal   ' Ex. 50% -> 50 / 100 = 0.5
        
        BlendRGBQUAD.rgbRed = RGB3.R
        BlendRGBQUAD.rgbGreen = RGB3.G
        BlendRGBQUAD.rgbBlue = RGB3.b
        
End Function

Private Sub CalculateRects( _
        Optional ByVal BorderH As Long = 4, _
        Optional ByVal BorderW As Long = 4)
'   Calculate areas where to draw the caption and the icon/picture
'   Uses the integer divisions because it does not round off the result when calculating,
'   thus making the result more accurate when performing specific calulations.
'   Although in theory, integer division is slower than the regular division, don't worry,
'   this procedure is called only when the control has encountered a decisive change.
    
    m_bCalculateRects = False
    
    Dim bh As Long          ' Height of button
    Dim bw As Long          ' Width of button
                            '
    Dim pr As RECT          ' Draw area of picture
                            '
    Dim s1 As Single        ' Temporary variables
    Dim s2 As Single        '
                            '
    Dim tn As Long          ' Length of text
    Dim tr As RECT          ' Draw area of caption
    Dim tx As String        ' Caption text
                            '
    With m_tButtonSettings  '
        bh = .Height        ' Get button size
        bw = .Width         '
                            '
        tx = m_tButtonProperty.Caption
        tn = Len(tx)        '
                            '
        SetRect .Focus, BorderW, BorderH, .Width - BorderW, .Height - BorderH
    End With                '
                            '
    If (tn > 0) Then        ' Set estimated drawing area of caption
        SetRect tr, 0, 0, bw, bh                    '
        If (m_bIsPlatformNT) Then                   '
            DrawTextW hDC, StrPtr(tx), tn, tr, DT_CALCFLAG
        Else                                        '
            DrawText hDC, tx, tn, tr, DT_CALCFLAG   ' Get width & height of area
        End If                                      ' that fits the text/caption
    End If                                          '
                                                    ' Move caption area to the center
    OffsetRect tr, (bw - tr.Right) \ 2, (bh - tr.Bottom) \ 2
                                                    '
    CopyRect m_tButtonSettings.Caption, tr          ' Save changes
                                                    '
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Exit Sub                                    ' Skip icon alignment when
    End If                                          ' no picture is to be aligned
    
    SetRect pr, 0, 0, m_tButtonProperty.PicSizeW, m_tButtonProperty.PicSizeH
    
    If (tn > 0) Then    ' Check if a caption is specified
        tn = 2          ' If set, then tn will be set to contain the value in
    Else                ' pixels the caption and the picture will be apart
        tn = 0          ' Else, just set to zero to retain the icon/picture
    End If              ' in place (center) when no caption is specified :)
    
    ' Note: OffsetRect moves the RECT coordinates by how many pixels
    '       from its current position. (it does not change its size)
    
    Select Case m_tButtonProperty.PicAlign                          ' Picture on center
        Case elaBehindText                                          ' but behind text
            OffsetRect pr, (bw - pr.Right) \ 2, (bh - pr.Bottom) \ 2
                                                                    '
        Case elaBottomEdge, elaBottomOfCaption                      ' Picture on
            OffsetRect pr, (bw - pr.Right) \ 2, 0                   ' bottom portion
            OffsetRect tr, 0, -tr.Top                               ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = elaBottomEdge) Then    '
                OffsetRect pr, 0, bh - pr.Bottom - BorderH          '
                If (tn > 0) Then                                    '
                    OffsetRect tr, 0, pr.Top - tr.Bottom '- tn      '
                    tn = tr.Top - BorderH                           '
                    If (tn > 1) Then                                '
                        OffsetRect tr, 0, -(tn \ 2)                 '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, 0, (bh - pr.Bottom) \ 2              '
            Else                                                    '
                OffsetRect pr, 0, tr.Bottom '+ tn                   '
                tn = (bh - pr.Bottom) \ 2                           '
                OffsetRect pr, 0, tn                                '
                OffsetRect tr, 0, tn                                '
            End If                                                  '
                                                                    '
        Case elaLeftEdge, elaLeftOfCaption                          ' Picture on
            OffsetRect pr, 0, (bh - pr.Bottom) \ 2                  ' left portion
            OffsetRect tr, -tr.Left, 0                              ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = elaLeftEdge) Then      '
                OffsetRect pr, BorderW, 0                           '
                If (tn > 0) Then                                    '
                    OffsetRect tr, pr.Right + tn, 0                 '
                    tn = bw - BorderW - tr.Right                    '
                    If (tn > 1) Then                                '
                        OffsetRect tr, tn \ 2, 0                    '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, (bw - pr.Right) \ 2, 0               '
            Else                                                    '
                OffsetRect tr, pr.Right + tn, 0                     '
                tn = (bw - tr.Right) \ 2                            '
                OffsetRect tr, tn, 0                                '
                OffsetRect pr, tn, 0                                '
            End If                                                  '
                                                                    '
        Case elaRightEdge, elaRightOfCaption                        ' Picture on
            OffsetRect pr, 0, (bh - pr.Bottom) \ 2                  ' right portion
            OffsetRect tr, -tr.Left, 0                              ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = elaRightEdge) Then     '
                OffsetRect pr, bw - pr.Right - BorderW, 0           '
                If (tn > 0) Then                                    '
                    OffsetRect tr, pr.Left - tr.Right - tn, 0       '
                    tn = tr.Left - BorderW                          '
                    If (tn > 1) Then                                '
                        OffsetRect tr, -(tn \ 2), 0                 '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, (bw - pr.Right) \ 2, 0               '
            Else                                                    '
                OffsetRect pr, tr.Right + tn, 0                     '
                tn = (bw - pr.Right) \ 2                            '
                OffsetRect pr, tn, 0                                '
                OffsetRect tr, tn, 0                                '
            End If                                                  '
                                                                    '
        Case elaTopEdge, elaTopOfCaption                            ' Picture on
            OffsetRect pr, (bw - pr.Right) \ 2, 0                   ' top portion
            OffsetRect tr, 0, -tr.Top                               ' of caption
                                                                    '
            If (m_tButtonProperty.PicAlign = elaTopEdge) Then       '
                OffsetRect pr, 0, BorderH                           '
                If (tn > 0) Then                                    '
                    OffsetRect tr, 0, pr.Bottom '+ tn               '
                    tn = bh - tr.Bottom - BorderH                   '
                    If (tn > 1) Then                                '
                        OffsetRect tr, 0, tn \ 2                    '
                    End If                                          '
                End If                                              '
            ElseIf (tn = 0) Then                                    '
                OffsetRect pr, 0, (bh - pr.Bottom) \ 2              '
            Else                                                    '
                OffsetRect tr, 0, pr.Bottom '+ tn                   '
                tn = (bh - tr.Bottom) \ 2                           '
                OffsetRect tr, 0, tn                                '
                OffsetRect pr, 0, tn                                '
            End If                                                  '
            
    End Select
    
    CopyRect m_tButtonSettings.Picture, pr  ' Save changes to drawing areas
    CopyRect m_tButtonSettings.Caption, tr  '
    
End Sub

Private Sub CreateButtonRegion(EllipseW As Long, EllipseH As Long)
'   Create a button region with rounded corners.
    
    With m_tButtonSettings
        Dim hRgn  As Long
        Dim hRgn2 As Long
            ' Create button region with rounded corners as requested
            hRgn = CreateRoundRectRgn(0, 0, .Width + 1, .Height + 1, EllipseW, EllipseH)
        
        If (m_tButtonProperty.Shape = elsCutLeft) Or _
           (m_tButtonProperty.Shape = elsCutSides) Then
            hRgn2 = CreateRectRgn(0, 0, .Width / 2, .Height + 1)
            CombineRgn hRgn, hRgn, hRgn2, RGN_OR
            DeleteObject hRgn2
        End If
        
        If (m_tButtonProperty.Shape = elsCutRight) Or _
           (m_tButtonProperty.Shape = elsCutSides) Then
            hRgn2 = CreateRectRgn(.Width / 2, 0, .Width + 1, .Height + 1)
            CombineRgn hRgn, hRgn, hRgn2, RGN_OR
            DeleteObject hRgn2
        End If
        
        SetWindowRgn hWnd, hRgn, True ' Set new window region
        DeleteObject hRgn
    End With
    
End Sub

Private Sub DrawButton(Optional State As eButtonStates = -1, Optional Force As Boolean)
'   Draws the button itself
    If (m_bControlHidden) Then Exit Sub ' Why draw if control was hidden?
    
    If (State = -1) Then
        State = m_tButtonSettings.State ' Get current button state
    End If
    
    If (Not Force) Then ' Always update when forced
        ' Check if same as previous state, if so then exit
        If (State = m_tButtonSettings.State) Then
            If (m_tButtonSettings.HasFocus = m_bButtonHasFocus) Then
                ' If the button has changed its focus state from
                ' last drawing then button display must be updated
                Exit Sub
            End If
        End If
    End If
    
    UserControl.Cls
    
    m_tButtonSettings.HasFocus = m_bButtonHasFocus
    m_tButtonSettings.State = State ' Store current button state
    
    DrawVistaButtonStyle State
    
End Sub

Private Sub DrawVistaButtonStyle(DrawState As eButtonStates)
'   Drawing procedure for the Vista button styles.
    
    If (m_bCalculateRects) Then ' Recreate region
        CreateButtonRegion 3, 3
        CalculateRects
    End If
    
    With m_tButtonProperty
        If (.CheckBox And .Value And .Enabled) Then
            If (DrawState = elsDown) Then
                DrawState = elsNormal
            End If
        End If
    End With
    
    Dim Color1 As Long
    Dim Color2 As Long
    
    With m_tButtonSettings
        Select Case DrawState ' Draw button background
            Case elsNormal, elsHot
                If (m_tButtonProperty.Value) Then
                    Color2 = ShiftColor(m_tButtonColors.BackColor, -0.1)
                    DrawGradientEx Color2, m_tButtonColors.StartColor
                    Color2 = BlendColors(Color2, m_tButtonColors.StartColor)
                Else
                    If DrawState = elsNormal Then
                        DrawGradientEx ShiftColor(m_tButtonColors.BackColor, 0.08), ShiftColor(m_tButtonColors.BackColor, 0.02), , , 2, , ScaleHeight / 2
                        DrawGradientEx m_tButtonColors.BackColor, ShiftColor(m_tButtonColors.BackColor, -0.1), , , ScaleHeight / 2, , ScaleHeight - 2
                        Color2 = BlendColors(m_tButtonColors.StartColor, m_tButtonColors.BackColor)
                    Else
                        DrawGradientEx RGB(234, 246, 253), m_tButtonColors.HoverColor, , , 2, , ScaleHeight / 2
                        DrawGradientEx RGB(190, 230, 253), RGB(167, 217, 245), , , ScaleHeight / 2, , ScaleHeight - 2
                        Color2 = BlendColors(m_tButtonColors.StartColor, m_tButtonColors.BackColor)
                    End If
                End If
            Case elsDown
                Color2 = m_tButtonColors.DownColor                          '
                FillButtonEx Color2                                         '
                
                DrawGradientEx RGB(229, 244, 252), m_tButtonColors.DownColor, , , 2, , ScaleHeight / 2
                DrawGradientEx RGB(152, 209, 239), RGB(104, 179, 219), , , ScaleHeight / 2, , ScaleHeight - 2
            
                DrawLine RGB(158, 176, 186), 1, 1, 1, ScaleHeight
                DrawLine RGB(158, 176, 186), ScaleWidth - 2, 1, ScaleWidth - 2, ScaleHeight
                DrawLine RGB(158, 176, 186), 1, 1, ScaleWidth - 1, 1
            Case elsDisabled
                FillButtonEx m_tButtonColors.GrayColor
        End Select
    End With
    
    With m_tButtonSettings
        If (.State = elsDisabled) Then ' Disabled state
            DrawIcon m_tButtonProperty.PicNormal
            Color1 = RGB(173, 178, 181) 'm_tButtonColors.GrayText
            DrawCaption m_tButtonColors.GrayText
        Else
            DrawIconEffect Color2
            DrawCaptionEffect m_tButtonColors.ForeColor, Color2
            
            ' If button is held down and cursor is moved out of the control,
            ' the Hot state border is displayed until mouse button is released
            ' or cursor is moved back to the button
            
            If (.State = elsHot) Or (m_bMouseIsDown And Not m_bMouseOnButton And Not DrawState = elsDown) Then
                Color1 = RGB(60, 127, 177)
                RectangleEx vbWhite, 1, 1, .Width - 1, .Height - 1
            Else
                If (m_bParentActive) And (Not .State = elsDown And (m_bButtonHasFocus Or .Default)) Then
                    Color1 = RGB(60, 127, 177)
                    RectangleEx RGB(72, 216, 251), 1, 1, .Width - 1, .Height - 1
                ElseIf .State = elsDown Then
                    Color1 = RGB(44, 98, 139)
                Else
                    Color1 = RGB(112, 112, 112)
                    RectangleEx vbWhite, 1, 1, .Width - 1, .Height - 1
                End If
            End If
        End If
        
        ' Draw main button border
        RectangleEx Color1, 0, 0, .Width, .Height
        
        If Not (m_tButtonProperty.Shape = elsCutLeft Or _
                m_tButtonProperty.Shape = elsCutSides) Then
            SetPixelV hDC, 1, 1, Color1
            SetPixelV hDC, 1, .Height - 2, Color1
        End If
        
        If Not (m_tButtonProperty.Shape = elsCutRight Or _
                m_tButtonProperty.Shape = elsCutSides) Then
            SetPixelV hDC, .Width - 2, 1, Color1
            SetPixelV hDC, .Width - 2, .Height - 2, Color1
        End If
    End With
    
End Sub

Private Sub DrawCaption(Color As Long, Optional MoveX As Long, Optional MoveY As Long)
'   Draw the button's caption
    
    Dim tx As String                            '
    Dim tn As Long                              '
        tx = m_tButtonProperty.Caption          ' Get caption and its length
        tn = Len(tx)                            '
                                                '
    If (tn = 0) Then Exit Sub                   ' Bail procedure if no
                                                ' defined caption
    Dim rc As RECT                              '
        CopyRect rc, m_tButtonSettings.Caption  ' Get drawing area
                                                '
    If (Not MoveX = 0) Or (Not MoveY = 0) Then  '
        OffsetRect rc, MoveX, MoveY             ' Move drawing area for how many pixels
    End If                                      ' from the original area if set
                                                '
    SetTextColor hDC, Color                     ' Set text color
                                                '
    If (m_bIsPlatformNT) Then                   '
        DrawTextW hDC, StrPtr(tx), tn, rc, DT_DRAWFLAG
    Else                                        '
        DrawText hDC, tx, tn, rc, DT_DRAWFLAG   ' Draw caption as ansi/unicode
    End If                                      '
    
End Sub

Private Sub DrawCaptionEffect( _
        TextColor As Long, _
        BackColor As Long, _
        Optional MoveX As Long, _
        Optional MoveY As Long)
'   Draw the button's caption with special effect defined applied.
    DrawCaption TextColor, MoveX, MoveY
    
End Sub

Private Sub DrawGradientEx( _
        StartColor As Long, _
        EndColor As Long, _
        Optional Center As Single = 50, _
        Optional ByVal X1 As Long, _
        Optional ByVal Y1 As Long, _
        Optional ByVal X2 As Long = -1, _
        Optional ByVal Y2 As Long = -1) ' Center in percent
'   Draw vertical gradient effect on the control on specified coordinates.
    
    If (X2 = -1) Then X2 = m_tButtonSettings.Width - 1
    If (Y2 = -1) Then Y2 = m_tButtonSettings.Height - 1
    
    X2 = TranslateNumber(X2, X1, X2)            ' X2 must not be < than X1
    Y2 = TranslateNumber(Y2, Y1, Y2)            ' Y2 must not be < than Y1
                                                '
    Dim Color As Long                           '
    Dim Step As Single                          '
                                                '
    Dim RGB1 As tRGB                            ' Start color
    Dim RGB2 As tRGB                            ' End color
    Dim RGB3 As tRGB                            ' Mid color
    Dim RGB4 As tRGB                            ' Gradient color
                                                '
    Center = TranslateNumber(Center, 0, 100)    ' Center should not exceed 0 to 100
                                                '
    RGB1 = GetRGB(StartColor)                   ' Get RGB color values
    RGB2 = GetRGB(EndColor)                     '
                                                '
    RGB3.R = RGB1.R + (RGB2.R - RGB1.R) * 0.5   ' Blend start and end color
    RGB3.G = RGB1.G + (RGB2.G - RGB1.G) * 0.5   '
    RGB3.b = RGB1.b + (RGB2.b - RGB1.b) * 0.5   '
                                                '
    Center = (Y2 - Y1 - 1) * Center / 100       ' Converts center in percent
    Center = (Y1 + Center)                      ' to actual pixel coordinate
                                                '
    If (Center = 0) Then Center = 1             ' Avoid errors
                                                '
    While (Y1 <= Y2)                            ' Draw from top to bottom
        If (Y1 <= Center) Then                  '
            Step = Y1 / Center                  '
            RGB4.R = RGB1.R + (RGB3.R - RGB1.R) * Step
            RGB4.G = RGB1.G + (RGB3.G - RGB1.G) * Step
            RGB4.b = RGB1.b + (RGB3.b - RGB1.b) * Step
        Else                                    '
            Step = (Y1 - Center) / (Y2 - Center)
            RGB4.R = RGB3.R + (RGB2.R - RGB3.R) * Step
            RGB4.G = RGB3.G + (RGB2.G - RGB3.G) * Step
            RGB4.b = RGB3.b + (RGB2.b - RGB3.b) * Step
        End If                                  '
                                                '
        Color = RGBEx(RGB4.R, RGB4.G, RGB4.b)   ' Prepare color
                                                '
        If (X1 = X2) Then                       '
            SetPixelV hDC, X1, Y1, Color        ' Draw point to device
        Else                                    '
            DrawLine Color, X1, Y1, X2, Y1      ' Draw line to device
        End If                                  '
                                                '
        Y1 = Y1 + 1                             '
    Wend                                        '
    
End Sub

Private Sub DrawIcon( _
        Picture As StdPicture, _
        Optional MoveX As Long, _
        Optional MoveY As Long, _
        Optional BrushColor As Long = -1)
'   Draw icon to button with the specified settings.
    Dim hBMBck As Long
    Dim hDCBck As Long
    Dim hOBBck As Long
    
    Dim hBMDst As Long
    Dim hDCDst As Long
    Dim hOBDst As Long
    
    Dim hBMPic As Long
    Dim hDCPic As Long
    Dim hOBPic As Long
    
    Dim hBMSrc As Long
    Dim hdcSrc As Long
    Dim hOBSrc As Long
    
    Dim brRGB As tRGB
    Dim drawH As Long
    Dim drawW As Long
    Dim hBrsh As Long
    Dim lMask As Long
    Dim pictH As Long
    Dim pictW As Long
    Dim tBtmp As BITMAPINFO ' Drawing information for the picture/icon
    Dim tCrop As POINTAPI   ' Points where to start copying/cropping the image
    Dim tPict As RECT       ' Area where to draw the picture
    
    Dim aPict() As RGBQUAD                      ' Picture color bits
    Dim aBack() As RGBQUAD                      ' Background color bits
                                                '
    CopyRect tPict, m_tButtonSettings.Picture   ' Get drawing area
                                                '
    If (Not MoveX = 0) Or (Not MoveY = 0) Then  '
        OffsetRect tPict, MoveX, MoveY          ' Move if set
    End If                                      '
    
    If (m_tButtonProperty.PicSize = elsNormal) Then
        ' Crop drawing area not visible in the button
        If (tPict.Left < 0) Then tCrop.X = -tPict.Left: tPict.Left = 0
        If (tPict.Top < 0) Then tCrop.Y = -tPict.Top: tPict.Top = 0
        If (tPict.Bottom > m_tButtonSettings.Height) Then tPict.Bottom = m_tButtonSettings.Height
        If (tPict.Right > m_tButtonSettings.Width) Then tPict.Right = m_tButtonSettings.Width
    End If
    
    drawH = tPict.Bottom - tPict.Top            ' Draw height
    drawW = tPict.Right - tPict.Left            ' Draw width
                                                '
    If (drawH < 1) Or (drawW < 1) Then Exit Sub ' Nowhere to draw
                                                '
    pictH = ScaleY(Picture.Height, 8, 3)        ' Picture height
    pictW = ScaleX(Picture.Width, 8, 3)         ' Picture width
                                                '
    hdcSrc = CreateCompatibleDC(hDC)            ' Create drawing DC
                                                '
    If (Picture.Type = 1) Or (Picture.Type > 1 And Not m_tButtonProperty.UseMask) Then
        hOBSrc = SelectObject(hdcSrc, Picture.Handle)
    End If                                      '
                                                '
    If (m_tButtonProperty.UseMask) Then         ' Check if we can use the maskcolor set
        lMask = m_tButtonColors.MaskColor       '
    ElseIf (Picture.Type > 1) Then              ' if not then check if we have an icon
        lMask = GetPixel(hdcSrc, 0, 0)          ' if it is then get top-left pixel color
        DeleteObject SelectObject(hdcSrc, hOBSrc)
    Else                                        '
        lMask = -1                              ' if it is a bitmap then use no maskcolor
    End If                                      '
    
    If (Picture.Type > 1) Then
        hBMSrc = CreateCompatibleBitmap(hDC, pictW, pictH)
        hOBSrc = SelectObject(hdcSrc, hBMSrc)
        hBrsh = CreateSolidBrush(lMask)
        
        ' Fill transparent areas of the icon with the defined maskcolor
        DrawIconEx hdcSrc, 0, 0, Picture.Handle, pictW, pictH, 0, hBrsh, DI_NORMAL
        DeleteObject hBrsh
    End If
    
    hDCBck = CreateCompatibleDC(hdcSrc)
    hDCDst = CreateCompatibleDC(hdcSrc)
    hDCPic = CreateCompatibleDC(hdcSrc)
    
    hBMBck = CreateCompatibleBitmap(hDC, drawW, drawH)
    hBMDst = CreateCompatibleBitmap(hDC, drawW, drawH)
    hBMPic = CreateCompatibleBitmap(hDC, drawW, drawH)
    
    hOBBck = SelectObject(hDCBck, hBMBck)
    hOBDst = SelectObject(hDCDst, hBMDst)
    hOBPic = SelectObject(hDCPic, hBMPic)
    
    If (tCrop.X > 0 Or tPict.Right = m_tButtonSettings.Width) Then pictW = drawW + tCrop.X
    If (tCrop.Y > 0 Or tPict.Bottom = m_tButtonSettings.Height) Then pictH = drawH + tCrop.Y
    
    ' Copy image to destination DC. Crop/resize if necessary.
    StretchBlt hDCDst, 0, 0, drawW, drawH, hdcSrc, tCrop.X, tCrop.Y, pictW - tCrop.X, pictH - tCrop.Y, SRCCOPY
    
    If (Not Picture.Type = 1) Then ' vbPicTypeBitmap
        DeleteObject SelectObject(hdcSrc, hOBSrc)
    End If
    
    DeleteDC hdcSrc
    
    ReDim aBack(0 To drawW * drawH * 1.5) As RGBQUAD
    ReDim aPict(0 To UBound(aBack)) As RGBQUAD
    
    ' Get background & picture bitmap image
    BitBlt hDCBck, 0, 0, drawW, drawH, hDC, tPict.Left, tPict.Top, SRCCOPY
    BitBlt hDCPic, 0, 0, drawW, drawH, hDCDst, 0, 0, SRCCOPY
    
    With tBtmp.bmiHeader
        .biBitCount = 24 ' bit
        .biCompression = BI_RGB ' = 0
        .biHeight = drawH
        .biPlanes = 1
        .biSize = Len(tBtmp.bmiHeader)
        .biWidth = drawW
    End With
    
    ' Get background & picture color bits
    GetDIBits hDCBck, hBMBck, 0, drawH, aBack(0), tBtmp, DIB_RGB_COLORS
    GetDIBits hDCPic, hBMPic, 0, drawH, aPict(0), tBtmp, DIB_RGB_COLORS
    
    DeleteObject SelectObject(hDCBck, hOBBck)       ' Clear bitmap objects from memory
    DeleteObject SelectObject(hDCPic, hOBPic)       ' immediately after being used
                                                    '
    DeleteDC hDCBck                                 ' Clear device context instances
    DeleteDC hDCPic                                 ' from memory immediately
                                                    '
    If (BrushColor > -1) Then                       ' Determine brush color
        brRGB = GetRGB(BrushColor)                  ' used to replace colors on the
    End If                                          ' image
                                                    '
    Dim lOpacity As Long                            '
    If (m_tButtonSettings.State = elsDisabled) Then ' For disabled buttons
        lOpacity = m_tButtonProperty.PicOpacity     ' We will just blend the picture
        m_tButtonProperty.PicOpacity = 0.2          ' On the button by 20%
    End If                                          '
    
    If (lMask = -1) And (BrushColor = -1) And (m_tButtonProperty.PicOpacity = 1) Then
        ' Skip bit by bit processing of image when not really necessary.
        ' Helps make things faster especially when loading large image/s
        GoTo Jmp_DrawImage
    End If
    
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    
    While (Y < drawH)
        X = 0
        While (X < drawW)
            ' GetNearestColor returns the actual value identifying a color from the
            ' system palette that will be displayed when the specified color is used
            
            If (GetNearestColor(hDCDst, RGBEx(aPict(Z).rgbRed, _
                                              aPict(Z).rgbGreen, _
                                              aPict(Z).rgbBlue)) = lMask) Then
                                                '
                aPict(Z) = aBack(Z)             ' Replace to background pixel color
                                                ' to make it look like transparent
            Else                                '
                If (BrushColor > -1) Then       '
                    aPict(Z).rgbRed = brRGB.R   ' Change all pixel color values
                    aPict(Z).rgbGreen = brRGB.G ' of an image with the specified
                    aPict(Z).rgbBlue = brRGB.b  ' brush color when set
                    
                ElseIf (Not m_tButtonProperty.PicOpacity = 1) Then
                    ' Results to an effect that blends the picture on the control
                    aPict(Z) = BlendRGBQUAD(aBack(Z), aPict(Z), m_tButtonProperty.PicOpacity)
                    
                End If
            End If
            
            X = X + 1
            Z = Z + 1 ' bit counter
        Wend
        
        Y = Y + 1
    Wend
    
Jmp_DrawImage:
    
    Erase aBack
    DeleteObject SelectObject(hDCDst, hOBDst)
    DeleteDC hDCDst
    
    SetDIBitsToDevice hDC, _
                      tPict.Left, _
                      tPict.Top, _
                      drawW, _
                      drawH, _
                      0, _
                      0, _
                      0, _
                      drawH, _
                      aPict(0), _
                      tBtmp, _
                      DIB_RGB_COLORS                ' Draw optimized image to button
    Erase aPict                                     '
                                                    ' Clear color bit arrays from memory
    If (m_tButtonSettings.State = elsDisabled) Then '
        m_tButtonProperty.PicOpacity = lOpacity     ' Restore opacity
    End If                                          '
    
End Sub

Private Sub DrawIconEffect( _
        BackColor As Long, _
        Optional MoveX As Long, _
        Optional MoveY As Long, _
        Optional BrushColor As Long = -1)
'   Draw icon to button with defined special effect being applied.
    If (m_tButtonProperty.PicNormal Is Nothing) Then
        Exit Sub
    End If
    
    Dim Picture As StdPicture
    
    If (m_tButtonSettings.State = elsHot And Not m_tButtonProperty.PicHot Is Nothing) Then
        Set Picture = m_tButtonProperty.PicHot
    ElseIf (m_tButtonSettings.State = elsDown) Or (m_tButtonProperty.Value) Then
        If (m_tButtonProperty.PicDown Is Nothing) Then
            Set Picture = m_tButtonProperty.PicHot
        Else
            Set Picture = m_tButtonProperty.PicDown
        End If
    End If
    
    If (Picture Is Nothing) Then
        Set Picture = m_tButtonProperty.PicNormal
    End If
    
    DrawIcon Picture, MoveX, MoveY, BrushColor
    
End Sub

Private Sub DrawLine(Color As Long, X1 As Long, Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
'   Draw a line with the specified color and coordinates
    Dim hOld As Long
    Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 1, Color)
        hOld = SelectObject(hDC, hPen)
    
    If (X1 = X2) Then Y2 = Y2 + 1   ' LineTo draws a line up to, but not
    If (Y1 = Y2) Then X2 = X2 + 1   ' including the defined point. But not now!
                                    '
    ' If (X1 = X2) Then             '
    '     If (Y2 >= Y1) Then Y2 = Y2 + 1 Else Y2 = Y2 - 1
    ' End If                        '
    ' If (Y1 = Y2) Then             '
    '     If (X2 >= X1) Then X2 = X2 + 1 Else X2 = X2 - 1
    ' End If                        '
                                    '
    MoveToEx hDC, X1, Y1, 0&        ' Set starting position of line
    LineTo hDC, X2, Y2              ' then draw a line to this point
                                    '
    SelectObject hDC, hOld          ' Restore the previous pen
    DeleteObject hPen               ' then clear pen object from memory
    
End Sub

Private Sub FillButtonEx(Color As Long)
'   Fill the control with the specified color.
    Dim hOldBr As Long
    Dim hBrush As Long
    Dim lpRect As RECT
    
    hBrush = CreateSolidBrush(Color)    ' Create new brush with the specified color
    hOldBr = SelectObject(hDC, hBrush)  ' Use the new brush but save previous one
                                        '
    SetRect lpRect, 0, 0, m_tButtonSettings.Width, m_tButtonSettings.Height
                                        '
    FillRect hDC, lpRect, hBrush        '
                                        '
    SelectObject hDC, hOldBr            ' Restore the previous brush
    DeleteObject hBrush                 ' Remove brush instance from memory
    
End Sub

Private Function GetAccessKey(Caption As String) As String
'   Get accesskey from a caption
    Dim iMax As Integer
    Dim iPos As Integer
    Dim sChr As String * 1 ' Helps conserve memory I guess
    
    iMax = Len(Caption)
    
    ' Of course you can't assign an accesskey
    ' with less than 2 characters
    If (iMax < 2) Then Exit Function
    
    iMax = iMax - 1 ' Start from the second to the last character
    
    Do  ' An accesskey is found after the last
        ' ampersand character found on the string
        iPos = InStrRev(Caption, "&", iMax)
        
        If (iPos = 0) Then Exit Do
        
        If (iPos = 1) Then
            GetAccessKey = Mid$(Caption, iPos + 1, 1)
            Exit Do
        Else
            ' Check if the character before the
            ' ampersand is also an ampersand
            sChr = Mid$(Caption, iPos - 1, 1)
            
            ' A series of two ampersand characters will draw
            ' an ampersand character as part of the caption
            ' and will not be considered as an accesskey
            If (Not StrComp(sChr, "&") = 0) Then
                GetAccessKey = Mid$(Caption, iPos + 1, 1)
                Exit Do
            End If
            
            iMax = iPos - 2 ' Find another character
        End If
        
    Loop While (iMax > 0)
    
    GetAccessKey = LCase$(GetAccessKey)
    
End Function

Private Function GetRGB(Color As Long) As tRGB
'   Returns the RGB color value of the specified color.
    GetRGB.R = Color And 255
    GetRGB.G = (Color \ 256) And 255
    GetRGB.b = (Color \ 65536) And 255
    
End Function

Private Function IsFunctionSupported(sFunction As String, sModule As String) As Boolean
'   Determines if the passed function is supported by a library
    Dim hModule As Long
        hModule = GetModuleHandleA(sModule) ' GetModuleHandle
        
    If (hModule = 0) Then
        hModule = LoadLibrary(sModule)
    End If
    
    If (hModule) Then
        If (GetProcAddress(hModule, sFunction)) Then
            IsFunctionSupported = True
        End If
        FreeLibrary hModule
    End If
    
End Function

Private Function IsPlatformNT() As Boolean
'   Determines if the system currently running the program has an NT platform.
    Dim OSINFO As OSVERSIONINFO
        OSINFO.dwOSVersionInfoSize = Len(OSINFO)
        
    If (GetVersionEx(OSINFO)) Then
        IsPlatformNT = (OSINFO.dwPlatformId = VER_PLATFORM_WIN32_NT)
    End If
    
End Function

Private Sub RectangleEx(Color As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
'   Draws a rectangle using the specified color and coordinates
    Dim hOld As Long
    Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 1, Color)    ' Create new pen
        hOld = SelectObject(hDC, hPen)          ' Use new pen
                                                '
        Rectangle hDC, X1, Y1, X2, Y2           ' Draw shape
                                                '
        SelectObject hDC, hOld                  ' Restore previous pen
        DeleteObject hPen                       ' Remove new pen from memory
    
End Sub

Private Function RGBEx( _
        ByVal Red As Long, _
        ByVal Green As Long, _
        ByVal Blue As Long) As Long
'   Returns a whole number representing an RGB color value.
    RGBEx = Red + 256& * Green + 65536 * Blue
    
End Function

Private Sub SetButtonColors()
'   Set button colors and translate to vb safe colors
    
    With m_tButtonColors ' Common to all styles
        .BackColor = TranslateColor(m_tButtonProperty.BackColor)
        .DownColor = RGB(196, 229, 246) 'ShiftColor(.BackColor, -0.05)
        .FocusBorder = RGB(72, 215, 251) '&HE4AD89
        .ForeColor = TranslateColor(m_tButtonProperty.ForeColor)
        .GrayColor = BlendColors(.BackColor, &HFFFFFF, 0.8)
        .GrayText = BlendColors(.ForeColor, &HFFFFFF, 0.6)
        .HoverColor = RGB(217, 240, 252) '&H30B3F8
        .MaskColor = TranslateColor(m_tButtonProperty.MaskColor)
        .StartColor = RGB(242, 242, 242) '&HFFFFFF ' vbWhite
    End With
    
End Sub

Private Function ShiftColor(Color As Long, PercentInDecimal As Single) As Long
'   Add or remove a certain color quantity by how many percent.
    Dim RGB1 As tRGB
        RGB1 = GetRGB(Color)
        
        RGB1.R = RGB1.R + PercentInDecimal * 255 ' Percent should already
        RGB1.G = RGB1.G + PercentInDecimal * 255 ' be translated.
        RGB1.b = RGB1.b + PercentInDecimal * 255 ' Ex. 50% -> 50 / 100 = 0.5
        
    If (PercentInDecimal > 0) Then ' RGB values must be between 0-255 only
        If (RGB1.R > 255) Then RGB1.R = 255
        If (RGB1.G > 255) Then RGB1.G = 255
        If (RGB1.b > 255) Then RGB1.b = 255
    Else
        If (RGB1.R < 0) Then RGB1.R = 0
        If (RGB1.G < 0) Then RGB1.G = 0
        If (RGB1.b < 0) Then RGB1.b = 0
    End If
    
    ShiftColor = RGBEx(RGB1.R, RGB1.G, RGB1.b) ' Return shifted color value
    
End Function

Private Sub TrackMouseTracking(hWnd As Long)
'   Start tracking of mouse leave event
    Dim lpEventTrack As TRACKMOUSEEVENTTYPE
    
    With lpEventTrack
        .cbSize = Len(lpEventTrack)
        .dwFlags = TME_LEAVE
        .hwndTrack = hWnd
    End With
    
    If (m_bTrackHandler32) Then
        TrackMouseEvent lpEventTrack
    Else
        TrackMouseEvent2 lpEventTrack
    End If
    
End Sub

Private Function TranslateNumber( _
        ByVal Value As Single, _
        Minimum As Long, _
        Maximum As Long) As Single
'   Ensure a number does not exceed the specified limits.
    
    If (Value > Maximum) Then
        TranslateNumber = Maximum
    ElseIf (Value < Minimum) Then
        TranslateNumber = Minimum
    Else
        TranslateNumber = Value
    End If
    
End Function

' //-- Subclassing Procedure --//

Private Sub Subclass_Proc( _
        ByVal bBefore As Boolean, _
        ByRef bHandled As Boolean, _
        ByRef lReturn As Long, _
        ByVal lng_hWnd As Long, _
        ByVal uMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long, _
        ByRef lParamUser As Long)
        
    Select Case uMsg
        Case WM_MOUSELEAVE
            ' Triggered as the cursor moved out the button. If mouse button is held
            ' down, is it triggered when the button is released outside the button
            
            m_bMouseOnButton = False
            m_bIsTracking = False
            
            If (Not m_bSpacebarIsDown) Then
                If (m_tButtonProperty.Enabled) Then
                    DrawButton elsNormal, (m_tButtonSettings.Button = 1)
                End If
                RaiseEvent MouseLeave
            End If
            
        Case WM_ACTIVATE, WM_NCACTIVATE ' Parent form is activated or deactivated
            m_bParentActive = (Not wParam = 0)
            
            If (m_bParentActive) Then  ' Activated
                If (m_tButtonProperty.Enabled) Then
                    ' Force update if set to default or was on focus
                    If (m_bButtonHasFocus Or m_tButtonSettings.Default) Then
                        DrawButton elsNormal, True
                    End If
                End If
            Else ' Deactivated
                Dim bFocus As Boolean
                Dim bForce As Boolean
                
                bFocus = m_bButtonHasFocus
                bForce = m_bButtonHasFocus Or m_tButtonSettings.Default Or m_bMouseOnButton
                
                m_bButtonHasFocus = False   ' Unset runtime settings
                m_bButtonIsDown = False     ' necessary to effectively
                m_bMouseIsDown = False      ' draw a normal button
                m_bMouseOnButton = False    '
                m_bSpacebarIsDown = False   '
                
                m_tButtonSettings.Default = False ' Temporary cancel DisplayAsDefault
                
                If (m_tButtonProperty.Enabled) Then
                    DrawButton elsNormal, bForce
                End If
                
                ' Restore neccesary settings used when parent form is reactivated again
                m_tButtonSettings.Default = Ambient.DisplayAsDefault
                m_bButtonHasFocus = bFocus
            End If
    End Select
    
End Sub

' //-- UserControl Procedures --//

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
'   Triggered if the accesskey is pressed (Alt + underlined character of caption)
'   Also on ENTER key or ESCAPE if Cancel property is set to True
    
    If (m_tButtonProperty.Enabled) Then
        If (m_bSpacebarIsDown) Then
            If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                ReleaseCapture                      ' input processing
            End If                                  ' of the window
        End If
        
        If (m_tButtonProperty.CheckBox) Then
            ' Checkboxes does not respond to Enter & Escape keys
            If (KeyAscii = 13) Or (KeyAscii = 27) Then ' vbKeyReturn, vbKeyEscape
                Exit Sub
            End If
        End If
        
        m_bButtonIsDown = False         ' Release button
        m_tButtonSettings.Button = 1    '
                                        '
        UserControl_Click               ' Trigger click event
    End If
    
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'   Usually triggered as the user changes focus on different controls on the window
    
    ' DisplayAsDefault returns True if the current control
    ' on focus is not another button or this button itself.
    
    m_tButtonSettings.Default = Ambient.DisplayAsDefault
    
    If (StrComp(PropertyName, "DisplayAsDefault") = 0) Then
        m_bButtonIsDown = False
        m_bMouseIsDown = False
        m_bSpacebarIsDown = False
        
        ' Prevent unneccessary drawing updates
        If (m_tButtonProperty.Enabled) And (Not m_bMouseOnButton) Then
            ' GotFocus event will just update the button display later
            If (Not m_bButtonHasFocus) Then
                DrawButton Force:=True
            End If
        End If
    End If
    
End Sub

Private Sub UserControl_Click()
'   Triggered normally when user clicks the control and release it inside the button
    If (m_bButtonIsDown) Or (Not m_tButtonSettings.Button = 1) Then Exit Sub
        m_bMouseIsDown = False
        m_bSpacebarIsDown = False
    
    ' Note: For a normal(unchecked) state of a checkbox mode, VALUE is FALSE
    '       While on a down(checked) state, VALUE returns TRUE.
    '
    '       In a command button mode, VALUE is always FALSE but not
    '       before the click event in which it should return TRUE
    
    With m_tButtonProperty                  '
        If (.CheckBox) Then                 ' Check if checkbox mode is on
            .Value = Not .Value             ' If so, then toggle button value
        End If                              '
                                            '
        If (Not m_bMouseOnButton) Then      ' Check if cursor is over the control
            DrawButton elsNormal, True      ' Redraw is necessary if it is not :)
        Else                                '
            DrawButton elsHot, .CheckBox    ' Force redraw for checkbox mode
        End If                              '
                                            '
        RaiseEvent Click                    '
                                            '
        If (Not .CheckBox) Then             ' Sometimes VALUE property is set to True
            .Value = False                  ' to trigger the click event
        End If                              ' So we should reset value if set
    End With                                '
    
End Sub

Private Sub UserControl_DblClick()
    If (m_tButtonProperty.HandPointer) Then
        SetCursor m_tButtonSettings.Cursor ' Set hand cursor
    End If
    
    If (m_tButtonSettings.Button = 1) Then ' vbLeftButton
        ' Draw a down button state which helps so emulate multiple clicks
        m_bButtonIsDown = True          ' Bug fixed: 07/27/06
        m_bMouseIsDown = True           ' Draw HOT state on dblclick, hold then mousemove
        DrawButton elsDown
        m_tButtonSettings.Button = 8    ' Just a double click flag for the MouseUp event
                                        '
        If (Not GetCapture = UserControl.hWnd) Then
            SetCapture UserControl.hWnd ' Send MouseUp event to the control
        End If                          '
                                        '
        RaiseEvent DblClick             '
    End If                              '
    
End Sub

Private Sub UserControl_GotFocus()
'   Only raised when last focused control is on the same window but not itself
    m_bButtonHasFocus = True
    
    If (Not m_bButtonIsDown) Then
        DrawButton elsNormal
    End If
    
End Sub

Private Sub UserControl_Hide()
    m_bControlHidden = True
    
End Sub

Private Sub UserControl_Initialize()
'   Called on design and run-time; when the form is getting ready for display
    
    m_bIsPlatformNT = IsPlatformNT() ' Needed for the unicode text support
    m_bRedrawOnResize = False
    
End Sub

Private Sub UserControl_InitProperties()
'   Called on design time only; everytime this control is added on the form
    
    With m_tButtonProperty
        .BackColor = &HE6E6E6   'RGB(230, 230, 230)
        .Caption = Ambient.DisplayName
        .CheckBox = False
        .Enabled = True
         UserControl.Font = Ambient.Font
        .ForeColor = Ambient.ForeColor
        .MaskColor = &HC0C0C0
        .PicAlign = elaLeftOfCaption
    Set .PicDown = Nothing
    Set .PicHot = Nothing
    Set .PicNormal = Nothing
        .PicOpacity = 1
        .PicSize = 0 ' elpNormal
        .UseMask = True
        .Value = False
    End With
    
    SetButtonColors
    
    m_bRedrawOnResize = True
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 32 ' vbKeySpace
            ' Some buttons seems to forget this thing, and it sucks me off!
            ' I find Alt+Space to be so useful especially in closing forms.
            If (Not Shift = 4) Then ' vbAltMask
                m_bButtonIsDown = True
                m_bSpacebarIsDown = True
                
                If (Not GetCapture = UserControl.hWnd) Then
                    SetCapture UserControl.hWnd ' Restrict user from selecting other
                End If                          ' controls while spacebar is held down
                                                '
                If (Not m_bMouseIsDown) Then    ' If mouse is up only
                    DrawButton elsDown
                End If
            End If
            
            ' Normally, this event is raised before drawing the
            ' pressed button state when spacebar is pressed.
            ' I moved it here to draw the down state of the button
            ' before the user can add some additional commands.
            RaiseEvent KeyDown(KeyCode, Shift)
            
        Case 37, 38, 39, 40 ' vbKeyLeft, vbKeyUp, vbKeyRight, vbKeyDown
            If (Shift = 0) Then
                If (KeyCode = 37) Or (KeyCode = 38) Then ' vbKeyLeft, vbKeyUp
                    SendKeys "+{TAB}"
                Else
                    SendKeys "{TAB}"
                End If
            End If
            
            If (m_bSpacebarIsDown) Then
                If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                    ReleaseCapture                      ' input processing
                End If                                  ' of the window
                
                ' If SPACEBAR is down then either of the arrow keys is pressed
                ' Process the arrow event first to transfer focus to the next
                ' available control then trigger the click event after
                DoEvents
                m_tButtonSettings.Button = 1
                UserControl_Click
            End If
            
        Case Else
            ' If spacebar is held down, then a key not included above is pressed
            ' should simulate a release to the button to an appropriate state.
            If (m_bSpacebarIsDown) Then
                m_bButtonIsDown = False
                m_bSpacebarIsDown = False
                
                If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                    ReleaseCapture                      ' input processing
                End If                                  ' of the window
                
                If (m_bMouseOnButton) Then
                    DrawButton elsHot
                Else
                    DrawButton elsNormal
                End If
            End If
            
            RaiseEvent KeyDown(KeyCode, Shift)
    End Select
    
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    
    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 32 ' vbKeySpace
            m_bButtonIsDown = m_bMouseIsDown
            
            If (m_bSpacebarIsDown) Then                 '
                m_bSpacebarIsDown = False               ' Space has been released
                                                        '
                m_tButtonSettings.Button = 1            '
                UserControl_Click                       '
            End If                                      '
                                                        '
            If (m_bButtonIsDown) Then                   ' Raise Mouse_Up event
                If (Not GetCapture = UserControl.hWnd) Then
                    SetCapture UserControl.hWnd         ' When the left-mouse button
                End If                                  ' will be finally released
            Else                                        '
                If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
                    ReleaseCapture                      ' input processing
                End If                                  ' of the window
            End If                                      '
                                                        '
            RaiseEvent KeyUp(KeyCode, Shift)            '
        Case Else
            RaiseEvent KeyUp(KeyCode, Shift)
    End Select
    
End Sub

Private Sub UserControl_LostFocus()
'   Only raised when another control of the same window is focused
    
    m_bButtonHasFocus = False           '
    m_bButtonIsDown = False             '
    m_bMouseIsDown = False              ' Release buttons
    m_bSpacebarIsDown = False           '
                                        '
    If (m_tButtonProperty.Enabled) Then '
        If (m_bParentActive) Then       ' We need to force redraw
            DrawButton elsNormal, True  ' button on lost of control focus
        End If                          ' only when parent window is active
    End If                              ' the other way is handled by the subclass proc
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (m_tButtonProperty.HandPointer) Then '
        SetCursor m_tButtonSettings.Cursor  ' Set hand cursor
    End If                                  '
                                            '
    m_tButtonSettings.Button = Button       ' Cache button to trigger double click event
                                            '
    If (Button = 1) Then                    ' vbLeftButton
        m_bButtonHasFocus = True            '
        m_bButtonIsDown = True              '
        m_bMouseIsDown = True               '
                                            '
        If (Not m_bSpacebarIsDown) Then     '
            DrawButton elsDown              '
        ElseIf (Not m_bMouseOnButton) Then  ' If mouse button is pressed outside
            m_bMouseIsDown = False          ' the control while the spacebar is
            DrawButton elsNormal            ' being held down then draw hot state
            m_bMouseIsDown = True           ' Unset/set m_bMouseIsDown to trick
        End If                              ' drawing procedures about the state
    End If                                  '
                                            '
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Parent form is inactive
    If (m_tButtonProperty.HandPointer) Then
        SetCursor m_tButtonSettings.Cursor ' Set hand cursor
    Else ' Ensure cursor is what we expected
        UserControl.MousePointer = UserControl.MousePointer
    End If
    
    ' Restrict HOT state if parent window/form is not in focus
    ' If (Not m_bParentActive) Then Exit Sub
    ' Not now, I allow MouseOver event to show HOT state
    ' even if parent window is not in focus
    
    Dim lpPoint As POINTAPI
        GetCursorPos lpPoint
        
    If (Not WindowFromPoint(lpPoint.X, lpPoint.Y) = UserControl.hWnd) Then
        ' It reaches here if the left mouse button is held down
        ' as the user moves the cursor off the control
        If (m_bMouseOnButton) Then
            m_bMouseOnButton = False
            
            If (Not m_bSpacebarIsDown And Not m_bMouseIsDown) Then ' Retain down state
                DrawButton elsNormal
            End If
            
            RaiseEvent MouseLeave
        ElseIf (m_bMouseIsDown) Then
            DrawButton elsHot
        End If
    Else
        m_bMouseOnButton = True
        
        If (Not m_bSpacebarIsDown) Then ' Check if spacebar is held down
            If (m_bButtonIsDown) Then   ' If not, draw appropriate button state
                DrawButton elsDown      '
            Else                        ' If button should be in down state then
                DrawButton elsHot       ' draw the down state else draw the hot state
            End If                      '
        ElseIf (m_bMouseIsDown) Then    ' Else, check if mouse is held down
            DrawButton elsDown          ' as it moves over the button to
        End If                          ' draw the down state
                                        '
        If (Not m_bIsTracking) Then     ' Trigger MouseEnter event the first time
                m_bIsTracking = True    ' the cursor has moved on the control
            
            TrackMouseTracking UserControl.hWnd
            RaiseEvent MouseEnter
        Else
            ' Succeeding move events will trigger the MouseMove event
            RaiseEvent MouseMove(Button, Shift, X, Y)
        End If
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (m_tButtonProperty.HandPointer) And (m_bMouseOnButton) Then
        SetCursor m_tButtonSettings.Cursor ' Set hand cursor
    End If
    
    If (Button = 1) Then ' vbLeftButton
        m_bMouseIsDown = False
        m_bButtonIsDown = m_bSpacebarIsDown
        
        If (m_bSpacebarIsDown) And (Not m_bMouseOnButton) Then
            DrawButton elsDown
        ElseIf (m_tButtonSettings.Button = 8) Then
            ' 8 -> Control had been double clicked
            If (m_bMouseOnButton) Then
                DrawButton elsHot
            Else
                DrawButton elsNormal
            End If
        End If
        
        If (GetCapture = UserControl.hWnd) Then ' Restore normal mouse
            ReleaseCapture                      ' input processing
        End If                                  ' of the window
                                                '
        RaiseEvent MouseUp(Button, Shift, X, Y) '
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'   Called everytime the properties are needed for display after being initialized
'   (On loading of form and before the is control is unloaded to design mode)
    
    With PropBag
        m_tButtonProperty.BackColor = .ReadProperty("BackColor", &HE6E6E6)
        m_tButtonProperty.Shape = .ReadProperty("ButtonShape", 0)
        m_tButtonProperty.Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_tButtonProperty.CheckBox = .ReadProperty("CheckBox", False)
        m_tButtonProperty.Enabled = .ReadProperty("Enabled", True)
        m_tButtonProperty.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
        m_tButtonProperty.HandPointer = .ReadProperty("HandPointer", False)
        m_tButtonProperty.MaskColor = .ReadProperty("MaskColor", &HC0C0C0)
        m_tButtonProperty.PicAlign = .ReadProperty("PicAlign", elaLeftOfCaption)
    Set m_tButtonProperty.PicDown = .ReadProperty("PicDown", Nothing)
    Set m_tButtonProperty.PicHot = .ReadProperty("PicHot", Nothing)
    Set m_tButtonProperty.PicNormal = .ReadProperty("PicNormal", Nothing)
        m_tButtonProperty.PicOpacity = .ReadProperty("PicOpacity", 1) ' Do not blend
        m_tButtonProperty.PicSize = .ReadProperty("PicSize", 0) ' elpNormal
        m_tButtonProperty.PicSizeH = .ReadProperty("PicSizeH", 0)
        m_tButtonProperty.PicSizeW = .ReadProperty("PicSizeW", 0)
        m_tButtonSettings.State = .ReadProperty("State", 0) ' elsNormal
        m_tButtonProperty.UseMask = .ReadProperty("UseMask", True)
        m_tButtonProperty.Value = .ReadProperty("Value", False)
        
    Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0) ' vbDefault
    End With
    
    UserControl.AccessKeys = GetAccessKey(m_tButtonProperty.Caption) ' Assign accesskey
    UserControl.Enabled = m_tButtonProperty.Enabled
    
    If (Ambient.UserMode) Then ' Start subclassing
        
        If (m_tButtonProperty.HandPointer) Then
            m_tButtonSettings.Cursor = LoadCursor(0, IDC_HAND) ' Load hand pointer
            ' If LoadCursor fails, the system does not support hand pointer option
            m_tButtonProperty.HandPointer = (Not m_tButtonSettings.Cursor = 0)
        End If
        
        m_bTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
        
        If (Not m_bTrackHandler32) Then
            If (Not IsFunctionSupported("_TrackMouseEvent", "Comctl32")) Then
                Err.Raise -1, "System does not support TrackMouseEvent."
                ' ...which is really neccessary for this control to work properly
                GoTo Jmp_Skip
            End If
        End If
        
        sc_Subclass hWnd               ' Subclass the control
        sc_AddMsg hWnd, WM_MOUSELEAVE  ' Detect mouse leave event
        
        Dim hParent As Long
        Dim hWindow As Long
            If (TypeOf Parent Is Form) Then             ' Check if parent is a form
                hParent = Parent.hWnd                   ' Get parent form handle
            ElseIf (TypeOf Parent.Parent Is Form) Then  ' If not check parent of it
                hParent = Parent.Parent.hWnd            ' and so on...
            ElseIf (TypeOf Parent.Parent.Parent Is Form) Then
                hParent = Parent.Parent.Parent.hWnd     ' If here still fails then
            End If                                      ' I quit. We just simply skip
                                                        ' subclass for parent form :)
            If (hParent) Then                           '
                sc_Subclass hParent                     '
                hWindow = GetWindowLong(hParent, GWL_EXSTYLE)
                                                        '
                If (hWindow And WS_EX_MDICHILD) Then    ' Bug fixed:
                    sc_AddMsg hParent, WM_NCACTIVATE    ' Now we check if the
                Else                                    ' parent form is an MDI
                    sc_AddMsg hParent, WM_ACTIVATE      ' child the API way :)
                End If                                  '
                
            End If
    End If
    
Jmp_Skip:
    SetButtonColors
    m_bCalculateRects = True
    DrawButton Force:=True
    
    m_bRedrawOnResize = True
    
End Sub

Private Sub UserControl_Resize()
'   Called everytime the control is resized on design and run time
'   (Also at the first time the control is added and when the form is
'   loaded/unloaded but only after the control has been initialized)
    
    With m_tButtonSettings
        Dim lpRect As RECT
        GetClientRect UserControl.hWnd, lpRect
        
        ' We use the control's hWnd thus returning its own coordinates
        ' lpRect.Left & lpRect.Top always returns zero(0) while
        ' lpRect.Bottom returns the control's height and
        ' lpRect.Right returns the control's width.
        
        .Height = lpRect.Bottom
        .Width = lpRect.Right
        
        Const MIN_HPX As Long = 15 ' Minimum button height (in pixels)
        Const MIN_WPX As Long = 15 ' Minimum button width  (in pixels)
        
        ' Defined minimum button height and width should be enough...
        ' What's the purpose of setting the button'size lower than that?
        ' Nothing? Or just to find bugs/defects??? Oh no you dont! :)
        
        If (.Height < MIN_HPX) Or (.Width < MIN_WPX) Then
            If (.Height < MIN_HPX) Then
                UserControl.Height = MIN_HPX * Screen.TwipsPerPixelY
            End If
            If (.Width < MIN_WPX) Then
                UserControl.Width = MIN_WPX * Screen.TwipsPerPixelX
            End If
            Exit Sub
        End If
        
        m_bCalculateRects = True
        
        If (Ambient.UserMode) Then      ' Always allow to redraw when running mode
            DrawButton Force:=True      '
        ElseIf (m_bRedrawOnResize) Then ' On IDE, some sort of filtering is done
            DrawButton Force:=True      ' to prevent control to redraw twice or more
        End If                          '
    End With
    
End Sub

Private Sub UserControl_Show()
    
    m_bControlHidden = False
    
End Sub

Private Sub UserControl_Terminate()
    
    If (m_tButtonProperty.HandPointer) Then
        DeleteObject m_tButtonSettings.Cursor ' Destry hand pointer handle
    End If
    
    On Error GoTo Jmp_Skip      ' Bug fixed:
                                ' IDE crash when UNLOAD called from its own event
    If (Ambient.UserMode) Then  '
        Call sc_Terminate       ' Stop subclassers
    End If                      '
                                ' Jump through here on occurence of error
Jmp_Skip:                       ' Error 398: Client Site not available...
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'   Called everytime the parent control/form is loaded on memory
'   but only when there are changes in the properties to be updated
    
    With PropBag
        .WriteProperty "BackColor", m_tButtonProperty.BackColor, &HE6E6E6
        .WriteProperty "ButtonShape", m_tButtonProperty.Shape, 0
        .WriteProperty "Caption", m_tButtonProperty.Caption, Ambient.DisplayName
        .WriteProperty "CheckBox", m_tButtonProperty.CheckBox, False
        .WriteProperty "Enabled", m_tButtonProperty.Enabled, True
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "ForeColor", m_tButtonProperty.ForeColor, Ambient.ForeColor
        .WriteProperty "HandPointer", m_tButtonProperty.HandPointer, False
        .WriteProperty "MaskColor", m_tButtonProperty.MaskColor, &HC0C0C0
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, 0 ' vbDefault
        .WriteProperty "PicAlign", m_tButtonProperty.PicAlign, elaLeftOfCaption
        .WriteProperty "PicDown", m_tButtonProperty.PicDown, Nothing
        .WriteProperty "PicHot", m_tButtonProperty.PicHot, Nothing
        .WriteProperty "PicNormal", m_tButtonProperty.PicNormal, Nothing
        .WriteProperty "PicOpacity", m_tButtonProperty.PicOpacity, 1 ' Do not blend
        .WriteProperty "PicSize", m_tButtonProperty.PicSize, 0 ' elpNormal
        .WriteProperty "PicSizeH", m_tButtonProperty.PicSizeH, 0
        .WriteProperty "PicSizeW", m_tButtonProperty.PicSizeW, 0
        .WriteProperty "State", m_tButtonSettings.State, 0 ' elsNormal
        .WriteProperty "UseMask", m_tButtonProperty.UseMask, True
        .WriteProperty "Value", m_tButtonProperty.Value, False
    End With
    
End Sub

' //-- SelfSub v2.1 (Apr 13) code by Paul Caton (paul_caton@hotmail.com) --//

' Compacted and comments removed: Refer to the original submission for complete info
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64867&lngWId=1

Private Function sc_Subclass(ByVal lng_hWnd As Long, Optional ByVal lParamUser As Long = 0, Optional ByVal nOrdinal As Long = 1, Optional ByVal oCallback As Object = Nothing, Optional ByVal bIdeSafety As Boolean = True) As Boolean
    Const CODE_LEN As Long = 260: Const MEM_LEN As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)): Const PAGE_RWX As Long = &H40&: Const MEM_COMMIT As Long = &H1000&: Const MEM_RELEASE As Long = &H8000&: Const IDX_EBMODE As Long = 3: Const IDX_CWP As Long = 4: Const IDX_SWL As Long = 5: Const IDX_FREE As Long = 6: Const IDX_BADPTR As Long = 7: Const IDX_OWNER As Long = 8: Const IDX_CALLBACK As Long = 10: Const IDX_EBX As Long = 16: Const SUB_NAME As String = "sc_Subclass"
    Dim nAddr As Long, nID As Long, nMyID As Long
    If IsWindow(lng_hWnd) = 0 Then Exit Function
    nMyID = GetCurrentProcessId
    GetWindowThreadProcessId lng_hWnd, nID
    If nID <> nMyID Then Exit Function
    If oCallback Is Nothing Then Set oCallback = Me
    nAddr = zAddressOf(oCallback, nOrdinal)
    If nAddr = 0 Then Exit Function
    If z_Funk Is Nothing Then
        Set z_Funk = New Collection
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
        z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&
        z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA"): z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA"): z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree"): z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)
    If z_ScMem <> 0 Then
        On Error GoTo ReleaseMemory
        z_Funk.Add z_ScMem, "h" & lng_hWnd
        On Error GoTo 0
        If bIdeSafety Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")
        z_Sc(IDX_EBX) = z_ScMem: z_Sc(IDX_HWND) = lng_hWnd: z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN: z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4): z_Sc(IDX_OWNER) = ObjPtr(oCallback): z_Sc(IDX_CALLBACK) = nAddr: z_Sc(IDX_PARM_USER) = lParamUser: nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)
        If nAddr = 0 Then GoTo ReleaseMemory
        z_Sc(IDX_WNDPROC) = nAddr: RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN: sc_Subclass = True
    End If
    Exit Function
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE
End Function
Private Sub sc_Terminate()
    Dim I As Long
    If Not (z_Funk Is Nothing) Then
        For I = z_Funk.Count To 1 Step -1
            z_ScMem = z_Funk.Item(I)
            If IsBadCodePtr(z_ScMem) = 0 Then sc_UnSubclass zData(IDX_HWND)
        Next I
        Set z_Funk = Nothing
    End If
End Sub
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
    If Not (z_Funk Is Nothing) Then
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
            zData(IDX_SHUTDOWN) = -1: zDelMsg ALL_MESSAGES, IDX_BTABLE: zDelMsg ALL_MESSAGES, IDX_ATABLE
        End If
        z_Funk.Remove "h" & lng_hWnd
    End If
End Sub
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        If When And MSG_BEFORE Then zAddMsg uMsg, IDX_BTABLE
        If When And MSG_AFTER Then zAddMsg uMsg, IDX_ATABLE
    End If
End Sub
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        If When And MSG_BEFORE Then zDelMsg uMsg, IDX_BTABLE
        If When And MSG_AFTER Then zDelMsg uMsg, IDX_ATABLE
    End If
End Sub
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam)
    End If
End Function
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        sc_lParamUser = zData(IDX_PARM_USER)
    End If
End Property
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        zData(IDX_PARM_USER) = NewValue
    End If
End Property
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long, nBase As Long, I As Long
    nBase = z_ScMem: z_ScMem = zData(nTable)
    If uMsg = ALL_MESSAGES Then
        nCount = ALL_MESSAGES
    Else
        nCount = zData(0)
        If nCount >= MSG_ENTRIES Then GoTo Bail
        For I = 1 To nCount
            If zData(I) = 0 Then
                zData(I) = uMsg: GoTo Bail
            ElseIf zData(I) = uMsg Then
                GoTo Bail
            End If
        Next I
        nCount = I: zData(nCount) = uMsg
    End If
    zData(0) = nCount
Bail:
    z_ScMem = nBase
End Sub
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
    Dim nCount As Long, nBase As Long, I As Long
    nBase = z_ScMem: z_ScMem = zData(nTable)
    If uMsg = ALL_MESSAGES Then
        zData(0) = 0
    Else
        nCount = zData(0)
        For I = 1 To nCount
            If zData(I) = uMsg Then
                zData(I) = 0: GoTo Bail
            End If
        Next I
    End If
Bail:
    z_ScMem = nBase
End Sub
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
End Function
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
    If Not (z_Funk Is Nothing) Then
        On Error GoTo Catch
        z_ScMem = z_Funk("h" & lng_hWnd): zMap_hWnd = z_ScMem
    End If
Catch:
End Function
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    Dim bSub As Byte, bVal As Byte, nAddr As Long, I As Long, J As Long
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4
    If Not zProbe(nAddr + &H1C, I, bSub) Then
        If Not zProbe(nAddr + &H6F8, I, bSub) Then
            If Not zProbe(nAddr + &H7A4, I, bSub) Then Exit Function
        End If
    End If
    I = I + 4: J = I + 1024
    Do While I < J
        RtlMoveMemory VarPtr(nAddr), I, 4
        If IsBadCodePtr(nAddr) Then
            RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4: Exit Do
        End If
        RtlMoveMemory VarPtr(bVal), nAddr, 1
        If bVal <> bSub Then
            RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4: Exit Do
        End If
        I = I + 4
    Loop
End Function
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal As Byte, nAddr As Long, nLimit As Long, nEntry As Long
    nAddr = nStart: nLimit = nAddr + 32
    Do While nAddr < nLimit
        RtlMoveMemory VarPtr(nEntry), nAddr, 4
        If Not nEntry = 0 Then
            RtlMoveMemory VarPtr(bVal), nEntry, 1
            If bVal = &H33 Or bVal = &HE9 Then
                nMethod = nAddr: bSub = bVal: zProbe = True: Exit Function
            End If
        End If
        nAddr = nAddr + 4
    Loop
End Function
Private Property Get zData(ByVal nIndex As Long) As Long
    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property
Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property
Private Sub zWndProc1(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
    ' Usercontrol procedure
    Subclass_Proc bBefore, bHandled, lReturn, lng_hWnd, uMsg, wParam, lParam, lParamUser
End Sub

' Created by Noel A. Dacara | noeldacara@hotmail.com | www.dacarasoftwares.cjb.net


