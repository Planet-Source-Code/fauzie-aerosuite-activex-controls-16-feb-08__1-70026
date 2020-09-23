VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuBarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Any) As Long
Private Declare Function PtInRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetMenuItemInfo_String Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO_String) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DrawFrameControl Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
Private Type LOGFONT
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
Private Type MENUITEMINFO_String        ' used to retrieve/store menu items
     cbSize As Long                      ' this structure is used with all O/S
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As String
     cch As Long
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
Private Type MSG
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Type HotTrackingStructure
    htk_Now As Long
    htk_Was As Long
    htk_Next As Long
    htk_ESC As Boolean
    htk_FromKeyBd As Boolean
    htk_Pt As POINTAPI
End Type
Private Type MenuItemDataStructure
    ID As Long
    mRect As RECT
    Flags As Long
'    isEnabled As Boolean
'    Caption As String
End Type


Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const MF_DISABLED As Long = &H2&
Private Const MIIM_ID As Long = &H2
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_STRING As Long = &H40
Private Const WM_CANCELMODE As Long = &H1F
Private Const WM_COMMAND As Long = &H111
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_MOUSEMOVE As Long = &H200
Private Const mouse_event_flag As Long = &H2
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const MOUSEEVENTF_MIDDLEDOWN As Long = &H20
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_XDOWN As Long = &H80
Private Const MOUSEEVENTF_ABSOLUTE As Long = &H8000
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_XBUTTONDOWN As Long = &H20B
Private Const WM_NCRBUTTONDOWN As Long = &HA4
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_NCLBUTTONUP As Long = &HA2

Private Const MNU_ResetAll As Long = -1
Private Const MNU_Reset As Long = -2    ' draw appropriate selection
Private Const MNU_Refresh As Long = -3  ' exit loop, update bar appropriately

Private mnuItem() As MenuItemDataStructure

Public Event RefreshMenuBar()
Public Event GetWindowRegion(lRegion As Long)

Private m_dcBitmap As Long
Private m_bmpMenuBar As Long ' static menu bar
Private m_MenuDC As Long

Private m_mbarOffset As POINTAPI
Private m_mbarLocked As Boolean
Private m_FarExtent As Long
Private m_FarExtentPrevious As Long
Private m_SingleLineMenu As Boolean
Private m_Menu As Long
Private m_HotKeys As String
Private m_Font As Long
Private p_Hwnd As Long

Private m_LastState As Long ' whether active or not
Private m_CurSelect As Long ' last menu item highlighted
Private m_TrackLevel As Integer   ' 1=hilight only, 2=select, 4= has sysicon
Private m_Tracking As Boolean
Private m_MBarXY As POINTAPI
Private HTrack As HotTrackingStructure

Private m_fColor(0 To 3) As Long
Private m_HiliteColor(0 To 1) As Long
Private m_HiliteStyle As Integer

Private Sub Class_Initialize()
m_LastState = 2
m_CurSelect = MNU_Reset
m_fColor(fcDisabled) = vbGrayText
m_fColor(fcEnabled) = vbMenuText
m_fColor(fcInActive) = vbGrayText
m_fColor(fcSelected) = vbMenuText
m_HiliteColor(0) = vbHighlightText
m_HiliteColor(1) = vb3DShadow
End Sub

Private Sub Class_Terminate()
TrackingState = False
If m_Font Then DeleteObject m_Font
If m_MenuDC Then DeleteDC m_MenuDC
If m_dcBitmap Then DeleteObject m_dcBitmap
If m_bmpMenuBar Then DeleteObject m_bmpMenuBar
If m_Menu Then SetMenu p_Hwnd, m_Menu
m_bmpMenuBar = 0
m_dcBitmap = 0
m_Menu = 0
m_Font = 0
Erase mnuItem
End Sub

Public Property Get TrackingState() As Boolean
TrackingState = m_Tracking
End Property
Public Property Let TrackingState(isTracking As Boolean)
If isTracking Then
    m_Tracking = True   ' no routine set this to true; really used to reset
Else
    m_Tracking = False
    'If SetInputHook(False, Nothing) Then ReleaseCapture
    SetInputHook False, Nothing
    SetMenuHook False, Nothing
    'Debug.Print "trackingstate reset via property /wm_cancelmode"
End If
End Property


Public Property Set Font(newFont As StdFont)
Dim nFont As LOGFONT
If m_Font Then
    DeleteObject m_Font
Else
    Dim hDC As Long
    hDC = GetWindowDC(GetDesktopWindow())
    m_MenuDC = CreateCompatibleDC(hDC)
    ReleaseDC GetDesktopWindow(), hDC
    SetBkMode m_MenuDC, 3
End If
If newFont Is Nothing Then
    Dim ncm As NONCLIENTMETRICS
    ncm.cbSize = Len(ncm)
    ' this will return the system menu font info
    SystemParametersInfo 41, 0, ncm, 0
    nFont = ncm.lfMenuFont
Else
    With newFont
        nFont.lfFaceName = .Name & String$(32, 0)
        nFont.lfHeight = (.Size * -20) / Screen.TwipsPerPixelY
        nFont.lfItalic = Abs(.Italic)
        nFont.lfStrikeOut = Abs(.Strikethrough)
        nFont.lfUnderline = Abs(.Underline)
        nFont.lfWeight = Abs(.Bold) * 300 + 400
    End With
End If
nFont.lfCharSet = 1
m_Font = CreateFontIndirect(nFont)
End Property

Public Property Get Font() As StdFont
Dim tFont As StdFont, nFont As LOGFONT
If m_Font Then
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
Else
    Set tFont = New StdFont
    tFont.Name = "Tahoma"
    tFont.Size = 8.25
End If
On Error Resume Next
Set Font = tFont
End Property

Public Property Let FontColor(lState As FontStateColorConstants, lColor As Long)
If lState < 0 Or lState > fcInActive Then Exit Property
m_fColor(lState) = lColor
ForceMenuBarRepaint
End Property
Public Property Get FontColor(lState As FontStateColorConstants) As Long
If lState < 0 Or lState > fcInActive Then Exit Property
FontColor = m_fColor(lState)
End Property

Public Property Get WindowMenu(dbgID As String) As Long
WindowMenu = m_Menu
End Property

Public Property Let WindowMenu(dbgID As String, Zero As Long)
Dim hMenu As Long
hMenu = GetMenu(p_Hwnd)
If hMenu Then
    m_Menu = hMenu
    'Debug.Print " resetting menu via mnubar resetmenu from "; dbgID
    SetMenu p_Hwnd, 0
    DrawMenuBar p_Hwnd
'    ForceMenuBarRepaint
'    Stop
End If

End Property

Public Property Let MenubarStatic(isStatic As Boolean)
m_mbarLocked = isStatic
End Property
Public Property Get MenubarStatic() As Boolean
MenubarStatic = m_mbarLocked
End Property

Friend Property Let ParentWIndow(hWnd As Long)
p_Hwnd = hWnd

Dim hMenu As Long
hMenu = GetMenu(p_Hwnd)
If hMenu Then m_Menu = hMenu
End Property

Private Property Get MenuItemCount(hSubMenu As Long) As Long
If m_Menu Then
    If hSubMenu Then
        MenuItemCount = GetMenuItemCount(hSubMenu)
    Else
        MenuItemCount = GetMenuItemCount(m_Menu)
    End If
End If
End Property

Public Property Let ReleaseRefreshLock(bRelease As Boolean)
m_LastState = -2
End Property

Public Function HiliteStyle(bSet As Boolean, isFlat As Boolean) As Boolean
If bSet Then
    m_HiliteStyle = Abs(isFlat)
Else
    HiliteStyle = (m_HiliteStyle = 1)
End If
    
End Function

Public Sub HiliteColors(Color1 As Long, Color2 As Long)
m_HiliteColor(0) = Color1
m_HiliteColor(1) = Color2
End Sub

Public Sub ForceMenuBarRepaint()
m_LastState = -2
RaiseEvent RefreshMenuBar
End Sub

Friend Sub DrawMenuBarBkg(hDC As Long, destDC As Long, mbarRight As Long, mBarHeight As Long, isActive As Boolean)
Dim bSizeChanged As Boolean, bUserDrawn As Boolean
Dim deskDC As Long, hOldBmp As Long, hNewBmp As Long

    ' ensure bitmap is the right size
    hOldBmp = ResizeBitmap(destDC, m_dcBitmap, mbarRight, mBarHeight, hDC, bSizeChanged)
    If bSizeChanged Or CLng(isActive) <> m_LastState Then
        m_LastState = CLng(isActive)
        hNewBmp = ResizeBitmap(destDC, m_bmpMenuBar, mbarRight, mBarHeight, m_MenuDC, False)
        Dim callBack As CustomWindowCalls
        If GetObjectFromPointer(GetProp(p_Hwnd, "lvImpCB_Ptr"), callBack) Then
            Call callBack.UserDrawnMenuBar(hDC, mbarRight, mBarHeight, isActive, bUserDrawn)
        End If
        If Not bUserDrawn Then
            Dim hBrush As Long, dcRect As RECT
            SetRect dcRect, 0, 0, mbarRight, mBarHeight
            hBrush = CreateSolidBrush(ConvertVBSysColor(vbButtonFace))
            FillRect hDC, dcRect, hBrush
            DeleteObject hBrush
        End If
        BitBlt m_MenuDC, 0, 0, mbarRight, mBarHeight, hDC, 0, 0, vbSrcCopy
    Else
        hNewBmp = SelectObject(m_MenuDC, m_bmpMenuBar)
    End If
    BitBlt hDC, 0, 0, mbarRight, mBarHeight, m_MenuDC, 0, 0, vbSrcCopy
    SelectObject m_MenuDC, hNewBmp
    SelectObject hDC, hOldBmp

End Sub

Friend Sub GetBarDimensions(hDC As Long, mBarWidth As Long, mBarHeight As Long, xLeft As Long, yTop As Long)

' gotta be a faster way, but can't think of one.
' Unless a skinned window has a fixed menubar dimension, in order
' to measure the menubar, we have to re-measure each menubar item.
' Measuring the menubar is triggered by window sizing, some
' window style changes, and changing a menu item's properties....
' This is understandable and can be easily
' trapped/identified. However, whenever DrawMenuBar is called,
' windows is sent a message that makes it appear the window resized.
' Do we ignore it? No, cause maybe a menu item was added/deleted or
' possibly the caption changed which causes menubar wrap. Open to
' suggestions; thankfully menu items on a menubar don't number in
' the dozens!  Caching the captions doesn't appear to be a way to
' measure quicker since you need to check each menu caption anyway
' and then do an IF to see if it changed or IF the properties changed.

' cache for use elsewhere. This the menubar left/top in window coords
m_MBarXY.X = xLeft
m_MBarXY.Y = yTop

If hDC = 0 Then Exit Sub

Dim nrItems As Long, deskDC As Long
Dim hOldBmp As Long, hNewBmp As Long
Dim bSizeChanged As Boolean, bUserDrawn As Boolean


m_HotKeys = ""  ' reset so it can be built on demand if needed
nrItems = MenuItemCount(0)

If nrItems Then
    
    If m_Font = 0 Then Set Font = Nothing
    
    
    If m_mbarLocked = False Then
            
        Dim itemNr As Long, hOldFont As Long
        Dim iRect As RECT, mnuCaption As String
        Dim lState As Long, mnuID As Long
        Const lineSpacer As Integer = 12
        
        
        ReDim Preserve mnuItem(0 To nrItems - 1)
        hOldFont = SelectObject(m_MenuDC, m_Font)
        
        On Error Resume Next
        
        For itemNr = 0 To nrItems - 1
            mnuCaption = GetMenuItemCaption(itemNr, mnuItem(itemNr).Flags, mnuItem(itemNr).ID)
            iRect.Left = iRect.Right
            DrawText m_MenuDC, mnuCaption, -1, iRect, DT_SINGLELINE Or DT_LEFT Or DT_CALCRECT
            iRect.Right = iRect.Right + lineSpacer
            If iRect.Right >= mBarWidth And iRect.Left > 0 Then
                OffsetRect iRect, -iRect.Left, -iRect.Top + iRect.Bottom + lineSpacer \ 2
            End If
            SetRect mnuItem(itemNr).mRect, iRect.Left, iRect.Top, iRect.Right, iRect.Bottom + lineSpacer \ 2
        Next
        SelectObject m_MenuDC, hOldFont
        mBarHeight = mnuItem(nrItems - 1).mRect.Bottom

        
'            'Debug.Print "recalculated menubar dimensions"
    End If
    
    
Else
    mBarHeight = 0
End If
End Sub

Private Function GetMenuItemCaption(itemPos As Long, Optional fState As Long, Optional fID As Long) As String
Dim MIS As MENUITEMINFO_String
' annotation needed:  menu captions max length of 128 characters
MIS.dwTypeData = String$(128, 0)
MIS.fMask = MIIM_ID Or MIIM_STATE Or MIIM_STRING
MIS.cbSize = LenB(MIS)
MIS.cch = 128 + 1
GetMenuItemInfo_String m_Menu, itemPos, True, MIS
''Debug.Print vbTab; "cch is "; MIS.cch; MIS.wID
If MIS.cch > 0 Then
'    'Debug.Print vbTab; Left$(MIS.dwTypeData, MIS.cch)
    GetMenuItemCaption = Left$(MIS.dwTypeData, MIS.cch)
End If
fID = MIS.wID
fState = MIS.fState
''Debug.Print MIS.wID, MIS.fType
End Function

Friend Sub DrawMenuBarItems(hDC As Long, destDC As Long, _
    menuRect As RECT, windowRect As RECT, _
    ByVal itemSelect As Long, ByVal RefreshFlag As Long)

If m_Font = 0 Or hDC = 0 Then Exit Sub
'Exit Sub

'RaiseEvent MenuBoundary(hDC, xOffset, yOffset, farRight)

''Debug.Print "drawing menu items"
If MenuItemCount(0) = 0 Then Exit Sub

Dim nrItems As Long, itemNr As Long
Dim dcRect As RECT, mBar As RECT, tRect As RECT
Dim hOldBmp As Long, hMnuBmp As Long
Dim hFont As Long, bSwapColor As Boolean, fColor As Long
Dim mCaption As String, clipRgn As Long


If menuRect.Right = menuRect.Left Then ' null rectangle passed
    SetRect menuRect, m_MBarXY.X, m_MBarXY.Y, m_MBarXY.X, m_MBarXY.Y
Else
    m_MBarXY.X = menuRect.Left
    m_MBarXY.Y = menuRect.Top
End If
If windowRect.Right = windowRect.Left Then ' null window rect passed
    GetWindowRect p_Hwnd, windowRect
End If

' convert menubar from screen coords to nonclient coords
mBar = menuRect
OffsetRect mBar, -windowRect.Left, -windowRect.Top

If itemSelect = MNU_ResetAll Then
    m_CurSelect = MNU_Reset     ' reset to no previous selection
    nrItems = UBound(mnuItem)
    itemNr = 0
    ' select cached bitmap and target bitmap
    hOldBmp = SelectObject(hDC, m_dcBitmap)
    hMnuBmp = SelectObject(m_MenuDC, m_bmpMenuBar)
    ' refresh target bitmap vs redrawing all over again
    BitBlt hDC, 0, 0, dcRect.Right, dcRect.Bottom, m_MenuDC, 0, 0, vbSrcCopy
    SelectObject m_MenuDC, hMnuBmp
    ' select the working menu font into DC
    hFont = SelectObject(hDC, m_Font)
    If m_LastState Then ' active titlebar
        fColor = ConvertVBSysColor(m_fColor(fcEnabled))
    Else
        fColor = ConvertVBSysColor(m_fColor(fcInActive))
    End If
    SetTextColor hDC, fColor
    
    ' create an offset for copying from target bitmap to window DC
    dcRect = mBar   ' copy that rectangle & offset to 0,0 for DC drawing
    OffsetRect dcRect, -mBar.Left, -mBar.Top

    For itemNr = 0 To nrItems
        With mnuItem(itemNr)
            ' todo: finish idea here.... for skinning and having a fixed menubar dimension
'            If .mRect.Right > dcRect.Right And m_mbarLocked = True Then Exit For
            ' move menuitem rectangle to window coordinates
            tRect = .mRect
        End With
        ' shift menu rectangle to window DC coords
        OffsetRect tRect, dcRect.Left, dcRect.Top
        ' get caption & print to DC
        mCaption = GetMenuItemCaption(itemNr)
        If ((mnuItem(itemNr).Flags And MF_DISABLED) = MF_DISABLED) And m_LastState <> 0 Then
            SetTextColor hDC, ConvertVBSysColor(m_fColor(fcDisabled))
            bSwapColor = True
        End If
        DrawText hDC, mCaption, -1, tRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        If bSwapColor Then
            SetTextColor hDC, fColor
            bSwapColor = False
        End If
    Next

    ' draw finished menubar onto the window
    BitBlt destDC, mBar.Left, mBar.Top, dcRect.Right, dcRect.Bottom, hDC, 0, 0, vbSrcCopy
    ' replaced font & original bitmap
    SelectObject hDC, hOldBmp
    SelectObject hDC, hFont
    m_TrackLevel = 0
    m_Tracking = False
    
Else    ' updating menubar due to keyboard/mouse navigation
    If m_Tracking Then Exit Sub
    
    If RefreshFlag = MNU_Refresh Or RefreshFlag = MNU_Reset * 100 Then
        Dim mPt As POINTAPI
        GetCursorPos mPt
        itemSelect = GetHitTest(mPt.X - m_MBarXY.X, mPt.Y - m_MBarXY.Y) - 1
        m_TrackLevel = 0
        RefreshFlag = MNU_Refresh
    Else
        If m_CurSelect = MNU_Reset And itemSelect = MNU_Reset Then Exit Sub
    End If
    
    RaiseEvent GetWindowRegion(clipRgn)
    If m_CurSelect <> itemSelect Or RefreshFlag <> 0 Then
        If m_CurSelect > -1 Then
            ' no repeating selection or removing any hilighted item
            ' setup DC and select appropriate bitmap/font
            destDC = GetWindowDC(p_Hwnd)
            SelectClipRgn destDC, clipRgn
            SetBkMode destDC, 3
            hFont = SelectObject(destDC, m_Font)
            hMnuBmp = SelectObject(m_MenuDC, m_bmpMenuBar)
            ' shift rect to dc coords
            dcRect = mnuItem(m_CurSelect).mRect
            OffsetRect dcRect, mBar.Left, mBar.Top
            With mnuItem(m_CurSelect).mRect
                ' get caption & print to DC
                mCaption = GetMenuItemCaption(m_CurSelect)
                If (mnuItem(m_CurSelect).Flags And MF_DISABLED) = MF_DISABLED Then
                    fColor = SetTextColor(destDC, ConvertVBSysColor(m_fColor(fcDisabled)))
                ElseIf m_LastState Then
                    fColor = SetTextColor(destDC, ConvertVBSysColor(m_fColor(fcEnabled)))
                Else
                    fColor = SetTextColor(destDC, ConvertVBSysColor(m_fColor(fcInActive)))
                End If
                ' erase previously hilighted item
                BitBlt destDC, dcRect.Left, dcRect.Top, dcRect.Right - dcRect.Left, dcRect.Bottom - dcRect.Top, m_MenuDC, .Left, .Top, vbSrcCopy
                DrawText destDC, mCaption, -1, dcRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            End With
            SetTextColor destDC, fColor
            ' unselect all
            SelectObject destDC, hFont
            If RefreshFlag = 0 Then
                m_CurSelect = MNU_Reset ' reset to none selected
            Else
                If RefreshFlag = MNU_Refresh Then m_CurSelect = itemSelect
            End If
        End If
    End If
    If (RefreshFlag <> 0 Or itemSelect <> m_CurSelect) And itemSelect > -1 Then
        ' no repeating and a new item will be selected
        If destDC = 0 Then  ' setup DC as needed
            destDC = GetWindowDC(p_Hwnd)
            hMnuBmp = SelectObject(m_MenuDC, m_bmpMenuBar)
            SetBkMode destDC, 3
            SelectClipRgn destDC, clipRgn
        End If
        m_CurSelect = itemSelect    ' cache newest selection
        ' shift rect to dc coords
        dcRect = mnuItem(m_CurSelect).mRect
        OffsetRect dcRect, mBar.Left, mBar.Top
        BitBlt destDC, dcRect.Left, dcRect.Top, dcRect.Right - dcRect.Left, dcRect.Bottom - dcRect.Top, m_MenuDC, mnuItem(m_CurSelect).mRect.Left, mnuItem(m_CurSelect).mRect.Top, vbSrcCopy
        InflateRect dcRect, -1, -1
        If (mnuItem(m_CurSelect).Flags And MF_DISABLED) <> MF_DISABLED Then DrawHilite destDC, dcRect, ((m_TrackLevel Or 2) = m_TrackLevel)
        If (mnuItem(m_CurSelect).Flags And MF_DISABLED) = MF_DISABLED Then
            SetTextColor destDC, ConvertVBSysColor(m_fColor(fcDisabled))
        ElseIf m_LastState Then
            fColor = SetTextColor(destDC, ConvertVBSysColor(m_fColor(fcSelected)))
        Else
            fColor = SetTextColor(destDC, ConvertVBSysColor(m_fColor(fcInActive)))
        End If
        If m_HiliteStyle = 0 Then
            If m_HiliteStyle = 0 And ((m_TrackLevel Or 2) = m_TrackLevel) Then OffsetRect dcRect, 1, 1
        End If
        hFont = SelectObject(destDC, m_Font)
        mCaption = GetMenuItemCaption(m_CurSelect)
        DrawText destDC, mCaption, -1, dcRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        SelectObject destDC, hFont
        SetTextColor destDC, fColor
    End If
    If destDC Then
        If hMnuBmp Then SelectObject m_MenuDC, hMnuBmp
        SelectClipRgn destDC, 0
        ReleaseDC p_Hwnd, destDC
    End If
End If

End Sub

Private Sub DrawHilite(hDC As Long, mRect As RECT, isPushed As Boolean)
Dim bColor1 As Long, bColor2 As Long
Dim xOffset As Long, yOffset As Long
Dim hOldPen As Long, hBrush As Long

If m_HiliteStyle Then
'    If isPushed Then
        hBrush = CreateSolidBrush(ConvertVBSysColor(m_HiliteColor(1)))
        FillRect hDC, mRect, hBrush
        DeleteObject hBrush
'    End If
    hBrush = CreateSolidBrush(ConvertVBSysColor(m_HiliteColor(0)))
    FrameRect hDC, mRect, hBrush
    DeleteObject hBrush
Else
    If isPushed Then
        bColor1 = m_HiliteColor(1)
        bColor2 = m_HiliteColor(0)
'        mRect.Bottom = mRect.Bottom + 1
    Else
        bColor1 = m_HiliteColor(0)
        bColor2 = m_HiliteColor(1)
    End If
    hOldPen = SelectObject(hDC, CreatePen(0, 1, ConvertVBSysColor(bColor1)))
    MoveToEx hDC, mRect.Left, mRect.Bottom, ByVal 0&
    LineTo hDC, mRect.Left, mRect.Top
    LineTo hDC, mRect.Right, mRect.Top
    DeleteObject SelectObject(hDC, CreatePen(0, 1, ConvertVBSysColor(bColor2)))
    LineTo hDC, mRect.Right, mRect.Bottom
    LineTo hDC, mRect.Left - 1, mRect.Bottom
    DeleteObject SelectObject(hDC, hOldPen)
End If
End Sub

Friend Function GetHitTest(X As Long, Y As Long) As Long
' identifies if mouse is over a menubar menu item
Dim itemNr As Long
For itemNr = 0 To UBound(mnuItem)
    If PtInRect(mnuItem(itemNr).mRect, X, Y) Then
        GetHitTest = itemNr + 1 ' append 1 here,it will be remove later
        Exit Function
    End If
Next
End Function

Friend Sub TrackSystemMenu(sysRegion As Long, leftAligned As Boolean)
' displays the system menu
Dim mPts As POINTAPI, sysRect As RECT
Dim X As Long, Y As Long, tpmFlags As Long
    
Const TPM_BOTTOMALIGN As Long = &H20&

If sysRegion = 0 Then
    GetCursorPos mPts
    X = mPts.X: Y = mPts.Y
Else
    GetRgnBox sysRegion, sysRect
    If leftAligned Then
        X = sysRect.Right + 2
        Y = sysRect.Bottom
        tpmFlags = TPM_BOTTOMALIGN
    Else
        X = sysRect.Left
        Y = sysRect.Bottom + 2
    End If
End If
If m_Tracking Then ' called as part of the menu loop
    TrackMenuBar -100, True, True, False, True, X, Y, tpmFlags
Else                ' default display method
    TrackPopupMenu GetSystemMenu(p_Hwnd, 0), tpmFlags, X, Y, 0, p_Hwnd, ByVal 0&
End If
End Sub

Friend Sub TrackMenuBar(ByVal menuPos As Long, ByVal isClicked As Boolean, _
    ByVal hasSysIcon As Boolean, ByVal hiliteOnly As Boolean, _
    Optional bSelectFirst As Boolean, _
    Optional ByVal X As Long, Optional ByVal Y As Long, _
    Optional ByVal pFlags As Long)

' This routine and the next four below are all part of the menu loop
' A bit unique from others out there cause it does handle keyboard
' loops to include the system menu & also handles menubar menu items
' that do not have submenus....

' THE MENU LOOP PART 1 - Showing the appropriate submenu/systemmenu

' Parameters....
' menuPos is the zero-based menubar item position
' isClicked is same as mouse down
' hasSysIcon indicates keybd navigation can go to the system menu in a loop
' hiliteOnly only will not display submenus; otherwise it will
' bSelectFirst will attempt to hilight the 1st menu item in submenu (disabled or not)
' X & Y are optional screen coords to display popup menus
' pFlags are optional and generally used here to redraw/refresh menubar

Dim subID As Long
Dim mnuRect As RECT, wndRect As RECT

If isClicked = True And hiliteOnly = False Then ' user clicked on the submenu item
    If menuPos > -1 Then    ' not the system menu which is either -99 or -100
        ' Test for upper level menus without submenus -- some people do this...
        subID = GetSubMenu(m_Menu, menuPos) ' see if it has a submenu
        If subID = 0 Then                   ' no submenu - fire the event
            TrackingState = False   ' release hooks & refresh menubar
            DrawMenuBarItems 1, 0, mnuRect, wndRect, MNU_Reset, MNU_Refresh
            ' send the click event
            If (mnuItem(menuPos).Flags And MF_DISABLED) <> MF_DISABLED Then _
                PostMessage p_Hwnd, WM_COMMAND, mnuItem(menuPos).ID, ByVal 0&
            HTrack.htk_FromKeyBd = False
            HTrack.htk_ESC = False
            Exit Sub
        End If
    End If
Else
    If isClicked Then
        subID = GetSubMenu(m_Menu, menuPos) ' see if it has a submenu
    Else
        ' special meaning and used internally in this class...only when not displaying items
        If pFlags = MNU_Refresh Or pFlags = MNU_Reset Or pFlags = MNU_Reset * 100 Then
            DrawMenuBarItems 1, 0, mnuRect, wndRect, menuPos, pFlags
            If pFlags = MNU_Reset * 100 Then
                ' special case: user exited menu with escape while submenus were
                ' displayed. Windows generally remains in the loop and this message
                ' will re=enter the loop to mimic that functionality
                TrackingState = False
                PostMessage p_Hwnd, &H112, &HF100&, -menuPos
            End If
            HTrack.htk_FromKeyBd = False
            HTrack.htk_ESC = False
            Exit Sub
        End If
    End If
End If

''Debug.Print "entering menu loop with "; menuPos

' build the hotkey list if needed
If Len(m_HotKeys) = 0 Then BuildHotKeyList

' calc tracking level :: 1=hilite only, 2=auto select submenu
m_TrackLevel = 1 - isClicked
If hasSysIcon Then m_TrackLevel = m_TrackLevel Or 4

Do
    
    m_Tracking = False     ' initially turn of tracking
    If menuPos = -99 Or menuPos = -100 Then ' system menu here
        ' erase any previously highlighted menubar items
        DrawMenuBarItems 1, 0, mnuRect, wndRect, MNU_Reset, 0
    Else    ' menubar item, highlight it
        DrawMenuBarItems 1, 0, mnuRect, wndRect, menuPos, pFlags
        pFlags = 0
    End If
    m_Tracking = True ' set property for subclass wndProc to check
    
    ' regardless whether or not a menu item has a submenu, if we are
    ' on the 1st level of tracking (hilighting only) then don't show submenu
    If (m_TrackLevel Or 1) = m_TrackLevel Then subID = 0
    
    SetProp p_Hwnd, "nextMnuPos", -1    ' set a default "failure" value
    SetInputHook False, Nothing         ' ensure keybd/mouse hook disabled
    
    GetCursorPos HTrack.htk_Pt    ' cache current cursor
    
    If subID > 0 Then   ' menubar item has a submenu - easy
    
        SetMenuHook True, Me    ' sneak a peek on mouse movements
        
        ' show the menubar submenu & select 1st item if appropriate
        If bSelectFirst Then
            PostMessage p_Hwnd, &H100&, vbKeyDown, ByVal 0&
            bSelectFirst = False
        End If
        TrackPopupMenu subID, pFlags, _
            m_MBarXY.X + mnuItem(menuPos).mRect.Left, _
            m_MBarXY.Y + mnuItem(menuPos).mRect.Bottom, 0, p_Hwnd, ByVal 0&
        
        SetMenuHook False, Nothing  ' disable menu hook
        ' get the next menu to display set in routine: SetMenuAction
        menuPos = GetProp(p_Hwnd, "nextMnuPos")
        m_Tracking = False ' reset variable
        
    ElseIf menuPos = -99 Then   ' keyboard left/right to system menu
        ' forward message back so routines can prevent VB's min/max/close drawing
        PostMessage p_Hwnd, &H313, &HF100&, ByVal 0&
        menuPos = -1    ' prevent looping; the above postmessage will get back here
        
    ElseIf menuPos = -100 Then  ' sysmenu is back and ready to display
        
        SetMenuHook True, Me    ' sneak a peek on mouse movements
        
        ' show the system menu with optional X,Y & alignment flags
        ' and also auto-select the 1st menu item
        PostMessage p_Hwnd, &H100&, vbKeyDown, ByVal 0&
        TrackPopupMenu GetSystemMenu(p_Hwnd, 0), pFlags, X, Y, 0, p_Hwnd, ByVal 0&
        
        menuPos = GetProp(p_Hwnd, "nextMnuPos") ' get next menu to display
        pFlags = 0      ' reset if ever used
        m_Tracking = False ' reset flag
        
    Else ' main menu item without submenus, or tracking in hilight mode only
         ' We'll need to track via mouse & keyboard hooks
        menuPos = -1            ' prevent looping for now
        m_Tracking = True ' set flag
        SetInputHook True, Me   ' start the keybd/mouse hooks
    End If
    
    HTrack.htk_FromKeyBd = False  ' flag to indicate menu navigated via keyboard
    If menuPos = -1 Then   ' is there a next menu to display?
        Exit Do ' no next menu or hilight tracking only; boogy out
    Else
        subID = 0           ' reset variables
        With HTrack         ' this is used to help navigate popups
            .htk_Was = 0    ' with keyboard and mouse simultaneously
            .htk_Now = 0
            .htk_Next = 0
            .htk_ESC = False        'flag to indicate ESC last pressed or not
        End With
        ' now see if the next menu item has a submenu or not
        If menuPos > -1 Then subID = GetSubMenu(m_Menu, menuPos)
        ' return back up to the loop
    End If
Loop
''Debug.Print "exiting menu loop"
End Sub

Friend Function SetMenuAction(lParam As Long) As Boolean

' THE MENU LOOP PART 2 - Normal popups including system menu

' This routine is the callback function from the MessageHook &
' is used only for menubar items with submenus.

''Debug.Print " getting menu messages "

If lParam = 0 Or m_Tracking = False Then Exit Function

Dim newMsg As MSG
Dim newMenuID As Long

' copy the message so we can look at it
CopyMemory newMsg, ByVal lParam, Len(newMsg)


With newMsg
    Select Case .message    ' these are common WM_flags
    Case WM_MOUSEMOVE
        ' ensure the point is somewhere in our window!
        If WindowFromPoint(.pt.X, .pt.Y) = p_Hwnd Then
            ' each popup usually is followed by 1 or more mouse moves
            ' Since we cached the current coords, we can prevent misfiring
            ' when cursor is over one item, but user keyboards to another
            ' and when it opens, the mouse move message would trigger the
            ' routines to go back to the item where the cursor is. Annoying.
            If HTrack.htk_Pt.X <> .pt.X Or HTrack.htk_Pt.Y <> .pt.Y Then
                ' ok, a new spot, get a hit test on menuitems
                newMenuID = GetHitTest(.pt.X - m_MBarXY.X, .pt.Y - m_MBarXY.Y) - 1
                ' ensure not already highlighted and is enabled...
                If newMenuID <> m_CurSelect And newMenuID > -1 Then
                    ' set flag & post cancel message will will then
                    ' release the loop so it can continue
                    SetProp p_Hwnd, "nextMnuPos", newMenuID
                    PostMessage p_Hwnd, WM_CANCELMODE, 0, ByVal 0&
                End If
            End If
        End If
    Case WM_KEYDOWN ' keyboard navigation
        HTrack.htk_ESC = False  ' indicate last key pressed is not the ESC key
        ' when navigated from submenu-less item to here, the right or left
        ' Key_Down message will follow unfortunately. So we use the
        ' HTTrack.htk_FromKeyBd set in that routine to warn us here.
        ' For description of other HTTrack elements, see further below.
        Select Case .wParam     ' virtual key codes
        Case vbKeyLeft
            If (HTrack.htk_Now = HTrack.htk_Was Or HTrack.htk_Was = 0) _
                And HTrack.htk_FromKeyBd = False Then
                ' set flag to next menu item, send cnx message to loop
                SetProp p_Hwnd, "nextMnuPos", NextMenuPosition(True, True)
                PostMessage p_Hwnd, WM_CANCELMODE, 0, ByVal 0&
                SetMenuAction = True    ' eat this message
            Else
                HTrack.htk_FromKeyBd = False    ' reset flag
            End If
        Case vbKeyRight
            If HTrack.htk_Next = 0 And HTrack.htk_FromKeyBd = False Then
                'set flag to next menu item, send cnx message to loop
                SetProp p_Hwnd, "nextMnuPos", NextMenuPosition(False, True)
                PostMessage p_Hwnd, WM_CANCELMODE, 0, ByVal 0&
                SetMenuAction = True    ' eat this message
            Else
                HTrack.htk_FromKeyBd = False    ' reset flag
            End If
        Case vbKeyReturn
            If (HTrack.htk_Now = 0 Or HTrack.htk_FromKeyBd = True) Then
                HTrack.htk_FromKeyBd = False
                SetMenuAction = True
            End If
        Case vbKeyEscape
            HTrack.htk_ESC = True   ' identify as last key pressed
        End Select
    
    Case &H11F  ' wm_menuselect
        If HiWord(.wParam) = -1 And .lParam = 0 Then ' menu is closing
            If m_CurSelect < 0 Then ' system menu
                ' if no other menu items exist then exit loop completely
                HTrack.htk_Now = NextMenuPosition(False, False)
                If HTrack.htk_Now < 0 Then HTrack.htk_ESC = False
            Else
                HTrack.htk_Now = m_CurSelect ' cache current item
            End If
            TrackingState = False           ' reset tracking variables
            If HTrack.htk_ESC Then          ' last key pressed was an ESC
                ' no need to check tracking level 'cause this routine is
                ' only called when tracking level = 2. So we go back to
                ' tracking level one...
                SendMessage p_Hwnd, WM_CANCELMODE, 0, ByVal 0&
                TrackMenuBar HTrack.htk_Now, False, False, False, , , , MNU_Reset * 100
                SetMenuAction = True
            Else    ' some other key or mouse stroke, loss of focus, etc
                    ' caused menu to close; completely exit menu loop
                TrackMenuBar 0, False, False, False, , , , MNU_Refresh
            End If
        Else    ' got a selection; track a few things...
            ' If this menu item as a submenu, we can let O/S handle
            ' right arrow message; otherwise, we move to next menubar item
            HTrack.htk_Next = ((HiWord(.wParam) And &H10&) = &H10&) 'GetSubMenu(.lParam, LoWord(.wParam))
            ' get current submenu handle this item belongs to
            ' It will help to determine whether the left arrow can be
            ' handled by O/S or we move to previous menubar item
            HTrack.htk_Now = .lParam
            ' Set a point where menu can no longer go back using left
            ' arrows, regardless how deep the submenus go which would
            ' change the .htk_Next & .htk_Now values
            If HTrack.htk_Was = 0 Then HTrack.htk_Was = .lParam
            With HTrack
                'Debug.Print .htk_Was; .htk_Now; .htk_Next;
            End With
            'Debug.Print .wParam; ((HiWord(.wParam) And &H8000&) = &H8000&)
        End If
    Case WM_LBUTTONDOWN ' clicking...
        HTrack.htk_ESC = False
        ' get a current hit test
        newMenuID = GetHitTest(.pt.X - m_MBarXY.X, .pt.Y - m_MBarXY.Y) - 1
        If newMenuID < 0 Then
            TrackMenuBar 0, False, False, False, , , , MNU_Refresh
        ElseIf newMenuID = m_CurSelect Then ' clicking on current item
            ' cancel the popup
            TrackMenuBar newMenuID, True, ((m_TrackLevel Or 4) = m_TrackLevel), True, , , , MNU_Refresh
            PostMessage p_Hwnd, WM_CANCELMODE, 0, ByVal 0&
            SetMenuAction = True    ' eat this one
        Else    ' clicking on a menu item most likely
            TrackMenuBar 0, False, False, False, , , , MNU_Refresh
        End If
    Case Else
'        'Debug.Print "menu hook msg"; .message; .wParam; .lParam
    End Select
End With
End Function


Friend Function SetMouseAction(wParam As Long, lParam As Long) As Boolean

If m_Tracking = False Then Exit Function

' THE MENU LOOP PART 3 - Menubar items without submenus (Mouse Hook)
' We are manually tracking the menubar at this point...

' This routine is the callback function from the Mouse Hook &
' is used only for menubar items without submenus.

' hittest of 30 is a custom value I forced on windows whenever a
' mouse rode over a menubar item....

''Debug.Print wMsg; wHitTest; menuItem


Dim mPt As POINTAPI, menuItem As Long
Select Case wParam
    Case WM_LBUTTONDOWN, &HA1, WM_NCRBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_XBUTTONDOWN
        TrackingState = False
        TrackMenuBar 0, False, False, False, , , , MNU_Refresh
        m_TrackLevel = 0
    Case WM_LBUTTONUP, WM_NCLBUTTONUP
        GetCursorPos mPt
        menuItem = GetHitTest(mPt.X - m_MBarXY.X, mPt.Y - m_MBarXY.Y) - 1
        If menuItem < 0 Then
            TrackingState = False
            TrackMenuBar 0, False, False, False, , , , MNU_Refresh
        Else
            TrackMenuBar menuItem, True, False, False
        End If
        SetMouseAction = True
    Case Else
        Dim inMsg As MOUSEHOOKSTRUCT
        CopyMemory inMsg, ByVal lParam, Len(inMsg)
        If inMsg.hWnd = p_Hwnd Then
            With inMsg
                If .wHitTestCode = 30 Then
                    SetMouseAction = True
                    GetCursorPos mPt
                    Select Case wParam
                        Case &HA0 'WM_NCMOUSEMOVE
                            If (HTrack.htk_Pt.X <> mPt.X Or HTrack.htk_Pt.Y <> mPt.Y) Then
                                ' get current hittest to a specific item
                                menuItem = GetHitTest(mPt.X - m_MBarXY.X, mPt.Y - m_MBarXY.Y) - 1
                                If menuItem < 0 Then Exit Function
                                If menuItem <> m_CurSelect And menuItem > -1 Then
                                    ' ok, not the same item
                                    ' tell message loop to restart with this menuitem
                                    TrackMenuBar menuItem, ((m_TrackLevel Or 2) = m_TrackLevel), ((m_TrackLevel Or 4) = m_TrackLevel), (GetSubMenu(m_Menu, menuItem) = 0)
                                End If
                            End If
                        Case Else
                    End Select
                End If
            End With
        End If
End Select

End Function

Friend Function SetKeyBdAction(wKeyCode As Long, lParam As Long) As Boolean

If m_Tracking = False Then
    'Debug.Print "entering keyboard hook "; m_Tracking
    TrackingState = False
    SetKeyBdAction = True
    Exit Function
End If
' THE MENU LOOP PART 4 - Menubar items without submenus (Keyboard Hook)
' We are manually tracking the menubar at this point...

' This routine is the callback function from the Keyboard Hook &
' is used only for menubar items without submenus.

Dim bKeyUp As Boolean, bShowMenuItem As Boolean, bHasSubMenu As Boolean
Dim newMenuPos As Long, tgtKey As String * 1

bKeyUp = ((lParam And &H80000000) = &H80000000)

If Not bKeyUp Then

    Select Case wKeyCode
    
    Case vbKeyF10, vbKeyMenu    ' second alt keydown; end menu loop
        TrackingState = False
        TrackMenuBar 0, False, False, False, , , , MNU_Refresh
    
    Case vbKeyEscape    ' not much to do here; just removehooks
        TrackingState = False
        TrackMenuBar 0, False, False, False, , , , MNU_Refresh
    
    Case vbKeyUp, vbKeyDown, vbKeyReturn ' no action on submenu-less items
        If GetSubMenu(m_Menu, m_CurSelect) Then
            HTrack.htk_FromKeyBd = (wKeyCode = vbKeyReturn)
            TrackMenuBar m_CurSelect, True, ((m_TrackLevel Or 4) = m_TrackLevel), False, (wKeyCode = vbKeyReturn), , , MNU_Reset
        ElseIf wKeyCode = vbKeyReturn Then
            TrackingState = False
            TrackMenuBar m_CurSelect, True, False, False
        End If

    Case vbKeySpace     ' if a sysicon then show it
        If (m_TrackLevel Or 4) = m_TrackLevel Then TrackMenuBar -99, True, True, False
        
    Case vbKeyLeft, vbKeyRight ' navigating at menubar
        ' get the next valid menu position
        newMenuPos = NextMenuPosition((vbKeyLeft = wKeyCode), ((m_TrackLevel Or 2) = m_TrackLevel))
        If newMenuPos > -1 Then ' got another menubar item...
            bShowMenuItem = True
            HTrack.htk_FromKeyBd = True
        ElseIf newMenuPos = -99 Then        ' got to the system menu
            m_Tracking = False
            HTrack.htk_FromKeyBd = True
            TrackMenuBar newMenuPos, True, True, False
        Else ' there isn't any other items; just 1 menu bar item
            If m_CurSelect = MNU_Reset Then ' system menu
                ' Windows keep user in menu loop;
                ' let's just exit out since user cannot really tell
                m_Tracking = False
                TrackMenuBar 0, False, False, False, , , , MNU_Reset
            End If
        End If
    Case Else ' accelerator key? let's check
        Dim bHasDups As Boolean, dupMenuPos As Long
        tgtKey = LCase(Chr$(wKeyCode))
        newMenuPos = InStr(m_HotKeys, tgtKey) - 1
        If newMenuPos > -1 Then
            If (InStr(newMenuPos + 2, m_HotKeys, tgtKey) > 0) Then
                ' we have more than one hotkey like this one!
                ' 1st see if we are currently on that one?
                If Mid$(m_HotKeys, m_CurSelect + 1, 1) = tgtKey Then
                    ' currently on a dup key, let's find next one
                    dupMenuPos = InStr(m_CurSelect + 2, m_HotKeys, tgtKey)
                    If dupMenuPos > 0 Then newMenuPos = dupMenuPos - 1
                End If
            End If
            bShowMenuItem = True
            HTrack.htk_FromKeyBd = False
        Else
            TrackingState = False
            TrackMenuBar 0, False, False, False, , , , MNU_Refresh
            If wKeyCode <> vbKeyTab Then
                Beep
                'Debug.Print "beep beep beep beep" 'cute? some testing done on PC w/o speakers
            End If
        End If
    End Select
    If bShowMenuItem Then
        m_Tracking = False ' reset flag
        bHasSubMenu = (GetSubMenu(m_Menu, newMenuPos) <> 0)
        If ((m_TrackLevel Or 1) = m_TrackLevel) Then HTrack.htk_FromKeyBd = False
        TrackMenuBar newMenuPos, ((m_TrackLevel Or 2) = m_TrackLevel), ((m_TrackLevel Or 4) = m_TrackLevel), Not bHasSubMenu, bHasSubMenu
    End If
End If
SetKeyBdAction = True
End Function

Friend Function TraceHotKey(wKeyCode As Long, hasSysIcon As Boolean, isSysCommand As Boolean) As Long

Dim mnuPos As Long, tgtKey As String * 1
Dim lRtnVal As Long

Const MNC_CLOSE As Long = 1
Const MNC_EXECUTE As Long = 2
Const MNC_IGNORE As Long = 0
Const MNC_SELECT As Long = 3

If Len(m_HotKeys) = 0 Then BuildHotKeyList
tgtKey = LCase(Chr$(wKeyCode))
mnuPos = InStr(m_HotKeys, tgtKey) - 1


' send a bad menu execute command for a valid menu item
' otherwise system will beep
lRtnVal = MakeDWord(-1, MNC_EXECUTE)

If mnuPos < 0 Then
    If isSysCommand Then
        Beep
        'Debug.Print "beep beep beep beep" 'cute? some testing done on PC w/o speakers
    Else
        lRtnVal = 0 ' system will beep for us
    End If
    ' beep manually if sent via wm_syscommand; otherwise wm_menuchar will beep for us
Else
    If isSysCommand = False And m_Tracking = True Then
        If InStr(m_HotKeys, tgtKey) = 0 Then lRtnVal = 0
    Else
        If InStr(mnuPos + 2, m_HotKeys, tgtKey) Then
            ' have more than one menu item with this hot key
            TrackMenuBar mnuPos, False, hasSysIcon, True
        Else
            If GetSubMenu(m_Menu, mnuPos) = 0 Then ' submenuless-menubar item
                ' simply call routine to execute the item
                TrackMenuBar mnuPos, True, False, False
            Else                                    ' has submenu
                If isSysCommand Then
                    ' display the submenu
                    TrackMenuBar mnuPos, True, hasSysIcon, False, True
                Else    ' received via wm_menuchar message, this message
                        ' will close the menu immediately since it no
                        ' longer exists via SetMenu commands. So we will
                        ' post a new message to activate it
                    lRtnVal = MakeDWord(mnuPos + 0, MNC_CLOSE)
                    PostMessage p_Hwnd, &H112, &HF100&, ByVal wKeyCode
                End If
            End If
        End If
    End If
End If
TraceHotKey = lRtnVal
End Function

Private Sub BuildHotKeyList()

Dim I As Integer, mnuCaption As String
Dim ampPos As Integer, lastPos As Integer, sList As String

For I = 0 To MenuItemCount(0) - 1
    ampPos = 0
    mnuCaption = LCase$(GetMenuItemCaption(I + 0, 0, 0)) & Chr$(0)
    lastPos = Len(mnuCaption)
    ampPos = InStr(ampPos + 1, mnuCaption, "&")
    Do While ampPos
        If ampPos Then
            If Mid$(mnuCaption, ampPos + 1, 1) = "&" Then
                ampPos = ampPos + 1
            Else
                lastPos = ampPos + 1
            End If
        End If
        ampPos = InStr(ampPos + 1, mnuCaption, "&")
    Loop
    If lastPos = Len(mnuCaption) Or ((mnuItem(I).Flags Or MF_DISABLED) = mnuItem(I).Flags) Then
        sList = sList & Chr$(0)
    Else
        sList = sList & Mid$(mnuCaption, lastPos, 1)
    End If
Next
m_HotKeys = sList
'Debug.Print "hot key list is >"; sList; "<"
End Sub

Private Function NextMenuPosition(bLeft As Boolean, includeSysMenu As Boolean) As Long

' Routine finds the next available menubar item, including system menu
' Although currently used only for keyboard navigation;
' includeSysMenu paramater is included to test vertically for
' owner-drawn menus when the time comes

Dim newPos As Long

If m_CurSelect = MNU_Reset Then ' on the system menu, not menubar
    ' move to the first or last item on menubar depending on direction
    If MenuItemCount(0) Then
        newPos = UBound(mnuItem) * Abs(bLeft = True)
        HTrack.htk_FromKeyBd = True
    Else
        newPos = -1
    End If
Else            ' on the menubar, goto 1 previous or 1 next
    If bLeft Then newPos = m_CurSelect - 1 Else newPos = m_CurSelect + 1
    If newPos < 0 Then
        If includeSysMenu Then newPos = -99 Else newPos = UBound(mnuItem)
    ElseIf newPos > UBound(mnuItem) Then
        If includeSysMenu Then newPos = -99 Else newPos = 0
    End If
End If
NextMenuPosition = newPos   ' return the next value
End Function
