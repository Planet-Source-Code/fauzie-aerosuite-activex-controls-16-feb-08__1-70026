Attribute VB_Name = "modSubClasser"
Option Explicit
' APIs used in this module
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

' used to convert VB system color variables to proper long color values
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
' used to create drawing pens/lines & DC movements
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long


' temporary - all border routines will be moved to a separate class
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Enum ButtonStateConstants
    bsNormal = 0
    bsHover = 1
    bsDown = 2
    bsDisabled = 3
End Enum
Public Enum TitlelBarBtnPosition
    tbPosDefault = 0
    tbPosLockX = 1
    tbPosLockY = 2
    tbPosStatic = 4
    tbNoFrame = 128
End Enum
Public Enum SysMenuItemConstants
    smClose = 2
    smMinimize = 4
    smMaximize = 8
    smSize = 16
    smMove = 32
    smSysIcon = 64
End Enum
Public Enum WindowBorderStyleConstants
    wbBlackEdge = 1
    wbThin = 2
    wbDialog = 3
    wbThick = 4
    wbCustom = 5
End Enum
Public Enum FontStateColorConstants
    fcEnabled = 0
    fcSelected = 1
    fcDisabled = 2
    fcInActive = 3
End Enum
Public Type SystemMenuItems
    ID As Long
    SysIcon As Long
    ItemType As Long
    Caption As String
End Type
Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
Private Type MSG
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const MSGF_MENU As Long = 2
Private Const WH_KEYBOARD As Long = 2
Private Const WH_MSGFILTER As Long = -1
Private Const WH_GETMESSAGE As Long = 3
Private Const WH_MOUSE As Long = 7

Private menuHK_ptr As Long
Private oldMenuHook As Long
Private inputHK_ptr As Long
Private oldKeyBdHook As Long
Private oldMouseHook As Long

Public Sub SetMenuHook(bSet As Boolean, callingClass As clsMenuBarControl)

If oldMenuHook Then UnhookWindowsHookEx oldMenuHook
If bSet Then
    Dim hookAddr As Long
    hookAddr = ReturnAddressOf(AddressOf MenuFilterProc)
    menuHK_ptr = ObjPtr(callingClass)
    oldMenuHook = SetWindowsHookEx(WH_MSGFILTER, hookAddr, 0, GetCurrentThreadId())
Else
    oldMenuHook = 0
    menuHK_ptr = 0
End If
End Sub

Public Sub SetInputHook(bSet As Boolean, callingClass As clsMenuBarControl)

If oldKeyBdHook Then ' currently existing hook; remove it
    UnhookWindowsHookEx oldKeyBdHook
    UnhookWindowsHookEx oldMouseHook
End If
If bSet Then
    Dim hookAddr As Long
    hookAddr = ReturnAddressOf(AddressOf KeybdFilterProc)
    inputHK_ptr = ObjPtr(callingClass)
    oldKeyBdHook = SetWindowsHookEx(WH_KEYBOARD, hookAddr, 0, GetCurrentThreadId())
    hookAddr = ReturnAddressOf(AddressOf MouseFilterProc)
    oldMouseHook = SetWindowsHookEx(WH_MOUSE, hookAddr, 0, GetCurrentThreadId())
Else
    oldKeyBdHook = 0
    oldMouseHook = 0
    inputHK_ptr = 0
End If
End Sub

Public Function GetSubClassAddr() As Long
GetSubClassAddr = ReturnAddressOf(AddressOf NewWndProc)
End Function

Private Function ReturnAddressOf(lAddress As Long) As Long
ReturnAddressOf = lAddress
End Function

Private Function NewWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim tgtClass As clsFrameControl
If GetObjectFromPointer(GetProp(hWnd, "lvCFrame_Optr"), tgtClass) Then
    NewWndProc = tgtClass.NewWndProc(wMsg, wParam, lParam)
End If
End Function

Private Function MenuFilterProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ncode = MSGF_MENU Then
    Dim tgtClass As clsMenuBarControl
    If GetObjectFromPointer(menuHK_ptr, tgtClass) Then
        If tgtClass.SetMenuAction(lParam) = True Then
            MenuFilterProc = 1
            Exit Function
        End If
    End If
End If
MenuFilterProc = CallNextHookEx(oldMenuHook, ncode, wParam, lParam)
End Function

Private Function MouseFilterProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ncode > -1 Then
    Dim tgtClass As clsMenuBarControl
    If GetObjectFromPointer(inputHK_ptr, tgtClass) Then
        'If tgtClass.SetMessageAction(wParam, lParam) = True Then
        If tgtClass.SetMouseAction(wParam, lParam) = True Then
            MouseFilterProc = 1
            Exit Function
        End If
    End If
End If
MouseFilterProc = CallNextHookEx(oldMouseHook, ncode, wParam, lParam)
End Function

Private Function KeybdFilterProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ncode > -1 Then
    Dim tgtClass As clsMenuBarControl
    If GetObjectFromPointer(inputHK_ptr, tgtClass) Then
        If tgtClass.SetKeyBdAction(wParam, lParam) = True Then
            KeybdFilterProc = 1
            Exit Function
        End If
    End If
End If
KeybdFilterProc = CallNextHookEx(oldKeyBdHook, ncode, wParam, lParam)
End Function

Public Function GetObjectFromPointer(oPtr As Long, outClass As Object) As Boolean
If oPtr Then
    Dim tgtClass As Object
    CopyMemory tgtClass, oPtr, &H4
    Set outClass = tgtClass
    CopyMemory tgtClass, 0&, &H4
    GetObjectFromPointer = True
End If
End Function

Public Function LoWord(DWord As Long) As Long
' =====================================================================
' function to return the LoWord of a Long value
' =====================================================================
     If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
     Else
        LoWord = DWord And &HFFFF&
     End If
End Function

Public Function HiWord(DWord As Long) As Long
' =====================================================================
' function to return the HiWord of a Long value
' =====================================================================
     HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
' =====================================================================
' function combines 2 Integers into a Long value
' =====================================================================
     MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Public Sub TempDrawBorders(hDC As Long, wRgn As Long, cRgn As Long)

Const BDR_RAISEDINNER As Long = &H4
Const BDR_RAISEDOUTER As Long = &H1
Const BDR_SUNKENINNER As Long = &H8
Const BDR_SUNKENOUTER As Long = &H2

Const BF_MIDDLE As Long = &H800
Const BF_LEFT As Long = &H1
Const BF_TOP As Long = &H2
Const BF_RIGHT As Long = &H4
Const BF_BOTTOM As Long = &H8

Dim eRect As RECT
Dim edgeRgn As Long


GetRgnBox wRgn, eRect
edgeRgn = CreateRectRgnIndirect(eRect) '  copy the overall window region

OffsetRgn edgeRgn, -eRect.Left, -eRect.Top  ' offset to 0,0
'OffsetRgn cRgn, -eRect.Left, -eRect.Top     ' offset client area to 0,0

CombineRgn edgeRgn, edgeRgn, cRgn, 4   ' exclude the client region

' use it for clipping region to prevent painting over client area
SelectClipRgn hDC, edgeRgn
DeleteObject edgeRgn

' draw the rectangular borders
OffsetRect eRect, -eRect.Left, -eRect.Top
DrawEdge hDC, eRect, BDR_RAISEDINNER Or BDR_RAISEDOUTER, BF_BOTTOM Or BF_LEFT Or BF_RIGHT Or BF_TOP Or BF_MIDDLE

SelectClipRgn hDC, 0
'DeleteObject cRgn

End Sub


Public Function ResizeBitmap(cDC As Long, hBmp As Long, _
        newCx As Long, newCy As Long, _
        selectInto As Long, bResized As Boolean) As Long
        
Dim bmpInfo As BITMAP

If hBmp Then GetGDIObject hBmp, Len(bmpInfo), bmpInfo
If bmpInfo.bmHeight <> newCy Or bmpInfo.bmWidth <> newCx Then
    If hBmp Then DeleteObject hBmp
    hBmp = CreateCompatibleBitmap(cDC, newCx, newCy)
    bResized = True
End If
If selectInto Then ResizeBitmap = SelectObject(selectInto, hBmp)
End Function


Public Function ConvertVBSysColor(inColor As Long) As Long
' converts a vbSystemColor variable to a long color variable

' I've never seen the GetSysColor API return an error, but just in case...
On Error GoTo ExitRoutine
If inColor < 0 Then
    ConvertVBSysColor = GetSysColor(inColor And &HFF&)
Else
    ConvertVBSysColor = inColor
End If
ExitRoutine:
End Function

Public Sub GradientFill(ByVal FromColor As Long, ByVal ToColor As Long, _
    hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, Cy As Long, _
    Optional ByVal Roughness As Byte)

' FromColor :: any valid RGB color or system color (i.e vbActiveTitleBar)
' ToColor :: any valid RGB color or system color (i.e vbInactiveTitleBar)
' hDC :: the DC to draw gradient on
' X :: left edge of gradient rectangle
' Y :: top edge of gradient rectangle
' to determine direction of gradient, pass Cx and/or Cy as follows
' Left>Right :: Cx is positive and right edge of gradient rectangle (i.e., Right)
' Right>Left :: Cx is negative and right edge of gradient rectangle (i.e., -Right)
' Bottom>Top :: Cy is negative and bottom edge of gradient rectangle (i.e., -Bottom)
' Top>Bottom :: Cx is negative & Cy is negative (i.e., -Right & -Bottom)
' Roughness :: 0=fine detail, 1-4 is lesser quality for larger rectangles
'               determines line thickness of 1,3,5,7 or 9

Dim bColor(0 To 3) As Byte, eColor(0 To 3) As Byte

'convert values like vbButtonFace to a proper RGB value
FromColor = ConvertVBSysColor(FromColor)
ToColor = ConvertVBSysColor(ToColor)

' quick easy way to convert long to RGB values
CopyMemory bColor(0), FromColor, &H3
CopyMemory eColor(0), ToColor, &H3

Dim lPtIncr As Long ' counter in positive values only
Dim lPenSize As Long ' size of drawing pen
Dim lWxHx As Long   ' adjusted width/height of gradient rectangle
Dim lPoint As Long  ' loop variables
Dim lPtStart As Long, lPtEnd As Long, lPtStep As Long
' values to add/subtracted from RGB to show next gradient color
Dim ratioRed As Single, ratioGreen As Single, ratioBlue As Single
' memory DC variables
Dim hPen As Long, hOldPen As Long

' set a maximum value. I think CreatePen API tends to max out around 10
' This value will help determine the line width/size
If Roughness > 4 Then Roughness = 4
' ensure an odd number; even number sizes may not step right in a loop
Roughness = Roughness * 2 + 1

' Setup the loop variables
If Cy < 0 Then ' vertical
    If Cx < 0 Then ' vertical top to bottom
        lPtStart = Y
        lPtEnd = Abs(Cy)
        lPtStep = Roughness
    Else            ' vertical bottom to top
        lPtStart = Abs(Cy)
        lPtEnd = Y
        lPtStep = -Roughness
    End If
Else        ' horizontal
    If Cx < 0 Then ' horizontal right to left
        lPtStep = -Roughness
        lPtStart = Abs(Cx)
        lPtEnd = X
    Else                ' horizontal left to right
        lPtStep = Roughness
        lPtStart = X
        lPtEnd = Cx
    End If
End If

' calculate the width & add a buffer of 1 to prevent RGB overflow possibility
lWxHx = Abs(lPtEnd - lPtStart) + 1
' ensure we can draw at least a minimum amount of lines
If lWxHx < Roughness Then
    ' if not, make the step value either +1 or -1 depending on current pos/neg sign
    lPtStep = lPtStep / Abs(lPtStep)
Else
' tweak to prevent situation where last line may not be drawn
' To combat this, we simply add an extra loop
    lPtEnd = lPtEnd - lPtStep * (Abs(lPtStep) > 1)
End If

' calculate color step value
ratioRed = ((eColor(0) + 0 - bColor(0)) / lWxHx)
ratioGreen = ((eColor(1) + 0 - bColor(1)) / lWxHx)
ratioBlue = ((eColor(2) + 0 - bColor(2)) / lWxHx)

' cache vs using the ABS function in the loop -- less calculations
Cx = Abs(Cx)
lPenSize = Abs(lPtStep)

' It is faster to have 2 separate loops (1 for vertical & 1 for horizontal)
' than to use one loop and put an IF statement in there to identify
' direction of drawing. Difference could be 100's of "IFs" processed.

' select the first color; then enter loop.
hOldPen = SelectObject(hDC, CreatePen(0, lPenSize, FromColor))

' these loops are pretty much identical with the only big difference
' of shifting X,Y coords to draw a vertical line or a horizontal line
If Cy < 0 Then  ' vertical loop
    For lPoint = lPtStart To lPtEnd Step lPtStep
        MoveToEx hDC, X, lPoint, ByVal 0&
        LineTo hDC, Cx, lPoint
        DeleteObject SelectObject(hDC, CreatePen(0, lPenSize, RGB( _
            bColor(0) + lPtIncr * ratioRed, _
            bColor(1) + lPtIncr * ratioGreen, _
            bColor(2) + lPtIncr * ratioBlue)))
        lPtIncr = lPtIncr + lPenSize
    Next
Else        ' horizontal loop
    For lPoint = lPtStart To lPtEnd Step lPtStep
        MoveToEx hDC, lPoint, Y, ByVal 0&
        LineTo hDC, lPoint, Cy
        DeleteObject SelectObject(hDC, CreatePen(0, lPenSize, RGB( _
            bColor(0) + lPtIncr * ratioRed, _
            bColor(1) + lPtIncr * ratioGreen, _
            bColor(2) + lPtIncr * ratioBlue)))
        lPtIncr = lPtIncr + lPenSize
    Next
End If
' destroy the last pen created & replace with original DC pen
DeleteObject SelectObject(hDC, hOldPen)

End Sub

Public Function IsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
  ' test to see if an array has been initialized
  ' Cannot be used on variants
  IsArrayEmpty = (lArrayPointer = -1)
End Function

