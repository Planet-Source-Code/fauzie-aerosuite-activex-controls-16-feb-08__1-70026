VERSION 5.00
Begin VB.UserControl AeroOptionButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   105
   ToolboxBitmap   =   "AeroOptionButton.ctx":0000
End
Attribute VB_Name = "AeroOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------
' AeroOptionButton ActiveX Control
'--------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'--------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'--------------------------------------------------------------------------

Option Explicit

'********************************'
'* Subclasser Declarations      *'
'********************************'
Private Const ALL_MESSAGES          As Long = -1
Private Const GMEM_FIXED            As Long = 0
Private Const GWL_WNDPROC           As Long = -4
Private Const PATCH_04              As Long = 88
Private Const PATCH_05              As Long = 93
Private Const PATCH_08              As Long = 132
Private Const PATCH_09              As Long = 137
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_MOVING             As Long = &H216
Private Const WM_SIZING             As Long = &H214
Private Const WM_EXITSIZEMOVE       As Long = &H232

Private Type tSubData
 hWnd                               As Long
 nAddrSub                           As Long
 nAddrOrig                          As Long
 nMsgCntA                           As Long
 nMsgCntB                           As Long
 aMsgTblA()                         As Long
 aMsgTblB()                         As Long
End Type

Private Enum eMsgWhen
 MSG_AFTER = 1
 MSG_BEFORE = 2
 MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
 TME_HOVER = &H1&
 TME_LEAVE = &H2&
 TME_QUERY = &H40000000
 TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
 cbSize                             As Long
 dwFlags                            As TRACKMOUSEEVENT_FLAGS
 hwndTrack                          As Long
 dwHoverTime                        As Long
End Type

Private bTrack                      As Boolean
Private bTrackUser32                As Boolean
Private bInCtrl                     As Boolean
Private sc_aSubData()               As tSubData

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
'*******************************************************'

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Enum AlignOption
    [Align Left] = 0
    [Align Right] = 1
End Enum

'events
Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Public Event MouseEnter()
Public Event MouseLeave()

'variables
Private m_Align As AlignOption

Private He As Long  'the height of the button
Private Wi As Long  'the width of the button
Private CheckRect As RECT

Private m_Caption As String     'current text

Private rc As RECT

Private LastButton As Byte, LastKeyDown As Byte
Private m_Enabled As Boolean
Private hasFocus As Boolean, m_ShowFocusRect As Boolean

Private lastStat As Byte, TE As String, isShown As Boolean  'used to avoid unnecessary repaints
Private isOver As Boolean

Private m_Value As Boolean

Private ThePics As New c32bppDIB
Private myObj As Object

'* ========================================================================================================
'*  Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'* ========================================================================================================
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
  '* Parameters:
  '*  bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  '*  bHandled - Set this variable to True in a before callback to prevent the message being subsequently processed by the default handler... and if set, an after callback
  '*  lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  '*  hWnd     - The window handle
  '*  uMsg     - The message number
  '*  wParam   - Message related data
  '*  lParam   - Message related data
  '* Notes:
  '*  If you really know what youre doing, its possible to change the values of the _
      hWnd, uMsg, wParam and lParam parameters in a before callback so that different _
      values get passed to the default handler.. and optionaly, the after callback.
  Select Case uMsg
  Case WM_MOUSELEAVE
    isOver = False
    Call Redraw(0, True)
    RaiseEvent MouseLeave
  End Select
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
  fAbout.Show vbModal
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  LastButton = 1
  Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Click()
  If LastButton = 1 And m_Enabled Then
    m_Value = True
    Call Redraw(0, True)
    UserControl.Refresh
  End If
End Sub

Private Sub UserControl_DblClick()
  If LastButton = 1 Then
    Call UserControl_MouseDown(1, 0, 0, 0)
    SetCapture hWnd
  End If
End Sub

Private Sub UserControl_GotFocus()
  hasFocus = True
  CheckAllValue False
  m_Value = True
  Call Redraw(lastStat, True)
  UserControl.Refresh
  RaiseEvent Click
End Sub

Private Sub UserControl_Hide()
  isShown = False
End Sub

Private Sub UserControl_Initialize()
  ThePics.LoadPicture_Resource "OPTION", "PNG"
  isShown = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
  
  LastKeyDown = KeyCode
  Select Case KeyCode
  Case 32 'spacebar pressed
    Call Redraw(2, False)
  Case 39, 40 'right and down arrows
    SendKeys "{Tab}"
  Case 37, 38 'left and up arrows
    SendKeys "+{Tab}"
  End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)

  If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed, and not cancelled by the user
    m_Value = IIf(m_Value <> 1, 1, 0)
    Call Redraw(0, False)
    UserControl.Refresh
    RaiseEvent Click
  End If
End Sub

Private Sub UserControl_LostFocus()
  hasFocus = False
  Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_InitProperties()
  m_Enabled = True: m_ShowFocusRect = True
  m_Caption = Ambient.DisplayName
  Set UserControl.Font = Ambient.Font
  UserControl.BackColor = Ambient.BackColor
  UserControl.ForeColor = vbBlack
  Call CalcTextRects
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
  LastButton = Button
  If Button = 1 Then Call Redraw(2, True)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
  If Not IsMouseOver Then
    Call Redraw(0, False)
  Else
    If Button <> 1 And Not isOver Then
      Call TrackMouseLeave(hWnd)
      isOver = True
      Call Redraw(0, True)
      RaiseEvent MouseEnter
    ElseIf Button = 1 Then
      isOver = True
      Call Redraw(2, False)
      isOver = False
    End If
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
  If Button = 1 Then Call Redraw(0, False)
End Sub

Public Property Get Align() As AlignOption
Attribute Align.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Align = m_Align
End Property

Public Property Let Align(ByVal New_Align As AlignOption)
  m_Align = New_Align
  Call Refresh
  PropertyChanged "Align"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
  UserControl.BackColor = theCol
  Call Redraw(lastStat, True)
  PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
  UserControl.ForeColor = theCol
  Call Redraw(lastStat, True)
  PropertyChanged "ForeColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
  m_Caption = NewValue
  Call SetAccessKeys
  Call CalcTextRects
  Call Redraw(0, True)
  PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
  On Error Resume Next
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  m_Enabled = NewValue
  Call Redraw(0, True)
  UserControl.Enabled = m_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef NewFont As Font)
  Set UserControl.Font = NewFont
  Call CalcTextRects
  Call Redraw(0, True)
  PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontBold.VB_MemberFlags = "400"
  FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal NewValue As Boolean)
  UserControl.FontBold = NewValue
  Call CalcTextRects
  Call Redraw(0, True)
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontItalic.VB_MemberFlags = "400"
  FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal NewValue As Boolean)
  UserControl.FontItalic = NewValue
  Call CalcTextRects
  Call Redraw(0, True)
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontUnderline.VB_MemberFlags = "400"
  FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal NewValue As Boolean)
  UserControl.FontUnderline = NewValue
  Call CalcTextRects
  Call Redraw(0, True)
End Property

Public Property Get FontSize() As Integer
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontSize.VB_MemberFlags = "400"
  FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal NewValue As Integer)
  UserControl.FontSize = NewValue
  Call CalcTextRects
  Call Redraw(0, True)
End Property

Public Property Get FontName() As String
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
  FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal NewValue As String)
  UserControl.FontName = NewValue
  Call CalcTextRects
  Call Redraw(0, True)
End Property

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = ";Misc"
  ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal NewValue As Boolean)
  m_ShowFocusRect = NewValue
  Call Redraw(lastStat, True)
  PropertyChanged "ShowFocusRect"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)
  UserControl.MousePointer = newPointer
  PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
  Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal newIcon As StdPicture)
  On Local Error Resume Next
  Set UserControl.MouseIcon = newIcon
  PropertyChanged "MouseIcon"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
  hWnd = UserControl.hWnd
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Behavior"
  Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Boolean)
  m_Value = NewValue
  Call Redraw(0, True)
  PropertyChanged "Value"
End Property

Private Sub UserControl_Resize()
  He = UserControl.ScaleHeight: Wi = UserControl.ScaleWidth
  Call CalcTextRects
  
  If He Then Call Redraw(0, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    m_Align = .ReadProperty("Align", 0)
    m_Caption = .ReadProperty("Caption", "")
    m_Enabled = .ReadProperty("Enabled", True)
    Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
    m_ShowFocusRect = .ReadProperty("ShowFocusRect", True)
    UserControl.BackColor = .ReadProperty("BackColor", RGB(240, 240, 240))
    UserControl.ForeColor = .ReadProperty("ForeColor", vbBlack)
    UserControl.MousePointer = .ReadProperty("MousePointer", 0)
    Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
    m_Value = .ReadProperty("Value", 0)
  End With
  UserControl.Enabled = m_Enabled
  Call CalcTextRects
  Call SetAccessKeys
  If Ambient.UserMode Then
    bTrack = True
    bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
    If Not (bTrackUser32 = True) Then
      If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
        bTrack = False
      End If
    End If
    If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
      '* Start subclassing the UserControl.
      Call Subclass_Start(hWnd)
      Call Subclass_AddMsg(hWnd, WM_MOUSELEAVE, MSG_AFTER)
    End If
  End If
End Sub

Private Sub UserControl_Show()
  isShown = True
  Call Redraw(0, True)
End Sub

Private Sub UserControl_Terminate()
  isShown = False
  On Error GoTo Catch
  Call Subclass_StopAll '* Stop all subclassing.
  Exit Sub
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("Align", m_Align)
    Call .WriteProperty("Caption", m_Caption)
    Call .WriteProperty("Enabled", m_Enabled)
    Call .WriteProperty("Font", UserControl.Font)
    Call .WriteProperty("ShowFocusRect", m_ShowFocusRect)
    Call .WriteProperty("BackColor", UserControl.BackColor)
    Call .WriteProperty("ForeColor", UserControl.ForeColor)
    Call .WriteProperty("MousePointer", UserControl.MousePointer)
    Call .WriteProperty("MouseIcon", UserControl.MouseIcon)
    Call .WriteProperty("Value", m_Value)
  End With
End Sub

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
  Dim i As Integer
  If Not Force Then  'check drawing redundancy
    If (curStat = lastStat) And (TE = m_Caption) Then Exit Sub
  End If

  If He = 0 Or Not isShown Then Exit Sub   'we don't want errors

  lastStat = curStat
  TE = m_Caption

  i = IIf(m_Value, 52, 0)
  UserControl.Cls
  If m_Enabled Then
    If curStat = 0 And isOver Then curStat = 1
    ThePics.Render UserControl.hdc, CheckRect.Left, CheckRect.Top, 13, 13, 0, (curStat * 13) + i, 13, 13
  Else
    ThePics.Render UserControl.hdc, CheckRect.Left, CheckRect.Top, 13, 13, 0, (3 * 13) + i, 13, 13
  End If
  Call DrawCaption
  DrawFocusR
End Sub

Private Sub DrawFocusR()
  If m_ShowFocusRect And hasFocus Then
    SetTextColor UserControl.hdc, vbBlack
    DrawFocusRect UserControl.hdc, rc
  End If
End Sub

Private Sub SetAccessKeys()
  Dim ampersandPos As Long
  
  UserControl.AccessKeys = ""
  
  If Len(m_Caption) > 1 Then
    ampersandPos = InStr(1, m_Caption, "&", vbTextCompare)
    If (ampersandPos < Len(m_Caption)) And (ampersandPos > 0) Then
      If Mid$(m_Caption, ampersandPos + 1, 1) <> "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
        UserControl.AccessKeys = LCase$(Mid$(m_Caption, ampersandPos + 1, 1))
      Else 'do only a second pass to find another ampersand character
        ampersandPos = InStr(ampersandPos + 2, m_Caption, "&", vbTextCompare)
        If Mid$(m_Caption, ampersandPos + 1, 1) <> "&" Then
          UserControl.AccessKeys = LCase$(Mid$(m_Caption, ampersandPos + 1, 1))
        End If
      End If
    End If
  End If
End Sub

Private Sub CalcChkPos()
  Select Case m_Align
  Case 0 'left
    SetRect CheckRect, 0, (He - 13) \ 2, 13, ((He - 13) \ 2) + 13
  Case 1 'right
    SetRect CheckRect, Wi - 13 - 1, (He - 13) \ 2, Wi, ((He - 13) \ 2) + 13
  End Select
End Sub

Private Sub CalcTextRects()
  Select Case m_Align
  Case 0
    rc.Left = 16: rc.Right = Wi - 2: rc.Top = 1: rc.Bottom = He - 2
  Case 1
    rc.Left = 1: rc.Right = Wi - 2 - 13: rc.Top = 1: rc.Bottom = He - 2
  End Select
  DrawText UserControl.hdc, m_Caption, Len(m_Caption), rc, DT_CALCRECT Or DT_WORDBREAK
  OffsetRect rc, 0, (He - rc.Bottom) \ 2

  Call CalcChkPos
End Sub

Public Sub DisableRefresh()
  isShown = False
End Sub

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
  Call CalcTextRects
  isShown = True
  Call Redraw(lastStat, True)
End Sub

Private Sub DrawCaption()
  With UserControl
    If m_Enabled Then
      SetTextColor .hdc, TranslateColor(.ForeColor)
    Else
      SetTextColor .hdc, RGB(128, 128, 128)
    End If
    DrawText .hdc, m_Caption, Len(m_Caption), rc, DT_LEFT
  End With
End Sub

Private Function IsMouseOver() As Boolean
  Dim pt As POINTAPI
  GetCursorPos pt
  IsMouseOver = (WindowFromPoint(pt.X, pt.Y) = hWnd)
End Function

Private Sub CheckAllValue(ByVal isValue As Boolean)
  For Each myObj In Parent.Controls
    If (TypeOf myObj Is AeroOptionButton) Then
      If Not (myObj.Container Is UserControl.Parent) Then
        If (myObj.hWnd = UserControl.hWnd) Then
          Call CheckContainerControls(myObj.Container, isValue)
          Exit Sub
        End If
      End If
    End If
  Next
  Call CheckContainerControls(UserControl.Parent, False)
End Sub

Private Sub CheckContainerControls(ByVal cContainer As Object, ByVal ctlValue As Boolean)
  For Each myObj In Parent.Controls
    If (TypeOf myObj Is AeroOptionButton) Then
      If (myObj.Container Is cContainer) Then
        If Not (myObj.hWnd = UserControl.hWnd) Then
          If (myObj.Value = True) Then myObj.Value = ctlValue
        End If
      End If
    End If
  Next
End Sub

'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hmod As Long, bLibLoaded As Boolean
  
  hmod = GetModuleHandleA(sModule)
  If (hmod = 0) Then
    hmod = LoadLibraryA(sModule)
    If (hmod) Then bLibLoaded = True
  End If
  If (hmod) Then
    If (GetProcAddress(hmod, sFunction)) Then IsFunctionExported = True
  End If
  If (bLibLoaded = True) Then Call FreeLibrary(hmod)
End Function

'* Track the mouse leaving the indicated window.
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If (bTrack = True) Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With
    If (bTrackUser32 = True) Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
End Sub

'* ============================================================================================================================
'*  Subclass code - The programmer may call any of the following Subclass_??? routines
'*  Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'* ============================================================================================================================
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  '* Parameters:
  '*  lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  '*  uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  '*  When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If (When) And (eMsgWhen.MSG_BEFORE) Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If (When) And (eMsgWhen.MSG_AFTER) Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'* Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  '* Parameters:
  '*  lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  '*  uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  '*  When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If (When) And (eMsgWhen.MSG_BEFORE) Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If (When) And (eMsgWhen.MSG_AFTER) Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'* Return whether were running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'* Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
  '* Parameters:
  '*  lng_hWnd  - The handle of the window to be subclassed.
  '*  Returns;
  '*  The sc_aSubData() index.
  Const CODE_LEN              As Long = 200
  Const FUNC_CWP              As String = "CallWindowProcA"
  Const FUNC_EBM              As String = "EbMode"
  Const FUNC_SWL              As String = "SetWindowLongA"
  Const MOD_USER              As String = "user32"
  Const MOD_VBA5              As String = "vba5"
  Const MOD_VBA6              As String = "vba6"
  Const PATCH_01              As Long = 18
  Const PATCH_02              As Long = 68
  Const PATCH_03              As Long = 78
  Const PATCH_06              As Long = 116
  Const PATCH_07              As Long = 121
  Const PATCH_0A              As Long = 186
  Static aBuf(1 To CODE_LEN)  As Byte
  Static pCWP                 As Long
  Static pEbMode              As Long
  Static pSWL                 As Long
  Dim i                       As Long
  Dim j                       As Long
  Dim nSubIdx                 As Long
  Dim sHex                    As String
  
  '* If its the first time through here..
  If (aBuf(1) = 0) Then
    '* The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
    '* Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While (j < CODE_LEN)
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
      i = i + 2
    Loop
    '* Get API function addresses.
    If (Subclass_InIDE = True) Then
      aBuf(16) = &H90
      aBuf(17) = &H90
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
      If (pEbMode = 0) Then pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
    End If
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
    ReDim sc_aSubData(0 To 0) As tSubData
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If (nSubIdx = -1) Then
      nSubIdx = UBound(sc_aSubData()) + 1
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
    End If
    Subclass_Start = nSubIdx
  End If
  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
  End With
End Function

'* Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  On Error GoTo myErr
  i = UBound(sc_aSubData())
  Do While (i >= 0)
    With sc_aSubData(i)
      If (.hWnd <> 0) Then Call Subclass_Stop(.hWnd)
    End With
    i = i - 1
  Loop
  Exit Sub
myErr:
End Sub

'* Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
  '* Parameters:
  '*  lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
    Call zPatchVal(.nAddrSub, PATCH_05, 0)
    Call zPatchVal(.nAddrSub, PATCH_09, 0)
    Call GlobalFree(.nAddrSub)
    .hWnd = 0
    .nMsgCntB = 0
    .nMsgCntA = 0
    Erase .aMsgTblB
    Erase .aMsgTblA
  End With
End Sub

'* ======================================================================================================
'*  These z??? routines are exclusively called by the Subclass_??? routines.
'*  Worker sub for Subclass_AddMsg
'* ======================================================================================================
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long, nOff1 As Long, nOff2 As Long
  
  If (uMsg = ALL_MESSAGES) Then
    nMsgCnt = ALL_MESSAGES
  Else
    Do While (nEntry < nMsgCnt)
      nEntry = nEntry + 1
      If (aMsgTbl(nEntry) = 0) Then
        aMsgTbl(nEntry) = uMsg
        Exit Sub
      ElseIf (aMsgTbl(nEntry) = uMsg) Then
        Exit Sub
      End If
    Loop
    nMsgCnt = nMsgCnt + 1
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
    aMsgTbl(nMsgCnt) = uMsg
  End If
  If (When = eMsgWhen.MSG_BEFORE) Then
    nOff1 = PATCH_04
    nOff2 = PATCH_05
  Else
    nOff1 = PATCH_08
    nOff2 = PATCH_09
  End If
  If (uMsg <> ALL_MESSAGES) Then Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
  Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

'* Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc
End Function

'* Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If (uMsg = ALL_MESSAGES) Then
    nMsgCnt = 0
    If (When = eMsgWhen.MSG_BEFORE) Then
      nEntry = PATCH_05
    Else
      nEntry = PATCH_09
    End If
    Call zPatchVal(nAddr, nEntry, 0)
  Else
    Do While (nEntry < nMsgCnt)
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then
        aMsgTbl(nEntry) = 0
        Exit Do
      End If
    Loop
  End If
End Sub

'* Get the sc_aSubData() array index of the passed hWnd.
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
  '* Get the upper bound of sc_aSubData() - If you get an error here, youre probably Subclass_AddMsg-ing before Subclass_Start.
  zIdx = UBound(sc_aSubData)
  Do While (zIdx >= 0)
    With sc_aSubData(zIdx)
      If (.hWnd = lng_hWnd) And Not (bAdd = True) Then
        Exit Function
      ElseIf (.hWnd = 0) And (bAdd = True) Then
        Exit Function
      End If
    End With
    zIdx = zIdx - 1
  Loop
  If Not (bAdd = True) Then Debug.Assert False
  '* If we exit here, were returning -1, no freed elements were found.
End Function

'* Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'* Patch the machine code buffer at the indicated offset with the passed value.
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'* Worker function for Subclass_InIDE.
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function
