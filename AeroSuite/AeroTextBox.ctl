VERSION 5.00
Begin VB.UserControl AeroTextBox 
   BackColor       =   &H00B2ACA5&
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   194
   ToolboxBitmap   =   "AeroTextBox.ctx":0000
End
Attribute VB_Name = "AeroTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------
' AeroTextBox ActiveX Control
'--------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'--------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'--------------------------------------------------------------------------

Option Explicit

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const CODE_LEN               As Long = 200                                      'Length of the machine code in bytes
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
  sCode                              As String
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array
Private sc_aBuf(1 To CODE_LEN)       As Byte                                            'Code buffer byte array
Private sc_pCWP                      As Long                                            'Address of the CallWindowsProc
Private sc_pEbMode                   As Long                                            'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                      As Long                                            'Address of the SetWindowsLong function
  
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

'\\

Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type
Private Const LOGPIXELSY        As Long = 90
Private Const FW_NORMAL         As Long = 400
Private Const FW_BOLD           As Long = 700

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Const SW_SHOW           As Long = 5
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Enum eScrollBar
  [None]
  [Horizontal]
  [Vertical]
  [Both]
End Enum

Private Const EM_GETPASSWORDCHAR = &HD2
Private Const EM_SETPASSWORDCHAR = &HCC

Private Const ES_NUMBER = &H2000&
Private Const ES_MULTILINE = &H4&
Private Const ES_AUTOHSCROLL = &H80&
Private Const ES_AUTOVSCROLL = &H40&
Private Const ES_READONLY = &H800&
Private Const ES_CENTER = &H1&
Private Const ES_LEFT = &H0&
Private Const ES_RIGHT = &H2&
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
Private Const WS_TABSTOP             As Long = &H10000
Private Const WS_BORDER              As Long = &H800000
Private Const WS_CHILD               As Long = &H40000000

Private m_tRect As RECT
Private m_bInitialized           As Boolean
Private m_hEditBox               As Long
Private m_lX                     As Long
Private m_lY                     As Long
Private m_hFont                  As Long

'Default Property Values:
Const m_def_BorderStyle = 1
Const m_def_ScrollBars = 0
Const m_def_MultiLine = 0

'Property Variables:
Dim m_BorderStyle As Integer
Dim m_Scrollbars As eScrollBar
Dim m_Multiline As Boolean
Dim m_PasswordChar As String
Dim m_Text As String

'Event Declarations:
Event Click() 'MappingInfo=Text1,Text1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Text1,Text1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Text1,Text1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."

'========================================================================================
' Subclass handler: MUST be the first Public routine in this file.
'                   That includes public properties also.
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

'Parameters:
'   bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'   bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'   lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'   lng_hWnd - The window handle
'   uMsg     - The message number
'   wParam   - Message related data
'   lParam   - Message related data
'
'Notes:
'   If you really know what you're doing, it's possible to change the values of the
'   hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'   values get passed to the default handler.. and optionaly, the 'after' callback
  
  
  Dim uPoint  As POINTAPI
'  Dim uRect   As RECT2
  Dim hNode   As Long
  Dim lfHit   As Long
  Dim hEdit   As Long
  Dim nCancel As Integer
  Dim sText   As String
  Dim X       As Long
  Dim Y       As Long
  
    Select Case lng_hWnd
                
        Case UserControl.hWnd
           
            Select Case uMsg
                
                Case WM_SETFOCUS
                
                    Call SetFocus(m_hEditBox)
            
            End Select
            
        Case m_hEditBox
            
            Select Case uMsg
            
                Case WM_KEYDOWN
                    
                    RaiseEvent KeyDown(wParam And &H7FFF&, pvShiftState())
                    
                Case WM_CHAR
                    
                    RaiseEvent Change
                    RaiseEvent KeyPress(wParam And &H7FFF&)
                    
                Case WM_KEYUP
                    
                    RaiseEvent KeyUp(wParam And &H7FFF&, pvShiftState())
                    
                Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
                    
                    Call pvGetClientCursorPos(X, Y)
                    RaiseEvent MouseDown(pvButton(uMsg), pvShiftState(), CSng(X), CSng(Y))
                    
                Case WM_MOUSEMOVE
                    
                    Call pvGetClientCursorPos(X, Y)
                    If (X <> m_lX Or X <> m_lY) Then
                        m_lX = X
                        m_lY = Y
                        RaiseEvent MouseMove(pvButton(uMsg), pvShiftState(), CSng(X), CSng(Y))
                    End If
                    
                Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP

                    Call pvGetClientCursorPos(X, Y)
                    RaiseEvent MouseUp(pvButton(uMsg), pvShiftState(), CSng(X), CSng(Y))
                    RaiseEvent Click
                  
            End Select
    End Select
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  Dim uLogFont As LOGFONT

  If (m_hEditBox) Then
    Call pvDestroyFont
    Call pvStdFontToLogFont(UserControl.Font, uLogFont)
    m_hFont = CreateFontIndirect(uLogFont)
    Call SendMessage(m_hEditBox, WM_SETFONT, m_hFont, 0)
  End If
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  UserControl.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
  hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
  PasswordChar = m_PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
  If Len(New_PasswordChar) > 1 Then Exit Property
  m_PasswordChar = New_PasswordChar
  If m_hEditBox Then Call SendMessage(m_hEditBox, EM_SETPASSWORDCHAR, Asc(m_PasswordChar), ByVal 0&)
  PropertyChanged "PasswordChar"
End Property

'******************************************************
'***** Please, help me for these properties ***********
'******************************************************
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Text1,Text1,-1,MaxLength
'Public Property Get maxLength() As Long
'  maxLength = Text1.maxLength
'End Property
'
'Public Property Let maxLength(ByVal New_MaxLength As Long)
'  Text1.maxLength() = New_MaxLength
'  PropertyChanged "MaxLength"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Text1,Text1,-1,Alignment
'Public Property Get Alignment() As Integer
'  Alignment = Text1.Alignment
'End Property
'
'Public Property Let Alignment(ByVal New_Alignment As Integer)
'  Text1.Alignment() = New_Alignment
'  PropertyChanged "Alignment"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Text1,Text1,-1,SelLength
'Public Property Get SelLength() As Long
'  SelLength = Text1.SelLength
'End Property
'
'Public Property Let SelLength(ByVal New_SelLength As Long)
'  Text1.SelLength() = New_SelLength
'  PropertyChanged "SelLength"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Text1,Text1,-1,SelStart
'Public Property Get SelStart() As Long
'  SelStart = Text1.SelStart
'End Property
'
'Public Property Let SelStart(ByVal New_SelStart As Long)
'  Text1.SelStart() = New_SelStart
'  PropertyChanged "SelStart"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Text1,Text1,-1,SelText
'Public Property Get SelText() As String
'  SelText = Text1.SelText
'End Property
'
'Public Property Let SelText(ByVal New_SelText As String)
'  Text1.SelText() = New_SelText
'  PropertyChanged "SelText"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
  If m_hEditBox <> 0 Then
    Dim tmpLen As Long, tmpBuff As String * 32767
    tmpLen = GetWindowText(m_hEditBox, tmpBuff, 32767)
    m_Text = Left(tmpBuff, tmpLen)
  End If
  Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
  m_Text = New_Text
  If Not Ambient.UserMode Then Call UserControl_Paint
  SetWindowText m_hEditBox, m_Text
  PropertyChanged "Text"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_Multiline = m_def_MultiLine
  m_BorderStyle = m_def_BorderStyle
  m_Scrollbars = m_def_ScrollBars
  m_Text = Ambient.DisplayName
End Sub

Private Sub UserControl_Paint()
  If Not Ambient.UserMode Then
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), vbWhite, BF
    DrawText hDC, m_Text, -1, m_tRect, 0
  End If
  If m_BorderStyle Then
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &HB2ACA5, B
  End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
'  Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
'  Text1.maxLength = PropBag.ReadProperty("MaxLength", 0)
  PasswordChar = PropBag.ReadProperty("PasswordChar", "")
'  Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
'  Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
'  Text1.SelText = PropBag.ReadProperty("SelText", "")
  m_Text = PropBag.ReadProperty("Text", "Text1")
  MultiLine = PropBag.ReadProperty("MultiLine", m_def_MultiLine)
  m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
  m_Scrollbars = PropBag.ReadProperty("ScrollBars", m_def_ScrollBars)
  If Ambient.UserMode Then Call Initialize
End Sub

Private Sub UserControl_Resize()
  SetRect m_tRect, m_BorderStyle + 2, m_BorderStyle, ScaleWidth - m_BorderStyle - 2, ScaleHeight - m_BorderStyle
'  Text1.Move m_BorderStyle, m_BorderStyle, ScaleWidth - (m_BorderStyle * 2), ScaleHeight - (m_BorderStyle * 2)
'  Height = (Text1.Height + (m_BorderStyle * 2)) * Screen.TwipsPerPixelY
  SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1, 2, 2), True
  If m_hEditBox <> 0 Then Call MoveWindow(m_hEditBox, m_BorderStyle, m_BorderStyle, ScaleWidth - (m_BorderStyle * 2), ScaleHeight - (m_BorderStyle * 2), True)
End Sub

Private Sub UserControl_Terminate()
  On Error Resume Next
  
  If (m_hEditBox) Then
    '-- Stop subclassing and destroy all
    Call Subclass_StopAll
    Call pvDestroyFont
    Call pvDestroyEditBox
  End If
  
  On Error GoTo 0
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000008)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
'  Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
'  Call PropBag.WriteProperty("MaxLength", Text1.maxLength, 0)
  Call PropBag.WriteProperty("PasswordChar", m_PasswordChar, "")
'  Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
'  Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
'  Call PropBag.WriteProperty("SelText", Text1.SelText, "")
  Call PropBag.WriteProperty("Text", m_Text, "Text1")
  Call PropBag.WriteProperty("MultiLine", m_Multiline, m_def_MultiLine)
  Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
  Call PropBag.WriteProperty("ScrollBars", m_Scrollbars, m_def_ScrollBars)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,0,0
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
  MultiLine = m_Multiline
End Property

Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
'  If Ambient.UserMode Then Err.Raise 382
  m_Multiline = New_MultiLine
  PropertyChanged "MultiLine"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  m_BorderStyle = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get ScrollBars() As eScrollBar
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
  ScrollBars = m_Scrollbars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As eScrollBar)
  If Ambient.UserMode Then Err.Raise 382
  m_Scrollbars = New_ScrollBars
  PropertyChanged "ScrollBars"
End Property

Public Function Initialize() As Boolean
  If (m_bInitialized = False) Then
    Initialize = pvCreateEditBox()
    If (m_hEditBox) Then
      '-- Subclass UserControl (parent)
      Call Subclass_Start(UserControl.hWnd)
      Call Subclass_AddMsg(UserControl.hWnd, WM_MOUSEACTIVATE, MSG_AFTER)
      Call Subclass_AddMsg(UserControl.hWnd, WM_SETFOCUS, MSG_AFTER)
'      Call Subclass_AddMsg(UserControl.hwnd, WM_SIZE, MSG_AFTER)
      
      '-- Subclass EditBox (child)
      Call Subclass_Start(m_hEditBox)
      Call Subclass_AddMsg(m_hEditBox, WM_SIZE, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_KEYDOWN, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_CHAR, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_KEYUP, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_LBUTTONDOWN, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_RBUTTONDOWN, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_MBUTTONDOWN, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_MOUSEMOVE, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_LBUTTONUP, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_RBUTTONUP, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_MBUTTONUP, MSG_BEFORE)
      Call Subclass_AddMsg(m_hEditBox, WM_SETFOCUS, MSG_BEFORE)
      
      m_lX = -1
      m_lY = -1
      
      m_bInitialized = True
    End If
  End If
End Function

Private Function pvCreateEditBox() As Boolean
  Dim lExStyle As Long
  Dim lStyle As Long
  
  '-- Define window style
  lStyle = WS_CHILD Or WS_TABSTOP Or IIf(m_Multiline, ES_MULTILINE, 0) Or ES_AUTOHSCROLL Or ES_AUTOVSCROLL
  If (m_Scrollbars = vbBoth Or m_Scrollbars = Horizontal) And m_Multiline Then lStyle = lStyle Or WS_HSCROLL
  If (m_Scrollbars = vbBoth Or m_Scrollbars = Vertical) And m_Multiline Then lStyle = lStyle Or WS_VSCROLL
  If m_Scrollbars = vbVertical And m_Multiline Then lStyle = lStyle And Not ES_AUTOHSCROLL
  
'  lExStyle = WS_EX_CLIENTEDGE
  
  '-- Create EditBox window
  m_hEditBox = CreateWindowEx(lExStyle, "EDIT", m_Text, lStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0)
  
  '-- Success [?]
  If (m_hEditBox) Then
    Call ShowWindow(m_hEditBox, SW_SHOW)
    Set Font = UserControl.Font
    pvCreateEditBox = True
  End If
End Function

Private Sub pvDestroyEditBox()
  If (m_hEditBox) Then
    If (DestroyWindow(m_hEditBox)) Then
      m_hEditBox = 0
    End If
  End If
End Sub

Private Sub pvDestroyFont()
  If (m_hFont) Then
    If (DeleteObject(m_hFont)) Then
      m_hFont = 0
    End If
  End If
End Sub

Private Sub pvStdFontToLogFont(oStdFont As StdFont, uLogFont As LOGFONT)
  Dim lChar As Long
  With uLogFont
    For lChar = 1 To Len(oStdFont.Name)
      .lfFaceName(lChar - 1) = CByte(Asc(Mid$(oStdFont.Name, lChar, 1)))
    Next lChar
    .lfHeight = -MulDiv(oStdFont.Size, GetDeviceCaps(UserControl.hDC, LOGPIXELSY), 72)
    .lfItalic = oStdFont.Italic
    .lfWeight = IIf(oStdFont.Bold, FW_BOLD, FW_NORMAL)
    .lfUnderline = oStdFont.Underline
    .lfStrikeOut = oStdFont.Strikethrough
    .lfCharSet = oStdFont.Charset
  End With
End Sub

Private Sub pvGetClientCursorPos(X As Long, Y As Long)
  Dim uPt As POINTAPI
  
  Call GetCursorPos(uPt)
  Call ScreenToClient(m_hEditBox, uPt)
  X = uPt.X
  Y = uPt.Y
End Sub

Private Function pvButton(ByVal uMsg As Long) As Integer
  Select Case uMsg
  Case WM_LBUTTONDOWN, WM_LBUTTONUP
    pvButton = vbLeftButton
  Case WM_RBUTTONDOWN, WM_RBUTTONUP
    pvButton = vbRightButton
  Case WM_MBUTTONDOWN, WM_MBUTTONUP
    pvButton = vbMiddleButton
  Case WM_MOUSEMOVE
    Select Case True
    Case GetAsyncKeyState(vbKeyLButton) < 0
      pvButton = vbLeftButton
    Case GetAsyncKeyState(vbKeyRButton) < 0
      pvButton = vbRightButton
    Case GetAsyncKeyState(vbKeyMButton) < 0
      pvButton = vbMiddleButton
    End Select
  End Select
End Function

Private Function pvShiftState() As Integer
  Dim lS As Integer
  
  If (GetAsyncKeyState(vbKeyShift) < 0) Then
    lS = lS Or vbShiftMask
  End If
  If (GetAsyncKeyState(vbKeyMenu) < 0) Then
    lS = lS Or vbAltMask
  End If
  If (GetAsyncKeyState(vbKeyControl) < 0) Then
    lS = lS Or vbCtrlMask
  End If
  pvShiftState = lS
End Function

'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lhWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lhWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
'Private Sub Subclass_DelMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
''Parameters:
'  'lhWnd  - The handle of the window for which the uMsg is to be removed from the callback table
'  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
'  'When      - Whether the msg is to be removed from the before, after or both callback tables
'  With sc_aSubData(zIdx(lhWnd))
'    If When And eMsgWhen.MSG_BEFORE Then
'      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
'    End If
'    If When And eMsgWhen.MSG_AFTER Then
'      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
'    End If
'  End With
'End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lhWnd As Long) As Long
'Parameters:
  'lhWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Dim I                       As Long                                                   'Loop index
  Dim J                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sSubCode                As String                                                 'Subclass code string
Const PUB_CLASSES             As Long = 0                                               'The number of UserControl public classes
Const GMEM_FIXED              As Long = 0                                               'Fixed memory GlobalAlloc flag
Const PAGE_EXECUTE_READWRITE  As Long = &H40&                                           'Allow memory to execute without violating XP SP2 Data Execution Prevention
Const PATCH_01                As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02                As Long = 68                                              'Address of the previous WndProc
Const PATCH_03                As Long = 78                                              'Relative address of SetWindowsLong
Const PATCH_06                As Long = 116                                             'Address of the previous WndProc
Const PATCH_07                As Long = 121                                             'Relative address of CallWindowProc
Const PATCH_0A                As Long = 186                                             'Address of the owner object
Const FUNC_CWP                As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
Const FUNC_EBM                As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL                As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER                As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5                As String = "vba5"                                        'Location of the EbMode function if running VB5
Const MOD_VBA6                As String = "vba6"                                        'Location of the EbMode function if running VB6

'If it's the first time through here..
  If sc_aBuf(1) = 0 Then

'Build the hex pair subclass string
    sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90" & _
               Hex$(&HA4 + (PUB_CLASSES * 12)) & "070000C3"
    
'Convert the string from hex pairs to bytes and store in the machine code buffer
    I = 1
    Do While J < CODE_LEN
      J = J + 1
      sc_aBuf(J) = CByte("&H" & Mid$(sSubCode, I, 2))                                   'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      sc_aBuf(16) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_aBuf(17) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                        'Get the address of EbMode in vba6.dll
      If sc_pEbMode = 0 Then                                                            'Found?
        sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                      'VB5 perhaps
      End If
    End If
    
    Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                            'Patch the address of this object instance into the static machine code buffer
    
    sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                             'Get the address of the CallWindowsProc function
    sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                             'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lhWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .sCode = sc_aBuf
    .nAddrSub = StrPtr(.sCode)
    '.nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, I)           'Mark memory as executable
    'Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                           'Copy the machine code from the static byte array to the code array in sc_aSubData
    
    .hWnd = lhWnd                                                                       'Store the hWnd
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    
    Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                                     'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                                        'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                                        'Patch the relative address of the CallWindowProc api function
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    I = I - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lhWnd As Long)
'Parameters:
  'lhWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lhWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    'Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

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
'Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'  Dim nEntry As Long
'
'  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
'    nMsgCnt = 0                                                                         'Message count is now zero
'    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
'      nEntry = PATCH_05                                                                 'Patch the before table message count location
'    Else                                                                                'Else after
'      nEntry = PATCH_09                                                                 'Patch the after table message count location
'    End If
'    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
'  Else                                                                                  'Else deleteting a specific message
'    Do While nEntry < nMsgCnt                                                           'For each table entry
'      nEntry = nEntry + 1
'      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
'        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
'        Exit Do                                                                         'Bail
'      End If
'    Loop                                                                                'Next entry
'  End If
'End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lhWnd Then                                                             'If the hWnd of this element is the one we're looking for
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


