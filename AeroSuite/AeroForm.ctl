VERSION 5.00
Begin VB.UserControl AeroForm 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AeroForm.ctx":0000
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   480
      Picture         =   "AeroForm.ctx":0312
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   3
      Top             =   2160
      Width           =   1500
   End
   Begin VB.PictureBox pCaptionButtons 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   1800
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   1320
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aero Form"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3090
   End
End
Attribute VB_Name = "AeroForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                                           'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                                          'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE                          'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                         'Subclass data type
    hwnd                   As Long                                            'Handle of the window being subclassed
    nAddrSub               As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig              As Long                                            'The address of the pre-existing WndProc
    nMsgCntA               As Long                                            'Msg after table entry count
    nMsgCntB               As Long                                            'Msg before table entry count
    aMsgTblA()             As Long                                            'Msg after table array
    aMsgTblB()             As Long                                            'Msg Before table array
End Type

Private sc_aSubData()      As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED   As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC  As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04     As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05     As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08     As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09     As Long = 137                                      'Table A (after) entry count patch offset

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'//

Private Enum eBorderStyle
  None = 0
  Fixed = 1
  Sizable = 2
  Dialog = 3
End Enum

Private Const LF_FACESIZE = 32
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
   lfFaceName(LF_FACESIZE) As Byte
End Type
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Const CaptionHeight = 25
Private Const LeftWidth = 9
Private Const RightWidth = 9
Private Const BottomHeight = 9

Private FBorderStyle As eBorderStyle, FMaxButton As Boolean, FMinButton As Boolean
Private WorkArea               As RECT

Private mDC As Long  ' Memory hDC
Private mainBitmap As Long ' Memory Bitmap
Private blendFunc32bpp As BLENDFUNCTION
Private Token As Long ' Needed to close GDI+
Private oldBitmap As Long

Private MainWnd As Long
Private MainRect As RECT

Private pIcon As New c32bppDIB
Private iFrame(3) As New c32bppDIB
Private bClose(1) As New c32bppDIB, bCloseS(1) As New c32bppDIB, bMaxRes(2) As New c32bppDIB, bMin(1) As New c32bppDIB
Private clH As Boolean, clHl As Boolean, clD As Boolean, clRct As RECT
Private mxH As Boolean, mxHl As Boolean, mxD As Boolean, mxRct As RECT
Private mnH As Boolean, mnHl As Boolean, mnD As Boolean, mnRct As RECT

Private R As RECT, CR As RECT, R1 As RECT
Private FormActive As Boolean
Private pCapt As New c32bppDIB, cRect As RECT

Private Sub pCaptionButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If clH Then
    clD = True
  ElseIf mxH Then
    mxD = True
  ElseIf mnH Then
    mnD = True
  End If
  UpdateButtons
End Sub

Private Sub pCaptionButtons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  clH = X > clRct.Left And X < clRct.Right And Y > clRct.Top And Y < clRct.Bottom
  If clH <> clHl Then UpdateButtons: clHl = clH: Exit Sub
  mxH = X > mxRct.Left And X < mxRct.Right And Y > mxRct.Top And Y < mxRct.Bottom And FMaxButton
  If mxH <> mxHl Then UpdateButtons: mxHl = mxH: Exit Sub
  mnH = X > mnRct.Left And X < mnRct.Right And Y > mnRct.Top And Y < mnRct.Bottom And FMinButton
  If mnH <> mnHl Then UpdateButtons: mnHl = mnH: Exit Sub
End Sub

Private Sub pCaptionButtons_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If clH Then
    SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0&
  ElseIf mxH Then
    If UserControl.Parent.WindowState = 0 Then
      SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0&
    Else
      SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_RESTORE, 0&
    End If
  ElseIf mnH Then
    If UserControl.Parent.WindowState <> 1 Then
      SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_MINIMIZE, 0&
    Else
      SendMessage UserControl.Parent.hwnd, WM_SYSCOMMAND, SC_RESTORE, 0&
    End If
  End If
  clH = False: mxH = False: mnH = False
  clD = False: mxD = False: mnD = False
  UpdateButtons
End Sub

Private Sub pCaptionButtons_Resize()
  With pCaptionButtons
    clRct.Left = .ScaleWidth - 10 - 44: clRct.Top = 14
    clRct.Right = clRct.Left + 44: clRct.Bottom = clRct.Top + 18
    mxRct.Left = .ScaleWidth - 10 - 44 - 26: mxRct.Top = 14
    mxRct.Right = mxRct.Left + 26: mxRct.Bottom = mxRct.Top + 18
    mnRct.Left = .ScaleWidth - 10 - 44 - 26 - 26: mnRct.Top = 14
    mnRct.Right = mnRct.Left + 26: mnRct.Bottom = mnRct.Top + 18
  End With
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lngReturnValue As Long
  If Button = 1 Then
    Call ReleaseCapture
    lngReturnValue = SendMessage(MainWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  If Ambient.UserMode Then
    Command1.Visible = False
    MainWnd = UserControl.Parent.hwnd
    
    Subclass_Start MainWnd
    Subclass_AddMsg MainWnd, WM_GETMINMAXINFO
    Subclass_AddMsg MainWnd, WM_SYSCOMMAND
    Subclass_AddMsg MainWnd, WM_MOVING
    Subclass_AddMsg MainWnd, WM_LBUTTONDOWN
    Subclass_AddMsg MainWnd, WM_SIZE
    Subclass_AddMsg MainWnd, WM_SHOWWINDOW
    Subclass_AddMsg MainWnd, WM_SETFOCUS
    Subclass_AddMsg MainWnd, MSM_NCACTIVATE
    Subclass_AddMsg MainWnd, WM_NCLBUTTONDOWN
    Subclass_AddMsg MainWnd, WM_PAINT
    Subclass_AddMsg MainWnd, WM_ACTIVATEAPP
    Subclass_AddMsg MainWnd, WM_CLOSE
    Subclass_AddMsg MainWnd, WM_DESTROY
    Subclass_AddMsg MainWnd, WM_KILLFOCUS
  
    If Not TypeOf UserControl.Parent Is MDIForm Then
      FBorderStyle = UserControl.Parent.BorderStyle
      If FBorderStyle = 4 Then FBorderStyle = Fixed
      If FBorderStyle = 5 Then FBorderStyle = Sizable
      
      FMaxButton = UserControl.Parent.MaxButton
      FMinButton = UserControl.Parent.MinButton
    Else
      FBorderStyle = Sizable
      FMaxButton = True
      FMinButton = True
    End If
    pIcon.LoadPicture_StdPicture UserControl.Parent.Icon
    iFrame(0).LoadPicture_File App.Path & "\Images\WindowFrameTop.png"
    iFrame(1).LoadPicture_File App.Path & "\Images\WindowFrameBottom.png"
    iFrame(2).LoadPicture_File App.Path & "\Images\WindowFrameLeft.png"
    iFrame(3).LoadPicture_File App.Path & "\Images\WindowFrameRight.png"
    bCloseS(0).LoadPicture_File App.Path & "\Images\CloseButtonSingle.png"
    bCloseS(1).LoadPicture_File App.Path & "\Images\close-s-glow.png"
    bClose(0).LoadPicture_File App.Path & "\Images\CloseButton.png"
    bClose(1).LoadPicture_File App.Path & "\Images\close-glow.png"
    bMaxRes(0).LoadPicture_File App.Path & "\Images\MaxButton.png"
    bMaxRes(1).LoadPicture_File App.Path & "\Images\max-glow.png"
    bMaxRes(2).LoadPicture_File App.Path & "\Images\ResButton.png"
    bMin(0).LoadPicture_File App.Path & "\Images\MinButton.png"
    bMin(1).LoadPicture_File App.Path & "\Images\min-glow.png"
    
'    RepaintWindow
  End If
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  If Ambient.UserMode Then
    Picture1.Move LeftWidth, 0, ScaleWidth - 28 - 75 - 68, CaptionHeight
    cRect.Right = Picture1.ScaleWidth: cRect.Bottom = Picture1.ScaleHeight
    pCaptionButtons_Resize
  Else
    UserControl.Width = 48 * Screen.TwipsPerPixelX
    UserControl.Height = 48 * Screen.TwipsPerPixelY
    Command1.Move 0, 0, ScaleWidth, ScaleHeight
  End If
End Sub

Private Sub UserControl_Show()
  On Error Resume Next
  If Not TypeOf UserControl.Parent Is MDIForm Then UserControl.Parent.BackColor = RGB(240, 240, 240)
End Sub

Private Sub UserControl_Terminate()
  On Error Resume Next
  Subclass_StopAll
  Call GdiplusShutdown(Token)
  SelectObject mDC, oldBitmap
  DeleteObject mainBitmap
  DeleteObject oldBitmap
  DeleteDC mDC
End Sub

Private Sub GetWorkArea()
    SystemParametersInfo 48&, 0&, WorkArea, 0&
End Sub

'========================================================================================
' UserControl subclass procedure
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, _
                          ByRef bHandled As Boolean, _
                          ByRef lReturn As Long, _
                          ByRef lhWnd As Long, _
                          ByRef uMsg As Long, _
                          ByRef wParam As Long, _
                          ByRef lParam As Long _
                          )
  On Error Resume Next
  Select Case lhWnd
  Case MainWnd
    Select Case uMsg
    Case WM_GETMINMAXINFO
        Dim MMI As MINMAXINFO, cy&, cx&
        cy = GetSystemMetrics(SM_CYCAPTION)
        cx = GetSystemMetrics(SM_CXFRAME)
        GetWorkArea

        CopyMemory MMI, ByVal lParam, LenB(MMI)

        'set the minmaxinfo data to the
        'minimum and maximum values set
        'by the option choice
        With MMI
            .ptMaxPosition.X = WorkArea.Left - cx + 8
            .ptMaxPosition.Y = WorkArea.Top - cy - cx + 24
            .ptMaxSize.X = (WorkArea.Right - WorkArea.Left) - .ptMaxPosition.X - cx '+ cX + cX - 16
            .ptMaxSize.Y = (WorkArea.Bottom - WorkArea.Top) - .ptMaxPosition.Y - cx '+ cX + cX - cY '- CaptionHeight
            .ptMinTrackSize.X = 200
            .ptMinTrackSize.Y = 100
        End With

        CopyMemory ByVal lParam, MMI, LenB(MMI)
    
    Case WM_NCLBUTTONDOWN
      Resize False, False, SWP_NOZORDER
    
    Case WM_SIZE
      Call Resize((uMsg = WM_SIZE), True, , True)
    
    Case WM_MOVING
      Call Resize((uMsg = WM_SIZE), True)
    
    Case WM_LBUTTONDOWN
      Call Resize((uMsg = WM_SIZE), False)
    
    Case WM_ACTIVATEAPP
      Select Case wParam
      Case WA_ACTIVE, WA_CLICKACTIVE
        FormActive = True
      Case WA_INACTIVE
        FormActive = False
      End Select
      Call Resize(True, False, , True)
      
    Case MSM_NCACTIVATE
      On Local Error Resume Next
      Select Case wParam
      Case WA_ACTIVE, WA_CLICKACTIVE
        FormActive = True
      Case WA_INACTIVE
        FormActive = False
      End Select
      Call Resize(True, False, , True)
    
    Case WM_SETFOCUS
      FormActive = True
      Call Resize(True, False, , False)
    
    Case WM_SHOWWINDOW
      Dim curWinLong As Long
      'Border
      curWinLong = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
      curWinLong = curWinLong Or WS_EX_TOOLWINDOW
      SetWindowLong UserControl.hwnd, GWL_EXSTYLE, curWinLong
      Call SetParent(UserControl.hwnd, GetParent(MainWnd))
      curWinLong = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
      curWinLong = curWinLong Or WS_EX_LAYERED
      SetWindowLong UserControl.hwnd, GWL_EXSTYLE, curWinLong
      
      'Caption Buttons
      curWinLong = GetWindowLong(pCaptionButtons.hwnd, GWL_EXSTYLE)
      curWinLong = curWinLong Or WS_EX_TOOLWINDOW
      SetWindowLong pCaptionButtons.hwnd, GWL_EXSTYLE, curWinLong
      Call SetParent(pCaptionButtons.hwnd, GetParent(MainWnd))
      curWinLong = GetWindowLong(pCaptionButtons.hwnd, GWL_EXSTYLE)
      curWinLong = curWinLong Or WS_EX_LAYERED
      SetWindowLong pCaptionButtons.hwnd, GWL_EXSTYLE, curWinLong
      Call Resize(True, True, , True)
    
    Case WM_CLOSE
    
    Case WM_DESTROY
    
    End Select
  End Select
End Sub

Public Sub Resize(SetWndRect As Boolean, SetPosition As Boolean, Optional lFlag As Long = SWP_FRAMECHANGED, Optional bRepaint As Boolean)
  Dim cy, lStyle, cx
  On Error Resume Next
  
  cy = GetSystemMetrics(SM_CYCAPTION)
  cx = GetSystemMetrics(SM_CXFRAME)
      
  GetWindowRect MainWnd, MainRect

  If SetWndRect = True Then
    Dim lRet As Long
    Dim GRET As Long
    
    ' form görünümü genel kesim
    lRet = CreateRoundRectRgn(cx, cy + cx, (MainRect.Right - MainRect.Left) - cx, (MainRect.Bottom - MainRect.Top) - cx, 0, 0)
    Call SetWindowRgn(MainWnd, lRet, True)
    Call DeleteObject(lRet)
  End If
  If SetPosition = True Then
    Call SetWindowPos(UserControl.hwnd, MainWnd, MainRect.Left + cx - LeftWidth, (MainRect.Top) + cy + cx - CaptionHeight, (MainRect.Right - MainRect.Left) - cx - cx + LeftWidth + RightWidth - 1, (MainRect.Bottom - MainRect.Top) - cy - cx - cx + CaptionHeight + BottomHeight - 1, lFlag)
    Call SetWindowPos(pCaptionButtons.hwnd, UserControl.hwnd, MainRect.Left + (MainRect.Right - MainRect.Left) + cx + cx - 122, (MainRect.Top) + cy + cx - CaptionHeight - 13, 120, 38, lFlag)
  End If
  If bRepaint Then RepaintWindow
  Call UserControl_Resize
End Sub

Private Sub RepaintWindow()
   Dim tempBI As BITMAPINFO
   Dim winSize As Size
   Dim srcPoint As POINTAPI
   
   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32    ' Each pixel is 32 bit's wide
      .biHeight = UserControl.ScaleHeight  ' Height of the form
      .biWidth = UserControl.ScaleWidth    ' Width of the form
      .biPlanes = 1   ' Always set to 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
   End With
   mainBitmap = CreateDIBSection(hdc, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
   oldBitmap = SelectObject(hdc, mainBitmap)   ' Select the new bitmap, track the old that was selected
   
  If FormActive Then 'Or GetActiveWindow = hwnd Or GetActiveWindow = UserControl.Parent.hwnd Then
    ' Top
    iFrame(0).Render hdc, 0, 0, LeftWidth, CaptionHeight, 0, 0, LeftWidth, CaptionHeight
    iFrame(0).Render hdc, LeftWidth, 0, ScaleWidth - LeftWidth - RightWidth, CaptionHeight, LeftWidth, 0, iFrame(0).Width - LeftWidth - RightWidth, CaptionHeight
    iFrame(0).Render hdc, ScaleWidth - RightWidth, 0, RightWidth, CaptionHeight, iFrame(0).Width - RightWidth, 0, RightWidth, CaptionHeight
    
    ' Bottom
    iFrame(1).Render hdc, 0, ScaleHeight - BottomHeight, LeftWidth, BottomHeight, 0, 0, LeftWidth, BottomHeight
    iFrame(1).Render hdc, LeftWidth, ScaleHeight - BottomHeight, ScaleWidth - LeftWidth - RightWidth, BottomHeight, LeftWidth, 0, iFrame(1).Width - LeftWidth - RightWidth, BottomHeight
    iFrame(1).Render hdc, ScaleWidth - RightWidth, ScaleHeight - BottomHeight, RightWidth, BottomHeight, iFrame(1).Width - RightWidth, 0, RightWidth, BottomHeight
    
    ' Left
    iFrame(2).Render hdc, 0, CaptionHeight, LeftWidth, ScaleHeight - CaptionHeight - BottomHeight, 0, 0, LeftWidth
    
    ' Right
    iFrame(3).Render hdc, ScaleWidth - RightWidth, CaptionHeight, RightWidth, ScaleHeight - CaptionHeight - BottomHeight, 0, 0, RightWidth
  Else
    ' Top
    iFrame(0).Render hdc, 0, 0, LeftWidth, CaptionHeight, 0, CaptionHeight, LeftWidth, CaptionHeight
    iFrame(0).Render hdc, LeftWidth, 0, ScaleWidth - LeftWidth - RightWidth, CaptionHeight, LeftWidth, CaptionHeight, iFrame(0).Width - LeftWidth - RightWidth, CaptionHeight
    iFrame(0).Render hdc, ScaleWidth - RightWidth, 0, RightWidth, CaptionHeight, iFrame(0).Width - RightWidth, CaptionHeight, RightWidth, CaptionHeight
    
    ' Bottom
    iFrame(1).Render hdc, 0, ScaleHeight - BottomHeight, LeftWidth, BottomHeight, 0, BottomHeight, LeftWidth, BottomHeight
    iFrame(1).Render hdc, LeftWidth, ScaleHeight - BottomHeight, ScaleWidth - LeftWidth - RightWidth, BottomHeight, LeftWidth, BottomHeight, iFrame(1).Width - LeftWidth - RightWidth, BottomHeight
    iFrame(1).Render hdc, ScaleWidth - RightWidth, ScaleHeight - BottomHeight, RightWidth, BottomHeight, iFrame(1).Width - RightWidth, BottomHeight, RightWidth, BottomHeight
    
    ' Left
    iFrame(2).Render hdc, 0, CaptionHeight, LeftWidth, ScaleHeight - CaptionHeight - BottomHeight, LeftWidth, 0, LeftWidth
    
    ' Right
    iFrame(3).Render hdc, ScaleWidth - RightWidth, CaptionHeight, RightWidth, ScaleHeight - CaptionHeight - BottomHeight, RightWidth, 0, RightWidth
  End If
   
  ' Icon
  If FBorderStyle <> Dialog Then
    pIcon.Render hdc, 10, 5, 16, 16
  End If
  
  ' Caption
  Picture1.Cls
  Picture1.ForeColor = IIf(FormActive, vbBlack, RGB(64, 64, 64))
  DrawText Picture1.hdc, UserControl.Parent.Caption, Len(UserControl.Parent.Caption), cRect, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER Or DT_END_ELLIPSIS Or DT_WORDBREAK
  pCapt.LoadPicture_StdPicture Picture1.Image
  pCapt.MakeTransparent vbWhite
  pCapt.RenderDropShadow_JIT hdc, IIf(FBorderStyle <> Dialog, 22, 2), -6, 7, vbWhite, 100 ', 100
  pCapt.RenderDropShadow_JIT hdc, IIf(FBorderStyle <> Dialog, 22, 2), -6, 7, vbWhite, 100 ', 100
  pCapt.Render hdc, IIf(FBorderStyle <> Dialog, 30, 10), 0

  ' Needed for updateLayeredWindow call
  srcPoint.X = 0
  srcPoint.Y = 0
  winSize.cx = UserControl.ScaleWidth
  winSize.cy = UserControl.ScaleHeight
  
  With blendFunc32bpp
    .AlphaFormat = AC_SRC_ALPHA ' 32 bit
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER
    .SourceConstantAlpha = 255
  End With
  
  Call UpdateLayeredWindow(UserControl.hwnd, UserControl.hdc, ByVal 0&, winSize, hdc, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
  
  SelectObject hdc, oldBitmap
  DeleteObject mainBitmap
  DeleteObject oldBitmap
  
  UpdateButtons
End Sub

Private Sub UpdateButtons()
   Dim tempBI As BITMAPINFO
   Dim winSize As Size
   Dim srcPoint As POINTAPI
   Dim mainBitmap2 As Long, oldBitmap2 As Long
   Debug.Print "UpdateButtons"; Rnd * 1000
   With tempBI.bmiHeader
      .biSize = Len(tempBI.bmiHeader)
      .biBitCount = 32    ' Each pixel is 32 bit's wide
      .biHeight = pCaptionButtons.ScaleHeight  ' Height of the form
      .biWidth = pCaptionButtons.ScaleWidth    ' Width of the form
      .biPlanes = 1   ' Always set to 1
      .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
   End With
   With pCaptionButtons
     mainBitmap2 = CreateDIBSection(.hdc, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
     oldBitmap2 = SelectObject(.hdc, mainBitmap2)   ' Select the new bitmap, track the old that was selected
     
    If FormActive Then
      ' Close Button
      If FMaxButton Or FMinButton Then
        bClose(0).Render .hdc, .ScaleWidth - 10 - 44, 14, 44, 0, Val(IIf(clH, IIf(clD, 88, 44), 0)), 0, 44
      ElseIf Not FMaxButton And Not FMinButton Then
        bCloseS(0).Render .hdc, .ScaleWidth - 10 - 44, 14, 44, 0, Val(IIf(clH, IIf(clD, 88, 44), 0)), 0, 44
      End If
    
      ' Max Button
      If FMaxButton Then
        bMaxRes(0).Render .hdc, .ScaleWidth - 10 - 44 - 26, 14, 26, 0, Val(IIf(mxH, IIf(mxD, 52, 26), 0)), 0, 26
      Else
        If FMinButton Then bMaxRes(0).Render .hdc, .ScaleWidth - 10 - 44 - 26, 14, 26, 0, 78, 0, 26
      End If
      
      ' Min Button
      If FMinButton Then
        bMin(0).Render .hdc, .ScaleWidth - 10 - 44 - 26 - 26, 14, 26, 0, Val(IIf(mnH, IIf(mnD, 52, 26), 0)), 0, 26
      Else
        If FMaxButton Then bMin(0).Render .hdc, .ScaleWidth - 10 - 44 - 26 - 26, 14, 26, 0, 78, 0, 26
      End If
    Else
      ' Close Button
      If FMaxButton Or FMinButton Then
        bClose(0).Render .hdc, .ScaleWidth - 10 - 44, 14, 44, 0, Val(IIf(clH, IIf(clD, 264, 220), 176)), 0, 44
      ElseIf Not FMaxButton And Not FMinButton Then
        bCloseS(0).Render .hdc, .ScaleWidth - 10 - 44, 14, 44, 0, Val(IIf(clH, IIf(clD, 264, 220), 176)), 0, 44
      End If
    
      ' Max Button
      If FMaxButton Then
        bMaxRes(0).Render .hdc, .ScaleWidth - 10 - 44 - 26, 14, 26, 0, Val(IIf(mxH, IIf(mxD, 156, 130), 104)), 0, 26
      Else
        If FMinButton Then bMaxRes(0).Render .hdc, .ScaleWidth - 10 - 44 - 26, 14, 26, 0, 182, 0, 26
      End If
    
      ' Min Button
      If FMinButton Then
        bMin(0).Render .hdc, .ScaleWidth - 10 - 44 - 26 - 26, 14, 26, 0, Val(IIf(mnH, IIf(mnD, 156, 130), 104)), 0, 26
      Else
        If FMaxButton Then bMin(0).Render .hdc, .ScaleWidth - 10 - 44 - 26 - 26, 14, 26, 0, 182, 0, 26
      End If
    End If
     
    ' Glow
    If clH Then
      If FMaxButton Or FMinButton Then
        bClose(1).Render .hdc, .ScaleWidth - 10 - 52, 4
      ElseIf Not FMaxButton And Not FMinButton Then
        bCloseS(1).Render .hdc, .ScaleWidth - 10 - 52, 4
      End If
    ElseIf mxH Then
      bMaxRes(1).Render .hdc, .ScaleWidth - 10 - 54 - 26, 4
    ElseIf mnH Then
      bMin(1).Render .hdc, .ScaleWidth - 10 - 55 - 26 - 26, 4
    End If
  
    ' Needed for updateLayeredWindow call
    srcPoint.X = 0
    srcPoint.Y = 0
    winSize.cx = .ScaleWidth
    winSize.cy = .ScaleHeight
    
    With blendFunc32bpp
      .AlphaFormat = AC_SRC_ALPHA ' 32 bit
      .BlendFlags = 0
      .BlendOp = AC_SRC_OVER
      .SourceConstantAlpha = 255
    End With
    
    Call UpdateLayeredWindow(.hwnd, 0, ByVal 0&, winSize, .hdc, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
    
    SelectObject .hdc, oldBitmap2
    DeleteObject mainBitmap2
    DeleteObject oldBitmap2
  End With
End Sub

Private Sub pOLEFontToLogFont(fntThis As StdFont, ByVal hdc As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer
Dim B() As Byte

   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
     sFont = fntThis.Name
     B = StrConv(sFont, vbFromUnicode)
     For iChar = 1 To Len(sFont)
       .lfFaceName(iChar - 1) = B(iChar - 1)
     Next iChar
     ' Based on the Win32SDK documentation:
     .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
     .lfItalic = fntThis.Italic
     If (fntThis.Bold) Then
       .lfWeight = FW_BOLD
     Else
       .lfWeight = FW_NORMAL
     End If
     .lfUnderline = fntThis.Underline
     .lfStrikeOut = fntThis.Strikethrough
     .lfCharSet = fntThis.Charset
   End With

End Sub

Public Function ShowMsgBox(ByVal MainText, Optional ByVal ContentText, Optional ByVal MsgBoxButton As eMsgBoxBtn, Optional ByVal MsgBoxIcon As eMsgBoxIcon, Optional ByVal Title = "") As eMsgBoxResult
  Dim fMsg As fMessage, mRes As eMsgBoxResult
  Set fMsg = New fMessage
  With fMsg
    .mButtons = MsgBoxButton
    .mIcon = MsgBoxIcon
    .lMain.Caption = MainText
    .lContent.Caption = ContentText
    If Title <> "" Then .Caption = Title
    .Show vbModal, UserControl.Parent
    ShowMsgBox = .mResult
    Unload fMsg
  End With
End Function

'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  
    With sc_aSubData(zIdx(lhWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lhWnd As Long) As Long

  Const CODE_LEN              As Long = 202
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
  
    If (aBuf(1) = 0) Then
  
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
    
        If (Subclass_InIDE) Then
            aBuf(16) = &H90
            aBuf(17) = &H90
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (pEbMode = 0) Then
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If
    
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
      Else
        nSubIdx = zIdx(lhWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lhWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
    End With
End Function

Private Sub Subclass_Stop(ByVal lhWnd As Long)
  
    With sc_aSubData(zIdx(lhWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hwnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB()
        Erase .aMsgTblA()
    End With
End Sub

Private Sub Subclass_StopAll()
  
  Dim i As Long
  
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If (.hwnd <> 0) Then
                Call Subclass_Stop(.hwnd)
            End If
        End With
        i = i - 1
    Loop
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  
  Dim nEntry  As Long
  Dim nOff1   As Long
  Dim nOff2   As Long
  
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGES
      Else
        Do While nEntry < nMsgCnt
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

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Function zIdx(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hwnd = lhWnd) Then
                If (Not bAdd) Then
                    Exit Function
                End If
            ElseIf (.hwnd = 0) Then
                If (bAdd) Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False
    End If
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

