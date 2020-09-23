VERSION 5.00
Begin VB.UserControl AeroBasicForm 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "BasicForm.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   3090
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aero Form"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox pBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   600
      Picture         =   "BasicForm.ctx":0312
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox pBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   600
      Picture         =   "BasicForm.ctx":2AB4
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox pBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   600
      Picture         =   "BasicForm.ctx":5256
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox pBtn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   600
      Picture         =   "BasicForm.ctx":79F8
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.PictureBox pMenu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      Picture         =   "BasicForm.ctx":A19A
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox pBottom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      Picture         =   "BasicForm.ctx":ABF4
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox pRight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   1200
      Picture         =   "BasicForm.ctx":B036
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pLeft 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   720
      Picture         =   "BasicForm.ctx":C0F8
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pTop 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   0
      Picture         =   "BasicForm.ctx":D1BA
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "AeroBasicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------
' AeroTextBox ActiveX Control (completely re-written on Feb 14)
' Uses the LaVolpe Custom Window I (modified)
'--------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'--------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'--------------------------------------------------------------------------

Option Explicit

Public Enum eMsgBoxResult
  bOK = 0
  bCancel = 1
  bYes = 2
  bNo = 3
  bAbort = 4
  bRetry = 5
  bIgnore = 6
End Enum

Public Enum eMsgBoxBtn
  bOKOnly = 0
  bOKcancel = 1
  bYesNo = 2
  bYesNoCancel = 3
  bRetryCancel = 4
  bAbortRetryIgnore = 5
End Enum

Public Enum eMsgBoxIcon
  iNone = 0
  iWarning = 1
  iError = 2
  iInformation = 3
  iQuestion = 4
End Enum

Private Enum eBorderStyle
  None = 0
  Fixed = 1
  Sizable = 2
  Dialog = 3
End Enum

Private FBorderStyle As eBorderStyle, FMaxButton As Boolean, FMinButton As Boolean
Private pCapt As New c32bppDIB

' optional statement & used only if real-time overriding drawing is performed
Implements CustomWindowCalls

' required & must be initialized somewhere
Private LaVolpe_Window As clsCustomWndow

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Sub CustomWindowCalls_BeforeDrawCaption(ByVal hDC As Long, Caption As String, ByVal X As Long, ByVal Y As Long)
  Dim cRect As RECT, tCaption As String
  tCaption = Caption
  With Picture1
    .Width = .TextWidth(Caption): .Height = .TextHeight(Caption)
    SetRect cRect, 0, 0, .ScaleWidth, .ScaleHeight
    .Cls
    DrawText .hDC, tCaption, Len(tCaption), cRect, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER Or &H8000&
    pCapt.LoadPicture_StdPicture .Image
  End With
  pCapt.MakeTransparent vbWhite
  pCapt.RenderDropShadow_JIT hDC, X - 5, Y - 1, 5, vbWhite, 100
End Sub

Private Sub CustomWindowCalls_EnterExitSizing(ByVal BeginSizing As Boolean, UserRedrawn As Boolean)

End Sub

Private Sub CustomWindowCalls_OnCreateRegion(hRgn As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long)
  On Error Resume Next
  hRgn = CreateRoundRectRgn(0, 0, CX + 1, CY, 9, 9)
  Call CombineRgn(hRgn, hRgn, CreateRectRgn(0, 29, CX, CY), 2)
End Sub

Private Sub CustomWindowCalls_UserButtonClick(ByVal ID As String)

End Sub

Private Sub CustomWindowCalls_UserDrawnBorder(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)
  Dim State As Long
  State = Abs(HasFocus)
  BitBlt hDC, 0, 0, 8, 29, pTop.hDC, 0, 29 * State, vbSrcCopy '8, 29, ByVal 0&
  StretchBlt hDC, 8, 0, CX - 16, 29, pTop.hDC, 8, 29 * State, 5, 29, vbSrcCopy
  BitBlt hDC, CX - 8, 0, 8, 29, pTop.hDC, 13, 29 * State, vbSrcCopy '8, 29, ByVal 0&
  
  StretchBlt hDC, 0, 29, 8, CY - 29, pLeft.hDC, 8 * State, 0, 8, 88, vbSrcCopy
  
  StretchBlt hDC, CX - 8, 29, 8, CY - 29, pRight.hDC, 8 * State, 0, 8, 88, vbSrcCopy
  
  BitBlt hDC, 0, CY - 8, 8, 8, pBottom.hDC, 0, 8 * State, vbSrcCopy '8, 8, ByVal 0&
  StretchBlt hDC, 8, CY - 8, CX - 16, 8, pBottom.hDC, 8, 8 * State, 5, 8, vbSrcCopy
  BitBlt hDC, CX - 8, CY - 8, 8, 8, pBottom.hDC, 13, 8 * State, vbSrcCopy '8, 8, ByVal 0&
  
  Modified = True
End Sub

Private Sub CustomWindowCalls_UserDrawnButton(ByVal ID As String, ByVal State As Integer, ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)
  Dim Inset As Integer
  Inset = IIf(HasFocus, 0, 112)
'  Exit Sub
  Select Case ID
  Case 1
    BitBlt hDC, X, Y, 28, 15, pBtn(0).hDC, (State * 28) + Inset, 0, vbSrcCopy
'  Case SC_RESTORE
'    BitBlt hdc, X, Y, 28, 15, pBtn(2).hdc, (State * 28) + Inset, 0, vbSrcCopy
  Case 2
    BitBlt hDC, X, Y, 28, 15, pBtn(IIf(IsZoomed(hWnd), 2, 1)).hDC, (State * 28) + Inset, 0, vbSrcCopy
  Case 3
    BitBlt hDC, X, Y, 28, 15, pBtn(3).hDC, (State * 28) + Inset, 0, vbSrcCopy
  End Select
  Modified = True
End Sub

Private Sub CustomWindowCalls_UserDrawnMenuBar(ByVal hDC As Long, ByVal CX As Long, ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)
  StretchBlt hDC, 0, 0, CX, CY, pMenu.hDC, 0, 0, 45, 19, vbSrcCopy
  Modified = True
End Sub

Private Sub CustomWindowCalls_UserDrawnTitleBar(ByVal hDC As Long, ByVal CX As Long, ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)
  StretchBlt hDC, 0, 0, CX, CY, pTop.hDC, 8, (29 * Abs(HasFocus)) + 8, 5, 21, vbSrcCopy
  Modified = True
End Sub

Private Sub UserControl_Initialize()
  If LaVolpe_Window Is Nothing Then Set LaVolpe_Window = New clsCustomWndow
End Sub

Private Sub UserControl_InitProperties()
  Parent.BackColor = &HF0F0F0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  If Not LaVolpe_Window Is Nothing And Ambient.UserMode Then
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
    
    Dim CX&, wRect As RECT
    GetWindowRect UserControl.Parent.hWnd, wRect
    CX = GetSystemMetrics(SM_CXFRAME)
    Parent.Height = 100 * 15 '((wRect.Bottom - wRect.Top) + (16 - CX - CX)) * 15
    
    With LaVolpe_Window
      .BorderStyle = wbCustom
      Set .Font_TBar = UserControl.Font
      UserControl.FontBold = False
      Set .Font_MenuBar = UserControl.Font
      .FontColor_TBar(True) = vbBlack
      .FontColor_TBar(False) = RGB(64, 64, 64)
      .SetMenuSelectColors RGB(51, 153, 255), RGB(51, 153, 255)
      .FontColor_Menubar(fcDisabled) = RGB(128, 128, 128)
      .FontColor_Menubar(fcSelected) = vbWhite
      .MenuSelect_FlatStyle = True
      .ShowInTaskBar = Parent.ShowInTaskBar
      .EnableTBarBtns(smSize) = (FBorderStyle = Sizable)
      .EnableTBarBtns(smMaximize) = FMaxButton
      .EnableTBarBtns(smMinimize) = FMinButton
      .EnableTBarBtns(smSysIcon) = (FBorderStyle <> Dialog)
      If Not FMaxButton And Not FMinButton Then .HideDisabledButtons = True
      .BeginCustomWindow UserControl.Parent, Me
    End With
  
'    Dim CX&, wRect As RECT
'    GetWindowRect UserControl.Parent.hWnd, wRect
'    CX = GetSystemMetrics(SM_CXFRAME)
'    Parent.Height = ((wRect.Bottom - wRect.Top) + (16 - CX - CX)) * 15
'    Debug.Print SetWindowPos(UserControl.Parent.hWnd, 0, wRect.Left - (8 - CX), wRect.Top - (8 - CX), wRect.Right - wRect.Left + (16 - CX - CX), (wRect.Bottom - wRect.Top) + (16 - CX - CX), SWP_FRAMECHANGED)
  End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 48 * Screen.TwipsPerPixelX
    UserControl.Height = 48 * Screen.TwipsPerPixelY
    Command1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_Terminate()
  If Not LaVolpe_Window Is Nothing Then Set LaVolpe_Window = Nothing
End Sub

Public Function ShowMsgBox(ByVal MainText, Optional ByVal ContentText, Optional ByVal MsgBoxButton As eMsgBoxBtn, Optional ByVal MsgBoxIcon As eMsgBoxIcon, Optional ByVal Title = "") As eMsgBoxResult
  Dim fMsg As fMessage, mRes As eMsgBoxResult
  Set fMsg = New fMessage
  With fMsg
    .mButtons = MsgBoxButton
    .mIcon = MsgBoxIcon
    .lMain = MainText
    .lContent = ContentText
    If Title <> "" Then .Caption = Title
    .Show vbModal, UserControl.Parent
    ShowMsgBox = .mResult
    Unload fMsg
  End With
End Function

Public Function ShowInputBox(ByVal Prompt, Optional ByVal Default = "", Optional ByVal Title = "") As String
  Dim fMsg As fInput
  Set fMsg = New fInput
  With fMsg
    .lPrompt = Prompt
    .tInput.Text = Default
    If Title <> "" Then .Caption = Title
    .Show vbModal, UserControl.Parent
    If Not .Canceled Then ShowInputBox = .tInput.Text
    Unload fMsg
  End With
End Function

