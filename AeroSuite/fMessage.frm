VERSION 5.00
Begin VB.Form fMessage 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AeroSuite.AeroBasicForm BasicForm1 
      Left            =   360
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.PictureBox pContainer 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   0
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   5400
      Begin VB.Label lContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Instruction"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   765
         TabIndex        =   5
         Top             =   615
         Width           =   4515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Instruction"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00993300&
         Height          =   315
         Left            =   765
         TabIndex        =   4
         Top             =   150
         Width           =   4455
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   150
         Picture         =   "fMessage.frx":000C
         Top             =   150
         Width           =   480
      End
   End
   Begin AeroSuite.AeroButton Button 
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AeroSuite.AeroButton Button 
      Height          =   345
      Index           =   2
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AeroSuite.AeroButton Button 
      Height          =   345
      Index           =   3
      Left            =   4080
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DADADA&
      X1              =   0
      X2              =   248
      Y1              =   88
      Y2              =   88
   End
End
Attribute VB_Name = "fMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MinHeight = 52

Public mButtons As eMsgBoxBtn
Public mIcon As eMsgBoxIcon
Public mResult  As eMsgBoxResult

Dim CX&, CY&

Private Sub Button_Click(Index As Integer)
  Select Case Button(Index).Tag
  Case "ok": mResult = bOK
  Case "cancel": mResult = bCancel
  Case "yes": mResult = bYes
  Case "no": mResult = bNo
  Case "abort": mResult = bAbort
  Case "retry": mResult = bRetry
  Case "ignore": mResult = bIgnore
  End Select
  Hide
End Sub

Private Sub Form_Load()
  Select Case mButtons
  Case bOKOnly
    Button(3).Visible = True
    Button(3).Cancel = True
    Button(3).Caption = "OK"
    Button(3).Tag = "ok"
  Case bOKcancel
    Button(3).Visible = True
    Button(3).Cancel = True
    Button(3).Caption = "Cancel"
    Button(3).Tag = "cancel"
    Button(2).Visible = True
    Button(2).Caption = "OK"
    Button(2).Tag = "ok"
  Case bYesNo
    Button(3).Visible = True
    Button(3).Cancel = True
    Button(3).Caption = "&No"
    Button(3).Tag = "no"
    Button(2).Visible = True
    Button(2).Caption = "&Yes"
    Button(2).Tag = "yes"
  Case bYesNoCancel
    Button(3).Visible = True
    Button(3).Cancel = True
    Button(3).Caption = "Cancel"
    Button(3).Tag = "cancel"
    Button(2).Visible = True
    Button(2).Caption = "&No"
    Button(2).Tag = "no"
    Button(1).Visible = True
    Button(1).Caption = "&Yes"
    Button(1).Tag = "yes"
  Case bRetryCancel
    Button(3).Visible = True
    Button(3).Cancel = True
    Button(3).Caption = "Cancel"
    Button(3).Tag = "cancel"
    Button(2).Visible = True
    Button(2).Caption = "&Retry"
    Button(2).Tag = "retry"
  Case bAbortRetryIgnore
    Button(3).Visible = True
    Button(3).Cancel = True
    Button(3).Caption = "&Ignore"
    Button(3).Tag = "ignore"
    Button(2).Visible = True
    Button(2).Caption = "&Retry"
    Button(2).Tag = "retry"
    Button(1).Visible = True
    Button(1).Caption = "&Abort"
    Button(1).Tag = "abort"
  End Select
  
  Select Case mIcon
  Case iNone
    Set imgIcon.Picture = LoadPicture
    lMain.Left = 10
  Case iWarning
    Set imgIcon.Picture = LoadResPicture("WARNICON", vbResBitmap)
  Case iError
    Set imgIcon.Picture = LoadResPicture("ERRORICON", vbResBitmap)
  Case iInformation
    Set imgIcon.Picture = LoadResPicture("INFOICON", vbResBitmap)
  Case iQuestion
    Set imgIcon.Picture = LoadResPicture("QUESICON", vbResBitmap)
  End Select
  
  CY = GetSystemMetrics(SM_CYCAPTION)
  CX = GetSystemMetrics(SM_CXFRAME)
  
  Call Form_Resize
End Sub

Private Sub Form_Resize()
  lContent.Move lMain.Left, lMain.Top + lMain.Height + 10
  If lContent <> "" Then
    pContainer.Height = lContent.Height + lMain.Height + 40
  Else
    pContainer.Height = lMain.Height + 40 'MinHeight
  End If
  
  Height = (pContainer.Height + CY + CX + CX + 52) * Screen.TwipsPerPixelY
  
  Button(3).Move ScaleWidth - Button(3).Width - 10, ScaleHeight - Button(3).Height - 10
  Button(2).Move Button(3).Left - Button(2).Width - 10, Button(3).Top
  Button(1).Move Button(2).Left - Button(1).Width - 10, Button(2).Top
End Sub

Private Sub lContent_Change()
  Form_Resize
End Sub

Private Sub lMain_Change()
  Form_Resize
End Sub

Private Sub pContainer_Resize()
  Line1.Y1 = pContainer.Height: Line1.Y2 = Line1.Y1
  Line1.X2 = ScaleWidth
End Sub
