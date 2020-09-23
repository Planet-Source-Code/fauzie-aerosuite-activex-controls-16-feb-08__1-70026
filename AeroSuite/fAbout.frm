VERSION 5.00
Begin VB.Form fAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AeroSuite ActiveX Controls"
   ClientHeight    =   3210
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5640
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fAbout.frx":000C
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AeroSuite.AeroBasicForm BasicForm1 
      Left            =   360
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin AeroSuite.AeroButton cmdOK 
      Height          =   345
      Left            =   4140
      TabIndex        =   3
      Top             =   2550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   "OK"
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
   Begin VB.Label LbEdition 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Free Source Edition"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   210
      Left            =   3285
      TabIndex        =   4
      Top             =   840
      Width           =   1770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DADADA&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   368
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00F0F0F0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2007-2008 by Fauzie's Software. All Rights Reserved."
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   690
      TabIndex        =   0
      Top             =   1485
      Width           =   3840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00F0F0F0&
      BackStyle       =   0  'Transparent
      Caption         =   "AeroSuite ActiveX Controls"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   690
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   368
      Y1              =   161
      Y2              =   161
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00F0F0F0&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.1"
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
      Left            =   4470
      TabIndex        =   2
      Top             =   630
      Width           =   300
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub
