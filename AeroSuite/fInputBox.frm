VERSION 5.00
Begin VB.Form fInput 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2010
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fInputBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AeroSuite.AeroBasicForm BasicForm1 
      Left            =   600
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin AeroSuite.AeroButton OKButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin AeroSuite.AeroTextBox tInput 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
   End
   Begin AeroSuite.AeroButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin VB.Label lPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Instruction"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean

Private Sub CancelButton_Click()
  Canceled = True
  Hide
End Sub

Private Sub lPrompt_Change()
  If lPrompt.Width > 273 Then lPrompt.Width = 273
End Sub

Private Sub OKButton_Click()
  Hide
End Sub
