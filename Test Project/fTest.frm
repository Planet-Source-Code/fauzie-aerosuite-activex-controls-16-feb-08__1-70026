VERSION 5.00
Object = "*\A..\AeroSuite\AeroSuite.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AeroSuite Controls Sample"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   3  'Windows Default
   Begin AeroSuite.AeroBasicForm AeroForm1 
      Left            =   3120
      Top             =   4320
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin AeroSuite.AeroButton bAbout 
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "About"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":0796
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":0AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":0E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1192
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":14E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":180A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":2206
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":255A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":27EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":2B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":2E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":31BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":350E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":3862
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":3B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":3E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":419E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1920
      Top             =   1920
   End
   Begin AeroSuite.AeroStatusBar AeroStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      Top             =   5295
      Width           =   6165
      _ExtentX        =   10874
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
      Style           =   1
      SimpleText      =   "Copyright © 2007-2008 by Fauzie's Software"
   End
   Begin AeroSuite.AeroTab AeroTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7858
      TabCount        =   6
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabCaption(3)   =   "Tab 3"
      TabCaption(4)   =   "Tab 4"
      TabCaption(5)   =   "Tab 5"
      ActiveTabBackEndColor=   16777215
      ActiveTabBackStartColor=   16777215
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ActiveTabForeColor=   0
      BackColor       =   16777215
      BottomRightInnerBorderColor=   10070188
      DisabledTabBackColor=   13355721
      DisabledTabForeColor=   10526880
      InActiveTabBackEndColor=   13619151
      InActiveTabBackStartColor=   15461355
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InActiveTabForeColor=   0
      OuterBorderColor=   9800841
      TabStyle        =   1
      TopLeftInnerBorderColor=   16777215
      UseMouseWheelScroll=   0   'False
      Begin AeroSuite.AeroTab AeroTab2 
         Height          =   3735
         Left            =   -49760
         TabIndex        =   29
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6588
         TabCount        =   2
         TabCaption(0)   =   "Tab 0"
         TabCaption(1)   =   "Tab 1"
         ActiveTabBackEndColor=   16777215
         ActiveTabBackStartColor=   16777215
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ActiveTabForeColor=   0
         BackColor       =   16777215
         DisabledTabBackColor=   13355721
         DisabledTabForeColor=   10526880
         InActiveTabBackEndColor=   13619151
         InActiveTabBackStartColor=   15461355
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         InActiveTabForeColor=   0
         OuterBorderColor=   9800841
         TabStripBackColor=   16777215
         TabStyle        =   1
         Begin AeroSuite.AeroTextBox tContent 
            Height          =   615
            Left            =   1800
            TabIndex        =   47
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1085
            BackColor       =   16777215
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   $"fTest.frx":44F2
            MultiLine       =   -1  'True
            ScrollBars      =   2
         End
         Begin AeroSuite.AeroGroupBox AeroGroupBox5 
            Height          =   615
            Left            =   -9760
            TabIndex        =   44
            Top             =   2880
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1085
            BorderColor     =   0
            BackColor       =   16777215
            BackColor2      =   0
            HeadColor1      =   0
            HeadColor2      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "InputBox Result"
            Begin AeroSuite.AeroTextBox tResult 
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   450
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
         End
         Begin AeroSuite.AeroButton bMsgBox 
            Height          =   375
            Left            =   3360
            TabIndex        =   39
            Top             =   3120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Show Dialog"
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
         Begin VB.ComboBox cboButton 
            Height          =   345
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   2280
            Width           =   3255
         End
         Begin VB.ComboBox cboIcon 
            Height          =   345
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1920
            Width           =   3255
         End
         Begin AeroSuite.AeroTextBox tTitle 
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Message Box Sample"
         End
         Begin AeroSuite.AeroTextBox tMainIns 
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   840
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Type The Main Instruction Here!"
         End
         Begin AeroSuite.AeroTextBox tTitle1 
            Height          =   255
            Left            =   -8200
            TabIndex        =   40
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Input Box Sample"
         End
         Begin AeroSuite.AeroButton bInputBox 
            Height          =   375
            Left            =   -6640
            TabIndex        =   43
            Top             =   3120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Show Dialog"
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
         Begin AeroSuite.AeroTextBox tDefault 
            Height          =   255
            Left            =   -8200
            TabIndex        =   14
            Top             =   1560
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Input Box Sample"
         End
         Begin AeroSuite.AeroTextBox tPrompt 
            Height          =   615
            Left            =   -8200
            TabIndex        =   48
            Top             =   840
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1085
            BackColor       =   16777215
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   $"fTest.frx":452F
            MultiLine       =   -1  'True
            ScrollBars      =   2
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Default Value :"
            Height          =   225
            Index           =   7
            Left            =   -9760
            TabIndex        =   15
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Message &Title :"
            Height          =   225
            Index           =   6
            Left            =   -9760
            TabIndex        =   42
            Top             =   480
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Prompt :"
            Height          =   225
            Index           =   5
            Left            =   -9760
            TabIndex        =   41
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Buttons :"
            Height          =   225
            Index           =   4
            Left            =   240
            TabIndex        =   37
            Top             =   2280
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Display &Icon :"
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   35
            Top             =   1920
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Message &Content :"
            Height          =   225
            Index           =   2
            Left            =   240
            TabIndex        =   34
            Top             =   1200
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Main Instruction :"
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Message &Title :"
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   1170
         End
      End
      Begin AeroSuite.AeroTextBox AeroTextBox1 
         Height          =   255
         Index           =   0
         Left            =   -9640
         TabIndex        =   26
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox1"
      End
      Begin AeroSuite.AeroGroupBox AeroGroupBox4 
         Height          =   2175
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3836
         BorderColor     =   0
         BackColor       =   16777215
         BackColor2      =   0
         HeadColor1      =   0
         HeadColor2      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Multiline"
         Begin AeroSuite.AeroOptionButton AeroOptionButton1 
            Height          =   495
            Left            =   240
            TabIndex        =   19
            Top             =   1560
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Align           =   0
            Caption         =   "AeroOptionButton1"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":456C
            Value           =   0   'False
         End
         Begin AeroSuite.AeroButton AeroButton4 
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
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
         Begin AeroSuite.AeroCheckBox AeroCheckBox5 
            Height          =   495
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            Align           =   0
            Caption         =   "AeroCheckBox5"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":4588
            Value           =   0
         End
      End
      Begin AeroSuite.AeroCheckBox AeroCheckBox4 
         Height          =   375
         Left            =   -26040
         TabIndex        =   11
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Align           =   0
         Caption         =   "Enabled"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         MousePointer    =   0
         MouseIcon       =   "fTest.frx":45A4
         Value           =   1
      End
      Begin AeroSuite.AeroGroupBox AeroGroupBox1 
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2566
         BorderColor     =   0
         BackColor       =   16777215
         BackColor2      =   0
         HeadColor1      =   0
         HeadColor2      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Regular Buttons"
         Begin AeroSuite.AeroButton AeroButton1 
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Caption         =   "Standard Button"
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
         Begin AeroSuite.AeroButton AeroButton2 
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Caption         =   "Button with picture"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   16711935
            PicNormal       =   "fTest.frx":45C0
            PicSizeH        =   16
            PicSizeW        =   16
         End
      End
      Begin AeroSuite.AeroGroupBox AeroGroupBox2 
         Height          =   1455
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2566
         BorderColor     =   0
         BackColor       =   16777215
         BackColor2      =   0
         HeadColor1      =   0
         HeadColor2      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CheckBoxes"
         Begin AeroSuite.AeroCheckBox AeroCheckBox1 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Align           =   0
            Caption         =   "Unchecked"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":4914
            Value           =   0
         End
         Begin AeroSuite.AeroCheckBox AeroCheckBox2 
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Align           =   0
            Caption         =   "Checked"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":4930
            Value           =   1
         End
         Begin AeroSuite.AeroCheckBox AeroCheckBox3 
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Align           =   0
            Caption         =   "Grayed"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":494C
            Value           =   2
         End
      End
      Begin AeroSuite.AeroGroupBox AeroGroupBox3 
         Height          =   1455
         Left            =   3000
         TabIndex        =   8
         Top             =   2040
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2566
         BorderColor     =   0
         BackColor       =   16777215
         BackColor2      =   0
         HeadColor1      =   0
         HeadColor2      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Option Buttons"
         Begin AeroSuite.AeroOptionButton AeroOptionButton2 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Align           =   0
            Caption         =   "Option 1"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":4968
            Value           =   0   'False
         End
         Begin AeroSuite.AeroOptionButton AeroOptionButton2 
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Align           =   0
            Caption         =   "Option 2"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":4984
            Value           =   0   'False
         End
         Begin AeroSuite.AeroOptionButton AeroOptionButton2 
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Align           =   0
            Caption         =   "Option 3"
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShowFocusRect   =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            MousePointer    =   0
            MouseIcon       =   "fTest.frx":49A0
            Value           =   0   'False
         End
      End
      Begin AeroSuite.AeroProgressBar AeroProgressBar1 
         Height          =   270
         Left            =   -39520
         Top             =   840
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   476
      End
      Begin AeroSuite.AeroProgressBar AeroProgressBar2 
         Height          =   270
         Left            =   -39520
         Top             =   1680
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   476
         Value           =   0
         MarqueeStyle    =   -1  'True
      End
      Begin AeroSuite.AeroScrollbar AeroHScrollbar1 
         Height          =   255
         Left            =   -29760
         Top             =   3360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         LargeChange     =   30
         Orientation     =   1
      End
      Begin AeroSuite.AeroScrollbar AeroVScrollbar1 
         Height          =   2895
         Left            =   -26880
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   5106
         LargeChange     =   30
      End
      Begin AeroSuite.AeroListBox AeroListBox1 
         Height          =   2760
         Left            =   -19760
         TabIndex        =   12
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4868
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ItemHeight      =   20
         ItemHeightAuto  =   0   'False
      End
      Begin AeroSuite.AeroListBox AeroListBox2 
         Height          =   2760
         Left            =   -17000
         TabIndex        =   23
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4868
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ItemHeight      =   18
      End
      Begin AeroSuite.AeroTextBox AeroTextBox1 
         Height          =   345
         Index           =   1
         Left            =   -9640
         TabIndex        =   27
         Top             =   1320
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox2"
      End
      Begin AeroSuite.AeroTextBox AeroTextBox1 
         Height          =   1095
         Index           =   2
         Left            =   -9640
         TabIndex        =   46
         Top             =   1920
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1931
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AeroTextBox3"
         MultiLine       =   -1  'True
         ScrollBars      =   3
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   $"fTest.frx":49BC
         Height          =   675
         Left            =   -9640
         TabIndex        =   28
         Top             =   3480
         Width           =   5160
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Without Icon:"
         Height          =   255
         Index           =   1
         Left            =   -17000
         TabIndex        =   25
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With Icon:"
         Height          =   255
         Index           =   0
         Left            =   -19760
         TabIndex        =   24
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marquee Style ProgressBar"
         Height          =   225
         Left            =   -39760
         TabIndex        =   10
         Top             =   1320
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard ProgressBar"
         Height          =   225
         Left            =   -39760
         TabIndex        =   9
         Top             =   480
         Width           =   1680
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AeroCheckBox4_Click()
  AeroVScrollbar1.Enabled = Abs(AeroCheckBox4.Value)
  AeroHScrollbar1.Enabled = Abs(AeroCheckBox4.Value)
End Sub

Private Sub bAbout_Click()
  bAbout.AboutBox
End Sub

Private Sub bInputBox_Click()
  tResult.Text = AeroForm1.ShowInputBox(tPrompt.Text, tDefault.Text, tTitle1.Text)
End Sub

Private Sub bMsgBox_Click()
  AeroForm1.ShowMsgBox tMainIns.Text, tContent.Text, cboButton.ListIndex, cboIcon.ListIndex, tTitle.Text
End Sub

Private Sub Form_Load()
  AeroTab1.TabCaption(0) = "Buttons"
  AeroTab1.TabCaption(1) = "TextBox"
  AeroTab1.TabCaption(2) = "ListBox"
  AeroTab1.TabCaption(3) = "ScrollBar"
  AeroTab1.TabCaption(4) = "ProgressBar"
  AeroTab1.TabCaption(5) = "DialogBoxes"
  AeroTab2.TabCaption(0) = "Message Box"
  AeroTab2.TabCaption(1) = "Input Box"
  
  AeroListBox1.SetImageList ImageList1
  Randomize
  For i = 0 To 500
    AeroListBox1.AddItem "This is item " & i, Int(Rnd * 20) + 1
    AeroListBox2.AddItem "This is item " & i
  Next
  AeroButton4.Caption = "This is a button" & vbCrLf & "with multiline text"
  AeroCheckBox5.Caption = "This is a checkbox" & vbCrLf & "with multiline text"
  AeroOptionButton1.Caption = "This is an optionbutton" & vbCrLf & "with multiline text"
  
  cboIcon.AddItem "iNone"
  cboIcon.AddItem "iWarning"
  cboIcon.AddItem "iError"
  cboIcon.AddItem "iInformation"
  cboIcon.AddItem "iQuestion"
  cboIcon.ListIndex = 3
  cboButton.AddItem "bOKOnly"
  cboButton.AddItem "bOKCancel"
  cboButton.AddItem "bYesNo"
  cboButton.AddItem "bYesNoCancel"
  cboButton.AddItem "bRetryCancel"
  cboButton.AddItem "bAbortRetryIgnore"
  cboButton.ListIndex = 0
End Sub

Private Sub Timer1_Timer()
  If AeroProgressBar2.Value < AeroProgressBar2.Max Then
    AeroProgressBar2.Value = AeroProgressBar2.Value + 1
  Else
    AeroProgressBar2.Value = 0
  End If
End Sub
