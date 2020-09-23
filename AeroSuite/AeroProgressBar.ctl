VERSION 5.00
Begin VB.UserControl AeroProgressBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AeroProgressBar.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2160
      Top             =   2520
   End
End
Attribute VB_Name = "AeroProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
' AeroProgressBar ActiveX Control
'-----------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'-----------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'-----------------------------------------------------------------------------

Option Explicit

Dim PicBack As New pcMemDC, PicChunk As New pcMemDC
Dim Pos As Integer, VWidth As Integer
Dim PFlash As New c32bppDIB, PMarquee As New c32bppDIB

Const m_def_Max = 100
Const m_def_Value = 100

Dim m_Max As Long
Dim m_Value As Long
Dim m_Marquee As Boolean

Public Sub About()
Attribute About.VB_UserMemId = -552
  fAbout.Show vbModal
End Sub

Public Property Get MarqueeStyle() As Boolean
Attribute MarqueeStyle.VB_ProcData.VB_Invoke_Property = ";Misc"
  MarqueeStyle = m_Marquee
End Property

Public Property Let MarqueeStyle(New_Value As Boolean)
  m_Marquee = New_Value
  Call DrawProg
End Property

Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Misc"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
    DrawProg
End Property

Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    If New_Value > m_Max Then Err.Raise 380
    m_Value = New_Value
    PropertyChanged "Value"
    DrawProg
End Property

Private Sub DrawProg()
  On Error Resume Next
  Cls
  PicBack.Draw hdc, 0, 0, 1, 18, 0, 0
  PicBack.StretchDraw hdc, 1, 0, ScaleWidth - 2, 18, 1, 0, 2, 18
  PicBack.Draw hdc, 3, 0, 1, 18, ScaleWidth - 1, 0
  
  If Value > 0 Then
    If Not m_Marquee Then
      PicChunk.StretchDraw hdc, 1, 1, ScaleWidth / m_Max * m_Value - 2, 16, 0, 0, 50, 16
      PFlash.Render hdc, Pos, , IIf(Pos + 123 > VWidth, VWidth - Pos, 0), , , , IIf(Pos + 123 > VWidth, VWidth - Pos, 0)
    Else
      PMarquee.Render hdc, ((ScaleWidth + PMarquee.Width) / m_Max * m_Value - 2) - PMarquee.Width, 0
      PicBack.Draw hdc, 0, 0, 1, 18, 0, 0
      PicBack.Draw hdc, 3, 0, 1, 18, ScaleWidth - 1, 0
    End If
  End If
End Sub

Private Sub Timer1_Timer()
  VWidth = ScaleWidth / m_Max * m_Value - 1
  Pos = Pos + 10
  DrawProg
'  PFlash.Render hDC, Pos, , IIf(Pos + 123 > VWidth, VWidth - Pos, 0), , , , IIf(Pos + 123 > VWidth, VWidth - Pos, 0)
  Refresh
  If Pos > VWidth Then Pos = -(123 * 7)
End Sub

Private Sub UserControl_Initialize()
    PicBack.CreateFromPicture LoadResPicture("PROGBACK", vbResBitmap)
    PicChunk.CreateFromPicture LoadResPicture("PROGCHUNK", vbResBitmap)
    PFlash.LoadPicture_Resource "PROGFLASH", "PNG"
    PMarquee.LoadPicture_Resource "MARQUEE", "PNG"
    Pos = -123
End Sub

Private Sub UserControl_InitProperties()
    m_Max = m_def_Max
    m_Value = m_def_Value
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Marquee = PropBag.ReadProperty("MarqueeStyle", False)
    Timer1 = (Ambient.UserMode And Not m_Marquee)
    Pos = -123
    DrawProg
End Sub

Private Sub UserControl_Resize()
  On Error GoTo ErrTrap
  UserControl.Extender.Height = ScaleY(18, vbPixels, UserControl.Extender.Container.ScaleMode)
  DrawProg
  Exit Sub
ErrTrap:
  UserControl.Height = 18 * Screen.TwipsPerPixelY
  DrawProg
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("MarqueeStyle", m_Marquee, False)
End Sub
