VERSION 5.00
Begin VB.UserControl AeroGroupBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   MaskColor       =   &H00C0C0C0&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AeroGroupBox.ctx":0000
End
Attribute VB_Name = "AeroGroupBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------
' AeroGroupBox ActiveX Control
'--------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'--------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'--------------------------------------------------------------------------

Option Explicit

Private Type tRGB
    R As Long
    G As Long
    b As Long
End Type

Public Enum eFrameAppearance
  aGroupBox = 0
  aFrame = 1
End Enum

Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private m_CRect As RECT

'Default Property Values:
Private Const m_def_Alignment = AlignLeft
Private Const m_def_Appearance = aGroupBox
Private Const m_def_Caption = ""
Private Const m_def_ForeColor = vbBlack

'Property Variables:
'Private WithEvents m_Font As StdFont
Private m_Alignment As AlignmentCts
Private m_Appearance As eFrameAppearance
Private m_Caption As String
Private m_BorderColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_BackColor2 As OLE_COLOR
Private m_HeadColor1 As OLE_COLOR
Private m_HeadColor2 As OLE_COLOR
Private m_HeaderHeight As Long

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607

Public Sub About()
Attribute About.VB_UserMemId = -552
  fAbout.Show vbModal
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  UserControl_Resize
End Sub

Private Sub UserControl_Paint()
  Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
  Dim m_Text As String
  On Error Resume Next
  m_Text = m_Caption
  With UserControl
    .Cls
    If m_Appearance = aGroupBox Then
      .ForeColor = vbWhite
      RoundRect UserControl.hDC, 1, (TextHeight("H") \ 2) + 1, ScaleWidth - 2, ScaleHeight - 1, 5, 5
      .ForeColor = RGB(213, 223, 229)
      RoundRect UserControl.hDC, 0, (TextHeight("H") \ 2), ScaleWidth - 1, ScaleHeight - 2, 5, 5
      .ForeColor = m_ForeColor
      If m_Text <> "" Then
        Select Case m_Alignment
        Case AlignLeft
          SetRect m_CRect, 6, 0, .TextWidth(m_Text) + 12, .TextHeight(m_Text)
        Case AlignCenter
          SetRect m_CRect, (.ScaleWidth - .TextWidth(m_Text) - 6) \ 2, 0, 0, .TextHeight(m_Text)
          SetRect m_CRect, m_CRect.Left, 0, .ScaleWidth - m_CRect.Left, m_CRect.Bottom
        Case AlignRight
          SetRect m_CRect, .ScaleWidth - .TextWidth(m_Text) - 12, 0, .ScaleWidth - 6, .TextHeight(m_Text)
        End Select
        UserControl.Line (m_CRect.Left, m_CRect.Top)-(m_CRect.Right, m_CRect.Bottom), m_BackColor, BF
        DrawText UserControl.hDC, m_Text, -1, m_CRect, DT_CENTER Or DT_VCENTER
      End If
    Else
      APILine 0, 0, ScaleWidth - 1, 0, TranslateColor(m_BorderColor)
      APILine 0, 0, 0, ScaleHeight - 1, TranslateColor(m_BorderColor)
      APILine 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, TranslateColor(m_BorderColor)
      APILine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, TranslateColor(m_BorderColor)
    
      APILine 0, m_HeaderHeight + 1, ScaleWidth - 1, m_HeaderHeight + 1, TranslateColor(m_BorderColor)
      
      DrawGradientEx TranslateColor(m_HeadColor1), TranslateColor(m_HeadColor2), 60, 1, 1, ScaleWidth - 1, m_HeaderHeight
    
      DrawGradientEx TranslateColor(m_BackColor), TranslateColor(m_BackColor2), 50, 1, m_HeaderHeight + 2, ScaleWidth - 1, ScaleHeight - 2
    
      .ForeColor = m_ForeColor
      If m_Text <> "" Then
        SetRect m_CRect, 6, 1, ScaleWidth - 6, m_HeaderHeight
        DrawText UserControl.hDC, m_Text, -1, m_CRect, m_Alignment Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
      End If
    End If
  End With
End Sub

Public Property Get Alignment() As AlignmentCts
  Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentCts)
  m_Alignment = New_Alignment
  PropertyChanged "Alignment"
  UserControl_Resize
End Property

Public Property Get Appearance() As eFrameAppearance
  Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eFrameAppearance)
  m_Appearance = New_Appearance
  PropertyChanged "Appearance"
  UserControl_Resize
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  m_BorderColor = New_BorderColor
  PropertyChanged "BorderColor"
  UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  m_BackColor = New_BackColor
  PropertyChanged "BackColor"
  UserControl.BackColor = m_BackColor
  UserControl_Resize
End Property

Public Property Get BackColor2() As OLE_COLOR
  BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor As OLE_COLOR)
  m_BackColor2 = New_BackColor
  PropertyChanged "BackColor2"
  UserControl_Resize
End Property

Public Property Get HeadColor1() As OLE_COLOR
  HeadColor1 = m_HeadColor1
End Property

Public Property Let HeadColor1(ByVal New_HeadColor As OLE_COLOR)
  m_HeadColor1 = New_HeadColor
  PropertyChanged "HeadColor1"
  UserControl_Resize
End Property

Public Property Get HeadColor2() As OLE_COLOR
  HeadColor2 = m_HeadColor2
End Property

Public Property Let HeadColor2(ByVal New_HeadColor As OLE_COLOR)
  m_HeadColor2 = New_HeadColor
  PropertyChanged "HeadColor2"
  UserControl_Resize
End Property

Public Property Get HeaderHeight() As Long
  HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal New_HeadColor As Long)
  m_HeaderHeight = New_HeadColor
  PropertyChanged "HeaderHeight"
  UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
  UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
  UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
  UserControl.Refresh
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute hWnd.VB_UserMemId = -515
  hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  m_Caption = New_Caption
  PropertyChanged "Caption"
  UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_Alignment = m_def_Alignment
  m_Appearance = m_def_Appearance
  Set UserControl.Font = Ambient.Font
  m_Caption = Ambient.DisplayName
  m_ForeColor = m_def_ForeColor
  m_BackColor = Ambient.BackColor
  m_BackColor2 = ShiftColorOXP(TranslateColor(m_BackColor), -100)
  m_BorderColor = ShiftColorOXP(TranslateColor(m_BackColor), -350)
  m_HeadColor1 = Ambient.BackColor
  m_HeadColor2 = ShiftColorOXP(TranslateColor(m_HeadColor1), -200)
  m_HeaderHeight = 22
  UserControl.BackColor = m_BackColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error Resume Next
  m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
  m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
  m_BorderColor = PropBag.ReadProperty("BorderColor")
  m_BackColor = PropBag.ReadProperty("BackColor")
  m_BackColor2 = PropBag.ReadProperty("BackColor2")
  m_HeadColor1 = PropBag.ReadProperty("HeadColor1")
  m_HeadColor2 = PropBag.ReadProperty("HeadColor2")
  m_HeaderHeight = PropBag.ReadProperty("HeaderHeight", 22)
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
  UserControl.BackColor = m_BackColor
End Sub

Private Sub UserControl_Show()
  UserControl_Paint
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
  Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
  Call PropBag.WriteProperty("BorderColor", m_BorderColor)
  Call PropBag.WriteProperty("BackColor", m_BackColor)
  Call PropBag.WriteProperty("BackColor2", m_BackColor2)
  Call PropBag.WriteProperty("HeadColor1", m_HeadColor1)
  Call PropBag.WriteProperty("HeadColor2", m_HeadColor2)
  Call PropBag.WriteProperty("HeaderHeight", m_HeaderHeight, 22)
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
  Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
End Sub

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
  Dim Red As Long, Blue As Long, Green As Long
  Dim Delta As Long
  
  Blue = ((theColor \ &H10000) Mod &H100)
  Green = ((theColor \ &H100) Mod &H100)
  Red = (theColor And &HFF)
  Delta = &HFF - Base
  
  Blue = Base + Blue * Delta \ &HFF
  Green = Base + Green * Delta \ &HFF
  Red = Base + Red * Delta \ &HFF
  
  If Red > 255 Then Red = 255
  If Green > 255 Then Green = 255
  If Blue > 255 Then Blue = 255
  
  ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

Private Sub APILine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)
  Dim pt As POINTAPI
  Dim hPen As Long, hPenOld As Long
  
  '   Convert the Color to RGB
  lColor = TranslateColor(lColor)
  '   Create a New Pen
  hPen = CreatePen(0, 1, lColor)
  '   Select the New Pen into the DC
  hPenOld = SelectObject(UserControl.hDC, hPen)
  '   Move to the Fist Point
  MoveToEx UserControl.hDC, X1, Y1, pt
  '   Draw the Line Segment
  LineTo UserControl.hDC, X2, Y2
  '   Select the Old Pen back into DC
  SelectObject UserControl.hDC, hPenOld
  '   Delete the New Pen
  DeleteObject hPen
End Sub

Private Function GetRGB(Color As Long) As tRGB
'  Returns the RGB color value of the specified color.
  GetRGB.R = Color And 255
  GetRGB.G = (Color \ 256) And 255
  GetRGB.b = (Color \ 65536) And 255
End Function

Private Sub DrawGradientEx( _
        StartColor As Long, _
        EndColor As Long, _
        Optional Center As Single = 50, _
        Optional ByVal X1 As Long, _
        Optional ByVal Y1 As Long, _
        Optional ByVal X2 As Long = -1, _
        Optional ByVal Y2 As Long = -1) ' Center in percent
'   Draw vertical gradient effect on the control on specified coordinates.
    
    If (X2 = -1) Then X2 = ScaleWidth - 1
    If (Y2 = -1) Then Y2 = ScaleHeight - 1
    
    X2 = TranslateNumber(X2, X1, X2)            ' X2 must not be < than X1
    Y2 = TranslateNumber(Y2, Y1, Y2)            ' Y2 must not be < than Y1
                                                '
    Dim Color As Long                           '
    Dim Step As Single                          '
                                                '
    Dim RGB1 As tRGB                            ' Start color
    Dim RGB2 As tRGB                            ' End color
    Dim RGB3 As tRGB                            ' Mid color
    Dim RGB4 As tRGB                            ' Gradient color
                                                '
    Center = TranslateNumber(Center, 0, 100)    ' Center should not exceed 0 to 100
                                                '
    RGB1 = GetRGB(StartColor)                   ' Get RGB color values
    RGB2 = GetRGB(EndColor)                     '
                                                '
    RGB3.R = RGB1.R + (RGB2.R - RGB1.R) * 0.5   ' Blend start and end color
    RGB3.G = RGB1.G + (RGB2.G - RGB1.G) * 0.5   '
    RGB3.b = RGB1.b + (RGB2.b - RGB1.b) * 0.5   '
                                                '
    Center = (Y2 - Y1 - 1) * Center / 100       ' Converts center in percent
    Center = (Y1 + Center)                      ' to actual pixel coordinate
                                                '
    If (Center = 0) Then Center = 1             ' Avoid errors
                                                '
    While (Y1 <= Y2)                            ' Draw from top to bottom
        If (Y1 <= Center) Then                  '
            Step = Y1 / Center                  '
            RGB4.R = RGB1.R + (RGB3.R - RGB1.R) * Step
            RGB4.G = RGB1.G + (RGB3.G - RGB1.G) * Step
            RGB4.b = RGB1.b + (RGB3.b - RGB1.b) * Step
        Else                                    '
            Step = (Y1 - Center) / (Y2 - Center)
            RGB4.R = RGB3.R + (RGB2.R - RGB3.R) * Step
            RGB4.G = RGB3.G + (RGB2.G - RGB3.G) * Step
            RGB4.b = RGB3.b + (RGB2.b - RGB3.b) * Step
        End If                                  '
                                                '
        Color = RGBEx(RGB4.R, RGB4.G, RGB4.b)   ' Prepare color
                                                '
        If (X1 = X2) Then                       '
            SetPixel hDC, X1, Y1, Color         ' Draw point to device
        Else                                    '
            APILine X1, Y1, X2, Y1, Color       ' Draw line to device
        End If                                  '
                                                '
        Y1 = Y1 + 1                             '
    Wend                                        '
    
End Sub

Private Function TranslateNumber( _
        ByVal Value As Single, _
        Minimum As Long, _
        Maximum As Long) As Single
'   Ensure a number does not exceed the specified limits.
  If (Value > Maximum) Then
    TranslateNumber = Maximum
  ElseIf (Value < Minimum) Then
    TranslateNumber = Minimum
  Else
    TranslateNumber = Value
  End If
End Function

Private Function RGBEx( _
        ByVal Red As Long, _
        ByVal Green As Long, _
        ByVal Blue As Long) As Long
' Returns a whole number representing an RGB color value.
  RGBEx = Red + 256& * Green + 65536 * Blue
End Function
