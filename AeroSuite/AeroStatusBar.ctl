VERSION 5.00
Begin VB.UserControl AeroStatusBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   270
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
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AeroStatusBar.ctx":0000
End
Attribute VB_Name = "AeroStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
' AeroStatusBar ActiveX Control
'-----------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'-----------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'-----------------------------------------------------------------------------

Option Explicit

Public Enum SbarStyleConstants
  sbrNormal
  sbrSimple
End Enum

Private Type tRGB
    R As Long
    G As Long
    B As Long
End Type

Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private pGrip As New pcMemDC, rctGrip As RECT, m_OnGrip As Boolean

'Default Property Values:
Const m_def_Style = 0
Const m_def_ShowTips = 1
'Property Variables:
Dim m_Style As SbarStyleConstants
Dim m_SimpleText As String
Dim m_ShowTips As Boolean
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Public Sub About()
Attribute About.VB_UserMemId = -552
  fAbout.Show vbModal
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
  DrawStatusBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
  DrawStatusBar
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
  DrawStatusBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
  UserControl.Refresh
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get Style() As SbarStyleConstants
Attribute Style.VB_Description = "Returns/sets the single (simple) or multiple panel style."
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Behavior"
  Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As SbarStyleConstants)
  m_Style = New_Style
  PropertyChanged "Style"
  DrawStatusBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get SimpleText() As String
Attribute SimpleText.VB_Description = "Returns/sets the text displayed when a StatusBar control's Style property is set to Simple."
Attribute SimpleText.VB_ProcData.VB_Invoke_Property = ";Misc"
  SimpleText = m_SimpleText
End Property

Public Property Let SimpleText(ByVal New_SimpleText As String)
  m_SimpleText = New_SimpleText
  PropertyChanged "SimpleText"
  DrawStatusBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ShowTips() As Boolean
Attribute ShowTips.VB_Description = "Enables/disables ToolTips for panels."
Attribute ShowTips.VB_ProcData.VB_Invoke_Property = ";Behavior"
  ShowTips = m_ShowTips
End Property

Public Property Let ShowTips(ByVal New_ShowTips As Boolean)
  m_ShowTips = New_ShowTips
  PropertyChanged "ShowTips"
End Property

Private Sub DrawStatusBar()
  Call DrawBackground
  Call DrawBar
  UserControl.Refresh
End Sub

Private Sub DrawBackground()
  On Error Resume Next
  With UserControl
    .Cls
    .BackColor = RGB(244, 244, 244)
    DrawGradientEx RGB(215, 215, 215), RGB(244, 244, 244), 50, 0, 0, .ScaleWidth, 6
    If UserControl.Parent.BorderStyle = vbSizable And UserControl.Parent.WindowState = vbNormal _
        And .Extender.Width = .Parent.ScaleWidth And .Extender.Top + .Extender.Height = .Parent.ScaleHeight Then
      SetRect rctGrip, .ScaleWidth - 15, .ScaleHeight - 15, .ScaleWidth - 1, .ScaleHeight - 1
      pGrip.Draw hdc, 0, 0, 17, 17, .ScaleWidth - 18, .ScaleHeight - 18, True
    Else
      SetRectEmpty rctGrip
    End If
  End With
End Sub

Private Sub DrawBar()
  Dim rBar As RECT, sText As String
  If m_Style = sbrNormal Then
    
  Else
    sText = m_SimpleText
    SetRect rBar, 5, 0, UserControl.ScaleWidth - (rctGrip.Right - rctGrip.Left), UserControl.ScaleHeight
    DrawText UserControl.hdc, sText, Len(sText), rBar, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
  End If
End Sub

Private Function GetRGB(Color As Long) As tRGB
'   Returns the RGB color value of the specified color.
    GetRGB.R = Color And 255
    GetRGB.G = (Color \ 256) And 255
    GetRGB.B = (Color \ 65536) And 255
    
End Function

Private Sub DrawLine(Color As Long, X1 As Long, Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
'   Draw a line with the specified color and coordinates
    Dim hOld As Long
    Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 1, Color)
        hOld = SelectObject(hdc, hPen)
    
    If (X1 = X2) Then Y2 = Y2 + 1   ' LineTo draws a line up to, but not
    If (Y1 = Y2) Then X2 = X2 + 1   ' including the defined point. But not now!
                                    '
    ' If (X1 = X2) Then             '
    '     If (Y2 >= Y1) Then Y2 = Y2 + 1 Else Y2 = Y2 - 1
    ' End If                        '
    ' If (Y1 = Y2) Then             '
    '     If (X2 >= X1) Then X2 = X2 + 1 Else X2 = X2 - 1
    ' End If                        '
                                    '
    MoveToEx hdc, X1, Y1, 0&        ' Set starting position of line
    LineTo hdc, X2, Y2              ' then draw a line to this point
                                    '
    SelectObject hdc, hOld          ' Restore the previous pen
    DeleteObject hPen               ' then clear pen object from memory
    
End Sub

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
    RGB3.B = RGB1.B + (RGB2.B - RGB1.B) * 0.5   '
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
            RGB4.B = RGB1.B + (RGB3.B - RGB1.B) * Step
        Else                                    '
            Step = (Y1 - Center) / (Y2 - Center)
            RGB4.R = RGB3.R + (RGB2.R - RGB3.R) * Step
            RGB4.G = RGB3.G + (RGB2.G - RGB3.G) * Step
            RGB4.B = RGB3.B + (RGB2.B - RGB3.B) * Step
        End If                                  '
                                                '
        Color = RGB(RGB4.R, RGB4.G, RGB4.B)   ' Prepare color
                                                '
        If (X1 = X2) Then                       '
            SetPixelV hdc, X1, Y1, Color        ' Draw point to device
        Else                                    '
            DrawLine Color, X1, Y1, X2, Y1      ' Draw line to device
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

Private Sub UserControl_AmbientChanged(PropertyName As String)
  DrawStatusBar
End Sub

Private Sub UserControl_Initialize()
  pGrip.CreateFromPicture LoadResPicture("RESGRIP", vbResBitmap)
End Sub

Private Sub UserControl_InitProperties()
  With UserControl
    .Extender.Align = vbAlignBottom
    .Extender.Height = ScaleY(23, vbPixels, .Extender.Container.ScaleMode)
  End With
  Set UserControl.Font = Ambient.Font
  m_Style = m_def_Style
  m_SimpleText = Ambient.DisplayName
  m_ShowTips = m_def_ShowTips
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
  If Button = 1 And m_OnGrip Then
    Call ReleaseCapture
    Call SendMessage(UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&)
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
  If X > rctGrip.Left And Y > rctGrip.Top And X < rctGrip.Right And Y < rctGrip.Bottom Then
    MousePointer = vbSizeNWSE
    If Not m_OnGrip Then m_OnGrip = True
  Else
    MousePointer = vbArrow
    If m_OnGrip Then m_OnGrip = False
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  m_Style = PropBag.ReadProperty("Style", m_def_Style)
  m_SimpleText = PropBag.ReadProperty("SimpleText", Ambient.DisplayName)
  m_ShowTips = PropBag.ReadProperty("ShowTips", m_def_ShowTips)
End Sub

Private Sub UserControl_Resize()
  DrawStatusBar
End Sub

Private Sub UserControl_Show()
  DrawStatusBar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
  Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
  Call PropBag.WriteProperty("SimpleText", m_SimpleText, Ambient.DisplayName)
  Call PropBag.WriteProperty("ShowTips", m_ShowTips, m_def_ShowTips)
End Sub
