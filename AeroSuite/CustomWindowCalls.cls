VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomWindowCalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' When is implementation Required?
' - When you want to custom draw a titlebar or menubar background
' - When you add custom titlebar buttons
' Otherwise, when call clsCustomWindow.BeginCustomWindow, you can
'   simply ignore the ClassCallBack parameter

' whichever forms will be subclassed add the following line to the
' Declaration section of those forms:

' Implements CustomWindowCalls      ' < optional see notes above
' Private myVariableName As New clsCustomWindow

' Then somewhere use the following line to trigger custom window:
' myVariableName.BeginCustomWindow Me, Me < using Implementation, or
' myVariableName.BeginCustomWindow Me     < not using Implementation

' Last but not least, you MUST activate the implementation & that is
' as simple as clicking the implementation name from the top left dropdown
' box on your form's code window. Then select each of the procedures in the
' the form's top right dropdown box. You don't need to add any code unless
' you want to.

Option Explicit


Public Sub UserDrawnTitleBar(ByVal hDC As Long, ByVal CX As Long, _
                ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)
' allows user to customize titlebar on the fly
' hDC is the DC to draw in, remember to unselect anything selected into the DC
' Cx is the width of the titlebar
' Cy is the height of the titlebar
' The left/top coordinates are always zero
' If custom drawn, set Modified = True to prevent overdrawing your work

End Sub

Public Sub UserDrawnMenuBar(ByVal hDC As Long, ByVal CX As Long, _
                ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)
' allows user to customize menubar on the fly
' hDC is the DC to draw in, remember to unselect anything selected into the DC
' Cx is the width of the menubar
' Cy is the height of the menubar
' The left/top coordinates are always zero
' If custom drawn, set Modified = True to prevent overdrawing your work

End Sub


Public Sub UserDrawnButton(ByVal ID As String, ByVal State As Integer, _
        ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal CX As Long, ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)
' allows user to draw custom titlebar buttons
' ID is button ID provided when button was created
' State : 0=normal, 1=down/depressed, 2=disabled
' hDC is the DC to draw in, remember to unselect anything selected into the DC
' X is the left coordinate, Y is the top coordinate
' Cx is the width & Cy is the height

End Sub

Public Sub UserButtonClick(ByVal ID As String)
' informs user that a custom titlebar button was clicked
' ID is button ID provided when button was created
End Sub

Public Sub UserDrawnBorder(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal HasFocus As Boolean, Modified As Boolean)

End Sub

Public Sub BeforeDrawCaption(ByVal hDC As Long, Caption As String, ByVal X As Long, ByVal Y As Long)

End Sub

Public Sub OnCreateRegion(hRgn As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long)

End Sub

Public Sub EnterExitSizing(ByVal BeginSizing As Boolean, UserRedrawn As Boolean)
' informs user that sizing either started or ended
' if BeginSizing=False, sizing has ended
' User can adjust quality of user-drawn items to speed up drawing
' and then call a Redraw when sizing has terminated so items can
' be drawn in fine detail.

' If redrawn by user set the UserRedrawn flag to true to prevent another
' redraw of the window
End Sub
