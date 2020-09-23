VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomWndow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' LAST TESTED FOR MEMORY LEAKS: 24 NOV 04 : None
'   Last leak found: 15 Nov - not destroying m_font:clsFrameControl

' HISTORY - Tested personally on Win2K & Win98, other O/S are up to you coders

' I. 1-24 NOV. Project began and 1st cut finished
'   -- given to psc to help troubleshoot on other O/S
'   Following are pending if project continues after feedback received
'       -- does not support MDI forms but will
'       -- does not contain true skinning ability but will
'       -- does not include owner-drawn menu class but may!
'   Next cut is expected to support MDIs and allow skinning abilities to
'   include on the fly user-callback drawing thru Implementation &
'   pre-loaded skins either via initial setup or possibly property pages
'       -- owner drawn menu class is last step
' II. To be determined. Note to self: split up major class; getting too big

' =============================== WARNING ===============================
' I have managed to pretty much prevent every instance of VB drawing
' on the window except when Me.Caption is used to set a caption & when
' maximizing/restoring/minimizing as noted below in Known Issues.
' In doing so, window styles are routinely added/removed which causes
' several more messages being processed and/or discarded.
' WARNING: This project is very heavily subclassed
' 1. Save changes often
' 2. Avoid editing your code in Break Mode (typed error can cause crash)
' 3. If placing a stop within the NewWndProc routine(s), be very cautious
' 4. The classes are designed to self-terminate subclassing when
'   Unload Me is used. Never use the END statement when subclassing
' 5. Should you want to control when subclassing is terminated,
'   then use this class's EndCustomWindow routine
' 6. Very little error trapping is performed. This is by design. While
'   building an app, I want to know where it breaks even if it means a crash
' 7. If marketing a similar project, retrofitting routines similar to VB
'   Accelerator's SSubTimer or the VB Thunking routines would be worthwhile
' ========================================================================

' ***************************** KNOWN ISSUES *****************************
' 1. Don't set window caption via Me.Caption : use class's SetCaption routine
'   -- Using Me.Caption is not trappable & is performed internally by VB
' 2. I have no control over what happens to the DC when passed to users for
'   custom on-the fly drawing. If user adds memory objects (fonts/pens/etc) &
'   doesn't unselect/destroy them -- leaks will occur. If user actually
'   unselects the bitmap then all hell breaks out -- leaks & possibly crash
' 3. In non-specific scenarios, the system menu may enable menu items that it
'   shouldn't, even though those items won't do anything. Haven't figured
'   this one out yet & using DrawMenuBar doesn't seem to help.
' 4. When maximizing, restoring or minimizing, the VB default blue titlebar
'   will be seen during the transition. Another internal thing done by VB.
' 5. Please ignore any/all Debug statements... This is sometimes the only
'   way to debug subclassed routines & will be removed when project is final
' **************************************************************************

' This class was designed as the user interface to all propeties/procedures
' dealing with menubars, titlebars, borders, & menus.  It can be converted
' to an usercontrol or DLL with minimal effort by copying & pasting.

' Usage. To use these classes, only reference this class directly, the others
' should not be directly accessed....
' 1. Should you want custom callback capability, read comments in
'       CustomWindowCalls and enter statement in your forms's
'       Declaration section as :  Implements CustomWindowCalls
'       Then go into each procedure added to your form and add
'       appropriate code or a simple tick if the procedure isn't used
' 2. Declare an instance of this class in form's Declaration section
' 3. Initialize the class at some point: Set myClass = New clsCustomWindow
' 4. These classes are designed to allow you to set all properties before
'       the window is even subclassed. This is a nice feature.


Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private c_Frame As clsFrameControl

Private Sub Class_Initialize()
Set c_Frame = New clsFrameControl
End Sub
Private Sub Class_Terminate()
If Not c_Frame Is Nothing Then c_Frame.EndSubclass
Set c_Frame = Nothing
End Sub

Public Sub BeginCustomWindow(ByVal formObject As Object, Optional CallBackClass As CustomWindowCalls)
' only accept forms; theoretically it could work with usercontrols & other hWnd objects
If TypeOf formObject Is Form Or _
   TypeOf formObject Is MDIForm Then
   
        If CallBackClass Is Nothing Then
            RemoveProp formObject.hWnd, "lvImpCB_Ptr"
        Else
            SetProp formObject.hWnd, "lvImpCB_Ptr", ObjPtr(CallBackClass)
        End If
        SetProp formObject.hWnd, "lvCFrame_Optr", ObjPtr(c_Frame)
        c_Frame.BeginSubclass formObject.hWnd
End If
End Sub
Public Sub EndCustomWindow()
' user-preempted subclass termination
    c_Frame.EndSubclass
    Set c_Frame = Nothing
    Set c_Frame = New clsFrameControl   ' for subsequent reference as needed
End Sub

' ------------ APPLICATION PROPERTIES -------------
Public Sub SetAppIcon(Icon32Pixels As Long)
c_Frame.SetApplicationIcon Icon32Pixels
End Sub
Public Property Get IsSubclassed(hWnd As Long)
If hWnd Then IsSubclassed = (GetProp(hWnd, "lvCFrame_Optr") <> 0)
End Property

' --------------- TITLE BAR PROPERTIES -------------
Public Property Let Font_TBar(newFont As StdFont)
Set c_Frame.Font = newFont
End Property
Public Property Set Font_TBar(newFont As StdFont)
Set c_Frame.Font = newFont
End Property
Public Property Get Font_TBar() As StdFont
Set Font_TBar = c_Frame.Font
End Property

Public Property Let EnableTBarBtns(ByVal buttonIDs As SysMenuItemConstants, isEnabled As Boolean)
c_Frame.sysMenuEnabled(buttonIDs) = isEnabled
End Property
Public Property Get EnableTBarBtns(ByVal buttonIDs As SysMenuItemConstants) As Boolean
EnableTBarBtns = c_Frame.sysMenuEnabled(buttonIDs)
End Property

Public Property Let BorderStyle(ByVal Style As WindowBorderStyleConstants)
c_Frame.SetBorderStyle True, Style
End Property
Public Property Get BorderStyle() As WindowBorderStyleConstants
BorderStyle = c_Frame.SetBorderStyle(False, 0)
End Property

Public Property Let VerticalCaption(isVertical As Boolean)
c_Frame.SetVerticalCaption True, isVertical
End Property
Public Property Get VerticalCaption() As Boolean
VerticalCaption = c_Frame.SetVerticalCaption(False, False)
End Property

Public Property Let CenteredCaption(isCentered As Boolean)
c_Frame.SetCenterCaption True, isCentered
End Property
Public Property Get CenteredCaption() As Boolean
CenteredCaption = c_Frame.SetCenterCaption(False, False)
End Property

Public Property Let HideDisabledButtons(bHide As Boolean)
c_Frame.SetHideDisabledBtns True, bHide
End Property
Public Property Get HideDisabledButtons() As Boolean
    HideDisabledButtons = c_Frame.SetHideDisabledBtns(False, False)
End Property

Public Property Let ShowInTaskBar(bShow As Boolean)
c_Frame.SetInTaskBar True, bShow
End Property
Public Property Get ShowInTaskBar() As Boolean
ShowInTaskBar = c_Frame.SetInTaskBar(False, False)
End Property

Public Property Let AlwaysActive(AllThreads As Boolean, KeepActive As Boolean)
c_Frame.SetKeepActive True, KeepActive, AllThreads
End Property
Public Property Get AlwaysActive(AllThreads As Boolean) As Boolean
AlwaysActive = c_Frame.SetKeepActive(False, False, AllThreads)
End Property

Public Property Let FontColor_TBar(ActiveColor As Boolean, newColor As Long)
c_Frame.SetFontColor True, ActiveColor, newColor
End Property
Public Property Get FontColor_TBar(ActiveColor As Boolean) As Long
FontColor_TBar = c_Frame.SetFontColor(False, ActiveColor, 0)
End Property

Public Sub SetCaption(newCaption As String, Optional hWnd As Long)
' if not yet subclassed, provide the hWnd; otherwise hWnd is not needed
c_Frame.SetTitle newCaption, hWnd
End Sub
Public Sub Refresh()
c_Frame.MenuBar.ForceMenuBarRepaint
End Sub

Public Sub AddTitleBarButton(ID As String, Position As Long, _
    Optional PositionType As TitlelBarBtnPosition, _
    Optional X As Long, Optional Y As Long, _
    Optional Width As Long, Optional Height As Long)
    c_Frame.AddToolBarButton ID, Position, PositionType, X, Y, Width, Height
End Sub

' -------------- MENU BAR PROPERTIES ---------------
Public Property Let Font_MenuBar(newFont As StdFont)
Set c_Frame.MenuBar.Font = newFont
End Property
Public Property Set Font_MenuBar(newFont As StdFont)
Set c_Frame.MenuBar.Font = newFont
'c_Frame.MenuBar.ForceMenuBarRepaint
End Property
Public Property Get Font_MenuBar() As StdFont
On Error Resume Next
Set Font_MenuBar = c_Frame.MenuBar.Font
End Property

Public Property Let FontColor_Menubar(State As FontStateColorConstants, newColor As Long)
c_Frame.MenuBar.FontColor(State) = newColor
End Property
Public Property Get FontColor_Menubar(State As FontStateColorConstants) As Long
FontColor_Menubar = c_Frame.MenuBar.FontColor(State)
End Property

Public Property Let MenuSelect_FlatStyle(isFlat As Boolean)
Call c_Frame.MenuBar.HiliteStyle(True, isFlat)
End Property
Public Property Get MenuSelect_FlatStyle() As Boolean
MenuSelect_FlatStyle = c_Frame.MenuBar.HiliteStyle(False, False)
End Property

Public Sub SetMenuSelectColors(Color1 As Long, Color2 As Long)
Call c_Frame.MenuBar.HiliteColors(Color1, Color2)
End Sub
