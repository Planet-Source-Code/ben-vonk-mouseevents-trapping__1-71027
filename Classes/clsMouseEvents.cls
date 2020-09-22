VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'MouseEvents Class
'
'Author Ben Vonk
'24-05-2007 First version
'02-06-2007 Second version improved and implement some suggestions of LaVolpe
'06-06-2007 Third version little changes are made and implement Mouse_Hover event of Tom Pydeski

' Public Events
Public Event Hover(hWnd As Long, Name As String)
Attribute Hover.VB_Description = "Raised when the mouse hover the hooked object."
Public Event Leave(hWnd As Long, Name As String)
Attribute Leave.VB_Description = "Raised when the mouse left the hooked object."
Public Event Wheel(hWnd As Long, Name As String, ScrollLines As Long)
Attribute Wheel.VB_Description = "Raised when the mouse wheel is used."

' Private Constant
Private Const SPI_GETWHEELSCROLLLINES As Long = 104

' Private Variables
Private MouseWheelIsHooked            As Boolean
Private m_HoverTime                   As Long
Private m_hWnd                        As Long
Private m_ScrollLines                 As Long
Private m_Name                        As String

' Private API
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Property Get HoverTime() As Long

   HoverTime = m_HoverTime

End Property

Public Property Let HoverTime(NewHoverTime As Long)

   m_HoverTime = NewHoverTime

End Property

Public Property Get hWnd() As String

   hWnd = m_hWnd

End Property

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name of the hooked object."

   Name = m_Name

End Property

Public Property Get ScrollLines() As Long
Attribute ScrollLines.VB_Description = "Returns/Sets the number of lines they must be scrolled."

   ScrollLines = m_ScrollLines

End Property

Public Property Let ScrollLines(ByVal NewScrollLines As Long)

   If NewScrollLines = 0 Then NewScrollLines = GetScrollLines
   
   m_ScrollLines = NewScrollLines

End Property

Public Function HookMouse(ByRef hWndObject As Object, Optional ByVal hInstance As Long, Optional ByVal ThreadID As Long) As Long
Attribute HookMouse.VB_Description = "Hooked the object for tracking mouse events."

Dim blnhWndParentUsed As Boolean
Dim blnIsHooked       As Boolean
Dim lnghWnd           As Long

   On Local Error GoTo GetParenthWnd
   m_Name = hWndObject.Name
   lnghWnd = hWndObject.hWnd
   
   GoTo Continue
   
GetParenthWnd:
   On Local Error GoTo ExitFunction
   lnghWnd = GetParentInfo(hWndObject)
   blnhWndParentUsed = True
   
Continue:
   If Not MouseEvent Is Nothing Then GoTo ExitFunction
   
   Set MouseEvent = Me
   blnIsHooked = MouseLeaveHook(lnghWnd, (Not hInstance And Not ThreadID), blnhWndParentUsed, m_HoverTime)
   
   ' if failed then clear MouseEvent and return 0 for hWnd
   If Not blnIsHooked Then
      Set MouseEvent = Nothing
      lnghWnd = 0
   End If
   
ExitFunction:
   On Local Error GoTo 0
   m_hWnd = lnghWnd
   HookMouse = m_hWnd
   
   If blnIsHooked And (hInstance <> 0) And (ThreadID <> 0) Then MouseWheelIsHooked = MouseWheelHook(hInstance, ThreadID)

End Function

Public Sub UnhookMouse(Optional ByVal ForceLeave As Boolean)
Attribute UnhookMouse.VB_Description = "Unhookes the object."

   ' used for hooked objects with no hWnd
   If ForceLeave Then Call MouseLeave
   
   Call MouseLeaveUnhook(m_hWnd)
   
   If MouseWheelIsHooked Then Call MouseWheelUnhook
   
   Set MouseEvent = Nothing
   MouseWheelIsHooked = False

End Sub

Friend Sub MouseHover()

Static lngInhWnd As Long

   If lngInhWnd <> m_hWnd Then
      lngInhWnd = m_hWnd
      RaiseEvent Hover(m_hWnd, m_Name)
   End If

End Sub

Friend Sub MouseLeave()

   ' first unhook the object
   Call UnhookMouse
   
   RaiseEvent Leave(m_hWnd, m_Name)

End Sub

Friend Sub MouseWheel(ByVal WheelUp As Boolean)

   RaiseEvent Wheel(m_hWnd, m_Name, -m_ScrollLines + ((m_ScrollLines * 2) And WheelUp))

End Sub

Private Function GetScrollLines() As Long

   ' get the mouse scrolllines set by system
   SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, m_ScrollLines, 0
   GetScrollLines = m_ScrollLines + (1 And (m_ScrollLines = 0)) ' if there are no scrollines set it to 1

End Function

Private Sub Class_Initialize()

   GetScrollLines

End Sub

Private Sub Class_Terminate()

   Call UnhookMouse

End Sub

