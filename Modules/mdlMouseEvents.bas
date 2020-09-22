Attribute VB_Name = "mdlMouseEvents"
Option Explicit

' Private Constants
Private Const GWL_WNDPROC   As Long = (-4)
Private Const PM_REMOVE     As Long = &H1
Private Const TME_CANCEL    As Long = &H80000000
Private Const TME_HOVER     As Long = &H1
Private Const TME_LEAVE     As Long = &H2&
Private Const WH_MOUSE      As Long = &H7
Private Const WM_MOUSEHOVER As Long = &H2A1
Private Const WM_MOUSELEAVE As Long = &H2A3&
Private Const WM_MOUSEMOVE  As Long = &H200
Private Const WM_MOUSEWHEEL As Long = &H20A

' Private Types
Private Type PointAPI
   X                       As Long
   Y                       As Long
End Type

Private Type Msg
   hWnd                    As Long
   Message                 As Long
   wParam                  As Long
   lParam                  As Long
   Time                    As Long
   Pt                      As PointAPI
End Type

Private Type Rect
   Left                     As Long
   Top                      As Long
   Right                    As Long
   Bottom                   As Long
End Type

Private Type TrackMouseEventType
   cbSize                   As Long
   dwFlags                  As Long
   hwndTrack                As Long
   dwHoverTime              As Long
End Type

' Public Class
Public MouseEvent           As clsMouseEvents

' Private Variables
Private IsFocusedWheel      As Boolean
Private IshWndParent        As Boolean
Private hWndMouseLeave      As Long
Private hWndMouseWheel      As Long
Private PrevMouseEvent      As Long
Private WindowMessage       As Msg
Private MouseXY             As PointAPI
Private ChildRect           As Rect

' Private API's
Private Declare Function TrackMouseEvent Lib "ComCtl32" Alias "_TrackMouseEvent" (lpEventTrack As TrackMouseEventType) As Long
Private Declare Function CallNextHookEx Lib "User32" (ByVal hWnd As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function PeekMessage Lib "User32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal IDHook As Long, ByVal lpFn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hWnd As Long) As Long

' gets info of the parent, set child rectangle and returns parent hWnd
Public Function GetParentInfo(ByRef hWndObject As Object, Optional ByRef ChildTop As Long, Optional ByRef ChildLeft As Long) As Long

Static objChild      As Object
Static lnghWndParent As Long

Dim lngParentBorder  As Long

   ' keep child object for use when form is reached
   If objChild Is Nothing Then Set objChild = hWndObject
   
   If TypeOf hWndObject.Container Is Form Then
      ' if form is reached then set child info
      With hWndObject.Container
         lngParentBorder = (.ScaleX(.Width, .ScaleMode, vbPixels) - .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)) / 2 - 1
         ChildTop = .ScaleY(ChildTop + .Top, .ScaleMode, vbPixels) + (.ScaleY(.Height, .ScaleMode, vbPixels) - .ScaleY(.ScaleHeight, .ScaleMode, vbPixels)) - lngParentBorder
         ChildLeft = .ScaleX(ChildLeft + .Left, .ScaleMode, vbPixels) + lngParentBorder
         ChildRect.Top = ChildTop + .ScaleY(objChild.Top, .ScaleMode, vbPixels)
         ChildRect.Left = ChildLeft + .ScaleX(objChild.Left, .ScaleMode, vbPixels)
         ChildRect.Right = ChildRect.Left + .ScaleX(objChild.Width, .ScaleMode, vbPixels)
         ChildRect.Bottom = ChildRect.Top + .ScaleY(objChild.Height, .ScaleMode, vbPixels)
         lnghWndParent = objChild.Parent.hWnd
         Set objChild = Nothing
      End With
      
   Else
      ' search for next parent
      GetParentInfo hWndObject.Container, ChildTop + hWndObject.Container.Top, ChildLeft + hWndObject.Container.Left
   End If
   
   GetParentInfo = lnghWndParent

End Function

' sets the info for tracking mouse events for the hooked object, returns False if failed
Public Function MouseLeaveHook(ByRef hWnd As Long, ByVal FocusedWheel As Boolean, ByVal hWndParentUsed As Boolean, ByVal HoverTime As Long) As Boolean

Dim tmeMouseEvent As TrackMouseEventType

   On Local Error GoTo ExitFunction
   
   With tmeMouseEvent
      .cbSize = Len(tmeMouseEvent)
      .hwndTrack = hWnd
      .dwFlags = TME_LEAVE Or TME_HOVER
      .dwHoverTime = HoverTime
   End With
   
   IsFocusedWheel = FocusedWheel
   IshWndParent = hWndParentUsed
   PrevMouseEvent = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf CheckMouseLeave)
   TrackMouseEvent tmeMouseEvent
   hWndMouseLeave = hWnd
   MouseLeaveHook = True
   
ExitFunction:
   On Local Error GoTo 0

End Function

' hook for tracking mouse wheel event, returns False if failed
Public Function MouseWheelHook(ByVal hInstance As Long, ByVal ThreadID As Long) As Boolean

   On Local Error GoTo ExitFunction
   hWndMouseWheel = SetWindowsHookEx(WH_MOUSE, AddressOf CheckMouseWheel, hInstance, ThreadID)
   MouseWheelHook = True
   
ExitFunction:
   On Local Error GoTo 0

End Function

' unhook mouse events
Public Sub MouseLeaveUnhook(ByVal hWnd As Long)

Dim tmeMouseEvent As TrackMouseEventType

   With tmeMouseEvent
      .cbSize = Len(tmeMouseEvent)
      .hwndTrack = hWnd
      .dwFlags = TME_LEAVE Or TME_CANCEL
      TrackMouseEvent tmeMouseEvent
      .dwFlags = TME_HOVER Or TME_CANCEL
      TrackMouseEvent tmeMouseEvent
   End With
   
   SetWindowLong hWndMouseLeave, GWL_WNDPROC, PrevMouseEvent

End Sub

' unhook MouseWheel event
Public Sub MouseWheelUnhook()

   UnhookWindowsHookEx hWndMouseWheel

End Sub

' check if mouse left the object
Private Function CheckMouseLeave(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   If IshWndParent Then
      GetCursorPos MouseXY
      
      If PtInRect(ChildRect, MouseXY.X, MouseXY.Y) Then
         Call MouseEvent.MouseHover
         
      Else
         Call MouseEvent.MouseLeave
      End If
      
   ElseIf wMsg = WM_MOUSELEAVE Then
      Call MouseEvent.MouseLeave
      
   ' if IsFocusedWheel is set then check if mouse wheel is used
   ElseIf wMsg = WM_MOUSEWHEEL Then
      If IsFocusedWheel Then Call MouseEvent.MouseWheel(wParam > 0)
      
   ElseIf wMsg = WM_MOUSEHOVER Then
      Call MouseEvent.MouseHover
   End If
   
   ' call next hook
   CheckMouseLeave = CallWindowProc(PrevMouseEvent, hWnd, wMsg, wParam, lParam)

End Function

' check if mouse wheel is used
Private Function CheckMouseWheel(ByVal IDHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   If (IDHook >= 0) And (wParam = WM_MOUSEWHEEL) Then If PeekMessage(WindowMessage, 0, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then Call MouseEvent.MouseWheel(WindowMessage.wParam > 0)
   
   ' call next hook
   CheckMouseWheel = CallNextHookEx(hWndMouseWheel, IDHook, wParam, ByVal lParam)

End Function

