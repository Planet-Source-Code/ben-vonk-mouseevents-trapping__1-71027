This is a small class for 3 Mouse events which are not incuded in the VB environment.

Events:
  Hover(hWnd As Long, Name As String)
   - Returns the hWnd and Name of hooked object, when mouse enters the object.
     (if the object has no hWnd, the parent hWnd will be returned)

  Leave(hWnd As Long, Name As String)
   - Returns the hWnd and Name of hooked object, when mouse left the object.

  Wheel(hWnd As Long, Name As String, ScrollLines As Long)
   - Returns the hWnd and Name of the hooked object and the number of ScrollLines
     are set by Mouse or ScrollLines property, when mouse wheel is used.

Properties:
  HoverTime
    - Returns/sets the hover time in miliseconds, to hover the object.

  hWnd
   - Returns the hWnd of the hooked object.
     (when the object has no hWnd, the parent hWnd of the object will be returned)

  Name
   - Returns the name of the hooked object.

  ScrollLines
   - Returns/sets the number of lines they must be scrolled. (when it's set to 0 the Mouse default ScrollLines will be used)

Function:
  HookMouse(ByRef HookedObject As Object, Optional ByVal hInstance As Long, Optional ByVal ThreadID As Long) As Long
   - Hooked the object and returns the hWnd of the hooked object. (if returns 0 hookes failed)
     HookedObject is the object wich must be hooked. (if the object has no hWnd, the parent hWnd will be used)
     hInstance and ThreadID will be set for track MouseWheel event when mouse hover the object.
     When not set, the MouseWheel event will only be trapped if the object has the focus.
     (normaly App.hInstance and App.ThreadID will be used)

Sub:
  UnhookMouse(Optional ByVal ForceLeave As Boolean)
   - Unhookes the hooked object.
     ForceLeave is used for raising Leave event before the hooked object will be unhooked.
     (only if the new object has no hWnd and the parent of that object will also be hooked)


To demonstrate;
  move the Mouse over the CommandButton and scrolls the MouseWheel.
  Then leaves the CommandButton and move the Mouse over the other controls to.
