VERSION 5.00
Begin VB.Form frmMouseEvents 
   AutoRedraw      =   -1  'True
   Caption         =   "MouseEvents - MouseHover, MouseLeave and MouseWheel  (v3)"
   ClientHeight    =   5532
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10452
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMouseEvents.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5532
   ScaleWidth      =   10452
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fraDemo 
      Caption         =   "Frame"
      Height          =   2172
      Left            =   7320
      TabIndex        =   0
      Top             =   3240
      Width           =   3012
      Begin VB.PictureBox picDemo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         Height          =   252
         Left            =   120
         ScaleHeight     =   204
         ScaleWidth      =   2724
         TabIndex        =   1
         Top             =   240
         Width           =   2772
      End
      Begin VB.TextBox txtDemo 
         BackColor       =   &H00FF80FF&
         Height          =   288
         Left            =   120
         TabIndex        =   2
         Text            =   "TextBox"
         Top             =   600
         Width           =   2772
      End
      Begin VB.CommandButton cmdDemo 
         BackColor       =   &H0080FFFF&
         Caption         =   "MouseEvents"
         Height          =   732
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   2772
      End
      Begin VB.Label lblDemo 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2772
      End
   End
End
Attribute VB_Name = "frmMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Class
Private WithEvents MouseEvent As clsMouseEvents
Attribute MouseEvent.VB_VarHelpID = -1

' Private Variables
Private MouseWheelIsHooked    As Boolean
Private hWndObject            As Long
Private hWndTextBox           As Long

' hook the given object and returns the hWnd of that object
Private Function MouseHook(ByRef HookedObject As Object) As Long

   ' if MouseEvent is set only returns hWnd of the object
   If Not MouseEvent Is Nothing Then
      MouseHook = hWndObject
      Exit Function
   End If
   
   ' set MouseEvent with the class and hook the object
   Set MouseEvent = New clsMouseEvents
   MouseEvent.HoverTime = 1
   MouseHook = MouseEvent.HookMouse(HookedObject, App.hInstance, App.ThreadID)
   
   ' if mouse hooking failed then clear MouseEvent for other objects
   If MouseHook = 0 Then Set MouseEvent = Nothing

End Function

Private Sub MouseUnhook()

   ' clear MouseEvent and unhookes the mouse events
   If Not MouseEvent Is Nothing Then Set MouseEvent = Nothing

End Sub

' show the specified mouse event for the hooked control
Private Sub ShowMouseEvent(ByVal hWnd As Long, ByVal Name As String, ByVal IsEvent As String, ByVal Color As Long)

Dim strText As String

   strText = IsEvent & " " & Name
   
   If Name = "fraDemo" Then
      fraDemo.Caption = strText
      
   ElseIf hWnd = hWndTextBox Then
      txtDemo.Text = strText
      
   ElseIf Name = "picDemo" Then
      picDemo.Cls
      picDemo.Print strText
      
   ElseIf Name = "lblDemo" Then
      lblDemo.Caption = strText
      
   ElseIf Name = "cmdDemo" Then
      cmdDemo.BackColor = Color
      cmdDemo.Caption = IsEvent & vbCrLf & "hWnd = " & hWnd & vbCrLf & "Name = " & Name
   End If

End Sub

Private Sub cmdDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' set the cmdDemo control for hooking the mouse events
   hWndObject = MouseHook(cmdDemo)

End Sub

Private Sub Form_Load()

   hWndTextBox = txtDemo.hWnd
   picDemo.Print "PictureBox"
   
   ' information of the MouseEvents class module
   Print
   Print " This is a small class for 3 Mouse events which are not incuded in the VB environment."
   Print
   Print " Events:"
   Print "  Hover(hWnd As Long, Name As String)"
   Print "  Leave(hWnd As Long, Name As String)"
   Print "  Wheel(hWnd As Long, Name As String, ScrollLines As Long)"
   Print
   Print " Properties:"
   Print "  HoverTime  (returns/sets the hover time in miliseconds, to hover the object)"
   Print "  hWnd  (returns the hWnd of the hooked object)"
   Print "  Name  (returns the name of the hooked object)"
   Print "  ScrollLines  (returns/sets the number of lines they must be scrolled)"
   Print
   Print " Function:"
   Print "  HookMouse(ByRef HookedObject As Object, Optional ByVal hInstance As Long, Optional ByVal ThreadID As Long) As Long"
   Print
   Print " Sub:"
   Print "  UnhookMouse(Optional ByVal ForceLeave As Boolean)"
   Print
   Print
   Print " *** For more details see: ReadMe.txt ***"
   Print
   Print
   Print
   Print " To demonstrate; move the Mouse over the CommandButton and scrolls the MouseWheel."
   Print " Then leaves the CommandButton and move the Mouse over the other controls to."

End Sub

Private Sub Form_Terminate()

   ' clear the MouseEvent and unhook it
   Set MouseEvent = Nothing
   
End Sub

Private Sub fraDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' set the fraDemo control for hooking the mouse events
   hWndObject = MouseHook(fraDemo)

End Sub

Private Sub lblDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' use it before hook a control without hWnd
   ' (only needed when the parent also will be hooked, because when the
   '   control do not have a hWnd the MouseLeave event will not determine it)
   If Not MouseEvent Is Nothing Then If MouseEvent.Name <> lblDemo.Name Then Call MouseEvent.UnhookMouse(True)
   
   ' set the lblDemo control for hooking the mouse events
   hWndObject = MouseHook(lblDemo)

End Sub

' raised when the mouse hover the control
Private Sub MouseEvent_Hover(hWnd As Long, Name As String)

   Call ShowMouseEvent(hWnd, Name, "MouseHover:", &H80FF80)

End Sub

' raised when the mouse left the control
Private Sub MouseEvent_Leave(hWnd As Long, Name As String)

   Call ShowMouseEvent(hWnd, Name, "MouseLeave:", &H8080FF)
   Call MouseUnhook

End Sub

' raised when the mouse wheel is used
Private Sub MouseEvent_Wheel(hWnd As Long, Name As String, ScrollLines As Long)

Dim strCaption As String

   If ScrollLines > 0 Then
      strCaption = "Up"
      
   Else
      strCaption = "Down"
   End If
   
   strCaption = "Total ScrollLines " & strCaption & " = " & ScrollLines
   
   If Name = "cmdDemo" Then
      cmdDemo.BackColor = &HFF8080
      cmdDemo.Caption = strCaption
      
   ElseIf Name = "picDemo" Then
      picDemo.Cls
      picDemo.Print strCaption
   End If

End Sub

Private Sub picDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' set the picDemo control for hooking the mouse events
   hWndObject = MouseHook(picDemo)

End Sub

Private Sub txtDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' set the txtDemo control for hooking the mouse events
   hWndObject = MouseHook(txtDemo)

End Sub

