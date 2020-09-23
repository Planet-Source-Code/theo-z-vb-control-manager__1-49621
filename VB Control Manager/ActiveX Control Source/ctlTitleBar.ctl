VERSION 5.00
Begin VB.UserControl ctlTitleBar 
   ClientHeight    =   216
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4776
   ScaleHeight     =   216
   ScaleWidth      =   4776
   Begin VBControlManager.ctlButton cbtnClose 
      Height          =   180
      Left            =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      Picture         =   "ctlTitleBar.ctx":0000
   End
End
Attribute VB_Name = "ctlTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : ctlTitleBar.ctl                                             **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A custom title bar ActiveX control                          **
'** Usage       : To provide interface for the user to move the control in    **
'**               ControlManager at run-time                                  **
'** Dependencies: ctlButton, mdlAPI                                           **
'** Members     :                                                             **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Property   : CloseVisible                                             **
'**   * Methods    : -                                                        **
'**   * Events     : Click, CloseClick, DblClick, MouseDown, MouseMove,       **
'**                  MouseUp, Move, MoveBegin, MoveEnd                        **
'** Last modified on November 14, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Property Variables
Private mblnCloseVisible As Boolean

'--- Private Constants
Private Const mconGapWidth = 2             'gap width in pixel between the title
                                           '          bar and its control button
'--- Private Variables
Private mblnDrag As Boolean                         'indicating whether the user
                                                    '  is dragging the title bar
Private mposLastDrag As mdlAPI.POINTAPI                'last (x,y) coordinate of
                                                       '     the dragging action

'-------------------------------
' ActiveX Control Custom Events
'-------------------------------

'Description: Occurs when the user presses and then realeses a mouse button over
'             the control
Public Event Click()

'Description: Occurs when the user presses and then realeses a mouse button and
'             then presses and releases it again over the control
Public Event DblClick()

'Description: Occurs when the user presses and then releases a mouse button over
'             the close button
Public Event CloseClick()

'Description: Occurs when the user presses a mouse button over the control
'Argument   : Button, Shift, X, Y (see reference for MouseDown event in MSDN for
'                                  the description of the arguments)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                       ByVal X As Single, ByVal Y As Single)

'Description: Occurs when the user moves the mouse over the control
'Argument   : Button, Shift, X, Y (see reference for MouseMove event in MSDN for
'                                  the description of the arguments)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                       ByVal X As Single, ByVal Y As Single)

'Description: Occurs when the user releases a mouse button over the control
'Argument   : Button, Shift, X, Y (see reference for MouseUp event in MSDN for
'                                  the description of the arguments)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                     ByVal X As Single, ByVal Y As Single)

'Description: Occurs when the user moves the mouse after the BeginMove event and
'             before the EndMove event
'Arguments  : Shift (an integer that corresponds to the state of the SHIFT,
'                    CTRL, and ALT keys)
Public Event Move(ByVal Shift As Integer)

'Description: Occurs when the user presses a mouse left-button over the custom-
'             drawing title bar
'Arguments  : Shift (an integer that corresponds to the state of the SHIFT,
'                    CTRL, and ALT keys)
Public Event MoveBegin(ByVal Shift As Integer)

'Description: Occurs when the user releases the mouse button after the BeginMove
'             event
'Arguments  : Shift (an integer that corresponds to the state of the SHIFT,
'                    CTRL, and ALT keys)
Public Event MoveEnd(ByVal Shift As Integer)

'------------------------
' ActiveX Control Events
'------------------------

' Purpose    : Raises custom event CloseClick
' Effect     : As specified
Private Sub cbtnClose_Click()
  RaiseEvent CloseClick
End Sub

' Purpose    : Raises custom event Click
' Effect     : As specified
Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

' Purpose    : Begins move the title bar virtually and raises custom event
'              MouseDown
' Effect     : * If Button = vbLeft Button, as specified
'              * Otherwise no effect
' Inputs     : Button, Shift, X, Y
Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
  If Button = vbLeftButton Then
    mblnDrag = True
    mposLastDrag.X = X
    mposLastDrag.Y = Y
    RaiseEvent MoveBegin(Shift)
  End If
End Sub

' Purpose    : Moves the title bar virtually and raises custom event Move
' Assumption : The UserControl_MouseDown procedure has been called
' Effect     : * If mblnDrag = true, as specified
'              * Otherwise, no effect
' Inputs     : Index, Button, Shift, x, y
Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, X As Single, Y As Single)
  If mblnDrag Then
    RaiseEvent Move(Shift)
    mposLastDrag.X = X
    mposLastDrag.Y = Y
  Else
    RaiseEvent MouseMove(Button, Shift, X, Y)
  End If
End Sub

' Purpose    : Ends the title bar virtual move action and raises custom event
'              Moved
' Effect     : * If mblnDrag = true, as specified
'              * Otherwise, no effect
' Inputs     : Index, Button, Shift, x, y
Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
  If mblnDrag Then
    mblnDrag = False
    RaiseEvent MoveEnd(Shift)
  End If
End Sub

' Purpose    : Draws the title bar
Private Sub UserControl_Paint()
  DrawTitleBar
End Sub

' Purpose    : Adjusts the components inside the control to agree with the
'              control's size
' Effect     : See the codes below
Private Sub UserControl_Resize()
  With cbtnClose
    .Height = UserControl.ScaleHeight
    .Width = .Height
    .Left = UserControl.ScaleWidth - .Width
  End With
  DrawTitleBar
End Sub

'----------------------------
' ActiveX Control Properties
'----------------------------

' Purpose    : Sets a value that determines whether a close button in the title
'              bar is visible
' Effect     : As specified
' Input      : blnCloseVisible (the new CloseVisible property value)
' Return     : As specified
Public Property Let CloseVisible(blnCloseVisible As Boolean)
  mblnCloseVisible = blnCloseVisible
  
  cbtnClose.Visible = mblnCloseVisible
End Property

' Purpose    : Returns a value that determines whether a close button in the
'              title bar is visible
' Return     : As specified
Public Property Get CloseVisible() As Boolean
  CloseVisible = mblnCloseVisible
End Property

'--------------------
' Private Procedures
'--------------------

' Purpose    : Draws the title bar
' Effect     : As specified
Private Sub DrawTitleBar()
  Const conMarginTop = 4
  Const conMarginBottom = 4

  Dim lngBarWidth As Long
  Dim lngCloseButtonWidth As Long
  Dim rec As mdlAPI.RECT
  
  If mblnCloseVisible Then
    lngCloseButtonWidth = cbtnClose.Width
  Else
    lngCloseButtonWidth = 0
  End If
  lngBarWidth = (UserControl.ScaleHeight \ Screen.TwipsPerPixelY) - _
                conMarginTop - conMarginBottom
  mdlAPI.SetRect lpRect:=rec, X1:=0, Y1:=conMarginTop, _
                 X2:=(UserControl.ScaleWidth - _
                      lngCloseButtonWidth - mconGapWidth) \ _
                     Screen.TwipsPerPixelX, _
                 Y2:=(conMarginTop + lngBarWidth)
  mdlAPI.DrawEdge hdc:=UserControl.hdc, qrc:=rec, _
                  edge:=mdlAPI.EDGE_RAISED, grfFlags:=mdlAPI.BF_RECT
End Sub
