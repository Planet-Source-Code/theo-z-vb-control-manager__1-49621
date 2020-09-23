VERSION 5.00
Begin VB.UserControl ControlManager 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   FillColor       =   &H00404040&
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlControlManager.ctx":0000
   Begin VBControlManager.ctlRect crecControl 
      Height          =   3195
      Left            =   435
      TabIndex        =   2
      Top             =   285
      Visible         =   0   'False
      Width           =   4110
      _ExtentX        =   7260
      _ExtentY        =   5630
   End
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3390
      Index           =   9999
      Left            =   105
      ScaleHeight     =   3396
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   180
   End
   Begin VBControlManager.ctlTitleBar ctbTitleBar 
      Height          =   180
      Index           =   9999
      Left            =   405
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   4185
      _ExtentX        =   7387
      _ExtentY        =   318
   End
End
Attribute VB_Name = "ControlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An ActiveX control to allow the user to resize docked controls at run time"
'*******************************************************************************
'** File Name   : ControlManager.ctl                                          **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : An ActiveX control to allow the user to dock controls and   **
'**               resize, move and show/hide them at run-time                 **
'** Dependencies: All files in the VB Control Manager and SubTimer6 project   **
'** Members     :                                                             **
'**   * Collections: Controls, Splitters                                      **
'**   * Objects    : -                                                        **
'**   * Properties : ActivateColor (r/w), BackColor (r/w), ClipCursor (r/w),  **
'**                  Enable (r/w), FillContainer (r/w), LiveUpdate (r/w),     **
'**                  MarginBottom (r/w), MarginLeft (r/w), MarginRight (r/w), **
'**                  MarginTop (r/w), Size (r/w),                             **
'**                  TitleBar_CloseVisible (r/w), TitleBar_Height (r/o),      **
'**                  TitleBar_Visible (r/w)                                   **
'**   * Methods    : CloseControl, MoveControl, MoveSplitters, OpenControl    **
'**   * Events     : ControlAfterClose, ControlBeforeClose, ControlMove,      **
'**                  ControlMoveBegin, ControlMoveEnd, SplitterClick,         **
'**                  SplitterDblClick, SplitterMouseDown, SplitterMouseMove,  **
'**                  SplitterMouseUp, SplitterMove, SplitterMoveBegin,        **
'**                  SplitterMoveEnd, TitleBarClick, TitleBarDblClick,        **
'**                  TitleBarMouseDown, TitleBarMouseMove, TitleBarMouseUp    **
'** Credits     : * To Steve McMahon for the SubTimer6.vbp and for the        **
'**                 excellent articles at http://www.vbaccelerator.com        **
'**               * To KPD-Team at http://www.allapi.net for their API-Guide  **
'**               * To Carles P.V. (carles_pv@terra.es), Vlad Vissoultchev    **
'**                 (wqw@bora.exco.net), Paul Caton (paul_caton@hotmail.com)  **
'**                 and umairata (umairata@hotmail.com) for various bugs      **
'**                 information.                                              **
'** Last modified on November 14, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Public Type Declaration
Public Enum genmMoveDestination
  mdControlTop
  mdControlRight
  mdControlBottom
  mdControlLeft
  mdEdgeTop
  mdEdgeRight
  mdEdgeBottom
  mdEdgeLeft
  mdSplitter
End Enum

'--- Collection Variables
' Note: These collection are used to represents virtual controls and splitters
Private WithEvents mControls As clsControls
Attribute mControls.VB_VarHelpID = -1
Private WithEvents mSplitters As clsSplitters
Attribute mSplitters.VB_VarHelpID = -1

'--- Property Variables
Private mblnFillContainer As Boolean
Private mlngMarginBottom As Long
Private mlngMarginLeft As Long
Private mlngMarginRight As Long
Private mlngMarginTop As Long

'--- PropBag Names
Private Const mconActiveColor As String = "ActiveColor"
Private Const mconBackColor As String = "BackColor"
Private Const mconClipCursor As String = "ClipCursor"
Private Const mconEnable As String = "Enable"
Private Const mconFillContainer As String = "FillContainer"
Private Const mconLiveUpdate As String = "LiveUpdate"
Private Const mconMarginBottom As String = "MarginBottom"
Private Const mconMarginLeft As String = "MarginLeft"
Private Const mconMarginRight As String = "MarginRight"
Private Const mconMarginTop As String = "MarginTop"
Private Const mconSize As String = "Size"
Private Const mconTitleBar_CloseVisible As String = "TitleBar_CloseVisible"
Private Const mconTitleBar_Height As String = "TitleBar_Height"
Private Const mconTitleBar_Visible As String = "TitleBar_Visible"

'--- Property Default Values
Private Const mconDefaultFillContainer As Boolean = True
Private Const mconDefaultMarginBottom As Long = 0
Private Const mconDefaultMarginLeft As Long = 0
Private Const mconDefaultMarginRight As Long = 0
Private Const mconDefaultMarginTop As Long = 0

'--- Te Variables below is used to save procedure-level variables value that
'    lost in subclassing process
' Procedure-level variables in MouseDown, MouseMove, and MouseUp events
Private mScIndex As Integer
Private mScButton As Integer
Private mScShift As Integer
Private mScX As Single
Private mScY As Single

'--- Other Variables
Private mblnDragSplitter As Boolean                 'indicating whether the user
                                                    '   is dragging the splitter
Private mblnControlMoved As Boolean            'indicating whether a control has
                                               '     just been moved by the user
Private mblnSplitterMoved As Boolean          'indicating whether a splitter has
                                              '      just been moved by the user
Private mblnVisibleSave As Boolean              'to restore the Visible property
                                                '      of the control's instance
Private mhwndParent As Long               'the handle of the control's container
Private mhwndRoot As Long          'the handle of the root window of the control
Private mlngDragStart As Long    'the x- or y- coordinate (depends on the active
                                 '                Splitter 's orientation) where
                                 '                    the user strats to drag it
Private muposPrev As mdlAPI.POINTAPI 'previous mouse pointer coordinate relative
                                     '   to the splitter (note: this variable is
                                     '        used to make sure the custom event
                                     '                 MouseMove works properly)

'--- Implements the Interface
Implements ISubclass
Implements TitleBar

'-------------------------------
' ActiveX Control Custom Events
'-------------------------------

'Description: Occurs when a control has just been closed by the user
'Argument   : IdControl (a value that uniquely identifies the control that has
'                        just been closed by the user)
Public Event ControlAfterClose(ByVal IdControl As Long)
Attribute ControlAfterClose.VB_Description = "Occurs when a control has just been closed by the user"

'Description: Occurs after the user presses a close button of certain control
'             and before the control is closed
'Arguments  : * IdControl (a value that uniquely identifies the control that
'                          about to be closed)
'             * Cancel (setting this argument to true stops the control from
'                       closing)
Public Event ControlBeforeClose(ByVal IdControl As Long, _
                                ByRef Cancel As Boolean)
Attribute ControlBeforeClose.VB_Description = "Occurs after the user presses a close button of certain control and before the control is closed"

'Description: Occurs when the user is moving a control
'Arguments  : * IdControl (a value that uniquely identifies the control that is
'                          being moved by the user)
'             * Shift (an integer that corresponds to the state of the SHIFT,
'                      CTRL, and ALT keys)
'             * Left (an integer indicating the x-coordinate for the left edge
'                     of the current position of the rectangle that represents
'                     the moving control)
'             * Top (an integer indicating the y-coordinate for the top edge of
'                    the current position of the rectangle that represents the
'                    moving control)
'             * Width (an integer indicating the current width of the rectangle
'                      that represents the moving control)
'             * Height (an integer indicating the current height of the
'                       rectangle that represents the moving control)
Public Event ControlMove(ByVal IdControl As Long, ByVal Shift As Integer, _
                         ByVal Left As Long, ByVal Top As Long, _
                         ByVal Width As Long, ByVal Height As Long)
Attribute ControlMove.VB_Description = "Occurs when the user is moving a control"

'Description: Occurs when the user is about to move a control, i.e. the first
'             time the rectangle that represents the moving control occurs
'Arguments  : * IdControl (a value that uniquely identifies the control that is
'                          ready to be moved)
'             * Shift (an integer that corresponds to the state of the SHIFT,
'                      CTRL, and ALT keys)
Public Event ControlMoveBegin(ByVal IdControl As Long, ByVal Shift As Integer)
Attribute ControlMoveBegin.VB_Description = "Occurs when the user is about to move a control, i.e. the first time the rectangle that represents the moving control occurs"

'Description: Occurs when the user is finished moving a control, i.e. when the
'             rectangle that represents the moving control disappears
'Arguments  : * IdControl (a value that uniquely identifies the control that
'                          has just been moved)
'             * Shift (an integer that corresponds to the state of the SHIFT,
'                      CTRL, and ALT keys)
'             * Moved (a value that determines whether the control is moved)
Public Event ControlMoveEnd(ByVal IdControl As Long, _
                            ByVal Shift As Integer, ByVal Moved As Boolean)
Attribute ControlMoveEnd.VB_Description = "Occurs when the user is finished moving a control, i.e. when the rectangle that represents the moving control disappears"

'Description: Occurs when the user presses and then realeses a mouse button over
'             a splitter
'Argument   : IdSplitter (a value that uniquely identifies the splitter that has
'                         just been clicked by the user)
Public Event SplitterClick(ByVal IdSplitter As Long)
Attribute SplitterClick.VB_Description = "Occurs when the user presses and then realeses a mouse button over a splitter"

'Description: Occurs when the user presses and then realeses a mouse button and
'             then presses and releases it again over a splitter
'Argument   : IdSplitter (a value that uniquely identifies the splitter that has
'                         just been double-clicked by the user)
Public Event SplitterDblClick(ByVal IdSplitter As Long)
Attribute SplitterDblClick.VB_Description = "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a splitter"

'Description: Occurs when the user presses a mouse button over a splitter
'Arguments  : * IdSplitter (a value that uniquely identifies the splitter where
'                           the user presses a mouse button over)
'             * Button, Shift, X, Y (see reference for MouseDown event in
'                                    MSDN Library for the description of the
'                                    arguments)
Public Event SplitterMouseDown( _
               ByVal IdSplitter As Long, ByVal Button As Integer, _
               ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
             )
Attribute SplitterMouseDown.VB_Description = "Occurs when the user presses a mouse button over a splitter"

'Description: Occurs when the user moves a mouse over a splitter without moving
'             the splitter
'Arguments  : * IdSplitter (a value that uniquely identifies a splitter where
'                           the user moves a mouse over)
'             * Button, Shift, X, Y (see reference for MouseMove event in
'                                    MSDN Library for the description of the
'                                    arguments)
Public Event SplitterMouseMove( _
               ByVal IdSplitter As Long, ByVal Button As Integer, _
               ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
             )
Attribute SplitterMouseMove.VB_Description = "Occurs when the user moves a mouse over a splitter without moving the splitter"

'Description: Occurs when the user releases a mouse button over a splitter
'             without previously moving the splitter
'Arguments  : * IdSplitter (a value that uniquely identifies the splitter where
'                           the user releases a mouse button over)
'             * Button, Shift, X, Y (see reference for MouseUp event in
'                                    MSDN for the description of the arguments)
Public Event SplitterMouseUp( _
               ByVal IdSplitter As Long, ByVal Button As Integer, _
               ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
             )
Attribute SplitterMouseUp.VB_Description = "Occurs when the user releases a mouse button over a splitter without previously moving the splitter"

'Description: Occurs when the user is moving a splitter
'Arguments  : * IdSplitter (a value that uniquely identifies the splitter that
'                           is being moved by the user)
'             * Shift, X, Y (see reference for MouseMove event in MSDN for the
'                            description of the arguments)
Public Event SplitterMove(ByVal IdSplitter As Long, ByVal Shift As Integer, _
                          ByVal X As Single, ByVal Y As Single)
Attribute SplitterMove.VB_Description = "Occurs when the user is moving a splitter"

'Description: Occurs when the user is about to move a splitter
'Arguments  : * IdSplitter (A value that uniquely identifies the splitter that
'                           is about to be moved by the user)
'             * Shift, X, Y (see reference for MouseDown event in MSDN for the
'                            description of the arguments)
Public Event SplitterMoveBegin( _
               ByVal IdSplitter As Long, ByVal Shift As Integer, _
               ByVal X As Single, ByVal Y As Single _
             )
Attribute SplitterMoveBegin.VB_Description = "Occurs when the user is about to move a splitter"

'Description: Occurs when the user is finished moving a splitter
'Arguments  : * IdSplitter (a value that uniquely identifies the splitter that
'                           has just been moved by the user)
'             * Shift, X, Y (see reference for MouseUp event in MSDN for the
'                            description of the arguments)
Public Event SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
Attribute SplitterMoveEnd.VB_Description = "Occurs when the user presses and then realeses a mouse button over a control title bar"
                             
'Description: Occurs when the user presses and then realeses a mouse button over
'             a control title bar
'Argument   : IdControl (a value that uniquely identifies the control whose
'                        title bar has just been clicked by the user)
Public Event TitleBarClick(ByVal IdControl As Long)
Attribute TitleBarClick.VB_Description = "Occurs when the user presses and then realeses a mouse button over a control title bar"

'Description: Occurs when the user presses and then realeses a mouse button and
'             then presses and releases it again over a control title bar
'Argument   : IdControl (a value that uniquely identifies the control that owns
'                        the title bar)
Public Event TitleBarDblClick(ByVal IdControl As Long)
Attribute TitleBarDblClick.VB_Description = "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a control title bar"

'Description: Occurs when the user presses a mouse button over a control title
'             bar
'Argument   : * IdControl (a value that uniquely identifies the control that own
'                          the title bar where the user presses a mouse button
'                          over)
'             * Button, Shift, X, Y (see reference for MouseDown event in
'                                    MSDN Library for the description of the
'                                    arguments)
Public Event TitleBarMouseDown( _
               ByVal IdControl As Long, ByVal Button As Integer, _
               ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
             )
Attribute TitleBarMouseDown.VB_Description = "Occurs when the user presses a mouse button over a control title bar"

'Description: Occurs when the user moves the mouse over a control title bar
'             without moving the control
'Argument   : * IdControl (A value that uniquely identifies the control that own
'                          the title bar where the user moves a mouse over)
'             * Button, Shift, X, Y (see reference for MouseMove event in
'                                    MSDN Library for the description of the
'                                    arguments)
Public Event TitleBarMouseMove( _
               ByVal IdControl As Long, ByVal Button As Integer, _
               ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
             )
Attribute TitleBarMouseMove.VB_Description = "Occurs when the user moves a mouse over a control title bar without moving the control"

'Description: Occurs when the user releases a mouse button over a control title
'             bar without previously moving the control
'Argument   : * IdControl (a value that uniquely identifies the control that own
'                          the title bar where the user releases a mouse button
'                          over)
'             * Button, Shift, X, Y (see reference for MouseUp event in
'                                    MSDN Library for the description of the
'                                    arguments)
Public Event TitleBarMouseUp( _
               ByVal IdControl As Long, ByVal Button As Integer, _
               ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
             )
Attribute TitleBarMouseUp.VB_Description = "Occurs when the user releases a mouse button over a control title bar without previously moving the control"

'-------------------------------------------------------
' Subclassing Interface's Public Members Implementation
'-------------------------------------------------------
' Note: Refer to http://www.vbaccelerator.com/home/VB/Code/Libraries/Subclassing/
'       SSubTimer/article.asp for the description of the property and the
'       function below

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
  '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
  ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc( _
                   ByVal hwnd As Long, ByVal iMsg As Long, _
                   ByVal wParam As Long, ByVal lParam As Long _
                 ) As Long
  Select Case iMsg
    Case WM_ACTIVATE
      '-- This subclassing is used to handle the possibility of the user
      '   swithing to another application while dragging the splitter or while
      '   moving the control
      If wParam = WA_INACTIVE Then
        If mblnDragSplitter Then _
          picSplitter_MouseUp mScIndex, mScButton, mScShift, mScX, mScY
        If crecControl.Visible Then _
          ctbTitleBar_MoveEnd mScIndex, mScShift
      End If
    Case WM_SHOWWINDOW, WM_SIZE
      '-- In VB Splitter, developers need to add one line of code in their form
      '   resize event to call the Activate method. Now with this subclassing,
      '   there is no need to add any code to form to use basic features of
      '   Control Manager ActiveX Control
      Activate
      If iMsg = WM_SHOWWINDOW Then DetachMessage Me, mhwndParent, WM_SHOWWINDOW
  End Select
End Function

'--------------------------------------------
' ActiveX Control Constructor and Destructor
'--------------------------------------------

Private Sub UserControl_Initialize()
  Set mControls = New clsControls
  Set mSplitters = New clsSplitters
End Sub

Private Sub UserControl_Terminate()
  Set mControls = Nothing
  Set mSplitters = Nothing
  
  DetachMessage Me, mhwndParent, WM_SIZE
End Sub

'---------------------------
' ActiveX Control About Box
'---------------------------

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
  dlgAbout.Show vbModal
  Unload dlgAbout
  Set dlgAbout = Nothing
End Sub

'-----------------------------------
' ActiveX Control Properties Events
'-----------------------------------

Private Sub UserControl_InitProperties()
  mblnFillContainer = mconDefaultFillContainer
  mlngMarginBottom = mconDefaultMarginBottom
  mlngMarginLeft = mconDefaultMarginLeft
  mlngMarginRight = mconDefaultMarginRight
  mlngMarginTop = mconDefaultMarginTop
  mSplitters.BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    mSplitters.ActiveColor = _
      .ReadProperty(Name:=mconActiveColor, _
                    DefaultValue:=mSplitters.DefaultActiveColor)
    mSplitters.BackColor = .ReadProperty(Name:=mconBackColor, _
                                         DefaultValue:=Ambient.BackColor)
    mSplitters.ClipCursor = _
      .ReadProperty(Name:=mconClipCursor, _
                    DefaultValue:=mSplitters.DefaultClipCursor)
    mSplitters.Enable = .ReadProperty(Name:=mconEnable, _
                                      DefaultValue:=mSplitters.DefaultEnable)
    mblnFillContainer = .ReadProperty(Name:=mconFillContainer, _
                                      DefaultValue:=mconDefaultFillContainer)
    mSplitters.LiveUpdate = _
      .ReadProperty(Name:=mconLiveUpdate, _
                    DefaultValue:=mSplitters.DefaultLiveUpdate)
    mlngMarginBottom = .ReadProperty(Name:=mconMarginBottom, _
                                     DefaultValue:=mconDefaultMarginBottom)
    mlngMarginLeft = .ReadProperty(Name:=mconMarginLeft, _
                                   DefaultValue:=mconDefaultMarginLeft)
    mlngMarginRight = .ReadProperty(Name:=mconMarginRight, _
                                    DefaultValue:=mconDefaultMarginRight)
    mlngMarginTop = .ReadProperty(Name:=mconMarginTop, _
                                  DefaultValue:=mconDefaultMarginTop)
    mSplitters.Size = .ReadProperty(Name:=mconSize, _
                                    DefaultValue:=mSplitters.DefaultSize)
    mControls.TitleBar_CloseVisible = _
      .ReadProperty(Name:=mconTitleBar_CloseVisible, _
                    DefaultValue:=mControls.DefaultTitleBar_CloseVisible)
    mControls.TitleBar_Height = _
      .ReadProperty(Name:=mconTitleBar_Height, _
                    DefaultValue:=mControls.DefaultTitleBar_Height)
    mControls.TitleBar_Visible = _
      .ReadProperty(Name:=mconTitleBar_Visible, _
                    DefaultValue:=mControls.DefaultTitleBar_Visible)
  End With
  
  gstrControlName = Ambient.DisplayName
  
  ' Hide the ActiveX control when initializing the controls in it to reduce the
  '   flickering
  If Ambient.UserMode Then
    mblnVisibleSave = Extender.Visible
    Extender.Visible = False
  End If

  mhwndParent = UserControl.Parent.hwnd
  If Ambient.UserMode Then
    AttachMessage Me, mhwndParent, WM_SHOWWINDOW
    AttachMessage Me, mhwndParent, WM_SIZE
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty Name:=mconActiveColor, _
                   Value:=mSplitters.ActiveColor, _
                   DefaultValue:=mSplitters.DefaultActiveColor
    .WriteProperty Name:=mconBackColor, Value:=mSplitters.BackColor, _
                   DefaultValue:=Ambient.BackColor
    .WriteProperty Name:=mconClipCursor, Value:=mSplitters.ClipCursor, _
                   DefaultValue:=mSplitters.DefaultClipCursor
    .WriteProperty Name:=mconEnable, Value:=mSplitters.Enable, _
                   DefaultValue:=mSplitters.DefaultEnable
    .WriteProperty Name:=mconFillContainer, Value:=mblnFillContainer, _
                   DefaultValue:=mconDefaultFillContainer
    .WriteProperty Name:=mconLiveUpdate, Value:=mSplitters.LiveUpdate, _
                   DefaultValue:=mSplitters.DefaultLiveUpdate
    .WriteProperty Name:=mconMarginBottom, Value:=mlngMarginBottom, _
                   DefaultValue:=mconDefaultMarginBottom
    .WriteProperty Name:=mconMarginLeft, Value:=mlngMarginLeft, _
                   DefaultValue:=mconDefaultMarginLeft
    .WriteProperty Name:=mconMarginRight, Value:=mlngMarginRight, _
                   DefaultValue:=mconDefaultMarginRight
    .WriteProperty Name:=mconMarginTop, Value:=mlngMarginTop, _
                   DefaultValue:=mconDefaultMarginTop
    .WriteProperty Name:=mconSize, Value:=mSplitters.Size, _
                   DefaultValue:=mSplitters.DefaultSize
    .WriteProperty Name:=mconTitleBar_CloseVisible, _
                   Value:=mControls.TitleBar_CloseVisible, _
                   DefaultValue:=mControls.DefaultTitleBar_CloseVisible
    .WriteProperty Name:=mconTitleBar_Height, _
                   Value:=mControls.TitleBar_Height, _
                   DefaultValue:=mControls.DefaultTitleBar_Height
    .WriteProperty Name:=mconTitleBar_Visible, _
                   Value:=mControls.TitleBar_Visible, _
                   DefaultValue:=mControls.DefaultTitleBar_Visible
  End With
End Sub

'-------------------------------
' Others ActiveX Control Events
'-------------------------------

' Purpose    : Raises custom event TitleBarClick
' Input      : Index
' Effect     : As specified
Private Sub ctbTitleBar_Click(Index As Integer)
  RaiseEvent TitleBarClick(CLng(Index))
End Sub

' Purpose    : Closes the control at run-time, re-arranges the other controls
'              and raises ControlBeforeClose and ControlAfterClose event
' Effect     : See the codes
' Input      : Index (the id of the control which will be closed)
Private Sub ctbTitleBar_CloseClick(Index As Integer)
  Dim blnCancel As Boolean

  blnCancel = False
  RaiseEvent ControlBeforeClose(Index, blnCancel)
  If Not blnCancel Then
    CloseControl IdControl:=Index
    RaiseEvent ControlAfterClose(Index)
  End If
End Sub

' Purpose    : Raises custom event TitleBarDblClick
' Input      : Index
' Effect     : As specified
Private Sub ctbTitleBar_DblClick(Index As Integer)
  RaiseEvent TitleBarDblClick(CLng(Index))
End Sub

' Purpose    : Raises custom event TitleBarMouseDown
' Inputs     : Index, Button, Shift, x, y
' Effect     : As specified
Private Sub ctbTitleBar_MouseDown( _
              Index As Integer, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  muposPrev.X = X
  muposPrev.Y = Y
  
  RaiseEvent TitleBarMouseDown(Index, Button, Shift, X, Y)
End Sub

' Purpose    : Raises custom event TitleBarMouseMove
' Inputs     : Index, Button, Shift, x, y
' Effect     : As specified
Private Sub ctbTitleBar_MouseMove( _
              Index As Integer, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  If Not mblnControlMoved And ((X <> muposPrev.X) Or (Y <> muposPrev.Y)) Then _
    RaiseEvent TitleBarMouseMove(Index, Button, Shift, X, Y)
End Sub

' Purpose    : Raises custom event TitleBarMouseUp
' Inputs     : Index, Button, Shift, x, y
' Effect     : As specified
Private Sub ctbTitleBar_MouseUp( _
              Index As Integer, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  If Not mblnControlMoved Then _
    RaiseEvent TitleBarMouseUp(Index, Button, Shift, X, Y)
End Sub

' Purpose    : Moves the control at run-time and raises ControlMove event
' Effects    : * If the cursor is on a splitter which doesn't belong to the
'                control or the cursor is on the edge of the ControlManager
'                control, then the drop guider rectangle has been shown
'              * Otherwise, the guider rectangle position has been adjusted
'                based on the cursor position
' Inputs     : * Index (the id of the control which will be moved)
'              * Shift (an integer that corresponds to the state of the SHIFT,
'                       CTRL, and ALT keys)
Private Sub ctbTitleBar_Move(Index As Integer, ByVal Shift As Integer)
  Dim blnShowDropRect As Boolean             'indicating whether the drop guider
                                             '         rectangle should be shown
  Dim IdCtl As Long
  Dim IdSpl As Long
  Dim udeTargetType As genmMoveDestination
  Dim uposCursor As mdlAPI.POINTAPI    'indicating the current cursor's position
                                       ' relative to the Control Manager control
  Dim urecDrop As mdlAPI.RECT              'indicating the drop guider rectangle
                                           '               size and and position
  
  mblnControlMoved = (crecControl.Left <> mControls(Index).Left) Or _
                     (crecControl.Top <> mControls(Index).Top) Or _
                     (crecControl.Width <> mControls(Index).Width) Or _
                     (crecControl.Height <> mControls(Index).Height)
  GetDropTarget blnTargetValid:=blnShowDropRect, _
                udeTargetType:=udeTargetType, lngIdCtl:=IdCtl, lngIdSpl:=IdSpl
  If blnShowDropRect Then _
    blnShowDropRect = blnShowDropRect And _
                      Not IsRectNearSource(IdCtlSource:=CLng(Index), _
                                           IdCtlDestination:=IdCtl, _
                                           IdSplDestination:=IdSpl, _
                                           udeTargetType:=udeTargetType)
  If blnShowDropRect Then
    uposCursor = GetCursorRelPos(UserControl.hwnd)
    urecDrop = GetDropRect(IdCtlSource:=CLng(Index), IdCtlDestination:=IdCtl, _
                           IdSplDestination:=IdSpl, _
                           udeTargetType:=udeTargetType)
    ' urecDrop.Left = gconUninitialized means that the drop guider rectangle's
    '   size is bigger than the minimum size of the control
    blnShowDropRect = urecDrop.Left <> gconUninitialized
  End If
  If blnShowDropRect Then
    '-- Show the drop guider rectangle
    crecControl.Move urecDrop.Left, urecDrop.Top, _
                     urecDrop.Right - urecDrop.Left, _
                     urecDrop.Bottom - urecDrop.Top
  Else
    '-- Update the guider rectangle position based on the cursor position
    crecControl.UpdatePosition
  End If
  
  If mblnControlMoved Then _
   RaiseEvent ControlMove(Index, Shift, crecControl.Left, crecControl.Top, _
                          crecControl.Width, crecControl.Height)
End Sub

' Purpose    : Initializes all things needed to move the control at run-time
' Effect     : The guider rectangle has been shown
' Inputs     : * Index (the id of the control which will be moved)
'              * Shift (an integer that corresponds to the state of the SHIFT,
'                       CTRL, and ALT keys)
Private Sub ctbTitleBar_MoveBegin(Index As Integer, ByVal Shift As Integer)
  ' This subclassing below is used to handle the possibility of the user
  '   swithing to another application while dragging the splitter
  mScIndex = Index
  mScShift = Shift
  AttachMessage Me, mhwndRoot, WM_ACTIVATE
  
  With mControls(Index)
    crecControl.Move .Left, .Top, .Width, .Height
  End With
  crecControl.ZOrder
  crecControl.Visible = True
  
  RaiseEvent ControlMoveBegin(Index, Shift)
End Sub

' Purpose    : Ends the run-time control move action
' Effect     : * The guider rectangle has been hidden
'              * If the drop target is valid, the control has been moved and
'                the other controls position and size have been re-arranged
' Inputs     : * Index (the id of the control which will be moved)
'              * Shift (an integer that corresponds to the state of the SHIFT,
'                       CTRL, and ALT keys)
Private Sub ctbTitleBar_MoveEnd(Index As Integer, ByVal Shift As Integer)
  Dim blnSuccess As Boolean     'indicating whether the move action is succesful
  
  ' Variables for GetDropTarget parameters
  Dim blnTargetValid As Boolean
  Dim IdCtl As Long
  Dim IdSpl As Long
  Dim udeTargetType As genmMoveDestination
  
  DetachMessage Me, mhwndRoot, WM_ACTIVATE
  
  GetDropTarget blnTargetValid:=blnTargetValid, _
                udeTargetType:=udeTargetType, lngIdCtl:=IdCtl, lngIdSpl:=IdSpl
  If blnTargetValid Then _
    blnTargetValid = blnTargetValid And _
                     Not IsRectNearSource(IdCtlSource:=CLng(Index), _
                                          IdCtlDestination:=IdCtl, _
                                          IdSplDestination:=IdSpl, _
                                          udeTargetType:=udeTargetType)
  If blnTargetValid Then
    MoveControl IdControlSource:=CLng(Index), MoveTo:=udeTargetType, _
                IdControlDestination:=IdCtl, IdSplitterDestination:=IdSpl, _
                Success:=blnSuccess
  Else
    blnSuccess = False
    crecControl.Visible = False
  End If
  
  If mblnControlMoved Then _
    RaiseEvent ControlMoveEnd(Index, Shift, blnTargetValid And blnSuccess)
  mblnControlMoved = False
End Sub

' Purpose    : Refreshes the control's title bar close button visibility
' Input      : IdControl (a value that uniquely identifies a control)
' Effect     : As specified
Private Sub mControls_TitleBarCloseVisibleChange(ByVal IdControl As Long)
  If IdControl <> gconUninitialized Then _
    ctbTitleBar(IdControl).CloseVisible = _
      mControls(IdControl).TitleBar_CloseVisible
End Sub

' Purpose    : Refreshes the control's title bar visibility
' Input      : IdControl (a value that uniquely identifies a control)
' Effects    : * As specified
'              * The maximum and minimum value of the corresponding splitters
'                have been adjusted
Private Sub mControls_TitleBarVisibleChange(ByVal IdControl As Long)
  If IdControl <> gconUninitialized Then
    With mControls(IdControl)
      If .TitleBar_Visible Then
        If .IdSplTop <> gconUninitialized Then _
          mSplitters(.IdSplTop).MaxYc = mSplitters(.IdSplTop).MaxYc - _
                                        .TitleBar_VisibleHeight
        If .IdSplBottom <> gconUninitialized Then _
          mSplitters(.IdSplBottom).MinYc = mSplitters(.IdSplBottom).MinYc + _
                                           .TitleBar_VisibleHeight
      Else
        If .IdSplTop <> gconUninitialized Then _
          mSplitters(.IdSplTop).MaxYc = mSplitters(.IdSplTop).MaxYc + _
                                        .TitleBar_VisibleHeight
        If .IdSplBottom <> gconUninitialized Then _
          mSplitters(.IdSplBottom).MinYc = mSplitters(.IdSplBottom).MinYc - _
                                           .TitleBar_VisibleHeight
      End If
    End With
  
    Refresh
  End If
End Sub

' Purpose    : Refreshes the control back color to match the splitter's back
'              color change
' Input      : IdSplitter (a value that uniquely identifies a splitter)
' Effect     : As specified
Private Sub mSplitters_BackColorChange(ByVal IdSplitter As Long)
  If IdSplitter <> gconUninitialized Then _
    picSplitter(IdSplitter).BackColor = mSplitters(IdSplitter).BackColor
End Sub

' Purpose    : Refreshes the splitter back color to match the new property value
' Input      : IdSplitter (a value that uniquely identifies a splitter)
' Effect     : As specified
Private Sub mSplitters_EnableChange(ByVal IdSplitter As Long)
  If IdSplitter <> gconUninitialized Then _
    picSplitter(IdSplitter).Enabled = mSplitters(IdSplitter).Enable
End Sub

' Purpose    : Raises custom event SplitterClick
' Input      : Index
' Effect     : As specified
Private Sub picSplitter_Click(Index As Integer)
  RaiseEvent SplitterClick(CLng(Index))
End Sub

' Purpose    : Raises custom event SplitterDblClick
' Input      : Index
' Effect     : As specified
Private Sub picSplitter_DblClick(Index As Integer)
  RaiseEvent SplitterDblClick(CLng(Index))
End Sub

' Purpose    : Initializes all things needed to move the splitter at run-time
'              and raises custom event SplitterMouseDown and SplitterMoveBegin
' Assumption : Picture Box control picSplitter(Index) which represents the
'              splitter exits
' Effects    : * mblnDrag = true
'              * mlngDragStart = x or y (see the codes)
'              * Control picSplitter(Index) is in front of the other controls
'              * If the splitter's LiveUpdate property is false, then the
'                picSpliter(Index) BackColor property has been set to the
'                splitter's ActiveColor property
'              * If the splitter's ClipCursor property is true, then the mouse
'                pointer has been confined based on the splitter's MinXc, MinYc,
'                MaxXc and MaxYc property value
'              * Custom event SplitterMouseDown has been raised
'              * If the user presses the left-button, then the SplitterMoveBegin
'                event has been raised
' Inputs     : Index, Button, Shift, X, Y
' Note       : Notes that this procedure may confine the mouse pointer to
'              certain area in the screen. If you call this procedure, don't
'              forget to free the mouse pointer afterwards using
'              mdlAPI.ClipCursorClear function.
Private Sub picSplitter_MouseDown(Index As Integer, Button As Integer, _
                                  Shift As Integer, X As Single, Y As Single)
  Dim uposCursor As mdlAPI.POINTAPI                  'another variable needed to
                                                     ' confine the mouse pointer
  Dim urecClipCursor As mdlAPI.RECT          'the rectangle area where the mouse
                                             '         pointer would be confined
  
  If Button = vbLeftButton Then
    ' This subclassing below is used to handle the possibility of the user
    '   swithing to another application while dragging the splitter
    mScIndex = Index
    mScButton = Button
    mScShift = Shift
    mScX = X
    mScY = Y
    AttachMessage Me, mhwndRoot, WM_ACTIVATE
    
    mblnDragSplitter = True
    Select Case mSplitters(Index).Orientation
      Case orHorizontal
        mlngDragStart = Y
      Case orVertical
        mlngDragStart = X
    End Select
    picSplitter(Index).ZOrder
    
    If Not mSplitters(Index).LiveUpdate Then
      picSplitter(Index).BackColor = mSplitters(Index).ActiveColor
      UserControl.BackColor = mSplitters(Index).BackColor
    End If
    
    If mSplitters(Index).ClipCursor Then
      mdlAPI.GetCursorPos uposCursor
      uposCursor.X = (uposCursor.X * Screen.TwipsPerPixelX) - _
                     (picSplitter(Index).Left + X)
      uposCursor.Y = (uposCursor.Y * Screen.TwipsPerPixelY) - _
                     (picSplitter(Index).Top + Y)
      With urecClipCursor
        Select Case mSplitters(Index).Orientation
          Case orHorizontal
            .Top = (uposCursor.Y + mSplitters(Index).MinYc) \ _
                   Screen.TwipsPerPixelY
            .Right = (uposCursor.X + mSplitters(Index).Right) \ _
                     Screen.TwipsPerPixelX
            .Bottom = (uposCursor.Y + mSplitters(Index).MaxYc) \ _
                      Screen.TwipsPerPixelY
            .Left = (uposCursor.X + mSplitters(Index).Left) \ _
                    Screen.TwipsPerPixelX
          Case orVertical
            .Top = (uposCursor.Y + mSplitters(Index).Top) \ _
                   Screen.TwipsPerPixelY
            .Right = (uposCursor.X + mSplitters(Index).MaxXc) \ _
                     Screen.TwipsPerPixelX
            .Bottom = (uposCursor.Y + mSplitters(Index).Bottom) \ _
                      Screen.TwipsPerPixelY
            .Left = (uposCursor.X + mSplitters(Index).MinXc) \ _
                    Screen.TwipsPerPixelX
        End Select
      End With
      mdlAPI.ClipCursor urecClipCursor
    End If
    
    RaiseEvent SplitterMoveBegin(CLng(Index), Shift, X, Y)
  End If
  
  muposPrev.X = X
  muposPrev.Y = Y
  
  RaiseEvent SplitterMouseDown(CLng(Index), Button, Shift, X, Y)
End Sub

' Purpose    : Moves the splitter at run-time and raises custom event
'              SplitterMouseMove or SplitterMove
' Assumption : The picSplitter_MouseDown procedure has been called
' Effects    : * If the user moves the splitter, custom event Moving has been
'                raised
'              * Otherwise, custom event MouseMove has been raised
'              * Other effect, as specified
' Inputs     : Index, Button, Shift, x, y
Private Sub picSplitter_MouseMove(Index As Integer, Button As Integer, _
                                  Shift As Integer, X As Single, Y As Single)
Attribute picSplitter_MouseMove.VB_Description = "Moves the splitter at run time"
  Dim blnSplitterMoved As Boolean      'indicating whether the splitter is moved
  Dim lngPos As Long              'to determine where the splitter will be moved
  
  Select Case mSplitters(Index).Orientation
    Case orHorizontal
      blnSplitterMoved = mblnDragSplitter And (Y <> mlngDragStart)
    Case orVertical
      blnSplitterMoved = mblnDragSplitter And (X <> mlngDragStart)
  End Select
  If blnSplitterMoved = True Then mblnSplitterMoved = True
  
  If blnSplitterMoved Then
    Select Case mSplitters(Index).Orientation
      Case orHorizontal
        lngPos = picSplitter(Index).Top + (Y - mlngDragStart)
        If (lngPos < mSplitters(Index).MinYc) Then
          lngPos = mSplitters(Index).MinYc
        ElseIf _
           (lngPos + picSplitter(Index).Height > mSplitters(Index).MaxYc) Then
          lngPos = mSplitters(Index).MaxYc - picSplitter(Index).Height
        End If
        picSplitter(Index).Top = lngPos
        If mSplitters(Index).LiveUpdate Then _
          MoveSplitter IdSplitter:=CLng(Index), _
                       MoveTo:=picSplitter(Index).Top + _
                               (picSplitter(Index).Height \ 2)
      Case orVertical
        lngPos = picSplitter(Index).Left + (X - mlngDragStart)
        If (lngPos < mSplitters(Index).MinXc) Then
          lngPos = mSplitters(Index).MinXc
        ElseIf _
           (lngPos + picSplitter(Index).Width > mSplitters(Index).MaxXc) Then
          lngPos = mSplitters(Index).MaxXc - picSplitter(Index).Width
        End If
        picSplitter(Index).Left = lngPos
        If mSplitters(Index).LiveUpdate Then _
          MoveSplitter IdSplitter:=CLng(Index), _
                       MoveTo:=picSplitter(Index).Left + _
                               (picSplitter(Index).Width \ 2)
    End Select
  End If
  
  If Not mblnDragSplitter And Not blnSplitterMoved And _
     ((X <> muposPrev.X) Or (Y <> muposPrev.Y)) Then
    RaiseEvent SplitterMouseMove(CLng(Index), Button, Shift, X, Y)
  ElseIf blnSplitterMoved Then
    RaiseEvent SplitterMove(CLng(Index), Shift, X, Y)
  End If
  
  muposPrev.X = X
  muposPrev.Y = Y
End Sub

' Purpose    : Ends the run-time splitter move action and raises custom event
'              SplitterMouseUp or SplitterMoveEnd
' Assumption : Picture Box control picSplitter(Index) which represents the
'              splitter exits
' Effects    : * mblnDrag = false
'              * Control picSplitter(Index) is in front of the other controls
'              * If the splitter's LiveUpdate property is false, then the
'                picSpliter(Index) BackColor property has been set to the
'                splitter's BackColor property
'              * The splitters minimum and maximum x- and y- coordinates have
'                been adjusted
'              * If the splitter's ClipCursor property is true, then the mouse
'                pointer has been freed from confinement
'              * If the splitter was moved then custom event Moved has been
'                raised, otherwise, custom event MouseUp has been raised
' Inputs     : Index, Button, Shift, x, y
Private Sub picSplitter_MouseUp(Index As Integer, Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
Attribute picSplitter_MouseUp.VB_Description = "Ends the run time splitter move action"
  DetachMessage Me, mhwndRoot, WM_ACTIVATE
    
  mblnDragSplitter = False
  If Not mSplitters(Index).LiveUpdate Then _
    picSplitter(Index).BackColor = mSplitters.BackColor
  Select Case mSplitters(Index).Orientation
    Case orHorizontal
      MoveSplitter IdSplitter:=CLng(Index), _
                   MoveTo:=picSplitter(Index).Top + _
                           (picSplitter(Index).Height \ 2)
    Case orVertical
      MoveSplitter IdSplitter:=CLng(Index), _
                   MoveTo:=picSplitter(Index).Left + _
                           (picSplitter(Index).Width \ 2)
  End Select
  If mSplitters(Index).ClipCursor Then mdlAPI.ClipCursorClear

  If mblnSplitterMoved Then
    RaiseEvent SplitterMoveEnd(CLng(Index), Shift, X, Y)
    mblnSplitterMoved = False
  Else
    RaiseEvent SplitterMouseUp(CLng(Index), Button, Shift, X, Y)
  End If
End Sub

' Purpose    : Adjusts the components inside the control to agree with the
'              control's size
' Effect     : See the codes below and see the effects of BuildManager and
'              Stretch procedures
Private Sub UserControl_Resize()
Attribute UserControl_Resize.VB_Description = "Adjusts the components inside the control to agree with the control's size"
  If ContainedControls.Count > 0 Then
    If (mControls.Count = 0) Or Not Ambient.UserMode Then
      BuildManager
      If Ambient.UserMode Then Extender.Visible = mblnVisibleSave
    Else
      Stretch
    End If
  End If
End Sub

'-----------------------------
' ActiveX Control Collections
'-----------------------------

Public Property Get Controls() As clsControls
Attribute Controls.VB_Description = "Returns a collection whose elements represent each virtual control in a Control Manager object"
  Set Controls = mControls
End Property

Public Property Get Splitters() As clsSplitters
Attribute Splitters.VB_Description = "Returns a collection whose elements represent each virtual splitter in a Control Manager object"
  Set Splitters = mSplitters
End Property

'----------------------------
' ActiveX Control Properties
'----------------------------

' Purpose    : Sets the background color used to display a splitter when the
'              user moves it in none live update mode
' Effect     : As specified
' Input      : lngActiveColor (the new ActiveColor property value)
Public Property Let ActiveColor(lngActiveColor As OLE_COLOR)
Attribute ActiveColor.VB_Description = "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
  mSplitters.ActiveColor = lngActiveColor
  PropertyChanged mconActiveColor
End Property

' Purpose    : Returns the background color used to display a splitter when the
'              user moves it in none live update mode
' Return     : As specified
Public Property Get ActiveColor() As OLE_COLOR
  ActiveColor = mSplitters.ActiveColor
End Property

' Purpose    : Sets the background color used to display all splitters
' Effect     : As specified
' Input      : lngBackColor (the new BackColor property value)
Public Property Let BackColor(lngBackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Returns/sets the background color used to display all splitters"
  mSplitters.BackColor = lngBackColor
  PropertyChanged mconBackColor
  
  Refresh
End Property

' Purpose    : Returns the background color used to display all splitters
' Return     : As specified
Public Property Get BackColor() As OLE_COLOR
  BackColor = mSplitters.BackColor
End Property

' Purpose    : Sets a value that determines whether the mouse pointer is
'              confined to the virtual splitter minimum and maximum x- and
'              y-coordinate when the user moves a splitter
' Effect     : As specified
' Input      : blnClipCursor (the new ClipCursor property value)
Public Property Let ClipCursor(ByVal blnClipCursor As Boolean)
Attribute ClipCursor.VB_Description = "Returns/sets a value that determines whether the mouse pointer is confined to the virtual splitter minimum and maximum x- and y-coordinate when the user moves a splitter"
  mSplitters.ClipCursor = blnClipCursor
  PropertyChanged mconClipCursor
End Property

' Purpose    : Returns a value that determines whether the mouse pointer is
'              confined to the virtual splitter minimum and maximum x- and
'              y-coordinate when the user moves a splitter
' Return     : As specified
Public Property Get ClipCursor() As Boolean
  ClipCursor = mSplitters.ClipCursor
End Property

' Purpose    : Sets a value that determines whether all splitters are movable
' Effect     : As specified
' Input      : blnEnable (the new Enable property value)
Public Property Let Enable(ByVal blnEnable As Boolean)
Attribute Enable.VB_Description = "Returns/sets a value that determines whether all splitters are movable"
  mSplitters.Enable = blnEnable
  PropertyChanged mconEnable
  
  Refresh
End Property

' Purpose    : Returns a value that determines whether all splitters are movable
' Return     : As specified
Public Property Get Enable() As Boolean
  Enable = mSplitters.Enable
End Property

' Purpose    : Sets a value that determines whether the ActiveX Control (along
'              with all controls inside it) will automatically adjust its size
'              to fill-up its container with respect to the margin properties
' Effect     : As specified
' Input      : blnFillContainer (the new FillContainer property value)
Public Property Let FillContainer(ByVal blnFillContainer As Boolean)
Attribute FillContainer.VB_Description = "Returns/sets a value that determines whether the ActiveX Control (along with all controls inside it) will automatically adjust its size to fill-up its container with respect to the margin properties"
  mblnFillContainer = blnFillContainer
  PropertyChanged mconFillContainer

  Activate
End Property

' Purpose    : Returns a value that determines whether the ActiveX Control
'              (along with all controls inside it) will automatically adjust its
'              size to fill-up its container with respect to the margin
'              properties
' Return     : As specified
Public Property Get FillContainer() As Boolean
  FillContainer = mblnFillContainer
End Property

' Purpose    : Sets a value that determines whether the controls should be
'              resized as a splitter is moved
' Effect     : As specified
' Input      : blnLiveUpdate (the new LiveUpdate property value)
Public Property Let LiveUpdate(ByVal blnLiveUpdate As Boolean)
Attribute LiveUpdate.VB_Description = "Returns/sets a value that determines whether the controls should be resized as a splitter is moved"
  mSplitters.LiveUpdate = blnLiveUpdate
  PropertyChanged mconLiveUpdate
End Property

' Purpose    : Returns a value that determines whether the controls should be
'              resized as a splitter is moved
' Return     : As specified
Public Property Get LiveUpdate() As Boolean
  LiveUpdate = mSplitters.LiveUpdate
End Property

' Purpose    : Sets the bottom margin of the ActiveX Control from its container
' Effect     : As specified
' Input      : lngMarginBottom (the new MarginBottom property value)
Public Property Let MarginBottom(ByVal lngMarginBottom As Long)
Attribute MarginBottom.VB_Description = "Returns/sets the bottom margin of the ActiveX Control from its container"
  mlngMarginBottom = lngMarginBottom
  PropertyChanged mconMarginBottom
  
  Activate
End Property

' Purpose    : Returns the bottom margin of the ActiveX Control from its
'              container
' Return     : As specified
Public Property Get MarginBottom() As Long
  MarginBottom = mlngMarginBottom
End Property

' Purpose    : Sets the left margin of the ActiveX Control from its container
' Effect     : As specified
' Input      : lngMarginLeft (the new MarginLeft property value)
Public Property Let MarginLeft(ByVal lngMarginLeft As Long)
Attribute MarginLeft.VB_Description = "Returns/sets the left margin of the ActiveX Control from its container"
  mlngMarginLeft = lngMarginLeft
  PropertyChanged mconMarginLeft
  
  Activate
End Property

' Purpose    : Returns the left margin of the ActiveX Control from its container
' Return     : As specified
Public Property Get MarginLeft() As Long
  MarginLeft = mlngMarginLeft
End Property

' Purpose    : Sets the right margin of the ActiveX Control from its container
' Effect     : As specified
' Input      : lngMarginRight (the new MarginRight property value)
Public Property Let MarginRight(ByVal lngMarginRight As Long)
Attribute MarginRight.VB_Description = "Returns/sets the right margin of the ActiveX Control from its container"
  mlngMarginRight = lngMarginRight
  PropertyChanged mconMarginRight

  Activate
End Property

' Purpose    : Returns the right margin of the ActiveX Control from its
'              container
' Return     : As specified
Public Property Get MarginRight() As Long
  MarginRight = mlngMarginRight
End Property

' Purpose    : Sets the top margin of the ActiveX Control from its container
' Effect     : As specified
' Input      : lngMarginTop (the new MarginTop property value)
Public Property Let MarginTop(ByVal lngMarginTop As Long)
Attribute MarginTop.VB_Description = "Returns/sets the top margin of the ActiveX Control from its container"
  mlngMarginTop = lngMarginTop
  PropertyChanged mconMarginTop
  
  Activate
End Property

' Purpose    : Returns the top margin of the ActiveX Control from its container
' Return     : As specified
Public Property Get MarginTop() As Long
  MarginTop = mlngMarginTop
End Property

' Purpose    : Sets the size of all splitters
' Effects    : * If Size is smaller than the splitters' minimum size then the
'                splitters' size has been set to their minimum size
'              * If there is a control with size less than its minimum size
'                then the error message has been raised
'              * Otherwise, as specified
' Input      : lngSize (the new Size property value)
Public Property Let Size(ByVal lngSize As Long)
Attribute Size.VB_Description = "Returns/sets the size of all splitters"
  Dim blnNeedToShow As Boolean
  Dim lngdSize As Long
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim oId As clsId                     'for enumerating all Id in Ids collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
  Dim ospl2 As clsSplitter                'for enumerating all virtual splitters
                                          '              in Splitters collection

  If lngSize < mSplitters.MinimumSize Then lngSize = mSplitters.MinimumSize
  lngdSize = lngSize - mSplitters.Size
  mSplitters.Size = lngSize
  PropertyChanged mconSize
  
  '-- Refresh the splitter size
  blnNeedToShow = (mControls.Count = 0) And Ambient.UserMode
  
  For Each octl In mControls
    If Not octl.Closed Then
      octl.Left = octl.Left + IIf(octl.Left <> mControls.Left, lngdSize \ 2, 0)
      octl.Top = octl.Top + IIf(octl.Top <> mControls.Top, lngdSize \ 2, 0)
      octl.Right = octl.Right - _
                   IIf(octl.Right <> mControls.Right, lngdSize \ 2, 0)
      octl.Bottom = octl.Bottom - _
                    IIf(octl.Bottom <> mControls.Bottom, lngdSize \ 2, 0)
    End If
  Next
  For Each ospl In mSplitters
    Select Case ospl.Orientation
      Case orHorizontal
        ospl.Height = lngSize
        
        '-- Adjust the width of the splitter if necessary
        If ospl.Left > 0 Then ospl.Left = ospl.Left + (lngdSize \ 2)
        If ospl.Right < Extender.Width Then _
          ospl.Right = ospl.Right - (lngdSize \ 2)
        
        '-- Adjust the minimum value of the splitter if necessary
        For Each oId In ospl.IdsCtlTop
          If mControls(oId).IdSplTop <> gconUninitialized Then
            ospl.MinYc = ospl.MinYc + (lngdSize \ 2)
            Exit For
          End If
        Next
        
        '-- Adjust the maximum value of the splitter if necessary
        For Each oId In ospl.IdsCtlBottom
          If mControls(oId).IdSplBottom <> gconUninitialized Then
            ospl.MaxYc = ospl.MaxYc - (lngdSize \ 2)
            Exit For
          End If
        Next
        
      Case orVertical
        ospl.Width = lngSize
        
        '-- Adjust the height of the splitter if necessary
        If ospl.Top > 0 Then ospl.Top = ospl.Top + (lngdSize \ 2)
        If ospl.Bottom < Extender.Height Then _
          ospl.Bottom = ospl.Bottom - (lngdSize \ 2)
        
        '-- Adjust the minimum value of the splitter if necessary
        For Each oId In ospl.IdsCtlLeft
          If mControls(oId).IdSplLeft <> gconUninitialized Then
            ospl.MinXc = ospl.MinXc + (lngdSize \ 2)
            Exit For
          End If
        Next
        
        '-- Adjust the maximum value of the splitter if necessary
        For Each oId In ospl.IdsCtlRight
          If mControls(oId).IdSplRight <> gconUninitialized Then
            ospl.MaxXc = ospl.MaxXc - (lngdSize \ 2)
            Exit For
          End If
        Next
    End Select
  Next
  If mControls.IsValid Then
    Refresh
  Else
    RaiseError udeErrNumber:=errResizeSplitter, strSource:="Size"
  End If
  
  If blnNeedToShow Then Extender.Visible = mblnVisibleSave
End Property

' Purpose    : Returns the size of all splitters
' Return     : As specified
Public Property Get Size() As Long
  Size = mSplitters.Size
End Property

' Purpose    : Sets a value that determines whether a close button in all
'              control title bars is visible
' Effect     : As specified
' Input      : blnTitleBar_CloseVisible (the new TitleBar_CloseVisible property
'                                        value)
Public Property Let TitleBar_CloseVisible( _
                       ByVal blnTitleBar_CloseVisible As Boolean _
                     )
Attribute TitleBar_CloseVisible.VB_Description = "Returns/sets a value that determines whether a close button in all control title bars is visible"
  If mControls.TitleBar_CloseVisible <> blnTitleBar_CloseVisible Then
    mControls.TitleBar_CloseVisible = blnTitleBar_CloseVisible
    PropertyChanged mconTitleBar_CloseVisible
    
    Refresh
  End If
End Property

' Purpose    : Returns a value that determines whether a close button in all
'              control title bars is visible
' Return     : As specified
Public Property Get TitleBar_CloseVisible() As Boolean
  TitleBar_CloseVisible = mControls.TitleBar_CloseVisible
End Property

' Purpose    : Sets the height of all control title bars
' Input      : lngTitleBar_Height (the new TitleBar_Height property value)
Private Property Let TitleBar_Height(ByVal lngTitleBar_Height As Long)
  mControls.TitleBar_Height = lngTitleBar_Height
  PropertyChanged mconTitleBar_Height
End Property

' Purpose    : Returns the height of the visible part of all control title bars
' Return     : As specified
Public Property Get TitleBar_Height() As Long
Attribute TitleBar_Height.VB_Description = "Returns/sets the height of all control title bars"
  TitleBar_Height = mControls.TitleBar_Height
End Property

' Purpose    : Sets a value that determines whether all control title bars are
'              visible
' Effect     : As specified
' Input      : blnblnTitleBar_Visible (the new TitleBar_Visible property value)
Public Property Let TitleBar_Visible(ByVal blnTitleBar_Visible As Boolean)
Attribute TitleBar_Visible.VB_Description = "Returns/sets a value that determines whether all control title bars are visible"
  Dim blnSuccess As Boolean
  Dim lngd As Long
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
    
  If mControls.TitleBar_Visible <> blnTitleBar_Visible Then
    mControls.TitleBar_Visible = blnTitleBar_Visible
    If blnTitleBar_Visible Then
      lngd = mControls.TitleBar_VisibleHeight
    Else
      lngd = -mControls.TitleBar_VisibleHeight
    End If
    
    blnSuccess = True
    For Each octl In mControls
      If octl.TitleBar_Visible <> mControls.TitleBar_Visible Then
        blnSuccess = False
        Exit For
      End If
    Next
    If blnSuccess Then blnSuccess = blnSuccess And mControls.IsValid
    If blnSuccess Then
      PropertyChanged mconTitleBar_Visible
      For Each ospl In mSplitters
        If ospl.Orientation = orHorizontal Then
          If ospl.IdsCtlTop.Count > 0 Then ospl.MinYc = ospl.MinYc + lngd
          If ospl.IdsCtlBottom.Count > 0 Then ospl.MaxYc = ospl.MaxYc - lngd
        End If
      Next
      Refresh
    Else
      mControls.TitleBar_Visible = Not mControls.TitleBar_Visible
    End If
  End If
End Property

' Purpose    : Returns a value that determines whether all control title bars
'              are visible
' Return     : As specified
Public Property Get TitleBar_Visible() As Boolean
  TitleBar_Visible = mControls.TitleBar_Visible
End Property

'-------------------------
' ActiveX Control Methods
'-------------------------

' Purposes   : Activates and resize the control to meet its container size with
'              respect to the control's margin property and FillContainer
'              property
' Assumption : The parent of the control has ScaleWidth and ScaleHeight property
' Effect     : As specified
' Note       : This is the main method of the control. This method should be
'              called whenever its container is loaded. Also this method should
'              be called everytime its container's size is changed so that the
'              FillContainer property would work. If the container is forms,
'              this method should be called in the form's resize event.
Private Sub Activate()
Attribute Activate.VB_Description = "Activates and resize the control to meet its container size with respect to the control's margin property and FillContainer property"
  Dim lngWidth As Long                             'the new width of the control
  Dim lngHeight As Long                           'the new height of the control

  mhwndRoot = mdlAPI.GetAncestor(Extender.Container.hwnd, mdlAPI.GA_ROOT)
    
  If mblnFillContainer Then
    lngWidth = UserControl.Parent.ScaleWidth - mlngMarginRight - mlngMarginLeft
    If lngWidth < 0 Then lngWidth = 0
    lngHeight = UserControl.Parent.ScaleHeight - _
                mlngMarginBottom - mlngMarginTop
    If lngHeight < 0 Then lngHeight = 0
    Extender.Move mlngMarginLeft, mlngMarginTop, lngWidth, lngHeight
  Else
    If mControls.Count = 0 Then UserControl_Resize
  End If
End Sub

' Purpose    : Closes (hides) a control
' Effects    : * If successful, the control has been closed
'              * If control IdControl doesn't exist, a run-time error has been
'                generated
'              * otherwise, no effect
' Input      : IdControl (a value that uniquely identifies the control the
'                         developer want to close)
' Return     : Success (a returned value that determines whether the Close
'                       method is successful)
Public Sub CloseControl(ByVal IdControl As Long, _
                        Optional ByRef Success As Boolean)
Attribute CloseControl.VB_Description = "Closes (hides) a control"
  If Not mControls.IsExist(IdControl) Then
    Success = False
    SecureRaiseError udeErrNumber:=errIdControl, strSource:="CloseControl"
  Else
    If mControls(IdControl).Closed Then
      Success = True
    Else
      '-- Backup the controls position in case the Control Manager couldn't be
      '   rebuilt
      mControls.Backup
      
      '-- Close the virtual control IdControl
      mControls(IdControl).Closed = True
        
      '-- Re-arrange the other virtual controls
      mControls.Compact
      mControls.RemoveHoles
      
      Success = IsSolid(blnIncludeSplitter:=False)
      If Success Then
        '-- Rebuild the splitters and applies the virtual controls and splitters
        ContainedControls(IdControl).Visible = False
        ctbTitleBar(IdControl).Visible = False
        BuildManager blnNewControl:=False
      Else
        '-- Restore the controls' close status, position and size
        mControls(IdControl).Closed = False
        mControls.Restore
        Refresh
      End If
    End If
  End If
End Sub

' Purpose    : Moves a control to certain area
' Effects    : * If successful, the control has been moved
'              * If control IdControl or splitter IdSplitter doesn't exist, a
'                run-time error has been generated
'              * otherwise, no effect
' Inputs     : * IdControlSource (A value that uniquely identifies the source
'                                 control the developer want to move)
'              * MoveTo (A value indicating the area type where the source
'                        control will be moved to)
'              * IdControlDestination (A value that uniquely identifies the
'                                      destination control the developer want to
'                                      move the source control to. This input is
'                                      required if the area type indicated by
'                                      MoveTo input is a control.)
'              * IdSplitterDestination (A value that uniquely identifies the
'                                       splitter the developer want to move the
'                                       source control to. This input is
'                                       required only if the are type indicated
'                                       by MoveTo is a splitter)
' Return     : Success (a returned value that determines whether the MoveControl
'                       method is successful)
Public Sub MoveControl(ByVal IdControlSource As Long, _
                       ByVal MoveTo As genmMoveDestination, _
                       Optional ByVal IdControlDestination _
                                        As Long = gconUninitialized, _
                       Optional ByVal IdSplitterDestination _
                                        As Long = gconUninitialized, _
                       Optional ByRef Success As Boolean)
Attribute MoveControl.VB_Description = "Moves a control to certain area"
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim udeRemoveHeapDirection As genmRemoveHeapDirection
  Dim urecControlBackup() As mdlAPI.RECT    'backup of the Controls collection's
                                            'size and position in case the
                                            'Control Manager couldn't be rebuilt
  Dim urecDrop As mdlAPI.RECT              'indicating the drop guider rectangle
                                           '               size and and position
  
  If Not mControls.IsExist(IdControlSource) Then
    Success = False
    SecureRaiseError udeErrNumber:=errIdControl, strSource:="MoveControl"
  ElseIf (MoveTo = mdSplitter) And _
         (Not mSplitters.IsExist(IdSplitterDestination)) Then
    Success = False
    SecureRaiseError udeErrNumber:=errIdSplitter, strSource:="MoveControl"
  ElseIf mControls(IdControlSource).Closed Then
    Success = False
    SecureRaiseError udeErrNumber:=errMoveControlClosed, _
                     strSource:="MoveControl"
  Else
    urecDrop = GetDropRect(IdCtlSource:=IdControlSource, _
                           IdCtlDestination:=IdControlDestination, _
                           IdSplDestination:=IdSplitterDestination, _
                           udeTargetType:=MoveTo)
    If urecDrop.Left = gconUninitialized Then
      Success = False
      If Not crecControl.Visible Then _
        SecureRaiseError udeErrNumber:=errMoveControlRoom, _
                         strSource:="MoveControl"
      crecControl.Visible = False
    Else
      crecControl.Visible = False
      '-- Backup the controls position in case the Control Manager couldn't be
      '   rebuilt
      mControls.Backup
      
      '-- Move the virtual control IdControl
      mControls(IdControlSource).Left = urecDrop.Left
      mControls(IdControlSource).Top = urecDrop.Top
      mControls(IdControlSource).Right = urecDrop.Right
      mControls(IdControlSource).Bottom = urecDrop.Bottom
      
      '-- Re-arrange the other virtual controls
      Select Case MoveTo
        Case mdControlTop, mdControlBottom, mdEdgeTop, mdEdgeBottom
          udeRemoveHeapDirection = rhdVertical
        Case mdControlLeft, mdControlRight, mdEdgeLeft, mdEdgeRight
          udeRemoveHeapDirection = rhdHorizontal
        Case mdSplitter
          Select Case mSplitters(IdSplitterDestination).Orientation
            Case orHorizontal
              udeRemoveHeapDirection = rhdVertical
            Case orVertical
              udeRemoveHeapDirection = rhdHorizontal
          End Select
      End Select
      mControls.RemoveHeap IdCtl:=IdControlSource, blnMaintainSize:=True, _
                           udeRemoveDirection:=udeRemoveHeapDirection
      mControls.Compact
      mControls.RemoveHoles
      
      Success = IsSolid(blnIncludeSplitter:=False)
      If Success Then
        '-- Rebuild the splitters and applies the virtual controls and splitters
        BuildManager blnNewControl:=False
      Else
        '-- Restore the controls' position and size
        mControls.Restore
        Refresh
      End If
    End If
  End If
End Sub

' Purpose    : Moves a splitter to the specified x- or y- (depending on the
'              splitter's Orientation property) coordinate
' Effects    : * If successful, the control has been moved and all other
'                effected splitters and controls' minimum and maximum x- and y-
'                coordinates have been adjusted
'              * If splitter IdSplitter doesn't exist, a run-time error has been
'                generated
' Inputs     : * IdSplitter (a value that uniquely identifies the splitter the
'                            developer want to move)
'              * MoveTo (an integer value that specifies the x- or y- coordinate
'                        (depending on the splitter's Orientation property)
'                        where the splitter will be moved)
Public Sub MoveSplitter(IdSplitter As Long, MoveTo As Long)
Attribute MoveSplitter.VB_Description = "Moves a splitter to the specified x- or y- (depending on the splitter's Orientation property) coordinate"
  Dim lngId As Long       'to determines the new friend control for the splitter
  Dim oId As clsId                     'for enumerating all Id in Ids collection
  Dim oid2 As clsId                    'for enumerating all Id in Ids collection
  
  If Not mSplitters.IsExist(IdSplitter) Then
    SecureRaiseError udeErrNumber:=errIdSplitter, strSource:="MoveSplitter"
  Else
    With mSplitters(IdSplitter)
      Select Case .Orientation
        Case orHorizontal
          '-- If the destination coordinate is beyond the splitter's minimum or
          '   maximum value, generates a custom run-time error
          If (MoveTo < .MinYc) Or (MoveTo > .MaxYc) Then
            If Ambient.UserMode Then
              SecureRaiseError udeErrNumber:=errMoveSplitter, _
                               strSource:="MoveSplitter"
            Else
              Exit Sub
            End If
          End If
                
          '-- Move the splitter
          .Yc = MoveTo
          
          '-- Resize the controls and splitters that effected by the splitter
          '   movement
          For Each oId In .IdsCtlTop
            mControls(oId).Bottom = .Top
          Next
          For Each oId In .IdsCtlBottom
            mControls(oId).Top = .Bottom
          Next
          For Each oId In .IdsSplTop
            mSplitters(oId).Bottom = .Top
          Next
          For Each oId In .IdsSplBottom
            mSplitters(oId).Top = .Bottom
          Next
          
          '-- Finalizes the splitter movement by adjusting the minimum and
          '   maximum y- coordinates of the splitters above or below the active
          '   splitter
          If Not mblnDragSplitter Then
            For Each oId In .IdsCtlTop
              If mControls(oId).IdSplTop <> gconUninitialized Then
                lngId = gconUninitialized
                For Each oid2 _
                    In mSplitters(mControls(oId).IdSplTop).IdsCtlBottom
                  If lngId = gconUninitialized Then
                    lngId = oid2
                  ElseIf mControls(oid2).Height - mControls(oid2).MinHeight < _
                         mControls(lngId).Height - _
                           mControls(lngId).MinHeight Then
                    lngId = oid2
                  End If
                Next
                mSplitters(mControls(oId).IdSplTop).MaxYc = _
                  mControls(lngId).Bottom - mControls(lngId).MinHeight
                mSplitters(mControls(oId).IdSplTop).IdCtlFriendBottom = lngId
              End If
            Next
            For Each oId In .IdsCtlBottom
              If mControls(oId).IdSplBottom <> gconUninitialized Then
                lngId = gconUninitialized
                For Each oid2 _
                    In mSplitters(mControls(oId).IdSplBottom).IdsCtlTop
                  If lngId = gconUninitialized Then
                    lngId = oid2
                  ElseIf mControls(oid2).Height - mControls(oid2).MinHeight < _
                         mControls(lngId).Height - _
                           mControls(lngId).MinHeight Then
                    lngId = oid2
                  End If
                Next
                mSplitters(mControls(oId).IdSplBottom).MinYc = _
                  mControls(lngId).Top + mControls(lngId).MinHeight
                mSplitters(mControls(oId).IdSplBottom).IdCtlFriendTop = lngId
              End If
            Next
          End If
        Case orVertical
          '-- If the destination coordinate is beyond the splitter's minimum or
          '   maximum value, generates a custom run-time error
          If (MoveTo < .MinXc) Or (MoveTo > .MaxXc) Then
            If Ambient.UserMode Then
              SecureRaiseError udeErrNumber:=errMoveSplitter, _
                               strSource:="MoveSplitter"
            Else
              Exit Sub
            End If
          End If
          
          ' Move the splitter
          .Xc = MoveTo
          
          '-- Resize the controls and splitters that effected by the splitter
          '   movement
          For Each oId In .IdsCtlLeft
            mControls(oId).Right = .Left
          Next
          For Each oId In .IdsCtlRight
            mControls(oId).Left = .Right
          Next
          For Each oId In .IdsSplLeft
            mSplitters(oId).Right = .Left
          Next
          For Each oId In .IdsSplRight
            mSplitters(oId).Left = .Right
          Next
          
          '-- Finalizes the splitter movement by adjusting the minimum and
          '   maximum x- coordinates of the splitters above or below the active
          '   splitter
          If Not mblnDragSplitter Then
            For Each oId In .IdsCtlLeft
              If mControls(oId).IdSplLeft <> gconUninitialized Then
                lngId = gconUninitialized
                For Each oid2 _
                    In mSplitters(mControls(oId).IdSplLeft).IdsCtlRight
                  If lngId = gconUninitialized Then
                    lngId = oid2
                  ElseIf mControls(oid2).Width - mControls(oid2).MinWidth < _
                         mControls(lngId).Width - mControls(lngId).MinWidth Then
                    lngId = oid2
                  End If
                Next
                mSplitters(mControls(oId).IdSplLeft).MaxXc = _
                  mControls(lngId).Right - mControls(lngId).MinWidth
                mSplitters(mControls(oId).IdSplLeft).IdCtlFriendRight = lngId
              End If
            Next
            For Each oId In .IdsCtlRight
              If mControls(oId).IdSplRight <> gconUninitialized Then
                lngId = gconUninitialized
                For Each oid2 _
                    In mSplitters(mControls(oId).IdSplRight).IdsCtlLeft
                  If lngId = gconUninitialized Then
                    lngId = oid2
                  ElseIf mControls(oid2).Width - mControls(oid2).MinWidth < _
                         mControls(lngId).Width - mControls(lngId).MinWidth Then
                    lngId = oid2
                  End If
                Next
                mSplitters(mControls(oId).IdSplRight).MinXc = _
                  mControls(lngId).Left + mControls(lngId).MinWidth
                mSplitters(mControls(oId).IdSplRight).IdCtlFriendLeft = lngId
              End If
            Next
          End If
      End Select
    End With
    Refresh
  End If
End Sub

' Purpose    : Opens (shows) a control and docks it to the ActiveX Control
' Effects    : * If successful, the control has been opened
'              * If control IdControl doesn't exist, a run-time error has been
'                generated
'              * otherwise, no effect
' Inputs     : * IdControl (a value that uniquely identifies the control the
'                developer want to open)
'              * MaintainSize (a value that determines whether the method will
'                              not change the size and position of the control
'                              the developer want to open)
' Return     : Success (a returned value that determines whether the OpenControl
'                       method is successful)
Public Sub OpenControl(ByVal IdControl As Long, _
                       Optional MaintainSize As Boolean = False, _
                       Optional ByRef Success As Boolean)
Attribute OpenControl.VB_Description = "Opens (shows) a control and docks it to the ActiveX Control"
  If Not mControls.IsExist(IdControl) Then
    Success = False
    SecureRaiseError udeErrNumber:=errIdControl, strSource:="OpenControl"
  Else
    If Not mControls(IdControl).Closed Then
      Success = True
    Else
      '-- Backup the controls position in case the Control Manager couldn't be
      '   rebuilt
      mControls.Backup
      
      '-- Open the virtual control IdControl
      mControls(IdControl).Closed = False
      
      '-- Re-arrange the other virtual controls
      mControls.RemoveHeap IdCtl:=IdControl, blnMaintainSize:=MaintainSize
      mControls.Compact
      mControls.RemoveHoles
      
      Success = IsSolid(blnIncludeSplitter:=False)
      If Success Then
        '-- Rebuild the splitters and applies the virtual controls and splitters
        BuildManager blnNewControl:=False
        ContainedControls(IdControl).Visible = True
        ctbTitleBar(IdControl).Visible = True
      Else
        '-- Restore the controls' close status, position and size
        mControls(IdControl).Closed = True
        mControls.Restore
        Refresh
      End If
    End If
  End If
End Sub

'----------------------------------
' Private Functions and Procedures
'----------------------------------

' Purpose    : Returns the adjusted height of control ctl
' Inputs     : * ctl
'              * octl (the virtual control of control ctl)
' Note       : This function is used to avoid flickering effect in LiveUpdate
'              mode for list box control or other controls that inherit it
Private Function AdjustedHeight(ByVal ctl As Control, _
                                ByVal octl As clsControl) As Long
Attribute AdjustedHeight.VB_Description = "Returns the adjusted height of control ctl"
  Dim lngAdjustedHeight                                          'returned value
  Dim lngHeightFactor As Long           'the height of each item in the list box
  
  If (TypeOf ctl Is ListBox) Or _
     (TypeOf ctl Is DirListBox) Or (TypeOf ctl Is FileListBox) Then
    lngHeightFactor = _
      mdlAPI.SendMessage(ctl.hwnd, mdlAPI.LB_GETITEMHEIGHT, 0&, 0&) * _
      Screen.TwipsPerPixelY
    lngAdjustedHeight = (((octl.Height - octl.MinHeight) \ lngHeightFactor) * _
                         lngHeightFactor) + octl.MinHeight
  Else
    lngAdjustedHeight = octl.Height
  End If
  AdjustedHeight = lngAdjustedHeight
End Function

' Purpose    : Build virtual controls (may with its title bar) and splitters and
'              applies it to the real controls and splitters
' Effect     : * If successed, as specified
'              * Otherwise, the custom error message has been raised
Private Sub BuildManager(Optional ByVal blnNewControl As Boolean = True)
  '-- Backup properties value for Controls collection
  Dim blnTitleBarCloseVisibleSave As Boolean
  Dim blnTitleBarVisibleSave As Boolean
  
  Dim i As Long       'for iterating all control in ContainedControls collection
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim oId As clsId                     'for enumerating all Id in Ids collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
  
  '-- VB Control Manager control can't have another VB Control Manager control
  '   inside it
  If IsSelfContained Then _
    SecureRaiseError udeErrNumber:=errSelfContained, strSource:="Init"
 
  On Error GoTo ErrorHandler
  
  If blnNewControl Then
    '-- Create new Controls collection
    blnTitleBarCloseVisibleSave = mControls.TitleBar_CloseVisible
    blnTitleBarVisibleSave = mControls.TitleBar_Visible
    Set mControls = New clsControls
    mControls.TitleBar_CloseVisible = blnTitleBarCloseVisibleSave
    mControls.TitleBar_Visible = blnTitleBarVisibleSave
    
    '-- Make the virtual controls solid and fill-up the VB Control Manager
    '   control's container
    For i = 0 To ContainedControls.Count - 1
      mControls.Add cctl:=ContainedControls, IdCtl:=i
    Next
    With mControls
      .Left = 0
      .Top = 0
      .Right = UserControl.ScaleWidth
      .Bottom = UserControl.ScaleHeight
      .RemoveHeaps
      .Compact
      .RemoveHoles
      .Stretch
    End With
    '-- Creates the new ctlTitleBar control instances to represent the control's
    '   title bar
    For i = 0 To ctbTitleBar.Count - 2
      Unload ctbTitleBar(i)
    Next
    For Each octl In mControls
      Load ctbTitleBar(octl)
    Next
  End If

  '-- Build virtual splitters and place it as the virtual controls' "border"
  For Each octl In mControls
    octl.IdSplTop = gconUninitialized
    octl.IdSplRight = gconUninitialized
    octl.IdSplBottom = gconUninitialized
    octl.IdSplLeft = gconUninitialized
  Next
  With mSplitters
    .Left = 0
    .Top = 0
    .Right = UserControl.ScaleWidth
    .Bottom = UserControl.ScaleHeight
    .Clear
    For Each octl In mControls
      If Not octl.Closed Then .Add octl:=octl, octls:=mControls
    Next
    For Each ospl In mSplitters
      ospl.IdsSplTop.RemoveDeleted lngLastPos:=.Count
      ospl.IdsSplRight.RemoveDeleted lngLastPos:=.Count
      ospl.IdsSplBottom.RemoveDeleted lngLastPos:=.Count
      ospl.IdsSplLeft.RemoveDeleted lngLastPos:=.Count
    Next
    For Each ospl In mSplitters
      Select Case ospl.Orientation
        Case orHorizontal
          For Each oId In ospl.IdsSplTop
            .Item(oId).Bottom = .Item(oId).Bottom - (ospl.Height \ 2)
          Next
          For Each oId In ospl.IdsSplBottom
            .Item(oId).Top = .Item(oId).Top + (ospl.Height \ 2)
          Next
        Case orVertical
          For Each oId In ospl.IdsSplLeft
            .Item(oId).Right = .Item(oId).Right - (ospl.Width \ 2)
          Next
          For Each oId In ospl.IdsSplRight
            .Item(oId).Left = .Item(oId).Left + (ospl.Width \ 2)
          Next
      End Select
    Next
  End With
  
  '-- Creates the new PictureBox control instances to represent the splitter
  For i = 0 To picSplitter.Count - 2
    Unload picSplitter(i)
  Next
  For Each ospl In mSplitters
    Load picSplitter(ospl)
    picSplitter(ospl).MousePointer = vbCustom
    Select Case ospl.Orientation
      Case orHorizontal
        picSplitter(ospl).MouseIcon = _
          LoadResPicture(gconCurHSplitter, vbResCursor)
      Case orVertical
        picSplitter(ospl).MouseIcon = _
          LoadResPicture(gconCurVSplitter, vbResCursor)
    End Select
    picSplitter(ospl).Visible = True
  Next
 
ErrorHandler:
  If (Err.Number <> 0) And (Not Ambient.UserMode) Then
    Resume Next
  ElseIf (Err.Number <> 0) Or (Not IsSolid) Or (Not mControls.IsValid) Then
    SecureRaiseError udeErrNumber:=errBuildSplitters, strSource:="Activate"
  End If
  
  Refresh
End Sub

' Purpose    : Retrieves the drop guider rectangle
' Inputs     : * IdCtlSource (the control's id that will be moved)
'              * IdCtlDestination (the control's id where the control
'                                  IdCtlSource will be moved to)
'              * IdSplDestination (the splitter's id where the control
'                                  IdCtlSource will be moved to)
'              * udeTargetType (the target type (an edge or a splitter) of the
'                               drop rect)
' Return     : As specified
Private Function GetDropRect( _
                   ByVal IdCtlSource As Long, _
                   ByVal IdCtlDestination As Long, _
                   ByVal IdSplDestination As Long, _
                   ByVal udeTargetType As genmMoveDestination _
                 ) As mdlAPI.RECT
  Const conMaxCtlAreaTaken = 0.5             'the maximum percentage area of the
                                             'other controls allowed to be taken
  Const conCtlAreaTaken = 0.25                'the percentage area of the target
                                              '    control that will be taken in
                                              '              move control action
  
  Dim lngHeight As Long           'indicating the drop guider rectangle's height
  Dim lngSplSize As Long
  Dim lngWidth As Long             'indicating the drop guider rectangle's width
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim octlMinHeight As clsControl                 'to get the value of lngHeight
  Dim octlMinWidth As clsControl                   'to get the value of lngWidth
  Dim oId As clsId                     'for enumerating all Id in Ids collection
  Dim urecGetDropRect As mdlAPI.RECT                             'returned value
  
  Set octlMinWidth = New clsControl
  Set octlMinHeight = New clsControl
  lngSplSize = 0
  Select Case udeTargetType
    Case mdControlTop
      '-- The cursor is on the control's top edge
      With mControls(IdCtlDestination)
        If .IdSplTop <> gconUninitialized Then _
          lngSplSize = mSplitters(.IdSplTop).Height \ 2
        If (.Width < mControls(IdCtlSource).MinWidth) Or _
           ((.Height * conCtlAreaTaken) - lngSplSize < _
            mControls(IdCtlSource).MinHeight) Then
          urecGetDropRect.Left = gconUninitialized
        Else
          mdlAPI.SetRect lpRect:=urecGetDropRect, _
                         X1:=.Left, Y1:=.Top, _
                         X2:=.Right, Y2:=(.Top + (.Height * conCtlAreaTaken))
        End If
      End With
    Case mdControlRight
      '-- The cursor is on the control's right edge
      With mControls(IdCtlDestination)
        If .IdSplRight <> gconUninitialized Then _
          lngSplSize = mSplitters(.IdSplRight).Width \ 2
        If ((.Width * conCtlAreaTaken) - lngSplSize < _
            mControls(IdCtlSource).MinWidth) Or _
           (.Height < mControls(IdCtlSource).MinHeight) Then
          urecGetDropRect.Left = gconUninitialized
        Else
          mdlAPI.SetRect lpRect:=urecGetDropRect, _
                         X1:=(.Right - (.Width * conCtlAreaTaken)), Y1:=.Top, _
                         X2:=.Right, Y2:=.Bottom
        End If
      End With
    Case mdControlBottom
      '-- The cursor is on the control's bottom edge
      With mControls(IdCtlDestination)
        If .IdSplBottom <> gconUninitialized Then _
          lngSplSize = mSplitters(.IdSplBottom).Height \ 2
        If (.Width < mControls(IdCtlSource).MinWidth) Or _
           ((.Height * conCtlAreaTaken) - lngSplSize < _
            mControls(IdCtlSource).MinHeight) Then
          urecGetDropRect.Left = gconUninitialized
        Else
          mdlAPI.SetRect lpRect:=urecGetDropRect, _
                         X1:=.Left, _
                         Y1:=(.Bottom - (.Height * conCtlAreaTaken)), _
                         X2:=.Right, Y2:=.Bottom
        End If
      End With
    Case mdControlLeft
      '-- The cursor is on the control's left edge
      With mControls(IdCtlDestination)
        If .IdSplLeft <> gconUninitialized Then _
          lngSplSize = mSplitters(.IdSplLeft).Width \ 2
        If ((.Width * conCtlAreaTaken) - lngSplSize < _
            mControls(IdCtlSource).MinWidth) Or _
           (.Height < mControls(IdCtlSource).MinHeight) Then
          urecGetDropRect.Left = gconUninitialized
        Else
          mdlAPI.SetRect lpRect:=urecGetDropRect, _
                         X1:=.Left, Y1:=.Top, _
                         X2:=(.Left + (.Width * conCtlAreaTaken)), Y2:=.Bottom
        End If
      End With
    Case mdEdgeTop
      '-- The cursor is on the top edge
      ' Get the height of the dropped control
      octlMinHeight.Height = gconLngInfinite
      For Each octl In mControls
        If (Not octl.Closed) And (octl <> mControls(IdCtlSource)) And _
           (octl.Bottom < octlMinHeight.Bottom) Then Set octlMinHeight = octl
      Next
      lngHeight = _
        GetMin(mControls(IdCtlSource).Height, _
               (octlMinHeight.Height * conMaxCtlAreaTaken) - mSplitters.Size)
      ' If the height is less than the minimum height, don't draw the drop guider
      '   rectangle, else draw the drop guider rectangle
      If (lngHeight < octlMinHeight.MinHeight) Or _
         (lngHeight < mControls(IdCtlSource).MinHeight) Then
        urecGetDropRect.Left = gconUninitialized
      Else
        mdlAPI.SetRect lpRect:=urecGetDropRect, _
                       X1:=0, Y1:=0, X2:=UserControl.ScaleWidth, Y2:=lngHeight
      End If
    Case mdEdgeRight
      '-- The cursor is on the right edge
      ' Get the width of the dropped control
      octlMinWidth.Width = 0
      For Each octl In mControls
        If (Not octl.Closed) And (octl <> mControls(IdCtlSource)) And _
           (octl.Left > octlMinWidth.Left) Then Set octlMinWidth = octl
      Next
      lngWidth = _
        GetMin(mControls(IdCtlSource).Width, _
               (octlMinWidth.Width * conMaxCtlAreaTaken) - mSplitters.Size)
      ' If the width is less than the minimum width, don't draw the drop guider
      '   rectangle, else draw the drop guider rectangle
      If (lngWidth < octlMinWidth.MinWidth) Or _
         (lngWidth < mControls(IdCtlSource).MinWidth) Then
        urecGetDropRect.Left = gconUninitialized
      Else
        mdlAPI.SetRect lpRect:=urecGetDropRect, _
                       X1:=UserControl.ScaleWidth - lngWidth, Y1:=0, _
                       X2:=UserControl.ScaleWidth, Y2:=UserControl.ScaleHeight
      End If
    Case mdEdgeBottom
      '-- The cursor is on the bottom edge
      ' Get the height of the dropped control
      octlMinHeight.Height = 0
      For Each octl In mControls
        If (Not octl.Closed) And (octl <> mControls(IdCtlSource)) And _
           (octl.Top > octlMinHeight.Top) Then Set octlMinHeight = octl
      Next
      lngHeight = _
        GetMin(mControls(IdCtlSource).Height, _
               (octlMinHeight.Height * conMaxCtlAreaTaken) - mSplitters.Size)
      ' If the height is less than the minimum height, don't draw the drop guider
      '   rectangle, else draw the drop guider rectangle
      If (lngHeight < octlMinHeight.MinHeight) Or _
         (lngHeight < mControls(IdCtlSource).MinHeight) Then
        urecGetDropRect.Left = gconUninitialized
      Else
        mdlAPI.SetRect lpRect:=urecGetDropRect, _
                       X1:=0, Y1:=UserControl.ScaleHeight - lngHeight, _
                       X2:=UserControl.ScaleWidth, Y2:=UserControl.ScaleHeight
      End If
    Case mdEdgeLeft
      '-- The cursor is on the left edge
      ' Get the width of the dropped control
      octlMinWidth.Width = gconLngInfinite
      For Each octl In mControls
        If (Not octl.Closed) And (octl <> mControls(IdCtlSource)) And _
           (octl.Right < octlMinWidth.Right) Then Set octlMinWidth = octl
      Next
      lngWidth = GetMin((octlMinWidth.Width * conMaxCtlAreaTaken) - _
                          mSplitters.Size, _
                        mControls(IdCtlSource).Width)
      ' If the width is less than the minimum width, don't draw the drop guider
      '   rectangle, else draw the drop guider rectangle
      If (lngWidth < octlMinWidth.MinWidth) Or _
         (lngWidth < mControls(IdCtlSource).MinWidth) Then
        urecGetDropRect.Left = gconUninitialized
      Else
        mdlAPI.SetRect lpRect:=urecGetDropRect, _
                       X1:=0, Y1:=0, X2:=lngWidth, Y2:=UserControl.ScaleHeight
      End If
    Case mdSplitter
      '-- The cursor is on a splitter
      Select Case mSplitters(IdSplDestination).Orientation
        Case orHorizontal
          ' Get the height of the dropped control
          octlMinHeight.Height = gconLngInfinite
          For Each oId In mSplitters(IdSplDestination).IdsCtlTop
            If (mControls(oId) <> mControls(IdCtlSource)) And _
               (mControls(oId).Height < octlMinHeight.Height) Then _
              Set octlMinHeight = mControls(oId)
          Next
          For Each oId In mSplitters(IdSplDestination).IdsCtlBottom
            If (mControls(oId) <> mControls(IdCtlSource)) And _
               (mControls(oId).Height < octlMinHeight.Height) Then _
              Set octlMinHeight = mControls(oId)
          Next
          lngHeight = _
            GetMin(mControls(IdCtlSource).Height, _
                   (octlMinHeight.Height * _
                      conMaxCtlAreaTaken) - mSplitters.Size)
          ' If the height is less than the minimum height, don't draw the drop
          '   guider rectangle, else draw the drop guider rectangle
          If (lngHeight < octlMinHeight.MinHeight) Or _
             (lngHeight < mControls(IdCtlSource).MinHeight) Then
            urecGetDropRect.Left = gconUninitialized
          Else
            With mSplitters(IdSplDestination)
              mdlAPI.SetRect lpRect:=urecGetDropRect, _
                             X1:=.Left, Y1:=.Top - lngHeight, _
                             X2:=.Right, Y2:=.Bottom + lngHeight
            End With
          End If
        Case orVertical
          ' Get the width of the dropped control
          octlMinWidth.Width = gconLngInfinite
          For Each oId In mSplitters(IdSplDestination).IdsCtlLeft
            If (mControls(oId) <> mControls(IdCtlSource)) And _
               (mControls(oId).Width < octlMinWidth.Width) Then _
              Set octlMinWidth = mControls(oId)
          Next
          For Each oId In mSplitters(IdSplDestination).IdsCtlRight
            If (mControls(oId) <> mControls(IdCtlSource)) And _
               (mControls(oId).Width < octlMinWidth.Width) Then _
              Set octlMinWidth = mControls(oId)
          Next
          lngWidth = _
            GetMin(mControls(IdCtlSource).Width, _
                   (octlMinWidth.Width * conMaxCtlAreaTaken) - mSplitters.Size)
          ' If the width is less than the minimum width, don't draw the drop
          '   guider rectangle, else draw the drop guider rectangle
          If (lngWidth < octlMinWidth.MinWidth) Or _
             (lngWidth < mControls(IdCtlSource).MinWidth) Then
            urecGetDropRect.Left = gconUninitialized
          Else
            With mSplitters(IdSplDestination)
              mdlAPI.SetRect lpRect:=urecGetDropRect, _
                             X1:=.Left - lngWidth, Y1:=.Top, _
                             X2:=.Right + lngWidth, Y2:=.Bottom
            End With
          End If
      End Select
  End Select
  Set octlMinWidth = Nothing
  Set octlMinHeight = Nothing
  
  GetDropRect = urecGetDropRect
End Function

' Purpose    : Retrieves the drop guider target type based on the current mouse
'              position
' Returns    : * blnTargetValid (indicating whether the current mouse position
'                                is on the valid target)
'              * udtTargetType (the target type: an edge or a control's edge or
'                               a splitter of the drop rect)
'              * lngIdCtl (the target control's id)
'              * lngIdSpl (the target splitter's id)
Private Sub GetDropTarget(ByRef blnTargetValid As Boolean, _
                          ByRef udeTargetType As genmMoveDestination, _
                          ByRef lngIdCtl As Long, ByRef lngIdSpl As Long)
  Const conCtlAreaTaken = 0.1                 'the percentage area of the target
                                              '    control that will be taken in
                                              '              move control action
  
  Dim octl As clsControl                   'for enumerating all virtual controls
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
  Dim uposCursor As mdlAPI.POINTAPI    'indicating the current cursor's position
                                       ' relative to the Control Manager control

  uposCursor = GetCursorRelPos(UserControl.hwnd)
  blnTargetValid = True
  lngIdCtl = gconUninitialized
  lngIdSpl = gconUninitialized
  
  '-- Check whether the cursor is on the edge of the Control Manager control
  If uposCursor.X <= mSplitters.Size Then
    udeTargetType = mdEdgeLeft
  ElseIf uposCursor.X >= UserControl.ScaleWidth - mSplitters.Size Then
    udeTargetType = mdEdgeRight
  ElseIf uposCursor.Y <= mSplitters.Size Then
    udeTargetType = mdEdgeTop
  ElseIf uposCursor.Y >= UserControl.ScaleHeight - mSplitters.Size Then
    udeTargetType = mdEdgeBottom
  Else
    blnTargetValid = False
    
    '-- Check whether the cursor is on the edge of a control
    If Not blnTargetValid Then
      For Each octl In mControls
        If Not octl.Closed Then
          If ((octl.Left <= uposCursor.X) And (uposCursor.X <= octl.Right) And _
              (octl.Top <= uposCursor.Y) And _
              (uposCursor.Y <= octl.Bottom)) And _
             ((uposCursor.X <= octl.Left + (octl.Width * conCtlAreaTaken)) Or _
              (uposCursor.X >= octl.Right - (octl.Width * conCtlAreaTaken)) Or _
              (uposCursor.Y <= octl.Top + (octl.Height * conCtlAreaTaken)) Or _
              (uposCursor.Y >= octl.Bottom - _
                (octl.Height * conCtlAreaTaken))) Then
            blnTargetValid = True
            lngIdCtl = octl.Id
            Select Case GetMin(uposCursor.Y - octl.Top, _
                               octl.Right - uposCursor.X, _
                               octl.Bottom - uposCursor.Y, _
                               uposCursor.X - octl.Left)
              Case uposCursor.Y - octl.Top
                udeTargetType = mdControlTop
              Case octl.Right - uposCursor.X
                udeTargetType = mdControlRight
              Case octl.Bottom - uposCursor.Y
                udeTargetType = mdControlBottom
              Case uposCursor.X - octl.Left
                udeTargetType = mdControlLeft
            End Select
          End If
        End If
      Next
    End If
    
    '-- Check whether the cursor is on a splitter
    If Not blnTargetValid Then
      For Each ospl In mSplitters
        If (ospl.Left <= uposCursor.X) And (uposCursor.X <= ospl.Right) And _
           (ospl.Top <= uposCursor.Y) And (uposCursor.Y <= ospl.Bottom) Then
          lngIdSpl = ospl.Id
          udeTargetType = mdSplitter
          blnTargetValid = True
          Exit For
        End If
      Next
    End If
  End If
End Sub

' Purpose    : Returns a valid x- and y- coordinate scale
' Return     : As specified
Private Sub GetValidStretchScale(ByRef sngXScale As Single, _
                                 ByRef sngYScale As Single)
Attribute GetValidStretchScale.VB_Description = "Returns a valid x- and y- coordinate scale"
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  
  sngXScale = UserControl.ScaleWidth / mControls.Width
  sngYScale = UserControl.ScaleHeight / mControls.Height
  For Each octl In mControls
    If Not octl.Closed Then
      If octl.Width * sngXScale < octl.MinWidth Then sngXScale = 1
      If octl.Height * sngYScale < octl.MinHeight Then sngYScale = 1
    End If
  Next
End Sub

' Purpose    : Returns a value indicating whether the target move action
'              udeTargetType is near the control IdCtlSource which will be moved
' Inputs     : * IdCtlSource (the id of the control which will be moved)
'              * IdCtlDestination (the id of the control where the control
'                                  IdCtlSource will be moved to [for MoveTo =
'                                  mdControlTop, mdControlRight, mdControlBottom
'                                  or mdControlLeft)
'              * IdSplDestination (the id of the splitter where the control
'                                  IdControlSource will be moved to [for
'                                  MoveTo = mdSplitter])
'              * udeTargetType (the type of the area [an edge, a control's edge
'                               or a splitter] where the control IdControl will
'                               be moved to)
' Return     : As specified
Private Function IsRectNearSource( _
                   ByVal IdCtlSource As Long, ByVal IdCtlDestination As Long, _
                   ByVal IdSplDestination As Long, _
                   ByVal udeTargetType As genmMoveDestination _
                 ) As Boolean
  Dim blnIsRectNearSource As Boolean

  Select Case udeTargetType
    Case mdControlTop, mdControlRight, mdControlBottom, mdControlLeft
      blnIsRectNearSource = (IdCtlSource = IdCtlDestination)
    Case mdSplitter
      blnIsRectNearSource = False
      With mControls(IdCtlSource)
        If .IdSplTop <> gconUninitialized Then _
          blnIsRectNearSource = blnIsRectNearSource Or _
                                ((IdSplDestination = .IdSplTop) And _
                                 (mSplitters(.IdSplTop).IdsCtlBottom.Count = 1))
        If .IdSplRight <> gconUninitialized Then _
          blnIsRectNearSource = blnIsRectNearSource Or _
                                ((IdSplDestination = .IdSplRight) And _
                                 (mSplitters(.IdSplRight).IdsCtlLeft.Count = 1))
        If .IdSplBottom <> gconUninitialized Then _
          blnIsRectNearSource = blnIsRectNearSource Or _
                                ((IdSplDestination = .IdSplBottom) And _
                                 (mSplitters(.IdSplBottom).IdsCtlTop.Count = 1))
        If .IdSplLeft <> gconUninitialized Then _
          blnIsRectNearSource = blnIsRectNearSource Or _
                                ((IdSplDestination = .IdSplLeft) And _
                                 (mSplitters(.IdSplLeft).IdsCtlRight.Count = 1))
      End With
  End Select
  IsRectNearSource = blnIsRectNearSource
End Function

' Purpose    : Returns a value indicating whether this VB Control Manager
'              control instance contains another VB Control Manager Controls
'              instance
' Return     : As specified
Private Function IsSelfContained() As Boolean
Attribute IsSelfContained.VB_Description = "Returns a value indicating whether this VB Splitter control instance contains another VB Splitter Controls instance"
  Dim blnIsSelfContained As Boolean                              'returned value
  Dim ctl As Control                            'for enumerating all controls in
                                                '   ContainedControls collection
  
  blnIsSelfContained = False
  For Each ctl In ContainedControls
    If TypeOf ctl Is ControlManager Then
      blnIsSelfContained = True
      Exit For
    End If
  Next
  IsSelfContained = blnIsSelfContained
End Function

' Purpose    : Returns a value indicating whether the virtual controls and
'              splitters are solid
' Input      : blnIncludeSplitter (indicating whether the splitters are included
'                                  to determine the solid state)
' Return     : As specified
' Note       : See VB Control Manager's documention for the definition of
'              "solid"
Private Function IsSolid( _
                   Optional ByVal blnIncludeSplitter As Boolean = True _
                 ) As Boolean
  Dim lngExtent As Long      'total extent of the virtual controls and splitters
  Dim lngSplTopHeight As Long         'the height of the virtual splitter on the
                                      '       top-side of the current enumerated
                                      '                          virtual control
  Dim lngSplRightWidth As Long         'the width of the virtual splitter on the
                                       '    right-side of the current enumerated
                                       '                         virtual control
  Dim lngSplBottomHeight As Long      'the height of the virtual splitter on the
                                      '    bottom-side of the current enumerated
                                      '                          virtual control
  Dim lngSplLeftWidth As Long          'the width of the virtual splitter on the
                                       '     left-side of the current enumerated
                                       '                         virtual control
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  
  lngExtent = 0
  For Each octl In mControls
    If Not octl.Closed Then
      If blnIncludeSplitter And (octl.IdSplTop <> gconUninitialized) Then
        lngSplTopHeight = mSplitters(octl.IdSplTop).Height
      Else
        lngSplTopHeight = 0
      End If
      If blnIncludeSplitter And (octl.IdSplRight <> gconUninitialized) Then
        lngSplRightWidth = mSplitters(octl.IdSplRight).Width
      Else
        lngSplRightWidth = 0
      End If
      If blnIncludeSplitter And (octl.IdSplBottom <> gconUninitialized) Then
        lngSplBottomHeight = mSplitters(octl.IdSplBottom).Height
      Else
        lngSplBottomHeight = 0
      End If
      If blnIncludeSplitter And (octl.IdSplLeft <> gconUninitialized) Then
        lngSplLeftWidth = mSplitters(octl.IdSplLeft).Width
      Else
        lngSplLeftWidth = 0
      End If
      lngExtent = lngExtent + ((octl.Width + (lngSplLeftWidth \ 2) + _
                               (lngSplRightWidth \ 2)) * _
                              (octl.Height + (lngSplTopHeight \ 2) + _
                               (lngSplBottomHeight \ 2)))
    End If
  Next
  IsSolid = (lngExtent = 0) Or _
            (lngExtent = (mControls.Width * mControls.Height))
End Function

' Purpose    : Applies the virtual controls and splitters to their real controls
'              and splitter
' Effect     : As specified
Private Sub Refresh()
  Const conErrHeightReadOnly = 383
  
  Dim lngErrNumber As Long                      'for the control with r/o height
  Dim lngHeight As Long                    'adjusted height for list box control
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
  
  On Error GoTo ErrorHandler
  
  '-- Applies all virtuals splitters to their real splitters
  For Each ospl In mSplitters
    With picSplitter(ospl)
      .Move ospl.Left, ospl.Top, ospl.Width, ospl.Height
      .BackColor = ospl.BackColor
      .Enabled = ospl.Enable
      .ZOrder
    End With
  Next
  
  '-- Applies all virtuals controls and its title bar to their real controls
  For Each octl In mControls
    If Not octl.Closed Then
      lngHeight = AdjustedHeight(ctl:=ContainedControls(octl.Id), octl:=octl)
      With ContainedControls(octl.Id)
        If octl.TitleBar_Visible Then
          With ctbTitleBar(octl)
            .Left = octl.Left
            .Top = octl.Top
            .Height = octl.TitleBar_Height
            .Width = octl.Width
            .Visible = True
            .CloseVisible = octl.TitleBar_CloseVisible
          End With
        Else
          ctbTitleBar(octl).Visible = False
        End If
        .Move octl.Left, octl.Top + octl.TitleBar_Height, octl.Width, _
              IIf(lngHeight - octl.TitleBar_Height >= 0, _
                  lngHeight - octl.TitleBar_Height, 0)
        If lngErrNumber = conErrHeightReadOnly Then
          .Move octl.Left, octl.Top, octl.Width
          lngErrNumber = 0
        End If
      End With
    End If
  Next
  
  UserControl.Refresh
  Exit Sub
  
ErrorHandler:
  If Err.Number = conErrHeightReadOnly Then
    lngErrNumber = Err.Number
    Resume Next
  Else
    Err.Raise Err.Number
  End If
End Sub

' Purpose    : Securely raises custom error udeErrNumber by firstly ends the
'              subclassing
' Assumptions: * Error message udeErrNumber exists in the resource file
'              * Global variable gstrControlName has been initialized
' Inputs     : * udeErrNumber
'              * strSource (the location in form ClassName.RoutinesName where
'                the error occur
' Effect     : As specified
Private Sub SecureRaiseError(ByVal udeErrNumber As genmErrNumber, _
                             Optional ByVal strSource As String = "")
  DetachMessage Me, mhwndParent, WM_SIZE
  RaiseError udeErrNumber:=udeErrNumber, strSource:=strSource
End Sub

' Purpose    : Stretchs the controls and splitters to fill-up their container
' Effect     : As specified
Private Sub Stretch()
Attribute Stretch.VB_Description = "Stretchs the controls and splitters to fill-up their container"
  Dim octl As clsControl                   'for enumerating all virtual controls
                                           '              in Controls collection
  Dim oId As clsId                     'for enumerating all Id in Ids collection
  Dim ospl As clsSplitter                 'for enumerating all virtual splitters
                                          '              in Splitters collection
  Dim sngXScale As Single                          'a valid x-coorindate's scale
  Dim sngYScale As Single                          'a valid y-coordinate's scale
  
  GetValidStretchScale sngXScale, sngYScale
  
  '-- Stretch the virtual splitters
  If (Abs(sngXScale - 1) > 0.001) Or (Abs(sngYScale - 1) > 0.001) Then
    mSplitters.Width = mSplitters.Width * sngXScale
    mSplitters.Height = mSplitters.Height * sngYScale
    For Each ospl In mSplitters
      With ospl
        Select Case .Orientation
          Case orHorizontal
            .Xc = CLng(.Xc * sngXScale)
            .Yc = CLng((.Top * sngYScale) + ((.Height * sngYScale) / 2))
            .Width = CLng(.Width * sngXScale)
            .MinYc = CLng((mControls(ospl.IdCtlFriendTop).Top * sngYScale) + _
                          mControls(ospl.IdCtlFriendTop).MinHeight)
            .MaxYc = CLng((mControls(ospl.IdCtlFriendBottom).Bottom * _
                           sngYScale) - _
                          mControls(ospl.IdCtlFriendBottom).MinHeight)
          Case orVertical
            .Xc = CLng((.Left * sngXScale) + ((.Width * sngXScale) / 2))
            .Yc = CLng(.Yc * sngYScale)
            .Height = CLng(.Height * sngYScale)
            .MinXc = CLng((mControls(ospl.IdCtlFriendLeft).Left * sngXScale) + _
                          mControls(ospl.IdCtlFriendLeft).MinWidth)
            .MaxXc = CLng((mControls(ospl.IdCtlFriendRight).Right * _
                           sngXScale) - _
                          mControls(ospl.IdCtlFriendRight).MinWidth)
        End Select
      End With
    Next
    For Each ospl In mSplitters
      Select Case ospl.Orientation
        Case orHorizontal
          For Each oId In ospl.IdsSplTop
            mSplitters(oId).Bottom = ospl.Top
          Next
          For Each oId In ospl.IdsSplBottom
            mSplitters(oId).Top = ospl.Bottom
          Next
        Case orVertical
          For Each oId In ospl.IdsSplLeft
            mSplitters(oId).Right = ospl.Left
          Next
          For Each oId In ospl.IdsSplRight
            mSplitters(oId).Left = ospl.Right
          Next
      End Select
    Next
    
    '-- Stretch the virtual controls
    mControls.Width = mControls.Width * sngXScale
    mControls.Height = mControls.Height * sngYScale
    For Each octl In mControls
      If Not octl.Closed Then
        If octl.IdSplTop = gconUninitialized Then
          octl.Top = mControls.Top
          If octl.IdSplLeft <> gconUninitialized Then _
            mSplitters(octl.IdSplLeft).Top = octl.Top
          If octl.IdSplRight <> gconUninitialized Then _
            mSplitters(octl.IdSplRight).Top = octl.Top
        Else
          octl.Top = mSplitters(octl.IdSplTop).Bottom
        End If
        If octl.IdSplRight = gconUninitialized Then
          octl.Right = mControls.Right
          If octl.IdSplTop <> gconUninitialized Then _
            mSplitters(octl.IdSplTop).Right = octl.Right
          If octl.IdSplBottom <> gconUninitialized Then _
            mSplitters(octl.IdSplBottom).Right = octl.Right
        Else
          octl.Right = mSplitters(octl.IdSplRight).Left
        End If
        If octl.IdSplBottom = gconUninitialized Then
          octl.Bottom = mControls.Bottom
          If octl.IdSplLeft <> gconUninitialized Then _
            mSplitters(octl.IdSplLeft).Bottom = octl.Bottom
          If octl.IdSplRight <> gconUninitialized Then _
            mSplitters(octl.IdSplRight).Bottom = octl.Bottom
        Else
          octl.Bottom = mSplitters(octl.IdSplBottom).Top
        End If
        If octl.IdSplLeft = gconUninitialized Then
          octl.Left = mControls.Left
          If octl.IdSplTop <> gconUninitialized Then _
            mSplitters(octl.IdSplTop).Left = octl.Left
          If octl.IdSplBottom <> gconUninitialized Then _
          mSplitters(octl.IdSplBottom).Left = octl.Left
        Else
          octl.Left = mSplitters(octl.IdSplLeft).Right
        End If
      End If
    Next
    
    Refresh
  End If
End Sub
