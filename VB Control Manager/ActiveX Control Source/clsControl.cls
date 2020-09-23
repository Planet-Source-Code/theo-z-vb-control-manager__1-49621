VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to represent a ""virtual"" control"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     : clsControl.cls                                            **
'** Language      : Visual Basic 6.0                                          **
'** Author        : Theo Zacharias (theo_yz@yahoo.com)                        **
'** Description   : A class module to represent a virtual control             **
'** Usage         : Instead of using a real control (like TextBox control)    **
'**                 properties directly, VB Control Manager control saves the **
'**                 necessary properties plus several custom properties in a  **
'**                 virtual control. Furthermore, VB Control Manager control  **
'**                 works with this virtual control to manipulate its         **
'**                 position and size and then write it back to the related   **
'                   real control.                                             **
'** Dependencies  : mdlGeneral                                                **
'** Public Members:                                                           **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Properties : Bottom (r/o), Closed (r/o), Height (r/o), Id (def. r/o), **
'**                  IdSplBottom (r/o), IdSplLeft (r/o), IdSplRight (r/o),    **
'**                  IdSplTop (r/o), Left (r/o), MinHeight (r/o),             **
'**                  MinWidth (r/o), Name (r/o), Right (r/o),                 **
'**                  TitleBar_CloseVisible (r/w), TitleBar_Height (r/o),      **
'**                  TitleBar_Visible (r/w), Top (r/o), Width (r/o),          **
'**                  Xc (r/o), Yc (r/o)                                       **
'**   * Methods    : -                                                        **
'**   * Events     : -                                                        **
'** Friend Members:                                                           **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Properties : Bottom (r/w), Closed (r/w), Height (r/w), Id (def. r/w), **
'**                  IdCtlFriendBottom (r/w), IdCtlFriendLeft (r/w),          **
'**                  IdCtlFriendRight (r/w), IdCtlFriendTop (r/w),            **
'**                  IdSplBottom (r/w), IdSplLeft (r/w), IdSplRight (r/w),    **
'**                  IdSplTop (r/w), Left (r/w), MinHeight (r/w),             **
'**                  MinWidth (r/w), Name (r/w), Parent (r/w), Right (r/w),   **
'**                  TitleBar_CloseVisible (r/w), TitleBar_Height (r/w),      **
'**                  TitleBar_Visible (r/w), TitleBar_VisibleHeight (r/o),    **
'**                  Top (r/w), Width (r/w), Xc (r/o), Yc (r/o)               **
'**   * Methods    : -                                                        **
'** Notes         : * Height and Yc properties value are generated from       **
'**                   Bottom and Top properties value                         **
'**                 * Width and Xc properties value are generated from Left   **
'**                   and Right properties value                              **
'** Last modified on September 2, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Property Variables
Private mId As Long
Private mName As String
Private mParent As clsControls
Private mblnClosed As Boolean

' Position and Size
Private mlngLeft As Long
Private mlngTop As Long
Private mlngRight As Long
Private mlngBottom As Long

' Minimum Size
Private mlngMinWidth As Long
Private mlngMinHeight As Long

'Title Bar
Private mblnTitleBar_CloseVisible As Boolean
Private mlngTitleBar_Height As Long
Private mblnTitleBar_Visible As Boolean

' Friends Controls' Id
' Note: Friend controls are the nearest control around the control
Private mIdCtlFriendTop As Long
Private mIdCtlFriendRight As Long
Private mIdCtlFriendBottom As Long
Private mIdCtlFriendLeft As Long

' Friend Splitters' Id
' Note: Friend splitters (called splitters in this class) are the splitters
'       around the control
Private mIdSplTop As Long
Private mIdSplRight As Long
Private mIdSplBottom As Long
Private mIdSplLeft As Long

'--- Property Default Values
Private Const mconDefaultClosed As Boolean = False

'--- Implements the clsTitleBar interface
Implements TitleBar

'-------------------
' Class Constructor
'-------------------

Private Sub Class_Initialize()
  mId = gconUninitialized
  mblnClosed = mconDefaultClosed
  
  mlngMinWidth = 0
  mlngMinHeight = 0
  
  mIdCtlFriendTop = gconUninitialized
  mIdCtlFriendRight = gconUninitialized
  mIdCtlFriendBottom = gconUninitialized
  mIdCtlFriendLeft = gconUninitialized
  
  mIdSplTop = gconUninitialized
  mIdSplRight = gconUninitialized
  mIdSplBottom = gconUninitialized
  mIdSplLeft = gconUninitialized
End Sub

'------------------
' Class Properties
'------------------

' Purpose    : Sets the distance between the internal bottom edge of the virtual
'              control and the top edge of the related Control Manager object
' Effect     : As specified
' Input      : lngBottom (the new Bottom property value)
Friend Property Let Bottom(ByVal lngBottom As Long)
  mlngBottom = lngBottom
End Property

' Purpose    : Returns the distance between the internal bottom edge of the
'              virtual control and the top edge of the related Control Manager
'              object
' Return     : As specified
Public Property Get Bottom() As Long
Attribute Bottom.VB_Description = "Returns the distance between the internal bottom edge of the virtual control and the top edge of the related Control Manager object"
  Bottom = mlngBottom
End Property

' Purpose    : Sets a value that determines whether the virtual control is
'              closed
' Effect     : As specified
' Input      : blnClosed (the new Closed property value)
Friend Property Let Closed(ByVal blnClosed As Boolean)
  mblnClosed = blnClosed
End Property

' Purpose    : Returns a value that determines whether the virtual control is
'              closed
' Return     : As specified
Public Property Get Closed() As Boolean
Attribute Closed.VB_Description = "Returns a value that determines whether the control is closed"
  Closed = mblnClosed
End Property

' Purpose    : Sets the height of the virtual control
' Effect     : As specified
' Input      : lngHeight(the new Height property value)
Friend Property Let Height(ByVal lngHeight As Long)
  mlngBottom = mlngTop + lngHeight
End Property

' Purpose    : Returns the height of the virtual control
' Return     : As specified
Public Property Get Height() As Long
Attribute Height.VB_Description = "Returns the height of the virtual control"
  Height = mlngBottom - mlngTop
End Property

' Purpose    : Sets a value that uniquely identifies the virtual control
' Effect     : As specified
' Input      : lngId(the new Id property value)
Friend Property Let Id(ByVal lngId As Long)
  mId = lngId
End Property

' Purpose    : Returns a value that uniquely identifies the virtual control
' Return     : As specified
Public Property Get Id() As Long
Attribute Id.VB_Description = "Returns a value that uniquely identifies the virtual control"
Attribute Id.VB_UserMemId = 0
  Id = mId
End Property

' Purpose    : Sets the control's bottom-side friend control's Id
' Effect     : As specified
' Input      : lngIdCtlFriendBottom (the new IdCtlFriendBottom property value)
Friend Property Let IdCtlFriendBottom(ByVal lngIdCtlFriendBottom As Long)
Attribute IdCtlFriendBottom.VB_Description = "Returns or sets the control's bottom-side friend control's Id"
  mIdCtlFriendBottom = lngIdCtlFriendBottom
End Property

' Purpose    : Returns the control's bottom-side friend control's Id
' Return     : As specified
Friend Property Get IdCtlFriendBottom() As Long
  IdCtlFriendBottom = mIdCtlFriendBottom
End Property

' Purpose    : Sets the control's left-side friend control's Id
' Effect     : As specified
' Input      : lngIdCtlFriendLeft (the new IdCtlFriendLeft property value)
Friend Property Let IdCtlFriendLeft(ByVal lngIdCtlFriendLeft As Long)
Attribute IdCtlFriendLeft.VB_Description = "Returns or sets the control's left-side friend control's Id"
  mIdCtlFriendLeft = lngIdCtlFriendLeft
End Property

' Purpose    : Returns the control's left-side friend control's Id
' Return     : As specified
Friend Property Get IdCtlFriendLeft() As Long
  IdCtlFriendLeft = mIdCtlFriendLeft
End Property

' Purpose    : Sets the control's right-side friend control's Id
' Effect     : As specified
' Input      : lngIdCtlFriendRight (the new IdCtlFriendRight property value)
Friend Property Let IdCtlFriendRight(ByVal lngIdCtlFriendRight As Long)
Attribute IdCtlFriendRight.VB_Description = "Returns or sets the control's right-side friend control's Id"
  mIdCtlFriendRight = lngIdCtlFriendRight
End Property

' Purpose    : Returns the control's right-side friend control's Id
' Return     : As specified
Friend Property Get IdCtlFriendRight() As Long
  IdCtlFriendRight = mIdCtlFriendRight
End Property

' Purpose    : Sets the control's top-side friend control's Id
' Input      : lngIdCtlFriendTop (the new IdCtlFriendTop property value)
Friend Property Let IdCtlFriendTop(ByVal lngIdCtlFriendTop As Long)
Attribute IdCtlFriendTop.VB_Description = "Returns or sets the control's top-side friend control's Id"
  mIdCtlFriendTop = lngIdCtlFriendTop
End Property

' Purpose    : Returns the control's top-side friend control's Id
Friend Property Get IdCtlFriendTop() As Long
  IdCtlFriendTop = mIdCtlFriendTop
End Property

' Purpose    : Sets a value that uniquely identifies the virtual splitter which
'              its movement will change the value of the virtual control's
'              Bottom property
' Effect     : As specified
' Input      : lngIdSplFriendBottom (the new IdSplFriendBottom property value)
Friend Property Let IdSplBottom(ByVal lngIdSplBottom As Long)
  mIdSplBottom = lngIdSplBottom
End Property

' Purpose    : Returns a value that uniquely identifies the virtual splitter
'              which its movement will change the value of the virtual control's
'              Bottom property
' Return     : As specified
Public Property Get IdSplBottom() As Long
Attribute IdSplBottom.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Bottom property"
  IdSplBottom = mIdSplBottom
End Property

' Purpose    : Sets a value that uniquely identifies the virtual splitter which
'              its movement will change the value of the virtual control's Left
'              property
' Effect     : As specified
' Input      : lngIdSplFriendLeft (the new IdSplFriendLeft property value)
Friend Property Let IdSplLeft(ByVal lngIdSplLeft As Long)
  mIdSplLeft = lngIdSplLeft
End Property

' Purpose    : Returns a value that uniquely identifies the virtual splitter
'              which its movement will change the value of the virtual control's
'              Left property
' Return     : As specified
Public Property Get IdSplLeft() As Long
Attribute IdSplLeft.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Left property"
  IdSplLeft = mIdSplLeft
End Property

' Purpose    : Sets a value that uniquely identifies the virtual splitter which
'              its movement will change the value of the virtual control's Right
'              property
' Effect     : As specified
' Input      : lngIdSplFriendRight (the new IdSplFriendRight property value)
Friend Property Let IdSplRight(ByVal lngIdSplRight As Long)
  mIdSplRight = lngIdSplRight
End Property

' Purpose    : Returns a value that uniquely identifies the virtual splitter
'              which its movement will change the value of the virtual control's
'              Right property
' Return     : As specified
Public Property Get IdSplRight() As Long
Attribute IdSplRight.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Right property"
  IdSplRight = mIdSplRight
End Property

' Purpose    : Sets a value that uniquely identifies the virtual splitter which
'              its movement will change the value of the virtual control's Top
'              property
' Effect     : As specified
' Input      : lngIdSplFriendTop (the new IdSplFriendTop property value)
Friend Property Let IdSplTop(ByVal lngIdSplTop As Long)
  mIdSplTop = lngIdSplTop
End Property

' Purpose    : Returns a value that uniquely identifies the virtual splitter
'              which its movement will change the value of the virtual control's
'              Top property
' Return     : As specified
Public Property Get IdSplTop() As Long
Attribute IdSplTop.VB_Description = "Returns a value that uniquely identifies the virtual splitter which its movement will change the value of the virtual control's Top property"
  IdSplTop = mIdSplTop
End Property

' Purpose    : Sets the distance between the internal left edge of the virtual
'              control and the left edge of the related Control Manager object
' Effect     : As specified
' Input      : lngLeft (the new Left property value)
Friend Property Let Left(ByVal lngLeft As Long)
  mlngLeft = lngLeft
End Property

' Purpose    : Returns the distance between the internal left edge of the
'              virtual control and the left edge of the related Control Manager
'              object
' Return     : As specified
Public Property Get Left() As Long
Attribute Left.VB_Description = "Returns the distance between the internal left edge of the virtual control and the left edge of the related Control Manager object"
  Left = mlngLeft
End Property

' Purpose    : Sets the minimum height of the virtual control
' Effect     : As specified
' Input      : lngMinHeight(the new MinHeight propety value)
Friend Property Let MinHeight(ByVal lngMinHeight As Long)
  mlngMinHeight = lngMinHeight
End Property

' Purpose    : Returns the minimum height of the virtual control
' Return     : As specified
Public Property Get MinHeight() As Long
Attribute MinHeight.VB_Description = "Returns the minimum height of the virtual control"
  MinHeight = mlngMinHeight + TitleBar_Height
End Property

' Purpose    : Sets the minimum width of the virtual control
' Effect     : As specified
' Input      : lngMinWidth(the new MinWidth propety value)
Friend Property Let MinWidth(ByVal lngMinWidth As Long)
  mlngMinWidth = lngMinWidth
End Property

' Purpose    : Returns the minimum width of the control
' Return     : As specified
Public Property Get MinWidth() As Long
Attribute MinWidth.VB_Description = "Returns the minimum width of the virtual control"
  MinWidth = mlngMinWidth
End Property

' Purpose    : Sets the name of the real control that the virtual control
'              represents
' Effect     : As specified
' Input      : strName (the new Name property value)
Friend Property Let Name(ByVal strName As String)
  mName = strName
End Property

' Purpose    : Returns the name of the real control that this virtual control
'              represents
' Return     : As specified
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name of the real control that the virtual control represents"
  Name = mName
End Property

' Purpose    : Sets the collection that contain this class
' Effect     : As specified
' Input      : splsParent (the new Parent property value)
Friend Property Set Parent(ByVal ctlsParent As clsControls)
  Set mParent = ctlsParent
End Property

' Purpose    : Returns the collection that contain this class
' Return     : As specified
Friend Property Get Parent() As clsControls
Attribute Parent.VB_Description = "Returns the collection that contain this class"
  Set Parent = mParent
End Property

' Purpose    : Sets the distance between the internal right edge of the virtual
'              control and the left edge of the related Control Manager object
' Effect     : As specified
' Input      : lngRight (the new Right property value)
Friend Property Let Right(ByVal lngRight As Long)
  mlngRight = lngRight
End Property

' Purpose    : Returns the distance between the internal right edge of the
'              virtual control and the left edge of the related Control Manager
'              object
' Return     : As specified
Public Property Get Right() As Long
Attribute Right.VB_Description = "Returns the distance between the internal right edge of the virtual control and the left edge of the related Control Manager object"
  Right = mlngRight
End Property

' Purpose    : Sets a value that determines whether a close button in the
'              virtual control title bar is visible
' Effect     : As specified
' Input      : blnTitleBar_CloseVisible (the new TitleBar_CloseVisible property
'                                        value)
Public Property Let TitleBar_CloseVisible( _
                       ByVal blnTitleBar_CloseVisible As Boolean _
                    )
Attribute TitleBar_CloseVisible.VB_Description = "Returns/sets a value that determines whether a close button in the virtual control title bar is visible"
  mblnTitleBar_CloseVisible = blnTitleBar_CloseVisible
  Parent.TitleBar_CloseVisible(mId) = mblnTitleBar_CloseVisible
End Property

' Purpose    : Returns a value that determines whether a close button in the
'              virtual control title bar is visible
' Return     : As specified
Public Property Get TitleBar_CloseVisible() As Boolean
  TitleBar_CloseVisible = mblnTitleBar_CloseVisible
End Property

' Purpose    : Sets the height of the virtual control title bar
' Input      : lngTitleBar_Height (the new TitleBar_Height property value)
Friend Property Let TitleBar_Height(ByVal lngTitleBar_Height As Long)
  mlngTitleBar_Height = lngTitleBar_Height
End Property

' Purpose    : Returns the height of the visible part of the virtual control
'              title bar
' Return     : As specified
Public Property Get TitleBar_Height() As Long
Attribute TitleBar_Height.VB_Description = "Returns the height of the virtual control title bars"
  If Not mblnTitleBar_Visible Then
    TitleBar_Height = 0
  Else
    TitleBar_Height = mlngTitleBar_Height
  End If
End Property

' Purpose    : Sets a value that determines whether the virtual control title
'              bar is visible
' Effect     : As specified
' Input      : blnblnTitleBar_Visible (the new blnTitleBar_Visible property
'                                      value)
Public Property Let TitleBar_Visible(ByVal blnTitleBar_Visible As Boolean)
Attribute TitleBar_Visible.VB_Description = "Returns/sets a value that determines whether the virtual control title bar is visible"
  If mblnTitleBar_Visible <> blnTitleBar_Visible Then
    mblnTitleBar_Visible = blnTitleBar_Visible
    If Parent.IsValid Then
      Parent.TitleBar_Visible(mId) = mblnTitleBar_Visible
    Else
      mblnTitleBar_Visible = Not mblnTitleBar_Visible
    End If
  End If
End Property

' Purpose    : Returns a value that determines whether the virtual control title
'              bar is visible
' Return     : As specified
Public Property Get TitleBar_Visible() As Boolean
  TitleBar_Visible = mblnTitleBar_Visible
End Property

' Purpose    : Returns the height of the title bar
' Return     : As specified
Friend Property Get TitleBar_VisibleHeight() As Long
  TitleBar_VisibleHeight = mlngTitleBar_Height
End Property

' Purpose    : Sets the distance between the internal top edge of the virtual
'              control and the top edge of the related Control Manager object
' Effect     : As specified
' Input      : lngTop (the new Top property value)
Friend Property Let Top(ByVal lngTop As Long)
  mlngTop = lngTop
End Property

' Purpose    : Returns the distance between the internal top edge of the virtual
'              control and the top edge of the related Control Manager object
' Return     : As specified
Public Property Get Top() As Long
Attribute Top.VB_Description = "Returns the distance between the internal top edge of the virtual control and the top edge of the related Control Manager object"
  Top = mlngTop
End Property

' Purpose    : Sets the width of the virtual control
' Effect     : As specified
' Input      : lngWidth(the new Width property value)
Friend Property Let Width(ByVal lngWidth As Long)
  mlngRight = mlngLeft + lngWidth
End Property

' Purpose    : Returns the width of the virtual control
' Return     : As specified
Public Property Get Width() As Long
Attribute Width.VB_Description = "Returns the width of the virtual control"
  Width = mlngRight - mlngLeft
End Property

' Purpose    : Returns the x-coordinate of the virtual control center
' Return     : As specified
Public Property Get Xc() As Long
Attribute Xc.VB_Description = "Returns the x-coordinate of the virtual control center"
  Xc = CLng((mlngLeft + mlngRight) / 2)
End Property

' Purpose    : Returns the y-coordinate of the virtual control center
' Return     : As specified
Public Property Get Yc() As Long
Attribute Yc.VB_Description = "Returns the y-coordinate of the virtual control center"
  Yc = CLng((mlngTop + mlngBottom) / 2)
End Property
