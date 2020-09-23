VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSplitters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to implement custom collection of class clsSplitter"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     : clsSplitters.cls                                          **
'** Language      : Visual Basic 6.0                                          **
'** Author        : Theo Zacharias (theo_yz@yahoo.com)                        **
'** Description   : A class module to implement custom collection of class    **
'**                 clsSplitter                                               **
'** Usage         : VB Control Manager control uses PictureBox control to for **
'**                 the splitters. Instead of using PictureBox properties     **
'**                 directly, VB Control Manager control saves the necessary  **
'**                 properties plus several custom properties in a virtual    **
'**                 splitter. Furthermore, VB Control Manager control works   **
'**                 with this virtual splitter (will be called splitter next) **
'**                 to manipulate its position and then write it back to the  **
'**                 PictureBox control.                                       **
'** Dependencies  : clsSplitter, mdlGeneral                                   **
'** Public Members:                                                           **
'**   * Collections: -                                                        **
'**   * Object     : Item (def. r/o)                                          **
'**   * Property   : Count (r/o)                                              **
'**   * Method     : IsExist                                                  **
'**   * Events     : BackColorChange                                          **
'** Friend Members:                                                           **
'**   * Collections: -                                                        **
'**   * Object     : Item (def. r/o)                                          **
'**   * Properties : ActiveColor (r/w), BackColor (r/w), Bottom (r/w),        **
'**                  ClipCursor (r/w), Count (r/o), DefaultActiveColor (r/o), **
'**                  DefaultClipCursor (r/o), DefaultEnable (r/o),            **
'**                  DefaultLiveUpdate (r/o), DefaultSize (r/o),              **
'**                  Enable (r/w), Height (r/w), Left (r/w),                  **
'**                  LiveUpdate (r/w), MinimumSize (r/o), Right (r/w),        **
'**                  Size (r/w), Top (r/w), Width (r/w)                       **
'**   * Methods    : Add, Clear, IsExist, Remove                              **
'** Last modified on September 5, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Custom Collection Class Variable
Private mcolSplitters As Collection

'--- Property Variables

' Property variables which appears in the property page
Private mblnClipCursor As Boolean
Private mblnEnable As Boolean
Private mblnLiveUpdate As Boolean
Private mlngActiveColor As OLE_COLOR
Private mlngBackColor As OLE_COLOR
Private mlngSize As Long

' Splitters frame area
Private mlngLeft As Long
Private mlngTop As Long
Private mlngRight As Long
Private mlngBottom As Long

'--- Property Default Values
Private Const mconDefaultActiveColor As Long = vbBlack
Private Const mconDefaultClipCursor As Boolean = False
Private Const mconDefaultEnable As Boolean = True
Private Const mconDefaultLiveUpdate As Boolean = True
Private Const mconDefaultSize As Long = 4

'--- Other Variables
Private mblnEventDisable As Boolean                'indicating whether the class
                                                   'event shouldn't be triggered
Private mlngMinimumSize As Long                     'the minimum splitters' size

'---------------------
' Class Custom Events
'---------------------

'Description: Occurs when the BackColor property of the class is changed
'Argument   : IdSplitter (a value that uniquely identifies a splitter)
Public Event BackColorChange(ByVal IdSplitter As Long)
Attribute BackColorChange.VB_Description = "Sets/returns the background color used to display the splitters"

'Description: Occurs when the Enable property of the class is changed
'Argument   : IdSplitter (a value that uniquely identifies a splitter)
Public Event EnableChange(ByVal IdSplitter As Long)

'----------------------------------
' Class Constructor and Destructor
'----------------------------------

Private Sub Class_Initialize()
  Set mcolSplitters = New Collection
  
  mlngActiveColor = mconDefaultActiveColor
  mblnClipCursor = mconDefaultClipCursor
  mblnEnable = mconDefaultEnable
  mblnLiveUpdate = mconDefaultLiveUpdate
  mlngMinimumSize = IIf(Screen.TwipsPerPixelX > Screen.TwipsPerPixelY, _
                        Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)
  mlngSize = DefaultSize
  
  mblnEventDisable = False
End Sub

Private Sub Class_Terminate()
  Set mcolSplitters = Nothing
End Sub

'-----------------------------------------
' Collection Class Properties and Methods
'-----------------------------------------

' Purpose    : Adds new items to the collection
' Effects    : * If does not exist, the top-, right-, bottom- and left-side
'                splitters of control octl have been added and their necessary
'                properties have been initialized
'              * Control octl have been shrunk to give space for the splitters
'                around it
'              * This method may left several splitters' id marked deleted that
'                should be removed after all the splitters have been added
' Inputs     : * octl (the control that trigger the splitters addition)
'              * octls (the controls collection of control octl)
Friend Sub Add(ByRef octl As clsControl, ByVal octls As clsControls)
Attribute Add.VB_Description = "Adds new items to the collection"
  Dim lngPos As Long                   'a temporary variable used in determining
                                       '               Splitter 's movement area
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  Dim osplNew As clsSplitter 'the new item which will be added to the collection
  
  '-- If the splitters for control octl already exist, assigns the splitters to
  '   controls octl's IdSplTop, IdSpRight, IdSplBottom or IdSplLeft property and
  '   adjust the splitters' size
  For Each ospl In Me
    If (ospl.Orientation = orHorizontal) And (octl.Top = ospl.Yc) Then
      If Not IsExistControlBetween(octlsSource:=octls, _
                                   octlSource:=octl, osplSource:=ospl) Then
        octl.IdSplTop = ospl.Id
        If octl.Left < ospl.Left Then ospl.Left = octl.Left
        If octl.Right > ospl.Right Then ospl.Right = octl.Right
      End If
    End If
    If (ospl.Orientation = orVertical) And (octl.Right = ospl.Xc) Then
      If Not IsExistControlBetween(octlsSource:=octls, _
                                   octlSource:=octl, osplSource:=ospl) Then
        octl.IdSplRight = ospl.Id
        If octl.Top < ospl.Top Then ospl.Top = octl.Top
        If octl.Bottom > ospl.Bottom Then ospl.Bottom = octl.Bottom
      End If
    End If
    If (ospl.Orientation = orHorizontal) And (octl.Bottom = ospl.Yc) Then
      If Not IsExistControlBetween(octlsSource:=octls, _
                                   octlSource:=octl, osplSource:=ospl) Then
        octl.IdSplBottom = ospl.Id
        If octl.Left < ospl.Left Then ospl.Left = octl.Left
        If octl.Right > ospl.Right Then ospl.Right = octl.Right
      End If
    End If
    If (ospl.Orientation = orVertical) And (octl.Left = ospl.Xc) Then
      If Not IsExistControlBetween(octlsSource:=octls, _
                                   octlSource:=octl, osplSource:=ospl) Then
        octl.IdSplLeft = ospl.Id
        If octl.Top < ospl.Top Then ospl.Top = octl.Top
        If octl.Bottom > ospl.Bottom Then ospl.Bottom = octl.Bottom
      End If
    End If
  Next
    
  '-- If the splitters for control octl doesn't exist, create a new splitter,
  '   add it to the collection, and assigns it to controls octl's IdSplTop,
  '   IdSplRight, IdSplBottom or IdSplLeft
  mblnEventDisable = True
  If (octl.IdSplTop = gconUninitialized) And (octl.Top <> mlngTop) Then
    Set osplNew = New clsSplitter
    With osplNew
      Set .Parent = Me
      .Id = mcolSplitters.Count
      .Orientation = orHorizontal
      .Xc = octl.Xc
      .Yc = octl.Top
      .Width = octl.Width
      .Height = mlngSize
      .ActiveColor = mlngActiveColor
      .BackColor = mlngBackColor
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplTop = .Id
    End With
    Set osplNew = Nothing
  End If
  If (octl.IdSplRight = gconUninitialized) And (octl.Right <> mlngRight) Then
    Set osplNew = New clsSplitter
    With osplNew
      Set .Parent = Me
      .Id = mcolSplitters.Count
      .Orientation = orVertical
      .Xc = octl.Right
      .Yc = octl.Yc
      .Width = mlngSize
      .Height = octl.Height
      .ActiveColor = mlngActiveColor
      .BackColor = mlngBackColor
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplRight = .Id
    End With
    Set osplNew = Nothing
  End If
  If (octl.IdSplBottom = gconUninitialized) And (octl.Bottom <> mlngBottom) Then
    Set osplNew = New clsSplitter
    With osplNew
      Set .Parent = Me
      .Id = mcolSplitters.Count
      .Orientation = orHorizontal
      .Xc = octl.Xc
      .Yc = octl.Bottom
      .Width = octl.Width
      .Height = mlngSize
      .ActiveColor = mlngActiveColor
      .BackColor = mlngBackColor
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplBottom = .Id
    End With
    Set osplNew = Nothing
  End If
  If (octl.IdSplLeft = gconUninitialized) And (octl.Left <> mlngLeft) Then
    Set osplNew = New clsSplitter
    With osplNew
      Set .Parent = Me
      .Id = mcolSplitters.Count
      .Orientation = orVertical
      .Xc = octl.Left
      .Yc = octl.Yc
      .Width = mlngSize
      .Height = octl.Height
      .ActiveColor = mlngActiveColor
      .BackColor = mlngBackColor
      .ClipCursor = mblnClipCursor
      .Enable = mblnEnable
      .LiveUpdate = mblnLiveUpdate
      mcolSplitters.Add Item:=osplNew, Key:=CStr(.Id)
      octl.IdSplLeft = .Id
    End With
    Set osplNew = Nothing
  End If
  mblnEventDisable = False
  
  '-- Shrinks the control octl's size to give space for the splitters
  octl.Left = octl.Left + IIf(octl.Left <> mlngLeft, mlngSize \ 2, 0)
  octl.Top = octl.Top + IIf(octl.Top <> mlngTop, mlngSize \ 2, 0)
  octl.Right = octl.Right - IIf(octl.Right <> mlngRight, mlngSize \ 2, 0)
  octl.Bottom = octl.Bottom - IIf(octl.Bottom <> mlngBottom, mlngSize \ 2, 0)
  
  '-- Assigns all controls and splitters' id that effected by splitter
  '   octl.IdSplTop, octl.IdSplRight, octl.IdSplBottom and octl.IdSplLeft to the
  '   related splitter
  If octl.IdSplTop <> gconUninitialized Then
    With Me(octl.IdSplTop)
      .IdsCtlBottom.Add octl.Id
      .IdsSplBottom.Add octl.IdSplRight
      If .IdsSplTop.IsExist(octl.IdSplRight) Then
        .IdsSplTop.MarkDeleted octl.IdSplRight
        .IdsSplBottom.MarkDeleted octl.IdSplRight
      End If
      .IdsSplBottom.Add octl.IdSplLeft
      If .IdsSplTop.IsExist(octl.IdSplLeft) Then
        .IdsSplTop.MarkDeleted octl.IdSplLeft
        .IdsSplBottom.MarkDeleted octl.IdSplLeft
      End If
      lngPos = octl.Bottom - octl.MinHeight
      If lngPos < .MaxYc Then
        .MaxYc = lngPos
        .IdCtlFriendBottom = octl.Id
      End If
    End With
  End If
  If octl.IdSplRight <> gconUninitialized Then
    With Me(octl.IdSplRight)
      .IdsCtlLeft.Add octl.Id
      .IdsSplLeft.Add octl.IdSplTop
      If .IdsSplRight.IsExist(octl.IdSplTop) Then
        .IdsSplRight.MarkDeleted octl.IdSplTop
        .IdsSplLeft.MarkDeleted octl.IdSplTop
      End If
      .IdsSplLeft.Add octl.IdSplBottom
      If .IdsSplRight.IsExist(octl.IdSplBottom) Then
        .IdsSplRight.MarkDeleted octl.IdSplBottom
        .IdsSplLeft.MarkDeleted octl.IdSplBottom
      End If
      lngPos = octl.Left + octl.MinWidth
      If lngPos > .MinXc Then
        .MinXc = lngPos
        .IdCtlFriendLeft = octl.Id
      End If
    End With
  End If
  If octl.IdSplBottom <> gconUninitialized Then
    With Me(octl.IdSplBottom)
      .IdsCtlTop.Add octl.Id
      .IdsSplTop.Add octl.IdSplRight
      If .IdsSplBottom.IsExist(octl.IdSplRight) Then
        .IdsSplBottom.MarkDeleted octl.IdSplRight
        .IdsSplTop.MarkDeleted octl.IdSplRight
      End If
      .IdsSplTop.Add octl.IdSplLeft
      If .IdsSplBottom.IsExist(octl.IdSplLeft) Then
        .IdsSplBottom.MarkDeleted octl.IdSplLeft
        .IdsSplTop.MarkDeleted octl.IdSplLeft
      End If
      lngPos = octl.Top + octl.MinHeight
      If lngPos > .MinYc Then
        .MinYc = lngPos
        .IdCtlFriendTop = octl.Id
      End If
    End With
  End If
  If octl.IdSplLeft <> gconUninitialized Then
    With Me(octl.IdSplLeft)
      .IdsCtlRight.Add octl.Id
      .IdsSplRight.Add octl.IdSplTop
      If .IdsSplLeft.IsExist(octl.IdSplTop) Then
        .IdsSplLeft.MarkDeleted octl.IdSplTop
        .IdsSplRight.MarkDeleted octl.IdSplTop
      End If
      .IdsSplRight.Add octl.IdSplBottom
      If .IdsSplLeft.IsExist(octl.IdSplBottom) Then
        .IdsSplLeft.MarkDeleted octl.IdSplBottom
        .IdsSplRight.MarkDeleted octl.IdSplBottom
      End If
      lngPos = octl.Right - octl.MinWidth
      If lngPos < .MaxXc Then
        .MaxXc = lngPos
        .IdCtlFriendRight = octl.Id
      End If
    End With
  End If
End Sub

' Purpose    : Removes all items in the collection
' Effect     : As specified
Friend Sub Clear()
Attribute Clear.VB_Description = "Removes all items in the collection"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    Remove ospl
  Next
End Sub

' Purpose    : Returns the number of items in the collection
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
  Count = mcolSplitters.Count
End Property

' Purpose    : Returns an item in the collection which has key IdSpl
' Assumption : Key IdSpl exists in the collection
' Input      : IdSpl
' Return     : As specified
Public Property Get Item(ByVal IdSpl As Long) As clsSplitter
Attribute Item.VB_Description = "Returns an item in the collection which has key IdSpl"
Attribute Item.VB_UserMemId = 0
  Set Item = mcolSplitters.Item(CStr(IdSpl))
End Property

' Purpose    : Enables For Each ... Next enumeration
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enables For Each ... Next enumeration"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolSplitters.[_NewEnum]
End Property

' Purpose    : Removes an item which has key IdSpl from the collection
' Assumption : Key IdSpl exists in the collection
' Effect     : As specified
' Input      : IdSpl
Friend Sub Remove(ByVal IdSpl As Long)
Attribute Remove.VB_Description = "Removes an item which has key IdSpl from the collection"
  mcolSplitters.Remove CStr(IdSpl)
End Sub

'------------------
' Class Properties
'------------------

' Purpose    : Sets the background color used to display the splitter when the
'              user moves it in none live update mode
' Effect     : As specified
' Input      : lngActiveColor (the new ActiveColor property value)
Friend Property Let ActiveColor(ByVal lngActiveColor As OLE_COLOR)
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    ospl.ActiveColor = lngActiveColor
  Next
  mlngActiveColor = lngActiveColor
End Property

' Purpose    : Returns the background color used to display the splitters when
'              the user moves it in none live update mode
' Return     : As specified
Friend Property Get ActiveColor() As OLE_COLOR
  ActiveColor = mlngActiveColor
End Property

' Purpose    : Sets the background color used to display the splitters
' Effect     : As specified
' Inputs     : * lngIdSplitter (a value that uniquely identifies a splitter)
'              * lngBackColor (the new BackColor property value)
Friend Property Let BackColor(Optional ByVal _
                                lngIdSplitter As Long = gconUninitialized, _
                              ByVal lngBackColor As OLE_COLOR)
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  If lngIdSplitter = gconUninitialized Then
    mblnEventDisable = True
    For Each ospl In Me
      ospl.BackColor = lngBackColor
    Next
    mblnEventDisable = False
    mlngBackColor = lngBackColor
  ElseIf Not mblnEventDisable Then
    RaiseEvent BackColorChange(lngIdSplitter)
  End If
End Property

' Purpose    : Returns the background color used to display the splitters
' Input      : lngIdSplitter (a value that uniquely identifies a splitter)
' Return     : As specified
Friend Property Get BackColor(Optional ByVal _
                                lngIdSplitter As Long = gconUninitialized _
                             ) As OLE_COLOR
  BackColor = mlngBackColor
End Property

' Purpose    : Sets the distance between the bottom edge of the splitters frame
'              area and the top edge of VB Control Manager control
' Effect     : As specified
' Input      : lngBottom (the new Bottom propety value)
Friend Property Let Bottom(ByVal lngBottom As Long)
Attribute Bottom.VB_Description = "Returns or sets the distance between the bottom edge of the splitters frame area and the top edge of VB Splitter control"
  mlngBottom = lngBottom
End Property

' Purpose    : Returns the distance between the bottom edge of the splitters
'              frame area and the top edge of VB Control Manager control
' Return     : As specified
Friend Property Get Bottom() As Long
  Bottom = mlngBottom
End Property

' Purpose    : Sets a value that determines whether the mouse pointer is
'              confined to the items' region
' Effects    : * The ClipCursor property of all items in the collection has been
'                set to blnClipCursor
'              * As specified
' Input      : blnClipCursor (the new ClipCursor property value)
Friend Property Let ClipCursor(ByVal blnClipCursor As Boolean)
Attribute ClipCursor.VB_Description = "Returns or sets a value that determines whether the mouse pointer is confined to the items' region\r\n"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    ospl.ClipCursor = blnClipCursor
  Next
  mblnClipCursor = blnClipCursor
End Property

' Purpose    : Returns a value that determines whether the mouse pointer is
'              confined to the items' region
' Return     : As specified
' Note       : Each item has its own ClipCursor properties. This property only
'              return the last ClipCursor property value set by the users. If
'              the users set the property individually for each item, this
'              property may not return the correct ClipCursor property value.
Friend Property Get ClipCursor() As Boolean
  ClipCursor = mblnClipCursor
End Property

' Purpose    : Returns the default ActiveColor property
Friend Property Get DefaultActiveColor() As Boolean
  DefaultActiveColor = mconDefaultActiveColor
End Property

' Purpose    : Returns the default ClipCursor property
Friend Property Get DefaultClipCursor() As Boolean
  DefaultClipCursor = mconDefaultClipCursor
End Property

' Purpose    : Returns the default Enable property
Friend Property Get DefaultEnable() As Boolean
  DefaultEnable = mconDefaultEnable
End Property

' Purpose    : Returns the default LiveUpdate property
Friend Property Get DefaultLiveUpdate() As Boolean
  DefaultLiveUpdate = mconDefaultLiveUpdate
End Property

' Purpose    : Returns the default Size property
Friend Property Get DefaultSize() As Long
  DefaultSize = mconDefaultSize * Screen.TwipsPerPixelX
End Property

' Purpose    : Sets a value that determines whether the splitter is movable
' Effects    : * The Enable property of all items in the collection has been
'                set to blnEnable
'              * As specified
' Inputs     : * lngIdSplitter (a value that uniquely identifies a splitter)
'              * blnEnable (the new Enable property value)
Friend Property Let Enable(Optional ByVal _
                             lngIdSplitter As Long = gconUninitialized, _
                           ByVal blnEnable As Boolean)
Attribute Enable.VB_Description = "Returns or sets a value that determines whether the items are movable"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  If lngIdSplitter = gconUninitialized Then
    mblnEventDisable = True
    For Each ospl In Me
      ospl.Enable = blnEnable
    Next
    mblnEventDisable = False
    mblnEnable = blnEnable
  ElseIf Not mblnEventDisable Then
    RaiseEvent EnableChange(lngIdSplitter)
  End If
End Property

' Purpose    : Returns a value that determines whether the items are movable
' Return     : As specified
' Note       : Each item has its own Enable properties. This property only
'              return the last Enable property value set by the users. If the
'              users set the property individually for each item, this property
'              may not return the correct Enable property value.
Friend Property Get Enable(Optional ByVal _
                             lngIdSplitter As Long = gconUninitialized _
                          ) As Boolean
  Enable = mblnEnable
End Property

' Purpose    : Sets the height of the splitters frame area
' Effect     : As specified
' Input      : lngHeight(the new Height property value)
Friend Property Let Height(ByVal lngHeight As Long)
Attribute Height.VB_Description = "Returns or sets the height of the splitters frame area"
  mlngBottom = mlngTop + lngHeight
End Property

' Purpose    : Returns the height of the splitters frame area
' Return     : As specified
Friend Property Get Height() As Long
  Height = mlngBottom - mlngTop
End Property

' Purpose    : Sets the distance between the left edge of the splitters frame
'              area and the left edge of VB Control Manager control
' Effect     : As specified
' Input      : lngLeft (the new Left propety value)
Friend Property Let Left(ByVal lngLeft As Long)
Attribute Left.VB_Description = "Returns or sets the distance between the left edge of the splitters frame area and the left edge of VB Splitter control"
  mlngLeft = lngLeft
End Property

' Purpose    : Returns the distance between the left edge of the splitters frame
'              area and the left edge of VB Control Manager control
' Return     : As specified
Friend Property Get Left() As Long
  Left = mlngLeft
End Property

' Purpose    : Sets a value that determines whether the controls should be
'              resized as the item is moved
' Effects    : * The LiveUpdate property of all items in the collection has been
'                set to blnLiveUpdate
'              * As specified
' Input      : blnLiveUpdate (the new LiveUpdate property value)
Friend Property Let LiveUpdate(ByVal blnLiveUpdate As Boolean)
Attribute LiveUpdate.VB_Description = "Returns or sets a value that determines whether the controls should be resized as the  item is moved"
  Dim ospl As clsSplitter           'for enumerating all items in the collection
  
  For Each ospl In Me
    ospl.LiveUpdate = blnLiveUpdate
  Next
  mblnLiveUpdate = blnLiveUpdate
End Property

' Purpose    : Returns a value that determines whether the controls should be
'              resized as the item is moved
' Return     : As specified
' Note       : Each item has its own LiveUpdate properties. This property only
'              return the last LiveUpdate property value set by the users. If
'              the users set the property individually for each item, this
'              property may not return the correct LiveUpdate property value.
Friend Property Get LiveUpdate() As Boolean
  LiveUpdate = mblnLiveUpdate
End Property

' Purpose    : Returns the minimum items' size
Friend Property Get MinimumSize() As Long
Attribute MinimumSize.VB_Description = "Returns the minimum items' size"
  MinimumSize = mlngMinimumSize
End Property

' Purpose    : Sets the distance between the right edge of the splitters frame
'              area and the left edge of VB Control Manager control
' Effect     : As specified
' Input      : lngRight (the new Right propety value)
Friend Property Let Right(ByVal lngRight As Long)
Attribute Right.VB_Description = "Returns or sets the distance between the right edge of the splitters frame area and the left edge of VB Splitter control"
  mlngRight = lngRight
End Property

' Purpose    : Returns the distance between the right edge of the splitters
'              frame area and the left edge of VB Control Manager control
' Return     : As specified
Friend Property Get Right() As Long
  Right = mlngRight
End Property

' Purpose    : Sets the items' size
' Effects    : * The Height or Width property (depend on the Orientation
'                property) of all items in the collection have been set to
'                lngSize
'              * As specified
' Input      : lngSize (the new Size property value)
Friend Property Let Size(ByVal lngSize As Long)
Attribute Size.VB_Description = "Returns or sets the items' size"
  Dim ospl As clsSplitter           'for enumerating all items in the collection

  For Each ospl In Me
    Select Case ospl.Orientation
      Case orHorizontal
        ospl.Height = lngSize
      Case orVertical
        ospl.Width = lngSize
    End Select
  Next
  mlngSize = lngSize
End Property

' Purpose    : Returns the items' size
' Return     : As specified
' Note       : Each item has its own Size properties. This property only return
'              the last Size property value set by the users. If the users set
'              the property individually for each item, this property may not
'              return the correct Size property value.
Friend Property Get Size() As Long
  Size = mlngSize
End Property

' Purpose    : Sets the distance between the top edge of the splitters frame
'              area and the top edge of VB Control Manager control
' Effect     : As specified
' Input      : lngTop (the new Top propety value)
Friend Property Let Top(ByVal lngTop As Long)
Attribute Top.VB_Description = "Returns or sets the distance between the top edge of the splitters frame area and the top edge of VB Splitter control"
  mlngTop = lngTop
End Property

' Purpose    : Returns the distance between the top edge of the splitters frame
'              area and the top edge of VB Control Manager control
' Return     : As specified
Friend Property Get Top() As Long
  Top = mlngTop
End Property

' Purpose    : Sets the width of the splitters frame area
' Effect     : As specified
' Input      : lngWidth(the new Width property value)
Friend Property Let Width(ByVal lngWidth As Long)
Attribute Width.VB_Description = "Returns or sets the width of the splitters frame area"
  mlngRight = mlngLeft + lngWidth
End Property

' Purpose    : Returns the width of the splitters frame area
' Return     : As specified
Friend Property Get Width() As Long
  Width = mlngRight - mlngLeft
End Property

'---------------
' Class Methods
'---------------

' Purpose    : Returns value indicating whether an id is exist in the collection
' Input      : Id
' Return     : As specified
Public Function IsExist(Id As Long) As Boolean
Attribute IsExist.VB_Description = "Returns a value indicating whether certain index that uniquely identifies a virtual splitter exist in the collection"
  Dim blnIsExist As Boolean
  Dim ospl As clsSplitter
  
  blnIsExist = False
  For Each ospl In Me
    If ospl.Id = Id Then
      blnIsExist = True
      Exit For
    End If
  Next
  IsExist = blnIsExist
End Function

'-------------------
' Private Functions
'-------------------

' Purpose    : Returns a value indicating whether there is a control between
'              control octlSource and splitter osplSource in controls collection
'              octlsSource
' Inputs     : * octlsSource
'              * octlSource
'              * osplSource
' Return     : As specified
Private Function IsExistControlBetween(ByVal octlsSource As clsControls, _
                                       ByVal octlSource As clsControl, _
                                       ByVal osplSource As clsSplitter)
  Dim blnIsExistControlBetween As Boolean                        'returned value
  Dim octl As clsControl            'for enumerating all items in the collection
  
  blnIsExistControlBetween = False
  For Each octl In octlsSource
    Select Case osplSource.Orientation
      Case orHorizontal
        If (octl.Top <= osplSource.Top) And _
           (octl.Bottom >= osplSource.Bottom) And _
           (((octl.Left >= osplSource.Right) And _
             (octl.Right <= octlSource.Left)) Or _
            ((octl.Left >= octlSource.Right) And _
             (octl.Right <= osplSource.Left))) Then
          blnIsExistControlBetween = True
          Exit For
        End If
      Case orVertical
        If (octl.Left <= osplSource.Left) And _
           (octl.Right >= osplSource.Right) And _
           (((octl.Top >= osplSource.Bottom) And _
             (octl.Bottom <= octlSource.Top)) Or _
            ((octl.Top >= octlSource.Bottom) And _
             (octl.Bottom <= osplSource.Top))) Then
          blnIsExistControlBetween = True
          Exit For
        End If
    End Select
  Next
  IsExistControlBetween = blnIsExistControlBetween
End Function
