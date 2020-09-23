Attribute VB_Name = "mdlGeneral"
Attribute VB_Description = "A module to handle general operations"
'*******************************************************************************
'** File Name   : mdlGeneral.bas                                              **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Dependency  : mdlAPI                                                      **
'** Description : A module to handle general operations                       **
'** Last modified on November 13, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Resource File Constants
' Splitter Cursor
Public Const gconCurHSplitter = 101                  'horizontal splitter cursor
Public Const gconCurVSplitter = 102                    'vertical splitter cursor

' Error Message Index
Public Enum genmErrNumber
  errBuildSplitters = 2000
  errSelfContained = 2001
  errMoveSplitter = 2002
  errResizeSplitter = 2003
  errMoveControlRoom = 2004
  errIdControl = 2005
  errIdSplitter = 2006
  errMoveControlClosed = 2007
End Enum

'--- Other Constants
Public Const gconUninitialized = -1      'represent the Id which is not exist or
                                         '           hasn't been initialized yet
Public Const gconLngInfinite = 2147483647

'--- Variable Declaration
Public gstrControlName As String                 'the name of VB Control Manager
                                                 '              control instance

' Purpose    : Retrieves the cursor's position in twips relative to certain
'              window
' Assumptions: Window hwnd exist (if hwnd is not omitted)
' Input      : hwnd (the window where the cursor will be retrieved relative to;
'                    if ommited, the screen will be used as the window)
' Return     : As specified
Public Function GetCursorRelPos( _
                  Optional ByVal hwnd As Long = gconUninitialized _
                ) As mdlAPI.POINTAPI
  Dim uposGetCursorRelPos As mdlAPI.POINTAPI
                
  mdlAPI.GetCursorPos uposGetCursorRelPos
  If hwnd <> gconUninitialized Then
    mdlAPI.ScreenToClient hwnd:=hwnd, lpPoint:=uposGetCursorRelPos
    With uposGetCursorRelPos
      .X = .X * Screen.TwipsPerPixelX
      .Y = .Y * Screen.TwipsPerPixelY
    End With
  End If
  GetCursorRelPos = uposGetCursorRelPos
End Function

' Purpose    : Gets minimum value of numbers in array lngValue()
' Assumptions: * Option base is set to 0
'              * Array lngValue() contains only numbers
' Input      : vntValue()
' Return     : * If no parameters passed to vntValue(), returns Empty
'              * Otherwise, returns as specified
Public Function GetMin(ParamArray vntValue() As Variant) As Variant
Attribute GetMin.VB_Description = "Gets minimum value of numbers in array lngValue()"
  Dim i As Long                              'for iterating the parameters value
  Dim vntGetMin As Variant                                       'returned value
  
  If Not IsMissing(vntValue) Then
    vntGetMin = vntValue(0)
    For i = 1 To UBound(vntValue)
      If vntValue(i) < vntGetMin Then vntGetMin = vntValue(i)
    Next
    GetMin = vntGetMin
  End If
End Function

' Purpose    : Raises custom error udeErrNumber
' Assumptions: * Error message udeErrNumber exists in the resource file
'              * Global variable gstrControlName has been initialized
' Inputs     : * udeErrNumber
'              * strSource (the location in form ClassName.RoutinesName where
'                the error occur
' Effect     : As specified
Public Sub RaiseError(ByVal udeErrNumber As genmErrNumber, _
                      Optional ByVal strSource As String = "")
Attribute RaiseError.VB_Description = "Raises custom error udeErrNumber"
  If strSource <> "." Then strSource = "." & strSource
  Err.Raise Number:=(vbObjectError + udeErrNumber), _
            Source:=gstrControlName & strSource, _
            Description:=LoadResString(udeErrNumber)
End Sub
