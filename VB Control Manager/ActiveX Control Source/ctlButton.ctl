VERSION 5.00
Begin VB.UserControl ctlButton 
   ClientHeight    =   312
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   276
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   312
   ScaleWidth      =   276
   Begin VB.Image imgButton 
      Height          =   255
      Left            =   15
      Stretch         =   -1  'True
      Top             =   30
      Width           =   225
   End
End
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : ctlButton.ctl                                               **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A custom button ActiveX control with hover effect           **
'** Usage       : As close button in ctlTitleBar ActiveX control              **
'** Dependencies: mdlAPI                                                      **
'** Members     :                                                             **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Property   : Picture (def. r/w)                                       **
'**   * Methods    : -                                                        **
'**   * Events     : Click                                                    **
'** Last modified on November 13, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Property Variables
Private mspPicture As StdPicture

'--- PropBag Names
Private Const mconPicture = "Picture"

'-------------------------------
' ActiveX Control Custom Events
'-------------------------------

'Description: Occurs when the user presses and then releases a mouse button over
'             the control
Public Event Click()

'--------------------------------------------
' ActiveX Control Constructor and Destructor
'--------------------------------------------

Private Sub UserControl_Initialize()
  Set mspPicture = New StdPicture
End Sub

Private Sub UserControl_Terminate()
  Set mspPicture = Nothing
End Sub

'-----------------------------------
' ActiveX Control Properties Events
'-----------------------------------

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set mspPicture = PropBag.ReadProperty(Name:=mconPicture, _
                                        DefaultValue:=Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty Name:=mconPicture, _
                        Value:=mspPicture, DefaultValue:=Nothing
End Sub

'-------------------------------
' Others ActiveX Control Events
'-------------------------------

' Purpose    : Refer to UserControl_MouseDown
' Effect     : As specified
' Inputs     : Button, Shift, X, Y
Private Sub imgButton_MouseDown(Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  UserControl_MouseDown Button, Shift, X, Y
End Sub

' Purpose    : Refer to UserControl_MouseDown
' Effect     : As specified
' Inputs     : Button, Shift, X, Y
Private Sub imgButton_MouseMove(Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  UserControl_MouseMove Button, Shift, X, Y
End Sub

' Purpose    : Refer to UserControl_MouseDown
' Effect     : As specified
' Inputs     : Button, Shift, X, Y
Private Sub imgButton_MouseUp(Button As Integer, _
                              Shift As Integer, X As Single, Y As Single)
  UserControl_MouseUp Button, Shift, X, Y
End Sub

' Purpose    : Draws pressed effect on the control
' Effect     : As specified
' Inputs     : Button, Shift, X, Y
Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, X As Single, Y As Single)
  Dim rec As mdlAPI.RECT     'rectangle area at where to draw the pressed effect

  If Button = vbLeftButton Then
    mdlAPI.SetRect lpRect:=rec, X1:=0, Y1:=0, _
                   X2:=UserControl.ScaleWidth \ Screen.TwipsPerPixelX, _
                   Y2:=UserControl.ScaleHeight \ Screen.TwipsPerPixelY
    mdlAPI.DrawEdge hdc:=UserControl.hdc, qrc:=rec, _
                    edge:=mdlAPI.EDGE_SUNKEN, grfFlags:=mdlAPI.BF_RECT
  End If
End Sub

' Purpose    : Draws hover effect on the control
' Effect     : As specified
' Inputs     : Button, Shift, X, Y
Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, X As Single, Y As Single)
  Dim pos As mdlAPI.POINTAPI      'used with mdlAPI.WindowFromPoint to determine
                                  '  whether the mouse pointer is on the control
  Dim rec As mdlAPI.RECT       'rectangle area at where to draw the hover effect

  If Button = 0 Then
    If mdlAPI.GetCapture() <> UserControl.hwnd Then
      mdlAPI.SetCapture UserControl.hwnd
      mdlAPI.SetRect lpRect:=rec, X1:=0, Y1:=0, _
                     X2:=UserControl.ScaleWidth \ Screen.TwipsPerPixelX, _
                     Y2:=UserControl.ScaleHeight \ Screen.TwipsPerPixelY
      mdlAPI.DrawEdge hdc:=UserControl.hdc, qrc:=rec, _
                      edge:=mdlAPI.EDGE_RAISED, grfFlags:=mdlAPI.BF_RECT
    Else
      pos.X = X
      pos.Y = Y
      mdlAPI.ClientToScreen hwnd:=UserControl.hwnd, lpPoint:=pos
      If mdlAPI.WindowFromPoint(pos.X, pos.Y) <> UserControl.hwnd Then
        UserControl.Cls
        mdlAPI.ReleaseCapture
      End If
    End If
  End If
End Sub

' Purposes   : * Clear all effect on the control
'              * Raise click event if the mouse pointer is on the control
' Effect     : As specified
' Inputs     : Button, Shift, X, Y
Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, X As Single, Y As Single)
  Dim pos As mdlAPI.POINTAPI
  
  If Button = vbLeftButton Then UserControl.Cls
  
  If (Button = vbLeftButton) And _
     (0 <= X) And (X <= UserControl.ScaleWidth) And _
     (0 <= Y) And (Y <= UserControl.ScaleHeight) Then _
    RaiseEvent Click
End Sub

' Purpose    : Refresh the picture on the imgButton
' Effect     : As specified
Private Sub UserControl_Paint()
  Set imgButton.Picture = mspPicture
End Sub

' Purpose    : Adjusts the imgButton size to match the control size
' Effect     : As specified
Private Sub UserControl_Resize()
  With imgButton
    .Left = 3 * Screen.TwipsPerPixelX
    .Top = 3 * Screen.TwipsPerPixelY
    .Width = UserControl.ScaleWidth - (6 * Screen.TwipsPerPixelX)
    .Height = UserControl.ScaleHeight - (6 * Screen.TwipsPerPixelY)
  End With
End Sub

'----------------------------
' ActiveX Control Properties
'----------------------------

' Purpose    : Sets a graphic to be displayed in the imgButton
' Effect     : As specified
' Input      : spPicture (the new Picture property value)
Public Property Set Picture(spPicture As StdPicture)
  Set mspPicture = spPicture
  PropertyChanged mconPicture
  
  UserControl_Paint
End Property

' Purpose    : Returns a graphic to be displayed in the imgButton
' Return     : As specified
Public Property Get Picture() As StdPicture
Attribute Picture.VB_UserMemId = 0
  Set Picture = mspPicture
End Property
