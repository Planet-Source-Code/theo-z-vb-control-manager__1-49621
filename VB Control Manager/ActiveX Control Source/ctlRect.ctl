VERSION 5.00
Begin VB.UserControl ctlRect 
   Appearance      =   0  'Flat
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2532
   ForeColor       =   &H00000000&
   ScaleHeight     =   2316
   ScaleWidth      =   2532
End
Attribute VB_Name = "ctlRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : ctlRect.ctl                                                 **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A custom rectangle ActiveX control                          **
'** Usage       : As the rectangle that represent the moving control          **
'** Dependencies: mdlGeneral, mdlAPI                                          **
'** Members     :                                                             **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Properties : -                                                        **
'**   * Method     : UpdatePosition                                           **
'**   * Events     : -                                                        **
'** Last modified on November 1, 2003                                         **
'*******************************************************************************

'-- Variables to save the original rectangle size and position
Private mlngRelLeft As Long
Private mlngRelTop As Long
Private mlngWidth As Long
Private mlngHeight As Long

Option Explicit

'------------------------------
' Others ActiveX Control Event
'------------------------------

' Purpose    : Saves the size and position (relative to the current position of
'              the cursor) of the original rectangle
' Effect     : As specified
Private Sub UserControl_Show()
  Dim uposCursor As mdlAPI.POINTAPI

  uposCursor = GetCursorRelPos(hwnd:=UserControl.hwnd)
  mlngRelLeft = uposCursor.X
  mlngRelTop = uposCursor.Y
  mlngWidth = Extender.Width
  mlngHeight = Extender.Height
End Sub

'------------------------
' ActiveX Control Method
'------------------------

' Purpose    : Update rectangle position based on its original size and position
'              (relative to the current position of the cursor when the first
'              time the rectangle is shown)
' Effect     : As specified
Public Sub UpdatePosition()
  Dim uposCursor As mdlAPI.POINTAPI
  
  uposCursor = GetCursorRelPos(hwnd:=UserControl.hwnd)
  Extender.Move Extender.Left + uposCursor.X - mlngRelLeft, _
                Extender.Top + uposCursor.Y - mlngRelTop, _
                mlngWidth, mlngHeight
End Sub

