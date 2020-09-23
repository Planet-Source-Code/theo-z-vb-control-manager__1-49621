VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
'** File Name     : TitleBar.cls                                              **
'** Language      : Visual Basic 6.0                                          **
'** Author        : Theo Zacharias (theo_yz@yahoo.com)                        **
'** Description   : An abstract class to provide the template for the         **
'**                 TitleBar interface                                        **
'** Last modified on September 1, 2003                                        **
'*******************************************************************************

Option Explicit

'--- Class Property
Public CloseVisible As Boolean
Public Height As Long
Public Visible As Boolean

