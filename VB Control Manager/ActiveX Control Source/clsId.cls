VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class module to represent an Id"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************************
'** File Name     : clsId.cls                                                 **
'** Language      : Visual Basic 6.0                                          **
'** Author        : Theo Zacharias (theo_yz@yahoo.com)                        **
'** Description   : A class module to represent an Id                         **
'** Usage         : Represents Ids of class clsControl's and clsSplitter's    **
'**                 instances                                                 **
'** Dependencies  : mdlGeneral                                                **
'** Public Members:                                                           **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Property   : Id (def. r/o)                                            **
'**   * Methods    : -                                                        **
'**   * Events     : -                                                        **
'** Friend Members:                                                           **
'**   * Collections: -                                                        **
'**   * Objects    : -                                                        **
'**   * Property   : Id (def. r/w)                                            **
'**   * Methods    : -                                                        **
'** Last modified on October 10, 2003                                         **
'*******************************************************************************

Option Explicit

'--- Property Variable
Private mId As Long

'-------------------
' Class Constructor
'-------------------

Private Sub Class_Initialize()
  mId = gconUninitialized
End Sub

'----------------
' Class Property
'----------------

' Purpose    : Sets the id
' Effect     : As specified
' Input      : lngId (the new Id property value)
Friend Property Let Id(ByVal lngId As Long)
  mId = lngId
End Property

' Purpose    : Returns the id
' Return     : As specified
Public Property Get Id() As Long
Attribute Id.VB_Description = "Returns the id"
Attribute Id.VB_UserMemId = 0
  Id = mId
End Property

