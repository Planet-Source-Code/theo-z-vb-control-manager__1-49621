VERSION 5.00
Begin VB.Form dlgAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB Control Manager"
   ClientHeight    =   2460
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5220
   ClipControls    =   0   'False
   Icon            =   "dlgAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -45
      TabIndex        =   3
      Top             =   1755
      WhatsThisHelpID =   10385
      Width           =   5385
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   1995
      WhatsThisHelpID =   10379
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "This ActiveX Control is freeware and open source. You may freely use and modify any part of the code for your personal needs."
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   255
      TabIndex        =   6
      Top             =   975
      Width           =   4740
   End
   Begin VB.Label Label4 
      Caption         =   " Thank you for using this control."
      Height          =   255
      Left            =   195
      TabIndex        =   5
      Top             =   1380
      Width           =   2505
   End
   Begin VB.Label lblEMail 
      Caption         =   "(theo_yz@yahoo.com)"
      Height          =   255
      Left            =   2910
      MouseIcon       =   "dlgAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   510
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "VB Control Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   915
      TabIndex        =   1
      Top             =   180
      WhatsThisHelpID =   10382
      Width           =   2745
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Written by Theo Zacharias"
      Height          =   225
      Left            =   930
      TabIndex        =   2
      Top             =   510
      WhatsThisHelpID =   10383
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   135
      Picture         =   "dlgAbout.frx":0316
      Stretch         =   -1  'True
      Top             =   165
      Width           =   630
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   135
      TabIndex        =   7
      Top             =   885
      Width           =   4935
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : dlgAbout.frm                                                **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : This is the VB Control Manager about box dialogue           **
'** Last modified on September 3, 2003                                        **
'*******************************************************************************

Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, X As Single, Y As Single)
  lblEMail.Font.Underline = False
End Sub

Private Sub lblEMail_Click()
  mdlAPI.ShellExecute hwnd:=Me.hwnd, lpOperation:=vbNullString, _
                      lpFile:="mailto:theo_yz@yahoo.com", _
                      lpParameters:=vbNullString, _
                      lpDirectory:=vbNullString, nShowCmd:=1
End Sub

Private Sub lblEMail_MouseMove(Button As Integer, _
                               Shift As Integer, X As Single, Y As Single)
  lblEMail.Font.Underline = True
End Sub
