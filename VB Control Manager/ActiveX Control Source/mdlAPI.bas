Attribute VB_Name = "mdlAPI"
Attribute VB_Description = "A module to declare Windows API procedures, functions, types and constants"
'*******************************************************************************
'** File Name   : mdlAPI.bas                                                  **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : A module to declare Windows API procedures, functions,      **
'**               types and constants                                         **
'** Last modified on November 13, 2003                                        **
'*******************************************************************************

Option Explicit

'------------------------------------------
' Refer to MSDN Library for the complete
' descripion of each API declaration below
'------------------------------------------

'--- API Functions Parameters Constant
Public Const GA_ROOT = &H2                                  'used by GetAncestor
Public Const LB_GETITEMHEIGHT = &H1A1                       'used by SendMessage
' Used by DrawEdge
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

'--- API Window Messages
Public Const WM_ACTIVATE = &H6                       'used to detect the current
                                                     '         form deactivation
Public Const WM_SHOWWINDOW = &H18                      'used to detect form show
'                                                      '(for the first time only)
Public Const WM_SIZE = &H5                           'used to detect form resize

'--- API Window Messages Parameter
Public Const WA_INACTIVE = &H0                         'parameter of WM_ACTIVATE

'--- API Types Declaration
Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'--- API Function/Sub Declaration

' Purpose    : Converts the client coordinates of a specified point to screen
'              coordinates
Public Declare Sub _
  ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)

' Purpose    : Confines the cursor to a rectangular area lpRect on the screen
Public Declare Sub ClipCursor Lib "user32" (lpRect As RECT)
Attribute ClipCursor.VB_Description = "Confines the cursor to a rectangular area lpRect on the screen"

' Purpose    : Frees the cursor to move anywhere on the screen
Public Declare Sub _
  ClipCursorClear Lib "user32" Alias "ClipCursor" _
    (Optional ByVal lpRect As Long = 0&)
Attribute ClipCursorClear.VB_Description = "Frees the cursor to move anywhere on the screen"

' Purpose    : Draws one or more edges of a rectangle
' Usage      : Draws the control's title bar
Public Declare Sub _
  DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, _
                         ByVal edge As Long, ByVal grfFlags As Long)

' Purpose    : Retrieves the handle to the ancestor of the specified window
Public Declare Function _
  GetAncestor Lib "user32" _
    (ByVal hwnd As Long, ByVal gaFlags As Long) As Long

' Purpose    : Retrieves the handle to the window (if any) that has captured the
'              mouse or stylus input
' Usage      : In ctlButton to provide hover effect along with SetCapture,
'              ReleaseCapture and WindowFromPoint
Public Declare Function GetCapture Lib "user32" () As Long

' Purpose    : Retrieves the cursor's position in screen coordinates
Public Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
Attribute GetCursorPos.VB_Description = "Retrieves the cursor's position in screen coordinates"

' Purpose    : Releases the mouse capture from a window in the current thread
'              and restores normal mouse input processing
' Usage      : In ctlButton to provide hover effect along with SetCapture,
'              GetCapture and WindowFormPoint
Public Declare Function ReleaseCapture Lib "user32" () As Long

' Purpose    : Converts the screen coordinates of a specified point on the
'              screen to client coordinates
Public Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

' Purpose    : Sends message wMsg to window hWnd
' Usage      : Gets the height of the item in list box control or other controls
'              that inherit it
Public Declare Function _
  SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, lParam As Long) As Long
Attribute SendMessage.VB_Description = "Sends message wMsg to window hWnd"

' Purpose    : Sets the mouse capture to the specified window belonging to the
'              current thread
' Usage      : In ctlButton to provide hover effect along with GetCapture,
'              ReleaseCapture and WindowFromPoint
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

' Purpose    : Sets the coordinates of the specified rectangle
' Usage      : Just to shorten the number of source code lines needed to set
'              the rectangle coordinate
Public Declare Sub _
  SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, _
                        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)

' Purpose    : Performs an operation on a specified file
' Usage      : Opens mail client for specified e-mail address
Public Declare Sub _
  ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
     ByVal lpFile As String, ByVal lpParameters As String, _
     ByVal lpDirectory As String, ByVal nShowCmd As Long)

' Purpose    : Retrieves the handle to the window that contains the specified
'              point
' Usage      : In ctlButton to provide hover effect along with SetCapture,
'              GetCapture and ReleaseCapture
Public Declare Function _
  WindowFromPoint Lib "user32" _
    (ByVal xPoint As Long, ByVal yPoint As Long) As Long
