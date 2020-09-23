VERSION 5.00
Object = "*\A..\ActiveX Control Source\VB Control Manager.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmDemoFeatures 
   Caption         =   "Features"
   ClientHeight    =   8955
   ClientLeft      =   165
   ClientTop       =   -1350
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrEvents 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   -60
      Top             =   6000
   End
   Begin VBControlManager.ControlManager ControlManager1 
      Height          =   8160
      Left            =   3600
      TabIndex        =   29
      Top             =   -15
      Width           =   7575
      _extentx        =   13361
      _extenty        =   14129
      liveupdate      =   0
      marginleft      =   3600
      titlebar_height =   0
      Begin VB.TextBox Text1 
         Height          =   5580
         Left            =   0
         TabIndex        =   39
         Text            =   "TextBox Sample"
         Top             =   180
         Width           =   1440
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6874
         Left            =   1500
         TabIndex        =   38
         Top             =   180
         Width           =   4898
         _ExtentX        =   8652
         _ExtentY        =   12118
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         FileName        =   "F:\My Creations\Software\Excellent\VB Control Manager\Unpublished Files\Features.rtf"
         TextRTF         =   $"frmDemoFeatures.frx":0000
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1005
         Left            =   1500
         TabIndex        =   37
         Top             =   7125
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   1773
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Column 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Column 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Column 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Column 4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Column 5"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5340
         Left            =   6465
         TabIndex        =   36
         Top             =   1710
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   9419
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   1
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1485
         Left            =   6465
         Picture         =   "frmDemoFeatures.frx":AD81
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1110
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2325
         Left            =   0
         Picture         =   "frmDemoFeatures.frx":D475
         Stretch         =   -1  'True
         Top             =   5805
         Width           =   1440
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Frame fraFeatures 
         Height          =   7695
         Index           =   6
         Left            =   495
         TabIndex        =   87
         Top             =   390
         Width           =   3030
         Begin VB.VScrollBar vsbSplProperties 
            Height          =   7545
            Left            =   2805
            TabIndex        =   133
            Top             =   105
            Width           =   210
         End
         Begin VB.Frame fraConSplProperties 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   7395
            Left            =   30
            TabIndex        =   132
            Top             =   135
            Width           =   2730
            Begin VB.Frame fraSplProperties 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   11595
               Left            =   0
               TabIndex        =   134
               Top             =   0
               Width           =   2700
               Begin VB.ComboBox cboSplOrientation 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "frmDemoFeatures.frx":FFD4
                  Left            =   1605
                  List            =   "frmDemoFeatures.frx":FFDE
                  Style           =   2  'Dropdown List
                  TabIndex        =   159
                  ToolTipText     =   "Returns the virtual splitter movement direction"
                  Top             =   11370
                  Width           =   915
               End
               Begin VB.TextBox txtSplYc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   158
                  ToolTipText     =   "Returns the y-coordinate of the virtual splitter center"
                  Top             =   14070
                  Width           =   915
               End
               Begin VB.TextBox txtSplXc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   157
                  ToolTipText     =   "Returns the x-coordinate of the virtual splitter center"
                  Top             =   13530
                  Width           =   915
               End
               Begin VB.TextBox txtSplWidth 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   156
                  ToolTipText     =   "Returns the width of the virtual splitter"
                  Top             =   13005
                  Width           =   915
               End
               Begin VB.TextBox txtSplTop 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   155
                  ToolTipText     =   "Returns the distance between the internal top edge of the virtual splitter and the top edge of the related Control Manager object"
                  Top             =   12465
                  Width           =   915
               End
               Begin VB.TextBox txtSplRight 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   154
                  ToolTipText     =   $"frmDemoFeatures.frx":FFFC
                  Top             =   11940
                  Width           =   915
               End
               Begin VB.TextBox txtSplMinYc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   153
                  ToolTipText     =   "Returns the minimum y-coordinate of the virtual splitter"
                  Top             =   10845
                  Width           =   915
               End
               Begin VB.TextBox txtSplMinXc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   152
                  ToolTipText     =   "Returns the minimum x-coordinate of the virtual splitter"
                  Top             =   10305
                  Width           =   915
               End
               Begin VB.TextBox txtSplMaxYc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   151
                  ToolTipText     =   "Returns the maximum y-coordinate of the virtual splitter"
                  Top             =   9780
                  Width           =   915
               End
               Begin VB.TextBox txtSplMaxXc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   150
                  ToolTipText     =   "Returns the maximum x-coordinate of the virtual splitter"
                  Top             =   9255
                  Width           =   915
               End
               Begin VB.ComboBox cboSplLiveUpdate 
                  Height          =   315
                  ItemData        =   "frmDemoFeatures.frx":10084
                  Left            =   1605
                  List            =   "frmDemoFeatures.frx":1008E
                  Style           =   2  'Dropdown List
                  TabIndex        =   149
                  ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
                  Top             =   8745
                  Width           =   915
               End
               Begin VB.TextBox txtSplLeft 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   148
                  ToolTipText     =   $"frmDemoFeatures.frx":1009F
                  Top             =   8280
                  Width           =   915
               End
               Begin VB.ListBox lstSplIdsSplTop 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   147
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's up-movement"
                  Top             =   7661
                  Width           =   900
               End
               Begin VB.ListBox lstSplIdsSplRight 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   146
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's right-movement"
                  Top             =   7042
                  Width           =   900
               End
               Begin VB.ListBox lstSplIdsSplLeft 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   145
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's left-movement"
                  Top             =   6423
                  Width           =   900
               End
               Begin VB.ListBox lstSplIdsSplBottom 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   144
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's down-movement"
                  Top             =   5804
                  Width           =   900
               End
               Begin VB.ListBox lstSplIdsCtlTop 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   143
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's up-movement"
                  Top             =   5185
                  Width           =   900
               End
               Begin VB.ListBox lstSplIdsCtlRight 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   142
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's right-movement"
                  Top             =   4566
                  Width           =   900
               End
               Begin VB.ListBox lstSplIdsCtlLeft 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   141
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's left-movement"
                  Top             =   3947
                  Width           =   900
               End
               Begin VB.ListBox lstSplIdsCtlBottom 
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   1605
                  TabIndex        =   140
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's down-movement"
                  Top             =   3328
                  Width           =   900
               End
               Begin VB.TextBox txtSplId 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   139
                  ToolTipText     =   "Returns a value that uniquely identifies the virtual splitter"
                  Top             =   2874
                  Width           =   915
               End
               Begin VB.TextBox txtSplHeight 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   138
                  ToolTipText     =   "Returns the height of the virtual splitter"
                  Top             =   2420
                  Width           =   915
               End
               Begin VB.ComboBox cboSplEnable 
                  Height          =   315
                  ItemData        =   "frmDemoFeatures.frx":10126
                  Left            =   1605
                  List            =   "frmDemoFeatures.frx":10130
                  Style           =   2  'Dropdown List
                  TabIndex        =   137
                  ToolTipText     =   "Returns/sets a value that determines whether the splitter is movable"
                  Top             =   1936
                  Width           =   915
               End
               Begin VB.ComboBox cboSplClipCursor 
                  Height          =   315
                  ItemData        =   "frmDemoFeatures.frx":10141
                  Left            =   1605
                  List            =   "frmDemoFeatures.frx":1014B
                  Style           =   2  'Dropdown List
                  TabIndex        =   136
                  ToolTipText     =   $"frmDemoFeatures.frx":1015C
                  Top             =   1452
                  Width           =   915
               End
               Begin VB.TextBox txtSplBottom 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1605
                  TabIndex        =   135
                  ToolTipText     =   $"frmDemoFeatures.frx":1020B
                  Top             =   998
                  Width           =   915
               End
               Begin VB.Label lblSplClick 
                  Caption         =   "(click the splitter)"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   315
                  TabIndex        =   190
                  Top             =   2970
                  Width           =   1245
               End
               Begin VB.Label Label76 
                  AutoSize        =   -1  'True
                  Caption         =   "Orientation:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   188
                  ToolTipText     =   "Returns the virtual splitter movement direction"
                  Top             =   11475
                  Width           =   810
               End
               Begin VB.Label Label75 
                  AutoSize        =   -1  'True
                  Caption         =   "Yc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   187
                  ToolTipText     =   "Returns the y-coordinate of the virtual splitter center"
                  Top             =   14160
                  Width           =   240
               End
               Begin VB.Label Label74 
                  AutoSize        =   -1  'True
                  Caption         =   "Xc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   186
                  ToolTipText     =   "Returns the x-coordinate of the virtual splitter center"
                  Top             =   13620
                  Width           =   240
               End
               Begin VB.Label Label73 
                  AutoSize        =   -1  'True
                  Caption         =   "Width:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   185
                  ToolTipText     =   "Returns the width of the virtual splitter"
                  Top             =   13080
                  Width           =   465
               End
               Begin VB.Label Label72 
                  AutoSize        =   -1  'True
                  Caption         =   "Top:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   184
                  ToolTipText     =   "Returns the distance between the internal top edge of the virtual splitter and the top edge of the related Control Manager object"
                  Top             =   12540
                  Width           =   330
               End
               Begin VB.Label Label71 
                  AutoSize        =   -1  'True
                  Caption         =   "Right:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   183
                  ToolTipText     =   $"frmDemoFeatures.frx":10293
                  Top             =   12015
                  Width           =   420
               End
               Begin VB.Label Label70 
                  AutoSize        =   -1  'True
                  Caption         =   "MinYc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   182
                  ToolTipText     =   "Returns the minimum y-coordinate of the virtual splitter"
                  Top             =   10935
                  Width           =   495
               End
               Begin VB.Label Label69 
                  AutoSize        =   -1  'True
                  Caption         =   "MinXc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   181
                  ToolTipText     =   "Returns the minimum x-coordinate of the virtual splitter"
                  Top             =   10395
                  Width           =   495
               End
               Begin VB.Label Label68 
                  AutoSize        =   -1  'True
                  Caption         =   "MaxYc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   180
                  ToolTipText     =   "Returns the maximum y-coordinate of the virtual splitter"
                  Top             =   9870
                  Width           =   540
               End
               Begin VB.Label Label67 
                  AutoSize        =   -1  'True
                  Caption         =   "MaxXc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   179
                  ToolTipText     =   "Returns the maximum x-coordinate of the virtual splitter"
                  Top             =   9330
                  Width           =   540
               End
               Begin VB.Label Label66 
                  AutoSize        =   -1  'True
                  Caption         =   "LiveUpdate:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   178
                  ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as the splitter is moved"
                  Top             =   8835
                  Width           =   870
               End
               Begin VB.Label Label65 
                  AutoSize        =   -1  'True
                  Caption         =   "Left:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   177
                  ToolTipText     =   $"frmDemoFeatures.frx":1031B
                  Top             =   8370
                  Width           =   315
               End
               Begin VB.Label Label64 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsCtlTop:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   176
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's up-movement"
                  Top             =   7650
                  Width           =   720
               End
               Begin VB.Label Label63 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsSplRight:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   175
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's right-movement"
                  Top             =   7035
                  Width           =   855
               End
               Begin VB.Label Label62 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsSplLeft:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   174
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's left-movement"
                  Top             =   6435
                  Width           =   750
               End
               Begin VB.Label Label61 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsSplBottom:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   173
                  ToolTipText     =   "Returns the Id collection of all virtual splitters which are effected by the virtual splitter's down-movement"
                  Top             =   5820
                  Width           =   975
               End
               Begin VB.Label Label60 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsCtlTop:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   172
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's up-movement"
                  Top             =   5220
                  Width           =   720
               End
               Begin VB.Label Label59 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsCtlRight:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   171
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's right-movement"
                  Top             =   4605
                  Width           =   810
               End
               Begin VB.Label Label58 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsCtlLeft:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   170
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's left-movement"
                  Top             =   4005
                  Width           =   705
               End
               Begin VB.Label Label57 
                  AutoSize        =   -1  'True
                  Caption         =   "IdsCtlBottom:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   169
                  ToolTipText     =   "Returns the Id collection of all virtual controls which are effected by the virtual splitter's down-movement"
                  Top             =   3390
                  Width           =   930
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
                  Caption         =   "Id:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   168
                  ToolTipText     =   "Returns a value that uniquely identifies the virtual splitter"
                  Top             =   2985
                  Width           =   180
               End
               Begin VB.Label Label54 
                  AutoSize        =   -1  'True
                  Caption         =   "Height:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   167
                  ToolTipText     =   "Returns the height of the virtual splitter"
                  Top             =   2505
                  Width           =   510
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  Caption         =   "Enable:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   166
                  ToolTipText     =   "Returns/sets a value that determines whether the splitter is movable"
                  Top             =   2040
                  Width           =   540
               End
               Begin VB.Label Label55 
                  AutoSize        =   -1  'True
                  Caption         =   "BackColor:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   165
                  ToolTipText     =   "Returns/sets the background color used to display the splitter"
                  Top             =   615
                  Width           =   780
               End
               Begin VB.Label lblSplBackColor 
                  BackColor       =   &H00404040&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   255
                  Left            =   1605
                  TabIndex        =   164
                  ToolTipText     =   "Returns/sets the background color used to display the splitter"
                  Top             =   574
                  Width           =   915
               End
               Begin VB.Label Label53 
                  AutoSize        =   -1  'True
                  Caption         =   "Active Color:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   163
                  ToolTipText     =   "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
                  Top             =   150
                  Width           =   900
               End
               Begin VB.Label lblSplActiveColor 
                  BackColor       =   &H00404040&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   255
                  Left            =   1605
                  TabIndex        =   162
                  ToolTipText     =   "Returns/sets the background color used to display the splitter when the user moves it in none live update mode"
                  Top             =   150
                  Width           =   915
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  Caption         =   "ClipCursor:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   161
                  ToolTipText     =   $"frmDemoFeatures.frx":103A2
                  Top             =   1560
                  Width           =   750
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  Caption         =   "Bottom:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   160
                  ToolTipText     =   $"frmDemoFeatures.frx":10451
                  Top             =   1095
                  Width           =   540
               End
            End
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7710
         Index           =   3
         Left            =   495
         TabIndex        =   88
         Top             =   390
         Width           =   3030
         Begin VB.VScrollBar vsbCtlProperties 
            Height          =   7530
            Left            =   2805
            TabIndex        =   131
            Top             =   105
            Width           =   210
         End
         Begin VB.Frame fraConCtlProperties 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   7455
            Left            =   30
            TabIndex        =   89
            Top             =   135
            Width           =   2760
            Begin VB.Frame fraCtlProperties 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   7425
               Left            =   0
               TabIndex        =   90
               Top             =   0
               Width           =   2745
               Begin VB.ComboBox cboCtlTitleBarVisible 
                  Height          =   315
                  ItemData        =   "frmDemoFeatures.frx":104D9
                  Left            =   1680
                  List            =   "frmDemoFeatures.frx":104E3
                  Style           =   2  'Dropdown List
                  TabIndex        =   110
                  ToolTipText     =   "Returns/sets a value that determines whether the virtual control title bar is visible"
                  Top             =   6600
                  Width           =   915
               End
               Begin VB.ComboBox cboCtlTitleBarCloseVisible 
                  Height          =   315
                  ItemData        =   "frmDemoFeatures.frx":104F4
                  Left            =   1680
                  List            =   "frmDemoFeatures.frx":104FE
                  Style           =   2  'Dropdown List
                  TabIndex        =   109
                  ToolTipText     =   "Returns/sets a value that determines whether a close button in the virtual control title bar is visible"
                  Top             =   5715
                  Width           =   915
               End
               Begin VB.TextBox txtCtlTitleBarHeight 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   108
                  ToolTipText     =   "Returns the height of the virtual control title bars"
                  Top             =   6165
                  Width           =   915
               End
               Begin VB.TextBox txtCtlBottom 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   107
                  ToolTipText     =   $"frmDemoFeatures.frx":1050F
                  Top             =   90
                  Width           =   915
               End
               Begin VB.ComboBox cboCtlClosed 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "frmDemoFeatures.frx":10596
                  Left            =   1680
                  List            =   "frmDemoFeatures.frx":105A0
                  Style           =   2  'Dropdown List
                  TabIndex        =   106
                  ToolTipText     =   "Returns a value that determines whether the virtual control is closed"
                  Top             =   525
                  Width           =   915
               End
               Begin VB.TextBox txtCtlHeight 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   105
                  ToolTipText     =   "Returns the height of the virtual control"
                  Top             =   975
                  Width           =   915
               End
               Begin VB.TextBox txtCtlId 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   104
                  ToolTipText     =   "Returns a value that uniquely identifies the virtual control"
                  Top             =   1410
                  Width           =   915
               End
               Begin VB.TextBox txtCtlIdSplBottom 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   103
                  ToolTipText     =   $"frmDemoFeatures.frx":105B1
                  Top             =   1845
                  Width           =   915
               End
               Begin VB.TextBox txtCtlIdSplLeft 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   102
                  ToolTipText     =   $"frmDemoFeatures.frx":10644
                  Top             =   2265
                  Width           =   915
               End
               Begin VB.TextBox txtCtlIdSplRight 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   101
                  ToolTipText     =   $"frmDemoFeatures.frx":106D5
                  Top             =   2700
                  Width           =   915
               End
               Begin VB.TextBox txtCtlIdSplTop 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   100
                  ToolTipText     =   $"frmDemoFeatures.frx":10767
                  Top             =   3135
                  Width           =   915
               End
               Begin VB.TextBox txtCtlLeft 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   99
                  ToolTipText     =   "Returns the distance between the internal left edge of the virtual control and the left edge of the related Control Manager object"
                  Top             =   3555
                  Width           =   915
               End
               Begin VB.TextBox txtCtlMinHeight 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   98
                  ToolTipText     =   "Returns the minimum height of the virtual control"
                  Top             =   3990
                  Width           =   915
               End
               Begin VB.TextBox txtCtlMinWidth 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   97
                  ToolTipText     =   "Returns the minimum width of the virtual control"
                  Top             =   4425
                  Width           =   915
               End
               Begin VB.TextBox txtCtlName 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   96
                  ToolTipText     =   "Returns the name of the real control that the virtual control represents"
                  Top             =   4845
                  Width           =   915
               End
               Begin VB.TextBox txtCtlRight 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   95
                  ToolTipText     =   $"frmDemoFeatures.frx":107F7
                  Top             =   5280
                  Width           =   915
               End
               Begin VB.TextBox txtCtlTop 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   94
                  ToolTipText     =   "Returns the distance between the internal top edge of the virtual control and the top edge of the related Control Manager object"
                  Top             =   7065
                  Width           =   915
               End
               Begin VB.TextBox txtCtlWidth 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   93
                  ToolTipText     =   "Returns the width of the virtual control"
                  Top             =   7485
                  Width           =   915
               End
               Begin VB.TextBox txtCtlXc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   92
                  ToolTipText     =   "Returns the x-coordinate of the virtual control center"
                  Top             =   7920
                  Width           =   915
               End
               Begin VB.TextBox txtCtlYc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   91
                  ToolTipText     =   "Returns the y-coordinate of the virtual control center"
                  Top             =   8355
                  Width           =   915
               End
               Begin VB.Label lblCtlClick 
                  Caption         =   "(click the title bar)"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   330
                  TabIndex        =   189
                  Top             =   1470
                  Width           =   1275
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "TitleBar_Visible:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   130
                  ToolTipText     =   "Returns/sets a value that determines whether the virtual control title bar is visible"
                  Top             =   6675
                  Width           =   1125
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  Caption         =   "TitleBar_Height:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   129
                  ToolTipText     =   "Returns the height of the virtual control title bars"
                  Top             =   6240
                  Width           =   1140
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  Caption         =   "TitleBar_CloseVisible:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   128
                  ToolTipText     =   "Returns/sets a value that determines whether a close button in the virtual control title bar is visible"
                  Top             =   5820
                  Width           =   1515
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Bottom:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   127
                  ToolTipText     =   $"frmDemoFeatures.frx":1087E
                  Top             =   165
                  Width           =   540
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Closed:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   126
                  ToolTipText     =   "Returns a value that determines whether the virtual control is closed"
                  Top             =   600
                  Width           =   525
               End
               Begin VB.Label Label38 
                  AutoSize        =   -1  'True
                  Caption         =   "Height:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   125
                  ToolTipText     =   "Returns the height of the virtual control"
                  Top             =   1035
                  Width           =   510
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  Caption         =   "Id:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   124
                  ToolTipText     =   "Returns a value that uniquely identifies the virtual control"
                  Top             =   1470
                  Width           =   180
               End
               Begin VB.Label Label40 
                  AutoSize        =   -1  'True
                  Caption         =   "IdSplBottom:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   123
                  ToolTipText     =   $"frmDemoFeatures.frx":10905
                  Top             =   1905
                  Width           =   900
               End
               Begin VB.Label Label41 
                  AutoSize        =   -1  'True
                  Caption         =   "IdSplLeft:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   122
                  ToolTipText     =   $"frmDemoFeatures.frx":10998
                  Top             =   2340
                  Width           =   675
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  Caption         =   "IdSplRight:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   121
                  ToolTipText     =   $"frmDemoFeatures.frx":10A29
                  Top             =   2775
                  Width           =   780
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "IdSplTop:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   120
                  ToolTipText     =   $"frmDemoFeatures.frx":10ABB
                  Top             =   3210
                  Width           =   690
               End
               Begin VB.Label Label44 
                  AutoSize        =   -1  'True
                  Caption         =   "Left:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   119
                  ToolTipText     =   "Returns the distance between the internal left edge of the virtual control and the left edge of the related Control Manager object"
                  Top             =   3630
                  Width           =   315
               End
               Begin VB.Label Label45 
                  AutoSize        =   -1  'True
                  Caption         =   "MinHeight:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   118
                  ToolTipText     =   "Returns the minimum height of the virtual control"
                  Top             =   4065
                  Width           =   765
               End
               Begin VB.Label Label46 
                  AutoSize        =   -1  'True
                  Caption         =   "MinWidth:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   117
                  ToolTipText     =   "Returns the minimum width of the virtual control"
                  Top             =   4500
                  Width           =   720
               End
               Begin VB.Label Label47 
                  AutoSize        =   -1  'True
                  Caption         =   "Name:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   116
                  ToolTipText     =   "Returns the name of the real control that the virtual control represents"
                  Top             =   4935
                  Width           =   465
               End
               Begin VB.Label Label48 
                  AutoSize        =   -1  'True
                  Caption         =   "Right:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   115
                  ToolTipText     =   $"frmDemoFeatures.frx":10B4B
                  Top             =   5385
                  Width           =   420
               End
               Begin VB.Label Label49 
                  AutoSize        =   -1  'True
                  Caption         =   "Top:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   114
                  ToolTipText     =   "Returns the distance between the internal top edge of the virtual control and the top edge of the related Control Manager object"
                  Top             =   7110
                  Width           =   330
               End
               Begin VB.Label Label50 
                  AutoSize        =   -1  'True
                  Caption         =   "Width:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   113
                  ToolTipText     =   "Returns the width of the virtual control"
                  Top             =   7545
                  Width           =   465
               End
               Begin VB.Label Label51 
                  AutoSize        =   -1  'True
                  Caption         =   "Xc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   112
                  ToolTipText     =   "Returns the x-coordinate of the virtual control center"
                  Top             =   7980
                  Width           =   240
               End
               Begin VB.Label Label52 
                  AutoSize        =   -1  'True
                  Caption         =   "Yc:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   111
                  ToolTipText     =   "Returns the y-coordinate of the virtual control center"
                  Top             =   8430
                  Width           =   240
               End
            End
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7695
         Index           =   0
         Left            =   495
         TabIndex        =   1
         Top             =   390
         Width           =   3030
         Begin VB.ComboBox cboCMTitleBarVisible 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10BD2
            Left            =   1725
            List            =   "frmDemoFeatures.frx":10BDC
            Style           =   2  'Dropdown List
            TabIndex        =   57
            ToolTipText     =   "Returns/sets a value that determines whether all control title bars are visible"
            Top             =   5625
            Width           =   915
         End
         Begin VB.ComboBox cboCMTitleBarCloseVisible 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10BED
            Left            =   1725
            List            =   "frmDemoFeatures.frx":10BF7
            Style           =   2  'Dropdown List
            TabIndex        =   56
            ToolTipText     =   "Returns/sets a value that determines whether a close button in all control title bars is visible"
            Top             =   4770
            Width           =   915
         End
         Begin VB.TextBox txtCMTitleBarHeight 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1725
            TabIndex        =   53
            ToolTipText     =   "Returns/sets the height of all control title bars"
            Top             =   5220
            Width           =   915
         End
         Begin VB.TextBox txtCMSize 
            Height          =   285
            Left            =   1725
            TabIndex        =   24
            ToolTipText     =   "Returns/sets the size of all splitters"
            Top             =   4365
            Width           =   915
         End
         Begin VB.TextBox txtCMMarginTop 
            Height          =   285
            Left            =   1725
            TabIndex        =   23
            ToolTipText     =   "Returns/sets the top margin of the ActiveX Control from its container"
            Top             =   3960
            Width           =   915
         End
         Begin VB.TextBox txtCMMarginRight 
            Height          =   285
            Left            =   1725
            TabIndex        =   22
            ToolTipText     =   "Returns/sets the right margin of the ActiveX Control from its container"
            Top             =   3555
            Width           =   915
         End
         Begin VB.TextBox txtCMMarginLeft 
            Height          =   285
            Left            =   1740
            TabIndex        =   21
            ToolTipText     =   "Returns/sets the left margin of the ActiveX Control from its container"
            Top             =   3150
            Width           =   915
         End
         Begin VB.TextBox txtCMMarginBottom 
            Height          =   285
            Left            =   1725
            TabIndex        =   20
            ToolTipText     =   "Returns/sets the bottom margin of the ActiveX Control from its container"
            Top             =   2745
            Width           =   915
         End
         Begin VB.ComboBox cboCMLiveUpdate 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10C08
            Left            =   1725
            List            =   "frmDemoFeatures.frx":10C12
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as a splitter is moved"
            Top             =   2310
            Width           =   915
         End
         Begin VB.ComboBox cboCMEnable 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10C23
            Left            =   1725
            List            =   "frmDemoFeatures.frx":10C2D
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Returns/sets a value that determines whether all splitters are movable"
            Top             =   1425
            Width           =   915
         End
         Begin VB.ComboBox cboCMFillContainer 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10C3E
            Left            =   1725
            List            =   "frmDemoFeatures.frx":10C48
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   $"frmDemoFeatures.frx":10C59
            Top             =   1875
            Width           =   915
         End
         Begin VB.ComboBox cboCMClipCursor 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10D23
            Left            =   1725
            List            =   "frmDemoFeatures.frx":10D2D
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   $"frmDemoFeatures.frx":10D3E
            Top             =   1005
            Width           =   915
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "TitleBar_Visible:"
            Height          =   195
            Left            =   150
            TabIndex        =   55
            ToolTipText     =   "Returns/sets a value that determines whether all control title bars are visible"
            Top             =   5715
            Width           =   1125
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "TitleBar_Height:"
            Height          =   195
            Left            =   150
            TabIndex        =   54
            ToolTipText     =   "Returns/sets the height of all control title bars"
            Top             =   5295
            Width           =   1140
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "TitleBar_CloseVisible:"
            Height          =   195
            Left            =   135
            TabIndex        =   52
            ToolTipText     =   "Returns/sets a value that determines whether a close button in all control title bars is visible"
            Top             =   4875
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Size:"
            Height          =   195
            Left            =   150
            TabIndex        =   19
            ToolTipText     =   "Returns/sets the size of all splitters"
            Top             =   4455
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "MarginTop:"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            ToolTipText     =   "Returns/sets the top margin of the ActiveX Control from its container"
            Top             =   4050
            Width           =   810
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "MarginRight:"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            ToolTipText     =   "Returns/sets the right margin of the ActiveX Control from its container"
            Top             =   3630
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "MarginLeft:"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            ToolTipText     =   "Returns/sets the left margin of the ActiveX Control from its container"
            Top             =   3210
            Width           =   795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "MarginBottom:"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            ToolTipText     =   "Returns/sets the bottom margin of the ActiveX Control from its container"
            Top             =   2790
            Width           =   1020
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Live Update:"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            ToolTipText     =   "Returns/sets a value that determines whether the controls should be resized as a splitter is moved"
            Top             =   2385
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Enable:"
            Height          =   195
            Left            =   150
            TabIndex        =   12
            ToolTipText     =   "Returns/sets a value that determines whether all splitters are movable"
            Top             =   1545
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fill Container:"
            Height          =   195
            Left            =   150
            TabIndex        =   10
            ToolTipText     =   $"frmDemoFeatures.frx":10DEB
            Top             =   1965
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Clip Cursor:"
            Height          =   195
            Left            =   150
            TabIndex        =   7
            ToolTipText     =   $"frmDemoFeatures.frx":10EB5
            Top             =   1125
            Width           =   795
         End
         Begin VB.Label lblCMBackColor 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1725
            TabIndex        =   6
            ToolTipText     =   "Returns/sets the background color used to display all splitters"
            Top             =   615
            Width           =   915
         End
         Begin VB.Label lblCMActiveColor 
            BackColor       =   &H00404040&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1740
            TabIndex        =   5
            ToolTipText     =   "Returns/sets the background color used to display a splitter when the user moves it in none live update mode"
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Active Color:"
            Height          =   195
            Left            =   150
            TabIndex        =   4
            ToolTipText     =   "Returns/sets the background color used to display a splitter when the user moves it in none live update mode"
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Back Color:"
            Height          =   195
            Left            =   150
            TabIndex        =   2
            ToolTipText     =   "Returns/sets the background color used to display all splitters"
            Top             =   720
            Width           =   825
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7680
         Index           =   1
         Left            =   495
         TabIndex        =   25
         Top             =   390
         Width           =   3030
         Begin VB.ComboBox cboOpenControlMaintainSize 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10F62
            Left            =   1275
            List            =   "frmDemoFeatures.frx":10F6C
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   4755
            Width           =   1485
         End
         Begin VB.ComboBox cboMoveControlMoveTo 
            Height          =   315
            ItemData        =   "frmDemoFeatures.frx":10F7D
            Left            =   1005
            List            =   "frmDemoFeatures.frx":10F9C
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1710
            Width           =   1710
         End
         Begin VB.CommandButton cmdMoveControl 
            Caption         =   "Call"
            Height          =   285
            Left            =   2070
            TabIndex        =   66
            Top             =   1005
            Width           =   660
         End
         Begin VB.CommandButton cmdOpenControl 
            Caption         =   "Call"
            Height          =   285
            Left            =   2070
            TabIndex        =   62
            Top             =   4080
            Width           =   660
         End
         Begin VB.CommandButton cmdCloseControl 
            Caption         =   "Call"
            Height          =   285
            Left            =   2070
            TabIndex        =   58
            Top             =   240
            Width           =   660
         End
         Begin VB.CommandButton cmdMoveSplitter 
            Caption         =   "Call"
            Height          =   285
            Left            =   2070
            TabIndex        =   44
            Top             =   2910
            Width           =   660
         End
         Begin VB.TextBox txtMoveSplitterMoveTo 
            Height          =   300
            Left            =   990
            TabIndex        =   42
            Top             =   3615
            Width           =   1725
         End
         Begin VB.Label Label33 
            Caption         =   "IdSplitterDesination:"
            Height          =   255
            Left            =   225
            TabIndex        =   72
            Tag             =   $"frmDemoFeatures.frx":1101C
            Top             =   2565
            Width           =   1410
         End
         Begin VB.Label lblIdSplitter 
            Caption         =   "(click the spliter)"
            Height          =   255
            Index           =   0
            Left            =   1695
            TabIndex        =   71
            Top             =   2565
            Width           =   1200
         End
         Begin VB.Label lblIdControl2 
            Caption         =   "(double-click the title bar)"
            Height          =   435
            Left            =   1695
            TabIndex        =   194
            Top             =   2175
            Width           =   1200
         End
         Begin VB.Label Label17 
            Caption         =   "IdControlDesination:"
            Height          =   255
            Left            =   210
            TabIndex        =   193
            ToolTipText     =   $"frmDemoFeatures.frx":110BF
            Top             =   2175
            Width           =   1410
         End
         Begin VB.Label Label29 
            Caption         =   "MaintainSize:"
            Height          =   255
            Left            =   195
            TabIndex        =   74
            ToolTipText     =   $"frmDemoFeatures.frx":1116D
            Top             =   4860
            Width           =   1050
         End
         Begin VB.Label Label32 
            Caption         =   "IdControlSource:"
            Height          =   255
            Left            =   195
            TabIndex        =   70
            ToolTipText     =   "Required. A value that uniquely identifies the source control the developer want to move"
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label Label31 
            Caption         =   "MoveTo:"
            Height          =   255
            Left            =   210
            TabIndex        =   69
            ToolTipText     =   "Required. A value indicating the area type where the source control will be moved to"
            Top             =   1800
            Width           =   810
         End
         Begin VB.Label lblIdControl 
            Caption         =   "(click the title bar)"
            Height          =   255
            Index           =   1
            Left            =   1455
            TabIndex        =   68
            Top             =   1410
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "MoveControl"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   195
            TabIndex        =   67
            ToolTipText     =   "Moves a control to certain area"
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label Label30 
            Caption         =   "IdControl:"
            Height          =   255
            Left            =   195
            TabIndex        =   65
            ToolTipText     =   "Required. A value that uniquely identifies the control the developer want to open"
            Top             =   4485
            Width           =   705
         End
         Begin VB.Label lblIdControl 
            Caption         =   "(click the title bar)"
            Height          =   255
            Index           =   2
            Left            =   975
            TabIndex        =   64
            Top             =   4485
            Width           =   1800
         End
         Begin VB.Label Label28 
            Caption         =   "OpenControl"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   195
            TabIndex        =   63
            ToolTipText     =   "Opens (shows) a control and docks it to the ActiveX Control"
            Top             =   4125
            Width           =   1590
         End
         Begin VB.Label Label27 
            Caption         =   "IdControl:"
            Height          =   255
            Left            =   195
            TabIndex        =   61
            ToolTipText     =   "A value that uniquely identifies the control the developer want to close"
            Top             =   645
            Width           =   705
         End
         Begin VB.Label lblIdControl 
            Caption         =   "(click the title bar)"
            Height          =   255
            Index           =   0
            Left            =   1005
            TabIndex        =   60
            Top             =   645
            Width           =   1800
         End
         Begin VB.Label Label25 
            Caption         =   "CloseControl"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   195
            TabIndex        =   59
            ToolTipText     =   "Closes (hides) a control"
            Top             =   285
            Width           =   1590
         End
         Begin VB.Label Label15 
            Caption         =   "MoveSplitter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   195
            TabIndex        =   43
            ToolTipText     =   "Moves a splitter to the specified x- or y- (depending on the splitter's Orientation property) coordinate"
            Top             =   2955
            Width           =   1590
         End
         Begin VB.Label lblIdSplitter 
            Caption         =   "(click the spliter)"
            Height          =   255
            Index           =   1
            Left            =   975
            TabIndex        =   41
            Top             =   3315
            Width           =   1800
         End
         Begin VB.Label Label4 
            Caption         =   "MoveTo:"
            Height          =   255
            Left            =   195
            TabIndex        =   40
            ToolTipText     =   $"frmDemoFeatures.frx":11211
            Top             =   3705
            Width           =   810
         End
         Begin VB.Label Label3 
            Caption         =   "IdSplitter:"
            Height          =   255
            Left            =   195
            TabIndex        =   26
            ToolTipText     =   "Required. A value that uniquely identifies the splitter the developer want to move"
            Top             =   3315
            Width           =   705
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7665
         Index           =   2
         Left            =   495
         TabIndex        =   27
         Top             =   390
         Width           =   3030
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "TitleBarMouseUp"
            Height          =   255
            Index           =   17
            Left            =   225
            TabIndex        =   86
            ToolTipText     =   "Occurs when the user releases a mouse button over a control title bar without previously moving the control"
            Top             =   6645
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "TitleBarMouseMove"
            Height          =   255
            Index           =   16
            Left            =   225
            TabIndex        =   85
            ToolTipText     =   "Occurs when the user moves a mouse over a control title bar without moving the control"
            Top             =   6285
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "TitleBarMouseDown"
            Height          =   225
            Index           =   15
            Left            =   225
            TabIndex        =   84
            ToolTipText     =   "Occurs when the user presses a mouse button over a control title bar"
            Top             =   5925
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "TitleBarDblClick"
            Height          =   255
            Index           =   14
            Left            =   225
            TabIndex        =   83
            ToolTipText     =   "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a control title bar"
            Top             =   5550
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "TitleBarClick"
            Height          =   255
            Index           =   13
            Left            =   225
            TabIndex        =   82
            ToolTipText     =   "Occurs when the user presses and then realeses a mouse button over a control title bar"
            Top             =   5175
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterMoveEnd"
            Height          =   255
            Index           =   12
            Left            =   225
            TabIndex        =   81
            ToolTipText     =   "Occurs when the user presses and then realeses a mouse button over a control title bar"
            Top             =   4800
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterMoveBegin"
            Height          =   255
            Index           =   11
            Left            =   225
            TabIndex        =   80
            ToolTipText     =   "Occurs when the user is about to move a splitter"
            Top             =   4425
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterMove"
            Height          =   255
            Index           =   10
            Left            =   225
            TabIndex        =   79
            ToolTipText     =   "Occurs when the user is moving a splitter"
            Top             =   4050
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterMouseUp"
            Height          =   255
            Index           =   9
            Left            =   225
            TabIndex        =   78
            ToolTipText     =   "Occurs when the user releases a mouse button over a splitter without previously moving the splitter"
            Top             =   3675
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterMouseMove"
            Height          =   255
            Index           =   8
            Left            =   225
            TabIndex        =   77
            ToolTipText     =   "Occurs when the user moves a mouse over a splitter without moving the splitter"
            Top             =   3300
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterMouseDown"
            Height          =   255
            Index           =   7
            Left            =   225
            TabIndex        =   76
            ToolTipText     =   "Occurs when the user presses a mouse button over a splitter"
            Top             =   2925
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterDblClick"
            Height          =   255
            Index           =   6
            Left            =   225
            TabIndex        =   35
            ToolTipText     =   "Occurs when the user presses and then realeses a mouse button and then presses and releases it again over a splitter"
            Top             =   2550
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "SplitterClick"
            Height          =   255
            Index           =   5
            Left            =   225
            TabIndex        =   34
            ToolTipText     =   "Occurs when the user presses and then realeses a mouse button over a splitter"
            Top             =   2175
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "ControlMoveEnd"
            Height          =   255
            Index           =   4
            Left            =   225
            TabIndex        =   33
            ToolTipText     =   "Occurs when the user is finished moving a control, i.e. when the rectangle that represents the moving control disappears"
            Top             =   1800
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "ControlMoveBegin"
            Height          =   255
            Index           =   3
            Left            =   225
            TabIndex        =   32
            ToolTipText     =   "Occurs when the user is about to move a control, i.e. the first time the rectangle that represents the moving control occurs"
            Top             =   1425
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "ControlMove"
            Height          =   255
            Index           =   2
            Left            =   225
            TabIndex        =   31
            ToolTipText     =   "Occurs when the user is moving a control"
            Top             =   1050
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "ControlBeforeClose"
            Height          =   255
            Index           =   1
            Left            =   225
            TabIndex        =   30
            ToolTipText     =   "Occurs after the user presses a close button of certain control and before the control is closed"
            Top             =   675
            Width           =   2580
         End
         Begin VB.Label lblEvents 
            Alignment       =   2  'Center
            Caption         =   "ControlAfterClose"
            Height          =   255
            Index           =   0
            Left            =   225
            TabIndex        =   28
            ToolTipText     =   "Occurs when a control has just been closed by the user"
            Top             =   300
            Width           =   2580
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7695
         Index           =   4
         Left            =   495
         TabIndex        =   46
         Top             =   390
         Width           =   3030
         Begin VB.Label lblNoCtlMethod 
            Alignment       =   2  'Center
            Caption         =   "No Method"
            Height          =   210
            Left            =   45
            TabIndex        =   51
            Top             =   3765
            Width           =   2925
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7695
         Index           =   8
         Left            =   495
         TabIndex        =   49
         Top             =   390
         Width           =   3030
         Begin VB.Label lblNoSplEvent 
            Alignment       =   2  'Center
            Caption         =   "No Event"
            Height          =   210
            Left            =   30
            TabIndex        =   192
            Top             =   3480
            Width           =   2925
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7695
         Index           =   7
         Left            =   495
         TabIndex        =   48
         Top             =   390
         Width           =   3030
         Begin VB.Label lblNoSplMethod 
            Alignment       =   2  'Center
            Caption         =   "No Method"
            Height          =   210
            Left            =   45
            TabIndex        =   191
            Top             =   3570
            Width           =   2925
         End
      End
      Begin VB.Frame fraFeatures 
         Height          =   7035
         Index           =   5
         Left            =   495
         TabIndex        =   47
         Top             =   390
         Width           =   3030
         Begin VB.Label lblNoCtlEvent 
            Alignment       =   2  'Center
            Caption         =   "No Event"
            Height          =   255
            Left            =   45
            TabIndex        =   50
            Top             =   3330
            Width           =   2895
         End
      End
      Begin MSComDlg.CommonDialog cdlColor 
         Left            =   -60
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.TabStrip tabFeatures 
         Height          =   330
         Left            =   495
         TabIndex        =   3
         Top             =   75
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   582
         TabWidthStyle   =   1
         MultiRow        =   -1  'True
         Style           =   2
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Properties"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Methods"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Events"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TabStrip tabMembers 
         Height          =   7605
         Left            =   75
         TabIndex        =   45
         Top             =   480
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   13414
         TabWidthStyle   =   1
         MultiRow        =   -1  'True
         HotTracking     =   -1  'True
         Placement       =   2
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Control Manager"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Control"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Splitter"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDemoFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conFraControlManagerProperties = 0
Private Const conFraControlManagerMethods = 1
Private Const conFraControlManagerEvents = 2
Private Const conFraControlProperties = 3
Private Const conFraControlMethods = 4
Private Const conFraControlEvents = 5
Private Const conFraSplitterProperties = 6
Private Const conFraSplitterMethods = 7
Private Const conFraSplitterEvents = 8

Private Const conLblEventControlAfterClose = 0
Private Const conLblEventControlBeforeClose = 1
Private Const conLblEventControlMove = 2
Private Const conLblEventControlMoveBegin = 3
Private Const conLblEventControlMoveEnd = 4
Private Const conLblEventSplitterClick = 5
Private Const conLblEventSplitterDblClick = 6
Private Const conLblEventSplitterMouseDown = 7
Private Const conLblEventSplitterMouseMove = 8
Private Const conLblEventSplitterMouseUp = 9
Private Const conLblEventSplitterMove = 10
Private Const conLblEventSplitterMoveBegin = 11
Private Const conLblEventSplitterMoveEnd = 12
Private Const conLblEventTitleBarClick = 13
Private Const conLblEventTitleBarDblClick = 14
Private Const conLblEventTitleBarMouseDown = 15
Private Const conLblEventTitleBarMouseMove = 16
Private Const conLblEventTitleBarMouseUp = 17

Private Const conTabProperties = 1
Private Const conTabMethods = 2
Private Const conTabEvents = 3

Private Const conTabControlManager = 1
Private Const conTabControl = 2
Private Const conTabSplitter = 3

Private lngSelectedFrame As Long

Private Sub cboCMClipCursor_Click()
  ControlManager1.ClipCursor = CBool(cboCMClipCursor)
  RefreshProperties
End Sub

Private Sub cboCMEnable_Click()
  ControlManager1.Enable = CBool(cboCMEnable)
  RefreshProperties
End Sub

Private Sub cboCMFillContainer_Click()
  ControlManager1.FillContainer = CBool(cboCMFillContainer)
End Sub

Private Sub cboCMLiveUpdate_Click()
  ControlManager1.LiveUpdate = CBool(cboCMLiveUpdate)
  RefreshProperties
End Sub

Private Sub cboCMTitleBarCloseVisible_Click()
  ControlManager1.TitleBar_CloseVisible = CBool(cboCMTitleBarCloseVisible)
  RefreshProperties
End Sub

Private Sub cboCMTitleBarVisible_Click()
  ControlManager1.TitleBar_Visible = CBool(cboCMTitleBarVisible)
  RefreshProperties
End Sub

Private Sub cboCtlTitleBarCloseVisible_Click()
  If txtCtlId <> "" Then _
    ControlManager1.Controls(CLng(txtCtlId)).TitleBar_CloseVisible = _
      CBool(cboCtlTitleBarCloseVisible)
End Sub

Private Sub cboCtlTitleBarVisible_Click()
  If txtCtlId <> "" Then
    ControlManager1.Controls(CLng(txtCtlId)).TitleBar_Visible = _
      CBool(cboCtlTitleBarVisible)
    RefreshProperties
  End If
End Sub

Private Sub cboSplClipCursor_Click()
  If txtSplId <> "" Then
    If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then _
      ControlManager1.Splitters(CLng(txtSplId)).ClipCursor = _
        CBool(cboSplClipCursor)
  End If
End Sub

Private Sub cboSplEnable_Click()
  If txtSplId <> "" Then
    If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then _
      ControlManager1.Splitters(CLng(txtSplId)).Enable = _
        CBool(cboSplEnable)
  End If
End Sub

Private Sub cboSplLiveUpdate_Click()
  If txtSplId <> "" Then
    If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then _
      ControlManager1.Splitters(CLng(txtSplId)).LiveUpdate = _
        CBool(cboSplLiveUpdate)
  End If
End Sub

Private Sub ChangeTab()
  Select Case tabMembers.SelectedItem.Index
    Case conTabControlManager
      Select Case tabFeatures.SelectedItem.Index
        Case conTabProperties
          lngSelectedFrame = conFraControlManagerProperties
        Case conTabMethods
          lngSelectedFrame = conFraControlManagerMethods
        Case conTabEvents
          lngSelectedFrame = conFraControlManagerEvents
      End Select
    Case conTabControl
      Select Case tabFeatures.SelectedItem.Index
        Case conTabProperties
          lngSelectedFrame = conFraControlProperties
        Case conTabMethods
          lngSelectedFrame = conFraControlMethods
        Case conTabEvents
          lngSelectedFrame = conFraControlEvents
      End Select
    Case conTabSplitter
      Select Case tabFeatures.SelectedItem.Index
        Case conTabProperties
          lngSelectedFrame = conFraSplitterProperties
        Case conTabMethods
          lngSelectedFrame = conFraSplitterMethods
        Case conTabEvents
          lngSelectedFrame = conFraSplitterEvents
      End Select
  End Select
  ShowSelectedFrame
  
  If RichTextBox1.Visible Then RichTextBox1.SetFocus
End Sub

Private Sub ClearEvent(ByVal lngId As Long)
  lblEvents(lngId).BorderStyle = vbBSNone
  lblEvents(lngId).Font.Bold = False
  tmrEvents(lngId).Enabled = False
End Sub

Private Sub cmdCloseControl_Click()
  Dim blnSuccess As Boolean
  
  On Error GoTo ErrorHandler
  
  ControlManager1.CloseControl IdControl:=CLng(lblIdControl(0)), _
                               Success:=blnSuccess
  If Not blnSuccess Then ShowErrMessage "Fail to close the control"
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
End Sub

Private Sub cmdMoveControl_Click()
  Dim blnSuccess As Boolean
  
  On Error GoTo ErrorHandler
  
  Select Case cboMoveControlMoveTo.ListIndex
    Case mdSplitter
      ControlManager1.MoveControl IdControlSource:=CLng(lblIdControl(1)), _
                                  MoveTo:=cboMoveControlMoveTo.ListIndex, _
                                  IdSplitterDestination:= _
                                    CLng(lblIdSplitter(0)), _
                                  Success:=blnSuccess
    Case mdControlTop, mdControlRight, mdControlBottom, mdControlLeft
      ControlManager1.MoveControl IdControlSource:=CLng(lblIdControl(1)), _
                                  MoveTo:=cboMoveControlMoveTo.ListIndex, _
                                  IdControlDestination:=CLng(lblIdControl2), _
                                  Success:=blnSuccess
    Case mdEdgeTop, mdEdgeRight, mdEdgeBottom, mdEdgeLeft
      ControlManager1.MoveControl IdControlSource:=CLng(lblIdControl(1)), _
                                  MoveTo:=cboMoveControlMoveTo.ListIndex, _
                                  Success:=blnSuccess
  End Select
  If Not blnSuccess Then ShowErrMessage "Fail to move the control"
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
End Sub

Private Sub cmdMoveSplitter_Click()
  On Error GoTo ErrorHandler
  
  ControlManager1.MoveSplitter IdSplitter:=CLng(lblIdSplitter(1)), _
                               MoveTo:=CLng(txtMoveSplitterMoveTo)
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
End Sub

Private Sub cmdOpenControl_Click()
  Dim blnSuccess As Boolean
  
  On Error GoTo ErrorHandler
  
  ControlManager1.OpenControl IdControl:=CLng(lblIdControl(2)), _
                              MaintainSize:=CBool(cboOpenControlMaintainSize), _
                              Success:=blnSuccess
  If Not blnSuccess Then ShowErrMessage "Fail to open the control"
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
End Sub

Private Sub ControlManager1_ControlAfterClose(ByVal IdControl As Long)
  HighlightEvent conLblEventControlAfterClose
  RefreshProperties
End Sub

Private Sub ControlManager1_ControlBeforeClose( _
              ByVal IdControl As Long, Cancel As Boolean _
            )
  HighlightEvent conLblEventControlBeforeClose
End Sub

Private Sub ControlManager1_ControlMove( _
              ByVal IdControl As Long, ByVal Shift As Integer, _
              ByVal Left As Long, ByVal Top As Long, _
              ByVal Width As Long, ByVal Height As Long _
            )
  HighlightEvent conLblEventControlMove
End Sub

Private Sub ControlManager1_ControlMoveBegin( _
              ByVal IdControl As Long, ByVal Shift As Integer _
            )
  HighlightEvent conLblEventControlMoveBegin
End Sub

Private Sub ControlManager1_ControlMoveEnd( _
              ByVal IdControl As Long, _
              ByVal Shift As Integer, ByVal Moved As Boolean _
            )
  HighlightEvent conLblEventControlMoveEnd
  RefreshProperties
End Sub

Private Sub ControlManager1_SplitterClick(ByVal IdSplitter As Long)
  Dim i As Long
  
  HighlightEvent conLblEventSplitterClick
  
  If fraFeatures(conFraControlManagerMethods).Visible Then
    For i = 0 To lblIdSplitter.UBound
      lblIdSplitter(i) = CStr(IdSplitter)
    Next
  End If
  txtSplId = CStr(IdSplitter)
End Sub

Private Sub ControlManager1_SplitterDblClick(ByVal IdSplitter As Long)
  HighlightEvent conLblEventSplitterDblClick
End Sub

Private Sub ControlManager1_SplitterMouseDown( _
              ByVal IdSplitter As Long, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventSplitterMouseDown
End Sub

Private Sub ControlManager1_SplitterMouseMove( _
              ByVal IdSplitter As Long, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventSplitterMouseMove
End Sub

Private Sub ControlManager1_SplitterMouseUp( _
              ByVal IdSplitter As Long, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventSplitterMouseUp
End Sub

Private Sub ControlManager1_SplitterMove( _
              ByVal IdSplitter As Long, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventSplitterMove
End Sub

Private Sub ControlManager1_SplitterMoveBegin( _
              ByVal IdSplitter As Long, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventSplitterMoveBegin
End Sub

Private Sub ControlManager1_SplitterMoveEnd( _
              ByVal IdSplitter As Long, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventSplitterMoveEnd
  RefreshProperties
End Sub

Private Sub ControlManager1_TitleBarClick(ByVal IdControl As Long)
  Dim i As Long
  
  HighlightEvent conLblEventTitleBarClick
  
  If fraFeatures(conFraControlManagerMethods).Visible Then
    For i = 0 To lblIdControl.UBound
      lblIdControl(i) = CStr(IdControl)
    Next
  End If
  txtCtlId = CStr(IdControl)
End Sub

Private Sub ControlManager1_TitleBarDblClick(ByVal IdControl As Long)
  HighlightEvent conLblEventTitleBarDblClick
  
  If fraFeatures(conFraControlManagerMethods).Visible Then
    lblIdControl2 = CStr(IdControl)
  End If
End Sub

Private Sub ControlManager1_TitleBarMouseDown( _
              ByVal IdControl As Long, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventTitleBarMouseDown
End Sub

Private Sub ControlManager1_TitleBarMouseMove( _
              ByVal IdControl As Long, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventTitleBarMouseMove
End Sub

Private Sub ControlManager1_TitleBarMouseUp( _
              ByVal IdControl As Long, ByVal Button As Integer, _
              ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single _
            )
  HighlightEvent conLblEventTitleBarMouseUp
End Sub

Private Sub Form_Load()
  Dim Index As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim nodi As Node
  Dim nodj As Node

  For i = lblEvents.LBound + 1 To lblEvents.UBound
    Load tmrEvents(i)
  Next
  
  For i = 1 To 10
    Set nodi = TreeView1.Nodes.Add(, , , "Node " & CStr(i))
    For j = 1 To 6
      Set nodj = TreeView1.Nodes.Add(nodi.Index, tvwChild, , _
                                     "Node " & CStr(i) & "." & CStr(j))
      For k = 1 To 3
        TreeView1.Nodes.Add nodj.Index, tvwChild, , _
                            "Node " & CStr(i) & "." & CStr(j) & "." & CStr(k)
      Next
    Next
  Next
    
  For i = 1 To 10
    ListView1.ListItems.Add , , "Item " & CStr(i) & ".1"
    For j = 1 To ListView1.ColumnHeaders.Count - 1
      ListView1.ListItems(i).SubItems(j) = "Item " & CStr(i) & "." & CStr(j + 1)
    Next
  Next
  
  Me.Show
  tabFeatures_Click
  InitFeatures
End Sub

Private Sub Form_Resize()
  Dim lngNewHeight As Long

  fraMain.Height = Me.ScaleHeight
  lngNewHeight = Me.ScaleHeight - (tabFeatures.Top + tabFeatures.Height)
  vsbCtlProperties.Visible = (txtCtlYc.Top + txtCtlYc.Height + _
                              (3 * Screen.TwipsPerPixelY) > lngNewHeight)
  vsbSplProperties.Visible = (txtSplYc.Top + txtSplYc.Height + _
                              (3 * Screen.TwipsPerPixelY) > lngNewHeight)
  If lngNewHeight - (11 * Screen.TwipsPerPixelY) > 0 Then
    fraFeatures(lngSelectedFrame).Height = lngNewHeight
    tabMembers.Height = lngNewHeight - (7 * Screen.TwipsPerPixelY)
    vsbCtlProperties.Height = lngNewHeight - (8 * Screen.TwipsPerPixelY)
    vsbSplProperties.Height = lngNewHeight - (8 * Screen.TwipsPerPixelY)
    fraConCtlProperties.Height = lngNewHeight - (11 * Screen.TwipsPerPixelY)
    fraConSplProperties.Height = lngNewHeight - (11 * Screen.TwipsPerPixelY)
    lblNoCtlMethod.Top = (lngNewHeight \ 2) - (lblNoCtlMethod.Height \ 4)
    lblNoCtlEvent.Top = lblNoCtlMethod.Top
    lblNoSplMethod.Top = lblNoCtlMethod.Top
    lblNoSplEvent.Top = lblNoCtlMethod.Top
  Else
    fraFeatures(lngSelectedFrame).Height = 0
    tabMembers.Height = 0
    vsbCtlProperties.Height = 0
    vsbSplProperties.Height = 0
    fraConCtlProperties.Height = 0
    fraConSplProperties.Height = 0
    lblNoCtlMethod.Top = 0
    lblNoCtlEvent.Top = 0
  End If
  RefreshScrollBar
End Sub

Private Sub HighlightEvent(ByVal lngId As Long)
  If fraFeatures(conFraControlManagerEvents).Visible Then
    lblEvents(lngId).Font.Bold = True
    lblEvents(lngId).BorderStyle = vbFixedSingle
    tmrEvents(lngId).Enabled = True
  End If
End Sub

Private Sub InitFeatures()
  fraConCtlProperties.Height = txtCtlYc.Top + _
                               txtCtlYc.Height + (7 * Screen.TwipsPerPixelY)
  fraCtlProperties.Height = fraConCtlProperties.Height
 
  fraConSplProperties.Height = txtSplYc.Top + _
                               txtSplYc.Height + (7 * Screen.TwipsPerPixelY)
  fraSplProperties.Height = fraConSplProperties.Height
  
  With ControlManager1
    lblCMActiveColor.BackColor = .ActiveColor
    lblCMBackColor.BackColor = .BackColor
    cboCMClipCursor = CStr(.ClipCursor)
    cboCMEnable = CStr(.Enable)
    cboCMFillContainer = CStr(.FillContainer)
    cboCMLiveUpdate = CStr(.LiveUpdate)
    cboCMTitleBarCloseVisible = CStr(.TitleBar_CloseVisible)
    cboCMTitleBarVisible = CStr(.TitleBar_Visible)
    txtCMMarginBottom = CStr(.MarginBottom)
    txtCMMarginLeft = CStr(.MarginLeft)
    txtCMMarginRight = CStr(.MarginRight)
    txtCMMarginTop = CStr(.MarginTop)
    txtCMTitleBarHeight = CStr(.TitleBar_Height)
    txtCMSize = CStr(.Size)
  End With
  
  cboMoveControlMoveTo = "mdSplitter"
  cboOpenControlMaintainSize = "True"
End Sub

Private Sub lblCMActiveColor_Click()
  cdlColor.Flags = cdlCCRGBInit
  cdlColor.Color = lblCMActiveColor.BackColor
  cdlColor.ShowColor
  lblCMActiveColor.BackColor = cdlColor.Color
  ControlManager1.ActiveColor = lblCMActiveColor.BackColor
End Sub

Private Sub lblCMBackColor_Click()
  cdlColor.Flags = cdlCCRGBInit
  cdlColor.Color = lblCMBackColor.BackColor
  cdlColor.ShowColor
  lblCMBackColor.BackColor = cdlColor.Color
  ControlManager1.BackColor = lblCMBackColor.BackColor
End Sub

Private Sub lblCtlClick_Click()
  MsgBox "Click a title bar to see its control properties", vbInformation
End Sub

Private Sub lblSplActiveColor_Click()
  If txtSplId <> "" Then
    If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then
      cdlColor.Flags = cdlCCRGBInit
      cdlColor.Color = lblSplActiveColor.BackColor
      cdlColor.ShowColor
      lblSplActiveColor.BackColor = cdlColor.Color
      ControlManager1.Splitters(CLng(txtSplId)).ActiveColor = _
        lblSplActiveColor.BackColor
    End If
  End If
End Sub

Private Sub lblSplBackColor_Click()
  If txtSplId <> "" Then
    If ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then
      cdlColor.Flags = cdlCCRGBInit
      cdlColor.Color = lblSplBackColor.BackColor
      cdlColor.ShowColor
      lblSplBackColor.BackColor = cdlColor.Color
      ControlManager1.Splitters(CLng(txtSplId)).BackColor = _
        lblSplBackColor.BackColor
    End If
  End If
End Sub

Private Sub lblSplClick_Click()
  MsgBox "Click a splitter to see its properties", vbInformation
End Sub

Private Sub RefreshProperties()
  Static blnRefreshing As Boolean
  Dim oid As clsid
  
  If Not blnRefreshing Then
    blnRefreshing = True
    With ControlManager1
      txtCMTitleBarHeight = CStr(.TitleBar_Height)
      cboCMTitleBarVisible = CStr(.TitleBar_Visible)
    End With
    
    If txtCtlId <> "" Then
      With ControlManager1.Controls(txtCtlId)
        txtCtlBottom = CStr(.bottom)
        cboCtlClosed = CStr(.Closed)
        txtCtlHeight = CStr(.Height)
        txtCtlIdSplBottom = CStr(.IdSplBottom)
        txtCtlIdSplLeft = CStr(.IdSplLeft)
        txtCtlIdSplRight = CStr(.IdSplRight)
        txtCtlIdSplTop = CStr(.IdSplTop)
        txtCtlLeft = CStr(.Left)
        txtCtlMinHeight = CStr(.MinHeight)
        txtCtlMinWidth = CStr(.MinWidth)
        txtCtlName = .Name
        txtCtlRight = CStr(.Right)
        cboCtlTitleBarCloseVisible = CStr(.TitleBar_CloseVisible)
        txtCtlTitleBarHeight = CStr(.TitleBar_Height)
        cboCtlTitleBarVisible = CStr(.TitleBar_Visible)
        txtCtlTop = CStr(.Top)
        txtCtlWidth = CStr(.Width)
        txtCtlXc = CStr(.Xc)
        txtCtlYc = CStr(.Yc)
      End With
    End If
    
    If txtSplId <> "" Then
      If Not ControlManager1.Splitters.IsExist(CLng(txtSplId)) Then
        lblSplActiveColor.BackColor = vbBlack
        lblSplBackColor.BackColor = vbBlack
        txtSplBottom = ""
        cboSplClipCursor.ListIndex = -1
        cboSplEnable.ListIndex = -1
        txtSplHeight = ""
        txtSplId = ""
        lstSplIdsCtlBottom.Clear
        lstSplIdsCtlLeft.Clear
        lstSplIdsCtlRight.Clear
        lstSplIdsCtlTop.Clear
        lstSplIdsSplBottom.Clear
        lstSplIdsSplLeft.Clear
        lstSplIdsSplRight.Clear
        lstSplIdsSplTop.Clear
        txtSplLeft = ""
        cboSplLiveUpdate.ListIndex = -1
        txtSplMaxXc = ""
        txtSplMaxYc = ""
        txtSplMinXc = ""
        txtSplMinYc = ""
        cboSplOrientation.ListIndex = -1
        txtSplRight = ""
        txtSplTop = ""
        txtSplWidth = ""
        txtSplXc = ""
        txtSplYc = ""
      Else
        With ControlManager1.Splitters(CLng(txtSplId))
          lblSplActiveColor.BackColor = .ActiveColor
          lblSplBackColor.BackColor = .BackColor
          txtSplBottom = CStr(.bottom)
          cboSplClipCursor = CStr(.ClipCursor)
          cboSplEnable = CStr(.Enable)
          txtSplHeight = CStr(.Height)
          
          lstSplIdsCtlBottom.Clear
          For Each oid In .IdsCtlBottom
            lstSplIdsCtlBottom.AddItem CStr(oid)
          Next
          
          lstSplIdsCtlLeft.Clear
          For Each oid In .IdsCtlLeft
            lstSplIdsCtlLeft.AddItem CStr(oid)
          Next
          
          lstSplIdsCtlRight.Clear
          For Each oid In .IdsCtlRight
            lstSplIdsCtlRight.AddItem CStr(oid)
          Next
          
          lstSplIdsCtlTop.Clear
          For Each oid In .IdsCtlTop
            lstSplIdsCtlTop.AddItem CStr(oid)
          Next
          
          lstSplIdsSplBottom.Clear
          For Each oid In .IdsSplBottom
            lstSplIdsSplBottom.AddItem CStr(oid)
          Next
          
          lstSplIdsSplLeft.Clear
          For Each oid In .IdsSplLeft
            lstSplIdsSplLeft.AddItem CStr(oid)
          Next
          
          lstSplIdsSplRight.Clear
          For Each oid In .IdsSplRight
            lstSplIdsSplRight.AddItem CStr(oid)
          Next
          
          lstSplIdsSplTop.Clear
          For Each oid In .IdsSplTop
            lstSplIdsSplTop.AddItem CStr(oid)
          Next
          
          txtSplLeft = CStr(.Left)
          cboSplLiveUpdate = CStr(.LiveUpdate)
          txtSplMaxXc = CStr(.MaxXc)
          txtSplMaxYc = CStr(.MaxYc)
          txtSplMinXc = CStr(.MinXc)
          txtSplMinYc = CStr(.MinYc)
          
          Select Case .Orientation
            Case orHorizontal
              cboSplOrientation = "orHorizontal"
            Case orVertical
              cboSplOrientation = "orVertical"
          End Select
          
          txtSplRight = CStr(.Right)
          txtSplTop = CStr(.Top)
          txtSplWidth = CStr(.Width)
          txtSplXc = CStr(.Xc)
          txtSplYc = CStr(.Yc)
        End With
      End If
    End If
    
    blnRefreshing = False
  End If
End Sub

Private Sub RefreshScrollBar()
  With vsbCtlProperties
    .Min = 0
    .Max = fraCtlProperties.Height - _
           fraFeatures(conFraControlProperties).Height + _
           (7 * Screen.TwipsPerPixelY)
    .SmallChange = Screen.TwipsPerPixelY * 10
    .LargeChange = Screen.TwipsPerPixelY * 100
  End With
  
  With vsbSplProperties
    .Min = 0
    .Max = fraSplProperties.Height - _
           fraFeatures(conFraSplitterProperties).Height + _
           (7 * Screen.TwipsPerPixelY)
    .SmallChange = Screen.TwipsPerPixelY * 10
    .LargeChange = Screen.TwipsPerPixelY * 100
  End With
End Sub

Private Sub ShowErrMessage(Optional strErrMessage As String = "")
  If strErrMessage = "" Then strErrMessage = Err.Description
  MsgBox strErrMessage, vbCritical + vbOKOnly
End Sub

Private Sub ShowSelectedFrame()
  Dim i As Integer
  
  For i = fraFeatures.LBound To fraFeatures.UBound
    fraFeatures(i).Visible = False
  Next
  Form_Resize
  fraFeatures(lngSelectedFrame).Visible = True
  fraFeatures(lngSelectedFrame).ZOrder
End Sub

Private Sub tabFeatures_Click()
  ChangeTab
End Sub

Private Sub tabMembers_Click()
  ChangeTab
End Sub

Private Sub tmrEvents_Timer(Index As Integer)
  ClearEvent Index
End Sub

Private Sub txtCMMarginBottom_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = ControlManager1.MarginBottom
  ControlManager1.MarginBottom = CLng(txtCMMarginBottom)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtCMMarginBottom = lngOldValue
  Cancel = True
End Sub

Private Sub txtCMMarginLeft_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = ControlManager1.MarginLeft
  ControlManager1.MarginLeft = CLng(txtCMMarginLeft)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtCMMarginLeft = lngOldValue
  Cancel = True
End Sub

Private Sub txtCMMarginRight_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = ControlManager1.MarginRight
  ControlManager1.MarginRight = CLng(txtCMMarginRight)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtCMMarginRight = lngOldValue
  Cancel = True
End Sub

Private Sub txtCMMarginTop_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = ControlManager1.MarginTop
  ControlManager1.MarginTop = CLng(txtCMMarginTop)
  Exit Sub

ErrorHandler:
  ShowErrMessage
  txtCMMarginTop = lngOldValue
  Cancel = True
End Sub

Private Sub txtCMSize_Validate(Cancel As Boolean)
  Dim lngOldValue As Long
  
  On Error GoTo ErrorHandler

  lngOldValue = ControlManager1.Size
  ControlManager1.Size = CLng(txtCMSize)
  txtCMSize = CStr(ControlManager1.Size)
  Exit Sub
  
ErrorHandler:
  ShowErrMessage
  txtCMSize = lngOldValue
  ControlManager1.Size = lngOldValue
  Cancel = True
End Sub

Private Sub txtCtlId_Change()
  RefreshProperties
End Sub

Private Sub txtSplId_Change()
  If txtSplId <> "" Then RefreshProperties
End Sub

Private Sub vsbCtlProperties_Change()
  fraCtlProperties.Top = -vsbCtlProperties.Value
End Sub

Private Sub vsbSplProperties_Change()
  fraSplProperties.Top = -vsbSplProperties.Value
End Sub
