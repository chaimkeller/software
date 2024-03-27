VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mapsearchfm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search for highest points"
   ClientHeight    =   8055
   ClientLeft      =   6495
   ClientTop       =   1815
   ClientWidth     =   5400
   Icon            =   "mapsearchfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   5400
   Begin VB.CheckBox chkProfiles 
      Caption         =   "profiles"
      Height          =   255
      Left            =   4560
      TabIndex        =   67
      ToolTipText     =   "Output profile files in x and y as a function of grid spacing"
      Top             =   7380
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdGoogleMap 
      Height          =   495
      Left            =   2040
      Picture         =   "mapsearchfm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Load coordinates of google map form"
      Top             =   720
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBarProg 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   63
      Top             =   7680
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCheckSkipped 
      Caption         =   "Check for skipped points"
      Height          =   275
      Left            =   1180
      TabIndex        =   61
      Top             =   7380
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar ProgressBarProfs 
      Height          =   375
      Left            =   120
      TabIndex        =   62
      Top             =   7320
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame frmProg 
      Height          =   7335
      Left            =   4680
      TabIndex        =   38
      Top             =   -40
      Visible         =   0   'False
      Width           =   615
      Begin VB.CommandButton cmdCancelSearch 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   40
         Top             =   6840
         Width           =   495
      End
      Begin MSComctlLib.ProgressBar progSearch 
         Height          =   6495
         Left            =   180
         TabIndex        =   39
         Top             =   300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   11456
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   180
      TabIndex        =   32
      Top             =   -60
      Width           =   4275
      Begin VB.CommandButton Command16 
         Height          =   315
         Left            =   3720
         Picture         =   "mapsearchfm.frx":1284
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "save new city(area) name"
         Top             =   180
         Width           =   435
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1020
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "city(area) name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   34
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      Picture         =   "mapsearchfm.frx":1386
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Load in the current 3D Explorer coordinates"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2600
      Picture         =   "mapsearchfm.frx":1690
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Load the Map & More center coordinates"
      Top             =   720
      Width           =   795
   End
   Begin VB.Frame Frame3 
      Caption         =   "search results"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2475
      Left            =   180
      TabIndex        =   11
      Top             =   4800
      Width           =   4275
      Begin MSFlexGridLib.MSFlexGrid sky2 
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   101
         Cols            =   5
         BackColor       =   12648447
         ForeColorSel    =   16777215
         GridColor       =   33023
         FocusRect       =   2
         HighLight       =   2
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "^Point #|^ latitude       |^longitude      |^height (m)|^distance(km)"
      End
      Begin VB.CommandButton cmdSaveAll 
         Height          =   375
         Left            =   480
         Picture         =   "mapsearchfm.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Save all the points"
         Top             =   2040
         Width           =   480
      End
      Begin VB.CommandButton cmdReLoad 
         Height          =   375
         Left            =   130
         Picture         =   "mapsearchfm.frx":1F14
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Reload saved points and begin analysis"
         Top             =   2040
         Width           =   360
      End
      Begin VB.CommandButton cmdPlotSearchPnts 
         Height          =   375
         Left            =   2520
         Picture         =   "mapsearchfm.frx":2446
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Plot all the search results"
         Top             =   2040
         Width           =   450
      End
      Begin VB.CommandButton Command14 
         Height          =   375
         Left            =   1320
         Picture         =   "mapsearchfm.frx":2750
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Load places in *.bat city file"
         Top             =   2040
         Width           =   390
      End
      Begin VB.CommandButton Command12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         Picture         =   "mapsearchfm.frx":2C82
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Set Map&More's center coordinates to coordinates of highlighted point"
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   2160
         Picture         =   "mapsearchfm.frx":30C4
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Move active map interface (3D Explorer or Google Mazps) back to Maps & More center coordinates"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   3750
         Picture         =   "mapsearchfm.frx":3216
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "View sunset horizon for selected point"
         Top             =   2040
         Width           =   360
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   3370
         Picture         =   "mapsearchfm.frx":3658
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "view sunrise horizon of selected point"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdClear 
         Height          =   375
         Left            =   960
         Picture         =   "mapsearchfm.frx":3A9A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Discard search results and save buffers"
         Top             =   2040
         Width           =   435
      End
      Begin VB.CommandButton cmdSave 
         Height          =   375
         Left            =   2980
         Picture         =   "mapsearchfm.frx":3B9C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Save search results"
         Top             =   2040
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "search parameters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1800
      Left            =   180
      TabIndex        =   5
      Top             =   3000
      Width           =   4275
      Begin VB.CheckBox chkValidity 
         Height          =   195
         Left            =   3960
         TabIndex        =   66
         ToolTipText     =   "Check for a valid true and unique peak"
         Top             =   1560
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame frmExtras 
         Caption         =   "Advanced"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   120
         TabIndex        =   56
         Top             =   540
         Width           =   4065
         Begin VB.OptionButton optClear 
            Caption         =   "Search all dir."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   3320
            TabIndex        =   60
            Top             =   200
            Width           =   700
         End
         Begin VB.OptionButton optWest 
            Caption         =   "Search westward only"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   59
            ToolTipText     =   "search only west of center point"
            Top             =   300
            Width           =   1615
         End
         Begin VB.OptionButton optEast 
            Caption         =   "Search eastward only"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   58
            ToolTipText     =   "Only include points east of center point"
            Top             =   120
            Width           =   1615
         End
         Begin VB.CheckBox chkLineSight 
            Caption         =   "Clear Line of Siight"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   120
            TabIndex        =   57
            ToolTipText     =   "Require clear line of sight from center point to search point"
            Top             =   140
            Width           =   2000
         End
      End
      Begin VB.TextBox txtIgnoreLarge 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3840
         TabIndex        =   55
         Top             =   140
         Width           =   375
      End
      Begin VB.CheckBox chkIgnoreLarge 
         Caption         =   "Ignore hgts>="
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   54
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtHgtLimit 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3840
         TabIndex        =   53
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox chkIgnoreZeros 
         Caption         =   "Ignore hgt<="
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   52
         Top             =   380
         Width           =   975
      End
      Begin VB.Frame frmSort 
         Caption         =   "sort by"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   48
         Top             =   1120
         Width           =   825
         Begin VB.OptionButton optSortDist 
            Caption         =   "dist."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optSortHgt 
            Caption         =   "height"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   49
            Top             =   180
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.TextBox txtMosaic 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   43
         Top             =   1450
         Width           =   435
      End
      Begin VB.OptionButton optMosaic 
         Caption         =   "Mosaic Search"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   42
         ToolTipText     =   "Search for max heights among mosaics of a certain size"
         Top             =   1520
         Width           =   1215
      End
      Begin VB.OptionButton optSimple 
         Caption         =   "Simple Search"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   41
         ToolTipText     =   "Search for maximum heights"
         Top             =   1200
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtStep 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3480
         TabIndex        =   36
         Top             =   1150
         Width           =   435
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   20
         BuddyControl    =   "Text4"
         BuddyDispid     =   196648
         OrigLeft        =   3420
         OrigTop         =   300
         OrigRight       =   3660
         OrigBottom      =   675
         Increment       =   10
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   15
         BuddyControl    =   "Text3"
         BuddyDispid     =   196650
         OrigLeft        =   2100
         OrigTop         =   240
         OrigRight       =   2340
         OrigBottom      =   615
         Max             =   200
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Text            =   "20"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1040
         Picture         =   "mapsearchfm.frx":3C9E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Begin search"
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   6
         Text            =   "15"
         ToolTipText     =   "Search radius (km)"
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "km"
         Height          =   255
         Left            =   3960
         TabIndex        =   46
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblMosaic 
         Caption         =   "stepsize"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2880
         TabIndex        =   44
         Top             =   1520
         Width           =   555
      End
      Begin VB.Label lblStep 
         Caption         =   "stepsize"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   37
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Num. of  points"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   220
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Search rad. (km)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   200
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "coordinates of center point"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   180
      TabIndex        =   0
      Top             =   1200
      Width           =   4275
      Begin VB.CommandButton cmdMoveGoogleMap 
         Height          =   375
         Left            =   1250
         Picture         =   "mapsearchfm.frx":40E0
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Position Google Maps to center coordinates"
         Top             =   900
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   1380
         Width           =   2655
      End
      Begin VB.CommandButton Command15 
         Height          =   315
         Left            =   3720
         Picture         =   "mapsearchfm.frx":4F22
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Add the center coordinates and name to city file"
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton Command13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Picture         =   "mapsearchfm.frx":5024
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Set Map&More's center coordinates to search center coordinates"
         Top             =   900
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1620
         Picture         =   "mapsearchfm.frx":5466
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Position 3D Explorer to the center point"
         Top             =   900
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Height          =   375
         Left            =   2760
         Picture         =   "mapsearchfm.frx":58A8
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "View sunset horizon for center point"
         Top             =   900
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Height          =   375
         Left            =   2220
         Picture         =   "mapsearchfm.frx":5CEA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "view sunrise horizon for center point"
         Top             =   900
         Width           =   555
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2580
         TabIndex        =   2
         Text            =   "0"
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2580
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "sub-city name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Longitude of center point"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   540
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Latitude of center point"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Load in center coordinates from:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   21
      Top             =   480
      Width           =   3195
   End
End
Attribute VB_Name = "mapsearchfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim searchhgts() As Double
Dim subcitnams$(200), nncity%
Dim CancelSearch As Boolean
Dim SearchType%
Dim HeightSort As Boolean
Dim PlotSearchPoints As Boolean

Private Sub chkIgnoreLarge_Click()
   If chkIgnoreLarge.value = vbUnchecked Then
      txtIgnoreLarge.Enabled = False
   Else
      txtIgnoreLarge.Enabled = True
      If txtIgnoreLarge.Text = sEmpty Then txtIgnoreLarge.Text = "0"
      End If
End Sub

Private Sub chkIgnoreZeros_Click()
   If chkIgnoreZeros.value = vbUnchecked Then
      txtHgtLimit.Enabled = False
   Else
      txtHgtLimit.Enabled = True
      If txtHgtLimit.Text = sEmpty Then txtHgtLimit.Text = "0"
      End If
End Sub

Private Sub chkValidity_Click()
   If chkValidity.value = vbChecked Then
      chkProfiles.Visible = True
   Else
      chkProfiles.Visible = False
      End If
End Sub

Private Sub cmdCancelSearch_Click()
   CancelSearch = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCheckSkipped_Click
' Author    : chaim
' Date      : 7/17/2020
' Purpose   : automatically checks for skipped points
'---------------------------------------------------------------------------------------
'
Private Sub cmdCheckSkipped_Click()

  Dim batnum%, maplistnum%, doclin$, doclin2$, nummissing%
  Dim SplitArray() As String
  Dim SplitBat() As String
  Dim LatCompare As Double, LonCompare As Double
  Dim LatBat As Double, LonBat As Double
  Dim found%, batnam$, tmpfile$, tmpnum%
  
  nummissing% = 0
  
    'find name of active bat file
   On Error GoTo cmdCheckSkipped_Click_Error

    Select Case MsgBox("This will check for skipped points for the following city:" _
                       & vbCrLf & "" _
                       & vbCrLf & mapsearchfm.Combo2.Text _
                       & vbCrLf & "" _
                       & vbCrLf & "Answer " _
                       & vbCrLf & "" _
                       & vbCrLf & "       Yes --to check for skipped sunrise profiles" _
                       & vbCrLf & "       No - to check for skipped sunset profiles" _
                       & vbCrLf & "" _
                       & vbCrLf & "       Cancel - cancel the chek" _
                       , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "Checking for skipped points")
    
        Case vbYes
            sunmode = 1
        Case vbNo
            sunmode = 0
        Case vbCancel
            Exit Sub
    
    End Select
    
    'back up the mappoints file
    FileCopy drivjk_c$ & "mappoints.sav", drivjk_c$ & "mappoints.bak"
    nummissing% = 0
    If world Then
       AddPath$ = "eros\"
    Else
       AddPath$ = sEmpty
       End If
       
    'find name of active bat file
    nummissing% = 0
    If sunmode >= 1 Then 'sunrises
       batnam$ = Dir(drivcities$ & AddPath & Combo2.Text & "\netz\*.bat")
       batnam$ = drivcities$ & AddPath & Combo2.Text & "\netz\" & batnam$
       tmpfile$ = drivjk_c$ & "mappoints_tmp_netz.sav"
    ElseIf sunmode <= 0 Then 'sunsets
       batnam$ = Dir(drivcities$ & AddPath & Combo2.Text & "\skiy\*.bat")
       batnam$ = drivcities$ & AddPath & Combo2.Text & "\netz\" & batnam$
       tmpfile$ = drivjk_c$ & "mappoints_tmp_skiy.sav"
       End If

    maplistnum% = FreeFile
    Open drivjk_c$ + "mappoints.sav" For Input As #maplistnum%
    tmpnum% = FreeFile
    Open tmpfile$ For Output As #tmpnum%
    
    Do Until EOF(maplistnum%)
       Line Input #maplistnum%, doclin$
       SplitArray = Split(doclin$, ",")

       'this is a vantage point entry, so compare
      LatCompare = Val(SplitArray(0))
      LonCompare = Val(SplitArray(1))
      batnum% = FreeFile
      Open batnam$ For Input As #batnum%
      found% = 0
      Do Until EOF(batnum%)
         Line Input #batnum%, docliln2$
         SplitBat = Split(docliln2$, ",")
         If UBound(SplitBat) > 0 Then
              If InStr(SplitBat(0), "netz") Or InStr(SplitBat(0), "skiy") Then
                 'this is a vantage point entry
                 LatBat = Val(SplitBat(1))
                 LonBat = -Val(SplitBat(2))
                 If Abs(LatCompare - LatBat) < 0.0001 And Abs(LonCompare - LonBat) < 0.0001 Then
                    'accounted for, skip to next entry
                    found% = 1
                    Exit Do
                    End If
                 End If
             End If
      Loop
      
      If found% = 0 Then 'add to missing list
         Print #tmpnum%, doclin$
         nummissing% = nummissing% + 1
         End If
         
     Close #batnum%
    Loop
    Close #tmpnum%

    If nummissing% = 0 Then
       Call MsgBox("No missing search points found!", vbInformation Or vbDefaultButton1, "Checking for skipped points")
       Exit Sub
    Else
        Select Case MsgBox("Found  " & Str(nummissing%) & " skipped search points!" _
                           & vbCrLf & "" _
                           & vbCrLf & "Proceed to calculate their profiles(s)?" _
                           , vbOKCancel Or vbQuestion Or vbDefaultButton1, "Checking for skipped points")
        
           Case vbOK
           
           Case vbCancel
             Exit Sub
        
        End Select
        End If
        
'   cmdCheckSkipped.Enabled = False
     'reload the saved points, and begin automatic analysis
   If Dir(tmpfile$) <> sEmpty Then
      'determine number of rows
      savfil% = FreeFile
      Open tmpfile$ For Input As #savfil%
      mRows& = 0
      Do Until EOF(savfil%)
         Line Input #savfil%, doclin$
         mRows& = mRows& + 1
      Loop
      Close #savfil%
      sky2.Rows = mRows& + 1
      
      sky2.Clear
      savfil% = FreeFile
      Open tmpfile$ For Input As #savfil%
      i& = 1
      Do Until EOF(savfil%)
         If i& > sky2.Rows - 1 Then Exit Do
         Input #savfil%, savlat, savlon, savhgt, savdis
         sky2.TextArray(skyp2(i&, 0)) = i&
         sky2.TextArray(skyp2(i&, 1)) = savlat
         sky2.TextArray(skyp2(i&, 2)) = savlon
         sky2.TextArray(skyp2(i&, 3)) = savhgt
         sky2.TextArray(skyp2(i&, 4)) = savdis
         i& = i& + 1
      Loop
      Close #savfil%
      
      If world = False Then
         sky2.FormatString = "^Point # |^    ITMx          |^    ITMy             |^ height(m)  |^distance(km)"
      Else
         sky2.FormatString = "^Point #|^ latitude       |^longitude      |^height (m)|^distance(km)       "
         End If
         
      End If
       
    mapsearchfm.cmdCheckSkipped.Visible = False
    mapsearchfm.ProgressBarProfs.Visible = True
    mapsearchfm.ProgressBarProfs.Min = 0
    mapsearchfm.ProgressBarProfs.Max = i& - 1
    mapsearchfm.ProgressBarProfs.value = 0
    mapsearchfm.StatusBarProg.Panels(1).Text = i& - 1
    mapsearchfm.StatusBarProg.Panels(2).Text = "0"
       
    'now begin the automatic run
   AutoProf = True 'run profile analysis automatically
   AutoVer = True 'automatically increment the version number
   
   starting& = 0
   If Dir(drivjk_c$ & "mapstatus.sav") <> sEmpty Then
      Close
      statfil% = FreeFile
      Open drivjk_c$ & "mapstatus.sav" For Output As #statfil%
      Do Until EOF(statfil%)
         Print #statfil%, statnum&
      Loop
      Close #statfil%
      End If
      
uoa100:
      
   If world = False Then GoTo W100
   
   'create eros.tm6 file
   dtmfile% = FreeFile
   Open ramdrive & ":\eros.tm6" For Output As #dtmfile%
   Select Case DTMflag
      Case 0, -1 'GTOPO30, SRTM30
         outdrive$ = worlddtm
      Case 1, 2 'SRTM
         outdrive$ = srtmdtm
      Case 3
         outdrive$ = alosdtm
   End Select
   Print #dtmfile%, outdrive$; ","; DTMflag
   Close #dtmfile%
   
   'create eros.tm7 flag -- to tell newreadDTM not to ask about missing tiles
   dtmfile% = FreeFile
   Open ramdrive & ":\eros.tm7" For Output As #dtmfile%
   Print #dtmfile%, 1
   Close #dtmfile%
   
   AutoNum& = starting&
50 savfil% = FreeFile
   If sunmode = 1 Then
   ElseIf sunmode = 0 Then
      End If
   Open tmpfile$ For Input As #savfil%
   looping% = 0
   i& = 0
   Do Until EOF(savfil%)
      Input #savfil%, savlat, savlon, savhgt, savdis
      i& = i& + 1
      mapsearchfm.ProgressBarProfs.value = i&
      mapsearchfm.StatusBarProg.Panels(2).Text = Str(i&)
      statfil% = FreeFile 'record current status
      Open drivjk_c$ & "mapstatus.sav" For Output As #statfil%
      Write #statfil%, i&
      Close #statfil%
      If i& > AutoNum& Then
         'inputed all the desired coordinates, so exit loop
         looping% = 1

         Exit Do
         End If
   Loop
   Close #savfil%
   
   If looping% = 0 Then
      'reached EOF--i.e., finished the profiles
      mapsearchfm.ProgressBarProfs.Visible = False
      mapsearchfm.cmdCheckSkipped.Visible = True
      cmdCheckSkipped.Enabled = True
      
      AutoProf = False
      AutoVer = False
      Delay% = 0
      Kill ramdrive & ":\eros.tm7"
      Exit Sub
      End If
   
   If savlon = 0 And savlat = 0 Then GoTo 50
   lon = savlon
   lat = savlat
   worldmove = True
   Call goto_click
   jumpworld = False
   skymove = False
   worldmove = False
   Skycoord% = 0
   Call BringWindowToTop(mapsearchfm.hwnd)
   searchhgt = savhgt
   viewsearch = True
   Call sunrisesunset(sunmode%)
   viewsearch = False
   AutoNum& = AutoNum& + 1
   AutoVer = False 'don't increment the version number again
   GoTo 50
    

   On Error GoTo 0
   Exit Sub
   
W100: 'run analysis of Eretz Yisroel places
   'ask for a default name
   Unload mapsearchfm
   Set mapsearchfm = Nothing
   Call EYsunrisesunset(sunmode%)
   Exit Sub

cmdCheckSkipped_Click_Error:

    Close
    cmdCheckSkipped.Enabled = True
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCheckSkipped_Click of Form mapsearchfm"
    
End Sub

Private Sub cmdGoogleMap_Click()
  If GoogleMapVis Then
  Else
     frmMap.Visible = True
     End If
  
  If frmMap.txtLat.Text <> sEmpty Or frmMap.txtLong.Text <> sEmpty Then
     Text1.Text = frmMap.txtLat.Text
     Text2.Text = frmMap.txtLong.Text
  Else
     MsgBox "Coordinates in the Google Map form are empty" & vbCrLf & "Nowhere to move to!", vbOKOnly, "Google Map error"
     
     End If
End Sub

Private Sub cmdMoveGoogleMap_Click()
   If GoogleMapVis = True Then
      frmMap.txtLat.Text = Text1.Text
      frmMap.txtLong.Text = Text2.Text
      frmMap.Command2.value = True
   Else
      frmMap.Visible = True
      Call BringWindowToTop(frmMap.hwnd)
      frmMap.txtLat.Text = Text1.Text
      frmMap.txtLong.Text = Text2.Text
      frmMap.Command2.value = True
      End If
End Sub

Private Sub cmdPlotSearchPnts_Click()
   'plot all the search results on the map

   On Error GoTo cmdPlotSearchPnts_Click_Error
   
   If PlotSearchPoints Then
      'button pressed twice, so erase plot points
      If SearchVis Then
         SearchVis = False
         blitpictures
         End If
      PlotSearchPoints = False
      Exit Sub
      End If

    For j& = 1 To sky2.Rows - 1
      nplachos& = j&
      If nplachos& = 0 Then Exit Sub
      If world = False Then
         If InStr(sky2.FormatString, "ITMx") <> 0 Then
            skymove = True
            Skycoord% = 2
            skyx = sky2.TextArray(skyp2(nplachos&, 1))
            skyy = sky2.TextArray(skyp2(nplachos&, 2))
            If skyx = "0" And skyy = "0" Then GoTo p500
            'convert kmx,kmy to screen coordinates
            Call ScreenToGeo(x, y, skyx, skyy, 2, ier%)
         ElseIf InStr(sky2.FormatString, "latitude") <> 0 Then
            lati = sky2.TextArray(skyp2(nplachos&, 1))
            lons = sky2.TextArray(skyp2(nplachos&, 2))
            'convert lon,lat to screen coordinates
            Call ScreenToGeo(x, y, lons, lati, 2, ier%)
            End If
        End If
      
      If world = True Then
         lons = sky2.TextArray(skyp2(nplachos&, 2))
         lati = sky2.TextArray(skyp2(nplachos&, 1))
         'convert lon,lat to screen coordinates
         Call ScreenToGeo(x, y, lons, lati, 2, ier%)
         End If
    
      'plot the points
      mapPictureform.mapPicture.DrawWidth = 2 '2 * mag
      mapPictureform.mapPicture.Circle (x, y), 20, 255 '20 * mag, 255
      mapPictureform.mapPicture.DrawWidth = 1 '1 * mag
    
p500:
    Next j&
    
    PlotSearchPoints = True
    
    SearchVis = True

Call BringWindowToTop(mapsearchfm.hwnd)

   On Error GoTo 0
   Exit Sub

cmdPlotSearchPnts_Click_Error:
    Close
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPlotSearchPnts_Click of Form mapsearchfm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdReLoad_Click
' DateTime  : 9/2/2003 20:09
' Author    : Chaim Keller
' Purpose   : reloads saved points for automatic profile analysis
'---------------------------------------------------------------------------------------
'
Private Sub cmdReLoad_Click()
   
   On Error GoTo cmdReLoad_Click_Error
   
   If Not SavedAll Then
      Select Case MsgBox("You haven't yet saved the buffer points." _
                         & vbCrLf & "(If you don't save them, then you won't be able to check for skipped points.)" _
                         & vbCrLf & "" _
                         & vbCrLf & "Save them now?" _
                         , vbOKCancel Or vbInformation Or vbDefaultButton1, "Save search points")
      
        Case vbOK
            cmdSaveAll.value = True
        Case vbCancel
      
      End Select
      End If

   'reload the saved points, and begin automatic analysis
   If Dir(drivjk_c$ & "mappoints.sav") <> sEmpty Then
      'determine number of rows
      savfil% = FreeFile
      Open drivjk_c$ & "mappoints.sav" For Input As #savfil%
      mRows& = 0
      Do Until EOF(savfil%)
         Line Input #savfil%, doclin$
         mRows& = mRows& + 1
      Loop
      Close #savfil%
      sky2.Rows = mRows& + 1
      
      sky2.Clear
      savfil% = FreeFile
      Open drivjk_c$ & "mappoints.sav" For Input As #savfil%
      i& = 1
      Do Until EOF(savfil%)
         If i& > sky2.Rows - 1 Then Exit Do
         Input #savfil%, savlat, savlon, savhgt, savdis
         sky2.TextArray(skyp2(i&, 0)) = i&
         sky2.TextArray(skyp2(i&, 1)) = savlat
         sky2.TextArray(skyp2(i&, 2)) = savlon
         sky2.TextArray(skyp2(i&, 3)) = savhgt
         sky2.TextArray(skyp2(i&, 4)) = savdis
         i& = i& + 1
      Loop
      Close #savfil%
      
      If world = False Then
         sky2.FormatString = "^Point # |^    ITMx          |^    ITMy             |^ height(m)  |^distance(km)"
      Else
         sky2.FormatString = "^Point #|^ latitude       |^longitude      |^height (m)|^distance(km)       "
         End If
      
      response = MsgBox("Begin automatic profile generation using these ponts for this metro area?", _
                        vbYesNoCancel + vbQuestion, "Maps & more")
      If response = vbYes Then
         With mapsearchfm
            .cmdCheckSkipped.Visible = False
            .ProgressBarProfs.Visible = True
            .ProgressBarProfs.Min = 0
            .ProgressBarProfs.Max = i& - 1
            .StatusBarProg.Panels(1).Text = i& - 1
            .StatusBarProg.Panels(2).Text = "0"
         End With
         RunOnAutomatic 'automatically process profiles
         End If
   Else
      MsgBox "Can't find the mappoints.sav file!", vbOKOnly + vbExclamation, "Maps & More"
      End If

   On Error GoTo 0
   Exit Sub

cmdReLoad_Click_Error:
    Close
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdReLoad_Click of Form mapsearchfm"
    Resume
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSave_Click
' DateTime  : 9/13/2003 20:55
' Author    : Chaim Keller
' Purpose   : Save search points to buffer
'---------------------------------------------------------------------------------------
'
Private Sub cmdSave_Click()
   On Error GoTo cmdSave_Click_Error

   savnum& = sky2.row
   If savnum& <> 0 Then
     'save the highlighted item's coordinates into the temporary buffer
     savlat = sky2.TextArray(skyp2(savnum&, 1))
     savlon = sky2.TextArray(skyp2(savnum&, 2))
     savhgt = sky2.TextArray(skyp2(savnum&, 3))
     savdist = sky2.TextArray(skyp2(savnum&, 4))
     If Dir(drivjk_c$ & "mappoints.sav") <> sEmpty Then
        'check that this point hasn't already been recorded
        savfil% = FreeFile
        Open drivjk_c$ & "mappoints.sav" For Input As #savfil%
        Do Until EOF(savfil%)
           Input #savfil%, savlat2, savlon2, savhgt2, savdist2
           If savlat2 = Val(savlat) And savlon2 = Val(savlon) Then
              MsgBox "Point already saved in buffer!", vbExclamation + vbOKOnly, "Maps & More"
              Close #savfil%
              Exit Sub 'don't record this point
              End If
        Loop
        Close #savfil%
        End If
     'now record this point
     savfil% = FreeFile
     Open drivjk_c$ & "mappoints.sav" For Append As #savfil%
     Write #savfil%, Val(savlat), Val(savlon), Val(savhgt), Val(savdist)
     Close #savfil%
     End If

   On Error GoTo 0
   Exit Sub

cmdSave_Click_Error:

    Close
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSave_Click of Form mapsearchfm"
End Sub

Private Sub cmdTotSearch_Click()
   'Global mosaic search over the highlighted metro area name
   'using two search databases:
   '(1) The coordinates of the existing profiles for the metro area
   '(2) The coordinates of the sub-city areas of that metro area
   'The mosaic search uses the search displayed search paramters,
   'i.e., the inputed search radius and search step size and
   'mosaic size.
   'The search results are dumped to a check list box
   
   'open the selected metro area directory
   'myfile = Dir(drivcities$ & "\eros\" & Combo2.Text, vbDirectory)
   'If myfile <> sEmpty Then
   '   'look for sav file
   '   myfile2 = Dir(drivcities$ & "\eros\" & Combo2.Text & ".sav")
   '   If myfile2 <> sEmpty Then
   '      savfil% = FreeFile
   '      Open drivcities$ & "\eros\" & Combo2.Text & ".sav" For Input As #savfil%
   '      Do Until EOF(savfil%)
         
   '   End If
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAll_Click
' Author    : chaim
' Date      : 7/21/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveAll_Click()
   On Error GoTo cmdSaveAll_Click_Error
   
   Dim VerifyMode As Boolean
   
   savfil% = FreeFile
   
'   Select Case MsgBox("Do you want to append these points to the former list?" _
'                      & vbCrLf & "" _
'                      & vbCrLf & "Answer:" _
'                      & vbCrLf & "" _
'                      & vbCrLf & "       Yes -- to append" _
'                      & vbCrLf & "       No -- to save/clear the buffer and start a new list" _
'                      & vbCrLf & "" _
'                      & vbCrLf & "       Cancel - to exit this routine" _
'                      , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "Saving search results")
'
'    Case vbYes
'        Open drivjk_c$ & "mappoints.sav" For Append As #savfil%
'
'    Case vbNo
'        Select Case MsgBox("This operation will clear the last search bufer unless you save it." _
'                           & vbCrLf & "" _
'                           & vbCrLf & "Do you want to save the last mappoints.sav list as a ""bak"" file with the currentt date added to the name?" _
'                           , vbYesNo Or vbInformation Or vbDefaultButton1, "Saving search results")
'
'            Case vbYes
'                bufout = FreeFile
'                BufOutName$ = drivjk_c$ & "mappoints_" & Format(Trim$(Month(Now)), "00") & Format(Trim$(Day(Now)), "00") & Format(Year(Now), "00") & ".sav"
'                FileCopy drivjk_c$ & "mappoints.sav", BufOutName$
'            Case vbNo
'
'        End Select
'        cmdClear.value = True
'        Open drivjk_c$ & "mappoints.sav" For Output As #savfil%
'
'    Case vbCancel
'        Close
'        Exit Sub
'
'   End Select
'
'   response = MsgBox("Do you want to verify each point?" & vbLf & vbLf & _
'                     "Answer: " & vbLf & _
'                     "Yes -- to verify each point" & vbLf & _
'                     "No  -- to store all the points w/o verification", _
'                     vbYesNoCancel + vbQuestion, "Maps & More")
'   Select Case response
'     Case vbYes
'       VerifyMode = True
'     Case Else
'       VerifyMode = False
'   End Select
   
   'new interface 061422 //////////////////////////////////////////////
   
   ret = SetWindowPos(mapsearchfm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)

   frmMsgBox.MsgCstm "Do you want to append these points to the former list?", "Saving search results", mbQuestion, 1, False, _
                     "Yes, append", "No, start a new list", "Cancel"
   Select Case frmMsgBox.g_lBtnClicked
       Case 1 'the 1st button in your list was clicked
            Open drivjk_c$ & "mappoints.sav" For Append As #savfil%

       Case 2 'the 2nd button in your list was clicked
            Select Case MsgBox("This operation will clear the last search bufer unless you save it." _
                               & vbCrLf & "" _
                               & vbCrLf & "Do you want to save the last mappoints.sav list as a ""bak"" file with the currentt date added to the name?" _
                               , vbYesNo Or vbInformation Or vbDefaultButton1, "Saving search results")
            
                Case vbYes
                    bufout = FreeFile
                    BufOutName$ = drivjk_c$ & "mappoints_" & Format(Trim$(Month(Now)), "00") & Format(Trim$(Day(Now)), "00") & Format(Year(Now), "00") & ".sav"
                    FileCopy drivjk_c$ & "mappoints.sav", BufOutName$
                Case vbNo
            
            End Select
'            cmdClear.value = True
            Open drivjk_c$ & "mappoints.sav" For Output As #savfil%
        
      Case 0, 3 'cancel.
            Close
            Exit Sub
   End Select
   
   frmMsgBox.MsgCstm "Do you want to verify each point?", "Verify?", mbQuestion, 1, False, _
                     "Yes, verify", "No, store all the points without verifcation", "Cancel"
   Select Case frmMsgBox.g_lBtnClicked
       Case 1 'the 1st button in your list was clicked
            VerifyMode = True

       Case 2 'the 2nd button in your list was clicked
            VerifyMode = False
        
      Case 0, 3 'cancel.
            Close
            Exit Sub
   End Select

   For i& = 1 To sky2.Rows - 1
       
     'save the highlighted item's coordinates into the temporary buffer
     savlat = sky2.TextArray(skyp2(i&, 1))
     savlon = sky2.TextArray(skyp2(i&, 2))
     savhgt = sky2.TextArray(skyp2(i&, 3))
     savdist = sky2.TextArray(skyp2(i&, 4))
     If Val(savlat) = 0 And Val(savlon) = 0 And Val(savhgt) = 0 Then
        'blank record, don't record
     Else
        If VerifyMode Then
           lon = Val(savlon)
           lat = Val(savlat)
           If lon > 180 Or lat > 180 Then
              'this is eretz israel--remove extra decimal places
              kmx = lat
              kmy = lon
              Maps.Text5 = kmx
              Maps.Text6 = kmy
              If savlat < 1000# Then savlat = CInt(1000 * kmx) * 0.001
              If savlon < 1000# Then savlon = CInt(1000 * (kmy - 1000000)) * 0.001
              skymove = False
              worldmove = False
              Call goto_click
           Else
              worldmove = True
              Call goto_click
              jumpworld = False
              skymove = False
              worldmove = False
              End If
           Skycoord% = 0
           Call BringWindowToTop(mapsearchfm.hwnd)
           response = MsgBox("Record this point in the buffer?", _
                       vbQuestion + vbYesNoCancel, "Maps & More")
           If response = vbYes Then
              Write #savfil%, Val(savlat), Val(savlon), Val(savhgt), Val(savdist)
           ElseIf response = vbCancel Then
              Close #savfil%
              Exit Sub
              End If
         Else
          If savlon > 180 Or savlat > 180 Then
             'this is eretz israel--remove extra decimal places]
             savlat = CLng((1000# * savlat) * 0.001)
             savlon = CLng((1000# * (savlon - 1000000#)) * 0.001) + 1000000
             End If
           Write #savfil%, Val(savlat), Val(savlon), Val(savhgt), Val(savdist)
           End If
        End If
   Next i&
   Close #savfil%
   
   SavedAll = True

   On Error GoTo 0
   Exit Sub

cmdSaveAll_Click_Error:
    response = MsgBox("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSaveAll_Click of Form mapsearchfm" & vbLf & _
               "lon,lat,hgt,dist=" & savlon & "," & savlat & "," & savhgt & "," & savdist & vbLf & vbLf & _
               "Do you want to retry or abort?", _
               vbQuestion + vbAbortRetryIgnore, "Maps & More")
    Select Case response
       Case vbRetry, vbIgnore
           Resume Next
       Case Else
           Close
    End Select

   On Error GoTo 0
   Exit Sub

End Sub

Private Sub Combo1_Click()
   If Combo2.Text = sEmpty Or Combo1.Text = sEmpty Then Exit Sub
   subcity$ = Combo1.Text
   tmpfil$ = drivcities$ + "eros\" + Combo2.Text + ".sav"
   erosfil% = FreeFile
   Open tmpfil$ For Input As #erosfil%
   Do Until EOF(erosfil%)
      Input #erosfil%, doclin$, latcity, loncity, hgtcity
      If doclin$ = subcity$ Then
         Text1 = latcity
         If world Then
            Text2 = -loncity
         Else
            Text2 = loncity
            End If
         Close #erosfil%
         Exit Sub
         End If
   Loop
   Close #erosfil%
End Sub

Private Sub Combo2_Click()
   On Error GoTo Combo2_Click_Error

      tmpfil$ = drivcities$ + "eros\" + Combo2.Text + ".sav"
      myfile = Dir(tmpfil$)
      nn% = 0
      If myfile <> sEmpty Then
        Combo1.Clear
        erosfil% = FreeFile
        nncity% = 0
        Open tmpfil$ For Input As #erosfil%
        Do Until EOF(erosfil%)
           Input #erosfil%, doclin$, latcity, loncity, hgtcity
           Combo1.AddItem doclin$
           subcitnams(nncity%) = doclin$
           nncity% = nncity% + 1
        Loop
        Close #erosfil%
        Combo1.ListIndex = Combo1.ListCount - 1
        If world = True And txtStep.Text = sEmpty Then
           txtStep.Text = "0.5"
           End If
        End If

   On Error GoTo 0
   Exit Sub

Combo2_Click_Error:
    Close #erosfil%
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Combo2_Click of Form mapsearchfm"
End Sub

Private Sub Combo2_DblClick()
      tmpfil$ = drivcities$ + "eros\" + Combo2.Text + ".sav"
      myfile = Dir(tmpfil$)
      nn% = 0
      If myfile <> sEmpty Then
        Combo1.Clear
        erosfil% = FreeFile
        nncity% = 0
        Open tmpfil$ For Input As #erosfil%
        Do Until EOF(erosfil%)
           Input #erosfil%, doclin$, latcity, loncity, hgtcity
           Combo1.AddItem doclin$
           subcitnams(nncity%) = doclin$
           nncity% = nncity% + 1
        Loop
        Close #erosfil%
        Combo1.ListIndex = Combo1.ListCount - 1
        End If
End Sub

Private Sub Command1_Click()
   'perform search
   Dim Dist As Double
   Dim cosang As Double
   Dim numpnts&
   
   On Error GoTo Command1_Click_Error

   If txtStep.Text = sEmpty Or Val(txtStep.Text) = 0 Then
      response = MsgBox("Search stepsize must be nonzero!", vbCritical + vbOKOnly, "Maps & More")
      Exit Sub
      End If
   If SearchType% = 1 And (Trim$(txtMosaic.Text) = sEmpty Or Val(txtMosaic.Text) = 0) Then
      response = MsgBox("Stepsize of search inside mosaics must be nonzero!", vbCritical + vbOKOnly, "Maps & More")
      Exit Sub
      End If
   If noheights = True Then
      response = MsgBox("Please load the appropriate DTM into the CD-drive!", vbOKOnly + vbCritical, "Maps & More")
      Exit Sub
      End If
      
   If SearchVis Then
      sky2.Clear
      SearchVis = False
      blitpictures
      End If
      
   SavedAll = False
      
   mapsearchfm.Width = 5490
   mapsearchfm.frmProg.Visible = True
   mapsearchfm.Refresh
   mapsearchfm.progSearch.Min = 0
   mapsearchfm.progSearch.value = 0
   
   Screen.MousePointer = vbHourglass
   numheights = 0
   
   del = Val(txtStep.Text)
   If del = 0 Then
      If SearchType% = 0 Then
         del = 0.05
         txtStep.Text = "0.05"
         End If
      If SearchType% = 1 Then
         del = 0.06
         txtStep.Text = "0.2"
         End If
      End If
   If world = True Then
      If SearchType% = 0 Then 'simple search
        ydegkm = 180 / (pi * 6371.315)   'degrees per km latitude
        xdegkm = 180 / (pi * 6371.315 * Cos(cd * Text1))   'degrees per km longitude
        If (Text3 < 1) Then ydegkm = xdegkm 'for a small search radius, use the same step size for x and y
        y11 = Val(Text1) - Val(Text3) * ydegkm '- del * ydegkm 'move beyond the borders by del
        y12 = Val(Text1) + Val(Text3) * ydegkm '+ del * ydegkm
        x11 = Val(Text2) - Val(Text3) * xdegkm '- del * xdegkm
        x12 = Val(Text2) + Val(Text3) * xdegkm '+ del * xdegkm
        BegMosaicX = 0
        EndMosaicX = 0
        BegMosaicY = 0
        EndMosaicY = 0
        MosaicStepX = 1
        MosaicStepy = 1

        numSlots& = Val(Text4)
        
        ReDim searchhgts(3, Val(Text4) + 1)
        
      ElseIf SearchType% = 1 Then 'mosaic search
        ydegkm = 180 / (pi * 6371.315) 'degrees per km latitude
        xdegkm = 180 / (pi * 6371.315 * Cos(cd * Text1)) 'degrees per km longitude
        MosaicStepX = Val(txtMosaic.Text) * xdegkm
        MosaicStepy = Val(txtMosaic.Text) * ydegkm
        BegMosaicY = Val(Text1) - Val(Text3) * ydegkm  '- MosaicStepY
        EndMosaicY = Val(Text1) + Val(Text3) * ydegkm  '+ MosaicStepY
        BegMosaicX = Val(Text2) - Val(Text3) * xdegkm  '- MosaicStepX
        EndMosaicX = Val(Text2) + Val(Text3) * xdegkm  '+ MosaicStepX
        y11 = BegMosaicY
        y12 = BegMosaicY + MosaicStepy
        x11 = BegMosaicX
        x12 = BegMosaicX + MosaicStepX
        numSlots& = 1
        End If
      
   Else
      If SearchType% = 0 Then 'simple search
        ydegkm = del * 1000
        xdegkm = del * 1000
        y11 = Val(Text1) - Val(Text3 * 1000)
        y12 = Val(Text1) + Val(Text3 * 1000)
        x11 = Val(Text2) - Val(Text3 * 1000)
        x12 = Val(Text2) + Val(Text3 * 1000)
        BegMosaicX = 0
        EndMosaicX = 0
        BegMosaicY = 0
        EndMosaicY = 0
        MosaicStepX = 1
        MosaicStepy = 1
        numSlots& = Val(Text4)
        
        ReDim searchhgts(3, Val(Text4) + 1)
      
      ElseIf SearchType% = 1 Then 'mosaic search
        ydegkm = del * 1000
        xdegkm = del * 1000
        MosaicStepX = Val(txtMosaic.Text) * 1000
        MosaicStepy = Val(txtMosaic.Text) * 1000
        BegMosaicY = Val(Text1) - Val(Text3 * 1000)
        EndMosaicY = Val(Text1) + Val(Text3 * 1000) - MosaicStepy
        BegMosaicX = Val(Text2) - Val(Text3 * 1000)
        EndMosaicX = Val(Text2) + Val(Text3 * 1000) - MosaicStepX
        y11 = BegMosaicY
        y12 = BegMosaicY + MosaicStepy
        x11 = BegMosaicX
        x12 = BegMosaicX + MosaicStepX
        numSlots& = 1
        End If
      End If
      
   numMosaic& = ((EndMosaicY - BegMosaicY) / MosaicStepy + 1) * ((EndMosaicX - BegMosaicX) / MosaicStepX + 1)
   numsearchpnts& = numMosaic& * (2 * (y12 - y11) / (ydegkm) + 1) * (2 * (x12 - x11) / (xdegkm) + 1)
   If SearchType% = 1 Then
      ReDim searchhgts(3, numMosaic& + 1)
      Text4.Text = Str$(numMosaic&)
      If world Then
         numsearchpnts& = numMosaic& * (2 * (y12 - y11) / (del * ydegkm) + 1) * (2 * (x12 - x11) / (del * xdegkm) + 1)
         End If
      End If
   mapsearchfm.progSearch.Max = numsearchpnts&
   
   nn& = 0
   numMos& = 0
   numpnts& = -1
   For iYMosaic = BegMosaicY To EndMosaicY Step MosaicStepy
   For iXMosaic = BegMosaicX To EndMosaicX Step MosaicStepX
      If SearchType% = 1 Then
         numMos& = numMos& + 1
         If numMos& > numMosaic& Then
            GoTo sr500
            End If
         y11 = iYMosaic
         y12 = iYMosaic + MosaicStepy
         x11 = iXMosaic
         x12 = iXMosaic + MosaicStepX
         End If
   init = 0
   If world = True Then
      multstep = del
   Else
      multstep = 1
      End If
   init = 0
   
   If SearchType% = 0 Then
       numsearchpnts& = ((y12 - y11) / (ydegkm * 0.5 * multstep) + 1) * ((x12 - x11) / (xdegkm * 0.5 * multstep) + 1)
       mapsearchfm.progSearch.Max = numsearchpnts&
    
       ReDim searchhgts(3, Val(Text4) + 1)
       End If
   
   For kmys = y11 To y12 Step ydegkm * 0.5 * multstep
      For kmxs = x11 To x12 Step xdegkm * 0.5 * multstep
        DoEvents
        If CancelSearch Then GoTo sr500
        If world = True Then
'           dist = Sqr(((kmxs - Val(Text2)) / xdegkm) ^ 2 + ((kmys - Val(Text1)) / ydegkm) ^ 2)
           X1 = Cos(Val(Text1) * cd) * Cos(Val(Text2) * cd)
           X2 = Cos(kmys * cd) * Cos(kmxs * cd)
           Y1 = Cos(Val(Text1) * cd) * Sin(Val(Text2) * cd)
           Y2 = Cos(kmys * cd) * Sin(kmxs * cd)
           Z1 = Sin(Val(Text1) * cd)
           Z2 = Sin(kmys * cd)
           'this is a calculation of
           'the shortest geodesic distance and is given by
           'Re * Angle between vectors
           'cos(Angle between unit vectors) = Dot product of unit vectors
           'this is considerably smaller than the straight line distance
           'for large distances.  To calculate that distance you
           'need to use the CrossSection option
           cosang = X1 * X2 + Y1 * Y2 + Z1 * Z2
           If Abs(cosang - 1) > 0.000000001 Then
              Dist = 6371.315 * DACOS(cosang)
           Else
              Dist = 0
              Dist = 6371.315 * Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
              End If
           
           'If SearchType% = 1 Then dist = dist * 0.5
        Else
           Dist = 0.0005 * Sqr((kmxs - Val(Text2)) ^ 2 + (kmys - Val(Text1)) ^ 2)
           End If
        If Dist <= Val(Text3) Then '/ 2# Then  'within the search radius
            If init = 0 Then numpnts& = numpnts& + 1
            If world = True Then
               Call worldheights(kmxs, kmys, hgts)
            Else
               kmxs1 = kmxs
               kmys1 = kmys
               Call heights(kmxs1, kmys1, hgts)
               End If
               
            If SearchType% = 0 Then
            If init = 0 Then
               If numheights <= numSlots& Then
                  numheights = numheights + 1
                  searchhgts(0, numheights - 1) = kmys
                  searchhgts(1, numheights - 1) = kmxs
                  searchhgts(2, numheights - 1) = hgts
                  searchhgts(3, numheights - 1) = Val(Format(Str$(Dist), "####0.0###"))
               Else
                  init = 1
                  End If
               End If
            If init = 1 Then 'find highest points
               For i& = 1 To numSlots&
                  If hgts > searchhgts(2, i& - 1) Then
                     'find minimum searchhgts, and replace it with this height
s50:                 minhgt0 = searchhgts(2, 1)
                     For j& = 2 To numSlots&
                        If searchhgts(2, j&) < minhgt0 Then
                           minhgt = searchhgts(2, j& - 1)
                           minkmy = searchhgts(0, j& - 1)
                           minkmx = searchhgts(1, j& - 1)
                           mindist = searchhgts(3, j& - 1)
                           minhgtindex = j&
                           searchhgts(2, j&) = searchhgts(2, 0)
                           searchhgts(0, j&) = searchhgts(0, 0)
                           searchhgts(1, j&) = searchhgts(1, 0)
                           searchhgts(3, j&) = searchhgts(3, 0)
                           searchhgts(2, 1) = minhgt
                           searchhgts(0, 1) = minkmy
                           searchhgts(1, 1) = minkmx
                           searchhgts(3, 1) = mindist
                           GoTo s50
                           End If
                     Next j&
                     'replace minimum height with newly found higher height
                     searchhgts(2, 1) = hgts
                     searchhgts(0, 1) = kmys
                     searchhgts(1, 1) = kmxs
                     searchhgts(3, 1) = Dist
                     Exit For
                     End If
                 Next i&
               End If
            ElseIf SearchType% = 1 Then
               If init = 0 Then
                  searchhgts(0, numpnts&) = kmys
                  searchhgts(1, numpnts&) = kmxs
                  searchhgts(2, numpnts&) = hgts
                  searchhgts(3, numpnts&) = Dist
                  init = 1
               Else
                  If hgts > searchhgts(2, numpnts&) Then
                     searchhgts(0, numpnts&) = kmys
                     searchhgts(1, numpnts&) = kmxs
                     searchhgts(2, numpnts&) = hgts
                     searchhgts(3, numpnts&) = Dist
                     End If
                  End If
              End If
            End If
sr250:
            nn& = nn& + 1
            newCaption$ = "Searching...  " & Trim$(Str$(CInt((nn& / numsearchpnts&) * 100))) & " % complete"
            If newCaption$ <> mapsearchfm.Caption Then mapsearchfm.Caption = newCaption$
            If nn& >= mapsearchfm.progSearch.Max Then
               GoTo sr300
               End If
            mapsearchfm.progSearch.value = nn&
sr300:
      Next kmxs
   Next kmys
   
   If chkValidity.value = vbChecked Then
      'check for unique valid peak, e.g., highest point is actually on slope, or near highest point beyond mosaic border
      Dim StepSize As Double
      Dim NumStepsToCheck As Integer
      
      StepSize = xdegkm * 0.5 * multstep

      'numstepstocheck is a quarter of a mosaic
      NumStepsToCheck = 0.25 * MosaicStepX / StepSize
      For kmxs = searchhgts(1, numpnts&) - NumStepsToCheck * StepSize To searchhgts(1, numpnts&) + NumStepsToCheck * StepSize Step StepSize
        kmys = searchhgts(0, numpnts&)
        If world = True Then
           Call worldheights(kmxs, kmys, hgts)
        Else
           kmxs1 = kmxs
           kmys1 = kmys
           Call heights(kmxs1, kmys1, hgts)
           End If
        If hgts > searchhgts(2, numpnts&) Then
           'remove this entry
           numpnts& = numpnts& - 1
           numsearchpnts& = numsearchpnts& - 1
           GoTo SkipEntry
           End If
      Next kmxs
      
      StepSize = ydegkm * 0.5 * multstep
      NumStepsToCheck = 0.25 * MosaicStepy / StepSize
      For kmys = searchhgts(0, numpnts&) - NumStepsToCheck * StepSize To searchhgts(0, numpnts&) + NumStepsToCheck * StepSize Step StepSize
        kmxs = searchhgts(1, numpnts&)
        If world = True Then
           Call worldheights(kmxs, kmys, hgts)
        Else
           kmxs1 = kmxs
           kmys1 = kmys
           Call heights(kmxs1, kmys1, hgts)
           End If
        If hgts > searchhgts(2, numpnts&) Then
           'remove this entry
           numpnts& = numpnts& - 1
           numsearchpnts& = numsearchpnts& - 1
           GoTo SkipEntry
           End If
      Next kmys
      End If
      
      'now do a proximity search
      For i& = 0 To numpnts& - 1
          'determine distance between high places
          If world = True Then
'            dist = Sqr(((kmxs - Val(Text2)) / xdegkm) ^ 2 + ((kmys - Val(Text1)) / ydegkm) ^ 2)
             X1 = Cos(searchhgts(0, i&) * cd) * Cos(searchhgts(0, i&) * cd)
             X2 = Cos(searchhgts(0, numpnts&) * cd) * Cos(searchhgts(1, numpnts&) * cd)
             Y1 = Cos(searchhgts(0, i&) * cd) * Sin(searchhgts(0, i&) * cd)
             Y2 = Cos(searchhgts(0, numpnts&) * cd) * Sin(searchhgts(1, numpnts&) * cd)
             Z1 = Sin(searchhgts(0, i&) * cd)
             Z2 = Sin(searchhgts(0, numpnts&) * cd)
             'this is a calculation of
             'the shortest geodesic distance and is given by
             'Re * Angle between vectors
             'cos(Angle between unit vectors) = Dot product of unit vectors
             'this is considerably smaller than the straight line distance
             'for large distances.  To calculate that distance you
             'need to use the CrossSection option
             cosang = X1 * X2 + Y1 * Y2 + Z1 * Z2
             If Abs(cosang - 1) > 0.000000001 Then
                Dist = 6371.315 * DACOS(cosang)
             Else
                Dist = 0
                Dist = 6371.315 * Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
                End If
          Else
             'distance in kilometers
             Dist = 0.001 * Sqr((searchhgts(1, numpnts&) - searchhgts(1, i&)) ^ 2 + (searchhgts(0, numpnts&) - searchhgts(0, i&)) ^ 2)
             End If
          If Dist < 0.25 * Val(txtMosaic.Text) Then
             If searchhgts(2, numpnts&) > searchhgts(2, i&) Then
                'remove previous points by switching
                searchhgts(0, i&) = searchhgts(0, numpnts&)
                searchhgts(1, i&) = searchhgts(1, numpnts&)
                searchhgts(2, i&) = searchhgts(2, numpnts&)
                searchhgts(3, i&) = searchhgts(3, numpnts&)
                numpnts& = numpnts& - 1
                numsearchpnts& = numsearchpnts& - 1
                GoTo SkipEntry
             ElseIf searchhgts(2, numpnts&) <= searchhgts(2, i&) Then
                'remove this point
                numpnts& = numpnts& - 1
                numsearchpnts& = numsearchpnts& - 1
                GoTo SkipEntry
                End If
            End If
      Next i&
      
SkipEntry:
   Next iXMosaic
   Next iYMosaic
   
   If chkProfiles.value = vbChecked Then
    'determine if peak is a true peak in x/y over a typical distance
    Dim SymmPeakArray() As Integer
    Dim numSymmPeaks&, numSlope%
    numSlope% = 15 'number of grid steps to check for descending slope
    Dim SlopeOut() As Single
    Dim FileSlopeName$, slopenum%
    ReDim SlopeOut(3, numSlope% - 1) As Single
    
    numSymmPeaks& = 0
    For i& = 0 To numpnts& - 1
        pkkmy = searchhgts(0, i&)
        pkkmx = searchhgts(1, i&)
        pkhgt = searchhgts(2, i&)
        
        hgts1 = pkhgt
        hgts2 = pkhgt
        hgts3 = pkhgt
        hgts4 = pkhgt
        
        For j& = 1 To numSlope%
        
          kmys = pkkmy
          kmxs = pkkmx + j& * xdegkm
          If world = True Then
             Call worldheights(kmxs, kmys, hgts)
          Else
             kmxs1 = kmxs
             kmys1 = kmys
             Call heights(kmxs1, kmys1, hgts)
             End If
             
          If hgts > hgts1 Then
             GoTo nextentry
          Else
             hgts1 = hgts
             SlopeOut(0, j& - 1) = hgts / pkhgt
             End If
        
          kmxs = pkkmx - j& * xdegkm
          If world = True Then
             Call worldheights(kmxs, kmys, hgts)
          Else
             kmxs1 = kmxs
             kmys1 = kmys
             Call heights(kmxs1, kmys1, hgts)
             End If
             
          If hgts > hgts2 Then
             GoTo nextentry
          Else
             hgst2 = hgts
             SlopeOut(1, j& - 1) = hgts / pkhgt
             End If
        
          kmxs = pkkmx
          kmys = pkkmy + j& * ydegkm
          If world = True Then
             Call worldheights(kmxs, kmys, hgts)
          Else
             kmxs1 = kmxs
             kmys1 = kmys
             Call heights(kmxs1, kmys1, hgts)
             End If
             
          If hgts > hgts3 Then
             GoTo nextentry
          Else
             hgts3 = hgts
             SlopeOut(2, j& - 1) = hgts / pkhgt
             End If
        
          kmys = pkkmy - j& * ydegkm
          If world = True Then
             Call worldheights(kmxs, kmys, hgts)
          Else
             kmxs1 = kmxs
             kmys1 = kmys
             Call heights(kmxs1, kmys1, hgts)
             End If
             
          If hgts > hgts4 Then
             GoTo nextentry
          Else
             hgts4 = hgts
             SlopeOut(3, j& - 1) = hgts / pkhgt
             End If
          
        Next j&
        
        numSymmPeaks& = numSymmPeaks& + 1
        ReDim Preserve SymmPeakArray(numSymmPeaks&)
        SymmPeakArray(numSymmPeaks& - 1) = i&
        
        'output slope files
        FileSlopeName$ = App.Path & "\FS" & Trim$(Str$(pkkmx)) & "-" & Trim$(Str$(pkkmy)) & "-xp.txt"
        slopenum% = FreeFile
        Open FileSlopeName$ For Output As #slopenum%
        Print #slopenum%, "0.0, 1.0"
        For j& = 1 To numSlope%
            Print #slopenum%, Str$(j& * xdegkm) & "," & Format(Str$(SlopeOut(0, j& - 1)), "#0.0###")
        Next j&
        Close #slopenum%
        FileSlopeName$ = App.Path & "\FS" & Trim$(Str$(pkkmx)) & "-" & Trim$(Str$(pkkmy)) & "-xn.txt"
        slopenum% = FreeFile
        Open FileSlopeName$ For Output As #slopenum%
        For j& = numSlope% To 1 Step -1
            Print #slopenum%, Str$(-j& * xdegkm) & "," & Format(Str$(SlopeOut(1, j& - 1)), "#0.0###")
        Next j&
        Print #slopenum%, "0.0, 1.0"
        Close #slopenum%
        FileSlopeName$ = App.Path & "\FS" & Trim$(Str$(pkkmx)) & "-" & Trim$(Str$(pkkmy)) & "-yp.txt"
        slopenum% = FreeFile
        Open FileSlopeName$ For Output As #slopenum%
        Print #slopenum%, "0.0, 1.0"
        For j& = 1 To numSlope%
            Print #slopenum%, Str$(j& * ydegkm) & "," & Format(Str$(SlopeOut(2, j& - 1)), "#0.0###")
        Next j&
        Close #slopenum%
        FileSlopeName$ = App.Path & "\FS" & Trim$(Str$(pkkmx)) & "-" & Trim$(Str$(pkkmy)) & "-yn.txt"
        slopenum% = FreeFile
        Open FileSlopeName$ For Output As #slopenum%
        For j& = numSlope% To 1 Step -1
            Print #slopenum%, Str$(-j& * ydegkm) & "," & Format(Str$(SlopeOut(3, j& - 1)), "#0.0###")
        Next j&
        Print #slopenum%, "0.0, 1.0"
        Close #slopenum%
            
nextentry:
    Next i&
    
    'now repack searchhgts array
    If numpnts& <> numSymmPeaks& Then
        For j& = 0 To numSymmPeaks& - 1
            searchhgts(0, j&) = searchhgts(0, SymmPeakArray(j&))
            searchhgts(1, j&) = searchhgts(1, SymmPeakArray(j&))
            searchhgts(2, j&) = searchhgts(2, SymmPeakArray(j&))
            searchhgts(3, j&) = searchhgts(3, SymmPeakArray(j&))
        Next j&
        numpnts& = numSymmPeaks&
        End If
        
    End If
   
sr500:
   mapsearchfm.frmProg.Visible = False
   mapsearchfm.Width = 4770
   CancelSearch = False
   mapsearchfm.Caption = "Search for highest points"
   
   'now sort the points from nearest to farthest distances
   'or form highest to lowest elevation
   'and place them into Flex-Grid
    If SearchType% = 0 Then numpnts& = numSlots&
    sky2.Rows = numpnts& + 1 'Val(Text4) + 1
    If world = True Then
        For i& = 0 To numpnts& - 1 'Val(Text4)
            
           If EastOnly Then
              If searchhgts(1, i&) > Text2 Then GoTo sr600
              End If
              
           If WestOnly Then
              If searchhgts(1, i&) < Text2 Then GoTo sr600
              End If
            
            If chkIgnoreZeros.value = vbChecked And _
               searchhgts(2, i&) <= Val(txtHgtLimit.Text) Then
               'ignore points with hgts<=0
               GoTo sr600
               End If
                
            If chkIgnoreLarge.value = vbChecked And _
               searchhgts(2, i&) >= Val(txtIgnoreLarge.Text) Then
               'point is larger than upper limit
               'so ignore it
               GoTo sr600
               End If
               
           If chkLineSight.value = vbChecked Then
              'check if within line of sight before adding to array
              crosssectionpnt(0, 0) = Text2 'center longitude
              crosssectionpnt(0, 1) = Text1 'center lat
              crosssectionhgt(0) = Maps.Text7 'height of center point
              crosssectionpnt(1, 0) = searchhgts(1, i&) 'search point longitude
              crosssectionpnt(1, 1) = searchhgts(0, i&) 'search point lat
              crosssectionhgt(1) = searchhgts(2, i&) 'height of search point
              'since small distance, use straight line on mercator projection
              greatcircle = False
              SearchCrossSection = True
              GoCrossSection = True
              mapCrossSections 'calculate if point is obstructed
              If SearchCrossObstruct Then GoTo sr600 'point is obstructed so ignore it
              End If
           
           sky2.TextArray(skyp2(i& + 1, 0)) = i& + 1
           sky2.TextArray(skyp2(i& + 1, 1)) = Format(Trim$(Str$(searchhgts(0, i&))), "###0.0#####")
           sky2.TextArray(skyp2(i& + 1, 2)) = Format(Trim$(Str$(searchhgts(1, i&))), "###0.0#####")
           sky2.TextArray(skyp2(i& + 1, 3)) = Format(Trim$(Str$(searchhgts(2, i&))), "###0.0")
           sky2.TextArray(skyp2(i& + 1, 4)) = Format(Trim$(Str$(searchhgts(3, i&))), "###0.0###")
sr600:
        Next i&
    Else
        For i& = 0 To numpnts& - 1 'Val(Text4)
        
           If EastOnly Then
              If searchhgts(1, i&) > Text2 Then GoTo sr650
              End If
              
           If WestOnly Then
              If searchhgts(1, i&) < Text2 Then GoTo sr650
              End If
            
           If chkIgnoreZeros.value = vbChecked And _
               searchhgts(2, i&) <= Val(txtHgtLimit.Text) Then
               'ignore points with hgts<=0
               'numPnts& = numPnts& - 1
               'sky2.Rows = numPnts& + 1
               'i& = i& - 1
               GoTo sr650
               End If
        
           If chkIgnoreLarge.value = vbChecked And _
               searchhgts(2, i&) >= Val(txtIgnoreLarge.Text) Then
               'point is larger than upper limit
               'so ignore it
               GoTo sr650
               End If
               
           If chkLineSight.value = vbChecked Then
              'check if within line of sight before adding to array
              crosssectionpnt(0, 0) = Text2 'starting longitude
              crosssectionpnt(0, 1) = Text1 'starting lat
              crosssectionhgt(0) = Maps.Text7 'height of center point
              crosssectionpnt(1, 0) = searchhgts(1, i&) 'search point longitude
              crosssectionpnt(1, 1) = searchhgts(0, i&) 'search point lat
              crosssectionhgt(1) = searchhgts(2, i&) 'height of search point
              'since small distance, use straight line on mercator projection
              greatcircle = False
              SearchCrossSection = True
              GoCrossSection = True
              mapCrossSections 'calculate if point is obstructed
              If SearchCrossObstruct Then GoTo sr650 'point is obstructed so ignore it
              End If
           
           sky2.TextArray(skyp2(i& + 1, 0)) = i& + 1
           sky2.TextArray(skyp2(i& + 1, 1)) = searchhgts(1, i&)
           sky2.TextArray(skyp2(i& + 1, 2)) = searchhgts(0, i&)
           sky2.TextArray(skyp2(i& + 1, 3)) = searchhgts(2, i&)
           sky2.TextArray(skyp2(i& + 1, 4)) = searchhgts(3, i&)
sr650:
        Next i&
       End If
    Call dosort
    If world = False Then
       sky2.FormatString = "^Point # |^    ITMx           |^    ITMy            |^ height(m)  |^distance(km)"
    Else
       sky2.FormatString = "^Point #|^ latitude       |^longitude      |^height (m)|^distance(km)       "
       End If
    Screen.MousePointer = vbDefault
    OverhWnd = FindWindow(vbNullString, "Overview")
'    If OverhWnd <> 0 Then ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    If OverhWnd <> 0 Then BringWindowToTop (OverhWnd)
    
    'highlight first row
    sky2.SelectionMode = flexSelectionByRow
    sky2.HighLight = flexHighlightWithFocus
    sky2.row = 1
    
    'reclaim memory
    ReDim searchhgts(0, 0)
    
Exit Sub

   On Error GoTo 0
   Exit Sub

Command1_Click_Error:
    Close
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command1_Click of Form mapsearchfm"
End Sub
Function skyp2(row As Long, col As Integer) As Long
     skyp2 = row * sky2.Cols + col
End Function

Private Sub Command10_Click()
   l2 = Val(Text1.Text)
   l1 = Val(Text2.Text)
   If noheights = False Then
       Call worldheights(l1, l2, hgt)
       If hgt = -9999 Then hgt = 0
       searchhgt = hgt
   Else
       searchhgt = 0
       End If
   viewsearch = True
   AutoProf = False
   Call sunrisesunset(0)
   viewsearch = False
'   ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (OverhWnd)
   Call BringWindowToTop(mapsearchfm.hwnd)
End Sub

Private Sub Command11_Click()
     If tblbuttons(3) = 1 Then 'De Lorme 3D-Explorer
        lResult = FindWindow(vbNullString, "Overview")
        If lResult <> 0 Then 'De Lorme 3D-Explorer detected
            TdxhWnd = 0
            bRtn = EnumWindows(AddressOf EnumWndProc, lParam)
            iposit% = InStr(Tdxname, "-  ")
            If iposit% <> 0 Then
              lat3d = Val(Mid$(Tdxname, iposit% + 4, 2)) + Val(Mid$(Tdxname, iposit% + 8, 4)) / 60
              lon3d = -(Val(Mid$(Tdxname, iposit% + 15, 3)) + Val(Mid$(Tdxname, iposit% + 19, 5)) / 60)
              'lat3d = Val(Text1.Text)
              'lon3d = Val(Text2.Text)
              OverhWnd = FindWindow(vbNullString, "Overview")
              'Call BringWindowToTop(OverhWnd)
'              ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              BringWindowToTop (OverhWnd)
              If Abs(Val(Text1.Text) - lat3d) > 0.37167 Or Abs(Val(Maps.Text5.Text) - lon3d > 0.37167) Then
                  'activate find window
                  Call BringWindowToTop(OverhWnd)
                  Call keybd_event(VK_F6, 0, 0, 0) 'activates alt key
                  Call keybd_event(VK_F6, 0, KEYEVENTF_KEYUP, 0)
               Else
                  'move mouse cursor to right place and
                  'depress right mouse key to go there
                  dx1 = -1000 '-30 '30
                  dy1 = -1000 '-240 '60
                  Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
                  waitime = Timer + 0.01
                  Do Until Timer > waitime
                     DoEvents
                  Loop
                  dx1 = (Val(Text2.Text) - lon3d) * 516.6 + 96
                  dy1 = -(Val(Text1.Text) - lat3d) * 516.6 + 156
                  Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
                  waitime = Timer + 0.01
                  Do Until Timer > waitime
                     DoEvents
                  Loop
                  Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'move mouse to Location item
                  Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0) 'move mouse to Location item
                  End If
               End If
            End If
      Else 'jump to this coordinate on the map
         Maps.Text5 = mapsearchfm.Text2
         Maps.Text6 = mapsearchfm.Text1
         Call goto_click
         End If
      If txtStep.Text = sEmpty And world = True Then
         txtStep.Text = "0.5"
         End If
End Sub

Private Sub Command12_Click()
   If sky2.row <> 0 Then
      nplachos& = sky2.row
      Maps.Text6.Text = sky2.TextArray(skyp2(nplachos&, 1))
      Maps.Text5.Text = sky2.TextArray(skyp2(nplachos&, 2))
      Call goto_click
      Call BringWindowToTop(mapsearchfm.hwnd)
      OverhWnd = FindWindow(vbNullString, "Overview")
      If OverhWnd <> 0 Then BringWindowToTop (OverhWnd) 'ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      End If
End Sub

Private Sub Command13_Click()
    Maps.Text6.Text = Text1.Text
    Maps.Text5.Text = Text2.Text
    Call goto_click
    Call BringWindowToTop(mapsearchfm.hwnd)
    OverhWnd = FindWindow(vbNullString, "Overview")
'    If OverhWnd <> 0 Then BringWindowToTop (OverhWnd) 'ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

Private Sub Command14_Click()
  OverhWnd = FindWindow(vbNullString, "Overview")
  If OverhWnd <> 0 Then
'     ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
     BringWindowToTop (OverhWnd)
  Else
     'Exit Sub
     End If
  On Error GoTo c3error
  CommonDialog1.CancelError = True
  CommonDialog1.Filter = "cities.bat files (*.bat)|*.bat|"
  CommonDialog1.FilterIndex = 1
  If world = True Then
     CommonDialog1.FileName = drivcities$ + "eros\*.bat"
  Else
     CommonDialog1.FileName = drivcities$ + "*.bat"
     End If
  CommonDialog1.ShowOpen
  batfile$ = CommonDialog1.FileName
  
  'determine number of rows
  numRows& = 0
  openfilnum% = FreeFile
  Open batfile$ For Input As #openfilnum%
  Line Input #openfilnum%, doclin$
  Do Until EOF(openfilnum%)
     Line Input #openfilnum%, doclin$
     numRows& = numRows& + 1
  Loop
  sky2.Rows = numRows& + 1
  
  If world = False Then
     sky2.FormatString = "^Point # |^    ITMx           |^    ITMy            |^ height(m)  |^distance(km)"
  Else
     sky2.FormatString = "^Point #|^ latitude       |^longitude      |^height (m)|^distance(km)       "
     End If

  Seek #openfilnum%, 1 'rewind
  If world = True Then Line Input #openfilnum%, doclin$
  i& = 0
  If world = True Then
      Do Until EOF(openfilnum%)
        Input #openfilnum%, citynam$, lata, lono, hgto
        If UCase(citynam$) = "VERSION" Then GoTo 500
        i& = i& + 1
        sky2.TextArray(skyp2(i&, 0)) = i&
        sky2.TextArray(skyp2(i&, 1)) = lata
        sky2.TextArray(skyp2(i&, 2)) = -lono
        sky2.TextArray(skyp2(i&, 3)) = hgto
        sky2.TextArray(skyp2(i&, 4)) = 0
500
     Loop
 Else
      Do Until EOF(openfilnum%)
        Input #openfilnum%, citynam$, lata, lono, hgto
        If UCase(citynam$) = "VERSION" Then GoTo 600
        i& = i& + 1
        sky2.TextArray(skyp2(i&, 0)) = i&
        sky2.TextArray(skyp2(i&, 1)) = lata * 1000
        sky2.TextArray(skyp2(i&, 2)) = 1000000 + lono * 1000
        sky2.TextArray(skyp2(i&, 3)) = hgto
        sky2.TextArray(skyp2(i&, 4)) = 0
600
      Loop
  End If
  Close #openfilnum%
c3error:
  Exit Sub
End Sub

Private Sub Command15_Click()
      'first check if this is new place
      For i% = 0 To nncity%
         If Combo1.Text = subcitnams(i%) Then
            Beep
            ret = SetWindowPos(mapsearchfm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            'lResult = FindWindow(vbNullString, mapsearchfm.hWnd)
            'Call BringWindowToTop(lResult)
            response = MsgBox("Neighborhood has already been recorded, do you wan't to overwrite?", vbExclamation + vbYesNo, "Maps & More Warning")
            If response = vbNo Then
'               ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'               ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               BringWindowToTop (OverhWnd)
               BringWindowToTop (mapPictureform.hwnd)
               Exit Sub
            Else
'               ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'               ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               BringWindowToTop (OverhWnd)
               BringWindowToTop (mapPictureform.hwnd)
               Exit For
               End If
            End If
      Next i%
      If Combo2.Text <> sEmpty Then
        tmpfil$ = drivcities$ + "eros\" & Combo2.Text & ".sav"
        myfile = Dir(tmpfil$)
        erosfil% = FreeFile
        Open tmpfil$ For Append As #erosfil%
        latcity = Val(Text1.Text)
        loncity = Val(Text2.Text)
        If world = True Then
           hgtcity = 0
           If Not noheights Then Call worldheights(loncity, latcity, hgtcity)
           Write #erosfil%, Combo1.Text, latcity, -loncity, hgtcity
        Else
           hgtcity = 0
           If Not noheights Then Call heights(loncity, latcity, hgtcity)
           Write #erosfil%, Combo1.Text, latcity, loncity, hgtcity
           End If
        Close #erosfil%
        Combo1.AddItem Combo1.Text
        nncity% = nncity% + 1
        subcitnams(nncity%) = Combo1.Text
        Combo1.ListIndex = Combo1.ListCount - 1
      Else
        Beep
        response = MsgBox("You must enter a name for the city(area)!", vbCritical + vbOKOnly, "Maps&More")
        End If
End Sub

Private Sub Command16_Click()
   'check format of name, first check for spaces
   'also check for "area" and for "_"
   If InStr(Combo2.Text, "area") = 0 Or InStr(Combo2.Text, "_") = 0 Then
      response = MsgBox("Directory name seems to have the wrong format! Example of proper format: Los_Angeles_area_CA_USA.  Do you wan't to continue?", vbExclamation + vbYesNoCancel + vbDefaultButton2, "Maps & More")
      If response <> vbYes Then Exit Sub
      End If
   For i% = 1 To Len(Combo2.Text)
      If Mid$(Combo2.Text, i%, 1) = " " Then
         response = MsgBox("Directory name can't have spaces!", vbCritical + vbOKOnly, "Maps & More")
         Exit Sub
         End If
   Next i%
   myfile = Dir(drivcities$ + "eros\eroscity.sav")
   erosfil% = FreeFile
   If myfile <> sEmpty Then
      Open drivcities$ + "eros\eroscity.sav" For Input As #erosfil%
      'check for duplicates
      Do Until EOF(erosfil%)
         Line Input #erosfil%, doclin$
         If Combo2.Text = doclin$ Then
            Select Case MsgBox("A city area of the same name is already recorded!" _
                               & vbCrLf & "" _
                               & vbCrLf & "Do you ant to nontheless record this place in the ""sav"" file?" _
                               & vbCrLf & "" _
                               & vbCrLf & "(Hint: if you answer ""Yes"", you should edit out the duplication.)" _
                               , vbYesNoCancel Or vbInformation Or vbDefaultButton1, "Duplicate city area or neighborhood")
            
                Case vbYes
            
                Case vbNo, vbCancel
                   Close #erosfil%
                   Exit Sub
            
            End Select
            End If
      Loop
      Close #erosfil%
      Open drivcities$ + "eros\eroscity.sav" For Append As #erosfil%
      Print #erosfil%, Combo2.Text
      Close #erosfil%
   Else
      Open drivcities$ + "eros\eroscity.sav" For Output As #erosfil%
      Write #erosfil%, Combo2.Text
      Close #erosfil%
      End If
   Combo2.AddItem Combo2.Text
   Combo2.ListIndex = Combo2.ListCount - 1
   Combo1.Clear
   nncity% = 0
End Sub

Private Sub Command2_Click()
   Text1 = Maps.Text6
   Text2 = Maps.Text5
   If world = True Then
      If Val(txtStep.Text) < 0.5 Then txtStep.Text = 0.5 'Str(90 / (pi * 6371.315))
   Else
      If Val(txtStep.Text) = 0 Then txtStep.Text = Str(0.0125)
      End If
End Sub

Private Sub dosort()
   sky2.row = 1
   If HeightSort Then 'sort descending by height
      sky2.col = 3
      sky2.ColSel = 3
      sky2.RowSel = sky2.Rows - 1
      sky2.Sort = 2 'generic descending
   Else 'sort ascending by distance
      sky2.col = 4
      sky2.ColSel = 4
      sky2.RowSel = sky2.Rows - 1
      sky2.Sort = 1 'generic ascending
      End If
End Sub

Private Sub cmdClear_Click()
   If mapsearchfm.sky2.Rows > 0 Then
      resp = MsgBox("Clear the search results and buffer?", vbQuestion + vbYesNo, "Maps&More")
      If resp = vbYes Then
         sky2.Clear
         If SearchVis Then
            SearchVis = False
            blitpictures
            End If
         End If
      End If
    Close
    If Dir(drivjk_c$ & "mappoints.sav") <> sEmpty Then
       Kill drivjk_c$ & "mappoints.sav"
       End If
    If Dir(drivjk_c$ & "mapstatus.sav") <> sEmpty Then
       Kill drivjk_c$ & "mapstatus.sav"
       End If
    If Dir(ramdrive & ":\eros.tm7") <> sEmpty Then
       Kill ramdrive & ":\eros.tm7"
       End If
       
    If world = False Then
       sky2.FormatString = "^Point # |^    ITMx           |^    ITMy            |^ height(m)  |^distance(km)"
    Else
       sky2.FormatString = "^Point #|^ latitude       |^longitude      |^height (m)|^distance(km)       "
       End If
       
    AutoProf = False
    SavedAll = False

End Sub


Private Sub Command5_Click()
   nplachos& = sky2.row
   lat = sky2.TextArray(skyp2(nplachos&, 1))
   lon = sky2.TextArray(skyp2(nplachos&, 2))
   searchhgt = sky2.TextArray(skyp2(nplachos&, 3))
   viewsearch = True
   AutoProf = False
   Call sunrisesunset(1)
   viewsearch = False
'   ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (OverhWnd)
   Call BringWindowToTop(mapsearchfm.hwnd)
End Sub

Private Sub Command6_Click()
   nplachos& = sky2.row
   lat = sky2.TextArray(skyp2(nplachos&, 1))
   lon = sky2.TextArray(skyp2(nplachos&, 2))
   searchhgt = sky2.TextArray(skyp2(nplachos&, 3))
   viewsearch = True
   AutoProf = False
   Call sunrisesunset(0)
   viewsearch = False
'   ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (OverhWnd)
   Call BringWindowToTop(mapsearchfm.hwnd)
End Sub

Private Sub Command7_Click()
     If tblbuttons(3) = 1 Then
        lResult = FindWindow(vbNullString, "Overview")
        If lResult <> 0 Then
            TdxhWnd = 0
            bRtn = EnumWindows(AddressOf EnumWndProc, lParam)
            iposit% = InStr(Tdxname, "-  ")
            Text1 = CSng(Val(Mid$(Tdxname, iposit% + 4, 2)) + Val(Mid$(Tdxname, iposit% + 8, 4)) / 60)
            Text2 = -CSng(Val(Mid$(Tdxname, iposit% + 15, 3)) + Val(Mid$(Tdxname, iposit% + 19, 5)) / 60)
        Else
            Exit Sub
            End If
        End If
   If world = True Then
      txtStep.Text = 0.5 'Str(90 / (pi * 6371.315))
   Else
      txtStep.Text = Str(0.0125)
      End If
End Sub

Private Sub Command8_Click()
   If D3dExplorerDir$ <> sEmpty And Not GoogleMapVis Then
     If tblbuttons(3) = 1 Then
        lResult = FindWindow(vbNullString, "Overview")
        If lResult <> 0 Then
            TdxhWnd = 0
            bRtn = EnumWindows(AddressOf EnumWndProc, lParam)
            iposit% = InStr(Tdxname, "-  ")
            If iposit% <> 0 Then
              lat3d = Val(Mid$(Tdxname, iposit% + 4, 2)) + Val(Mid$(Tdxname, iposit% + 8, 4)) / 60
              lon3d = -(Val(Mid$(Tdxname, iposit% + 15, 3)) + Val(Mid$(Tdxname, iposit% + 19, 5)) / 60)
              'lat3d = Val(Maps.Text6.Text)
              'lon3d = Val(Maps.Text5.Text)
              OverhWnd = FindWindow(vbNullString, "Overview")
              'Call BringWindowToTop(OverhWnd)
'              ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              BringWindowToTop (OverhWnd)
              If Abs(Val(Maps.Text6.Text) - lat3d) > 0.37167 Or Abs(Val(Maps.Text5.Text) - lon3d > 0.37167) Then
                  'activate find window
                  Call BringWindowToTop(OverhWnd)
                  Call keybd_event(VK_F6, 0, 0, 0) 'activates alt key
                  Call keybd_event(VK_F6, 0, KEYEVENTF_KEYUP, 0)
               Else
                  'move mouse cursor to right place and
                  'depress right mouse key to go there
                  dx1 = -1000 '-30 '30
                  dy1 = -1000 '-240 '60
                  Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
                  waitime = Timer + 0.01
                  Do Until Timer > waitime
                     DoEvents
                  Loop
                  dx1 = (Val(Maps.Text5.Text) - lon3d) * 516.6 + 96
                  dy1 = -(Val(Maps.Text6.Text) - lat3d) * 516.6 + 156
                  Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
                  waitime = Timer + 0.01
                  Do Until Timer > waitime
                     DoEvents
                  Loop
                  Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'move mouse to Location item
                  Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0) 'move mouse to Location item
                  End If
               End If
            End If
         End If
    Else
       If GoogleMapVis Then
       Else
          frmMap.Visible = True
          End If
          
        If sky2.row <> 0 Then
           nplachos& = sky2.row
           frmMap.txtLat.Text = sky2.TextArray(skyp2(nplachos&, 1))
           frmMap.txtLong.Text = sky2.TextArray(skyp2(nplachos&, 2))
           frmMap.Command2.value = True
           
           Call goto_click
           Call BringWindowToTop(mapsearchfm.hwnd)
           OverhWnd = FindWindow(vbNullString, "Map")
'           If OverhWnd <> 0 Then BringWindowToTop (OverhWnd) 'ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           End If
           
       End If
       
End Sub

Private Sub Command9_Click()
   l2 = Val(Text1.Text)
   l1 = Val(Text2.Text)
   If noheights = False Then
       Call worldheights(l1, l2, hgt)
       If hgt = -9999 Then hgt = 0
       searchhgt = hgt
   Else
       searchhgt = 0
       End If
   viewsearch = True
   AutoProf = False
   Call sunrisesunset(1)
   viewsearch = False
'   ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (overhtwnd)
   Call BringWindowToTop(mapsearchfm.hwnd)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : form_load
' Author    : chaim
' Date      : 7/21/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub form_load()
   On Error GoTo form_load_Error

   mapsearchfm.Width = 4770
   CancelSearch = False
   mapSearchVis = True
   HeightSort = True
   
   If D3dExplorerDir$ <> sEmpty Then
      Command7.Enabled = True
      Command11.Enabled = True
      End If
   
   myfile = Dir(drivcities$ + "eros\eroscity.sav")
   If myfile <> sEmpty Then
      erosfil% = FreeFile
      Open drivcities$ + "eros\eroscity.sav" For Input As #erosfil%
      Do Until EOF(erosfil%)
         Line Input #erosfil%, doclin$
         Combo2.AddItem doclin$
      Loop
      Close #erosfil%
      Combo2.ListIndex = Combo2.ListCount - 1
      tmpfil$ = drivcities$ + "eros\" + Combo2.Text + ".sav"
      myfile = Dir(tmpfil$)
      nn% = 0
      If myfile <> sEmpty Then
        erosfil% = FreeFile
        nncity% = 0
        Open tmpfil$ For Input As #erosfil%
        Do Until EOF(erosfil%)
           Input #erosfil%, doclin$, latcity, loncity, hgtcity
           Combo1.AddItem doclin$
           subcitnams(nncity%) = doclin$
           nncity% = nncity% + 1
        Loop
        Close #erosfil%
        Combo1.ListIndex = Combo1.ListCount - 1
        End If
      End If
      
   optSimple_Click 'simple search is default

   On Error GoTo 0
   Exit Sub

form_load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure form_load of Form mapsearchfm"
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lResult As Long
    Unload mapsearchfm
    Set mapsearchfm = Nothing
    mapSearchVis = False
    
    OverhWnd = FindWindow(vbNullString, "Overview")
    If OverhWnd = 0 Then
'       ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       BringWindowToTop (mapPictureform.hwnd)
       End If
    If world = False Then
      lResult = FindWindow(vbNullString, terranam$)
      If lResult > 0 Then
'         ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (lResult)
         End If
    Else
      lResult = FindWindow(vbNullString, "3D Viewer")
      If lResult > 0 Then
'         ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (lResult)
         End If
      End If
      
End Sub

Private Sub optClear_Click()
   EastOnly = False
   WestOnly = False
End Sub

Private Sub optEast_Click()
   EastOnly = True
   WestOnly = False
End Sub

Private Sub optMosaic_Click()
   lblStep = "Stepsize in mosaic(km)"
   txtStep.ToolTipText = "Step size within the mosaic (km)"
   lblMosaic.Enabled = True
   txtMosaic.Enabled = True
   Text4.Enabled = False
   UpDown2.Enabled = False
   chkValidity.Visible = True
   chkProfiles.Visible = True
   If Trim$(txtMosaic) = sEmpty Then txtMosaic = "1"
   If Trim$(txtStep) = sEmpty Then txtStep = "0.1"
   SearchType% = 1
End Sub

Private Sub optSimple_Click()
   Text4.Enabled = True
   UpDown2.Enabled = True
   lblStep = "Search stepsize(km)"
   txtStep.ToolTipText = "Step size between heights"
   If Trim$(txtStep) = sEmpty Then
      If world Then
         txtStep = "1"
      Else
         txtStep = "0.0125"
         End If
      End If
   SearchType% = 0
   lblMosaic.Enabled = False
   txtMosaic.Enabled = False
   chkValidity.Visible = False
   chkProfiles.Visible = False
   chkProfiles.value = vbUnchecked
End Sub

Private Sub optSortDist_Click()
   HeightSort = False
End Sub

Private Sub optSortHgt_Click()
   HeightSort = True
End Sub

Private Sub optWest_Click()
   WestOnly = True
   EastOnly = False
End Sub

Private Sub sky2_DblClick()
  Dim lResult As Long
  
  nplachos& = sky2.MouseRow
  If nplachos& = 0 Then Exit Sub
  placdblclk = True
  jumpworld = False
  If world = False Then
     If InStr(sky2.FormatString, "ITMx") <> 0 Then
        skymove = True
        Skycoord% = 2
        skyx = sky2.TextArray(skyp2(nplachos&, 1))
        skyy = sky2.TextArray(skyp2(nplachos&, 2))
     ElseIf InStr(sky2.FormatString, "latitude") <> 0 Then
        jumpworld = True
        lat = sky2.TextArray(skyp2(nplachos&, 1))
        lon = sky2.TextArray(skyp2(nplachos&, 2))
        End If
    lResult = FindWindow(vbNullString, terranam$)
    If lResult > 0 And terranam$ <> "" Then
        dx1 = 500 'position cursor off screen above dbl click
        dy1 = -500
        Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
        End If
     End If
  If world = True Then
     If tblbuttons(3) = 1 Then
        lResult = FindWindow(vbNullString, "Overview")
        If lResult <> 0 Then
            TdxhWnd = 0
            bRtn = EnumWindows(AddressOf EnumWndProc, lParam)
            iposit% = InStr(Tdxname, "-  ")
            lat3d = Val(Mid$(Tdxname, iposit% + 4, 2)) + Val(Mid$(Tdxname, iposit% + 8, 4)) / 60
            lon3d = Val(Mid$(Tdxname, iposit% + 15, 3)) + Val(Mid$(Tdxname, iposit% + 19, 5)) / 60
            OverhWnd = FindWindow(vbNullString, "Overview")
'            ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            BringWindowToTop (OverhWnd)
'            ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            BringWindowToTop (mapPictureform.hwnd)
            'Call BringWindowToTop(OverhWnd)
            dx1 = -1000 '-30 '30
            dy1 = -1000 '-240 '60
            Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
            waitime = Timer + 0.01
            Do Until Timer > waitime
               DoEvents
            Loop
            lon3dnew = -Val(sky2.TextArray(skyp2(nplachos&, 2)))
            lat3dnew = Val(sky2.TextArray(skyp2(nplachos&, 1)))
            dx1 = -(lon3dnew - lon3d) * 516.6 + 96
            dy1 = -(lat3dnew - lat3d) * 516.6 + 156
            Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
            waitime = Timer + 0.01
            Do Until Timer > waitime
               DoEvents
            Loop
            Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'move mouse to Location item
            Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0) 'move mouse to Location item
'            ret = SetWindowPos(mapsearchfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            BringWindowToTop (mapsearchfm.hwnd)
            waitime = Timer + 2
            Do Until Timer > waitime
               DoEvents
            Loop
            Call BringWindowToTop(mapsearchfm.hwnd)
            Exit Sub
            End If
         End If
     worldmove = True
     lon = sky2.TextArray(skyp2(nplachos&, 2))
     lat = sky2.TextArray(skyp2(nplachos&, 1))
     'Call form_queryunload(i%, j%)
     Call goto_click
     jumpworld = False
     skymove = False
     worldmove = False
     Skycoord% = 0
     'Call BringWindowToTop(OverhWnd)
     'ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
     'ret = SetWindowPos(mapsearchfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
     Call BringWindowToTop(mapsearchfm.hwnd)
     Exit Sub
     End If
  'Call form_queryunload(i%, j%)
  Call goto_click
  jumpworld = False
  worldmove = False
  skymove = False
  Skycoord% = 0
  Call BringWindowToTop(mapsearchfm.hwnd)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RunOnAutomatic
' DateTime  : 09/01/2003 17:01
' Author    : Chaim Keller
' Purpose   : automatic processing of profiles
'---------------------------------------------------------------------------------------
'
Sub RunOnAutomatic()

  Dim batnum%, maplistnum%, doclin$, doclin2$, nummissing%
  Dim Checking As Boolean
  Dim SplitArray() As String
  Dim SplitBat() As String
  Dim LatCompare As Double, LonCompare As Double
  Dim LatBat As Double, LonBat As Double
  Dim found%, batnam$, tmpfile$, tmpnum%
  
   On Error GoTo RunOnAutomatic_Error

   AutoProf = True 'run profile analysis automatically
   AutoVer = True 'automatically increment the version number
   
'   Select Case MsgBox("Sunrise profiles?" _
'                      & vbCrLf & "" _
'                      & vbCrLf & "Answer:" _
'                      & vbCrLf & "   Yes -- for sunrise profiles" _
'                      & vbCrLf & "   No --- for sunset profiles" _
'                      & vbCrLf & "" _
'                      , vbYesNoCancel + vbQuestion + vbSystemModal + vbDefaultButton1, App.Title)
'
'    Case vbYes
'       sunmode% = 1
'    Case vbNo
'       sunmode% = 0
'    Case vbCancel
'       Exit Sub
'   End Select

   'New interface
'   ret = SetWindowPos(mapsearchfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (mapsearchfm.hwnd)
   
   frmMsgBox.MsgCstm "Click sunrise or sunset horizon profile buttons:", "Horizon calculations", mbInformation, 1, False, _
                     "Sunrise profiles", "Sunset profiles", "Cancel"
   Select Case frmMsgBox.g_lBtnClicked
       Case 1 'the 1st button in your list was clicked
          sunmode% = 1
       Case 2 'the 2nd button in your list was clicked
          sunmode% = 0
       Case 0, 3 'cancel.
          Exit Sub
       Case Else
          Exit Sub
   End Select
   
   starting& = 0
   If Dir(drivjk_c$ & "mapstatus.sav") <> sEmpty Then
      Close
      statfil% = FreeFile
      Open drivjk_c$ & "mapstatus.sav" For Input As #statfil%
      Do Until EOF(statfil%)
         Input #statfil%, statnum&
      Loop
      Close #statfil%
      
'      Select Case MsgBox("Start at point #:" & Str$(statnum&) _
'                         & vbCrLf & "" _
'                         & vbCrLf & "Answer:" _
'                         & vbCrLf & "    Yes -- to start at the above number" _
'                         & vbCrLf & "    No  -- to start from the beginning" _
'                         , vbYesNoCancel + vbQuestion + vbDefaultButton1, App.Title)
'
'        Case vbYes
'           starting& = statnum&
'        Case vbNo
'           starting& = 0
'        Case vbCancel
'           Exit Sub
'
'      End Select
      
      frmMsgBox.MsgCstm "Start at point #:" & Str$(statnum&) & " ?", "Start at?", mbQuestion, 2, False, _
                      "Yes -- start at above number", "Start from beginning", "Cancel"
      Select Case frmMsgBox.g_lBtnClicked
        Case 1 'the 1st button in your list was clicked
           starting& = statnum&
        Case 2 'the 2nd button in your list was clicked
           starting& = 0
        Case 0, 3 'cancel.
           Exit Sub
        Case Else
           Exit Sub
      End Select
      End If
      
uoa100:
      
   mapsearchfm.cmdCheckSkipped.Visible = False
   mapsearchfm.ProgressBarProfs.Visible = True
   
   If world = False Then GoTo W100
   
   'create eros.tm6 file
   dtmfile% = FreeFile
   Open ramdrive & ":\eros.tm6" For Output As #dtmfile%
   Select Case DTMflag
      Case 0, -1 'GTOPO30, SRTM30
         outdrive$ = worlddtm
      Case 1 'SRTM1
         outdrive$ = srtmdtm
      Case 2 'SRTM3
         outdrive$ = d3asdtm
      Case 3
         outdrive$ = alosdtm
   End Select
   Print #dtmfile%, outdrive$; ","; DTMflag
   Close #dtmfile%
   
   'create eros.tm7 flag -- to tell newreadDTM not to ask about missing tiles
   dtmfile% = FreeFile
   Open ramdrive & ":\eros.tm7" For Output As #dtmfile%
   Print #dtmfile%, 1
   Close #dtmfile%
   
   AutoNum& = starting&
50 savfil% = FreeFile
   Open drivjk_c$ & "mappoints.sav" For Input As #savfil%
   looping% = 0
   i& = 0
   Do Until EOF(savfil%)
      Input #savfil%, savlat, savlon, savhgt, savdis
      i& = i& + 1
      mapsearchfm.ProgressBarProfs.value = i&
      mapsearchfm.StatusBarProg.Panels(2).Text = Str(i&)
      statfil% = FreeFile 'record current status
      Open drivjk_c$ & "mapstatus.sav" For Output As #statfil%
      Write #statfil%, i&
      Close #statfil%
      If i& > AutoNum& Then
         'inputed all the desired coordinates, so exit loop
         looping% = 1

         Exit Do
         End If
   Loop
   Close #savfil%
   
   If looping% = 0 Then
      'reached EOF--i.e., finished the profiles
      AutoProf = False
      AutoVer = False
      Delay% = 0
      Kill ramdrive & ":\eros.tm7"
      
      mapsearchfm.ProgressBarProfs.Visible = False
      mapsearchfm.cmdCheckSkipped.Visible = True
      mapsearchfm.cmdCheckSkipped.Enabled = True
      
      Select Case MsgBox("Do you want to check for skipped points?", vbYesNo Or vbQuestion Or vbDefaultButton1, "Search for skipped points")
      
        Case vbYes
        
            'test that all the places have been done, and if not start new loop
            GoTo ComparePoints
      
        Case vbNo
      
      End Select
      
      Exit Sub
      End If
   
   If savlon = 0 And savlat = 0 Then GoTo 50
   lon = savlon
   lat = savlat
   worldmove = True
   Call goto_click
   jumpworld = False
   skymove = False
   worldmove = False
   Skycoord% = 0
   Call BringWindowToTop(mapsearchfm.hwnd)
   searchhgt = savhgt
   viewsearch = True
   Call sunrisesunset(sunmode%)
   viewsearch = False
   AutoNum& = AutoNum& + 1
   AutoVer = False 'don't increment the version number again
   GoTo 50
   
   On Error GoTo 0
   Exit Sub
   
W100: 'run analysis of Eretz Yisroel places
   'ask for a default name
   Unload mapsearchfm
   Set mapsearchfm = Nothing
   Call EYsunrisesunset(sunmode%)
   Exit Sub
   
   
ComparePoints:

    mapsearchfm.cmdCheckSkipped.Visible = False
    mapsearchfm.ProgressBarProfs.Visible = True

    Close 'close any open file
    If world Then
        AddPath$ = "eros\"
    Else
       AddPath$ = sEmpty
       End If
       
    'find name of active bat file
    nummissing% = 0
    If sunmode >= 1 Then 'sunrises
       batnam$ = Dir(drivcities$ & AddPath & Combo2.Text & "\netz\*.bat")
       batnam$ = drivcities$ & AddPath & Combo2.Text & "\netz\" & batnam$
       tmpfile$ = drivjk_c$ & "mappoints_tmp_netz.sav"
    ElseIf sunmode <= 0 Then 'sunsets
       batnam$ = Dir(drivcities$ & AddPath & Combo2.Text & "\skiy\*.bat")
       batnam$ = drivcities$ & AddPath & Combo2.Text & "\netz\" & batnam$
       tmpfile$ = drivjk_c$ & "mappoints_tmp_skiy.sav"
       End If

    maplistnum% = FreeFile
    Open drivjk_c$ + "mappoints.sav" For Input As #maplistnum%
    tmpnum% = FreeFile
    Open tmpfile$ For Output As #tmpnum%
    
    Do Until EOF(maplistnum%)
       Line Input #maplistnum%, doclin$
       SplitArray = Split(doclin$, ",")

       'this is a vantage point entry, so compare
      LatCompare = Val(SplitArray(0))
      LonCompare = Val(SplitArray(1))
      batnum% = FreeFile
      Open batnam$ For Input As #batnum%
      found% = 0
      Do Until EOF(batnum%)
         Line Input #batnum%, docliln2$
         SplitBat = Split(docliln2$, ",")
         If UBound(SplitBat) > 0 Then
              If InStr(SplitBat(0), "netz") Or InStr(SplitBat(0), "skiy") Then
                 'this is a vantage point entry
                 LatBat = Val(SplitBat(1))
                 LonBat = -Val(SplitBat(2))
                 If Abs(LatCompare - LatBat) < 0.0001 And Abs(LonCompare - LonBat) < 0.0001 Then
                    'accounted for, skip to next entry
                    found% = 1
                    Exit Do
                    End If
                 End If
             End If
      Loop
      
      If found% = 0 Then 'add to missing list
         Print #tmpnum%, doclin$
         nummissing% = nummissing% + 1
         End If
         
     Close #batnum%
    Loop
    Close #tmpnum%
    CloseErrors% = 1

    If nummissing% > 0 Then
       Close
       
       mapsearchfm.ProgressBarProfs.Max = nummissing%
       mapsearchfm.StatusBarProg.Panels(1).Text = nummissing%
       mapsearchfm.StatusBarProg.Panels(2).Text = "0"

       mapsearchfm.ProgressBarProfs.value = 0
       
       'reload list and restart analysis
       FileCopy drivjk_c$ + "mappoints.sav", drivjk_c$ + "mappoints_old.sav"
       Kill drivjk_c$ + "mappoints.sav"
       FileCopy tmpfile$, drivjk_c$ + "mappoints.sav"
       Checking = True
       starting& = 0
       GoTo uoa100
    Else 'restore mappoints.sav
       Kill drivjk_c$ + "mappoints.sav"
       FileCopy drivjk_c$ + "mappoints_old.sav", drivjk_c$ + "mappoints.sav"
       Kill drivjk_c$ + "mappoints_old.sav"
       mapsearchfm.ProgressBarProfs.Visible = False
       mapsearchfm.cmdCheckSkipped.Visible = True
       mapsearchfm.cmdCheckSkipped.Enabled = True
       End If
       
    Return

RunOnAutomatic_Error:
    Close
   
    mapsearchfm.cmdCheckSkipped.Enabled = True
    
    If CloseErrors% = 1 Then Resume Next

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RunOnAutomatic of Form mapsearchfm"

End Sub
