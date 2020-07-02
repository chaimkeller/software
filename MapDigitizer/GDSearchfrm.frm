VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GDSearchfrm 
   BorderStyle     =   0  'None
   Caption         =   "Search Wizard"
   ClientHeight    =   6765
   ClientLeft      =   660
   ClientTop       =   660
   ClientWidth     =   11280
   Icon            =   "GDSearchfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "__"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      TabIndex        =   164
      ToolTipText     =   "Minimize"
      Top             =   480
      Width           =   285
   End
   Begin MSComCtl2.Animation anmSearch 
      Height          =   435
      Left            =   10740
      TabIndex        =   135
      Top             =   1200
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   767
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   30
      FullHeight      =   29
   End
   Begin VB.PictureBox picAnimation 
      Height          =   555
      Left            =   10680
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   136
      Top             =   1140
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdSearchReport 
      Height          =   555
      Left            =   10710
      Picture         =   "GDSearchfrm.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Search"
      Top             =   3720
      Width           =   495
   End
   Begin TabDlg.SSTab tbSearch 
      Height          =   6645
      Left            =   0
      TabIndex        =   92
      Top             =   60
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   11721
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   794
      MouseIcon       =   "GDSearchfrm.frx":074C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Coordinates && Names"
      TabPicture(0)   =   "GDSearchfrm.frx":0768
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmCoordinates"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmNames"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Sample's Source"
      TabPicture(1)   =   "GDSearchfrm.frx":0BBA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Fossil Type"
      TabPicture(2)   =   "GDSearchfrm.frx":0D14
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Geo. Formation"
      TabPicture(3)   =   "GDSearchfrm.frx":0E6E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Geo. Age Dates"
      TabPicture(4)   =   "GDSearchfrm.frx":0FC8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Clients, Analysts, etc."
      TabPicture(5)   =   "GDSearchfrm.frx":141A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(1)=   "Frame13"
      Tab(5).ControlCount=   2
      Begin VB.Frame frmNames 
         Caption         =   "Names of Places/Wells to Ssearch"
         Height          =   1575
         Left            =   -74280
         TabIndex        =   170
         Top             =   4800
         Width           =   9255
         Begin VB.CheckBox chkCase 
            Caption         =   "Match Case"
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   184
            ToolTipText     =   "Check in order to find strings with the exact case "
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Frame frmDictionary 
            Caption         =   "Dictionary"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   5520
            TabIndex        =   177
            Top             =   120
            Width           =   3615
            Begin VB.CommandButton cmdPasteDictionary 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2040
               Picture         =   "GDSearchfrm.frx":186C
               Style           =   1  'Graphical
               TabIndex        =   183
               ToolTipText     =   "Paste the Dictionary Name to the selected textbox"
               Top             =   760
               Width           =   315
            End
            Begin VB.ComboBox cmbDictionary 
               Enabled         =   0   'False
               Height          =   315
               Left            =   840
               TabIndex        =   180
               ToolTipText     =   "Suggested names for searching"
               Top             =   360
               Width           =   2655
            End
            Begin VB.OptionButton optName2 
               Caption         =   "#2"
               Enabled         =   0   'False
               ForeColor       =   &H00004000&
               Height          =   195
               Left            =   120
               TabIndex        =   179
               ToolTipText     =   "Paste dictionary word into Name #2 textbox"
               Top             =   650
               Width           =   615
            End
            Begin VB.OptionButton optName1 
               Caption         =   "#1"
               Enabled         =   0   'False
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   178
               ToolTipText     =   "Paste dictionary word into Name #1 textbox"
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdName2 
            Caption         =   "Enable"
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   176
            ToolTipText     =   "Click to enable searches over String #2"
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdName1 
            Caption         =   "Enable"
            Height          =   255
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   175
            ToolTipText     =   "Click to enable searches over String #1"
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cmbName2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   720
            Width           =   1150
         End
         Begin VB.ComboBox cmbName1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   360
            Width           =   1150
         End
         Begin VB.TextBox txtName2 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2640
            TabIndex        =   172
            ToolTipText     =   "Search String #2"
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txtName1 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2640
            TabIndex        =   171
            ToolTipText     =   "Search String #1"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblName2 
            Caption         =   "#2"
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   182
            Top             =   720
            Width           =   255
         End
         Begin VB.Label lblName1 
            Caption         =   "#1"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   181
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H8000000A&
         Height          =   6075
         Left            =   -66960
         TabIndex        =   127
         Top             =   480
         Width           =   2535
         Begin VB.Frame Frame18 
            BackColor       =   &H8000000A&
            Caption         =   "Analysts"
            Height          =   5835
            Left            =   60
            TabIndex        =   128
            Top             =   120
            Width           =   2415
            Begin VB.ListBox lstAnalystUnsorted 
               Height          =   255
               Left            =   240
               MultiSelect     =   2  'Extended
               TabIndex        =   140
               Top             =   1300
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.ListBox lstAnalyst 
               Enabled         =   0   'False
               Height          =   4935
               Left            =   60
               MultiSelect     =   2  'Extended
               Sorted          =   -1  'True
               TabIndex        =   91
               Top             =   240
               Width           =   2295
            End
            Begin VB.CommandButton cmdAnalyst 
               Caption         =   "Enable"
               Height          =   375
               Left            =   420
               TabIndex        =   90
               Top             =   5320
               Width           =   1635
            End
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Height          =   6075
         Left            =   -74880
         TabIndex        =   102
         Top             =   480
         Width           =   7875
         Begin VB.Frame Frame17 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Company/Division"
            Height          =   5835
            Left            =   120
            TabIndex        =   121
            Top             =   120
            Width           =   2535
            Begin VB.ListBox lstCompanyUnsorted 
               Height          =   255
               Left            =   180
               MultiSelect     =   2  'Extended
               TabIndex        =   138
               Top             =   1300
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.CommandButton cmdCompany 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Enable"
               Height          =   375
               Left            =   420
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   5320
               Width           =   1695
            End
            Begin VB.ListBox lstCompany 
               Enabled         =   0   'False
               Height          =   4935
               Left            =   60
               MultiSelect     =   2  'Extended
               Sorted          =   -1  'True
               TabIndex        =   83
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Date"
            Height          =   5835
            Left            =   2700
            TabIndex        =   120
            Top             =   120
            Width           =   1395
            Begin VB.CommandButton cmdDate 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Enable"
               Height          =   375
               Left            =   95
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   5320
               Width           =   1215
            End
            Begin VB.ListBox lstDate 
               Enabled         =   0   'False
               Height          =   4935
               Left            =   60
               MultiSelect     =   2  'Extended
               TabIndex        =   85
               Top             =   240
               Width           =   1275
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Project No."
            Height          =   5835
            Left            =   4140
            TabIndex        =   119
            Top             =   120
            Width           =   1335
            Begin VB.ListBox lstFormProject 
               Height          =   255
               Left            =   120
               TabIndex        =   129
               Top             =   1320
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton cmdProject 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Enable"
               Height          =   375
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   5320
               Width           =   1080
            End
            Begin VB.ListBox lstProject 
               Enabled         =   0   'False
               Height          =   4935
               Left            =   60
               MultiSelect     =   2  'Extended
               TabIndex        =   87
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Clients"
            Height          =   5835
            Left            =   5520
            TabIndex        =   118
            Top             =   120
            Width           =   2295
            Begin VB.ListBox lstClientUnsorted 
               Height          =   255
               Left            =   240
               MultiSelect     =   2  'Extended
               TabIndex        =   139
               Top             =   1300
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CommandButton cmdClient 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Enable"
               Height          =   375
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   88
               Top             =   5320
               Width           =   1575
            End
            Begin VB.ListBox lstClient 
               Enabled         =   0   'False
               Height          =   4935
               Left            =   120
               MultiSelect     =   2  'Extended
               Sorted          =   -1  'True
               TabIndex        =   89
               Top             =   240
               Width           =   2055
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   101
         Top             =   480
         Width           =   10395
         Begin VB.ListBox lstDates 
            Height          =   255
            Left            =   1980
            TabIndex        =   111
            Top             =   2880
            Visible         =   0   'False
            Width           =   2895
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5145
            Left            =   1200
            TabIndex        =   71
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   9075
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame Frame12 
            Height          =   555
            Left            =   1200
            TabIndex        =   110
            Top             =   120
            Width           =   8055
            Begin VB.CommandButton cmdEnabDate 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Press here to enable searches over a range of geologic age dates"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   540
               Style           =   1  'Graphical
               TabIndex        =   70
               Top             =   140
               Width           =   7035
            End
         End
         Begin VB.Frame Frame11 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4650
            Left            =   5760
            TabIndex        =   109
            Top             =   900
            Width           =   3495
            Begin VB.ComboBox cmbLQuest 
               Enabled         =   0   'False
               ForeColor       =   &H000040C0&
               Height          =   315
               Left            =   2820
               TabIndex        =   75
               Top             =   1080
               Width           =   570
            End
            Begin VB.ComboBox cmbLPre 
               Enabled         =   0   'False
               ForeColor       =   &H000040C0&
               Height          =   315
               Left            =   120
               TabIndex        =   73
               Top             =   1080
               Width           =   915
            End
            Begin VB.ComboBox cmbEQuest 
               Enabled         =   0   'False
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   2820
               TabIndex        =   79
               Text            =   " "
               Top             =   2760
               Width           =   570
            End
            Begin VB.ComboBox cmbEPre 
               Enabled         =   0   'False
               ForeColor       =   &H00C00000&
               Height          =   315
               Left            =   120
               TabIndex        =   77
               Top             =   2760
               Width           =   855
            End
            Begin VB.OptionButton optExact 
               Caption         =   "&Search for the exact Dates"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   80
               ToolTipText     =   "Search for the defined dates only"
               Top             =   3660
               Width           =   3135
            End
            Begin VB.OptionButton optRange 
               Caption         =   "&Search for range of Dates"
               Enabled         =   0   'False
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
               Left            =   240
               TabIndex        =   81
               ToolTipText     =   "Search for a range of dates"
               Top             =   3900
               Value           =   -1  'True
               Width           =   3075
            End
            Begin VB.TextBox txtLater 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   345
               Left            =   1080
               TabIndex        =   74
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton cmdLater 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Press to accept as the &Later Date"
               Enabled         =   0   'False
               Height          =   375
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   72
               ToolTipText     =   "Press to ""fix"" selection of later age date"
               Top             =   360
               Width           =   2775
            End
            Begin VB.TextBox txtEarlier 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   1020
               TabIndex        =   78
               Top             =   2760
               Width           =   1755
            End
            Begin VB.CommandButton cmdEarlier 
               BackColor       =   &H00FFFF80&
               Caption         =   "Press to accept as the &Earlier Date"
               Enabled         =   0   'False
               Height          =   375
               Left            =   420
               Style           =   1  'Graphical
               TabIndex        =   76
               ToolTipText     =   "Press to ""fix"" selection for earlier age date"
               Top             =   2040
               Width           =   2775
            End
            Begin VB.Label lblEQuest 
               Alignment       =   2  'Center
               Caption         =   "?"
               Height          =   195
               Left            =   2880
               TabIndex        =   117
               Top             =   2520
               Width           =   255
            End
            Begin VB.Label lblLater 
               Alignment       =   2  'Center
               Caption         =   "Later Age Date"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1140
               TabIndex        =   116
               Top             =   840
               Width           =   1635
            End
            Begin VB.Label lblEPre 
               Alignment       =   2  'Center
               Caption         =   "Pre-Age"
               Enabled         =   0   'False
               Height          =   195
               Left            =   120
               TabIndex        =   115
               Top             =   2520
               Width           =   855
            End
            Begin VB.Label lblLQuest 
               Alignment       =   2  'Center
               Caption         =   "?"
               Enabled         =   0   'False
               Height          =   195
               Left            =   2880
               TabIndex        =   114
               Top             =   840
               Width           =   255
            End
            Begin VB.Label lblEarlier 
               Alignment       =   2  'Center
               Caption         =   "Earlier Age Date"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1080
               TabIndex        =   113
               Top             =   2520
               Width           =   1635
            End
            Begin VB.Label lblLPre 
               Alignment       =   2  'Center
               Caption         =   "Pre-Age"
               Enabled         =   0   'False
               Height          =   195
               Left            =   120
               TabIndex        =   112
               Top             =   840
               Width           =   855
            End
            Begin VB.Line Line10 
               X1              =   180
               X2              =   3420
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Line Line7 
               X1              =   120
               X2              =   3360
               Y1              =   1740
               Y2              =   1740
            End
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   -74880
         TabIndex        =   100
         Top             =   460
         Width           =   10335
         Begin VB.Frame frmFormation 
            Height          =   5115
            Left            =   840
            TabIndex        =   106
            Top             =   780
            Width           =   8715
            Begin VB.ListBox lstFormationUnsorted 
               Height          =   255
               Left            =   2400
               MultiSelect     =   2  'Extended
               TabIndex        =   137
               Top             =   1440
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.Frame Frame10 
               Height          =   3555
               Left            =   5640
               TabIndex        =   108
               Top             =   780
               Width           =   1455
               Begin VB.CommandButton cmdClearForm 
                  Caption         =   "&Clear All"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   177
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   210
                  Style           =   1  'Graphical
                  TabIndex        =   67
                  ToolTipText     =   "Clear all choices"
                  Top             =   1900
                  Width           =   1035
               End
               Begin VB.CommandButton cmdAllForm 
                  Caption         =   "Select &All"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   177
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   210
                  Style           =   1  'Graphical
                  TabIndex        =   66
                  ToolTipText     =   "Search records with a non-empty formation field"
                  Top             =   1000
                  Width           =   1035
               End
            End
            Begin VB.ListBox lstFormation 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4380
               Left            =   1980
               MultiSelect     =   2  'Extended
               Sorted          =   -1  'True
               TabIndex        =   68
               Top             =   540
               Width           =   3495
            End
            Begin VB.Label lblFormation 
               Alignment       =   2  'Center
               Caption         =   "Searches will be conducted over any formation"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   540
               TabIndex        =   107
               Top             =   180
               Width           =   7215
            End
         End
         Begin VB.CommandButton cmdActivateForm 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Press to enable searches over specific formations"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   240
            Width           =   6915
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6015
         Left            =   120
         TabIndex        =   99
         Top             =   480
         Width           =   10395
         Begin VB.CommandButton cmdEditdb2 
            Caption         =   "&Edit"
            Height          =   220
            Left            =   9780
            TabIndex        =   168
            ToolTipText     =   "Edit Sql clause for fossil categories of scanned database"
            Top             =   5740
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdEditdb1 
            Caption         =   "&Edit"
            Height          =   220
            Left            =   9780
            TabIndex        =   167
            ToolTipText     =   "Edit Sql clause for fossil categories of active database"
            Top             =   5500
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtSQLdb2 
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
            Height          =   240
            Left            =   2820
            TabIndex        =   166
            Text            =   "txtSQLdb2"
            ToolTipText     =   "Scanned database's Fossil Type SQL clause"
            Top             =   5725
            Visible         =   0   'False
            Width           =   6915
         End
         Begin VB.TextBox txtSQLdb1 
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
            Height          =   240
            Left            =   2820
            TabIndex        =   165
            Text            =   "txtSQLdb1"
            ToolTipText     =   "Active database's Fossil Type SQL clause"
            Top             =   5490
            Visible         =   0   'False
            Width           =   6915
         End
         Begin VB.CheckBox chkShekef 
            Caption         =   "&Shekef"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1740
            TabIndex        =   163
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   5620
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox Combo15 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   161
            ToolTipText     =   "Boolean search operators (Shekef)"
            Top             =   5580
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox Combo14 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8280
            TabIndex        =   56
            Text            =   "Combo14"
            ToolTipText     =   "Boolean search operators (palynology fossil names)"
            Top             =   4920
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Combo13 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8280
            TabIndex        =   51
            Text            =   "Combo13"
            ToolTipText     =   "Boolean search operators (ostracoda fossil names)"
            Top             =   4140
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Combo12 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8280
            TabIndex        =   46
            Text            =   "Combo12"
            ToolTipText     =   "Boolean search operators (nannoplankton fossil names)"
            Top             =   3420
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Combo11 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8280
            TabIndex        =   41
            Text            =   "Combo11"
            ToolTipText     =   "Boolean search operators (megafauna fossil names)"
            Top             =   2640
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Combo10 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8280
            TabIndex        =   36
            Text            =   "Combo10"
            ToolTipText     =   "Boolean search operators (foraminifera fossil names)"
            Top             =   1860
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Combo9 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8280
            TabIndex        =   31
            Text            =   "Combo9"
            ToolTipText     =   "Boolean search operators (diatom fossil names)"
            Top             =   1080
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Combo8 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8280
            TabIndex        =   26
            Text            =   "Combo8"
            ToolTipText     =   "Boolean search operators (conodonta fossil names)"
            Top             =   360
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox Combo1 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   58
            ToolTipText     =   "Boolean search operators (Conodonta)"
            Top             =   300
            Width           =   1095
         End
         Begin VB.ListBox lstDiatomsdic 
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   900
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstPalyndic 
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   4740
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstOstracoddic 
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   3960
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstNanodic 
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   3240
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstMegadic 
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   2460
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstForamsdic 
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   1680
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstDiatom 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   57
            Top             =   900
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstConodsdic 
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.CheckBox chkDicPaly 
            Caption         =   "Species"
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3300
            TabIndex        =   54
            Top             =   5100
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CheckBox chkDicOstra 
            Caption         =   "Species"
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3300
            TabIndex        =   49
            Top             =   4320
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CheckBox chkDicNano 
            Caption         =   "Species"
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3300
            TabIndex        =   44
            Top             =   3600
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CheckBox chkDicMega 
            Caption         =   "Species"
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3300
            TabIndex        =   39
            Top             =   2820
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CheckBox chkDicForam 
            Caption         =   "Species"
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3300
            TabIndex        =   34
            Top             =   2040
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CheckBox chkDicDiatom 
            Caption         =   "Species"
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3300
            TabIndex        =   29
            Top             =   1260
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.CheckBox chkDicCono 
            Caption         =   "Species"
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3300
            TabIndex        =   24
            Top             =   540
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ListBox lstPalynology 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   126
            Top             =   4740
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstOstracoda 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   125
            Top             =   3960
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstNannoplankton 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   124
            Top             =   3240
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstMegafauna 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   123
            Top             =   2460
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstForaminifera 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   122
            Top             =   1680
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ListBox lstConodonta 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   645
            Left            =   4080
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   120
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.CheckBox chkActPaly 
            Caption         =   "Zones"
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
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   3300
            TabIndex        =   53
            Top             =   4800
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CheckBox chkActOstra 
            Caption         =   "Zones"
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
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   3300
            TabIndex        =   48
            Top             =   4020
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CheckBox chkActNano 
            Caption         =   "Zones"
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
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   3300
            TabIndex        =   43
            Top             =   3300
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CheckBox chkActMega 
            Caption         =   "Zones"
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
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   3300
            TabIndex        =   38
            Top             =   2520
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CheckBox chkActForam 
            Caption         =   "Zones"
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
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   3300
            TabIndex        =   33
            Top             =   1740
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CheckBox chkActDiatom 
            Caption         =   "Zones"
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
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   3300
            TabIndex        =   28
            Top             =   960
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CheckBox chkActCono 
            Caption         =   "Zones"
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
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   3300
            TabIndex        =   23
            Top             =   240
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "&Clear All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9210
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Clear all choices"
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton cmdAll 
            Caption         =   "&Pick All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9210
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Search all the fossil categories"
            Top             =   960
            Width           =   975
         End
         Begin VB.Frame Frame9 
            Caption         =   "Fossil Types"
            ForeColor       =   &H00FF0000&
            Height          =   5355
            Left            =   9120
            TabIndex        =   105
            Top             =   120
            Width           =   1155
            Begin VB.CommandButton cmdEditSql 
               Caption         =   "Show Fossil Types &SQL"
               Height          =   675
               Left            =   90
               Style           =   1  'Graphical
               TabIndex        =   169
               ToolTipText     =   "Show the Fossil Type SQL clause(s)"
               Top             =   4200
               Width           =   975
            End
            Begin VB.CommandButton cmdAllOr 
               Caption         =   "All ""&OR"""
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   90
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   """OR"" search over all fossil categories"
               Top             =   3060
               Width           =   975
            End
            Begin VB.CommandButton cmdAllAnd 
               Caption         =   "All ""&AND"""
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   90
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   """And"" search over all fossil categories"
               Top             =   2460
               Width           =   975
            End
         End
         Begin VB.ComboBox Combo7 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   64
            ToolTipText     =   "Boolean search operators (Palynology)"
            Top             =   4920
            Width           =   1095
         End
         Begin VB.ComboBox Combo6 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   63
            ToolTipText     =   "Boolean search operators (Ostracoda)"
            Top             =   4140
            Width           =   1095
         End
         Begin VB.ComboBox Combo5 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   62
            ToolTipText     =   "Boolean search operators (Nannoplankton)"
            Top             =   3360
            Width           =   1095
         End
         Begin VB.ComboBox Combo4 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   61
            ToolTipText     =   "Boolean search operators (Megafauna)"
            Top             =   2640
            Width           =   1095
         End
         Begin VB.ComboBox Combo3 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   60
            ToolTipText     =   "Boolean search operators (Foraminifera)"
            Top             =   1860
            Width           =   1095
         End
         Begin VB.ComboBox Combo2 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   59
            ToolTipText     =   "Boolean search operators (Daitom)"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chkPalynology 
            Caption         =   "&Palynology"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   52
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   4920
            Width           =   1335
         End
         Begin VB.CheckBox chkOstracoda 
            Caption         =   "&Ostracoda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   47
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   4140
            Width           =   1455
         End
         Begin VB.CheckBox chkNanoplankton 
            Caption         =   "&Nannoplankton"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   42
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   3420
            Width           =   1575
         End
         Begin VB.CheckBox chkMegafauna 
            Caption         =   "&Megafauna"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   37
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CheckBox chkForaminifera 
            Caption         =   "&Foraminifera"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1740
            TabIndex        =   32
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   1860
            Width           =   1395
         End
         Begin VB.CheckBox chkDiatom 
            Caption         =   "&Diatom"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1740
            TabIndex        =   27
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   1080
            Width           =   1155
         End
         Begin VB.CheckBox chkConodonta 
            Caption         =   "&Conodonta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1680
            TabIndex        =   22
            ToolTipText     =   "Check the box to search for this fossil"
            Top             =   300
            Width           =   1275
         End
         Begin VB.Line lineSQL 
            Visible         =   0   'False
            X1              =   2760
            X2              =   2760
            Y1              =   5460
            Y2              =   6000
         End
         Begin VB.Image imShekef 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":1D9E
            Stretch         =   -1  'True
            Top             =   5580
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblShekef 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   162
            Top             =   5640
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Line Line18 
            X1              =   0
            X2              =   9060
            Y1              =   5460
            Y2              =   5460
         End
         Begin VB.Label Label19 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   154
            Top             =   4980
            Width           =   195
         End
         Begin VB.Label Label18 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   153
            Top             =   4200
            Width           =   195
         End
         Begin VB.Label Label17 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   152
            Top             =   3420
            Width           =   195
         End
         Begin VB.Label Label16 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   151
            Top             =   2700
            Width           =   195
         End
         Begin VB.Label Label15 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   150
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label Label14 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   149
            Top             =   1140
            Width           =   195
         End
         Begin VB.Label Label13 
            Caption         =   "-->"
            Height          =   195
            Left            =   1140
            TabIndex        =   148
            Top             =   360
            Width           =   195
         End
         Begin VB.Label lblFosPaly 
            Caption         =   "<--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8100
            TabIndex        =   147
            Top             =   4980
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblFosOstra 
            Caption         =   "<--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8100
            TabIndex        =   146
            Top             =   4200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblFosNano 
            Caption         =   "<--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8100
            TabIndex        =   145
            Top             =   3480
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblFosMega 
            Caption         =   "<--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8100
            TabIndex        =   144
            Top             =   2700
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblFosForam 
            Caption         =   "<--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8100
            TabIndex        =   143
            Top             =   1920
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblFosDiatom 
            Caption         =   "<--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8100
            TabIndex        =   142
            Top             =   1140
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblFosCono 
            Caption         =   "<--"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8100
            TabIndex        =   141
            Top             =   420
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Image Image9 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":1EE8
            Stretch         =   -1  'True
            Top             =   4920
            Width           =   315
         End
         Begin VB.Image Image8 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":2032
            Stretch         =   -1  'True
            Top             =   4140
            Width           =   315
         End
         Begin VB.Image Image7 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":217C
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   315
         End
         Begin VB.Image Image6 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":22C6
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   315
         End
         Begin VB.Image Image5 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":2410
            Stretch         =   -1  'True
            Top             =   1860
            Width           =   315
         End
         Begin VB.Image Image4 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":255A
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   315
         End
         Begin VB.Image Image3 
            Height          =   315
            Left            =   1380
            Picture         =   "GDSearchfrm.frx":26A4
            Stretch         =   -1  'True
            Top             =   300
            Width           =   315
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   9060
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   9060
            Y1              =   3900
            Y2              =   3900
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   9120
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   9120
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   9120
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   9120
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5895
         Left            =   -74820
         TabIndex        =   98
         Top             =   600
         Width           =   10275
         Begin VB.CommandButton cmdAllSources 
            Caption         =   "&Search All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6420
            Picture         =   "GDSearchfrm.frx":27EE
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Search records from all sources"
            Top             =   4620
            Width           =   2175
         End
         Begin VB.CommandButton cmdOutcroppings 
            Caption         =   "Search &Outcroppings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3840
            Picture         =   "GDSearchfrm.frx":2938
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Search records from outcroppings"
            Top             =   4620
            Width           =   2595
         End
         Begin VB.CommandButton cmdWells 
            Caption         =   "Search &Wells"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1560
            Picture         =   "GDSearchfrm.frx":2A82
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Search records from  wells"
            Top             =   4620
            Width           =   2295
         End
         Begin VB.Frame Frame8 
            Caption         =   "Outcropings"
            Height          =   1995
            Left            =   120
            TabIndex        =   104
            Top             =   2220
            Width           =   9975
            Begin VB.CheckBox chkOutcroppings 
               Caption         =   "Search samples from &outcroppings"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   3060
               TabIndex        =   17
               Top             =   900
               Width           =   4575
            End
            Begin VB.Image Image2 
               Height          =   615
               Left            =   1740
               Picture         =   "GDSearchfrm.frx":2BCC
               Stretch         =   -1  'True
               Top             =   780
               Width           =   675
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Wells"
            Height          =   1995
            Left            =   120
            TabIndex        =   103
            Top             =   180
            Width           =   9975
            Begin MSComCtl2.UpDown udLimdo 
               Height          =   315
               Left            =   7140
               TabIndex        =   158
               Top             =   1260
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtLimdo"
               BuddyDispid     =   196781
               OrigLeft        =   8880
               OrigTop         =   1140
               OrigRight       =   9120
               OrigBottom      =   1455
               Increment       =   10
               Max             =   99999
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtLimdo 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   6480
               TabIndex        =   157
               Text            =   "0"
               ToolTipText     =   "meters"
               Top             =   1260
               Width           =   675
            End
            Begin MSComCtl2.UpDown udLimup 
               Height          =   315
               Left            =   7140
               TabIndex        =   156
               Top             =   840
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtLimup"
               BuddyDispid     =   196782
               OrigLeft        =   8880
               OrigTop         =   540
               OrigRight       =   9120
               OrigBottom      =   855
               Increment       =   10
               Max             =   99999
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtLimup 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   6480
               TabIndex        =   155
               Text            =   "0"
               ToolTipText     =   "meters"
               Top             =   840
               Width           =   675
            End
            Begin VB.CheckBox chkJustCores 
               Caption         =   "Just &Cores"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3000
               TabIndex        =   15
               Top             =   1380
               Width           =   1815
            End
            Begin VB.CheckBox chkJustCuttings 
               Caption         =   "&Just Cuttings"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3000
               TabIndex        =   14
               Top             =   1020
               Width           =   1875
            End
            Begin VB.CheckBox chkAllWells 
               Caption         =   "&All wells"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3000
               TabIndex        =   16
               Top             =   770
               Width           =   1635
            End
            Begin VB.CheckBox chkWells 
               Caption         =   "Search samples from &wells"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3420
               TabIndex        =   13
               Top             =   300
               Width           =   3675
            End
            Begin VB.Label lblLimdo 
               Caption         =   "Maximum Depth:"
               Height          =   195
               Left            =   5220
               TabIndex        =   160
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lblLimup 
               Caption         =   "Minimum Depth:"
               Height          =   195
               Left            =   5220
               TabIndex        =   159
               Top             =   900
               Width           =   1215
            End
            Begin VB.Image Image1 
               Height          =   555
               Left            =   1740
               Picture         =   "GDSearchfrm.frx":2D16
               Stretch         =   -1  'True
               Top             =   900
               Width           =   615
            End
         End
      End
      Begin VB.Frame frmCoordinates 
         Caption         =   "Coordinate Boundaries for Search"
         Height          =   4035
         Left            =   -74280
         TabIndex        =   93
         Top             =   600
         Width           =   9195
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Choose over any coordinates"
            Top             =   3000
            Width           =   915
         End
         Begin VB.CommandButton cmdFullExtent 
            Caption         =   "&Entire Map"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5100
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Search over entire map"
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox txtNorthMax 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   5
            Text            =   "0.0"
            Top             =   1500
            Width           =   1935
         End
         Begin VB.TextBox txtNorthMin 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   7
            Text            =   "0.0"
            Top             =   2460
            Width           =   1935
         End
         Begin VB.TextBox txtEastMax 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   6
            Text            =   "0.0"
            Top             =   1980
            Width           =   1935
         End
         Begin VB.TextBox txtEastMin 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   4
            Text            =   "0.0"
            Top             =   1020
            Width           =   1935
         End
         Begin VB.Label lblYMin 
            Caption         =   "ITMy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7260
            TabIndex        =   134
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label lblYMax 
            Caption         =   "ITMy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7260
            TabIndex        =   133
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label lblXMax 
            Caption         =   "ITMx"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7260
            TabIndex        =   132
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label lblXMin 
            Caption         =   "ITMx"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7260
            TabIndex        =   131
            Top             =   1080
            Width           =   615
         End
         Begin VB.Line Line17 
            BorderColor     =   &H0000FFFF&
            X1              =   3180
            X2              =   3120
            Y1              =   2340
            Y2              =   2400
         End
         Begin VB.Line Line11 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   3480
            X2              =   3120
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line9 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   2340
            X2              =   2400
            Y1              =   1440
            Y2              =   1380
         End
         Begin VB.Line Line8 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   2340
            X2              =   3420
            Y1              =   1440
            Y2              =   1260
         End
         Begin VB.Image Image10 
            Height          =   2775
            Left            =   1740
            Picture         =   "GDSearchfrm.frx":2E60
            Stretch         =   -1  'True
            Top             =   780
            Width           =   1515
         End
         Begin VB.Line Line20 
            X1              =   3600
            X2              =   3720
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Line Line16 
            X1              =   3600
            X2              =   3720
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Line Line15 
            X1              =   3600
            X2              =   3600
            Y1              =   1020
            Y2              =   1860
         End
         Begin VB.Line Line14 
            X1              =   3600
            X2              =   3720
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line13 
            X1              =   3600
            X2              =   3720
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Line Line12 
            X1              =   3600
            X2              =   3600
            Y1              =   1980
            Y2              =   2820
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Coordinate boundaries to search"
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
            Left            =   4755
            TabIndex        =   130
            Top             =   660
            Width           =   2805
         End
         Begin VB.Label Label4 
            Caption         =   "North Maximum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   97
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "North Mininum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   96
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "East Maximum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   95
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "East Mininum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   94
            Top             =   1080
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10800
      Picture         =   "GDSearchfrm.frx":10516
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   180
      Width           =   285
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Next"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Back"
      Top             =   2460
      Width           =   495
   End
End
Attribute VB_Name = "GDSearchfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PasteName1 As Boolean 'used for pasting suggested names to name search text boxes
Dim PasteName2 As Boolean
Private Sub chkActCono_Click()
   If chkActCono.value = vbChecked And GDSearchfrm.chkConodonta.value = vbChecked Then
      GDSearchfrm.lstConodonta.Enabled = True
      GDSearchfrm.lstConodonta.Selected(0) = True
      GDSearchfrm.lstConodonta.Visible = True
      GDSearchfrm.lstConodsdic.Visible = False
      GDSearchfrm.chkDicCono.value = vbUnchecked
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstConodonta.Enabled = False
      For i& = 1 To GDSearchfrm.lstConodonta.ListCount
         GDSearchfrm.lstConodonta.Selected(i& - 1) = False
      Next i&
      chkActCono.value = vbUnchecked
      GDSearchfrm.lstConodonta.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkActForam_Click()
   If chkActForam.value = vbChecked And GDSearchfrm.chkForaminifera.value = vbChecked Then
      GDSearchfrm.lstForaminifera.Enabled = True
      GDSearchfrm.lstForaminifera.Selected(0) = True
      GDSearchfrm.lstForaminifera.Visible = True
      GDSearchfrm.lstForamsdic.Visible = False
      GDSearchfrm.chkDicForam.value = vbUnchecked
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstForaminifera.Enabled = False
      For i& = 1 To GDSearchfrm.lstForaminifera.ListCount
         GDSearchfrm.lstForaminifera.Selected(i& - 1) = False
      Next i&
      chkActForam.value = vbUnchecked
      GDSearchfrm.lstForaminifera.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkActMega_Click()
   If chkActMega.value = vbChecked And GDSearchfrm.chkMegafauna.value = vbChecked Then
      GDSearchfrm.lstMegafauna.Enabled = True
      GDSearchfrm.lstMegafauna.Selected(0) = True
      GDSearchfrm.lstMegafauna.Visible = True
      GDSearchfrm.lstMegadic.Visible = False
      GDSearchfrm.chkDicMega.value = vbUnchecked
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstMegafauna.Enabled = False
      For i& = 1 To GDSearchfrm.lstMegafauna.ListCount
         GDSearchfrm.lstMegafauna.Selected(i& - 1) = False
      Next i&
      GDSearchfrm.lstMegafauna.Visible = False
      chkActMega.value = vbUnchecked
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkActNano_Click()
   If chkActNano.value = vbChecked And GDSearchfrm.chkNanoplankton.value = vbChecked Then
      GDSearchfrm.lstNannoplankton.Enabled = True
      GDSearchfrm.lstNannoplankton.Selected(0) = True
      GDSearchfrm.lstNannoplankton.Visible = True
      GDSearchfrm.lstNanodic.Visible = False
      GDSearchfrm.chkDicNano.value = vbUnchecked
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstNannoplankton.Enabled = False
      For i& = 1 To GDSearchfrm.lstNannoplankton.ListCount
         GDSearchfrm.lstNannoplankton.Selected(i& - 1) = False
      Next i&
      chkActNano.value = vbUnchecked
      GDSearchfrm.lstNannoplankton.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkActOstra_Click()
   If chkActOstra.value = vbChecked And GDSearchfrm.chkOstracoda.value = vbChecked Then
      GDSearchfrm.lstOstracoda.Enabled = True
      GDSearchfrm.lstOstracoda.Selected(0) = True
      GDSearchfrm.lstOstracoda.Visible = True
      GDSearchfrm.lstOstracoddic.Visible = False
      GDSearchfrm.chkDicOstra.value = vbUnchecked
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstOstracoda.Enabled = False
      For i& = 1 To GDSearchfrm.lstOstracoda.ListCount
         GDSearchfrm.lstOstracoda.Selected(i& - 1) = False
      Next i&
      chkActOstra.value = vbUnchecked
      GDSearchfrm.lstOstracoda.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkActPaly_Click()
   If chkActPaly.value = vbChecked And GDSearchfrm.chkPalynology.value = vbChecked Then
      GDSearchfrm.lstPalynology.Enabled = True
      GDSearchfrm.lstPalynology.Selected(0) = True
      GDSearchfrm.lstPalynology.Visible = True
      GDSearchfrm.lstPalyndic.Visible = False
      GDSearchfrm.chkDicPaly.value = vbUnchecked
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstPalynology.Enabled = False
      For i& = 1 To GDSearchfrm.lstPalynology.ListCount
         GDSearchfrm.lstPalynology.Selected(i& - 1) = False
      Next i&
      chkActPaly.value = vbUnchecked
      GDSearchfrm.lstPalynology.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkAllWells_Click()
   chkJustCores.value = vbUnchecked
   chkJustCuttings.value = vbUnchecked
End Sub

Private Sub chkActDiatom_Click()
   If chkActDiatom.value = vbChecked And GDSearchfrm.chkDiatom.value = vbChecked Then
      GDSearchfrm.lstDiatom.Enabled = True
      GDSearchfrm.lstDiatom.Selected(0) = True
      GDSearchfrm.lstDiatom.Visible = True
      GDSearchfrm.lstDiatomsdic.Visible = False
      GDSearchfrm.chkDicDiatom.value = vbUnchecked
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstDiatom.Enabled = False
      For i& = 1 To GDSearchfrm.lstDiatom.ListCount
         GDSearchfrm.lstDiatom.Selected(i& - 1) = False
      Next i&
      chkActDiatom.value = vbUnchecked
      GDSearchfrm.lstDiatom.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkConodonta_Click()
   If chkConodonta.value = vbUnchecked Then
   'disenable searches over zones and fossils
      chkActCono_Click
      chkDicCono_Click
      chkActCono.Visible = False
      chkDicCono.Visible = False
      chkActCono.Visible = False
      chkDicCono.Visible = False
   Else
      chkActCono.Visible = True
      chkDicCono.Visible = True
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDiatom_Click()
   If chkDiatom.value = vbUnchecked Then
   'disenable searches over zones and fossils
      chkActDiatom_Click
      chkDicDiatom_Click
      chkActDiatom.Visible = False
      chkDicDiatom.Visible = False
      lstDiatom.Visible = False
      lstDiatomsdic.Visible = False
   Else
      chkActDiatom.Visible = True
      chkDicDiatom.Visible = True
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDicCono_Click()
   If chkDicCono.value = vbChecked And GDSearchfrm.chkConodonta.value = vbChecked Then
      GDSearchfrm.lstConodonta.Visible = False
      GDSearchfrm.lstConodonta.Enabled = False
      GDSearchfrm.lstConodonta.Selected(0) = False
      GDSearchfrm.lstConodsdic.Visible = True
      GDSearchfrm.lstConodsdic.Enabled = True
      GDSearchfrm.lstConodsdic.Selected(0) = True
      GDSearchfrm.chkActCono.value = vbUnchecked
      Combo8.Visible = True
      lblFosCono.Visible = True
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstConodsdic.Enabled = False
      For i& = 1 To GDSearchfrm.lstConodsdic.ListCount
         GDSearchfrm.lstConodsdic.Selected(i& - 1) = False
      Next i&
      GDSearchfrm.chkDicCono.value = vbUnchecked
      GDSearchfrm.lstConodsdic.Visible = False
      Combo8.Visible = False
      lblFosCono.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDicDiatom_Click()
   If chkDicDiatom.value = vbChecked And GDSearchfrm.chkDiatom.value = vbChecked Then
      GDSearchfrm.lstDiatom.Visible = False
      GDSearchfrm.lstDiatom.Enabled = False
      GDSearchfrm.lstDiatom.Selected(0) = False
      GDSearchfrm.lstDiatomsdic.Visible = True
      GDSearchfrm.lstDiatomsdic.Enabled = True
      GDSearchfrm.lstDiatomsdic.Selected(0) = True
      GDSearchfrm.chkActDiatom.value = vbUnchecked
      Combo9.Visible = True
      lblFosDiatom.Visible = True
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstDiatomsdic.Enabled = False
      For i& = 1 To GDSearchfrm.lstDiatomsdic.ListCount
         GDSearchfrm.lstDiatomsdic.Selected(i& - 1) = False
      Next i&
      GDSearchfrm.chkDicDiatom.value = vbUnchecked
      GDSearchfrm.lstDiatomsdic.Visible = False
      Combo9.Visible = False
      lblFosDiatom.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDicForam_Click()
   If chkDicForam.value = vbChecked And GDSearchfrm.chkForaminifera.value = vbChecked Then
      GDSearchfrm.lstForaminifera.Visible = False
      GDSearchfrm.lstForaminifera.Enabled = False
      GDSearchfrm.lstForaminifera.Selected(0) = False
      GDSearchfrm.lstForamsdic.Visible = True
      GDSearchfrm.lstForamsdic.Enabled = True
      GDSearchfrm.lstForamsdic.Selected(0) = True
      GDSearchfrm.chkActForam.value = vbUnchecked
      Combo10.Visible = True
      lblFosForam.Visible = True
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstForamsdic.Enabled = False
      For i& = 1 To GDSearchfrm.lstForamsdic.ListCount
         GDSearchfrm.lstForamsdic.Selected(i& - 1) = False
      Next i&
      chkDicForam.value = vbUnchecked
      GDSearchfrm.lstForamsdic.Visible = False
      Combo10.Visible = False
      lblFosForam.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDicMega_Click()
   If chkDicMega.value = vbChecked And GDSearchfrm.chkMegafauna.value = vbChecked Then
      GDSearchfrm.lstMegafauna.Enabled = False
      GDSearchfrm.lstMegafauna.Selected(0) = False
      GDSearchfrm.lstMegafauna.Visible = False
      GDSearchfrm.lstMegadic.Visible = True
      GDSearchfrm.lstMegadic.Enabled = True
      GDSearchfrm.lstMegadic.Selected(0) = True
      GDSearchfrm.chkActMega.value = vbUnchecked
      Combo11.Visible = True
      lblFosMega.Visible = True
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstMegadic.Enabled = False
      For i& = 1 To GDSearchfrm.lstMegadic.ListCount
         GDSearchfrm.lstMegadic.Selected(i& - 1) = False
      Next i&
      chkDicMega.value = vbUnchecked
      GDSearchfrm.lstMegadic.Visible = False
      Combo11.Visible = False
      lblFosMega.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDicNano_Click()
   If chkDicNano.value = vbChecked And GDSearchfrm.chkNanoplankton.value = vbChecked Then
      GDSearchfrm.lstNannoplankton.Enabled = False
      GDSearchfrm.lstNannoplankton.Selected(0) = False
      GDSearchfrm.lstNannoplankton.Visible = False
      GDSearchfrm.lstNanodic.Visible = True
      GDSearchfrm.lstNanodic.Enabled = True
      GDSearchfrm.lstNanodic.Selected(0) = True
      GDSearchfrm.chkActNano.value = vbUnchecked
      Combo12.Visible = True
      lblFosNano.Visible = True
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstNanodic.Enabled = False
      For i& = 1 To GDSearchfrm.lstNanodic.ListCount
         GDSearchfrm.lstNanodic.Selected(i& - 1) = False
      Next i&
      chkDicNano.value = vbUnchecked
      GDSearchfrm.lstNanodic.Visible = False
      Combo12.Visible = False
      lblFosNano.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDicOstra_Click()
   If chkDicOstra.value = vbChecked And GDSearchfrm.chkOstracoda.value = vbChecked Then
      GDSearchfrm.lstOstracoda.Enabled = False
      GDSearchfrm.lstOstracoda.Selected(0) = False
      GDSearchfrm.lstOstracoda.Visible = False
      GDSearchfrm.lstOstracoddic.Visible = True
      GDSearchfrm.lstOstracoddic.Enabled = True
      GDSearchfrm.lstOstracoddic.Selected(0) = True
      GDSearchfrm.chkActOstra.value = vbUnchecked
      Combo13.Visible = True
      lblFosOstra.Visible = True
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstOstracoddic.Enabled = False
      For i& = 1 To GDSearchfrm.lstOstracoddic.ListCount
         GDSearchfrm.lstOstracoddic.Selected(i& - 1) = False
      Next i&
      chkDicOstra.value = vbUnchecked
      GDSearchfrm.lstOstracoddic.Visible = False
      Combo13.Visible = False
      lblFosOstra.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkDicPaly_Click()
   If chkDicPaly.value = vbChecked And GDSearchfrm.chkPalynology.value = vbChecked Then
      GDSearchfrm.lstPalynology.Enabled = False
      GDSearchfrm.lstPalynology.Selected(0) = False
      GDSearchfrm.lstPalynology.Visible = False
      GDSearchfrm.lstPalyndic.Visible = True
      GDSearchfrm.lstPalyndic.Enabled = True
      GDSearchfrm.lstPalyndic.Selected(0) = True
      GDSearchfrm.chkActPaly.value = vbUnchecked
      Combo14.Visible = True
      lblFosPaly.Visible = True
      'MsgBox "Tip: in order to obtain meaningful results for the paleo zones of this fossil" & vbLf & _
      '       "be sure to limit your search to this fossil only " & vbLf & _
      '       "(i.e., uncheck the other fossils)", vbInformation + vbOKOnly, "MapDigitizer"
   Else
      GDSearchfrm.lstPalyndic.Enabled = False
      For i& = 1 To GDSearchfrm.lstPalyndic.ListCount
         GDSearchfrm.lstPalyndic.Selected(i& - 1) = False
      Next i&
      chkDicPaly.value = vbUnchecked
      GDSearchfrm.lstPalyndic.Visible = False
      Combo14.Visible = False
      lblFosPaly.Visible = False
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkForaminifera_Click()
   If chkForaminifera.value = vbUnchecked Then
   'disenable searches over zones and fossils
      chkActForam_Click
      chkDicForam_Click
      chkActForam.Visible = False
      chkDicForam.Visible = False
      lstForaminifera.Visible = False
      lstForamsdic.Visible = False
   Else
      chkActForam.Visible = True
      chkDicForam.Visible = True
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkJustCores_Click()
   chkAllWells.value = vbUnchecked
   chkJustCuttings.value = vbUnchecked
End Sub

Private Sub chkJustCuttings_Click()
   chkJustCores.value = vbUnchecked
   chkAllWells.value = vbUnchecked
End Sub

Private Sub chkMegafauna_Click()
   If chkMegafauna.value = vbUnchecked Then
   'disenable searches over zones and fossils
      chkActMega_Click
      chkDicMega_Click
      chkActMega.Visible = False
      chkDicMega.Visible = False
      lstMegafauna.Visible = False
      lstMegadic.Visible = False
   Else
      chkActMega.Visible = True
      chkDicMega.Visible = True
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkNanoplankton_Click()
   If chkNanoplankton.value = vbUnchecked Then
   'disenable searches over zones and fossils
      chkActNano_Click
      chkDicNano_Click
      chkActNano.Visible = False
      chkDicNano.Visible = False
      lstNannoplankton.Visible = False
      lstNanodic.Visible = False
   Else
      chkActNano.Visible = True
      chkDicNano.Visible = True
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
 End Sub

Private Sub chkOstracoda_Click()
   If chkOstracoda.value = vbUnchecked Then
   'disenable searches over zones and fossils
      chkActOstra_Click
      chkDicOstra_Click
      chkActOstra.Visible = False
      chkDicOstra.Visible = False
      lstOstracoda.Visible = False
      lstOstracoddic.Visible = False
   Else
      chkActOstra.Visible = True
      chkDicOstra.Visible = True
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkOutcroppings_Click()
   If chkOutcroppings.value = vbChecked Then
      chkWells.Enabled = False
      chkJustCuttings.Enabled = False
      chkJustCores.Enabled = False
      chkAllWells.Enabled = False
      GDMDIform.Toolbar1.Buttons(25).Enabled = True
      GDSearchfrm.tbSearch.TabEnabled(3) = True
   Else
      chkWells.Enabled = True
      End If
End Sub

Private Sub chkPalynology_Click()
   If chkPalynology.value = vbUnchecked Then
   'disenable searches over zones and fossils
      chkActPaly_Click
      chkDicPaly_Click
      chkActPaly.Visible = False
      chkDicPaly.Visible = False
      lstPalynology.Visible = False
      lstPalyndic.Visible = False
   Else
      chkActPaly.Visible = True
      chkDicPaly.Visible = True
      End If
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkShekef_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub chkWells_Click()
   If chkWells.value = vbChecked Then
      chkOutcroppings.Enabled = False
      chkJustCuttings.Enabled = True
      chkJustCores.Enabled = True
      chkAllWells.Enabled = True
      chkAllWells.value = vbChecked
   Else
      chkOutcroppings.Enabled = True
      chkJustCuttings.Enabled = False
      chkJustCores.Enabled = False
      chkAllWells.Enabled = False
      End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdActivateForm_Click
' DateTime  : 9/30/2002 19:34
' Purpose   : Enable/Disenable searches over formation names
'---------------------------------------------------------------------------------------
'
Private Sub cmdActivateForm_Click()
   If lstFormation.Enabled = False Then
        lblFormation.Enabled = True
        lstFormation.Enabled = True
        lblFormation.Caption = "Searches will be conducted only over the selected formations"
        cmdActivateForm.Caption = "&Press to disenable searches over specific formations"
        cmdAllForm.Enabled = True
        cmdClearForm.Enabled = True
   Else
        cmdActivateForm.Caption = "&Press to enable searches over specific formations"
        lblFormation.Enabled = False
        lstFormation.Enabled = False
        For i& = 1 To lstFormation.ListCount
           lstFormation.Selected(i& - 1) = False
        Next i&
        lblFormation.Caption = "Searches will be conducted over any formation"
        cmdAllForm.Enabled = False
        cmdClearForm.Enabled = False
        End If
End Sub

Private Sub cmdAll_Click()
   chkConodonta.value = vbChecked
   chkDiatom.value = vbChecked
   chkForaminifera.value = vbChecked
   chkMegafauna.value = vbChecked
   chkNanoplankton.value = vbChecked
   chkOstracoda.value = vbChecked
   chkPalynology.value = vbChecked
   If SearchDBs% <> 2 Then chkShekef.value = vbChecked
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAll_MouseMove
' DateTime  : 9/30/2002 17:10
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAll.BackColor = &H8000000F Then
        cmdAll.BackColor = &HC0C0C0
        cmdAllAnd.BackColor = &H8000000F
        cmdAllOr.BackColor = &H8000000F
        cmdClearAll.BackColor = &H8000000F
        cmdEditSql.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = "Search over all fossil categories"
End Sub

Private Sub cmdAllAnd_Click()
   Combo1.ListIndex = 1
   Combo2.ListIndex = 1
   Combo3.ListIndex = 1
   Combo4.ListIndex = 1
   Combo5.ListIndex = 1
   Combo6.ListIndex = 1
   Combo7.ListIndex = 1
   If SearchDBs% <> 2 Then Combo15.ListIndex = 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAllAnd_MouseMove
' DateTime  : 9/30/2002 17:14
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdAllAnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAllAnd.BackColor = &H8000000F Then
        cmdAllAnd.BackColor = &HC0C0C0
        cmdAll.BackColor = &H8000000F
        cmdAllOr.BackColor = &H8000000F
        cmdClearAll.BackColor = &H8000000F
        cmdEditSql.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = "Boolean AND search over all fossil categories"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAllForm_Click
' DateTime  : 9/30/2002 19:36
' Purpose   : Pick all the formation names to search for
'---------------------------------------------------------------------------------------
'
Private Sub cmdAllForm_Click()
   For i& = 1 To GDSearchfrm.lstFormation.ListCount
      GDSearchfrm.lstFormation.Selected(i& - 1) = True
   Next i&
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAllForm_MouseMove
' DateTime  : 9/30/2002 17:25
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdAllForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAllForm.BackColor = &H8000000F Then
        cmdAllForm.BackColor = &HC0C0C0
        cmdClearForm.BackColor = &H8000000F
        End If
     GDMDIform.StatusBar1.Panels(1) = "Search over all records with a recorded formation name"
End Sub

Private Sub cmdAllOr_Click()
   Combo1.ListIndex = 0
   Combo2.ListIndex = 0
   Combo3.ListIndex = 0
   Combo4.ListIndex = 0
   Combo5.ListIndex = 0
   Combo6.ListIndex = 0
   Combo7.ListIndex = 0
   If SearchDBs% <> 2 Then Combo15.ListIndex = 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAllOr_MouseMove
' DateTime  : 9/30/2002 17:15
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdAllOr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAllOr.BackColor = &H8000000F Then
        cmdAllAnd.BackColor = &H8000000F
        cmdAll.BackColor = &H8000000F
        cmdAllOr.BackColor = &HC0C0C0
        cmdClearAll.BackColor = &H8000000F
        cmdEditSql.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = "Boolean OR search over all fossil categories"
End Sub

Private Sub cmdAllSources_Click()
  chkWells.value = vbChecked
  chkOutcroppings.value = vbChecked
  chkWells.Enabled = True
  chkOutcroppings.Enabled = True
  txtLimup.Enabled = True
  udLimup.Enabled = True
  txtLimdo.Enabled = True
  udLimdo.Enabled = True
  lblLimdo.Enabled = True
  lblLimup.Enabled = True
  chkAllWells.Enabled = False
  chkJustCores.Enabled = False
  chkJustCuttings.Enabled = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAllSources_MouseMove
' DateTime  : 9/30/2002 16:40
' Purpose   : Make button face color change as mouse moves across it
'---------------------------------------------------------------------------------------
'
Private Sub cmdAllSources_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If cmdAllSources.BackColor = &H8000000F Then
      cmdOutcroppings.BackColor = &H8000000F
      cmdWells.BackColor = &H8000000F
      cmdAllSources.BackColor = &HC0C0C0
      End If
   GDMDIform.StatusBar1.Panels(1) = "Search all wells and outcroppings"
End Sub

Private Sub cmdAnalyst_Click()
  If lstAnalyst.Enabled = False Then
     lstAnalyst.Enabled = True
     cmdAnalyst.Caption = "Disenable"
     'MsgBox "Tip: in order to obtain meaningful results for this search" & vbLf & _
     '       "be sure that you selected no more than one fossil under ""Fossil Type""." & vbLf _
     '       , vbInformation + vbOKOnly, "MapDigitizer"
  Else
     lstAnalyst.Enabled = False
     cmdAnalyst.Caption = "Enable"
     For i& = 1 To lstAnalyst.ListCount
        lstAnalyst.Selected(i& - 1) = False
     Next i&
     sAnalystSearch = sEmpty
     End If
End Sub

Private Sub cmdBack_Click()
   cmdNext.Enabled = True
   stepsearch& = stepsearch& - 1
   If stepsearch& = 0 Then cmdBack.Enabled = False
   tbSearch.Tab = stepsearch&
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdBack_MouseMove
' DateTime  : 9/30/2002 16:55
' Purpose   : Highlight button as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdBack.BackColor = &H8000000F Then
        cmdBack.BackColor = &HC0C0C0
        cmdNext.BackColor = &H8000000F
        cmdSearchReport.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = "Previous tab"
End Sub

Private Sub cmdClear_Click()
   txtEastMin = "0.0"
   txtEastMax = "0.0"
   txtNorthMin = "0.0"
   txtNorthMax = "0.0"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdClear_MouseMove
' DateTime  : 9/30/2002 17:29
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdClear.BackColor = &H8000000F Then
        cmdClear.BackColor = &HC0C0C0
        cmdFullExtent.BackColor = &H8000000F
        End If
     GDMDIform.StatusBar1.Panels(1) = "Clear search boundaries (search over all records notwithstanding their coordinates)"
End Sub

Private Sub cmdClearAll_Click()
   chkConodonta.value = vbUnchecked
   chkDiatom.value = vbUnchecked
   chkForaminifera.value = vbUnchecked
   chkMegafauna.value = vbUnchecked
   chkNanoplankton.value = vbUnchecked
   chkOstracoda.value = vbUnchecked
   chkPalynology.value = vbUnchecked
   If SearchDBs% <> 2 Then chkShekef.value = vbUnchecked
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdClearAll_MouseMove
' DateTime  : 9/30/2002 17:12
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdClearAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdClearAll.BackColor = &H8000000F Then
        cmdAllAnd.BackColor = &H8000000F
        cmdAll.BackColor = &H8000000F
        cmdAllOr.BackColor = &H8000000F
        cmdClearAll.BackColor = &HC0C0C0
        cmdEditSql.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = "Clear all fossil categories"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdClearForm_Click
' DateTime  : 9/30/2002 19:36
' Purpose   : Clear all the highlighted formation names
'---------------------------------------------------------------------------------------
'
Private Sub cmdClearForm_Click()
   For i& = 1 To GDSearchfrm.lstFormation.ListCount
      GDSearchfrm.lstFormation.Selected(i& - 1) = False
   Next i&
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdClearForm_MouseMove
' DateTime  : 9/30/2002 17:25
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdClearForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdClearForm.BackColor = &H8000000F Then
        cmdClearForm.BackColor = &HC0C0C0
        cmdAllForm.BackColor = &H8000000F
        End If
     GDMDIform.StatusBar1.Panels(1) = "Unselect highlighted formations"
End Sub

Private Sub cmdClient_Click()
  If lstClient.Enabled = False Then
     lstClient.Enabled = True
     cmdClient.Caption = "Disenable"
     lstProject.Enabled = False
     cmdProject.Caption = "Enable"
     lstCompany.Enabled = False
     cmdCompany.Caption = "Enable"
     lstDate.Enabled = False
     cmdDate.Caption = "Enable"
  Else
     lstClient.Enabled = False
     cmdClient.Caption = "Enable"
     For i& = 1 To lstClient.ListCount
        lstClient.Selected(i& - 1) = False
     Next i&
     sFormSearch = sEmpty
     End If
End Sub

Private Sub cmdCompany_Click()
  If lstCompany.Enabled = False Then
     lstCompany.Enabled = True
     cmdCompany.Caption = "Disenable"
     lstDate.Enabled = False
     cmdDate.Caption = "Enable"
     lstClient.Enabled = False
     cmdClient.Caption = "Enable"
     lstProject.Enabled = False
     cmdProject.Caption = "Enable"
  Else
     lstCompany.Enabled = False
     cmdCompany.Caption = "Enable"
     For i& = 1 To lstCompany.ListCount
        lstCompany.Selected(i& - 1) = False
     Next i&
    sFormSearch = sEmpty
    sFormSearchOld = sEmpty
    End If
End Sub

Private Sub cmdDate_Click()
  If lstDate.Enabled = False Then
     lstDate.Enabled = True
     cmdDate.Caption = "Disenable"
     lstCompany.Enabled = False
     cmdCompany.Caption = "Enable"
     lstClient.Enabled = False
     cmdClient.Caption = "Enable"
     lstProject.Enabled = False
     cmdProject.Caption = "Enable"
  Else
     lstDate.Enabled = False
     cmdDate.Caption = "Enable"
     For i& = 1 To lstDate.ListCount
        lstDate.Selected(i& - 1) = False
     Next i&
     sFormSearch = sEmpty
     sFormSearchOld = sEmpty
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdEarlier_Click
' DateTime  : 9/30/2002 19:31
' Purpose   : "Fix" earlier geologic age date
'---------------------------------------------------------------------------------------
'
Private Sub cmdEarlier_Click()
   If txtEarlier.Enabled = True Then
      txtEarlier.Enabled = False
      cmbEPre.Enabled = False
      cmbEQuest.Enabled = False
      cmdEarlier.Caption = "Press to change the &Earlier Date"
   Else
      txtEarlier.Enabled = True
      cmbEPre.Enabled = True
      cmbEQuest.Enabled = True
      cmdEarlier.Caption = "Press to accept as the &Earlier Date"
      End If
End Sub

Private Sub cmdEditdb1_Click()

   'form the BOOLEAN search string over fossil tables
   'of the active database
   
    With GDSearchfrm
   
         .txtSQLdb1.Enabled = True
        
          Createdb1FossilSql
            
         .txtSQLdb1.Text = strSql1
    
    End With
    
End Sub

Private Sub cmdEditSql_Click()
  If lineSQL.Visible = False Then
    'show Fossil Category SQL if searching over that database
    lineSQL.Visible = True
    If SearchDBs% = 1 Or SearchDBs% = 2 Then
        cmdEditdb1.Visible = True
        txtSQLdb1.Visible = True
        txtSQLdb1.Enabled = False
      
        'present the SQL clause if it exists
        If txtSQLdb1.Enabled = True Then
           If linked Then cmdEditdb1.value = True
        Else
           If linked Then cmdEditdb1.value = True
           txtSQLdb1.Enabled = False
           End If
        
        End If
        
    If SearchDBs% = 1 Or SearchDBs% = 3 Then
        cmdEditdb2.Visible = True
        txtSQLdb2.Visible = True
        txtSQLdb2.Enabled = False
      
        'present the SQL clause if it exists
        If txtSQLdb2.Enabled = True Then
           If linkedOld Then cmdEditdb2.value = True
        Else
           If linkedOld Then cmdEditdb2.value = True
           txtSQLdb2.Enabled = False
           End If
        
        End If
        
    cmdEditSql.Caption = "Don't show Fossil Type &SQL"
    cmdEditSql.ToolTipText = "Don't show the Fossil Type SQL clause"
  Else
    'don't show Fossil Category SQL
    lineSQL.Visible = False
    txtSQLdb1.Visible = False
    cmdEditdb1.Visible = False
    txtSQLdb2.Visible = False
    cmdEditdb2.Visible = False
    cmdEditSql.Caption = "Show Fossil Types &SQL"
    cmdEditSql.ToolTipText = "Show the Fossil Type SQL clause"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdEditSql_MouseMove
' DateTime  : 12/26/2002 14:50
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdEditSql_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdEditSql.BackColor = &H8000000F Then
        cmdEditSql.BackColor = &HC0C0C0
        cmdAll.BackColor = &H8000000F
        cmdAllAnd.BackColor = &H8000000F
        cmdAllOr.BackColor = &H8000000F
        cmdClearAll.BackColor = &H8000000F
        End If
     If Searching Then Exit Sub 'don't reset status bar during searches
     If lineSQL.Visible = True Then
        GDMDIform.StatusBar1.Panels(1) = "Disenable the editing of the Fossil Type SQL clauses"
     Else
        GDMDIform.StatusBar1.Panels(1) = "Enable editing the Fossil Type SQL clauses (for advanced users only)"
        End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdEnabDate_Click
' DateTime  : 9/30/2002 19:29
' Purpose   : Enable/Disenable searches over geologic age dates
'---------------------------------------------------------------------------------------
'
Private Sub cmdEnabDate_Click()
   If TreeView1.Enabled = False Then
      TreeView1.Enabled = True
      cmdEarlier.Enabled = True
      txtEarlier.Enabled = True
      cmdLater.Enabled = True
      txtLater.Enabled = True
      cmbEPre.Enabled = True
      cmbEQuest.Enabled = True
      cmbLPre.Enabled = True
      cmbLQuest.Enabled = True
      optExact.Enabled = True
      optRange.Enabled = True
      lblEPre.Enabled = True
      lblEQuest.Enabled = True
      lblEarlier.Enabled = True
      lblLPre.Enabled = True
      lblLQuest.Enabled = True
      lblLater.Enabled = True
      cmdEnabDate.Caption = "&Press here to disenable searches over geolgic age dates"
   Else
      TreeView1.Enabled = False
      cmdEarlier.Enabled = False
      txtEarlier.Enabled = False
      txtEarlier = sEmpty
      cmbEPre.Text = sEmpty
      cmbEQuest.Text = sEmpty
      cmdLater.Enabled = False
      txtLater.Enabled = False
      txtLater = sEmpty
      cmbLPre.Text = sEmpty
      cmbLQuest.Text = sEmpty
      cmbEPre.Enabled = False
      cmbEQuest.Enabled = False
      cmbLPre.Enabled = False
      cmbLQuest.Enabled = False
      optExact.Enabled = False
      optRange.Enabled = False
      lblEPre.Enabled = False
      lblEQuest.Enabled = False
      lblEarlier.Enabled = False
      lblLPre.Enabled = False
      lblLQuest.Enabled = False
      lblLater.Enabled = False
      cmdEnabDate.Caption = "&Press here to enable searches over a range of geolgic age dates"
      End If
End Sub

Private Sub cmdFullExtent_Click()
   If picnam$ <> sEmpty Then
     txtEastMin = x10
     txtEastMax = x20
     txtNorthMin = y20
     txtNorthMax = y10
   Else
     txtEastMin = 0
     txtEastMax = 0
     txtNorthMin = 0
     txtNorthMax = 0
     End If
End Sub


Private Sub cmdCancel_Click()
   Call Form_QueryUnload(0, 0)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdFullExtent_MouseMove
' DateTime  : 9/30/2002 17:28
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdFullExtent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdFullExtent.BackColor = &H8000000F Then
        cmdFullExtent.BackColor = &HC0C0C0
        cmdClear.BackColor = &H8000000F
        End If
     GDMDIform.StatusBar1.Panels(1) = "Search for records with coordinates anywhere on the geologic map"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLater_Click
' DateTime  : 9/30/2002 19:30
' Purpose   : "Fix" later geologic date
'---------------------------------------------------------------------------------------
'
Private Sub cmdLater_Click()
   If txtLater.Enabled = True Then
      txtLater.Enabled = False
      cmbLPre.Enabled = False
      cmbLQuest.Enabled = False
      cmdLater.Caption = "Press to change the &Later Date"
   Else
      txtLater.Enabled = True
      cmbLPre.Enabled = True
      cmbLQuest.Enabled = True
      cmdLater.Caption = "Press to accept as the &Later Date"
      End If
End Sub

Private Sub cmdMinimize_Click()
   'user requested to minimize this form
   ret = ShowWindow(GDSearchfrm.hWnd, SW_MINIMIZE)
End Sub

Private Sub cmdName1_Click()
   'button used to enabling string searches over string Name1
   If cmdName1.Caption = "Enable" Then
      cmdName1.ToolTipText = "Click to disenable"
      cmdName1.Caption = "Disenable"
      cmbName1.Enabled = True
      txtName1.Enabled = True
      If frmDictionary.Enabled = True Then optName1.Enabled = True
      chkCase.Enabled = True
      cmdName1.BackColor = &HC0C0FF
      txtName1.BackColor = &HC0FFFF
      cmbName1.BackColor = &HC0FFFF
      cmdName2.Enabled = True
      
   Else
      cmdName1.Caption = "Enable"
      cmdName1.ToolTipText = "Click to enable searches over String #1"
      cmbName1.Enabled = False
      txtName1.Enabled = False
      If frmDictionary.Enabled = True Then optName1.Enabled = False
      chkCase.Enabled = False
      cmdName1.BackColor = &H8000000F
      txtName1.BackColor = &H80000005
      cmbName1.BackColor = &H80000005
      
      'If this string is diseanbled then all 2nd string also disenabled
      cmdName2.Enabled = False
      cmdName2.Caption = "Enable"
      cmdName2.ToolTipText = "Click to enable searches over String #2"
      cmbName2.Enabled = False
      txtName2.Enabled = False
      If frmDictionary.Enabled = True Then optName2.Enabled = False
      cmdName2.BackColor = &H8000000F
      txtName2.BackColor = &H80000005
      cmbName2.BackColor = &H80000005
      
      End If
End Sub

Private Sub cmdName2_Click()
   'button used to enabling string searches over string Name2
   If cmdName2.Caption = "Enable" Then
      cmdName2.ToolTipText = "Click to disenable"
      cmdName2.Caption = "Disenable"
      cmbName2.Enabled = True
      txtName2.Enabled = True
      If frmDictionary.Enabled = True Then optName2.Enabled = True
      cmdName2.BackColor = &HC0C0FF
      txtName2.BackColor = &HC0FFFF
      cmbName2.BackColor = &HC0FFFF
   Else
      cmdName2.Caption = "Enable"
      cmdName2.ToolTipText = "Click to enable searches over String #2"
      cmbName2.Enabled = False
      txtName2.Enabled = False
      If frmDictionary.Enabled = True Then optName2.Enabled = False
      cmdName2.BackColor = &H8000000F
      txtName2.BackColor = &H80000005
      cmbName2.BackColor = &H80000005
      End If
End Sub

Private Sub cmdNext_Click()
   cmdBack.Enabled = True
   stepsearch& = stepsearch& + 1
   If stepsearch& = tbSearch.Tabs - 1 Then cmdNext.Enabled = False
   tbSearch.Tab = stepsearch&
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdNext_MouseMove
' DateTime  : 9/30/2002 16:56
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdNext.BackColor = &H8000000F Then
        cmdBack.BackColor = &H8000000F
        cmdNext.BackColor = &HC0C0C0
        cmdSearchReport.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = "Next tab"
End Sub

Private Sub cmdOutcroppings_Click()
   If chkOutcroppings.value = vbChecked Then
      chkOutcroppings.Enabled = True
      chkWells.value = vbUnchecked
      chkAllWells.value = vbUnchecked
      chkJustCores.value = vbUnchecked
      chkWells.Enabled = False
      chkAllWells.Enabled = False
      chkJustCores.Enabled = False
      txtLimup.Enabled = False
      udLimup.Enabled = False
      txtLimdo.Enabled = False
      udLimdo.Enabled = False
      lblLimdo.Enabled = False
      lblLimup.Enabled = False
   Else
      chkWells.Enabled = False
      chkJustCuttings.Enabled = False
      chkJustCores.Enabled = False
      chkAllWells.Enabled = False
      txtLimup.Enabled = False
      udLimup.Enabled = False
      txtLimdo.Enabled = False
      udLimdo.Enabled = False
      lblLimdo.Enabled = False
      lblLimup.Enabled = False
      chkWells.value = vbUnchecked
      chkAllWells.value = vbUnchecked
      chkJustCores.value = vbUnchecked
      chkOutcroppings.Enabled = True
      chkOutcroppings.value = vbChecked
      GDMDIform.Toolbar1.Buttons(25).Enabled = True
      GDSearchfrm.tbSearch.TabEnabled(3) = True
      End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdOutcroppings_MouseMove
' DateTime  : 9/30/2002 16:39
' Purpose   : change button color as cursor moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdOutcroppings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If cmdOutcroppings.BackColor = &H8000000F Then
      cmdOutcroppings.BackColor = &HC0C0C0
      cmdWells.BackColor = &H8000000F
      cmdAllSources.BackColor = &H8000000F
      End If
   GDMDIform.StatusBar1.Panels(1) = "Search all outcroppings"
End Sub

Private Sub cmdPasteDictionary_Click()
   If PasteName1 Then
      txtName1 = cmbDictionary.Text
   ElseIf PasteName2 Then
      txtName2 = cmbDictionary.Text
      End If
End Sub

Private Sub cmdProject_Click()
  If lstProject.Enabled = False Then
     lstProject.Enabled = True
     cmdProject.Caption = "Disenable"
     lstCompany.Enabled = False
     cmdCompany.Caption = "Enable"
     lstDate.Enabled = False
     cmdDate.Caption = "Enable"
     lstClient.Enabled = False
     cmdClient.Caption = "Enable"
  Else
     lstProject.Enabled = False
     cmdProject.Caption = "Enable"
     For i& = 1 To lstProject.ListCount
        lstProject.Selected(i& - 1) = False
     Next i&
     sFormSearch = sEmpty
     sFormSearchOld = sEmpty
     End If
End Sub

Private Sub cmdSearchReport_Click()

'****************************************
 'this routine starts the database search
 '****************************************

On Error GoTo errhand
   
'If an old search report is still visible, then erase its
'arrays and any plotted points:
If PicSum Then
   'press the clear button on the Search Report Form
   GDReportfrm.cmdClear.value = True
   End If

With GDSearchfrm
   
   '_______________________________________________________
   'first check inputs
   If val(.txtEastMin) > val(.txtEastMax) Or val( _
       .txtNorthMin) > val(.txtNorthMax) Then
      Screen.MousePointer = vbDefault
      BringWindowToTop (.hWnd)
      .tbSearch.Tab = 0
      MsgBox "The Coordinate boundaries have not been defined properly!", _
          vbExclamation + vbOKOnly, "MapDigitizer"
      If ArcDump Then numArc& = -100
      Exit Sub
      End If
      
   If val(.txtLimup) > val(.txtLimdo) Then
      Screen.MousePointer = vbDefault
      BringWindowToTop (.hWnd)
      .tbSearch.Tab = 1
      MsgBox "Maximum Depth must be greater or equal to Minimum Depth!", _
          vbExclamation + vbOKOnly, "MapDigitizer"
      If ArcDump Then numArc& = -100
      Exit Sub
      End If
   
   If .chkWells.value = vbUnchecked And _
       .chkOutcroppings.value = vbUnchecked Then
      Screen.MousePointer = vbDefault
      BringWindowToTop (.hWnd)
      .tbSearch.Tab = 1
      MsgBox "You must choose a sample source!" & vbLf & _
          "That is, choose either Wells, or Outcroppings, or both to search.", _
          vbExclamation + vbOKOnly, "MapDigitizer"
      If ArcDump Then numArc& = -100
      Exit Sub
      End If

   If .TreeView1.Enabled = True Then 'date searches enabled
      If .txtEarlier = sEmpty And .txtLater = sEmpty Then
         'no range of dates were defined
         Screen.MousePointer = vbDefault
         BringWindowToTop (.hWnd)
         .tbSearch.Tab = 4
         MsgBox "You have enabled the search over geologic ages." & vbLf & _
             "However, a range of dates has not been defined!" & vbLf & _
             "To search over dates, you must choose a range of dates." & vbLf & _
             "If you don't wish to search over dates, disenable it" & vbLf & _
             "by pressing the large button above.", vbExclamation + vbOKOnly, _
             "MapDigitizer"
         If ArcDump Then numArc& = -100
         Exit Sub
         End If
      End If
      
   '------------finished checking inputs---------------------

  
   '-----------------Parameters for Active and Old database search-----------
   '_________________________________________________________
   'if searching over dates, write search string
   sDateRange = sEmpty
   If .TreeView1.Enabled = True Then
      If RangeOfDates Then 'range of dates
         'find earlier Date
         If .txtEarlier = sEmpty Then
            'use earliest date as boundary
            numEarlier& = 0
         Else
            numEarlier& = 0
            For i& = 0 To .lstDates.ListCount - 1
                If .lstDates.List(i&) = Trim$( _
                    .txtEarlier) Then
                   numEarlier& = i&
                   Exit For
                   End If
            Next i&
            If numEarlier& = 0 Then
               Screen.MousePointer = vbDefault
               BringWindowToTop (.hWnd)
               .tbSearch.Tab = 4
               MsgBox "I don't recognize the earlier date!" & vbLf & _
                   "You must use one of the dates in the provided Date Line.", _
                   vbExclamation + vbOKOnly, "MapDigitizer"
               Exit Sub
               End If
            End If

         'find later Date
         If .txtLater = sEmpty Then
            'use earliest date as boundary
            numLater& = .lstDates.ListCount - 1
         Else
            numLater& = 0
            For i& = 0 To .lstDates.ListCount - 1
                If .lstDates.List(i&) = Trim$( _
                    .txtLater) Then
                   numLater& = i&
                   Exit For
                   End If
            Next i&
            If numLater& = 0 Then
               Screen.MousePointer = vbDefault
               BringWindowToTop (.hWnd)
               .tbSearch.Tab = 4
               MsgBox "I don't recognize the later date!" & vbLf & _
                   "You must use one of the dates in the provided Date Line.", _
                   vbExclamation + vbOKOnly, "MapDigitizer"
               Exit Sub
               End If
            End If

         If numLater& > numEarlier& Then
            Screen.MousePointer = vbDefault
            BringWindowToTop (.hWnd)
            .tbSearch.Tab = 4
            MsgBox "Your earlier date seems to be later than your later date!" _
                & vbLf & "Please check your entries, and try again.", _
                vbExclamation + vbOKOnly, "MapDigitizer"
            Exit Sub
            End If

         If numEarlier& - numLater& > 70 And SearchDBs% <> 2 And SearchDBs% <> 0 Then
            'The old database age range search is done using SQL
            'so it is limited to a finite number of clauses
            Screen.MousePointer = vbDefault
            BringWindowToTop (.hWnd)
            .tbSearch.Tab = 4
            MsgBox "The geologic age date range is too large!" & vbLf & _
                "Please narrow the range, and try again.", vbExclamation + vbOKOnly, _
                "MapDigitizer"
            Exit Sub
            End If

         sDateRange = sEmpty
         For i& = numEarlier& To numLater& Step -1
             sDateRange = sDateRange & .lstDates.List(i&) & " ?"
         Next i&
      Else 'exact dates
      End If
   
      If SearchDBs% = 1 Or SearchDBs% = 2 Then 'searching active database
         'check that are querying at least one fossil
         If .chkConodonta.value = vbChecked Or _
             .chkDiatom.value = vbChecked Or _
             .chkForaminifera.value = vbChecked Or _
             .chkMegafauna.value = vbChecked Or _
             .chkNanoplankton.value = vbChecked Or _
             .chkOstracoda.value = vbChecked Or _
             .chkPalynology.value = vbChecked Then
         Else
            If SearchDBs% = 2 Then 'only searching over active database
                Screen.MousePointer = vbDefault
                BringWindowToTop (.hWnd)
                .tbSearch.Tab = 2
                MsgBox "When searching the active database for geologic ages" & vbLf & _
                       "you must choose fossil type(s)!", _
                        vbExclamation + vbOKOnly, "MapDigitizer"
                If ArcDump Then numArc& = -100
                Exit Sub
             ElseIf SearchDBs% = 1 Then 'searching over both databases
                'give warning and have user decide whether to abort
                Screen.MousePointer = vbDefault
                resp = MsgBox("Warning! You must choose fossil type(s) to search the active database for geo. ages!" & vbLf & _
                       "If you continue with this search, then only the scanned database will be searched." & vbLf & vbLf & _
                       "Continue this search?" & vbLf, vbExclamation + vbYesNo, "MapDigitizer")
                If resp = vbNo Then
                    If ArcDump Then numArc& = -100
                    Exit Sub
                Else
                    Screen.MousePointer = vbHourglass
                    End If
                End If
            End If
         End If
         
   End If

   '____________________________________________________________
   'Select the formation names in the unselected list that
   'is used for searches
   Screen.MousePointer = vbHourglass
   'select the new selected records
   If .lstFormation.Enabled = True Then
      'clear any old selected records
      For j& = 1 To .lstFormationUnsorted.ListCount
         .lstFormationUnsorted.Selected(j& - 1) = False
      Next j&
      found% = 0
      'mm& = 0
      For i& = 1 To .lstFormation.ListCount
         If .lstFormation.Selected(i& - 1) Then
            For j& = 1 To .lstFormationUnsorted.ListCount
                If .lstFormationUnsorted.List( _
                    j& - 1) = .lstFormation.List(i& - 1) Then
                   .lstFormationUnsorted.Selected(j& - 1) = True
                   found% = 1
                   'mm& = mm& + 1
                   'If SearchDBs% = 1 Or SearchDBs% = 3 Then
                   '   'check for maximum number of formations
                   '   If mm& >= 120 Then
                   '     Screen.MousePointer = vbDefault
                   '     BringWindowToTop (.hWnd)
                   '     .tbSearch.Tab = 3
                   '     MsgBox "When searching over the old database, you are limited to" & vbLf & '            "searching over a maximum of 120 formations at one time!" & vbLf & '            "You have exceeded this limit.  Select fewer formations." & vbLf & vbLf & '            "If you wish to search over all the formations," & vbLf & '            "then press the large button above to disenable searches over" & vbLf & '            "individual formations.", vbOKOnly + vbExclamation, "GSI_GDB"
                   '     Exit Sub
                   '     End If
                   Exit For
                   End If
            Next j&
            End If
       Next i&
       If found% = 0 Then
         Screen.MousePointer = vbDefault
         BringWindowToTop (.hWnd)
         .tbSearch.Tab = 3
         MsgBox "You have activated the Formation List Box," & vbLf & _
             "but none have been selected!" & vbLf & _
             "If you don't wish to search over Formations," & vbLf & _
             "then press the large button above.", vbOKOnly + vbExclamation, _
             "GSI_GDB"
         Exit Sub
         End If
       
       End If

   '___________________________________________________________
   'Now form string from Form table fields that user is searching over
   sFormSearch = sEmpty
   sFormSearchOld = sEmpty
   If .lstClient.Enabled = True Then
      'clear any old selected records
      'and check that at least one record is selected
      For j& = 1 To .lstClientUnsorted.ListCount
         .lstClientUnsorted.Selected(j& - 1) = False
      Next j&
      found% = 0
      For i& = 1 To .lstClient.ListCount
         If .lstClient.Selected(i& - 1) Then
            For j& = 1 To .lstClientUnsorted.ListCount
                If .lstClientUnsorted.List( _
                    j& - 1) = .lstClient.List(i& - 1) Then
                   .lstClientUnsorted.Selected(j& - 1) = True
                   found% = 1
                   Exit For
                   End If
            Next j&
            End If
       Next i&
       If found% = 0 Then
          Screen.MousePointer = vbDefault
          BringWindowToTop (.hWnd)
          .tbSearch.Tab = 5
          MsgBox "You have activated the Client List Box," & vbLf & _
              "but none have been selected!", vbOKOnly + vbExclamation, "GSI_GDB"
          Exit Sub
          End If
   ElseIf .lstDate.Enabled = True Then
      found% = 0
      For i& = 1 To .lstDate.ListCount
         If .lstDate.Selected(i& - 1) Then
            If InStr(.lstDate.List(i& - 1), "-") = 0 Then
                sFormSearch = sFormSearch & "(" & .lstDate.List( _
                    i& - 1) & ")"
                found% = 1
                End If
            If InStr(.lstDate.List(i& - 1), "-") <> 0 Then
                sFormSearchOld = sFormSearchOld & "(" & _
                    .lstDate.List(i& - 1) & ")"
                found% = 1
                End If
            End If
      Next i&
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         BringWindowToTop (.hWnd)
         .tbSearch.Tab = 5
         MsgBox "You have activated the Project Date List Box," & vbLf & _
             "but none have been selected!", vbOKOnly + vbExclamation, "GSI_GDB"
         Exit Sub
         End If
   ElseIf .lstProject.Enabled = True Then
      found% = 0
      For i& = 1 To .lstProject.ListCount
         If .lstProject.Selected(i& - 1) Then
             If InStr(.lstProject.List(i& - 1), "*") = 0 Then
                sFormSearch = sFormSearch & "(" & _
                    .lstFormProject.List(i& - 1) & ")"
                found% = 1
                End If
             If InStr(.lstProject.List(i& - 1), "*") <> 0 Then
                sFormSearchOld = sFormSearchOld & "(" & Mid$( _
                    .lstProject.List(i& - 1), 2, _
                    Len(.lstProject.List(i& - 1)) - 1) & ")"
                found% = 1
                End If
             End If
      Next i&
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         BringWindowToTop (.hWnd)
         .tbSearch.Tab = 5
         MsgBox "You have activated the Project Number List Box," & vbLf & _
             "but none have been selected!", vbOKOnly + vbExclamation, "GSI_GDB"
         Exit Sub
         End If
   ElseIf .lstCompany.Enabled = True Then
      'clear any old selected records
      For j& = 1 To .lstCompanyUnsorted.ListCount
         .lstCompanyUnsorted.Selected(j& - 1) = False
      Next j&
      found% = 0
      For i& = 1 To .lstCompany.ListCount
         If .lstCompany.Selected(i& - 1) Then
            For j& = 1 To .lstCompanyUnsorted.ListCount
                If .lstCompanyUnsorted.List( _
                    j& - 1) = .lstCompany.List(i& - 1) Then
                   .lstCompanyUnsorted.Selected(j& - 1) = True
                   found% = 1
                   Exit For
                   End If
            Next j&
            End If
       Next i&
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         BringWindowToTop (.hWnd)
         .tbSearch.Tab = 5
         MsgBox "You have activated the Company/Division List Box," & vbLf & _
             "but none have been selected!", vbOKOnly + vbExclamation, "GSI_GDB"
         Exit Sub
         End If
      End If

   'flag that Form table is being searched
   If .lstCompany.Enabled = True Or .lstProject.Enabled = _
       True Or .lstDate.Enabled = True Or .lstClient.Enabled _
       = True Then
      If sFormSearch = sEmpty Then sFormSearch = "enabled"
      End If

   '______________________________________________________
   'Do the same thing if the user is searching over analysts
   sAnalystSearch = sEmpty
   If .lstAnalyst.Enabled = True Then
      'clear any old selected records
      For j& = 1 To .lstAnalystUnsorted.ListCount
         .lstAnalystUnsorted.Selected(j& - 1) = False
      Next j&
      found% = 0
      For i& = 1 To .lstAnalyst.ListCount
         If .lstAnalyst.Selected(i& - 1) Then
            For j& = 1 To .lstAnalystUnsorted.ListCount
                If .lstAnalystUnsorted.List( _
                    j& - 1) = .lstAnalyst.List(i& - 1) Then
                   .lstAnalystUnsorted.Selected(j& - 1) = True
                   sAnalystSearch = "enabled"
                   found% = 1
                   Exit For
                   End If
            Next j&
            End If
       Next i&
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         BringWindowToTop (.hWnd)
         .tbSearch.Tab = 5
         MsgBox "You have activated the Analysts List Box," & vbLf & _
             "but none have been selected!", vbOKOnly + vbExclamation, "GSI_GDB"
         Exit Sub
         End If
      If sAnalystSearch <> sEmpty Then
         'check that are querying at least one fossil
         If .chkConodonta.value = vbChecked Or _
             .chkDiatom.value = vbChecked Or _
             .chkForaminifera.value = vbChecked Or _
             .chkMegafauna.value = vbChecked Or _
             .chkNanoplankton.value = vbChecked Or _
             .chkOstracoda.value = vbChecked Or _
             .chkPalynology.value = vbChecked Then
         Else
            Screen.MousePointer = vbDefault
            BringWindowToTop (.hWnd)
            .tbSearch.Tab = 5
            MsgBox "You cannot perform a search of the active database for analysts " & vbLf & _
                "without also picking an active database fossil type to search! " & vbLf & _
                "(Tip: check one of the fossils in ""Fossil Type"" other then ""shekef"".)", _
                vbExclamation + vbOKOnly, "MapDigitizer"
            Exit Sub
            End If
         End If
      End If

   '__________________________________________________
   'load up Paleo Zone search arrays
    LoadPaleoArrays 'load paleo zones into search arrays

   If .lstConodonta.Enabled = True Or .lstDiatom.Enabled _
       = True Or .lstForaminifera.Enabled = True Or _
       .lstMegafauna.Enabled = True Or _
       .lstNannoplankton.Enabled = True Or _
       .lstOstracoda.Enabled = True Or .lstPalynology.Enabled _
       = True Then
      bPaleoZone = True
   Else
      bPaleoZone = False
      End If

   '__________________________________________________
   'load up fossil names into search arrays
    loadFossilArrays 'load fossil names into search arrays

   If .lstConodsdic.Enabled = True Or _
       .lstDiatomsdic.Enabled = True Or .lstForamsdic.Enabled _
       = True Or .lstMegadic.Enabled = True Or _
       .lstNanodic.Enabled = True Or .lstOstracoddic.Enabled _
       = True Or .lstPalyndic.Enabled = True Then
      bFossilNames = True
   Else
      bFossilNames = False
      End If
      
srchOld:
   '------------------------------------------------
   'if the user edited the Fossil Type SQL syntax and is using it for searches
   'then check the syntax for equal numbers of open and closed parenthesis
   'and for a beginning "AND " clause
   If (.txtSQLdb1.Enabled = True And .txtSQLdb1.Visible = True) Or _
      (.txtSQLdb2.Enabled = True And .txtSQLdb2.Visible = True) Then
        Call CheckFossilSQLSyntax(ier%)
        If ier% < 0 Then
           Screen.MousePointer = vbDefault
           BringWindowToTop (.hWnd)
           .tbSearch.Tab = 2
           Select Case ier%
              Case -1
                 MsgBox "Your edited active database Fossil Type SQL clause has error(s) !" & vbLf & _
                    "Check for equal numbers of left and right parenthesis.", _
                    vbExclamation + vbOKOnly, "MapDigitizer"
              Case -2
                 MsgBox "Your edited scanned database Fossil Type SQL clause has error(s) !" & vbLf & _
                    "Check for equal numbers of left and right parenthesis.", _
                    vbExclamation + vbOKOnly, "MapDigitizer"
              Case -3
                 MsgBox "Your edited Fossil Type SQL clauses have errors !" & vbLf & _
                    "Check for equal numbers of left and right parenthesis.", _
                    vbExclamation + vbOKOnly, "MapDigitizer"
              Case -4
                 MsgBox "Your edited active database Fossil Type SQL clause has error(s) !" & vbLf & _
                     "This SQL clause is part of a larger clause, so it must start with ""AND "".", _
                     vbExclamation + vbOKOnly, "MapDigitizer"
              Case -5
                 MsgBox "Your edited scanned database Fossil Type SQL clause has error(s) !" & vbLf & _
                     "This SQL clause is part of a larger clause, so it must start with ""AND "".", _
                     vbExclamation + vbOKOnly, "MapDigitizer"
              Case -9
                 MsgBox "Your edited Fossil Type SQL clauses have errors !" & vbLf & _
                     "The edited SQL clauses are parts of larger clauses," & vbLf & _
                     "so they must start with ""AND "".", _
                     vbExclamation + vbOKOnly, "MapDigitizer"
           End Select
           Exit Sub
           End If
       End If
   
   '-------------SQL statement for Scanned database---------------
   If SearchDBs% = 1 Or SearchDBs% = 3 Then
      CreateOldDbSql 'write SQL string for the old database
      If strSQLOld = gsEmpty And SearchDBs% = 3 Then
        'nothing to search for
        Screen.MousePointer = vbDefault
        MsgBox "No records were found!" & vbLf _
              & "Try using different search conditions.", vbExclamation + vbOKOnly, "MapDigitizer"
        Exit Sub
        End If
      End If
      
   '------------record average coordinates of search area for Google dump-----------
   AveEastSearch = (txtEastMax - txtEastMin) * 0.5 + txtEastMin
   AveNorthSearch = (txtNorthMax - txtNorthMin) * 0.5 + txtNorthMin
   
   '___________________________________________________
   'now do the actual searching
     
   .cmdSearchReport.Enabled = False 'disenable search button
                                   'during duration of search
   Call SearchDataBase

   If ArcDump Then Exit Sub 'don't show search results

   '_______________________________________________________
   'set flag that search report form is visible
   PicSum = True
   If numReport& > 0 Then
      ret = ShowWindow(GDReportfrm.hWnd, SW_MAXIMIZE)
      End If

   Exit Sub
   
End With

errhand:
   Screen.MousePointer = vbDefault
   If Searching Then
     Screen.MousePointer = vbDefault
     GDMDIform.prbSearch.Enabled = False
     GDMDIform.prbSearch.Visible = False
     GDMDIform.cmdCancelSearch.Enabled = False
     GDMDIform.cmdCancelSearch.Visible = False
     GDMDIform.StatusBar1.Panels(1) = sEmpty
     
    'stop the animation if activated
     GDSearchfrm.picAnimation.Visible = False
     GDSearchfrm.anmSearch.Visible = False
     GDReportfrm.picAnimation.Visible = False
     GDReportfrm.anmReport.Visible = False
     GDReportfrm.anmReport.Stop
     GDSearchfrm.anmSearch.Stop
     
     'enable search options
     GDSearchfrm.cmdSearchReport.Enabled = True
     GDSearchfrm.cmdCancel.Enabled = True
     GDSearchfrm.cmdMinimize.Enabled = True
     Searching = False
     
     numHighlighted& = 0
     numReport& = 0
     End If
   If PicSum Then
      Unload GDReportfrm
      End If
   If ArcDump Then numArc& = -100
   MsgBox "Encountered error #: " & Err.Number & vbLf & Err.Description & vbLf _
       & "Error occured in module: GDSearchfrm.cmdSearchReport." & vbLf & sEmpty, _
       vbCritical + vbOKOnly, "MapDigitizer"
       
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSearchReport_MouseMove
' DateTime  : 9/30/2002 16:57
' Purpose   : Highlight button face as mouse moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdSearchReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdSearchReport.BackColor = &H8000000F Then
        cmdBack.BackColor = &H8000000F
        cmdNext.BackColor = &H8000000F
        cmdSearchReport.BackColor = &HC0C0C0
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = "Execute search"
End Sub

Private Sub cmdWells_Click()
  
   If chkWells.value = vbChecked Then
      chkOutcroppings.Enabled = False
      chkOutcroppings.value = vbUnchecked
      chkWells.Enabled = True
      chkAllWells.Enabled = True
      chkJustCores.Enabled = True
      chkJustCuttings.Enabled = True
      txtLimup.Enabled = True
      udLimup.Enabled = True
      txtLimdo.Enabled = True
      udLimdo.Enabled = True
      lblLimdo.Enabled = True
      lblLimup.Enabled = True
   Else
      chkOutcroppings.Enabled = False
      chkOutcroppings.value = vbUnchecked
      chkJustCuttings.Enabled = True
      chkJustCores.Enabled = True
      chkAllWells.Enabled = True
      chkAllWells.value = vbChecked
      chkWells.Enabled = True
      chkWells.value = vbChecked
      txtLimup.Enabled = True
      udLimup.Enabled = True
      txtLimdo.Enabled = True
      udLimdo.Enabled = True
      lblLimdo.Enabled = True
      lblLimup.Enabled = True
      End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : cmdWells_MouseMove
' DateTime  : 9/30/2002 16:38
' Purpose   : Change button color as cursor moves over it
'---------------------------------------------------------------------------------------
'
Private Sub cmdWells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If cmdWells.BackColor = &H8000000F Then
      cmdOutcroppings.BackColor = &H8000000F
      cmdWells.BackColor = &HC0C0C0
      cmdAllSources.BackColor = &H8000000F
      End If
   GDMDIform.StatusBar1.Panels(1) = "Search all wells"
End Sub


Private Sub cmdEditdb2_Click()
    'display the SQL fossil category string for the
    'scanned database.
    
    With GDSearchfrm
    
        .txtSQLdb2.Enabled = True
        
        Createdb2SqlFossil
           
        txtSQLdb2.Text = strSqlCategory
    
    End With

End Sub

Private Sub Combo1_click()
   
   ShowSQL 'display SQL clauses if flagged

End Sub

Private Sub Combo15_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub Combo2_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub Combo3_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub Combo4_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub Combo5_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub Combo6_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub Combo7_Click()
      
   ShowSQL 'display SQL clauses if flagged
   
End Sub

Private Sub Form_Load()
   'This is the search wizard form
   'It enables the performing of searches over the database
   
   GDSearchfrm.Top = 0   'set default starting position
   GDSearchfrm.Left = 0

   cmdBack.Enabled = False
   
   lblXMin = lblX
   lblXMax = lblX
   lblYMin = LblY
   lblYMax = LblY
   
   'load boolean combo boxes for fossil types
   Combo1.Clear
   Combo1.AddItem "OR"
   Combo1.AddItem "AND"
   Combo1.AddItem "AND NOT"
   Combo1.Text = Combo1.List(0)
   
   Combo2.Clear
   Combo2.AddItem "OR"
   Combo2.AddItem "AND"
   Combo2.AddItem "AND NOT"
   Combo2.Text = Combo1.List(0)
   
   Combo3.Clear
   Combo3.AddItem "OR"
   Combo3.AddItem "AND"
   Combo3.AddItem "AND NOT"
   Combo3.Text = Combo1.List(0)
   
   Combo4.Clear
   Combo4.AddItem "OR"
   Combo4.AddItem "AND"
   Combo4.AddItem "AND NOT"
   Combo4.Text = Combo1.List(0)
   
   Combo5.Clear
   Combo5.AddItem "OR"
   Combo5.AddItem "AND"
   Combo5.AddItem "AND NOT"
   Combo5.Text = Combo1.List(0)
   
   Combo6.Clear
   Combo6.AddItem "OR"
   Combo6.AddItem "AND"
   Combo6.AddItem "AND NOT"
   Combo6.Text = Combo1.List(0)
   
   Combo7.Clear
   Combo7.AddItem "OR"
   Combo7.AddItem "AND"
   Combo7.AddItem "AND NOT"
   Combo7.Text = Combo1.List(0)
   
   Combo15.Clear
   Combo15.AddItem "OR"
   Combo15.AddItem "AND"
   Combo15.AddItem "AND NOT"
   Combo15.Text = Combo1.List(0)
   
   'load boolean combo boxes for fossil species
   Combo8.Clear
   Combo8.AddItem "OR"
   Combo8.AddItem "AND"
   Combo8.Text = Combo1.List(0)
   
   Combo9.Clear
   Combo9.AddItem "OR"
   Combo9.AddItem "AND"
   Combo9.Text = Combo1.List(0)
   
   Combo10.Clear
   Combo10.AddItem "OR"
   Combo10.AddItem "AND"
   Combo10.Text = Combo1.List(0)
   
   Combo11.Clear
   Combo11.AddItem "OR"
   Combo11.AddItem "AND"
   Combo11.Text = Combo1.List(0)
   
   Combo12.Clear
   Combo12.AddItem "OR"
   Combo12.AddItem "AND"
   Combo12.Text = Combo1.List(0)
   
   Combo13.Clear
   Combo13.AddItem "OR"
   Combo13.AddItem "AND"
   Combo13.Text = Combo1.List(0)
   
   Combo14.Clear
   Combo14.AddItem "OR"
   Combo14.AddItem "AND"
   Combo14.Text = Combo1.List(0)
   
   'load boolean combo boxes for name searches
   cmbName1.Clear
   cmbName1.AddItem "OR"
   cmbName1.AddItem "AND"
   cmbName1.AddItem "AND NOT"
   cmbName1.Text = Combo1.List(0)
   
   cmbName2.Clear
   cmbName2.AddItem "OR"
   cmbName2.AddItem "AND"
   cmbName2.AddItem "AND NOT"
   cmbName2.Text = Combo1.List(0)
   
   If SearchDBs% <> 2 Then
      Combo15.Visible = True
      imShekef.Visible = True
      chkShekef.Visible = True
      lblShekef.Visible = True
      End If
      
   'SQL syntax edit boxes
   txtSQLdb1.Text = sEmpty
   txtSQLdb2.Text = sEmpty
   If linked Then txtSQLdb1.Enabled = True
   If linkedOld Then txtSQLdb2.Enabled = True
   
   Screen.MousePointer = vbHourglass
   
   'load Zone genera and species names into Zone list boxes
   LoadZones
   
   'load Fossil genera and species names into Fossil list boxes
   'from the Fossil dictionaries
   LoadFossilDic
   
   'load formation names in lstFormation
   cmbLoadFormations
   
   'load geolgic dates into TreeView
   LoadSearchTree
   
   'load Client,Date,Project,Company,Analyst list boxes
   LoadTheRest
   
   'load combo boxes for date searches
   cmbEPre.AddItem "Lower"
   cmbEPre.AddItem sEmpty
   cmbEPre.AddItem "Upper"
   cmbEPre.AddItem "Middle"
   cmbEPre.ListIndex = 1
   cmbLPre.AddItem "Lower"
   cmbLPre.AddItem sEmpty
   cmbLPre.AddItem "Upper"
   cmbLPre.AddItem "Middle"
   cmbLPre.ListIndex = 1
   cmbEQuest.AddItem sEmpty
   cmbEQuest.AddItem " ?"
   cmbEQuest.ListIndex = 0
   cmbLQuest.AddItem sEmpty
   cmbLQuest.AddItem " ?"
   cmbLQuest.ListIndex = 0
    
   SearchVis = True
   
   If SearchDBs% = 1 Or SearchDBs% = 3 Then
      'load arrays for searching old paleontological data base
      LoadOldDbArrays
      End If
   
   RangeOfDates = True 'Searching over range of dates is the default
   
   sFormSearch = sEmpty 'zero search string for Form table
   sFormSearchOld = sEmpty 'zero search string for Old database Project numbers
   sAnalystSearch = sEmpty 'zero search string for Analysts
   
   StopSearch = False
   
   'enable search buttons
   With GDMDIform
     For i& = 18 To 27
        .Toolbar1.Buttons(i&).Enabled = True
     Next i&
     .Toolbar1.Refresh
     .mnuMapSearch.Enabled = True
     .mnuAllSources.Enabled = True
     .mnuWells.Enabled = True
     .mnuOutcroppings.Enabled = True
     .mnuFossils.Enabled = True
     .mnuFormations.Enabled = True
     .mnuDates.Enabled = True
     .mnuClients.Enabled = True
     
   End With
   
   Screen.MousePointer = vbDefault
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_MouseMove
' DateTime  : 9/30/2002 17:00
' Purpose   : Return control button face colors to normal as mouse leaves them
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAll.BackColor <> &H8000000F Or _
        cmdAllAnd.BackColor <> &H8000000F Or _
        cmdAllOr.BackColor <> &H8000000F Or _
        cmdClearAll.BackColor <> &H8000000F Or _
        cmdEditSql.BackColor <> &H8000000F Or _
        cmdNext.BackColor <> &H8000000F Or _
        cmdBack.BackColor <> &H8000000F Or _
        cmdSearchReport.BackColor <> &H8000000F Then
        
        cmdAllAnd.BackColor = &H8000000F
        cmdAll.BackColor = &H8000000F
        cmdAllOr.BackColor = &H8000000F
        cmdClearAll.BackColor = &H8000000F
        cmdEditSql.BackColor = &H8000000F
        cmdNext.BackColor = &H8000000F
        cmdBack.BackColor = &H8000000F
        cmdSearchReport.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = sEmpty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   If (GeoMap Or TopoMap) And Not CloseSearchWizard And Not Closing Then
      response = MsgBox("Tip: if you are still planning to search over" & vbLf & _
               "the maps, to save time don't close the search wizard." & vbLf & _
               "Just minimize the search wizard to bring the maps to the foreground," & vbLf & _
               "If the maps are in the foreground, you can drag over them" & vbLf & _
               "to define new coordinate boundaries." & vbLf & vbLf & _
               "Do you still want to close the search wizard?", _
               vbYesNoCancel + vbInformation, "MapDigitizer")
      If response = vbYes Then
         CloseSearchWizard = True
      ElseIf response = vbNo Or vbCancel Then
         Cancel = True
         Exit Sub
         End If
   Else
      CloseSearchWizard = False
      End If
      
   'release memory from Fossil Name arrays
   'ReDim sArrConoNames(0)
   'ReDim sArrDiatomNames(0)
   'ReDim sArrForamNames(0)
   'ReDim sArrMegaNames(0)
   'ReDim sArrNanoNames(0)
   'ReDim sArrOstraNames(0)
   'ReDim sArrPalyNames(0)
   '
   'ReDim lArrCono(0)
   'numFosCono = 0
   'ReDim lArrDiatom(0)
   'numFosDiatom = 0
   'ReDim lArrForam(0)
   'numFosForam = 0
   'ReDim lArrMega(0)
   'numFosMega = 0
   'ReDim lArrNano(0)
   'numFosNano = 0
   'ReDim lArrOstra(0)
   'numFosOstra = 0
   'ReDim lArrPaly(0)
   'numFosPaly = 0
   
   'release memory from Fossil Zone arrays
   ReDim sArrConodonta(1, 0)
   ReDim sArrDiatom(1, 0)
   ReDim sArrForaminifera(1, 0)
   ReDim sArrMegafauna(1, 0)
   ReDim sArrNannoplankton(1, 0)
   ReDim sArrOstracoda(1, 0)
   ReDim sArrPalynology(1, 0)
   
   numN07 = 0: numN10 = 0: numN11 = 0: numOldDates = 0
   
   Unload Me
   Set GDSearchfrm = Nothing
   
   SearchVis = False
   
   With GDMDIform
   
     'disenable search buttons
     For i& = 17 To 33
        .Toolbar1.Buttons(i&).Enabled = False
     Next i&
     .mnuMapSearch.Enabled = False
     .mnuAllSources.Enabled = False
     .mnuWells.Enabled = False
     .mnuOutcroppings.Enabled = False
     .mnuFossils.Enabled = False
     .mnuFormations.Enabled = False
     .mnuDates.Enabled = False
     .mnuClients.Enabled = False
     .mnuPrintReport.Enabled = False
     .mnuSave.Enabled = False
     .mnuArcGIS.Enabled = False
     .mnuGoogle.Enabled = False
     'renable some of the button
     .Toolbar1.Buttons(28).Enabled = True 'reports and previeiwng
     If PicSum Then 'reenable some buttons
        .Toolbar1.Buttons(29).Enabled = True 'print previewing
        .mnuPrintReport.Enabled = True
        .Toolbar1.Buttons(30).Enabled = True 'saving
        .mnuSave.Enabled = True
        .Toolbar1.Buttons(32).Enabled = True 'ArcMap
        .mnuArcGIS.Enabled = True
        .Toolbar1.Buttons(33).Enabled = True 'Google
        .mnuGoogle.Enabled = True
        End If
        
     If GeoMap Then
        .Toolbar1.Buttons(2).Enabled = True 'Input Geo button
    '    buttonstate(18) = 0
        buttonstate(2) = 1
        .Toolbar1.Buttons(2).value = tbrPressed
        End If
        
    ' If GeoMap Or TopoMap Then
    '    .mnuGotoRetrieve.Enabled = False
    '    End If
       
   End With
   
   Closing = False
   
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Frame1_MouseMove
' DateTime  : 9/30/2002 17:30
' Purpose   : Restore button face colors to normal when mouse leaves them
'---------------------------------------------------------------------------------------
'
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdFullExtent.BackColor <> &H8000000F Or cmdClear.BackColor <> _
         &H8000000F Then
        cmdFullExtent.BackColor = &H8000000F
        cmdClear.BackColor = &H8000000F
        End If
     GDMDIform.StatusBar1.Panels(1) = sEmpty

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Frame10_MouseMove
' DateTime  : 9/30/2002 17:26
' Purpose   : Restore normal button face colors to buttons as mouse leaves them
'---------------------------------------------------------------------------------------
'
Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAllForm.BackColor <> &H8000000F Or cmdClearForm.BackColor <> _
         &H8000000F Then
        cmdAllForm.BackColor = &H8000000F
        cmdClearForm.BackColor = &H8000000F
        End If
     GDMDIform.StatusBar1.Panels(1) = sEmpty
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Frame2_MouseMove
' DateTime  : 9/30/2002 16:47
' Purpose   : Restore Button Faces to Normal Color
'---------------------------------------------------------------------------------------
'
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdOutcroppings.BackColor <> &H8000000F Or cmdWells.BackColor <> _
        &H8000000F Or cmdAllSources.BackColor <> &H8000000F Then
        cmdOutcroppings.BackColor = &H8000000F
        cmdWells.BackColor = &H8000000F
        cmdAllSources.BackColor = &H8000000F
        End If
     GDMDIform.StatusBar1.Panels(1) = sEmpty
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Frame3_MouseMove
' DateTime  : 9/30/2002 17:17
' Purpose   : Restore normal button face colors as mouse leaves them
'---------------------------------------------------------------------------------------
'
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAll.BackColor <> &H8000000F Or _
        cmdAllAnd.BackColor <> &H8000000F Or _
        cmdAllOr.BackColor <> &H8000000F Or _
        cmdClearAll.BackColor <> &H8000000F Or _
        cmdEditSql.BackColor <> &H8000000F Or _
        cmdNext.BackColor <> &H8000000F Or _
        cmdBack.BackColor <> &H8000000F Or _
        cmdSearchReport.BackColor <> &H8000000F Then
        
        cmdAllAnd.BackColor = &H8000000F
        cmdAll.BackColor = &H8000000F
        cmdAllOr.BackColor = &H8000000F
        cmdClearAll.BackColor = &H8000000F
        cmdEditSql.BackColor = &H8000000F
        cmdNext.BackColor = &H8000000F
        cmdBack.BackColor = &H8000000F
        cmdSearchReport.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = sEmpty

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Frame9_MouseMove
' DateTime  : 9/30/2002 17:17
' Purpose   : Restore normal button face colors as mouse leaves them
'---------------------------------------------------------------------------------------
'
Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If cmdAll.BackColor <> &H8000000F Or _
        cmdAllAnd.BackColor <> &H8000000F Or _
        cmdAllOr.BackColor <> &H8000000F Or _
        cmdClearAll.BackColor <> &H8000000F Or _
        cmdEditSql.BackColor <> &H8000000F Or _
        cmdNext.BackColor <> &H8000000F Or _
        cmdBack.BackColor <> &H8000000F Or _
        cmdSearchReport.BackColor <> &H8000000F Then
        
        cmdAllAnd.BackColor = &H8000000F
        cmdAll.BackColor = &H8000000F
        cmdAllOr.BackColor = &H8000000F
        cmdClearAll.BackColor = &H8000000F
        cmdEditSql.BackColor = &H8000000F
        cmdNext.BackColor = &H8000000F
        cmdBack.BackColor = &H8000000F
        cmdSearchReport.BackColor = &H8000000F
        End If
     If Not Searching Then GDMDIform.StatusBar1.Panels(1) = sEmpty
End Sub

Private Sub lstConodonta_GotFocus()
   lstConodonta.Height = 1290
   BringWindowToTop (lstConodonta.hWnd)
End Sub

Private Sub lstConodonta_LostFocus()
   lstConodonta.Height = 645
End Sub

Private Sub lstConodsdic_GotFocus()
   lstConodsdic.Height = 1290
   BringWindowToTop (lstConodsdic.hWnd)
End Sub

Private Sub lstConodsdic_LostFocus()
   lstConodsdic.Height = 645
End Sub

Private Sub lstDiatom_GotFocus()
   lstDiatom.Height = 1290
   BringWindowToTop (lstDiatom.hWnd)
End Sub

Private Sub lstDiatom_LostFocus()
   lstDiatom.Height = 645
End Sub

Private Sub lstDiatomsdic_GotFocus()
   lstDiatomsdic.Height = 1290
   BringWindowToTop (lstDiatomsdic.hWnd)
End Sub

Private Sub lstDiatomsdic_LostFocus()
   lstDiatomsdic.Height = 645
End Sub

Private Sub lstForaminifera_GotFocus()
   lstForaminifera.Height = 1290
   BringWindowToTop (lstForaminifera.hWnd)
End Sub

Private Sub lstForaminifera_LostFocus()
   lstForaminifera.Height = 645
End Sub

Private Sub lstForamsdic_GotFocus()
   lstForamsdic.Height = 1290
   BringWindowToTop (lstForamsdic.hWnd)
End Sub

Private Sub lstForamsdic_LostFocus()
   lstForamsdic.Height = 645
End Sub

Private Sub lstMegadic_GotFocus()
   lstMegadic.Height = 1290
   BringWindowToTop (lstMegadic.hWnd)
End Sub

Private Sub lstMegadic_LostFocus()
   lstMegadic.Height = 645
End Sub

Private Sub lstMegafauna_GotFocus()
   lstMegafauna.Height = 1290
   BringWindowToTop (lstMegafauna.hWnd)
End Sub

Private Sub lstMegafauna_LostFocus()
   lstMegafauna.Height = 645
End Sub

Private Sub lstNannoplankton_GotFocus()
   lstNannoplankton.Height = 1290
   BringWindowToTop (lstNannoplankton.hWnd)
End Sub

Private Sub lstNannoplankton_LostFocus()
   lstNannoplankton.Height = 645
End Sub

Private Sub lstNanodic_GotFocus()
   lstNanodic.Height = 1290
   BringWindowToTop (lstNanodic.hWnd)
End Sub

Private Sub lstNanodic_LostFocus()
   lstNanodic.Height = 645
End Sub

Private Sub lstOstracoda_GotFocus()
   lstOstracoda.Height = 1290
   BringWindowToTop (lstOstracoda.hWnd)
End Sub

Private Sub lstOstracoda_LostFocus()
   lstOstracoda.Height = 645
End Sub

Private Sub lstOstracoddic_GotFocus()
   lstOstracoddic.Height = 1290
   BringWindowToTop (lstOstracoddic.hWnd)
End Sub

Private Sub lstOstracoddic_LostFocus()
   lstOstracoddic.Height = 645
End Sub

Private Sub lstPalyndic_GotFocus()
   lstPalyndic.Height = 1290
   BringWindowToTop (lstPalyndic.hWnd)
End Sub

Private Sub lstPalyndic_LostFocus()
   lstPalyndic.Height = 645
End Sub

Private Sub lstPalynology_GotFocus()
   lstPalynology.Height = 1290
   BringWindowToTop (lstPalynology.hWnd)
End Sub

Private Sub lstPalynology_LostFocus()
   lstPalynology.Height = 645
End Sub

Private Sub optExact_Click()
   RangeOfDates = False
End Sub

Private Sub optName1_Click()
   PasteName1 = True
   PasteName2 = False
End Sub

Private Sub optName2_Click()
   PasteName1 = False
   PasteName2 = True
End Sub

Private Sub optRange_Click()
   RangeOfDates = True
End Sub

Private Sub tbSearch_Click(PreviousTab As Integer)
   Select Case tbSearch.Tab
      Case 0 'maps
        stepsearch& = 0
        cmdBack.Enabled = False
        cmdNext.Enabled = True
      Case 1 'wells/outcroppings
        stepsearch& = 1
        cmdBack.Enabled = True
        cmdNext.Enabled = True
      Case 2 'fossils
        stepsearch& = 2
        cmdBack.Enabled = True
        cmdNext.Enabled = True
      Case 3 'formations
        stepsearch& = 3
        cmdBack.Enabled = True
        cmdNext.Enabled = True
      Case 4 'dates
        stepsearch& = 4
        cmdBack.Enabled = True
        cmdNext.Enabled = True
      Case 5 'clients
        stepsearch& = 5
        cmdBack.Enabled = True
        cmdNext.Enabled = False
     Case Else
    End Select
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TreeView1_NodeClick
' DateTime  : 9/30/2002 19:32
' Purpose   : Fill Earlier/Later age date edit boxes with clicked date
'---------------------------------------------------------------------------------------
'
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  If txtEarlier.Enabled = True Then txtEarlier = Node.Key
  If txtLater.Enabled = True Then txtLater = Node.Key
End Sub

Private Sub LoadPaleoArrays()
   'load the paleo names into the search arrays
   
   If lstConodonta.Enabled = True Then
      Call LoadZoneArrays(lstConodonta, sArrConodonta)
      End If
   
   If lstDiatom.Enabled = True Then
      Call LoadZoneArrays(lstDiatom, sArrDiatom)
      End If
   
   If lstForaminifera.Enabled = True Then
      Call LoadZoneArrays(lstForaminifera, sArrForaminifera)
      End If
   
   If lstMegafauna.Enabled = True Then
      Call LoadZoneArrays(lstMegafauna, sArrMegafauna)
      End If
   
   If lstNannoplankton.Enabled = True Then
      Call LoadZoneArrays(lstNannoplankton, sArrNannoplankton)
      End If
   
   If lstOstracoda.Enabled = True Then
      Call LoadZoneArrays(lstOstracoda, sArrOstracoda)
      End If
   
   If lstPalynology.Enabled = True Then
      Call LoadZoneArrays(lstPalynology, sArrPalynology)
      End If
      
End Sub
Private Sub loadFossilArrays()
   'load the selected fossil names into the search arrays
   
'   Dim i&, pos1&, pos2&, n&
   
   If lstConodsdic.Enabled = True Then
      Call LoadFosNameArrays(numFosCono, lArrCono, lstConodsdic)
      End If
   
   If lstDiatomsdic.Enabled = True Then
      Call LoadFosNameArrays(numFosDiatom, lArrDiatom, lstDiatomsdic)
      End If
   
   If lstForamsdic.Enabled = True Then
      Call LoadFosNameArrays(numFosForam, lArrForam, lstForamsdic)
      End If
   
   If lstMegadic.Enabled = True Then
      Call LoadFosNameArrays(numFosMega, lArrMega, lstMegadic)
      End If
   
   If lstNanodic.Enabled = True Then
      Call LoadFosNameArrays(numFosNano, lArrNano, lstNanodic)
      End If
   
   If lstOstracoddic.Enabled = True Then
      Call LoadFosNameArrays(numFosOstra, lArrOstra, lstOstracoddic)
      End If
   
   If lstPalyndic.Enabled = True Then
      Call LoadFosNameArrays(numFosPaly, lArrPaly, lstPalyndic)
      End If
      
         
End Sub

Sub LoadZoneArrays(LstBox, sArr() As String)
     'load selected zones into string array
   
      Dim i&, pos1&, pos2&, N&
      
      'zero search array before start of new search
      For i& = 1 To LstBox.ListCount
         sArr(1, i&) = sEmpty
      Next i&

      For i& = 1 To LstBox.ListCount
         If LstBox.Selected(i& - 1) Then
            pos1& = InStr(LstBox.List(i& - 1), "id:")
            pos2& = InStr(pos1&, LstBox.List(i& - 1), ")")
            N& = val(Mid$(LstBox.List(i& - 1), pos1& + 4, pos2& - pos1& - 4))
            sArr(1, N&) = Trim$(Mid$(LstBox.List(i& - 1), 1, pos1& - 2))
            End If
      Next i&
      
End Sub

Sub LoadFosNameArrays(numFos As Long, lArr() As Long, LstBox)
      'load fossil name arrays
      
      Dim i&, pos1&, pos2&, N&
      
      numFos = 0
      ReDim lArr(0)
      For i& = 1 To LstBox.ListCount
         pos1& = InStr(LstBox.List(i& - 1), "id:")
         pos2& = InStr(pos1&, LstBox.List(i& - 1), ")")
         N& = val(Mid$(LstBox.List(i& - 1), pos1& + 4, pos2& - pos1& - 4))
         If LstBox.Selected(i& - 1) Then
            numFos = numFos + 1
            ReDim Preserve lArr(numFos)
            lArr(numFos - 1) = N&
            End If
      Next i&
      
End Sub
Sub ShowSQL()
   'display the Fossil Types SQL clauses
   
   If lineSQL.Visible = True Then 'refresh the displayed SQL
      If txtSQLdb1.Enabled = True Then
         If linked Then cmdEditdb1.value = True
      Else
         If linked Then cmdEditdb1.value = True
         txtSQLdb1.Enabled = False
         End If
      
      If txtSQLdb2.Enabled = True Then
         If linkedOld Then cmdEditdb2.value = True
      Else
         If linkedOld Then cmdEditdb2.value = True
         txtSQLdb2.Enabled = False
         End If
         
      End If

End Sub
