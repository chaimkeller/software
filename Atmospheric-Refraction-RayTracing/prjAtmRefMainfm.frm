VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form prjAtmRefMainfm 
   Caption         =   "Ray Tracing Utility"
   ClientHeight    =   12090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17835
   Icon            =   "prjAtmRefMainfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12090
   ScaleWidth      =   17835
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab TabRef 
      Height          =   11775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   20770
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   529
      TabCaption(0)   =   "Parameters"
      TabPicture(0)   =   "prjAtmRefMainfm.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paramfrm"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Temperature Profile"
      TabPicture(1)   =   "prjAtmRefMainfm.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Tempfrm"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pressure Profile"
      TabPicture(2)   =   "prjAtmRefMainfm.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Pressfrm"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Suns"
      TabPicture(3)   =   "prjAtmRefMainfm.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Sunsfrm"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Ray Tracing"
      TabPicture(4)   =   "prjAtmRefMainfm.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Rayfrm"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Transfer Curve"
      TabPicture(5)   =   "prjAtmRefMainfm.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Tcfrm"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Terrestrial Refraction"
      TabPicture(6)   =   "prjAtmRefMainfm.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Terrfrm"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "van der Werf"
      TabPicture(7)   =   "prjAtmRefMainfm.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "frmVDW"
      Tab(7).ControlCount=   1
      Begin VB.Frame frmVDW 
         Caption         =   "van der Werf graphics page"
         Height          =   10935
         Left            =   -74520
         TabIndex        =   118
         Top             =   600
         Width           =   16215
         Begin VB.PictureBox picVDW 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            Height          =   10095
            Left            =   480
            ScaleHeight     =   10035
            ScaleMode       =   0  'User
            ScaleWidth      =   15195
            TabIndex        =   119
            Top             =   480
            Width           =   15255
         End
      End
      Begin VB.Frame Sunsfrm 
         Caption         =   "Suns"
         Height          =   10695
         Left            =   -74400
         TabIndex        =   59
         Top             =   900
         Width           =   16215
         Begin VB.CommandButton cmdShowSuns 
            Caption         =   "Show Suns"
            Height          =   1155
            Left            =   6000
            TabIndex        =   60
            Top             =   4440
            Width           =   4395
         End
      End
      Begin VB.Frame Terrfrm 
         Caption         =   "Terrestrial Refraction Calculator"
         Height          =   10815
         Left            =   -74520
         TabIndex        =   6
         Top             =   780
         Width           =   16815
         Begin VB.Frame frmFit2 
            Caption         =   "Plot and Fit"
            Height          =   3375
            Left            =   11520
            TabIndex        =   216
            Top             =   6840
            Width           =   5055
            Begin VB.CommandButton cmdRefFiles_browse 
               Caption         =   "Browse"
               Height          =   255
               Left            =   4200
               TabIndex        =   235
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox txtRefFileDir 
               Height          =   285
               Left            =   3120
               TabIndex        =   234
               Text            =   "e:/atmref"
               Top             =   480
               Width           =   1095
            End
            Begin VB.CheckBox chkRefFiles_dip 
               Caption         =   "dip"
               Height          =   255
               Left            =   3360
               TabIndex        =   233
               ToolTipText     =   "Plot and calculate local dip angle as a function of the ray height using the refraction files"
               Top             =   1320
               Width           =   735
            End
            Begin VB.CheckBox chkRefFiles_Ref 
               Caption         =   "lev ref"
               Height          =   255
               Left            =   3360
               TabIndex        =   232
               ToolTipText     =   "plot and fit leveling refraction as a function of ray height using the refraction files"
               Top             =   960
               Width           =   735
            End
            Begin MSComCtl2.UpDown updwnVA 
               Height          =   255
               Left            =   2640
               TabIndex        =   231
               Top             =   1320
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               _Version        =   393216
               Max             =   48
               Min             =   -48
               Wrap            =   -1  'True
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtVA 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1800
               TabIndex        =   229
               Text            =   "0.0"
               Top             =   1320
               Width           =   840
            End
            Begin VB.CheckBox chkVA 
               Height          =   375
               Left            =   240
               TabIndex        =   227
               ToolTipText     =   "Fix the View Angle"
               Top             =   1200
               Width           =   255
            End
            Begin VB.CheckBox chkHgt 
               Height          =   255
               Left            =   240
               TabIndex        =   226
               ToolTipText     =   "Fix the Observer's height"
               Top             =   900
               Width           =   255
            End
            Begin MSComCtl2.UpDown updwnfit1 
               Height          =   255
               Left            =   2640
               TabIndex        =   225
               Top             =   480
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               _Version        =   393216
               Max             =   20
               Wrap            =   -1  'True
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown updwnhgtfit 
               Height          =   285
               Left            =   2640
               TabIndex        =   224
               Top             =   900
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtHgtFit"
               BuddyDispid     =   196622
               OrigLeft        =   4440
               OrigTop         =   480
               OrigRight       =   4695
               OrigBottom      =   735
               Increment       =   30
               Max             =   3000
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtHgtFit 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1800
               TabIndex        =   223
               Text            =   "0.0"
               ToolTipText     =   "Observer Height )m)"
               Top             =   900
               Width           =   840
            End
            Begin VB.TextBox txtFit1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1800
               TabIndex        =   221
               Text            =   "260"
               ToolTipText     =   "Temperature at observation place"
               Top             =   480
               Width           =   840
            End
            Begin VB.CheckBox chkFit1 
               Height          =   255
               Left            =   240
               TabIndex        =   219
               ToolTipText     =   "Fix temperature and height and plot and fit refraction vs. view angle"
               Top             =   480
               Width           =   255
            End
            Begin VB.CommandButton cmdPlotFit 
               Caption         =   "Plot and Fit"
               Height          =   255
               Left            =   480
               TabIndex        =   218
               Top             =   1800
               Width           =   4095
            End
            Begin VB.TextBox txtFitResults 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   480
               TabIndex        =   217
               Top             =   2280
               Width           =   4095
            End
            Begin VB.Label ErrorLbl 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   230
               Top             =   2760
               Width           =   4095
            End
            Begin VB.Label lblVA 
               Caption         =   "View Angle (deg.)"
               Height          =   255
               Left            =   480
               TabIndex        =   228
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lblHeight 
               Caption         =   "Height (m)"
               Height          =   255
               Left            =   600
               TabIndex        =   222
               Top             =   900
               Width           =   735
            End
            Begin VB.Label txtTK 
               Caption         =   "Temp (K)"
               Height          =   255
               Left            =   600
               TabIndex        =   220
               Top             =   480
               Width           =   735
            End
         End
         Begin MSChart20Lib.MSChart MSChartTR 
            Height          =   5895
            Left            =   360
            OleObjectBlob   =   "prjAtmRefMainfm.frx":0522
            TabIndex        =   209
            Top             =   480
            Width           =   15495
         End
         Begin VB.Frame frmTR 
            Caption         =   "Terrestrial Refraction"
            Height          =   3735
            Left            =   360
            TabIndex        =   100
            Top             =   6600
            Width           =   16455
            Begin VB.TextBox txtError 
               Height          =   285
               Left            =   7080
               TabIndex        =   213
               Text            =   "txtError"
               Top             =   3360
               Width           =   3855
            End
            Begin VB.TextBox txtAs 
               Height          =   285
               Left            =   7080
               TabIndex        =   212
               Text            =   "txtAs"
               Top             =   3000
               Width           =   3855
            End
            Begin VB.CommandButton cmdTest 
               Caption         =   "Plot"
               Height          =   375
               Left            =   10320
               TabIndex        =   211
               Top             =   1200
               Width           =   495
            End
            Begin VB.CheckBox chkUseDll 
               Caption         =   "Use iterative search via the dll"
               Height          =   255
               Left            =   7800
               TabIndex        =   210
               ToolTipText     =   "Instead of interpolating among the TR files, use faster and more accurate iterative search via the dll"
               Top             =   1920
               Value           =   1  'Checked
               Width           =   2535
            End
            Begin VB.TextBox txtStepT 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   6720
               TabIndex        =   208
               Text            =   "3"
               ToolTipText     =   "step size in temperature (K)"
               Top             =   1920
               Width           =   735
            End
            Begin VB.TextBox txtStepD1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   6720
               TabIndex        =   207
               Text            =   "5"
               ToolTipText     =   "step size in kms"
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox txtStepH2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   6720
               TabIndex        =   206
               Text            =   "10"
               ToolTipText     =   "step size in meters"
               Top             =   920
               Width           =   615
            End
            Begin VB.TextBox txtStepH1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   6720
               TabIndex        =   205
               Text            =   "10"
               ToolTipText     =   "Stjep size in meters"
               Top             =   360
               Width           =   615
            End
            Begin VB.Frame TRprogfrm 
               Caption         =   "Progress"
               Height          =   735
               Left            =   6960
               TabIndex        =   199
               Top             =   2280
               Visible         =   0   'False
               Width           =   3975
               Begin VB.PictureBox picProgBarTR 
                  Height          =   375
                  Left            =   120
                  ScaleHeight     =   315
                  ScaleWidth      =   3675
                  TabIndex        =   200
                  Top             =   240
                  Width           =   3735
               End
            End
            Begin VB.CommandButton cmdBrowseTR 
               Caption         =   "Browse"
               Height          =   375
               Left            =   9720
               TabIndex        =   196
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDirTR 
               Height          =   375
               Left            =   7680
               TabIndex        =   195
               Text            =   "Browse for TR files"
               ToolTipText     =   "Directory containing TR files"
               Top             =   360
               Width           =   2055
            End
            Begin MSComCtl2.UpDown updwnT12 
               Height          =   285
               Left            =   5640
               TabIndex        =   194
               Top             =   1920
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               Value           =   240
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtT12"
               BuddyDispid     =   196644
               OrigLeft        =   6480
               OrigTop         =   2040
               OrigRight       =   6735
               OrigBottom      =   2295
               Max             =   350
               Min             =   240
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtT12 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4680
               TabIndex        =   193
               Text            =   "320"
               Top             =   1920
               Width           =   960
            End
            Begin MSComCtl2.UpDown updwnD2 
               Height          =   285
               Left            =   5640
               TabIndex        =   192
               Top             =   1440
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtD2"
               BuddyDispid     =   196645
               OrigLeft        =   6360
               OrigTop         =   1560
               OrigRight       =   6615
               OrigBottom      =   1815
               Max             =   10000
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtD2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4680
               TabIndex        =   191
               Text            =   "0.0"
               Top             =   1440
               Width           =   960
            End
            Begin MSComCtl2.UpDown updwnH22 
               Height          =   285
               Left            =   5640
               TabIndex        =   190
               Top             =   920
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtH22"
               BuddyDispid     =   196646
               OrigLeft        =   6360
               OrigTop         =   960
               OrigRight       =   6615
               OrigBottom      =   1215
               Max             =   3000
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtH22 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4680
               TabIndex        =   189
               Text            =   "0.0"
               Top             =   920
               Width           =   960
            End
            Begin MSComCtl2.UpDown updwnH21 
               Height          =   285
               Left            =   5640
               TabIndex        =   188
               Top             =   360
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtH12"
               BuddyDispid     =   196647
               OrigLeft        =   6360
               OrigTop         =   360
               OrigRight       =   6615
               OrigBottom      =   615
               Max             =   3000
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtH12 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4680
               TabIndex        =   187
               Text            =   "0.0"
               Top             =   360
               Width           =   960
            End
            Begin VB.CheckBox chkTemp 
               Caption         =   "to"
               Height          =   375
               Left            =   4080
               TabIndex        =   186
               Top             =   1920
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.CheckBox chkDist 
               Caption         =   "to"
               Height          =   255
               Left            =   4080
               TabIndex        =   185
               Top             =   1440
               Width           =   855
            End
            Begin VB.CheckBox chkH2 
               Caption         =   "to"
               Height          =   255
               Left            =   4080
               TabIndex        =   184
               Top             =   930
               Width           =   615
            End
            Begin VB.CheckBox chkH1 
               Caption         =   "to"
               Height          =   255
               Left            =   4080
               TabIndex        =   183
               Top             =   360
               Width           =   735
            End
            Begin VB.Frame Atmfrm 
               Caption         =   "Other Atmospheres"
               Height          =   3255
               Left            =   11160
               TabIndex        =   172
               Top             =   240
               Width           =   3855
               Begin VB.Label Label20 
                  Caption         =   "winter"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   182
                  Top             =   2760
                  Width           =   1455
               End
               Begin VB.Label Label19 
                  Caption         =   "summer"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   181
                  Top             =   2520
                  Width           =   1815
               End
               Begin VB.Label Label18 
                  Caption         =   "Menat Atmospheres (view-angle: deg.)"
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
                  Left            =   360
                  TabIndex        =   180
                  Top             =   2280
                  Width           =   3375
               End
               Begin VB.Label Label17 
                  Caption         =   "6. US - standard"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   179
                  Top             =   1800
                  Width           =   2655
               End
               Begin VB.Label Label16 
                  Caption         =   "5. Suubartic winter"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   178
                  Top             =   1560
                  Width           =   2775
               End
               Begin VB.Label Label15 
                  Caption         =   "4. Subartic summer"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   177
                  Top             =   1320
                  Width           =   2775
               End
               Begin VB.Label Label14 
                  Caption         =   "3. Midlatitude winter:"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   176
                  Top             =   1080
                  Width           =   2775
               End
               Begin VB.Label Label13 
                  Caption         =   "2. Midlatitude summer"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   175
                  Top             =   840
                  Width           =   3015
               End
               Begin VB.Label Label12 
                  Caption         =   "1. Tropical:"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   174
                  Top             =   600
                  Width           =   2655
               End
               Begin VB.Label Label11 
                  Caption         =   "Lowtran Atmospheres (view angle - deg.)"
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
                  Left            =   240
                  TabIndex        =   173
                  Top             =   360
                  Width           =   3615
               End
            End
            Begin MSComCtl2.UpDown updwnTemp 
               Height          =   285
               Left            =   3600
               TabIndex        =   171
               Top             =   1920
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               Value           =   263
               OrigLeft        =   3720
               OrigTop         =   2040
               OrigRight       =   3975
               OrigBottom      =   2295
               Max             =   323
               Min             =   243
               Wrap            =   -1  'True
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtT11 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   170
               Text            =   "260"
               Top             =   1920
               Width           =   840
            End
            Begin VB.CommandButton cmdCalcTR 
               Caption         =   "Calculate and Plot"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   8040
               TabIndex        =   112
               Top             =   960
               Width           =   2055
            End
            Begin MSComCtl2.UpDown updwnTR_Dis 
               Height          =   285
               Left            =   3600
               TabIndex        =   109
               Top             =   1440
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               OrigLeft        =   3960
               OrigTop         =   1560
               OrigRight       =   4215
               OrigBottom      =   1815
               Max             =   500
               Wrap            =   -1  'True
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtD1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   108
               Text            =   "61.066"
               Top             =   1440
               Width           =   840
            End
            Begin MSComCtl2.UpDown updwnTR_Obs 
               Height          =   285
               Left            =   3600
               TabIndex        =   106
               Top             =   920
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               OrigLeft        =   3480
               OrigTop         =   960
               OrigRight       =   3735
               OrigBottom      =   1455
               Max             =   15000
               Wrap            =   -1  'True
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtH21 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   105
               Text            =   "959.2"
               Top             =   920
               Width           =   840
            End
            Begin MSComCtl2.UpDown updwnTR_OH 
               Height          =   285
               Left            =   3600
               TabIndex        =   103
               Top             =   360
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               OrigLeft        =   3720
               OrigTop         =   360
               OrigRight       =   3975
               OrigBottom      =   735
               Max             =   15000
               Wrap            =   -1  'True
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtH11 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   102
               Text            =   "756.5"
               Top             =   360
               Width           =   840
            End
            Begin VB.Label Label4 
               Caption         =   "Step:"
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
               Left            =   6120
               TabIndex        =   204
               Top             =   1920
               Width           =   495
            End
            Begin VB.Label Label3 
               Caption         =   "Step:"
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
               Left            =   6120
               TabIndex        =   203
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Step:"
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
               Left            =   6120
               TabIndex        =   202
               Top             =   920
               Width           =   495
            End
            Begin VB.Label lblStepH 
               Caption         =   "Step:"
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
               Left            =   6120
               TabIndex        =   201
               Top             =   360
               Width           =   495
            End
            Begin VB.Label LablOE 
               Caption         =   "View Angle with Ter Ref (old estimate) (deg.)"
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
               Left            =   480
               TabIndex        =   198
               Top             =   3240
               Width           =   6135
            End
            Begin VB.Label lblTemp 
               Caption         =   "Temperature (deg K)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   169
               Top             =   1920
               Width           =   2175
            End
            Begin VB.Label lblTR_Ref 
               Caption         =   "View Angle with Ter. Ref. estimate (deg.)"
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
               Left            =   480
               TabIndex        =   111
               Top             =   2880
               Width           =   6135
            End
            Begin VB.Label lblTR_Est 
               Caption         =   "View Angle w/o Ref. (deg.)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   110
               Top             =   2520
               Width           =   6135
            End
            Begin VB.Label lblTR_Dis 
               Caption         =   "Dist. to Obstruction (km)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   107
               Top             =   1440
               Width           =   2175
            End
            Begin VB.Label lblTR_Obs 
               Caption         =   "Obstruction's Height (m):"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   104
               Top             =   920
               Width           =   2295
            End
            Begin VB.Label lblTR_OH 
               Caption         =   "Observer's Height (m):"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   101
               Top             =   360
               Width           =   2175
            End
         End
      End
      Begin VB.Frame Tcfrm 
         Caption         =   "Transfer Curve"
         Height          =   10935
         Left            =   -74880
         TabIndex        =   5
         Top             =   780
         Width           =   16695
         Begin MSChart20Lib.MSChart MSCharttc 
            Height          =   9855
            Left            =   480
            OleObjectBlob   =   "prjAtmRefMainfm.frx":2878
            TabIndex        =   89
            Top             =   600
            Width           =   15735
         End
      End
      Begin VB.Frame Rayfrm 
         Caption         =   "Ray Tracing"
         Height          =   11055
         Left            =   -74640
         TabIndex        =   4
         Top             =   780
         Width           =   16455
         Begin MSComCtl2.FlatScrollBar VScroll1 
            Height          =   9855
            Left            =   16000
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   17383
            _Version        =   393216
            Orientation     =   1638400
         End
         Begin MSComCtl2.FlatScrollBar HScroll1 
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   10680
            Visible         =   0   'False
            Width           =   15615
            _ExtentX        =   27543
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin VB.ComboBox cmbAlt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8880
            TabIndex        =   19
            ToolTipText     =   "Choose an observed view angle to follow"
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox cmbSun 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8160
            TabIndex        =   18
            ToolTipText     =   "Choose Sun"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdRight 
            Height          =   375
            Left            =   6840
            Picture         =   "prjAtmRefMainfm.frx":4D14
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "translate right"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdDown 
            Height          =   375
            Left            =   6240
            Picture         =   "prjAtmRefMainfm.frx":5156
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "translate down"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdup 
            Height          =   375
            Left            =   5760
            Picture         =   "prjAtmRefMainfm.frx":5598
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Translate Up"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton cmdleft 
            Height          =   375
            Left            =   5280
            Picture         =   "prjAtmRefMainfm.frx":59DA
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "translate left"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtStartMult 
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
            Left            =   2640
            TabIndex        =   13
            ToolTipText     =   "Current Multiplication"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdSmaller 
            Height          =   375
            Left            =   1440
            Picture         =   "prjAtmRefMainfm.frx":5E1C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Zoom Out"
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdLarger 
            Height          =   375
            Left            =   360
            Picture         =   "prjAtmRefMainfm.frx":5F66
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Zoom In"
            Top             =   240
            Width           =   975
         End
         Begin VB.PictureBox picture1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000A&
            Height          =   9975
            Left            =   240
            ScaleHeight     =   661
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   1044
            TabIndex        =   9
            Top             =   720
            Width           =   15720
            Begin VB.PictureBox Picture2 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               Height          =   9975
               Left            =   0
               Picture         =   "prjAtmRefMainfm.frx":60B0
               ScaleHeight     =   661
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   1037
               TabIndex        =   22
               Top             =   0
               Width           =   15615
               Begin VB.PictureBox picRef 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   2175
                  Left            =   4560
                  ScaleHeight     =   145
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   209
                  TabIndex        =   23
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   3135
               End
            End
         End
      End
      Begin VB.Frame Pressfrm 
         Caption         =   "Pressure Profile"
         Height          =   10935
         Left            =   -74760
         TabIndex        =   3
         Top             =   780
         Width           =   16455
         Begin MSChart20Lib.MSChart MSChartPress 
            Height          =   9735
            Left            =   360
            OleObjectBlob   =   "prjAtmRefMainfm.frx":3882B
            TabIndex        =   90
            Top             =   600
            Width           =   15615
         End
      End
      Begin VB.Frame Tempfrm 
         Caption         =   "Temperature Profile"
         Height          =   10815
         Left            =   -74520
         TabIndex        =   2
         Top             =   900
         Width           =   16455
         Begin MSChart20Lib.MSChart MSChartTemp 
            Height          =   10095
            Left            =   360
            OleObjectBlob   =   "prjAtmRefMainfm.frx":3AB7F
            TabIndex        =   197
            Top             =   360
            Width           =   15615
         End
      End
      Begin VB.Frame paramfrm 
         Caption         =   "Parameters"
         Height          =   11175
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   17175
         Begin VB.Frame progressfrm2 
            Height          =   1215
            Left            =   14640
            TabIndex        =   161
            Top             =   8280
            Visible         =   0   'False
            Width           =   2400
            Begin VB.PictureBox picProgBar3 
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   2115
               TabIndex        =   164
               ToolTipText     =   "Progress for height loop"
               Top             =   720
               Width           =   2175
            End
            Begin VB.PictureBox picProgBar2 
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   2115
               TabIndex        =   163
               ToolTipText     =   "Progress for temperature loop"
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Frame frmT 
            Caption         =   "Terrestrialr Refract. Loop"
            Height          =   7095
            Left            =   14640
            TabIndex        =   146
            Top             =   240
            Width           =   2415
            Begin VB.CheckBox chkRefFile 
               Caption         =   "Write  Ref. File"
               Height          =   195
               Left            =   360
               TabIndex        =   215
               ToolTipText     =   "Create Total Refraction file to be used for fitting"
               Top             =   5640
               Width           =   1695
            End
            Begin VB.CheckBox chkSkipDone 
               Caption         =   "skip when file exists"
               Height          =   195
               Left            =   360
               TabIndex        =   214
               ToolTipText     =   "Don't repeart recorded ray tracing"
               Top             =   5400
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox chkTRef 
               Caption         =   "Record Ref vs Temp"
               Height          =   255
               Left            =   360
               TabIndex        =   168
               Top             =   3480
               Width           =   1815
            End
            Begin VB.TextBox txtDir 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   240
               TabIndex        =   167
               Top             =   4680
               Width           =   1935
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse forOutput Directory"
               Height          =   615
               Index           =   1
               Left            =   240
               TabIndex        =   166
               Top             =   3960
               Width           =   1935
            End
            Begin VB.CheckBox chkPause 
               Caption         =   "Pause"
               Height          =   375
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   165
               ToolTipText     =   "Press to pause calculations"
               Top             =   6480
               Width           =   1575
            End
            Begin VB.CheckBox chkAtmRefDll 
               Caption         =   "Use dll for calc."
               Height          =   195
               Left            =   360
               TabIndex        =   160
               ToolTipText     =   "Perform raytracing via dll instead of VB code"
               Top             =   5160
               Width           =   1455
            End
            Begin VB.TextBox txtSHgt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1200
               TabIndex        =   159
               Text            =   "30"
               Top             =   2880
               Width           =   975
            End
            Begin VB.TextBox txtEHgt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1200
               TabIndex        =   157
               Text            =   "3000"
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox txtBHgt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1200
               TabIndex        =   155
               Text            =   "0.0"
               Top             =   1920
               Width           =   975
            End
            Begin VB.TextBox txtTS 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1200
               TabIndex        =   153
               Text            =   "3"
               Top             =   1440
               Width           =   975
            End
            Begin VB.CommandButton cmpTLoop 
               Caption         =   "Start Calculation"
               Height          =   495
               Left            =   480
               TabIndex        =   151
               Top             =   5880
               Width           =   1575
            End
            Begin VB.TextBox txtET 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1200
               TabIndex        =   150
               Text            =   "320"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtST 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1200
               TabIndex        =   148
               Text            =   "260"
               Top             =   480
               Width           =   975
            End
            Begin VB.Label lblSH 
               Caption         =   "Step (m)"
               Height          =   255
               Left            =   200
               TabIndex        =   158
               Top             =   2880
               Width           =   735
            End
            Begin VB.Label lblEH 
               Caption         =   "End Hgt.(m)"
               Height          =   255
               Left            =   120
               TabIndex        =   156
               Top             =   2400
               Width           =   975
            End
            Begin VB.Label lblBH 
               Caption         =   "Beg Hgt (m)"
               Height          =   255
               Left            =   120
               TabIndex        =   154
               ToolTipText     =   "Beggining observer height (m)"
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label lblTS 
               Caption         =   "Step (K)"
               Height          =   255
               Left            =   120
               TabIndex        =   152
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label lblET 
               Caption         =   "End T (K)"
               Height          =   255
               Left            =   120
               TabIndex        =   149
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lblst 
               Caption         =   "Start T (K)"
               Height          =   255
               Left            =   120
               TabIndex        =   147
               Top             =   480
               Width           =   735
            End
         End
         Begin VB.Frame frmNewvdw 
            Caption         =   "van der Werf parameters"
            Height          =   6375
            Left            =   9480
            TabIndex        =   117
            ToolTipText     =   "Show van der Werf's info screens and plots from the original  program"
            Top             =   240
            Width           =   5055
            Begin VB.CheckBox chkVDW_Show 
               Caption         =   "Show VDW info screens"
               Height          =   255
               Left            =   1560
               TabIndex        =   145
               Top             =   5880
               Width           =   2295
            End
            Begin VB.CommandButton cmdVDW 
               Caption         =   "Perform van der Werf Ray Tracing"
               Height          =   375
               Left            =   840
               TabIndex        =   144
               ToolTipText     =   "Use Siebren van der Werf formulation for ray tracing"
               Top             =   5280
               Width           =   3375
            End
            Begin VB.TextBox txtNSTEPS 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   143
               Text            =   "500"
               Top             =   4320
               Width           =   1215
            End
            Begin VB.TextBox txtOBSLAT 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   141
               Text            =   "32.0"
               Top             =   3960
               Width           =   1215
            End
            Begin VB.TextBox txtBETAST 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3480
               TabIndex        =   139
               Text            =   "0.1"
               Top             =   3600
               Width           =   1215
            End
            Begin VB.TextBox txtBETAHI 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3480
               TabIndex        =   137
               Text            =   "2.5"
               Top             =   3240
               Width           =   1215
            End
            Begin VB.TextBox txtBETALO 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3480
               TabIndex        =   135
               Text            =   "-2.5"
               Top             =   2880
               Width           =   1215
            End
            Begin VB.TextBox txtRELHUM 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   133
               Text            =   "0"
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtPress0 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   131
               Text            =   "1013.25"
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox txtTHIGH 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   129
               Text            =   "400.0"
               Top             =   1800
               Width           =   1215
            End
            Begin VB.TextBox txtTLOW 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   127
               Text            =   "0.0"
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox txtHMAXT 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   125
               Text            =   "1000"
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox txtTGROUND 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   123
               Text            =   "288.15"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox txtHOBS 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3480
               TabIndex        =   121
               Text            =   "756.7"
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblInv 
               Alignment       =   2  'Center
               Caption         =   "Use the Ducting Layer controls to add an inversion and the Min Wavelength text box to change the wavelength"
               ForeColor       =   &H00000080&
               Height          =   375
               Left            =   360
               TabIndex        =   237
               Top             =   4800
               Width           =   4335
            End
            Begin VB.Label lbl13 
               Caption         =   "Number of steps up till HMAXT"
               Height          =   255
               Left            =   240
               TabIndex        =   142
               Top             =   4320
               Width           =   2775
            End
            Begin VB.Label lbl12 
               Caption         =   "Latitude of observer ((degrees)"
               Height          =   255
               Left            =   240
               TabIndex        =   140
               Top             =   3960
               Width           =   2415
            End
            Begin VB.Label lbl10 
               Caption         =   "Stepsize in apparent altitude (arcmin)"
               Height          =   255
               Left            =   240
               TabIndex        =   138
               Top             =   3600
               Width           =   2775
            End
            Begin VB.Label lbl9 
               Caption         =   "Highest apparent altitude (arcmin')"
               Height          =   255
               Left            =   240
               TabIndex        =   136
               Top             =   3240
               Width           =   2655
            End
            Begin VB.Label lbl8 
               Caption         =   "Lowest apparent altitude (arcmin)"
               Height          =   255
               Left            =   240
               TabIndex        =   134
               Top             =   2880
               Width           =   2535
            End
            Begin VB.Label lbl7 
               Caption         =   "Relative humidity (%) in troposphere"
               Height          =   255
               Left            =   240
               TabIndex        =   132
               Top             =   2520
               Width           =   2655
            End
            Begin VB.Label lbl6 
               Caption         =   "Atmospheric pressure (hPa) at h=0"
               Height          =   255
               Left            =   240
               TabIndex        =   130
               Top             =   2160
               Width           =   2655
            End
            Begin VB.Label lbl5 
               Caption         =   "highest value (K) for which to show T-profile"
               Height          =   255
               Left            =   240
               TabIndex        =   128
               Top             =   1800
               Width           =   3255
            End
            Begin VB.Label lbl4 
               Caption         =   "Show temperature profile from TLOW (K) till.."
               Height          =   255
               Left            =   240
               TabIndex        =   126
               Top             =   1440
               Width           =   3135
            End
            Begin VB.Label lbl3 
               Caption         =   "Max.Hgt (m) for curv. plot "
               Height          =   255
               Left            =   240
               TabIndex        =   124
               Top             =   1080
               Width           =   3015
            End
            Begin VB.Label lbl2 
               Caption         =   "Temperature (K) at height = 0"
               Height          =   255
               Left            =   240
               TabIndex        =   122
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label lblq 
               Caption         =   "Observer Eye Height (m)"
               Height          =   255
               Left            =   240
               TabIndex        =   120
               Top             =   360
               Width           =   1935
            End
         End
         Begin VB.Frame frmRef 
            Caption         =   "Refraction"
            Height          =   640
            Left            =   360
            TabIndex        =   91
            Top             =   8480
            Width           =   4695
            Begin VB.OptionButton optMenat 
               Caption         =   "Menat"
               Height          =   255
               Left            =   1680
               TabIndex        =   93
               ToolTipText     =   "Use Menat's wavelength dependent simplified refraction expression"
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton optCiddor 
               Caption         =   "Ciddor (Bruton)"
               Height          =   195
               Left            =   1680
               TabIndex        =   92
               ToolTipText     =   "Use Ciddor's wavelenght and humidity dependent expression"
               Top             =   160
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.Frame frmModel 
            Caption         =   "Other layered atmospheres"
            Height          =   6255
            Left            =   5280
            TabIndex        =   68
            Top             =   1440
            Width           =   3975
            Begin VB.OptionButton opt4 
               Caption         =   "Lowtran (Selby) Mid latitude summer"
               Height          =   195
               Left            =   240
               TabIndex        =   245
               Top             =   2120
               Width           =   3135
            End
            Begin VB.OptionButton opt3 
               Caption         =   "Lowtran (Selby) Midl latitude winter"
               Height          =   195
               Left            =   240
               TabIndex        =   244
               Top             =   1820
               Width           =   3255
            End
            Begin VB.CheckBox chkHgtProfile 
               Caption         =   "Make atmosphere follow height profile"
               Height          =   195
               Left            =   240
               TabIndex        =   243
               ToolTipText     =   "Ground hugging atmosphere"
               Top             =   5760
               Width           =   3495
            End
            Begin VB.CheckBox chkReNorm 
               Caption         =   "Renormalize hgts so that first height is zero"
               Height          =   240
               Left            =   240
               TabIndex        =   242
               Top             =   5400
               Width           =   3300
            End
            Begin VB.CheckBox chkMeters 
               Caption         =   "Elevations in meters"
               Height          =   195
               Left            =   1200
               TabIndex        =   238
               ToolTipText     =   "Check box if the atmosphere's elevation data is in meters"
               Top             =   4560
               Width           =   1815
            End
            Begin MSComCtl2.UpDown updwnHumid 
               Height          =   285
               Left            =   2760
               TabIndex        =   88
               Top             =   4880
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtHumid"
               BuddyDispid     =   196757
               OrigLeft        =   2760
               OrigTop         =   4680
               OrigRight       =   3015
               OrigBottom      =   4935
               Max             =   100
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtHumid 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2280
               TabIndex        =   87
               Text            =   "0"
               ToolTipText     =   "Percent Humidity (0-100%)"
               Top             =   4880
               Width           =   480
            End
            Begin VB.TextBox txtGroundPressure 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2280
               TabIndex        =   85
               Text            =   "1013.25"
               ToolTipText     =   "Atmospheric pressure on ground (mb)"
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox txtGroundTemp 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2280
               TabIndex        =   83
               Text            =   "298.8"
               ToolTipText     =   "Ground Temperature for Standard Atmosphere (degrees Kelvin)"
               Top             =   720
               Width           =   855
            End
            Begin MSComDlg.CommonDialog comdlgOther 
               Left            =   3360
               Top             =   1080
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton cmdOtherBrowse 
               Caption         =   "Browse"
               Height          =   255
               Left            =   3000
               TabIndex        =   78
               Top             =   4200
               Width           =   735
            End
            Begin VB.TextBox txtOther 
               Height          =   285
               Left            =   960
               TabIndex        =   77
               Top             =   4200
               Width           =   2055
            End
            Begin VB.OptionButton opt10 
               Caption         =   "Other"
               Height          =   255
               Left            =   240
               TabIndex        =   76
               ToolTipText     =   "Atmosphere file w/ 3 columns: Height(m), Temperature ((K), Pressure (mbar)"
               Top             =   4200
               Width           =   3255
            End
            Begin VB.OptionButton opt9 
               Caption         =   "Menat Eretz Yisroel summer atmosphere"
               Height          =   255
               Left            =   240
               TabIndex        =   75
               Top             =   3720
               Width           =   3375
            End
            Begin VB.OptionButton opt8 
               Caption         =   "Menat Eretz Yisroel winter atmosphere"
               Height          =   195
               Left            =   240
               TabIndex        =   74
               Top             =   3420
               Width           =   3015
            End
            Begin VB.OptionButton opt7 
               Caption         =   "Lowtran US Standard Atmosphere"
               Height          =   255
               Left            =   240
               TabIndex        =   73
               Top             =   3080
               Width           =   3015
            End
            Begin VB.OptionButton opt6 
               Caption         =   "Lowtran (Selby) subarctic summer"
               Height          =   195
               Left            =   240
               TabIndex        =   72
               Top             =   2780
               Width           =   3015
            End
            Begin VB.OptionButton opt5 
               Caption         =   "Lowtran (Sellby) subarctic winter"
               Height          =   195
               Left            =   240
               TabIndex        =   71
               Top             =   2440
               Width           =   2775
            End
            Begin VB.OptionButton opt2 
               Caption         =   "Lowtran (Selby) Tropical Atmosphere"
               Height          =   255
               Left            =   240
               TabIndex        =   70
               Top             =   1500
               Width           =   3495
            End
            Begin VB.OptionButton opt1 
               Caption         =   "Bruton's Thesis standard atmosphere"
               Height          =   255
               Left            =   240
               TabIndex        =   69
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label lblHumid 
               Caption         =   "Percent Humidity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   720
               TabIndex        =   86
               Top             =   4880
               Width           =   1575
            End
            Begin VB.Label lblGroundPressure 
               Caption         =   "Ground Pressure. (mb)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   84
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label lblGroundTemp 
               Caption         =   "Ground Temperature (K)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   600
               TabIndex        =   82
               Top             =   720
               Width           =   1575
            End
         End
         Begin VB.Frame progressfrm 
            Caption         =   "Progress"
            Height          =   855
            Left            =   14640
            TabIndex        =   67
            Top             =   7440
            Visible         =   0   'False
            Width           =   2415
            Begin VB.PictureBox picProgBar 
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   2115
               TabIndex        =   162
               ToolTipText     =   "Progress for view angle loop"
               Top             =   280
               Width           =   2175
            End
         End
         Begin VB.Frame frmSize 
            Caption         =   "Sun's picture size"
            Height          =   1215
            Left            =   5280
            TabIndex        =   61
            Top             =   240
            Width           =   3975
            Begin VB.TextBox txtYSize 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2280
               TabIndex        =   65
               Text            =   "0"
               ToolTipText     =   "Y pixel size (leave blank to let suns define it)"
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtXSize 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2280
               TabIndex        =   64
               Text            =   "1500"
               ToolTipText     =   "X pixel size (leave blank to define by suns)"
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblYSize 
               Caption         =   "Max View Angle (deg)"
               Height          =   300
               Left            =   480
               TabIndex        =   63
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label lblXSize 
               Caption         =   "X Pixel Size:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   62
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Ductingfrm 
            Caption         =   "Ducting layer"
            Height          =   1815
            Left            =   360
            TabIndex        =   48
            Top             =   9120
            Width           =   4695
            Begin MSComCtl2.UpDown UpDownEInv 
               Height          =   285
               Left            =   4081
               TabIndex        =   58
               Top             =   1320
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               Value           =   100
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtEInv"
               BuddyDispid     =   196781
               OrigLeft        =   4080
               OrigTop         =   1320
               OrigRight       =   4335
               OrigBottom      =   1575
               Max             =   99999
               Min             =   1
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   0   'False
            End
            Begin VB.TextBox txtEInv 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   2760
               TabIndex        =   57
               Text            =   "300"
               Top             =   1320
               Width           =   1320
            End
            Begin MSComCtl2.UpDown UpDownSInv 
               Height          =   285
               Left            =   4080
               TabIndex        =   55
               Top             =   960
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtSInv"
               BuddyDispid     =   196782
               OrigLeft        =   4200
               OrigTop         =   960
               OrigRight       =   4455
               OrigBottom      =   1215
               Max             =   99999
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   0   'False
            End
            Begin VB.TextBox txtSInv 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   2760
               TabIndex        =   54
               Text            =   "20"
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.UpDown UpDownDInv 
               Height          =   285
               Left            =   4081
               TabIndex        =   52
               Top             =   600
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               Value           =   5
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtDInv"
               BuddyDispid     =   196783
               OrigLeft        =   4440
               OrigTop         =   600
               OrigRight       =   4695
               OrigBottom      =   855
               Max             =   1000
               Min             =   1
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   0   'False
            End
            Begin VB.TextBox txtDInv 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   2760
               TabIndex        =   51
               Text            =   "50"
               Top             =   600
               Width           =   1320
            End
            Begin VB.CheckBox chkDucting 
               Caption         =   "Add Inversion Layer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1320
               TabIndex        =   49
               Top             =   120
               Width           =   2295
            End
            Begin VB.Label lblEInv 
               Caption         =   "End Height of Inversion (m)"
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
               Height          =   255
               Left            =   240
               TabIndex        =   56
               Top             =   1320
               Width           =   2535
            End
            Begin VB.Label lblSInv 
               Caption         =   "Start Height of Inversion (m)"
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
               Height          =   255
               Left            =   240
               TabIndex        =   53
               Top             =   960
               Width           =   2535
            End
            Begin VB.Label lblInvStep 
               Caption         =   "Inv. Lapse Rate (deg K/km)"
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
               Height          =   375
               Left            =   240
               TabIndex        =   50
               Top             =   600
               Width           =   2535
            End
         End
         Begin VB.Frame AtmModelfrm 
            Caption         =   "Atmospheric Model structure and Terrain"
            Height          =   2680
            Left            =   360
            TabIndex        =   42
            Top             =   5800
            Width           =   4695
            Begin VB.CommandButton cmdBrowseHgtProfile 
               Caption         =   "Browse"
               Height          =   255
               Left            =   3480
               TabIndex        =   240
               ToolTipText     =   "Browse for ground height profile (dist in m, hgt in m)"
               Top             =   2280
               Width           =   735
            End
            Begin VB.TextBox txtHgtProfile 
               Height          =   285
               Left            =   480
               TabIndex        =   239
               Text            =   "External terrain profile file path (m,m)"
               ToolTipText     =   "terrain profile to east (km,m)"
               Top             =   2280
               Width           =   2895
            End
            Begin MSComCtl2.UpDown UpDownHeightStepSize 
               Height          =   285
               Left            =   3360
               TabIndex        =   81
               Top             =   1880
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               Value           =   10
               BuddyControl    =   "txtHeightStepSize"
               BuddyDispid     =   196791
               OrigLeft        =   3120
               OrigTop         =   1920
               OrigRight       =   3375
               OrigBottom      =   2175
               Max             =   10000
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtHeightStepSize 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   79
               Text            =   "10"
               Top             =   1880
               Width           =   600
            End
            Begin VB.OptionButton OptionSelby 
               Caption         =   "Other layered Atm. models (define step size below and choose a model in ------------------>)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   240
               TabIndex        =   47
               ToolTipText     =   "Layered models"
               Top             =   1320
               Width           =   4215
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse"
               Height          =   375
               Index           =   0
               Left            =   3240
               TabIndex        =   46
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox TextExternal 
               Height          =   285
               Left            =   480
               TabIndex        =   45
               Top             =   880
               Width           =   2775
            End
            Begin VB.OptionButton OptionRead 
               Caption         =   "External Bruton format layer model"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   44
               Top             =   520
               Width           =   3495
            End
            Begin VB.OptionButton OptionLayer 
               Caption         =   "Bruton Layer Model (Temp vs. Height)"
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
               Left            =   240
               TabIndex        =   43
               ToolTipText     =   "Defined by Temperature Profile Graph"
               Top             =   280
               Value           =   -1  'True
               Width           =   3975
            End
            Begin VB.Label lblHeightStepSize 
               Caption         =   "Height Step Size (m)"
               Height          =   255
               Left            =   1200
               TabIndex        =   80
               Top             =   1880
               Width           =   1695
            End
         End
         Begin VB.Frame frmControl 
            Caption         =   "Calculate ray tracing"
            Height          =   4215
            Left            =   9480
            TabIndex        =   8
            Top             =   6720
            Width           =   5055
            Begin VB.CommandButton cmdMenat 
               Caption         =   "Perform Menat rray tracing method"
               Height          =   375
               Left            =   960
               TabIndex        =   241
               Top             =   3480
               Width           =   3375
            End
            Begin VB.CheckBox chkLapse 
               Height          =   195
               Left            =   840
               TabIndex        =   116
               Top             =   2760
               Width           =   255
            End
            Begin MSComCtl2.UpDown updwnLapse 
               Height          =   285
               Left            =   4320
               TabIndex        =   115
               Top             =   2760
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               Value           =   6
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtLapse"
               BuddyDispid     =   196800
               OrigLeft        =   3480
               OrigTop         =   2640
               OrigRight       =   3735
               OrigBottom      =   2895
               Max             =   20
               Min             =   -20
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   0   'False
            End
            Begin VB.TextBox txtLapse 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3600
               TabIndex        =   114
               Text            =   "6.5"
               Top             =   2760
               Width           =   720
            End
            Begin VB.CheckBox chkFudge 
               Caption         =   "Use fix for theta"
               Height          =   615
               Left            =   3840
               TabIndex        =   99
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox chkMenat 
               Caption         =   "Use Menat's approx. of Gardne'sr rindex"
               Enabled         =   0   'False
               Height          =   195
               Left            =   840
               TabIndex        =   98
               Top             =   1680
               Width           =   3255
            End
            Begin VB.CheckBox chkCiddor 
               Caption         =   "Use Ciddor's index of refraction instead of HS"
               Enabled         =   0   'False
               Height          =   255
               Left            =   840
               TabIndex        =   97
               Top             =   2040
               Width           =   3495
            End
            Begin VB.CheckBox chkHSoatm 
               Caption         =   "Use other atmospheres (above) for HS ray tracing"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   840
               TabIndex        =   96
               ToolTipText     =   "Instead of HS atmospheric model, use atmospheric models above for HS ray tracing"
               Top             =   2400
               Width           =   3855
            End
            Begin VB.CommandButton cmdRefWilson 
               Caption         =   "Perform HS Ray tracing"
               Height          =   435
               Left            =   1080
               TabIndex        =   94
               ToolTipText     =   "Use method of Hohenkerk and Sinclair"
               Top             =   1080
               Width           =   2895
            End
            Begin VB.CommandButton cmdCalc 
               Caption         =   "Perform Brtuon Ray tracing"
               Height          =   495
               Left            =   600
               TabIndex        =   12
               ToolTipText     =   "Perform Ray tracing using these parameters"
               Top             =   240
               Width           =   2895
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   4800
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Label lblLapse 
               Caption         =   "Lapse Rate (deg K/km)"
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
               Height          =   255
               Left            =   1440
               TabIndex        =   113
               Top             =   2760
               Width           =   2175
            End
            Begin VB.Line Line1 
               BorderStyle     =   6  'Inside Solid
               X1              =   120
               X2              =   4920
               Y1              =   960
               Y2              =   960
            End
         End
         Begin VB.Frame frmInit 
            Caption         =   "Calculation Parameters"
            Height          =   5535
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   4695
            Begin VB.TextBox txtKStep 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   41
               Text            =   "10"
               Top             =   4920
               Width           =   1200
            End
            Begin VB.TextBox txtKmax 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   39
               Text            =   "585"
               Top             =   4440
               Width           =   1200
            End
            Begin VB.TextBox txtKmin 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   37
               Text            =   "585"
               Top             =   3960
               Width           =   1200
            End
            Begin VB.TextBox txtPPAM 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   35
               Text            =   "2.0"
               Top             =   3120
               Width           =   1335
            End
            Begin VB.TextBox txtXmax 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   33
               Text            =   "1500"
               ToolTipText     =   "How far the rays are traced along the earth's circumference (km)"
               Top             =   2400
               Width           =   1200
            End
            Begin VB.TextBox txtNumSuns 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   31
               Text            =   "30"
               ToolTipText     =   "Number of suns to draw (number of steps)"
               Top             =   1920
               Width           =   1200
            End
            Begin VB.TextBox txtDelAlt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   29
               Text            =   "-5"
               ToolTipText     =   "Step in solar altitude in arc minutes"
               Top             =   1440
               Width           =   1200
            End
            Begin VB.TextBox txtStartAlt 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   27
               Text            =   "0.0"
               ToolTipText     =   "Starting non refraction solar altitude in arc minutes"
               Top             =   960
               Width           =   1200
            End
            Begin VB.TextBox txtHeight 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2760
               TabIndex        =   25
               Text            =   "100"
               ToolTipText     =   "Observer's height in meters"
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label lblWavelength 
               Caption         =   "Wavelength (mu), 0.65=R,0.589=Y (Sodium),0.52=G"
               Height          =   255
               Left            =   480
               TabIndex        =   236
               Top             =   3600
               Width           =   3855
            End
            Begin VB.Label lblKStep 
               Caption         =   "Wavelength Step (nm)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   40
               Top             =   4920
               Width           =   2055
            End
            Begin VB.Label lblKMax 
               Caption         =   "Wavelength Max.(nm)"
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
               Left            =   480
               TabIndex        =   38
               Top             =   4440
               Width           =   1935
            End
            Begin VB.Label lblWavMin 
               Caption         =   "Wavelength Min.(nm)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   36
               Top             =   3960
               Width           =   1935
            End
            Begin VB.Label lblResol 
               Caption         =   "Pixels per arcminute"
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
               Left            =   480
               TabIndex        =   34
               Top             =   3120
               Width           =   1935
            End
            Begin VB.Label lblXmax 
               Caption         =   "Max. disance.along earth circumfeence to trace path of rays ( km)"
               Height          =   615
               Left            =   480
               TabIndex        =   32
               Top             =   2400
               Width           =   2055
            End
            Begin VB.Label lbltSuns 
               Caption         =   "Number Steps (suns)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   30
               Top             =   1920
               Width           =   2055
            End
            Begin VB.Label lblDelAlt 
               Caption         =   "Delta Altitutde ( ' )"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   28
               Top             =   1440
               Width           =   1935
            End
            Begin VB.Label Label1 
               Caption         =   "Starting Altitude ( ' )"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   26
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label lblObserver 
               Caption         =   "Observer Height (m)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   24
               Top             =   480
               Width           =   1935
            End
         End
         Begin VB.Label lblRef 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   5400
            TabIndex        =   95
            Top             =   7920
            Width           =   3735
         End
         Begin VB.Label lblHorizon 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   5280
            TabIndex        =   66
            Top             =   9480
            Width           =   3855
            WordWrap        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "prjAtmRefMainfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Private HasSolution As Boolean
    Private PtX As Collection
    Private PtY As Collection
    Private BestCoeffs As Collection
'Private Declare Function GdipAddPathEllipse Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
'Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mFillMode As Long, ByRef mpath As Long) As Long
'Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPolyGradient As Long) As Long
'Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
'Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
'Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
'Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColors As Long, ByRef mCount As Long) As Long
'Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
'Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef graphics As Long) As Long
'Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
'
'Dim m_PtList(2) As POINTAPI
'Dim ZoomValue As Single
'
'Private Sub AlphaSun_DblClick()
'    HScrollPan.Value = 0
'    ZoomValue = 100
'    VScrollPan.Value = 0
'End Sub
'
'Private Sub AlphaSun_PrePaint(hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long, HitTestRgn As Long, Cancel As Boolean)
'    If HScrollPan.Enabled Then  ' else animation sample being run
'        Dim zoomOffset As Single
'        zoomOffset = ZoomValue / 100
'        lblZoom.Caption = "Zoom: " & Format(zoomOffset, "Percent")
'        Cancel = True ' prevent rendering image, we will be rendering it below
'        AlphaSun.Picture.Render hdc, (AlphaSun.Width - (AlphaSun.Width * zoomOffset)) \ 2 + HScrollPan.Value * zoomOffset, _
'            (AlphaSun.Height - (AlphaSun.Height * zoomOffset)) \ 2 + VScrollPan.Value * zoomOffset, _
'            AlphaSun.Width * zoomOffset, AlphaSun.Height * zoomOffset, , , , , , _
'            AlphaSun.Effects.AttributesHandle, , AlphaSun.Effects.EffectsHandle(AlphaSun.Effect)
'    Else
'        AlphaSun.Picture.RenderSkewed hdc, m_PtList(0).X, m_PtList(0).Y, m_PtList(1).X, m_PtList(1).Y, m_PtList(2).X, m_PtList(2).Y
'        Cancel = True
'    End If
'End Sub
'dll called to calculate Van der Werf raytracing
'Private Declare Function RayTracing Lib "AtmRef.dll" (StarAng As Double, EndAng As Double, StepAng As Double, NAngles As Long, _
'                                                                  HOBS As Double, TGROUND As Double, HMAXT As Double, _
'                                                                  GPress As Double, WAVELN As Double, HUMID As Double, OBSLAT As Double, NSTEPS As Long, _
'                                                                   ByVal pFunc As Long) As Long
'RayTracing(double *BETALO, double *BETAHI, double *BETAST,int *NAngles,
'                                            double *HOBSERVER, double *TEMPGROUND, double *HMAXT,
'                                            double *Press0, double *WAVELN, double *RELHUM, double *OBSLATITUDE, int *NSTEPS,
'                                           long cbAddress)
                                                                  



Private Sub chkDist_Click()
   If chkDist.Value = vbChecked Then
      chkH1.Value = vbUnchecked
      chkH2.Value = vbUnchecked
      chkTemp.Value = vbUnchecked
      End If
End Sub

Private Sub chkDucting_Click()
   With prjAtmRefMainfm
       If .chkDucting.Value = vbChecked Then
         .txtDInv.Enabled = True
         .txtSInv.Enabled = True
         .txtEInv.Enabled = True
         .lblInvStep.Enabled = True
         .lblEInv.Enabled = True
         .lblSInv.Enabled = True
         .UpDownDInv.Enabled = True
         .UpDownEInv.Enabled = True
         .UpDownSInv.Enabled = True
         INVFLAG = 1
         OptionSelby.Value = False
       Else
         .txtDInv.Enabled = False
         .txtSInv.Enabled = False
         .txtEInv.Enabled = False
         .lblInvStep.Enabled = False
         .lblEInv.Enabled = False
         .lblSInv.Enabled = False
         .UpDownDInv.Enabled = False
         .UpDownEInv.Enabled = False
         .UpDownSInv.Enabled = False
         INVFLAG = 0
         End If
   End With
End Sub

Private Sub chkH1_Click()
   If chkH1.Value = vbChecked Then
      chkH2.Value = vbUnchecked
      chkDist.Value = vbUnchecked
      chkTemp.Value = vbUnchecked
      End If
End Sub

Private Sub chkH2_Click()
   If chkH2.Value = vbChecked Then
      chkH1.Value = vbUnchecked
      chkDist.Value = vbUnchecked
      chkTemp.Value = vbUnchecked
      End If
End Sub

Private Sub chkLapse_Click()
  If chkLapse.Value = vbChecked Then
     lblLapse.Enabled = True
     txtLapse.Enabled = True
     updwnLapse.Enabled = True
  Else
     lblLapse.Enabled = False
     txtLapse.Enabled = False
     updwnLapse.Enabled = False
     End If
End Sub

Private Sub chkTemp_Click()
   If chkTemp.Value = vbChecked Then
      chkH2.Value = vbUnchecked
      chkDist.Value = vbUnchecked
      chkH1.Value = vbUnchecked
      End If
End Sub

Private Sub chkTRef_Click()
   If Trim$(txtDir.Text) = sEmpty Then txtDir = App.Path
End Sub

'Private Sub chkVDW_Click()
'   If chkVDW.Value = vbChecked Then
'      chkDucting.Value = vbUnchecked 'if adding an inversion layer then can't use special atmospheres
'      End If
'End Sub
'
'Private Sub chkHgtProfile_Click()
''   If prjAtmRefMainfm.chkHgtProfile.Value = vbChecked Then
''      prjAtmRefMainfm.txtHOBS.Text = DistModel(0)
''   Else
''      prjAtmRefMainfm.txtAs.Text = 0#
''      End If
'End Sub

Private Sub cmdBrowse_Click(Index As Integer)
   On Error GoTo errhand

'   TextExternal = BrowseForFolder(prjAtmRefMainfm.hwnd, "Choose Directory")
    With comdlgOther
        .CancelError = True
        .Filter = "dat fiels (*.dat)|*.dat|text files (*.txt)|*.txt|All fiels (*.*)|*.*"
        .filename = App.Path & "\*.dat"
        .ShowOpen
        TextExternal = .filename
   End With
   
errhand:
End Sub

Private Sub cmdBrowseTR_Click()
   On Error GoTo errhand

   txtDirTR = BrowseForFolder(prjAtmRefMainfm.hwnd, "Choose Directory with TR files")
   
errhand:
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCalc_Click
' Author    : Dr-John-K-Hall
' Date      : 12/24/2018
' Purpose   : Performs Ray Tracing
'---------------------------------------------------------------------------------------
'
Private Sub cmdCalc_Click()

'   On Error GoTo cmdCalc_Click_Error
   
   cmdCalc.Enabled = False
   cmdRefWilson.Enabled = False
   cmdMenat.Enabled = False
   cmdVDW.Enabled = False
   
   RefCalcType% = 0

 'Option Explicit
'C                  APPENDIX  B
'C     MULTILAYER MODEL SOURCE CODE
'C
'C   TRANFER CURVE
'C   DAN BRUTON (ASTRO@TAMU.EDU)
'C   APRIL 4, 1996
'C
'C   THIS PROGRAM CONVERTS A TRUE IMAGE OF THE SUN OR OBJECT TO AN
'C   APPARENT IMAGE FOR A GIVEN TEMPERATURE PROFILE (INPUT FILE) USING
'C   A PARABOLIC RAY PATH, MULTILAYER ATMOSPHERE MODEL.  A PORTABLE
'C   PIXMAP (PPM) IMAGE FILE IS GENERATED.
'C
'C   ALT = ALTITUDE OF THE OBJECT IN ARCMINUTES
'C   AZM = RELATIVE AZIMUTH OF THE OBJECT IN ARCMINUTES
'C   ALFA = APPARENT ALTITUDE IN ARCMINUTES
'C   ALFT = TRUE ALTITUDE IN ARCMINUTES
'C   RI = REFRACTIVE INDEX
'C   PRS = PRESSURE IN PASCALS
'C   TMP = ABSOLUTE TEMPERATURE
'C   RC = RADIUS OF CURVATURE OF THE RAY IN METERS
'C   HOBS = HEIGHT OF THE OBSERVER ABOVE THE EARTH'S SURFACE IN METERS
'C   ROBJ = RADIUS OF THE OBJECT IN ARCMINUTES
'C   RE = RADIUS OF THE EARTH IN METERS
'C   CV = PIXEL COLOR RGB VALUES (ARRAY)
'C   Image1 = MAKE IMAGE FLAG (TRUE OR FALSE: 1 OR 0)
'C   ITRAN = TRANSFER FUNCTION (COMPUTED OR STANDARD: 1 OR 0)
'C   BETA = INCIDENT ANGLE RELATIVE TO LAYER IN RADIANS
'C
'      IMPLICIT DOUBLE PRECISION (A-H,O-Z)
'      Dim CV(601, 601, 4) As Double
'      Dim ELV(2001) As Double, TMP(2001) As Double, RCV(82, 2001) As Double
'      Dim ALFA(82, 501) As Double, ALFT(82, 501) As Double, SSR(82, 501) As Double
'      Dim AA(2001) As Double, AT(2001) As Double, DEN As Double
'      Dim EDIS(82) As Double, IDCT(501) As Double, IEND(501) As Double
'      Dim AIRM(501) As Double, ADEN(2001) As Double, RC As Double
'      Dim SINV As Double, EINV As Double, HGTSCALE As Double, DTINV As Double
'      Dim ALT(11) As Double, AZM(11) As Double ', CNST(1000) As Double
      CalcComplete = False
      
      Dim StatusMes As String
      Dim FNM As String, lR As Double, Theta_M As Double, el As Double
      Dim NNN As Long, II As Long, k As Long, AtmType As Integer, AtmNumber As Integer
      Dim lpsrate As Double, tst As Double, pst As Double
      
'      COMMON /TC/ CNST
'      DATA ALT/0.,-21.,-30.,-42.,-51./
'      DATA AZM/19.,51.,83.,115.,147./
'      DATA ALT/100.,75.,51.,42.,21.,0.,-21.,-30.,-42.,-51./
'      DATA AZM/19.,51.,83.,115.,147.,179.,211.,243.,275.,307./
'      ALT(0) = -20#: ALT(1) = -30: ALT(2) = -40#: ALT(3) = -50#: ALT(4) = -60#
'      AZM(0) = 19#: AZM(1) = 51#: AZM(2) = 83#: AZM(3) = 115#: AZM(4) = 147#
       
       cmbSun.Clear
'       ALT(1) = 100#: ALT(2) = 75#: ALT(3) = 51#: ALT(4) = 42#: ALT(5) = 21#: ALT(6) = 0#: ALT(7) = -21#: ALT(8) = -30#: ALT(9) = -42#: ALT(10) = -51#
'       AZM(1) = 19#: AZM(2) = 51#: AZM(3) = 83#: AZM(4) = 115#: AZM(5) = 147#: AZM(6) = 179#: AZM(7) = 211#: AZM(8) = 243#: AZM(9) = 275#: AZM(10) = 307#
'C
'C     PPM IMAGE CONSTANTS AND OTHER INPUT PARAMETERS
'C     HEIGHT, WIDTH, PIXEL DEPTH, PIXELS PER ARCMINUTE
'C
      
      Dim ISEED As Long
      ISEED = 5
'      Dim RE As Double
      RE = 6371000#  '6378140# 'use average radius, not radius at equator
      
      pi = 4# * Atn(1#) '3.141592654
      CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
      cd = pi / 180# 'conversion of degrees to radians

      ROBJ = 15#
      KMIN = 1
      KMAX = 81
      KSTEP = 40
      GAM = 0.8
      max = 255
      HDEG = 1.5
      PPAM = CDbl(n) / (HDEG * 60#)
      DELTA = 1000#
      XMAX = 30000#
      SSRMAX = 0#
      HOBS = 10.01
      DELALT = -10#
      DELAZM = 32#
      STARTAZM = 19#
      RMX = 0#
      TSUN = 5800#
      ITRAN = 1
      Image1 = 1
      IPLOT = 0
      If Val(txtYSize) = 0 Then txtYSize = "1"
      n = 2 * Val(txtYSize) * 60 * PPAM '300 'height of ppm image of suns, determines range of viewing angles
      m = Val(txtXSize) 'width of pm image of suns
      If m = 0 Then m = 1000
      
      Dim ier As Long
      
'C   INVERSION LAYER CONSTANTS
    INVFLAG = 0
    DTINV = 5
    SINV = 0
    EINV = 100
    KSTART = 0
'C
'C   HEIGHT SCALING
    HGTSCALE = 0
    
     '------------------progress bar initialization
    With prjAtmRefMainfm
      '------fancy progress bar settings---------
      .progressfrm.Visible = True
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------

'C
'C     READ INPUT PARAMETERS
'C
'      PRINT *, '    Reading input parameters from paramd.dat'
'GoTo pAR10 'skip reading param.dat, use the parameters instead

'      Dim Txt As String
'      Dim ParamFile As String
'
'      StatusMes = "Choose parameter file to open (default: param.dat)"
'      Call StatusMessage(StatusMes, 1, 0)
''
''      Call FileDialog("Choose parameter file (param.dat)", "param.dat", App.Path, ParamFile, ".dat", 0, ier)
''      If ier < 0 Then Exit Sub
'      ParamFile = App.Path & "\paramd.dat" 'param.dat"
'
'      Dim filnum%
'      filnum% = FreeFile
'      Open ParamFile For Input As #filnum%
'
''      OPEN(UNIT=20,FILE='paramd.dat',STATUS='UNKNOWN')
''      READ(20,*)  HOBS,ITRAN,Image1,IPLOT,STARTALT,DELALT,XMAX
'      Dim STARTALT As Double
'      Input #filnum%, HOBS, ISSR, ITRAN, Image1, IPLOT, STARTALT, DELALT, XMAX
'
''2     Format (A80)
'      Dim doclin$
''      READ(20,2) TXT
'      Line Input #filnum%, doclin$
'      'READ(20,*) N,M,PPAM,KMIN,KMAX,KSTEP
'      Input #filnum%, n, M, PPAM, KMIN, KMAX, KSTEP
'      'READ(20,2) TXT
'      Line Input #filnum%, doclin$
''3     Format (A15)
''      READ(20,3) FNM
'      Line Input #filnum%, FNM
'      If (M = 0) Then M = Int(n * 6 / 3)
'      If (PPAM = 0) Then PPAM = CDbl(n) / (HDEG * 60#)
''      PRINT *, '    Pixels per arcminute ', PPAM
''      PRINT *, '    Maximum height (degrees) ', (N/(120.D0*PPAM))
'      StatusMes = "Pixels per arcminute " & Str(PPAM) & ", Maximum height (degrees) " & Str(n / (120# * PPAM))
'      Call StatusMessage(StatusMes, 1, 0)
''      Close (20)
'      Close #filnum%
      
pAR10:
      HOBS = Val(txtHeight.Text)
      If HOBS = 0 Then HOBS = 0.001 'this code doesn't work for hobs = 0 so add epsiolon of height
      ISSR = 0
      ITRAN = 1
      Image1 = 1
      IPLOT = 1
      STARTALT = Val(txtStartAlt.Text)
      DELALT = Val(txtDelAlt.Text)
      XMAX = Val(txtXmax.Text) * 1000 'convert km to meters
      PPAM = Val(txtPPAM.Text)
        KMIN = CInt((Val(txtKmin.Text) - 380) / 5# + 1#)
        KMAX = CInt((Val(txtKmax.Text) - 380) / 5# + 1#)
        KSTEP = CInt(Val(txtKStep.Text) * 0.1)
      STARTAZM = 19
      DELAZM = 32
      If INVFLAG = 1 Then
         SINV = Val(txtSInv.Text)
         EINV = Val(txtEInv.Text)
         DTINV = Val(txtDInv.Text)
         End If
         
      n = 500 '300 'height of ppm image of suns, determines range of viewing angles
      m = 20 + Val(txtNumSuns.Text) * 32 * PPAM
      
      If Trim$(txtXSize.Text) <> sEmpty Then
         m = Val(txtXSize.Text)
         End If
      If Trim$(txtYSize.Text) <> sEmpty Then
         n = 2 * Val(txtYSize.Text) * 60 * PPAM
         End If
         
      StatusMes = "Pixels per arcminute " & Str(PPAM) & ", Maximum height (degrees) " & Str(n / (120# * PPAM))
      Call StatusMessage(StatusMes, 1, 0)
      
      Dim KA As Long
      For KA = 1 To NumSuns
         ALT(KA) = STARTALT + CDbl(KA - 1) * DELALT
         AZM(KA) = STARTAZM + CDbl(KA - 1) * DELAZM
      Next KA
'C
'C   *** MAKE PLOT PARAMETERS ***
'C
'      If (IPLOT = 1) Then
'         N = 10
'         PPAM = 0.8333
'         End If
'C
'C     READ TEMPERATURE PROFILE
'C
'      StatusMes = "Choose temperature profile file to open (default: " & FNM & ")"
'      Call StatusMessage(StatusMes, 1, 0)
      
'      Call FileDialog("Choose temperature profiile file (" & FNM & ")", FNM, App.Path, FNM, ".dat", 0, ier)
'      If ier < 0 Then
'         Call MsgBox("Couldn't find the requested file", vbCritical, "file tempprof.dat")
'         Exit Sub
'         End If

     'specify atmosphere type and the file containing the atmosphere profile
      StatusMes = "Calculating and Storing multilayer atmospheric details"
      Call StatusMessage(StatusMes, 1, 0)
     
     If OptionLayer.Value = True Then
        AtmType = 1
        FNM = App.Path & "\stmod1.dat"
     ElseIf OptionRead.Value = True Then
        AtmType = 1
        FNM = TextExternal.Text
     ElseIf OptionSelby.Value = True Then
        AtmType = 2
        If prjAtmRefMainfm.opt1.Value = True Then
           AtmNumber = 1
        ElseIf prjAtmRefMainfm.opt2.Value = True Then
           AtmNumber = 2
        ElseIf prjAtmRefMainfm.opt3.Value = True Then
           AtmNumber = 3
        ElseIf prjAtmRefMainfm.opt4.Value = True Then
           AtmNumber = 4
        ElseIf prjAtmRefMainfm.opt5.Value = True Then
           AtmNumber = 5
        ElseIf prjAtmRefMainfm.opt6.Value = True Then
           AtmNumber = 6
        ElseIf prjAtmRefMainfm.opt7.Value = True Then
           AtmNumber = 7
        ElseIf prjAtmRefMainfm.opt8.Value = True Then
           AtmNumber = 8
        ElseIf prjAtmRefMainfm.opt9.Value = True Then
           AtmNumber = 9
        ElseIf prjAtmRefMainfm.opt10.Value = True Then
           AtmNumber = 10
           FNM = txtOther.Text
           End If
        End If

'     FNM = App.Path & "\stmod1.dat" 'stmod2.dat.txt"
     ier = LoadAtmospheres(FNM, AtmType, AtmNumber, lpsrate, tst, pst, NNN, 1)
     
     If ier < 0 Then
        Screen.MousePointer = vbDefault
        Close
        cmdVDW.Enabled = True
        cmdCalc.Enabled = True
        cmdMenat.Enabled = True
        cmdRefWilson.Enabled = True
        Exit Sub
        End If
        
     'now load up temperature and pressure charts
      ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
      
      For j = 1 To NNN
         TransferCurve(j, 1) = " " & CStr(ELV(j - 1) * 0.001)
'         TransferCurve(J, 2) = ELV(J - 1) * 0.001
         TransferCurve(j, 2) = TMP(j - 1)
      Next j
      
      With MSChartTemp
        .chartType = VtChChartType2dLine
        .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN - 1
'        .RowLabel = "Height (km)"
'        .ColumnLabel = "Temperature (Kelvin)"
        .ChartData = TransferCurve
      End With
     
      For j = 1 To NNN
         TransferCurve(j, 2) = PRSR(j - 1)
      Next j
      
      With MSChartPress
        .chartType = VtChChartType2dLine
        .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN - 1
'        .RowLabel = "Height (km)"
'        .ColumnLabel = "Pressure (Kelvin)"
        .ChartData = TransferCurve
      End With
     '///////////////////////////////////////////////////////////////
         
'     Screen.MousePointer = vbHourglass
'
''      OPEN(UNIT=20,FILE=FNM,STATUS='UNKNOWN')
'      filnum% = FreeFile
'      Open FNM For Input As #filnum%
''      PRINT *, '    Reading temperature profile from ',FNM
'      StatusMes = "Reading temperature profile from " & FNM
'      Call StatusMessage(StatusMes, 1, 0)
'
'      'READ(20,*) NNN
'      Input #filnum%, NNN
'      NNN = NNN - 1
'      Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, 0)
'      For II = 0 To NNN
'         'READ(20,*) ELV(II),TMP(II)
'         Input #filnum%, ELV(II), TMP(II)
'
'         If II = 0 Then
'            MinTemp = TMP(II)
'            MaxTemp = MinTemp
'         Else
'            If TMP(II) > MaxTemp Then MaxTemp = TMP(II)
'            If TMP(II) < MinTemp Then MinTemp = TMP(II)
'            End If
'
'         Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, CLng(100# * II / (NNN - 1)))
'
'      Next II
'      'Close (20)
'      Close #filnum%
'      prjAtmRefMainfm.progressfrm.Visible = False
'      Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, 0)
      
      '////////////////////////////////////////////////////////////////////
      Screen.MousePointer = vbHourglass
      
      NumTemp = NNN + 1
      
      If ((ELV(0) < HOBS) And (HOBS < 0)) Then
'C         USING BELOW SEA LEVEL ATMOSPHERE TERMS, SO
'C         RESCALE EARTH RADIUS AND OBSERVER HEIGHT
          IISTART = Int((HOBS - ELV(0)) * 0.1)
          HGTSCALE = ELV(IISTART)
          RE = RE + HGTSCALE
          HOBS = HOBS - HGTSCALE
'C         RESCALE INVERSION CONSTANTS
          SINV = SINV - HGTSCALE
          EINV = EINV - HGTSCALE
'C         NOW RESCALE THE TEMPERATURE PROFILE
'C         ALSO ADD IN INVERSION IF FLAGGED
          For II = IISTART To NNN
             ELV(II - IISTART) = ELV(II) - HGTSCALE
             TMP(II - IISTART) = TMP(II)
          Next II
          NNN = NNN - IISTART
          End If

      If (INVFLAG <> 0) Then
'C        ADD TEMPERATURE INVERSION
         For II = 0 To NNN
            If ((ELV(II) >= SINV) And (ELV(II) <= EINV)) Then
               TMP(II) = TMP(II) + DTINV
               End If
         Next II
         End If
         
      
'      PRINT *, '    Computing radius of curvatures. '
      StatusMes = "Computing radius of curvatures"
      Call StatusMessage(StatusMes, 1, 0)
      
      prjAtmRefMainfm.progressfrm.Visible = True
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
      
      Dim wl As Double
      
      For II = 0 To NNN
         For k = 1 To 81
            wl = 380# + CDbl(k - 1) * 5#
'            Call RADCUR(II, RC, DEN, ELV(), TMP(), WL, AtmType)
'            Call RADCUR(II, RC, DEN, WL, AtmType)
            Call RADCUR_new(II, RC, den, wl, AtmType)
            RCV(k, II) = RC
            ADEN(II) = den
         Next k
         Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * II / NNN))
      Next II
      prjAtmRefMainfm.progressfrm.Visible = False
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)

'C
'C     COMPUTE VERTICAL AIRMASS
'C
      Dim AMZ As Double
      AMZ = 0#
      For II = 1 To NNN
         AMZ = AMZ + ADEN(II) * (ELV(II) - ELV(II - 1))
      Next II
'      PRINT *, '    Vertical Air Mass (kg/m^2)',AMZ
      StatusMes = "Vertical Air Mass (kg/m^2 = " & Str(AMZ) & ", Finding transfer function, writing to test.dat"
      Call StatusMessage(StatusMes, 1, 0)
'C
'C     FIND ALFT IN TERMS OF ALFA
'C
'      PRINT *, '    Finding transfer function.'
'      Dim KSTOP As Long, III As Long, IIS As Long, J As Long
'      Dim AMZOBS As Double, BETA As Double, x As Double, DIS As Double
'      Dim THETA As Double, H As Double, UP As Double, XD As Double, HD As Double
'      Dim R1 As Double, R2 As Double, Z As Double, ARGT As Double
'      Dim DZ As Double, AMGAM As Double, XC As Double, YC As Double, ZC As Double
'      Dim ISTND As Long, BETAD As Double, TOP As Double, BOT As Double, DAB As Double
'      Dim ZL As Double, ZR As Double, ISTOP As Long, CON As Double, U As Double
'      Dim RMINTMP As Double, RMAXTMP As Double, RMINELV As Double, RMAXELV As Double
'      Dim IJK As Long, MMM As Long, IX As Long, IY As Long, IPER As Long
'      Dim IG As Long, IR As Long, IB As Long, MM As Long, NN As Long
'      Dim r As Double, g As Double, b As Double
      
      If Dir(App.Path & "\CIEColorMatchingFunctions.txt") <> sEmpty Then
         'upload the CIE Color Matching functions values
         filnum% = FreeFile
         Open App.Path & "\CIEColorMatchingFunctions.txt" For Input As #filnum%
         For j = 1 To 81
             Input #filnum%, WXYZ(1, j), WXYZ(2, j), WXYZ(3, j)
         Next j
         Close #filnum%
      Else
        Call MsgBox("Can't find the CIEColorMatchingFunctions.txt file in the program directory." _
                    & vbCrLf & sEmpty _
                    & vbCrLf & "Please find it...." _
                    , vbInformation, "missing file")
      
        Call FileDialog("Choose CIE Color Matching Functions File", "CIEColorMatchingFunctions.txt", App.Path, FNM, ".txt", 0, ier)
        If ier < 0 Then
           Screen.MousePointer = vbDefault
           Call MsgBox("Couldn't find the requested file", vbCritical, "file tempprof.dat")
           Exit Sub
           End If

        filnum% = FreeFile
        Open FNM For Input As #filnum%
        For j = 1 To 81
            Input #filnum%, WXYZ(1, j), WXYZ(2, j), WXYZ(3, j)
        Next j
        Close #filnum%

        End If
      
'C
'C     FIND ALFT IN TERMS OF ALFA
'C
      StatusMes = "Finding transfer function."
      Call StatusMessage(StatusMes, 1, 0)
      
      filnum% = FreeFile
      For k = KMIN To KMAX Step KSTEP   '<1
         KSTOP = k
'         OPEN(UNIT=20,FILE='test.dat',STATUS='UNKNOWN')
         If Dir(App.Path & "\test.dat") <> sEmpty Then
            Kill App.Path & "\test.dat"
            End If
         Open App.Path & "\test.dat" For Output As #filnum%
'1        FORMAT(F15.5,1X,F15.5)
         If (ITRAN = 1) Then
             wl = 380# + CDbl(k - 1) * 5#
    '         PRINT *, ' Wavelength (nm) ',WL
             StatusMes = "Wavelength (nm) " & Str(wl)
             Call StatusMessage(StatusMes, 1, 0)
      

'C   FIND THE STARTING LAYER
         For III = 0 To NNN
            If (HOBS >= ELV(III)) Then IIS = III + 1
         Next III
         AMZOBS = 0#
         For III = 1 To IIS
            AMZOBS = AMZOBS + ADEN(III) * (ELV(III) - ELV(III - 1))
         Next III
         AMZOBS = AMZOBS - ADEN(II) * (ELV(II) - HOBS)
'C        RAY TRACE THROUGH THE FIRST LAYER

        If IPLOT = 1 Then
            prjAtmRefMainfm.progressfrm.Visible = True
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
            End If

         For j = 1 To n + 1 '<2

         If (IPLOT = 1) Then
'            PRINT *, J
'            StatusMes = Str(J)
'            Call StatusMessage(StatusMes, 1, 0)
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * j / (n + 1)))
            DoEvents
            End If
         II = IIS
         IDCT(j) = 0
         AIRM(j) = 0#
         SSR(k, j) = 0#
         ALFA(k, j) = (CDbl(n / 2 - (j - 1)) / PPAM)
'c      ALFA is the observed viewing angle, and is simply defined by the pixel height of the image
'c      starting from largest positive angles looking upward to smallest negative angles looking downward
'c      Program loops over all of these observed angles and records the ray tracing corresponding to that
'c      angle when flagged, also recording the corresponding "true" angle (i.e., without refraction), ALFT, when flagged
         
         BETA = ALFA(k, j) * CONV
         RC = RCV(k, II) * Cos(BETA) * Cos(BETA) * Cos(BETA)
         x = 0#
'C        INITIAL RAY - LEHN'S EQUATION 12 & 17
         If (BETA >= 0#) Then
            DIS = Tan(BETA) ^ 2# + 2# * (ELV(II) - HOBS) * (RC - RE) / (RC * RE)
            If (DIS <= 0#) Then
               DIS = Tan(BETA) ^ 2# - 2# * (HOBS - ELV(II - 1)) * (RC - RE) / (RC * RE)
               UP = (-Tan(BETA) - Sqr(DIS)) * (RE * RC / (RC - RE))
               IFLAG = 0
            Else
               UP = (-Tan(BETA) + Sqr(DIS)) * (RE * RC / (RC - RE))
               IFLAG = 1
               End If
         Else
            DIS = Tan(BETA) ^ 2# - 2# * (HOBS - ELV(II - 1)) * (RC - RE) / (RC * RE)
            If (DIS <= 0#) Then
               DIS = Tan(BETA) ^ 2# + 2# * (ELV(II) - HOBS) * (RC - RE) / (RC * RE)
               UP = (-Tan(BETA) + Sqr(DIS)) * (RE * RC / (RC - RE))
               IFLAG = 1
            Else
               UP = (-Tan(BETA) - Sqr(DIS)) * (RE * RC / (RC - RE))
               IFLAG = 0
               End If
            End If
            
'C       CONTINUE RAY TRACING THROUGH THE LAYERS
            Theta = 0#
            Theta_M = 0#
            H = HOBS
            Do While ((II >= 1) And (II <= NNN) And (x <= XMAX))
            
'               IF (IPLOT = 1) WRITE (20,1) X,H
               If (IPLOT = 1) Then Print #filnum%, Format(x, "######0.0####"), Format(H, "######0.0####"), Format(ALFA(k, j), "######0.0####")

               XD = 0#
               HD = H
               
'C              FILL IN THE GAPS DURING RAY TRACING
               Do While ((IPLOT = 1) And (XD <= UP) And ((x + XD) <= XMAX))
'                 IF (IPLOT = 1) WRITE(20,1) X+XD,HD
                 If (IPLOT = 1 And XD <> 0) Then
                    Print #filnum%, Format(x + XD, "######0.0####"), Format(HD, "######0.0####"), Format(ALFA(k, j), "######0.0####")
                    End If
                XD = XD + DELTA
                HD = ((1# / RE) - (1# / RC)) * (XD ^ 2#) / 2# + XD * Tan(BETA) + H
               Loop
               
'C              FIND NEW POSITION ANGLE THETA WRT EARTH'S CENTER
               R1 = RE + ELV(II - 1)
               R2 = RE + ELV(II)
               z = -0.5 * Cos(BETA) * (UP ^ 2#) / RC + UP * Tan(BETA)
               ARGT = (R1 ^ 2 + R2 ^ 2# - UP ^ 2# - z ^ 2#) / (2# * R1 * R2)
               
               If (Abs(ARGT) <= 1#) Then Theta = Theta + DACOS(ARGT)
               
'C                AIRMASS CALCULATION
                  If (IFLAG = 1) Then
                     DZ = ELV(II) - (UP ^ 2#) / (2# * RE) - H
                  Else
                     DZ = ELV(II - 1) - (UP ^ 2#) / (2# * RE) - H
                     End If
                     
'C                AIRMASS
                  dd = Sqr(UP * UP + DZ * DZ)
                  
                  If (RC <> 0#) Then
                     AMGAM = 2# * DASIN(dd / (2# * RC))
                     AIRM(j) = AIRM(j) + Abs(ADEN(II) * RC * AMGAM)
                  Else
                     AIRM(j) = AIRM(j) + dd
                     End If
                     
'C                 DETERMINE THE NEW PARAMETERS
                  BETA = Atn(Tan(BETA) - (UP / RC) + (UP / RE)) 'view angle for this increment
                  
                  If chkFudge.Value = vbChecked Then
                    'calculate theat a different way
                    el = SQT(BETA / cd, R2 - R1, R1) 'approximate path length in this increment
                    lR = el * Cos(BETA) 'path along the last increment's radius
                    Theta_M = Theta_M + DASIN(lR / R2) 'angle subtended on circumference of Earth by last increment in r in radians
                    ALFT(k, j) = BETA / CONV - 0.5 * (Theta + Theta_M) / CONV 'use the average theta as a fudge solution
                  Else
                    ALFT(k, j) = (BETA - Theta) / CONV  'this is current angle of ray, so refraction at this point will be ALFA - ALFT
                    End If
                  
                  If (IFLAG = 1) Then
                     H = ELV(II)
                     II = II + 1
                  Else
                     H = ELV(II - 1)
                     II = II - 1
                     End If
                     
                  RC = RCV(k, II) * Cos(BETA) * Cos(BETA)
                  x = x + UP
                  
'C                 CALCULATE NEW UP VALUE
                  If ((II >= 1) And (II <= NNN)) Then
                  
                     If (IFLAG = 1) Then
'C                       LEHN'S EQUATION 12 & 13
                        DIS = Tan(BETA) ^ 2# + 2# * (ELV(II) - ELV(II - 1)) * (RC - RE) / (RC * RE)
                        
                        If (DIS <= 0#) Then
                           UP = 2# * RE * RC * Tan(BETA) / (RE - RC)
                           IFLAG = 0
                        Else
                           UP = (-Tan(BETA) + Sqr(DIS)) * (RE * RC / (RC - RE))
                           IFLAG = 1
                           ZL = ELV(II)
                           ZR = ELV(II - 1)
                           End If
                           
                     Else
'C                      LEHN'S EQUATION 17 & 13
                        DIS = Tan(BETA) ^ 2# - 2# * (ELV(II) - ELV(II - 1)) * (RC - RE) / (RC * RE)
                        
                        If (DIS <= 0#) Then
                           UP = 2# * RE * RC * Tan(BETA) / (RE - RC)
                           IFLAG = 1
                        Else
                           UP = (-Tan(BETA) - Sqr(DIS)) * (RE * RC / (RC - RE))
                           IFLAG = 0
                           End If
                           
                        End If
                        
                     End If
                     
                  Loop
                  
               If (II = 0) Then ALFT(k, j) = -1000#
               If (x >= XMAX) Then IDCT(j) = 1
               IEND(j) = II
            Next j '<2
            
'C           FOR DUCTING PHENOMENON LIKE NOVAYA ZEMLA
            ISTND = 0
            If (ISTND = 1) Then
               For j = 1 To n + 1
                  If (ALFT(k, j) <> -1000#) Then
                     BETA = ALFT(k, j)
                     BETAD = BETA / 60#
                     Top = (0.001594 + (1.96 - 4 * BETAD) + (2# - 7 * BETAD * BETAD))
                     BOT = (1# + (0.505 * BETAD) + (0.0845 * BETAD * BETAD))
                     DAB = (ADEN(IEND(j)) * 8.31451 / 28.96 - 3) * 60# * Top / BOT
                     ALFT(k, j) = BETA - DAB
                     End If
               Next j
               End If
''C
''C           TRANSFER FUNCTION FOR A STANDARD ATMOSPHERE
''C
            ElseIf (ITRAN = 0) Then
               For j = 1 To n + 1
                  ALFA(k, j) = (CDbl(n / 2 - (j - 1)) / PPAM)
                  BETA = (CDbl(n / 2 - (j - 1)) / PPAM)
                  BETAD = BETA / 60#
                  Pr = 99975#
                  TR = 273#
                  Top = Pr * (0.001594 + (0.000196 * BETAD) + (0.0000002 * BETAD * BETAD))
                  BOT = TR * (1# + (0.505 * BETAD) + (0.0845 * BETAD * BETAD))
                  DAB = 60# * Top / BOT
                  ALFT(k, j) = BETA - DAB
               Next j
               End If
         Next k
         Close #filnum%
         Screen.MousePointer = vbDefault
         If IPLOT = 1 Then
            prjAtmRefMainfm.progressfrm.Visible = False
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
            End If
         
'C
'C    READ OR WRITE TRANSFER CURVE TO A FILE
'C       ASSUME DESCENDING ORDER
'C
         If (ITRAN = 2) Then
'         PRINT *, ' Reading transfer curve.'
         StatusMes = "Reading transfer curve."
         Call StatusMessage(StatusMes, 1, 0)
         filnum% = FreeFile
'         OPEN(UNIT=20,FILE='tcin.dat',STATUS='UNKNOWN')
         Dim NM As Long
         Call FileDialog("Choose transfer curve filename", "tcin.dat", App.Path, FNM, ".dat", 0, ier)
        If ier < 0 Then
           Call MsgBox("Couldn't find the requested file", vbCritical, "file tempprof.dat")
           Exit Sub
           End If
           
         Open FNM For Input As #filnum%
'         READ (20,*) NM
         Input #filnum%, NM
         For jm = 1 To NM
'            READ(20,*) AA(JM),AT(JM)
            Input #filnum%, AA(jm), AT(jm)
         Next jm
         For k = KMIN To KMAX Step KSTEP
            For j = 1 To n + 1
               ALFA(k, j) = (CDbl(n / 2 - (j - 1)) / PPAM)
               ALFT(k, j) = 0#
               For jm = 2 To (NM - 1)
                  If ((ALFA(k, j) >= AA(jm)) And (ALFA(k, j) <= AA(jm - 1))) Then
                     X1 = AA(jm)
                     X2 = AA(jm - 1)
                     Y1 = AT(jm)
                     Y2 = AT(jm - 1)
                     ALFT(k, j) = Y1 + (Y2 - Y1) * (ALFA(k, j) - X1) / (X2 - X1)
                     End If
               Next jm
         Next j
      Next k
      Close #filnum%

      End If
      
       Dim NumTc As Long
       
'      PRINT *, '    Writing transfer curve.'
        StatusMes = "Writing transfer curve."
        Call StatusMessage(StatusMes, 1, 0)
        filnum% = FreeFile
'      OPEN(UNIT=20,FILE='tc.dat',STATUS='UNKNOWN')
      If Dir(App.Path & "\tc.dat") <> sEmpty Then
         Kill App.Path & "\tc.dat"
         End If
      Open App.Path & "\tc.dat" For Output As #filnum%
'      WRITE (20,*) N
      Print #filnum%, n
      For j = 1 To n + 1
'        WRITE(20,1) ALFA(KMIN,J),ALFT(KMIN,J)
        Print #filnum%, ALFA(KMIN, j), ALFT(KMIN, j)
        If ALFA(KMIN, j) = 0 Then 'display the refraction value for the zero view angle ray
           prjAtmRefMainfm.lblRef.Caption = "Atms. refraction (deg.) = " & Abs(ALFT(KMIN, j)) / 60# & vbCrLf & "Atms. refraction (mrad) = " & Abs(ALFT(KMIN, j)) * 1000# * cd / 60#
           End If
        'store all view angles that contribute to sun's orb
        If ALFT(KMIN, j) <> -1000 Then
            NumTc = NumTc + 1
            For KA = 1 To NumSuns
               y = ALFT(KMIN, j) - ALT(KA)
               If Abs(y) <= ROBJ Then
                  'only accept rays that pass over the horizon (ALFT(KMIN, J) <> -1000) and are within the solar disk
                  SunAngles(KA - 1, NumSunAlt(KA - 1)) = j
                  NumSunAlt(KA - 1) = NumSunAlt(KA - 1) + 1
                  End If
            Next KA
            End If
'        cmbAlt.AddItem ALFA(KMIN, J)
'        cmbAlt.ListIndex = cmbAlt.ListCount - 1
'        cmbAlt.Refresh
      Next j
      Close #filnum%
      
      'now load up transfercurve array for plotting
      ReDim TransferCurve(1 To NumTc, 1 To 2) As Variant
      
      For j = 1 To NumTc
         TransferCurve(j, 1) = " " & CStr(ALFA(KMIN, j))
         TransferCurve(j, 2) = ALFT(KMIN, j)
'         TransferCurve(J, 1) = " " & CStr(ALFT(KMIN, J))
'         TransferCurve(J, 2) = ALFA(KMIN, J)
      Next j
      
      With MSCharttc
        .chartType = VtChChartType2dLine
        .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN
'        .RowLabel = "True angle (min)"
'        .ColumnLabel = "View angle (min)"
        .ChartData = TransferCurve
      End With
      
'      For KA = 1 To NumSuns
'         cc = NumSunAlt(KA - 1)
'      Next KA
''C
''C   FIND THE ALTITUDE OF THE HORIZON
''C
      
      For j = n To 1 Step -1
         If (ALFT(KMIN, j) = -1000#) Then ISTOP = j
      Next j
'      PRINT *, '    Apparent Altitude of the Horizon (arcminutes)',
'     *      ALFA(KSTOP,ISTOP)
'      PRINT *, '    True Altitude of the Horizon (arcminutes)',
'     *      (-DACOS(RE/(RE+HOBS))/CONV)
     StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KSTOP, ISTOP)) & vbCrLf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(RE / (RE + HOBS)) / CONV))
     Call StatusMessage(StatusMes, 1, 0)
'     ier = MsgBox(StatusMes, vbInformation + vbOKOnly, "Horizon")
     lblHorizon.Caption = StatusMes
     DoEvents
'C
'C   MAKE Image1 OF THE OBJECT USE LIMB DARKENING
'C
      If (Image1 = 1) Then
'      PRINT *, '    Making Image1.'
     StatusMes = "Creating Image1."
     Call StatusMessage(StatusMes, 1, 0)
     
     Screen.MousePointer = vbHourglass
      
      For j = 1 To n + 1
         For i = 1 To m
            CV(i, j, 1) = 0#
            CV(i, j, 2) = 0#
            CV(i, j, 3) = 0#
         Next i
      Next j


      prjAtmRefMainfm.progressfrm.Visible = True
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)

      For j = 1 To n + 1 '<1
         IPER = Int(j * 10# / n)
         If (IPER = (j * 10# / n)) Then
'            PRINT *, IPER*10,'%'
            StatusMes = "Writing sun images: " & Str(IPER * 10) & "%"
            Call StatusMessage(StatusMes, 1, 0)
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(j * 100# / (n + 1)))
            DoEvents
            End If
         For i = 1 To m
'            Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, CLng(I * 100# / M))
            For k = KMIN To KMAX Step KSTEP
               wl = 380# + CDbl(k - 1) * 5#
               CON = 1240# / 0.00008617
'C          BLACKBODY RADIATION
               EDIS1 = 3.74183E-16 * (1# / wl ^ 5) / (Exp(CON / (wl * TSUN)) - 1#)
'C EXTINCTION
               TAUN = (283# / wl) ^ 4
               EDIS(k) = 0#
'C          DRAW SUNS
               For KA = 1 To NumSuns
                  y = ALFT(k, j) - ALT(KA)
                  x = (CDbl(i) / PPAM) - AZM(KA)
                  RHO = Sqr(x * x + y * y)
                  If (RHO <= ROBJ) Then
                     U = 0.6
                     DARK = (1# - U * (1# - Sqr(1# - (RHO / ROBJ) ^ 2))) / (1# - U / 3#)
                     EDIS(k) = EDIS(k) + EDIS1 * DARK * Exp(-TAUN * AIRM(j) / AMZ)
                  Else
'C            APPROXIMATE BACKGROUND SKY
                     If (ALFT(k, j) <> -1000#) Then
                        EDIS(k) = EDIS(k) + EDIS1 * Exp(-TAUN * AIRM(j) / AMZ) / 20#
                        End If
                     End If
                Next KA
            Next k
            Call EXYZ(xc, YC, ZC, EDIS())
            Call XYZTORGB(xc, YC, ZC, r, g, b)
            CV(i, j, 1) = CV(i, j, 1) + r
            CV(i, j, 2) = CV(i, j, 2) + g
            CV(i, j, 3) = CV(i, j, 3) + b
            If (DMAX1(CV(i, j, 1), CV(i, j, 2), CV(i, j, 3)) > RMX) Then
               RMX = DMAX1(CV(i, j, 1), CV(i, j, 2), CV(i, j, 3))
               End If
         Next i
      Next j
      prjAtmRefMainfm.progressfrm.Visible = False
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
      
'C
'C   WRITE TO PPM FILE
'C
'       PRINT *, '   Writing to PPM file.'
       StatusMes = " Writing to PPM sky view file."
       Call StatusMessage(StatusMes, 1, 0)
'       OPEN(UNIT=20,FILE='temp.ppm',STATUS='UNKNOWN')
       If Dir(App.Path & "\temp.ppm") <> sEmpty Then
          Kill App.Path & "\temp.ppm"
          End If
       filnum% = FreeFile
       Open App.Path & "\temp.ppm" For Output As #filnum%
'2001      Format (A2)
'       WRITE(20,2001) 'P3'
       Print #filnum%, "P3"
'1002      Format (A12)
'       WRITE(20,1002) '# temp.ppm'
       Print #filnum%, "# temp.ppm"
'1003      FORMAT (I3,1X,I3)
'       WRITE(20,1003) M,N
       Print #filnum%, m, n
'1004      Format (I3)
'       WRITE(20,1004) MAX
       Print #filnum%, max
       If (RMX = 0#) Then
          'PRINT *, ' Sun is not in the window.'
          StatusMes = " Sun is not in the window."
          Call StatusMessage(StatusMes, 1, 0)
          End If
          
       prjAtmRefMainfm.progressfrm.Visible = True
       Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
          
       For j = 1 To n + 1
          IPER = Int(j * 10# / n)
          If (IPER = (j * 10# / n)) Then
             'PRINT *, IPER*10,'%'
'            StatusMes = Str(IPER * 10) & "%"
'            Call StatusMessage(StatusMes, 1, 0)
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(j * 100# / (n + 1)))
            DoEvents
            End If
          For i = 1 To m
'1005      FORMAT (I3,1X,I3,1X,I3)
           If (RMX > 0#) Then
              IR = Int(CDbl(max) * (183# / 183#) * (CV(i, j, 1) / RMX) ^ GAM)
              IG = Int(CDbl(max) * (183# / 255#) * (CV(i, j, 2) / RMX) ^ GAM)
              IB = Int(CDbl(max) * (183# / 246#) * (CV(i, j, 3) / RMX) ^ GAM)
              IR = Int(CDbl(max) * (CV(i, j, 1) / RMX) ^ GAM)
              IG = Int(CDbl(max) * (CV(i, j, 2) / RMX) ^ GAM)
              IB = Int(CDbl(max) * (CV(i, j, 3) / RMX) ^ GAM)
           Else
              IR = 0#
              IG = 0#
              IB = 0#
              End If
'C          DRAW ASTRONOMICAL HORIZON
            If ((Abs(ALFA(KSTOP, j)) <= 1# / PPAM) And (i <= 6)) Then
               IR = 255
               IG = 255
               IB = 255
               End If
'C          DRAW SURFACE OF THE EARTH
            If ((j >= (ISTOP - 1)) And (i <= m + 1) And (ISTOP > 1)) Then
               Factor = Exp(-AIRM(j) / AIRM(n))
               Factor = Abs(CDbl((n - ISTOP) - (j - ISTOP)) / CDbl(n - ISTOP))
               IR = Int(50# * Factor)
               IG = Int(50# * Factor)
               RR = -Log(RAN(ISEED)) / 100#
               IB = Int(75# + 150# * RR)
               End If
'C          SHOW DUCTING REGION
           If ((IDCT(j) = 1) And (i <= 3)) Then IG = 255
'           WRITE(20,1005) IR,IG,IB
'           IG = IG - 68 'fix overly green image
'           If IG < 0 Then IG = 0
           Print #filnum%, IR, IG, IB
          Next i
       Next j
       Close #filnum%
       
       prjAtmRefMainfm.progressfrm.Visible = False
       Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
       
'       BrutonAtmReffm.picRef.Cls
'       DoPPM App.Path & "\temp.ppm", AtmRefPicSunfm.picRef
       'load up the file to the alphaimage box
'       Set prjAtmRefMainfm.AlphaSun.Picture = LoadPictureGDIplus(App.Path & "\temp.ppm")
'       prjAtmRefMainfm.Sunsfrm.Visible = True
'       prjAtmRefMainfm.TabRef.Tab = 3
       
       NN = 100
       MM = 200
'       PRINT *, '   Writing to PPM plot file.'
       StatusMes = " Writing to PPM temperature plot file."
       Call StatusMessage(StatusMes, 1, 0)
       For j = 1 To NN
          For i = 1 To MM
              CV(i, j, 1) = 0#
              CV(i, j, 2) = 0#
              CV(i, j, 3) = 0#
          Next i
       Next j
       RMINTMP = 1E+20
       RMAXTMP = -1E+20
       RMINELV = 1E+20
       RMAXELV = -1E+20
       For III = 0 To NNN
          If (ELV(III) <= 500#) Then
             MMM = III
             If (ELV(III) > RMAXELV) Then RMAXELV = ELV(III)
             If (ELV(III) <= RMINELV) Then RMINELV = ELV(III)
             If (TMP(III) > RMAXTMP) Then RMAXTMP = TMP(III)
             If (TMP(III) <= RMINTMP) Then RMINTMP = TMP(III)
             End If
       Next III
       For III = 0 To MMM
          IX = 3 + Int((ELV(III) - RMINELV) * (MM - 10) / (RMAXELV - RMINELV))
          IY = NN - (3 + Int((TMP(III) - RMINTMP) * (NN - 10) / (RMAXTMP - RMINTMP)))
          For IJK = 1 To 3
             CV(IX, IY, IJK) = 255
          Next IJK
       Next III
       For j = 1 To NN
          IX = 3 + Int((HOBS - RMINELV) * (MM - 10) / (RMAXELV - RMINELV))
          IY = j
         CV(IX, IY, 1) = 255
       Next j
       
       If Dir(App.Path & "\plot.ppm") <> sEmpty Then
          Kill App.Path & "\plot.ppm"
          End If
       filnum% = FreeFile
'       OPEN(UNIT=20,FILE='plot.ppm',
'     * STATUS='UNKNOWN')
       Open App.Path & "\plot.ppm" For Output As #filnum%
'       WRITE(20,2001) 'P3'
       Print #filnum%, "P3"
'       WRITE(20,1002) '# plot.ppm'
       Print #filnum%, "# plot.ppm"
'       WRITE(20,1003) MM,NN
       Print #filnum%, MM, NN
'       WRITE(20,1004) MAX
       Print #filnum%, max
       
      prjAtmRefMainfm.progressfrm.Visible = True
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
       
       For j = 1 To NN
          IPER = Int(j * 10# / NN)
          If (IPER = (j * 10# / NN)) Then
'             PRINT *, IPER*10,'%'
'            StatusMes = Str(IPER * 10) & "%"
'            Call StatusMessage(StatusMes, 1, 0)
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(j * 100# / NN))
            DoEvents
             End If
          For i = 1 To MM
             IR = Int(CV(i, j, 1))
             IG = Int(CV(i, j, 2))
             IB = Int(CV(i, j, 3))
'             WRITE(20,1005) IR,IG,IB
             Print #filnum%, IR, IG, IB
          Next i
       Next j
       Close #filnum%
      
      prjAtmRefMainfm.progressfrm.Visible = False
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
       
      
'       PRINT *, '   Writing to PPM transfer curve.'
        StatusMes = "Writing to PPM transfer curve."
        Call StatusMessage(StatusMes, 1, 0)
       For j = 1 To NN
          For i = 1 To MM
             CV(i, j, 1) = 0#
             CV(i, j, 2) = 0#
             CV(i, j, 3) = 0#
          Next i
       Next j
       RMINA = 1E+20
       RMAXA = -1E+20
       RMINT = 1E+20
       RMAXT = -1E+20
       For j = 1 To n + 1
          If (ALFT(KSTOP, j) > -900#) Then
             If (ALFT(KSTOP, j) > RMAXT) Then RMAXT = ALFT(KSTOP, j)
             If (ALFT(KSTOP, j) <= RMINT) Then RMINT = ALFT(KSTOP, j)
             If (ALFA(KSTOP, j) > RMAXA) Then RMAXA = ALFA(KSTOP, j)
             If (ALFA(KSTOP, j) <= RMINA) Then RMINA = ALFA(KSTOP, j)
             End If
       Next j
       For j = 1 To n + 1
          If (ALFT(KSTOP, j) > -900#) Then
             IX = 3 + Int((ALFA(KSTOP, j) - RMINA) * (MM - 10) / (RMAXA - RMINA))
             IY = NN - (3 + Int((ALFT(KSTOP, j) - RMINT) * (NN - 10) / (RMAXT - RMINT)))
             For IJK = 1 To 3
                 CV(IX, IY, IJK) = 255
             Next IJK
           End If
       Next j
       For jm = 1 To NM
          If (AT(jm) > -900#) Then
             IX = 3 + Int((AA(jm) - RMINA) * (MM - 10) / (RMAXA - RMINA))
             IY = NN - (3 + Int((AT(jm) - RMINT) * (NN - 10) / (RMAXT - RMINT)))
             For IJK = 1 To 1
                CV(IX, IY, IJK) = 255
             Next IJK
             End If
       Next jm
'       OPEN(UNIT=20,FILE='tran.ppm', STATUS='UNKNOWN')
       If Dir(App.Path & "\tran.ppm") <> sEmpty Then
          Kill App.Path & "\tran.ppm"
          End If
       filnum% = FreeFile
       Open App.Path & "\tran.ppm" For Output As #filnum%
'       WRITE(20,2001) 'P3'
       Print #filnum%, "P3"
'       WRITE(20,1002) '# tran.ppm'
       Print #filnum%, "# tran.ppm"
'       WRITE(20,1003) MM,NN
       Print #filnum%, MM, NN
'       WRITE(20,1004) MAX
       Print #filnum%, max
       
      prjAtmRefMainfm.progressfrm.Visible = True
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
       
       For j = 1 To NN
          IPER = Int(j * 10# / NN)
          If (IPER = (j * 10# / NN)) Then
'             PRINT *, IPER*10,'%'
'            StatusMes = Str(IPER * 10) & "%"
'            Call StatusMessage(StatusMes, 1, 0)
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(j * 100# / NN))
            DoEvents
            End If
          For i = 1 To MM
             IR = Int(CV(i, j, 1))
             IG = Int(CV(i, j, 2))
             IB = Int(CV(i, j, 3))
'             WRITE(20,1005) IR,IG,IB
'             IG = IG - 68 'fix overly green image
'             If IG < 0 Then IG = 0
             Print #filnum%, IR, IG, IB
          Next i
       Next j
       Close #filnum%
       
      prjAtmRefMainfm.progressfrm.Visible = False
      Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
       
       End If
       
       Screen.MousePointer = vbDefault

'       PRINT *, '   Program complete.'
        StatusMes = " Program complete."
        Call StatusMessage(StatusMes, 1, 0)
'       Stop

    CalcComplete = True
    PlotMode = 0
    TracesLoaded = False
    
   cmdCalc.Enabled = True
   cmdRefWilson.Enabled = True
   cmdMenat.Enabled = True
   cmdVDW.Enabled = True
    
    StatusMes = "Ray tracing calculation complete"
    Call StatusMessage(StatusMes, 1, 0)
    
    'load angle combo boxes
'    AtmRefPicSunfm.WindowState = vbMinimized
'    BrutonAtmReffm.WindowState = vbMaximized
    'set size of picref by size of earth
'    Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
'    prjAtmRefMainfm.cmbSun.Clear
'    prjAtmRefMainfm.cmbAlt.Clear
'    For i = 1 To NumSuns
'       If NumSunAlt(i - 1) > 0 Then prjAtmRefMainfm.cmbSun.AddItem i
'    Next i
   
   n_size = n
'   cmbSun.ListIndex = 0
   
   Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
   prjAtmRefMainfm.cmbSun.Clear
   prjAtmRefMainfm.cmbAlt.Clear
   For i = 1 To NumSuns
      If NumSunAlt(i - 1) > 0 Then prjAtmRefMainfm.cmbSun.AddItem i
   Next i
     
   prjAtmRefMainfm.TabRef.Tab = 4
   DoEvents
    
  cmbSun.ListIndex = 0
    
   On Error GoTo 0
   Exit Sub

cmdCalc_Click_Error:
    Close
    Screen.MousePointer = vbDefault
    
    StatusMes = sEmpty
    Call StatusMessage(StatusMes, 1, 0)
    Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
    prjAtmRefMainfm.progressfrm.Visible = False
    
   cmdCalc.Enabled = True
   cmdRefWilson.Enabled = True
   cmdMenat.Enabled = True
   cmdVDW.Enabled = True
   
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdCalc_Click of Form prjAtmRefMainfm", vbCritical + vbOKOnly
End Sub

Private Sub cmdZoomIn_Click()
   ZoomValue = 1.1 * ZoomValue
   AlphaSun.Refresh
End Sub

Private Sub cmdZoomOut_Click()
   ZoomValue = (1# / 1.1) * ZoomValue
   AlphaSun.Refresh
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCalcTR_Click
' Author    : Dr-John-K-Hall
' Date      : 8/8/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdCalcTR_Click()
   'view angle without refraction, including curvature of earth
    Dim H1 As Double, H2 As Double
    Dim lt1 As Double, lg1 As Double
    Dim lt2 As Double, lg2 As Double
    Dim Rearth As Double, X1 As Double, X2 As Double, Y1 As Double, Y2 As Double
    Dim z1 As Double, z2 As Double, re1 As Double, re2 As Double
    Dim dist1 As Double, dist2 As Double, ANGLE As Double
    Dim viewang As Double, H11 As Double, H12 As Double
    Dim SLOPEfound As Double
    Dim H21 As Double, H22 As Double
    Dim D1 As Double, D2 As Double, T1 As Double, T2 As Double
    Dim JustTLOW As Boolean
    Dim JustTHIGH As Boolean
    Dim JustHLOW As Boolean
    Dim JustHHIGH As Boolean
    Dim VATHIGH As Double
    Dim VATLOW As Double
    Dim Difference As Double
    Dim DistTolerance As Double 'meters
    Dim VA11 As Double, VA21 As Double, VA12 As Double, VA22 As Double, VAfinal As Double
    Dim Dist11 As Double, Dist21 As Double, Dist12 As Double, Dist22 As Double
    Dim height11 As Double, height21 As Double, height12 As Double, height22 As Double
    Dim j As Long
    Dim FileMode As Integer
    Dim PL1 As Double, PL2 As Double, PL10 As Double, PL20 As Double, PL11 As Double
    Dim PL12 As Double, PL21 As Double, PL22 As Double, PATHLENGTH As Double
    Dim VA1 As Double, VA2 As Double, VA10 As Double, VA20 As Double
    Dim hgt1 As Double, hgt2 As Double
    
    Dim StartAng As Double, EndAng As Double, StepAng As Double, NAngles As Long
    Dim TempStart As Double, TempEnd As Double, RecordTLoop As Boolean
    Dim FilePath As String, StepSize As Integer, Press0 As Double, WAVELN As Double
    Dim HUMID As Double, OBSLAT As Double, NSTEPS As Long, RELHUM As Double
    Dim VA As Double, LastVA As Double, distR As Double, TC As Double, dlon As Double

    Dim TempValue As Double, DistValue As Double, CalcDist As Double, TotalDist As Double
    
    Set PtX = New Collection
    Set PtY = New Collection
  
   On Error GoTo cmdCalcTR_Click_Error

    DistTolerance = 1 'search tolerance in meters of obstruction height
    
    FileMode = 1 '= 1 for 5 numbers per line
                 '= 2 for 6 numbers per line (newer file type with PATHLENGTH)
   
    pi = 4# * Atn(1#) '3.141592654
    cd = pi / 180# 'conversion from degrees to radians
    Rearth = 6356766#
    RE = Rearth

   If txtDirTR = "Browse for TR files" Then
      'first try G:/AtmRef
      If FileMode = 1 Then
         txtDirTR = App.Path
      Else
        If Dir("E:/AtmRef", vbDirectory) <> sEmpty Then
           txtDirTR = "E:/AtmRef"
           txtDirTR.Refresh
           DoEvents
        Else
      
            Call MsgBox("You must define the path to the TR files!" _
                        & vbCrLf & sEmpty _
                        & vbCrLf & "Use the ""Browse"" button" _
                        , vbExclamation, "Missing TR folder")
            Exit Sub
            
            End If
            
        End If
        
      End If
     
     'use first place as Jerusalem
     'calculate the coordinates of the second place
'     lg1 = -35.2166667
'     lt1 = 31.7833333 ',762
'     lg2 = lg1 - (D1 / Rearth) / cd 'place second point to the East of first point and at same latitude
'     lt2 = lt1
     'now check the distance
     
'     lg2 = -35.8612429850206
'     lt2 = 31.8909320460135
'     H21 = 959.2
'
'     'Rabbi Druk's shul at Armon Hanatziv
'     lg1 = -35.238133306709
'     lt1 = 31.7487155576439
'     H11 = 756.5 <-- added 1.8, should be 754.7
'
'     H21 = 100 + H21 - H11
'     H11 = 100
     
'     D1 = Rearth * DistTrav(lt1, lg1, lt2, lg2, 1)
     

'    ++numdtm;
'    dtms[numdtm].ver[0] = azi;
'    dtms[numdtm].ver[1] = viewang / cd + avref;
'    dtms[numdtm].ver[2] = kmy;

    'convert to mrad
  
'    lblTR_Est = "View Angle w/o Ref. (deg.): " & Format(Str(Val(viewang / cd)), "##0.0####")
'    lblTR_Ref = "View Angle with Ter. Ref. estimate (deg.): " & Format(Str(Val(VAfinal / 60#)), "##0.0####")
   
'   Rearth = 6378136.6
'   VA = DACOS(RE / (Rearth + (H1 - H2)))
'   'estimate of terrestrial refraction
'   TR = (16.3 * l * P / (T ^ 2)) * (0.0342 + ALPHA) * Cos(VA)
           
      
   If chkH1.Value = vbChecked Then
   
      T1 = Val(txtT11.Text)
      T2 = T1
      D1 = Val(txtD1.Text) * 1000
      D2 = D1
      H11 = Val(txtH11.Text)
      H12 = Val(txtH12.Text)
      H21 = Val(txtH21.Text)
      H22 = H21
      
    '-------------------------------------------------
    With prjAtmRefMainfm
      '------fancy progress bar settings---------
      .TRprogfrm.Visible = True
      .picProgBarTR.AutoRedraw = True
      .picProgBarTR.BackColor = &H8000000B 'light grey
      .picProgBarTR.DrawMode = 10
    
      .picProgBarTR.FillStyle = 0
      .picProgBarTR.ForeColor = &H400000 'dark blue
      .picProgBarTR.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
   
   ElseIf chkH2.Value = vbChecked Then
      T1 = Val(txtT11.Text)
      T2 = T1
      D1 = Val(txtD1.Text) * 1000
      D2 = D1
      H11 = Val(txtH11.Text)
      H12 = H11
      H22 = Val(txtH22.Text)
      H21 = Val(txtH21.Text)
      
    '-------------------------------------------------
    With prjAtmRefMainfm
      '------fancy progress bar settings---------
      .TRprogfrm.Visible = True
      .picProgBarTR.AutoRedraw = True
      .picProgBarTR.BackColor = &H8000000B 'light grey
      .picProgBarTR.DrawMode = 10
    
      .picProgBarTR.FillStyle = 0
      .picProgBarTR.ForeColor = &H400000 'dark blue
      .picProgBarTR.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
      
   ElseIf chkDist.Value = vbChecked Then
      
        Select Case MsgBox("Did you set the temperature for the calculation?" _
                & vbCrLf & "(It is set in the T1 box of the temperature scan)" _
                & vbCrLf & "" _
                & vbCrLf & "Did you set the observer height?" _
                & vbCrLf & "(It is set in the first box of the observer height scan)" _
                & vbCrLf & "" _
                & vbCrLf & "Dd  you set the obstruction height?" _
                & vbCrLf & "(It is set in the first box of the obsruction height scan)" _
                , vbQuestion + vbOKCancel, "Set the temperature of the calculation")
                  
            Case vbOK
            
            Case vbCancel
               prjAtmRefMainfm.picProgBarTR.Visible = False
               prjAtmRefMainfm.TRprogfrm.Visible = False
               Screen.MousePointer = vbDefault
               Exit Sub
        End Select
                 
          T1 = Val(txtT11.Text)
          TempValue = T1
          T2 = T1
          D1 = Val(txtD1.Text) * 1000
          H11 = Val(txtH11.Text)
          H12 = H11
          D2 = Val(txtD2.Text) * 1000
          H21 = Val(txtH21.Text)
          H22 = H21
          
        'Rabbi Druk's shul at Armon Hanatziv
        lg1 = -35.238133306709
        lt1 = 31.7487155576439
          
'        distR = D1 / Rearth
'        distR = distR / cd
'
'        TC = 0
'        lt2 = DASIN(Sin(lt1 * cd) * Cos(distR * cd) + Cos(lt1 * cd) * Sin(distR * cd) * Cos(TC))
'        lt2 = lt2 / cd
'        lg2 = lg1
     
     
'     If (Cos(lt2) = 0) Then
'        lg2 = lg1    ' endpoint a pole
'     Else
'        lg2 = ((lg1 * cd - DASIN(Sin(TC) * Sin(distR) / Cos(lt2 * cd)) + PI) Mod (2 * PI)) - PI
'        lg2 = lg2 / cd
'     End If
     
'        TC = 0 'parralel longitudes
'        lt2 = DASIN(Sin(lt1 * cd) * Cos(distR) + Cos(lt1) * Sin(distR) * Cos(TC))
'        dlon = Atan2(Sin(TC) * Sin(distR) * Cos(lt1 * cd), Cos(distR) - Sin(lt1 * cd) * Sin(lt1 * cd))
'        lg2 = (lg1 * cd - dlon + PI) Mod (2 * PI) - PI
       
'        D1 = Rearth * DistTrav(lt1, lg1, lt2, lg2, 1)
          
        '-------------------------------------------------
        With prjAtmRefMainfm
          '------fancy progress bar settings---------
          .TRprogfrm.Visible = True
          .picProgBarTR.AutoRedraw = True
          .picProgBarTR.BackColor = &H8000000B 'light grey
          .picProgBarTR.DrawMode = 10
        
          .picProgBarTR.FillStyle = 0
          .picProgBarTR.ForeColor = &H400000 'dark blue
          .picProgBarTR.Visible = True
        End With
        pbScaleWidth = 100
        '-------------------------------------------------
       
    '     D1 = Rearth * DistTrav(lt1, lg1, lt2, lg2, 1)
         
        Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0) 'reset
        DoEvents
        
        'now save the results in a file
        fileout% = FreeFile
        RecordTempTR$ = txtDirTR & "/TRC_T-" & Trim$(Val(T1)) & "-" & Trim$(Val(T2)) & "_HOSV-" & Trim$(Val(H11)) & "-" & Trim$(Val(H12)) & "_DOBST-" & Trim$(Val(D1)) & "-" & Trim$(Val(D2)) & "_HOBST-" & Trim$(Val(H21)) & "-" & Trim$(Val(H22)) & ".dat"
        
        Open RecordTempTR$ For Output As #fileout%
        
        NNN = (D2 - D1) / Val(txtStepD1 * 1000) + 1
        
        'now load up temperature and pressure charts
        
    '    ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
        
        Call UpdateStatus(prjAtmRefMainfm, picProgBarTR, 1, 0)
    
        j = 0
    '    maxva = -9999999
    '    minva = 9999999
        
        FilePath = txtDirTR
        StepSie = 1
        RecordTLoop = True
        FileMode = 1 'mode used for determination of terrestrial refraction using the dll
        
        Press0 = Val(txtPress0)
        HMAXT = Val(txtHMAXT)
        RELHUM = Val(txtRELHUM)
        StartAng = Val(txtBETAHI) * 60# 'convert to arc minutes
        EndAng = Val(txtBETALO) * 60#
        StepAng = Val(txtBETAST) * 60#
        WAVELN = Val(txtKmin) * 0.001 'Val(txtWAVELN)
        OBSLAT = Val(txtOBSLAT)
        NSTEPS = Val(txtNSTEPS)
        If NSTEPS < 5000 Then NSTEPS = 5000
        HUMID = RELHUM
        HOBS = H11
        StepSize = Val(prjAtmRefMainfm.txtHeightStepSize.Text)
        NAngles = 2 * StartAng / StepAng + 1
        LastVA = 9999999 'insure proper temperature progression, which should be proportional to the inverse square of the temperature
        
        Screen.MousePointer = vbHourglass
                
        Call UpdateStatus(prjAtmRefMainfm, picProgBarTR, 1, 0) 'reset
        
        For DistValue = D1 To D2 Step Val(txtStepD1 * 1000#)
        
        
          distR = DistValue / Rearth
          distR = distR / cd
        
          TC = 0
          lt2 = DASIN(Sin(lt1 * cd) * Cos(distR * cd) + Cos(lt1 * cd) * Sin(distR * cd) * Cos(TC)) 'modified from Aviation Forumulary
          lt2 = lt2 / cd
          lg2 = lg1
       
          CalcDist = Rearth * DistTrav(lt1, lg1, lt2, lg2, 1)
            
          If chkUseDll.Value = vbUnchecked Then
            'interpolate among the TR files (slow and inaccurate)
            GoSub SingleCalc
          Else
            'use the dll for a much faster and more accurate calculation
                
            GoSub VAsub 'calculate the vuew angle without refraction
    
            VA = viewang
            ier = RayTracing(StartAng, EndAng, StepAng, LastVA, NAngles, _
                             CalcDist, VA, H21, DistTolerance, FileMode, _
                             H11, TempValue, HMAXT, FilePath, StepSize, _
                             Press0, WAVELN, HUMID, OBSLAT, NSTEPS, _
                             RecordTLoop, T1, T2, AddressOf MyCallback)
            VAfinal = LastVA 'calculated view angle in radians
            End If
            
          'add to buffer
          j = j + 1
    
          'increment progress bar
          Call UpdateStatus(prjAtmRefMainfm, picProgBarTR, 1, CLng(100# * j / NNN))
    
    '      TransferCurve(j, 1) = " " & CStr(TempValue)
    '      TR = VAfinal / cd - viewang / cd
    '      TransferCurve(j, 2) = TR
    '
    '      If TR > maxva Then maxva = TR
    '      If TR < minva Then minva = TR
          
          Print #fileout%, Format(Str$(CalcDist * 0.001), "##0.0###"), Format(Str$(VAfinal / cd), "##0.0#####"), Format(Str$(viewang / cd), "##0.0#####"), Format(Str$((VAfinal - viewang) / cd), "##0.0#####")
          
        Next DistValue
        
        Close #fileout%
        
        Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0) 'reset
        TRprogfrm.Visible = False
        
       'calculate a polynomial fit using now plot the terrestrial refraction calculation results
       ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
       
       maxva = -999999999
       minva = 999999999
         
       filein% = FreeFile
       Open RecordTempTR$ For Input As #filein%
         
       j = 0
       TotalDist = 0
       Do While Not EOF(filein%)
          Input #filein%, Temp, A, b, TR
          TotalDist = D1 * 0.001 + j * Val(txtStepD1)
          j = j + 1
          TransferCurve(j, 1) = " " & CStr(TotalDist)
          TransferCurve(j, 2) = TR
          PtX.Add TotalDist
          'caculate the approximate ray distance
'            //use Lehn's parabolic path approx to ray trajectory and Brutton equation 58
          PATHLENGTH = Sqr(TotalDist ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (TotalDist ^ 2#) / (Rearth * 0.001)) ^ 2#)
          
          PtY.Add TotalDist * (TR * TempValue * TempValue) / (0.0083 * PATHLENGTH * Press0)
        
          If TR > maxva Then maxva = TR
          If TR < minva Then minva = TR
          
       Loop
       Close #filein%
       
        ' Find a good fit.
        degree = 1
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
    
    '    stop_time = Timer
    '    Debug.Print Format$(stop_time - start_time, "0.0000")
    
        Txt = ""
        For i = 1 To BestCoeffs.Count
            Txt = Txt & " " & BestCoeffs.Item(i)
        Next i
        If Len(Txt) > 0 Then Txt = Mid$(Txt, 2)
        txtAs.Text = Txt
    
        ' Display the error.
        Call ShowError(txtError)
    
        ' We have a solution.
        HasSolution = True
    '    picGraph.Refresh
       
        With MSChartTR
          .chartType = VtChChartType2dLine
          .RandomFill = False
          .RowCount = 2
          .ColumnCount = NNN
          .RowLabel = "Distance between observer and obstruction (kms)"
          .ColumnLabel = "Terrestrial refraction (degrees)"
          .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
          .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
          '      .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = Format(((maxva - minva) \ NNN), "##0.####0")
          '      .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 10
          .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1.1 * maxva
          .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0.9 * minva
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = D1
          .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = D2
          .ChartData = TransferCurve
        End With
    
        Screen.MousePointer = vbDefault
      
   ElseIf chkTemp.Value = vbChecked Then
   
        lg2 = -35.8612429850206
        lt2 = 31.8909320460135
        H21 = 959.2
        
        'Rabbi Druk's shul at Armon Hanatziv
        lg1 = -35.238133306709
        lt1 = 31.7487155576439
        H11 = 756.5
        
        H21 = 100 + H21 - H11
        H11 = 100
     
          T1 = Val(txtT11.Text)
    '      D1 = Val(txtD1.Text) * 1000
          D2 = D1
    '      H11 = Val(txtH11.Text)
          H12 = H11
          T2 = Val(txtT12.Text)
    '      H21 = Val(txtH21.Text)
          H22 = H21
          
        GoSub VAsub 'calculate the vuew angle without refraction
        
         D1 = Rearth * DistTrav(lt1, lg1, lt2, lg2, 1)
         
      
        lblTR_Est = "View Angle w/o Ref. (deg.): " & Format(Str(Val(viewang / cd)), "##0.0####")
        lblTR_Ref = "View Angle with Ter. Ref. estimate (deg.): " & Format(Str(Val(VAfinal / 60#)), "##0.0####")
        
        
        '-------------------------------------------------
        With prjAtmRefMainfm
          '------fancy progress bar settings---------
          .TRprogfrm.Visible = True
          .picProgBarTR.AutoRedraw = True
          .picProgBarTR.BackColor = &H8000000B 'light grey
          .picProgBarTR.DrawMode = 10
        
          .picProgBarTR.FillStyle = 0
          .picProgBarTR.ForeColor = &H400000 'dark blue
          .picProgBarTR.Visible = True
        End With
        pbScaleWidth = 100
        '-------------------------------------------------
        
        Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0) 'reset
        DoEvents
        
        'now save the results in a file
        fileout% = FreeFile
        RecordTempTR$ = txtDirTR & "/TRC_T-" & Trim$(Val(T1)) & "-" & Trim$(Val(T2)) & "_HOSV-" & Trim$(Val(H11)) & "-" & Trim$(Val(H12)) & "_DOBST-" & Trim$(Val(D1)) & "-" & Trim$(Val(D2)) & "_HOBST-" & Trim$(Val(H21)) & "-" & Trim$(Val(H22)) & ".dat"
        
        Open RecordTempTR$ For Output As #fileout%
        
        NNN = (T2 - T1) / Val(txtStepT) + 1
        
        'now load up temperature and pressure charts
        
    '    ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
        
        Call UpdateStatus(prjAtmRefMainfm, picProgBarTR, 1, 0)
    
        j = 0
    '    maxva = -9999999
    '    minva = 9999999
        
        FilePath = txtDirTR
        StepSie = 1
        RecordTLoop = True
        FileMode = 1 'mode used for determination of terrestrial refraction using the dll
        
        Press0 = Val(txtPress0)
        HMAXT = Val(txtHMAXT)
        RELHUM = Val(txtRELHUM)
        StartAng = Val(txtBETAHI) * 60# 'convert to arc minutes
        EndAng = Val(txtBETALO) * 60#
        StepAng = Val(txtBETAST) * 60#
        WAVELN = Val(txtKmin) * 0.001 'Val(txtWAVELN)
        OBSLAT = Val(txtOBSLAT)
        NSTEPS = Val(txtNSTEPS)
        If NSTEPS < 5000 Then NSTEPS = 5000
        HUMID = RELHUM
        HOBS = H11
        StepSize = Val(prjAtmRefMainfm.txtHeightStepSize.Text)
        NAngles = 2 * StartAng / StepAng + 1
        LastVA = 9999999 'insure proper temperature progression, which should be proportional to the inverse square of the temperature
        
        Screen.MousePointer = vbHourglass
                
        Call UpdateStatus(prjAtmRefMainfm, picProgBarTR, 1, 0) 'reset
        
        For TempValue = T1 To T2 Step Val(txtStepT)
        
          If chkUseDll.Value = vbUnchecked Then
            'interpolate among the TR files (slow and inaccurate)
            GoSub SingleCalc
          Else
            'use the dll for a much faster and more accurate calculation
            
            VA = viewang
            ier = RayTracing(StartAng, EndAng, StepAng, LastVA, NAngles, _
                             D1, VA, hgt2, DistTolerance, FileMode, _
                             HOBS, TempValue, HMAXT, FilePath, StepSize, _
                             Press0, WAVELN, HUMID, OBSLAT, NSTEPS, _
                             RecordTLoop, T1, T2, AddressOf MyCallback)
            If ier = 0 Then
                VAfinal = LastVA 'calculated view angle in radians
            Else
                'didn't converge,no TR
                VAfinal = viewang
                End If
            End If
            
          'add to buffer
          j = j + 1
    
          'increment progress bar
          Call UpdateStatus(prjAtmRefMainfm, picProgBarTR, 1, CLng(100# * j / NNN))
    
    '      TransferCurve(j, 1) = " " & CStr(TempValue)
    '      TR = VAfinal / cd - viewang / cd
    '      TransferCurve(j, 2) = TR
    '
    '      If TR > maxva Then maxva = TR
    '      If TR < minva Then minva = TR
          
          Print #fileout%, Format(Str$(TempValue), "##0.0#"), Format(Str$(VAfinal / cd), "##0.0#####"), Format(Str$(viewang / cd), "##0.0#####"), Format(Str$((VAfinal - viewang) / cd), "##0.0#####")
          
        Next TempValue
        
        Close #fileout%
        
        Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0) 'reset
        TRprogfrm.Visible = False
        
        
       'calculate a polynomial fit using now plot the terrestrial refraction calculation results
       ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
       
       maxva = -999999999
       minva = 999999999
         
       filein% = FreeFile
       Open RecordTempTR$ For Input As #filein%
         
       j = 0
       Do While Not EOF(filein%)
          Input #filein%, Temp, A, b, TR
          j = j + 1
          TransferCurve(j, 1) = " " & CStr(Temp)
          TransferCurve(j, 2) = TR
          
          PATHLENGTH = Sqr(TotalDist ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (TotalDist ^ 2#) / (Rearth * 0.001)) ^ 2#)
          
          PtX.Add Temp
          PtY.Add Temp * (TR * Temp * Temp) / (0.0083 * PATHLENGTH * 0.001 * Press0)
          
          If TR > maxva Then maxva = TR
          If TR < minva Then minva = TR
          
       Loop
       Close #filein%
       
        ' Find a good fit.
        degree = 1
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
    
    '    stop_time = Timer
    '    Debug.Print Format$(stop_time - start_time, "0.0000")
    
        Txt = ""
        For i = 1 To BestCoeffs.Count
            Txt = Txt & " " & BestCoeffs.Item(i)
        Next i
        If Len(Txt) > 0 Then Txt = Mid$(Txt, 2)
        txtAs.Text = Txt
    
        ' Display the error.
        Call ShowError(txtError)
    
        ' We have a solution.
        HasSolution = True
    '    picGraph.Refresh
       
        With MSChartTR
          .chartType = VtChChartType2dLine
          .RandomFill = False
          .RowCount = 2
          .ColumnCount = NNN
          .RowLabel = "Temperature (degrees Kelvin)"
          .ColumnLabel = "Terrestrial refraction (degrees)"
          .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
          .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
          '      .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = Format(((maxva - minva) \ NNN), "##0.####0")
          '      .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 10
          .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1.1 * maxva
          .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0.9 * minva
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = T1
          .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = T2
          .ChartData = TransferCurve
        End With
    
        Screen.MousePointer = vbDefault
        
    
   Else
   
      T1 = Val(txtT11.Text)
      D1 = Val(txtD1.Text) * 1000
      H11 = Val(txtH11.Text)
      H21 = Val(txtH21.Text)
      
'      T1 = 298.2
'      D1 = 23.2 * 1000
'      H11 = 232.5
'      H21 = 2832

    TempValue = T1
      
    GoSub SingleCalc
      
    End If
    
    Exit Sub
    
'this is calculation for one series of parameters
SingleCalc:
      
   Screen.MousePointer = vbHourglass
     
     PATHLENGTH = 0#
     
     'now open correct files and interpolate
     'determine which temperatures to interpolate
     TMax = Val(txtET)
     TMin = Val(txtST)
     TSTEP = Val(txtTS)
     'determine between which temperatures to interpolate
     TLOW = Fix((TempValue - TMin) / TSTEP) * TSTEP + TMin
     THIGH = TLOW + TSTEP
     JustTLOW = False
     JustTHIGH = False
     If TempValue = TLOW Then JustTLOW = True
     If TempValue = THIGH Then JustTHIGH = True
     
     'now do the same for the observer's height
     HGTMAX = Val(txtEHgt)
     HGTMIN = Val(txtBHgt)
     HGTSTP = Val(txtSHgt)
     
     H1LOW = Fix((H11 - HGTMIN) / HGTSTP) * HGTSTP + HGTMIN
     H1HIGH = H1LOW + HGTSTP
     If H11 = H1LOW Then JustHLOW = True
     If H11 = H1HIGH Then JustHHIGH = True
     
     'now do the same for the obstruction's height
     H2LOW = Fix((H21 - HGTMIN) / HGTSTP) * HGTSTP + HGTMIN
     H2HIGH = H2LOW + HGTSTP
     
     VATHIGH = 0
     VATLOW = 0
     found1% = 0
     found2% = 0
     Difference = 999999
         
     If JustTLOW Then GoTo 1500
     
     'generate the name of the two files
     FileT1H1High$ = txtDirTR.Text & "\TR_VDW_" & Trim$(Str$(THIGH)) & "_" & Trim$(Str$(H1HIGH)) & "_32.dat"
     FileLow$ = txtDirTR.Text & "\TR_VDW_" & Trim$(Str$(THIGH)) & "_" & Trim$(Str$(H1LOW)) & "_32.dat"
     
     If Dir(FileT1H1High$) = sEmpty Or _
        Dir(FileLow$) = sEmpty Then
        Screen.MousePointer = vbDefault
        Call MsgBox("Can't find file(s)", vbCritical, "File doesn't exist")
     Else
        'move file to hard disk on computer to speed up analysis
        If Not JustHLOW Then
            FileCopy FileT1H1High$, App.Path & "\TR_VDW_" & Trim$(Str$(THIGH)) & "_" & Trim$(Str$(H1HIGH)) & "_32.dat"
            FileT1H1High$ = App.Path & "\TR_VDW_" & Trim$(Str$(THIGH)) & "_" & Trim$(Str$(H1HIGH)) & "_32.dat"
            End If
        If Not JustHHIGH Then
            FileCopy FileLow$, App.Path & "\TR_VDW_" & Trim$(Str$(THIGH)) & "_" & Trim$(Str$(H1LOW)) & "_32.dat"
            FileLow$ = App.Path & "\TR_VDW_" & Trim$(Str$(THIGH)) & "_" & Trim$(Str$(H1LOW)) & "_32.dat"
            End If
        End If
        
    If JustHLOW Then GoTo 500
     
     'now open files and find the interpolated view angle for the requsted DIST and H2
     
     file1% = FreeFile
     Open FileT1H1High$ For Input As #file1%
     found1% = 0
     Difference = 9999999
     npnt& = 0
     Do Until EOF(file1%)
        If FileMode = 1 Then
            Input #file1%, dist1, height1, VA1, Dip1, DIF1
        ElseIf FileMode = 2 Then
            Input #file1%, PL1, dist1, height1, VA1, Dip1, DIF1
            If npnt& <> 0 Then
               PATHLENGTH = PATHLENGTH + Sqr(((Rearth + H2) * Sin(VA10)) ^ 2# + 2# * Rearth * (H2 - H1) + H2 * H2 - H1 * H1) - (Rearth + H1) * Sin(VA10)
               End If
            End If
            
        If height1 = -1000 Then Exit Do 'ray hit ground
        
        If (npnt& = 0) Then
            dist10 = dist1
            height10 = height1
            VA10 = VA1
            Dip10 = Dip1
            DIF10 = DIF1
            PL10 = PL1
        Else
            If D1 >= dist10 And D1 < dist1 Then
               If H21 >= height10 And H21 < height1 Then
               
                  'interpolate
                  slope1 = (height1 - height10) / (dist1 - dist10)
                  height1fit1 = (D1 - dist10) * slope1 + height10
                  found1% = 1
                  VA11 = VA1
                  Dist11 = dist1
                  height11 = height1
                  PL11 = PL1
                  Exit Do
                  
               Else
                  'interpolate and look for best fit
                  SLOPE2 = (height1 - height10) / (dist1 - dist10)
                  height1fit2 = (D1 - dist10) * SLOPE2 + height10
                  If Abs(H21 - height1fit2) < Difference Then
                     Difference = Abs(H21 - height1fit2)
                     VA11 = VA1
                     Dist11 = dist1
                     height11 = height1
                     PL11 = PL1
                     End If
                  End If
                End If

            dist10 = dist1
            height10 = height1
            VA10 = VA1
            Dip10 = Dip1
            DIF10 = DIF1
            PL10 = PL1
            End If
            
        npnt& = npnt& + 1
        
     Loop
     Close #file1%
     
     If found1% = 0 And Difference <= DistTolerance Then
        found1% = 1
     ElseIf found1% = 0 And Difference > DistTolerance Then
        Call MsgBox("Search at THIGH and HHIGH was unsuccessful" _
                    & vbCrLf & sEmpty _
                    & vbCrLf & "Increase the Distance Tolerance!" _
                    , vbExclamation, "No result returned")
        
        End If
     
500:
     If JustHHIGH Then GoTo 1000
     found2% = 0
     Difference = 999999
     file2% = FreeFile
     Open FileLow$ For Input As #file2%
     npnt& = 0
     Do Until EOF(file2%)
        If FileMode = 1 Then
            Input #file2%, dist2, height2, VA2, Dip2, DIF2
        ElseIf FileMode = 2 Then
            Input #file2%, PL2, dist2, height2, VA2, Dip2, DIF2
            End If
        If height2 = -1000 Then Exit Do 'ray hit ground
        
        If (npnt& = 0) Then
            dist20 = dist2
            height20 = height2
            VA20 = VA2
            Dip20 = Dip2
            DIF20 = DIF2
            PL20 = PL2
        Else
            If D1 >= dist20 And D1 < dist2 Then
               If H21 >= height20 And H21 < height2 Then
               
                  'interpolate
                  SLOPE2 = (height2 - height20) / (dist2 - dist20)
                  height1fit2 = (D1 - dist20) * SLOPE2 + height20
                  VA12 = VA2
                  Dist12 = dist2
                  height12 = height2
                  PL12 = PL2
                  found2% = 1
                  Exit Do
                  
               Else
                  'interpolate and look for best fit
                  SLOPE2 = (height2 - height20) / (dist2 - dist20)
                  height1fit2 = (D1 - dist20) * SLOPE2 + height20
                  If Abs(H21 - height1fit2) < Difference Then
                     Difference = Abs(H21 - height1fit2)
                     VA12 = VA2
                     Dist12 = dist2
                     height12 = height2
                     PL12 = PL2
                     End If
                  End If
                End If

            dist20 = dist2
            height20 = height2
            VA20 = VA2
            Dip20 = Dip2
            DIF20 = DIF2
            PL20 = PL2
            End If
            
        npnt& = npnt& + 1
        
     Loop
     Close #file2%
     
     If found2% = 0 And Difference <= DistTolerance Then
        found2% = 1
     ElseIf found2% = 0 And Difference > DistTolerance Then
        Call MsgBox("Search at THIGH and HLOW was unsuccessful" _
                    & vbCrLf & sEmpty _
                    & vbCrLf & "Increase the Distance Tolerance!" _
                    , vbExclamation, "No result returned")
        
        End If
        
1000:
     If JustHHIGH Then
        VATHIGH = VA11
     ElseIf JustHLOW Then
        VATHIGH = VA12
     Else
     
        If found1% = 1 And found2% = 1 Then
           SLOPEfound = (VA11 - VA12) / (H1HIGH - H1LOW)
           VATHIGH = (H11 - H1LOW) * SLOPEfound + VA12
           End If
           
        End If
     
1500:
     If Dir(FileT1H1High$) <> sEmpty Then Kill FileT1H1High$
     If Dir(FileLow$) <> sEmpty Then Kill FileLow$
     
     'now redo for lower temperature
     found1% = 0
     found2% = 0
     Difference = 999999
     
     If JustTHIGH Then GoTo 3000
     
     'now open files and find the interpolated view angle for the requsted DIST and H2
     FileT1H1High$ = txtDirTR.Text & "\TR_VDW_" & Trim$(Str$(TLOW)) & "_" & Trim$(Str$(H1HIGH)) & "_32.dat"
     FileLow$ = txtDirTR.Text & "\TR_VDW_" & Trim$(Str$(TLOW)) & "_" & Trim$(Str$(H1LOW)) & "_32.dat"
     If Dir(FileT1H1High$) = sEmpty Or _
        Dir(FileLow$) = sEmpty Then
        Screen.MousePointer = vbDefault
        Call MsgBox("Can't find file(s)", vbCritical, "File doesn't exist")
     Else
        'move file to hard disk on computer to speed up analysis
        If Not JustHLOW Then
            FileCopy FileT1H1High$, App.Path & "\TR_VDW_" & Trim$(Str$(TLOW)) & "_" & Trim$(Str$(H1HIGH)) & "_32.dat"
            FileT1H1High$ = App.Path & "\TR_VDW_" & Trim$(Str$(TLOW)) & "_" & Trim$(Str$(H1HIGH)) & "_32.dat"
            End If
        If Not JustHHIGH Then
            FileCopy FileLow$, App.Path & "\TR_VDW_" & Trim$(Str$(TLOW)) & "_" & Trim$(Str$(H1LOW)) & "_32.dat"
            FileLow$ = App.Path & "\TR_VDW_" & Trim$(Str$(TLOW)) & "_" & Trim$(Str$(H1LOW)) & "_32.dat"
            End If
        End If
        
     'now open files and find the interpolated view angle for the requsted DIST and H2
     If JustHLOW Then GoTo 2000
     file1% = FreeFile
     Open FileT1H1High$ For Input As #file1%
     npnt& = 0
     found1% = 0
     Difference = 99999999
     Do Until EOF(file1%)
        If FileMode = 1 Then
            Input #file1%, dist1, height1, VA1, Dip1, DIF1
        ElseIf FileMode = 2 Then
            Input #file1%, PL1, dist1, height1, VA1, Dip1, DIF1
            End If
        If height1 = -1000 Then Exit Do 'ray hit ground
        
        If (npnt& = 0) Then
            dist10 = dist1
            height10 = height1
            VA10 = VA1
            Dip10 = Dip1
            DIF10 = DIF1
            PL10 = PL1
        Else
            If D1 >= dist10 And D1 < dist1 Then
               If H21 >= height10 And H21 < height1 Then
               
                  'interpolate
                  slope1 = (height1 - height10) / (dist1 - dist10)
                  height1fit1 = (D1 - dist10) * slope1 + height10
                  VA21 = VA1
                  Dist21 = dist1
                  height21 = height1
                  PL21 = PL1
                  found1% = 1
                  Exit Do
                  
               Else
                  'interpolate and look for best fit
                  SLOPE2 = (height1 - height10) / (dist1 - dist10)
                  height1fit2 = (D1 - dist10) * SLOPE2 + height10
                  If Abs(H21 - height1fit2) < Difference Then
                     Difference = Abs(H21 - height1fit2)
                     VA21 = VA1
                     Dist21 = dist1
                     height21 = height1
                     PL21 = PL1
                     End If
                  End If
                End If

            dist10 = dist1
            height10 = height1
            VA10 = VA1
            Dip10 = Dip1
            DIF10 = DIF1
            PL10 = PL1
            End If
            
        npnt& = npnt& + 1
        
     Loop
     Close #file1%
     
     If found1% = 0 And Difference <= DistTolerance Then
        found1% = 1
     ElseIf found1% = 0 And Difference > DistTolerance Then
        Call MsgBox("Search at TLOW and HHIGH was unsuccessful" _
                    & vbCrLf & sEmpty _
                    & vbCrLf & "Increase the Distance Tolerance!" _
                    , vbExclamation, "No result returned")
        
        End If
        
2000:
     If JustHHIGH Then GoTo 2500
     Difference = 999999
     found2% = 0
     file2% = FreeFile
     Open FileLow$ For Input As #file2%
     npnt& = 0
     Do Until EOF(file2%)
        If FileMode = 1 Then
            Input #file2%, dist2, height2, VA2, Dip2, DIF2
        ElseIf FileMode = 2 Then
            Input #file2%, PL2, dist2, height2, VA2, Dip2, DIF2
            End If
        If height2 = -1000 Then Exit Do 'ray hit ground
        
        If (npnt& = 0) Then
            dist20 = dist2
            height20 = height2
            VA20 = VA2
            Dip20 = Dip2
            DIF20 = DIF2
            PL20 = PL2
        Else
            If D1 >= dist20 And D1 < dist2 Then
               If H21 >= height20 And H21 < height2 Then
               
                  'interpolate
                  SLOPE2 = (height2 - height20) / (dist2 - dist20)
                  height1fit2 = (D1 - dist20) * SLOPE2 + height20
                  VA22 = VA2
                  Dist22 = dist2
                  height22 = height2
                  PL22 = PL2
                  found2% = 1
                  Exit Do
                  
               Else
                  'interpolate and look for best fit
                  SLOPE2 = (height2 - height20) / (dist2 - dist20)
                  height1fit2 = (D1 - dist20) * SLOPE2 + height20
                  If Abs(H21 - height1fit2) < Difference Then
                     Difference = Abs(H21 - height1fit2)
                     VA22 = VA2
                     Dist22 = dist2
                     height22 = height2
                     PL22 = PL2
                     End If
                  End If
                End If

            dist20 = dist2
            height20 = height2
            VA20 = VA2
            Dip20 = Dip2
            DIF20 = DIF2
            PL20 = PL2
            End If
            
        npnt& = npnt& + 1
        
     Loop
     Close #file2%
     
     If found2% = 0 And Difference <= DistTolerance Then
        found2% = 1
     ElseIf found2% = 0 And Difference > DistTolerance Then
        Call MsgBox("Search at TLOW and HLOW was unsuccessful" _
                    & vbCrLf & sEmpty _
                    & vbCrLf & "Increase the Distance Tolerance!" _
                    , vbExclamation, "No result returned")
        
        End If
        
2500:

     If JustHHIGH Then
        VATLOW = VA21
     ElseIf JustHLOW Then
        VATLOW = VA22
     Else
     
        If found1% = 1 And found2% = 1 Then
           SLOPEfound = (VA21 - VA22) / (H1HIGH - H1LOW)
           VATLOW = (H11 - H1LOW) * SLOPEfound + VA22
           End If
           
        End If
      
3000:
     'now intterpolate the VA for the right temperature
     If JustTHIGH Then
        VAfinal = VATHIGH
     ElseIf JustTLOW Then
        VAfinal = VATLOW
     Else
        If VATHIGH <> 0 And VATLOW <> 0 Then
           SlopeVA = (VATHIGH - VATLOW) / (THIGH - TLOW)
           VAfinal = (T1 - TLOW) * SlopeVA + VATLOW
           End If
        End If
        
     If Dir(FileT1H1High$) <> sEmpty Then Kill FileT1H1High$
     If Dir(FileLow$) <> sEmpty Then Kill FileLow$
     
    LablOE = "View Angle with Ter Ref (old estimate) (deg.): " & Format(Str(Val(viewang / cd + avref)), "##0.0####")
     
         
    Screen.MousePointer = vbDefault


Return

VAsub:

'   RE = 6378136.6
    'use same latitudes, longitude difference due to length on circumference of earth -- approximate to distance between the two places
    'so lg1, lg1 + L/RE
   
'    RE = Rearth
    hgt1 = H11
    hgt2 = H21
    X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
    X2 = Cos(lt2 * cd) * Cos(lg2 * cd)
    Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
    Y2 = Cos(lt2 * cd) * Sin(lg2 * cd)
    z1 = Sin(lt1 * cd)
    z2 = Sin(lt2 * cd)
'    Rearth = 6371315#
    re1 = (hgt1 + RE)
    re2 = (hgt2 + RE)
    X1 = re1 * X1
    Y1 = re1 * Y1
    z1 = re1 * z1
    X2 = re2 * X2
    Y2 = re2 * Y2
    z2 = re2 * z2
    dist1 = re1
    dist2 = re2
    ANGLE = DACOS((X1 * X2 + Y1 * Y2 + z1 * z2) / (dist1 * dist2))
    viewang = Atn((-re1 + re2 * Cos(ANGLE)) / (re2 * Sin(ANGLE)))
    
'    /* Computing 2nd power */
    d__1 = X1 - X2
'    /* Computing 2nd power */
    d__2 = Y1 - Y2
'    /* Computing 2nd power */
    d__3 = z1 - z2
    distd = Sqr(d__1 * d__1 + d__2 * d__2 + d__3 * d__3) * 0.001 'convert to kms
'    re1 = hgt + re;
'    re2 = hgt2 + re;
    deltd = hgt1 - hgt2
'    x1 = re1 * x1;
'    y1 = re1 * y1;
'    z1 = re1 * z1;
'    x2 = re2 * x2;
'    y2 = re2 * y2;
'    z2 = re2 * z2;
'    dist1 = re1;
'    dist2 = re2;
'    angle = acos((x1 * x2 + y1 * y2 + z1 * z2) / (dist1 * dist2));
'/*          view angle in radians */
'    viewang = atan((-re1 + re2 * cos(angle)) / (re2 * sin(angle)));
'    d__ = (dist1 - dist2 * cos(angle)) / dist1;
'    x1d = x1 * (1 - d__) - x2;
'    y1d = y1 * (1 - d__) - y2;
'    z1d = z1 * (1 - d__) - z2;
'    x1p = -sin(-lg * cd);
'    y1p = cos(-lg * cd);
'    z1p = 0.;
'    azicos = x1p * x1d + y1p * y1d;
'    x1s = -cos(-lg * cd) * sin(lt * cd);
'    y1s = -sin(-lg * cd) * sin(lt * cd);
'    z1s = cos(lt * cd);
'    azisin = x1s * x1d + y1s * y1d + z1s * z1d;
'    azi = atan(azisin / azicos);
'/*      azimuth in degrees */
'    azi /= cd;
'/*      add contribution of atmospheric refraction to view angle */
    If (deltd <= 0#) Then
        defm = 0.000782 - deltd * 0.000000311
        defb = deltd * 0.000034 - 0.0141
    ElseIf (deltd > 0#) Then
        defm = deltd * 0.000000309 + 0.000764
        defb = -0.00915 - deltd * 0.0000269
        End If
    avref = defm * distd + defb
    If (avref < 0#) Then avref = 0#
    
Return

   On Error GoTo 0
   Exit Sub

cmdCalcTR_Click_Error:

    If err.Number = 53 Then Resume Next 'trying to kill a nonexistent file
    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdCalcTR_Click of Form prjAtmRefMainfm"

   On Error GoTo 0
   Exit Sub

End Sub

Private Sub cmdCiddor_Click()

  Dim H As Double
  Dim PDRY As Double
  Dim PVAP As Double
  Dim NumLayers As Long

    RELHUM = Val(txtRELHUM) 'relative humidity
    RELH = RELHUM / 100
    For i = 2 To 50
       If HL(i) = 0 Then
          NumLayers = i - 1
          Exit For
          End If
    Next i

    H = H * 1000#  'convert to meters

'    PDRY = fFNDPD1(H, PRESSD1, Dist, NumLayers) 'to get this to work need to reference this function globally in a module
'    PVAP = RELH * fVAPOR(H, Dist, NumLayers)
    
    txtCiddorDry.Text = " "
    txtCiddorDry.Text = PDRY
    txtCiddorWet.Text = " "
    txtCiddorWet.Text = PVAP
    
End Sub

Private Sub cmdDown_Click()
   Yorigin = Yorigin + prjAtmRefMainfm.height / 20
   Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
End Sub

Private Sub cmdHS_Click()
'      SUBROUTINE sla_REFRO ( ZOBS, HM, TDK, PMB, RH, WL, PHI, TLR,
'     :                       EPS, REF )
'*+
'*     - - - - - -
'*      R E F R O
'*     - - - - - -
'*
'*  Atmospheric refraction for radio and optical/IR wavelengths.
'*
'*  Given:
'*    ZOBS    d  observed zenith distance of the source (radian)
'*    HM      d  height of the observer above sea level (metre)
'*    TDK     d  ambient temperature at the observer (K)
'*    PMB     d  pressure at the observer (millibar)
'*    RH      d  relative humidity at the observer (range 0-1)
'*    WL      d  effective wavelength of the source (micrometre)
'*    PHI     d  latitude of the observer (radian, astronomical)
'*    TLR     d  temperature lapse rate in the troposphere (K/metre)
'*    EPS     d  precision required to terminate iteration (radian)
'*
'*  Returned:
'*    REF     d  refraction: in vacuo ZD minus observed ZD (radian)
'*
'*  Notes:
'*
'*  1  A suggested value for the TLR argument is 0.0065D0.  The
'*     refraction is significantly affected by TLR, and if studies
'*     of the local atmosphere have been carried out a better TLR
'*     value may be available.  The sign of the supplied TLR value
'*     is ignored.
'*
'*  2  A suggested value for the EPS argument is 1D-8.  The result is
'*     usually at least two orders of magnitude more computationally
'*     precise than the supplied EPS value.
'*
'*  3  The routine computes the refraction for zenith distances up
'*     to and a little beyond 90 deg using the method of Hohenkerk
'*     and Sinclair (NAO Technical Notes 59 and 63, subsequently adopted
'*     in the Explanatory Supplement, 1992 edition - see section 3.281).
'*
'*  4  The code is a development of the optical/IR refraction subroutine
'*     AREF of C.Hohenkerk (HMNAO, September 1984), with extensions to
'*     support the radio case.  Apart from merely cosmetic changes, the
'*     following modifications to the original HMNAO optical/IR refraction
'*     code have been made:
'*
'*     .  The angle arguments have been changed to radians.
'*
'*     .  Any value of ZOBS is allowed (see note 6, below).
'*
'*     .  Other argument values have been limited to safe values.
'*
'*     .  Murray's values for the gas constants have been used
'*        (Vectorial Astrometry, Adam Hilger, 1983).
'*
'*     .  The numerical integration phase has been rearranged for
'*        extra clarity.
'*
'*     .  A better model for Ps(T) has been adopted (taken from
'*        Gill, Atmosphere-Ocean Dynamics, Academic Press, 1982).
'*
'*     .  More accurate expressions for Pwo have been adopted
'*        (again from Gill 1982).
'*
'*     .  The formula for the water vapour pressure, given the
'*        saturation pressure and the relative humidity, is from
'*        Crane (1976), expression 2.5.5.
'*
'*     .  Provision for radio wavelengths has been added using
'*        expressions devised by A.T.Sinclair, RGO (private
'*        communication 1989).  The refractivity model currently
'*        used is from J.M.Rueger, "Refractive Index Formulae for
'*        Electronic Distance Measurement with Radio and Millimetre
'*        Waves", in Unisurv Report S-68 (2002), School of Surveying
'*        and Spatial Information Systems, University of New South
'*        Wales, Sydney, Australia.
'*
'*     .  The optical refractivity for dry air is from Resolution 3 of
'*        the International Association of Geodesy adopted at the XXIIth
'*        General Assembly in Birmingham, UK, 1999.
'*
'*     .  Various small changes have been made to gain speed.
'*
'*  5  The radio refraction is chosen by specifying WL > 100 micrometres.
'*     Because the algorithm takes no account of the ionosphere, the
'*     accuracy deteriorates at low frequencies, below about 30 MHz.
'*
'*  6  Before use, the value of ZOBS is expressed in the range +/- pi.
'*     If this ranged ZOBS is -ve, the result REF is computed from its
'*     absolute value before being made -ve to match.  In addition, if
'*     it has an absolute value greater than 93 deg, a fixed REF value
'*     equal to the result for ZOBS = 93 deg is returned, appropriately
'*     signed.
'*
'*  7  As in the original Hohenkerk and Sinclair algorithm, fixed values
'*     of the water vapour polytrope exponent, the height of the
'*     tropopause, and the height at which refraction is negligible are
'*     used.
'*
'*  8  The radio refraction has been tested against work done by
'*     Iain Coulson, JACH, (private communication 1995) for the
'*     James Clerk Maxwell Telescope, Mauna Kea.  For typical conditions,
'*     agreement at the 0.1 arcsec level is achieved for moderate ZD,
'*     worsening to perhaps 0.5-1.0 arcsec at ZD 80 deg.  At hot and
'*     humid sea-level sites the accuracy will not be as good.
'*
'*  9  It should be noted that the relative humidity RH is formally
'*     defined in terms of "mixing ratio" rather than pressures or
'*     densities as is often stated.  It is the mass of water per unit
'*     mass of dry air divided by that for saturated air at the same
'*     temperature and pressure (see Gill 1982).
'*
'*  10 The algorithm is designed for observers in the troposphere.  The
'*     supplied temperature, pressure and lapse rate are assumed to be
'*     for a point in the troposphere and are used to define a model
'*     atmosphere with the tropopause at 11km altitude and a constant
'*     temperature above that.  However, in practice, the refraction
'*     values returned for stratospheric observers, at altitudes up to
'*     25km, are quite usable.
'*
'*  Called:  sla_DRANGE, sla__ATMT, sla__ATMS
'*
'*  Last revision:   5 December 2005
'*
'*  Copyright P.T.Wallace.  All rights reserved.
'*
'*  License:
'*    This program is free software; you can redistribute it and/or modify
'*    it under the terms of the GNU General Public License as published by
'*    the Free Software Foundation; either version 2 of the License, or
'*    (at your option) any later version.
'*
'*    This program is distributed in the hope that it will be useful,
'*    but WITHOUT ANY WARRANTY; without even the implied warranty of
'*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'*    GNU General Public License for more details.
'*
'*    You should have received a copy of the GNU General Public License
'*    along with this program (see SLA_CONDITIONS); if not, write to the
'*    Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'*    Boston, MA  02111-1307  USA
'*
'*-
'      IMPLICIT NONE
      Dim ZOBS As Double, HM As Double, TDK As Double, PMB As Double, RH As Double
      Dim wl As Double, Phi As Double, TLR As Double, EPS As Double, ref As Double
'*
'*  Fixed parameters
'*
      Dim D93 As Double, gcr As Double, DMD As Double, DMW As Double, s As Double
      Dim DELTA As Double, ht As Double, hs As Double
      Dim ISMAX As Integer
'*  93 degrees in radians
      D93 = 1.623156204
'*  Universal gas constant
      gcr = 8314.32
'*  Molecular weight of dry air
      DMD = 28.9644
'*  Molecular weight of water vapour
      DMW = 18.0152
'*  Mean Earth radius (metre)
      s = 6378120#
'*  Exponent of temperature dependence of water vapour pressure
      DELTA = 18.36
'*  Height of tropopause (metre)
      ht = 11000#
'*  Upper limit for refractive effects (metre)
      hs = 80000#
      
      pi = 4# * Atn(1#) '3.141592654
      cd = pi / 180# 'conversion from degrees to radians

'*  Numerical integration: maximum number of strips.
      ISMAX = 16384
      Dim iss As Integer, k As Integer, n As Integer, i As Integer, j As Integer
      Dim OPTIC As Boolean, LOOPS As Boolean
      Dim ZOBS1 As Double, ZOBS2 As Double, HMOK As Double, TDKOK As Double, PMBOK As Double, RHOK As Double, WLOK As Double, ALPHA As Double
      Dim Tol As Double, wlsq As Double, gb As Double, A As Double, GAMAL As Double, gamma As Double, GAMM2 As Double, DELM2 As Double
      Dim TDC As Double, psat As Double, PWO As Double, W As Double
      Dim C1 As Double, C2 As Double, C3 As Double, C4 As Double, C5 As Double, C6 As Double, r0 As Double, TEMPO As Double, DN0 As Double, RDNDR0 As Double, sk0 As Double, f0 As Double
      Dim rt As Double, TT As Double, DNT As Double, RDNDRT As Double, SINE As Double, zt As Double, ft As Double, DNTS As Double, RDNDRP As Double, zts As Double, fts As Double
      Dim rs As Double, DNS As Double, RDNDRS As Double, zs As Double, FS As Double, REFOLD As Double, z0 As Double, ZRANGE As Double, fb As Double, ff As Double, fo As Double, fe As Double
      Dim H As Double, r As Double, SZ As Double, rg As Double, DR As Double, tg As Double, DN As Double, RDNDR As Double, T As Double, F As Double, refp As Double, reft As Double

      ZOBS = 90# * cd 'zenith angle in radians, use straight angle for now
      HM = HOBS
      TDK = prjAtmRefMainfm.txtGroundTemp
      PMB = prjAtmRefMainfm.txtGroundPressure
      RH = prjAtmRefMainfm.txtHumid / 100
      wl = 0.574 'micrometers - middle of starlight wavelength range
      Phi = 32.1 'latitude
      TLR = 0.0065
      EPS = 0.00000001

'      Dim sla_DRANGE  As Double
'*  The refraction integrand
'      Dim REFI As Double
'      REFI(DN, RDNDR) = RDNDR / (DN + RDNDR)
'*  Transform ZOBS into the normal range.
      ZOBS1 = sla_DRANGE(ZOBS)
      ZOBS2 = Minimum(Abs(ZOBS1), D93)
'*  Keep other arguments within safe bounds.
      HMOK = Minimum(Maximum(HM, -1000#), hs)
      TDKOK = Minimum(Maximum(TDK, 100#), 500#)
      PMBOK = Minimum(Maximum(PMB, 0#), 10000#)
      RHOK = Minimum(Maximum(RH, 0#), 1#)
      WLOK = Maximum(wl, 0.1)
      ALPHA = Minimum(Maximum(Abs(TLR), 0.001), 0.01)
'*  Tolerance for iteration.
      Tol = Minimum(Maximum(Abs(EPS), 0.000000000001), 0.1) / 2#
'*  Decide whether optical/IR or radio case - switch at 100 microns.
      If WLOK <= 100# Then
         OPTIC = True
         End If
'      OPTIC = WLOK.LE.100D0
'*  Set up model atmosphere parameters defined at the observer.
      wlsq = WLOK * WLOK
      gb = 9.784 * (1# - 0.0026 * Cos((Phi + Phi) * cd) - 0.00000028 * HMOK)
      If (OPTIC) Then
         A = (287.6155 + (1.62887 + 0.0136 / wlsq) / wlsq) * 0.00027315 / 1013.25
      Else
         A = 0.000077689
      End If
      GAMAL = (gb * DMD) / gcr
      gamma = GAMAL / ALPHA
      GAMM2 = gamma - 2#
      DELM2 = DELTA - 2#
      TDC = TDKOK - 273.15
      psat = 10# ^ ((0.7859 + 0.03477 * TDC) / (1# + 0.00412 * TDC)) * (1# + PMBOK * (0.0000045 + 0.0000000006 * TDC * TDC))
      If (PMBOK > 0#) Then
         PWO = RHOK * psat / (1# - (1# - RHOK) * psat / PMBOK)
      Else
         PWO = 0#
      End If
      W = PWO * (1# - DMW / DMD) * gamma / (DELTA - gamma)
      C1 = A * (PMBOK + W) / TDKOK
      If (OPTIC) Then
         C2 = (A * W + 0.0000112684 * PWO) / TDKOK
      Else
         C2 = (A * W + 0.0000063938 * PWO) / TDKOK
      End If
      C3 = (gamma - 1#) * ALPHA * C1 / TDKOK
      C4 = (DELTA - 1#) * ALPHA * C2 / TDKOK
      If (OPTIC) Then
         C5 = 0#
         C6 = 0#
      Else
         C5 = 0.375463 * PWO / TDKOK
         C6 = C5 * DELM2 * ALPHA / (TDKOK * TDKOK)
      End If
'*  Conditions at the observer.
      r0 = s + HMOK
      Call sla__ATMT(r0, TDKOK, ALPHA, GAMM2, DELM2, C1, C2, C3, C4, C5, C6, r0, TEMPO, DN0, RDNDR0)
      sk0 = DN0 * r0 * Sin(ZOBS2)
      f0 = refi(DN0, RDNDR0)
'*  Conditions in the troposphere at the tropopause.
      rt = s + Maximum(ht, HMOK)
      Call sla__ATMT(r0, TDKOK, ALPHA, GAMM2, DELM2, C1, C2, C3, C4, C5, C6, rt, TT, DNT, RDNDRT)
      SINE = sk0 / (rt * DNT)
      zt = Minimum(2 * Atn(SINE) / cd, 90#)
      If zt > 90# Then zt = 90#
'      zt = Atan2(SINE, Sqr(Maximum(1# - SINE * SINE, 0#))) / cd
      ft = refi(DNT, RDNDRT)
'*  Conditions in the stratosphere at the tropopause.
      Call sla__ATMS(rt, TT, DNT, GAMAL, rt, DNTS, RDNDRP)
      SINE = sk0 / (rt * DNTS)
      zts = Minimum(2# * Atn(SINE) / cd, 90#)
'      zts = Atan2(SINE, Sqr(Maximum(1# - SINE * SINE, 0#))) / cd
      fts = refi(DNTS, RDNDRP)
'*  Conditions at the stratosphere limit.
      rs = s + hs
      Call sla__ATMS(rt, TT, DNT, GAMAL, rs, DNS, RDNDRS)
      SINE = sk0 / (rs * DNS)
      zs = Minimum(2# * Atn(SINE) / cd, 90#)
'      zs = Atan2(SINE, Sqr(Maximum(1# - SINE * SINE, 0#))) / cd
      FS = refi(DNS, RDNDRS)
'*  Variable initialization to avoid compiler warning.
      reft = 0#
'*  Integrate the refraction integral in two parts;  first in the
'*  troposphere (K=1), then in the stratosphere (K=2).
      For k = 1 To 2
'*     Initialize previous refraction to ensure at least two iterations.
         REFOLD = 1#
'*     Start off with 8 strips.
         iss = 8
'*     Start Z, Z range, and start and end values.
         If (k = 1) Then
            z0 = ZOBS2 / cd
            ZRANGE = zt - z0
            fb = f0
            ff = ft
         Else
            z0 = zts
            ZRANGE = zs - z0
            fb = fts
            ff = FS
         End If
'*     Sums of odd and even values.
         fo = 0#
         fe = 0#
'*     First time through the loop we have to do every point.
         n = 1
'*     Start of iteration loop (terminates at specified precision).
         LOOPS = True
         Do While LOOPS = True
'*        Strip width.
            H = ZRANGE / CDbl(iss)
'*        Initialize distance from Earth centre for quadrature pass.
            If (k = 1) Then
               r = r0
            Else
               r = rt
            End If
'*        One pass (no need to compute evens after first time).
            For i = 1 To iss - 1 Step n
'*           Sine of observed zenith distance.
               SZ = Sin(z0 * cd + H * CDbl(i))
'*           Find R (to the nearest metre, maximum four iterations).
               If (SZ > 1E-20) Then
                  W = sk0 / SZ
                  rg = r
                  DR = 1000000#
                  j = 0
                  Do While (Abs(DR) > 1# And j < 4)
                     j = j + 1
                     If (k = 1) Then
                        Call sla__ATMT(r0, TDKOK, ALPHA, GAMM2, DELM2, C1, C2, C3, C4, C5, C6, rg, tg, DN, RDNDR)
                     Else
                        Call sla__ATMS(rt, TT, DNT, GAMAL, rg, DN, RDNDR)
                     End If
                     DR = (rg * DN - W) / (DN + RDNDR)
                     rg = rg - DR
                  Loop
                  r = rg
               End If
'*           Find the refractive index and integrand at R.
               If (k = 1) Then
                  Call sla__ATMT(r0, TDKOK, ALPHA, GAMM2, DELM2, C1, C2, C3, C4, C5, C6, r, T, DN, RDNDR)
               Else
                  Call sla__ATMS(rt, TT, DNT, GAMAL, r, DN, RDNDR)
               End If
               F = refi(DN, RDNDR)
'*           Accumulate odd and (first time only) even values.
               If (n = 1# And i Mod 2 = 0) Then
                  fe = fe + F
               Else
                  fo = fo + F
               End If
            Next i
'*        Evaluate the integrand using Simpson's Rule.
            refp = H * (fb + 4# * fo + 2# * fe + ff) / 3# 'rad

'*        Has the required precision been achieved (or can't be)?
            If (Abs(refp - REFOLD) > Tol And iss < ISMAX) Then
'*           No: prepare for next iteration.
'*           Save current value for convergence test.
               REFOLD = refp
'*           Double the number of strips.
               iss = iss + iss
'*           Sum of all current values = sum of next pass's even values.
               fe = fe + fo

'*           Prepare for new odd values.
               fo = 0#

'*           Skip even values next time.
               n = 2
            Else

'*           Yes: save troposphere component and terminate the loop.
               If (k = 1) Then reft = refp
               LOOPS = False
            End If
         Loop
      Next k

'*  Result.
      ref = reft + refp
      If (ZOBS1 < 0) Then ref = -ref

End
End Sub
Function refi(DN As Double, RDNDR As Double) As Double
   refi = RDNDR / (DN + RDNDR)
End Function
Function sla_DRANGE(ANGLE As Double) As Double

'*+
'*     - - - - - -
'*      R A N G E
'*     - - - - - -
'*
'*  Normalize angle into range +/- pi  (single precision)
'*
'*  Given:
'*     ANGLE     dp      the angle in radians
'*
'*  The result is ANGLE expressed in the +/- pi (single
'*  precision).
'*
'*  P.T.Wallace   Starlink   23 November 1995
'*
'*  Copyright (C) 1995 Rutherford Appleton Laboratory
'*
'*  License:
'*    This program is free software; you can redistribute it and/or modify
'*    it under the terms of the GNU General Public License as published by
'*    the Free Software Foundation; either version 2 of the License, or
'*    (at your option) any later version.
'*
'*    This program is distributed in the hope that it will be useful,
'*    but WITHOUT ANY WARRANTY; without even the implied warranty of
'*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'*    GNU General Public License for more details.
'*
'*    You should have received a copy of the GNU General Public License
'*    along with this program (see SLA_CONDITIONS); if not, write to the
'*    Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'*    Boston, MA  02111-1307  USA
'*
'*-

'      IMPLICIT NONE

'      Real ANGLE

      Dim API As Double, A2PI As Double
      API = 3.14159265358979
      A2PI = 6.28318530717959


      sla_DRANGE = ANGLE - (Int(ANGLE / A2PI) * A2PI)
      If (Abs(sla_DRANGE) >= API) Then sla_DRANGE = sla_DRANGE - A2PI * Sgn(ANGLE)

End Function
      Sub sla__ATMT(r0 As Double, t0 As Double, ALPHA As Double, GAMM2 As Double, _
      DELM2 As Double, C1 As Double, C2 As Double, C3 As Double, C4 As Double, _
      C5 As Double, C6 As Double, r As Double, T As Double, DN As Double, RDNDR As Double)
'*+
'*     - - - - -
'*      A T M T
'*     - - - - -
'*
'*  Internal routine used by REFRO
'*
'*  Refractive index and derivative with respect to height for the
'*  troposphere.
'*
'*  Given:
'*    R0      d    height of observer from centre of the Earth (metre)
'*    T0      d    temperature at the observer (K)
'*    ALPHA   d    alpha          )
'*    GAMM2   d    gamma minus 2  ) see HMNAO paper
'*    DELM2   d    delta minus 2  )
'*    C1      d    useful term  )
'*    C2      d    useful term  )
'*    C3      d    useful term  ) see source
'*    C4      d    useful term  ) of sla_REFRO
'*    C5      d    useful term  )
'*    C6      d    useful term  )
'*    R       d    current distance from the centre of the Earth (metre)
'*
'*  Returned:
'*    T       d    temperature at R (K)
'*    DN      d    refractive index at R
'*    RDNDR   d    R * rate the refractive index is changing at R
'*
'*  Note that in the optical case C5 and C6 are zero.
'*
'*  Last revision:   26 December 2004
'*
'*  Copyright P.T.Wallace.  All rights reserved.
'*
'*  License:
'*    This program is free software; you can redistribute it and/or modify
'*    it under the terms of the GNU General Public License as published by
'*    the Free Software Foundation; either version 2 of the License, or
'*    (at your option) any later version.
'*
'*    This program is distributed in the hope that it will be useful,
'*    but WITHOUT ANY WARRANTY; without even the implied warranty of
'*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'*    GNU General Public License for more details.
'*
'*    You should have received a copy of the GNU General Public License
'*    along with this program (see SLA_CONDITIONS); if not, write to the
'*    Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'*    Boston, MA  02111-1307  USA
'*
'*-
'
'      IMPLICIT NONE

'      dim R0,T0,ALPHA,GAMM2,DELM2,C1,C2,C3,C4,C5,C6,
':                      R , T, DN, RDNDR

      Dim tt0 As Double, TT0GM2 As Double, TT0DM2 As Double


      T = Maximum(Minimum(t0 - ALPHA * (r - r0), 320#), 100#)
      tt0 = T / t0
      TT0GM2 = tt0 ^ GAMM2
      TT0DM2 = tt0 ^ DELM2
      DN = 1# + (C1 * TT0GM2 - (C2 - C5 / T) * TT0DM2) * tt0
      RDNDR = r * (-C3 * TT0GM2 + (C4 - C6 / tt0) * TT0DM2)

End Sub
      Sub sla__ATMS(rt As Double, TT As Double, DNT As Double, GAMAL As Double, r As Double, DN As Double, RDNDR As Double)

'*+
'*     - - - - -
'*      A T M S
'*     - - - - -
'*
'*  Internal routine used by REFRO
'*
'*  Refractive index and derivative with respect to height for the
'*  stratosphere.
'*
'*  Given:
'*    RT      d    height of tropopause from centre of the Earth (metre)
'*    TT      d    temperature at the tropopause (K)
'*    DNT     d    refractive index at the tropopause
'*    GAMAL   d    constant of the atmospheric model = G*MD/R
'*    R       d    current distance from the centre of the Earth (metre)
'*
'*  Returned:
'*    DN      d    refractive index at R
'*    RDNDR   d    R * rate the refractive index is changing at R
'*
'*  Last revision:   26 December 2004
'*
'*  Copyright P.T.Wallace.  All rights reserved.
'*
'*  License:
'*    This program is free software; you can redistribute it and/or modify
'*    it under the terms of the GNU General Public License as published by
'*    the Free Software Foundation; either version 2 of the License, or
'*    (at your option) any later version.
'*
'*    This program is distributed in the hope that it will be useful,
'*    but WITHOUT ANY WARRANTY; without even the implied warranty of
'*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'*    GNU General Public License for more details.
'*
'*    You should have received a copy of the GNU General Public License
'*    along with this program (see SLA_CONDITIONS); if not, write to the
'*    Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'*    Boston, MA  02111-1307  USA
'*
'*-
'
'      IMPLICIT NONE

'      DOUBLE PRECISION RT,TT,DNT,GAMAL,R,DN,RDNDR

      Dim b As Double, W As Double


   On Error GoTo sla__ATMS_Error

      b = GAMAL / TT
      W = (DNT - 1#) * Exp(-b * (r - rt))
      DN = 1# + W
      RDNDR = -r * b * W

   On Error GoTo 0
   Exit Sub

sla__ATMS_Error:
    If err.Number = 6 Then
       W = 0
       Resume Next
       End If
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure sla__ATMS of Form prjAtmRefMainfm"

End Sub

Private Sub cmdLarger_Click()
   Dim StatusMes As String
   RefZoom.LastZoom = Mult
   Mult = Mult * 2#
   RefZoom.Zoom = Mult
   StatusMes = "Multiplication = " & Format(Mult, "#############0.0#")
   Call StatusMessage(StatusMes, 1, 0)
   txtStartMult = Mult
   'determine canvas size, i.e., size of picture2 and then values of scroll bars
   
   Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
End Sub

Private Sub cmdleft_Click()
   Xorigin = Xorigin + prjAtmRefMainfm.Width / 20
   Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdMenat_Click
' Author    : Dr-John-K-Hall
' Date      : 2/18/2019
' Purpose   : VB6 enactment of Menat ray tracing code, source: menat.cpp
'---------------------------------------------------------------------------------------
'
Private Sub cmdMenat_Click()
'/* menat.f -- translated by f2c (version 20021022).
'   You must link the resulting object file with the libraries:
'    -lf2c -lm   (in that order)
'*/
'
'//#include "f2c.h"
'#include <stdio.h>
'#include <math.h>
'#include <stdio.h>
'#include <stdlib.h>
'#include <ctype.h>
'#include <string.h>
'
'/* Common Block Declarations */
'
'struct {
'Dim hj(50) As Double, tj(50) As Double, pj(50) As Double, at(50) As Double, ct(50) As Double

Dim zz_1 As zz

'} zz_
'
'#define zz_1 zz_
'
'Dim layerheights(50) As Double
'Dim temps(50) As Double
'Dim press(50) As Double
Dim NumLayers As Integer
'
'//declare extermanl file
'
'/* Table of constant values */
'
'/*       PROGRAM MENAT--CALCULATES ASTRONOMICAL ATMOSPHERIC REFRACTION */
'/*       FOR ANY OBSERVER HEIGHT by using ray tracing through a */
'/*       simplified layered atmosphered */
'/*       ITS OUTPUT IS REPRESENTED in the files menatsum.ren, menatwin.ren */
'/* Main program */ //int MAIN__(void)
'int main(int argc, char* argv[])
Dim pi2 As Double, fr As Double, hz As Double, dh As Double, hsof As Double, epg As Double
Dim co As Double, pie As Double, ramg As Double, q3 As Double, q6 As Double, ra As Double
Dim b As Double, d__ As Double, e As Double, g As Double, h__ As Double, jstop As Long, j As Long
Dim k As Long, l As Long, n As Long, KWAV As Double, KMIN As Double, KMAX As Double, KSTEP As Double
Dim s As Double, T As Double, a3 As Double, e1 As Double, e2 As Double, g1 As Double, d6 As Double, S1 As Double, s2 As Double
'    //char ch[1]
Dim dg As Double, bn As Double, el As Double, en As Double, bz As Double, cz As Double, em As Double, StatusMes As String

'    //int nn
Dim rt As Double
Dim XP  As Double
'   On Error GoTo cmdMenat_Click_Error

XP = 0#
Dim nz As Long
Dim ru As Double
'    //char ch1[1], ch2[2], ch3[3]
Dim ep1 As Double, at2 As Double, en1 As Double, dt1 As Double
'    //double en2
'    //double en3//
'    //double en4,
Dim hz1 As Double, hz2 As Double
'    //logical beg
Dim den As Double
'    //logical neg
Dim hen As Double, dtg As Double, hev As Double
Dim kgr As Integer
Dim dhz As Double
'    //double ent[8000]  /* was [4][2000] */
Dim sbn As Double
'    extern double fun_(double *)
Dim entry As Double
Dim isn As Integer
isn = 1
Dim iam As Integer
iam = 0
Dim epz As Double
'    extern double sqt_(double *, double *, double *)
'    extern int LoadAtmospheres(char filnam[])
Dim hen1 As Double, epg1 As Double, epg2 As Double, dtg2 As Double, fieg As Double
'    //logical angl
'    //int mang,
Dim nang As Integer
'    //int ccc
Dim fiem As Double, hmin As Double, hmax As Double
Dim nent As Integer, nhgt As Integer '//,mhgt
Dim rsof As Double, epzm As Double
'    //char finam[10]
'    //double chisn
Dim nkhgt As Integer
Dim estep As Double
Dim ANGLE As Double
ANGLE = 0#
Dim A1 As Double
A1 = 0#
Dim A2 As Double
A2 = 0#
Dim START As Boolean
START = False
'    char filnam[255] = sempty
Dim filnam As String
'    char chr[2] = sempty
Dim ier As Integer
ier = 0
'    int ier = 0

Dim FNM As String, AtmType As Integer, AtmNumber As Integer, lpsrate As Double, tst As Double, pst As Double, NNN As Long

   cmdCalc.Enabled = False
   cmdRefWilson.Enabled = False
   cmdMenat.Enabled = False
   cmdVDW.Enabled = False
   
   RefCalcType% = 2
   CalcComplete = False

STARTALT = Val(txtStartAlt.Text)
DELALT = Val(txtDelAlt.Text)
XMAX = Val(txtXmax.Text) * 1000 'convert km to meters
PPAM = Val(txtPPAM.Text)
KMIN = CInt((Val(txtKmin.Text) - 380) / 5# + 1#)
KMAX = CInt((Val(txtKmax.Text) - 380) / 5# + 1#)
KSTEP = CInt(Val(txtKStep.Text) * 0.1)
STARTAZM = 19
DELAZM = 32
If INVFLAG = 1 Then
   SINV = Val(txtSInv.Text)
   EINV = Val(txtEInv.Text)
   DTINV = Val(txtDInv.Text)
   End If
   
StatusMes = "Pixels per arcminute " & Str(PPAM) & ", Maximum height (degrees) " & Str(n / (120# * PPAM))
Call StatusMessage(StatusMes, 1, 0)

n_size = 500
msize = 20 + Val(txtNumSuns.Text) * 32 * PPAM

If Trim$(txtXSize.Text) <> sEmpty Then
   msize = Val(txtXSize.Text)
   End If
If Trim$(txtYSize.Text) <> sEmpty Then
   n_size = 2 * Val(txtYSize.Text) * PPAM * 60
   End If
   
Dim KA As Long
For KA = 1 To NumSuns
   ALT(KA) = STARTALT + CDbl(KA - 1) * DELALT
   AZM(KA) = STARTAZM + CDbl(KA - 1) * DELAZM
Next KA

myfile$ = Dir(App.Path & "\test_M.dat")
If myfile$ <> sEmpty Then
   Kill App.Path & "\test_M.dat"
   End If
myfile$ = Dir(App.Path & "\tc_M.dat")
If myfile$ <> sEmpty Then
   Kill App.Path & "\tc_M.dat"
   End If


pi = 4# * Atn(1#) '3.141592654
CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
cd = pi / 180# 'conversion of degrees to radians
ROBJ = 15# 'half size of sun in minutes of arc
'{
'    /* Initialized data */
'
 pi2 = pi * 0.5 '1.570796
 co = cd '0.017453293
 fr = 1.001
 hz = 0.801
 dh = 0.002
 hsof = 65#
 epg = 0#
' hs[7] = { 0.,13.,18.,25.,47.,50.,70. }
 pie = pi '3.1415926
 ramg = 0.05729578
 q3 = 1000#
 q6 = 1000000#
'    /*
'    double hw[7] = { 0.,10.,19.,25.,30.,50.,70. }
'    double ts[7] = { 299.,215.5,216.5,224.,273.,276.,218. }
'    double tw[7] = { 284.,220.,215.,216.,217.,266.,231. }
'    double ps[7] = { 1013.,179.,81.2,27.7,1.29,.951,.067 }
'    double pw[7] = { 1018.,256.8,62.8,24.3,11.1,.682,.0467 }
'    */

ra = 6371 ' 6378.1366     '6371.
RE = ra * 1000#
'
'    /* System generated locals */
'    int i__1, i__2, i__3//, i__4
'    double d__1//, d__2
'
'    /* Local variables */
'    double b, d__, e, g, h__
'    int k, l, n
'    double s, t, a2, a3, e1, e2, g1, d6, s1, s2
'    //char ch[1]
'    double dg, bn, el, en, bz, cz, em
'    //int nn
'    double rt
'    double XP = 0.0
'    int nz
'    double ru
'    //char ch1[1], ch2[2], ch3[3]
'    double ep1, at2, en1, dt1
'    //double en2
'    //double en3//
'    //double en4,
'    double  hz1, hz2
'    //logical beg
'    double den
'    //logical neg
'    double hen, dtg, hev
'    int kgr
'    double dhz
'    //double ent[8000]  /* was [4][2000] */
'    double sbn
'    extern double fun_(double *)
'    double entry
'    int isn = 1
'    int iam = 0
'    double epz
'    extern double sqt_(double *, double *, double *)
'    extern int LoadAtmospheres(char filnam[])
'    double hen1, epg1, epg2, dtg2, fieg
'    //logical angl
'    //int mang,
'    int nang
'    //int ccc
'    double fiem, hmin, hmax
'    int nent,nhgt//,mhgt
'    double rsof, epzm
'    //char finam[10]
'    //double chisn
'    int nkhgt
'    double estep
'    double ANGLE = 0.0
'    double A1 = 0.0
'    double A2 = 0.0
'    bool START = false
'    char filnam[255] = sempty
'    char chr[2] = sempty
'    int ier = 0
'
'
'    FILE *stream
'
'L0:
    NumLayers = 0
    
     '------------------progress bar initialization
    With prjAtmRefMainfm
      '------fancy progress bar settings---------
      .progressfrm.Visible = True
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
   
     'specify atmosphere type and the file containing the atmosphere profile
      StatusMes = "Calculating and Storing multilayer atmospheric details"
      Call StatusMessage(StatusMes, 1, 0)
     
     If OptionLayer.Value = True Then
        AtmType = 1
        FNM = App.Path & "\stmod1.dat"
     ElseIf OptionRead.Value = True Then
        AtmType = 1
        FNM = TextExternal.Text
     ElseIf OptionSelby.Value = True Then
        AtmType = 2
        If prjAtmRefMainfm.opt1.Value = True Then
           AtmNumber = 1
        ElseIf prjAtmRefMainfm.opt2.Value = True Then
           AtmNumber = 2
        ElseIf prjAtmRefMainfm.opt3.Value = True Then
           AtmNumber = 3
        ElseIf prjAtmRefMainfm.opt4.Value = True Then
           AtmNumber = 4
        ElseIf prjAtmRefMainfm.opt5.Value = True Then
           AtmNumber = 5
        ElseIf prjAtmRefMainfm.opt6.Value = True Then
           AtmNumber = 6
        ElseIf prjAtmRefMainfm.opt7.Value = True Then
           AtmNumber = 7
        ElseIf prjAtmRefMainfm.opt8.Value = True Then
           AtmNumber = 8
        ElseIf prjAtmRefMainfm.opt9.Value = True Then
           AtmNumber = 9
        ElseIf prjAtmRefMainfm.opt10.Value = True Then
           AtmNumber = 10
           FNM = txtOther.Text
           End If
        End If

'     FNM = App.Path & "\stmod1.dat" 'stmod2.dat.txt"
     ier = LoadAtmospheres(FNM, AtmType, AtmNumber, lpsrate, tst, pst, NNN, 4)
     NumLayers = NNN + 1
     NumTemp = NumLayers
     
     If ier < 0 Then
        Screen.MousePointer = vbDefault
        Close
        cmdVDW.Enabled = True
        cmdCalc.Enabled = True
        cmdMenat.Enabled = True
        cmdRefWilson.Enabled = True
        Exit Sub
        End If
        
     '///////////////////////////////////////////////////////////////
'
'    //angl = TRUE_
'    printf(" SUMMER=1 (DEF), WINTER=2 --> ")
'    scanf("%lg", &entry)
'    isn = (int)entry
'    if (isn <> 1 and isn <>2) isn = 1
'
'    printf(" Choose Atmosphere: Menat-(1) Selby: Tropical-(2) Mid latitude-(3) subartic-(4) US Standard-(5) -->")
'    scanf("%lg", &entry)
'    iam = (int)entry
'    iam = iam - 1
'    if (iam < 0) iam = 0
'
'    //load the atmosphere files
'    if (isn == 1) {
'
'        switch (iam) {
'
'            Case 0:
'                strcpy(filnam, "Menat-EY-summer.txt")
'                break
'
'            Case 1:
'                strcpy(filnam, "Selby-tropical.txt")
'                break
'
'            Case 2:
'                strcpy(filnam, "Selby-midlatitude-summer.txt")
'                break
'
'            Case 3:
'                strcpy(filnam, "Selby-subartic-summer.txt")
'                break
'
'            Case 4:
'                strcpy(filnam, "Selby-US-standard.txt")
'                break
'
'default:
'                strcpy(filnam, "Menat-EY-summer.txt")
'                break
'        }
'    }
'
'    else if (isn == 2) {
'
'        switch (iam) {
'
'            Case 0:
'                strcpy(filnam, "Menat-EY-winter.txt" )
'                break
'
'            Case 1:
'                strcpy(filnam, "Selby-tropical.txt")
'                break
'
'            Case 2:
'                strcpy(filnam, "Selby-midlatitude-winter.txt")
'                break
'
'            Case 3:
'                strcpy(filnam, "Selby-subartic-winter.txt")
'                break
'
'            Case 4:
'                strcpy(filnam, "Selby-US-standard.txt")
'                break
'
'default:
'                strcpy(filnam, "Menat-EY-winter.txt" )
'                break
'        }
'    }
'
'    ier = LoadAtmospheres(filnam)
'    if (ier == -1) return -1
'
'    hsof = layerheights[numlayers - 1]
     hsof = ELV(NNN)
'
'    //output file
     fileout% = FreeFile
     Open App.Path & "\test_M.dat" For Output As #fileout%
'    if ( !( stream = fopen( "MENAT.OUT", "w")) )
'    {
'        return -1
'    }
'
'    printf(" INPUT BEGINNING HEIGHT FOR CALCULATION (M)--> ")
'    scanf("%lg", &hz1)
'
    hz1 = prjAtmRefMainfm.txtHeight
    hz1 = hz1 / 1000# 'beginning observer height to kms
'
'    printf(" INPUT END HEIGHT FOR CALCULATION (M)--> ")
'    scanf("%lg", &hz2)
    hz2 = hz2 / 1000# 'end observer height in kms
'
'    printf(" INPUT STEP HEIGHT FOR CALCULATION (M)--> ")
'    scanf("%lg", &dhz)
    dhz = 100 'meters
    dhz = dhz / 1000# 'stepsize in observer height in kms
    d__1 = (hz2 - hz1) / dhz
    nhgt = CInt(d__1) + 1
'
'/*        WRITE (*,'(A,F6.1)')' PRESENT HEIGHT FOR CALCULATION =',HZ*1000 */
'/*       WRITE (*,'(A\)')' WANT NEW HEIGHT ? (Y/N)' */
'/*        READ(*,'(A)')CH */
'/*        IF ((CH.EQ.'Y').OR.(CH.EQ.'y')) THEN */
'/* 2               WRITE(*,'(A\)')' INPUT NEW HEIGHT (M)-->' */
'/*                READ(*,*,ERR=2)HZ */
'/*                HZ=HZ/1.0D3 */
'/*                END IF */
'/*       DO 550 ISN=1,1 */
'/*       WRITE(*,5)ISN */
'/* 5       FORMAT('2',5X,'SUMMER=1 @ WINTER=2','   ISN=',I1//) */
'
'    //printf(" SUMMER=1 (DEF), WINTER=2 --> ")
'    //scanf("%d", &isn)
'    //if (isn <> 1 and isn <>2) isn = 1
'
'    /*
    For k = 1 To NumLayers
        zz_1.hj(k - 1) = ELV(k - 1) ' //hs(k - 1)
        zz_1.tj(k - 1) = TMP(k - 1) ' //ts(k - 1)
        zz_1.pj(k - 1) = PRSR(k - 1) ' //ps(k - 1)
    Next k
'    /*
'    if (isn == 1) {
'        goto L7
'    }
'    zz_1.hj(k - 1) = hw(k - 1)
'    zz_1.tj(k - 1) = tw(k - 1)
'    zz_1.pj(k - 1) = pw(k - 1)
'    */
'//L7:
'    /*
'
'    }
    For k = 1 To NumLayers
        l = k + 1
        If (k < NumLayers) Then
            zz_1.AT(k - 1) = (zz_1.tj(l - 1) - zz_1.tj(k - 1)) / (zz_1.hj(l - _
                1) - zz_1.hj(k - 1))
            End If
        If (k < NumLayers) Then
            If (zz_1.tj(l - 1) <> zz_1.tj(k - 1)) Then '//non-isothermic region
            
                zz_1.ct(k - 1) = Log(zz_1.pj(l - 1) / zz_1.pj(k - 1)) / Log( _
                    zz_1.tj(l - 1) / zz_1.tj(k - 1))
            
            Else '//isothermic region  -- interpolate between the two layer's pressures
            
                zz_1.ct(k - 1) = (zz_1.pj(l - 1) - zz_1.pj(k - 1)) / (zz_1.hj(l - _
                1) - zz_1.hj(k - 1))
            
                End If

'    printf("%i3, %f6, %f9, %f7, AT=%f7, CT=%f9\n", k, zz_1.hj(k - 1), zz_1.pj(k - 1), zz_1.tj(k - 1), zz_1.at(k - 1), zz_1.ct(k - 1))
          End If
     Next k
'    */
'
'    printf(" INPUT MINIMUM ANG ALT (DEG)--> ")
'    scanf("%lg", &epg1)
'
'
'    printf(" INPUT MAXIMUM ANG ALT (DEG)--> ")
'    scanf("%lg", &epg2)
'
'
'    printf(" INPUT STEP ANG ALT (DEG)--> ")
'    scanf("%lg", &estep)
'
'
    n_size = 500
    msize = 20 + Val(txtNumSuns.Text) * 32 * PPAM
    
    If Trim$(txtXSize.Text) <> sEmpty Then
       msize = Val(txtXSize.Text)
       End If
    If Trim$(txtYSize.Text) <> sEmpty Then
       n_size = 2 * Val(txtYSize.Text) * PPAM * 60
       End If
       
   PPAM = prjAtmRefMainfm.txtPPAM
   epg1 = CDbl(n_size * 0.5 / PPAM)
   epg2 = -epg1
   estep = CDbl(1 / PPAM)
   nang = n_size
   
Screen.MousePointer = vbHourglass

For KWAV = KMIN To KMAX Step KSTEP   '<1

   wl = 380# + CDbl(KWAV - 1) * 5#
   wl = wl / 1000# 'convert to nm
   
    nhgt = 1 'use so far only one height '<<<<<<<<<<<<<<
    i__1 = nhgt
    For nkhgt = 1 To i__1
'    {
'
        hz = hz1 + (nkhgt - 1) * dhz
'        //mhgt = floor(hz*1000)
'
        XP = 0#
'
'        //fprintf(stream, "%lg, %lg, %lg, %lg\n", XP, hz, A1, A2)
'
'        printf("Observer height (m) = %lg\n", hz * 1000.0)
'
        d__1 = (epg2 - epg1) / estep
        nang = CInt(Abs(d__1)) + 1
'
        START = True
'
'
        'for testing '<<<<<<<<<<<<<<<<<
'        nang = 1
'        epg1 = 0#
        
        Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset

        i__2 = nang
        For kgr = 1 To i__2
            dtg2 = 0#
            dh = 0.002
            epg = epg1 - (kgr - 1) * estep
            ALFA(KWAV, kgr) = epg
            
            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * kgr / nang))
            DoEvents

            If (epg > 0# And epg < estep) Then
               epg = 0#
               End If

            epz = epg / 60# * co
'    /*           write(*,*)' epz=',epz */
            epzm = epz * q3
            rsof = ra + hsof
            If (epz < 0#) Then
                If (hz - dh / 2# > 0#) Then
                    hen1 = hz - dh / 2#
                    dh = -0.002
                ElseIf (hz - dh / 2# <= 0#) Then
                    hen1 = hz
                    End If
            ElseIf (epz >= 0#) Then
                hen1 = hz + dh / 2#
                End If
                
            If hen1 = 0 And zz_1.hj(0) = 0 Then 'hit rock bottom
               jstop = kgr
               Exit For
               End If
            
            If (START) Then
                Write #fileout%, 0#, hz1 * 1000#, ALFA(KWAV, kgr), ALFA(KWAV, kgr), 0#
                START = False
                End If
            
            cz = Cos(epz) * (ra + hz) * (fun_(hen1, zz_1, NumLayers) + 1#)
            bz = pi2 - epz '//zenith angle
            bn = bz
'    /*           write(*,*)' bn=',bn */
            ep1 = q3 * epz

            T = 0#
            S1 = 0#
            at2 = 0#
            nz = CLng((hsof + Abs(dh) - hz) / Abs(dh) + 1.1)  '//number of steps

            h__ = hz - dh '//start at this height

            nent = 0

            i__3 = nz
            For n = 1 To i__3 'trace over all the layers

                b = bn 'initial zenith angle
                e = pi2 - b 'initial view angle

                e1 = q3 * e 'view angle in mrad
                
                d__1 = e / co

                If epg <> 0# And d__1 < 0.00005 Then
                    e = 0#
                    dh = 0.002
                    dtg2 = dtg
                    End If

                If (dh < 0#) Then
                    nz = nz + 1
                    End If
                h__ = h__ + dh
                
                If (h__ < 0#) Then
                    h__ = 0#
                    End If
                    
                rt = ra + h__
'        /* L25: */
                ru = rt + dh
                hen = h__ + dh / 2#
'        /* >>          write(*,*)' hen=',hen */

                If (n = 1) Then
                    hmin = hen
                    hmax = hen
                Else
                    If (hen < hmin) Then
                        hmin = hen
                        End If
                    If (hen > hmax) Then
                        hmax = hen
                        End If
                    End If
                    
                hev = hen + dh
'
                el = sqt_(e, dh, rt) 'path length of light ray from last height to current height
'
                s = el * Cos(e)  'begin law of sin calculation of the subtended cylindrical angle at the Earth's center
'
                A2 = DASIN(s / ru) 'this is the subtended angle, dtheta for this last step
'        /* L30: */
                den = A2 * ra  'this is length along the circumference of the earth for the last subtended angle increment
'
                XP = XP + den
'                'diagnostics
'                /*
'                if (XP >= 249.111) {
'                    ccc = 1
'                }
'                */
                at2 = at2 + A2  'this is the total cylindrical angle, theta
                s2 = at2 * ra 'this is the total length along the circumference of the earth
                a3 = q6 * A2 'ditto in mrad
                en = fun_(hen, zz_1, NumLayers) 'variable portion of the index of refraction at this height
                en1 = en * 10000000# 'normalized to 1.0, where n = 1 + en1
'        /* L35: */
                g = b - A2  'incident angle
'        /* >>          WRITE(*,*)' G(DEG)=',G/CO */
                g1 = q3 * g
                e2 = q3 * (pi2 - g)
                em = fun_(hev, zz_1, NumLayers) + 1#

                sbn = (en + 1#) * Sin(g) / em 'Snell's law, where new angle asin(sbn) which is g + incremental refraction
'        /* >>          WRITE(*,*)' SBN=',SBN */
                If (sbn > 1#) Then
                    sbn = 1#
                    End If

                bn = DASIN(sbn)
                If (g > pi2) Then
                    bn = pie - bn
                    d__1 = g - pi2
                    dh = -el * Abs(d__1)
                    End If

                d__ = bn - g
                d6 = q6 * d__
                T = T + d__
                dt1 = q3 * T
                dtg = dt1 * ramg
                fr = 1#
                If (h__ >= 1.5) Then
                    fr = 1.0001
                    End If
                dh = fr * dh
                dg = dh * q3

                ANGLE = XP / ra 'angle XP subtends in radians
                A1 = XP * Cos(ANGLE) + ru * Sin(ANGLE)
                A2 = -XP * Sin(ANGLE) + ru * Cos(ANGLE) - ra
                

'                'fprintf(stream, "%lg, %lg, %lg, %lg\n", XP, ru - ra, A1, A2)
'                fprintf(stream, "%lg, %lg\n", A1, A2)
'                Write #fileout%, XP, ru 'A1, A2
                
                If ru < ra Then 'collided with the surface
                   Write #fileout%, XP * 1000#, -1000, ALFA(KWAV, kgr), e2, dt1 * 0.001 * 180# * 60 / pi
                   jstop = kgr
                Else
                   'limit the recording to every other step
                   If n Mod Val(prjAtmRefMainfm.txtHeightStepSize.Text) = 0 Then
                      Write #fileout%, XP * 1000#, (ru - ra) * 1000#, ALFA(KWAV, kgr), e2, dt1 * 0.001 * 180# * 60 / pi
                      End If
                   jstop = -1
                   End If
                
                If (rt >= rsof) Then
                    Exit For
                    End If
                    
                fiem = epzm - 4.665 - dt1
                fieg = fiem * ramg
                
                If jstop <> -1 Then Exit For

            Next n
                        
            If jstop <> -1 Then Exit For
'
'            printf("View Angle (deg.) = %lg, Accumlated refraction (mrad) = %lg\n", epg, dt1)
'            If epg = 0 Then
'               prjAtmRefMainfm.lblRef = "View Angle (deg.) = " & epg & vbCrLf & "Accumlated refraction (mrad) = " & dt1
'               End If
            
            ALFT(KWAV, kgr) = ALFA(KWAV, kgr) - dt1 * 0.001 * 180# * 60 / pi 'true depression angle in minutes of degree
            
'            If ALFA(KWAV, kgr) = 0 Then
'               ccc = 1
'               End If

            START = True
            XP = 0#

        Next kgr
        If jstop <> -1 Then Exit For
    Next nkhgt
   If jstop <> -1 Then Exit For
   
Next KWAV

Close #fileout%

Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
prjAtmRefMainfm.progressfrm.Visible = False
     
StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KWAV, jstop - 1)) & vbCrLf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(ra / (ra + hz)) / CONV))
Call StatusMessage(StatusMes, 1, 0)
prjAtmRefMainfm.lblHorizon.Caption = StatusMes
prjAtmRefMainfm.lblHorizon.Refresh
DoEvents
'     ier = MsgBox(StatusMes, vbInformation + vbOKOnly, "Horizon")
StatusMes = "Ray tracing calculation complete..."
Call StatusMessage(StatusMes, 1, 0)

CalcComplete = True

'increase resolution of atmospheric models and load to charts and convert elevations to meters
Dim hgt As Double, Pr As Double, Te As Double
NumTemp = 0
For j = 1 To NNN
   For hgt = zz_1.hj(j - 1) To zz_1.hj(j) - Val(prjAtmRefMainfm.txtHeightStepSize.Text) * 0.001 Step Val(prjAtmRefMainfm.txtHeightStepSize.Text) * 0.001
      Call layers_int(hgt, zz_1, NumLayers, Pr, Te)
      ELV(NumTemp) = hgt * 1000#
      TMP(NumTemp) = Te
      PRSR(NumTemp) = Pr
      NumTemp = NumTemp + 1
   Next hgt
Next j
ELV(NumTemp) = zz_1.hj(NNN - 1) * 1000#
TMP(NumTemp) = zz_1.tj(NNN - 1)
PRSR(NumTemp) = zz_1.pj(NNN - 1)
NNN = NumTemp
'now load up temperature and pressure charts
    
 ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
 
 For j = 1 To NNN
    TransferCurve(j, 1) = " " & CStr(ELV(j - 1) * 0.001)
'         TransferCurve(J, 2) = ELV(J - 1) * 0.001
    TransferCurve(j, 2) = TMP(j - 1)
 Next j
 
 With MSChartTemp
   .chartType = VtChChartType2dLine
   .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN - 1
'        .RowLabel = "Height (km)"
'        .ColumnLabel = "Temperature (Kelvin)"
   .ChartData = TransferCurve
 End With

 For j = 1 To NNN
    TransferCurve(j, 2) = PRSR(j - 1)
 Next j
 
 With MSChartPress
   .chartType = VtChChartType2dLine
   .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN - 1
'        .RowLabel = "Height (km)"
'        .ColumnLabel = "Pressure (Kelvin)"
   .ChartData = TransferCurve
 End With

Screen.MousePointer = vbDefault
    
StatusMes = "Writing transfer curve."
Call StatusMessage(StatusMes, 1, 0)
filnum% = FreeFile
Open App.Path & "\tc_M.dat" For Output As #filnum%
'      WRITE (20,*) N
NumTc = 0
Print #filnum%, n_size
For j = 1 To jstop - 1
'        WRITE(20,1) ALFA(KMIN,J),ALFT(KMIN,J)
    Print #filnum%, ALFA(KMIN, j), ALFT(KMIN, j)
    If ALFA(KMIN, j) = 0 Then 'display the refraction value for the zero view angle ray
       prjAtmRefMainfm.lblRef.Caption = "Atms. refraction (deg.) = " & Abs(ALFT(KMIN, j)) / 60# & vbCrLf & "Atms. refraction (mrad) = " & Abs(ALFT(KMIN, j)) * 1000# * cd / 60#
       prjAtmRefMainfm.lblRef.Refresh
       DoEvents
       End If
'store all view angles that contribute to sun's orb
    NumTc = NumTc + 1
    For KA = 1 To NumSuns
       y = ALFT(KMIN, j) - ALT(KA)
       If Abs(y) <= ROBJ Then
          'only accept rays that pass over the horizon (ALFT(KMIN, J) <> -1000) and are within the solar disk
          SunAngles(KA - 1, NumSunAlt(KA - 1)) = j
          NumSunAlt(KA - 1) = NumSunAlt(KA - 1) + 1
          End If
    Next KA
Next j
Close #filnum%


'now load up transfercurve array for plotting
ReDim TransferCurve(1 To NumTc, 1 To 2) As Variant

For j = 1 To NumTc
 TransferCurve(j, 1) = " " & CStr(ALFA(KMIN, j))
 TransferCurve(j, 2) = ALFT(KMIN, j)
'         TransferCurve(J, 1) = " " & CStr(ALFT(KMIN, J))
'         TransferCurve(J, 2) = ALFA(KMIN, J)
Next j

With MSCharttc
.chartType = VtChChartType2dLine
.RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN
'        .RowLabel = "True angle (min)"
'        .ColumnLabel = "View angle (min)"
.ChartData = TransferCurve
End With

 
 StatusMes = "Drawing the rays on the sky simulation, please wait...."
 Call StatusMessage(StatusMes, 1, 0)
 'load angle combo boxes
'    AtmRefPicSunfm.WindowState = vbMinimized
'    BrutonAtmReffm.WindowState = vbMaximized
 'set size of picref by size of earth
 Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
 prjAtmRefMainfm.cmbSun.Clear
 prjAtmRefMainfm.cmbAlt.Clear
 For i = 1 To NumSuns
    If NumSunAlt(i - 1) > 0 Then prjAtmRefMainfm.cmbSun.AddItem i
 Next i
 
 prjAtmRefMainfm.TabRef.Tab = 4
 DoEvents

cmbSun.ListIndex = 0

   cmdCalc.Enabled = True
   cmdRefWilson.Enabled = True
   cmdMenat.Enabled = True
   cmdVDW.Enabled = True
'
'    printf("Do you want to a new calculation? (y/n) -->")
'    scanf("%s", chr)
'    if (strstr(chr, "y")) goto L0
'
'
'    return 0
'} /* MAIN__ */
'
'
   On Error GoTo 0
   Exit Sub

cmdMenat_Click_Error:
    Close
    Screen.MousePointer = vbDefault
    
    StatusMes = sEmpty
    Call StatusMessage(StatusMes, 1, 0)
    Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
    prjAtmRefMainfm.progressfrm.Visible = False
    
   cmdCalc.Enabled = True
   cmdRefWilson.Enabled = True
   cmdMenat.Enabled = True
   cmdVDW.Enabled = True
   
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdMenat_Click of Form prjAtmRefMainfm"
End Sub
    
    
'---------------------------------------------------------------------------------------
' Procedure : sqt_
' Author    : Dr-John-K-Hall
' Date      : 2/18/2019
' Purpose   : Calculates ray path length for Menat ray tracing using law of cosines
'---------------------------------------------------------------------------------------
'
Public Function sqt_(e As Double, dh As Double, r__ As Double) As Double
'{
'    /* System generated locals */
    Dim ret_val As Double, d__1 As Double
'
'    /* Builtin functions */
'    'double sin(double), sqrt(double)
'
'    /* Local variables */
    Dim q As Double, y As Double

    q = r__ * Sin(e)
    y = r__ * 2# + dh
'/* >>     WRITE(*,*)' DH=',DH */
    If (dh < 0#) Then
        d__1 = q * q + dh * y
        ret_val = -q - Sqr(Abs(d__1))
        End If
    If (dh >= 0#) Then
        d__1 = q * q + dh * y
        ret_val = -q + Sqr(Abs(d__1))
        End If
        
'/*        if (q*q+dh*y .lt. 0) then */
'/*           write(*,*)'-q,quo,sqt=',-q,q*q+dh*y,sqt */
'/*           end if */
    sqt_ = ret_val
    
    
End Function
'} /* sqt_ */
'
'

'} /* fun_ */
'
'
'int LoadAtmospheres(char filnam[])
'{
'    FILE *stream
'
'    if ( !( stream = fopen( filnam, "r")) )
'    {
'        return -1
'    }
'
'    numlayers = 0
'    while (!feof(stream) ) {
'
'        fscanf(stream, "%lg %lg %lg\n", &layerheights[numlayers], &temps[numlayers], &press[numlayers])
'        numlayers ++
'
'    }
'    fclose (stream)
'    return 0
'
'}
'/* Main program alias */ //int menat_ () { MAIN__ () return 0 }



Private Sub cmdOtherBrowse_Click()
   On Error GoTo errhand
   comdlgOther.CancelError = True
   comdlgOther.Filter = "Text files (.txt)|*.txt|All files (*.*)|*.*"
   comdlgOther.ShowOpen
   txtOther.Text = comdlgOther.filename
   If InStr(txtOther.Text, "-sondes.txt") Then 'need to convert from meters of height to kms
      prjAtmRefMainfm.chkMeters.Value = vbChecked
      opt10.Value = True
      End If
errhand:
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdPlotFit_Click
' Author    : Dr-John-K-Hall
' Date      : 9/15/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdPlotFit_Click()
  Dim FileNameIn As String
  Dim VAFit() As Double, Refr() As Double, j As Long
  Dim HgtFit() As Double, TKFit() As Double
'  Dim TotHgtFit() As Double, TotRefr() As Double
  Dim BestM As Double, BestB As Double ', N500 As Integer
  
  Set PtX = New Collection
  Set PtY = New Collection
  
  On Error GoTo cmdPlotFit_Click_Error

  FileNameIn = App.Path & "\TR_VDW_Total_Refraction.dat"
  
  If Dir(FileNameIn) = sEmpty And chkRefFiles_Ref.Value = vbUnchecked And chkRefFiles_dip.Value = vbUnchecked Then
     Call MsgBox("Can't find the refraction data file:" _
                 & vbCrLf & FileNameIn _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: it needs to be in " & App.Path _
                 , vbExclamation, "File Missing")
     Exit Sub
     End If
     
  filein% = FreeFile
  Open FileNameIn For Input As #filein%
  j = 0
  
  maxref = -999999999
  minref = 999999999
  maxva = -999999999
  minva = 999999999
  maxhgt = -9999999999#
  minhgt = 9999999999#
  mintk = 999999999999#
  maxtk = -999999999999#
  Screen.MousePointer = vbHourglass
  If chkFit1.Value = vbChecked And chkHgt.Value = vbChecked And chkRefFiles_Ref.Value = vbUnchecked And chkRefFiles_dip.Value = vbUnchecked Then
     'add refraction as a function of view angle for chosen temperature and observer height
      Do Until EOF(filein%)
         Input #filein%, Tfit, Hfit, LObs, VA, ref
         If Tfit = Val(txtFit1) And Hfit = Val(txtHgtFit) Then
            found% = 1
            ReDim Preserve VAFit(j + 1)
            ReDim Preserve Refr(j + 1)
            VAFit(j) = VA / 60#
            Refr(j) = ref
            PtX.Add VAFit(j)
            PtY.Add Refr(j)
            If ref > maxref Then maxref = ref
            If ref < minref Then minref = ref
            If VAFit(j) > maxva Then maxva = VAFit(j)
            If VAFit(j) < minva Then minva = VAFit(j)
            j = j + 1
       Else
            If found% = 1 Then Exit Do
            End If
      Loop
      Close #filein%
      
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         Call MsgBox("Nothing found!" _
                     & vbCrLf & "" _
                     & vbCrLf & "Try different temperature and height" _
                     , vbInformation, "Nothing found")
         
         Exit Sub
      Else
        NNN = j - 1
        End If
        
        ' Find a good fit.
        degree = 6
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
        
        '    stop_time = Timer
        '    Debug.Print Format$(stop_time - start_time, "0.0000")
        
        txtFitResults = ""
        For i = 1 To BestCoeffs.Count
            txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
        Next i
        If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
        
        ' Display the error.
        Call ShowError(, ErrorLbl, 1)
        
        ' We have a solution.
        HasSolution = Tru
        
      Screen.MousePointer = vbDefault
      
     'set up chart and plot the data then fit it
     ReDim TransferCurve(0, 0) As Variant
      ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
      n = 0
      For j = NNN - 1 To 0 Step -1
         TransferCurve(n, 0) = " " + Format(CStr(VAFit(j + 1)), "##0.0#")
         TransferCurve(n, 1) = Refr(j + 1)
         TransferCurve(n, 2) = BestCoeffs.Item(1) + _
                               BestCoeffs.Item(2) * VAFit(j + 1) + _
                               BestCoeffs.Item(3) * VAFit(j + 1) ^ 2# + _
                               BestCoeffs.Item(4) * VAFit(j + 1) ^ 3# + _
                               BestCoeffs.Item(5) * VAFit(j + 1) ^ 4# + _
                               BestCoeffs.Item(6) * VAFit(j + 1) ^ 5# + _
                               BestCoeffs.Item(7) * VAFit(j + 1) ^ 6#
         n = n + 1
      Next j
      
    With MSChartTR
        .chartType = VtChChartType2dLine
        .Title = "Total Atmospheric Refraction (degrees)"
        .RandomFill = False
        .ShowLegend = True
        .ChartData = TransferCurve
        With .Plot
            With .Wall.Brush
                .Style = VtBrushStyleSolid
                .FillColor.Set 255, 255, 255
            End With
            With .Axis(VtChAxisIdX)
                .AxisTitle = "View Angle (mrad)"
            End With
            With .Axis(VtChAxisIdX).CategoryScale
                .Auto = False
                .DivisionsPerLabel = NNN * 0.1
                .DivisionsPerTick = NNN * 0.1
                .LabelTick = True
            End With
            With .Axis(VtChAxisIdY)
                .AxisTitle = "Refraction (degrees)"
                With .ValueScale
                    .Auto = False
                    .Minimum = 0.9 * minref
                    .Maximum = 1.1 * maxref
                    .MajorDivision = 10
                End With
            End With
            With .Axis(VtChAxisIdY2)
                .AxisTitle = "Fit to Atmospheric Refraction (degrees)"
                With .ValueScale
                    .Auto = False
                    .Minimum = 0.9 * minref
                    .Maximum = 1.1 * maxref
                    .MajorDivision = 10
                End With
            End With
            With .SeriesCollection(1)
                .SeriesMarker.Show = False
                .LegendText = "Refraction Calculation"
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 0, 0, 255 'blue
                End With
                With .DataPoints(-1).Marker
                    .Style = VtMarkerStyleDiamond
                End With
            End With
            With .SeriesCollection(2)
                .LegendText = "Refraction Fit"
                .SecondaryAxis = True
                .SeriesMarker.Show = False
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 255, 0, 0 'red
                End With
            End With
        End With
    End With

    ElseIf chkFit1.Value = vbChecked And chkVA.Value = vbChecked And chkRefFiles_Ref.Value = vbUnchecked And chkRefFiles_dip.Value = vbUnchecked Then
    
       'choose all the refraction values as a function of observer height for a fixed temperature and a fixed view angle
      j = 0

      Do Until EOF(filein%)
         Input #filein%, Tfit, Hfit, LObs, VA, ref
         If Tfit = Val(txtFit1) And VA = Val(txtVA) * 60 Then
            found% = 1
            ReDim Preserve HgtFit(j + 1)
            ReDim Preserve Refr(j + 1)
            HgtFit(j) = Hfit
            Refr(j) = ref
            PtX.Add HgtFit(j)
            PtY.Add Refr(j)
            If ref > maxref Then maxref = ref
            If ref < minref Then minref = ref
            If HgtFit(j) > maxhgt Then maxhgt = HgtFit(j)
            If HgtFit(j) < minhgt Then minhgt = HgtFit(j)
            j = j + 1
            End If
      Loop
      Close #filein%

      If found% = 0 Then
         Screen.MousePointer = vbDefault
         Call MsgBox("Nothing found!" _
                     & vbCrLf & "" _
                     & vbCrLf & "Try different temperature and height" _
                     , vbInformation, "Nothing found")
         
         Exit Sub
      Else
        NNN = j - 1
        End If
        
        ' Find a good fit.
        degree = 3
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
        
        '    stop_time = Timer
        '    Debug.Print Format$(stop_time - start_time, "0.0000")
        
        txtFitResults = ""
        For i = 1 To BestCoeffs.Count
            txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
        Next i
        If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
        
        ' Display the error.
        Call ShowError(, ErrorLbl, 1)
        
        ' We have a solution.
        HasSolution = True
        
      Screen.MousePointer = vbDefault
      
        'set up chart and plot the data then fit it
        ReDim TransferCurve(0, 0) As Variant
         
        ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
        n = 0
        For j = 0 To NNN - 1
           TransferCurve(j, 0) = " " + Format(CStr(HgtFit(j + 1)), "###0.0#")
           TransferCurve(j, 1) = Refr(j + 1)
           TransferCurve(j, 2) = BestCoeffs.Item(1) + _
                                 BestCoeffs.Item(2) * TKFit(j + 1) + _
                                 BestCoeffs.Item(3) * TKFit(j + 1) ^ 2# + _
                                 BestCoeffs.Item(4) * TKFit(j + 1) ^ 3#
        Next j
      
        With MSChartTR
            .chartType = VtChChartType2dLine
            .Title = "Total Atmospheric Refraction (degrees)"
            .RandomFill = False
            .ShowLegend = True
            .ChartData = TransferCurve
            With .Plot
                With .Wall.Brush
                    .Style = VtBrushStyleSolid
                    .FillColor.Set 255, 255, 255
                End With
                With .Axis(VtChAxisIdX)
                    .AxisTitle = "Observer's Height (m)"
                End With
                With .Axis(VtChAxisIdX).CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = NNN * 0.1
                    .DivisionsPerTick = NNN * 0.1
                    .LabelTick = True
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle = "Refraction (degrees)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = "Fit to Refraction (degrees)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .SeriesCollection(1)
                    .SeriesMarker.Show = True
                    .LegendText = "Refraction"
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 0, 0, 255
                    End With
                    With .DataPoints(-1).Marker
                        .Style = VtMarkerStyleDiamond
                    End With
                End With
                With .SeriesCollection(2)
                    .LegendText = "Fit"
                    .SecondaryAxis = True
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 155, 0, 0
                    End With
                End With
            End With
        End With
      
      
    ElseIf chkHgt.Value = vbChecked And chkVA.Value = vbChecked And chkRefFiles_Ref.Value = vbUnchecked And chkRefFiles_dip.Value = vbUnchecked Then
    
       'choose all the refraction values as a function of temperature for a fixed height and a fixed view angle
      j = 0

      Do Until EOF(filein%)
         Input #filein%, Tfit, Hfit, LObs, VA, ref
         If VA = Val(txtVA) And Hfit = Val(txtHgtFit) Then
            found% = 1
            ReDim Preserve TKFit(j + 1)
            ReDim Preserve Refr(j + 1)
            TKFit(j) = Tfit
            Refr(j) = ref
            PtX.Add TKFit(j)
            PtY.Add Refr(j)
            If ref > maxref Then maxref = ref
            If ref < minref Then minref = ref
            If TKFit(j) > maxtk Then maxtk = TKFit(j)
            If TKFit(j) < mintk Then mintk = TKFit(j)
            j = j + 1
            End If
      Loop
      Close #filein%
      
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         Call MsgBox("Nothing found!" _
                     & vbCrLf & "" _
                     & vbCrLf & "Try different temperature and height" _
                     , vbInformation, "Nothing found")
         
         Exit Sub
      Else
        NNN = j - 1
        End If
        
        ' Find a good fit.
        degree = 3
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
        
        '    stop_time = Timer
        '    Debug.Print Format$(stop_time - start_time, "0.0000")
        
        txtFitResults = ""
        For i = 1 To BestCoeffs.Count
            txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
        Next i
        If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
        
        ' Display the error.
        Call ShowError(, ErrorLbl, 1)
        
        ' We have a solution.
        HasSolution = True
        
      Screen.MousePointer = vbDefault
      
      'set up chart and plot the data then fit it
       ReDim TransferCurve(0, 0) As Variant
      
        ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
        For j = 0 To NNN - 1
           TransferCurve(j, 0) = " " + Format(CStr(TKFit(j + 1)), "###0.0#")
           TransferCurve(j, 1) = Refr(j + 1)
           TransferCurve(j, 2) = BestCoeffs.Item(1) + _
                                 BestCoeffs.Item(2) * TKFit(j + 1) + _
                                 BestCoeffs.Item(3) * TKFit(j + 1) ^ 2# + _
                                 BestCoeffs.Item(4) * TKFit(j + 1) ^ 3#
        Next j
        
        With MSChartTR
            .chartType = VtChChartType2dLine
            .Title = "Total Atmospheric Refraction (degrees)"
            .RandomFill = False
            .ShowLegend = True
            .ChartData = TransferCurve
            With .Plot
                With .Wall.Brush
                    .Style = VtBrushStyleSolid
                    .FillColor.Set 255, 255, 255
                End With
                With .Axis(VtChAxisIdX)
                    .AxisTitle = "Temperature at Observer (deg. K)"
                End With
                With .Axis(VtChAxisIdX).CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = NNN * 0.1
                    .DivisionsPerTick = NNN * 0.1
                    .LabelTick = True
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle = "Refraction (degrees)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = "Fit to Refraction (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .SeriesCollection(1)
                    .SeriesMarker.Show = True
                    .LegendText = "Refraction"
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 255, 0, 0
                    End With
                    With .DataPoints(-1).Brush
                        .FillColor.Set 0, 0, 255
                    End With
                    With .DataPoints(-1).Marker
                        .Style = VtMarkerStyleFilledCircle
                    End With
                End With
                With .SeriesCollection(2)
                    .LegendText = "Fit"
                    .SecondaryAxis = True
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 255, 0, 0
                    End With
                End With
            End With
        End With
    ElseIf Not chkRefFiles_Ref.Value = vbChecked And Not chkRefFiles_dip.Value = vbChecked Then
    
        Call MsgBox("You didn't check one of the possible options:" _
                    & vbCrLf & "" _
                    & vbCrLf & "1. Temperature and Height check boxes" _
                    & vbCrLf & "(Refraction vs. view angle for fixed temp and height)" _
                    & vbCrLf & "" _
                    & vbCrLf & "2. Temperature and View Angle checkboxes" _
                    & vbCrLf & "(Refraction vs observer height for fixed temp and view angle)" _
                    & vbCrLf & "" _
                    & vbCrLf & "3. Height and View Angle checkboxes" _
                    & vbCrLf & "(refraction vs temperature for fixed view angle, height)" _
                    , vbInformation, "Improper choice")
                    
        Exit Sub
        
    
    ElseIf chkRefFiles_Ref.Value = vbChecked And chkRefFiles_dip.Value = vbChecked Then
        Close
        FilePath = txtRefFileDir
        FileNameIn = FilePath & "\TR_VDW_260-320_0_32.dat"
        
        If Dir(FileNameIn) = sEmpty Then
          Call MsgBox("Can't open the file:" _
                      & vbCrLf & FileNameIn _
                      & vbCrLf & "" _
                      & vbCrLf & "It doesn't seem to exit" _
                      , vbInformation, "Missing file")
          Exit Sub
          End If
          
       'choose all the refraction values as a function of observer height for a fixed temperature and a fixed view angle
      filein% = FreeFile
      Open FileNameIn For Input As #filein%
      j = 0
       
      Do Until EOF(filein%)
      
        Input #filein%, Tfit, ref
        ReDim Preserve TKFit(j + 1)
        ReDim Preserve Refr(j + 1)
        TKFit(j) = Tfit
        Refr(j) = ref
        PtX.Add TKFit(j)
        PtY.Add Refr(j)
        If ref > maxref Then maxref = ref
        If ref < minref Then minref = ref
        If TKFit(j) > maxtk Then maxtk = TKFit(j)
        If TKFit(j) < mintk Then mintk = TKFit(j)
        j = j + 1

      Loop
      Close #filein%
      NNN = j - 1
        
        ' Find a good fit.
        degree = 3
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
        
        '    stop_time = Timer
        '    Debug.Print Format$(stop_time - start_time, "0.0000")
        
        txtFitResults = ""
        For i = 1 To BestCoeffs.Count
            txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
        Next i
        If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
        
        ' Display the error.
        Call ShowError(, ErrorLbl, 1)
        
        ' We have a solution.
        HasSolution = True
        
      Screen.MousePointer = vbDefault
      
      'set up chart and plot the data then fit it
       ReDim TransferCurve(0, 0) As Variant
      
        ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
        For j = 0 To NNN - 1
           TransferCurve(j, 0) = " " + Format(CStr(TKFit(j + 1)), "###0.0#")
           TransferCurve(j, 1) = Refr(j + 1)
           TransferCurve(j, 2) = BestCoeffs.Item(1) + _
                                 BestCoeffs.Item(2) * TKFit(j + 1) + _
                                 BestCoeffs.Item(3) * TKFit(j + 1) ^ 2# + _
                                 BestCoeffs.Item(4) * TKFit(j + 1) ^ 3#
        Next j
        
        With MSChartTR
            .chartType = VtChChartType2dLine
            .Title = "Total Atmospheric Refraction (degrees)"
            .RandomFill = False
            .ShowLegend = True
            .ChartData = TransferCurve
            With .Plot
                With .Wall.Brush
                    .Style = VtBrushStyleSolid
                    .FillColor.Set 255, 255, 255
                End With
                With .Axis(VtChAxisIdX)
                    .AxisTitle = "Temperature at Observer (deg. K)"
                End With
                With .Axis(VtChAxisIdX).CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = NNN * 0.1
                    .DivisionsPerTick = NNN * 0.1
                    .LabelTick = True
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle = "Refraction (degrees)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = "Fit to Refraction (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .SeriesCollection(1)
                    .SeriesMarker.Show = True
                    .LegendText = "Refraction"
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 255, 0, 0
                    End With
                    With .DataPoints(-1).Brush
                        .FillColor.Set 0, 0, 255
                    End With
                    With .DataPoints(-1).Marker
                        .Style = VtMarkerStyleFilledCircle
                    End With
                End With
                With .SeriesCollection(2)
                    .LegendText = "Fit"
                    .SecondaryAxis = True
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 255, 0, 0
                    End With
                End With
            End With
        End With

        
        'plot and fit the totatl atmospheric refraction as a function of temperature
'    260.00000     11.40697
'    263.00000     11.17969
'    266.00000     10.96167
'    269.00000     10.75072
'    272.00000     10.54709
'    275.00000     10.35096
'    278.00000     10.16096
'    281.00000      9.97698
'    284.00000      9.79948
'    287.00000      9.62726
'    290.00000      9.46016
'    293.00000      9.29830
'    296.00000      9.14155
'    299.00000      8.98936
'    302.00000      8.84135
'    305.00000      8.69739
'    308.00000      8.55760
'    311.00000      8.42181
'    314.00000      8.28967
'    317.00000      8.16102
'    320.00000      8.03602
    ElseIf chkRefFiles_Ref.Value = vbChecked And chkFit1.Value = vbChecked And chkVA.Value = vbChecked Then
        'plot and fit the leveling refraction as a function of ray (observer) height for a chosen temperature
        Close
        FilePath = txtRefFileDir
        
        Tfit = Val(txtFit1)
        VAngFit = Val(txtVA)
        FileNameIn = FilePath & "\TR_VDW_" & Trim$(txtFit1) & "_0_32.dat"
        FilePath = txtRefFileDir
        
        If Dir(FileNameIn) = sEmpty Then
          Call MsgBox("Can't open the file:" _
                      & vbCrLf & FileNameIn _
                      & vbCrLf & "" _
                      & vbCrLf & "It doesn't seem to exit" _
                      , vbInformation, "Missing file")
          Exit Sub
          End If
          
      found% = 0
      filein% = FreeFile
      j = 0
      Open FileNameIn For Input As #filein%
      Do Until EOF(filein%)
         Input #filein%, Dist, Hfit, VAO, Dip, ref
         If VAngFit * 60 = VAO And Hfit <= 4000 Then
            found% = 1
            If Hfit <= 500 Then
               N500 = j
               End If
            ReDim Preserve HgtFit(j + 1)
            ReDim Preserve Refr(j + 1)
            HgtFit(j) = Hfit
            Refr(j) = ref
            
            PtX.Add HgtFit(j)
            PtY.Add Refr(j)

            If ref > maxref Then maxref = ref
            If ref < minref Then minref = ref
            If HgtFit(j) > maxhgt Then maxhgt = HgtFit(j)
            If HgtFit(j) < minhgt Then minhgt = HgtFit(j)
            j = j + 1
         ElseIf VAngFit * 60 <> VAO And found% = 1 Then
            'passed the selected VA, so exit loop
            Exit Do
            End If
      Loop
      Close #filein%
      
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         Call MsgBox("Nothing found!" _
                     & vbCrLf & "" _
                     & vbCrLf & "Try different temperature and height" _
                     , vbInformation, "Nothing found")
         
         Exit Sub
      Else
        NNN = j - 1
'        'improve accuracy of the fit by adding reflection around zero to 500 m
'        NewNum = 0
'        For j = 0 To N500
'           ReDim Preserve TotRefr(j + 1)
'           ReDim Preserve TotHgtFit(j + 1)
'           TotRefr(NewNum) = -Refr(N500 - NewNum)
'           TotHgtFit(NewNum) = -HgtFit(N500 - NewNum)
'           PtX.Add TotHgtFit(NewNum)
'           PtY.Add TotRefr(NewNum)
'           NewNum = NewNum + 1
'        Next j
'        NewNum2 = NewNum
'        For j = NewNum To NNN + NewNum - 1
'           ReDim Preserve TotRefr(j)
'           ReDim Preserve TotHgtFit(j)
'           TotRefr(j) = Refr(j - NewNum + 1)
'           TotHgtFit(j) = HgtFit(j - NewNum + 1)
'           PtX.Add TotHgtFit(j)
'           PtY.Add TotRefr(j)
'           NewNum2 = NewNum2 + 1
'        Next j
'        NNN = NewNum2 - 1
'        minhgt = 0.9 * TotHgtFit(0)
'        minref = 0.9 * TotRefr(0)
        
'        'improve accuracy by adding points for x<0
'        For j = 250 To 1 Step -1
'           PtX.Add -HgtFit(j)
'           PtY.Add -Refr(j)
'        Next j
'        For j = 0 To NNN
'            PtX.Add HgtFit(j)
'            PtY.Add Refr(j)
'        Next j
        
        End If
        
        ' Find a good fit.
        degree = 7
        
        If degree > 1 Then
           Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
           
            txtFitResults = ""
            For i = 1 To BestCoeffs.Count
                txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
            Next i
            If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
            
            ' Display the error.
            Call ShowError(, ErrorLbl, 1)
            
            ' We have a solution.
            HasSolution = True
            
            Screen.MousePointer = vbDefault
          
            'set up chart and plot the data then fit it
            ReDim TransferCurve(0, 0) As Variant
             
            ReDim TransferCurve(0 To NNN, 0 To 2) As Variant
            
            fileout% = FreeFile
            Open App.Path & "\Point-100-fit.txt" For Output As #fileout%
            For j = 0 To NNN
               TransferCurve(j, 0) = " " + Format(CStr(HgtFit(j)), "###0.0#")
               TransferCurve(j, 1) = Refr(j)
               TransferCurve(j, 2) = BestCoeffs.Item(1) + _
                                     BestCoeffs.Item(2) * HgtFit(j) + _
                                     BestCoeffs.Item(3) * HgtFit(j) ^ 2# + _
                                     BestCoeffs.Item(4) * HgtFit(j) ^ 3# + _
                                     BestCoeffs.Item(5) * HgtFit(j) ^ 4# + _
                                     BestCoeffs.Item(6) * HgtFit(j) ^ 5# + _
                                     BestCoeffs.Item(7) * HgtFit(j) ^ 6# + _
                                     BestCoeffs.Item(8) * HgtFit(j) ^ 7#
                Write #fileout%, HgtFit(j), Refr(j) ', TransferCurve(j, 2)
            Next j
            Close #fileout%
        Else
           Call FindLinearLeastSquaresFit(PtX, PtY, BestM, BestB)
           
            txtFitResults = ""
            txtFitResults = BestB & " " & BestM
            If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
           
           Call ShowLinError(, ErrorLbl, 1)
           
           HasSolution = True
           
            'set up chart and plot the data then fit it
            ReDim TransferCurve(0, 0) As Variant
'            BestB = 0
'            BestM = Refr(NNN - 1) / HgtFit(NNN - 1)
            ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
            For j = 0 To NNN - 1
               TransferCurve(j, 0) = " " + Format(CStr(HgtFit(j + 1)), "###0.0#")
               TransferCurve(j, 1) = Refr(j + 1)
               TransferCurve(j, 2) = BestB + BestM * HgtFit(j + 1)
            Next j
           End If
        
        With MSChartTR
            .chartType = VtChChartType2dLine
            .Title = "Local Refraction (mrad)"
            .RandomFill = False
            .ShowLegend = True
            .ChartData = TransferCurve
            With .Plot
                With .Wall.Brush
                    .Style = VtBrushStyleSolid
                    .FillColor.Set 255, 255, 255
                End With
                With .Axis(VtChAxisIdX)
                    .AxisTitle = "Observer's Height (m)"
                End With
                With .Axis(VtChAxisIdX).CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = NNN * 0.1
                    .DivisionsPerTick = NNN * 0.1
                    .LabelTick = True
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle = "Refraction (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = "Fit to Refraction (degrees)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .SeriesCollection(1)
                    .SeriesMarker.Show = False 'True
                    .LegendText = "Refraction"
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 0, 0, 255
                    End With
                    With .DataPoints(-1).Marker
                        .Style = VtMarkerStyleDiamond
                    End With
                End With
                With .SeriesCollection(2)
                    .LegendText = "Fit"
                    .SecondaryAxis = True
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 155, 0, 0
                    End With
                End With
            End With
        End With
        
        Screen.MousePointer = vbDefault
       
' TR_VDW_296_0_32.dat
' TR_VDW_293_0_32.dat
'
'
'dist, height, view angle, dip, ref
'1153592.81106  97160.06975      0.00000    172.17649      9.29830
'1154675.39167  97351.30537      0.00000    172.34680      9.29830
'1155758.98096  97542.91798      0.00000    172.51726      9.29830
'1156843.57977  97734.90832      0.00000    172.68788      9.29830
'1157929.18893  97927.27712      0.00000    172.85866      9.29830
'1159015.80927  98120.02511      0.00000    173.02960      9.29830
'1160103.44164  98313.15304      0.00000    173.20070      9.29830
    
    ElseIf chkRefFiles_dip.Value = vbChecked And chkFit1.Value = vbChecked And chkVA.Value = vbChecked Then
        'plot and fit the leveling dip as a function of ray (observer) height for a chosen temperature
        Close
        FilePath = txtRefFileDir
        
        Tfit = Val(txtFit1)
        VAngFit = Val(txtVA)
        FileNameIn = FilePath & "\TR_VDW_" & Trim$(txtFit1) & "_0_32.dat"
        FilePath = txtRefFileDir
        
        If Dir(FileNameIn) = sEmpty Then
          Call MsgBox("Can't open the file:" _
                      & vbCrLf & FileNameIn _
                      & vbCrLf & "" _
                      & vbCrLf & "It doesn't seem to exit" _
                      , vbInformation, "Missing file")
          Exit Sub
          End If
          
      filein% = FreeFile
      Open FileNameIn For Input As #filein%
      j = 0
      Do Until EOF(filein%)
         Input #filein%, Dist, Hfit, VAO, Dip, ref
         If VAngFit * 60 = VAO And Hfit <= 4000 Then
            found% = 1
            ReDim Preserve HgtFit(j + 1)
            ReDim Preserve Refr(j + 1)
            HgtFit(j) = Hfit
            Refr(j) = Dip 'local dip in mrad
            PtX.Add HgtFit(j)
            PtY.Add Refr(j)
            If Dip > maxref Then maxref = Dip
            If Dip < minref Then minref = Dip
            If HgtFit(j) > maxhgt Then maxhgt = HgtFit(j)
            If HgtFit(j) < minhgt Then minhgt = HgtFit(j)
            j = j + 1
         ElseIf VAngFit * 60 <> VAO And found% = 1 Then
            'passed the selected VA, so exit loop
            Exit Do
            End If
      Loop
      Close #filein%
      
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         Call MsgBox("Nothing found!" _
                     & vbCrLf & "" _
                     & vbCrLf & "Try different temperature and height" _
                     , vbInformation, "Nothing found")
         
         Exit Sub
      Else
        NNN = j - 1
        End If
        ' Find a good fit.
        degree = 6
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
        
        '    stop_time = Timer
        '    Debug.Print Format$(stop_time - start_time, "0.0000")
        
        txtFitResults = ""
        For i = 1 To BestCoeffs.Count
            txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
        Next i
        If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
        
        ' Display the error.
        Call ShowError(, ErrorLbl, 1)
        
        ' We have a solution.
        HasSolution = True
        
      Screen.MousePointer = vbDefault
      
        'set up chart and plot the data then fit it
        ReDim TransferCurve(0, 0) As Variant
         
        ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
        For j = 0 To NNN - 1
           TransferCurve(j, 0) = " " + Format(CStr(HgtFit(j + 1)), "###0.0#")
           TransferCurve(j, 1) = Refr(j + 1)
           TransferCurve(j, 2) = BestCoeffs.Item(1) + _
                                 BestCoeffs.Item(2) * HgtFit(j + 1) + _
                                 BestCoeffs.Item(3) * HgtFit(j + 1) ^ 2# + _
                                 BestCoeffs.Item(4) * HgtFit(j + 1) ^ 3# + _
                                 BestCoeffs.Item(5) * HgtFit(j + 1) ^ 4# + _
                                 BestCoeffs.Item(6) * HgtFit(j + 1) ^ 5# + _
                                 BestCoeffs.Item(7) * HgtFit(j + 1) ^ 6#
        Next j
      
        With MSChartTR
            .chartType = VtChChartType2dLine
            .Title = "Dip (degrees)"
            .RandomFill = False
            .ShowLegend = True
            .ChartData = TransferCurve
            With .Plot
                With .Wall.Brush
                    .Style = VtBrushStyleSolid
                    .FillColor.Set 255, 255, 255
                End With
                With .Axis(VtChAxisIdX)
                    .AxisTitle = "Observer's Height (m)"
                End With
                With .Axis(VtChAxisIdX).CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = NNN * 0.1
                    .DivisionsPerTick = NNN * 0.1
                    .LabelTick = True
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle = "Dip (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = "Fit to Dip (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .SeriesCollection(1)
                    .SeriesMarker.Show = False 'True
                    .LegendText = "Refraction"
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 0, 0, 255
                    End With
                    With .DataPoints(-1).Marker
                        .Style = VtMarkerStyleDiamond
                    End With
                End With
                With .SeriesCollection(2)
                    .LegendText = "Fit"
                    .SecondaryAxis = True
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 155, 0, 0
                    End With
                End With
            End With
        End With
    ElseIf chkRefFiles_Ref.Value = vbChecked And chkFit1.Value = vbChecked And chkVA.Value = vbChecked And chkHgt.Value = vbChecked Then
        'plot and fit the leveling refraction as a function of ray (observer) height for a chosen temperature
        Close
        FilePath = txtRefFileDir
        
        Tfit = Val(txtFit1)
        VAngFit = Val(txtVA)
        HeightFit = Val(txtHgtFit)
        FileNameIn = FilePath & "\TR_VDW_" & Trim$(txtFit1) & "_" & Trim$(txtHgtFit) & "_32.dat"
        FilePath = txtRefFileDir
        
        If Dir(FileNameIn) = sEmpty Then
          Call MsgBox("Can't open the file:" _
                      & vbCrLf & FileNameIn _
                      & vbCrLf & "" _
                      & vbCrLf & "It doesn't seem to exit" _
                      , vbInformation, "Missing file")
          Exit Sub
          End If
          
      found% = 0
      filein% = FreeFile
      j = 0
      Open FileNameIn For Input As #filein%
      Do Until EOF(filein%)
         Input #filein%, Dist, Hfit, VAO, Dip, ref
         If VAngFit * 60 = VAO And Hfit <= 4000 Then
            found% = 1
            ReDim Preserve HgtFit(j + 1)
            ReDim Preserve Refr(j + 1)
            HgtFit(j) = Hfit
            Refr(j) = ref
            PtX.Add HgtFit(j)
            PtY.Add Refr(j)
            If ref > maxref Then maxref = ref
            If ref < minref Then minref = ref
            If HgtFit(j) > maxhgt Then maxhgt = HgtFit(j)
            If HgtFit(j) < minhgt Then minhgt = HgtFit(j)
            j = j + 1
         ElseIf VAngFit * 60 <> VAO And found% = 1 Then
            'passed the selected VA, so exit loop
            Exit Do
            End If
      Loop
      Close #filein%
      
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         Call MsgBox("Nothing found!" _
                     & vbCrLf & "" _
                     & vbCrLf & "Try different temperature and height" _
                     , vbInformation, "Nothing found")
         
         Exit Sub
      Else
        NNN = j - 1
        End If
        
        ' Find a good fit.
        degree = 6
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
        
        '    stop_time = Timer
        '    Debug.Print Format$(stop_time - start_time, "0.0000")
        
        txtFitResults = ""
        For i = 1 To BestCoeffs.Count
            txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
        Next i
        If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
        
        ' Display the error.
        Call ShowError(, ErrorLbl, 1)
        
        ' We have a solution.
        HasSolution = True
        
      Screen.MousePointer = vbDefault
      
        'set up chart and plot the data then fit it
        ReDim TransferCurve(0, 0) As Variant
         
        ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
        For j = 0 To NNN - 1
           TransferCurve(j, 0) = " " + Format(CStr(HgtFit(j + 1)), "###0.0#")
           TransferCurve(j, 1) = Refr(j + 1)
           TransferCurve(j, 2) = BestCoeffs.Item(1) + _
                                 BestCoeffs.Item(2) * HgtFit(j + 1) + _
                                 BestCoeffs.Item(3) * HgtFit(j + 1) ^ 2# + _
                                 BestCoeffs.Item(4) * HgtFit(j + 1) ^ 3# + _
                                 BestCoeffs.Item(5) * HgtFit(j + 1) ^ 4# + _
                                 BestCoeffs.Item(6) * HgtFit(j + 1) ^ 5# + _
                                 BestCoeffs.Item(7) * HgtFit(j + 1) ^ 6#
        Next j
      
        With MSChartTR
            .chartType = VtChChartType2dLine
            .Title = "Local Refraction (mrad)"
            .RandomFill = False
            .ShowLegend = True
            .ChartData = TransferCurve
            With .Plot
                With .Wall.Brush
                    .Style = VtBrushStyleSolid
                    .FillColor.Set 255, 255, 255
                End With
                With .Axis(VtChAxisIdX)
                    .AxisTitle = "Observer's Height (m)"
                End With
                With .Axis(VtChAxisIdX).CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = NNN * 0.1
                    .DivisionsPerTick = NNN * 0.1
                    .LabelTick = True
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle = "Refraction (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = "Fit to Refraction (degrees)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .SeriesCollection(1)
                    .SeriesMarker.Show = False 'True
                    .LegendText = "Refraction"
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 0, 0, 255
                    End With
                    With .DataPoints(-1).Marker
                        .Style = VtMarkerStyleDiamond
                    End With
                End With
                With .SeriesCollection(2)
                    .LegendText = "Fit"
                    .SecondaryAxis = True
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 155, 0, 0
                    End With
                End With
            End With
        End With
       
' TR_VDW_296_0_32.dat
' TR_VDW_293_0_32.dat
'
'
'dist, height, view angle, dip, ref
'1153592.81106  97160.06975      0.00000    172.17649      9.29830
'1154675.39167  97351.30537      0.00000    172.34680      9.29830
'1155758.98096  97542.91798      0.00000    172.51726      9.29830
'1156843.57977  97734.90832      0.00000    172.68788      9.29830
'1157929.18893  97927.27712      0.00000    172.85866      9.29830
'1159015.80927  98120.02511      0.00000    173.02960      9.29830
'1160103.44164  98313.15304      0.00000    173.20070      9.29830
    
    ElseIf chkRefFiles_dip.Value = vbChecked And chkFit1.Value = vbChecked And chkVA.Value = vbChecked And chkHgt.Value = vbChecked Then
        'plot and fit the leveling dip as a function of ray (observer) height for a chosen temperature
        Close
        FilePath = txtRefFileDir
        
        Tfit = Val(txtFit1)
        VAngFit = Val(txtVA)
        HeightFit = Val(txtHgtFit)
        FileNameIn = FilePath & "\TR_VDW_" & Trim$(txtFit1) & "_" & Trim$(txtHgtFit) & "_32.dat"
        FilePath = txtRefFileDir
        
        If Dir(FileNameIn) = sEmpty Then
          Call MsgBox("Can't open the file:" _
                      & vbCrLf & FileNameIn _
                      & vbCrLf & "" _
                      & vbCrLf & "It doesn't seem to exit" _
                      , vbInformation, "Missing file")
          Exit Sub
          End If
          
      filein% = FreeFile
      Open FileNameIn For Input As #filein%
      j = 0
      Do Until EOF(filein%)
         Input #filein%, Dist, Hfit, VAO, Dip, ref
         If VAngFit * 60 = VAO And Hfit <= 4000 Then
            found% = 1
            ReDim Preserve HgtFit(j + 1)
            ReDim Preserve Refr(j + 1)
            HgtFit(j) = Hfit
            Refr(j) = Dip 'local dip in mrad
            PtX.Add HgtFit(j)
            PtY.Add Refr(j)
            If Dip > maxref Then maxref = Dip
            If Dip < minref Then minref = Dip
            If HgtFit(j) > maxhgt Then maxhgt = HgtFit(j)
            If HgtFit(j) < minhgt Then minhgt = HgtFit(j)
            j = j + 1
         ElseIf VAngFit * 60 <> VAO And found% = 1 Then
            'passed the selected VA, so exit loop
            Exit Do
            End If
      Loop
      Close #filein%
      
      If found% = 0 Then
         Screen.MousePointer = vbDefault
         Call MsgBox("Nothing found!" _
                     & vbCrLf & "" _
                     & vbCrLf & "Try different temperature and height" _
                     , vbInformation, "Nothing found")
         
         Exit Sub
      Else
        NNN = j - 1
        End If
        ' Find a good fit.
        degree = 6
        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
        
        '    stop_time = Timer
        '    Debug.Print Format$(stop_time - start_time, "0.0000")
        
        txtFitResults = ""
        For i = 1 To BestCoeffs.Count
            txtFitResults = txtFitResults & " " & BestCoeffs.Item(i)
        Next i
        If Len(txtFitResults) > 0 Then txtFitResults = Mid$(txtFitResults, 2)
        
        ' Display the error.
        Call ShowError(, ErrorLbl, 1)
        
        ' We have a solution.
        HasSolution = True
        
      Screen.MousePointer = vbDefault
      
        'set up chart and plot the data then fit it
        ReDim TransferCurve(0, 0) As Variant
         
        ReDim TransferCurve(0 To NNN - 1, 0 To 2) As Variant
        For j = 0 To NNN - 1
           TransferCurve(j, 0) = " " + Format(CStr(HgtFit(j + 1)), "###0.0#")
           TransferCurve(j, 1) = Refr(j + 1)
           TransferCurve(j, 2) = BestCoeffs.Item(1) + _
                                 BestCoeffs.Item(2) * HgtFit(j + 1) + _
                                 BestCoeffs.Item(3) * HgtFit(j + 1) ^ 2# + _
                                 BestCoeffs.Item(4) * HgtFit(j + 1) ^ 3# + _
                                 BestCoeffs.Item(5) * HgtFit(j + 1) ^ 4# + _
                                 BestCoeffs.Item(6) * HgtFit(j + 1) ^ 5# + _
                                 BestCoeffs.Item(7) * HgtFit(j + 1) ^ 6#
        Next j
      
        With MSChartTR
            .chartType = VtChChartType2dLine
            .Title = "Dip (degrees)"
            .RandomFill = False
            .ShowLegend = True
            .ChartData = TransferCurve
            With .Plot
                With .Wall.Brush
                    .Style = VtBrushStyleSolid
                    .FillColor.Set 255, 255, 255
                End With
                With .Axis(VtChAxisIdX)
                    .AxisTitle = "Observer's Height (m)"
                End With
                With .Axis(VtChAxisIdX).CategoryScale
                    .Auto = False
                    .DivisionsPerLabel = NNN * 0.1
                    .DivisionsPerTick = NNN * 0.1
                    .LabelTick = True
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle = "Dip (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = "Fit to Dip (mrad)"
                    With .ValueScale
                        .Auto = False
                        .Minimum = 0.9 * minref
                        .Maximum = 1.1 * maxref
                        .MajorDivision = 10
                    End With
                End With
                With .SeriesCollection(1)
                    .SeriesMarker.Show = False 'True
                    .LegendText = "Refraction"
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 0, 0, 255
                    End With
                    With .DataPoints(-1).Marker
                        .Style = VtMarkerStyleDiamond
                    End With
                End With
                With .SeriesCollection(2)
                    .LegendText = "Fit"
                    .SecondaryAxis = True
                    With .Pen
                        .Width = ScaleX(1, vbPixels, vbTwips)
                        .VtColor.Set 155, 0, 0
                    End With
                End With
            End With
        End With
        
        
    ElseIf chkRefFiles_Ref.Value = vbChecked Or chkRefFiles_dip.Value = vbChecked And chkFit1.Value = vbUnchecked And chkVA.Value = vbUnchecked Then
       
       
       Call MsgBox("You must designate a temperature and view angle" _
                   & vbCrLf & "by checking the Temp. and View Angle Checkboxes" _
                   & vbCrLf & "and picking a obs. temperature and observed view angle." _
                   , vbExclamation, "Missing designations")
       
        
    End If

   On Error GoTo 0
   Exit Sub

cmdPlotFit_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdPlotFit_Click of Form prjAtmRefMainfm"
      
End Sub

Private Sub cmdRefFiles_browse_Click()
    txtRefFileDir = BrowseForFolder(prjAtmRefMainfm.hwnd, "Choose Directory")
End Sub

Private Sub cmdRefWilson_Click()
'Hohenkerk & Sinclair Refraction
'
'This program computes atmospheric refraction on the horizon using
'the algorithm described by H&S in their 2008 technical note:
'The Calculation of Angular Atsmospheric Refraction at Large Zenith Angles
'NAO Technical Note 63, and 2008 bug fixes

'
'SUBROUTINE hmnaoref ( z0 , h0 , t0 , p0 , ups , wl , ph , as , eps , ref )
'IMPLICIT none
Dim i As Long, j As Long, k As Long, inn As Long, iss As Long, istart As Long
Dim z0 As Double, h0 As Double, t0 As Double, p0 As Double, ups As Double, wl As Double
Dim PH As Double, ass As Double, EPS As Double, ref As Double, refi As Double
Dim r As Double, hepsr As Double, wlsq  As Double, z1 As Double, psat As Double
Dim A(10) As Double, gb As Double, z0r As Double, t0c As Double

Dim pw0 As Double, r0 As Double, sk0 As Double, f0 As Double, t0O As Double, n0 As Double
Dim dndr0 As Double, rt As Double, nt As Double, TT As Double
Dim dndrt As Double, zt As Double, ft As Double, dndrts As Double, nts As Double
Dim zts As Double, fts As Double, rs As Double, ns As Double
Dim dndrs As Double, zs As Double, FS As Double, ref0 As Double, refp As Double
Dim reft As Double, fb As Double, H As Double, step As Double
Dim z As Double, rg As Double, T As Double, tg As Double, n As Double, dndr As Double
Dim F As Double, fe As Double, fo As Double, ff As Double, ex1 As Double, ex2 As Double
Dim gcr As Double, md As Double, mw As Double, gamma As Double, z2 As Double, s As Double
Dim ht As Double, hs As Double, dgr As Double, refmrad As Double
Dim Theta As Double, P As Double, pt As Double, DeltaR As Double
Dim lR As Double, LE As Double, lpsrate As Double
Dim r00 As Double, jstop As Long, tst As Double, pst As Double
Dim XR As Double, yr As Double, XR0 As Double
'Dim Theta0 As Double, ThetaTot As Double
'tt,pt are temperature and pressure at the Tropopause in the Stratosphere
'Dim numPoints As Long, FirstPhi As Boolean, Phi0 As Double

Dim StatusMes As String
Dim FNM As String
Dim NNN As Long, II As Long, AtmType As Integer, AtmNumber As Integer, KWAV As Integer, jstep As Long
Dim ih As Long, ih0 As Long, iht As Long, ihs As Long
Dim p0t As Double, ps As Double, ts As Double, Isothermic As Boolean, CONP As Double, pps As Double
Dim DA As Double, DW As Double, Index As Double, Index2 As Double, n2 As Double, IndexMenat As Double
Dim e As Double, el As Double, BETA As Double
'Dim n_size As Long, msize As Long

cmdCalc.Enabled = False
cmdRefWilson.Enabled = False
cmdMenat.Enabled = False
cmdVDW.Enabled = False
   
CalcComplete = False

RefCalcType% = 1

cmbSun.Clear

pi = 4# * Atn(1#) '3.141592654
CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
cd = pi / 180# 'conversion of degrees to radians
ROBJ = 15# 'half size of sun in minutes of arc

'   On Error GoTo cmdRefWilson_Click_Error
   
 '------------------progress bar initialization
With prjAtmRefMainfm
  '------fancy progress bar settings---------
  .progressfrm.Visible = True
  .picProgBar.AutoRedraw = True
  .picProgBar.BackColor = &H8000000B 'light grey
  .picProgBar.DrawMode = 10

  .picProgBar.FillStyle = 0
  .picProgBar.ForeColor = &H400000 'dark blue
  .picProgBar.Visible = True
End With
pbScaleWidth = 100
'-------------------------------------------------

DeltaR = 10
   
HOBS = Val(txtHeight.Text)
If HOBS = 0 Then HOBS = 0.001 'this code doesn't work for hobs = 0 so add epsiolon of height
ISSR = 0
ITRAN = 1
Image1 = 1
IPLOT = 1
STARTALT = Val(txtStartAlt.Text)
DELALT = Val(txtDelAlt.Text)
XMAX = Val(txtXmax.Text) * 1000 'convert km to meters
PPAM = Val(txtPPAM.Text)
KMIN = CInt((Val(txtKmin.Text) - 380) / 5# + 1#)
KMAX = CInt((Val(txtKmax.Text) - 380) / 5# + 1#)
KSTEP = CInt(Val(txtKStep.Text) * 0.1)
STARTAZM = 19
DELAZM = 32
If INVFLAG = 1 Then
   SINV = Val(txtSInv.Text)
   EINV = Val(txtEInv.Text)
   DTINV = Val(txtDInv.Text)
   End If
   
StatusMes = "Pixels per arcminute " & Str(PPAM) & ", Maximum height (degrees) " & Str(n / (120# * PPAM))
Call StatusMessage(StatusMes, 1, 0)

n_size = 500
msize = 20 + Val(txtNumSuns.Text) * 32 * PPAM

If Trim$(txtXSize.Text) <> sEmpty Then
   msize = Val(txtXSize.Text)
   End If
If Trim$(txtYSize.Text) <> sEmpty Then
   n_size = 2 * Val(txtYSize.Text) * PPAM * 60
   End If
   
Dim KA As Long
For KA = 1 To NumSuns
   ALT(KA) = STARTALT + CDbl(KA - 1) * DELALT
   AZM(KA) = STARTAZM + CDbl(KA - 1) * DELAZM
Next KA

Screen.MousePointer = vbHourglass

'loop in z0 according to size of sun picture just like Bruton does

'=============inputs=============
'z0 , h0 , t0 , p0 , ups , wl , ph , as , eps , ref
      z0 = 90#  'use straight ahead for test, this is observed view angle
      h0 = HOBS
      t0 = prjAtmRefMainfm.txtGroundTemp
      p0 = prjAtmRefMainfm.txtGroundPressure
      ups = prjAtmRefMainfm.txtHumid / 100
      wl = 0.574 'micrometers - average wavelength from stars
      PH = 32.1 'latitude
      ass = 0.0065 'lapse rate degrees K/m
      EPS = 0.00000001
'==============================================

If prjAtmRefMainfm.chkHSoatm.Value = vbChecked Then

    'record atmospheric model in steps of 10 meters until the stratosphere, then every 100 meters
    StatusMes = "Determine lapse rate and beginning temperature and pressure of other atmospheric models"
    Call StatusMessage(StatusMes, 1, 0)
'
'     'specify atmosphere type and the file containing the atmosphere profile
'      StatusMes = "Calculating and Storing multilayer atmospheric details"
'      Call StatusMessage(StatusMes, 1, 0)
'
     If OptionLayer.Value = True Then
        AtmType = 1
        FNM = App.Path & "\stmod1.dat"
     ElseIf OptionRead.Value = True Then
        AtmType = 1
        FNM = TextExternal.Text
     ElseIf OptionSelby.Value = True Then
        AtmType = 2
        If prjAtmRefMainfm.opt1.Value = True Then
           AtmNumber = 1
        ElseIf prjAtmRefMainfm.opt2.Value = True Then
           AtmNumber = 2
        ElseIf prjAtmRefMainfm.opt3.Value = True Then
           AtmNumber = 3
        ElseIf prjAtmRefMainfm.opt4.Value = True Then
           AtmNumber = 4
        ElseIf prjAtmRefMainfm.opt5.Value = True Then
           AtmNumber = 5
        ElseIf prjAtmRefMainfm.opt6.Value = True Then
           AtmNumber = 6
        ElseIf prjAtmRefMainfm.opt7.Value = True Then
           AtmNumber = 7
        ElseIf prjAtmRefMainfm.opt8.Value = True Then
           AtmNumber = 8
        ElseIf prjAtmRefMainfm.opt9.Value = True Then
           AtmNumber = 9
        ElseIf prjAtmRefMainfm.opt10.Value = True Then
           AtmNumber = 10
           FNM = txtOther.Text
           End If
        End If

'     FNM = App.Path & "\stmod1.dat" 'stmod2.dat.txt"
     ier = LoadAtmospheres(FNM, AtmType, AtmNumber, lpsrate, tst, pst, NNN, 3)
     
     t0 = tst
     p0 = pst
     ass = lpsrate
     If chkLapse.Value = vbChecked Then 'use the inputed value instead of the above
        ass = Val(txtLapse.Text) * 0.001 'convert to deg K/m
        End If
'
'     NumTemp = NNN + 1 'number of layers are NumTemp - 1
'
     If ier < 0 Then
        Close
        prjAtmRefMainfm.cmdBrowse(0).Enabled = True
        Screen.MousePointer = vbDefault
        Close
        cmdVDW.Enabled = True
        cmdCalc.Enabled = True
        cmdMenat.Enabled = True
        cmdRefWilson.Enabled = True
        Exit Sub
        End If

   End If

'C - - - - - Constants of the atmosphere .
gcr = 8314.32 'Universal Gas Constant
md = 28.9644 'mol. weight of dry air
mw = 18.0152 'mol weight of water vapor)
gamma = 18.36 'Delta in HS, exponent of temperature dependence of water vapor pressure
z2 = 0.0000112684 'multiplicative factor in second term in expression of n
'C - - - - - Equatorial radius of the Earth .
s = 6371000#  '6378136.6 'use average radius, not radius at equator
RE = s
'C - - - - - Height of tropopause and extent of atmosphere in metres .
ht = 11000#
hs = 80000#
'C - - - - - Degrees to radians .
dgr = 1.74532925199433E-02
'C - - - - - Integration limits in radians .
hepsr = 0.5 * EPS / 3600#
'C - - - - - Convert ZD to radians
'z0r = z0 * dgr

'C - - - - - Set up parameters defined at the observer for the atmosphere .
'C
If prjAtmRefMainfm.chkHSoatm.Value = vbChecked Then
  GoSub LayeredAtmospheres
  End If
  
KMIN = (Val(prjAtmRefMainfm.txtKmin.Text) + 1) / 5# - 380#
KMAX = (Val(prjAtmRefMainfm.txtKmax.Text) + 1) / 5# - 380#
KSTEP = Val(prjAtmRefMainfm.txtKStep.Text) / 10#
  
For KWAV = KMIN To KMAX Step KSTEP   '<1

   wl = 380# + CDbl(KWAV - 1) * 5#
   wl = wl / 1000# 'convert to nm

Gravity:
gb = 9.806248 * (1# - 0.0026442 * Cos(2# * PH * dgr))
gb = gb - 0.000003086 * h0

'C - - - - - Factor for optical wavelengths .
wlsq = wl * wl
z1 = (287.6155 + (1.62887 + 0.0136 / wlsq) / wlsq)
z1 = z1 * 0.00027315 / 1013.25 'defined as A in HS
A(1) = Abs(ass) 'lapse rate
If chkLapse Then A(1) = ass
A(2) = (gb * md) / gcr 'gamma in HS
A(3) = A(2) / A(1)
A(4) = gamma ' delta in HS
'C - - - - - Wet air :
t0c = t0 - 273.15 'Degrees Celsius
ex1 = (0.7859 + 0.03477 * t0c) / (1# + 0.00412 * t0c)
ex2 = (1# + p0 * (0.0000045 + 0.0000000006 * t0c * t0c))
psat = 10# ^ (ex1 * ex2)
pw0 = ups * psat / (1# - (1.1 - ups) * psat / p0)
A(5) = pw0 * (1# - mw / md) * A(3) / (A(4) - A(3))
A(6) = p0 + A(5) 'full coeficient for pressure of air with humidity ups
A(7) = z1 * A(6) / t0 'humid air pressure contribution to n
A(8) = (z1 * A(5) + z2 * pw0) / t0 'water vapor contribution to n
A(9) = (A(3) - 1#) * A(1) * A(7) / t0 'first coeficient in expression for dndr
A(10) = (A(4) - 1#) * A(1) * A(8) / t0 'second coeficient in expression for dndr

'C - - - - - At the observer .
r0 = s + h0
Call atmostro(r0, t0, A(), r0, t0O, n0, dndr0, P)
If chkCiddor.Value = vbChecked Then
   Call INDEX_CIDDOR(wl * 1000, t0O, P * 100, ups, 402, DA, DW, Index)
   'now move up by 100 meters to calculate dn/dr
   Call atmostro(r0, t0, A(), r0 + DeltaR, t0O, n2, dndr0, P)
   dndr0 = (n2 - n0) / DeltaR
   Call INDEX_CIDDOR(wl * 1000, t0O, P * 100, ups, 402, DA, DW, Index2)
   dndr0 = (Index2 - Index) / DeltaR
   n0 = Index
   End If
'
sk0 = n0 * r0 * Sin(z0r)
f0 = refii(r0, n0, dndr0)
'
''C - - - - - At the Tropopause in the Troposphere .
rt = s + ht
Call atmostro(r0, t0, A(), rt, TT, nt, dndrt, pt)
If chkCiddor.Value = vbChecked Then
   Call INDEX_CIDDOR(wl * 1000, TT, pt * 100, ups, 402, DA, DW, Index)
   'now move up by 100 meters to calculate dn/dr
   Call atmostro(r0, t0, A(), rt + DeltaR, TT, n2, dndrt, P)
   Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index2)
   dndrt = (Index2 - Index) / DeltaR
   nt = Index
   End If


zt = DASIN(sk0 / (rt * nt)) / dgr 'convert back to degrees
ft = refii(rt, nt, dndrt)
'
''C - - - - - At the Tropopause in the Stratosphere .
Call atmosstr(rt, TT, nt, A(2), rt, nts, dndrts, pt, P)
If chkCiddor.Value = vbChecked Then
   Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index)
   'now move up by 100 meters to calculate dn/dr
   Call atmosstr(rt, TT, nt, A(2), rt + DeltaR, n2, dndrts, pt, P)
   Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index2)
   dndrts = (Index2 - Index) / DeltaR
   nts = Index
   End If

zts = DASIN(sk0 / (rt * nts)) / dgr 'convert back to degrees
fts = refii(rt, nts, dndrts)
'
''C - - - - - At the stratosphere limit
rs = s + hs
Call atmosstr(rt, TT, nt, A(2), rs, ns, dndrs, pt, P)
If chkCiddor.Value = vbChecked Then
   Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index)
   'now move up by 100 meters to calculate dn/dr
   Call atmosstr(rt, TT, nt, A(2), rs + DeltaR, n2, dndrs, pt, P)
   Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index2)
   dndrs = (Index2 - Index) / DeltaR
   ns = Index
   End If

zs = DASIN(sk0 / (rs * ns)) / dgr 'convert back to degrees
FS = refii(rs, ns, dndrs)

'record atmospheric model in steps of 10 meters until the stratosphere, then every 100 meters
StatusMes = "Recording atmospheric model temperature, pressure, and index of refraction as a function of elevation"
Call StatusMessage(StatusMes, 1, 0)

'increase resolution of temperature-pressure atmospheric model for plotting
numpoints = Int((rt - 25 - s) / 25 + (rs - rt - 200) / 200) + 3
NNN = 0
For jr = s To rt - 25 Step 25
   Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * NNN / numpoints))
   Call atmostro(r0, t0, A(), CDbl(jr), T, n, dndr, P)
   ELV(NNN) = CDbl(jr) - s
   TMP(NNN) = T
   PRSR(NNN) = P
   IndexRefraction(NNN) = n
   If chkCiddor.Value = vbChecked Then
      Call INDEX_CIDDOR(wl * 1000, T, P * 100, ups, 402, DA, DW, Index)
      IndexRefraction(NNN) = Index
   ElseIf chkMenat.Value = vbChecked Then
      IndexMenat = 1 + 0.000001 * (77.46 + 0.459 * (1# / wl) ^ 2) * (PRSR(NNN) / TMP(NNN))
      IndexRefraction(NNN) = IndexMenat
      End If
   NNN = NNN + 1
Next jr
   ELV(NNN) = rt - s
   TMP(NNN) = TT
   PRSR(NNN) = pt
   IndexRefraction(NNN) = nts
   If chkCiddor.Value = vbChecked Then
      Call INDEX_CIDDOR(wl * 1000, TT, pt * 100, ups, 402, DA, DW, Index)
      IndexRefraction(NNN) = Index
   ElseIf chkMenat.Value = vbChecked Then
      IndexMenat = 1 + 0.000001 * (77.46 + 0.459 * (1# / wl) ^ 2) * (PRSR(NNN) / TMP(NNN))
      IndexRefraction(NNN) = IndexMenat
      End If

   NNN = NNN + 1
For jr = rt + 200 To rs Step 200
   Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * NNN / numpoints))
   Call atmosstr(rt, TT, nt, A(2), CDbl(jr), n, dndrt, pt, P)
   ELV(NNN) = CDbl(jr) - s
   TMP(NNN) = TT
   PRSR(NNN) = P
   IndexRefraction(NNN) = n
   If chkCiddor.Value = vbChecked Then
      Call INDEX_CIDDOR(wl * 1000, T, P * 100, ups, 402, DA, DW, Index)
      IndexRefraction(NNN) = Index
   ElseIf chkMenat.Value = vbChecked Then
      IndexMenat = 1 + 0.000001 * (77.46 + 0.459 * (1# / wl) ^ 2) * (PRSR(NNN) / TMP(NNN))
      IndexRefraction(NNN) = IndexMenat
      End If
   NNN = NNN + 1
Next jr
NNN = NNN - 1

NumTemp = NNN + 1
    
        
myfile$ = Dir(App.Path & "\test_HS.dat")
If myfile$ <> sEmpty Then
   Kill App.Path & "\test_HS.dat"
   End If
myfile$ = Dir(App.Path & "\tc_HS.dat")
If myfile$ <> sEmpty Then
   Kill App.Path & "\tc_HS.dat"
   End If

Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset

'C - - - - - Integrate the refraction integral in the troposphere and
'------------stratosphere . ie Ref = Ref troposphere + Ref stratosphere .
'------------do this by taking steps in z (zenith angle) defined w.r.t observer at r
'------------and finding the corresponding r to that zenith angle.
'------------Accumulate the refraction until r is equal to height of atmosphere

'C - - - - - Initial step lengths etc .

StatusMes = "Calculating refraction as a function of view angle for wavelength " & Str(wl * 1000) & " nm"
Call StatusMessage(StatusMes, 1, 0)

   For jstep = 1 To n_size + 1 '<2

        If (IPLOT = 1) Then
        '            PRINT *, J
        '            StatusMes = Str(J)
        '            Call StatusMessage(StatusMes, 1, 0)
           Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * jstep / n_size))
           End If
        '         II = IIS
        '         IDCT(j) = 0
        '         AIRM(j) = 0#
        '         SSR(k, j) = 0#
        ALFA(KWAV, jstep) = (CDbl(n_size / 2 - (jstep - 1)) / PPAM)
'        If ALFA(KWAV, jstep) = 0 Then
'           ccc = 1
'           End If
        z0 = 90# - ALFA(KWAV, jstep) / 60#
        z0r = z0 * dgr
        
        'C - - - - - At the observer .
        r0 = s + h0
        Call atmostro(r0, t0, A(), r0, t0O, n0, dndr0, P)
        If chkCiddor.Value = vbChecked Then
           Call INDEX_CIDDOR(wl * 1000, t0O, P * 100, ups, 402, DA, DW, Index)
           'now move up by 100 meters to calculate dn/dr
           Call atmostro(r0, t0, A(), r0 + DeltaR, t0O, n2, dndr0, P)
           Call INDEX_CIDDOR(wl * 1000, t0O, P * 100, ups, 402, DA, DW, Index2)
           dndr0 = (Index2 - Index) / DeltaR
           n0 = Index
           End If
        sk0 = n0 * r0 * Sin(z0r)
        f0 = refii(r0, n0, dndr0)
        
        'C - - - - - At the Tropopause in the Troposphere .
        rt = s + ht
        Call atmostro(r0, t0, A(), rt, TT, nt, dndrt, pt)
        If chkCiddor.Value = vbChecked Then
           Call INDEX_CIDDOR(wl * 1000, TT, pt * 100, ups, 402, DA, DW, Index)
           'now move up by 100 meters to calculate dn/dr
           Call atmostro(r0, t0, A(), rt + DeltaR, TT, n2, dndrt, P)
           Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index2)
           dndrt = (Index2 - Index) / DeltaR
           nt = Index
           End If
        zt = DASIN(sk0 / (rt * nt)) / dgr 'convert back to degrees
        ft = refii(rt, nt, dndrt)
        
        'C - - - - - At the Tropopause in the Stratosphere .
        Call atmosstr(rt, TT, nt, A(2), rt, nts, dndrts, pt, P)
        If chkCiddor.Value = vbChecked Then
           Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index)
           'now move up by 100 meters to calculate dn/dr
           Call atmosstr(rt, TT, nt, A(2), rt + DeltaR, n2, dndrts, pt, P)
           Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index2)
           dndrts = (Index2 - Index) / DeltaR
           nts = Index
           End If
        zts = DASIN(sk0 / (rt * nts)) / dgr 'convert back to degrees
        fts = refii(rt, nts, dndrts)
        
        'C - - - - - At the stratosphere limit
        rs = s + hs
        Call atmosstr(rt, TT, nt, A(2), rs, ns, dndrs, pt, P)
        If chkCiddor.Value = vbChecked Then
           Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index)
           'now move up by 100 meters to calculate dn/dr
           Call atmosstr(rt, TT, nt, A(2), rs + DeltaR, n2, dndrs, pt, P)
           Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index2)
           dndrs = (Index2 - Index) / DeltaR
           ns = Index
           End If
        zs = DASIN(sk0 / (rs * ns)) / dgr 'convert back to degrees
        FS = refii(rs, ns, dndrs)

'ThetaTot = 0#
LE = 0#
lR = 0#
r00 = r0

ref0 = -999.999
iss = 16
For k = 1 To 2
    istart = 0
    fe = 0#
    fo = 0#
    
    If (k = 1) Then
        H = (zt - z0r / dgr) / CDbl(iss) 'iniital step size in zenith angle -- within the troposphere
        fb = f0
        ff = ft
    ElseIf (k = 2) Then
        H = (zs - zts) / CDbl(iss) 'initial step size in zenith angle --within the starosphere
        fb = fts
        ff = FS
        End If
    
    inn = iss - 1
    iss = iss / 2
    step = H
'100 CONTINUE
100:
    For i = 1 To inn
    
        If (i = 1 And k = 1) Then
           z = z0r / dgr + H
           r = r0
            
'           Theta0 = 90 - z0 'starting cylindrical angle defined from vertical y axis
           Close
           fileout% = FreeFile
           Open App.Path & "\testHS_t.dat" For Output As #fileout%
           lR = 0
           LE = 0
           XR = 0
           yr = 0
           XR0 = 0
           Write #fileout%, 0#, r0 - s, ALFA(KWAV, jstep), ALFA(KWAV, jstep), 0 'observer's X,Y position - defined as the Y axis, Y = r0
           
        ElseIf (i = 1 And k = 2) Then
            z = zts + H
            r = rt
            XR = XR0
            Close
            fileout% = FreeFile
            Open App.Path & "\testHS_s.dat" For Output As #fileout%
        Else
            z = z + step
            End If

'C - - - - - Given the zenith distance ( z ) find r by Newton - Raphson iteration (Magnume eqn 31).
        rg = r
        r00 = rg
        For j = 1 To 4
            If (k = 1) Then
                Call atmostro(r0, t0, A(), rg, tg, n, dndr, P)
                
                If chkCiddor.Value = vbChecked Then
                   Call INDEX_CIDDOR(wl * 1000, tg, P * 100, ups, 402, DA, DW, Index)
                   'now move up by 100 meters to calculate dn/dr
                   Call atmostro(r0, t0, A(), rg + DeltaR, tg, n2, dndr, P)
                   Call INDEX_CIDDOR(wl * 1000, tg, P * 100, ups, 402, DA, DW, Index2)
                   dndr = (Index2 - Index) / DeltaR
                   n = Index
                   End If
                
                
            ElseIf (k = 2) Then
                Call atmosstr(rt, TT, nt, A(2), rg, n, dndr, pt, P)
                
                If chkCiddor.Value = vbChecked Then
                   Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index)
                   'now move up by 100 meters to calculate dn/dr
                   Call atmosstr(rt, TT, nt, A(2), rg + DeltaR, n2, dndr, pt, P)
                   Call INDEX_CIDDOR(wl * 1000, T, P * 100, ups, 402, DA, DW, Index2)
                   dndr = (Index2 - Index) / DeltaR
                   n = Index
                   End If
                
                End If
            rg = rg - ((rg * n - sk0 / Sin(z * dgr)) / (n + rg * dndr))
        Next j
        r = rg
        
'C - - - - - Find refractive index and integrand at r .

        If (k = 1) Then
            Call atmostro(r0, t0, A(), r, T, n, dndr, P)
            
            If chkCiddor.Value = vbChecked Then
               Call INDEX_CIDDOR(wl * 1000, T, P * 100, ups, 402, DA, DW, Index)
               'now move up by 100 meters to calculate dn/dr
               Call atmostro(r0, t0, A(), r + DeltaR, T, n2, dndr, P)
               Call INDEX_CIDDOR(wl * 1000, T, P * 100, ups, 402, DA, DW, Index2)
               dndr = (Index2 - Index) / DeltaR
               n = Index
               End If
            
            pt = P
            
        ElseIf (k = 2) Then
            Call atmosstr(rt, TT, nt, A(2), r, n, dndr, pt, P)
            
            If chkCiddor.Value = vbChecked Then
               Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index)
               Call atmosstr(rt, TT, nt, A(2), r + DeltaR, n, dndr, pt, P)
               Call INDEX_CIDDOR(wl * 1000, TT, P * 100, ups, 402, DA, DW, Index2)
               dndr = (Index2 - Index) / DeltaR
               n = Index
               End If
            
            T = TT
            
            End If
            
            
        'calculate the distance taveled, and cylindrical coordinates of the ray
        'need to follow Menat's procedure, i.e., calculate the approximate path length using law of cosines and then
        'calculate the cylindrical angle subtended using the laws of sines.
        e = 90# - z 'view angle for this increment
        el = SQT(e, r - r00, r00) 'approximate path length in this increment
        lR = el * Cos(e * cd) 'path along the last increment's radius
        BETA = DASIN(lR / r) 'angle subtended on circumference of Earth by last increment in r in radians
'        BTOT = BTOT + BETA  'total angle subtended on circumference of Earth until now in radians
        LE = BETA * s
        XR = XR + LE 'totalt path along Earth's circumference until now
'        ANGLE = XP / RE
        yr = r - s
'        If ccc = 1 And iss = 256 And k = 1 Then
'           cccc = 1
'        ElseIf ccc = 1 And iss = 1024 And k = 2 Then
'           ccccc = 1
'           End If
        If k = 1 Then XR0 = XR
'        A1 = (XP * Cos(ANGLE) + (RE + YP) * Sin(ANGLE))
'        A2 = (-XP * Sin(ANGLE) + (RE + YP) * Cos(ANGLE))
              
'           Theta = (90# - z) 'current cylindrical angle of ray
'           ThetaTot = Theta - Theta0 'total difference in cyclindrical angle that ray traveled as it propogated through the atmosphere
'           'LE is distance along circumference of Earth, LR is path length
'           LE = r00 * ThetaTot * dgr
'           lR = lR + Sqr(LE * LE + (r - r00) * (r - r00)) 'use Euclidean geometry (Pythagorean Theorem) approx. to non-spherical curved surface distance for small r differences
'                                                          'better approximation is to approximate the light path as a portion of a great circle
'                                                          'and find the distance
'           'determine the X,Y of the endpoints
'           XR = r * Sin(ThetaTot * dgr)
'           YR = r * Cos(ThetaTot * dgr)
          
           'check that ray doesn't intersect with the earth, if it does, then signal that with YR = -1000
           If r < s Then 'collided with the surface
              Write #fileout%, XR, -1000, ALFA(KWAV, jstep), ALFA(KWAV, jstep), 0
              jstop = jstep
           Else
              Write #fileout%, XR, yr, ALFA(KWAV, jstep), ALFA(KWAV, jstep), 0
              jstop = -1
              End If
               
        F = refii(r, n, dndr)
        If (istart = 0 And i Mod 2 = 0) Then
            fe = fe + F
        Else
            fo = fo + F
            End If
            
        If jstop <> -1 Then Exit For
    Next i
    
    If jstop <> -1 Then Exit For

'C - - - - - Evaluate the integrand using Simpsons Rule .
    refp = H * (fb + 4# * fo + 2# * fe + ff) / 3#
'C - - - - - Test for convergence .
    If (Abs(refp - ref0) > hepsr) Then
        iss = 2 * iss
        inn = iss
        step = H
        H = H / 2#
        fe = fe + fo
        fo = 0#
        ref0 = refp
        If (istart = 0) Then istart = 1
        GoTo 100
        End If
        
    If (k = 1) Then reft = refp

Next k
Close #fileout%

If jstop <> -1 Then
   Close
   GoTo crw500
   End If

'append files
'if first append file

'FileCopy App.Path & "\testHS_t.dat", App.Path & "\testHS.dat"

filein% = FreeFile
Open App.Path & "\testHS_t.dat" For Input As #filein%

fileout% = FreeFile
Open App.Path & "\test_HS.dat" For Append As #fileout%

Do Until EOF(filein%)
   Line Input #filein%, doclin$
   Print #fileout%, doclin$
Loop
Close #filein%
'Close #fileout%

filein% = FreeFile
Open App.Path & "\testHS_s.dat" For Input As #filein%

'fileout% = FreeFile
'Open App.Path & "\test_HS.dat" For Append As #fileout%

Do Until EOF(filein%)
   Line Input #filein%, doclin$
   Print #fileout%, doclin$
Loop
Close #filein%
Close #fileout%

Kill App.Path & "\testHS_t.dat"
Kill App.Path & "\testHS_s.dat"

'C - - - - - Refraction in the troposphere + stratosphere in degrees .
ref = (reft + refp)  'degrees
refmrad = ref * dgr * 1000 'refraction in mrad

'now append the refraction value to the transfer file
ALFT(KWAV, jstep) = ALFA(KWAV, jstep) - ref * 60# 'true depression angle
'fileout% = FreeFile
'Open App.Path & "\tc_HS.dat" For Append As #fileout%
'Print #fileout%, ALFA(KWAV, jstep), ALFT(KWAV, jstep) 'units in minutes of arc
'Close #fileout%

If z0 = 90# Then
   prjAtmRefMainfm.lblRef.Caption = "Atms. refraction (deg.) = " & ref & vbCrLf & "Atms. refraction (mrad) = " & refmrad
   End If
   
crw500:
If jstop <> -1 Then
   Close
   
   StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KWAV, jstop - 1)) & vbCrLf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(s / (s + h0)) / CONV))
     Call StatusMessage(StatusMes, 1, 0)
'     ier = MsgBox(StatusMes, vbInformation + vbOKOnly, "Horizon")
     lblHorizon.Caption = StatusMes

   Exit For
   End If

Next jstep

Next KWAV

CalcComplete = True

'Screen.MousePointer = vbDefault
If IPLOT = 1 Then
   prjAtmRefMainfm.progressfrm.Visible = False
   Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
   End If

'NNN = numpoints + numNewPoints - 1

'record atmosphere as "HS-atmosphere.txt"
fileout% = FreeFile
Open App.Path & "\HS-atmosphere.txt" For Output As #fileout%
For j = 0 To NNN
   Print #fileout%, ELV(j), TMP(j), PRSR(j), IndexRefraction(j)
   
    If j = 0 Then
       MinTemp = TMP(0)
       MaxTemp = MinTemp
    Else
       If TMP(j) > MaxTemp Then MaxTemp = TMP(j)
       If TMP(j) < MinTemp Then MinTemp = TMP(j)
       End If
   
Next j
Close #fileout%

''record atmosphere as "HS-atmosphere.dat"
'filein% = FreeFile
'Open App.Path & "\testHS.dat" For Input As #filein%
'fileout% = FreeFile
'Open App.Path & "\HS-atmosphere.txt" For Output As #fileout%
'NNN = 0
'Do Until EOF(filein%)
'   Input #filein%, aaa, bbb, Ccc, Ddd, IndexRefraction(NNN)
'   Print #fileout%, ELV(NNN), TMP(NNN), PRSR(NNN), IndexRefraction(NNN)
'   NNN = NNN + 1
'Loop
'Close #filein%
'Close #fileout%

'load temperature and pressure charts

 ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
 
 For j = 1 To NNN
    TransferCurve(j, 1) = " " & CStr(ELV(j - 1) * 0.001)
'         TransferCurve(J, 2) = ELV(J - 1) * 0.001
    TransferCurve(j, 2) = TMP(j - 1)
 Next j
 
 With MSChartTemp
   .chartType = VtChChartType2dLine
   .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN - 1
'        .RowLabel = "Height (km)"
'        .ColumnLabel = "Temperature (Kelvin)"
   .ChartData = TransferCurve
 End With

 For j = 1 To NNN
    TransferCurve(j, 2) = PRSR(j - 1)
 Next j
 
 With MSChartPress
  .chartType = VtChChartType2dLine
  .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN - 1
'        .RowLabel = "Height (km)"
'        .ColumnLabel = "Pressure (Kelvin)"
  .ChartData = TransferCurve
End With

StatusMes = "Writing transfer curve."
Call StatusMessage(StatusMes, 1, 0)
filnum% = FreeFile
Open App.Path & "\tc_HS.dat" For Output As #filnum%
'      WRITE (20,*) N
NumTc = 0
Print #filnum%, n_size
For j = 1 To jstop - 1
'        WRITE(20,1) ALFA(KMIN,J),ALFT(KMIN,J)
    Print #filnum%, ALFA(KMIN, j), ALFT(KMIN, j)
    If ALFA(KMIN, j) = 0 Then 'display the refraction value for the zero view angle ray
       prjAtmRefMainfm.lblRef.Caption = "Atms. refraction (deg.) = " & Abs(ALFT(KMIN, j)) / 60# & vbCrLf & "Atms. refraction (mrad) = " & Abs(ALFT(KMIN, j)) * 1000# * cd / 60#
       End If
'store all view angles that contribute to sun's orb
    NumTc = NumTc + 1
    For KA = 1 To NumSuns
       y = ALFT(KMIN, j) - ALT(KA)
       If Abs(y) <= ROBJ Then
          'only accept rays that pass over the horizon (ALFT(KMIN, J) <> -1000) and are within the solar disk
          SunAngles(KA - 1, NumSunAlt(KA - 1)) = j
          NumSunAlt(KA - 1) = NumSunAlt(KA - 1) + 1
          End If
    Next KA
Next j
Close #filnum%


'now load up transfercurve array for plotting
ReDim TransferCurve(1 To NumTc, 1 To 2) As Variant

For j = 1 To NumTc
 TransferCurve(j, 1) = " " & CStr(ALFA(KMIN, j))
 TransferCurve(j, 2) = ALFT(KMIN, j)
'         TransferCurve(J, 1) = " " & CStr(ALFT(KMIN, J))
'         TransferCurve(J, 2) = ALFA(KMIN, J)
Next j

With MSCharttc
.chartType = VtChChartType2dLine
.RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN
'        .RowLabel = "True angle (min)"
'        .ColumnLabel = "View angle (min)"
.ChartData = TransferCurve
End With

Screen.MousePointer = vbDefault

cmdRefWilson.Enabled = True

cmdCalc.Enabled = True
cmdRefWilson.Enabled = True
cmdMenat.Enabled = True
cmdVDW.Enabled = True
   
 StatusMes = "Ray tracing calculation complete"
 Call StatusMessage(StatusMes, 1, 0)
 
 
 StatusMes = "Drawing the rays on the sky simulation, please wait...."
 Call StatusMessage(StatusMes, 1, 0)
 'load angle combo boxes
'    AtmRefPicSunfm.WindowState = vbMinimized
'    BrutonAtmReffm.WindowState = vbMaximized
 'set size of picref by size of earth
 Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
 prjAtmRefMainfm.cmbSun.Clear
 prjAtmRefMainfm.cmbAlt.Clear
 For i = 1 To NumSuns
    If NumSunAlt(i - 1) > 0 Then prjAtmRefMainfm.cmbSun.AddItem i
 Next i
 
 prjAtmRefMainfm.TabRef.Tab = 4
 DoEvents

cmbSun.ListIndex = 0

   On Error GoTo 0
   Exit Sub
   
'///////////////////extend HS code for layered atmosphers //////////////////////////////////////////////////////
   
LayeredAtmospheres:

''integrate through each layer of the layer atmosphers
'Dim jb As Long
'
''record atmospheric model in steps of 10 meters until the stratosphere, then every 100 meters
'StatusMes = "Recording atmospheric model temperature, pressure, and index of refraction as a function of elevation"
'Call StatusMessage(StatusMes, 1, 0)
'
'     'specify atmosphere type and the file containing the atmosphere profile
'      StatusMes = "Calculating and Storing multilayer atmospheric details"
'      Call StatusMessage(StatusMes, 1, 0)
'
'     If OptionLayer.Value = True Then
'        AtmType = 1
'        FNM = App.Path & "\stmod1.dat"
'     ElseIf OptionRead.Value = True Then
'        AtmType = 1
'        FNM = TextExternal.Text
'     ElseIf OptionSelby.Value = True Then
'        AtmType = 2
'        If prjAtmRefMainfm.opt1.Value = True Then
'           AtmNumber = 1
'        ElseIf prjAtmRefMainfm.opt2.Value = True Then
'           AtmNumber = 2
'        ElseIf prjAtmRefMainfm.opt3(0).Value = True Then
'           AtmNumber = 3
'        ElseIf prjAtmRefMainfm.opt4(1).Value = True Then
'           AtmNumber = 4
'        ElseIf prjAtmRefMainfm.opt5.Value = True Then
'           AtmNumber = 5
'        ElseIf prjAtmRefMainfm.opt6.Value = True Then
'           AtmNumber = 6
'        ElseIf prjAtmRefMainfm.opt7.Value = True Then
'           AtmNumber = 7
'        ElseIf prjAtmRefMainfm.opt8.Value = True Then
'           AtmNumber = 8
'        ElseIf prjAtmRefMainfm.opt9.Value = True Then
'           AtmNumber = 9
'        ElseIf prjAtmRefMainfm.opt10.Value = True Then
'           AtmNumber = 10
'           FNM = txtOther.Text
'           End If
'        End If
'
''     FNM = App.Path & "\stmod1.dat" 'stmod2.dat.txt"
'     ier = LoadAtmospheres(FNM, AtmType, AtmNumber, NNN, 2)
'
'     NumTemp = NNN + 1 'number of layers are NumTemp - 1
'
'     If ier < 0 Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'        End If
'
'
'     Dim flayers() As Double 'refraction values at layer boundaries
'     ReDim flayers(NumTemp - 1)
'
'KMIN = Val(prjAtmRefMainfm.txtKmin.Text)
'KMAX = Val(prjAtmRefMainfm.txtKmax.Text)
'KSTEP = Val(prjAtmRefMainfm.txtKStep.Text)
'
'For KWAV = KMIN To KMAX Step KSTEP   '<1
'
'   wl = 380# + CDbl(KWAV - 1) * 5# 'convert to nm
'   wl = wl / 1000#
'
''Gravity:
'gb = 9.806248 * (1# - 0.0026442 * Cos(2# * ph * dgr))
'gb = gb - 0.000003086 * h0
'
''C - - - - - Factor for optical wavelengths .
'wlsq = wl * wl
'z1 = (287.6155 + (1.62887 + 0.0136 / wlsq) / wlsq)
'z1 = z1 * 0.00027315 / 1013.25 'defined as A in HS
'a(1) = Abs(ass) 'lapse rate
'a(2) = (gb * md) / gcr 'gamma in HS
'a(3) = a(2) / a(1)
'a(4) = gamma ' delta in HS
''C - - - - - Wet air :
't0c = t0 - 273.15 'Degrees Celsius
'ex1 = (0.7859 + 0.03477 * t0c) / (1# + 0.00412 * t0c)
'ex2 = (1# + p0 * (0.0000045 + 0.0000000006 * t0c * t0c))
'psat = 10# ^ (ex1 * ex2)
'pw0 = ups * psat / (1# - (1.1 - ups) * psat / p0)
'a(5) = pw0 * (1# - mw / md) * a(3) / (a(4) - a(3))
'a(6) = p0 + a(5) 'full coeficient for pressure of air with humidity ups
'a(7) = z1 * a(6) / t0 'humid air pressure contribution to n
'a(8) = (z1 * a(5) + z2 * pw0) / t0 'water vapor contribution to n
'a(9) = (a(3) - 1#) * a(1) * a(7) / t0 'first coeficient in expression for dndr
'a(10) = (a(4) - 1#) * a(1) * a(8) / t0 'second coeficient in expression for dndr
'
''C - - - - - At the observer .
'r0 = s + h0
'Call atmostro(r0, t0, a(), r0, t0O, n0, dndr0, P)
'
'sk0 = n0 * r0 * Sin(z0r)
'f0 = refii(r0, n0, dndr0)
''
'''C - - - - - At the Tropopause in the Troposphere .
'rt = s + ht
'Call atmostro(r0, t0, a(), rt, tt, nt, dndrt, pt)
'zt = DASIN(sk0 / (rt * nt)) / dgr 'convert back to degrees
'ft = refii(rt, nt, dndrt)
'
''C - - - - - At the Tropopause in the Stratosphere .
'Call atmosstr(rt, tt, nt, a(2), rt, nts, dndrts, pt, P)
'zts = DASIN(sk0 / (rt * nts)) / dgr 'convert back to degrees
'fts = refii(rt, nts, dndrts)
'
''C - - - - - At the stratosphere limit
'rs = s + hs
'Call atmosstr(rt, tt, nt, a(2), rs, ns, dndrs, pt, P)
'zs = DASIN(sk0 / (rs * ns)) / dgr 'convert back to degrees
'fs = refii(rs, ns, dndrs)
'
''record atmospheric model in steps of 10 meters until the stratosphere, then every 100 meters
'StatusMes = "Recording atmospheric model temperature, pressure, and index of refraction as a function of elevation"
'Call StatusMessage(StatusMes, 1, 0)
'
''use Ciddor values for the refraction index for its range of applicability.  Afterwards, HS expression for the stratosphere
''to calculate dn/dr calculate n values for
'
'If prjAtmRefMainfm.chkHSoatm.Value = vbChecked Then
'
''     'specify atmosphere type and the file containing the atmosphere profile
'      StatusMes = "Calculating and Storing multilayer atmospheric details"
'      Call StatusMessage(StatusMes, 1, 0)
'
'     If OptionLayer.Value = True Then
'        AtmType = 1
'        FNM = App.Path & "\stmod1.dat"
'     ElseIf OptionRead.Value = True Then
'        AtmType = 1
'        FNM = TextExternal.Text
'     ElseIf OptionSelby.Value = True Then
'        AtmType = 2
'        If prjAtmRefMainfm.opt1.Value = True Then
'           AtmNumber = 1
'        ElseIf prjAtmRefMainfm.opt2.Value = True Then
'           AtmNumber = 2
'        ElseIf prjAtmRefMainfm.opt3(0).Value = True Then
'           AtmNumber = 3
'        ElseIf prjAtmRefMainfm.opt4(1).Value = True Then
'           AtmNumber = 4
'        ElseIf prjAtmRefMainfm.opt5.Value = True Then
'           AtmNumber = 5
'        ElseIf prjAtmRefMainfm.opt6.Value = True Then
'           AtmNumber = 6
'        ElseIf prjAtmRefMainfm.opt7.Value = True Then
'           AtmNumber = 7
'        ElseIf prjAtmRefMainfm.opt8.Value = True Then
'           AtmNumber = 8
'        ElseIf prjAtmRefMainfm.opt9.Value = True Then
'           AtmNumber = 9
'        ElseIf prjAtmRefMainfm.opt10.Value = True Then
'           AtmNumber = 10
'           FNM = txtOther.Text
'           End If
'        End If
'
''     FNM = App.Path & "\stmod1.dat" 'stmod2.dat.txt"
'     ier = LoadAtmospheres(FNM, AtmType, AtmNumber, NNN, 2)
'
'     NumTemp = NNN + 1
'
'     If ier < 0 Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'        End If
'
'Else
'
'    numpoints = Int((rt - 25 - s) / 25 + (rs - rt - 200) / 200) + 3
'    NNN = 0
'    For jr = s To rt - 25 Step 25
'       Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, CLng(100# * NNN / numpoints))
'       Call atmostro(r0, t0, a(), CDbl(jr), t, n, dndr, P)
'       ELV(NNN) = CDbl(jr) - s
'       TMP(NNN) = t
'       PRSR(NNN) = P
'       IndexRefraction(NNN) = n
'       NNN = NNN + 1
'    Next jr
'       ELV(NNN) = rt - s
'       TMP(NNN) = tt
'       PRSR(NNN) = pt
'       IndexRefraction(NNN) = nts
'       NNN = NNN + 1
'    For jr = rt + 200 To rs Step 200
'       Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, CLng(100# * NNN / numpoints))
'       Call atmosstr(rt, tt, nt, a(2), CDbl(jr), n, dndrt, pt, P)
'       ELV(NNN) = CDbl(jr) - s
'       TMP(NNN) = tt
'       PRSR(NNN) = P
'       IndexRefraction(NNN) = n
'       NNN = NNN + 1
'    Next jr
'    NNN = NNN - 1
'
'    NumTemp = NNN + 1
'
'    End If
'
'myfile$ = Dir(App.Path & "\test_HS.dat")
'If myfile$ <> sempty Then
'   Kill App.Path & "\test_HS.dat"
'   End If
'myfile$ = Dir(App.Path & "\tc_HS.dat")
'If myfile$ <> sempty Then
'   Kill App.Path & "\tc_HS.dat"
'   End If
'
'Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, 0) 'reset

'C - - - - - Integrate the refraction integral in the troposphere and
'------------stratosphere . ie Ref = Ref troposphere + Ref stratosphere .
'------------do this by taking steps in z (zenith angle) defined w.r.t observer at r
'------------and finding the corresponding r to that zenith angle.
'------------Accumulate the refraction until r is equal to height of atmosphere

'C - - - - - Initial step lengths etc .
'
'StatusMes = "Calculating refraction as a function of view angle for wavelength " & Str(wl * 1000) & " nm"
'Call StatusMessage(StatusMes, 1, 0)
'
'   For jstep = 1 To n_size  '<2
'
'        If (IPLOT = 1) Then
'        '            PRINT *, J
'        '            StatusMes = Str(J)
'        '            Call StatusMessage(StatusMes, 1, 0)
'           Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, CLng(100# * jstep / n_size))
'           End If
'        '         II = IIS
'        '         IDCT(j) = 0
'        '         AIRM(j) = 0#
'        '         SSR(k, j) = 0#
'        ALFA(KWAV, jstep) = (CDbl(n_size / 2 - jstep) / PPAM)
'        z0 = 90# - ALFA(KWAV, jstep) / 60#
'        z0r = z0 * dgr
'
'        If prjAtmRefMainfm.chkHSoatm.Value = vbChecked Then
'           'find the temperature and pressure and calculate the index of refraction and dndr
'           r0 = s + h0
'           rt = s + ht
'           rs = s + hs
'           For ih = 0 To NNN - 1
'              If ELV(ih - 1) <= r0 And ELV(ih) > r0 Then
'                 'at the observer
'                 ih0 = ih
'                 t0 = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (TMP(ih) - TMP(ih - 1)) + TMP(ih - 1)
'                 p0t = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (PRSR(ih) - PRSR(ih - 1)) + PRSR(ih - 1)
'                 'calculate the exponent delta in HS
'                 'check for isothermal regions
'                 If TMP(ih) <> TMP(ih - 1) Then
'                    Isothermic = False
'                    a(1) = Abs((TMP(ih) - TMP(ih - 1)) / (ELV(ih) - ELV(ih - 1))) 'lapse rate
'                    a(3) = Log(PRSR(ih) / PRSR(ih - 1)) / Log(TMP(ih) / TMP(ih - 1))
'                    a(5) = pw0 * (1# - mw / md) * a(3) / (a(4) - a(3))
'                    a(6) = p0 + a(5) 'full coeficient for pressure of air with humidity ups
'                    a(7) = z1 * a(6) / t0 'humid air pressure contribution to n
'                    a(8) = (z1 * a(5) + z2 * pw0) / t0 'water vapor contribution to n
'                    a(9) = (a(3) - 1#) * a(1) * a(7) / t0 'first coeficient in expression for dndr
'                    a(10) = (a(4) - 1#) * a(1) * a(8) / t0 'second coeficient in expression for dndr
'
'                    Call atmostro(r0, t0, a(), r0, t0O, n0, dndr0, P)
'
'                    sk0 = n0 * r0 * Sin(z0r)
'                    f0 = refii(r0, n0, dndr0)
'                 Else
'                    Isothermic = True
'                    n0 = 1# + z1 * p0t
'                    dndr0 = 0#
'                    sk0 = n0 * r0 * Sin(z0r)
'                    f0 = refii(r0, n0, dndr0)
'                    End If
'
'
'              ElseIf ELV(ih - 1) <= rt And ELV(ih) > rt Then
'                 'at the tropopause
'                 iht = ih
'                 tt = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (TMP(ih) - TMP(ih - 1)) + TMP(ih - 1)
'                 pt = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (PRSR(ih) - PRSR(ih - 1)) + PRSR(ih - 1)
'                 If TMP(ih) <> TMP(ih - 1) Then
'                    Isothermic = False
'                    a(1) = Abs((TMP(ih) - TMP(ih - 1)) / (ELV(ih) - ELV(ih - 1))) 'lapse rate
'                    a(3) = Log(PRSR(ih) / PRSR(ih - 1)) / Log(TMP(ih) / TMP(ih - 1))
'                    a(5) = pw0 * (1# - mw / md) * a(3) / (a(4) - a(3))
'                    a(6) = p0 + a(5) 'full coeficient for pressure of air with humidity ups
'                    a(7) = z1 * a(6) / t0 'humid air pressure contribution to n
'                    a(8) = (z1 * a(5) + z2 * pw0) / t0 'water vapor contribution to n
'                    a(9) = (a(3) - 1#) * a(1) * a(7) / t0 'first coeficient in expression for dndr
'                    a(10) = (a(4) - 1#) * a(1) * a(8) / t0 'second coeficient in expression for dndr
'                    'C - - - - - At the Tropopause in the Troposphere .
'                    Call atmostro(r0, t0, a(), rt, tt, nt, dndrt, pt)
'                    zt = DASIN(sk0 / (rt * nt)) / dgr 'convert back to degrees
'                    ft = refii(rt, nt, dndrt)
'
'                    'C - - - - - At the Tropopause in the Stratosphere .
'                    nts = nt
'                    zts = DASIN(sk0 / (rt * nts)) / dgr 'convert back to degrees
'                    fts = refii(rt, nts, dndrts)
'
'                 Else
'                    Isothermic = True
'                    'at the tropopause in the troposphere
'                    n = 1# + z1 * pt
'                    dndrt = 0#
'                    zt = DASIN(sk0 / (rt * nt)) / dgr 'convert back to degrees
'                    ft = refii(rt, nt, dndrt)
'                    'at the troppause in the stratosphere
'                    nts = nt
'                    dndrts = 0#
'                    zts = DASIN(sk0 / (rt * nts)) / dgr 'convert back to degrees
'                    fts = refii(rt, nts, dndrts)
'                    End If
'
'              ElseIf ELV(ih - 1) <= rs And ELV(ih) > rs Then
'                 ihs = ih 'end of stratosphere
'                 ts = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (TMP(ih) - TMP(ih - 1)) + TMP(ih - 1)
'                 ps = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (PRSR(ih) - PRSR(ih - 1)) + PRSR(ih - 1)
'                 ns = 1# + nt * ps / pt
'                 dndrs = nt * (PRSR(ih) / pt - PRSR(ih - 1) / pt) / (ELV(ih) - ELV(ih - 1))
'                 zs = DASIN(sk0 / (rs * ns)) / dgr 'convert back to degrees
'                 fs = refii(rs, ns, dndrs)
'
'                 Exit For
'
'                 End If
'
'           Next ih
'
'        Else
'
'            'C - - - - - At the observer .
'            r0 = s + h0
'            Call atmostro(r0, t0, a(), r0, t0O, n0, dndr0, P)
'
'            sk0 = n0 * r0 * Sin(z0r)
'            f0 = refii(r0, n0, dndr0)
'
'            'C - - - - - At the Tropopause in the Troposphere .
'            rt = s + ht
'            Call atmostro(r0, t0, a(), rt, tt, nt, dndrt, pt)
'            zt = DASIN(sk0 / (rt * nt)) / dgr 'convert back to degrees
'            ft = refii(rt, nt, dndrt)
'
'            'C - - - - - At the Tropopause in the Stratosphere .
'            Call atmosstr(rt, tt, nt, a(2), rt, nts, dndrts, pt, P)
'            zts = DASIN(sk0 / (rt * nts)) / dgr 'convert back to degrees
'            fts = refii(rt, nts, dndrts)
'
'            'C - - - - - At the stratosphere limit
'            rs = s + hs
'            Call atmosstr(rt, tt, nt, a(2), rs, ns, dndrs, pt, P)
'            zs = DASIN(sk0 / (rs * ns)) / dgr 'convert back to degrees
'            fs = refii(rs, ns, dndrs)
'
'            End If
'
'ThetaTot = 0#
'LE = 0#
'lR = 0#
'r00 = r0
'
'ref0 = -999.999
'iss = 16
'For k = 1 To 2
'    istart = 0
'    fe = 0#
'    fo = 0#
'
'    If (k = 1) Then
'        h = (zt - z0r / dgr) / CDbl(iss) 'iniital step size in zenith angle -- within the troposphere
'        fb = f0
'        ff = ft
'    ElseIf (k = 2) Then
'        h = (zs - zts) / CDbl(iss) 'initial step size in zenith angle --within the starosphere
'        fb = fts
'        ff = fs
'        End If
'
'    inn = iss - 1
'    iss = iss / 2
'    step = h
''100 CONTINUE
'100:
'    For i = 1 To inn
'
''        If i = 1 Then 'new loop, record z, r
''           Close
''           fileout% = FreeFile
''           Open App.Path & "\testHS.dat" For Output As #fileout%
''           L = 0
''           numPoints = 0
''           End If
'
'        If (i = 1 And k = 1) Then
'            z = z0r / dgr + h
'            r = r0
'
'           Close
'           fileout% = FreeFile
'           Open App.Path & "\testHS_t.dat" For Output As #fileout%
''           L = 0
''           Phi = 0
'
''           Phi = -Atn(r * dndr0) / dgr
''           Theta = 90# - z0 - ThetaTot
'           Theta = 0
'           ThetaTot = 90# - z0
'           LE = r00 * Theta * dgr
'           lR = Sqr(LE * LE + (r - s - h0) * (r - s - h0)) 'use euclidean geometry (pythagorean theorem) approximation to non-spherical curved surface distance
'           'determine the X,Y of the endpoints
'           X0 = 0
'           Y0 = h0
''           A1 = (XP * Cos(ANGLE) + (RE + YP) * Sin(ANGLE)) * Multiplication
''           A2 = (-XP * Sin(ANGLE) + (RE + YP) * Cos(ANGLE)) * Multiplication
'           XR = X0 + r00 * Sin(Theta * dgr)
'           YR = h0 + r00 * Cos(Theta * dgr) + s
''           r00 = r
'           X0 = 0
'           Y0 = r0
'           Theta0 = ThetaTot
'
''           Print #fileout%, Format(Str(z0), "##0.0####"), Format(Str(r0 - s), "######0.0####"), Format(Str(0#), "#########0.0####"), _
''                            Format(Str(0#), "#########0.0####"), Format(Str(X0 * 0.001), "#########0.0####"), _
''                            Format(Str(h0 * 0.001), "#########0.0####")
'           Write #fileout%, X0, r0
'
''           numpoints = 0
''           ELV(numpoints) = h0
''           TMP(numpoints) = t0
''           PRSR(numpoints) = p0
''           IndexRefraction(numpoints) = n0
'
'        ElseIf (i = 1 And k = 2) Then
'            z = zts + h
'            r = rt
'            Close
'            fileout% = FreeFile
'            Open App.Path & "\testHS_s.dat" For Output As #fileout%
''            numNewPoints = 1
'            LE = LE0
'            lR = LR0
'            X0 = X00
'        Else
'            z = z + step
'            End If
'
''C - - - - - Given the zenith distance ( z ) find r by Newton - Raphson iteration (Magnume eqn 31).
'        rg = r
'        r00 = rg
'        For j = 1 To 4
'            If prjAtmRefMainfm.chkHSoatm.Value = vbChecked Then
'
'               If k = 1 Then 'troposphere
'                  For ih = ih0 To iht
'                     If rg >= ELV(ih) And rg < ELV(ih - 1) Then
'                        If TMP(ih) <> TMP(ih - 1) Then
'                           Isothermic = False
'                           a(1) = Abs((TMP(ih) - TMP(ih - 1)) / (ELV(ih) - ELV(ih - 1))) 'lapse rate
'                           a(3) = Log(PRSR(ih) / PRSR(ih - 1)) / Log(TMP(ih) / TMP(ih - 1))
'                           a(5) = pw0 * (1# - mw / md) * a(3) / (a(4) - a(3))
'                           a(6) = p0 + a(5) 'full coeficient for pressure of air with humidity ups
'                           a(7) = z1 * a(6) / t0 'humid air pressure contribution to n
'                           a(8) = (z1 * a(5) + z2 * pw0) / t0 'water vapor contribution to n
'                           a(9) = (a(3) - 1#) * a(1) * a(7) / t0 'first coeficient in expression for dndr
'                           a(10) = (a(4) - 1#) * a(1) * a(8) / t0 'second coeficient in expression for dndr
'                           'C - - - - - At the Tropopause in the Troposphere .
'                           Call atmostro(r0, t0, a(), rt, tt, nt, dndrt, pt)
'                        Else
'                            n = 1# + z1 * p0t
'                            dndr = 0#
'                           End If
'
'                        Exit For
'                        End If
'                  Next ih
'
'               ElseIf k = 2 Then 'stratosphere
'                  For ih = iht To ihs
'                     If rg >= ELV(ih) And rg < ELV(ih - 1) Then
'                        pps = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (PRSR(ih) - PRSR(ih - 1)) + PRSR(ih - 1)
'                        n = 1# + nt * pps / pt
'                        dndr = nt * (PRSR(ih) / pt - PRSR(ih - 1) / pt) / (ELV(ih) - ELV(ih - 1))
'                        Exit For
'                        End If
'                  Next ih
'                  End If
'
'            Else
'                If (k = 1) Then
'                    Call atmostro(r0, t0, a(), rg, tg, n, dndr, P)
'                ElseIf (k = 2) Then
'                    Call atmosstr(rt, tt, nt, a(2), rg, n, dndr, pt, P)
'                    End If
'                End If
'            rg = rg - ((rg * n - sk0 / Sin(z * dgr)) / (n + rg * dndr))
'        Next j
'        r = rg
'
''C - - - - - Find refractive index and integrand at r .
'        If prjAtmRefMainfm.chkHSoatm.Value = vbChecked Then
'
'           If k = 1 Then 'troposphere
'                  For ih = ih0 To iht
'                     If rg >= ELV(ih) And rg < ELV(ih - 1) Then
'                        If TMP(ih) <> TMP(ih - 1) Then
'                           Isothermic = False
'                           a(1) = Abs((TMP(ih) - TMP(ih - 1)) / (ELV(ih) - ELV(ih - 1))) 'lapse rate
'                           a(3) = Log(PRSR(ih) / PRSR(ih - 1)) / Log(TMP(ih) / TMP(ih - 1))
'                           a(5) = pw0 * (1# - mw / md) * a(3) / (a(4) - a(3))
'                           a(6) = p0 + a(5) 'full coeficient for pressure of air with humidity ups
'                           a(7) = z1 * a(6) / t0 'humid air pressure contribution to n
'                           a(8) = (z1 * a(5) + z2 * pw0) / t0 'water vapor contribution to n
'                           a(9) = (a(3) - 1#) * a(1) * a(7) / t0 'first coeficient in expression for dndr
'                           a(10) = (a(4) - 1#) * a(1) * a(8) / t0 'second coeficient in expression for dndr
'                           'C - - - - - At the Tropopause in the Troposphere .
'                           Call atmostro(r0, t0, a(), rt, tt, nt, dndrt, pt)
'                        Else
'                            n = 1# + z1 * p0t
'                            dndr = 0#
'                           End If
'
'                        Exit For
'                        End If
'                  Next ih
'           ElseIf k = 2 Then 'stratosphere
'                  For ih = iht To ihs
'                     If rg >= ELV(ih) And rg < ELV(ih - 1) Then
'                        pps = (r0 - ELV(ih - 1)) / (ELV(ih) - ELV(ih - 1)) * (PRSR(ih) - PRSR(ih - 1)) + PRSR(ih - 1)
'                        n = 1# + nt * pps / pt
'                        dndr = nt * (PRSR(ih) / pt - PRSR(ih - 1) / pt) / (ELV(ih) - ELV(ih - 1))
'                        Exit For
'                        End If
'                  Next ih
'
'              End If
'
'        Else
'            If (k = 1) Then
'                Call atmostro(r0, t0, a, r, t, n, dndr, P)
'                pt = P
'
'    '            Strat = 0#
'            ElseIf (k = 2) Then
'    '            If r = rt Then
'    '               ccc = 1
'    '               End If
'                Call atmosstr(rt, tt, nt, a(2), r, n, dndr, pt, P)
'                t = tt
'
'    '            BegNum = numpoints
'    '            Strat = 1#
'
'                End If
'             End If
'
''        If Strat = 0# Then
''            numpoints = numpoints + 1
''            ELV(numpoints) = r - s
''            TMP(numpoints) = t
''            PRSR(numpoints) = P
''            IndexRefraction(numpoints) = n
''        ElseIf Strat = 1# Then
''            ELV(BegNum + numNewPoints) = r - s
''            TMP(BegNum + numNewPoints) = t
''            PRSR(BegNum + numNewPoints) = P
''            IndexRefraction(BegNum + numNewPoints) = n
''            numNewPoints = numNewPoints + 1
''            End If
'
'
'
'        'calculate the distance taveled, and cylindrical coordinates of the ray
'
''           Theta0 = Theta
'           Theta = (90# - z)
'           ThetaTot = Theta - Theta0
''           ThetaTot = 90# - z
'           LE = r00 * Theta * dgr
'           lR = lR + Sqr(LE * LE + (r - s - h0) * (r - s - h0)) 'use euclidean geometry (pythagorean theorem) approximation to non-spherical curved surface distance
'           'determine the X,Y of the endpoints
'           LE = LE + r00 * Theta * dgr
''           XR = X0 + r * Sin(ThetaTot * dgr)
'           XR = r * Sin(ThetaTot * dgr)
'
''           YR = h0 + r * Cos(ThetaTot * dgr) - s
'           YR = r * Cos(ThetaTot * dgr)
'           X0 = XR
'           Y0 = YR
'
'           'check that ray doesn't intersect with the earth, if it does, then signal that with YR = -1000
'
''           Print #fileout%, Format(Str(z), "##0.0####"), Format(Str(r - s), "######0.0####"), Format(Str(LE * 0.001), "#########0.0####"), _
''                            Format(Str(lR * 0.001), "#########0.0####"), Format(Str(XR * 0.001), "#########0.0####"), _
''                            Format(Str(YR * 0.001), "#########0.0####")
'           If r < s Then 'collided with the surface
'              Write #fileout%, XR, -1000
'              jstop = jstep
'           Else
'              Write #fileout%, XR, YR
'              jstop = -1
'              End If
'
'           If k = 1 Then
'            'keep track of last troposphere values
'            LE0 = LE
'            LR0 = lR
'            X00 = X0
'            End If
'
'
'        f = refii(r, n, dndr)
'        If (istart = 0 And i Mod 2 = 0) Then
'            fe = fe + f
'        Else
'            fo = fo + f
'            End If
'
'        If jstop <> -1 Then Exit For
'    Next i
'
'    If jstop <> -1 Then Exit For
'
''C - - - - - Evaluate the integrand using Simpsons Rule .
'    refp = h * (fb + 4# * fo + 2# * fe + ff) / 3#
''C - - - - - Test for convergence .
'    If (Abs(refp - ref0) > hepsr) Then
'        iss = 2 * iss
'        inn = iss
'        step = h
'        h = h / 2#
'        fe = fe + fo
'        fo = 0#
'        ref0 = refp
'        If (istart = 0) Then istart = 1
'        GoTo 100
'        End If
'
'    If (k = 1) Then reft = refp
'
'Next k
'Close #fileout%
'
'If jstop <> -1 Then
'   Close
'   GoTo crw500
'   End If
'
''append files
''if first append file
'
''FileCopy App.Path & "\testHS_t.dat", App.Path & "\testHS.dat"
'
'filein% = FreeFile
'Open App.Path & "\testHS_t.dat" For Input As #filein%
'
'fileout% = FreeFile
'Open App.Path & "\test_HS.dat" For Append As #fileout%
'
'Do Until EOF(filein%)
'   Line Input #filein%, doclin$
'   Print #fileout%, doclin$
'Loop
'Close #filein%
''Close #fileout%
'
'filein% = FreeFile
'Open App.Path & "\testHS_s.dat" For Input As #filein%
'
''fileout% = FreeFile
''Open App.Path & "\test_HS.dat" For Append As #fileout%
'
'Do Until EOF(filein%)
'   Line Input #filein%, doclin$
'   Print #fileout%, doclin$
'Loop
'Close #filein%
'Close #fileout%
'
'Kill App.Path & "\testHS_t.dat"
'Kill App.Path & "\testHS_s.dat"
'
''C - - - - - Refraction in the troposphere + stratosphere in degrees .
'ref = (reft + refp)  'degrees
'refmrad = ref * dgr * 1000 'refraction in mrad
'
''now append the refraction value to the transfer file
'ALFT(KWAV, jstep) = ALFA(KWAV, jstep) - ref * 60# 'true depression angle
''fileout% = FreeFile
''Open App.Path & "\tc_HS.dat" For Append As #fileout%
''Print #fileout%, ALFA(KWAV, jstep), ALFT(KWAV, jstep) 'units in minutes of arc
''Close #fileout%
'
'If z0 = 90# Then
'   prjAtmRefMainfm.lblRef.Caption = "Atms. refraction (deg.) = " & ref & vbCrLf & "Atms. refraction (mrad) = " & refmrad
'   End If
'
'crw500:
'If jstop <> -1 Then
'   Close
'
'   StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KWAV, jstop - 1))& vbcrlf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(s / (s + h0)) / CONV))
'     Call StatusMessage(StatusMes, 1, 0)
''     ier = MsgBox(StatusMes, vbInformation + vbOKOnly, "Horizon")
'     lblHorizon.Caption = StatusMes
'
'   Exit For
'   End If
'
'Next jstep
'
'Next KWAV
'
'CalcComplete = True

Return

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

cmdRefWilson_Click_Error:

    Close
    
   cmdCalc.Enabled = True
   cmdRefWilson.Enabled = True
   cmdMenat.Enabled = True
   cmdVDW.Enabled = True

    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdRefWilson_Click of Form prjAtmRefMainfm"

End Sub
'*******************************************************
Sub atmostro(r0 As Double, t0 As Double, A() As Double, r As Double, T As Double, n As Double, dndr As Double, P As Double)

'DOUBLE PRECISION r0 , t0 , a(10) , r , t , n
Dim tt0 As Double, tt01 As Double, tt02 As Double
T = t0 - A(1) * (r - r0) 'temprerature in troposphere
tt0 = T / t0
tt01 = tt0 ^ (A(3) - 2#)
tt02 = tt0 ^ (A(4) - 2#)
n = 1# + (A(7) * tt01 - A(8) * tt02) * tt0
dndr = -1# * A(9) * tt01 + A(10) * tt02
P = A(6) * tt01 - A(5) * tt02 'total pressure in troposphere = humid air pressure + water vapor pressure (mb)
End Sub
'*******************************************************
Sub atmosstr(rt As Double, TT As Double, nt As Double, taw As Double, r As Double, n As Double, dndr As Double, pt As Double, P As Double)
'IMPLICIT none
'DOUBLE PRECISION rt , tt , nt , taw , r , n , dndr
Dim b As Double, lognm1 As Double
Dim exp1 As Double
Dim logdndrp1 As Double
'Dim p00 As Double
'Dim exp2 As Double, p00 As Double

b = taw / TT 'temperature constant = tt (K)
exp1 = Exp(-b * (r - rt))
n = 1# + (nt - 1#) * exp1
dndr = -1# * b * (nt - 1#) * exp1

'p00 = P
'pt is pressure at tropopause of torposphere
P = pt * exp1 'Pressure in the stratosphere (mb)
'If P > p00 Then
'   ccc = 1
'   End If
   
End Sub
'*******************************************************
Function refii(r As Double, n As Double, dndr As Double) As Double
'IMPLICIT none
'DOUBLE PRECISION r , n , dndr
refii = r * dndr / (n + r * dndr)
End Function

Private Sub cmdRight_Click()
   Xorigin = Xorigin - prjAtmRefMainfm.Width / 20
   Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
End Sub

Private Sub cmdShowSuns_Click()
Dim ier As Long

If Dir(App.Path & "\temp.ppm") <> sEmpty Then
    'read the width and height
    filnum% = FreeFile
    Open App.Path & "\temp.ppm" For Input As #filnum%
    Line Input #filnum%, doclin$
    Line Input #filnum%, doclin$
    Input #filnum%, m, n
    Close #filnum%
    
    'now plot it
    Dim AspectRatio As Double
    AspectRatio = GetScreenAspectRatio()
    
    With frmZoomPan.AlphaImgCtl1
       .Width = m
       .height = n
    End With
    
    With frmZoomPan
       .Width = (.AlphaImgCtl1.Left + .AlphaImgCtl1.Width + .VScrollPan.Width + 30) * Screen.TwipsPerPixelX '.cmdSkew.Width + 100) * Screen.TwipsPerPixelX
       .height = (.AlphaImgCtl1.Top + .AlphaImgCtl1.height + .HScrollPan.height + .HScrollZoom.height + 100) * Screen.TwipsPerPixelY
    End With
    
    With frmZoomPan.VScrollPan
       .Top = frmZoomPan.AlphaImgCtl1.Top
       .Left = frmZoomPan.AlphaImgCtl1.Left + frmZoomPan.AlphaImgCtl1.Width + 10
       .height = frmZoomPan.AlphaImgCtl1.height
    End With

    With frmZoomPan.HScrollPan
        .Top = frmZoomPan.AlphaImgCtl1.Top + frmZoomPan.AlphaImgCtl1.height + 10
        .Width = frmZoomPan.AlphaImgCtl1.Left + frmZoomPan.AlphaImgCtl1.Width
        .Left = frmZoomPan.AlphaImgCtl1.Left
    End With

    With frmZoomPan.HScrollZoom
       .Left = frmZoomPan.HScrollPan.Left
       .Width = frmZoomPan.HScrollPan.Width
       .Top = frmZoomPan.HScrollPan.Top + frmZoomPan.HScrollPan.height + 10
    End With
    
    With frmZoomPan.lblZoom
       .Top = frmZoomPan.HScrollZoom.Top + frmZoomPan.HScrollZoom.height + 10
       .Left = frmZoomPan.HScrollPan.Left + frmZoomPan.HScrollPan.Width * 0.5
    End With
    
    With frmZoomPan
       .lblPan.Visible = False
       .cmdMask.Visible = False
       .cmdSkew.Visible = False
    End With
    
    frmZoomPan.Visible = True
    
    Set frmZoomPan.AlphaImgCtl1.Picture = LoadPictureGDIplus(App.Path & "\temp.ppm")
    frmZoomPan.Refresh
    
    ier = BringWindowToTop(frmZoomPan.hwnd)
    
    End If
    
End Sub

Private Sub cmdSmaller_Click()
   Dim StatusMes As String
   If Mult * 0.5 >= 1 Then
     RefZoom.LastZoom = Mult
     Mult = Mult * 0.5
     RefZoom.Zoom = Mult
     StatusMes = "Multiplication = " & Format(Mult, "#############0.0#")
     Call StatusMessage(StatusMes, 1, 0)
     txtStartMult = Mult

     Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
     End If
End Sub

Private Sub cmdTest_Click()
   Dim Temp As Double, TR As Double, RETH As Double, cd As Double, bbb As Double
   RETH = 6356.766
   cd = 3.14159265359 / 180#
   
    Set PtX = New Collection
    Set PtY = New Collection
   
   filein% = FreeFile


'   FileName = App.Path & "/TRC_T-288.15-288.15_HOSV-756.5-756.5_DOBST-15000-120000_HOBST-959.2-959.2.dat"
'    txtT11 = 288.15
'    txtH11 = 756.5
'    txtH21 = 959.2
'    txtD1 = 15
'    txtD2 = 120
'    chkDist.Value = vbChecked
    
'   FileName = App.Path & "/TRC_T-288.15-288.15_HOSV-250-250_DOBST-15000-120000_HOBST-3000-3000.dat"
'    txtT11 = 288.15
'    txtH11 = 250
'    txtH21 = 3000
'    txtD1 = 15
'    txtD2 = 120
'    chkDist.Value = vbChecked
    
   filename = App.Path & "/TRC_T-260-320_HOSV-100-100_DOBST-60824.1061819471-0_HOBST-302.7-302.7.dat"
   TotalDist = 60.8241061819471
   H21 = 302.7
   H11 = 100
    chkTemp.Value = vbChecked
    
'   FileName = App.Path & "/TRC_T-260-320_HOSV-756.5-756.5_DOBST-60824.106181947-60824.106181947_HOBST-959.2-959.2.dat"
'   TotalDist = 60.8241061819471
'   H21 = 959.2
'   H11 = 756.5
'   chkTemp.Value = vbChecked

   NNN = 0
   Open filename For Input As #filein%
   Do Until EOF(filein%)
      Line Input #filein%, doclin$
      NNN = NNN + 1
   Loop
   Close #filein%
   filein% = FreeFile
   Open filename For Input As #filein%
   
   If chkTemp.Value = vbChecked Then
   
        ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
        
        maxva = -999999999
        minva = 999999999
        TMax = -1E+17
        TMin = 1E+19
           
        j = 0
        Do While Not EOF(filein%)
           Input #filein%, Temp, A, b, TR
           j = j + 1
           TransferCurve(j, 1) = " " & CStr(Temp)
           TransferCurve(j, 2) = TR
           
           PATHLENGTH = Sqr(TotalDist ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (TotalDist ^ 2#) / RETH) ^ 2#)
           
           PtX.Add Temp
           bbb = (TR * Temp * Temp) / (0.0083 * PATHLENGTH * Val(txtPress0))
           PtY.Add bbb * Temp
           
           If TR > maxva Then maxva = TR
           If TR < minva Then minva = TR
           
        Loop
        Close #filein%
        
           With MSChartTR
             .chartType = VtChChartType2dLine
             .RandomFill = False
             .RowCount = 2
             .ColumnCount = NNN
             .RowLabel = "Temperature (degrees Kelvin)"
             .ColumnLabel = "Terrestrial refraction (degrees)"
             .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
             .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
             '      .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = Format(((maxva - minva) \ NNN), "##0.####0")
             '      .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 10
             .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1.1 * maxva
             .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0.9 * minva
             .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = TMin
             .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = TMax
             .ChartData = TransferCurve
             
           End With
           
    ElseIf chkDist.Value = vbChecked Then
       ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
       
       maxva = -999999999
       minva = 999999999
         
       filein% = FreeFile
       Open filename For Input As #filein%
         
       j = 0
       TotalDist = 0
       Do While Not EOF(filein%)
          Input #filein%, Temp, A, b, TR
          TotalDist = Val(txtD1) + j * Val(txtStepD1)
          j = j + 1
          TransferCurve(j, 1) = " " & CStr(TotalDist)
          TransferCurve(j, 2) = TR
          PtX.Add TotalDist
          'caculate the approximate ray distance
'            //use Lehn's parabolic path approx to ray trajectory and Brutton equation 58
          PATHLENGTH = Sqr(TotalDist ^ 2# + ((Val(txtH21) - Val(txtH11)) * 0.001 - 0.5 * (TotalDist ^ 2#) / RETH) ^ 2#)
          
          bbb = (TR * Val(txtT11) * Val(txtT11)) / (0.0083 * PATHLENGTH * Val(txtPress0))
          PtY.Add bbb * TotalDist
        
          If TR > maxva Then maxva = TR
          If TR < minva Then minva = TR
          
       Loop
       Close #filein%
       
        ' Find a good fit.
'        degree = 2
'        Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)
    
    '    stop_time = Timer
    '    Debug.Print Format$(stop_time - start_time, "0.0000")
    
'        Txt = ""
'        For i = 1 To BestCoeffs.Count
'            Txt = Txt & " " & BestCoeffs.Item(i)
'        Next i
'        If Len(Txt) > 0 Then Txt = Mid$(Txt, 2)
'    '    txtAs.Text = Txt
'
'        ' Display the error.
'        ShowError
'
'        ' We have a solution.
'        HasSolution = True
'    '    picGraph.Refresh
       
        With MSChartTR
          .chartType = VtChChartType2dLine
          .RandomFill = False
          .RowCount = 2
          .ColumnCount = NNN
          .RowLabel = "Distance between observer and obstruction (kms)"
          .ColumnLabel = "Terrestrial refraction (degrees)"
          .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
          .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
          '      .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = Format(((maxva - minva) \ NNN), "##0.####0")
          '      .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 10
          .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 1.1 * maxva
          .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0.9 * minva
          .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = D1
          .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = D2
          .ChartData = TransferCurve
        End With
    
       End If
      
    ' Find a good fit.
    degree = 1
    Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree)

'    stop_time = Timer
'    Debug.Print Format$(stop_time - start_time, "0.0000")

    Txt = ""
    For i = 1 To BestCoeffs.Count
        Txt = Txt & " " & BestCoeffs.Item(i)
    Next i
    If Len(Txt) > 0 Then Txt = Mid$(Txt, 2)
    txtAs.Text = Txt

    ' Display the error.
    Call ShowError(txtError)

    ' We have a solution.
    HasSolution = True
      
End Sub

Private Sub cmdup_Click()
   Yorigin = Yorigin - prjAtmRefMainfm.height / 20
   Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdVDW_Click
' Author    : Dr-John-K-Hall
' Date      : 4/1/2019
' Purpose   : Ray Tracing using van der Werf's formulation
'---------------------------------------------------------------------------------------
'
Private Sub cmdVDW_Click()

'DefDbl A-H, O-Z
'Option Explicit
'
Dim PRESSD1(99999) As Double, PRESSD2(99999) As Double
Dim HMAXT As Double, TLOW As Double, THIGH As Double, Press0 As Double
Dim RELHUM As Double, BETALO As Double, BETAHI As Double, BETAST As Double, WAVELN As Double
Dim NSTEPS As Long, i As Long, FSize As Long, BETAM As Double, StatusMes As String, NumLayers As Long
Dim NoShow As Boolean, j As Long, jstep As Long, TempStart As Double, TempEnd As Double, TempStep As Double, TLoop As Double, NTLoop As Long
Dim Dist As Double, REFRAC As Double, AIRDRY As Double, AIRVAP As Double, PHI1 As Double
Dim BETA1 As Double, H1 As Double, ier As Long, PATHLENGTH As Double
Dim Pathh As Double
Dim FKP1 As Double, FKR1 As Double, FKB1 As Double, FKAD1 As Double, FKAV1 As Double
Dim PHINEW As Double, RNEW As Double, BETANEW As Double, HNEW As Double
Dim FKP2 As Double, FKR2 As Double, FKB2 As Double, FKAD2 As Double, FKAV2 As Double
Dim FKP3 As Double, FKR3 As Double, FKB3 As Double, FKAD3 As Double, FKAV3 As Double
Dim FKP4 As Double, FKR4 As Double, FKB4 As Double, FKAD4 As Double, FKAV4 As Double
Dim PHI2 As Double, R2 As Double, BETA2 As Double, H2 As Double, DREFR As Double, NumBet As Long
Dim HgtStart As Double, HgtEnd As Double, HgtStep As Double, HLoop As Double, NHloop As Long
Dim FilePath As String, a0 As Double, e0 As Double
Dim StartedNumber As Boolean, AA(6) As Double

   On Error GoTo cmdVDW_Click_Error
   
'set error flag
cmdVDW_error = 0

MDIAtmRef.StatusBar.Panels(2).Text = vbNullString 'clear status bar

If chkVDW_Show.Value = vbUnchecked Or TempLoop Then
   frmVDW.Enabled = False
   picVDW.Visible = False
   NoShow = True
Else
   frmVDW.Enabled = True
   picVDW.Visible = True
   NoShow = False
   End If

FSize = picVDW.FontSize

RefCalcType% = 3
cmdVDW.Enabled = False
cmdCalc.Enabled = False
cmdMenat.Enabled = False
cmdRefWilson.Enabled = False

NumLayers = 0

'///////////////zero global elv, temp, pressure, refraction arrays/////////////////////////
For II = 0 To MaxViewSteps& - 1
   ELV(II) = 0
   TMP(II) = 0
   PRSR(II) = 0
   For jj = 0 To 81
     RCV(jj, II) = 0
   Next jj
Next II

For II = 0 To 82
   For jj = 0 To MaxViewSteps& - 1
    ALFA(II, jj) = 0
    ALFT(II, jj) = 0
   Next jj
Next II

'zero raytracing display arrays
For II = 0 To NumSuns - 1
   For jj = 0 To TotNumSunAlt - 1
      SunAngles(II, jj) = 0
   Next jj
   NumSunAlt(II) = 0
Next II

'////////////////////////////////////

'
'DECLARE FUNCTION f.VAPOR (H)
'DECLARE FUNCTION f.STNDATM (H) 'US1976 Standard atmosphere
'DECLARE FUNCTION f.FNDPD1 (H) 'Find INT[1/T] from lookup table PRESSD1
'DECLARE FUNCTION f.FNDPD2 (H) 'Find INT[/T] from lookup table PRESSD2
'DECLARE FUNCTION f.DLNTDH (H) 'd(ln(T))/dh = (1/T) dT/dH
'DECLARE FUNCTION f.RCINV (H) 'Curvature (1/r) of a horizontal light ray
'DECLARE FUNCTION f.TEMP (H) 'Temperature as a function of elevation
'DECLARE FUNCTION f.REFIND (H) 'Index of refraction
'DECLARE FUNCTION f.DTDH (H) 'dT/dH
'DECLARE FUNCTION f.GUESSL (BETAM, height) 'Guess horizontal distance needed to reach HEIGHT.
'DECLARE FUNCTION f.GUESSP (BETA, height) 'Guess path length needed to reach HEIGHT.
'DECLARE FUNCTION f.ASN (X) 'Arc-sine
'DECLARE FUNCTION f.ACS (X) 'Arc-cosine
'DECLARE FUNCTION f.DENSDRY (H) ' Density of dry air as function of elevation
'DECLARE FUNCTION f.DENSVAP (H) ' Density of water vapor as function of elevation
'DECLARE FUBCTION f.GRAVRAT(H) 'GRAVITY/GRAVITY(0) AS FUNCTION OF HEIGHT
'DECLARE FUNCTION f.PRESSURE(H) 'Find total pressure from f.FNDPD2 and vapor pressure
'DECLARE FUNCTION f.DVAPDT(H) 'Derived from vapor pressure to T as a function of H
'DECLARE FUNCTION f.DNDH (H) ' Derivative of refraction index
'Screen 12

     '------------------progress bar initialization
    With prjAtmRefMainfm
      '------fancy progress bar settings---------
      .progressfrm.Visible = True
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
    
    If OptionSelby.Value = True Then
      StatusMes = "Calculating and Storing multilayer atmospheric details"
      Call StatusMessage(StatusMes, 1, 0)
      
     Dim FNM As String
     Dim AtmType As Integer, AtmNumber As Integer, lpsrate As Double, tst As Double, pst As Double
     Dim NNN As Long, Mode As Integer
     
     'zero the HL and TL and LRL arrays
     For ih = 0 To MaxViewSteps& - 1 '0 To 49
        ELV(ih) = 0#
        TMP(ih) = 0#
        PRSR(ih) = 0#
'        HL(ih) = 0#
'        TL(ih) = 0#
'        LRL(ih) = 0#
     Next ih
     NNN = 0
     
     If OptionLayer.Value = True Then
        AtmType = 1
        FNM = App.Path & "\stmod1.dat"
     ElseIf OptionRead.Value = True Then
        AtmType = 1
        FNM = TextExternal.Text
     ElseIf OptionSelby.Value = True Then
        AtmType = 2
        If prjAtmRefMainfm.opt1.Value = True Then
           AtmNumber = 1
        ElseIf prjAtmRefMainfm.opt2.Value = True Then
           AtmNumber = 2
        ElseIf prjAtmRefMainfm.opt3.Value = True Then
           AtmNumber = 3
        ElseIf prjAtmRefMainfm.opt4.Value = True Then
           AtmNumber = 4
        ElseIf prjAtmRefMainfm.opt5.Value = True Then
           AtmNumber = 5
        ElseIf prjAtmRefMainfm.opt6.Value = True Then
           AtmNumber = 6
        ElseIf prjAtmRefMainfm.opt7.Value = True Then
           AtmNumber = 7
        ElseIf prjAtmRefMainfm.opt8.Value = True Then
           AtmNumber = 8
        ElseIf prjAtmRefMainfm.opt9.Value = True Then
           AtmNumber = 9
        ElseIf prjAtmRefMainfm.opt10.Value = True Then
           AtmNumber = 10
           FNM = txtOther.Text
        Else 'use vdw values
           AtmNumber = 0
           End If
        End If

'     FNM = App.Path & "\stmod1.dat" 'stmod2.dat.txt"
     ier = LoadAtmospheres(FNM, AtmType, AtmNumber, lpsrate, tst, pst, NNN, 4) '5)
     
     If ier < 0 Then
        'error signaled, so exit sub without writting anything
        If LoopingAtmTracing Then 'looping through sondes, so flag that this one is finished
           FinishedTracing = True
           End If
        Exit Sub
        End If
     
     If OptionSelby.Value = True And AtmNumber = 10 Then
        'convert to kilometers if necessary
        If ELV(NNN - 1) < 150 Then
           'scale is in kilmoeters, so OK
        Else
           'convert heights to kilmoeters similar to Selby atmospheres
           For i = 1 To NNN - 1
              ELV(i - 1) = ELV(i - 1) * 0.001
           Next i
           End If
        
'        If HL(NNN - 1) < 150 Then
'           'scale is in kilmoeters, so OK
'        Else
'           'convert heights to kilmoeters similar to Selby atmospheres
'           For i = 1 To NNN - 1
'              HL(i - 1) = HL(i - 1) * 0.001
'           Next i
'           End If
        End If
     
     NumTemp = NNN + 1

     If ier < 0 Then
        Screen.MousePointer = vbDefault
        Close
        cmdVDW.Enabled = True
        cmdCalc.Enabled = True
        cmdMenat.Enabled = True
        cmdRefWilson.Enabled = True
        Exit Sub
     Else
        NumLayers = NNN + 1
        For i = 1 To NumLayers
'           If HL(i) > 10 And i <> 1 Then
           If TMP(i) = TMP(i - 1) And ELV(i) > 9 Then 'this is by definition the cross over region, i.e., the beginning of the isothermal region
             HCROSS = ELV(i)
             TGROUND = txtTGROUND.Text
             If HCROSS > 12 Then 'determination of HCROSS from atmosspheric file failed, i.e., troposphere is full sure < 12 km high
                HCROSS = (TGROUND - 216.65) / 0.0065
                End If
'             HCROSS = (Val(prjAtmRefMainfm.txtTGROUND.Text) - TL(i)) / (Abs(LRL(i)) * 0.001)
             Exit For
             End If


'           If LRL(i) = 0 And HL(i) > 9 Then 'this is by definition the cross over region, i.e., the beginning of the isothermal region
'             HCROSS = HL(i)
'             TGROUND = txtTGROUND.Text
'             If HCROSS > 12 Then 'determination of HCROSS from atmosspheric file failed, i.e., troposphere is full sure < 12 km high
'                HCROSS = (TGROUND - 216.65) / 0.0065
'                End If
''             HCROSS = (Val(prjAtmRefMainfm.txtTGROUND.Text) - TL(i)) / (Abs(LRL(i)) * 0.001)
'             Exit For
'             End If
        Next i
        If HCROSS = 0 Then
             TGROUND = txtTGROUND.Text
             HCROSS = (TGROUND - 216.65) / 0.0065
             End If
        End If

     End If '<<<
     
     If chkDucting.Value = vbChecked Then
        'define constants for inversion layer
        LRL0 = Val(txtDInv.Text) * 0.001 'lapse rate degress K/m
        TL0 = Val(txtEInv.Text) 'maximum height of inversion layer (m)
        HL0 = Val(txtSInv.Text) 'starting height of inversion layer (m)
        AInv = 1 / TL0
        BInv = 1 / HL0
        CInv = LRL0 + 0.0065
        End If


If Not NoShow Then

    picVDW.Cls
    prjAtmRefMainfm.TabRef.Tab = 7
    'Cls
    
    picVDW.Scale (-20, 120)-(100, 0)
    picVDW.DrawMode = 13
    picVDW.ForeColor = QBColor(14)
    picVDW.Line (2, 80)-(10, 81), QBColor(15), BF
    picVDW.Line (2, 80)-(3, 60), QBColor(15), BF
    picVDW.Line (2, 72)-(10, 71), QBColor(15), BF
    picVDW.Line (10, 80)-(9, 71), QBColor(15), BF
    picVDW.Line (7, 71)-(8, 68), QBColor(15), BF
    picVDW.Line (7, 68)-(11, 67), QBColor(15), BF
    picVDW.Line (10, 67)-(11, 60), QBColor(15), BF
    
    
    picVDW.Line (15, 80)-(23, 81), QBColor(15), BF
    picVDW.Line (15, 80)-(16, 60), QBColor(15), BF
    picVDW.Line (15, 60)-(23, 61), QBColor(15), BF
    picVDW.Line (15, 70)-(20, 71), QBColor(15), BF
    
    picVDW.Line (28, 80)-(36, 81), QBColor(15), BF
    picVDW.Line (28, 80)-(29, 40), QBColor(15), BF
    picVDW.Line (28, 70)-(32, 71), QBColor(15), BF
    
    picVDW.Line (43, 75)-(51, 74), QBColor(15), BF
    picVDW.Line (51, 74)-(50, 68), QBColor(15), BF
    picVDW.Line (41, 67)-(51, 68), QBColor(15), BF
    picVDW.Line (41, 67)-(42, 61), QBColor(15), BF
    picVDW.Line (41, 60)-(52, 61), QBColor(15), BF
    
    picVDW.Line (53, 75)-(61, 74), QBColor(15), BF
    picVDW.Line (53, 75)-(54, 60), QBColor(15), BF
    picVDW.Line (53, 60)-(61, 61), QBColor(15), BF
    picVDW.Line (61, 75)-(60, 60), QBColor(15), BF
    
    picVDW.Line (63, 75)-(64, 60), QBColor(15), BF
    
    picVDW.Line (72, 75)-(73, 60), QBColor(15), BF
    picVDW.Line (66, 75)-(72, 74), QBColor(15), BF
    '
    picVDW.ForeColor = QBColor(13)
    'LOCATE 22, 15
    picVDW.CurrentX = 22: picVDW.CurrentY = 15
    picVDW.Print "REFRACTION BASED ON THE MODIFIED US1976 ATMOSPHERE"
    picVDW.CurrentX = 27: picVDW.CurrentY = 13
    picVDW.Print "AUTHOR: Siebren van der Werf"
    picVDW.CurrentX = 28: picVDW.CurrentY = 11
    picVDW.Print "last update: March 2019"
    picVDW.ForeColor = QBColor(15)
    'Sleep 3000 'sleep for 3 seconds
    'Sleep 3
    picVDW.ForeColor = QBColor(14)
    picVDW.FontSize = 14
    picVDW.CurrentX = 28: picVDW.CurrentY = 5: picVDW.Print " press any key to proceed"
    KeyPressed = 0
    Do While KeyPressed = 0
    'in$ = INPUT$(1)
        DoEvents
    Loop
    KeyPressed = 0
    picVDW.Cls
    '=========================================
    'SCREEN 12
    picVDW.Cls
    picVDW.FontSize = 14
    picVDW.ForeColor = QBColor(13)
    'picVDW.Scale (-20, 120)-(100, 0)
    picVDW.CurrentX = 28: picVDW.CurrentY = 113
    picVDW.Print "==========RAYTRACING IN THE MODIFIED US1976 ATMOSPHERE============="
    'picVDW.CurrentX = 28: picVDW.CurrentY = 110
    picVDW.Print "The starting template is that of the US 1976 Standard atmosphere, as described"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 110
    picVDW.Print "for instance in the Handbook of Chemistry and Physics, 81th ed., 2000 - 2001."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 107
    picVDW.Print "In this program some parameters may be adjusted. These are: the observer's"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 104
    picVDW.Print "height, the temperature and pressure at ground level, the wave length of"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 101
    picVDW.Print "the light, the relative humidity and the latitude."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 98
    picVDW.Print "Setting HOBS=0, TGROUND=283.15,PRESS0=1010, RELHUM=0 and OBSLAT=52"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 95
    picVDW.Print "will reproduce the new Nautical Almanac tables, as from 2005."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 92
    picVDW.Print "More information may be found in:"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 89
    picVDW.Print "Siebren Y. van der Werf, Raytracing and refraction in the modified"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 86
    picVDW.Print "US1976 atmosphere, Applied Optics 42(2003)354-366."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 83
    picVDW.Print "Backward raytracing up to 85 km is done using 4th order Runge-Kutta"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 80
    picVDW.Print "numerical integration, using path length as the integration variable."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 77
    picVDW.Print "Details on this method are given in:"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 74
    picVDW.Print "Siebren Y. van der Werf, Comment on `Improved ray tracing air mass"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 71
    picVDW.Print "numbers model', Applied Optics 47(2008)153-156."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 68
    picVDW.Print "Up till a height HMAXT, which is asked for as an input, the calculation will"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 65
    picVDW.Print "be made in a number of steps that must be specified: NSTEPS."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 62
    picVDW.Print "From HMAXT till 100 km the step size will be gradually increased."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 59
    picVDW.Print "HMAXT is further used for preparing a fine-step lookup table for the atm."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 56
    picVDW.Print "pressure and for displaying the ray's curvature for H=0-HMAXT."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 53
    picVDW.Print "Natural constants are taken from the Handbook of Chemistry and Physics, 81th ed."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 50
    picVDW.Print "Refractivities for dry air and for water vapor are taken from:"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 47
    picVDW.Print "P.E. Ciddor, Refractive index of air: new equations for the visible and near"
    'picVDW.CurrentX = 28: picVDW.CurrentY = 43
    picVDW.Print "infrared, Applied Optics 35(1996)1566-1573."
    'picVDW.CurrentX = 28: picVDW.CurrentY = 40
    picVDW.ForeColor = QBColor(14)
    picVDW.CurrentX = 28: picVDW.CurrentY = 20: picVDW.Print " press any key to proceed"
    picVDW.ForeColor = QBColor(15)
    KeyPressed = 0
    Do While KeyPressed = 0
    'in$ = INPUT$(1)
        DoEvents
    Loop
    KeyPressed = 0
    picVDW.Cls
    
    End If
    
'=================LIST OF CONSTANTS===========================
RBOLTZ = 8314.472  'Univ. gas const. = Avogadro's number x Boltzmann's const.
AMASSD = 28.964  'Molar weight of dry air
AMASSW = 18.016  'Molar weight of water
Rearth = 6356766# 'The Earth's mean radius in meters
RE = Rearth
pi = 4 * Atn(1)
RADCON = pi / (60 * 180) 'Converts arcminutes into radians
cd = pi / 180# 'converts degrees into radians
HLIMIT = 100000# 'Maximum height till which the rays are followed
HMAXP1 = 30000 'height where f.FNDPD2 (steps of 10 m) takes over from f.FNDPD1 (steps of 1 m)

STARTALT = Val(txtStartAlt.Text)
DELALT = Val(txtDelAlt.Text)
XMAX = Val(txtXmax.Text) * 1000 'convert km to meters
PPAM = Val(txtPPAM.Text)
KMIN = CInt((Val(txtKmin.Text) - 380) / 5# + 1#)
KMAX = CInt((Val(txtKmax.Text) - 380) / 5# + 1#)
KSTEP = CInt(Val(txtKStep.Text) * 0.1)
      
If TempLoop Then
   If txtYSize = sEmpty Then
      txtYSize = 12
      End If
   prjAtmRefMainfm.txtPPAM = 1#
   prjAtmRefMainfm.txtHeightStepSize = 1#
   txtNSTEPS = 1000
   Press0 = 1013
   txtPress0 = Press0
   End If

n_size = 2 * txtYSize * 60 * PPAM 'convert full width from degrees to minutes of arc and then to pixels
msize = 20 + Val(txtNumSuns.Text) * 32 * PPAM
n = n_size 'height of ppm image of suns, determines range of viewing angles
'm = width of pm image of suns
If Val(txtXSize) = 0 Then
  m = 1000 'width of pm image of suns
Else
  m = Val(txtXSize)
  End If
  
If Dir(App.Path & "\REF2017.OUT") <> sEmpty Then
   Kill App.Path & "\REF2017.OUT"
   End If
If Dir(App.Path & "\REF2017-ATM.OUT") <> sEmpty Then
   Kill App.Path & "\REF2017-ATM.OUT"
   End If


prjAtmRefMainfm.txtBETALO.Text = Format(Str(-n_size * 0.5 / PPAM), "###0.0###") 'convert arcminutes to degrees
prjAtmRefMainfm.txtBETAHI.Text = Format(Str(n_size * 0.5 / PPAM), "###0.0###")
prjAtmRefMainfm.txtBETAST.Text = Format(Str(1# / PPAM), "###0.0###")

Dim KA As Long
For KA = 1 To NumSuns
   ALT(KA) = STARTALT + CDbl(KA - 1) * DELALT
   AZM(KA) = STARTAZM + CDbl(KA - 1) * DELAZM
Next KA
'===================OPTION-MENU OF FORMULA SATURATED VAPOR PRESSURE: =====================
'OPTVAP=1: PL2, POWER LAW
'OPTVAP=2: CC2, CLAUSIUS-CLAPEYRON 2 PAR.
'OPTVAP=3: CC4, CLAUSIUS-CLAPEYRON 4 PAR
'OPTVAP=4: ST, SACKUR-TETRODE, 4 PAR.
OPTVAP = 4
'============INITIALIZATION: THE US 1976 STANDARD ATMOSPHERE=======
'SCREEN 12
picVDW.Cls
HOBS = 0#
TGROUND = 283.15
HMAXT = 1000#
TLOW = 0#
THIGH = 400#
Press0 = 1010#
RELHUM = 0#
BETALO = 0
BETAHI = 60
BETAST = 10
WAVELN = 0.574
OBSLAT = 52#
NSTEPS = 10000
ROBJ = 15#
'====================================================================
PARAMETERLIST:
'SCREEN 12
If Not NoShow Then
    picVDW.Cls
    picVDW.FontSize = 14
    picVDW.ForeColor = QBColor(13)
    picVDW.CurrentX = 28: picVDW.CurrentY = 113
    picVDW.CurrentX = 5: picVDW.CurrentX = 2: picVDW.Print "===============SPECIFICATION OF THE CASE================"
    picVDW.CurrentX = 7: picVDW.CurrentX = 2: picVDW.Print "01. HOBS = "; HOBS; "          Observer's eye height (m)"
    picVDW.CurrentX = 8: picVDW.CurrentX = 2: picVDW.Print "02. TGROUND = "; TGROUND; "  Temperature (K) at height = 0"
    picVDW.CurrentX = 9: picVDW.CurrentX = 2: picVDW.Print "03. HMAXT = "; HMAXT; "  Max. height (m) for curvature display (multiple of 100 m) "
    picVDW.CurrentX = 10: picVDW.CurrentX = 2: picVDW.Print "04. TLOW = "; TLOW; "        Show temperature profile from TLOW (K) till.."
    picVDW.CurrentX = 11: picVDW.CurrentX = 2: picVDW.Print "05. THIGH = "; THIGH; "       highest value (K) for which to show T-profile"
    picVDW.CurrentX = 12: picVDW.CurrentX = 2: picVDW.Print "06. PRESS0 = "; Press0; "  Atmospheric pressure (hPa) at h=0"
    picVDW.CurrentX = 13: picVDW.CurrentX = 2: picVDW.Print "07. RELHUM = "; RELHUM; "        Relative humidity (%) in troposphere"
    picVDW.CurrentX = 14: picVDW.CurrentX = 2: picVDW.Print "08. BETALO = "; BETALO; "         Lowest apparent altitude (arcmin)"
    picVDW.CurrentX = 15: picVDW.CurrentX = 2: picVDW.Print "09. BETAHI = "; BETAHI; "         Highest apparent altitude (arcmin)"
    picVDW.CurrentX = 16: picVDW.CurrentX = 2: picVDW.Print "10. BETAST = "; BETAST; "        Stepsize in apparent altitude (arcmin)"
    picVDW.CurrentX = 17: picVDW.CurrentX = 2: picVDW.Print "11. WAVELN = "; WAVELN; "      Wavelength (mu), 0.65=R,0.589=Y (Sodium),0.52=G"
    picVDW.CurrentX = 18: picVDW.CurrentX = 2: picVDW.Print "12. OBSLAT = "; OBSLAT; "       Latitude of observer ((degrees)"
    picVDW.CurrentX = 19: picVDW.CurrentX = 2: picVDW.Print "13. NSTEPS = "; NSTEPS; "      Number of steps up till HMAXT"
MODIFYPARAMETERS:
    picVDW.ForeColor = QBColor(14)
    picVDW.CurrentX = 25: picVDW.CurrentY = 20: picVDW.Print "Run with these parameters (Y/any other key) ?"
    'a$ = INPUT$(1)
    KeyPressed = 0
    Do While KeyPressed = 0
       DoEvents
    Loop
    If Chr(KeyPressed) = "Y" Or A$ = "y" Then
        GoTo ACCEPTPARAMETERS
    Else
    End If
    End If

'picVDW.CurrentX = 25: picVDW.CurrentX = 10: picVDW.Print "Input the parameter number (2 digits) to change"
'Color 15
'KeyPressed = 0
'Do While KeyPressed = 0
'   dovents
'Loop
''ANSWER$ = INPUT$(2)
'PARNUM = KeyPressed 'Val(ANSWER$)
'Loop Until PARNUM > 0 And PARNUM < 14
'Cls
'On PARNUM GoTo P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, P13
'P01:
'INPUT "HOBS = "; HOBS
'GoTo PARAMETERLIST
'P02:
'INPUT "TGROUND = "; TGROUND
'GoTo PARAMETERLIST
'P03:
'INPUT "HMAXT = "; HMAXT
'GoTo PARAMETERLIST
'P04:
'INPUT "TLOW = "; TLOW
'GoTo PARAMETERLIST
'P05:
'INPUT "THIGH = "; THIGH
'GoTo PARAMETERLIST
'P06:
'INPUT "PRESS0 = "; PRESS0
'GoTo PARAMETERLIST
'P07:
'INPUT "RELHUM = "; RELHUM
'GoTo PARAMETERLIST
'P08:
'INPUT "BETALO = "; BETALO
'GoTo PARAMETERLIST
'P09:
'INPUT "BETAHI = "; BETAHI
'GoTo PARAMETERLIST
'P10:
'INPUT "BETAST = "; BETAST
'GoTo PARAMETERLIST
'P11:
'INPUT "WAVELN = "; WAVELN
'GoTo PARAMETERLIST
'P12:
'INPUT "OBSLAT = "; OBSLAT
'GoTo PARAMETERLIST
'P13:
'INPUT "NSTEPS = "; NSTEPS
'GoTo PARAMETERLIST
ACCEPTPARAMETERS:
If Not LoopingAtmTracing And prjAtmRefMainfm.chkHgtProfile.Value = vbChecked And Val(txtHOBS.Text) = 0# Then
   Select Case MsgBox("You are using an alternative atmosphere and hugging the ground," _
                      & vbCrLf & "but the observer hieght is set to zero." _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to change the height?" _
                      , vbYesNo Or vbQuestion Or vbDefaultButton1, "Observer Height")
   
    Case vbYes
       txtHOBS = InputBox("Enter observer's height in meters", "Observer's height", 800.5)
    Case vbNo
   
   End Select
   End If

HOBS = Val(txtHOBS)
TGROUND = Val(txtTGROUND)
HMAXT = Val(txtHMAXT)
TLOW = Val(txtTLOW)
THIGH = Val(txtTHIGH)
Press0 = Val(txtPress0)
RELHUM = Val(txtRELHUM)
BETALO = Val(txtBETALO) * 60# 'convert to arc minutes
BETAHI = Val(txtBETAHI) * 60#
BETAST = Val(txtBETAST) * 60#
WAVELN = Val(txtKmin) * 0.001 'Val(txtWAVELN)
OBSLAT = Val(txtOBSLAT)
NSTEPS = Val(txtNSTEPS)
RELHUM = Val(txtHumid)
RELH = RELHUM / 100


'====CALCULATE HCROSS ========
If OptionSelby.Value = False And NoShow Then
   HCROSS = (TGROUND - 216.65) / 0.0065
   MinTemp = 9999999
   MaxTemp = -999999
   'create temperature layers for ray tracing plots
    For H = 0 To 85000 Step 100
        NumTemp = NumTemp + 1
        ELV(NumTemp) = H
        TMP(NumTemp) = fTEMP(H, -1, NumLayers)
        If TMP(NumTemp) > MaxTemp Then MaxTemp = TMP(NumTemp)
        If TMP(NumTemp) < MinTemp Then MinTemp = TMP(NumTemp)
    Next H
   End If

If Not NoShow Then
    picVDW.ForeColor = QBColor(12)
    picVDW.CurrentX = 28: picVDW.CurrentY = 15: picVDW.Print " press any key to proceed"
    picVDW.ForeColor = QBColor(15)
    KeyPressed = 0
    Do While KeyPressed = 0
       DoEvents
    Loop
    'SCREEN 0: WIDTH 80
    picVDW.Cls
    End If

'===================================================
If Trim(txtDir) = sEmpty And TempLoop Then
   Select Case MsgBox("You didn't select an output directory to write the files!" _
                      & vbCrLf & sEmpty _
                      & vbCrLf & "If you don't choose a directory, then the files" _
                      & vbCrLf & "will be written onto the default directory (see below):" _
                      & vbCrLf & sEmpty _
                      & vbCrLf & App.Path _
                      & vbCrLf & sEmpty _
                      & vbCrLf & "Proceed?" _
                      , vbYesNoCancel Or vbInformation Or vbDefaultButton1, "Output directory")
   
    Case vbYes
    
       FilePath = App.Path
   
    Case vbNo
       Exit Sub
    Case vbCancel
       progressfrm.Visible = False
       progressfrm2.Visible = False
       Screen.MousePointer = vbDefault
       Exit Sub
   End Select
Else
   FilePath = Trim(txtDir)
   End If
   
If TempLoop Then

   TempStart = txtST
   TempEnd = txtET
   TempStep = txtTS
   HgtStart = txtBHgt
   HgtEnd = txtEHgt
   HgtStep = txtSHgt
   cmpTLoop.Enabled = False
   prjAtmRefMainfm.progressfrm.Visible = True
   
    With prjAtmRefMainfm
      '------fancy progress bar settings---------
      .progressfrm2.Visible = True
      .picProgBar2.AutoRedraw = True
      .picProgBar2.BackColor = &H8000000B 'light grey
      .picProgBar2.DrawMode = 10
    
      .picProgBar2.FillStyle = 0
      .picProgBar2.ForeColor = &H808000   'dark green
      .picProgBar2.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
    With prjAtmRefMainfm
      '------fancy progress bar settings---------
      .progressfrm2.Visible = True
      .picProgBar3.AutoRedraw = True
      .picProgBar3.BackColor = &H8000000B 'light grey
      .picProgBar3.DrawMode = 10

      .picProgBar3.FillStyle = 0
      .picProgBar3.ForeColor = &H8000&     'dark green
      .picProgBar3.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
    
Else
    Close
    fileout1% = FreeFile
    Open App.Path & "\REF2017.OUT" For Output As #fileout1%
    Print #fileout1%, "BETAM, REFRAC, TRUALT, AIRDRY, AIRVAP"
    fileout3% = FreeFile
    Open App.Path & "\REF2017-ATM.OUT" For Output As #fileout3%
    Print #fileout3%, "height,temperature,dry-air pressure, vapor pressure, gravitational acceleration"
    fileout% = FreeFile
'    Open App.Path & "\TR_VDW_" & Trim(Str(TGROUND)) & "_" & Trim(Str(HOBS)) & "_" & Trim(Str(OBSLAT)) & ".dat" For Output As #fileout%
    Dim NumTc As Long
    NumTc = 0
   TempStart = TGROUND
   TempEnd = TGROUND
   TempStep = 1#
   HgtStart = HOBS
   HgtEnd = HOBS
   HgtStep = 1#
   End If
'===================================================
'===================Using DLL=========================
If chkAtmRefDll.Value = vbChecked And OptionSelby.Value = False Then  'as of now, can't use the dll for custom atmospheres -- TO DO
    Dim StartAng As Double, EndAng As Double, StepAng As Double
    Dim NAngles As Long, StepSize As Integer, RecordTLoop As Boolean
    Dim DistTo As Double, VAwo As Double, FileMode As Integer, Tol As Double, HOBSTR As Double
    Dim LastVA As Double
                
    Dim NewVA As Boolean, AAA(5) As Double, ref0, CurrentVA As Double
    
    DistTo = 0#
    VAwo = 0#
    Tol = 0#
    HOBSTR = 0#
    FileMode = 0
    Press0 = 1013
    LastVA = 0#
    
    StepSize = Val(prjAtmRefMainfm.txtHeightStepSize.Text)
    
    Screen.MousePointer = vbHourglass
   
    StatusMes = "Ray tracing in progress....press any key to abort next calculation"
    Call StatusMessage(StatusMes, 1, 0)
     
    NHloop = 0
    Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0) 'reset
    For HLoop = HgtStart To HgtEnd Step HgtStep
       Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 1, CLng(100# * (HLoop - HgtStart) / (HgtEnd - HgtStart + HgtStep)))

       txtHOBS = HLoop
       HOBS = HLoop
        
       NTLoop = 0
       Call UpdateStatus(prjAtmRefMainfm, picProgBar2, 0, 0) 'reset
       
        If chkTRef.Value = vbChecked Then
            RecordTLoop = True
        Else
            RecordTLoop = False
            End If
        
       For TLoop = TempStart To TempEnd Step TempStep
                                         
            Call UpdateStatus(prjAtmRefMainfm, picProgBar2, 1, CLng(100# * (TLoop - TempStart) / (TempEnd - TempStart + TempStep)))
    
            StartAng = CDbl(n_size / 2) / PPAM
            EndAng = -StartAng
            StepAng = 1 / PPAM
            NAngles = 2 * StartAng / StepAng + 1
            HUMID = RELHUM
            
'            If Dir(FilePath & "\TR_VDW_" & Trim(Str(TLoop)) & "_" & Trim(Str(HLoop)) & "_" & Trim(Str(OBSLAT)) & ".dat") <> SEmpty Then
            If chkSkipDone.Value = vbChecked And _
               Dir(FilePath & "\TR_VDW_" & Trim(Str(Fix(TLoop))) & "_" & Trim(Str(Fix(HLoop))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat") <> sEmpty Then
            
               If chkRefFile.Value = vbChecked Then
                    If Dir(App.Path & "\TR_VDW_Total_Refraction.dat") <> sEmpty Then
                       fileout% = FreeFile
                       Open App.Path & "\TR_VDW_Total_Refraction.dat" For Append As #fileout%
                    Else
                       fileout% = FreeFile
                       Open App.Path & "\TR_VDW_Total_Refraction.dat" For Output As #fileout%
                       End If
                       
                    filein% = FreeFile
                    
                    Open FilePath & "\TR_VDW_" & Trim(Str(Fix(TLoop))) & "_" & Trim(Str(Fix(HLoop))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat" For Input As #filein%
                    NewVA = True
                    Do Until EOF(filein%)
'                        sprintf(buff, "%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf%s", DIST, PATHLENGTH, H2, BETAM, BETA2 * 1000.0, REFRAC * (1000.0 * RADCON), "\n\0");
'                        sprintf(buff, "%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf%s", DIST, PATHLENGTH, -1000.0, BETAM, BETA2 * 1000.0, REFRAC * (1000.0 * RADCON), "\n\0");
                         Input #filein%, AAA(1), AAA(2), AAA(3), AAA(4), AAA(5)
                         
                         If NewVA Then
                            CurrentVA = AAA(3)
                            ref0 = AAA(5)
                            NewVA = False
                            If AAA(2) = -1000 Then Exit Do
                         Else
                         
                            If AAA(3) <> CurrentVA Then
                               NewVA = True
                               Print #fileout%, TLoop, HLoop, OBSLAT, CurrentVA, ref0
                               CurrentVA = AAA(3)
                               End If
                               
                            ref0 = AAA(5)
                            
                            If AAA(2) = -1000 Then Exit Do
                            
                            End If
                               
                    Loop
                    Close #filein%
                    Close fileout%
                    
                Else
                       'not recording total refraction, and already done ray tracing, so skip it,
                       End If
                       
            Else
            
                If FilePath = sEmpty Then FilePath = App.Path
            
                Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset
                ier = RayTracing(StartAng, EndAng, StepAng, LastVA, NAngles, _
                                 DistTo, VAwo, HOBSTR, Tol, FileMode, _
                                 HOBS, TLoop, HMAXT, FilePath, StepSize, _
                                 Press0, WAVELN, HUMID, OBSLAT, NSTEPS, _
                                 RecordTLoop, TempStart, TempEnd, AddressOf MyCallback)
                             
                End If
                             
            DoEvents 'look for messages
            If chkPause.Value = vbChecked Then
               chkPause.Value = vbUnchecked
               chkPause.Refresh
               DoEvents
               Select Case MsgBox("Do you want to abort?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Pause calculation")
                Case vbYes
                  GoTo exitcalc
                Case vbNo
               End Select
               End If
               
        Next TLoop
        
        If chkTRef.Value = vbChecked Then Close #filetemp%
        
    Next HLoop
    
exitcalc:
    KeyPressed = 0
    cmdVDW.Enabled = True
    cmdCalc.Enabled = True
    cmdMenat.Enabled = True
    cmdRefWilson.Enabled = True
    Screen.MousePointer = vbDefault
    Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset
    Call UpdateStatus(prjAtmRefMainfm, picProgBar2, 0, 0) 'reset
    Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0) 'reset
    progressfrm.Visible = False
    progressfrm2.Visible = False
    cmpTLoop.Enabled = True
    If chkRefFile.Value = vbChecked Then Exit Sub
    If FileMode = 0 Then GoTo DisplayResults
    Exit Sub
End If

'============End Using DLL======================
           
NHloop = 0
Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0)
For HLoop = HgtStart To HgtEnd Step HgtStep
   txtHOBS = HLoop
   HOBS = HLoop
   Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 1, CLng(100# * (HLoop - HgtStart) / (HgtEnd - HgtStart + HgtStep)))
           
'    If Dir(App.Path & "\TL_VDW_" & Trim(Str(TGROUND)) & "_" & Trim(Str(HOBS)) & "_" & Trim(Str(OBSLAT)) & ".dat") <> sempty Then
'       Kill App.Path & "\TL_VDW_" & Trim(Str(TGROUND)) & "_" & Trim(Str(HOBS)) & "_" & Trim(Str(OBSLAT)) & ".dat"
'       End If
'    filetemp% = FreeFile
'    Open App.Path & "\TL_VDW_" & Trim(Str(TGROUND)) & "_" & Trim(Str(HOBS)) & "_" & Trim(Str(OBSLAT)) & ".dat" For Output As #filetemp%

    NTLoop = 0
    Call UpdateStatus(prjAtmRefMainfm, picProgBar2, 0, 0)
    
    If chkTRef.Value = vbChecked Then
        'record the total atmospheric refraction as a function of temperature for any height
        filetemp% = FreeFile
        If OptionSelby.Value = False And chkDucting.Value = vbUnchecked Then
            FilNm = App.Path & "\TR_VDW_" & Trim(Str(Fix(TempStart))) & "-" & Trim(Str(Fix(TempEnd))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
            Open FilNm For Output As #filetemp%
        ElseIf chkDucting.Value = vbChecked Then
            FilNm = App.Path & "\TR_VDW_INV_" & Trim(Str(Fix(TempStart))) & "-" & Trim(Str(Fix(TempEnd))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
            Open FilNm For Output As #filetemp%
            'use van der Werf's formula to simulate an inversion
            Print #filetemp%, "Inversion layer added starting at" & Str(HL0) & " meters and ending at" & Str(TL0) & " meters, lapse rate:" & Str(LRL0) & " degrees K/m"
        ElseIf OptionSelby.Value = True Then
            FilNm = App.Path & "\TR_VDW_LAYERS_" & Trim(Str(Fix(TempStart))) & "-" & Trim(Str(Fix(TempEnd))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
            Open FilNm For Output As #filetemp%
            Print #filetemp%, "Custom atmospheric layer file type" & Str(AtmNumber); " used with" & Str(NumLayers) & " layers"
            End If
        End If
    
    For TLoop = TempStart To TempEnd Step TempStep
        Call UpdateStatus(prjAtmRefMainfm, picProgBar2, 1, CLng(100# * (TLoop - TempStart) / (TempEnd - TempStart + TempStep)))
        txtTGROUND = TLoop
        TGROUND = TLoop
        txtPress0 = Press0
        NumTc = 0
        If OptionSelby.Value = False Then HCROSS = (TGROUND - 216.65) / 0.0065

        If Dir(App.Path & "\TR_VDW_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat") <> sEmpty Then
           Kill App.Path & "\TR_VDW_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
           End If
        fileout% = FreeFile
        If OptionSelby.Value = False And chkDucting.Value = vbUnchecked Then
            FilNm = App.Path & "\TR_VDW_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
            Open FilNm For Output As #fileout%
        ElseIf chkDucting.Value = vbChecked Then
            FilNm = App.Path & "\TR_VDW_INV_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
            Open FilNm For Output As #fileout%
            'use van der Werf's formula to simulate an inversion
            Print #fileout%, "Inversion layer added starting at" & Str(HL0) & " meters and ending at" & Str(TL0) & " meters, lapse rate:" & Str(LRL0) & " degrees K/meter"
        ElseIf OptionSelby.Value = True Then
            FilNm = App.Path & "\TR_VDW_LAYERS_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
            Open FilNm For Output As #fileout%
            Print #fileout%, "Custom atmospheric layer file type" & Str(AtmNumber); " used with" & Str(NumLayers) & " layers"
            End If
        

'============LIST OF DERIVED CONSTANTS==================================
'Dim OLAT As Double, GRAVC As Double, BD As Double, BW As Double, s2 As Double, AD As Double, AW As Double, MAXIND As Double
OLAT = OBSLAT * 60 * RADCON
GRAVC = 9.780356 * (1 + 0.0052885 * (Sin(OLAT)) ^ 2 - 0.0000059 * (Sin(2 * OLAT)) ^ 2)           'Gravitat. const.
BD = GRAVC * AMASSD / RBOLTZ 'Dry air exponent
BW = GRAVC * AMASSW / RBOLTZ 'Water exponent
s2 = 1 / WAVELN ^ 2
'CIDDOR'S FORMULAS FOR DRY AIR AND WATER VAPOUR
AD = 0.00000001 * (5792105# / (238.0185 - s2) + 167917# / (57.362 - s2)) * 288.15 / 1013.25
AW = 0.00000001022 * (295.235 + 2.6422 * s2 - 0.03238 * s2 ^ 2 + 0.004028 * s2 ^ 3) * 293.15 / 13.33
MAXIND = HMAXT + 1

StatusMes = "Setting up pressure lookup tables...please wait"
Call StatusMessage(StatusMes, 1, 0)

'==================================================================
'modification -- 040520 -- if using layered atmosphere with pressure such as sondes, then
'skip filling in the arrays, and use the measured pressures
'If PRSR(0) > 500 Then GoTo SkipPressArrays

'==FILL ARRAY PRESSD1 (PARTIAL PRESSURE OF DRY AIR)==
'== IN STEPS OF 1 METER========
Dim P1 As Double, P2 As Double, DP2DH As Double, FK1 As Double
Dim FK2 As Double, FK3 As Double, FK4 As Double, HSTEP As Double
'Dim T As Double, fGRAVRATH As Double, fVAPORH As Double, fDVAPDTH As Double, fDTDHH As Double

'filtstout = FreeFile 'diagnostics
'Open App.Path & "\press_test.txt" For Output As #filtstout

H = 0#
PRESSD1(1) = Press0 - RELH * fVAPOR(0#, -1, NumLayers)
For i = 1 To (31000 + 15) Step 1
'    If i = 31000 Then
'       ccc = 1
'       End If
    If OptionSelby Then 'used measured pressure instead
       found% = 0

       If (i + 1) * 0.001 < ELV(0) Then
          PRESSD1(i + 1) = PRSR(0)
          found% = 1
       Else
            For j = 0 To NumLayers - 2
'                If ELV(j) = 0 Then
'                   ccc = 1
'                   End If
               If (i + 1) * 0.001 >= ELV(j) And (i + 1) * 0.001 < ELV(j + 1) Then
                  If ELV(j + 1) <> ELV(j) Then
                     PRESSD1(i + 1) = ((PRSR(j + 1) - PRSR(j)) / (ELV(j + 1) - ELV(j))) * ((i + 1) * 0.001 - ELV(j)) + PRSR(j)
                     found% = 1
                     Exit For
                  Else
                     PRESSD1(i + 1) = PRSR(j)
                     found% = 1
                     Exit For
                     End If
                  End If
            Next j
          End If
          If found% = 1 Then GoTo NextPressStep
'          PressTst = PRESSD1(i + 1) 'diagnostics
       End If
    '===Fill PRESSD1. I=1 -> H=0, I=2 -> h=1 m. etc..===
    '===INTEGRATION BY 4TH ORDER RUNGE-KUTTA===
    HSTEP = 1#

    'STEP 1
    P1 = PRESSD1(i)
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK1 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 1
    'STEP 2
    H = (i - 1) / 1# + HSTEP / 2

    P1 = PRESSD1(i) + FK1 * HSTEP / 2
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK2 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 2
    'STEP 3
    H = (i - 1) / 1# + HSTEP / 2

    P1 = PRESSD1(i) + FK2 * HSTEP / 2
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK3 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 3
    'STEP 4
    H = (i - 1) / 1# + HSTEP

    P1 = PRESSD1(i) + FK3 * HSTEP
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK4 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 4
    PRESSD1(i + 1) = PRESSD1(i) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
NextPressStep:
'    Print #filtstout, PressTst, PRESSD1(i + 1) 'diagnostics
Next i
'Close #filtstout

'===FIND PDM1 AT -1 METER===
If OptionSelby.Value = True Then
   'extrapolate to -1 meters
'   PDM01 = ((PRSR(1) - PRSR(0)) / (ELV(1) - ELV(0))) * (-1# * 0.001 - ELV(0)) + PRSR(0)
   PDM01 = PRSR(0)
Else
HSTEP = -1#
'STEP 1
H = 0#
P1 = PRESSD1(1)
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK1 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 1
'STEP 2
H = HSTEP / 2
P1 = PRESSD1(1) + FK1 * HSTEP / 2
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK2 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 2
'STEP 3
H = HSTEP / 2
P1 = PRESSD1(1) + FK2 * HSTEP / 2
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK3 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 3
'STEP 4
H = HSTEP
P1 = PRESSD1(1) + FK3 * HSTEP
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK4 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 4
PDM01 = PRESSD1(1) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
'==================================================================
End If
'=======END OF STORAGE PRESSURE ARRAY PRESSD1 FOR DRY AIR========

'==FILL ARRAY PRESSD2 (PARTIAL PRESSURE OF DRY AIR)==
'== IN STEPS OF 10 METER========
If OptionSelby.Value = False Then
    NumTemp = 0
    ELV(NumTemp) = 0
    PRSR(NumTemp) = PRESSD1(1)
    TMP(NumTemp) = fTEMP(0, -1, NumLayers)
    If TMP(NumTemp) < MinTemp Then MinTemp = TMP(NumTemp)
    If TMP(NumTemp) > MaxTemp Then MaxTemp = TMP(NumTemp)
    NumTemp = NumTemp + 1
    End If

PRESSD2(1) = PRESSD1(1)
Dim I2LOW As Long
I2LOW = CInt(HMAXP1 / 10)

If OptionSelby.Value = False Then
    MinTemp = TMP(0)
    MaxTemp = TMP(0)
    
    For i = 0 To I2LOW Step 1
        PRESSD2(i + 1) = PRESSD1(10 * i + 1)
        ELV(NumTemp) = i * 10
        PRSR(NumTemp) = PRESSD2(i + 1)
        TMP(NumTemp) = fTEMP(ELV(NumTemp), -1, NumLayers)
        If TMP(NumTemp) < MinTemp Then MinTemp = TMP(NumTemp)
        If TMP(NumTemp) > MaxTemp Then MaxTemp = TMP(NumTemp)
        NumTemp = NumTemp + 1
        NNN = NumTemp
    Next i
Else
    For i = 0 To I2LOW Step 1
        PRESSD2(i + 1) = PRESSD1(10 * i + 1)
    Next i
    
    End If
'filtstout = FreeFile 'diagnostics
'Open App.Path & "\press_test.txt" For Output As #filtstout
    
For i = I2LOW To (HLIMIT / 10 + 5) Step 1
    '===Fill PRESSD2. I=1 -> H=0, I=2 -> h=1 m. etc..===
    '===INTEGRATION BY 4TH ORDER RUNGE-KUTTA===
    
    If OptionSelby Then 'used measured pressure instead
       found% = 0
       If (i + 10) * 0.01 < ELV(0) Then
          PRESSD2(i + 1) = PRSR(0)
          found% = 1
       Else
            For j = 0 To NumLayers - 2
               If (i + 10) * 0.01 >= ELV(j) And (i + 10) * 0.01 < ELV(j + 1) Then
                  If ELV(j + 1) <> ELV(j) Then
                     PRESSD2(i + 1) = ((PRSR(j + 1) - PRSR(j)) / (ELV(j + 1) - ELV(j))) * ((i + 10) * 0.01 - ELV(j)) + PRSR(j)
                     found% = 1
                     Exit For
                  Else
                     PRESSD2(i + 1) = PRSR(j)
                     found% = 1
                     Exit For
                     End If
                  End If
            Next j
'            PressTst = PRESSD2(i + 1) 'diagnostics

          End If
          If found% = 1 Then GoTo NextPressStep2
       End If
       
    HSTEP = 10#
    'STEP 1
    H = (i - 1) * 10
    P1 = PRESSD2(i)
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK1 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 1
    'STEP 2
    H = (i - 1) * 10 + HSTEP / 2
    P1 = PRESSD2(i) + FK1 * HSTEP / 2
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK2 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 2
    'STEP 3
    H = (i - 1) * 10 + HSTEP / 2
    P1 = PRESSD2(i) + FK2 * HSTEP / 2
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK3 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 3
    'STEP 4
    H = (i - 1) * 10 + HSTEP
    P1 = PRESSD2(i) + FK3 * HSTEP
    P2 = RELH * fVAPOR(H, -1, NumLayers)
    DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
    FK4 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
    'END STEP 4
    PRESSD2(i + 1) = PRESSD2(i) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
    
    If NumTemp < MaxViewSteps& And H <> ELV(NumTemp - 1) And OptionSelby.Value = False Then
        ELV(NumTemp) = H
        PRSR(NumTemp) = PRESSD2(i + 1)
        TMP(NumTemp) = fTEMP(H, -1, NumLayers)
        If TMP(NumTemp) < MinTemp Then MinTemp = TMP(NumTemp)
        If TMP(NumTemp) > MaxTemp Then MaxTemp = TMP(NumTemp)
        NumTemp = NumTemp + 1
        NNN = NumTemp
        End If
'
NextPressStep2:
'    Print #filtstout, PressTst, PRESSD2(i + 1) 'diagnostics
Next i
'Close #filtstout

'SkipPressArrays:
NumTemp = NNN
If OptionSelby Then NumTemp = NNN + 1

''pressures
'fileout% = FreeFile
'Open App.Path & "\vdW-Pressures-VP-4.txt" For Output As #fileout%
'HUMD = 1#
'For H = 1 To 100000
'
'    If H < HMAXP1 Then
'        PP = fFNDPD1(H, PRESSD1, -1, NumLayers) + HUMD * fVAPOR(H, -1, NumLayers)
'    Else
'        PP = fFNDPD2(H, PRESSD2, -1, NumLayers) + HUMD * fVAPOR(H, -1, NumLayers)
'    End If
'
'    If H = 1 Then
'       Print #fileout%, H, PP
'       End If
'
'    If H Mod 200 = 0 Then
'       Print #fileout%, H, PP
'       End If
'
'Next H
'Close #fileout%

NNN = NumTemp

StatusMes = "Plotting the pressure and temperature layers...."
Call StatusMessage(StatusMes, 1, 0)

'-------------fill in mscharts array and plot it--------------
MultFac = 0.001
If OptionSelby Then MultFac = 1#
If NNN > 1000 Then
     ReDim TransferCurve(1 To NNN / 10, 1 To 2) As Variant
     filetmp = FreeFile 'record the temperature profile as function of height
     Open App.Path & "\VDW-INV.dat" For Output As #filetmp
     Print #filetmp, Int(NNN / 10) - 1
     For j = 1 To NNN / 10
        TransferCurve(j, 1) = " " & CStr(ELV((j - 1) * 10) * MultFac)
    '         TransferCurve(J, 2) = ELV(J - 1) * 0.001
        TransferCurve(j, 2) = TMP((j - 1) * 10)
        Print #filetmp, ELV((j - 1) * 10), TMP((j - 1) * 10)
     Next j
     Close #filetmp
     
     NumTemp = NNN / 10
     
Else 'no need to use less resolution
     ReDim TransferCurve(1 To NNN, 1 To 2) As Variant
     filetmp = FreeFile 'record the temperature profile as function of height
     Open App.Path & "\VDW-INV.dat" For Output As #filetmp
     Print #filetmp, NNN - 1
     For j = 1 To NNN
        TransferCurve(j, 1) = " " & CStr(ELV((j - 1)) * MultFac)
    '         TransferCurve(J, 2) = ELV(J - 1) * 0.001
        TransferCurve(j, 2) = TMP(j - 1)
        Print #filetmp, ELV(j - 1), TMP(j - 1)
     Next j
     Close #filetmp
     
     NumTemp = NNN
   End If
 
 With MSChartTemp
   .chartType = VtChChartType2dLine
   .RandomFill = False
   .Title = "Atmospheric Temperature (degrees K)"
   .ShowLegend = True
   .ChartData = TransferCurve
    With .Plot
        With .Wall.Brush
            .Style = VtBrushStyleSolid
            .FillColor.Set 255, 255, 255
        End With
        With .Axis(VtChAxisIdX)
            .ValueScale.Auto = False
            .AxisTitle = "Elevation (km)"
        End With
        With .Axis(VtChAxisIdX).CategoryScale
            .Auto = False
            .DivisionsPerLabel = NumTemp * 0.1
            .DivisionsPerTick = NumTemp * 0.1
            .LabelTick = True
        End With
        .AutoLayout = True
        With .Axis(VtChAxisIdY)
            .AxisTitle = "Temperature (degrees K)"
            With .ValueScale
                .Auto = False
                .MajorDivision = 10
                .Maximum = MaxTemp * 1.01
                .Minimum = MinTemp
            End With
        End With
'        With .Axis(VtChAxisIdY2)
'            .AxisTitle = "Elevation (km)"
'            With .ValueScale
'                .Auto = False
'                .MajorDivision = 10
'            End With
'        End With
        With .SeriesCollection(1)
            .SeriesMarker.Show = False
            .LegendText = "Temperature (K)"
            With .Pen
                .Width = ScaleX(1, vbPixels, vbTwips)
                .VtColor.Set 0, 0, 255 'blue
            End With
            With .DataPoints(-1).Marker
                .Style = VtMarkerStyleDiamond
            End With
        End With
    End With
End With
 
If NNN > 1000 Then
   For j = 1 To NNN / 10
     TransferCurve(j, 2) = PRSR((j - 1) * 10)
   Next j
Else
   For j = 1 To NNN
     TransferCurve(j, 2) = PRSR((j - 1) * 10)
   Next j
   End If
 
 With MSChartPress
   .chartType = VtChChartType2dLine
   .RandomFill = False
   .Title = "Atmospheric Pressure (mbar)"
   .ShowLegend = True
   .ChartData = TransferCurve
    With .Plot
        With .Wall.Brush
            .Style = VtBrushStyleSolid
            .FillColor.Set 255, 255, 255
        End With
        With .Axis(VtChAxisIdX)
            .AxisTitle = "Elevation (km)"
        End With
        With .Axis(VtChAxisIdX).CategoryScale
            .Auto = False
            .DivisionsPerLabel = NumTemp * 0.1
            .DivisionsPerTick = NumTemp * 0.1
            .LabelTick = True
        End With
        With .Axis(VtChAxisIdY)
            .AxisTitle = "Pressure (mbar)"
            With .ValueScale
                .Auto = False
                .MajorDivision = 10
                .Maximum = PRSR(0) * 1.01
                .Minimum = 0
            End With
        End With
'        With .Axis(VtChAxisIdY2)
'            .AxisTitle = "Elevation (km)"
'            With .ValueScale
'                .Auto = False
'                .MajorDivision = 10
'            End With
'        End With
        With .SeriesCollection(1)
            .SeriesMarker.Show = False
            .LegendText = "Pressure (mbar)"
            With .Pen
                .Width = ScaleX(1, vbPixels, vbTwips)
                .VtColor.Set 0, 0, 255 'blue
            End With
            With .DataPoints(-1).Marker
                .Style = VtMarkerStyleDiamond
            End With
        End With
    End With
End With

 '-----------------------------------------------------
'If PRSR(0) > 500 Then GoTo NextStep

'===FIND PDM10 AT -10 METER===
If OptionSelby.Value = True Then
   'extrapolate to -1 meters
'   PDM01 = ((PRSR(1) - PRSR(0)) / (ELV(1) - ELV(0))) * (-0.01 - ELV(0)) + PRSR(0)
   PDM01 = PRSR(0)
Else
HSTEP = -10#
'STEP 1
H = 0#
P1 = PRESSD2(1)
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK1 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 1
'STEP 2
H = HSTEP / 2
P1 = PRESSD2(1) + FK1 * HSTEP / 2
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK2 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 2
'STEP 3
H = HSTEP / 2
P1 = PRESSD2(1) + FK2 * HSTEP / 2
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK3 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 3
'STEP 4
H = HSTEP
P1 = PRESSD2(1) + FK3 * HSTEP
P2 = RELH * fVAPOR(H, -1, NumLayers)
DP2DH = RELH * fDVAPDT(H, -1, NumLayers) * fDTDH(H, -1, NumLayers)
FK4 = -DP2DH - BD * P1 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers) - BW * P2 * fGRAVRAT(H) / fTEMP(H, -1, NumLayers)
'END STEP 4
PDM10 = PRESSD2(1) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
'==================================================================
End If
'=======END OF STORAGE PRESSURE ARRAY FOR DRY AIR========
'NextStep:
'===================Temperature profile====================
If Not NoShow Then
    'SCREEN 12
    picVDW.Cls
    picVDW.FontSize = FSize
    picVDW.ForeColor = QBColor(15)
    picVDW.Scale (TLOW - 20, HLIMIT + 4000)-(THIGH + 20, -4000)
    picVDW.Line (TLOW, 0)-(THIGH, HLIMIT), QBColor(1), B
    picVDW.Line (TLOW, HCROSS)-(THIGH, HCROSS), QBColor(1)
    picVDW.Line (TLOW, 20000)-(THIGH, 20000), QBColor(1)
    picVDW.Line (TLOW, 32000)-(THIGH, 32000), QBColor(1)
    picVDW.Line (TLOW, 47000)-(THIGH, 47000), QBColor(1)
    picVDW.Line (TLOW, 51000)-(THIGH, 51000), QBColor(1)
    picVDW.Line (TLOW, 71000)-(THIGH, 71000), QBColor(1)
    picVDW.Line (TLOW, 85000)-(THIGH, 85000), QBColor(1)
    'picVDW.CurrentX = 24: picVDW.CurrentY = 100: picVDW.Print "TEMPERATURE PROFILE :  T=" & Str(Int(TLOW)) & "K -" & Str(Int(THIGH)) & "K"
    
    picVDW.CurrentX = 3: picVDW.CurrentY = 15: picVDW.Print "TEMPERATURE PROFILE :  T=" & Str(Int(TLOW)) & "K -" & Str(Int(THIGH)) & "K"
    picVDW.CurrentX = 5: picVDW.CurrentY = 20000: picVDW.Print "20 km"
    picVDW.CurrentX = 5: picVDW.CurrentY = 32000: picVDW.Print "32 km"
    picVDW.CurrentX = 5: picVDW.CurrentY = 47000: picVDW.Print "47 km"
    picVDW.CurrentX = 5: picVDW.CurrentY = 51000: picVDW.Print "51 km"
    picVDW.CurrentX = 5: picVDW.CurrentY = 71000: picVDW.Print "71 km"
    picVDW.CurrentX = 5: picVDW.CurrentY = 85000: picVDW.Print "85 km"
    picVDW.CurrentX = 5: picVDW.CurrentY = 11000: picVDW.Print "height troposphere:"
    picVDW.CurrentX = 35: picVDW.CurrentY = 11000: picVDW.Print Str(Int(HCROSS) / 1000) & "km"
    For i = 1 To HLIMIT Step 100
        H = i - 1
        R1 = fTEMP(H, -1, NumLayers)
        picVDW.Circle (R1, H), 0.01, QBColor(12)
    Next i
    '==================================================================
    picVDW.ForeColor = QBColor(14)
    picVDW.FontSize = 14
    picVDW.CurrentX = 0.5 * (THIGH - TLOW) - 50: picVDW.CurrentY = 2: picVDW.Print " press any key to proceed"
    picVDW.ForeColor = QBColor(15)
    KeyPressed = 0
    Do While KeyPressed = 0
       DoEvents
    Loop
    'in$ = INPUT$(1)
    'SCREEN 0: WIDTH 80
    picVDW.Cls
    End If
'=================R_curve/R-Earth======================

'The ratio R_curve/R-Earth will be shown on a scale -5 to 5 (horizontal)
'versus elevation from 0 to the maximum tabulated elevation HMAXT.

If Not NoShow Then
    'SCREEN 12
    picVDW.Cls
    picVDW.FontSize = FSize
    picVDW.Scale (-11, HMAXT * 1.1)-(11, -0.1 * HMAXT)
    picVDW.Line (-10, 0)-(10, HMAXT), QBColor(1), B
    picVDW.CurrentX = -8: picVDW.CurrentY = -2: picVDW.Print "-8"
    picVDW.CurrentX = -7: picVDW.CurrentY = -2: picVDW.Print "-7"
    picVDW.CurrentX = -6: picVDW.CurrentY = -2: picVDW.Print "-6"
    picVDW.CurrentX = -5: picVDW.CurrentY = -2: picVDW.Print "-5"
    picVDW.CurrentX = -4: picVDW.CurrentY = -2: picVDW.Print "-4"
    picVDW.CurrentX = -3: picVDW.CurrentY = -2: picVDW.Print "-3"
    picVDW.CurrentX = -2: picVDW.CurrentY = -2: picVDW.Print "-2"
    picVDW.CurrentX = -1: picVDW.CurrentY = -2: picVDW.Print "-1"
    picVDW.CurrentX = 0: picVDW.CurrentY = -2: picVDW.Print "0"
    picVDW.CurrentX = 1: picVDW.CurrentY = -2: picVDW.Print "1"
    picVDW.CurrentX = 2: picVDW.CurrentY = -2: picVDW.Print "2"
    picVDW.CurrentX = 3: picVDW.CurrentY = -2: picVDW.Print "3"
    picVDW.CurrentX = 4: picVDW.CurrentY = -2: picVDW.Print "4"
    picVDW.CurrentX = 5: picVDW.CurrentY = -2: picVDW.Print "5"
    picVDW.CurrentX = 6: picVDW.CurrentY = -2: picVDW.Print "6"
    picVDW.CurrentX = 7: picVDW.CurrentY = -2: picVDW.Print "7"
    picVDW.CurrentX = 8: picVDW.CurrentY = -2: picVDW.Print "8"
    For i = -9 To 9 Step 1
        picVDW.Line (i, 0)-(i, HMAXT), QBColor(1)
    Next i
    picVDW.CurrentX = 2: picVDW.CurrentY = 30: picVDW.Print "Horizontal  R_curv/R_Earth for a horizontal ray, Vertical: H=" & Str(0) & "-" & Str(Int(HMAXT)) & "m"
    
    For i = 1 To HMAXT Step 1
        H = i - 1
        R1 = 1 / (fRCINV(H, PRESSD1, PRESSD2, -1, NumLayers) * Rearth)
        picVDW.Circle (R1, H), 0.01, QBColor(12)
    Next i
    picVDW.ForeColor = QBColor(14)
    picVDW.FontSize = 14
    picVDW.CurrentX = 0: picVDW.CurrentY = -10: picVDW.Print " press any key to proceed"
    picVDW.ForeColor = QBColor(15)
    KeyPressed = 0
    Do While KeyPressed = 0
       DoEvents
    Loop
    End If
    
'in$ = INPUT$(1)
'SCREEN 0: WIDTH 80
'PRINT OUTPUT FILE #3, DATA ON THE ATMOSPHERE
If Not TempLoop Then
    Dim PDRY As Double, PVAP As Double, GRAVH As Double
    For H = 0 To 80000 Step 1000
        PDRY = fFNDPD2(H, PRESSD2, -1, NumLayers)
        PVAP = RELH * fTEMP(H, -1, NumLayers)
        GRAVH = GRAVC * fGRAVRAT(H)
        Print #fileout3%, Format(Str(H / 1000), " ####0.0#####"), Format(Str(fTEMP(H, -1, NumLayers)), " ####0.0#####"), Format(Str(PDRY), " ####0.0#####"), Format(Str(PVAP), " ####0.0#####"), Format(Str(GRAVH), " ####0.0#####")
    Next H
    Close #fileout3%
    End If

'=============SETUP GRAPHICS PART================================

If Not NoShow Then
    'SCREEN 12
    picVDW.Cls
    Dim LRANGE As Long, LDKM As Long
    LRANGE = 1.25 * fGUESSL(BETALO / 60#, HLIMIT)
    LDKM = Int(LRANGE / 1000)
    picVDW.Scale (-LRANGE * 0.01, HLIMIT + 2000)-(LRANGE * 1.01, -2000)
    picVDW.Line (0, 0)-(LRANGE, HLIMIT), QBColor(1), B
    
    picVDW.CurrentX = 2: picVDW.CurrentY = 2: picVDW.Print "Paths of the rays. Horizontal distance 0 - " & Str(LDKM) & "km, Vertical:  0 -100 km"
    End If

'===================CALCULATION====================================
If Not NoShow Then
    picVDW.ForeColor = QBColor(14)
    picVDW.CurrentX = LRANGE * 0.5 - 2000: picVDW.CurrentY = 20: picVDW.Print "To interrupt, hit the spacebar"
    picVDW.ForeColor = QBColor(15)
    End If
    
NumBet = 0

If Not NoShow Then
    picVDW.Scale (-LRANGE * 0.01, HLIMIT + 2000)-(LRANGE * 1.01, -2000)
    End If
    
StatusMes = "Ray tracing commencing...please wait"
Call StatusMessage(StatusMes, 1, 0)

Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset
'
'zero refraction arrays
'For II = 0 To 82
'   For jj = 0 To MaxViewSteps& - 1
'    ALFA(II, jj) = 0
'    ALFT(II, jj) = 0
'   Next jj
'Next II

Dim nloop As Long

'For BETAM = BETALO To BETAHI Step BETAST
For jstep = 1 To n_size + 1

    ALFA(KMIN, jstep) = (CDbl(n_size / 2 - (jstep - 1)) / PPAM)
    BETAM = ALFA(KMIN, jstep)

    DPATH = fGUESSP(BETAM, HMAXT) / NSTEPS
    
    If Not NoShow Then
        picVDW.FontSize = 14
        picVDW.ForeColor = QBColor(15)
        picVDW.FontBold = False
        picVDW.CurrentX = LRANGE * 0.75: picVDW.CurrentY = 2: picVDW.Print String(150, Chr(219))
        picVDW.CurrentX = LRANGE * 0.75: picVDW.CurrentY = 2: picVDW.Print "Step size (m) = " & String(100, Chr(219))
        picVDW.ForeColor = QBColor(13)
        picVDW.FontBold = True
        picVDW.CurrentX = LRANGE * 0.75: picVDW.CurrentY = 2: picVDW.Print "Step size (m) = " & Str(DPATH)
        End If
        
    NumBet = NumBet + 1
    If NumBet > 14 Then NumBet = 1
    Dist = 0#
    REFRAC = 0#
    AIRDRY = 0#
    AIRVAP = 0#
    PHI1 = 0#
    BETA1 = BETAM * RADCON
    R1 = Rearth + HOBS
    H1 = HOBS
    PATHLENGTH = 0#
    
    If Not NoShow Then
        picVDW.PSet (Dist, H1), QBColor(NumBet)
        picVDW.Refresh
        End If
        
'    Pathh = -DPATH
    Pathh = 0#
    '===============================
    'DO-LOOP OVER PATH
'    DoEvents
'    If Chr(KeyPressed) = Chr$(32) Then GoTo ENDOFCALCULATION
    nloop = 0
    Do
'        DoEvents
'        If Chr(KeyPressed) = Chr$(32) Then GoTo ENDOFCALCULATION
'        Pathh = Pathh + DPATH
        '
        ' FOURTH ORDER RUNGE-KUTTA
        ' THE THREE COUPLED FIRST ORDER DIFFERENTIAL EQUATIONS ARE:
        ' 1)  dPHI/dPATH=cos(BETA)/R
        ' 2)  dR/dPATH = sin(BETA)
        ' 3)  dBETA/dPATH = cos(BETA)[1//R+(1/n) dn/dR]
        ' WITH (1/n) [dn/dR]=f.RCINV(H)
        '
        ' STEP 1
        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
        '
        '<<<<<<<<note: Beta is the complement of the local zenith angle, i.e., the dip>>>>>>>>>>>>>>>>
        
        FKP1 = Cos(BETA1) / R1
        FKR1 = Sin(BETA1)
        FKB1 = Cos(BETA1) * (1 / R1 + fRCINV(H1, PRESSD1, PRESSD2, Dist, NumLayers))
        FKAD1 = fDENSDRY(H1, PRESSD1, PRESSD2, Dist, NumLayers)
        FKAV1 = fDENSVAP(H1, Dist, NumLayers)
        '
        'END OF FIRST STEP
        '
        'STEP 2
        PHINEW = PHI1 + FKP1 * DPATH / 2
        RNEW = R1 + FKR1 * DPATH / 2
        BETANEW = BETA1 + FKB1 * DPATH / 2
        HNEW = RNEW - Rearth 'elevation halfway step
        '
        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
        '
        FKP2 = Cos(BETANEW) / RNEW
        FKR2 = Sin(BETANEW)
        FKB2 = Cos(BETANEW) * (1 / RNEW + fRCINV(HNEW, PRESSD1, PRESSD2, Dist, NumLayers)) '<<<<<<<<<<<<<<<
        FKAD2 = fDENSDRY(HNEW, PRESSD1, PRESSD2, Dist, NumLayers)
        FKAV2 = fDENSVAP(HNEW, Dist, NumLayers)
        '
        'END OF SECOND STEP
        '
        'STEP 3
        PHINEW = PHI1 + FKP2 * DPATH / 2
        RNEW = R1 + FKR2 * DPATH / 2
        BETANEW = BETA1 + FKB2 * DPATH / 2
        HNEW = RNEW - Rearth 'elevation halfway step
        '
        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
        '
        FKP3 = Cos(BETANEW) / RNEW
        FKR3 = Sin(BETANEW)
        FKB3 = Cos(BETANEW) * (1 / RNEW + fRCINV(HNEW, PRESSD1, PRESSD2, Dist, NumLayers))
        FKAD3 = fDENSDRY(HNEW, PRESSD1, PRESSD2, Dist, NumLayers)
        FKAV3 = fDENSVAP(HNEW, Dist, NumLayers)

        '
        'END OF THIRD STEP
        '
        'STEP 4
        PHINEW = PHI1 + FKP3 * DPATH
        RNEW = R1 + FKR3 * DPATH
        BETANEW = BETA1 + FKB3 * DPATH
        HNEW = RNEW - Rearth
        H = HNEW 'elevation at full step
        '
        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
        '
'        If BETA2 * 1000 >= 6.3427 Then
'           ccc = 1
'           End If
        FKP4 = Cos(BETANEW) / RNEW
        FKR4 = Sin(BETANEW)
        FKB4 = Cos(BETANEW) * (1 / RNEW + fRCINV(HNEW, PRESSD1, PRESSD2, Dist, NumLayers))
        FREF4 = Cos(BETANEW) * fRCINV(HNEW, PRESSD1, PRESSD2, Dist, NumLayers)
        FKAD4 = fDENSDRY(HNEW, PRESSD1, PRESSD2, Dist, NumLayers)
        FKAV4 = fDENSVAP(HNEW, Dist, NumLayers)
        '
        'END OF FOURTH AND FINAL STEP
        '
        'FIND R2 AND PHI2
        PHI2 = PHI1 + (FKP1 + 2 * FKP2 + 2 * FKP3 + FKP4) * DPATH / 6
        R2 = R1 + (FKR1 + 2 * FKR2 + 2 * FKR3 + FKR4) * DPATH / 6
        BETA2 = BETA1 + (FKB1 + 2 * FKB2 + 2 * FKB3 + FKB4) * DPATH / 6
        AIRDRY = AIRDRY + (FKAD1 + 2 * FKAD2 + 2 * FKAD3 + FKAD4) * DPATH / 6
        AIRVAP = AIRVAP + (FKAV1 + 2 * FKAV2 + 2 * FKAV3 + FKAV4) * DPATH / 6
        H2 = R2 - Rearth
        DREFR = -BETA2 + BETA1 + PHI2 - PHI1
        'Stop this ray if it hits the ground, or if it seems to never end
        'as may occur for a Novaya-Zemlya atmosphere.
        
'        If H2 >= 800 Then
'           ccc = 1
'           End If
        
'        'calculate approx pathlength (from Wikipedia's article on "Air Mass Astronomy", where cos(z) -> sin(BETA1) )
'        'the pathlength will only be a good approximation for small step sizes in height
        If H2 > 0 Then
           PATHLENGTH = PATHLENGTH + Sqr(((Rearth + H2) * Sin(BETA1)) ^ 2# + 2# * Rearth * (Abs(H2 - H1)) + Abs(H2 * H2 - H1 * H1)) - (Rearth + H1) * Sin(BETA1)
           End If
           
        Pathh = Pathh + fGUESSP(BETA2 / RADCON, Abs(H2 - H1)) 'estimate of distance along the ray, which can be smaller than the
                                                              'distance along the surface of the Earth if the ray's radius of curviature
                                                              'is smaller's than the radius of the Earth, Beta2 is the local zenith angle

        'change of concept
        'the "Path" calculated in this program is in fact an approximate total path length of the ray in the atmosphere
        'and the approx above is wrong since it needs to use path along the earth instead
        'from now on PATHLENGTH will be used for distance along the earth.
                
        If nloop = 0 Then
'           If Not TempLoop Then Print #fileout%, 0, FormatNumber(HOBS, 1, vbFalse, vbFalse, vbFalse)
'           If Not TempLoop Then Write #fileout%, 0, HOBS, ALFA(KMIN, jstep)
            TRUALT = BETAM - REFRAC
           Print #fileout%, Format(0, "######0.0####"), Format(0, "######0.0####"), Format(HOBS, "######0.0####"), Format(BETAM, "######0.0####"), Format(BETA2 * 1000#, "######0.0####"), Format(REFRAC * (1000# * RADCON), "######0.0####")
           End If
           
        If H2 < 0 Or (Dist > 10 * fGUESSL(0, HMAXT) And H2 < HMAXT) Then
'            If Not TempLoop Then Print #fileout%, FormatNumber(DIST + REARTH * (PHI2 - PHI1), 1, vbFalse, vbFalse, vbFalse), -1000
'            If Not TempLoop Then Write #fileout%, DIST + REARTH * (PHI2 - PHI1), -1000, ALFA(KMIN, jstep)
'            Print #fileout%, Format(DIST + Rearth * (PHI2 - PHI1), "######0.0####"), Format(PATHLENGTH, "######0.0####"), Format(-1000, "######0.0####"), Format(BETAM, "######0.0####"), Format(BETA2 * 1000#, "######0.0####"), Format(REFRAC * (1000# * RADCON), "######0.0####")
'            PATHLENGTH = Dist + Rearth * (PHI2 - PHI1) 'pathlength along the earth, Path > Pathlength
'            Pathh = Pathh + fGUESSP(BETA2 / RADCON, H2 - H1) 'Pathh is estimate of actual distance along the ray
            Print #fileout%, Format(Dist + Rearth * (PHI2 - PHI1), "######0.0####"), Format(Pathh, "######0.0####"), Format(-1000, "######0.0####"), Format(BETAM, "######0.0####"), Format(BETA2 * 1000#, "######0.0####"), Format(REFRAC * (1000# * RADCON), "######0.0####")
            jstop = jstep
            GoTo NEXTRAY
        Else
            If H2 > HLIMIT Then
                ' H2 PASSED HLIMIT METER
                TRUALT = BETAM - REFRAC
'                Print #fileout1%, Format(Str(BETAM), " ####0.0#####"), ",", Format(Str(REFRAC), " ####0.0#####"), ",", Format(Str(TRUALT), " ####0.0#####"), ",", Format(Str(AIRDRY), " ####0.0#####"), ",", Format(Str(AIRVAP), " ####0.0#####")
                If Not TempLoop Then Write #fileout1%, BETAM, REFRAC, TRUALT, AIRDRY, AIRVAP
                '  BETAM, REFRAC,TRUALT, AIRDRY AND AIRVAP HAVE BEEN STORED IN REF2017.OUT
                jstop = -1
                GoTo NEXTRAY
            Else
            End If
        End If

        '==============================
'        picVDW.Scale (0, HLIMIT)-(LRANGE, 0)
'        picVDW.Scale (-LRANGE * 0.01, HLIMIT + 2000)-(LRANGE * 1.01, -2000)
        Dist = Dist + Rearth * (PHI2 - PHI1) 'distance along the earth
'        Pathh = Pathh + fGUESSP(BETA2 / RADCON, H2 - H1) 'estimae of distance along the ray
        REFRAC = REFRAC + DREFR / RADCON
       
        If (nloop + 1) Mod Val(prjAtmRefMainfm.txtHeightStepSize.Text) = 0 Then
'           Print #fileout%, FormatNumber(DIST, 1, vbFalse, vbFalse, vbFalse), FormatNumber(h2, 1, vbFalse, vbFalse, vbFalse)
'           Write #fileout%, DIST, h2, ALFA(KMIN, jstep)
            TRUALT = BETAM - REFRAC
'           Print #fileout%, Format(DIST, "######0.0####"), Format(PATHLENGTH, "######0.0####"), Format(H2, "######0.0####"), Format(BETAM, "######0.0####"), Format(BETA2 * 1000#, "######0.0####"), Format(REFRAC * (1000# * RADCON), "######0.0####")
'           PATHLENGTH = Dist 'distance along earth
           Print #fileout%, Format(Dist, "######0.0####"), Format(Pathh, "######0.0####"), Format(H2, "######0.0####"), Format(BETAM, "######0.0####"), Format(BETA2 * 1000#, "######0.0####"), Format(REFRAC * (1000# * RADCON), "######0.0####")
           End If
           
        If H2 > HLIMIT Then
            GoTo NORAYPRINT
        Else
        End If
        If Dist > LRANGE Then
            GoTo NORAYPRINT
        Else
        End If
'        picVDW.PSet (DIST, h2), QBColor(10)
        If Not NoShow Then picVDW.Line -(Dist, H2), QBColor(NumBet)
'        picVDW.Refresh 'uncomment to follow the ray path as it builds up
'        DoEvents
NORAYPRINT:
           
        PHI1 = PHI2
        R1 = R2
        H1 = H2
        BETA1 = BETA2
        If H2 > HMAXT Then
            DPATH = fGUESSP(BETAM, H2) / NSTEPS
        Else
        End If

        'END DO-LOOP OVER PATH
        nloop = nloop + 1
    Loop
NEXTRAY:
'    a$ = INKEY$
     If Not NoShow Then
        DoEvents
        If Chr$(KeyPressed) = Chr$(32) Then GoTo ENDOFCALCULATION
        End If
        
    If jstop <> -1 And n_step > 0 Then
'        ALFA(KMIN, NumTc) = BETAM 'arc minutes
        ALFT(KMIN, jstep) = -1000
        jstop = jstep
        Exit For
    ElseIf jstop = -1 Then
'        ALFA(KMIN, NumTc) = BETAM 'arc minutes
        ALFT(KMIN, jstep) = BETAM - REFRAC 'arc minutes
        NumTc = NumTc + 1
        
        If chkTRef.Value = vbChecked And ALFA(KMIN, jstep) = 0 Then
           If TempLoop Then
              Write #filetemp%, TGROUND, ALFT(KMIN, jstep)
              End If
           End If
           
        End If
           
'   Call UpdateStatus(prjAtmRefMainfm, picProgBar,1, CLng(100# * CInt(Abs(BETAM - BETALO) / BETAST + 1) / CInt(Abs(BETAHI - BETALO) / BETAST + 1)))
    Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * jstep / (n_size + 1)))

'Next BETAM
Next jstep
ENDOFCALCULATION:
'========================END OF CALCULATION=====================

    If Not TempLoop Then
        Close #fileout1%
        Close #fileout4%
        Close #fileout%
    Else
       Close #fileout%
       NTLoop = NTLoop + 1
       End If

    DoEvents
    If chkPause.Value = vbChecked Then
       chkPause.Value = vbUnchecked
       chkPause.Refresh
       DoEvents
       Select Case MsgBox("Do you want to abort?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Pause calculation")
        Case vbYes
          GoTo exitcalc
        Case vbNo
       End Select
       End If
               
Next TLoop

If TempLoop And chkTRef.Value = vbChecked Then
   Dim Ta As Double, RefA As Double
   If chkTRef.Value = vbChecked Then Close #filetemp%
   
   'populate the MSchart box for TC with this scan values
   ReDim TransferCurve(1 To NTLoop, 1 To 2) As Variant
   
'   filetemp% = FreeFile
'   Open App.Path & "\TL_VDW_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat" For Input As #filetemp%
    filetemp% = FreeFile
    If OptionSelby.Value = False And chkDucting.Value = vbUnchecked Then
        FilNm = App.Path & "\TR_VDW_" & Trim(Str(Fix(TempStart))) & "-" & Trim(Str(Fix(TempEnd))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
        Open FilNm For Input As #filetemp%
    ElseIf chkDucting.Value = vbChecked Then
        FilNm = App.Path & "\TR_VDW_INV_" & Trim(Str(Fix(TempStart))) & "-" & Trim(Str(Fix(TempEnd))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
        Open FilNm For Input As #filetemp%
        Line Input #filetemp%, doclin$
        'use van der Werf's formula to simulate an inversion
    ElseIf OptionSelby.Value = True Then
        FilNm = App.Path & "\TR_VDW_LAYERS_" & Trim(Str(Fix(TempStart))) & "-" & Trim(Str(Fix(TempEnd))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
        Open FilNm For Input As #filetemp%
        Line Input #filetemp%, doclin$
        End If
   
   j = 0
   Do Until EOF(filetemp%)
      
      Input #filetemp%, Ta, RefA
      j = j + 1

      TransferCurve(j, 1) = " " & CStr(Ta)
      TransferCurve(j, 2) = RefA
      
   Loop
   Close #filetemp%

    With MSCharttc
      .chartType = VtChChartType2dLine
      .RandomFill = False
      .ChartData = TransferCurve
    End With
    
    TabRef.Tab = 5
    
'   Exit Sub
   End If
   
NHloop = NHloop + 1

Next HLoop

If TempLoop Then
   Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset
   Call UpdateStatus(prjAtmRefMainfm, picProgBar2, 0, 0) 'reset
   Call UpdateStatus(prjAtmRefMainfm, picProgBar3, 0, 0) 'reset
   prjAtmRefMainfm.progressfrm.Visible = False
   TempLoop = False
   CalcComplete = True
   cmpTLoop.Enabled = True
   progressfrm2.Visible = False
   Exit Sub
   End If

Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
'prjAtmRefMainfm.progressfrm.Visible = False
     
DisplayResults:

If chkAtmRefDllvalue = vbChecked Then 'chkUseDll.Value = vbChecked Then
   'open the refraction info file and load up the ALFA and ALFT arrays, as well as determining the angle to the horizon
   'determine the filename
    
    filein% = FreeFile
    If FilePath = "" Then FilePath = App.Path
    Open FilePath & "\TR_VDW_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat" For Input As #fileout%
    jstep = 0
    NumTc = 0
    Do Until EOF(filein%)
   
       Line Input #filein%, doclin$
       doclin$ = Trim$(doclin$)
       '------------parse the line------------------
       nu$ = ""
       numNu% = 6
       StartedNumber = False
       For j = Len(doclin$) To 1 Step -1
          num1$ = Trim$(Mid$(doclin$, j, 1))
          
          If Trim$(num1$) = "" Then
             If StartedNumber Then
                StartedNumber = False
                AA(numNu%) = Val(nu$)
                numNu% = numNu% - 1
                If numNu% = 0 Then Exit For
                nu$ = ""
                End If
          Else
             StartedNumber = True
             nu$ = num1$ + nu$
             End If
             
       Next j
       '-----------end parsing----------------
       If numNu% = 1 And StartedNumber And j = 0 Then
          AA(1) = Val(nu$)
          End If
          
       If jstep = 0 Then
          a0 = AA(4)
          e0 = AA(6)
          jstep = 1
       Else
          If AA(4) <> a0 Then
             ALFA(KMIN, jstep) = a0
             ALFT(KMIN, jstep) = ALFA(KMIN, jstep) - e0 / (1000 * RADCON)

             NumTc = NumTc + 1
             jstep = jstep + 1
             End If
             
          a0 = AA(4)
          e0 = AA(6)
             
          End If
          
       If AA(3) = -1000 Then
          jstop = jstep
          Exit Do
       Else
          jstop = -1
          End If
          
    Loop
   Close #filein%
   End If
   
If n_size = 0 And jstep <> 0 Then jstop = jstep
If jstep > 1 Then jstop = jstep - 1

If jstop <> -1 Then
    StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KMIN, jstop - 1)) & vbCrLf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(Rearth / (Rearth + HOBS)) / RADCON))
    End If
Call StatusMessage(StatusMes, 1, 0)
prjAtmRefMainfm.lblHorizon.Caption = StatusMes
prjAtmRefMainfm.lblHorizon.Refresh
DoEvents
'     ier = MsgBox(StatusMes, vbInformation + vbOKOnly, "Horizon")
StatusMes = "Ray tracing calculation complete..."
Call StatusMessage(StatusMes, 1, 0)

CalcComplete = True
'prjAtmRefMainfm.frmFindPressure.Visible = True
Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset
prjAtmRefMainfm.progressfrm.Visible = False

Close

'      OPEN(UNIT=20,FILE='tc.dat',STATUS='UNKNOWN')

    If Dir(App.Path & "\tc-VDW.dat") <> sEmpty Then
       Kill App.Path & "\tc-VDW.dat"
       End If
   
    'zero raytracing display arrays
'    For II = 0 To NumSuns - 1
'       For jj = 0 To TotNumSunAlt - 1
'          SunAngles(II, jj) = 0
'       Next jj
'       NumSunAlt(II) = 0
'    Next II
   
  StatusMes = "Writing transfer curve."
  Call StatusMessage(StatusMes, 1, 0)
  filnum% = FreeFile
  Open App.Path & "\tc_VDW.dat" For Output As #filnum%
  '      WRITE (20,*) N
  NumTc = 0
  Print #filnum%, n_size
  found% = 0
  For j = 1 To jstop
  '        WRITE(20,1) ALFA(KMIN,J),ALFT(KMIN,J)
      Print #filnum%, ALFA(KMIN, j), ALFT(KMIN, j)
      If ALFA(KMIN, j) = 0 And found% = 0 Then 'display the refraction value for the zero view angle ray
         prjAtmRefMainfm.lblRef.Caption = "Atms. refraction (deg.) = " & Abs(ALFT(KMIN, j)) / 60# & vbCrLf & "Atms. refraction (mrad) = " & Abs(ALFT(KMIN, j)) * 1000# * cd / 60#
         prjAtmRefMainfm.lblRef.Refresh
         found% = 1
         
         If LoopingAtmTracing Then 'loopin in radiosondes atmosphere files, so record the 90 degrees zenith angle refracing
            fileoutatm = FreeFile
            Open FileNameAtmOut For Append As #fileoutatm
         
            Print #fileoutatm, DateNameAtm, Abs(ALFT(KMIN, j)) * 1000# * cd / 60#
            Close #fileoutatm
            End If

         DoEvents
         End If
  'store all view angles that contribute to sun's orb
      NumTc = NumTc + 1
      For KA = 1 To NumSuns
         y = ALFT(KMIN, j) - ALT(KA)
         If Abs(y) <= ROBJ Then
            'only accept rays that pass over the horizon (ALFT(KMIN, J) <> -1000) and are within the solar disk
            SunAngles(KA - 1, NumSunAlt(KA - 1)) = j
            NumSunAlt(KA - 1) = NumSunAlt(KA - 1) + 1
            End If
      Next KA
  Next j
  Close #filnum%
  'makes copy appended with sondes date if it is sondes atmosphere
  With prjAtmRefMainfm
     If OptionSelby.Value = True Then
        If .opt10.Value = True Then
           If InStr(.txtOther.Text, "-sondes.txt") And LoopingAtmTracing Then
           
              If .chkHgtProfile.Value = vbChecked Then
              
                 'extract date from sondes file
                 SondesDate$ = Mid$(Trim$(.txtOther.Text), 1, Len(Trim$(.txtOther.Text)) - 4)
                 If BARParametersfm.optAllSeasons.Value = True Then
                    FileCopy App.Path & "\tc_VDW.dat", SondesDate$ & "-tc-2-VDW.dat"
                 ElseIf BARParametersfm.optAllOrigPress.Value = True Then
                    FileCopy App.Path & "\tc_VDW.dat", SondesDate$ & "-tc-3-VDW.dat"
                    End If
                    
              Else 'mark it as not be hill hugging
              
                 'extract date from sondes file
                 SondesDate$ = Mid$(Trim$(.txtOther.Text), 1, Len(Trim$(.txtOther.Text)) - 4)
                 
                 If ZeroRefTesting Then
                    If BARParametersfm.optAllSeasons.Value = True Then
                       FileCopy App.Path & "\tc_VDW.dat", SondesDate$ & "-no-tc-VDW.dat"
                    ElseIf BARParametersfm.optAllOrigPress.Value = True Then
                       FileCopy App.Path & "\tc_VDW.dat", SondesDate$ & "-no-tc-3-VDW.dat"
                       End If
                       
                 Else
                    If BARParametersfm.optAllSeasons.Value = True Then
                       FileCopy App.Path & "\tc_VDW.dat", SondesDate$ & "-tc-VDW.dat"
                    ElseIf BARParametersfm.optAllOrigPress.Value = True Then
                       FileCopy App.Path & "\tc_VDW.dat", SondesDate$ & "-tc-3-VDW.dat"
                       End If
                       
                    End If
                    
                 End If
                 
              End If
           End If
        End If
  End With

      'now load up transfercurve array for plotting
      ReDim TransferCurve(1 To NumTc, 1 To 2) As Variant

      For j = 1 To NumTc
         TransferCurve(j, 1) = " " & CStr(ALFA(KMIN, j - 1))
         TransferCurve(j, 2) = ALFT(KMIN, j - 1)
'         TransferCurve(J, 1) = " " & CStr(ALFT(KMIN, J))
'         TransferCurve(J, 2) = ALFA(KMIN, J)
      Next j

      With MSCharttc
        .chartType = VtChChartType2dLine
        .RandomFill = False
'        .RowCount = 2
'        .ColumnCount = IncN
'        .RowLabel = "True angle (min)"
'        .ColumnLabel = "View angle (min)"
        .ChartData = TransferCurve
      End With

'      For KA = 1 To NumSuns
'         cc = NumSunAlt(KA - 1)
'      Next KA
''C
''C   FIND THE ALTITUDE OF THE HORIZON
''C

'      For j = NumTc To 1 Step -1
'         If (ALFT(KMIN, j - 1) = -1000#) Then ISTOP = j - 2
'      Next j
'      PRINT *, '    Apparent Altitude of the Horizon (arcminutes)',
'     *      ALFA(KSTOP,ISTOP)
'      PRINT *, '    True Altitude of the Horizon (arcminutes)',
'     *      (-DACOS(RE/(RE+HOBS))/CONV)
'     StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KMIN, jstop - 1))& vbcrlf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(REARTH / (REARTH + HOBS)) / RADCON))
'     StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KWAV, jstop - 1))& vbcrlf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(s / (s + h0)) / CONV))
'     Call StatusMessage(StatusMes, 1, 0)
     
'     ier = MsgBox(StatusMes, vbInformation + vbOKOnly, "Horizon")
'     lblHorizon.Caption = StatusMes
'     DoEvents
'
'     StatusMes = "Ray tracing calculation complete"
'     Call StatusMessage(StatusMes, 1, 0)
     
 StatusMes = "Drawing the rays on the sky simulation, please wait...."
 Call StatusMessage(StatusMes, 1, 0)
 'load angle combo boxes
'    AtmRefPicSunfm.WindowState = vbMinimized
'    BrutonAtmReffm.WindowState = vbMaximized
 'set size of picref by size of earth
 Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
 prjAtmRefMainfm.cmbSun.Clear
 prjAtmRefMainfm.cmbAlt.Clear
 For i = 1 To NumSuns
    If NumSunAlt(i - 1) > 0 Then prjAtmRefMainfm.cmbSun.AddItem i
 Next i
 
 prjAtmRefMainfm.TabRef.Tab = 4
 DoEvents

cmbSun.ListIndex = 0

'========HOLD THE SCREEN TILL HITTING ANY KEY==================

If Not NoShow Then

    picVDW.ForeColor = QBColor(12)
    picVDW.CurrentX = LRANGE * 0.5 - 3500: picVDW.CurrentY = 2000: picVDW.Print "Ray tracing complete, press any key to proceed"
    picVDW.ForeColor = QBColor(15)
    KeyPressed = 0
    Do While KeyPressed = 0
       DoEvents
    Loop
    'in$ = INPUT$(1)
    picVDW.Cls

    'Plot refraction (vertical), versus apparent angle
    '
    Dim BMIN As Double, BMAX As Double, NTICKS As Long
    If BETAHI < 312.5 Then
        BMIN = -6.25
        BMAX = 300
        NTICKS = 10
    Else
        If BETAHI < 625 Then
            BMIN = -12.5
            BMAX = 600
            NTICKS = 30
        Else
            If BETAHI < 1350 Then
                BMIN = -25
                BMAX = 1200
                NTICKS = 60
            Else
                If BETAHI < 2700 Then
                    BMIN = -50
                    BMAX = 2700
                    NTICKS = 150
                Else
                    If BETAHI <= 5400 Then
                        BMIN = -100
                        BMAX = 5400
                        NTICKS = 300
                    Else
                    End If
                End If
            End If
        End If
    End If
    picVDW.Scale (BMIN, 60)-(BMAX, -10)
    picVDW.CurrentX = 2: picVDW.CurrentY = 5: picVDW.Print "Vertical: refraction, tickmarks 10 arcmin"
    If BETAHI <= 300 Then
        picVDW.CurrentX = 3: picVDW.CurrentY = 5: picVDW.Print "Horizontal: apparent observer angle, 0 - 5 deg"
        picVDW.CurrentX = 4: picVDW.CurrentY = 5: picVDW.Print "Horizontal tickmarks 10 arcmin"
    Else
        If BETAHI <= 600 Then
            picVDW.CurrentX = 3: picVDW.CurrentY = 5: picVDW.Print "Horizontal: apparent observer angle, 0 - 10 deg"
            picVDW.CurrentX = 4: picVDW.CurrentY = 5: picVDW.Print "Horizontal tickmarks 30 arcmin"
        Else
            If BETAHI <= 1200 Then
                picVDW.CurrentX = 3: picVDW.CurrentY = 5: picVDW.Print "Horizontal: apparent observer angle, 0 - 20 deg"
                picVDW.CurrentX = 4: picVDW.CurrentY = 5: picVDW.Print "Horizontal tickmarks 1 deg"
            Else
                If BETAHI <= 2700 Then
                    picVDW.CurrentX = 3: picVDW.CurrentY = 5: picVDW.Print "Horizontal: apparent observer angle, 0 - 45 deg"
                    picVDW.CurrentX = 4: picVDW.CurrentY = 5: picVDW.Print "Horizontal tickmarks 2.5 deg "
                Else
                    If BETAHI <= 5400 Then
                        picVDW.CurrentX = 3: picVDW.CurrentY = 5: picVDW.Print "Horizontal: apparent observer angle, 0 - 90 deg"
                        picVDW.CurrentX = 4: picVDW.CurrentY = 5: picVDW.Print "Horizontal tickmarks 5 deg"
                    Else
                    End If
                End If
            End If
        End If
    End If
          
    picVDW.Line (BMIN, -10)-(BMAX, 60), QBColor(1), B
    picVDW.Line (BMIN, 0)-(BMAX, 0), QBColor(1)
    For XX = NTICKS To BMAX Step NTICKS
        picVDW.Line (XX, -1)-(XX, 1), QBColor(1)
    Next XX
    picVDW.Line (0, -10)-(0, 60), QBColor(1)
    For YY = -10 To 60 Step 10
        picVDW.Line (-NTICKS / 5, YY)-(NTICKS / 5, YY), QBColor(1)
    Next YY
    '
    filein2% = FreeFile
    Open App.Path & "\REF2017.OUT" For Input As #filein2%  'BETAM, REFRAC, TRUALT, AIRDRY, AIRVAP
    Line Input #filein2%, doclin$
    Do While Not EOF(filein2%)
        Input #filein2%, BETAM, REFRAC, TRUALT, AIRDRY, AIRVAP
        picVDW.Circle (BETAM, REFRAC), 0.1 * NTICKS, QBColor(7)
    Loop
    Close #filein2%
    
    '========HOLD THE SCREEN TILL HITTING ANY KEY==================
    picVDW.ForeColor = QBColor(14)
    picVDW.CurrentX = 28: picVDW.CurrentY = 40: picVDW.Print "press any key to proceed"
    picVDW.ForeColor = QBColor(15)
    KeyPressed = 0
    Do While KeyPressed = 0
       DoEvents
    Loop
    'in$ = INPUT$(1)
    'SCREEN 0: WIDTH 80
    
    picVDW.Cls
    picVDW.FontSize = 14
    picVDW.ForeColor = QBColor(15)
    picVDW.Scale (0, 150)-(100, 50)
    picVDW.CurrentX = 3: picVDW.CurrentY = 110: picVDW.Print "Output written to REF2017.OUT as: BETA, REFRAC, TRUALT, AIRDRY, AIRVAP"
    picVDW.CurrentX = 3: picVDW.CurrentY = 105: picVDW.Print "where BETA = apparent altitude (arcmin)"
    picVDW.CurrentX = 3: picVDW.CurrentY = 100: picVDW.Print "      REFRAC = path integral of refraction (arcmin)"
    picVDW.CurrentX = 3: picVDW.CurrentY = 95: picVDW.Print "      TRUALT = true altitude (arcmin)"
    picVDW.CurrentX = 3: picVDW.CurrentY = 90: picVDW.Print "      AIRDRY = path integral of dry-air density (kg/m2/100) "
    picVDW.CurrentX = 3: picVDW.CurrentY = 85: picVDW.Print "      AIRVAP = path integral of water vapor density (kg/m2/100)"
    picVDW.ForeColor = QBColor(14)
    picVDW.CurrentX = 15: picVDW.CurrentY = 65: picVDW.Print "Change parameters? (Y/any other key), else end of program "
    picVDW.ForeColor = QBColor(15)
    KeyPressed = 0
    Do While KeyPressed = 0
       DoEvents
    Loop
    If Chr(KeyPressed) = "Y" Or Chr(KeyPressed) = "y" Then
        GoTo PARAMETERLIST
    Else
        TabRef.Tab = 0
    End If
    End If

cmdVDW.Enabled = True
cmdCalc.Enabled = True
cmdMenat.Enabled = True
cmdRefWilson.Enabled = True
Screen.MousePointer = vbDefault

FinishedTracing = True
cmdVDW_error = 0
   On Error GoTo 0
   Exit Sub

cmdVDW_Click_Error:
'    Resumej 'diagnostics
    cmdVDW_error = -1
    Close
    
    Screen.MousePointer = vbDefault
    cmdVDW.Enabled = True
    cmdCalc.Enabled = True
    cmdMenat.Enabled = True
    cmdRefWilson.Enabled = True
    
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdVDW_Click of Form prjAtmRefMainfm"

End Sub

'===================FUNCTIONS=================================
Function fACS(x As Double) As Double
'Arccosine
fACS = 2 * Atn(1) - Atn(x / Sqr(1# - x * x))
End Function

Function fASN(x As Double) As Double
'Arcsine
fASN = Atn(x / Sqr(1# - x * x))
End Function


Function fDLNTDH(H As Double, Dist As Double, NumLayers As Long) As Double
'The deriative of ln(T): (1/T)(dT/dh)
fDLNTDH = fDTDH(H, Dist, NumLayers) / fTEMP(H, Dist, NumLayers)
End Function

Function fDTDH(H As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
Dim dh As Double, T1 As Double, T2 As Double, T3 As Double, T4 As Double
dh = 0.01
T1 = fTEMP(H - 3 * dh / 2, Dist, NumLayers)
T2 = fTEMP(H - dh / 2, Dist, NumLayers)
T3 = fTEMP(H + dh / 2, Dist, NumLayers)
T4 = fTEMP(H + 3 * dh / 2, Dist, NumLayers)
fDTDH = ((T3 - T2) * (27 / 24) - (T4 - T1) / 24) / dh
End Function

Function fFNDPD1(H As Double, PRESSD1() As Double, Dist As Double, NumLayers As Long) As Double
'Interpolation in the array PRESSD1
'DefDbl A-H, O-Z
'SHARED HLIMIT, PDM1, RELH, BD, BW, HMAXP1
Dim y As Double, i As Long, P1 As Double, P2 As Double, DP2DY As Double, FK1 As Double, FK2 As Double, FK3 As Double, FK4 As Double
Dim YSTEP As Double
  
If (H < (HMAXP1 + 1) And H >= 0) Then
    y = H
    YSTEP = y - Int(y)
    'STEP 1
    i = Int(y)
    P1 = PRESSD1(i + 1)
    P2 = RELH * fVAPOR(y, Dist, NumLayers)
    DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
    FK1 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
    'END STEP 1
    'STEP 2
    y = i + YSTEP / 2
    P1 = PRESSD1(i + 1) + FK1 * YSTEP / 2
    P2 = RELH * fVAPOR(y, Dist, NumLayers)
    DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
    FK2 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
    'END STEP 2
    'STEP 3
    y = i + YSTEP / 2
    P1 = PRESSD1(i + 1) + FK2 * YSTEP / 2
    P2 = RELH * fVAPOR(y, Dist, NumLayers)
    DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
    FK3 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
    'END STEP 3
    'STEP 4
    y = i + YSTEP
    P1 = PRESSD1(i + 1) + FK3 * YSTEP
    P2 = RELH * fVAPOR(y, Dist, NumLayers)
    DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
    FK4 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
    'END STEP 4
    fFNDPD1 = PRESSD1(i + 1) + (YSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
Else
End If
End Function

Function fFNDPD2(H As Double, PRESSD2() As Double, Dist As Double, NumLayers As Long) As Double
'Interpolation in the array PRESSD2
'DefDbl A-H, O-Z
'SHARED HLIMIT, PDM10
Dim y As Double, i As Long, P1 As Double, P2 As Double, DP2DY As Double, FK1 As Double, FK2 As Double, FK3 As Double, FK4 As Double
Dim YSTEP As Double
y = H
YSTEP = (y - Int(y)) * 10
'STEP 1
i = Int(y / 10)
P1 = PRESSD2(i + 1)
P2 = RELH * fVAPOR(y, Dist, NumLayers)
DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
FK1 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
'END STEP 1
'STEP 2
y = i + YSTEP / 2
P1 = PRESSD2(i + 1) + FK1 * YSTEP / 2
P2 = RELH * fVAPOR(y, Dist, NumLayers)
DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
FK2 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
'END STEP 2
'STEP 3
y = i + YSTEP / 2
P1 = PRESSD2(i + 1) + FK2 * YSTEP / 2
P2 = RELH * fVAPOR(y, Dist, NumLayers)
DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
FK3 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
'END STEP 3
'STEP 4
y = i + YSTEP
P1 = PRESSD2(i + 1) + FK3 * YSTEP
P2 = RELH * fVAPOR(y, Dist, NumLayers)
DP2DY = RELH * fDVAPDT(y, Dist, NumLayers) * fDTDH(y, Dist, NumLayers)
FK4 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers) - BW * P2 * fGRAVRAT(y) / fTEMP(y, Dist, NumLayers)
'END STEP 4
fFNDPD2 = PRESSD2(i + 1) + (YSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
PD2READY:
End Function

Function fGUESSL(BETAM As Double, height As Double) As Double
'Estimate for the distance along the Earth's surface
'for a ray with tilt angle BETA to cover the height interval
'from H=0 to H=HEIGHT.
'SHARED REARTH, RADCON
Dim A As Double, b As Double, C As Double
b = BETAM * RADCON
C = fASN((Rearth * Cos(b)) / (Rearth + height))
A = (2 * Atn(1) - b - C)
fGUESSL = A * Rearth
End Function

Function fGUESSP(BETA As Double, height As Double) As Double
'Estimate for the pathlength from H=0 to H=HEIGHT
'for a ray with tilt angle BETA
'SHARED REARTH, RADCON
Dim A As Double, b As Double, C As Double
A = 1
b = 2 * Rearth * Sin(BETA * RADCON)
C = -2 * Rearth * height - height ^ 2
fGUESSP = (-b + Sqr(Abs(b ^ 2 - 4 * A * C))) / 2
End Function

Function fRCINV(H As Double, PRESSD1() As Double, PRESSD2() As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'Inverse curvature of a light ray: 1/r for a horizontal ray
'SHARED HCROSS, HLIMIT, REARTH
fRCINV = fDNDH(H, PRESSD1, PRESSD2, Dist, NumLayers) / fREFIND(H, PRESSD1, PRESSD2, Dist, NumLayers) '<<<<<<<<<<<<<<<<<<<
End Function

Function fREFIND(H As Double, PRESSD1() As Double, PRESSD2() As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'SHARED HLIMIT, AD, AW, RELH, HMAXP1
Dim PD As Double, PW As Double
If H < HMAXP1 Then
    PD = fFNDPD1(H, PRESSD1, Dist, NumLayers)
    PW = RELH * fVAPOR(H, Dist, NumLayers)
    fREFIND = 1 + (AD * PD + AW * PW) / fTEMP(H, Dist, NumLayers)
Else
    PD = fFNDPD2(H, PRESSD2, Dist, NumLayers)
    PW = RELH * fVAPOR(H, Dist, NumLayers)
    fREFIND = 1 + (AD * PD + AW * PW) / fTEMP(H, Dist, NumLayers)
End If
End Function

Function fSTNDATM(H As Double, Dist As Double, NumLayers As Long) As Double
'STANDARD MUSA76 ATMOSPHERE WITH TROPOSPHERE AT HCROSS = (TGROUND-216.65)/0.0065
'SHARED HCROSS, TGROUND
    Dim GrndHgt As Double

    If OptionSelby.Value = False And chkDucting.Value = vbUnchecked Then
        If H < HCROSS Then
            fSTNDATM = 216.65 + 0.0065 * (HCROSS - H)
        Else
            If H < 20000# Then
                fSTNDATM = 216.65
            Else
                If H < 32000# Then
                    fSTNDATM = 216.65 + 0.001 * (H - 20000#)
                Else
                    If H < 47000# Then
                        fSTNDATM = 228.65 + 0.0028 * (H - 32000#)
                    Else
                        If H < 51000# Then
                            fSTNDATM = 270.65
                        Else
                            If H < 71000# Then
                                fSTNDATM = 270.65 - 0.0028 * (H - 51000#)
                            Else
                                If H < 85000 Then
                                    fSTNDATM = 214.65 - 0.002 * (H - 71000#)
                                Else
                                    fSTNDATM = 186.65
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    ElseIf OptionSelby.Value = True Then
      Dim i As Long, HN As Double, found%
   
     found% = 0
     '<<<<<<<<<<<<<important change -- detect meterse or km as scale---
     HN = H * 0.001 'convert to kilometers
     'import model 2D model of ground height vs. distance Dist (m) from observer
     If chkHgtProfile.Value = vbChecked And Dist <> -1 Then
        GrndHgt = DistModel(Dist)
        'make the atmospheric layers approx. hug the ground, i.e., they dip with valleys, rise with hills
        'but they don't dip completely, and if observerheight - ground height > 400 meters, use standard atmosphere
        If Dist >= 80000 Then
           'use standard vdw atmosphere after 10 km from place
            If H < HCROSS Then
                fSTNDATM = 216.65 + 0.0065 * (HCROSS - H)
            Else
                If H < 20000# Then
                    fSTNDATM = 216.65
                Else
                    If H < 32000# Then
                        fSTNDATM = 216.65 + 0.001 * (H - 20000#)
                    Else
                        If H < 47000# Then
                            fSTNDATM = 228.65 + 0.0028 * (H - 32000#)
                        Else
                            If H < 51000# Then
                                fSTNDATM = 270.65
                            Else
                                If H < 71000# Then
                                    fSTNDATM = 270.65 - 0.0028 * (H - 51000#)
                                Else
                                    If H < 85000 Then
                                        fSTNDATM = 214.65 - 0.002 * (H - 71000#)
                                    Else
                                        fSTNDATM = 186.65
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            Exit Function
         Else
'           'use relaxed grnhgt to determine the temperature in the layered atmosphere
'           If GrndHgt < HOBS - 100 Then
'              'keep the inversion going but at 200 meters lower
'              GrndHgt = HOBS - 100
'              End If
           GrndHgt = GrndHgt * 0.001
           End If
           
     Else
        GrndHgt = 0#
        End If
        
      For i = 1 To NumLayers - 1
         If HN < ELV(i) + GrndHgt Then
            fSTNDATM = ((TMP(i) - TMP(i - 1)) / (ELV(i) - ELV(i - 1))) * (HN - ELV(i - 1) - GrndHgt) + TMP(i - 1)
            found% = 1
            Exit For
            End If
      Next i
      If found% = 0 Then
         If HN >= ELV(NumLayers - 1) + GrndHgt Then
            fSTNDATM = TMP(NumLayers - 1) 'make isothermal for H >= last layer boundary
            End If
         End If
      
'         If HN < HL(i) + GrndHgt Then
'            fSTNDATM = TL(i - 1) - LRL(i - 1) * (HN - HL(i - 1) - GrndHgt)
'            found% = 1
'            Exit For
'            End If
'      Next i
'      If found% = 0 Then
'         If HN >= HL(NumLayers - 1) + GrndHgt Then
'            fSTNDATM = TL(NumLayers - 1) 'make isothermal for H >= last layer boundary
'            End If
'         End If
         
   ElseIf chkDucting.Value = vbChecked Then
      'add inversion
      
      If H < Val(txtEInv.Text) Then
      
            'use van der Werf's formula to simulate an inversion
            'fSTNDATM = TGROUND - H * 0.0065 + (LRL0 + 0.0065) * (H * (TL0 - H) / (TL0 * (1 + H / HL0)))
            fSTNDATM = TGROUND - H * 0.0065 + CInv * H * (1# - H * AInv) / (1 + H * BInv)
      
      Else 'outside the inversion layer, so use the standard vdW atmosphere
      
        If H < HCROSS Then
            fSTNDATM = 216.65 + 0.0065 * (HCROSS - H)
        Else
            If H < 20000# Then
                fSTNDATM = 216.65
            Else
                If H < 32000# Then
                    fSTNDATM = 216.65 + 0.001 * (H - 20000#)
                Else
                    If H < 47000# Then
                        fSTNDATM = 228.65 + 0.0028 * (H - 32000#)
                    Else
                        If H < 51000# Then
                            fSTNDATM = 270.65
                        Else
                            If H < 71000# Then
                                fSTNDATM = 270.65 - 0.0028 * (H - 51000#)
                            Else
                                If H < 85000 Then
                                    fSTNDATM = 214.65 - 0.002 * (H - 71000#)
                                Else
                                    fSTNDATM = 186.65
                                End If
                            End If
                        End If
                    End If
                End If
            End If
          End If
        End If
        
     End If

End Function

Function fTEMP(H As Double, Dist As Double, NumLayers As Long) As Double
'SHARED HCROSS, TGROUND
fTEMP = fSTNDATM(H, Dist, NumLayers)
End Function

Function fVAPOR(H As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'SHARED OPTVAP
Dim T As Double
T = fTEMP(H, Dist, NumLayers)
On OPTVAP GoTo opt1, opt2, opt3, opt4
opt1:
fVAPOR = (T / 247.1) ^ 18.36 'PL2 VORM
GoTo VAPORREADY:
opt2:
fVAPOR = Exp(21.39 - 5349# / T)  'CC2 FORM
GoTo VAPORREADY:
opt3:
fVAPOR = Exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T) 'CC4 FORM
GoTo VAPORREADY:
opt4:
fVAPOR = (T / 273.15) ^ (2.5) * Exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T) 'SACKUR-TETRODE FORM, FIT.

GoTo VAPORREADY:
VAPORREADY:
End Function
Function fGRAVRAT(H As Double) As Double
'DefDbl A-H, O-Z
'gravitation at H/ gravitation at H=0
'SHARED GRAVC, OBSLAT, DEG2RAD, REARTH
fGRAVRAT = (Rearth / (Rearth + H)) ^ 2
End Function

Function fPRESSURE(H As Double, PRESSD1() As Double, PRESSD2() As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'SHARED HLIMIT, RELH, HMAXP1

    If OptionSelby Then 'used measured pressure instead
       found% = 0

       If H * 0.001 < ELV(0) Then
          fPRESSURE = PRSR(0)
          found% = 1
       Else
            For j = 0 To NumLayers - 2
               If H * 0.001 >= ELV(j) And H * 0.001 < ELV(j + 1) Then
                  If ELV(j + 1) <> ELV(j) Then
                     fPRESSURE = ((PRSR(j + 1) - PRSR(j)) / (ELV(j + 1) - ELV(j))) * (H * 0.001 - ELV(j)) + PRSR(j)
                     found% = 1
                     Exit For
                  Else
                     fPRESSURE = PRSR(j)
                     found% = 1
                     Exit For
                     End If
                  End If
            Next j
          End If
          If found% = 1 Then Exit Function
       End If

   
If H < HMAXP1 Then
    fPRESSURE = fFNDPD1(H, PRESSD1, Dist, NumLayers) + RELH * fVAPOR(H, Dist, NumLayers)
Else
    fPRESSURE = fFNDPD2(H, PRESSD2, Dist, NumLayers) + RELH * fVAPOR(H, Dist, NumLayers)
End If
End Function

Function fDVAPDT(H As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'SHARED OPTVAP
Dim T As Double, HV As Double
T = fTEMP(H, Dist, NumLayers)
On OPTVAP GoTo opt1, opt2, opt3, opt4
opt1:
HV = (T / 247.1) ^ 18.36 'PL2 FORM
fDVAPDT = 18.36 / 247.1 * (T / 247.1) ^ 17.36
GoTo DVAPDTREADY:
opt2:
HV = Exp(21.39 - 5349# / T)  'CC2 FORM
fDVAPDT = HV * 5349 / T ^ 2
GoTo DVAPDTREADY:
opt3:
HV = Exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T) 'CC4 FORM
fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / T ^ 2)
GoTo DVAPDTREADY:
opt4:
HV = (T / 273.15) ^ (2.5) * Exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T) 'SACKUR-TETRODE FORM, FIT.
fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / T ^ 2)
DVAPDTREADY:
End Function

Function fDENSDRY(H As Double, PRESSD1() As Double, PRESSD2() As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'SHARED HLIMIT, AMASSD, RBOLTZ, RELH
Dim PVAP As Double, PDRY As Double
PVAP = RELH * fVAPOR(H, Dist, NumLayers)
PDRY = fPRESSURE(H, PRESSD1, PRESSD2, Dist, NumLayers) - PVAP
fDENSDRY = (AMASSD / RBOLTZ) * PDRY / fTEMP(H, Dist, NumLayers)
End Function

Function fDENSVAP(H As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'SHARED RELH, AMASSW, RBOLTZ
Dim PVAP As Double
PVAP = RELH * fVAPOR(H, Dist, NumLayers)
fDENSVAP = (AMASSW / RBOLTZ) * PVAP / fTEMP(H, Dist, NumLayers)
End Function


Function fDNDH(H As Double, PRESSD1() As Double, PRESSD2() As Double, Dist As Double, NumLayers As Long) As Double
'DefDbl A-H, O-Z
'SHARED AD, AW, BD, BW, RELH
Dim T As Double, PW As Double, PD As Double, DPWDH As Double, DPDDH As Double, HV1 As Double, HV2 As Double
T = fTEMP(H, Dist, NumLayers)
PW = RELH * fVAPOR(H, Dist, NumLayers)
PD = fPRESSURE(H, PRESSD1, PRESSD2, Dist, NumLayers) - PW
DPWDH = RELH * fDVAPDT(H, Dist, NumLayers) * fDTDH(H, Dist, NumLayers)
DPDDH = -DPWDH - BD * PD * fGRAVRAT(H) / T - BW * PW * fGRAVRAT(H) / T
HV1 = (AD * DPDDH + AW * DPWDH) / T
HV2 = -(AD * PD + AW * PW) / T / T * fDTDH(H, Dist, NumLayers)
fDNDH = HV1 + HV2
End Function

'Private Sub cmdVDW_Click()
'
''DefDbl A-H, O-Z
''
'Dim PRESSD1(99999), PRESSD2(99999)
''
''DECLARE FUNCTION f.VAPOR (H)
''DECLARE FUNCTION f.STNDATM (H) 'US1976 Standard atmosphere
''DECLARE FUNCTION f.FNDPD1 (H) 'Find INT[1/T] from lookup table PRESSD1
''DECLARE FUNCTION f.FNDPD2 (H) 'Find INT[/T] from lookup table PRESSD2
''DECLARE FUNCTION f.DLNTDH (H) 'd(ln(T))/dh = (1/T) dT/dH
''DECLARE FUNCTION f.RCINV (H) 'Curvature (1/r) of a horizontal light ray
''DECLARE FUNCTION f.TEMP (H) 'Temperature as a function of elevation
''DECLARE FUNCTION f.REFIND (H) 'Index of refraction
''DECLARE FUNCTION f.DTDH (H) 'dT/dH
''DECLARE FUNCTION f.GUESSL (BETAM, height) 'Guess horizontal distance needed to reach HEIGHT.
''DECLARE FUNCTION f.GUESSP (BETA, height) 'Guess path length needed to reach HEIGHT.
''DECLARE FUNCTION f.ASN (X) 'Arc-sine
''DECLARE FUNCTION f.ACS (X) 'Arc-cosine
''DECLARE FUNCTION f.DENSDRY (H) ' Density of dry air as function of elevation
''DECLARE FUNCTION f.DENSVAP (H) ' Density of water vapor as function of elevation
''DECLARE FUBCTION f.GRAVRAT(H) 'GRAVITY/GRAVITY(0) AS FUNCTION OF HEIGHT
''DECLARE FUNCTION f.PRESSURE(H) 'Find total pressure from f.FNDPD2 and vapor pressure
''DECLARE FUNCTION f.DVAPDT(H) 'Derived from vapor pressure to T as a function of H
''DECLARE FUNCTION f.DNDH (H) ' Derivative of refraction index
''Screen 12
'
'picVDW.Cls
''Cls
'
'picVDW.Scale (-20, 0)-(100, 120)
'picVDW.Line (2, 80)-(10, 81), 14, BF
'picVDW.Line (2, 80)-(3, 60), 14, BF
'picVDW.Line (2, 72)-(10, 71), 14, BF
'picVDW.Line (10, 80)-(9, 71), 14, BF
'picVDW.Line (7, 71)-(8, 68), 14, BF
'picVDW.Line (7, 68)-(11, 67), 14, BF
'picVDW.Line (10, 67)-(11, 60), 14, BF
'
'
'picVDW.Line (15, 80)-(23, 81), 14, BF
'picVDW.Line (15, 80)-(16, 60), 14, BF
'picVDW.Line (15, 60)-(23, 61), 14, BF
'picVDW.Line (15, 70)-(20, 71), 14, BF
'
'picVDW.Line (28, 80)-(36, 81), 14, BF
'picVDW.Line (28, 80)-(29, 40), 14, BF
'picVDW.Line (28, 70)-(32, 71), 14, BF
'
'picVDW.Line (43, 75)-(51, 74), 14, BF
'picVDW.Line (51, 74)-(50, 68), 14, BF
'picVDW.Line (41, 67)-(51, 68), 14, BF
'picVDW.Line (41, 67)-(42, 61), 14, BF
'picVDW.Line (41, 60)-(52, 61), 14, BF
'
'picVDW.Line (53, 75)-(61, 74), 14, BF
'picVDW.Line (53, 75)-(54, 60), 14, BF
'picVDW.Line (53, 60)-(61, 61), 14, BF
'picVDW.Line (61, 75)-(60, 60), 14, BF
'
'picVDW.Line (63, 75)-(64, 60), 14, BF
'
'picVDW.Line (72, 75)-(73, 60), 14, BF
'picVDW.Line (66, 75)-(72, 74), 14, BF
''
'picVDW.ForeColor = QBColor(14)
''LOCATE 22, 15
'picVDW.CurrentX = 22: picVDW.CurrentY = 15
'picVDW.Print "REFRACTION BASED ON THE MODIFIED US1976 ATMOSPHERE"
'picVDW.CurrentX = 27: picVDW.CurrentY = 50
'picVDW.Print "AUTHOR: Siebren van der Werf"
'picVDW.CurrentX = 28: picVDW.CurrentY = 50
'picVDW.Print "last update: March 2019"
'picVDW.ForeColor = QBColor(15)
'Sleep 3000 'sleep for 3 seconds
''Sleep 3
''=========================================
''SCREEN 12
'picVDW.Cls
'picVDW.Print "==========RAYTRACING IN THE MODIFIED US1976 ATMOSPHERE============="
'picVDW.Print "The starting template is that of the US 1976 Standard atmosphere, as described"
'picVDW.Print "for instance in the Handbook of Chemistry and Physics, 81th ed., 2000 - 2001."
'picVDW.Print "In this program some parameters may be adjusted. These are: the observer's"
'picVDW.Print "height, the temperature and pressure at ground level, the wave length of"
'picVDW.Print "the light, the relative humidity and the latitude."
'picVDW.Print "Setting HOBS=0, TGROUND=283.15,PRESS0=1010, RELHUM=0 and OBSLAT=52"
'picVDW.Print "will reproduce the new Nautical Almanac tables, as from 2005."
'picVDW.Print "More information may be found in:"
'picVDW.Print "Siebren Y. van der Werf, Raytracing and refraction in the modified"
'picVDW.Print "US1976 atmosphere, Applied Optics 42(2003)354-366."
'picVDW.Print "Backward raytracing up to 85 km is done using 4th order Runge-Kutta"
'picVDW.Print "numerical integration, using path length as the integration variable."
'picVDW.Print "Details on this method are given in:"
'picVDW.Print "Siebren Y. van der Werf, Comment on `Improved ray tracing air mass"
'picVDW.Print "numbers model', Applied Optics 47(2008)153-156."
'picVDW.Print "Up till a height HMAXT, which is asked for as an input, the calculation will"
'picVDW.Print "be made in a number of steps that must be specified: NSTEPS."
'picVDW.Print "From HMAXT till 100 km the step size will be gradually increased."
'picVDW.Print "HMAXT is further used for preparing a fine-step lookup table for the atm."
'picVDW.Print "pressure and for displaying the ray's curvature for H=0-HMAXT."
'picVDW.Print "Natural constants are taken from the Handbook of Chemistry and Physics, 81th ed."
'picVDW.Print "Refractivities for dry air and for water vapor are taken from:"
'picVDW.Print "P.E. Ciddor, Refractive index of air: new equations for the visible and near"
'picVDW.Print "infrared, Applied Optics 35(1996)1566-1573."
'picVDW.ForeColor = QBColor(14)
'picVDW.CurrentX = 28: picVDW.CurrentY = 2: picVDW.Print " press any key to proceed"
'picVDW.ForeColor = QBColor(15)
'KeyPressed = 0
'Do While KeyPressed = 0
''in$ = INPUT$(1)
'    DoEvents
'Loop
'KeyPressed = 0
'picVDW.Cls
''=================LIST OF CONSTANTS===========================
'RBOLTZ = 8314.472  'Univ. gas const. = Avogadro's number x Boltzmann's const.
'AMASSD = 28.964  'Molar weight of dry air
'AMASSW = 18.016  'Molar weight of water
'REARTH = 6356766# 'The Earth's mean radius
'PI = 4 * Atn(1)
'RADCON = PI / (60 * 180) 'Converts arcminutes into radians
'HLIMIT = 100000# 'Maximum height till which the rays are followed
'HMAXP1 = 30000 'height where f.FNDPD2 (steps of 10 m) takes over from f.FNDPD1 (steps of 1 m)
''===================OPTION-MENU OF FORMULA SATURATED VAPOR PRESSURE: =====================
''OPTVAP=1: PL2, POWER LAW
''OPTVAP=2: CC2, CLAUSIUS-CLAPEYRON 2 PAR.
''OPTVAP=3: CC4, CLAUSIUS-CLAPEYRON 4 PAR
''OPTVAP=4: ST, SACKUR-TETRODE, 4 PAR.
'OPTVAP = 4
''============INITIALIZATION: THE US 1976 STANDARD ATMOSPHERE=======
''SCREEN 12
'picVDW.Cls
'HOBS = 0#
'TGROUND = 283.15
'HMAXT = 1000#
'TLOW = 0#
'THIGH = 400#
'Press0 = 1010#
'RELHUM = 0#
'BETALO = 0
'BETAHI = 60
'BETAST = 10
'WAVELN = 0.574
'OBSLAT = 52#
'NSTEPS = 10000
''====================================================================
'PARAMETERLIST:
''SCREEN 12
'picVDW.Cls
'picVDW.CurrentX = 5: picVDW.CurrentX = 2: picVDW.Print "===============SPECIFICATION OF THE CASE================"
'picVDW.CurrentX = 7: picVDW.CurrentX = 2: picVDW.Print "01. HOBS = "; HOBS; "          Observer's eye height (m)"
'picVDW.CurrentX = 8: picVDW.CurrentX = 2: picVDW.Print "02. TGROUND = "; TGROUND; "  Temperature (K) at height = 0"
'picVDW.CurrentX = 9: picVDW.CurrentX = 2: picVDW.Print "03. HMAXT = "; HMAXT; "  Max. height (m) for curvature display (multiple of 100 m) "
'picVDW.CurrentX = 10: picVDW.CurrentX = 2: picVDW.Print "04. TLOW = "; TLOW; "        Show temperature profile from TLOW (K) till.."
'picVDW.CurrentX = 11: picVDW.CurrentX = 2: picVDW.Print "05. THIGH = "; THIGH; "       highest value (K) for which to show T-profile"
'picVDW.CurrentX = 12: picVDW.CurrentX = 2: picVDW.Print "06. PRESS0 = "; Press0; "  Atmospheric pressure (hPa) at h=0"
'picVDW.CurrentX = 13: picVDW.CurrentX = 2: picVDW.Print "07. RELHUM = "; RELHUM; "        Relative humidity (%) in troposphere"
'picVDW.CurrentX = 14: picVDW.CurrentX = 2: picVDW.Print "08. BETALO = "; BETALO; "         Lowest apparent altitude (arcmin)"
'picVDW.CurrentX = 15: picVDW.CurrentX = 2: picVDW.Print "09. BETAHI = "; BETAHI; "         Highest apparent altitude (arcmin)"
'picVDW.CurrentX = 16: picVDW.CurrentX = 2: picVDW.Print "10. BETAST = "; BETAST; "        Stepsize in apparent altitude (arcmin)"
'picVDW.CurrentX = 17: picVDW.CurrentX = 2: picVDW.Print "11. WAVELN = "; WAVELN; "      Wavelength (mu), 0.65=R,0.589=Y (Sodium),0.52=G"
'picVDW.CurrentX = 18: picVDW.CurrentX = 2: picVDW.Print "12. OBSLAT = "; OBSLAT; "       Latitude of observer ((degrees)"
'picVDW.CurrentX = 19: picVDW.CurrentX = 2: picVDW.Print "13. NSTEPS = "; NSTEPS; "      Number of steps up till HMAXT"
'MODIFYPARAMETERS:
'picVDW.ForeColor = QBColor(14)
'picVDW.CurrentX = 25: picVDW.CurrentX = 10: picVDW.Print "Run with these parameters (Y/any other key) ?"
''a$ = INPUT$(1)
'KeyPressed = 0
'Do While KeyPressed = 0
'   DoEvents
'Loop
'If Str(KeyPressed) = "Y" Or a$ = "y" Then
'    GoTo ACCEPTPARAMETERS
'Else
'End If
'
''picVDW.CurrentX = 25: picVDW.CurrentX = 10: picVDW.Print "Input the parameter number (2 digits) to change"
''Color 15
''KeyPressed = 0
''Do While KeyPressed = 0
''   dovents
''Loop
'''ANSWER$ = INPUT$(2)
''PARNUM = KeyPressed 'Val(ANSWER$)
''Loop Until PARNUM > 0 And PARNUM < 14
''picVDW.Cls
''On PARNUM GoTo P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, P13
''P01:
''INPUT "HOBS = "; HOBS
''GoTo PARAMETERLIST
''P02:
''INPUT "TGROUND = "; TGROUND
''GoTo PARAMETERLIST
''P03:
''INPUT "HMAXT = "; HMAXT
''GoTo PARAMETERLIST
''P04:
''INPUT "TLOW = "; TLOW
''GoTo PARAMETERLIST
''P05:
''INPUT "THIGH = "; THIGH
''GoTo PARAMETERLIST
''P06:
''INPUT "PRESS0 = "; PRESS0
''GoTo PARAMETERLIST
''P07:
''INPUT "RELHUM = "; RELHUM
''GoTo PARAMETERLIST
''P08:
''INPUT "BETALO = "; BETALO
''GoTo PARAMETERLIST
''P09:
''INPUT "BETAHI = "; BETAHI
''GoTo PARAMETERLIST
''P10:
''INPUT "BETAST = "; BETAST
''GoTo PARAMETERLIST
''P11:
''INPUT "WAVELN = "; WAVELN
''GoTo PARAMETERLIST
''P12:
''INPUT "OBSLAT = "; OBSLAT
''GoTo PARAMETERLIST
''P13:
''INPUT "NSTEPS = "; NSTEPS
''GoTo PARAMETERLIST
'ACCEPTPARAMETERS:
'RELH = RELHUM / 100
''====CALCULATE HCROSS ========
'HCROSS = (TGROUND - 216.65) / 0.0065
'picVDW.ForeColor = QBColor(14)
'picVDW.CurrentX = 28: picVDW.CurrentY = 2: picVDW.Print " press any key to proceed"
'picVDW.ForeColor = QBColor(15)
''in$ = INPUT$(1)
''SCREEN 0: WIDTH 80
'picVDW.Cls
'
''===================================================
'
''OPEN "O", #1, "REF2017.OUT" 'BETAM, REFRAC, TRUALT, AIRDRY, AIRVAP
'fileout1% = FreeFile
'Open App.Path & "\REF2017.OUT" For Output As #fileout1%
''OPEN "O", #3, "REF2017-ATM.OUT" 'height,temperature,dry-air pressure, vapor pressure, gravitational acceleration
'fileout3% = FreeFile
'Open App.Path & "\REF2017-ATM.OUT" For Output As #fileout3%
''===================================================
'
''============LIST OF DERIVED CONSTANTS==================================
'OLAT = OBSLAT * 60 * RADCON
'GRAVC = 9.780356 * (1 + 0.0052885 * (Sin(OLAT)) ^ 2 - 0.0000059 * (Sin(2 * OLAT)) ^ 2)           'Gravitat. const.
'BD = GRAVC * AMASSD / RBOLTZ 'Dry air exponent
'BW = GRAVC * AMASSW / RBOLTZ 'Water exponent
's2 = 1 / WAVELN ^ 2
''CIDDOR'S FORMULAS FOR DRY AIR AND WATER VAPOUR
'AD = 0.00000001 * (5792105# / (238.0185 - s2) + 167917# / (57.362 - s2)) * 288.15 / 1013.25
'AW = 0.00000001022 * (295.235 + 2.6422 * s2 - 0.03238 * s2 ^ 2 + 0.004028 * s2 ^ 3) * 293.15 / 13.33
'MAXIND = HMAXT + 1
'
''==================================================================
''==FILL ARRAY PRESSD1 (PARTIAL PRESSURE OF DRY AIR)==
''== IN STEPS OF 1 METER========
'PRESSD1(1) = Press0 - RELH * f.VAPOR(0#)
'For i = 1 To (31000 + 15) Step 1
'    '===Fill PRESSD1. I=1 -> H=0, I=2 -> h=1 m. etc..===
'    '===INTEGRATION BY 4TH ORDER RUNGE-KUTTA===
'    HSTEP = 1#
'    'STEP 1
'    P1 = PRESSD1(i)
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK1 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 1
'    'STEP 2
'    H = (i - 1) / 1# + HSTEP / 2
'    P1 = PRESSD1(i) + FK1 * HSTEP / 2
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK2 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 2
'    'STEP 3
'    H = (i - 1) / 1# + HSTEP / 2
'    P1 = PRESSD1(i) + FK2 * HSTEP / 2
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK3 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 3
'    'STEP 4
'    H = (i - 1) / 1# + HSTEP
'    P1 = PRESSD1(i) + FK3 * HSTEP
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK4 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 4
'    PRESSD1(i + 1) = PRESSD1(i) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
'Next i
''===FIND PDM1 AT -1 METER===
'HSTEP = -1#
''STEP 1
'H = 0#
'P1 = PRESSD1(1)
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK1 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 1
''STEP 2
'H = HSTEP / 2
'P1 = PRESSD1(1) + FK1 * HSTEP / 2
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK2 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 2
''STEP 3
'H = HSTEP / 2
'P1 = PRESSD1(1) + FK2 * HSTEP / 2
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK3 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 3
''STEP 4
'H = HSTEP
'P1 = PRESSD1(1) + FK3 * HSTEP
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK4 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 4
'PDM01 = PRESSD1(1) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
''==================================================================
''=======END OF STORAGE PRESSURE ARRAY PRESSD1 FOR DRY AIR========
'
''==FILL ARRAY PRESSD2 (PARTIAL PRESSURE OF DRY AIR)==
''== IN STEPS OF 10 METER========
'PRESSD2(1) = PRESSD1(1)
'I2LOW = CInt(HMAXP1 / 10)
'For i = 0 To I2LOW Step 1
'    PRESSD2(i + 1) = PRESSD1(10 * i + 1)
'Next i
'For i = I2LOW To (HLIMIT / 10 + 5) Step 1
'    '===Fill PRESSD2. I=1 -> H=0, I=2 -> h=1 m. etc..===
'    '===INTEGRATION BY 4TH ORDER RUNGE-KUTTA===
'    HSTEP = 10#
'    'STEP 1
'    H = (i - 1) * 10
'    P1 = PRESSD2(i)
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK1 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 1
'    'STEP 2
'    H = (i - 1) * 10 + HSTEP / 2
'    P1 = PRESSD2(i) + FK1 * HSTEP / 2
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK2 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 2
'    'STEP 3
'    H = (i - 1) * 10 + HSTEP / 2
'    P1 = PRESSD2(i) + FK2 * HSTEP / 2
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK3 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 3
'    'STEP 4
'    H = (i - 1) * 10 + HSTEP
'    P1 = PRESSD2(i) + FK3 * HSTEP
'    P2 = RELH * f.VAPOR(H)
'    DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'    FK4 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
'    'END STEP 4
'    PRESSD2(i + 1) = PRESSD2(i) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
'Next i
''===FIND PDM10 AT -10 METER===
'HSTEP = -10#
''STEP 1
'H = 0#
'P1 = PRESSD2(1)
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK1 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 1
''STEP 2
'H = HSTEP / 2
'P1 = PRESSD2(1) + FK1 * HSTEP / 2
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK2 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 2
''STEP 3
'H = HSTEP / 2
'P1 = PRESSD2(1) + FK2 * HSTEP / 2
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK3 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 3
''STEP 4
'H = HSTEP
'P1 = PRESSD2(1) + FK3 * HSTEP
'P2 = RELH * f.VAPOR(H)
'DP2DH = RELH * f.DVAPDT(H) * f.DTDH(H)
'FK4 = -DP2DH - BD * P1 * f.GRAVRAT(H) / f.Temp(H) - BW * P2 * f.GRAVRAT(H) / f.Temp(H)
''END STEP 4
'PDM10 = PRESSD2(1) + (HSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
''==================================================================
''=======END OF STORAGE PRESSURE ARRAY FOR DRY AIR========
'
''===================Temperature profile====================
'
''SCREEN 12
'picVDW.Cls
'picVDW.Scale (TLOW, 0)-(THIGH, HLIMIT)
'Line (TLOW, 0)-(THIGH, HLIMIT), 1, B
'Line (TLOW, HCROSS)-(THIGH, HCROSS), 1
'Line (TLOW, 20000)-(THIGH, 20000), 1
'Line (TLOW, 32000)-(THIGH, 32000), 1
'Line (TLOW, 47000)-(THIGH, 47000), 1
'Line (TLOW, 51000)-(THIGH, 51000), 1
'Line (TLOW, 71000)-(THIGH, 71000), 1
'Line (TLOW, 85000)-(THIGH, 85000), 1
'LOCATE 3, 15: Print "TEMPERATURE PROFILE :"; "   T="; Int(TLOW); "K"; "-"; Int(THIGH); "K";
'LOCATE 24, 2: Print "20 km";
'LOCATE 20, 2: Print "32 km";
'LOCATE 17, 2: Print "47 km";
'LOCATE 15, 2: Print "51 km";
'LOCATE 9, 2: Print "71 km";
'LOCATE 5, 2: Print "85 km";
'LOCATE 26, 2: Print ; "height troposphere:"
'LOCATE 27, 2: Print ; Int(HCROSS) / 1000; "km";
'For i = 1 To HLIMIT Step 100
'    H = i - 1
'    R1 = f.Temp(H)
'    picVDW.Circle (R1, H), 0.01, 7
'Next i
''==================================================================
'picVDW.ForeColor = QBColor(14)
'picVDW.CurrentX = 28: picVDW.CurrentY = 2: picVDW.Print " press any key to proceed"
'picVDW.ForeColor = QBColor(15)
'KeyPressed = 0
'Do While KeyPressed = 0
''in$ = INPUT$(1)
'    DoEvents
'Loop
'KeyPressed = 0
'picVDW.Cls
''SCREEN 0: WIDTH 80
''=================R_curve/R-Earth======================
'
''The ratio R_curve/R-Earth will be shown on a scale -5 to 5 (horizontal)
''versus elevation from 0 to the maximum tabulated elevation HMAXT.
'
''SCREEN 12
'picVDW.Cls
'picVDW.Scale (-10, 0)-(10, HMAXT)
'Line (-10, 0)-(10, HMAXT), 1, B
'LOCATE 28, 8: Print "-8"
'LOCATE 28, 12: Print "-7"
'LOCATE 28, 16: Print "-6"
'LOCATE 28, 20: Print "-5"
'LOCATE 28, 24: Print "-4"
'LOCATE 28, 28: Print "-3"
'LOCATE 28, 32: Print "-2"
'LOCATE 28, 36: Print "-1"
'LOCATE 28, 40: Print "0"
'LOCATE 28, 44: Print "1"
'LOCATE 28, 48: Print "2"
'LOCATE 28, 52: Print "3"
'LOCATE 28, 56: Print "4"
'LOCATE 28, 60: Print "5"
'LOCATE 28, 64: Print "6"
'LOCATE 28, 68: Print "7"
'LOCATE 28, 72: Print "8"
'For i = -9 To 9 Step 1
'    Line (i, 0)-(i, HMAXT), 1
'Next i
'LOCATE 2, 30: Print "Horizontal: R_curv/R_Earth for a horizontal ray, Vertical: H="; 0; "-"; Int(HMAXT); "m"
'
'For i = 1 To HMAXT Step 1
'    H = i - 1
'    R1 = 1 / (f.RCINV(H) * REARTH)
'    picVDW.Circle (R1, H), 0.01, 7
'Next i
'Color 14
'LOCATE 26, 40: Print "press any key to proceed"
'Color 15
'in$ = INPUT$(1)
''SCREEN 0: WIDTH 80
'picVDW.Cls
''PRINT OUTPUT FILE #3, DATA ON THE ATMOSPHERE
'For H = 0 To 80000 Step 1000
'    PDRY = f.FNDPD2(H)
'    PVAP = RELH * f.VAPOR(H)
'    GRAVH = GRAVC * f.GRAVRAT(H)
'    Print #3, USING; " #####.######"; H / 1000, f.Temp(H), PDRY, PVAP, GRAVH
'Next H
'Close #3
'
''=============SETUP GRAPHICS PART================================
'
''SCREEN 12
'picVDW.Cls
'LRANGE = 1.25 * f.GUESSL(BETALO, HLIMIT)
'LDKM = Int(LRANGE / 1000)
'picVDW.Scale (0, 0)-(LRANGE, HLIMIT)
'picVDW.Line (0, 0)-(LRANGE, HLIMIT), 1, B
'
'picVDW.CurrentX = 2: picVDW.CurrentY = 2: picVDW.Print "Paths of the rays. Horizontal distance 0 - " & Str(LDKM) & "km, Vertical:  0 -100 km"
'
'
''===================CALCULATION====================================
'picVDW.ForeColor = QBColor(14)
'picVDW.CurrentX = 27: picVDW.CurrentY = 40: picVDW.Print "To interrupt, hit the spacebar"
'picVDW.ForeColor = QBColor(15)
'For BETAM = BETALO To BETAHI Step BETAST
'    DPATH = f.GUESSP(BETAM, HMAXT) / NSTEPS
'    picVDW.CurrentX = 3: picVDW.CurrentY = 2: picVDW.Print "Step size (m) = " & Str(DPATH)
'    '
'    DIST = 0#
'    REFRAC = 0#
'    AIRDRY = 0#
'    AIRVAP = 0#
'    PHI1 = 0#
'    BETA1 = BETAM * RADCON
'    R1 = REARTH + HOBS
'    h1 = HOBS
'    picVDW.PSet (DIST, h1), 10
'    Path = -DPATH
'    '===============================
'    'DO-LOOP OVER PATH
'    Do
'        Path = Path + DPATH
'        '
'        ' FOURTH ORDER RUNGE-KUTTA
'        ' THE THREE COUPLED FIRST ORDER DIFFERENTIAL EQUATIONS ARE:
'        ' 1)  dPHI/dPATH=cos(BETA)/R
'        ' 2)  dR/dPATH = sin(BETA)
'        ' 3)  dBETA/dPATH = cos(BETA)[1//R+(1/n) dn/dR]
'        ' WITH (1/n) [dn/dR]=f.RCINV(H)
'        '
'        ' STEP 1
'        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
'        '
'        FKP1 = Cos(BETA1) / R1
'        FKR1 = Sin(BETA1)
'        FKB1 = Cos(BETA1) * (1 / R1 + f.RCINV(h1))
'        FKAD1 = f.DENSDRY(h1)
'        FKAV1 = f.DENSVAP(h1)
'        '
'        'END OF FIRST STEP
'        '
'        'STEP 2
'        PHINEW = PHI1 + FKP1 * DPATH / 2
'        RNEW = R1 + FKR1 * DPATH / 2
'        BETANEW = BETA1 + FKB1 * DPATH / 2
'        HNEW = RNEW - REARTH 'elevation halfway step
'        '
'        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
'        '
'        FKP2 = Cos(BETANEW) / RNEW
'        FKR2 = Sin(BETANEW)
'        FKB2 = Cos(BETANEW) * (1 / RNEW + f.RCINV(HNEW))
'        FKAD2 = f.DENSDRY(HNEW)
'        FKAV2 = f.DENSVAP(HNEW)
'        '
'        'END OF SECOND STEP
'        '
'        'STEP 3
'        PHINEW = PHI1 + FKP2 * DPATH / 2
'        RNEW = R1 + FKR2 * DPATH / 2
'        BETANEW = BETA1 + FKB2 * DPATH / 2
'        HNEW = RNEW - REARTH 'elevation halfway step
'        '
'        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
'        '
'        FKP3 = Cos(BETANEW) / RNEW
'        FKR3 = Sin(BETANEW)
'        FKB3 = Cos(BETANEW) * (1 / RNEW + f.RCINV(HNEW))
'        FKAD3 = f.DENSDRY(HNEW)
'        FKAV3 = f.DENSVAP(HNEW)
'
'        '
'        'END OF THIRD STEP
'        '
'        'STEP 4
'        PHINEW = PHI1 + FKP3 * DPATH
'        RNEW = R1 + FKR3 * DPATH
'        BETANEW = BETA1 + FKB3 * DPATH
'        HNEW = RNEW - REARTH
'        H = HNEW 'elevation at full step
'        '
'        ' FIND THE RUNGE-KUTTA K-COEFFICIENTS
'        '
'        FKP4 = Cos(BETANEW) / RNEW
'        FKR4 = Sin(BETANEW)
'        FKB4 = Cos(BETANEW) * (1 / RNEW + f.RCINV(HNEW))
'        FREF4 = Cos(BETANEW) * fRCINV(HNEW)
'        FKAD4 = fDENSDRY(HNEW)
'        FKAV4 = fDENSVAP(HNEW)
'        '
'        'END OF FOURTH AND FINAL STEP
'        '
'        'FIND R2 AND PHI2
'        PHI2 = PHI1 + (FKP1 + 2 * FKP2 + 2 * FKP3 + FKP4) * DPATH / 6
'        R2 = R1 + (FKR1 + 2 * FKR2 + 2 * FKR3 + FKR4) * DPATH / 6
'        BETA2 = BETA1 + (FKB1 + 2 * FKB2 + 2 * FKB3 + FKB4) * DPATH / 6
'        AIRDRY = AIRDRY + (FKAD1 + 2 * FKAD2 + 2 * FKAD3 + FKAD4) * DPATH / 6
'        AIRVAP = AIRVAP + (FKAV1 + 2 * FKAV2 + 2 * FKAV3 + FKAV4) * DPATH / 6
'        h2 = R2 - REARTH
'        DREFR = -BETA2 + BETA1 + PHI2 - PHI1
'        'Stop this ray if it hits the ground, or if it seems to never end
'        'as may occur for a Novaya-Zemlya atmosphere.
'        If h2 < 0 Or (DIST > 10 * fGUESSL(0, HMAXT) And h2 < HMAXT) Then
'            GoTo NEXTRAY
'        Else
'            If h2 > HLIMIT Then
'                ' H2 PASSED HLIMIT METER
'                TRUALT = BETAM - REFRAC
'                Print #1, USING; " #####.######"; BETAM; REFRAC; TRUALT; AIRDRY; AIRVAP
'                '  BETAM, REFRAC,TRUALT, AIRDRY AND AIRVAP HAVE BEEN STORED IN REF2017.OUT
'                GoTo NEXTRAY
'            Else
'            End If
'        End If
'
'        '==============================
'        picVDW.Scale (0, 0)-(LRANGE, HLIMIT)
'        DIST = DIST + REARTH * (PHI2 - PHI1)
'        If h2 > HLIMIT Then
'            GoTo NORAYPRINT
'        Else
'        End If
'        If DIST > LRANGE Then
'            GoTo NORAYPRINT
'        Else
'        End If
'        picVDW.PSet (DIST, h2), 10
'NORAYPRINT:
'        REFRAC = REFRAC + DREFR / RADCON
'        PHI1 = PHI2
'        R1 = R2
'        h1 = h2
'        BETA1 = BETA2
'        If h2 > HMAXT Then
'            DPATH = fGUESSP(BETAM, h2) / NSTEPS
'        Else
'        End If
'
'        'END DO-LOOP OVER PATH
'    Loop
'NEXTRAY:
'    a$ = INKEY$
'    If a$ = Chr$(32) Then GoTo ENDOFCALCULATION
'Next BETAM
'ENDOFCALCULATION:
''========================END OF CALCULATION=====================
'Close #1
''========HOLD THE SCREEN TILL HITTING ANY KEY==================
'Color 14
'LOCATE 28, 50: Print "press any key to proceed"
'Color 15
'in$ = INPUT$(1)
'picVDW.Cls
'
''Plot refraction (vertical), versus apparent angle
''
'If BETAHI < 312.5 Then
'    BMIN = -6.25
'    BMAX = 300
'    NTICKS = 10
'Else
'    If BETAHI < 625 Then
'        BMIN = -12.5
'        BMAX = 600
'        NTICKS = 30
'    Else
'        If BETAHI < 1350 Then
'            BMIN = -25
'            BMAX = 1200
'            NTICKS = 60
'        Else
'            If BETAHI < 2700 Then
'                BMIN = -50
'                BMAX = 2700
'                NTICKS = 150
'            Else
'                If BETAHI <= 5400 Then
'                    BMIN = -100
'                    BMAX = 5400
'                    NTICKS = 300
'                Else
'                End If
'            End If
'        End If
'    End If
'End If
'picVDW.Scale (BMIN, -10)-(BMAX, 60)
'LOCATE 2, 5: Print "Vertical: refraction, tickmarks 10 arcmin"
'If BETAHI <= 300 Then
'    LOCATE 3, 5: Print "Horizontal: apparent observer angle, 0 - 5 deg"
'    LOCATE 4, 5: Print "Horizontal tickmarks 10 arcmin"
'Else
'    If BETAHI <= 600 Then
'        LOCATE 3, 5: Print "Horizontal: apparent observer angle, 0 - 10 deg"
'        LOCATE 4, 5: Print "Horizontal tickmarks 30 arcmin"
'    Else
'        If BETAHI <= 1200 Then
'            LOCATE 3, 5: Print "Horizontal: apparent observer angle, 0 - 20 deg"
'            LOCATE 4, 5: Print "Horizontal tickmarks 1 deg"
'        Else
'            If BETAHI <= 2700 Then
'                LOCATE 3, 5: Print "Horizontal: apparent observer angle, 0 - 45 deg"
'                LOCATE 4, 5: Print "Horizontal tickmarks 2.5 deg "
'            Else
'                If BETAHI <= 5400 Then
'                    LOCATE 3, 5: Print "Horizontal: apparent observer angle, 0 - 90 deg"
'                    LOCATE 4, 5: Print "Horizontal tickmarks 5 deg"
'                Else
'                End If
'            End If
'        End If
'    End If
'End If
'
'Line (BMIN, -10)-(BMAX, 60), 1, B
'Line (BMIN, 0)-(BMAX, 0), 1
'For XX = NTICKS To BMAX Step NTICKS
'    Line (XX, -1)-(XX, 1), 1
'Next XX
'Line (0, -10)-(0, 60), 1
'For YY = -10 To 60 Step 10
'    Line (-NTICKS / 5, YY)-(NTICKS / 5, YY), 1
'Next YY
''
'OPEN "I", #2, "REF2017.OUT" 'BETAM, REFRAC, TRUALT, AIRDRY, AIRVAP
'Do While Not EOF(2)
'    Input #2, BETAM, REFRAC, TRUALT, AIRDRY, AIRVAP
'    picVDW.Circle (BETAM, REFRAC), 0.1 * NTICKS, 7
'Loop
'Close #2
'
''========HOLD THE SCREEN TILL HITTING ANY KEY==================
'Color 14
'LOCATE 28, 40: Print "press any key to proceed"
'Color 15
'in$ = INPUT$(1)
''SCREEN 0: WIDTH 80
'
'picVDW.Cls
'LOCATE 3, 5: Print "Output written to REF2017.OUT as: BETA, REFRAC, TRUALT, AIRDRY, AIRVAP"
'LOCATE 4, 5: Print "where BETA = apparent altitude (arcmin)"
'LOCATE 5, 5: Print "      REFRAC = path integral of refraction (arcmin)"
'LOCATE 6, 5: Print "      TRUALT = true altitude (arcmin)"
'LOCATE 7, 5: Print "      AIRDRY = path integral of dry-air density (kg/m2/100) "
'LOCATE 8, 5: Print "      AIRVAP = path integral of water vapor density (kg/m2/100)"
'Color 14
'LOCATE 15, 5: Print "Change parameters? (Y/any other key), else end of program "
'Color 15
'a$ = INPUT$(1)
'If a$ = "Y" Or a$ = "y" Then
'    GoTo PARAMETERLIST
'Else
'End If
'
'End Sub
'
''===================FUNCTIONS=================================
'Function fACS(X As Double) As Double
''Arccosine
'fACS = 2 * Atn(1) - Atn(X / Sqr(1# - X * X))
'End Function
'
'Function fASN(X As Double) As Double
''Arcsine
'fASN = Atn(X / Sqr(1# - X * X))
'End Function
'
'
'Function fDLNTDH(H As Double) As Double
''The deriative of ln(T): (1/T)(dT/dh)
'fDLNTDH = fDTDH(H) / fTEMP(H)
'End Function
'
'Function fDTDH(H As Double) As Double
''DefDbl A-H, O-Z
'Dim dh As Double, T1 As Double, T2 As Double, T3 As Double, T4 As Double
'dh = 0.01
'T1 = fTEMP(H - 3 * dh / 2)
'T2 = fTEMP(H - dh / 2)
'T3 = fTEMP(H + dh / 2)
'T4 = fTEMP(H + 3 * dh / 2)
'fDTDH = ((T3 - T2) * (27 / 24) - (T4 - T1) / 24) / dh
'End Function
'
'Function fFNDPD1(H)
''Interpolation in the array PRESSD1
''DefDbl A-H, O-Z
''SHARED HLIMIT, PDM1, RELH, BD, BW, HMAXP1
'If (H < (HMAXP1 + 1) And H >= 0) Then
'    y = H
'    YSTEP = y - Int(y)
'    'STEP 1
'    i = Int(y)
'    P1 = PRESSD1(i + 1)
'    P2 = RELH * fVAPOR(y)
'    DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'    FK1 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
'    'END STEP 1
'    'STEP 2
'    y = i + YSTEP / 2
'    P1 = PRESSD1(i + 1) + FK1 * YSTEP / 2
'    P2 = RELH * fVAPOR(y)
'    DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'    FK2 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
'    'END STEP 2
'    'STEP 3
'    y = i + YSTEP / 2
'    P1 = PRESSD1(i + 1) + FK2 * YSTEP / 2
'    P2 = RELH * fVAPOR(y)
'    DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'    FK3 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
'    'END STEP 3
'    'STEP 4
'    y = i + YSTEP
'    P1 = PRESSD1(i + 1) + FK3 * YSTEP
'    P2 = RELH * fVAPOR(y)
'    DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'    FK4 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
'    'END STEP 4
'    fFNDPD1 = PRESSD1(i + 1) + (YSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
'Else
'End If
'End Function
'
'Function fFNDPD2(H)
''Interpolation in the array PRESSD2
''DefDbl A-H, O-Z
''SHARED HLIMIT, PDM10
'y = H
'YSTEP = (y - Int(y)) * 10
''STEP 1
'i = Int(y / 10)
'P1 = PRESSD2(i + 1)
'P2 = RELH * fVAPOR(y)
'DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'FK1 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
''END STEP 1
''STEP 2
'y = i + YSTEP / 2
'P1 = PRESSD2(i + 1) + FK1 * YSTEP / 2
'P2 = RELH * fVAPOR(y)
'DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'FK2 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
''END STEP 2
''STEP 3
'y = i + YSTEP / 2
'P1 = PRESSD2(i + 1) + FK2 * YSTEP / 2
'P2 = RELH * fVAPOR(y)
'DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'FK3 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
''END STEP 3
''STEP 4
'y = i + YSTEP
'P1 = PRESSD2(i + 1) + FK3 * YSTEP
'P2 = RELH * fVAPOR(y)
'DP2DY = RELH * fDVAPDT(y) * fDTDH(y)
'FK4 = -DP2DY - BD * P1 * fGRAVRAT(y) / fTEMP(y) - BW * P2 * fGRAVRAT(y) / fTEMP(y)
''END STEP 4
'fFNDPD2 = PRESSD2(i + 1) + (YSTEP / 6) * (FK1 + 2 * FK2 + 2 * FK3 + FK4)
'PD2READY:
'End Function
'
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : cmdMenat_Click
'' Author    : Dr-John-K-Hall
'' Date      : 2/18/2019
'' Purpose   : VB6 enactment of Menat ray tracing code, source: menat.cpp
''---------------------------------------------------------------------------------------
''
'Private Sub cmdMenat_Click()
''/* menat.f -- translated by f2c (version 20021022).
''   You must link the resulting object file with the libraries:
''    -lf2c -lm   (in that order)
''*/
''
''//#include "f2c.h"
''#include <stdio.h>
''#include <math.h>
''#include <stdio.h>
''#include <stdlib.h>
''#include <ctype.h>
''#include <string.h>
''
''/* Common Block Declarations */
''
''struct {
''Dim hj(50) As Double, tj(50) As Double, pj(50) As Double, at(50) As Double, ct(50) As Double
'
'Dim zz_1 As zz
'
''} zz_
''
''#define zz_1 zz_
''
''Dim layerheights(50) As Double
''Dim temps(50) As Double
''Dim press(50) As Double
'Dim numLayers As Integer
''
''//declare extermanl file
''
''/* Table of constant values */
''
''/*       PROGRAM MENAT--CALCULATES ASTRONOMICAL ATMOSPHERIC REFRACTION */
''/*       FOR ANY OBSERVER HEIGHT by using ray tracing through a */
''/*       simplified layered atmosphered */
''/*       ITS OUTPUT IS REPRESENTED in the files menatsum.ren, menatwin.ren */
''/* Main program */ //int MAIN__(void)
''int main(int argc, char* argv[])
'Dim pi2 As Double, fr As Double, hz As Double, dh As Double, hsof As Double, epg As Double
'Dim co As Double, pie As Double, ramg As Double, q3 As Double, q6 As Double, ra As Double
'Dim b As Double, d__ As Double, e As Double, g As Double, h__ As Double, jstop As Long, j As Long
'Dim k As Long, l As Long, n As Long, KWAV As Double, KMIN As Double, KMAX As Double, KSTEP As Double
'Dim s As Double, T As Double, a3 As Double, e1 As Double, e2 As Double, g1 As Double, d6 As Double, s1 As Double, s2 As Double
''    //char ch[1]
'Dim dg As Double, bn As Double, el As Double, en As Double, bz As Double, cz As Double, em As Double, StatusMes As String
'
''    //int nn
'Dim rt As Double
'Dim XP  As Double
''   On Error GoTo cmdMenat_Click_Error
'
'XP = 0#
'Dim nz As Long
'Dim ru As Double
''    //char ch1[1], ch2[2], ch3[3]
'Dim ep1 As Double, at2 As Double, en1 As Double, dt1 As Double
''    //double en2
''    //double en3//
''    //double en4,
'Dim hz1 As Double, hz2 As Double
''    //logical beg
'Dim den As Double
''    //logical neg
'Dim hen As Double, dtg As Double, hev As Double
'Dim kgr As Integer
'Dim dhz As Double
''    //double ent[8000]  /* was [4][2000] */
'Dim sbn As Double
''    extern double fun_(double *)
'Dim entry As Double
'Dim isn As Integer
'isn = 1
'Dim iam As Integer
'iam = 0
'Dim epz As Double
''    extern double sqt_(double *, double *, double *)
''    extern int LoadAtmospheres(char filnam[])
'Dim hen1 As Double, epg1 As Double, epg2 As Double, dtg2 As Double, fieg As Double
''    //logical angl
''    //int mang,
'Dim nang As Integer
''    //int ccc
'Dim fiem As Double, hmin As Double, hmax As Double
'Dim nent As Integer, nhgt As Integer '//,mhgt
'Dim rsof As Double, epzm As Double
''    //char finam[10]
''    //double chisn
'Dim nkhgt As Integer
'Dim estep As Double
'Dim ANGLE As Double
'ANGLE = 0#
'Dim A1 As Double
'A1 = 0#
'Dim A2 As Double
'A2 = 0#
'Dim START As Boolean
'START = False
''    char filnam[255] = sempty
'Dim filnam As String
''    char chr[2] = sempty
'Dim ier As Integer
'ier = 0
''    int ier = 0
'
'Dim FNM As String, AtmType As Integer, AtmNumber As Integer, lpsrate As Double, tst As Double, pst As Double, NNN As Long
'
'   cmdCalc.Enabled = False
'   cmdRefWilson.Enabled = False
'   cmdMenat.Enabled = False
'
'   RefCalcType% = 2
'   CalcComplete = False
'
'STARTALT = Val(txtStartAlt.Text)
'DELALT = Val(txtDelAlt.Text)
'XMAX = Val(txtXmax.Text) * 1000 'convert km to meters
'PPAM = Val(txtPPAM.Text)
'KMIN = Val(txtKmin.Text)
'KMAX = Val(txtKmax.Text)
'KSTEP = Val(txtKStep.Text)
'STARTAZM = 19
'DELAZM = 32
'If INVFLAG = 1 Then
'   SINV = Val(txtSInv.Text)
'   EINV = Val(txtEInv.Text)
'   DTINV = Val(txtDInv.Text)
'   End If
'
'StatusMes = "Pixels per arcminute " & Str(PPAM) & ", Maximum height (degrees) " & Str(n / (120# * PPAM))
'Call StatusMessage(StatusMes, 1, 0)
'
'n_size = 500
'msize = 20 + Val(txtNumSuns.Text) * 32 * PPAM
'
'If Trim$(txtXSize.Text) <> sempty Then
'   msize = Val(txtXSize.Text)
'   End If
'If Trim$(txtYSize.Text) <> sempty Then
'   n_size = Val(txtYSize.Text)
'   End If
'
'Dim KA As Long
'For KA = 1 To NumSuns
'   ALT(KA) = STARTALT + CDbl(KA - 1) * DELALT
'   AZM(KA) = STARTAZM + CDbl(KA - 1) * DELAZM
'Next KA
'
'myfile$ = Dir(App.Path & "\test_M.dat")
'If myfile$ <> sempty Then
'   Kill App.Path & "\test_M.dat"
'   End If
'myfile$ = Dir(App.Path & "\tc_M.dat")
'If myfile$ <> sempty Then
'   Kill App.Path & "\tc_M.dat"
'   End If
'
'
'PI = 4# * Atn(1#) '3.141592654
'CONV = PI / (180# * 60#) 'conversion of minutes of arc to radians
'cd = PI / 180# 'conversion of degrees to radians
'ROBJ = 15# 'half size of sun in minutes of arc
''{
''    /* Initialized data */
''
' pi2 = PI * 0.5 '1.570796
' co = cd '0.017453293
' fr = 1.001
' hz = 0.801
' dh = 0.002
' hsof = 65#
' epg = 0#
'' hs[7] = { 0.,13.,18.,25.,47.,50.,70. }
' pie = PI '3.1415926
' ramg = 0.05729578
' q3 = 1000#
' q6 = 1000000#
''    /*
''    double hw[7] = { 0.,10.,19.,25.,30.,50.,70. }
''    double ts[7] = { 299.,215.5,216.5,224.,273.,276.,218. }
''    double tw[7] = { 284.,220.,215.,216.,217.,266.,231. }
''    double ps[7] = { 1013.,179.,81.2,27.7,1.29,.951,.067 }
''    double pw[7] = { 1018.,256.8,62.8,24.3,11.1,.682,.0467 }
''    */
'
'ra = 6371 ' 6378.1366     '6371.
'RE = ra * 1000#
''
''    /* System generated locals */
''    int i__1, i__2, i__3//, i__4
''    double d__1//, d__2
''
''    /* Local variables */
''    double b, d__, e, g, h__
''    int k, l, n
''    double s, t, a2, a3, e1, e2, g1, d6, s1, s2
''    //char ch[1]
''    double dg, bn, el, en, bz, cz, em
''    //int nn
''    double rt
''    double XP = 0.0
''    int nz
''    double ru
''    //char ch1[1], ch2[2], ch3[3]
''    double ep1, at2, en1, dt1
''    //double en2
''    //double en3//
''    //double en4,
''    double  hz1, hz2
''    //logical beg
''    double den
''    //logical neg
''    double hen, dtg, hev
''    int kgr
''    double dhz
''    //double ent[8000]  /* was [4][2000] */
''    double sbn
''    extern double fun_(double *)
''    double entry
''    int isn = 1
''    int iam = 0
''    double epz
''    extern double sqt_(double *, double *, double *)
''    extern int LoadAtmospheres(char filnam[])
''    double hen1, epg1, epg2, dtg2, fieg
''    //logical angl
''    //int mang,
''    int nang
''    //int ccc
''    double fiem, hmin, hmax
''    int nent,nhgt//,mhgt
''    double rsof, epzm
''    //char finam[10]
''    //double chisn
''    int nkhgt
''    double estep
''    double ANGLE = 0.0
''    double A1 = 0.0
''    double A2 = 0.0
''    bool START = false
''    char filnam[255] = sempty
''    char chr[2] = sempty
''    int ier = 0
''
''
''    FILE *stream
''
''L0:
'    numLayers = 0
'
'     '------------------progress bar initialization
'    With prjAtmRefMainfm
'      '------fancy progress bar settings---------
'      .progressfrm.Visible = True
'      .picProgBar.AutoRedraw = True
'      .picProgBar.BackColor = &H8000000B 'light grey
'      .picProgBar.DrawMode = 10
'
'      .picProgBar.FillStyle = 0
'      .picProgBar.ForeColor = &H400000 'dark blue
'      .picProgBar.Visible = True
'    End With
'    pbScaleWidth = 100
'    '-------------------------------------------------
'
'     'specify atmosphere type and the file containing the atmosphere profile
'      StatusMes = "Calculating and Storing multilayer atmospheric details"
'      Call StatusMessage(StatusMes, 1, 0)
'
'     If OptionLayer.Value = True Then
'        AtmType = 1
'        FNM = App.Path & "\stmod1.dat"
'     ElseIf OptionRead.Value = True Then
'        AtmType = 1
'        FNM = TextExternal.Text
'     ElseIf OptionSelby.Value = True Then
'        AtmType = 2
'        If prjAtmRefMainfm.opt1.Value = True Then
'           AtmNumber = 1
'        ElseIf prjAtmRefMainfm.opt2.Value = True Then
'           AtmNumber = 2
'        ElseIf prjAtmRefMainfm.opt3(0).Value = True Then
'           AtmNumber = 3
'        ElseIf prjAtmRefMainfm.opt4(1).Value = True Then
'           AtmNumber = 4
'        ElseIf prjAtmRefMainfm.opt5.Value = True Then
'           AtmNumber = 5
'        ElseIf prjAtmRefMainfm.opt6.Value = True Then
'           AtmNumber = 6
'        ElseIf prjAtmRefMainfm.opt7.Value = True Then
'           AtmNumber = 7
'        ElseIf prjAtmRefMainfm.opt8.Value = True Then
'           AtmNumber = 8
'        ElseIf prjAtmRefMainfm.opt9.Value = True Then
'           AtmNumber = 9
'        ElseIf prjAtmRefMainfm.opt10.Value = True Then
'           AtmNumber = 10
'           FNM = txtOther.Text
'           End If
'        End If
'
''     FNM = App.Path & "\stmod1.dat" 'stmod2.dat.txt"
'     ier = LoadAtmospheres(FNM, AtmType, AtmNumber, lpsrate, tst, pst, NNN, 4)
'     numLayers = NNN + 1
'     NumTemp = numLayers
'
'     If ier < 0 Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'        End If
'
'     '///////////////////////////////////////////////////////////////
''
''    //angl = TRUE_
''    printf(" SUMMER=1 (DEF), WINTER=2 --> ")
''    scanf("%lg", &entry)
''    isn = (int)entry
''    if (isn <> 1 and isn <>2) isn = 1
''
''    printf(" Choose Atmosphere: Menat-(1) Selby: Tropical-(2) Mid latitude-(3) subartic-(4) US Standard-(5) -->")
''    scanf("%lg", &entry)
''    iam = (int)entry
''    iam = iam - 1
''    if (iam < 0) iam = 0
''
''    //load the atmosphere files
''    if (isn == 1) {
''
''        switch (iam) {
''
''            Case 0:
''                strcpy(filnam, "Menat-EY-summer.txt")
''                break
''
''            Case 1:
''                strcpy(filnam, "Selby-tropical.txt")
''                break
''
''            Case 2:
''                strcpy(filnam, "Selby-midlatitude-summer.txt")
''                break
''
''            Case 3:
''                strcpy(filnam, "Selby-subartic-summer.txt")
''                break
''
''            Case 4:
''                strcpy(filnam, "Selby-US-standard.txt")
''                break
''
''default:
''                strcpy(filnam, "Menat-EY-summer.txt")
''                break
''        }
''    }
''
''    else if (isn == 2) {
''
''        switch (iam) {
''
''            Case 0:
''                strcpy(filnam, "Menat-EY-winter.txt" )
''                break
''
''            Case 1:
''                strcpy(filnam, "Selby-tropical.txt")
''                break
''
''            Case 2:
''                strcpy(filnam, "Selby-midlatitude-winter.txt")
''                break
''
''            Case 3:
''                strcpy(filnam, "Selby-subartic-winter.txt")
''                break
''
''            Case 4:
''                strcpy(filnam, "Selby-US-standard.txt")
''                break
''
''default:
''                strcpy(filnam, "Menat-EY-winter.txt" )
''                break
''        }
''    }
''
''    ier = LoadAtmospheres(filnam)
''    if (ier == -1) return -1
''
''    hsof = layerheights[numlayers - 1]
'     hsof = ELV(NNN)
''
''    //output file
'     fileout% = FreeFile
'     Open App.Path & "\test_M.dat" For Output As #fileout%
''    if ( !( stream = fopen( "MENAT.OUT", "w")) )
''    {
''        return -1
''    }
''
''    printf(" INPUT BEGINNING HEIGHT FOR CALCULATION (M)--> ")
''    scanf("%lg", &hz1)
''
'    hz1 = prjAtmRefMainfm.txtHeight
'    hz1 = hz1 / 1000# 'beginning observer height to kms
''
''    printf(" INPUT END HEIGHT FOR CALCULATION (M)--> ")
''    scanf("%lg", &hz2)
'    hz2 = hz2 / 1000# 'end observer height in kms
''
''    printf(" INPUT STEP HEIGHT FOR CALCULATION (M)--> ")
''    scanf("%lg", &dhz)
'    dhz = 100 'meters
'    dhz = dhz / 1000# 'stepsize in observer height in kms
'    d__1 = (hz2 - hz1) / dhz
'    nhgt = CInt(d__1) + 1
''
''/*        WRITE (*,'(A,F6.1)')' PRESENT HEIGHT FOR CALCULATION =',HZ*1000 */
''/*       WRITE (*,'(A\)')' WANT NEW HEIGHT ? (Y/N)' */
''/*        READ(*,'(A)')CH */
''/*        IF ((CH.EQ.'Y').OR.(CH.EQ.'y')) THEN */
''/* 2               WRITE(*,'(A\)')' INPUT NEW HEIGHT (M)-->' */
''/*                READ(*,*,ERR=2)HZ */
''/*                HZ=HZ/1.0D3 */
''/*                END IF */
''/*       DO 550 ISN=1,1 */
''/*       WRITE(*,5)ISN */
''/* 5       FORMAT('2',5X,'SUMMER=1 @ WINTER=2','   ISN=',I1//) */
''
''    //printf(" SUMMER=1 (DEF), WINTER=2 --> ")
''    //scanf("%d", &isn)
''    //if (isn <> 1 and isn <>2) isn = 1
''
''    /*
'    For k = 1 To numLayers
'        zz_1.hj(k - 1) = ELV(k - 1) ' //hs(k - 1)
'        zz_1.tj(k - 1) = TMP(k - 1) ' //ts(k - 1)
'        zz_1.pj(k - 1) = PRSR(k - 1) ' //ps(k - 1)
'    Next k
''    /*
''    if (isn == 1) {
''        goto L7
''    }
''    zz_1.hj(k - 1) = hw(k - 1)
''    zz_1.tj(k - 1) = tw(k - 1)
''    zz_1.pj(k - 1) = pw(k - 1)
''    */
''//L7:
''    /*
''
''    }
'    For k = 1 To numLayers
'        l = k + 1
'        If (k < numLayers) Then
'            zz_1.at(k - 1) = (zz_1.tj(l - 1) - zz_1.tj(k - 1)) / (zz_1.hj(l - _
'                1) - zz_1.hj(k - 1))
'            End If
'        If (k < numLayers) Then
'            If (zz_1.tj(l - 1) <> zz_1.tj(k - 1)) Then '//non-isothermic region
'
'                zz_1.ct(k - 1) = Log(zz_1.pj(l - 1) / zz_1.pj(k - 1)) / Log( _
'                    zz_1.tj(l - 1) / zz_1.tj(k - 1))
'
'            Else '//isothermic region  -- interpolate between the two layer's pressures
'
'                zz_1.ct(k - 1) = (zz_1.pj(l - 1) - zz_1.pj(k - 1)) / (zz_1.hj(l - _
'                1) - zz_1.hj(k - 1))
'
'                End If
'
''    printf("%i3, %f6, %f9, %f7, AT=%f7, CT=%f9\n", k, zz_1.hj(k - 1), zz_1.pj(k - 1), zz_1.tj(k - 1), zz_1.at(k - 1), zz_1.ct(k - 1))
'          End If
'     Next k
''    */
''
''    printf(" INPUT MINIMUM ANG ALT (DEG)--> ")
''    scanf("%lg", &epg1)
''
''
''    printf(" INPUT MAXIMUM ANG ALT (DEG)--> ")
''    scanf("%lg", &epg2)
''
''
''    printf(" INPUT STEP ANG ALT (DEG)--> ")
''    scanf("%lg", &estep)
''
''
'    n_size = 500
'    msize = 20 + Val(txtNumSuns.Text) * 32 * PPAM
'
'    If Trim$(txtXSize.Text) <> sempty Then
'       msize = Val(txtXSize.Text)
'       End If
'    If Trim$(txtYSize.Text) <> sempty Then
'       n_size = Val(txtYSize.Text)
'       End If
'
'   PPAM = prjAtmRefMainfm.txtPPAM
'   epg1 = CDbl(n_size * 0.5 / PPAM)
'   epg2 = -epg1
'   estep = CDbl(1 / PPAM)
'   nang = n_size
'
'Screen.MousePointer = vbHourglass
'
'For KWAV = KMIN To KMAX Step KSTEP   '<1
'
'   wl = 380# + CDbl(KWAV - 1) * 5#
'   wl = wl / 1000# 'convert to nm
'
'    nhgt = 1 'use so far only one height '<<<<<<<<<<<<<<
'    i__1 = nhgt
'    For nkhgt = 1 To i__1
''    {
''
'        hz = hz1 + (nkhgt - 1) * dhz
''        //mhgt = floor(hz*1000)
''
'        XP = 0#
''
''        //fprintf(stream, "%lg, %lg, %lg, %lg\n", XP, hz, A1, A2)
''
''        printf("Observer height (m) = %lg\n", hz * 1000.0)
''
'        d__1 = (epg2 - epg1) / estep
'        nang = CInt(Abs(d__1)) + 1
''
'        START = True
''
''
'        'for testing '<<<<<<<<<<<<<<<<<
''        nang = 1
''        epg1 = 0#
'
'        Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0) 'reset
'
'        i__2 = nang
'        For kgr = 1 To i__2
'            dtg2 = 0#
'            dh = 0.002
'            epg = epg1 - (kgr - 1) * estep
'            ALFA(KWAV, kgr) = epg
'
'            Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, CLng(100# * kgr / nang))
'            DoEvents
'
'            If (epg > 0# And epg < estep) Then
'               epg = 0#
'               End If
'
'            epz = epg / 60# * co
''    /*           write(*,*)' epz=',epz */
'            epzm = epz * q3
'            rsof = ra + hsof
'            If (epz < 0#) Then
'                If (hz - dh / 2# > 0#) Then
'                    hen1 = hz - dh / 2#
'                    dh = -0.002
'                ElseIf (hz - dh / 2# <= 0#) Then
'                    hen1 = hz
'                    End If
'            ElseIf (epz >= 0#) Then
'                hen1 = hz + dh / 2#
'                End If
'
'            If hen1 = 0 And zz_1.hj(0) = 0 Then 'hit rock bottom
'               jstop = kgr
'               Exit For
'               End If
'
'            If (START) Then
'                Write #fileout%, 0#, hz1 * 1000#
'                START = False
'                End If
'
'            cz = Cos(epz) * (ra + hz) * (fun_(hen1, zz_1, Dist, NumLayers) + 1#)
'            bz = pi2 - epz '//zenith angle
'            bn = bz
''    /*           write(*,*)' bn=',bn */
'            ep1 = q3 * epz
'
'            T = 0#
'            s1 = 0#
'            at2 = 0#
'            nz = CLng((hsof + Abs(dh) - hz) / Abs(dh) + 1.1)  '//number of steps
'
'            h__ = hz - dh '//start at this height
'
'            nent = 0
'
'            i__3 = nz
'            For n = 1 To i__3 'trace over all the layers
'
'                b = bn 'initial zenith angle
'                e = pi2 - b 'initial view angle
'
'                e1 = q3 * e 'view angle in mrad
'
'                d__1 = e / co
'
'                If epg <> 0# And d__1 < 0.00005 Then
'                    e = 0#
'                    dh = 0.002
'                    dtg2 = dtg
'                    End If
'
'                If (dh < 0#) Then
'                    nz = nz + 1
'                    End If
'                h__ = h__ + dh
'
'                If (h__ < 0#) Then
'                    h__ = 0#
'                    End If
'
'                rt = ra + h__
''        /* L25: */
'                ru = rt + dh
'                hen = h__ + dh / 2#
''        /* >>          write(*,*)' hen=',hen */
'
'                If (n = 1) Then
'                    hmin = hen
'                    hmax = hen
'                Else
'                    If (hen < hmin) Then
'                        hmin = hen
'                        End If
'                    If (hen > hmax) Then
'                        hmax = hen
'                        End If
'                    End If
'
'                hev = hen + dh
''
'                el = sqt_(e, dh, rt) 'path length of light ray from last height to current height
''
'                s = el * Cos(e)  'begin law of sin calculation of the subtended cylindrical angle at the Earth's center
''
'                A2 = DASIN(s / ru) 'this is the subtended angle, dtheta for this last step
''        /* L30: */
'                den = A2 * ra  'this is length along the circumference of the earth for the last subtended angle increment
''
'                XP = XP + den
''                'diagnostics
''                /*
''                if (XP >= 249.111) {
''                    ccc = 1
''                }
''                */
'                at2 = at2 + A2  'this is the total cylindrical angle, theta
'                s2 = at2 * ra 'this is the total length along the circumference of the earth
'                a3 = q6 * A2 'ditto in mrad
'                en = fun_(hen, zz_1, Dist, NumLayers) 'variable portion of the index of refraction at this height
'                en1 = en * 10000000# 'normalized to 1.0, where n = 1 + en1
''        /* L35: */
'                g = b - A2  'incident angle
''        /* >>          WRITE(*,*)' G(DEG)=',G/CO */
'                g1 = q3 * g
'                e2 = q3 * (pi2 - g)
'                em = fun_(hev, zz_1, Dist, NumLayers) + 1#
'
'                sbn = (en + 1#) * Sin(g) / em 'Snell's law, where new angle asin(sbn) which is g + incremental refraction
''        /* >>          WRITE(*,*)' SBN=',SBN */
'                If (sbn > 1#) Then
'                    sbn = 1#
'                    End If
'
'                bn = DASIN(sbn)
'                If (g > pi2) Then
'                    bn = pie - bn
'                    d__1 = g - pi2
'                    dh = -el * Abs(d__1)
'                    End If
'
'                d__ = bn - g
'                d6 = q6 * d__
'                T = T + d__
'                dt1 = q3 * T
'                dtg = dt1 * ramg
'                fr = 1#
'                If (h__ >= 1.5) Then
'                    fr = 1.0001
'                    End If
'                dh = fr * dh
'                dg = dh * q3
'
'                ANGLE = XP / ra 'angle XP subtends in radians
'                A1 = XP * Cos(ANGLE) + ru * Sin(ANGLE)
'                A2 = -XP * Sin(ANGLE) + ru * Cos(ANGLE) - ra
'
'
''                'fprintf(stream, "%lg, %lg, %lg, %lg\n", XP, ru - ra, A1, A2)
''                fprintf(stream, "%lg, %lg\n", A1, A2)
''                Write #fileout%, XP, ru 'A1, A2
'
'                If ru < ra Then 'collided with the surface
'                   Write #fileout%, XP * 1000#, -1000
'                   jstop = kgr
'                Else
'                   'limit the recording to every other step
'                   If n Mod Val(prjAtmRefMainfm.txtHeightStepSize.Text) = 0 Then
'                      Write #fileout%, XP * 1000#, (ru - ra) * 1000#
'                      End If
'                   jstop = -1
'                   End If
'
'                If (rt >= rsof) Then
'                    Exit For
'                    End If
'
'                fiem = epzm - 4.665 - dt1
'                fieg = fiem * ramg
'
'                If jstop <> -1 Then Exit For
'
'            Next n
'
'            If jstop <> -1 Then Exit For
''
''            printf("View Angle (deg.) = %lg, Accumlated refraction (mrad) = %lg\n", epg, dt1)
''            If epg = 0 Then
''               prjAtmRefMainfm.lblRef = "View Angle (deg.) = " & epg & vbCrLf & "Accumlated refraction (mrad) = " & dt1
''               End If
'
'            ALFT(KWAV, kgr) = ALFA(KWAV, kgr) - dt1 * 0.001 * 180# * 60 / PI 'true depression angle in minutes of degree
'
''            If ALFA(KWAV, kgr) = 0 Then
''               ccc = 1
''               End If
'
'            START = True
'            XP = 0#
'
'        Next kgr
'        If jstop <> -1 Then Exit For
'    Next nkhgt
'   If jstop <> -1 Then Exit For
'
'Next KWAV
'
'Close #fileout%
'
'Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
'prjAtmRefMainfm.progressfrm.Visible = False
'
'StatusMes = "Apparent Altitude of the Horizon (arcminutes) = " & Str(ALFA(KWAV, jstop - 1))& vbcrlf & "True Altitude of the Horizon (arcminutes) = " & Str((-DACOS(ra / (ra + hz)) / CONV))
'Call StatusMessage(StatusMes, 1, 0)
'prjAtmRefMainfm.lblHorizon.Caption = StatusMes
'prjAtmRefMainfm.lblHorizon.Refresh
'DoEvents
''     ier = MsgBox(StatusMes, vbInformation + vbOKOnly, "Horizon")
'StatusMes = "Ray tracing calculation complete..."
'Call StatusMessage(StatusMes, 1, 0)
'
'CalcComplete = True
'
''increase resolution of atmospheric models and load to charts and convert elevations to meters
'Dim hgt As Double, Pr As Double, Te As Double
'NumTemp = 0
'For j = 1 To NNN
'   For hgt = zz_1.hj(j - 1) To zz_1.hj(j) - Val(prjAtmRefMainfm.txtHeightStepSize.Text) * 0.001 Step Val(prjAtmRefMainfm.txtHeightStepSize.Text) * 0.001
'      Call layers_int(hgt, zz_1, Dist, NumLayers, Pr, Te)
'      ELV(NumTemp) = hgt * 1000#
'      TMP(NumTemp) = Te
'      PRSR(NumTemp) = Pr
'      NumTemp = NumTemp + 1
'   Next hgt
'Next j
'ELV(NumTemp) = zz_1.hj(NNN - 1) * 1000#
'TMP(NumTemp) = zz_1.tj(NNN - 1)
'PRSR(NumTemp) = zz_1.pj(NNN - 1)
'NNN = NumTemp
''now load up temperature and pressure charts
'
' ReDim TransferCurve(1 To NNN, 1 To 3) As Variant
'
' For j = 1 To NNN
'    TransferCurve(j, 1) = " " & CStr(ELV(j - 1) * 0.001)
''         TransferCurve(J, 2) = ELV(J - 1) * 0.001
'    TransferCurve(j, 2) = TMP(j - 1)
' Next j
'
' With MSChartTemp
'   .chartType = VtChChartType2dLine
'   .RandomFill = False
''        .RowCount = 2
''        .ColumnCount = IncN - 1
''        .RowLabel = "Height (km)"
''        .ColumnLabel = "Temperature (Kelvin)"
'   .ChartData = TransferCurve
' End With
'
' For j = 1 To NNN
'    TransferCurve(j, 2) = PRSR(j - 1)
' Next j
'
' With MSChartPress
'   .chartType = VtChChartType2dLine
'   .RandomFill = False
''        .RowCount = 2
''        .ColumnCount = IncN - 1
''        .RowLabel = "Height (km)"
''        .ColumnLabel = "Pressure (Kelvin)"
'   .ChartData = TransferCurve
' End With
'
'Screen.MousePointer = vbDefault
'
'StatusMes = "Writing transfer curve."
'Call StatusMessage(StatusMes, 1, 0)
'filnum% = FreeFile
'Open App.Path & "\tc_M.dat" For Output As #filnum%
''      WRITE (20,*) N
'NumTc = 0
'Print #filnum%, n_size
'For j = 1 To jstop - 1
''        WRITE(20,1) ALFA(KMIN,J),ALFT(KMIN,J)
'    Print #filnum%, ALFA(KMIN, j), ALFT(KMIN, j)
'    If ALFA(KMIN, j) = 0 Then 'display the refraction value for the zero view angle ray
'       prjAtmRefMainfm.lblRef.Caption = "Atms. refraction (deg.) = " & Abs(ALFT(KMIN, j)) / 60# & vbCrLf & "Atms. refraction (mrad) = " & Abs(ALFT(KMIN, j)) * 1000# * cd / 60#
'       prjAtmRefMainfm.lblRef.Refresh
'       DoEvents
'       End If
''store all view angles that contribute to sun's orb
'    NumTc = NumTc + 1
'    For KA = 1 To NumSuns
'       y = ALFT(KMIN, j) - ALT(KA)
'       If Abs(y) <= ROBJ Then
'          'only accept rays that pass over the horizon (ALFT(KMIN, J) <> -1000) and are within the solar disk
'          SunAngles(KA - 1, NumSunAlt(KA - 1)) = j
'          NumSunAlt(KA - 1) = NumSunAlt(KA - 1) + 1
'          End If
'    Next KA
'Next j
'Close #filnum%
'
'
''now load up transfercurve array for plotting
'ReDim TransferCurve(1 To NumTc, 1 To 2) As Variant
'
'For j = 1 To NumTc
' TransferCurve(j, 1) = " " & CStr(ALFA(KMIN, j))
' TransferCurve(j, 2) = ALFT(KMIN, j)
''         TransferCurve(J, 1) = " " & CStr(ALFT(KMIN, J))
''         TransferCurve(J, 2) = ALFA(KMIN, J)
'Next j
'
'With MSCharttc
'.chartType = VtChChartType2dLine
'.RandomFill = False
''        .RowCount = 2
''        .ColumnCount = IncN
''        .RowLabel = "True angle (min)"
''        .ColumnLabel = "View angle (min)"
'.ChartData = TransferCurve
'End With
'
'
' StatusMes = "Drawing the rays on the sky simulation, please wait...."
' Call StatusMessage(StatusMes, 1, 0)
' 'load angle combo boxes
''    AtmRefPicSunfm.WindowState = vbMinimized
''    BrutonAtmReffm.WindowState = vbMaximized
' 'set size of picref by size of earth
' Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
' prjAtmRefMainfm.cmbSun.Clear
' prjAtmRefMainfm.cmbAlt.Clear
' For i = 1 To NumSuns
'    If NumSunAlt(i - 1) > 0 Then prjAtmRefMainfm.cmbSun.AddItem i
' Next i
'
' prjAtmRefMainfm.TabRef.Tab = 4
' DoEvents
'
'cmbSun.ListIndex = 0
'
'   cmdCalc.Enabled = True
'   cmdRefWilson.Enabled = True
'   cmdMenat.Enabled = True
''
''    printf("Do you want to a new calculation? (y/n) -->")
''    scanf("%s", chr)
''    if (strstr(chr, "y")) goto L0
''
''
''    return 0
''} /* MAIN__ */
''
''
'   On Error GoTo 0
'   Exit Sub
'
'cmdMenat_Click_Error:
'    Close
'    Screen.MousePointer = vbDefault
'
'    StatusMes = sempty
'    Call StatusMessage(StatusMes, 1, 0)
'    Call UpdateStatus(prjAtmRefMainfm, picProgBar, 1, 0)
'    prjAtmRefMainfm.progressfrm.Visible = False
'
'   cmdCalc.Enabled = True
'   cmdRefWilson.Enabled = True
'   cmdMenat.Enabled = True
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMenat_Click of Form prjAtmRefMainfm"
'
'End Sub

Private Sub cmpTLoop_Click()
   Select Case MsgBox("Did you Set the Latitude?", vbYesNoCancel Or vbInformation Or vbDefaultButton1, "Temperature Loop")
   
    Case vbYes
      'proceed to calculation
    Case vbNo
      Exit Sub
    Case vbCancel
      Exit Sub
   End Select
   
   TempLoop = True
   OptionSelby.Value = False
   Call cmdVDW_Click
'   cmpTLoop.Enabled = False
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If TabRef.Tab = 7 Then
'        KeyPressed = KeyAscii
'        End If
'End Sub

Private Sub Form_Load()
   
   Mult = 1#
   RefZoom.LastZoom = Mult
   RefZoom.Zoom = Mult
   Xorigin = 0
   Yorigin = 0
   
   twipsx = 1 'Screen.TwipsPerPixelX
   twipsy = 1 'Screen.TwipsPerPixelY
   
   RefZoom.LastZoom = 1#
   RefZoom.Zoom = 1#
   
'   pixwi = 1162
'   pixhi = 3046
   prjAtmRefMainfm.WindowState = vbMaximized
   
   With TabRef
     .Left = 10
     .Top = 10
     .Width = prjAtmRefMainfm.ScaleWidth - 20
     .height = prjAtmRefMainfm.ScaleHeight - 20
   End With
   
      
   With paramfrm
'    .Visible = True
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With


   
'    AlphaSun.SetRedraw = False
'    With HScrollPan
'        .Min = -AlphaSun.Width
'        .Max = -.Min
'        .Value = 0
'    End With
'    With VScrollPan
'        .Min = -AlphaSun.Height
'        .Max = -.Min
'        .Value = 0
'    End With
'
'    AlphaSun.WantPrePostEvents = True
'    AlphaSun.SetRedraw = True
'    AlphaSun.Refresh
   
   With Tempfrm
    .Visible = False
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With
   
   With Pressfrm
    .Visible = False
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With
   
    
   With Sunsfrm
    .Visible = False
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With
   
   With Rayfrm
    .Visible = False
    .Width = TabRef.Width - 20
'    .Height = TabRef.Height - 20
    .height = TabRef.height - .Top - cmdLarger.Top - cmdLarger.height - 20

'    picture1.Width = .Width - prjAtmRefMainfm.VScroll1.Width - 10
'    picture1.Height = .Height - HScroll1.Height - cmdLarger.Top - cmdLarger.Height - 200
'    picture1.Left = 100
'    Picture2.Left = picture1.Left
'    Picture2.Width = Picture2.Width
'    picture1.Top = cmdLarger.Top + cmdLarger.Height + 100
'    Picture2.Top = picture1.Top
'    VScroll1.Left = picture1.Left + picture1.Width + 10
'    HScroll1.Left = picture1.Left
'    HScroll1.Width = picture1.Width
'    HScroll1.Top = picture1.Top + picture1.Height + 10
'    pixwi = picture1.Width
'    pixhi = picture1.Height
   End With
   
   With Tcfrm
    .Visible = False
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With
   
   With Terrfrm
    .Visible = False
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With
   
   With paramfrm
    .Visible = True
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With
   
   With frmVDW
    .Visible = False
    .Left = 500
    .Top = 500
    .Width = TabRef.Width - 200
    .height = TabRef.height - 200
   End With
   
    Dim label As MSChart20Lib.label

    With MSChartTemp
        .chartType = VtChChartType2dLine
        .RandomFill = False
        With .Title
            .Text = "Temperature Profile"
            .VtFont.Size = .VtFont.Size * 1.5
        End With
        With .Plot
            .UniformAxis = True

            'Set the Wall to white:
            With .Wall.Brush
                .FillColor.Set 255, 255, 255
                .Style = VtBrushStyleSolid
            End With

            'Set axis gridline colors, styles:
            With .Axis(VtChAxisIdX).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With
            With .Axis(VtChAxisIdY).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With

            'Put the vertical (X-axis) lines of the grid ON the values rather
            'than BORDERING the values (as we want for bar charts):
            .Axis(VtChAxisIdX).CategoryScale.LabelTick = True

            'Format value (Y-axis) labels as currency:
            For Each label In .Axis(VtChAxisIdY).Labels
                label.Format = "###0.00"
            Next
'            For Each label In .Axis(VtChAxisIdY2).Labels
'                label.Format = "##0,##0.00"
'            Next

            'Though the documentation says Width values are in Points they are
            'actually in Twips.
            '
            'Set Legend text and plot the values as 1-pixel lines of specific
            'colors:
            With .SeriesCollection(1)
                .LegendText = "Temperature"
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 255, 0, 0
                End With
            End With
'            With .SeriesCollection(2)
'                .LegendText = "Second"
'                With .Pen
'                    .Width = ScaleX(1, vbPixels, vbTwips)
'                    .VtColor.Set 0, 192, 0
'                End With
'            End With
        End With

        'Format the legend area:
        With .Legend.Backdrop
            With .Fill
                With .Brush
                    .FillColor.Set 255, 255, 255
                    .Style = VtBrushStyleSolid
                End With
                .Style = VtFillStyleBrush
            End With
            With .Frame
                .Style = VtFrameStyleSingleLine
            End With
        End With
        .ShowLegend = True
    End With

    With MSChartPress
        .RandomFill = False
        .chartType = VtChChartType2dLine
        With .Title
            .Text = "Pressure Profile"
            .VtFont.Size = .VtFont.Size * 1.5
        End With
        With .Plot
            .UniformAxis = True

            'Set the Wall to white:
            With .Wall.Brush
                .FillColor.Set 255, 255, 255
                .Style = VtBrushStyleSolid
            End With

            'Set axis gridline colors, styles:
            With .Axis(VtChAxisIdX).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With
            With .Axis(VtChAxisIdY).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With

            'Put the vertical (X-axis) lines of the grid ON the values rather
            'than BORDERING the values (as we want for bar charts):
            .Axis(VtChAxisIdX).CategoryScale.LabelTick = True

            'Format value (Y-axis) labels as currency:
            For Each label In .Axis(VtChAxisIdY).Labels
                label.Format = "######0.0#"
            Next
'            For Each label In .Axis(VtChAxisIdY2).Labels
'                label.Format = "######0,##0.00"
'            Next

            'Though the documentation says Width values are in Points they are
            'actually in Twips.
            '
            'Set Legend text and plot the values as 1-pixel lines of specific
            'colors:
        
            With .SeriesCollection(1)
                .LegendText = "Pressure"
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 255, 0, 0
                End With
            End With
'            With .SeriesCollection(2)
'                .LegendText = "Second"
'                With .Pen
'                    .Width = ScaleX(1, vbPixels, vbTwips)
'                    .VtColor.Set 0, 192, 0
'                End With
'            End With
        End With

        'Format the legend area:
        With .Legend.Backdrop
            With .Fill
                With .Brush
                    .FillColor.Set 255, 255, 255
                    .Style = VtBrushStyleSolid
                End With
                .Style = VtFillStyleBrush
            End With
            With .Frame
                .Style = VtFrameStyleSingleLine
            End With
        End With
        .ShowLegend = True
    End With

    With MSCharttc
        .chartType = VtChChartType2dLine
        .RandomFill = False
        With .Title
            .Text = "Transfer Curve"
            .VtFont.Size = .VtFont.Size * 1.5
        End With
        With .Plot
            .UniformAxis = True

            'Set the Wall to white:
            With .Wall.Brush
                .FillColor.Set 255, 255, 255
                .Style = VtBrushStyleSolid
            End With

            'Set axis gridline colors, styles:
            With .Axis(VtChAxisIdX).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With
            With .Axis(VtChAxisIdY).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With

            'Put the vertical (X-axis) lines of the grid ON the values rather
            'than BORDERING the values (as we want for bar charts):
            .Axis(VtChAxisIdX).CategoryScale.LabelTick = True

            'Format value (Y-axis) labels as currency:
            For Each label In .Axis(VtChAxisIdY).Labels
                label.Format = "##0.00"
            Next
'            For Each label In .Axis(VtChAxisIdY2).Labels
'                label.Format = "##0,##0.00"
'            Next

            'Though the documentation says Width values are in Points they are
            'actually in Twips.
            '
            'Set Legend text and plot the values as 1-pixel lines of specific
            'colors:
        
            With .SeriesCollection(1)
                .LegendText = "Transfer Curve"
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 255, 0, 0
                End With
            End With
'            With .SeriesCollection(2)
'                .LegendText = "Second"
'                With .Pen
'                    .Width = ScaleX(1, vbPixels, vbTwips)
'                    .VtColor.Set 0, 192, 0
'                End With
            End With
'        End With

        'Format the legend area:
        With .Legend.Backdrop
            With .Fill
                With .Brush
                    .FillColor.Set 255, 255, 255
                    .Style = VtBrushStyleSolid
                End With
                .Style = VtFillStyleBrush
            End With
            With .Frame
                .Style = VtFrameStyleSingleLine
            End With
        End With
        .ShowLegend = True
    End With
    
    With MSChartTR
        
        .chartType = VtChChartType2dLine
        .RandomFill = False
        With .Title
            .Text = "Terrestrial Refraction"
            .VtFont.Size = .VtFont.Size * 1.5
        End With
        With .Plot
            .UniformAxis = True

            'Set the Wall to white:
            With .Wall.Brush
                .FillColor.Set 255, 255, 255
                .Style = VtBrushStyleSolid
            End With

            'Set axis gridline colors, styles:
            With .Axis(VtChAxisIdX).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With
            With .Axis(VtChAxisIdY).AxisGrid.MajorPen
                .VtColor.Set 224, 224, 255
                .Style = VtPenStyleDotted
            End With

            'Put the vertical (X-axis) lines of the grid ON the values rather
            'than BORDERING the values (as we want for bar charts):
            .Axis(VtChAxisIdX).CategoryScale.LabelTick = True
            
            For Each label In .Axis(VtChAxisIdX).Labels
                label.Format = "##0.0#"
            Next
            'Format value (Y-axis) labels as currency:
            For Each label In .Axis(VtChAxisIdY).Labels
                label.Format = "##0.0####"
            Next
'            For Each label In .Axis(VtChAxisIdY2).Labels
'                label.Format = "##0,##0.00"
'            Next

            'Though the documentation says Width values are in Points they are
            'actually in Twips.
            '
            'Set Legend text and plot the values as 1-pixel lines of specific
            'colors:
            With .SeriesCollection(1)
                .LegendText = "Terrestrial Refraction"
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 255, 0, 0
                End With
            End With
'            With .SeriesCollection(2)
'                .LegendText = "Second"
'                With .Pen
'                    .Width = ScaleX(1, vbPixels, vbTwips)
'                    .VtColor.Set 0, 192, 0
'                End With
'            End With
        End With

        'Format the legend area:
        With .Legend.Backdrop
            With .Fill
                With .Brush
                    .FillColor.Set 255, 255, 255
                    .Style = VtBrushStyleSolid
                End With
                .Style = VtFillStyleBrush
            End With
            With .Frame
                .Style = VtFrameStyleSingleLine
            End With
        End With
        .ShowLegend = True
    End With
    
End Sub

Private Sub Form_Resize()

'  If prjAtmRefMainfm.WindowState <> vbMinimized Then
'     TabRef.Move 0, 0, ScaleWidth, ScaleHeight
'     paramfrm.Move paramfrm.Left, paramfrm.Top, ScaleWidth, ScaleHeight
'     Tempfrm.Move Tempfrm.Left, Tempfrm.Top, ScaleWidth, ScaleHeight
'     Pressfrm.Move Pressfrm.Left, Pressfrm.Top, ScaleWidth, ScaleHeight
'     Rayfrm.Move Rayfrm.Left, Rayfrm.Top, ScaleWidth, ScaleHeight
'     Sunsfrm.Move Sunsfrm.Left, Sunsfrm.Top, ScaleWidth, ScaleHeight
'     cmdShowSuns.Move cmdShowSuns.Left, cmdShowSuns.Top, ScaleWidth, ScaleHeight
'     Tempfrm.Move Tempfrm.Left, Tempfrm.Top, ScaleWidth, ScaleHeight
'     Tcfrm.Move Tcfrm.Left, Tcfrm.Top, ScaleWidth, ScaleHeight
'     picture1.Move picture1.Left, picture1.Top, ScaleWidth, ScaleHeight
'     Picture2.Move Picture2.Left, Picture2.Top, ScaleWidth, ScaleHeight
'     HScroll1.Move Picture2.Left, Picture2.Top + Picture2.Height, ScaleWidth, ScaleHeight
'     VScroll1.Move Picture2.Left + Picture2.Width, Picture2.Top, ScaleWidth, ScaleHeight
'     MSChartTemp.Move MSChartTemp.Left, MSChartTemp.Top, ScaleWidth, ScaleHeight
'     MSChartPress.Move MSChartPress.Left, MSChartPress.Top, ScaleWidth, ScaleHeight
'     MSCharttc.Move MSCharttc.Left, MSCharttc.Top, ScaleWidth, ScaleHeight
'     End If

   On Error GoTo Form_Resize_Error

  If prjAtmRefMainfm.WindowState = vbMaximized Then

     TabRef.Top = 10
     TabRef.Left = 10
     TabRef.Width = prjAtmRefMainfm.Width - 20
     TabRef.height = prjAtmRefMainfm.height - 20

      With paramfrm
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
      End With

      With Tempfrm
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
       MSChartTemp.Width = .Width - 2 * MSChartTemp.Left
       MSChartTemp.height = .height - 2 * MSChartTemp.Top
      End With

      With Pressfrm
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
       MSChartPress.Width = .Width - 2 * MSChartPress.Left
       MSChartPress.height = .height - 2 * MSChartPress.Top
      End With

      With Rayfrm
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
      End With

      'format picture boxes for ray tracing
      With prjAtmRefMainfm
        .picture1.Left = 10 * Screen.TwipsPerPixelX
        .picture1.Width = .Width - .VScroll1.Width - 60 * Screen.TwipsPerPixelX
        .picture1.height = .height - .cmdLarger.Top - .cmdLarger.height - .HScroll1.height - .Rayfrm.Top - 70 * Screen.TwipsPerPixelY
        .Picture2.Top = 0
        .Picture2.Left = 0
        .Picture2.Width = .picture1.Width / Screen.TwipsPerPixelX - 10
        .Picture2.height = .picture1.height / Screen.TwipsPerPixelY - 10
        .VScroll1.Left = .picture1.Left + .picture1.Width + 10
        .VScroll1.Top = .picture1.Top
        .VScroll1.height = .picture1.height
        .HScroll1.Left = .picture1.Left
        .HScroll1.Width = .picture1.Width
        .HScroll1.Top = .picture1.Top + .picture1.height + 10

'        .picRef.Left = .Picture2.Left
'        .picRef.Top = .Picture2.Top
'        .picRef.Width = .Picture2.Width
'        .picRef.Height = .Picture2.Height
     End With


      With Sunsfrm
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
       prjAtmRefMainfm.cmdShowSuns.Left = .Left + .Width / 2 - prjAtmRefMainfm.cmdShowSuns.Width / 2
       prjAtmRefMainfm.cmdShowSuns.Top = .Top + .height / 2 - prjAtmRefMainfm.cmdShowSuns.height / 2
      End With


      With Terrfrm
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
'       frmTR.Width = .Width - 2 * frmTR.Left
       
       MSChartTR.Width = .Width - 2 * MSChartTR.Left
       MSChartTR.height = 0.75 * .height
       MSChartTR.Top = 170
       
       frmTR.Top = MSChartTR.Top + MSChartTR.height
       frmTR.Left = MSChartTR.Left
       frmTR.Width = MSChartTR.Width
'       frmTR.Left = .Width * 0.5 - 0.5 * frmTR.Width
'       frmTR.Width = MSChartTR.Width
       
       frmFit2.Left = Atmfrm.Left + Atmfrm.Width + 700
       frmFit2.Top = frmTR.Top + 200
       frmFit2.height = Atmfrm.height
      End With


      With Tcfrm
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
       MSCharttc.Width = .Width - 2 * MSCharttc.Left
       MSCharttc.height = .height - 2 * MSCharttc.Top
      End With
      
      With frmVDW
       .Left = 10 * Screen.TwipsPerPixelX
       .Top = 25 * Screen.TwipsPerPixelY
       .Width = TabRef.Width - 35 * Screen.TwipsPerPixelX
       .height = TabRef.height - 50 * Screen.TwipsPerPixelY - MDIAtmRef.StatusBar.height
       picVDW.Top = 20
       picVDW.Left = 20
       picVDW.Width = .Width - 40
       picVDW.height = .height - 40
      End With
      End If
      
      prjAtmRefMainfm.Refresh
      DoEvents

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:
    Resume Next
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form prjAtmRefMainfm"

End Sub
Private Sub cmbAlt_Change()
     PlotMode = 1
     If CalcComplete Then Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
End Sub

Private Sub cmbAlt_Click()
      PlotMode = 2
      If CalcComplete Then Call PlotRayTracing(prjAtmRefMainfm, prjAtmRefMainfm.Picture2, prjAtmRefMainfm.cmbSun, prjAtmRefMainfm.cmbAlt)
End Sub

Private Sub cmbSun_Change()
   If Not UsingHSatmosphere Then
      n = n_size
      End If
   cmbAlt.Clear
   'now load cmbalt with relevant view angles
   Dim NA As Long
   If cmbSun.ListIndex = -1 Then Exit Sub
   NA = Val(cmbSun.List(cmbSun.ListIndex))
   For i = 1 To NumSunAlt(NA - 1)
      'ALFA(K, J) = (CDbl(N / 2 - SunAngles(NA, I)) / PPAM)
       cmbAlt.AddItem (CDbl(n / 2 - SunAngles(NA - 1, i - 1) + 1) / PPAM)
   Next i
   cmbAlt.AddItem "All"
   cmbAlt.ListIndex = 0
End Sub

Private Sub cmbSun_Click()
   On Error GoTo cmbSun_Click_Error

   If Not UsingHSatmosphere Then
      n = n_size
      End If
   cmbAlt.Clear
   'now load cmbalt with relevant view angles
   Dim NA As Long
   If cmbSun.ListIndex = -1 Then Exit Sub
   NA = Val(cmbSun.List(cmbSun.ListIndex))
   For i = 1 To NumSunAlt(NA - 1)
      'ALFA(K, J) = (CDbl(N / 2 - SunAngles(NA, I)) / PPAM)
      cmbAlt.AddItem (CDbl(n / 2 - SunAngles(NA - 1, i - 1) + 1) / PPAM)
   Next i
   cmbAlt.AddItem "All"
   cmbAlt.ListIndex = 0

   On Error GoTo 0
   Exit Sub

cmbSun_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmbSun_Click of Form prjAtmRefMainfm"
End Sub
'///////////////////////////////////////////////////////

Private Sub HScrollPan_Change()
   AlphaSun.Refresh
End Sub

Private Sub HScrollPan_Scroll()
   AlphaSun.Refresh
End Sub

Private Sub HScroll1_Change()
   prjAtmRefMainfm.Picture2.Left = prjAtmRefMainfm.Picture2.Left + HScroll1.Value
End Sub

Private Sub opt1_Click()
   If opt1.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt10_Click()
   If opt10.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt2_Click()
   If opt2.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt3_Click()
   If opt3.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt4_Click()
   If opt4.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt5_Click()
   If opt5.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt6_Click()
   If opt6.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt7_Click()
   If opt7.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt8_Click()
   If opt8.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub opt9_Click()
   If opt9.Value = True Then
      OptionSelby.Value = True
      End If
End Sub

Private Sub OptionSelby_Click()
   If OptionSelby.Value = True Then
      chkDucting.Value = vbUnchecked 'can't add ducting if using layer atmospheres
      End If
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
'   a = KeyCode
   'up arrow = 38
   'down arrow = 40
   'left arrow = 37
   'right arrow = 39
   'PgUp = 33
   'PgDown = 34
   If KeyCode = Asc("Z") Or KeyCode = Asc("z") Then
      'zoom out
      Call PictureBoxZoom(prjAtmRefMainfm.picture1, 0, -120, 0, 0, 0)
      End If
      
   If KeyCode = Asc("X") Or KeyCode = Asc("x") Then
      'zoom in
      Call PictureBoxZoom(prjAtmRefMainfm.picture1, 0, 120, 0, 0, 0)
      
   ElseIf KeyCode = vbKeyLeft Then 'left arrow
      H1 = -10
      GoSub ScrollHoriz
   ElseIf KeyCode = vbKeyUp Then 'up arrow
      H2 = -10
      GoSub ScrollVert
   ElseIf KeyCode = vbKeyRight Then 'right arrow
      H1 = 10
      GoSub ScrollHoriz
   ElseIf KeyCode = vbKeyDown Then 'down arrow
      H2 = 10
      GoSub ScrollVert
   ElseIf KeyCode = vbKeyPageUp Then  'PgUp
'      DigiPage = DigiPage + 1
      If prjAtmRefMainfm.VScroll1.Visible Then prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.max
   ElseIf KeyCode = vbKeyPageDown Then  'PgDown
'      DigiPage = DigiPage - 1
      If prjAtmRefMainfm.VScroll1.Visible Then prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.Min
   ElseIf KeyCode = vbKeyHome Then 'Home
      If prjAtmRefMainfm.HScroll1.Visible Then prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.Min
   ElseIf KeyCode = vbKeyEnd Then 'End
      If prjAtmRefMainfm.HScroll1.Visible Then prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.max
      End If
       
Exit Sub

ScrollHoriz:
    If prjAtmRefMainfm.HScroll1.Value + H1 < prjAtmRefMainfm.HScroll1.Min Or prjAtmRefMainfm.HScroll1.Value + H1 > prjAtmRefMainfm.HScroll1.max Then
          'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
          If prjAtmRefMainfm.picture1.Width > prjAtmRefMainfm.HScroll1.Width Then

             If prjAtmRefMainfm.HScroll1.Value + H1 < prjAtmRefMainfm.HScroll1.Min Then
                prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.Min
             ElseIf prjAtmRefMainfm.HScroll1.Value + H1 > prjAtmRefMainfm.HScroll1.max Then
                prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.max
                End If
             End If
    Else
       prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.Value + H1
       End If
Return

ScrollVert:
    If prjAtmRefMainfm.VScroll1.Value + H2 < 0 Or prjAtmRefMainfm.VScroll1.Value + H2 > prjAtmRefMainfm.VScroll1.max Then
         'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
          If prjAtmRefMainfm.picture1.height > prjAtmRefMainfm.VScroll1.height Then

             If prjAtmRefMainfm.VScroll1.Value + H2 < prjAtmRefMainfm.VScroll1.Min Then
                prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.Min
             ElseIf prjAtmRefMainfm.VScroll1.Value + H2 > prjAtmRefMainfm.VScroll1.max Then
                prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.max
                End If

             End If
    Else
       prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.Value + H2
       End If
                   
Return
      
End Sub



Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim ier As Integer
   ier = MouseDown(Button, Shift, x, y)
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim ier As Integer
   ier = MouseMove(Button, Shift, x, y)
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim ier As Integer
   ier = MouseUp(Button, Shift, x, y)
End Sub

Private Sub picVDW_KeyPress(KeyAscii As Integer)
   If TabRef.Tab = 7 Then
      KeyPressed = KeyAscii
      End If
End Sub

'Private Sub picVDW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If TabRef.Tab = 7 Then
'      KeyPressed = 1
'      End If
'End Sub

'Private Sub picVDW_Click()
'   If TabRef.Tab = 7 Then
'        KeyPressed = 1
'        End If
'End Sub

Private Sub picVDW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If TabRef.Tab = 7 Then
        KeyPressed = 1
        End If
End Sub

Private Sub TabRef_Click(PreviousTab As Integer)

   Select Case TabRef.Tab
      Case 0
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = True
         frmVDW.Visible = False
      Case 1
         Tempfrm.Visible = True
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         frmVDW.Visible = False
         paramfrm.Visible = False
      Case 2
         Tempfrm.Visible = False
         Pressfrm.Visible = True
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 3
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = True
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 4
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = True
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
        
      Case 5
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = True
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 6
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = True
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 7
         frmVDW.Visible = True
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
   End Select
End Sub

Private Sub TabRef_DblClick()
   Select Case TabRef.Tab
      Case 0
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = True
         frmVDW.Visible = False
      Case 1
         Tempfrm.Visible = True
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         frmVDW.Visible = False
         paramfrm.Visible = False
      Case 2
         Tempfrm.Visible = False
         Pressfrm.Visible = True
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 3
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = True
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 4
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = True
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 5
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = True
         Terrfrm.Visible = False
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 6
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = True
         paramfrm.Visible = False
         frmVDW.Visible = False
      Case 7
         frmVDW.Visible = True
         Tempfrm.Visible = False
         Pressfrm.Visible = False
         Sunsfrm.Visible = False
         Rayfrm.Visible = False
         Tcfrm.Visible = False
         Terrfrm.Visible = False
         paramfrm.Visible = False
   End Select

End Sub

Private Sub updwnfit1_Change()
   txtFit1 = 260 + updwnfit1.Value * 3
End Sub

Private Sub updwnVA_Change()
   txtVA = updwnVA.Value * 0.05
End Sub

'Private Sub TabRef_KeyPress(KeyAscii As Integer)
'   If TabRef.Tab = 7 Then
'      KeyPressed = 1
'      End If
'End Sub
'
'Private Sub TabRef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If TabRef.Tab = 7 Then
'      KeyPressed = 1
'      End If
'End Sub
'
'Private Sub TabRef_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If TabRef.Tab = 7 Then
'      KeyPressed = 1
'      End If
'End Sub

'Private Sub TabRef_Click()
'   Select Case TabRef.Tab
'      Case 0
'         Tempfrm.Visible = False
'         Pressfrm.Visible = False
'         Sunsfrm.Visible = False
'         Rayfrm.Visible = False
'         Tcfrm.Visible = False
'         Terrfrm.Visible = False
'         paramfrm.Visible = True
'      Case 1
'         Tempfrm.Visible = True
'         Pressfrm.Visible = False
'         Sunsfrm.Visible = False
'         Rayfrm.Visible = False
'         Tcfrm.Visible = False
'         Terrfrm.Visible = False
'         paramfrm.Visible = False
'      Case 2
'         Tempfrm.Visible = False
'         Pressfrm.Visible = True
'         Sunsfrm.Visible = False
'         Rayfrm.Visible = False
'         Tcfrm.Visible = False
'         Terrfrm.Visible = False
'         paramfrm.Visible = False
'      Case 3
'         Tempfrm.Visible = False
'         Pressfrm.Visible = False
'         Sunsfrm.Visible = True
'         Rayfrm.Visible = False
'         Tcfrm.Visible = False
'         Terrfrm.Visible = False
'         paramfrm.Visible = False
'      Case 4
'         Tempfrm.Visible = False
'         Pressfrm.Visible = False
'         Sunsfrm.Visible = False
'         Rayfrm.Visible = True
'         Tcfrm.Visible = False
'         Terrfrm.Visible = False
'         paramfrm.Visible = False
'      Case 5
'         Tempfrm.Visible = False
'         Pressfrm.Visible = False
'         Sunsfrm.Visible = False
'         Rayfrm.Visible = False
'         Tcfrm.Visible = True
'         Terrfrm.Visible = False
'         paramfrm.Visible = False
'      Case 6
'         Tempfrm.Visible = False
'         Pressfrm.Visible = False
'         Sunsfrm.Visible = False
'         Rayfrm.Visible = False
'         Tcfrm.Visible = False
'         Terrfrm.Visible = True
'         paramfrm.Visible = False
'   End Select
'End Sub

Private Sub VScroll1_Change()
   prjAtmRefMainfm.Picture2.Top = prjAtmRefMainfm.Picture2.Top - prjAtmRefMainfm.VScroll1.Value
End Sub

Public Function MinValue(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Min = v1
Else: Min = v2
End If
End Function

Public Function MaxValue(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  max = v2
Else: max = v1
End If
End Function

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' source : wheel mouse hook: http://www.vbforums.com/showthread.php?388222-VB6-MouseWheel-with-Any-Control-%28originally-just-MSFlexGrid-Scrolling%29
'          two files examples: WheelHook-AllControls.zip, WheelHook-NestedControls.zip

' ===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    PictureBoxZoom prjAtmRefMainfm.picture1, MouseKeys, Rotation, Xpos, Ypos, 0
 
'  'original WheelWheel code for interacting with very controls on the form is below
'  Dim ctl As Control, cContainerCtl As Control
'  Dim bHandled As Boolean
'  Dim bOver As Boolean
'
'  For Each ctl In Controls
'    ' Is the mouse over the control
'    On Error Resume Next
'    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
'    On Error GoTo 0
'
'    If bOver Then
'      ' If so, respond accordingly
'      bHandled = True
'      Select Case True
'
'        Case TypeOf ctl Is MSFlexGrid
'          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
'
'        Case TypeOf ctl Is PictureBox, TypeOf ctl Is Frame
'          Set cContainerCtl = ctl
'          bHandled = False
'
'        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
'          ' These controls already handle the mousewheel themselves, so allow them to:
'          If ctl.Enabled Then ctl.SetFocus
'
'        Case Else
'          bHandled = False
'
'      End Select
'      If bHandled Then Exit Sub
'    End If
'    bOver = False
'    Debug.Print ctl.Name
'  Next ctl
'
'  If Not cContainerCtl Is Nothing Then
'    If TypeOf cContainerCtl Is PictureBox Then PictureBoxZoom prjAtmRefMainfm.Picture2, MouseKeys, Rotation, Xpos, Ypos, 0
'  Else
'    ' Scroll was not handled by any controls, so treat as a general message send to the form
'    GDMDIform.StatusBar1.Panels(1) = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
'  End If
End Sub
'MouseDown Event for Picture2
Public Function MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) As Integer
   On Error GoTo errhand
        
    Dim twipscX As Long, twipscY As Long
    twipscX = Screen.TwipsPerPixelX
    twipscY = Screen.TwipsPerPixelY

   
   
   If Button = 1 And _
      Not DigitizerEraser Then
      drag1x = x
      drag1y = y
      dragbegin = True
      drag2x = drag1x
      drag2y = drag1y
      End If

   MouseDown = 0
   Exit Function
   
errhand:
   MsgBox "Encountered error #: " & err.Number & vbLf & _
          err.Description & vbLf & _
          "in module: prjAtmRefMainfm.Picture2_MouseDown", _
          vbCritical + vbOKOnly, "AtmRef"
   MouseDown = -1

End Function
'mouseup event for prjAtmRefMainfm.picture1
Public Function MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) As Integer
  Dim VarD As Double
  Dim BytePosit As Long
  
  Dim RecNum&
  Dim Byte0 As Byte
  Dim Byte1 As Byte
  Dim Byte2 As Byte
  Dim Byte3 As Byte
  Dim Byte4 As Byte
  Dim Byte5 As Byte
  
  Byte0 = 0
  Byte1 = 1
  Byte2 = 2
  Byte3 = 3
  Byte4 = 4
  Byte5 = 5
  
  Dim color_line As Long, colornum%
  
  'heights
  Dim kmx As Long, kmy As Long
  Dim lt2 As Double, lg2 As Double, hgt2 As Integer
  
  Dim SearchCoord(1) As POINTAPI
  Dim ContourCoord(1) As POINTAPI
  Dim SmoothCoord(1) As POINTAPI
  Dim DTMCoord(1) As POINTAPI
  
  On Error GoTo errhand
  
  nearmouse_digi.x = x
  nearmouse_digi.y = y
  
  prjAtmRefMainfm.picture1.MousePointer = vbCrosshair 'restore crosshair cursor
  
    Xcoord = x
    Ycoord = y
    
    Select Case Button
       Case 1  'left button
          'shift this point to middle of screen
          'this will be the case when (X,Y) = (picture1.width/2, picture1.height/2)
          
gd50:
          If (drag1x = drag2x And drag1y = drag2y) Then 'And Not DigitizeOn Then
              dragbegin = False
              dragbox = False
              
              'reset center timer if flagged
              If ce& = 1 Then 'blinker was shut down during drag, so reenable it
                 ce& = 0 'reset blinker flag
                 GDMDIform.CenterPointTimer.Enabled = True
                 End If
          Else 'signales end of drag
             End If
             
                 
          prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.Value + H1
              
              
          H2 = drag1y - drag2y
          If prjAtmRefMainfm.VScroll1.Value + H2 < 0 Or prjAtmRefMainfm.VScroll1.Value + H2 > prjAtmRefMainfm.VScroll1.max Then
'                      'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
                If prjAtmRefMainfm.picture1.height > prjAtmRefMainfm.VScroll1.height Then

                   If prjAtmRefMainfm.VScroll1.Value + H2 < prjAtmRefMainfm.VScroll1.Min Then
                      prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.Min
                   ElseIf prjAtmRefMainfm.VScroll1.Value + H2 > prjAtmRefMainfm.VScroll1.max Then
                      prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.max
                      End If

                   End If
          Else
             prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.Value + H2
             End If
          
         'reset the drag flags
        
         picture1.DrawMode = 13
         drawbox = False
         dragbegin = False
         'reset drag coordinates
         drag1x = 0
         drag2x = 0
         drag1y = 0
         drag2y = 0
    
       Case Else
     End Select
   
    MouseUp = 0
                
    Exit Function
    
errhand:
         MsgBox "Encountered error #: " & err.Number & vbLf & _
             err.Description & vbLf & _
             "in module: prjAtmRefMainfm.Picture2_MouseUp", _
             vbCritical + vbOKOnly, "AtmRef"
   
   MouseUp = -1

End Function
'mousemouse of prjAtmRefMainfm.picture1
Public Function MouseMove(Button As Integer, step As Integer, x As Single, y As Single) As Integer
  'As cursor moves over map, display readout of coordinates.
  
  On Error GoTo errhand
  
  Dim next_mouse As POINTAPI
  
  nearmouse_digi.x = x
  nearmouse_digi.y = y
  
'  GetCursorPos next_mouse
'  hDnext = GetDC(0)
'  R = GetPixel(hDnext, next_mouse.X, next_mouse.Y)
'  If R <> -1 Then
'     Next_Color = recupcouleur(R)
'     GDMDIform.StatusBar1.Panels(1).Text = "RGB: " & Next_Color.R & "," & Next_Color.v & "," & Next_Color.b
'     End If
'
'  Call ReleaseDC(0, hDnext)
  

   '<<<<<<<<<<<<<<<<<<<<<<<end of new>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
      
   If dragbegin = True And Button = 1 And dragbox = True Then 'dragging continues, draw box
      'continue dragging
      picture1.DrawMode = 7
      picture1.DrawStyle = vbDot
      picture1.DrawWidth = 1
      
      'erase last drag box
      prjAtmRefMainfm.picture1.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
      
      'check if cursor left the picture frame, if so move scroll bars
      'to allow for dragging over entire map
      '(move the picture by the smallest increment = 1)
      If x / twipsx < Picture2.Left + HScroll1.Value Then
         'scroll map to right
         If HScroll1.Value - 1 >= HScroll1.Min Then
            HScroll1.Value = HScroll1.Value - 1
            End If
      ElseIf x / twipsx > Picture2.Width + Picture2.Left + HScroll1.Value Then
         'scroll map to left
         If HScroll1.Value + 1 <= HScroll1.max Then
            HScroll1.Value = HScroll1.Value + 1
            End If
         End If
      If y / twipsy < Picture2.Top + VScroll1.Value Then
         'scroll map down
         If VScroll1.Value - 1 >= VScroll1.Min Then
            VScroll1.Value = VScroll1.Value - 1
            End If
      ElseIf y / twipsy > Picture2.Top + Picture2.height + VScroll1.Value Then
         'scroll map up
         If VScroll1.Value + 1 <= VScroll1.max Then
            VScroll1.Value = VScroll1.Value + 1
            End If
         End If
      
      'draw new drag box
      prjAtmRefMainfm.picture1.Line (x, y)-(drag1x, drag1y), QBColor(15), B
'      GDMDIform.StatusBar1.Panels(1).Text = GDMDIform.StatusBar1.Panels(1).Text & "X,Y,drag1x,drag1y= " & str(x) + ", " & str(Y) & ", " & str(drag1x) & ", " & str(drag1y)

      prjAtmRefMainfm.picture1.Refresh
      
      'record new drag end coordinates
      drag2x = x: drag2y = y
      
      End If
      
  'Convert coordinates to pixels
  Xcoord = x / (twipsx * RefZoom.LastZoom)
  Ycoord = y / (twipsy * RefZoom.LastZoom)
  
  'Convert pixel coordinates to ITM
  ITMx = ((LRGeoX - ULGeoX) / pixwi) * Xcoord + ULGeoX
  ITMy = ULGeoY - ((ULGeoY - LRGeoY) / pixhi) * Ycoord
    
    
  MouseMove = 0
  
  Exit Function

errhand:
   Exit Function  '>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<
         MsgBox "Encountered error #: " & err.Number & vbLf & _
               err.Description & vbLf & _
               sEmpty, vbCritical + vbOKOnly, "AtmRef"
               

   MouseMove = -1

End Function

Private Sub VScrollPan_Change()
   AlphaSun.Refresh
End Sub

Private Sub VScrollPan_Scroll()
   AlphaSun.Refresh
End Sub
Public Function SQT(e As Double, dh As Double, r As Double) As Double
'calculates the approximte path length (approximate since path is actually
'curved and is not even spherical, and the law of cosines assumes Eucledan geometry)
'E is the angle the ray makes with horizontal = 90 - z, where z is the zenith angle
'DH is the incremental increase in the radia distance from the center of the earth from the last ray vertex
'RT is the radial distance of the last ray vertex.
Dim q As Double, y As Double

q = r * Sin(e * cd)
y = 2# * r + dh
If (dh < 0) Then SQT = -q - Sqr(Abs(q * q + dh * y))
If (dh >= 0) Then SQT = -q + Sqr(Abs(q * q + dh * y))

End Function

' Display the error.
'Mode% = 0 display error in textbox ErrorText
'      = 1 display error in label ErrorLbl
Private Sub ShowError(Optional ErrorText As TextBox, Optional ErrorLbl As label, Optional Mode%)
Dim err As Double

    ' Get the error.
    err = Sqr(ErrorSquared(PtX, PtY, BestCoeffs))
    If Mode% = 0 Then
       'form is text box
       ErrorText.Text = "MSD: " & Format$(err, "#0.0###")
    ElseIf Mode% = 1 Then
       ErrorLbl.Caption = "MSD: " & Format$(err, "#0.0###")
       End If
End Sub

' Display the error.
Private Sub ShowLinError(Optional ErrorText As TextBox, Optional ErrorLbl As label, Optional Mode%)
Dim err As Double

    ' Get the error.
    err = Sqr(LinErrorSquared( _
        PtX, PtY, BestM, BestB))
    If Mode% = 0 Then
       'form is text box
       ErrorText.Text = "MSD: " & Format$(err, "#0.0###")
    ElseIf Mode% = 1 Then
       ErrorLbl.Caption = "MSD: " & Format$(err, "#0.0###")
       End If
End Sub
