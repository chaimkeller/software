VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form GDOptionsfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paths/Options"
   ClientHeight    =   8475
   ClientLeft      =   6780
   ClientTop       =   1035
   ClientWidth     =   5265
   Icon            =   "GDOptionsfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   5265
   Begin TabDlg.SSTab tabOptions 
      Height          =   8500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   15002
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Maps"
      TabPicture(0)   =   "GDOptionsfrm.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmLR"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDefaults"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAddMap"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSaveMaps"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frmGrid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "frmLL"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmUR"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Settings"
      TabPicture(1)   =   "GDOptionsfrm.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmHardy"
      Tab(1).Control(1)=   "frmClickCenter"
      Tab(1).Control(2)=   "frmContours"
      Tab(1).Control(3)=   "Frame15"
      Tab(1).Control(4)=   "Frame9"
      Tab(1).Control(5)=   "frmAnalysis"
      Tab(1).Control(6)=   "Frame2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Paths"
      TabPicture(2)   =   "GDOptionsfrm.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame20"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame21"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "frmKML"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "frmNewDTM"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "DTM"
      TabPicture(3)   =   "GDOptionsfrm.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmDTM"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame frmNewDTM 
         Caption         =   "generated DTM files folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1185
         Left            =   -74760
         TabIndex        =   149
         Top             =   3760
         Width           =   4695
         Begin VB.TextBox txtNewDTM 
            Height          =   285
            Left            =   360
            TabIndex        =   151
            ToolTipText     =   "Path to 1 arcsec hgt format files"
            Top             =   720
            Width           =   3975
         End
         Begin VB.CommandButton cmdBrowseNewDTM 
            Caption         =   "&Browse"
            Height          =   315
            Left            =   3240
            TabIndex        =   150
            ToolTipText     =   "Define folder for placing generated DTM files"
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame frmHardy 
         Caption         =   " Hardy Quadratic Surfaces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   520
         Left            =   -74760
         TabIndex        =   142
         Top             =   5320
         Width           =   4695
         Begin VB.CheckBox chkSave_xyz 
            Caption         =   "Write xyz data and horizon profiles to files"
            Height          =   195
            Left            =   650
            TabIndex        =   148
            ToolTipText     =   "Save xyz data and horizon prfiles at conclusion of Hardy Quadratic Surfaces calculation"
            Top             =   220
            Width           =   3375
         End
         Begin VB.OptionButton optGaussian 
            Caption         =   "Use Gauss.Elimination (fast)"
            Height          =   195
            Left            =   2330
            TabIndex        =   144
            ToolTipText     =   "solve system of equations using Gaussian eliminiation"
            Top             =   240
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.OptionButton optCramer 
            Caption         =   "Use matrix inversion (slow)"
            Height          =   195
            Left            =   120
            TabIndex        =   143
            ToolTipText     =   "Solve matrix equation by matrix inversion"
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.Frame frmDTM 
         Caption         =   "DTM creating - grid size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   7815
         Left            =   -74760
         TabIndex        =   130
         Top             =   480
         Width           =   4695
         Begin VB.Frame frmUnits 
            Caption         =   "DTM elevation units"
            Height          =   1815
            Left            =   2640
            TabIndex        =   165
            ToolTipText     =   "Digitized input is converted to meters "
            Top             =   1560
            Width           =   1815
            Begin VB.OptionButton optDefault 
               Caption         =   "meters"
               Height          =   195
               Left            =   360
               TabIndex        =   171
               ToolTipText     =   "Record DTM heights using meters (default)"
               Top             =   260
               Width           =   975
            End
            Begin VB.TextBox txtCustom 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   170
               Text            =   "1.0"
               ToolTipText     =   "Conversion factor from meters to the custom unit"
               Top             =   1420
               Width           =   1215
            End
            Begin VB.OptionButton optother 
               Caption         =   "custom"
               Height          =   195
               Left            =   360
               TabIndex        =   169
               ToolTipText     =   "Enter custom coversion factor from meters"
               Top             =   1200
               Width           =   975
            End
            Begin VB.OptionButton optdecimeters 
               Caption         =   "decimeters"
               Height          =   375
               Left            =   360
               TabIndex        =   168
               ToolTipText     =   "record DTM's heights using decimeters"
               Top             =   880
               Width           =   1215
            End
            Begin VB.OptionButton optfathoms 
               Caption         =   "fathoms"
               Height          =   315
               Left            =   360
               TabIndex        =   167
               ToolTipText     =   "Record DTM heights using fathoms"
               Top             =   680
               Width           =   975
            End
            Begin VB.OptionButton optfeet 
               Caption         =   "feet"
               Height          =   255
               Left            =   360
               TabIndex        =   166
               ToolTipText     =   "Record DTM heights using feet"
               Top             =   460
               Width           =   975
            End
         End
         Begin VB.Frame frmhgt 
            Caption         =   "Elevation output place accuracy"
            Height          =   1215
            Left            =   240
            TabIndex        =   160
            Top             =   5520
            Width           =   4215
            Begin VB.OptionButton optDouble 
               Caption         =   "8 bytes double precision (meters)"
               Height          =   375
               Left            =   840
               TabIndex        =   163
               ToolTipText     =   "Uses 16 times the memory"
               Top             =   760
               Width           =   2655
            End
            Begin VB.OptionButton optFloat 
               Caption         =   "4 bytes floating point (meters)"
               Height          =   195
               Left            =   840
               TabIndex        =   162
               ToolTipText     =   "Uses four times the memory"
               Top             =   560
               Width           =   2415
            End
            Begin VB.OptionButton optInteger 
               Caption         =   "2 bytes integer (meters)"
               Height          =   255
               Left            =   840
               TabIndex        =   161
               ToolTipText     =   "Use this resolution to save memory and allow calculating larger areas"
               Top             =   260
               Width           =   2415
            End
         End
         Begin VB.Frame frmProfile 
            Caption         =   "Azimuth range for profile viewer"
            Height          =   1935
            Left            =   240
            TabIndex        =   153
            Top             =   3480
            Width           =   4215
            Begin VB.TextBox txtaprn 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2280
               TabIndex        =   158
               Text            =   "0.5"
               ToolTipText     =   "Amount to shave off (nearest approach) in kilometers"
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txtStepAzi 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2280
               TabIndex        =   155
               Text            =   "0.1"
               ToolTipText     =   "azimuth spacing in generated profile"
               Top             =   840
               Width           =   1455
            End
            Begin VB.TextBox txtAzi 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2280
               TabIndex        =   154
               Text            =   "45"
               ToolTipText     =   "Total range is double this value"
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label lblaprn 
               Caption         =   "Amount to shave (km):"
               Height          =   255
               Left            =   360
               TabIndex        =   159
               Top             =   1380
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "Azimuth Spacing (deg.)"
               Height          =   255
               Left            =   360
               TabIndex        =   157
               Top             =   900
               Width           =   1695
            End
            Begin VB.Label lblAzimuth 
               Caption         =   "Half azimuth range (deg.)"
               Height          =   255
               Left            =   360
               TabIndex        =   156
               Top             =   400
               Width           =   1815
            End
         End
         Begin VB.CheckBox chkNewDTM 
            Caption         =   "Use heights based on generated DTM if available"
            Height          =   195
            Left            =   360
            TabIndex        =   152
            ToolTipText     =   "If generated a new DTM of a region, use its heights when moving the mouse"
            Top             =   6840
            Width           =   3855
         End
         Begin VB.CommandButton cmdSaveNewDTM 
            Caption         =   "Save"
            Height          =   495
            Left            =   1320
            TabIndex        =   133
            Top             =   7200
            Width           =   2055
         End
         Begin VB.Frame frmGeo 
            Caption         =   "New lat/lon grid"
            Height          =   1815
            Left            =   240
            TabIndex        =   132
            Top             =   1560
            Width           =   2330
            Begin VB.TextBox txtDTMlat 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   137
               Text            =   "0"
               Top             =   1200
               Width           =   1815
            End
            Begin VB.TextBox txtDTMlon 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   136
               Text            =   "0"
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label lbllat 
               Caption         =   "Y grid spacing (arc secs)"
               Height          =   255
               Left            =   300
               TabIndex        =   141
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label lbllon 
               Caption         =   "X grid spacing (arc secs)"
               Height          =   255
               Left            =   300
               TabIndex        =   140
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame frmNewITM 
            Caption         =   "New ITM or UTM grid"
            Height          =   1215
            Left            =   240
            TabIndex        =   131
            Top             =   240
            Width           =   4215
            Begin VB.TextBox txtDTMitmy 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2280
               TabIndex        =   135
               Text            =   "0"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox txtDTMitmx 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2280
               TabIndex        =   134
               Text            =   "0"
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbliTMy 
               Caption         =   "Y grid spacing (meters)"
               Height          =   255
               Left            =   480
               TabIndex        =   139
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label lblITMx 
               Caption         =   "X grid spacing (meters)"
               Height          =   255
               Left            =   480
               TabIndex        =   138
               Top             =   360
               Width           =   1695
            End
         End
      End
      Begin VB.Frame frmClickCenter 
         Caption         =   "Point Digitizing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   -74760
         TabIndex        =   124
         Top             =   6650
         Width           =   4695
         Begin VB.CheckBox chkCenterClick 
            Caption         =   "Don't center on clicking while digitizing points"
            Height          =   255
            Left            =   600
            TabIndex        =   125
            ToolTipText     =   "If checked, clicking doesn't recenter the screen when digitizing points"
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame frmContours 
         Caption         =   "Semi-automatic contour detection method"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   720
         Left            =   -74760
         TabIndex        =   121
         Top             =   5880
         Width           =   4695
         Begin VB.OptionButton optBug 
            Caption         =   "Four directional Freeman chain code method"
            Height          =   255
            Left            =   600
            TabIndex        =   123
            ToolTipText     =   "Use the Freeman 4 directional chain code for digitizing contours"
            Top             =   200
            Width           =   3615
         End
         Begin VB.OptionButton optFreeman 
            Caption         =   "Eight directional Freeman chain code method"
            Height          =   195
            Left            =   600
            TabIndex        =   122
            ToolTipText     =   "Use the eight directiona Freeman chain code for digitizing contours"
            Top             =   440
            Width           =   3615
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Save settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   525
         Left            =   -74760
         TabIndex        =   119
         Top             =   7800
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CheckBox chkSave 
            Caption         =   "&Save these settings upon exiting"
            Height          =   195
            Left            =   960
            TabIndex        =   120
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Warnings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   -74760
         TabIndex        =   116
         Top             =   7200
         Width           =   4695
         Begin VB.CheckBox chkErrorMessage 
            Caption         =   "&Ignore missing paths"
            Height          =   195
            Left            =   300
            TabIndex        =   118
            ToolTipText     =   "Don't report missing paths"
            Top             =   240
            Width           =   1755
         End
         Begin VB.CheckBox chkAutoRedraw 
            Caption         =   "Ignore &AutoRedraw errors"
            Height          =   195
            Left            =   2280
            TabIndex        =   117
            ToolTipText     =   "Ignore autoredraw error due to insufficient graphics memory"
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame frmKML 
         Caption         =   "Google Earth"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   -74760
         TabIndex        =   104
         Top             =   5040
         Width           =   4695
         Begin MSComDlg.CommonDialog cmdlgArc 
            Left            =   2400
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtGoogle 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   111
            Top             =   560
            Width           =   4395
         End
         Begin VB.CommandButton cmdBrowseGoogle 
            Caption         =   "&Browse"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3240
            TabIndex        =   110
            Top             =   200
            Width           =   915
         End
         Begin VB.TextBox txtOutCropIcon 
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
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   120
            TabIndex        =   109
            Text            =   "http://maps.google.com/mapfiles/kml/pal4/icon49.png"
            ToolTipText     =   "enter complete URL of outcroppings' icon"
            Top             =   2040
            Width           =   4395
         End
         Begin VB.TextBox txtWellIcon 
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
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   120
            TabIndex        =   108
            Text            =   "http://maps.google.com/mapfiles/kml/pal4/icon48.png"
            ToolTipText     =   "enter complete URL of wells' icon"
            Top             =   2640
            Width           =   4275
         End
         Begin VB.TextBox txtkml 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   107
            ToolTipText     =   "path of kml files"
            Top             =   1240
            Width           =   4275
         End
         Begin VB.CommandButton cmdBrowseKML 
            Caption         =   "&Browse"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3240
            TabIndex        =   106
            Top             =   880
            Width           =   915
         End
         Begin VB.CheckBox chkKML 
            Caption         =   "Edit icons"
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
            Height          =   195
            Left            =   120
            TabIndex        =   105
            ToolTipText     =   "Click to enable editing of icons' URL"
            Top             =   1600
            Width           =   2295
         End
         Begin VB.Label lblGoogle 
            Caption         =   "Path to GoogleEarth.exe"
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   120
            TabIndex        =   115
            Top             =   320
            Width           =   2055
         End
         Begin VB.Label lblOutCropIcon 
            Caption         =   "URL of Google Icon to use for Points"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   114
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Label lblWellIcon 
            Caption         =   "URL of Google Icon to use for Lines/Contours"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   113
            Top             =   2400
            Width           =   3735
         End
         Begin VB.Label lblKmlPath 
            Caption         =   "kml files directory"
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   120
            TabIndex        =   112
            Top             =   1000
            Width           =   2055
         End
      End
      Begin VB.Frame frmAnalysis 
         Caption         =   "Analysis parameters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2350
         Left            =   -74760
         TabIndex        =   92
         Top             =   2940
         Width           =   4695
         Begin VB.ComboBox cmbContour 
            Height          =   315
            ItemData        =   "GDOptionsfrm.frx":04B2
            Left            =   3720
            List            =   "GDOptionsfrm.frx":04D7
            TabIndex        =   164
            Text            =   "cmbContour"
            Top             =   1240
            Width           =   735
         End
         Begin MSComCtl2.UpDown updwnEraseBrushSize 
            Height          =   285
            Left            =   4200
            TabIndex        =   147
            Top             =   1930
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtEraserBrushSize"
            BuddyDispid     =   196671
            OrigLeft        =   4200
            OrigTop         =   1950
            OrigRight       =   4455
            OrigBottom      =   2235
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtEraserBrushSize 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3720
            TabIndex        =   146
            Text            =   "1"
            Top             =   1950
            Width           =   465
         End
         Begin MSComCtl2.UpDown UpDownPixelSearch 
            Height          =   285
            Left            =   4200
            TabIndex        =   128
            Top             =   1600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   10
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDistPixelSearch"
            BuddyDispid     =   196672
            OrigLeft        =   4080
            OrigTop         =   1920
            OrigRight       =   4335
            OrigBottom      =   2295
            Max             =   100
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDistPixelSearch 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3720
            TabIndex        =   127
            Text            =   "10"
            ToolTipText     =   "Maximum hover distance to search in pixels for a digitized point, contour,  or line"
            Top             =   1600
            Width           =   480
         End
         Begin MSComCtl2.UpDown UpDownSensitivity 
            Height          =   285
            Left            =   4200
            TabIndex        =   102
            Top             =   900
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   50
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtSensitivity"
            BuddyDispid     =   196673
            OrigLeft        =   3960
            OrigTop         =   1200
            OrigRight       =   4215
            OrigBottom      =   1575
            Max             =   100
            Min             =   1
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtSensitivity 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3720
            TabIndex        =   101
            Text            =   "50"
            Top             =   900
            Width           =   480
         End
         Begin MSComCtl2.UpDown UpDownDistLines 
            Height          =   285
            Left            =   4200
            TabIndex        =   99
            Top             =   550
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   5
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDistLines"
            BuddyDispid     =   196674
            OrigLeft        =   4080
            OrigTop         =   720
            OrigRight       =   4335
            OrigBottom      =   975
            Max             =   20
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDistLines 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3720
            TabIndex        =   98
            Text            =   "5"
            Top             =   550
            Width           =   480
         End
         Begin MSComCtl2.UpDown UpDownContour 
            Height          =   285
            Left            =   4200
            TabIndex        =   96
            Top             =   225
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   5
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDistContour"
            BuddyDispid     =   196675
            OrigLeft        =   3840
            OrigTop         =   360
            OrigRight       =   4095
            OrigBottom      =   615
            Max             =   15
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDistContour 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3720
            TabIndex        =   95
            Text            =   "5"
            Top             =   225
            Width           =   480
         End
         Begin VB.Label lblEraserBrushSize 
            Caption         =   "Minimum Digitizing Eraser Brush Size"
            Height          =   255
            Left            =   240
            TabIndex        =   145
            Top             =   2000
            Width           =   3255
         End
         Begin VB.Label lblHighlightSearch 
            Caption         =   "Hoverdistance in pixels while in Edit mode"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   1690
            Width           =   3375
         End
         Begin VB.Label lblContours 
            Caption         =   "Starting contour interval for displaying contours of Hardy quadratic surfaces"
            Height          =   495
            Left            =   240
            TabIndex        =   103
            Top             =   1280
            Width           =   3375
         End
         Begin VB.Label lblInitSens 
            Caption         =   "Starting Euclidean color sensitivity for contour color Euclidean comparison detection"
            Height          =   495
            Left            =   240
            TabIndex        =   100
            ToolTipText     =   "Sensitivity for accepting color as the same as the  contour's starting color"
            Top             =   855
            Width           =   3375
         End
         Begin VB.Label lblDistLines 
            Caption         =   "Distance between points of a digitized line"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            ToolTipText     =   "When converting digitized lines to points, create a point every this many pixels"
            Top             =   550
            Width           =   3135
         End
         Begin VB.Label lblDistPoints 
            Caption         =   "Distance between points of a digitized contour"
            Height          =   240
            Left            =   240
            TabIndex        =   94
            ToolTipText     =   "When converting a contour to points, create a point every this many pixels"
            Top             =   225
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Plotting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2595
         Left            =   -74760
         TabIndex        =   80
         Top             =   360
         Width           =   4695
         Begin MSComCtl2.UpDown UpDownMaxRecords 
            Height          =   360
            Left            =   3960
            TabIndex        =   87
            Top             =   2100
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   635
            _Version        =   393216
            Value           =   12000
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtMaxHighlight"
            BuddyDispid     =   196683
            OrigLeft        =   3960
            OrigTop         =   2280
            OrigRight       =   4215
            OrigBottom      =   2535
            Max             =   32768
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMaxHighlight 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3360
            TabIndex        =   86
            Text            =   "12000"
            Top             =   2100
            Width           =   660
         End
         Begin VB.Frame Frame18 
            ForeColor       =   &H00C00000&
            Height          =   1820
            Left            =   380
            TabIndex        =   81
            Top             =   200
            Width           =   3960
            Begin VB.CheckBox chkRainbow 
               Caption         =   "*use colors based on elevation"
               Height          =   195
               Left            =   250
               TabIndex        =   129
               ToolTipText     =   "Use elevation colors instead of fixed colors for digitized lines"
               Top             =   1460
               Width           =   3455
            End
            Begin VB.CommandButton cmdRSColor 
               Caption         =   "&Color"
               Height          =   315
               Left            =   3000
               TabIndex        =   91
               Top             =   1050
               Width           =   675
            End
            Begin VB.CommandButton cmdContours 
               Caption         =   "&Color"
               Height          =   315
               Left            =   3000
               TabIndex        =   90
               Top             =   380
               Width           =   675
            End
            Begin VB.CommandButton cmdPointColor 
               Caption         =   "&Color"
               Height          =   315
               Left            =   1140
               TabIndex        =   83
               Top             =   380
               Width           =   675
            End
            Begin VB.CommandButton cmdLineColor 
               Caption         =   "&Color"
               Height          =   315
               Left            =   1140
               TabIndex        =   82
               Top             =   1050
               Width           =   675
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Caption         =   "Rubber Sheeting"
               Height          =   195
               Left            =   2100
               TabIndex        =   93
               Top             =   810
               Width           =   1455
            End
            Begin VB.Shape shpRS 
               FillColor       =   &H0000FFFF&
               FillStyle       =   0  'Solid
               Height          =   315
               Left            =   2040
               Top             =   1050
               Width           =   855
            End
            Begin VB.Shape shpContours 
               FillColor       =   &H00800000&
               FillStyle       =   0  'Solid
               Height          =   315
               Left            =   2040
               Top             =   380
               Width           =   855
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "Contours"
               Height          =   195
               Left            =   2040
               TabIndex        =   89
               Top             =   150
               Width           =   1455
            End
            Begin VB.Shape shpPoints 
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   315
               Left            =   180
               Top             =   380
               Width           =   855
            End
            Begin VB.Shape shpLines 
               FillColor       =   &H0000FF00&
               FillStyle       =   0  'Solid
               Height          =   315
               Left            =   180
               Top             =   1050
               Width           =   855
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               Caption         =   "Points"
               Height          =   195
               Left            =   300
               TabIndex        =   85
               Top             =   150
               Width           =   1455
            End
            Begin VB.Label Label25 
               Alignment       =   2  'Center
               Caption         =   "Lines*"
               Height          =   195
               Left            =   600
               TabIndex        =   84
               Top             =   810
               Width           =   855
            End
         End
         Begin MSComDlg.CommonDialog cmdlgColor 
            Left            =   1920
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin VB.Label Label28 
            Caption         =   "Maximum no. of search results to plot simultaneously (depends on computer)"
            Height          =   375
            Left            =   480
            TabIndex        =   88
            Top             =   2055
            Width           =   2715
         End
      End
      Begin VB.Frame frmUR 
         Caption         =   "Map Geographic Coordinates of Upper-Left Grid Intersection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         TabIndex        =   68
         Top             =   6100
         Width           =   4935
         Begin VB.TextBox txtULGridY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   78
            Text            =   "txtULGridY"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtULGridX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   77
            Text            =   "txtULGridX"
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdULGridXchange 
            Enabled         =   0   'False
            Height          =   240
            Left            =   2280
            Picture         =   "GDOptionsfrm.frx":0503
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Convert coordiate display"
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton cmdULGridYchange 
            Enabled         =   0   'False
            Height          =   240
            Left            =   4440
            Picture         =   "GDOptionsfrm.frx":064D
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Convert coordinate display"
            Top             =   480
            Width           =   255
         End
         Begin VB.Line LineHor2 
            X1              =   240
            X2              =   740
            Y1              =   550
            Y2              =   550
         End
         Begin VB.Line LineHor1 
            X1              =   240
            X2              =   720
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Line LineGrid3 
            X1              =   600
            X2              =   600
            Y1              =   320
            Y2              =   680
         End
         Begin VB.Line LineGrid2 
            X1              =   480
            X2              =   480
            Y1              =   320
            Y2              =   680
         End
         Begin VB.Line LineGrid1 
            X1              =   360
            X2              =   360
            Y1              =   320
            Y2              =   680
         End
         Begin VB.Shape Shape2 
            Height          =   375
            Index           =   2
            Left            =   220
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "Y-Geo"
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
            Left            =   3600
            TabIndex        =   72
            Top             =   220
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "X-Geo"
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
            Index           =   7
            Left            =   1440
            TabIndex        =   71
            Top             =   220
            Width           =   495
         End
         Begin VB.Shape Shape18 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   300
            Shape           =   3  'Circle
            Top             =   365
            Width           =   135
         End
         Begin VB.Shape Shape17 
            Height          =   615
            Left            =   120
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame frmLL 
         Caption         =   "Map Geographic Coordinates of Lower-Right Grid Intersection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         TabIndex        =   61
         Top             =   7000
         Width           =   4935
         Begin VB.TextBox txtLRGridY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   65
            Text            =   "txtLRGridY"
            ToolTipText     =   "Use ""-"" for separte deg-min-sec"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtLRGridX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   64
            Text            =   "txtLRGridX"
            ToolTipText     =   "Use ""-"" for separte deg-min-sec"
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdLRGridXchange 
            Enabled         =   0   'False
            Height          =   240
            Left            =   2280
            Picture         =   "GDOptionsfrm.frx":0797
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Convert coordinate display"
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton cmdLRGridYchange 
            Enabled         =   0   'False
            Height          =   240
            Left            =   4440
            Picture         =   "GDOptionsfrm.frx":08E1
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Convert coordinate display"
            Top             =   480
            Width           =   255
         End
         Begin VB.Line LineHor4 
            X1              =   240
            X2              =   720
            Y1              =   550
            Y2              =   550
         End
         Begin VB.Line LineHor3 
            X1              =   240
            X2              =   720
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Line LineVer3 
            X1              =   600
            X2              =   600
            Y1              =   320
            Y2              =   680
         End
         Begin VB.Line LineVer2 
            X1              =   480
            X2              =   480
            Y1              =   320
            Y2              =   680
         End
         Begin VB.Line LineVer1 
            X1              =   360
            X2              =   360
            Y1              =   320
            Y2              =   680
         End
         Begin VB.Shape Shape2 
            Height          =   375
            Index           =   1
            Left            =   240
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label26 
            Caption         =   "Y-Geo"
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
            Left            =   3600
            TabIndex        =   67
            Top             =   220
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "X-Geo"
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
            Index           =   6
            Left            =   1440
            TabIndex        =   66
            Top             =   220
            Width           =   495
         End
         Begin VB.Shape Shape16 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   540
            Shape           =   3  'Circle
            Top             =   500
            Width           =   135
         End
         Begin VB.Shape Shape15 
            Height          =   615
            Left            =   120
            Top             =   200
            Width           =   735
         End
      End
      Begin VB.Frame frmGrid 
         Caption         =   "Num. major grid lines (include the outside margins if they are on the grid)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   680
         Left            =   120
         TabIndex        =   42
         Top             =   5430
         Width           =   4935
         Begin MSComCtl2.UpDown UpDownGridY 
            Height          =   285
            Left            =   4200
            TabIndex        =   47
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtGridY"
            BuddyDispid     =   196728
            OrigLeft        =   3960
            OrigTop         =   290
            OrigRight       =   4215
            OrigBottom      =   545
            Max             =   10000
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtGridY 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3240
            TabIndex        =   46
            Text            =   "0"
            ToolTipText     =   "Number of y grids"
            Top             =   290
            Width           =   960
         End
         Begin MSComCtl2.UpDown UpDownGridX 
            Height          =   285
            Left            =   1920
            TabIndex        =   44
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtGridX"
            BuddyDispid     =   196729
            OrigLeft        =   1680
            OrigTop         =   240
            OrigRight       =   1935
            OrigBottom      =   615
            Max             =   10000
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtGridX 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1080
            TabIndex        =   43
            Text            =   "0"
            ToolTipText     =   "Number of major x grids"
            Top             =   290
            Width           =   960
         End
         Begin VB.Label lblGridY 
            Caption         =   "Y Grid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   48
            Top             =   315
            Width           =   615
         End
         Begin VB.Label lblXgrid 
            Caption         =   "X Grid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   360
            TabIndex        =   45
            Top             =   315
            Width           =   615
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   " ground elevations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   830
         Left            =   -74760
         TabIndex        =   39
         Top             =   2880
         Width           =   4695
         Begin VB.OptionButton optDTM 
            Caption         =   "1 arcsec hgt format (SRTM) DTM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            TabIndex        =   41
            ToolTipText     =   "Use 1 arcsec SRTM hgt file format DTM"
            Top             =   420
            Width           =   2895
         End
         Begin VB.OptionButton optAster 
            Caption         =   "1 arcsec bil format (ASTER) DTM"
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
            Left            =   960
            TabIndex        =   40
            ToolTipText     =   "Use ASTER ground elevations"
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "ASTER V2 1 arc sec Digital Terrain Model"
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
         Height          =   1155
         Left            =   -74760
         TabIndex        =   35
         Top             =   480
         Width           =   4695
         Begin VB.TextBox txtAster 
            Height          =   285
            Left            =   360
            TabIndex        =   37
            ToolTipText     =   "Path to the ASTER DTM"
            Top             =   720
            Width           =   3975
         End
         Begin VB.CommandButton cmdBrowseASTER 
            Caption         =   "&Browse"
            Height          =   315
            Left            =   3240
            TabIndex        =   36
            ToolTipText     =   "Browse for Aster DTM"
            Top             =   240
            Width           =   915
         End
         Begin MSComDlg.CommonDialog cmdlgASTER 
            Left            =   3420
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label21 
            Caption         =   "Path to ASTER type 1 arcsec bil files"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   400
            Width           =   2835
         End
      End
      Begin VB.CommandButton cmdSaveMaps 
         Caption         =   "&Save maps"
         Height          =   315
         Left            =   2640
         TabIndex        =   34
         Top             =   8040
         Width           =   1515
      End
      Begin VB.CommandButton cmdAddMap 
         Caption         =   "&Add map"
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   "Save the map name and parameters"
         Top             =   8040
         Width           =   1275
      End
      Begin VB.CommandButton cmdDefaults 
         Caption         =   "&Restore default map"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Restore the default map and map boundaries"
         Top             =   8040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame6 
         Caption         =   "1 arcsec hgt format (SRTM) files"
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
         Height          =   1155
         Left            =   -74760
         TabIndex        =   29
         Top             =   1680
         Width           =   4695
         Begin VB.CommandButton cmdBrowseDTM 
            Caption         =   "&Browse"
            Height          =   315
            Left            =   3240
            TabIndex        =   31
            ToolTipText     =   "Browse for 1 arcsec hgt format files"
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox txtdtm 
            Height          =   285
            Left            =   360
            TabIndex        =   30
            ToolTipText     =   "Path to 1 arcsec hgt format files"
            Top             =   720
            Width           =   3975
         End
         Begin MSComDlg.CommonDialog cmdlgDTM 
            Left            =   3420
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbldtm 
            Caption         =   "Path to SRTM type 1 arcsec hgt files"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   32
            Top             =   400
            Width           =   2835
         End
      End
      Begin VB.Frame frmLR 
         Caption         =   "Map Geographic Coordinates of Lower-Right Map Border"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   120
         TabIndex        =   25
         Top             =   4410
         Width           =   4935
         Begin VB.CommandButton cmdPasteLRPixY 
            Height          =   255
            Left            =   2400
            Picture         =   "GDOptionsfrm.frx":0A2B
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Click to paste las clicked screen y coordinatge"
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdPasteLRPixX 
            Height          =   255
            Left            =   2400
            Picture         =   "GDOptionsfrm.frx":0B75
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Click to paste las clicked screen x coordinatge"
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtLRPixY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   58
            Text            =   "txtLRPixY"
            ToolTipText     =   "y pixel coordinate of lower right  corner of map (max: height - 1)"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtLRPixX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   57
            Text            =   "txtLRPixX"
            ToolTipText     =   "x pixel coordinate of lower right corner of map (max: Width - 1)"
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdLRYconvert 
            Enabled         =   0   'False
            Height          =   240
            Left            =   4480
            Picture         =   "GDOptionsfrm.frx":0CBF
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Convert coordinate display"
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdLRXconvert 
            Enabled         =   0   'False
            Height          =   240
            Left            =   4480
            Picture         =   "GDOptionsfrm.frx":0E09
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Convert coordinate display"
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtLRGeoX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3280
            TabIndex        =   9
            Text            =   "txtLRGeoX"
            ToolTipText     =   "Use ""-"" for separte deg-min-sec"
            Top             =   260
            Width           =   1215
         End
         Begin VB.TextBox txtLRGeoY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3280
            TabIndex        =   10
            Text            =   "txtLRGeoY"
            ToolTipText     =   "Use ""-"" for separte deg-min-sec"
            Top             =   620
            Width           =   1215
         End
         Begin VB.Label lblLRXpix 
            Caption         =   "Y Pixel"
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
            Left            =   920
            TabIndex        =   60
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblLRX 
            Caption         =   "X Pixel"
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
            Left            =   920
            TabIndex        =   59
            Top             =   285
            Width           =   495
         End
         Begin VB.Shape ShapeB3 
            Height          =   615
            Left            =   120
            Top             =   240
            Width           =   735
         End
         Begin VB.Shape Circle4 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   640
            Shape           =   3  'Circle
            Top             =   640
            Width           =   135
         End
         Begin VB.Label Label5 
            Caption         =   "X-Geo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   28
            Top             =   285
            Width           =   555
         End
         Begin VB.Label Label7 
            Caption         =   "Y-Geo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   27
            Top             =   600
            Width           =   495
         End
         Begin VB.Shape Shape3 
            Height          =   375
            Left            =   220
            Top             =   345
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   15
            Left            =   960
            TabIndex        =   26
            Top             =   780
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Geographic Coordinates of Upper-Left Map Border"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   120
         TabIndex        =   22
         Top             =   3370
         Width           =   4935
         Begin VB.CommandButton cmdPasteULPixY 
            Height          =   255
            Left            =   2400
            Picture         =   "GDOptionsfrm.frx":0F53
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Click to paste las clicked screen y coordinatge"
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdPasteULPixX 
            Height          =   255
            Left            =   2400
            Picture         =   "GDOptionsfrm.frx":109D
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Click to paste las clicked screen x coordinatge"
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtULPixY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   54
            Text            =   "txtULPixY"
            ToolTipText     =   "y pixel coordinate of upper left corner of map (min: 0)"
            Top             =   620
            Width           =   975
         End
         Begin VB.TextBox txtULPixX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   53
            Text            =   "txtULPixX"
            ToolTipText     =   "x pixel coordinate of upper left corner of map (min: 0)"
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdULYconvert 
            Enabled         =   0   'False
            Height          =   240
            Left            =   4480
            Picture         =   "GDOptionsfrm.frx":11E7
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Conver coordinate display"
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdULXconvert 
            Enabled         =   0   'False
            Height          =   240
            Left            =   4480
            Picture         =   "GDOptionsfrm.frx":1331
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "convert coordinate display"
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtULGeoX 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3280
            TabIndex        =   7
            Text            =   "txtULGeoX"
            ToolTipText     =   "Use ""-"" for separte deg-min-sec"
            Top             =   260
            Width           =   1215
         End
         Begin VB.TextBox txtULGeoY 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3280
            TabIndex        =   8
            Text            =   "txtULGeoY"
            ToolTipText     =   "Use ""-"" for separte deg-min-sec"
            Top             =   620
            Width           =   1215
         End
         Begin VB.Label lblULPixY 
            Caption         =   "Y Pixel"
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
            Left            =   920
            TabIndex        =   56
            Top             =   645
            Width           =   495
         End
         Begin VB.Label lbULPixX 
            Caption         =   "X Pixel"
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
            Left            =   920
            TabIndex        =   55
            Top             =   315
            Width           =   495
         End
         Begin VB.Shape ShapeB1 
            Height          =   615
            Left            =   120
            Top             =   270
            Width           =   735
         End
         Begin VB.Shape Circle3 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   200
            Shape           =   3  'Circle
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label3 
            Caption         =   "X-Geo"
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
            Index           =   1
            Left            =   2760
            TabIndex        =   24
            Top             =   285
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Y-Geo"
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
            Left            =   2760
            TabIndex        =   23
            Top             =   645
            Width           =   555
         End
         Begin VB.Shape Shape2 
            Height          =   375
            Index           =   0
            Left            =   220
            Top             =   405
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Current Map"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1900
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   4935
         Begin VB.ComboBox cmbMaps 
            Height          =   315
            Left            =   360
            TabIndex        =   3
            Text            =   "cmbMaps"
            Top             =   240
            Width           =   4335
         End
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   255
            TabIndex        =   17
            Top             =   660
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtPixWidth 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   5
            Text            =   "txtPixWidth"
            ToolTipText     =   "Map's X Size in Pixels"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtPixHeight 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   6
            Text            =   "txtPixHeight"
            ToolTipText     =   "Map's Y Size in Pixels"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.CommandButton cmdPixelSize 
            Caption         =   "&Auto Determine Pixel Size"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1120
            TabIndex        =   4
            ToolTipText     =   "Automatically determine map's pixel size"
            Top             =   660
            Width           =   2475
         End
         Begin MSComDlg.CommonDialog cmdlgGeoMap 
            Left            =   3900
            Top             =   540
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image ImageOptions 
            Height          =   300
            Left            =   660
            Picture         =   "GDOptionsfrm.frx":147B
            Stretch         =   -1  'True
            Top             =   1360
            Width           =   600
         End
         Begin VB.Shape ShapeB0 
            Height          =   320
            Left            =   660
            Top             =   1350
            Width           =   620
         End
         Begin VB.Label Label9 
            Caption         =   "Map's Pixel Width"
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
            Left            =   1560
            TabIndex        =   21
            Top             =   1260
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Map's Pixel Height"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1560
            TabIndex        =   20
            Top             =   1500
            Width           =   1515
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   540
            Top             =   1260
            Width           =   855
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   300
            X2              =   480
            Y1              =   1280
            Y2              =   1280
         End
         Begin VB.Line Line3 
            Index           =   0
            X1              =   300
            X2              =   480
            Y1              =   1740
            Y2              =   1740
         End
         Begin VB.Line Line4 
            Index           =   0
            X1              =   540
            X2              =   540
            Y1              =   1080
            Y2              =   1320
         End
         Begin VB.Line Line5 
            Index           =   0
            X1              =   1380
            X2              =   1380
            Y1              =   1080
            Y2              =   1320
         End
         Begin VB.Label Label6 
            Caption         =   "Y"
            Height          =   195
            Left            =   340
            TabIndex        =   19
            Top             =   1420
            Width           =   135
         End
         Begin VB.Label Label8 
            Caption         =   "X"
            Height          =   195
            Left            =   940
            TabIndex        =   18
            Top             =   1050
            Width           =   195
         End
         Begin VB.Line Line6 
            Index           =   0
            X1              =   540
            X2              =   840
            Y1              =   1150
            Y2              =   1150
         End
         Begin VB.Line Line7 
            Index           =   0
            X1              =   1080
            X2              =   1380
            Y1              =   1150
            Y2              =   1150
         End
         Begin VB.Line Line8 
            Index           =   0
            X1              =   390
            X2              =   390
            Y1              =   1650
            Y2              =   1740
         End
         Begin VB.Line Line9 
            Index           =   0
            X1              =   390
            X2              =   390
            Y1              =   1320
            Y2              =   1440
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Coordinate System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   4935
         Begin VB.CheckBox chkFathoms 
            BackColor       =   &H00808080&
            Caption         =   "Fathoms"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3960
            TabIndex        =   173
            ToolTipText     =   "Click if elevation units of map is in fathoms"
            Top             =   580
            Width           =   975
         End
         Begin VB.CheckBox chkfeet 
            BackColor       =   &H00808080&
            Caption         =   "Feet"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3240
            TabIndex        =   172
            ToolTipText     =   "Click if elevation units of map is in feet"
            Top             =   600
            Width           =   615
         End
         Begin VB.CheckBox chkLatLon 
            BackColor       =   &H00808080&
            Caption         =   "degrees lat. && lon."
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3240
            TabIndex        =   79
            ToolTipText     =   "Clheck the box if using degrees latitude and longitude"
            Top             =   120
            Width           =   1575
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   315
            Left            =   2160
            TabIndex        =   2
            ToolTipText     =   "Y Coordinate Label"
            Top             =   540
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   315
            Left            =   2160
            TabIndex        =   1
            ToolTipText     =   "X coordinate label"
            Top             =   180
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label20 
            BackColor       =   &H00808080&
            Caption         =   "Coordinate System"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   60
            TabIndex        =   33
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "X Coordinate Label"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            Caption         =   "Y Coordinate Label"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   600
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "GDOptionsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DefaultMap As Boolean, ier%
Dim AddedMaps As Boolean

Private Sub chkFathoms_Click()
   MapUnits = 1.8288002
End Sub

Private Sub chkfeet_Click()
   MapUnits = 0.30479999798832
End Sub

Private Sub chkNewDTM_Click()
   If chkNewDTM.value = vbChecked Then
      UseNewDTM% = 1
   Else
      UseNewDTM% = 0
      End If
End Sub

Private Sub chkSave_xyz_Click()
   'Turn off on on automatic saving of settings
   If chkSave_xyz.value = vbChecked Then
      Save_xyz% = 1
   Else
      Save_xyz% = 0
      End If

End Sub

Private Sub chkCenterClick_Click()
  If chkCenterClick.value = vbChecked Then
     PointCenterClick = 1
  Else
     PointCenterClick = 0
     End If
End Sub

Private Sub chkLatLon_Click()
  If chkLatLon.value = vbChecked Then
     With GDOptionsfrm
        .MaskEdBox1.Text = "lon."
        .MaskEdBox2.Text = "lat."
        .cmdULGridXchange.Enabled = True
        .cmdULGridYchange.Enabled = True
        .cmdULXconvert.Enabled = True
        .cmdULYconvert.Enabled = True
        .cmdLRGridXchange.Enabled = True
        .cmdLRGridYchange.Enabled = True
        .cmdLRXconvert.Enabled = True
        .cmdLRYconvert.Enabled = True
     End With
  Else
     With GDOptionsfrm
        .MaskEdBox1.Text = "    "
        .MaskEdBox2.Text = "    "
        .cmdULGridXchange.Enabled = False
        .cmdULGridYchange.Enabled = False
        .cmdULXconvert.Enabled = False
        .cmdULYconvert.Enabled = False
        .cmdLRGridXchange.Enabled = False
        .cmdLRGridYchange.Enabled = False
        .cmdLRXconvert.Enabled = False
        .cmdLRYconvert.Enabled = False
     End With
     End If
End Sub

Private Sub chkAutoRedraw_Click()
   If chkAutoRedraw.value = vbChecked And IgnoreAutoRedrawError% = 0 Then
      MsgBox "Warning--in the case the the screen memory is insufficient to handle a map," & vbLf & _
             "then the following graphic features will be affected:" & vbLf & vbLf & _
             "   1) The blinking cursor (will not appear)." & vbLf & _
             "   2) Drag box during dragging operations (will not appear)." & vbLf & _
             "   3) The magnification window (might not function properly).", vbInformation + vbOKOnly, "MapDigitizer Settings"
      End If
End Sub

Private Sub chkKML_Click()

   With GDOptionsfrm
   
      If chkKML.value = vbChecked Then
         .txtWellIcon.Enabled = True
         .txtOutCropIcon.Enabled = True
         .lblWellIcon(1).Enabled = True
         .lblOutCropIcon(0).Enabled = True
      Else
         .txtWellIcon.Enabled = False
         .txtOutCropIcon.Enabled = False
         .lblWellIcon(1).Enabled = False
         .lblOutCropIcon(0).Enabled = False
         End If
      
   End With

End Sub

Private Sub chkErrorMessage_Click()
   'Turn off on on reporting on path errors
   'This controls if the subroutine ShowError is called.
   If chkErrorMessage.value = vbChecked Then
      ReportPaths& = 1
   Else
      ReportPaths& = 0
      End If
End Sub

Private Sub chkRainbow_Click()
   If chkRainbow.value = vbChecked Then
      LineElevColors& = 1
   Else
      LineElevColors& = 0
      End If
End Sub

Private Sub chkSave_Click()
   'Turn off on on automatic saving of settings
   If chkSave.value = vbChecked Then
      SaveClose% = 1
   Else
      resp = MsgBox("Warning, any changes will not be recorded automatically!" & vbLf & _
                  "Do you wish to continue?", vbYesNoCancel + vbExclamation, "MapDigitizer")
      If resp = vbYes Then
         SaveClose% = 0
      Else
         SaveClose% = 1
         chkSave.value = vbChecked
         End If
      End If

End Sub

'Private Sub chkSquare1_Click()
'   txtULGridX = txtLRGeoX
'   txtULGridY = txtULGeoY
'End Sub

Private Sub chkTif_Click()
   If chkTif.value = vbChecked Then
      frmTifViewer.Enabled = True
      cmdBrowsePath.Enabled = True
      cmdBrowseViewer.Enabled = True
      txttifpath.Enabled = True
      txtTifViewer.Enabled = True
      txtCommandLine.Enabled = True
      cmdSave.Enabled = True
      cmdRestoreDefault.Enabled = True
      lbltifFiles.Enabled = True
      lblTifViewer.Enabled = True
      lblCommandLine.Enabled = True
   Else
      frmTifViewer.Enabled = False
      cmdBrowsePath.Enabled = False
      cmdBrowseViewer.Enabled = False
      txttifpath.Enabled = False
      txtTifViewer.Enabled = False
      txtCommandLine.Enabled = False
      cmdSave.Enabled = False
      cmdRestoreDefault.Enabled = False
      lbltifFiles.Enabled = False
      lblTifViewer.Enabled = False
      lblCommandLine.Enabled = False
      End If
End Sub

Private Sub cmbMaps_Click()
   'load up new map parameters
   
   On Error GoTo errhand
   
    If magvis And cmbMaps.Text <> picnam$ Then 'close magnify window first
       MsgBox "Can't close or switch maps until you close the magnification window!", _
              vbOKOnly + vbExclamation, "MapDigitizer"
       Exit Sub
       End If
   
50:
      If Dir(MapParms(0, cmbMaps.ListIndex)) <> sEmpty Then
      
         DigiGDIfailed = False
         RSMethod0 = False
         
         MaskEdBox1.Text = MapParms(1, cmbMaps.ListIndex)
         MaskEdBox2.Text = MapParms(2, cmbMaps.ListIndex)
         txtULGeoX = MapParms(3, cmbMaps.ListIndex)
         txtULGeoY = MapParms(4, cmbMaps.ListIndex)
         txtLRGeoX = MapParms(5, cmbMaps.ListIndex)
         txtLRGeoY = MapParms(6, cmbMaps.ListIndex)
         txtPixWidth = MapParms(7, cmbMaps.ListIndex)
         txtPixHeight = MapParms(8, cmbMaps.ListIndex)
         txtGridX = MapParms(9, cmbMaps.ListIndex)
         txtGridY = MapParms(10, cmbMaps.ListIndex)
         txtULPixX = MapParms(11, cmbMaps.ListIndex)
         txtULPixY = MapParms(12, cmbMaps.ListIndex)
         txtLRPixX = MapParms(13, cmbMaps.ListIndex)
         txtLRPixY = MapParms(14, cmbMaps.ListIndex)
         txtLRGridX = MapParms(15, cmbMaps.ListIndex)
         txtLRGridY = MapParms(16, cmbMaps.ListIndex)
         txtULGridX = MapParms(17, cmbMaps.ListIndex)
         txtULGridY = MapParms(18, cmbMaps.ListIndex)
         MapUnits = val(MapParms(19, cmbMaps.ListIndex))
         If MapUnits = 0 Then MapUnits = 1#
         
      Else
         'try adding the app.path and checking again
         testpath$ = App.Path & "\" & MapParms(0, cmbMaps.ListIndex)
         If Dir(testpath$) <> sEmpty Then
            MapParms(0, cmbMaps.ListIndex) = testpath$
            GoTo 50
            End If
            
         MsgBox "Can't find the requested map at the recorded path!" & vbLf & _
                "Use the browse button to find it", vbExclamation + vbOKOnly, "MapDigitizer"
         cmbMaps.ListIndex = 0 'restore original values
         End If
         
      If MapParms(0, cmbMaps.ListIndex) <> picnam$ Or DefaultMap Then
         'unload the last map, and load this map if desired
         picnam$ = MapParms(0, cmbMaps.ListIndex)
         lblX = MaskEdBox1.Text
         LblY = MaskEdBox2.Text
         ULGeoX = val(MapParms(3, cmbMaps.ListIndex))
         ULGeoY = val(MapParms(4, cmbMaps.ListIndex))
         LRGeoX = val(MapParms(5, cmbMaps.ListIndex))
         LRGeoY = val(MapParms(6, cmbMaps.ListIndex))
         pixwi = val(MapParms(7, cmbMaps.ListIndex))
         pixhi = val(MapParms(8, cmbMaps.ListIndex))
         NX_CALDAT = val(MapParms(9, cmbMaps.ListIndex))
         NY_CALDAT = val(MapParms(10, cmbMaps.ListIndex))
         ULPixX = val(MapParms(11, cmbMaps.ListIndex))
         ULPixY = val(MapParms(12, cmbMaps.ListIndex))
         LRPixX = val(MapParms(13, cmbMaps.ListIndex))
         LRPixY = val(MapParms(14, cmbMaps.ListIndex))
         LRGridX = val(MapParms(15, cmbMaps.ListIndex))
         LRGridY = val(MapParms(16, cmbMaps.ListIndex))
         ULGridX = val(MapParms(17, cmbMaps.ListIndex))
         ULGridY = val(MapParms(18, cmbMaps.ListIndex))
         MapUnits = val(MapParms(19, cmbMaps.ListIndex))
         
         picnam0$ = picnam$
         x10 = ULGeoX
         y10 = ULGeoY
         x20 = LRGeoX
         y20 = LRGeoY
         pixwi0 = pixwi
         pixhi0 = pixhi
         If buttonstate&(2) = 1 Or buttonstate&(18) = 1 Then
            'Geo map visible, so unload old one and load new one
            'unload old one
            GDMDIform.mnuMapInput_Click
            'load new Geo map
            GDMDIform.mnuMapInput_Click
         Else 'don't do anything
            End If
         BringWindowToTop (GDOptionsfrm.hwnd)
         DefaultMap = False
         End If
         
         chkLatLon.value = vbUnchecked
         
         If lblX = "lon." And LblY = "lat." Then
            GDOptionsfrm.chkLatLon.value = vbChecked
         Else
            GDOptionsfrm.MaskEdBox1.Text = lblX
            GDOptionsfrm.MaskEdBox2.Text = LblY
            End If
            
    If MapUnits = 0 Or MapUnits = 1 Then
       chkfeet.value = vbUnchecked
       chkFathoms.value = vbUnchecked
       MapUnits = 1#
       GDMDIform.Text3.ToolTipText = "Elevation (meters)"
       GDMDIform.Text7.ToolTipText = "Elevation (meters) at center of clicked point"
    ElseIf MapUnits = 0.30479999798832 Then
       chkFathoms.value = vbUnchecked
       chkfeet.value = vbChecked
       GDMDIform.Text3.ToolTipText = "Elevation (feet)"
       GDMDIform.Text7.ToolTipText = "Elevation (feet) at center of clicked point"
    ElseIf MapUnits = 1.8288002 Then
       chkfeet.value = vbUnchecked
       chkFathoms.value = vbChecked
       GDMDIform.Text3.ToolTipText = "Elevation (fathoms)"
       GDMDIform.Text7.ToolTipText = "Elevation (fathoms) at center of clicked point"
       End If
            
        'now open file
            
      
Exit Sub

errhand:
   DefaultMap = False
   Screen.MousePointer = vbDefault
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          sEmpty, vbCritical + vbOKOnly, "MapDigitizer"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLRGridXconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLRGridXconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtLRGridX = ConvertDegToNumber(txtLRGridX)
  Else
  
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLRGridYconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLRGridYconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtLRGridY = ConvertDegToNumber(txtLRGridY)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLRGridYchange_Click
' Author    : Dr-John-K-Hall
' Date      : 3/5/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLRGridYchange_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtLRGridY = ConvertDegToNumber(txtLRGridY)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLRXconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLRXconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtLRGeoX = ConvertDegToNumber(txtLRGeoX)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLRYconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLRYconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtLRGeoY = ConvertDegToNumber(txtLRGeoY)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdULGridXconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdULGridXconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtULGridX = ConvertDegToNumber(txtULGridX)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdULGridYconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdULGridYconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtULGridY = ConvertDegToNumber(txtULGridY)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!", vbExclamation, "Coordinate Conversion Error")
     End If
End Sub

Private Sub cmdPasteULPixX_Click()
   txtULPixX = GDMDIform.Text5
End Sub

Private Sub cmdPasteLRPixX_Click()
   txtLRPixX = GDMDIform.Text5
End Sub
Private Sub cmdPasteLRPixY_Click()
   txtLRPixY = GDMDIform.Text6
End Sub

Private Sub cmdSaveNewDTM_Click()

      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Output As #infonum&
      Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
      Write #infonum&, dirNewDTM
      Write #infonum&, val(txtEraserBrushSize)
      Write #infonum&, NEDdir
      Write #infonum&, dtmdir
      Write #infonum&, ChainCodeMethod
      Write #infonum&, val(txtDistContour), val(txtDistLines), val(txtSensitivity), val(cmbContour.Text) ' arcdir, mxddir
      Write #infonum&, PointCenterClick
      Write #infonum&, picnam$
      Write #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
      Write #infonum&, ReportPaths&, val(txtDistPixelSearch), numMaxHighlight&, Save_xyz%
      Write #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Write #infonum&, IgnoreAutoRedrawError%
      Write #infonum&, UseNewDTM%, nOtherCheck%
      Write #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
      Write #infonum&, val(txtGridX), val(txtGridY) 'NX_CALDAT, NY_CALDAT
      Write #infonum&, RSMethod0, RSMethod1, RSMethod2
      Write #infonum&, val(txtULPixX), val(txtULPixY), val(txtLRPixX), val(txtLRPixY), val(JustConvertDegToNumber(txtLRGridX)), val(JustConvertDegToNumber(txtLRGridY)), val(JustConvertDegToNumber(txtULGridX)), val(JustConvertDegToNumber(txtULGridY))
      Write #infonum&, val(txtDTMitmx), val(txtDTMitmy), val(txtDTMlon), val(txtDTMlat), val(txtAzi), val(txtStepAzi), val(txtaprn), HeightPrecision, val(txtCustom)
      Close #infonum&
      
      XStepITM = txtDTMitmx
      YStepITM = txtDTMitmy
      XStepDTM = txtDTMlon
      YStepDTM = txtDTMlat
      HalfAzi = txtAzi
      StepAzi = txtStepAzi
      Apprn = txtaprn
      DigiConvertToMeters = val(txtCustom)
      
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdULXconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdULXconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtULGeoX = ConvertDegToNumber(txtULGeoX)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdULYconvert_Click
' Author    : Dr-John-K-Hall
' Date      : 3/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdULYconvert_Click()
  If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
     txtULGeoY = ConvertDegToNumber(txtULGeoY)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
     
End Sub

Private Sub cmdAddMap_Click()
    'add new map and parameters to gdbmap.sav map info
  
     On Error GoTo errhand2
  
     cmdlgGeoMap.CancelError = True
     cmdlgGeoMap.Filter = "Bmp files (*.bmp)|*.bmp|" & _
                          "Gif files (*.gif)|*.gif|" & _
                          "JPG files (*.jpg)|*.jpg|" & _
                          "JPEG files (*.jpeg)|(*.jpeg)|" & _
                          "Mega files (*.wmf)|(*.wmf)|" & _
                          "Icon files (*.ico)|(*.ico)|" & _
                          "All files (*.*)|*.*"
     'specify default filter
     cmdlgGeoMap.FilterIndex = 1
     cmdlgGeoMap.FileName = "*.bmp" 'direct$ + "\*.bmp"
     cmdlgGeoMap.ShowOpen
        
     'check that selected map is not a tif file, since it is not supported
     If InStr(cmdlgGeoMap.FileName, ".tif") <> 0 Then
        MsgBox "Tif files are not supported!", vbExclamation + vbOKOnly, "MapDigitizer"
        Exit Sub
        End If
    
    'before displaying this name, make sure last info is recorded
    
    On Error GoTo errhand
    
    If Dir(cmdlgGeoMap.FileName) = sEmpty Then
       MsgBox "Map not found!", vbExclamation + vbOKOnly, "Map not found"
       Exit Sub
       End If
    
    If picnam$ = sEmpty Then GoTo cbgm500
       
    found% = 0
    For i% = 0 To UBound(MapParms, 2)
       If MapParms(0, i%) = cmdlgGeoMap.FileName Then
          found% = 1
          Exit For
          End If
    Next i%
    If found% = 0 And picnam$ <> sEmpty Then
    
      txtULGeoX = sEmpty
      txtULGeoY = sEmpty
      txtLRGeoX = sEmpty
      txtLRGeoY = sEmpty
      txtPixWidth = sEmpty
      txtPixHeight = sEmpty
      txtGridX = sEmpty
      txtGridY = sEmpty
      txtULPixX = sEmpty
      txtULPixY = sEmpty
      txtLRPixX = sEmpty
      txtLRPixY = sEmpty
      txtLRGridX = sEmpty
      txtLRGridY = sEmpty
      txtULGridX = sEmpty
      txtULGridY = sEmpty
      txtCustom = 1#
    
       nn& = cmbMaps.ListCount
'       ReDim Preserve MapParms(18, nn&)
'       MapParms(0, nn&) = cmbMaps.Text
'       MapParms(1, nn&) = MaskEdBox1.Text
'       MapParms(2, nn&) = MaskEdBox2.Text
'       MapParms(3, nn&) = JustConvertDegToNumber(txtULGeoX)
'       MapParms(4, nn&) = JustConvertDegToNumber(txtULGeoY)
'       MapParms(5, nn&) = JustConvertDegToNumber(txtLRGeoX)
'       MapParms(6, nn&) = ConvertDegToNumber(txtLRGeoY)
'       MapParms(7, nn&) = Trim$(txtPixWidth)
'       MapParms(8, nn&) = Trim$(txtPixHeight)
'       MapParms(9, nn&) = JustConvertDegToNumber(txtGridX)
'       MapParms(10, nn&) = JustConvertDegToNumber(txtGridY)
'       MapParms(11, nn&) = JustConvertDegToNumber(txtULPixX)
'       MapParms(12, nn&) = JustConvertDegToNumber(txtULPixY)
'       MapParms(13, nn&) = JustConvertDegToNumber(txtLRPixX)
'       MapParms(14, nn&) = JustConvertDegToNumber(txtLRPixY)
'       MapParms(15, nn&) = JustConvertDegToNumber(txtLRGridX)
'       MapParms(16, nn&) = JustConvertDegToNumber(txtLRGridY)
'       MapParms(17, nn&) = JustConvertDegToNumber(txtULGridX)
'       MapParms(18, nn&) = JustConvertDegToNumber(txtULGridY)
'       cmbMaps.ListIndex = cmbMaps.ListCount - 1
       End If
cbgm500:
  Screen.MousePointer = vbHourglass
  
  cmbMaps.Text = cmdlgGeoMap.FileName
  cmbMaps.AddItem cmdlgGeoMap.FileName
  If cmbMaps.ListCount = 0 Then GoTo cbgm600
     
  found% = 0 'check if map's parameters haven't already been added to map parameter array, mapParams
  For i% = 0 To UBound(MapParms, 2)
     If cmbMaps.Text = MapParms(0, i%) Then
        found% = 1
        nn& = i% 'just load it
        Exit For
        End If
  Next i%
  
cbgm600:
  If found% = 0 Then
    'determine its pixel size
    cmdPixelSize_Click
    If ier% < 0 Then 'error detected (probably in file format)
       'remove this pictures's name
       cmbMaps.Text = sEmpty
       ier% = 0
       Exit Sub
       End If
    
    'Add it to combo list, and save its temporary
    'parameters to the array MapParams
'    cmbMaps.AddItem cmbMaps.Text
'    nn& = cmbMaps.ListCount - 1
    
    'set default coordinate system
    MaskEdBox1.Text = sEmpty
    MaskEdBox2.Text = sEmpty
    
    ReDim Preserve MapParms(19, nn&)
    MapParms(0, nn&) = cmbMaps.Text
'    MapParms(1, nn&) = MaskEdBox1.Text
'    MapParms(2, nn&) = MaskEdBox2.Text
'    MapParms(3, nn&) = JustConvertDegToNumber(txtULGeoX)
'    MapParms(4, nn&) = JustConvertDegToNumber(txtULGeoY)
'    MapParms(5, nn&) = JustConvertDegToNumber(txtLRGeoX)
'    MapParms(6, nn&) = JustConvertDegToNumber(txtLRGeoY)
'    MapParms(7, nn&) = JustConvertDegToNumber(txtPixWidth)
'    MapParms(8, nn&) = JustConvertDegToNumber(txtPixHeight)
'    MapParms(9, nn&) = Trim$(txtGridX)
'    MapParms(10, nn&) = Trim$(txtGridY)
'    MapParms(11, nn&) = JustConvertDegToNumber(txtULPixX)
'    MapParms(12, nn&) = JustConvertDegToNumber(txtULPixY)
'    MapParms(13, nn&) = JustConvertDegToNumber(txtLRPixX)
'    MapParms(14, nn&) = JustConvertDegToNumber(txtLRPixY)
'    MapParms(15, nn&) = JustConvertDegToNumber(txtLRGridX)
'    MapParms(16, nn&) = JustConvertDegToNumber(txtLRGridY)
'    MapParms(17, nn&) = JustConvertDegToNumber(txtULGridX)
'    MapParms(18, nn&) = JustConvertDegToNumber(txtULGridY)
    
    'give warning that parameters must be edited
    MsgBox "You have successfully added a new map!" & vbLf & vbLf & _
           "Edit the values of the Coordinate System Labels and Coordinate Boundaries." & vbLf & _
           "Afterwards, save this the settings with the ""Save maps"" button.", _
           vbInformation + vbOKOnly, "MapDigitizer"
    
    AddedMaps = True 'remember that a new map was added
    
    End If
  
'  'make the chosen item the current map
'  cmbMaps.ListIndex = nn&
'  'load the map if map form is visible
'  cmbMaps_Click
  
  Screen.MousePointer = vbDefault

Exit Sub

errhand:
'   Close
   If infonum& > 0 Then
      Close #infonum&
      infonum& = 0
      End If
   Screen.MousePointer = vbDefault
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          sEmpty, vbCritical + vbOKOnly, "MapDigitizer"
   Exit Sub
   
errhand2:

End Sub

Private Sub cmdBrowseASTER_Click()
  On Error GoTo c4error
  cmdlgASTER.CancelError = True
  cmdlgASTER.Filter = "ASTER files (*.bil)|*.bil|"
  cmdlgASTER.FilterIndex = 1
  cmdlgASTER.FileName = Replace(ASTERdir + "\N31E035.bil", "\\", "\")
  cmdlgASTER.ShowOpen
  filn$ = cmdlgASTER.FileName
  pos& = InStr(LCase(filn$), "\n31e035.bil")
  txtAster = Mid$(filn$, 1, pos& - 1)
  If Dir(filn$) <> sEmpty Then
     ASTERdir = txtAster
     optAster.Enabled = True
     End If

c4error:
  Exit Sub

End Sub

Private Sub cmdBrowseGoogle_Click()
  On Error GoTo c5error
  cmdlgArc.CancelError = True
  cmdlgArc.Filter = "googleearth.exe (*.exe)|*.exe|"
  cmdlgArc.FilterIndex = 1
  cmdlgArc.FileName = Replace(googledir + "\googleearth.exe", "\\", "\")
  cmdlgArc.ShowOpen
  filn$ = cmdlgArc.FileName
  pos& = InStr(LCase(filn$), "\googleearth.exe")
  txtGoogle = Mid$(filn$, 1, pos& - 1)

c5error:
  Exit Sub

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdBrowseKML_Click
' DateTime  : 1/7/2009 18:02
' Author    : Chaim Keller
' Purpose   : Common Dialog controls that picks directories instead of files
'             taken from the site "Microsoft Help and Support, article:
'             "How To Select a Directory Without the Common Dialog Control"
'             URL = http://support.microsoft.com/kb/179497
'---------------------------------------------------------------------------------------
'
Private Sub cmdBrowseKML_Click()

   'Opens a Treeview control that displays the directories in a computer

      Dim lpIDList As Long
      Dim sBuffer As String
      Dim szTitle As String
      Dim tBrowseInfo As BrowseInfo

      szTitle = "Choose a directory to store kml files"
      With tBrowseInfo
         .hWndOwner = Me.hwnd
         .lpszTitle = lstrcat(szTitle, "")
         .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
      End With

      lpIDList = SHBrowseForFolder(tBrowseInfo)

      If (lpIDList) Then
         sBuffer = Space(MAX_PATH)
         SHGetPathFromIDList lpIDList, sBuffer
         sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
         txtkml.Text = sBuffer
      End If
    
Exit Sub

' taken from the site "Visual Basic Explorer", article:
' "Using Common Dialog to Select Directories Instead of Files"
' URL = http://www.vbexplorer.com/VBExplorer/tips/src27.asp

'    With cmdlgArc
'        .Flags = cdlOFNPathMustExist
'        .Flags = .Flags Or cdlOFNHideReadOnly
'        .Flags = .Flags Or cdlOFNNoChangeDir
'        .Flags = .Flags Or cdlOFNExplorer
'        .Flags = .Flags Or cdlOFNNoValidate
'        .FileName = "*.*"
'    End With
'
'
'    Dim x As Integer
'    '-- Cheap way to use the common dialog box as a directory-picker
'    x = 3
'
'    cmdlgArc.CancelError = True      'do not terminate on error
'
'    On Error Resume Next        'I will hande errors
'
'    cmdlgArc.Action = 1              'Present "open" dialog
'
'    '-- If FileTitle is null, user did not override the default (*.*)
'    If cmdlgArc.FileTitle <> "" Then x = Len(cmdlgArc.FileTitle)
'
'    If Err = 0 Then
'        ChDrive cmdlgArc.FileName
'        txtkml.text = Left(cmdlgArc.FileName, Len(cmdlgArc.FileName) - x)
'    Else
'     '-- User pressed "Cancel"
'    End If

End Sub



Private Sub cmdBrowseDTM_Click()
  On Error GoTo c4error
  cmdlgDTM.CancelError = True
  cmdlgDTM.Filter = "NED hgt files (*.hgt)|*.hgt|"
  cmdlgDTM.FilterIndex = 1
  cmdlgDTM.FileName = Replace(NEDdir + "\Z000000.hgt", "\\", "\")
  cmdlgDTM.ShowOpen
  filn$ = cmdlgDTM.FileName
  pos& = InStr(LCase(filn$), "\z000000.hgt")
  txtdtm = Mid$(filn$, 1, pos& - 1)
  If Dir(filn$) <> sEmpty Then
     NEDdir = txtdtm
     optDTM.Enabled = True
     End If
     
c4error:
  Exit Sub

End Sub


Private Sub cmdDefaults_Click()
   'restore default map parameters
   
   On Error GoTo errhand:
   
      If Dir(direct$ & "\gdbinfo.sav") = sEmpty Then
         resp = MsgBox("Sorry, no map information found!" & vbLf & _
                "Do you want to load the default parameter values?", _
                vbYesNoCancel + vbExclamation, "MapDigitizer")
         If resp = vbYes Then
            GoTo def500
         Else
            Exit Sub
            End If
         End If
      
      DefaultMap = True 'flag for loading the default map
      
        If infonum& > 0 Then
           Close #infonum&
           infonum& = 0
           End If

      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Input As #infonum&
      Input #infonum&, doclin$
      Input #infonum&, doclin$
      Input #infonum&, MinDigiEraserBrushSize
      Input #infonum&, NEDdirtmp
      Input #infonum&, dtmdirtmp
      Input #infonum&, ChainCodeMethod
      Input #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
      Input #infonum&, PointCenterClick
      Input #infonum&, picnam$
      Input #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
      Input #infonum&, ReportPaths&, DigiSearchRegion, numMaxHighlight&, Save_xyz%
      Input #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Input #infonum&, IgnoreAutoRedrawError%
      Input #infonum&, UseNewDTM%, nOtherCheck%
      Input #infonum&, googledirtmp, URL_OutCroptmp, URL_Welltmp, kmldirtmp, ASTERdirtmp, DTMtypetmp
      Input #infonum&, NX_CALDAT, NY_CALDAT
      Input #infonum&, RSMethod0, RSMethod1, RSMethod2
      Input #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
      Input #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
      Close #infonum&
      
      infonum& = 0
     
      txtDistContour = numDistContour
      txtDistLines = numDistLines
      txtSensitivity = numSensitivity
      cmbContour.Text = str$(numContours)
      txtDistPixelSearch = DigiSearchRegion
      txtEraserBrushSize = MinDigiEraserBrushSize
      txtITMx = XStepITM
      txtITMy = YStepITM
      txtDTMlon = XStepDTM
      txtDTMlat = YStepDTM
      
      If ChainCodeMethod = 0 Then
         optFreeman.value = True
      ElseIf ChainCodeMethod = 1 Then
         optBug.value = True
         End If
         
      If PointCenterClick = 1 Then
         chkCenterClick.value = vbChecked
         End If
         
      If LineElevColors& = 1 Then
         GDOptionsfrm.chkRainbow = vbChecked
         End If
      
def500:
      'default map and map boundaries and geo map pixel sizes
        
      If lblX = sEmpty Then lblX = "lon." 'ITMx"
      If LblY = sEmpty Then LblY = "lat." '"ITMy"
      
      If lblX = "lon." And LblY = "lat." Then GDOptionsfrm.chkLatLon.value = vbChecked
    
      If IsNull(ULGeoX) Then ULGeoX = 0 '80000#
      If IsNull(ULGeoY) Then ULGeoY = 0 '1300000#
      If IsNull(LRGeoX) Then LRGeoX = 0 '240000#
      If IsNull(LRGeoY) Then LRGeoY = 0 '880000#
'      If X1 = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'         X1 = 80000#
'         ulgeoy = 1300000#
'         lrgeox = 240000#
'         lrgeoy = 880000#
'         End If
'      If IsNull(pixwi) Then pixwi = 1268
'      If IsNull(pixhi) Then pixhi = 3338
'      If pixwi = 0 Then pixwi = 1268
'      If pixhi = 0 Then pixhi = 3338
      
      If IsNull(NX_CALDAT) Then NX_CALDAT = 0
      If IsNull(NY_CALDAT) Then NY_CALDAT = 0
               
      found% = 0
      For i% = 1 To cmbMaps.ListCount
         If cmbMaps.List(i% - 1) = picnam$ Then
            If txtULGeoX = "0" And txtULGeoY = "0" And txtLRGeoX = "0" And txtLRGeoY = "0" Then
               txtULGeoX = ULGeoX
               txtULGeoY = ULGeoY
               txtLRGeoX = LRGeoX
               txtLRGeoY = LRGeoY
               End If
            
            cmbMaps.ListIndex = i% - 1 'just go to it
            found% = 1
            Exit For
            End If
      Next i%
      If found% = 0 And picnam$ <> sEmpty Then
         cmbMaps.AddItem picnam$
         nn& = cmbMaps.ListIndex
         ReDim Preserve MapParms(19, nn&)
         MapParms(0, nn&) = picnam$
         MapParms(1, nn&) = lblX
         MapParms(2, nn&) = LblY
         MapParms(3, nn&) = str$(ULGeoX)
         MapParms(4, nn&) = str$(ULGeoY)
         MapParms(5, nn&) = str$(LRGeoX)
         MapParms(6, nn&) = str$(LRGeoY)
         MapParms(7, nn&) = str$(pixwi)
         MapParms(8, nn&) = str$(pixhi)
         MapParms(9, nn&) = str$(NX_CALDAT)
         MapParms(10, nn&) = str$(NY_CALDAT)
         MapParms(11, nn&) = str$(ULPixX)
         MapParms(12, nn&) = str$(ULPixY)
         MapParms(13, nn&) = str$(LRPixX)
         MapParms(14, nn&) = str$(LRPixY)
         MapParms(15, nn&) = str$(LRGridX)
         MapParms(16, nn&) = str$(LRGridY)
         MapParms(17, nn&) = str$(ULGridX)
         MapParms(18, nn&) = str$(ULGridY)
         If MapUnits = 0 Then MapUnits = 1#
         MapParms(19, nn&) = str$(MapUnits)
         
         cmbMaps.ListIndex = cmbMaps.ListCount - 1
         End If
         
     If picnam$ = sEmpty Or Dir(picnam$) = sEmpty Then
        MaskEdBox1.Text = lblX
        MaskEdBox2.Text = LblY
        
        txtULGeoX.Text = str(ULGeoX)
        txtULGeoY.Text = str(ULGeoY)
        txtLRGeoX.Text = str(LRGeoX)
        txtPixWidth.Text = str(pixwi)
        txtLRGeoY.Text = str(LRGeoY)
        txtPixHeight.Text = str(pixhi)
        MsgBox "You need to enter the path of the default map: new5.bmp" & vbLf & _
               "Use the browse button", vbInformation + vbOKOnly, "MapDigitizer"
        End If
        
Exit Sub

errhand:
   DefaultMap = False
   If infonum& > 0 Then
      Close #infonum&
      infonum& = 0
      End If
   Screen.MousePointer = vbDefault
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          sEmpty, vbCritical + vbOKOnly, "MapDigitizer"

End Sub

'Private Sub cmdGeoOutColor_Click()
'   On Error GoTo errhand
'   cmdlgColor.ShowColor
'   PointColor& = cmdlgColor.color
'   shpPoints.FillColor = PointColor&
'   Exit Sub
'errhand:
'End Sub
'
'Private Sub cmdGeoWellColor_Click()
'   On Error GoTo errhand
'   cmdlgColor.ShowColor
'   LineColor& = cmdlgColor.color
'   shpLines.FillColor = LineColor&
'   Exit Sub
'errhand:
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdPassword_Click
' DateTime  : 12/18/2008 21:09
' Author    : Chaim Keller
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdPassword_Click()

   On Error GoTo cmdPassword_Click_Error
   
   If Dir(accdir & "\MSaccess.exe") = gsEmpty And Dir(Trim(txtacc.Text)) = gsEmpty Then
      MsgBox "You must first provide a valid path to MS Access!", vbOKOnly + vbExclamation, "MS Access Toolbar Activation"
      Exit Sub
      End If

   CurUserPassTrial$ = mskEdPass1.Text & mskEdPass2.Text & mskEdPass3.Text
   
   If Len(CurUserPassTrial$) > 12 Then
      'too many numbers, try removing any "_" in case they were
      'inadvertently added (happens if repeat the entering of a password)
      CurUserPassTrial$ = Replace(CurUserPassTrial$, "_", sEmpty)
      End If
   
   If Not CheckPassword(CurUserPassTrial$) Then
      
      mskEdPass1.Mask = sEmpty
      mskEdPass2.Mask = sEmpty
      mskEdPass3.Mask = sEmpty
      MsgBox "Incorrect Password!" & vbLf & vbLf & _
             "See the system manager."
      Exit Sub
   
   Else
      If Not ActivatedVersion Then 'program is unregistered until now, so thank you, enable toolbars, menus
         
         Me.Visible = False
         
         MsgBox "Activation is complete!" & vbLf & vbLf & _
                "Be sure to save your password in a safe place!" & vbLf & _
                "In case it is lost, it can be retrieved from the registry at:" & vbLf & vbLf & _
                "HKEY_CURRENT_USER/Software/VB and VBA Program Settings/MapDigitizer/Security", _
                vbInformation + vbOKOnly, "Soid Ha'ibur Thanks"
      
         'save the correct password in the registry
         SaveSetting App.Title, "Security", "Password", CurUserPassTrial$
      
         ActivatedVersion = True
    
         Unload Me
         End If
       
      End If

'and check if inputed digits match
'if match, then set DemoVersion = false and restart program initialization
'else give message that password is incorrect


   On Error GoTo 0
   Exit Sub

cmdPassword_Click_Error:
   
   Screen.MousePointer = vbDefault
   
   If Err.Number = 5 Then 'can't write to registry
      Resume Next
      End If
      
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPassword_Click of Form GDOptionsfrm"

End Sub

Private Sub cmdPixelSize_Click()
   On Error GoTo errhand
   
   myfile = Dir(cmbMaps.Text)
   If myfile = sEmpty Then
      response = MsgBox("Can't find map picture file!", vbCritical + vbOKOnly, "MapDigitizer")
      cmbMaps.Text = sEmpty
      picnam$ = sEmpty
      Exit Sub
      End If
      
   Screen.MousePointer = vbHourglass
   Picture1.Picture = LoadPicture(cmbMaps.Text)
   Picture1.AutoSize = True
   Picture1.BorderStyle = 0
   txtPixWidth.Text = str(Picture1.Width) / Screen.TwipsPerPixelX
   txtPixHeight.Text = str(Picture1.Height) / Screen.TwipsPerPixelY
   'now release picture memory
   Picture1.Picture = LoadPicture(sEmpty)
   Screen.MousePointer = vbDefault
   'Call Form_QueryUnload(0, 0)
   Exit Sub
   
errhand:
   Screen.MousePointer = vbDefault
   Select Case Err.Number
      Case 481 'invalid picture
        ier% = -1
        MsgBox "Picture format of the map is not supported!" & vbLf & _
               "(Supported formats are: bmp, icon, metafile, gif, jpeg, jpg.)", _
               vbCritical + vbOKOnly, "MapDigitizer"
      Case Else
        MsgBox "Encountered error #: " & Err.Number & vbLf & _
               "in program module GDOptionsfrm:cmdPixelSize_Click." & vbLf & _
               Err.Description, vbCritical + vbOKOnly, "MapDigitizer"
   End Select
End Sub

Private Sub cmdRestoreDefault_Click()
    'restore the default values
    If txttifpath <> sEmpty Or txtTifViewer <> sEmpty Or txtCommandLine <> sEmpty Then
       response = MsgBox("This will overwrite the current values!" & vbLf & _
                       "Continue?", vbYesNoCancel + vbInformation, "MapDigitizer")
       If response = vbYes Then
          tifDir$ = direct$
          tifViewerDir$ = GetSystemPath & "\shimgvw.dll"
          txtTifViewer = tifViewerDir$
          tifCommandLine$ = "RUNDLL32.EXE " & tifViewerDir$ & ", ImageView_Fullscreen"
          txtCommandLine = tifCommandLine$
          End If
    Else
        tifDir$ = NEDdir
        tifViewerDir$ = GetSystemPath & "\shimgvw.dll"
        txtTifViewer = tifViewerDir$
        tifCommandLine$ = "RUNDLL32.EXE " & tifViewerDir$ & ", ImageView_Fullscreen"
        txtCommandLine = tifCommandLine$
        End If
End Sub

Private Sub cmdSave_Click()
    
    'record the changes to the tif paths
    tifDirff$ = txttifpath
    tifViewerDirff$ = txtTifViewer
    tifCommandLineff$ = txtCommandLine
    
    If Dir(direct$ & "\gdb_tif.sav") <> sEmpty Then
         'check for changes in the tif viewer files
         filin% = FreeFile
         Open direct$ & "\gdb_tif.sav" For Input As #filin%
         Line Input #filin%, tifDirf$
         Line Input #filin%, tifViewerDirf$
         Line Input #filin%, tifCommandLinef$
         Close #filin%
         
         If tifDirf$ <> tifDirff$ Or tifViewerDirf$ <> tifViewerDirff$ Or tifCommandLinef$ <> tifCommandLineff$ Then
            response = MsgBox("Overwrite the old parameters?", vbYesNoCancel + vbQuestion, App.Title)
            If response <> vbYes Then Exit Sub
            End If
       
        'record the changes to the tif paths
        tifDir$ = txttifpath
        tifViewerDir$ = txtTifViewer
        tifCommandLine$ = txtCommandLine
        
        'save the changes to the tif paths
        filin% = FreeFile
        Open direct$ & "\gdb_tif.sav" For Output As #filin%
        Print #filin%, tifDir$
        Print #filin%, tifViewerDir$
        Print #filin%, tifCommandLine$
        Close #filin%
    Else
        'record the changes to the tif paths
        tifDir$ = txttifpath
        tifViewerDir$ = txtTifViewer
        tifCommandLine$ = txtCommandLine
        
        'save the changes to the tif paths
        filin% = FreeFile
        Open direct$ & "\gdb_tif.sav" For Output As #filin%
        Print #filin%, tifDir$
        Print #filin%, tifViewerDir$
        Print #filin%, tifCommandLine$
        Close #filin%
        End If
        
End Sub

Private Sub cmdSaveMaps_Click()

   'make sure everything is in decimal numbers
   
   If Mid$(LCase(MaskEdBox1.Text), 1, 3) = "lon" And Mid$(LCase(MaskEdBox2.Text), 1, 3) = "lat" Then
     txtULGeoX = ConvertDegToNumber(txtULGeoX)
     txtULGeoY = ConvertDegToNumber(txtULGeoY)
     txtLRGeoX = ConvertDegToNumber(txtLRGeoX)
     txtLRGeoY = ConvertDegToNumber(txtLRGeoY)
     txtULGridX = ConvertDegToNumber(txtULGridX)
     txtULGridY = ConvertDegToNumber(txtULGridY)
     txtLRGridX = ConvertDegToNumber(txtLRGridX)
     txtLRGridY = ConvertDegToNumber(txtLRGridY)
     End If
     
  'first check the limits
  If val(JustConvertDegToNumber(txtULGeoX.Text)) > val(JustConvertDegToNumber(txtLRGeoX.Text)) Then
     Call MsgBox("The UL X Geo coordinate must be less then the LR X Geo coordinate." _
                   & vbCrLf & "" _
                   & vbCrLf & "(For Western Hemisphere, use negative longitudes;" _
                   & vbCrLf & "For Eastern Hemisphere, use positive longitudes)" _
                   , vbExclamation, "X Geo LImits error")
     Exit Sub
     End If
       
  If val(JustConvertDegToNumber(txtULGeoY.Text)) < val(JustConvertDegToNumber(txtLRGeoY.Text)) Then
     Call MsgBox("The UL Y Geo coordinate must be greater then the LR Y Geo coordinate." _
                   , vbExclamation, "Y Geo LImits error")
     Exit Sub
     End If

  'check the paramaters
'  If val(txtULPixX) = val(txtLRPixX) Then
'     txtULPixX = 0
'     txtLRPixX = pixwi
'     End If
'
'  If val(txtULPixY) = val(txtLRPixY) Then
'     txtULPixY = 0
'     txtLRPixY = pixhi
'     LblY = "Pixl"
'     End If
'
'  If val(txtULGeoX) = val(txtLRGeoX) Then
'     lblX = "Pixl"
'     txtULGeoX = 0
'     txtLRGeoX = txtLRPixX
'     End If
'
'  If val(txtULGridY) = val(txtLRGeoY) Then
'     LblY = "Pixl"
'     txtULGridY = 0
'     txtLRGridY = txtLRPixY
'     End If
  
  If Trim$(txtLRGridX) <> sEmpty And Trim$(txtULGridX) <> sEmpty Then
     If val(txtLRGridX) = val(txtULGridX) Then
        RSMethod1 = False
        RSMethod2 = False
        If (ULGeoX = LRGeoX And ULGeoY = LRGeoY) Or (ULPixX = LRPixX And ULPixY = LRPixY) Then RSMethod0 = False
        End If
     End If
     
  If Trim$(txtLRGridY) <> sEmpty And Trim$(txtULGridY) <> sEmpty Then
     If val(txtLRGridY) = val(txtULGridY) Then
        RSMethod1 = False
        RSMethod2 = False
        If (ULGeoX = LRGeoX And ULGeoY = LRGeoY) Or (ULPixX = LRPixX And ULPixY = LRPixY) Then RSMethod0 = False
        End If
     End If
     
  'save the map parameters
  'first save visible settings for current map
  nn& = cmbMaps.ListCount - 1
  If nn& = -1 Then nn& = cmbMaps.ListCount - 1
  ReDim Preserve MapParms(19, 0 To UBound(MapParms, 2))
  MapParms(0, nn&) = cmbMaps.Text
  MapParms(1, nn&) = MaskEdBox1.Text
  MapParms(2, nn&) = MaskEdBox2.Text
  MapParms(3, nn&) = JustConvertDegToNumber(txtULGeoX)
  MapParms(4, nn&) = JustConvertDegToNumber(txtULGeoY)
  MapParms(5, nn&) = JustConvertDegToNumber(txtLRGeoX)
  MapParms(6, nn&) = JustConvertDegToNumber(txtLRGeoY)
  MapParms(7, nn&) = Trim$(txtPixWidth)
  MapParms(8, nn&) = Trim$(txtPixHeight)
  MapParms(9, nn&) = JustConvertDegToNumber(txtGridX)
  MapParms(10, nn&) = JustConvertDegToNumber(txtGridY)
  MapParms(11, nn&) = JustConvertDegToNumber(txtULPixX)
  MapParms(12, nn&) = JustConvertDegToNumber(txtULPixY)
  MapParms(13, nn&) = JustConvertDegToNumber(txtLRPixX)
  MapParms(14, nn&) = JustConvertDegToNumber(txtLRPixY)
  MapParms(15, nn&) = JustConvertDegToNumber(txtLRGridX)
  MapParms(16, nn&) = JustConvertDegToNumber(txtLRGridY)
  MapParms(17, nn&) = JustConvertDegToNumber(txtULGridX)
  MapParms(18, nn&) = JustConvertDegToNumber(txtULGridY)
  If MpaUnits = 0 Then MapUnits = 1
  MapParms(19, nn&) = str$(MapUnits)
  
  ULPixX = val(txtULPixX)
  ULPixY = val(txtULPixY)
  LRPixX = val(txtLRPixX)
  LRPixY = val(txtLRPixY)
  LRGridX = val(txtLRGridX)
  LRGridY = val(txtLRGridY)
  ULGridX = val(txtULGridX)
  ULGridY = val(txtULGridY)
  ULGeoX = val(txtULGeoX)
  ULGeoY = val(txtULGeoY)
  LRGeoX = val(txtLRGeoX)
  LRGeoY = val(txtLRGeoY)
  NX_CALDAT = val(txtGridX)
  NY_CALDAT = val(txtGridY)
  pixwi = val(txtPixWidth)
  pixhi = val(txtPixHeight)
  DigiConvertToMeters = val(txtCustom)

  'now store all the maps
    
'  If Dir(direct$ & "\new5.bmp") <> sEmpty Then
'     found% = 0 'check if it is in cmbMaps combo box
'     For i% = 1 To cmbMaps.ListCount
'       If InStr(cmbMaps.List(i% - 1), "\new5.bmp") <> 0 Then
'          found% = 1
'          Exit For
'          End If
'     Next i%
'  Else
'     found% = 1 'can't add it to list since it doesn't exist
'     End If
'
  filemap% = FreeFile
  Open direct$ & "\gdbmap.sav" For Output As #filemap%
  Write #filemap%, "This file is used by the MapDigitizer program. Don't erase it!"
'
'  If found% = 0 Then 'didn't find default geo file in cmbMaps, so add it
'     Write #filemap%, direct$ & "\new5.bmp", "ITMx", "ITMy", _
'           "80000", "240000", "1300000", "880000", "1268", "3338", "0", "0", "", "", "", "", "", "", "", ""
'     End If
  'write all stored map parameters
  For i% = 1 To cmbMaps.ListCount
     Write #filemap%, MapParms(0, i% - 1), MapParms(1, i% - 1), MapParms(2, i% - 1), _
                      MapParms(3, i% - 1), MapParms(5, i% - 1), _
                      MapParms(4, i% - 1), MapParms(6, i% - 1), _
                      MapParms(7, i% - 1), MapParms(8, i% - 1), _
                      MapParms(9, i% - 1), MapParms(10, i% - 1), _
                      MapParms(11, i% - 1), MapParms(12, i% - 1), _
                      MapParms(13, i% - 1), MapParms(14, i% - 1), _
                      MapParms(15, i% - 1), MapParms(16, i% - 1), _
                      MapParms(17, i% - 1), MapParms(18, i% - 1), MapParms(19, i% - 1)
 Next i%
  Close #filemap%
  
  AddedMaps = False 'remember that saved the maps

End Sub

'Private Sub cmdTopoOut_Click()
'   On Error GoTo errhand
'   cmdlgColor.ShowColor
'   ContourColor& = cmdlgColor.color
'   shpContours.FillColor = ContourColor&
'   Exit Sub
'errhand:
'End Sub
'
'Private Sub cmdTopoWell_Click()
'   On Error GoTo errhand
'   cmdlgColor.ShowColor
'   RSColor& = cmdlgColor.color
'   shpRS.FillColor = RSColor&
'   Exit Sub
'errhand:
'End Sub

Private Sub cmdBrowsePath_Click()
  'browse for tif file directory
  On Error GoTo c3error
  cmdlgdtb.CancelError = True
  cmdlgdtb.Filter = "Tif files (*.tif) |*.tif|Jpg, Jpeg files (*.jpg)|*.jpg|All files (*.*)|*.*"
  cmdlgdtb.FilterIndex = 1
  cmdlgdtb.FileName = direct$ + "\*.tif"
  cmdlgdtb.ShowOpen
  filn$ = cmdlgdtb.FileName
  
  For i% = Len(filn$) To 1 Step -1
     If Mid$(filn$, i%, 1) = "\" Then
        Exit For
        End If
  Next i%
  
  If i% > 1 Then
     dirtmp$ = Mid$(filn$, 1, i% - 1)
     If InStr(UCase$(dirtmp$), "\CD_") <> 0 Then
        'this is almost certainly a mistake, so remove the CD_ part
        'and retain only the root directory
        dirtmp$ = Mid$(filn$, 1, InStr(UCase$(dirtmp$), "\CD_") - 1)
        End If
     txttifpath = dirtmp$
  Else
     MsgBox "Not a valid path!", vbExclamation + vbOKOnly, "MapDigitizer"
     Exit Sub
     End If

c3error:
  Exit Sub

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLRGridXchange_Click
' Author    : Dr-John-K-Hall
' Date      : 3/5/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLRGridXchange_Click()
   If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
      txtLRGridX = ConvertDegToNumber(txtLRGridX)
   Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdULGridXchange_Click
' Author    : Dr-John-K-Hall
' Date      : 3/5/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdULGridXchange_Click()
   If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
      txtULGridX = ConvertDegToNumber(txtULGridX)
   Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdULGridYchange_Click
' Author    : Dr-John-K-Hall
' Date      : 3/5/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdULGridYchange_Click()
   If MaskEdBox1.Text = "lon." And MaskEdBox2.Text = "lat." Then
      txtULGridY = ConvertDegToNumber(txtULGridY)
   Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                 & vbCrLf & "" _
                 & vbCrLf & "(Hint: change the coordinates labels above to "".lon."", "".lat"")" _
                 , vbInformation, "Coordinate Conversion Error")
     End If
      
End Sub

Private Sub cmdPasteULPixY_Click()
   txtULPixY = GDMDIform.Text6
End Sub

Private Sub cmdPointColor_Click()
   On Error GoTo errhand
   cmdlgColor.ShowColor
   PointColor& = cmdlgColor.color
   shpPoints.FillColor = PointColor&
   Exit Sub
errhand:
End Sub

Private Sub cmdLineColor_Click()
   On Error GoTo errhand
   cmdlgColor.ShowColor
   LineColor& = cmdlgColor.color
   shpLines.FillColor = LineColor&
   Exit Sub
errhand:
End Sub

Private Sub cmdContours_Click()
   On Error GoTo errhand
   cmdlgColor.ShowColor
   ContourColor& = cmdlgColor.color
   shpContours.FillColor = ContourColor&
   Exit Sub
errhand:
End Sub

Private Sub cmdRSColor_Click()
   On Error GoTo errhand
   cmdlgColor.ShowColor
   RSColor& = cmdlgColor.color
   shpRS.FillColor = RSColor&
   Exit Sub
errhand:
End Sub

Private Sub cmdBrowseNewDTM_Click()

   'Opens a Treeview control that displays the directories in a computer

      Dim lpIDList As Long
      Dim sBuffer As String
      Dim szTitle As String
      Dim tBrowseInfo As BrowseInfo

      szTitle = "Choose a directory to store generated dtm files"
      With tBrowseInfo
         .hWndOwner = Me.hwnd
         .lpszTitle = lstrcat(szTitle, "")
         .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
      End With

      lpIDList = SHBrowseForFolder(tBrowseInfo)

      If (lpIDList) Then
         sBuffer = Space(MAX_PATH)
         SHGetPathFromIDList lpIDList, sBuffer
         sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
         txtNewDTM.Text = sBuffer
      End If

End Sub

Private Sub Form_Activate()
   'make sure that button is pressed when form is visible
   GDMDIform.Toolbar1.Buttons(1).value = tbrPressed
   buttonstate&(1) = 1
End Sub

Private Sub Form_Deactivate()
   'when form is invisible, then unpress the button
   GDMDIform.Toolbar1.Buttons(1).value = tbrUnpressed
   buttonstate&(1) = 0
End Sub

Private Sub Form_Load()
   
   On Error GoTo errhand
   
   OptionsVis = True
   GDMDIform.Toolbar1.Buttons(1).value = tbrPressed
   buttonstate&(1) = 1
   
   With GDOptionsfrm
     .tabOptions.Tab = 0
     .Top = 0
     .Left = 0
     .txtULGeoX = sEmpty
     .txtULGeoY = sEmpty
     .txtLRGeoX = sEmpty
     .txtPixWidth = sEmpty
     .txtLRGeoY = sEmpty
     .txtPixHeight = sEmpty
     .txtGridX = "0"
     .txtGridY = "0"
     .txtULPixX = sEmpty
     .txtULPixY = sEmpty
     .txtLRPixX = sEmpty
     .txtLRPixY = sEmpty
     .txtLRGridX = sEmpty
     .txtLRGridY = sEmpty
     .txtULGridX = sEmpty
     .txtULGridY = sEmpty
     .txtDTMitmx = sEmpty
     .txtDTMitmy = sEmpty
     .txtDTMlon = sEmpty
     .txtDTMlat = sEmpty
     .txtAzi = sEmpty
     .txtStepAzi = sEmpty
     .txtaprn = sEmpty
     .txtCustom = 1#
     MapUnits = 1#
     
     If GaussMethod Then
        .optGaussian.value = True
     Else
        .optCramer.value = True
        End If
     
     .cmbMaps.Clear
      filemap% = FreeFile
      If Dir(direct$ & "\gdbmap.sav") <> sEmpty Then
        Open direct$ & "\gdbmap.sav" For Input As #filemap%
        Input #filemap%, doclin$
        nn& = 0
        WarningFlag% = 0
        Do Until EOF(filemap%)
           Input #filemap%, MapsName$, s_LblXt$, s_LblYt$, s_x1t$, s_x2t$, s_y1t$, s_y2t$, s_pixwit$, s_pixhit$, s_NX_CALDAT$, s_NY_CALDAT$, _
                            s_ULPixX$, s_ULPixY$, s_LRPixX$, s_LRPixY$, s_LRGridX$, s_LRGridY$, s_ULGridX$, s_ULGridY$, s_MapUnits$
           'check if map exists before adding it to array
           If Dir(MapsName$) = sEmpty Then
              'try adding path of this program
              testpath$ = App.Path & "\" & MapsName$
              If Dir(testpath$) <> sEmpty Then
                 MapsName$ = testpath$
              Else
                 WarningFlag% = 1
                 GoTo 100 'skip this entry
                 End If
              End If
                 
           ReDim Preserve MapParms(19, nn&)
           MapParms(0, nn&) = MapsName$
           MapParms(1, nn&) = s_LblXt$
           MapParms(2, nn&) = s_LblYt$
           MapParms(3, nn&) = s_x1t$
           MapParms(4, nn&) = s_y1t$
           MapParms(5, nn&) = s_x2t$
           MapParms(6, nn&) = s_y2t$
           MapParms(7, nn&) = s_pixwit$
           MapParms(8, nn&) = s_pixhit$
           MapParms(9, nn&) = s_NX_CALDAT$
           MapParms(10, nn&) = s_NY_CALDAT$
           MapParms(11, nn&) = s_ULPixX$
           MapParms(12, nn&) = s_ULPixY$
           MapParms(13, nn&) = s_LRPixX$
           MapParms(14, nn&) = s_LRPixY$
           MapParms(15, nn&) = s_LRGridX$
           MapParms(16, nn&) = s_LRGridY$
           MapParms(17, nn&) = s_ULGridX$
           MapParms(18, nn&) = s_ULGridY$
           MapParms(19, nn&) = s_MapUnits$
           cmbMaps.AddItem MapsName$
           nn& = nn& + 1
100:
        Loop
        Close #filemap%
        End If
        
   End With
   
   If WarningFlag% = 1 And ReportPaths& = 0 Then
      Call MsgBox("Not all maps whose names are stored in the ""gdbmaps.sav"" file" _
                  & vbCrLf & "could be added to the map list since their path is wrong." _
                  & vbCrLf & "" _
                  & vbCrLf & "(Hint: edit paths using the ""Options"" button)" _
                  , vbExclamation, "File error")
      WarningFlag% = 0
      End If
   
   '----------Geologic Map option defaults----------
   
   '-------load in previously recorded settings
   myfile = Dir(direct$ + "\gdbinfo.sav")
   
   If myfile = sEmpty Then
   
'      If lblX = sEmpty Then lblX = "lon." '"ITMx"
'      If LblY = sEmpty Then LblY = "lat." '"ITMy"
      MaskEdBox1.Text = s_LblXt$
      MaskEdBox2.Text = s_LblYt$
      
      'default map and map boundaries and geo map pixel sizes
      picnam$ = sEmpty
'      If Dir(direct$ & "\IsraelShadedReliefMap.jpg") <> sEmpty Then
'         picnam$ = direct$ & "\new5.bmp"
'         If IsNull(ULGeoX) Then ulgeox = 80000#
'         If IsNull(ulgeoy) Then ulgeoy = 1300000#
'         If IsNull(lrgeox) Then lrgeox = 240000#
'         If IsNull(lrgeoy) Then lrgeoy = 880000#
'         If ulgeox = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'            ulgeox = 80000#
'            ulgeoy = 1300000#
'            lrgeox = 240000#
'            lrgeoy = 880000#
'            End If
'         If IsNull(pixwi) Then pixwi = 1162
'         If IsNull(pixhi) Then pixhi = 3046
'         If pixwi = 0 Then pixwi = 1162
'         If pixhi = 0 Then pixhi = 3046
'         End If
        
      txtULGeoX.Text = s_ULGeoX$
      txtULGeoY.Text = s_ULGeoY$
      txtLRGeoX.Text = s_LRGeoX$
      txtPixWidth.Text = s_pixwi$
      txtLRGeoY.Text = s_LRGeoY$
      txtPixHeight.Text = s_pixhi$
      txtGridX.Text = "0"
      txtGridY.Text = "0"
      txtMaxHighlight = "20000"
      NX_CALDAT = 0
      NY_CALDAT = 0
      
      RSMethod1 = False
      RSMethod2 = False
      RSMethod0 = False

'      ReDim SX_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
'      ReDim SY_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
'      ReDim GX_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
'      ReDim GY_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
         
      chkSave.value = vbChecked 'save on closing
      
      'default plot mark colors
      PointColor& = 255 'red
      LineColor& = 65280 'green'65535 'yellow
      ContourColor& = 10485760 'dark blue
      RSColor& = 65535 'yellow '65280 'green
'      UnknownColor& = 8388736 'purple
      shpPoints.FillColor = PointColor&
      shpLines.FillColor = LineColor&
      shpContours.FillColor = ContourColor&
      shpRS.FillColor = RSColor&
'      shpUnknown.FillColor = UnknownColor&

      numDistLines = 5
      numDistContour = 5
      numContours = 5
      numSensitivity = 50
      txtDistLines = numDistLines
      txtDistContour = numDistContour
      txtSensitivity = numSensitivity
      cmbContour.Text = str$(numContours)
      
      DigiSearchRegion = 10
      txtDistPixelSearch = DigiSearchRegion
      txtEraserBrushSize = MinDigiEraserBrushSize
      
      If ChainCodeMethod = 0 Then
         optFreeman.value = True
      ElseIf ChainCodeMethod = 1 Then
         optBug.value = True
         End If
         
    If MapUnits = 0 Or MapUnits = 1 Then
       chkfeet.value = vbUnchecked
       chkFathoms.value = vbUnchecked
       MapUnits = 1#
       GDMDIform.Text3.ToolTipText = "Elevation (meters)"
       GDMDIform.Text7.ToolTipText = "Elevation (meters) at center of clicked point"
    ElseIf MapUnits = 0.30479999798832 Then
       chkFathoms.value = vbUnchecked
       chkfeet.value = vbChecked
       GDMDIform.Text3.ToolTipText = "Elevation (feet)"
       GDMDIform.Text7.ToolTipText = "Elevation (feet) at center of clicked point"
    ElseIf MapUnits = 1.8288002 Then
       chkfeet.value = vbUnchecked
       chkFathoms.value = vbChecked
       GDMDIform.Text3.ToolTipText = "Elevation (fathoms)"
       GDMDIform.Text7.ToolTipText = "Elevation (fathoms) at center of clicked point"
       End If
           
           
      'skip loading map parameters if they don't exist
      If ULGeoX = 0 And LRGeoX = 0 And ULGeoY = 0 And LRGeoY = 0 Then GoTo ld500
      
      'record default map parameters in map array
      cmbMaps.AddItem picnam$
      If cmbMaps.ListIndex = -1 Then nn& = cmbMaps.ListCount - 1
      ReDim Preserve MapParms(18, cmbMaps.ListCount - 1)
      MapParms(0, nn&) = cmbMaps.Text
      MapParms(1, nn&) = MaskEdBox1.Text
      MapParms(2, nn&) = MaskEdBox2.Text
      MapParms(3, nn&) = JustConvertDegToNumber(txtULGeoX)
      MapParms(4, nn&) = JustConvertDegToNumber(txtULGeoY)
      MapParms(5, nn&) = JustConvertDegToNumber(txtLRGeoX)
      MapParms(6, nn&) = JustConvertDegToNumber(txtLRGeoY)
      MapParms(7, nn&) = Trim$(txtPixWidth)
      MapParms(8, nn&) = Trim$(txtPixHeight)
      MapParms(9, nn&) = JustConvertDegToNumber(txtGridX)
      MapParms(10, nn&) = JustConvertDegToNumber(txtGridY)
      MapParms(11, nn&) = JustConvertDegToNumber(txtULPixX)
      MapParms(12, nn&) = JustConvertDegToNumber(txtULPixY)
      MapParms(13, nn&) = JustConvertDegToNumber(txtLRPixX)
      MapParms(14, nn&) = JustConvertDegToNumber(txtLRPixY)
      MapParms(15, nn&) = JustConvertDegToNumber(txtLRGridX)
      MapParms(16, nn&) = JustConvertDegToNumber(txtLRGridY)
      MapParms(17, nn&) = JustConvertDegToNumber(txtULGridX)
      MapParms(18, nn&) = JustConvertDegToNumber(txtULGridY)
      If MpaUnits = 0 Then MapUnits = 1
      MapParms(19, nn&) = str$(MapUnits)
      cmbMaps.ListIndex = nn&
      
ld500:
      'attempt to fill in other defaults
      defdb1$ = direct$
      If Dir(defdb1$ & "\pal_pr.mdb") <> sEmpty Then
         txtdb1 = defdb1$
         End If
         
      If Installation_Type = 1 Then
         defdb2$ = direct$
         txtdb2 = direct$
         
         defdb3$ = direct$
         txtdb3 = direct$
         
      Else
         
        'Try v,w,x,y,z as default mapped network drive letters
        'of the two databases (they are usually on the same directory)
        For i% = Asc("v") To Asc("z")
           defdb2$ = Chr$(i%) & ":"
           If Dir(defdb2$ & "\pal_dt.mdb") <> sEmpty Then
              txtdb2 = defdb2$
              Exit For
              End If
        Next i%
                
        For i% = Asc("v") To Asc("z")
           defdb3$ = Chr$(i%) & ":"
           If Dir(defdb3$ & "\pal_old.mdb") <> sEmpty Then
              txtdb3 = defdb3$
              Exit For
              End If
        Next i%
        End If
            
'      If Installation_Type = 0 Then
'        defdtm$ = "d:\dtm"
'        If Dir(defdtm$ & "\dtm-map.loc") <> sEmpty Then
'           txtdtm = defdtm$
'           End If
'      Else
'        defdtm$ = direct$ & "\dtm"
'        txtdtm = direct$ & "\dtm"
'        If Dir(txtdtm & "\dtm-map.loc") <> sEmpty Then
'           optDTM.Enabled = True
'        Else
'           txtdtm.Text = sEmpty
'           End If
'        txtAster = direct$ & "\aster"
'        If Dir(txtAster + "\N31E035.bil") <> sEmpty Then
'           optAster.Enabled = True
'        Else
'           txtAster = sEmpty
'           End If
'        End If
         
      
      defmxd$ = direct$
      If Dir(defmxd$ & "\IsraelESRI.mxd") <> sEmpty Then
         txtmxd = defmxd$ & "\IsraelESRI.mxd"
         End If
         
      defgoogle$ = Mid$(direct$, 1, 3) & "Program Files\Google\client\Google Earth"
      If Dir(defgoogle$ & "\googleearth.exe") <> sEmpty Then
         txtGoogle = defgoogle$
         google = True
         End If
         
      'default URL's for outcroppings and wells icons to use in the kml output file
      txtOutCropIcon = "http://maps.google.com/mapfiles/kml/pal4/icon49.png"
      URL_OutCrop = txtOutCropIcon
      txtWellIcon = "http://maps.google.com/mapfiles/kml/pal4/icon48.png"
      URL_Well = txtWellIcon
      
      txtEraserBrushSize = "1"
      MinDigiEraserBrushSize = 1
         
   Else
   
        If infonum& > 0 Then
          Close #infonum&
          infonum& = 0
          End If

      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Input As #infonum&
      Input #infonum&, doclin$
      Input #infonum&, dirNewDTMtmp
      Input #infonum&, MinDigiEraserBrushSize
      Input #infonum&, NEDdirtmp
      Input #infonum&, dtmdirtmp
      Input #infonum&, ChainCodeMethod
      Input #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
      Input #infonum&, PointCenterClick
      Input #infonum&, picnam$
      Input #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
      Input #infonum&, ReportPaths&, DigiSearchRegion, numMaxHighlight&, Save_xyz%
      Input #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Input #infonum&, IgnoreAutoRedrawError%
      Input #infonum&, UseNewDTM%, nOtherCheck%
      Input #infonum&, googledirtmp, URL_OutCroptmp, URL_Welltmp, kmldirtmp, ASTERdirtmp, DTMtypetmp
      Input #infonum&, NX_CALDAT, NY_CALDAT
      Input #infonum&, RSMethod0, RSMethod1, RSMethod2
      Input #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
      Input #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
      Close #infonum&
      
      'if directory name is only letter, like d:\ then truncate
      If InStr(dirNewDTMtmp, "\") <> 0 And Len(dirNewDTMtmp) = 3 Then dirNewDTMtmp = Mid$(dirNewDTMtmp, 1, 2)
'      If InStr(dbdir2tmp, "\") <> 0 And Len(dbdir2tmp) = 3 Then dbdir2tmp = Mid$(dbdir2tmp, 1, 2)
      If InStr(NEDdirtmp, "\") <> 0 And Len(NEDdirtmp) = 3 Then NEDdirtmp = Mid$(NEDdirtmp, 1, 2)
      If InStr(dtmdirtmp, "\") <> 0 And Len(dtmdirtmp) = 3 Then dtmdirtmp = Mid$(dtmdirtmp, 1, 2)
      If InStr(ASTERdirtmp, "\") <> 0 And Len(ASTERdirtmp) = 3 Then ASTERdirtmp = Mid$(ASTERdirtmp, 1, 2)
'      If InStr(topodirtmp, "\") <> 0 And Len(topodirtmp) = 3 Then topodirtmp = Mid$(topodirtmp, 1, 2)
'      If InStr(arcdirtmp, "\") <> 0 And Len(arcdirtmp) = 3 Then arcdirtmp = Mid$(arcdirtmp, 1, 2)
'      If InStr(accdirtmp, "\") <> 0 And Len(accdirtmp) = 3 Then accdirtmp = Mid$(accdirtmp, 1, 2)
      If InStr(googledirtmp, "\") <> 0 And Len(googledirtmp) = 3 Then googledirtmp = Mid$(googledirtmp, 1, 2)
      If InStr(kmldirtmp, "\") <> 0 And Len(kmldirtmp) = 3 Then kmldirtmp = Mid$(kmldirtmp, 1, 2)
      
      If MinDigiEraserBrushSize = 0 Then MinDigiEraserBrushSize = 1
       
      If lblX = "lon." And LblY = "lat." Then GDOptionsfrm.chkLatLon.value = vbChecked
       
      If Not Digitizing Then
        If lblX = sEmpty Then lblX = "ITMx"
        If LblY = sEmpty Then LblY = "ITMy"
        End If
      MaskEdBox1.Text = lblX
      MaskEdBox2.Text = LblY
      txtULPixX = ULPixX
      txtULPixY = ULPixY
      txtLRPixX = LRPixX
      txtLRPixY = LRPixY
      txtLRGridX = LRGridX
      txtLRGridY = LRGridY
      txtULGridX = ULGridX
      txtULGridY = ULGridY
      
      txtDistContour = numDistContour
      txtDistLines = numDistLines
      txtSensitivity = numSensitivity
      cmbContour.Text = str$(numContours)
      txtDistPixelSearch = DigiSearchRegion
      txtEraserBrushSize = MinDigiEraserBrushSize
      txtCustom = DigiConvertToMeters
      
      If PointCenterClick = 1 Then
         chkCenterClick.value = vbChecked
         End If
         
      If LineElevColors& = 1 Then
         chkRainbow.value = vbChecked
         End If
      
      If ChainCodeMethod = 0 Then
         optFreeman.value = True
      ElseIf ChainCodeMethod = 1 Then
         optBug.value = True
         End If
        
      If picnam$ = sEmpty Then
         Call MsgBox("No maps have been added yet." _
                     & vbCrLf & "Please add a map." _
                     , vbInformation, "Map Error")
'         picnam$ = "IsraelShadedReliefMap.jpg"
         End If
         
      cmbMaps.Text = picnam$
        
      'coordinate boundaries and geo map pixel size
      If IsNull(ULGeoX) Then ULGeoX = 0 '80000#
      If IsNull(ULGeoY) Then ULGeoY = 0 '1300000#
      If IsNull(LRGeoX) Then LRGeoX = 0 '240000#
      If IsNull(LRGeoY) Then LRGeoY = 0 '880000#
      If IsNull(NX_CALDAT) Then NX_CALDAT = 0
      If IsNull(NY_CALDAT) Then NY_CALDAT = 0
      
'      If X1 = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'         X1 = 80000#
'         ulgeoy = 1300000#
'         lrgeox = 240000#
'         lrgeoy = 880000#
'         End If
'      If IsNull(pixwi) Then pixwi = 1268
'      If IsNull(pixhi) Then pixhi = 3338
'      If pixwi = 0 Then pixwi = 1268
'      If pixhi = 0 Then pixhi = 3338
      
      If Not RSMethod1 And Not RSMethod2 And Not RSMethod0 Then
         RSMethod1 = False
         RSMethod2 = False
         RSMethod0 = False
         End If
        
      txtULGeoX.Text = str(ULGeoX)
      txtULGeoY.Text = str(ULGeoY)
      txtLRGeoX.Text = str(LRGeoX)
      txtPixWidth.Text = str(pixwi)
      txtLRGeoY.Text = str(LRGeoY)
      txtPixHeight.Text = str(pixhi)
      txtGridX.Text = str(NX_CALDAT)
      txtGridY.Text = str(NY_CALDAT)
      
      Select Case HeightPrecision
      
        Case 0
           optInteger.value = True
        Case 1
           optFloat.value = True
        Case 2
           optDouble.value = True
        
      End Select
      
    If MapUnits = 0 Or MapUnits = 1 Then
       chkfeet.value = vbUnchecked
       chkFathoms.value = vbUnchecked
       MapUnits = 1#
       GDMDIform.Text3.ToolTipText = "Elevation (meters)"
       GDMDIform.Text7.ToolTipText = "Elevation (meters) at center of clicked point"
    ElseIf MapUnits = 0.30479999798832 Then
       chkFathoms.value = vbUnchecked
       chkfeet.value = vbChecked
       GDMDIform.Text3.ToolTipText = "Elevation (feet)"
       GDMDIform.Text7.ToolTipText = "Elevation (feet) at center of clicked point"
    ElseIf MapUnits = 1.8288002 Then
       chkfeet.value = vbUnchecked
       chkFathoms.value = vbChecked
       GDMDIform.Text3.ToolTipText = "Elevation (fathoms)"
       GDMDIform.Text7.ToolTipText = "Elevation (fathoms) at center of clicked point"
       End If
       
'      If cmbMaps.ListCount = 0 Then
'         'no gdbmaps.sav file found, so begin new list of maps
'         nn& = 0
'         ReDim Preserve MapParms(18, nn&)
'         MapParms(0, nn&) = picnam$
'         MapParms(1, nn&) = lblX
'         MapParms(2, nn&) = LblY
'         MapParms(3, nn&) = Trim$(str$(ULGeoX))
'         MapParms(4, nn&) = Trim$(str$(ULGeoY))
'         MapParms(5, nn&) = Trim$(str$(LRGeoX))
'         MapParms(6, nn&) = Trim$(str$(LRGeoY))
'         MapParms(7, nn&) = Trim$(str$(pixwi))
'         MapParms(8, nn&) = Trim$(str$(pixhi))
'         MapParms(9, nn&) = Trim$(str$(NX_CALDAT))
'         MapParms(10, nn&) = Trim$(str$(NY_CALDAT))
'         MapParms(11, nn&) = Trim$(str$(ULPixX))
'         MapParms(12, nn&) = Trim$(str$(ULPixY))
'         MapParms(13, nn&) = Trim$(str$(LRPixX))
'         MapParms(14, nn&) = Trim$(str$(LRPixY))
'         MapParms(15, nn&) = Trim$(str$(LRGridX))
'         MapParms(16, nn&) = Trim$(str$(LRGridY))
'         MapParms(17, nn&) = Trim$(str$(ULGridX))
'         MapParms(18, nn&) = Trim$(str$(ULGridY))
'
'         If NX_CALDAT > 0 And NY_CALDAT > 0 Then
'            ReDim SX_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
'            ReDim SY_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
'            ReDim GX_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
'            ReDim GY_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
'            End If
'
'         cmbMaps.AddItem picnam$
'         cmbMaps.ListIndex = nn&
'      Else 'see if default picnam is in map list
'         found% = 0
'         For i% = 1 To cmbMaps.ListCount
'            If picnam$ = cmbMaps.List(i% - 1) Then
'               cmbMaps.ListIndex = i% - 1 'move to it on list
'               found% = 1
'               Exit For
'               End If
'         Next i%
'         If found% = 0 Then 'default picnam$ not in list, so store its parameters
'            cmbMaps.AddItem picnam$
'            nn& = cmbMaps.ListIndex
'            ReDim Preserve MapParms(18, nn&)
'            MapParms(0, nn&) = picnam$
'            MapParms(1, nn&) = lblX
'            MapParms(2, nn&) = LblY
'            MapParms(3, nn&) = Trim$(str$(ULGeoX))
'            MapParms(4, nn&) = Trim$(str$(ULGeoY))
'            MapParms(5, nn&) = Trim$(str$(LRGeoX))
'            MapParms(6, nn&) = Trim$(str$(LRGeoY))
'            MapParms(7, nn&) = Trim$(str$(pixwi))
'            MapParms(8, nn&) = Trim$(str$(pixhi))
'            MapParms(9, nn&) = Trim$(str$(NX_CALDAT))
'            MapParms(10, nn&) = Trim$(str$(NY_CALDAT))
'            MapParms(11, nn&) = Trim$(str$(ULPixX))
'            MapParms(12, nn&) = Trim$(str$(ULPixY))
'            MapParms(13, nn&) = Trim$(str$(LRPixX))
'            MapParms(14, nn&) = Trim$(str$(LRPixY))
'            MapParms(15, nn&) = Trim$(str$(LRGridX))
'            MapParms(16, nn&) = Trim$(str$(LRGridY))
'            MapParms(17, nn&) = Trim$(str$(ULGridX))
'            MapParms(18, nn&) = Trim$(str$(ULGridY))
'            cmbMaps.ListIndex = nn&
'            End If
'
'         End If
      
      'plot mark colors
      shpPoints.FillColor = PointColor&
      shpLines.FillColor = LineColor&
      shpContours.FillColor = ContourColor&
      shpRS.FillColor = RSColor&
'      shpUnknown.FillColor = UnknownColor&
      
      'maximum results highlighted
      If numMaxHighlight& = 0 Then numMaxHighlight& = 32767
      
      If ReportPaths& = 1 Then chkErrorMessage.value = vbChecked 'error warnings
      If SaveClose% = 1 Then chkSave.value = vbChecked 'save on closing
      If Save_xyz% = 1 Then chkSave_xyz.value = vbChecked 'save on closing
      If IgnoreAutoRedrawError% = 1 Then chkAutoRedraw.value = vbChecked 'ignore auto redraw errors
      txtMaxHighlight.Text = Trim$(str$(numMaxHighlight&)) 'max plotted records
      
'      Select Case SearchDBs% 'click database search option
'         Case 1
'            optAll_Click
'         Case 2
'            optActive_Click
'         Case 3
'            optInactive_Click
'      End Select
'
'      If linked = False Then
'         dbdir2 = dbdir2tmp
'         End If
      txtNewDTM = dirNewDTM
'      txtdb2 = dbdir2tmp
'
'      If linkedOld = False Then
'         NEDdir = NEDdirtmp
'         End If
'      txtdb3 = NEDdirtmp
'
      If heights = False Then
         dtmdir = dtmdirtmp
         NEDdir = NEDdirtmp
         ASTERdir = ASTERdirtmp
         End If
         
      If Dir(NEDdirtmp & "\z000000.hgt") <> sEmpty Then optDTM.Enabled = True
      
      If Dir(dtmdir & "\dtm-map.loc") <> sEmpty Then
         optDTM.Enabled = True
         If Not JKHDTM Then InitializeDTM
         End If
'
      txtAster = ASTERdirtmp
      If Dir(txtAster & "\N31E035.bil") <> sEmpty Then optAster.Enabled = True
      
      txtdtm = NEDdirtmp
'
'      If txtdtm = sEmpty Then 'try default
'         If Dir(direct$ & "\dtm\dtm-map.loc") <> sEmpty Then
'            txtdtm = direct$ & "\dtm"
'            optDTM.Enabled = True
'            End If
'         End If
      If txtAster = sEmpty Then 'try default
         If Dir(direct$ & "\aster\N31E035.bil") <> sEmpty Then
            txtdtm = direct$ & "\aster"
            optAster.Enabled = True
            End If
         End If
'
      DTMtype = DTMtypetmp
      If DTMtype <> 0 Then
         If DTMtype = 2 Then
            optDTM.value = True
         ElseIf DTMtype = 1 Then
            optAster.value = True
            End If
         End If
'
'      If topos = False Then
'         topodir = topodirtmp
'         End If
'      txttopo = topodirtmp
'
'      If arcs = False Then
'         arcdir = arcdirtmp
'         End If
'      txtarc = arcdirtmp
'      txtmxd = mxddirtmp
'
'      If acc = False Then
'         accdir = accdirtmp
'         End If
'      txtacc = accdirtmp

      If JKHDTM Then
         If val(XStepITM) = 0 Then XStepITM = 25
         If val(YStepITM) = 0 Then YStepITM = 30
      Else
         If val(XStepITM) = 0 Then XStepITM = 30
         If val(YStepITM) = 0 Then YStepITM = 30
         End If
          
      If val(XStepDTM) = 0 Then XStepDTM = 1#  '8.33333333333333E-04 / 3#
      If val(YStepDTM) = 0 Then YStepDTM = 1#  '8.33333333333333E-04 / 3#
      
      txtDTMitmx = XStepITM
      txtDTMitmy = YStepITM
      txtDTMlon = XStepDTM
      txtDTMlat = YStepDTM
      txtAzi = HalfAzi
      txtStepAzi = StepAzi
      txtaprn = Apprn
      
      If google = False Then
         'try default
         defgoogle$ = Mid$(direct$, 1, 3) & "Program Files\Google\Google Earth"
         If Dir(defgoogle$ & "\googleearth.exe") <> sEmpty Then
            googledirtmp = defgoogle$
         Else
            googledir = googledirtmp
            End If
         URL_OutCrop = URL_OutCroptmp
         URL_Well = URL_Welltmp
         kmldir = kmldirtmp
         If Trim$(kmldir) = sEmpty Then
            kmldir = direct$
            kmldirtmp = kmldir
         Else
            kmldir = kmldirtmp
            End If
         End If
      
      txtGoogle = googledirtmp
      txtkml = kmldirtmp
      
      If UseNewDTM% = 1 Then
         chkNewDTM.value = vbChecked
         UsingNewDTM = True
         End If
      
'      txtOutCropIcon = URL_OutCrop
'      If URL_OutCrop = sEmpty Then
'         txtOutCropIcon = "http://maps.google.com/mapfiles/kml/pal4/icon49.png"
'         URL_OutCrop = txtOutCropIcon
'         End If
'
'      txtWellIcon = URL_Well
'      If URL_Well = Empty Then
'         txtWellIcon = "http://maps.google.com/mapfiles/kml/pal4/icon48.png"
'         URL_Well = txtWellIcon
'         End If
'
'      'automatic replace null GL's with DTM heights
'      If nWellCheck% = 1 Then
'         chkWellsReplace.Value = vbChecked
'         ReplaceWellZ = True
'      Else
'         chkWellsReplace.Value = vbUnchecked
'         ReplaceWellZ = False
'         End If
'      If nOtherCheck% = 1 Then
'         chkOtherReplace.Value = vbChecked
'         ReplaceOtherZ = True
'      Else
'         chkOtherReplace.Value = vbUnchecked
'         ReplaceOtherZ = False
'         End If
   
   End If
   
'   'now input tif viewer path and name and command line
'   myfile = Dir(direct$ + "\gdb_tif.sav")
'
'   If myfile = sEmpty Then
'      'see if shimgvw.dll exists in the windows/system32 directory
'      tifDir$ = NEDdir
'      txttifpath = tifDir$
'
'      tifViewerDir$ = GetSystemPath & "\SHIMGVW.DLL"
'      If Dir(tifViewerDir$) <> sEmpty Then
'         txtTifViewer = tifViewerDir$
'         tifCommandLine$ = "RUNDLL32.EXE " & tifViewerDir$ & ", ImageView_Fullscreen"
'         txtCommandLine = tifCommandLine$
'         End If
'   Else
'      filin% = FreeFile
'      Open direct$ & "\gdb_tif.sav" For Input As #filin%
'      Line Input #filin%, tifDir$
'      Line Input #filin%, tifViewerDir$
'      Line Input #filin%, tifCommandLine$
'      Close #filin%
      
'      txttifpath = tifDir$
'      txtTifViewer = tifViewerDir$
'      txtCommandLine = tifCommandLine$
'      End If
'
'   If ActivatedVersion Then
'      frmPassword.Visible = False
'      GDOptionsfrm.Refresh
'      End If
      
   Exit Sub
      
errhand:
   Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo errhand
   
'----------------Save Geo Map settings-----------------

    'if changing Lblx to anything from ITM, then can't display heights
'    If heights And (Trim$(MaskEdBox1.Text) <> "ITMx" Or Trim$(MaskEdBox2.Text) <> "ITMy") Then '1
'       response = MsgBox("Heights will not be displayed with this coordinate system!" & vbLf & _
'                         "Do you wish to save these settings?", vbYesNoCancel + vbExclamation, "MapDigitizer")
'       If response <> vbYes Then '2
'          MaskEdBox1.Text = "ITMx"
'          MaskEdBox2.Text = "ITMy"
'          Cancel = True
'          Exit Sub
'          End If '2
'       End If '1
    
    
'       GDMDIform.Label1 = RTrim$(LTrim$(MaskEdBox1.Text))
'       GDMDIform.Label5 = RTrim$(LTrim$(MaskEdBox1.Text))
       lblX = RTrim$(LTrim$(MaskEdBox1.Text))
       
'       GDMDIform.Label2 = RTrim$(LTrim$(MaskEdBox2.Text))
'       GDMDIform.Label6 = RTrim$(LTrim$(MaskEdBox2.Text))
       LblY = RTrim$(LTrim$(MaskEdBox2.Text))
       
'    If SearchVis Then 'update map lables in search wizard '1
'       GDSearchfrm.lblXMax = lblX
'       GDSearchfrm.lblXMin = lblX
'       GDSearchfrm.lblYMin = LblY
'       GDSearchfrm.lblYMax = LblY
'       End If '1
       
    If Trim$(cmbMaps.Text) <> picnam$ Then picnam$ = Trim$(cmbMaps.Text)
'    If val(txtULGeoX.Text) >= val(txtLRGeoX.Text) Then '1
'       response = MsgBox("The current map's Lower-Right X Value must be larger" & vbLf & _
'                         "then the Upper-Left X Value!" & vbLf & _
'                         "Want to fix it? (You must fix it to use this map.)", vbExclamation + vbYesNoCancel, "MapDigitizer")
'       Select Case response
'          Case vbYes
'            Cancel = True
'            tabOptions.Tab = 0
'            Exit Sub
'          Case vbNo
'            GoTo op900
'          Case vbCancel
'            Cancel = True
'            Exit Sub
'       End Select
'       End If '1
'    If val(txtULGeoY.Text) <= val(txtLRGeoY.Text) Then '1
'       response = MsgBox("The current map's Lower-Right Y Value must be smaller" & vbLf & _
'                         "then the Upper-Left Y Value!" & vbLf & _
'                         "Want to fix it? (You must fix it to use this map.)", vbExclamation + vbYesNoCancel, "MapDigitizer")
'       Select Case response
'          Case vbYes
'            Cancel = True
'            tabOptions.Tab = 0
'            Exit Sub
'          Case vbNo
'            GoTo op900
'          Case vbCancel
'            Cancel = True
'            Exit Sub
'       End Select
'       End If '1
    
50:
    myfile = Dir(Trim$(cmbMaps.Text))
    If myfile = sEmpty Or Trim$(cmbMaps.Text) = sEmpty Then '1
       'try adding the current path
       myfile = Dir(Trim$(App.Path & "\" & cmbMaps.Text))
       If myfile <> sEmpty Then
          cmbMaps.Text = App.Path & "\" & cmbMaps.Text
          GoTo 50
          End If
       response = MsgBox("Can't find the current map!" & vbLf & _
                  "If you don't enter a picture name, you won't be able to use the maps." & vbLf & _
                  "Do you want to fix the name?", vbExclamation + vbYesNoCancel, "MapDigitizer")
       If response = vbYes Then '2
          Cancel = True
          tabOptions.Tab = 0
          Exit Sub
          End If '2
          
       cmbMaps.Text = sEmpty
       picnam$ = sEmpty
       picnam0$ = sEmpty
       GDMDIform.Toolbar1.Buttons(10).Enabled = False  'print maps
       GDMDIform.mnuPrintMap.Enabled = False
       GDMDIform.Toolbar1.Buttons(2).Enabled = False 'disenable large scale maps
       GDMDIform.Toolbar1.Buttons(3).Enabled = False 'disenable 1:50000 scale maps
       ULGeoX = 0: LRGeoX = 0: ULGeoY = 0: LRGeoY = 0: pixwi = 0: pixhi = 0
    Else '1
       picnam$ = Trim$(cmbMaps.Text)
       picnam0$ = Trim$(cmbMaps.Text)
       GDMDIform.Toolbar1.Buttons(10).Enabled = True  'print maps
       GDMDIform.mnuPrintMap.Enabled = True
       GDMDIform.Toolbar1.Buttons(2).Enabled = True
       If topos Then GDMDIform.Toolbar1.Buttons(3).Enabled = True 'enable 1:50000 scale maps
       GDMDIform.Toolbar1.Buttons(36).Enabled = True
       End If '1
       
    Screen.MousePointer = vbHourglass
    If (Trim$(txtPixWidth.Text) = sEmpty Or Trim$(txtPixHeight.Text) = sEmpty) And _
        picnam$ <> sEmpty And Dir(picnam$) <> sEmpty Then '1
        'determine pixel size of picture
        Picture1.Picture = LoadPicture(cmbMaps.Text)
        Picture1.AutoSize = True
        Picture1.BorderStyle = 0
        txtPixWidth.Text = str(Picture1.Width) / Screen.TwipsPerPixelX
        txtPixHeight.Text = str(Picture1.Height) / Screen.TwipsPerPixelY
        'now release picture memory
        Picture1.Picture = LoadPicture(sEmpty)
        End If '1
        
    'reset coordinates
    'first check
     If val(JustConvertDegToNumber(txtULGeoX.Text)) > val(JustConvertDegToNumber(txtLRGeoX.Text)) Then
        Call MsgBox("The UL X Geo coordinate must be less then the LR X Geo coordinate." _
                    & vbCrLf & "" _
                    & vbCrLf & "(For Western Hemisphere, use negative longitudes;" _
                    & vbCrLf & "For Eastern Hemisphere, use positive longitudes)" _
                    , vbExclamation, "X Geo LImits error")
        Exit Sub
        End If
        
     If val(JustConvertDegToNumber(txtULGeoY.Text)) < val(JustConvertDegToNumber(txtLRGeoY.Text)) Then
        Call MsgBox("The UL Y Geo coordinate must be greater then the LR Y Geo coordinate." _
                    , vbExclamation, "Y Geo LImits error")
        Exit Sub
        End If
        
     pixwi = val(txtPixWidth.Text)
     pixhi = val(txtPixHeight.Text)
     pixwi0 = pixwi
     pixhi0 = pixhi
     ULGeoX = val(JustConvertDegToNumber(txtULGeoX.Text)) 'record coordinate boundary values
     ULGeoY = val(JustConvertDegToNumber(txtULGeoY.Text))
     LRGeoX = val(JustConvertDegToNumber(txtLRGeoX.Text))
     LRGeoY = val(JustConvertDegToNumber(txtLRGeoY.Text))
     x10 = ULGeoX
     y10 = ULGeoY
     x20 = LRGeoX
     y20 = LRGeoY
    
    If AddedMaps Then 'didn't save the added maps yet '1
       resp = MsgBox("You added geologic maps but haven't saved them!" & vbLf & _
                     "Do you want to save the new maps?", vbQuestion + vbYesNoCancel, "MapDigitizer")
       If resp = vbYes Then '2
          cmdSaveMaps_Click
       ElseIf resp = vbCancel Then 'cancel the unload '2
          Cancel = True
          Exit Sub
          End If '2
          
       End If '1
    
''--------------------Other Paths--------------------
'      'check each inputed values
       dirNewDTM = txtNewDTM
       If Trim$(txtNewDTM) = sEmpty Then
       Else
          If Dir(txtNewDTM & "\Boundaries.txt") <> sEmpty Then
             UsingNewDTM = True
          Else
             UsingNewDTM = False
             End If
          End If
'
'      dbdir1 = txtdb1
'      If Trim$(txtdb1) = sEmpty Then '1
'         'so disenable buttons and menus pretaining to MS ACCESS
'         GDMDIform.mnuAccess.Enabled = False
'         GDMDIform.Toolbar1.Buttons(11).Enabled = False
'         acc = False
'       Else '1
'         If Dir(txtdb1 + "\Pal_pr.mdb") = gsEmpty Then '2
'            'so disenable buttons and menus pretaining to MS ACCESS
'            GDMDIform.mnuAccess.Enabled = False
'            GDMDIform.Toolbar1.Buttons(11).Enabled = False
'            End If '2
'          End If '1
'
'      If linked = False Or Trim$(txtdb2) = sEmpty Then '1
'         If linked Then 'user erased path to database so close it '2
'            CloseDatabase
'            End If '2
'         dbdir2 = txtdb2
'         'so disenable database buttons and menus
'         GDMDIform.mnuAccess.Enabled = False
'         If Not linkedOld Then '2
'            GDMDIform.mnuAddScannedFiles.Enabled = False
'            GDMDIform.mnuUpdateLink.Enabled = False
'            GDMDIform.retrieveinfofm.Enabled = False
'            GDMDIform.inputinfofm.Enabled = False
'            For i& = 15 To 33
'              GDMDIform.Toolbar1.Buttons(i&).Enabled = False
'            Next i&
'            End If '2
'
'      Else '1
'         If dbdir2 <> txtdb2 Then '2
'            response = MsgBox("Database already opened!" & vbLf & _
'            "Close it and reopen it using the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
'            If response = vbYes Then '3
'               dbdir2 = txtdb2
'               LinkTables
'               End If '3
'            End If '2
'          End If '1
'
'      If linkedOld = False Or Trim$(txtdb3) = sEmpty Then '1
'         If linkedOld Then 'user erased path to old database so close it '2
'            CloseDatabaseOld
'            If linkedpiv Then CloseDatabasepiv 'also close clone database
'            GDMDIform.Toolbar1.Buttons(12).Enabled = False
'            GDMDIform.mnuEditScannedDB.Enabled = False
'            GDMDIform.mnuLocations.Enabled = False
'            If Not linked Then 'disenable search menus and buttons '3
'               GDMDIform.mnuAddScannedFiles.Enabled = False
'               GDMDIform.mnuUpdateLink.Enabled = False
'               GDMDIform.retrieveinfofm.Enabled = False
'               GDMDIform.inputinfofm.Enabled = False
'               For i& = 15 To 33
'                 GDMDIform.Toolbar1.Buttons(i&).Enabled = False
'               Next i&
'               End If '3
'            End If '2
'
'         NEDdir = txtdb3
'
'      Else '1
'         If NEDdir <> txtdb3 And SearchDBs% <> 2 Then '2
'            response = MsgBox("Scanned Database already opened!" & vbLf & _
'            "Close it and reopen it using the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
'            If response = vbYes Then '3
'               NEDdir = txtdb3
'               LinkTablesOld
'               LinkDBpiv
'               End If '3
'           End If '2
'         End If '1
'
'      If SearchDBs% = 0 Then
'         If linked And linkedOld Then
'            SearchDBs% = 1
'         ElseIf linked And Not linkedOld Then
'            SearchDBs% = 2
'         ElseIf Not linked And linkedOld Then
'            SearchDBs% = 3
'            End If
'         End If

      If UseNewDTM% = 1 And basedtm% = 0 Then
        'initiate base DTM for reading
        ier = OpenCloseBaseDTM(0)
        End If
         
      If DTMtype = 2 Then 'using JKH's DTM '1
      
        If heights = False Or ASTERbilOpen Or Trim$(txtdtm) = sEmpty Then '2
           NEDdir = txtdtm
           
           If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
              If Dir(dtmdir + "\dtm-map.loc") <> gsEmpty Then
                 heights = True
                 If DTMtype = 0 Then DTMtype = 2
                 If Not JKHDTM Then InitializeDTM
              Else
                 If Dir(NEDdir + "\z000000.hgt") <> gsEmpty Then
                    heights = True
                    If DTMtype = 0 Then DTMtype = 2
                    End If
                 End If
              
              
              If ASTERbilOpen Then
                'switched from aster, so close it
                If ASTERfilnum > 0 Then Close #ASTERfilnum
                ASTERbilOpen = False
                End If

              'allow replacement of zero heights with DTM heights
              If PicSum Then GDReportfrm.chkGL.Enabled = True
           Else '3
              If DTMtype = 0 Then heights = False
              'don't allow replacement of zero heights with DTM heights
              If PicSum Then GDReportfrm.chkGL.Enabled = True
              End If '3
        Else '2
           If NEDdir <> txtdtm Then '3
              response = MsgBox("DTM already found!" & vbLf & _
              "Read it from the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
              If response = vbYes Then '4
                 NEDdir = txtdtm
                 If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then '4.5
                    If Dir(dtmdir + "\dtm-map.loc") <> gsEmpty Then '5
                       heights = True
                       If Not JKHDTM Then InitializeDTM
                       If DTMtype = 0 Then DTMtype = 2
                       'allow replacement of zero heights with DTM heights
                    Else '5
                         If DTMtype = 0 Then heights = False
                         'don't allow replacement of zero heights with DTM heights
                         If PicSum Then GDReportfrm.chkGL.Enabled = True
                         End If '5
                  Else '4.5
                     If Dir(NEDdir & "\z000000.hgt") <> gsEmpty Then
                        heights = True
                        If DTMtype = 0 Then DTMtype = 2
                     Else '5
                         If DTMtype = 0 Then heights = False
                         'don't allow replacement of zero heights with DTM heights
                         If PicSum Then GDReportfrm.chkGL.Enabled = True
                         End If '5
                        
                      End If '4.5
                    If PicSum Then GDReportfrm.chkGL.Enabled = True
                 End If '4
               End If '3
            End If '2
           
      ElseIf DTMtype = 1 Then 'using ASTER DTM '1
      
        'using ASTER
        If heights = False Or Trim$(txtAster) = sEmpty Then '2
           ASTERdir = txtAster
           If Dir(ASTERdir + "\N31E035.bil") <> gsEmpty Then '3
              heights = True
              If DTMtype = 0 Then DTMtype = 1
              'allow replacement of zero heights with DTM heights
              If PicSum Then GDReportfrm.chkGL.Enabled = True
           Else '3
              If DTMtype = 0 Then heights = False
              'don't allow replacement of zero heights with DTM heights
              If PicSum Then GDReportfrm.chkGL.Enabled = True
              End If '3
        Else '2
           If ASTERdir <> txtAster Then '3
              response = MsgBox("ASTER DTM already found!" & vbLf & _
              "Read it from the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
              If response = vbYes Then '4
                 ASTERdir = txtAster
                 If Dir(ASTERdir + "\N31E035.bil") <> gsEmpty Then '5
                    heights = True
                    If DTMtype = 0 Then DTMtype = 1
                    'allow replacement of zero heights with DTM heights
                    If PicSum Then GDReportfrm.chkGL.Enabled = True
                 Else '5
                    If DTMtype = 0 Then heights = False
                    'don't allow replacement of zero heights with DTM heights
                    If PicSum Then GDReportfrm.chkGL.Enabled = True
                    End If '5
                 End If '4
               End If '3
             End If '2
           End If '1
         
     If ((Mid$(LCase(lblX), 1, 3) <> "itm" And Mid$(LCase(LblY), 1, 3) <> "itm") And _
         (Mid$(LCase(lblX), 1, 3) <> "lon" And Mid$(LCase(LblY), 1, 3) <> "lat")) Then
         heights = False 'utm coord conversion not yet supported
         End If
      
      If JKHDTM Then
         If val(XStepITM) = 0 Then XStepITM = 25
         If val(YStepITM) = 0 Then YStepITM = 30
      Else
         If val(XStepITM) = 0 Then XStepITM = 30
         If val(YStepITM) = 0 Then YStepITM = 30
         End If
          
      If val(XStepDTM) = 0 Then XStepDTM = 1#  '8.33333333333333E-04 / 3#
      If val(YStepDTM) = 0 Then YStepDTM = 1#  '8.33333333333333E-04 / 3#
      
      txtDTMitmx = XStepITM
      txtDTMitmy = YStepITM
      txtDTMlon = XStepDTM
      txtDTMlat = YStepDTM
      txtAzi = HalfAzi
      txtStepAzi = StepAzi
      txtaprn = Apprn
      
'      If topos = False Or Trim$(txttopo) = sEmpty Then
'         topodir = txttopo
'         If Dir(topodir + "\cli0707.bmp") <> gsEmpty Then
'            topos = True
'            GDMDIform.Toolbar1.Buttons(3).Enabled = True
'         Else
'            topos = False
'            GDMDIform.Toolbar1.Buttons(3).Enabled = False
'            End If
'      Else
'         If topodir <> txttopo Then
'         response = MsgBox("Topo maps already found!" & vbLf & _
'            "Open them from the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
'            If response = vbYes Then
'               topodir = txttopo
'               If Dir(topodir + "\cli0707.bmp") <> gsEmpty Then
'                  topos = True
'               Else
'                  topos = False
'                  End If
'             End If
'          End If
'      End If
               
'     If arcs = False Then
'        arcdir = txtarc
'        If Dir(arcdir + "\ArcMap.exe") <> gsEmpty Then
'           arcs = True
'        Else
'           arcs = False
'           End If
'      Else
'         If arcdir <> txtarc Then
'         response = MsgBox("ArcMap already found!" & vbLf & _
'            "Open it from the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
'            If response = vbYes Then
'               arcdir = txtarc
'               If Dir(arcdir + "\ArcMap.exe") <> gsEmpty Then
'                  arcs = True
'               Else
'                  arcs = False
'                  End If
'             End If
'          End If
'      End If
'      mxddir = txtmxd
      
     If google = False Then
        googledir = txtGoogle
        If Trim$(kmldir) = sEmpty Then
           kmldir = direct$
        Else
           kmldir = txtkml
           End If
        If Dir(googledir + "\googleearth.exe") <> gsEmpty Then
           google = True
        Else
           google = False
           End If
      Else
         kmldir = txtkml.Text
         If Trim$(kmldir) = sEmpty Then kmldir = direct$
         If googledir <> txtGoogle Then
         response = MsgBox("Google Earth already found!" & vbLf & _
            "Open it from the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
            If response = vbYes Then
               googledir = txtGoogle
               If Dir(googledir + "\googleearth.exe") <> gsEmpty Then
                  google = True
               Else
                  google = False
                  End If
             End If
          End If
      End If
      
dskacc:
'     If acc = False Or Trim$(txtacc) = sEmpty Then
'        accdir = txtacc
'        If Dir(accdir + "\Msaccess.exe") <> gsEmpty Then
'           If linked Then
'              If Dir(dbdir1 + "\Pal_pr.mdb") <> gsEmpty And ActivatedVersion Then
'                acc = True
'                GDMDIform.Toolbar1.Buttons(11).Enabled = True
'                GDMDIform.mnuAccess.Enabled = True
'                End If
'              End If
'        Else
'           acc = False
'           GDMDIform.Toolbar1.Buttons(11).Enabled = False
'           GDMDIform.mnuAccess.Enabled = False
'           End If
'      Else
'         If accdir <> txtacc Then
'         response = MsgBox("MS Access already found!" & vbLf & _
'            "Open it from the new path?", vbExclamation + vbYesNoCancel, "MapDigitizer")
'            If response = vbYes Then
'               accdir = txtacc
'               If Dir(accdir + "\Msaccess.exe") <> gsEmpty And ActivatedVersion Then
'                  If Dir(dbdir1 + "\Pal_pr.mdb") <> gsEmpty Then
'                     acc = True
'                     End If
'               Else
'                  acc = False
'                  End If
'             End If
'          End If
'      End If
     
'     '**********Link to Paleontological Database**********
'     'Attempt to link tables in paleontolgical database
'     'residing on the server using the newly inputed path
'     If linked = False And dbdir2 <> sEmpty Then
'        LinkTables
'        If linked Then 'try to enable the access button again
'           GoTo dskacc
'           End If
'        End If
'     '******************************************
'
'     '**********Attempt Link to Scanned Access Paleontological Database**********
'     'Attempt to link tables in paleontolgical database
'     'residing on the server using the newly inputed path
'     If linkedOld = False And NEDdir <> sEmpty And SearchDBs% <> 2 Then
'        LinkTablesOld
'        LinkDBpiv
'        End If
'     '******************************************
     
     Screen.MousePointer = vbDefault
     
     'display any error messages in paths
     ShowError
           
'----------------Other Settings-------------------
     
     numMaxHighlight& = val(txtMaxHighlight.Text) 'maximum records to plot
     If numMaxHighlight& > 32768 Then
        txtMaxHighlight.Text = "32768"
        numMaxHighlight& = 32768
        End If
     SaveClose% = 0 'save settings to hard disk flag
     If chkSave.value = vbChecked Then
        SaveClose% = 1
        End If
     If chkSave_xyz.value = vbChecked Then
        Save_xyz% = 1
        End If
     IgnoreAutoRedrawError% = 0
     If chkAutoRedraw.value = vbChecked Then
        IgnoreAutoRedrawError% = 1
        End If
        
'     If SearchDBs% = 0 And (linked Or linkedOld) Then
'        Screen.MousePointer = vbDefault
'        resp = MsgBox("You haven't yet defined which database to search over!" & vbLf & _
'               "If you don't define that setting you won't be able to search!" & vbLf & vbLf & _
'               "Do you want to try again?", _
'               vbExclamation + vbYesNoCancel, "MapDigitizer")
'        If resp = vbYes Then
'           tabOptions.Tab = 7
'           Cancel = True
'           Exit Sub
'           End If
'        End If
        
      UseNewDTM% = 0
      If chkNewDTM.value = vbChecked Then
         UseNewDTM% = 1
         UsingNewDTM = True
         ier = OpenCloseBaseDTM(0)
         End If
     'replace null ground levels with DTM heights parameters
'     nWellCheck% = 0
'     nOtherCheck% = 0
'     ReplaceWellZ = False
'     ReplaceOtherZ = False

'     If chkWellsReplace.Value = vbChecked Then
'        nWellCheck% = 1
'        ReplaceWellZ = True
'        End If
'     If chkOtherReplace.Value = vbChecked Then
'        nOtherCheck% = 1
'        ReplaceOtherZ = True
'        End If
           
   NEDdir = txtdtm.Text
   If Dir(dtmdir & "\dtm-map.loc") <> gsEmpty Then
   Else
      If Dir(direct$ & "\dtm\dtm-map.loc") <> gsEmpty Then
         dtmdir = direct$ & "\dtm" 'hidden JKH dtm directory
         End If
      End If
         
   If chkSave.value = vbChecked Then 'save settings onto the hard disk
      '******update the directory paths********
        
        If infonum& > 0 Then
           Close #infonum&
           infonum& = 0
           End If
           
      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Output As #infonum&
      Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
      Write #infonum&, dirNewDTM
      Write #infonum&, val(txtEraserBrushSize)
      Write #infonum&, NEDdir
      Write #infonum&, dtmdir
      Write #infonum&, ChainCodeMethod
      Write #infonum&, val(txtDistContour), val(txtDistLines), val(txtSensitivity), val(cmbContour.Text) ' arcdir, mxddir
      Write #infonum&, PointCenterClick
      Write #infonum&, picnam$
      Write #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
      Write #infonum&, ReportPaths&, val(txtDistPixelSearch), numMaxHighlight&, Save_xyz%
      Write #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Write #infonum&, IgnoreAutoRedrawError%
      Write #infonum&, UseNewDTM%, nOtherCheck%
      Write #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
      Write #infonum&, val(txtGridX), val(txtGridY) 'NX_CALDAT, NY_CALDAT
      Write #infonum&, RSMethod0, RSMethod1, RSMethod2
      Write #infonum&, val(txtULPixX), val(txtULPixY), val(txtLRPixX), val(txtLRPixY), val(JustConvertDegToNumber(txtLRGridX)), val(JustConvertDegToNumber(txtLRGridY)), val(JustConvertDegToNumber(txtULGridX)), val(JustConvertDegToNumber(txtULGridY))
      Write #infonum&, val(txtDTMitmx), val(txtDTMitmy), val(txtDTMlon), val(txtDTMlat), val(txtAzi), val(txtStepAzi), val(txtaprn), HeightPrecision, val(txtCustom)
      Close #infonum&
      
      'now reload values into their variables
      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Input As #infonum&
      Input #infonum&, doclin$
      Input #infonum&, dirNewDTMt
      Input #infonum&, MinDigiEraserBrushSize
      Input #infonum&, NEDdirt
      Input #infonum&, dtmdirt
      Input #infonum&, ChainCodeMethod
      Input #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
      Input #infonum&, PointCenterClick
      Input #infonum&, picnamt$
      Input #infonum&, LblXt, LblYt, x1t, x2t, y1t, y2t, pixwit, pixhit, MapUnitst
      Input #infonum&, ReportPathst&, SearchPixelst%, numMaxHighlightt&, Save_xyzt%
      Input #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Input #infonum&, IgnoreAutoRedrawError%
      Input #infonum&, nWellCheckt%, nOtherCheckt%
      Input #infonum&, googledirt, URL_OutCropt, URL_Wellt, kmldirt, ASTERdirt, DTMtypet
      Input #infonum&, NX_CALDAT, NY_CALDAT
      Input #infonum&, RSMethod0, RSMethod1, RSMethod2
      Input #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
      Input #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
      Close #infonum&
      
      
'      'resave tif file parameters if flagged
'      If frmTifViewer.Enabled = True Then
'        filout% = FreeFile
'        Open direct$ & "\gdb_tif.sav" For Output As #filout%
'        Print #filout%, tifDir$
'        Print #filout%, tifViewerDir$
'        Print #filout%, tifCommandLine$
'        Close #filout%
'        End If
      
   Else 'check if settings changed from stored values and ask for verification of closing without saving
   
        If infonum& > 0 Then
           Close #infonum&
           infonum& = 0
           End If
      
      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Input As #infonum&
      Input #infonum&, doclin$
      Input #infonum&, dirNewDTMt
      Input #infonum&, MinDigiEraserBrushSizet
      Input #infonum&, NEDdirt
      Input #infonum&, dtmdirt
      Input #infonum&, ChainCodeMethodt
      Input #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
      Input #infonum&, PointCenterClickt
      Input #infonum&, picnamt$
      Input #infonum&, LblXt, LblYt, x1t, x2t, y1t, y2t, pixwit, pixhit, MapUnitst
      Input #infonum&, ReportPathst&, SearchPixelst%, numMaxHighlightt&, Save_xyzt%
      Input #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Input #infonum&, IgnoreAutoRedrawError%
      Input #infonum&, nWellCheckt%, nOtherCheckt%
      Input #infonum&, googledirt, URL_OutCropt, URL_Wellt, kmldirt, ASTERdirt, DTMtypet
      Input #infonum&, NX_CALDAT, NY_CALDAT
      Input #infonum&, RSMethod0, RSMethod1, RSMethod2
      Input #infonum&, ULPixXt, ULPixYt, LRPixXt, LRPixYt, LRGridXt, LRGridYt, ULGridXt, ULGridYt
      Input #infonum&, XStepITMt, YStepITMt, XStepDTMt, YStepDTMt, HalfAzit, StepAzit, Apprnt, HeightPrecisiont, DigiConvertToMeterst
      Close #infonum&
      
      If ChainCodeMethod = 0 Then
         optFreeman.value = True
      ElseIf ChainCodeMethod = 1 Then
         optBug.value = True
         End If
      
      If dirNewDTMt <> dirNewDTM Or MinDigiEraserBrushSizet <> MinDigiEraserBrushSize Or _
         NEDdirt <> NEDdir Or _
         dtmdirt <> dtmdir Or _
         ChainCodeMethodt <> ChainCodeMethod Or _
         PointCenterClickt <> PointCenterClick Or _
         XStepITMt <> XStepITM Or picnamt$ <> picnam$ Or _
         YStepITMt <> YStepITM Or XStepDTMt <> XStepDTM Or _
         YStepDTMt <> YStepDTM Or HalfAzit <> HalfAzi Or StepAzit <> StepAzi Or _
         Apprnt <> Apprn Or HeightPrecisiont <> HeightPrecision Or _
         DigiConvertToMeterst <> DigiConvertToMeters Or _
         MapUnitst <> MapUnits Or _
         googledirt <> googledir Or _
         URL_OutCropt <> URL_OutCrop Or URL_Wellt <> URL_Well Or kmldirt <> kmldir Or _
         tLblXt <> lblX Or LblYt <> LblY Or x1t <> ULGeoX Or x2t <> LRGeoX Or _
         y1t <> ULGeoY Or y2t <> LRGeoY Or pixwit <> pixwi Or pixhit <> pixhi Or _
         ReportPathst& <> ReportPaths& Or SearchPixelst% <> DigiSearchRegion Or _
         numMaxHighlightt& <> numMaxHighlight& Or Save_xyzt% <> Save_xyz% Or _
         UseNewDTMt% <> UseNewDTM% Or nOtherCheckt% <> nOtherCheck% Or ASTERdirt <> ASTERdir Or _
         ULPixXt <> ULPixX Or ULPixYt <> ULPixY Or LRPixXt <> LRPixX Or LRGridXt <> LRGridX Or _
         LRGridYt <> LRGridY Or ULGridXt <> ULGridX Or ULGridYt <> ULGridY Or _
         DTMtypet <> DTMtype Then
         'settings were changed, so warn the user
         Screen.MousePointer = vbDefault
         resp = MsgBox("Your settings have changed, do you wish to record them?", _
                vbQuestion + vbYesNoCancel, "MapDigitizer")
         If resp = vbYes Then
         
            If infonum& > 0 Then
               Close #infonum&
               infonum& = 0
               End If
               
            infonum& = FreeFile
            Open direct$ + "\gdbinfo.sav" For Output As #infonum&
            Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
            Write #infonum&, dirNewDTM
            Write #infonum&, val(txtEraserBrushSize)
            Write #infonum&, NEDdir
            Write #infonum&, dtmdir
            Write #infonum&, ChainCodeMethod
            Write #infonum&, val(txtDistContour), val(txtDistLines), val(txtSensitivity), val(cmbContour.Text) ' arcdir, mxddir
            Write #infonum&, PointCenterClick
            Write #infonum&, picnam$
            Write #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
            Write #infonum&, ReportPaths&, val(txtDistPixelSearch), numMaxHighlight&, Save_xyz%
            Write #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
            Write #infonum&, IgnoreAutoRedrawError%
            Write #infonum&, UseNewDTM%, nOtherCheck%
            Write #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
            Write #infonum&, NX_CALDAT, NY_CALDAT
            Write #infonum&, RSMethod0, RSMethod1, RSMethod2
            Write #infonum&, val(txtULPixX), val(txtULPixY), val(txtLRPixX), val(txtLRPixY), val(JustConvertDegToNumber(txtLRGridX)), val(JustConvertDegToNumber(txtLRGridY)), val(JustConvertDegToNumber(txtULGridX)), val(JustConvertDegToNumber(txtULGridY))
            Write #infonum&, val(txtITMx), val(txtITMy), val(txtDTMlon), val(txtDTMlat), val(txtAzi), val(txtStepAzi), val(txtaprn), HeightPrecision, val(txtCustom)
            Close #infonum&
            End If
         End If
         
'         'check for changes in the tif viewer files
'         filin% = FreeFile
'         Open direct$ & "\gdb_tif.sav" For Input As #filin%
'         Line Input #filin%, tifDirf$
'         Line Input #filin%, tifViewerDirf$
'         Line Input #filin%, tifCommandLinef$
'         Close #filin%
'
'         If tifDirf$ <> tifDir$ Or tifViewerDirf$ <> tifViewerDir$ Or tifCommandLinef$ <> tifCommandLine$ Then
'            tabOptions.Tab = 4
'            response = MsgBox("Do you want to save your changes to the tif file/viewer parameters?", vbYesNoCancel + vbQuestion, "MapDigitizer")
'            If response = vbYes Then
'               'save the changes to the tif paths, etc
'               filin% = FreeFile
'               Open direct$ & "\gdb_tif.sav" For Output As #filin%
'               Print #filin%, tifDir$
'               Print #filin%, tifViewerDir$
'               Print #filin%, tifCommandLine$
'               Close #filin%
'               End If
'            End If
      
      End If
      
    If numcpt = 0 And LineElevColors& = 1 Then
         
         myfile = Dir(App.Path & "\rainbow.cpt")
         If myfile = sEmpty Then
            GoTo op850
            End If
         
         '-----------------------load color palette--------------------------
         numpercent = -1
         numloop% = 0
         nowread = True
         num% = 0
    
         ier = 0
         
         ReDim cpt(3, 0)
         
         filenum% = FreeFile
         Open App.Path & "\rainbow.cpt" For Input As #filenum%
         
         Do Until EOF(filenum%)
            Line Input #filenum%, doclin$
            colorattributes = Split(doclin$, " ")
            For i = 0 To 10
              cc$ = colorattributes(i)
              If Trim$(cc$) <> vbNullString Then
                 If numloop% = 0 Then
                    If val(cc$) >= numpercent Then
                        num% = val(cc$)
                        
                        If num% - 1 > UBound(cpt, 2) Then
                           ReDim Preserve cpt(3, UBound(cpt, 2) + 1)
                           End If
                        
                        cpt(0, num% - 1) = val(cc$)
                        numloop% = 1
                        numpercent = val(cc$)
                        nowread = True
                    Else
                        nowread = False
                        End If
                 ElseIf numloop% = 1 Then
                    If nowread Then cpt(1, num% - 1) = val(cc$)
                    numloop% = 2
                 ElseIf numloop% = 2 Then
                    If nowread Then cpt(2, num% - 1) = val(cc$)
                    numloop% = 3
                 ElseIf numloop% = 3 Then
                    If nowread Then cpt(3, num% - 1) = val(cc$)
                    numloop% = 0
                    nowread = False
                    Exit For
                    End If
                 End If
                 
                 numcpt = num%
                 
            Next i
         Loop
         Close #filenum%
         End If

         
op850:
         
    If Dir(picnam$) = sEmpty Then
       'try adding app.path
       If Dir(App.Path & "\" & picnam$) <> sEmpty Then
          picnam$ = App.Path & "\" & picnam$
       Else
          buttonstate&(2) = 0
          GDMDIform.Toolbar1.Buttons(2).Enabled = False
          GDMDIform.StatusBar1.Panels(1).Text = "No stored map file could be found, define one using the ""Options"" dialog (click first button on toolbar)..."
          End If
       End If
       
op900:
   Screen.MousePointer = vbDefault
   Set GDOptionsfrm = Nothing
   OptionsVis = False
   GDMDIform.Toolbar1.Buttons(1).value = tbrUnpressed
   buttonstate&(1) = 0
     
   Exit Sub
     
errhand:
   Screen.MousePointer = vbDefault
   Select Case Err.Number
       Case 62
            'eof passed 'assume this is due to new version
            Resume Next
       Case 52
            MsgBox "One of the paths you inputed can't be accessed!" & vbLf & _
                   "Reenter the path, or choose a different one.", vbExclamation + vbOKOnly, "MapDigitizer"
            Resume Next
       Case Else
            MsgBox "Encountered error #: " & Err.Number & vbLf & _
                   Err.Description & vbLf & _
                   sEmpty, vbCritical + vbOKOnly, "MapDigitizer"
            Resume Next
   End Select
   
   If infonum& > 0 Then
      Close #infonum&
      infonum& = 0
      End If
   
End Sub

'Private Sub optActive_Click()
'  If linked Then
'    SearchDBs0% = SearchDBs%
'    SearchDBs% = 2
'    optActive.value = True
'    If SearchVis And SearchDBs% <> SearchDBs0% Then
'        resp = MsgBox("In order to activate this change," & vbLf & _
'               "the search wizard must be reopened." & vbLf & vbLf & _
'               "Reopen the search wizard now?", vbQuestion + vbYesNo, "MapDigitizer")
'        If resp = vbYes Then
'           tabsearch% = GDSearchfrm.tbSearch.Tab 'save the current tab
'           CloseSearchWizard = True
'           Unload GDSearchfrm 'unload the form
'           GDSearchfrm.Visible = True 'reload it
'           GDSearchfrm.tbSearch.Tab = tabsearch% 'restore the tab
'           BringWindowToTop (GDOptionsfrm.hWnd)
'           End If
'       End If
'  Else
'     'attempt to link it
'     '**********Link to Paleontological Database**********
'     'Attempt to link tables in paleontolgical database
'     'residing on the server using the newly inputed path
'     If dbdir1 = sEmpty And Trim$(txtdb1) <> sEmpty Then
'        'try this value as dbdir1
'        dbdir1 = Trim$(txtdb1)
'        End If
'     If dbdir2 = sEmpty And Trim$(txtdb2) <> sEmpty Then
'        'try this value as dbdir2
'        dbdir2 = Trim$(txtdb2)
'        End If
'     If linked = False And dbdir2 <> sEmpty Then
'        LinkTables
'        If linked Then 'try to enable the access button/menu
'
'            accdir = txtacc
'            If Dir(accdir + "\Msaccess.exe") <> gsEmpty Then
'               If Dir(dbdir1 + "\Pal_pr.mdb") <> gsEmpty Then
'                  acc = True
'                  GDMDIform.Toolbar1.Buttons(11).Enabled = True
'                  GDMDIform.mnuAccess.Enabled = True
'                  End If
'            Else
'               acc = False
'               End If
'
'           End If
'        End If
'     '******************************************
'
'
'     If Not linked Then
'        optActive.value = False
'        GDOptionsfrm.tabOptions.Tab = 7
'        MsgBox "You have attempted to enable searches over the active database" & vbLf & _
'               "However, the path to that databases is not correct." & vbLf & _
'               "Define that paths, and then return to this option.", _
'               vbExclamation + vbOKOnly, "MapDigitizer"
'     Else
'        SearchDBs% = 2
'        End If
'     End If
'End Sub

'Private Sub optAll_Click()
'  If linked And linkedOld Then
'     SearchDBs0% = SearchDBs%
'     SearchDBs% = 1
'     optAll.value = True
'     If SearchVis And SearchDBs% <> SearchDBs0% Then
'        resp = MsgBox("In order to activate this change," & vbLf & _
'               "the search wizard must be reopened." & vbLf & vbLf & _
'               "Reopen the search wizard now?", vbQuestion + vbYesNo, "MapDigitizer")
'        If resp = vbYes Then
'           tabsearch% = GDSearchfrm.tbSearch.Tab 'save the current tab
'           CloseSearchWizard = True
'           Unload GDSearchfrm 'unload the form
'           GDSearchfrm.Visible = True 'reload it
'           GDSearchfrm.tbSearch.Tab = tabsearch% 'restore the tab
'           BringWindowToTop (GDOptionsfrm.hWnd)
'           End If
'        End If
'  Else
'     'find out which one is not linked, and attempt to link it
'
'     '**********Link to Paleontological Database**********
'     'Attempt to link tables in paleontolgical database
'     'residing on the server using the newly inputed path
'     If dbdir1 = sEmpty And Trim$(txtdb1) <> sEmpty Then
'        'try this value as dbdir1
'        dbdir1 = Trim$(txtdb1)
'        End If
'     If dbdir2 = sEmpty And Trim$(txtdb2) <> sEmpty Then
'        'try this value as dbdir2
'        dbdir2 = Trim$(txtdb2)
'        End If
'     If linked = False And dbdir2 <> sEmpty Then
'        LinkTables
'        If linked Then 'try to enable the access button/menu
'
'            accdir = txtacc
'            If Dir(accdir + "\Msaccess.exe") <> gsEmpty Then
'               If Dir(dbdir1 + "\Pal_pr.mdb") <> gsEmpty Then
'                  acc = True
'                  GDMDIform.Toolbar1.Buttons(11).Enabled = True
'                  GDMDIform.mnuAccess.Enabled = True
'                  End If
'            Else
'               acc = False
'               End If
'
'           End If
'        End If
'     '******************************************
'
'     '**********Attempt Link to Scanned Access Paleontological Database**********
'     'Attempt to link tables in paleontolgical database
'     'residing on the server using the newly inputed path
'     If Trim$(txtdb3) <> sEmpty And NEDdir = sEmpty Then
'        'try txtdb3 as the new NEDdir
'        NEDdir = Trim$(txtdb3)
'        End If
'     If linkedOld = False And NEDdir <> sEmpty Then
'        LinkTablesOld
'        LinkDBpiv
'        End If
'     '******************************************
'
'     If Not linked Or Not linkedOld Then
'        optAll.value = False
'        GDOptionsfrm.tabOptions.Tab = 7
'        MsgBox "You have attempted to enable searches over both the" & vbLf & _
'               "Active and Scanned (inactive) database." & vbLf & _
'               "However, the path to one/both of those databases is (are) incorrect." & vbLf & _
'               "Define those paths, and then return to this option.", _
'               vbExclamation + vbOKOnly, "MapDigitizer"
'     Else
'        SearchDBs% = 1
'        End If
'     End If
'End Sub

Private Sub frmhgt_DragDrop(Source As Control, X As Single, Y As Single)
   HeightPrecision = 1
End Sub

Private Sub optAster_Click()
   DTMtype = 1
End Sub

Private Sub optBug_Click()
   If optBug.value = True Then
      ChainCodeMethod = 1
      End If
End Sub

Private Sub optCramer_Click()
   GaussMethod = False
End Sub

Private Sub optdecimeters_Click()
   DigiConvertToMeters = 0.1
   txtconvert = DigiConvertToMeters
   txtCustom.Enabled = False
End Sub

Private Sub optDefault_Click()
   DigiConvertToMeters = 1#
   txtCustom = DigiConvertToMeters
   txtCustom.Enabled = False
End Sub

Private Sub optDouble_Click()
   HeightPrecision = 2
End Sub

Private Sub optDTM_Click()
   DTMtype = 2
End Sub

'Private Sub optInactive_Click()
'  If linkedOld Then
'    SearchDBs0% = SearchDBs%
'    SearchDBs% = 3
'    optInactive.value = True
'    If SearchVis And SearchDBs% <> SearchDBs0% Then
'        resp = MsgBox("In order to activate this change," & vbLf & _
'               "the search wizard must be reopened." & vbLf & vbLf & _
'               "Reopen the search wizard now?", vbQuestion + vbYesNo, "MapDigitizer")
'        If resp = vbYes Then
'           tabsearch% = GDSearchfrm.tbSearch.Tab 'save the current tab
'           CloseSearchWizard = True
'           Unload GDSearchfrm 'unload the form
'           GDSearchfrm.Visible = True 'reload it
'           GDSearchfrm.tbSearch.Tab = tabsearch% 'restore the tab
'           BringWindowToTop (GDOptionsfrm.hWnd)
'           End If
'       End If
'  Else
'     'attempt to link it
'
'     '**********Attempt Link to Scanned Access Paleontological Database**********
'     'Attempt to link tables in paleontolgical database
'     'residing on the server using the newly inputed path
'     If Trim$(txtdb3) <> sEmpty And NEDdir = sEmpty Then
'        'try txtdb3 as the new NEDdir
'        NEDdir = Trim$(txtdb3)
'        End If
'     If linkedOld = False And NEDdir <> sEmpty Then
'        LinkTablesOld
'        LinkDBpiv
'        End If
'     '******************************************
'
'     If Not linkedOld Then
'        optInactive.value = False
'        GDOptionsfrm.tabOptions.Tab = 7
'        MsgBox "You have attempted to activate searches over the old" & vbLf & _
'               "database.  However, the path to the scanned database is not correct." & vbLf & _
'               "Define that path, and then return to this option.", _
'               vbExclamation + vbOKOnly, "MapDigitizer"
'     Else
'        SearchDBs% = 3
'        End If
'     End If
'End Sub

Private Sub optfathoms_Click()
   DigiConvertToMeters = 0.546806652777778  'convert from meters to fathoms
   txtCustom = DigiConvertToMeters
   txtCustom.Enabled = False
End Sub

Private Sub optfeet_Click()
   DigiConvertToMeters = 3.28083991666667
   txtCustom = DigiConvertToMeters
   txtCustom.Enabled = False
End Sub

Private Sub optFloat_Click()
   HeightPrecision = 1
End Sub

Private Sub optFreeman_Click()
   If optFreeman.value = True Then
      ChainCodeMethod = 0
      End If
End Sub

Private Sub optGaussian_Click()
  GaussMethod = True
End Sub

Private Sub optInteger_Click()
   HeightPrecision = 0
End Sub

Private Sub optOther_Click()
   txtCustom.Enabled = True
   txtCustom = 0#
End Sub

