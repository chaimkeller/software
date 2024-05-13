VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form mapdiskDTMfm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set directory locations of data and programs"
   ClientHeight    =   3360
   ClientLeft      =   5385
   ClientTop       =   3060
   ClientWidth     =   6495
   Icon            =   "mapdiskDTMfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   8
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Israel DTM"
      TabPicture(0)   =   "mapdiskDTMfm.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmSource"
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "1 km DEM"
      TabPicture(1)   =   "mapdiskDTMfm.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "30 m DEM"
      TabPicture(2)   =   "mapdiskDTMfm.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame7(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmUSA"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "frmALOS"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "frmALOSfolder"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "TerraExplorer"
      TabPicture(3)   =   "mapdiskDTMfm.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4(1)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "calculation drive"
      TabPicture(4)   =   "mapdiskDTMfm.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3(1)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "mouse movements"
      TabPicture(5)   =   "mapdiskDTMfm.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5(1)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "enabled Window calculations"
      TabPicture(6)   =   "mapdiskDTMfm.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame6(1)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "90 m DEM"
      TabPicture(7)   =   "mapdiskDTMfm.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame7(0)"
      Tab(7).Control(1)=   "frm3AS"
      Tab(7).ControlCount=   2
      Begin VB.Frame frm3AS 
         Caption         =   "30m DTM folder"
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
         Height          =   1455
         Left            =   -70320
         TabIndex        =   62
         Top             =   960
         Width           =   1455
         Begin VB.CommandButton cmdBrowse3AS 
            Caption         =   "Browse"
            Height          =   315
            Left            =   240
            TabIndex        =   64
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txt3AS 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Text            =   "3AS"
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frmALOSfolder 
         Caption         =   "ALOS Folder"
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
         Height          =   975
         Left            =   4680
         TabIndex        =   56
         Top             =   1800
         Width           =   1575
         Begin VB.CommandButton cmdBrowseBIL 
            Caption         =   "Browse"
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtBIL 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   60
            Text            =   "BIL"
            ToolTipText     =   "ALOS BIL folder"
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmALOS 
         Caption         =   "ALOS"
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
         Height          =   975
         Left            =   240
         TabIndex        =   55
         Top             =   1800
         Width           =   4335
         Begin VB.CommandButton cmdBrowseALOS 
            Caption         =   "Browse"
            Height          =   255
            Left            =   3240
            TabIndex        =   59
            Top             =   240
            Width           =   855
         End
         Begin VB.DriveListBox DriveALOS 
            Height          =   315
            Left            =   1560
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label LabelALOSdrive 
            Caption         =   "disk location"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   280
            TabIndex        =   58
            Top             =   280
            Width           =   1095
         End
      End
      Begin VB.Frame frmUSA 
         Caption         =   "SRTM=1 Folder"
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
         Height          =   1095
         Left            =   4680
         TabIndex        =   51
         Top             =   720
         Width           =   1455
         Begin VB.CommandButton cmdBrowseUSA 
            Caption         =   "Browse"
            Height          =   260
            Left            =   240
            TabIndex        =   53
            Top             =   640
            Width           =   975
         End
         Begin VB.TextBox txtUSA 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   52
            Text            =   "USA"
            ToolTipText     =   "SRTM 1 arcsecond directory"
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frmSource 
         Caption         =   "Elevation source for EY maps and calculations"
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
         Height          =   615
         Left            =   -74760
         TabIndex        =   48
         Top             =   2040
         Width           =   5775
         Begin VB.OptionButton optSRTM3 
            Caption         =   "SRTM3 / MERIT"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3200
            TabIndex        =   65
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optALOS 
            Caption         =   "ALOS"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4860
            TabIndex        =   54
            ToolTipText     =   "Use ALOS elevations"
            Top             =   240
            Width           =   800
         End
         Begin VB.OptionButton optSRTM 
            Caption         =   "NED / SRTM1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1680
            TabIndex        =   50
            ToolTipText     =   "Use SRTM elevations"
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optJK 
            Caption         =   "JK's 25-m DTM"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Use JKH's DTM"
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "SRTM3 / MERIT 30m DTM"
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
         Height          =   1455
         Index           =   0
         Left            =   -74640
         TabIndex        =   40
         Top             =   960
         Width           =   4215
         Begin VB.CheckBox chkSRTM30 
            Caption         =   "Use this DTM and not the GTOPO30"
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.DriveListBox Drive5 
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   44
            Top             =   300
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   "&Browse"
            Height          =   315
            Index           =   0
            Left            =   3120
            TabIndex        =   43
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton Option6 
            Caption         =   "This is a CD drive"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   42
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option5 
            Caption         =   "This is a disk drive"
            Height          =   195
            Index           =   0
            Left            =   2040
            TabIndex        =   41
            Top             =   720
            Value           =   -1  'True
            Width           =   1635
         End
         Begin MSComDlg.CommonDialog CommonDialog5 
            Index           =   0
            Left            =   3420
            Top             =   660
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Caption         =   "disk location"
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
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Eretz Yisroel DTM calculations"
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
         Height          =   735
         Index           =   1
         Left            =   -73920
         TabIndex        =   38
         Top             =   1080
         Width           =   4335
         Begin VB.CheckBox Check1 
            Caption         =   "DTM calculations from Windows enabled"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   39
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "mouse movement "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   975
         Index           =   1
         Left            =   -74040
         TabIndex        =   31
         Top             =   960
         Width           =   4335
         Begin VB.TextBox Text1 
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
            Height          =   375
            Index           =   1
            Left            =   1140
            TabIndex        =   35
            Text            =   "1.0"
            Top             =   360
            Width           =   555
         End
         Begin VB.TextBox Text2 
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
            Height          =   375
            Index           =   1
            Left            =   3000
            TabIndex        =   33
            Text            =   "1.0"
            Top             =   360
            Width           =   375
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   375
            Index           =   1
            Left            =   3376
            TabIndex        =   32
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "Text2(1)"
            BuddyDispid     =   196638
            BuddyIndex      =   1
            OrigLeft        =   3600
            OrigTop         =   360
            OrigRight       =   3840
            OrigBottom      =   735
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   34
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "Text1(1)"
            BuddyDispid     =   196637
            BuddyIndex      =   1
            OrigLeft        =   1740
            OrigTop         =   360
            OrigRight       =   1980
            OrigBottom      =   735
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "A*dx1, A="
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
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   400
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "B*dy1, B="
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
            Index           =   1
            Left            =   2100
            TabIndex        =   36
            Top             =   400
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "drive for calculations (ramdrive/hard drive)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Index           =   1
         Left            =   -74040
         TabIndex        =   27
         Top             =   1200
         Width           =   4335
         Begin VB.DriveListBox Drive3 
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   29
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command6 
            Caption         =   "&Browse"
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   28
            Top             =   360
            Width           =   915
         End
         Begin MSComDlg.CommonDialog CommonDialog3 
            Index           =   1
            Left            =   120
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Caption         =   "location"
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
            Index           =   9
            Left            =   180
            TabIndex        =   30
            Top             =   420
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "TerraExplorer"
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
         Height          =   1095
         Index           =   1
         Left            =   -73920
         TabIndex        =   21
         Top             =   1080
         Width           =   4335
         Begin VB.DriveListBox Drive4 
            Height          =   315
            Index           =   1
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Browse"
            Height          =   315
            Index           =   1
            Left            =   3180
            TabIndex        =   23
            Top             =   240
            Width           =   915
         End
         Begin VB.DirListBox Dir1 
            Height          =   315
            Index           =   1
            Left            =   1560
            TabIndex        =   22
            Top             =   600
            Width           =   2355
         End
         Begin MSComDlg.CommonDialog CommonDialog4 
            Index           =   1
            Left            =   3480
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Caption         =   "drive location"
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
            Index           =   8
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "directory name"
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
            Index           =   7
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   1275
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "NED DEM / SRTM1 DTM"
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
         Height          =   1095
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   4335
         Begin VB.OptionButton Option5 
            Caption         =   "This is a disk drive"
            Height          =   195
            Index           =   1
            Left            =   2340
            TabIndex        =   19
            Top             =   720
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton Option6 
            Caption         =   "This is a CD drive"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   18
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton Command7 
            Caption         =   "&Browse"
            Height          =   315
            Index           =   1
            Left            =   3180
            TabIndex        =   17
            Top             =   300
            Width           =   915
         End
         Begin VB.DriveListBox Drive5 
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   16
            Top             =   300
            Width           =   1455
         End
         Begin MSComDlg.CommonDialog CommonDialog5 
            Index           =   1
            Left            =   3420
            Top             =   660
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Caption         =   "disk location"
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
            Index           =   6
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "GTOPO30 / SRTM30 DTM"
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
         Height          =   1455
         Index           =   1
         Left            =   -74040
         TabIndex        =   9
         Top             =   840
         Width           =   4335
         Begin VB.CheckBox chkGTOPO30 
            Caption         =   "Use this DTM and not the SRTM30"
            Height          =   255
            Left            =   840
            TabIndex        =   47
            Top             =   1080
            Width           =   2895
         End
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   13
            Top             =   300
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Browse"
            Height          =   315
            Index           =   1
            Left            =   3180
            TabIndex        =   12
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton Option3 
            Caption         =   "This is a CD drive"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   11
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "This is a disk drive"
            Height          =   195
            Index           =   1
            Left            =   2340
            TabIndex        =   10
            Top             =   720
            Value           =   -1  'True
            Width           =   1635
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Index           =   1
            Left            =   3420
            Top             =   660
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Caption         =   "disk location"
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
            Index           =   5
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Israel DTM"
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
         Height          =   1095
         Index           =   1
         Left            =   -74040
         TabIndex        =   3
         Top             =   840
         Width           =   4335
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   7
            Top             =   300
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Browse"
            Height          =   315
            Index           =   1
            Left            =   3180
            TabIndex        =   6
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "This is a CD drive"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   5
            Top             =   720
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "This is a disk drive"
            Height          =   195
            Index           =   1
            Left            =   2340
            TabIndex        =   4
            Top             =   720
            Width           =   1635
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Index           =   1
            Left            =   3480
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Caption         =   "disk location"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Accept && &Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   3000
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
End
Attribute VB_Name = "mapdiskDTMfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)
   If Check1(1).value = vbChecked Then
      'test if enough RAM memory is available
      RdHalYes = RdHalTrue
      If Not RdHalYes Then
         Check1(1).value = vbUnchecked
         MsgBox "You need two RAM drives with 32MG memory for this option", vbExclamation + vbOKOnly, "Maps & More"
         End If
   Else
      RdHalYes = False
      End If
End Sub

Private Sub chkGTOPO30_Click()
   If chkGTOPO30.value = vbChecked Then
      chkSRTM30.value = vbUnchecked
   Else
      chkSRTM30.value = vbChecked
      End If
      
End Sub

Private Sub chkSRTM30_Click()
   If chkSRTM30.value = vbChecked Then
      chkGTOPO30.value = vbUnchecked
   Else
      chkGTOPO30.value = vbChecked
      End If

End Sub

Private Sub cmdBrowse3AS_Click()
Dim DirPath$

        DirPath$ = BrowseForFolder(mapdiskDTMfm.hwnd, "Choose Directory")
        If Dir(DirPath$, vbDirectory) <> "" Then
           D3ASDir$ = DirPath$ & "\"
           txt3AS = Mid$(D3ASDir$, 4, Len(D3ASDir$) - 4)
           alosdtm = Mid$(D3ASDir$, 1, 1)
           Drive5(0).Drive = alosdtm
           optALOS.Enabled = True
        Else
           ier = -1
           Exit Sub
           End If
End Sub

Private Sub cmdBrowseALOS_Click()

   On Error GoTo cmdBrowseALOS_Click_Error

  CommonDialog5(1).CancelError = True
  CommonDialog5(1).Filter = "ALOS files (*.bil)|*.bil|"
  CommonDialog5(1).FilterIndex = 1
  CommonDialog5(1).FileName = alosdtm + ":\BIL\ALPSMLC30_N032E035_DSM.bil"
  CommonDialog5(1).ShowOpen

   On Error GoTo 0
   Exit Sub

cmdBrowseALOS_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdBrowseALOS_Click of Form mapdiskDTMfm"
End Sub

Private Sub cmdBrowseBIL_Click()
Dim DirPath$

        DirPath$ = BrowseForFolder(mapdiskDTMfm.hwnd, "Choose Directory")
        If Dir(DirPath$, vbDirectory) <> "" Then
           BILDir$ = DirPath$ & "\"
           txtBIL = Mid$(BILDir$, 4, Len(BILDir$) - 4)
           alosdtm = Mid$(BILDir$, 1, 1)
           DriveALOS.Drive = alosdtm
           optALOS.Enabled = True
        Else
           ier = -1
           Exit Sub
           End If
End Sub

Private Sub cmdBrowseUSA_Click()
Dim DirPath$

        DirPath$ = BrowseForFolder(mapdiskDTMfm.hwnd, "Choose Directory")
        If Dir(DirPath$, vbDirectory) <> "" Then
           USADir$ = DirPath$ & "\"
           txtUSA = Mid$(USADir$, 4, Len(USADir$) - 4)
           srtmdtm = Mid$(USADir$, 1, 1)
           Drive5(1).Drive = srtmdtm
           optSRTM.Enabled = True
        Else
           ier = -1
           Exit Sub
           End If

End Sub

Private Sub Command1_Click()
      israeldtm = Drive1(1).Drive
      If chkSRTM30.value = vbChecked Then
         worlddtm = Drive5(0).Drive
      Else
         worlddtm = Drive2(1).Drive
         End If
      ramdrive = Drive3(1).Drive
      srtmdtm = Drive5(1).Drive
      alosdtm = DriveALOS.Drive
      'terradrive = Drive4.Drive
      terradir$ = Dir1(1).List(Dir1(1).ListIndex - 1)
      If israeldtmcd = True Then
         isrealdtmcdnum = 1
      Else
         isrealdtmcdnum = 0
         End If
      If worlddtmcd = True Then
         worlddtmcdnum = 1
      Else
         worlddtmcdnum = 0
         End If
      If srtmdtmcd = True Then
         srtmdtmcdnum = 1
      Else
         srtmdtmcdnum = 0
         End If
      adx1 = Val(Text1(1).Text)
      bdy1 = Val(Text2(1).Text)
      Call Form_QueryUnload(i%, j%)
End Sub

Private Sub Command2_Click()
      israeldtm = Drive1(1).Drive
      israeldtmf = israeldtm
      If chkSRTM30.value = vbChecked Then
         worlddtm = Drive5(0).Drive
      Else
         worlddtm = Drive2(1).Drive
         End If
      ramdrive = Drive3(1).Drive
      srtmdtm = Drive5(1).Drive
      worlddtmf = worlddtm
      ramdrivef = ramdrive
      terradirf$ = terradir$
      If israeldtmcd = True Then
         israeldtmcdf = True
         israeldtmcdnum = 1
         israeldtmcdnumf = 1
      Else
         israeldtmcdf = False
         israeldtmcdnum = 0
         israeldtmcdnumf = 0
         End If
      If worlddtmcd = True Then
         worlddtmcdf = True
         worlddtmcdnum = 1
         worlddtmcdnumf = 1
      Else
         worlddtmcdf = False
         worlddtmcdnum = 0
         worlddtmcdnumf = 0
         End If
      adx1f = Val(Text1(1).Text)
      bdy1f = Val(Text2(1).Text)
      
      mapinfonum% = FreeFile
      Close
      Open drivjk_c$ + "mapcdinfo.sav" For Output As #mapinfonum%
      Print #mapinfonum%, israeldtmf; ","; israeldtmcdnumf
      Print #mapinfonum%, worlddtmf; ","; worlddtmcdnumf
      Print #mapinfonum%, ramdrivef
      Print #mapinfonum%, terradirf$
      Write #mapinfonum%, adx1f, bdy1f
      Write #mapinfonum%, RdHalYes
      Write #mapinfonum%, IsraelDTMsource%
      Close #mapinfonum%
      
      If srtmdtmcd = True Then
         srtmdtmcdnum = 1
      Else
         srtmdtmcdnum = 0
         End If
      mapinfonum% = FreeFile
      Open drivjk_c$ & "mapSRTMinfo.sav" For Output As #mapinfonum%
      Print #mapinfonum%, srtmdtm; ","; srtmdtmcdnum
      Print #mapinfonum%, txtUSA
      Print #mapinfonum%, d3asdtm; ","; Trim$(txt3AS.Text)
      Print #mapinfonum%, alosdtm; ","; Trim$(txtBIL.Text)
      Close #mapinfonum%
      
      Call Form_QueryUnload(i%, j%)

End Sub

Private Sub Command3_Click(Index As Integer)
  On Error GoTo c3error
  CommonDialog1(1).CancelError = True
  CommonDialog1(1).Filter = "dtm-map.loc files (*.loc)|*.loc|"
  CommonDialog1(1).FilterIndex = 1
  CommonDialog1(1).FileName = israeldtm + ":\dtm\dtm-map.loc"
  CommonDialog1(1).ShowOpen
c3error:
  Exit Sub
End Sub

Private Sub Command4_Click(Index As Integer)
  On Error GoTo c4error
  CommonDialog2(1).CancelError = True
  CommonDialog2(1).Filter = "E020N40.GIF files (*.gif)|*.gif|" '"Gt30dem.gif files (*.gif)|*.gif|"
  CommonDialog2(1).FilterIndex = 1
  CommonDialog2(1).FileName = worlddtm + ":\E020N40\E020N40.GIF" 'worlddtm + ":\Gt30dem.gif"
  CommonDialog2(1).ShowOpen
c4error:
  Exit Sub
End Sub

Private Sub Command5_Click(Index As Integer)
  On Error GoTo c5error
  CommonDialog4(1).CancelError = True
  CommonDialog4(1).Filter = "TerraExplorer.exe (*.exe)|*.exe|"
  CommonDialog4(1).FilterIndex = 1
  CommonDialog4(1).FileName = terradir$ + "\TerraExplorer.exe"
  CommonDialog4(1).ShowOpen
  terradir$ = Mid$(CommonDialog4(1).FileName, 1, Len(CommonDialog4(1).FileName) - 18)
c5error:
  Exit Sub
End Sub

Private Sub Command7_Click(Index As Integer)
  On Error GoTo c3error
  CommonDialog5(1).CancelError = True
  CommonDialog5(1).Filter = "srtm 30m files (*.hgt)|*.hgt|"
  CommonDialog5(1).FilterIndex = 1
  'CommonDialog5(1).FileName = srtmdtm + ":\3AS\N31E035.hgt"
  CommonDialog5(1).FileName = "N31E035.hgt"
  CommonDialog5(1).ShowOpen
c3error:
  Exit Sub

End Sub

Private Sub Dir1_Change(Index As Integer)
   On Error GoTo errhand:
   Dir1(1).Path = Drive4(1).Drive    ' When drive changes, set directory path.
   ChDir Drive4(1).Drive
   'Dir1.ListIndex = 0
   'terradir$ = Drive1.Drive + Dir1.List(Dir1.ListIndex)
errhand:
End Sub


Private Sub Drive4_Change(Index As Integer)
   'Dir1.Path = Drive4.Drive    ' When drive changes, set directory path.
   On Error GoTo Drive4_Change_Error

   ChDir Mid$(terradir$, 1, 3)
   ChDrive Drive4(1).Drive
   Dir1(1).Path = Drive4(1).Drive
   For i% = 0 To Dir1(1).ListCount - 1
      If LCase$(Dir1(1).List(i%)) = LCase$(terradir$) Then
         Dir1(1).ListIndex = i% + 1
         Exit For
         End If
   Next i%
   'ChDir Drive4.Drive + "\"
   'Dir1.ListIndex = 0

   On Error GoTo 0
   Exit Sub

Drive4_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Drive4_Change of Form mapdiskDTMfm"
End Sub

Private Sub Drive5_Change(Index As Integer)
   If Dir(Drive5(1).Drive & "\usa", vbDirectory) <> sEmpty Then
      txtUSA.Text = "USA"
   Else
      txtUSA.Text = sEmpty
      End If
End Sub

Private Sub DriveALOS_Change()
   If Dir(DriveALOS.Drive & "\bil", vbDirectory) <> sEmpty Then
      txtBIL.Text = "BIL"
   Else
      txtBIL.Text = sEmpty
      End If
End Sub

Private Sub form_load()
   On Error GoTo errhand
   
   myfile = Dir(drivjk_c$ + "mapcdinfo.sav")
   If myfile = sEmpty Then
      israeldtmcd = True
      worlddtmcd = True
      srtmdtmcd = True
      'Turbo2cdDir$, USADir$, GEOTOPO30Dir$, D3ASDir$
      If drivdtm$ <> sEmpty Then
         Drive1(1).Drive = Mid$(drivdtm$, 1, 1)
         israeldtm = Mid$(drivdtm$, 1, 1)
      Else
         Drive1(1).Drive = MainDir$
         israeldtm = MainDir$
         End If
      If GEOTOPO30Dir$ <> sEmpty Then
         worldtm = Mid$(GEOTOPO30Dir$, 1, 1)
         Drive2(1).Drive = Mid$(GEOTOPO30Dir$, 1, 1)
      Else
         worlddtm = MainDir$
         Drive2(1).Drive = MainDir$
         End If
      If USADir$ <> sEmpty Then
         srtmdtm = Mid$(USADir$, 1, 1)
         Drive5(1).Drive = Mid$(USADir$, 1, 1)
      Else
         srtmdtm = "e"
         Drive2(1).Drive = MainDir$
         End If
      If D3ASDir$ <> sEmpty Then
         Drive5(0).Drive = Mid$(D3ASDir$, 1, 1)
         d3asdtm = Mid$(D3ASDir$, 1, 1)
      Else
         Drive5(0).Drive = MainDir$
         End If
      If BILDir$ <> sEmpty Then
         alosdtm = Mid$(BILDir$, 1, 1)
         DriveALOS.Drive = Mid$(BILDir$, 1, 1)
      Else
         alosdtm = "c"
         DriveALOS.Drive = MainDir$
         End If
         
      ramdrive = MainDir$
      
      If WinVer = 5 Then 'Or WinVer = 261 Then
         'Windows 2000 or XP
         ramdrivef = MainDir$
         Drive3(1).Drive = MainDir$
         Drive4(1).Drive = MainDir$
      ElseIf WinVer > 5 Then
         ramdrivef = MainDir$
         Drive3(1).Drive = MainDir$
         Drive4(1).Drive = MainDir$
      Else
         Drive3(1).Drive = "g"
         Drive4(1).Drive = "e"
         End If
      ChDir Drive4(1).Drive + "\"
      For i% = 0 To Dir1(1).ListCount - 1
         If Dir1(1).List(i%) = "terraviewer" Then Exit For
      Next i%
      terradir$ = "e:\terraviewer"
      Text1(1).Text = 1#: Text2(1).Text = 1#
      Check1(1).value = vbUnchecked
   Else
      mapinfonum% = FreeFile
      Open drivjk_c$ + "mapcdinfo.sav" For Input As #mapinfonum%
      Input #mapinfonum%, israeldtm, israeldtmcdnum
      Drive1(1).Drive = israeldtm
      If israeldtmcdnum = 0 Then
         israeldtmcd = False
         Option2(1).value = True
         optJK.Enabled = True
         'Option2_Click
      Else
         israeldtmcd = True
         Option1(1).value = True
         'Option1_Click
         End If
      Input #mapinfonum%, worlddtm, worlddtmcdnum
      Drive2(1).Drive = worlddtm
      Drive5(0).Drive = worlddtm
      If worlddtmcdnum = 0 Then
         worlddtmcd = False
         Option4(1).value = True
         Option5(0).value = True
         'Option4_Click
      Else
         worlddtmcd = True
         Option3(1).value = True
         Option6(0).value = True
         'Option3_Click
         End If
      Input #mapinfonum%, ramdrive
      If WinVer = 5 Or WinVer = 261 Then
        'Windows 2000 identified, use hard drive i
        'since Ramdrive.sys is not supported in Windows 2000.
        ramdrive = MainDir$ '"e"
      ElseIf WinVer > 5 Then
        'windows vista,7,8....
        ramdrive = MainDir$ '"c"
        End If
      Drive3(1).Drive = ramdrive
      Input #mapinfonum%, terradir$
      Drive4(1).Drive = Mid$(terradir$, 1, 3)
      ChDir Mid$(terradir$, 1, 1)
      ChDrive Drive4(1).Drive
      Dir1(1).Path = Drive4(1).Drive
      For i% = 0 To Dir1(1).ListCount - 1
         If LCase$(Dir1(1).List(i%)) = LCase$(terradir$) Then
         Dir1(1).ListIndex = i% + 1
         Exit For
         End If
      Next i%
      Input #mapinfonum%, adx1f, bdy1f
      If adx1 = 0 Or bdy1 = 0 Then
         Text1(1).Text = adx1f: Text2(1).Text = bdy1f
         adx1 = adx1f: bdy1 = bdy1f
      Else
         Text1(1).Text = adx1: Text2(1).Text = bdy1
         End If
      Input #mapinfonum%, RdHalYes
      If RdHalYes Then
         Check1(1).value = vbChecked
      Else
         Check1(1).value = vbUnchecked
         End If
      Input #mapinfonum%, IsraelDTMsource%
      If IsraelDTMsource% = 1 Then 'SRTM is source of Israel elevations
         optSRTM.value = True
      ElseIf IsraelDTMsource% = 2 Then 'SRTM3 - MERIT DEM is source of elevations in Israel
         optSRTM3.value = True
      ElseIf IsraelDTMsource% = 3 Then 'ALOS DEM is source of elevations in Israel
         optALOS.value = True
         End If
      Close #mapinfonum%
      End If
      
      If Dir(israeldtm + ":\dtm\N31E035.hgt") <> sEmpty Then
         Close 'close any open files
         CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
         optJK.Enabled = True
         End If
      
      'now open SRTM info file
      If Dir(drivjk_c$ & "mapSRTMinfo.sav") <> sEmpty Then
         mapinfonum% = FreeFile
         Open drivjk_c$ & "mapSRTMinfo.sav" For Input As #mapinfonum%
         Input #mapinfonum%, srtmdtm, srtmdtmcdnum
         Drive5(1).Drive = srtmdtm
         If srtmdtmcdnum = 0 Then
           srtmdtmcd = False
           Option5(1).value = True
         Else
           srtmdtmcd = True
           Option6(1).value = True
           End If
         Input #mapinfonum%, USADirtmp$
         If USADirtmp$ <> sEmpty Then
            USADir$ = srtmdtm & ":\" & USADirtmp$ & "\"
            txtUSA.Text = Mid$(USADir$, 4, Len(USADir$) - 4)
            optSRTM.Enabled = True
         Else
            If Dir(USADir$, vbDirectory) <> sEmpty Then
               txtUSA = Mid$(USADir$, 4, Len(USADir$) - 4)
               End If
            End If
         If Not EOF(mapinfonum%) Then
            Input #mapinfonum%, d3asdtm, D3ASDirtmp$
            If D3ASDirtmp$ <> sEmpty Then
               D3ASDir$ = d3asdtm & ":\" & D3ASDirtmp$ & "\"
               txt3AS.Text = D3ASDirtmp$
               optSRTM3.Enabled = True
               End If
            End If
         If Not EOF(mapinfonum%) Then
            Input #mapinfonum%, alosdtm, BILDirtmp$
            If BILDirtmp$ <> sEmpty Then
               BILDir$ = alosdtm & ":\" & BILDirtmp$ & "\"
               txtBIL.Text = BILDirtmp$
               optALOS.Enabled = True
               End If
            End If
         
         Close #mapinfonum%
      Else
         Drive5(1).Drive = "d" 'default SRTM data drive"
         srtmdtm = "d"
         srtmdtmcd = False 'disk drive is default
         Option5(1).value = True
         End If
         
      If IsraelDTMsource% = 1 Then
         optSRTM.value = True
      ElseIf IsraelDTMsource% = 2 Then
         optSRTM3.value = True
      ElseIf IsraelDTMsource% = 3 Then
         optALOS.value = True
         End If
         
      Exit Sub
      
errhand:
   Resume Next
End Sub

Private Sub optALOS_Click()
   'use ALOS DTM for Eretz Israel
'   If Dir(israeldtm + ":\dtm\N31E035.hgt") <> sEmpty Then
'      Close 'close any open files
'      CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
      IsraelDTMsource% = 3
'   Else
'      MsgBox "Israel SRTM DTM not found is expected location", vbExclamation + vbOKOnly, "Maps&More"
'      End If
End Sub

Private Sub Option1_Click(Index As Integer)
  israeldtmcd = True
End Sub

Private Sub Option2_Click(Index As Integer)
  israeldtmcd = False
End Sub

Private Sub Option3_Click(Index As Integer)
  worlddtmcd = True
End Sub

Private Sub Option4_Click(Index As Integer)
  worlddtmcd = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
End Sub

Private Sub Option5_Click(Index As Integer)
   srtmdtmcd = False
End Sub

Private Sub Option6_Click(Index As Integer)
   srtmdtmcd = True
End Sub

Private Sub optJK_Click()
   'use JKH's 25-m DTM
   If Dir(israeldtm + ":\dtm\dtm-map.loc") <> sEmpty Then
      Close 'close any open files
      CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
      IsraelDTMsource% = 0
      optJK.Enabled = True
   Else
      MsgBox "Israel 25-m DTM not found is expected location", vbExclamation + vbOKOnly, "Maps&More"
      End If
End Sub

Private Sub optSRTM_Click()
   'use SRTM1 DTM for Eretz Israel
'   If Dir(israeldtm + ":\dtm\N31E035.hgt") <> sEmpty Then
'      Close 'close any open files
'      CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
      IsraelDTMsource% = 1
'   Else
'      MsgBox "Israel SRTM DTM not found is expected location", vbExclamation + vbOKOnly, "Maps&More"
'      End If
End Sub

Private Sub optSRTM3_Click()
   'use SRTM3 DTM for Eretz Israel
'   If Dir(israeldtm + ":\dtm\N31E035.hgt") <> sEmpty Then
'      Close 'close any open files
'      CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
      IsraelDTMsource% = 2
'   Else
'      MsgBox "Israel SRTM DTM not found is expected location", vbExclamation + vbOKOnly, "Maps&More"
'      End If
End Sub
