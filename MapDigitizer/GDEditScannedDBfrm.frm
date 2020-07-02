VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GDEditScannedDBfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit scanned database"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   Icon            =   "GDEditScannedDBfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10635
   Begin VB.CheckBox chkSaveVerify 
      Caption         =   "Confirm changes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   108
      ToolTipText     =   "Check to ask for confirmation to commit changes"
      Top             =   5880
      Value           =   1  'Checked
      Width           =   1360
   End
   Begin MSComctlLib.StatusBar stbrEdSDB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   104
      Top             =   6225
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdTifView 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10050
      Picture         =   "GDEditScannedDBfrm.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   96
      ToolTipText     =   "View the scanned tif file"
      Top             =   5500
      Width           =   315
   End
   Begin VB.Frame frmModify 
      Caption         =   "Last modified"
      Height          =   675
      Left            =   8880
      TabIndex        =   91
      Top             =   4740
      Width           =   1695
      Begin VB.Label lblModified 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   92
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame frmWizard 
      Caption         =   "Suggested ITM"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   8880
      TabIndex        =   86
      Top             =   2640
      Width           =   1695
      Begin VB.ListBox lstWizard 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1560
         Left            =   120
         TabIndex        =   87
         ToolTipText     =   "Click to accept"
         Top             =   300
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame frmStatus 
      Caption         =   "Missing Entries"
      Height          =   2595
      Left            =   8880
      TabIndex        =   75
      Top             =   0
      Width           =   1695
      Begin VB.ListBox lstMissing 
         ForeColor       =   &H000000C0&
         Height          =   2205
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdPreview 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9725
      Picture         =   "GDEditScannedDBfrm.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Preview"
      Top             =   5500
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   9400
      Picture         =   "GDEditScannedDBfrm.frx":0986
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Undo"
      Top             =   5500
      Width           =   315
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   9055
      Picture         =   "GDEditScannedDBfrm.frx":0EB8
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Save changes"
      Top             =   5500
      Width           =   315
   End
   Begin VB.Frame frmEdit 
      Height          =   6255
      Left            =   60
      TabIndex        =   60
      Top             =   -60
      Width           =   8715
      Begin VB.CheckBox chkRefreshTif 
         Caption         =   "Refresh doc. image on stepping"
         Enabled         =   0   'False
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
         Left            =   6080
         TabIndex        =   110
         ToolTipText     =   "Check to automatically refresh the document's image file image on stepping"
         Top             =   1960
         Width           =   2295
      End
      Begin VB.Frame frmSource 
         Caption         =   "Sample Source"
         Height          =   850
         Left            =   120
         TabIndex        =   88
         Top             =   1210
         Width           =   2535
         Begin VB.CommandButton cmdUndoSources 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2100
            Picture         =   "GDEditScannedDBfrm.frx":13EA
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Undo"
            Top             =   160
            Width           =   315
         End
         Begin VB.OptionButton optWells 
            Caption         =   "&Well"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   540
            Width           =   675
         End
         Begin VB.Frame frmCore 
            Height          =   375
            Left            =   840
            TabIndex        =   89
            Top             =   420
            Width           =   1635
            Begin VB.OptionButton optCuttings 
               Caption         =   "Cu&ttings"
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
               Left            =   780
               TabIndex        =   50
               Top             =   130
               Width           =   795
            End
            Begin VB.OptionButton optCore 
               Caption         =   "&Core"
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
               Left            =   60
               TabIndex        =   49
               Top             =   130
               Width           =   615
            End
         End
         Begin VB.OptionButton optOutcroppings 
            Caption         =   "&Outcropping"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   220
            Width           =   2055
         End
      End
      Begin VB.OptionButton optStepOKeyNo 
         Caption         =   "Step in O_KEY"
         Height          =   195
         Left            =   6900
         TabIndex        =   5
         Top             =   1500
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optStepSerialNo 
         Caption         =   "Step in Doc. No."
         Height          =   195
         Left            =   6900
         TabIndex        =   4
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Frame frmCoordinates 
         Caption         =   "Coordinates and Place Names"
         Height          =   1455
         Left            =   120
         TabIndex        =   77
         Top             =   2100
         Width           =   8475
         Begin VB.CommandButton cmdFathoms 
            Height          =   285
            Left            =   2130
            TabIndex        =   115
            ToolTipText     =   "Convert depth in fathoms to meters"
            Top             =   900
            Width           =   100
         End
         Begin VB.CommandButton cmdConvert 
            Height          =   265
            Left            =   3060
            TabIndex        =   114
            ToolTipText     =   "Convert depth in feet to depth in meters"
            Top             =   920
            Width           =   115
         End
         Begin VB.CommandButton cmdCopyCoord 
            Height          =   315
            Left            =   1240
            Picture         =   "GDEditScannedDBfrm.frx":191C
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Copy coordinates from maps"
            Top             =   1020
            Width           =   315
         End
         Begin VB.OptionButton optWellscat 
            Caption         =   "WellsCat Table"
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
            Left            =   7080
            TabIndex        =   94
            Top             =   600
            Width           =   1275
         End
         Begin VB.CommandButton cmdPasteCoord 
            Height          =   315
            Left            =   4380
            Picture         =   "GDEditScannedDBfrm.frx":1E4E
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Replace with catalogue name"
            Top             =   1020
            Width           =   315
         End
         Begin VB.TextBox txtDepth 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2240
            TabIndex        =   12
            Text            =   "txtDepth"
            Top             =   900
            Width           =   855
         End
         Begin VB.OptionButton optLoadPlacesCat 
            Caption         =   "Places/Wells Cat."
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
            Left            =   5700
            TabIndex        =   59
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtGLCat 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   7695
            TabIndex        =   20
            Text            =   "txtGLCat"
            ToolTipText     =   "Gournd Level (meters)"
            Top             =   1100
            Width           =   675
         End
         Begin VB.TextBox txtITMx 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   540
            TabIndex        =   6
            Text            =   "txtITMx"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtITMy 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   540
            TabIndex        =   7
            Text            =   "txtITMy"
            Top             =   660
            Width           =   1455
         End
         Begin VB.CommandButton cmdExchange 
            Height          =   315
            Left            =   930
            Picture         =   "GDEditScannedDBfrm.frx":2380
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exchange x and y"
            Top             =   1020
            Width           =   315
         End
         Begin VB.ComboBox cmbdbNames 
            Height          =   315
            Left            =   5700
            Sorted          =   -1  'True
            TabIndex        =   14
            Text            =   "cmbdbNames"
            Top             =   180
            Width           =   2655
         End
         Begin VB.ComboBox cmbPlaceNames 
            Height          =   315
            Left            =   5700
            Sorted          =   -1  'True
            TabIndex        =   17
            Text            =   "cmbPlaceNames"
            Top             =   780
            Width           =   2675
         End
         Begin VB.CommandButton cmdUndoCoordinates 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4380
            Picture         =   "GDEditScannedDBfrm.frx":259E
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Undo"
            Top             =   595
            Width           =   315
         End
         Begin VB.TextBox txtGL 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3180
            TabIndex        =   13
            Text            =   "txtGL"
            Top             =   900
            Width           =   915
         End
         Begin VB.CommandButton cmdMap 
            Height          =   315
            Left            =   1560
            Picture         =   "GDEditScannedDBfrm.frx":26A0
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Locate on map"
            Top             =   1020
            Width           =   315
         End
         Begin VB.CommandButton cmdWizard 
            Height          =   315
            Left            =   600
            Picture         =   "GDEditScannedDBfrm.frx":27A2
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Coordinate wizard"
            Top             =   1020
            Width           =   315
         End
         Begin VB.TextBox txtNames 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2200
            TabIndex        =   11
            Text            =   "txtNames"
            Top             =   360
            Width           =   1915
         End
         Begin VB.CommandButton cmdEditCoordinates 
            Height          =   315
            Left            =   4380
            Picture         =   "GDEditScannedDBfrm.frx":292C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Replace with database name"
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox txtITMxCat 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5700
            TabIndex        =   18
            Text            =   "txtITMxCat"
            ToolTipText     =   "ITMx"
            Top             =   1100
            Width           =   1035
         End
         Begin VB.TextBox txtITMyCat 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6720
            TabIndex        =   19
            Text            =   "txtITMyCat"
            ToolTipText     =   "ITMy"
            Top             =   1100
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "ITMx"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   420
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "ITMy"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   720
            Width           =   555
         End
         Begin VB.Label lbldbNames 
            Alignment       =   2  'Center
            Caption         =   "Names in data base"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4980
            TabIndex        =   83
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblPlaceNames 
            Alignment       =   2  'Center
            Caption         =   "Place  && Well Catqlogues"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   82
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblGL 
            Alignment       =   2  'Center
            Caption         =   " Depth (m)     Ground Level (m)"
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
            Left            =   2160
            TabIndex        =   81
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "below the surface   w.r.t. sea level"
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
            Left            =   2100
            TabIndex        =   80
            Top             =   1200
            Width           =   2115
         End
         Begin VB.Label lblNames 
            Caption         =   "Currently recorded Name"
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
            Left            =   2400
            TabIndex        =   79
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label lblCoordCat 
            Alignment       =   2  'Center
            Caption         =   "ITMx, ITMy, GL"
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
            Left            =   4800
            TabIndex        =   78
            Top             =   1080
            Width           =   915
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   1680
         TabIndex        =   70
         Top             =   5530
         Width           =   6915
         Begin VB.PictureBox picProgBar 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   245
            Left            =   180
            ScaleHeight     =   210
            ScaleWidth      =   6525
            TabIndex        =   71
            Top             =   235
            Visible         =   0   'False
            Width           =   6555
         End
      End
      Begin VB.Frame frmFormations 
         Caption         =   "Geologic Formations"
         Height          =   795
         Left            =   1680
         TabIndex        =   69
         Top             =   4740
         Width           =   6915
         Begin VB.CommandButton cmdUndoFormation 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3840
            Picture         =   "GDEditScannedDBfrm.frx":2E5E
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Undo"
            Top             =   300
            Width           =   315
         End
         Begin VB.CommandButton cmdEditFormation 
            Height          =   315
            Left            =   3540
            Picture         =   "GDEditScannedDBfrm.frx":3390
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Replace"
            Top             =   300
            Width           =   315
         End
         Begin VB.TextBox txtFormation 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   780
            TabIndex        =   31
            Text            =   "txtFormation"
            Top             =   320
            Width           =   2655
         End
         Begin VB.ComboBox cmbFormation 
            Height          =   315
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   34
            Text            =   "cmbFormation"
            Top             =   300
            Width           =   2655
         End
         Begin VB.Label lblFormation 
            Alignment       =   2  'Center
            Caption         =   "Recorded Formation"
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
            TabIndex        =   74
            Top             =   300
            Width           =   675
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblFormations 
            Alignment       =   2  'Center
            Caption         =   "Catalogue of Formations"
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
            Left            =   4740
            TabIndex        =   73
            Top             =   120
            Width           =   1515
         End
      End
      Begin VB.Frame frmAges 
         Caption         =   "Geologic Ages"
         Height          =   1195
         Left            =   1680
         TabIndex        =   66
         Top             =   3540
         Width           =   6915
         Begin VB.CommandButton cmdPasteLstAges 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5950
            Picture         =   "GDEditScannedDBfrm.frx":38C2
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Copy  last recorded ages"
            Top             =   750
            Width           =   315
         End
         Begin VB.CommandButton cmdUndoCopyEtoL 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3550
            Picture         =   "GDEditScannedDBfrm.frx":3DF4
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Undo Copy"
            Top             =   680
            Width           =   315
         End
         Begin VB.CommandButton cmdCopyEarlytoLate 
            Height          =   315
            Left            =   3550
            Picture         =   "GDEditScannedDBfrm.frx":4326
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Copy early age to later age"
            Top             =   360
            Width           =   315
         End
         Begin VB.CommandButton cmdUndoAge 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6400
            Picture         =   "GDEditScannedDBfrm.frx":4858
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Undo"
            Top             =   750
            Width           =   315
         End
         Begin VB.CommandButton cmdEditAge 
            Height          =   315
            Left            =   5650
            Picture         =   "GDEditScannedDBfrm.frx":4D8A
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Replace"
            Top             =   750
            Width           =   315
         End
         Begin VB.TextBox txtLaterAge 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Text            =   "txtLaterAge"
            Top             =   720
            Width           =   1635
         End
         Begin VB.TextBox txtPreL 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1140
            TabIndex        =   23
            Text            =   "txtPreL"
            Top             =   720
            Width           =   675
         End
         Begin VB.TextBox txtEarlyAge 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1800
            TabIndex        =   22
            Text            =   "txtEarlyAge"
            Top             =   300
            Width           =   1635
         End
         Begin VB.TextBox txtPreE 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1140
            TabIndex        =   21
            Text            =   "txtPreE"
            Top             =   300
            Width           =   675
         End
         Begin VB.OptionButton optLate 
            Caption         =   "Later Age"
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
            Left            =   4500
            TabIndex        =   28
            Top             =   900
            Width           =   1035
         End
         Begin VB.OptionButton optEarly 
            Caption         =   "Earlier Age"
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
            Left            =   4500
            TabIndex        =   27
            Top             =   720
            Width           =   1035
         End
         Begin VB.ComboBox cmbAge 
            Height          =   315
            Left            =   5040
            Sorted          =   -1  'True
            TabIndex        =   26
            Text            =   "cmbAge"
            Top             =   360
            Width           =   1815
         End
         Begin VB.ComboBox cmbPreAge 
            Height          =   315
            Left            =   4200
            TabIndex        =   25
            Text            =   "cmbPreAge1"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblAges 
            Alignment       =   2  'Center
            Caption         =   "Catalogue of Ages"
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
            Left            =   4740
            TabIndex        =   72
            Top             =   180
            Width           =   1515
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Recorded Later Date"
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
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEarlier 
            Alignment       =   2  'Center
            Caption         =   "Recorded Early Date"
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
            Left            =   120
            TabIndex        =   67
            Top             =   315
            Width           =   1035
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frmFossils 
         Caption         =   "Fossil Categories"
         Height          =   2605
         Left            =   120
         TabIndex        =   65
         Top             =   3540
         Width           =   1515
         Begin VB.CommandButton cmdUndoFos 
            Enabled         =   0   'False
            Height          =   315
            Left            =   600
            Picture         =   "GDEditScannedDBfrm.frx":52BC
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Undo"
            Top             =   2220
            Width           =   315
         End
         Begin VB.OptionButton optShekef 
            Caption         =   "&Shekef"
            Height          =   195
            Left            =   80
            TabIndex        =   42
            Top             =   1920
            Width           =   975
         End
         Begin VB.OptionButton optPaly 
            Caption         =   "&Palynology"
            Height          =   195
            Left            =   80
            TabIndex        =   41
            Top             =   1680
            Width           =   1095
         End
         Begin VB.OptionButton optOstra 
            Caption         =   "&Ostracoda"
            Height          =   195
            Left            =   80
            TabIndex        =   40
            Top             =   1440
            Width           =   1155
         End
         Begin VB.OptionButton optNano 
            Caption         =   "&Nannoplankton"
            Height          =   195
            Left            =   80
            TabIndex        =   39
            Top             =   1200
            Width           =   1395
         End
         Begin VB.OptionButton optMega 
            Caption         =   "&Megafauna"
            Height          =   195
            Left            =   80
            TabIndex        =   38
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optForam 
            Caption         =   "&Foraminifera"
            Height          =   195
            Left            =   80
            TabIndex        =   37
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optDiatom 
            Caption         =   "&Diatom"
            Height          =   195
            Left            =   80
            TabIndex        =   36
            Top             =   480
            Width           =   915
         End
         Begin VB.OptionButton optCono 
            Caption         =   "&Conodonta"
            Height          =   195
            Left            =   80
            TabIndex        =   35
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbSerialNo 
         Height          =   315
         Left            =   4500
         TabIndex        =   1
         Text            =   "cmbSerialNo"
         Top             =   1560
         Width           =   1395
      End
      Begin VB.ComboBox cmbOKEYNo 
         Height          =   315
         Left            =   3060
         TabIndex        =   0
         Text            =   "cmbOKEYNo"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         TabIndex        =   3
         ToolTipText     =   "Next record"
         Top             =   1560
         Width           =   315
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6060
         TabIndex        =   2
         ToolTipText     =   "Previous record"
         Top             =   1560
         Width           =   315
      End
      Begin VB.Frame frmDataSource 
         Caption         =   "Data Source"
         Height          =   1095
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   8475
         Begin VB.Frame frmSearchType 
            Height          =   375
            Left            =   5100
            TabIndex        =   111
            Top             =   240
            Width           =   1215
            Begin VB.OptionButton optOKeys 
               Caption         =   "OKy"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   220
               Left            =   680
               TabIndex        =   113
               ToolTipText     =   "Search over range of O_Keys"
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton optDigits 
               Caption         =   "dig."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   220
               Left            =   80
               TabIndex        =   112
               ToolTipText     =   "Search over range of coordinate digits"
               Top             =   120
               Value           =   -1  'True
               Width           =   480
            End
         End
         Begin VB.CheckBox chkHeights 
            Caption         =   "Include only depths<>0"
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
            Left            =   4600
            TabIndex        =   106
            ToolTipText     =   "Check to search only for records with non-zero recorded  depths"
            Top             =   755
            Width           =   1770
         End
         Begin VB.Frame frmQuick 
            Caption         =   "Quick"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   940
            Left            =   7320
            TabIndex        =   102
            Top             =   100
            Width           =   1020
            Begin VB.CommandButton cmdQuickOptions 
               BackColor       =   &H8000000B&
               Caption         =   "&Options"
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
               Style           =   1  'Graphical
               TabIndex        =   116
               ToolTipText     =   "Pick which fields to paste"
               Top             =   170
               Width           =   735
            End
            Begin VB.CommandButton cmdSaveTemplate 
               Caption         =   "&Save"
               Height          =   235
               Left            =   120
               TabIndex        =   109
               ToolTipText     =   "Save current template"
               Top             =   380
               Width           =   735
            End
            Begin VB.CommandButton cmdQuick 
               Caption         =   "&Quick"
               Height          =   250
               Left            =   120
               TabIndex        =   103
               ToolTipText     =   "&Apply last recorded template"
               Top             =   620
               Width           =   735
            End
         End
         Begin VB.Frame frmCombineDB 
            Caption         =   "Combine DBs"
            Height          =   735
            Left            =   120
            TabIndex        =   99
            ToolTipText     =   "Start automatic combination process "
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
            Begin VB.CommandButton cmdStartCombineDB 
               Caption         =   "&Start "
               Enabled         =   0   'False
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
               Left            =   120
               TabIndex        =   101
               ToolTipText     =   "Click here to begin automatic combination of the two pal_old databases"
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton cmdEnableCombineDB 
               Caption         =   "&Enable"
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
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   100
               ToolTipText     =   "Enable combining the two versions of Pal_old.mdb"
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.OptionButton optReset 
            Caption         =   "&Reset"
            Height          =   195
            Left            =   6480
            TabIndex        =   95
            Top             =   840
            Width           =   735
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Edit records from &Search Report"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1800
            TabIndex        =   58
            Top             =   840
            Width           =   2595
         End
         Begin VB.OptionButton optAllSearch 
            Caption         =   "Edit records from &All sources"
            Height          =   195
            Left            =   1800
            TabIndex        =   57
            Top             =   600
            Width           =   2355
         End
         Begin VB.OptionButton optOutcroppingSearch 
            Caption         =   "Edit records only from &Outcroppings"
            Height          =   195
            Left            =   1800
            TabIndex        =   56
            Top             =   360
            Width           =   2835
         End
         Begin VB.OptionButton optWellSearch 
            Caption         =   "Edit records only from &Wells"
            Height          =   195
            Left            =   1800
            TabIndex        =   55
            Top             =   140
            Width           =   2295
         End
         Begin MSComCtl2.UpDown udMaxDigits 
            Height          =   285
            Left            =   6900
            TabIndex        =   54
            Top             =   480
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   7
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtMaxDigits"
            BuddyDispid     =   196718
            OrigLeft        =   6540
            OrigTop         =   480
            OrigRight       =   6780
            OrigBottom      =   735
            Max             =   7
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMaxDigits 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   53
            Text            =   "7"
            Top             =   480
            Width           =   435
         End
         Begin MSComCtl2.UpDown udMinDigits 
            Height          =   285
            Left            =   6900
            TabIndex        =   52
            Top             =   180
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   4
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtMinDigits"
            BuddyDispid     =   196719
            OrigLeft        =   5220
            OrigTop         =   300
            OrigRight       =   5460
            OrigBottom      =   555
            Max             =   7
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMinDigits 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   51
            Text            =   "4"
            Top             =   180
            Width           =   435
         End
         Begin VB.Line Line5 
            X1              =   4800
            X2              =   5160
            Y1              =   480
            Y2              =   780
         End
         Begin VB.Line Line4 
            X1              =   4800
            X2              =   5160
            Y1              =   480
            Y2              =   180
         End
         Begin VB.Line Line3 
            X1              =   4800
            X2              =   4800
            Y1              =   180
            Y2              =   740
         End
         Begin VB.Line Line2 
            X1              =   4560
            X2              =   4810
            Y1              =   740
            Y2              =   740
         End
         Begin VB.Line Line1 
            X1              =   4560
            X2              =   4800
            Y1              =   180
            Y2              =   180
         End
         Begin VB.Label lblMax 
            Caption         =   "Max No. of Digits"
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
            Left            =   5180
            TabIndex        =   64
            Top             =   620
            Width           =   1215
         End
         Begin VB.Label lblMin 
            Caption         =   "Min No of Digits"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   5200
            TabIndex        =   63
            Top             =   110
            Width           =   1155
         End
      End
      Begin VB.Label Label1 
         Caption         =   "           O_KEY            Document Serial #           "
         Height          =   195
         Left            =   2940
         TabIndex        =   61
         Top             =   1320
         Width           =   3015
      End
   End
End
Attribute VB_Name = "GDEditScannedDBfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fillEarlyAge As Boolean
Dim fillLateAge As Boolean
'declare variables used for undoing editing of entries
Dim txtITMx0$
Dim txtITMy0$
Dim txtPreE0$
Dim txtEarlyAge0$
Dim txtPreL0$
Dim txtLaterAge0$
Dim txtFormation0$
Dim txtNames0$
Dim txtGL0$
Dim txtDepth0$
Dim LoadPlaces%
Dim PasteLstAges As Boolean
Dim txtPreELast$
Dim txtPreEAgeLast$
Dim txtPreLLast$
Dim txtpreLAgeLast$




Private Sub cmbOKEYNo_Click()
   'check input
   If val(Trim$(cmbOKEYNo.Text)) <= 0 Then
      MsgBox "The O_KEY number must be greater than zero!", vbExclamation + vbOKOnly, "MapDigitizer/Edit DB Error"
      Exit Sub
      End If
   'load up Edit Form with information for this O_KEY (id field
   'of old scanned database)
   If Not LoadingEditForm Then
      'check for changes before moving to different record
      cmdSave_Click
      'now go to chosen record
      OKeyClick = True
      ONameClick = False
      minDigits = 0 'allow for any input
      maxDigits = 7
      cc = cmbOKEYNo.ListIndex
      LoadEditForm
      End If
   
   'display the record number clicked
   GDEditScannedDBfrm.stbrEdSDB.Panels(2).Text = sEmpty
   GDEditScannedDBfrm.stbrEdSDB.Panels(2).Text = "OKEY Record #: " & Trim$(str$(cmbOKEYNo.ListIndex + 1))
   
End Sub

Private Sub cmbOKEYNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then 'space inputed--> works like carraige return
     'load up Edit Form with information for this O_KEY (id field
     'of old scanned database)
     modeEdit% = 0
     minDigits = 0
     maxDigits = 7
     txtTmp$ = cmbOKEYNo.Text
     cmbSerialNo.Clear
     cmbOKEYNo.Clear
     cmbOKEYNo.Text = txtTmp$
     cmbOKEYNo_Click
   Else
      DoEvents
      End If
End Sub

Private Sub cmbPlaceNames_Click()
   If cmbPlaceNames.ListCount > 1 Then
      'determine the corresponding ITMx,ITMy,GL
      Call LoadCatCoord(GDEditScannedDBfrm, LoadPlaces%)
      End If
End Sub

Private Sub cmbSerialNo_Click()
   'load up Edit Form with information for this O_NAME (Document
   'name field of old scanned database)
   If Not LoadingEditForm Then
      'check for changes before moving to different record
      cmdSave_Click
      'now go to chosen record
      ONameClick = True
      OKeyClick = False
      minDigits = 0 'allow for any input
      maxDigits = 7
      LoadEditForm
      End If
      
   'display the record number clicked
   GDEditScannedDBfrm.stbrEdSDB.Panels(2).Text = sEmpty
   GDEditScannedDBfrm.stbrEdSDB.Panels(2).Text = "DOC Record #: " & Trim$(str$(cmbSerialNo.ListIndex + 1))
      
End Sub

Private Sub cmbSerialNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 32 Then 'space entered-->works like carriage return
    'load up Edit Form with information for this O_NAME (Document
    'name field of old scanned database)
    modeEdit% = 0
    minDigits = 0
    maxDigits = 7
    txtTmp$ = cmbSerialNo.Text
    cmbSerialNo.Clear
    cmbOKEYNo.Clear
    cmbSerialNo.Text = txtTmp$
    cmbSerialNo_Click
  Else
    DoEvents
    End If
End Sub

Private Sub cmdCancel_Click()

  'restore forms to original values
  With GDEditScannedDBfrm
     .txtITMx = txtITMx00$
     .txtITMy = txtITMy00$
     .txtPreE = txtPreE00$
     .txtEarlyAge = txtEarlyAge00$
     .txtPreL = txtPreL00$
     .txtLaterAge = txtLaterAge00$
     .txtFormation = txtFormation00$
     .txtNames = txtNames00$
     .txtGL = txtGL00$
     .txtDepth = txtDepth00$
     
     Select Case foscat00%
       Case 0 'blank
         .optCono.value = False
         .optDiatom.value = False
         .optForam.value = False
         .optMega.value = False
         .optNano.value = False
         .optOstra.value = False
         .optPaly.value = False
         .optShekef.value = False
       Case 1 'foram
         .optForam.value = True
       Case 2 'foram (shekef)
         .optShekef.value = True
       Case 3 'ostra
         .optOstra.value = True
       Case 4 'paly
         .optPaly.value = True
       Case 5 'mega
         .optMega.value = True
       Case 6 'nanno
         .optNano.value = True
       Case 7 'diatom
         .optDiatom.value = True
       Case 8 'cono
         .optCono.value = True
     End Select
     
   .optOutcroppings.value = False
   .optWells.value = False
   .optCore.value = False
   .optCuttings.value = False
   Select Case oldSource00%
      Case 0
         oldSource% = 0
      Case 1
         .optWells.value = True
      Case 2
         .optWells.value = True
         .optCore.value = True
      Case 3
         .optWells.value = True
         .optCuttings.value = True
      Case 4
         .optOutcroppings.value = True
   End Select
   .cmdUndoSources.Enabled = False
   .cmdUndoAge.Enabled = False
   .cmdUndoCoordinates.Enabled = False
   .cmdUndoFos.Enabled = False
   .cmdUndoFormation.Enabled = False
   .cmdCancel.Enabled = False
   
   .cmdSave.Enabled = False 'nothing more to save
     
  End With

End Sub

Private Sub cmdCoord_Click()
   Select Case cmdCoord.Caption
      Case "<-- o_key -->"
        'search over coordinates
        'prompt to change back to digits
        cmdCoord.Caption = "<-- digit -->"
        cmdCoord.ToolTipText = "Press to switch to search over range of digits"
        lblMin.Caption = "Min O_KEY"
        lblMax.Caption = "Max O_KEY"
        txtMinDigits.Left = txtMinDigits.Left - 300
        txtMinDigits.Width = txtMinDigits.Width + 300
        udMinDigits.Max = 60000
        txtMaxDigits.Left = txtMaxDigits.Left - 300
        txtMaxDigits.Width = txtMaxDigits.Width + 300
        udMaxDigits.Max = 60000
      Case "<-- digit -->"
        'search over number of digits in coordinates
        'prompt to change back to coordinates
        cmdCoord.Caption = "<-- o_key -->"
        cmdCoord.ToolTipText = "Press to switch to search over range of O_KEYs"
        lblMin.Caption = "Min No of Digits"
        lblMax.Caption = "Max No. of Digits"
        txtMinDigits.Left = txtMinDigits.Left + 300
        txtMinDigits.Width = txtMinDigits.Width - 300
        udMinDigits.Max = 7
        txtMaxDigits.Left = txtMaxDigits.Left + 300
        txtMaxDigits.Width = txtMaxDigits.Width - 300
        udMaxDigits.Max = 7
   End Select
End Sub

Private Sub cmdConvert_Click()
 'convert depth in feet to depth in meters
  
  'first record current parameter values for undo
  txtNames0$ = txtNames
  txtITMx0$ = txtITMx
  txtITMy0$ = txtITMy
  txtGL0$ = txtGL
  txtDepth0$ = txtDepth
  cmdUndoCoordinates.Enabled = False

  'now convert depth in feet to depth in meters
  txtDepth = CInt(val(txtDepth) * 3.048) * 0.1
  
End Sub

Private Sub cmdCopyCoord_Click()
   
   'copy coordinates from the maps
   
   'first store the original values
   txtITMx0$ = txtITMx
   txtITMy0$ = txtITMy
   txtGL0$ = txtGL
   txtDepth0$ = txtDepth
   txtNames0$ = txtNames
   
   'copy the coordinates
   txtITMx = GDMDIform.Text5
   txtITMy = GDMDIform.Text6
   
   'allow for undoing the changes
   cmdUndoCoordinates.Enabled = True
   cmdCancel.Enabled = True

End Sub

Private Sub cmdEditAge_Click()
   'use this ages for database values
  With GDEditScannedDBfrm
    If fillEarlyAge Then
       txtPreE0$ = .txtPreE
       txtEarlyAge0$ = .txtEarlyAge
       .txtPreE = .cmbPreAge.Text
       .txtEarlyAge = .cmbAge.Text
    ElseIf fillLateAge Then
       txtPreL0$ = .txtPreL
       txtLaterAge0$ = .txtLaterAge
       .txtPreL = .cmbPreAge.Text
       .txtLaterAge = .cmbAge.Text
       End If
       
    .cmdUndoAge.Enabled = True
    .cmdCancel.Enabled = True
   End With
End Sub

Private Sub cmdEditCoordinates_Click()
   With GDEditScannedDBfrm
      txtNames0$ = .txtNames
      .txtNames = .cmbdbNames.Text
      txtITMx0$ = .txtITMx
      txtITMy0$ = .txtITMy
      .cmdUndoCoordinates.Enabled = True
      .cmdCancel.Enabled = True
   End With
End Sub

Private Sub cmdEditFormation_Click()
   'load up txtFormation with current cmbEditFormation entry
   txtFormation0$ = txtFormation
   txtFormation = cmbFormation.Text
   cmdUndoFormation.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub cmdEnableCombineDB_Click()
   If cmdStartCombineDB.Enabled = False Then
      'enable the automatic combination of the two
      'databases: pal_old_piv.mdb into pal_old.mdb
      '(the latest revision date determines the final record)
      
      'load merg database's name catalogue
      LoadOldDbArrays_piv
      
      nCombine& = -1
      cmdStartCombineDB.Enabled = True
      cmdEnableCombineDB.Caption = "&Disenable"
      cmdEnableCombineDB.BackColor = &HC0C0FF
      cmdEnableCombineDB.ToolTipText = "Click to disenable automatic DB combination mode"
   Else
      'disenable the automatic combination process
      cmdStartCombineDB.Enabled = False
      cmdEnableCombineDB.Caption = "&Enable"
      cmdEnableCombineDB.BackColor = &H8000000F
      cmdEnableCombineDB.ToolTipText = "Enable combining the two versions of Pal_old.mdb"
      End If
End Sub

Private Sub cmdExchange_Click()
   'store the original values
   txtITMx0$ = txtITMx
   txtITMy0$ = txtITMy
   txtGL0$ = txtGL
   txtDepth0$ = txtDepth
   txtNames0$ = txtNames
   
   'exchange the coordinates
   TmpX = val(txtITMx)
   txtITMx = txtITMy
   txtITMy = TmpX
   
   'allow for undoing the changes
   cmdUndoCoordinates.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub cmdFathoms_Click()
 'convert depth in fathoms to depth in meters
  
  'first record current parameter values for undo
  txtNames0$ = txtNames
  txtITMx0$ = txtITMx
  txtITMy0$ = txtITMy
  txtGL0$ = txtGL
  txtDepth0$ = txtDepth
  cmdUndoCoordinates.Enabled = False

  'now convert depth in feet to depth in meters
  txtDepth = CInt(val(txtDepth) * 18.288) * 0.1
  

End Sub

Private Sub cmdMap_Click()
   'attempt to locate the coordinates on the maps
   
    'if map is not visible, then make geo map visible
    If Not GeoMap And Not TopoMap Then
       'display the geo map
        myfile = Dir(picnam$)
        If myfile = sEmpty Or Trim$(picnam$) = sEmpty Then
           response = MsgBox("Can't find map!" & vbLf & _
                      "Use the Files/Geologic map options menu to help find it.", _
                      vbExclamation + vbOKOnly, "GSIDB")
           'take further response
            GeoMap = False
            Exit Sub
        Else
            Screen.MousePointer = vbHourglass
            buttonstate&(3) = 0
            GDMDIform.Toolbar1.Buttons(3).value = tbrUnpressed
            For i& = 4 To 7
              GDMDIform.Toolbar1.Buttons(i&).Enabled = False
            Next i&
            GDMDIform.Toolbar1.Buttons(8).Enabled = True
            GDMDIform.Toolbar1.Buttons(9).Enabled = False
            If buttonstate&(15) = 1 Then 'search still activated
               GDMDIform.Toolbar1.Buttons(15).value = tbrPressed
               End If
                  
            GDMDIform.mnuGeo.Enabled = False 'disenable menu of geo. coordinates display
            GDMDIform.Toolbar1.Buttons(2).value = tbrPressed
            If topos = True Then GDMDIform.Toolbar1.Buttons(3).Enabled = True
            buttonstate&(2) = 1
            GDMDIform.Label1 = lblX
            GDMDIform.Label5 = lblX
            GDMDIform.Label2 = LblY
            GDMDIform.Label6 = LblY
              
            'load up Geo map
            Call ShowGeoMap(0)
            
            End If
            
       End If
       
       'now attempt to place the map at the recorded coordinates
       If Not Digitizing Then
         GDMDIform.Text5 = txtITMx
         GDMDIform.Text6 = txtITMy
         If PicSum Then ret = ShowWindow(GDReportfrm.hWnd, SW_MINIMIZE)
         ce& = 1 'flag to remove old cursor mark and draw new one at the new location
         Call gotocoord 'move the map to the record's coordinates
         'remove the clutter of windows from the screen and bring
         'the map to the top of the Z order
         End If
      
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdNext_Click
' DateTime  : 5/5/2005 07:14
' Author    : Chaim Keller
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdNext_Click()
   On Error GoTo cmdNext_Click_Error

   'next/previous button changes the cmb boxes, which in turn requeries the database
   If StepDocNo Then 'step in serial no
   
        If cmbSerialNo.ListIndex + 2 >= cmbSerialNo.ListCount Then
           cmdNext.Enabled = False
           cmdPrevious.Enabled = True
        Else
           cmdNext.Enabled = True
           cmdPrevious.Enabled = True
           End If
           
        If cmbSerialNo.ListIndex + 1 > cmbSerialNo.ListCount - 1 Then Exit Sub
        cmbSerialNo.ListIndex = cmbSerialNo.ListIndex + 1
        
    Else 'step in id number (O_KEY)
        If cmbOKEYNo.ListIndex + 2 >= cmbOKEYNo.ListCount Then
           cmdNext.Enabled = False
           cmdPrevious.Enabled = True
        Else
           cmdNext.Enabled = True
           cmdPrevious.Enabled = True
           End If
           
        If cmbOKEYNo.ListIndex + 1 > cmbOKEYNo.ListCount - 1 Then
           If cmdStartCombineDB.Enabled = True Then
              cmdStartCombineDB.value = 1
              End If
           Exit Sub
           End If
           
        If cmdStartCombineDB.Enabled = True Then
           'increment in O_Key
           nCombine& = nCombine& + 1
           cmbOKEYNo.ListIndex = nCombine&
        Else
           cmbOKEYNo.ListIndex = cmbOKEYNo.ListIndex + 1
           End If
       
       End If
       
    If chkRefreshTif.Enabled = True And chkRefreshTif.value = vbChecked Then
       'automatically refresh the tif image on stepping
       cmdTifView.value = 1
       End If

   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:

    Select Case Err.Number
       Case 380 'finished merging
          Screen.MousePointer = vbDefault
          cmdStartCombineDB.Enabled = False
          Exit Sub
       Case Else
          Screen.MousePointer = vbDefault
          MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdNext_Click of Form GDEditScannedDBfrm"
    End Select

End Sub

Private Sub cmdPasteCoord_Click()
   With GDEditScannedDBfrm
      txtNames0$ = .txtNames
      txtGL0$ = .txtGL
      txtDepth0$ = .txtDepth
      txtITMx0$ = .txtITMx
      txtITMy0$ = .txtITMy
     
      'remove the "(place)" or "(well)" from the catalogue name
      If InStr(Trim$(.cmbPlaceNames.Text), " (place)") Then
        placeName$ = Mid$(Trim$(.cmbPlaceNames.Text), 1, InStr(Trim$(.cmbPlaceNames.Text), " (place)") - 1)
      ElseIf InStr(Trim$(.cmbPlaceNames.Text), " (well)") Then
        placeName$ = Mid$(Trim$(.cmbPlaceNames.Text), 1, InStr(Trim$(.cmbPlaceNames.Text), " (well)") - 1)
      Else
        placeName$ = Trim$(.cmbPlaceNames.Text)
      End If
      .txtNames = placeName$
      .txtITMx = .txtITMxCat
      .txtITMy = .txtITMyCat
      .txtGL = .txtGLCat
      
      .cmdUndoCoordinates.Enabled = True
      .cmdCancel.Enabled = True
   End With
End Sub

Private Sub cmdPasteLstAges_Click()
   'use last saved ages
   With GDEditScannedDBfrm
      
      .txtPreE = txtPreELast$
      .txtEarlyAge = txtPreEAgeLast$
      .txtPreL = txtPreLLast$
      .txtLaterAge = txtpreLAgeLast$
       txtPreE0$ = .txtPreE
       txtEarlyAge0$ = .txtEarlyAge
       txtPreL0$ = .txtPreL
       txtLaterAge0$ = .txtLaterAge
      
      .cmdUndoAge.Enabled = True
      .cmdCancel.Enabled = True
   End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdPreview_Click
' DateTime  : 4/4/2005 15:04
' Author    : Chaim Keller
' Purpose   : show print preview of the current record
'---------------------------------------------------------------------------------------
'
Private Sub cmdPreview_Click()
   On Error GoTo cmdPreview_Click_Error
    
   If Trim$(cmbOKEYNo.Text) = sEmpty Then
      MsgBox "No file specified (O_Key is blank)!", vbInformation + vbOKOnly, App.Title
      Exit Sub
      End If

   'show or update print preview of the present record
   If Not Previewing Then 'no print preview visible at this time
      PreviewOrderNum& = -val(cmbOKEYNo.Text)
      MaxOrder& = PreviewOrderNum&
      MinOrder& = PreviewOrderNum&
      PrintPreview.Visible = True
   Else 'clear old print preview and present this one
      PreviewOrderNum& = -val(cmbOKEYNo.Text)
      FillPrintCombo
      If Not LoadInit Then
         PreviewPrint
         End If
      ret = ShowWindow(PrintPreview.hWnd, SW_MAXIMIZE)
      End If

   On Error GoTo 0
   Exit Sub

cmdPreview_Click_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPreview_Click of Form GDEditScannedDBfrm"
End Sub

Private Sub cmdPrevious_Click()
   'next/previous button changes the cmb boxes, which in turn requeries the database
   
   If StepDocNo Then 'step in Serial no
   
        If cmbSerialNo.ListIndex - 1 < 0 Then
           cmdPrevious.Enabled = False
           cmdNext.Enabled = True
        Else
           cmdNext.Enabled = False
           cmdPrevious.Enabled = True
           End If
           
        If cmbSerialNo.ListIndex - 1 < 0 Then Exit Sub
        cmbSerialNo.ListIndex = cmbSerialNo.ListIndex - 1
    
    Else 'step in id number (O_KEY)
    
        If cmbOKEYNo.ListIndex - 1 < 0 Then
           cmdPrevious.Enabled = False
           cmdNext.Enabled = True
        Else
           cmdNext.Enabled = False
           cmdPrevious.Enabled = True
           End If
           
        If cmbOKEYNo.ListIndex - 1 < 0 Then Exit Sub
        cmbOKEYNo.ListIndex = cmbOKEYNo.ListIndex - 1
        
       End If
       
    If chkRefreshTif.Enabled = True And chkRefreshTif.value = vbChecked Then
       'automatically refresh the tif image on stepping
       cmdTifView.value = 1
       End If
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdQuick_Click
' DateTime  : 4/4/2005 15:30
' Author    : Chaim Keller
' Purpose   : Perform multiple operations based on stored templates
'             This allows for "quick" editing of files having alike parameters
'---------------------------------------------------------------------------------------
'
Private Sub cmdQuick_Click()
   On Error GoTo cmdQuick_Click_Error

   If Dir$(direct$ & "\EditTemplate.txt") <> sEmpty Then
      'read the template and perform the stored operations
   
      filtemp% = FreeFile
      Open direct$ & "\EditTemplate.txt" For Input As #filtemp%
      Line Input #filtemp%, doclin$
      Line Input #filtemp%, doclin$
      
      'read outcropping or well designation
      Input #filtemp%, oldSourceT%
      If bSampleSources Then
         Select Case oldSourceT%
            Case 0 'neither outcropping or well
            Case 4
               optOutcroppings.value = True
            Case 1, 2, 3
               optWells.value = True
               If oldSourceT% = 2 Then
                  optCore.value = True
               ElseIf oldSourceT% = 3 Then
                  optCuttings.value = True
                  End If
         End Select
         End If
      
      'read fossil type
      Input #filtemp%, foscatT%
      If bFossilTypes Then
         Select Case foscatT%
            Case 1 'Forams
               optForam.value = True
            Case 2 'shekef
               optShekef.value = True
            Case 3 'ostrocodo
               optOstra.value = True
            Case 4 'palynology
               optPaly.value = True
            Case 5 'megafauna
               optMega.value = True
            Case 6 'Nannoplankton
               optNano.value = True
            Case 7 'diatoms
               optDiatom.value = True
            Case 8 'conodonta
               optCono.value = True
         End Select
         End If
      
      'read coordinates flag
      Input #filtemp%, coordflag%
'      Select Case coordflag%
'         Case 0 'don't do anything
'         Case 1 'paste from place/wells catalogue
'            cmdPasteCoord.Value = 1
'         Case 2 'paste from name catalogue
'            cmdEditCoordinates.Value = 1
'      End Select
      Input #filtemp%, txtITMxT$, txtITMyT$, txtNamesT$, txtGLT$
      If bCoordinates Then
         txtITMx = txtITMxT$
         txtITMy = txtITMyT$
         End If
         
      If bNames Then
         txtNames = txtNamesT$
         End If
         
      If bGroundLevels Then
         txtGL = txtGLT$
         End If
      
      Input #filtemp%, tmp1$, tmp2$ 'don't use stored age information
      Input #filtemp%, tmp3$, tmp4$
      If bGeologicAges Then
         txtPreE = tmp1$
         txtEarlyAge = tmp2$
         txtPreL = tmp3$
         txtLaterAge = tmp4$
         End If
         
      Input #filtemp%, txtFormationT$ 'use stored formation information
      If bFormations Then
         txtFormation = txtFormationT$
         End If
      
'      'restore early date (prefix, age)
'      Input #filtemp%, txtPreET$, txtEarlyAgeT$
'      txtPreE = txtPreET
'      txtEarlyAge = txtEarlyAgeT$
'
'      'restore latter date (prefix, age)
'      Input #filtemp%, txtPreLT$, txtLaterAgeT$
'      txtPreL = txtPreLT$
'      txtLaterAge = txtLaterAgeT$

      Input #filtemp%, txtDepths$ 'use sotred depth
      If bDepths Then
         txtDepth = txtDepths$
         End If
         
      Close #filtemp%
        
   Else
      MsgBox "Can't find the template file!" & _
             vbCrLf & vbCrLf & "Hint: ""Save"" a template before using this option." & _
             vbCrLf & "(For Help, consult the documentation)", _
             vbExclamation + vbOKOnly, App.Title
      Exit Sub
      End If

'------------------------former uses for this button------------------
'   'used as a patch for quick work of 2 digit coordinate files
'
'   'first store the original coordinate values
'   txtITMx0$ = txtITMx
'   txtITMy0$ = txtITMy
'   txtGL0$ = txtGL
'   txtDepth0$ = txtDepth
'   txtNames0$ = txtNames
'
'   With GDEditScannedDBfrm
'
'      'Western Galilee series
'      .cmdPasteCoord.Value = 1
'      .optWells.Value = True
'      .optCore.Value = True
'      .cmdTifView.Value = True
'      .optForam.Value = True
'      Exit Sub
'
'
'      If .txtMinDigits = "2" And .txtMaxDigits = "2" Then
'         .txtITMx = .txtITMx + "0000"
'         If Val(.txtITMy) < 30 Then
'            .txtITMy = "1" & .txtITMy + "0000"
'         Else 'Negev, Sinai
'            .txtITMy = .txtITMy + "0000"
'            End If
'         .optMega.Value = True
'         .optOutcroppings.Value = True
'         'view tif file
'         .cmdTifView.Value = 1
'      ElseIf .txtMinDigits = "2" And .txtMaxDigits = "3" Then
'         If Val(.txtITMx) < 100 And Mid$(.txtITMx, 1, 1) <> "0" Then
'            .txtITMx = "10" & .txtITMx & "00"
'         ElseIf Mid$(.txtITMx, 1, 1) = "0" Then
'            If Val(.txtITMx) > 9 Then
'               .txtITMx = "1" & .txtITMx & "000"
'            Else
'               .txtITMx = "1" & .txtITMx & "00"
'               End If
'         'ElseIf Val(.txtITMx) >= 80 And Val(.txtITMx) < 100 Then
'         '   .txtITMx = .txtITMx & "0000"
'         ElseIf Val(.txtITMx) > "799" Then 'Sinai
'            .txtITMx = .txtITMx & "00"
'         Else 'regular coordinates
'            .txtITMx = .txtITMx & "000"
'            End If
'
'         If Val(.txtITMy) < 100 And Mid$(.txtITMy, 1, 1) <> "0" Then
'            .txtITMy = "10" & .txtITMy & "000"
'         ElseIf Mid$(.txtITMy, 1, 1) = "0" Then
'            If Val(.txtITMy) > 9 Then
'               .txtITMy = "1" & .txtITMy & "000"
'            Else
'               .txtITMy = "1" & .txtITMy & "0000"
'               End If
'         ElseIf Val(.txtITMy) >= 800 Then
'            .txtITMy = .txtITMy & "000"
'         Else
'            .txtITMy = "1" & .txtITMy & "000"
'            End If
'
'         End If
'
'      'other defaults
'      .optMega.Value = True
'      .optOutcroppings.Value = True
'      .cmdTifView.Value = True
'
'   End With

   On Error GoTo 0
   Exit Sub

cmdQuick_Click_Error:

    Close
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdQuick_Click of Form GDEditScannedDBfrm"
End Sub

Private Sub cmdQuickOptions_Click()
  If QuickVis Then
     BringWindowToTop (GDQuickButtonfrm.hWnd)
  Else
     GDQuickButtonfrm.Visible = True
     End If
End Sub

Private Sub cmdSave_Click()
   'save changes to the scanned database record
   
   If cmdStartCombineDB.Enabled = True Or GotPassword = True Then GoTo GotPWD 'password already inputed

   'if this is the first time changes are made, then
   'need to enter password
   Dim frmPWD As New GDfrmDBPWD
   Dim Pwrd As String
   Pwrd = "neal"

GetPWD:
   frmPWD.Show
GetPWD2:
   DoEvents
   If Len(frmPWD.PWD) > 0 And Not PwdCancel Then
      If LCase$(frmPWD.PWD) = Pwrd Then 'success
         GotPassword = True 'got the password
         Unload frmPWD 'unload the password form, and reactivate the editor
         Set frmPWD = Nothing
         ret = ShowWindow(GDEditScannedDBfrm.hWnd, SW_NORMAL)
         GoTo GotPWD
      Else
         Select Case MsgBox("The password you entered was not correct!" & vbLf _
                            & "(See the documentation for more information)" & vbLf _
                            & vbCrLf & "Try again?" _
                            , vbOKCancel Or vbExclamation Or vbSystemModal Or vbDefaultButton1, App.Title)
         
            Case vbOK
                Unload frmPWD
                Set frmPWD = Nothing
                GoTo GetPWD
         
            Case vbCancel
                PwdCancel = False
                Unload frmPWD
                Set frmPWD = Nothing
                Exit Sub
         
         End Select
         End If
   ElseIf PwdCancel Then
     'they cancelled the pwd dialog so we need to exit
     PwdCancel = False
     Unload frmPWD
     Set frmPWD = Nothing
     ret = ShowWindow(GDEditScannedDBfrm.hWnd, SW_NORMAL)
     Exit Sub
   ElseIf Len(frmPWD.PWD) = 0 Then 'loop
     GoTo GetPWD2
   End If
   
GotPWD:
   
   With GDEditScannedDBfrm
   
       
'--------------------fix up last editing--------------------
       If cmdStartCombineDB.Enabled = True Then
       
 '        '(1)fix up Senon replacing Turonian to Maastrichtian
 '        'with Coniacian to Campanian
 '        If Trim$(.txtEarlyAge) = "Turonian" And _
 '           Trim$(.txtLaterAge) = "Maastrichtian" Then
 '           DoEvents
 '           .txtEarlyAge = "Coniacian"
 '           .txtLaterAge = "Campanian"
 '           End If
            
         '(2)if well and core/cuttings empty, make it cuttings
         If oldSource% = 1 Then
            .optCuttings.value = True
            'This operation sets oldSource% = 3, i.e.,  cuttings
            End If
         
         End If
'-------------------------------------------------------
                  
      'save last age parameters
      txtPreELast$ = .txtPreE
      txtPreEAgeLast$ = .txtEarlyAge
      txtPreLLast$ = .txtPreL
      txtpreLAgeLast$ = .txtLaterAge
      .cmdPasteLstAges.Enabled = True
      
      If .txtITMx.Text <> txtITMx00$ Or .txtITMy.Text <> txtITMy00$ Or _
         .txtPreE.Text <> txtPreE00$ Or .txtEarlyAge.Text <> txtEarlyAge00$ Or _
         .txtPreL.Text <> txtPreL00$ Or .txtLaterAge.Text <> txtLaterAge00$ Or _
         .txtFormation.Text <> txtFormation00$ Or .txtNames.Text <> txtNames00$ Or _
         foscat% <> foscat00% Or oldSource% <> oldSource00% Or _
         .txtGL.Text <> txtGL00$ Or .txtDepth.Text <> txtDepth00$ Then
         
         If cmdStartCombineDB.Enabled = False And chkSaveVerify.value = vbChecked Then
            res = MsgBox("Commit changes to the database?", vbQuestion + vbYesNo, "Save changes to scanned database?")
            If res <> vbYes Then Exit Sub
         Else 'automatically combining databases, so don't ask
            End If
            
         'check for illegal characters
         If InStr(.txtNames, "&") Then
            Select Case MsgBox("An illegal ""&"" character was found in the place name." _
                               & vbCrLf & "This character will cause problems if the result is plotted using ""Google Earth""" _
                               & vbCrLf & "" _
                               & vbCrLf & "Is it OK to replace the ""&"" with a ""+""?" _
                               , vbOKCancel Or vbQuestion Or vbDefaultButton1, "Editing scannied database")
            
               Case vbOK
               
                  .txtNames = Replace(.txtNames, "&", "+")
            
               Case vbCancel
            
            End Select
            End If
            
         'save the changes
         SaveEditChanges
         
         End If
   
   End With
   
   If EditScannedDBVis Then 'restore the GDAddScannedFiles form in the z order, etc
      GDMDIform.WindowState = vbNormal 'not maximized and not minimized
      GDMDIform.Width = GDAddScannedFiles.Width + 250
      GDMDIform.Height = Screen.Height - 400
      BringWindowToTop (GDAddScannedFiles.hWnd)
      End If
   
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSaveTemplate_Click
' DateTime  : 4/4/2005 14:33
' Author    : Chaim Keller
' Purpose   : Save current settings as template
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveTemplate_Click()

   On Error GoTo cmdSaveTemplate_Click_Error
   
   'before writing changes warn the user
   
   If Dir$(direct$ & "\EditTemplate.txt") = sEmpty Then
   
      Select Case MsgBox("Store this the current template?", vbYesNo Or vbQuestion Or vbDefaultButton1, App.Title)
   
        Case vbYes
    
        Case vbNo
          Exit Sub
      End Select
  
  Else
      
      Select Case MsgBox("Overwrite the last template and store this one in its place?", vbYesNoCancel Or vbQuestion Or vbDefaultButton1, App.Title)
    
        Case vbYes
    
        Case vbNo, vbCancel
           Exit Sub
    
      End Select
  
      End If
      
   Screen.MousePointer = vbHourglass
   
   'determine coordinate flag
   'if coordinates and names are the same as in the catalogues, then coordinate flag=1
   If txtITMx = txtITMxCat And txtITMy = txtITMyCat And txtNames = cmbPlaceNames.Text And _
      txtGL = txtGLCat Then
      'this means that pasted from catalogue of places/wells
      coordflag% = 1
   ElseIf txtNames = cmbdbNames.Text And _
      (txtITMx <> txtITMxCat Or txtITMy <> txtITMyCat Or txtGL <> txtGLCat) Then
      'this menas that pasted from catalogue of names
      coordflag% = 2
   Else
      coordflag% = 0
      End If
      
   'store the template
   filtemp% = FreeFile
   Open direct$ & "\EditTemplate.txt" For Output As #filtemp%
   Print #filtemp%, "This file is used by the MapDigitizer program, don't erase it! "
   Print #filtemp%, "Scanned database editing template, revised: " & Date
   'store outcropping or well designation
   Print #filtemp%, oldSource%
   'store fossil type
   Print #filtemp%, foscat%
   'store coordinates flag, coordinates, and names
   Print #filtemp%, coordflag% '(this option is not used as of yet)
   Write #filtemp%, txtITMx, txtITMy, txtNames, txtGL
   'store early date (prefix, age)
   Write #filtemp%, txtPreE, txtEarlyAge
   'store latter date (prefix, age)
   Write #filtemp%, txtPreL, txtLaterAge
   'store formation name
   Write #filtemp%, txtFormation
   Write #filtemp%, txtDepth
   Close #filtemp%
   
   Screen.MousePointer = vbDefault
   
   On Error GoTo 0
   Exit Sub

cmdSaveTemplate_Click_Error:

    Screen.MousePointer = vbDefault
    Close
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSaveTemplate_Click of Form GDEditScannedDBfrm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdStartCombineDB_Click
' DateTime  : 3/4/2005 15:46
' Author    : Chaim Keller
' Purpose   : Combine the pal_old.mdb and pal_old_piv.mdb databases
'             based on which has the most recent revisions
'---------------------------------------------------------------------------------------
'
Private Sub cmdStartCombineDB_Click()
     
   On Error GoTo cmdStartCombineDB_Click_Error

   Do Until cmdStartCombineDB.Enabled = False
     'compare and then go to next entry if combining data bases
     DoEvents
     'compare the values of pal_old_piv.mdb
     'and use the most recently revised records.
     LoadEditFormpiv
     'go to next entry
     GDEditScannedDBfrm.cmdNext.value = 1
   Loop

   On Error GoTo 0
   Exit Sub

cmdStartCombineDB_Click_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdStartCombineDB_Click of Form GDEditScannedDBfrm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdTifView_Click
' DateTime  : 5/5/2004 10:47
' Author    : Chaim Keller
' Purpose   : Queries OBJECTS_VER2 table to find the tif file path
'             corresponding to the current O_KEY
'---------------------------------------------------------------------------------------
'
Private Sub cmdTifView_Click()

   On Error GoTo cmdTifView_Click_Error
   
   If Dir(tifViewerDir$) <> sEmpty And Trim$(cmbOKEYNo.Text) <> sEmpty Then
      numOKey& = CLng(GDEditScannedDBfrm.cmbOKEYNo.Text)
      Call FindTifPath(numOKey&, numOFile$)
      
      Select Case numOFile$
         Case "-1" 'error flag
            MsgBox "Tif file not found!" & _
                   vbCrLf & "(Apparently no tif file is associated with this record)", _
                   vbInformation + vbOKOnly, App.Title
         Case Else 'view the file
            If Dir(tifDir$ & "\" & UCase$(numOFile$)) <> sEmpty Then
               Shell (tifCommandLine$ & " " & tifDir$ & "\" & numOFile$)
            Else
               Call MsgBox("The path: " & tifDir$ & "\" & UCase$(numOFile$) & " was not found or is not accessible!" & vbLf & vbLf & _
                           "Check the defined path to the tif files in the options menu, and try again", vbExclamation + vbOKOnly, App.Title)
               End If
      End Select
      
   ElseIf Dir(tifViewerDir$) = sEmpty Then
      cmdTifView.Enabled = False
      chkRefreshTif.Enabled = False
      cmdTifView.Enabled = False
      MsgBox "The path to the tif file viewer is no longer defined!" & _
             vbCrLf & "(See help documentation on how to set it.)", _
             vbExclamation + vbOKOnly, App.Title
             
   ElseIf Trim$(cmbOKEYNo.Text) = sEmpty Then
      MsgBox "No file specified (O_Key is blank)!", vbInformation + vbOKOnly, App.Title
      
      End If

   On Error GoTo 0
   Exit Sub

cmdTifView_Click_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdTifView_Click of Form GDEditScannedDBfrm"
End Sub

Private Sub cmdUndoAge_Click()
   'undo the last changes
   With GDEditScannedDBfrm
    
    If PasteLstAges Then 'unpaste both ages
       .txtPreE = txtPreE0$
       .txtEarlyAge = txtEarlyAge0$
       .txtPreL = txtPreL0$
       .txtLaterAge = txtLaterAge0$
       .cmdUndoAge.Enabled = False
       Exit Sub
       End If
    
    If fillEarlyAge Then
       .txtPreE = txtPreE0$
       .txtEarlyAge = txtEarlyAge0$
    ElseIf fillLateAge Then
       .txtPreL = txtPreL0$
       .txtLaterAge = txtLaterAge0$
       End If
    .cmdUndoAge.Enabled = False
   End With
End Sub

Private Sub cmdUndoCoordinates_Click()
   txtNames = txtNames0$
   txtITMx = txtITMx0$
   txtITMy = txtITMy0$
   txtGL = txtGL0$
   txtDepth = txtDepth0$
   cmdUndoCoordinates.Enabled = False
End Sub

Private Sub cmdUndoCopyEtoL_Click()
  'undo last copy of early to late ages
  With GDEditScannedDBfrm
     .txtPreL = txtPreL0$
     .txtLaterAge = txtLaterAge0$
     .cmdUndoCopyEtoL.Enabled = False
  End With

End Sub

Private Sub cmdUndoFormation_Click()
   txtFormation = txtFormation0$
   cmdUndoFormation.Enabled = False
End Sub

Private Sub cmdUndoFos_Click()
   'undo the last change to the fossil category
     Select Case foscat00%
       Case 0 'blank
         optCono.value = False
         optDiatom.value = False
         optForam.value = False
         optMega.value = False
         optNano.value = False
         optOstra.value = False
         optPaly.value = False
         optShekef.value = False
       Case 1 'foram
         optForam.value = True
       Case 2 'foram (shekef)
         optShekef.value = True
       Case 3 'ostra
         optOstra.value = True
       Case 4 'paly
         optPaly.value = True
       Case 5 'mega
         optMega.value = True
       Case 6 'nanno
         optNano.value = True
       Case 7 'diatom
         optDiatom.value = True
       Case 8 'cono
         optCono.value = True
     End Select
     
     cmdUndoFos.Enabled = False
     
End Sub

Private Sub cmdUndoSources_Click()
   'undo the last change to the Sample Source
   optOutcroppings.value = False
   optWells.value = False
   optCore.value = False
   optCuttings.value = False
   Select Case oldSource00%
      Case 0
         oldSource% = 0
      Case 1
         optWells.value = True
      Case 2
         optWells.value = True
         optCore.value = True
      Case 3
         optWells.value = True
         optCuttings.value = True
      Case 4
         optOutcroppings.value = True
   End Select
   cmdUndoSources.Enabled = False
End Sub

Private Sub cmdWizard_Click()
   'suggest possible coordinates
   SuggestEditCoord
End Sub

Private Sub cmdCopyEarlytoLate_Click()
  'backup dates
  With GDEditScannedDBfrm
     txtPreL0$ = .txtPreL
     txtLaterAge0$ = .txtLaterAge
     'copy early ages to latter ages
     .txtLaterAge = .txtEarlyAge
     .txtPreL = .txtPreE
     .cmdUndoCopyEtoL.Enabled = True
     .cmdCancel.Enabled = True
  End With
End Sub

Private Sub Form_Activate()
   'make button stay pressed
   GDMDIform.Toolbar1.Buttons(12).value = tbrPressed
   buttonstate&(12) = 1
   'ret = ShowWindow(GDEditScannedDBfrm.hWnd, SW_NORMAL)
End Sub

Private Sub Form_Deactivate()
   'unpress button to encourage user to press it again inorder to activate form
   GDMDIform.Toolbar1.Buttons(12).value = tbrUnpressed
   buttonstate&(12) = 0
   'ret = ShowWindow(GDEditScannedDBfrm.hWnd, SW_MINIMIZE)
End Sub

Private Sub Form_Load()
  'load up forms
  EditDBVis = True
  GDMDIform.Toolbar1.Buttons(12).value = tbrPressed
  buttonstate&(12) = 1
  LoadingEditForm = True
  optEDS% = 1 'default searches is using number of digits in coordinates
  
  StepDocNo = False 'default next/previous step is in document O_KEY number
  
  Screen.MousePointer = vbHourglass

  With GDEditScannedDBfrm
     .Top = 0
     .Left = 0
     
     OKeyClick = False
     ONameClick = False
     If Not Previewing And Not EditScannedDBVis Then modeEdit% = 0
     
     ReloadEditForm 'load up all the arrays and combo boxes
     
    '------Progress Bar Settings--------------
    picProgBar.AutoRedraw = True
    picProgBar.BackColor = &H8000000B 'light grey
    picProgBar.DrawMode = 10
    
    picProgBar.FillStyle = 0
    picProgBar.ForeColor = &H400000 'dark blue
    
    'other defaults
    .cmbOKEYNo.Text = sEmpty
    .cmbSerialNo.Text = sEmpty
    .txtITMx = sEmpty
    .txtITMy = sEmpty
    .cmbPlaceNames.Text = sEmpty
    .txtITMxCat = sEmpty
    .txtITMyCat = sEmpty
    .txtGLCat = sEmpty
    .txtPreE = sEmpty
    .txtEarlyAge = sEmpty
    .txtPreL = sEmpty
    .txtLaterAge = sEmpty
    .txtFormation = sEmpty
    .txtNames = sEmpty
    .txtGL = sEmpty
    .txtDepth = sEmpty
    
    'set defaults
    txtITMx00$ = sEmpty
    txtITMy00$ = sEmpty
    txtGL00$ = sEmpty
    txtPreE00$ = sEmpty
    txtEarlyAge00$ = sEmpty
    txtPreL00$ = sEmpty
    txtLaterAge00$ = sEmpty
    txtFormation00$ = sEmpty
    txtNames00$ = sEmpty
    txtDepth00$ = sEmpty
    foscat% = 0
    foscat00% = 0
    oldSource% = 0
    oldSource00% = 0
    
    'disenable editing boxes
    .frmSource.Enabled = False
    .frmCoordinates.Enabled = False
    .frmAges.Enabled = False
    .frmFormations.Enabled = False
    .frmFossils.Enabled = False
    .frmModify.Enabled = False
      
    'if got to this routine from PrintPreview then
    'disenable the DataSource frame and load up the relevant
    'record
    If Previewing Or EditScannedDBVis Then
       
       '.optAllSearch.Enabled = False
       '.optWellSearch.Enabled = False
       '.optOutcroppingSearch.Enabled = False
       '.optReset.Enabled = False
       '.lblMax.Enabled = False
       '.lblMin.Enabled = False
       '.txtMaxDigits.Enabled = False
       '.txtMinDigits.Enabled = False
       '.upMaxDigits.Enabled = False
       '.udMinDigits.Enabled = False
       .cmbOKEYNo.Clear
       .cmbSerialNo.Clear
       .cmdNext.Enabled = False
       .cmdPrevious.Enabled = False
       BringWindowToTop (GDEditScannedDBfrm.hWnd)
       '.Line1.BorderColor = &H8000000C
       '.Line2.BorderColor = &H8000000C
       '.Line3.BorderColor = &H8000000C
       '.Line4.BorderColor = &H8000000C
       '.Line5.BorderColor = &H8000000C
       LoadEditForm
       End If
       
    .cmdSave.Enabled = False
       
    If PicSum Then optReport.Enabled = True 'can search over search results
    
    LoadingEditForm = False
    
    'if tif viewer available, then enable button
    If Dir(tifViewerDir$) <> sEmpty Then
       cmdTifView.Enabled = True
       chkRefreshTif.Enabled = True
       End If
    
    'if Report Form not available then disenable optoin to load Search Report records
    If Not PicSum Then optReport.Enabled = False
    
    'if pal_old_piv.mdb available for combining to pal_old.mdb, then activate features
    If linkedpiv Then
       frmCombineDB.Visible = True
       End If
     
  End With

  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   modeEdit% = -1 'unflag any type of search
   
   If GDEditScannedDBfrm.picProgBar.Enabled Then
      'can't unload during search
      Cancel = True
      Exit Sub
      End If

   Set GDEditScannedDBfrm = Nothing
   EditDBVis = False
   GDMDIform.Toolbar1.Buttons(12).value = tbrUnpressed
   buttonstate&(12) = 0
   CheckDuplicatePoints = False 'now recheck the points for duplicates
        'for being off the map, etc. (This checking was shut off as
        'long as this form was visible to allow for placing the points
        'on the map without getting error messages).
        
   If EditScannedDBVis Then 'restore the GDAddScannedFiles form in the z order, etc
      GDMDIform.WindowState = vbNormal 'not maximized and not minimized
      GDMDIform.Width = GDAddScannedFiles.Width + 250
      GDMDIform.Height = Screen.Height - 400
      BringWindowToTop (GDAddScannedFiles.hWnd)
      End If
        
End Sub

Private Sub lstWizard_Click()
   'store the old values
   txtITMx0$ = txtITMx
   txtITMy0$ = txtITMy
   txtGL0$ = txtGL
   txtDepth0$ = txtDepth
   txtNames0$ = txtNames
   
   'load the selected suggestion to the ITMx, ITMy textboxes
   For i& = 1 To lstWizard.ListCount
      If lstWizard.Selected(i& - 1) Then
         doclin$ = lstWizard.List(i& - 1)
         pos% = InStr(1, doclin$, ",")
         txtITMx = Mid$(doclin$, 1, pos% - 1)
         txtITMy = Mid$(doclin$, pos% + 2, Len(doclin$) - pos% - 1)
         End If
   Next i&
   
   'allow for undoing the changes
   cmdUndoCoordinates.Enabled = True
   cmdCancel.Enabled = True
   
End Sub

Private Sub optAllSearch_Click()
   'load the Edit forms with all the records
   If Not LoadingEditForm Then
      modeEdit% = 0
      OKeyClick = False
      ONameClick = False
      minDigits = val(txtMinDigits)
      maxDigits = val(txtMaxDigits)
      LoadEditForm
      End If
End Sub

Private Sub optCono_Click()
   foscat% = 8
   If foscat% <> foscat00% Then
      cmdUndoFos.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub optCore_Click()
   oldSource% = 2
   If oldSource% <> oldSource00% Then
      cmdUndoSources.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub optCuttings_Click()
   oldSource% = 3
   If oldSource% <> oldSource00% Then
      cmdUndoSources.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub optDiatom_Click()
   foscat% = 7
   If foscat% <> foscat00% Then
      cmdUndoFos.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub optDigits_Click()
   'search using number of digits in coordinates
   optEDS% = 1
   
   lblMin.Caption = "Min No of Digits"
   lblMax.Caption = "Max No. of Digits"
   txtMinDigits.Left = txtMinDigits.Left + 150
   txtMinDigits.Width = txtMinDigits.Width - 150
   udMinDigits.Max = 7
   If val(txtMinDigits) > 7 Then txtMinDigits = 0
   txtMaxDigits.Left = txtMaxDigits.Left + 150
   txtMaxDigits.Width = txtMaxDigits.Width - 150
   udMaxDigits.Max = 7
   If val(txtMaxDigits) > 7 Then txtMaxDigits = 7

End Sub

Private Sub optEarly_Click()
   fillEarlyAge = True
   fillLateAge = False
End Sub

Private Sub optForam_Click()
  foscat% = 1
  If foscat% <> foscat00% Then
     cmdUndoFos.Enabled = True
     cmdCancel.Enabled = True
     cmdSave.Enabled = True
     End If
End Sub

Private Sub optLate_Click()
   fillEarlyAge = False
   fillLateAge = True
End Sub

Private Sub optLoadPlacesCat_Click()
 'load up catalogue of wells and NIMA places
 LoadPlaces% = 1
 LoadPlaceCat
End Sub

Private Sub optMega_Click()
  foscat% = 5
  If foscat% <> foscat00% Then
     cmdUndoFos.Enabled = True
     cmdCancel.Enabled = True
     cmdSave.Enabled = True
     End If
End Sub

Private Sub optNano_Click()
   foscat% = 6
  If foscat% <> foscat00% Then
     cmdUndoFos.Enabled = True
     cmdCancel.Enabled = True
     cmdSave.Enabled = True
     End If
End Sub

Private Sub optOKeys_Click()
   optEDS% = 2 'search for records over range of O_Keys

   lblMin.Caption = "Min O_KEY"
   lblMax.Caption = "Max O_KEY"
   txtMinDigits.Left = txtMinDigits.Left - 150
   txtMinDigits.Width = txtMinDigits.Width + 150
   udMinDigits.Max = 60000
   txtMaxDigits.Left = txtMaxDigits.Left - 150
   txtMaxDigits.Width = txtMaxDigits.Width + 150
   udMaxDigits.Max = 60000
   
End Sub

Private Sub optOstra_Click()
   foscat% = 3
  If foscat% <> foscat00% Then
     cmdUndoFos.Enabled = True
     cmdCancel.Enabled = True
     cmdSave.Enabled = True
     End If
End Sub

Private Sub optOutcroppings_Click()
   optCore.value = False
   optCuttings.value = False
   frmCore.Enabled = False
   optCore.Enabled = False
   optCuttings.Enabled = False
   oldSource% = 4
   If oldSource% <> oldSource00% Then
      cmdUndoSources.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub optOutcroppingSearch_Click()
   'load the Edit forms with records from the surface
   If Not LoadingEditForm Then
      modeEdit% = 2
      minDigits = val(txtMinDigits)
      maxDigits = val(txtMaxDigits)
      OKeyClick = False
      ONameClick = False
      LoadEditForm
      End If
End Sub

Private Sub optPaly_Click()
   foscat% = 4
  If foscat% <> foscat00% Then
     cmdUndoFos.Enabled = True
     cmdCancel.Enabled = True
     cmdSave.Enabled = True
     End If
End Sub

Private Sub optReport_Click()
   'load the Edit forms with all the search report records
   'from the old database
   
   If Not LoadingEditForm Then
      
      LoadingEditForm = True
   
      Screen.MousePointer = vbHourglass
      
      If Previewing Then 'save current order numbers for redisplay
         OKeyVal$ = cmbOKEYNo.Text
         ONameVal$ = cmbSerialNo.Text
         End If
      
      cmbOKEYNo.Clear
      cmbSerialNo.Clear
      
      GDEditScannedDBfrm.picProgBar.Visible = True
      GDEditScannedDBfrm.picProgBar.Enabled = True
      pbScaleWidth = numReport&
      GDEditScannedDBfrm.picProgBar.ScaleWidth = pbScaleWidth
      
      'scan the report form for order numbers of the old database
      numLoaded& = 0
      For i& = 1 To numReport&
         If InStr(GDReportfrm.lvwReport.ListItems(i&).SubItems(4), "*") Then
            'old scanned database record
            pos1& = InStr(GDReportfrm.lvwReport.ListItems(i&).SubItems(4), "*")
            pos2& = InStr(GDReportfrm.lvwReport.ListItems(i&).SubItems(4), "/")
            cmbOKEYNo.AddItem Mid$(GDReportfrm.lvwReport.ListItems(i&).SubItems(4), _
                    pos1& + 1, pos2& - pos1& - 1)
            cmbSerialNo.AddItem Mid$(GDReportfrm.lvwReport.ListItems(i&).SubItems(4), _
                    pos2& + 1, Len(GDReportfrm.lvwReport.ListItems(i&).SubItems(4)) - pos2&)
            numLoaded& = numLoaded& + 1
            End If
            
            Call UpdateStatus(GDEditScannedDBfrm, 1, i&)
      Next i&
      
      GDEditScannedDBfrm.picProgBar.Visible = False
      GDEditScannedDBfrm.picProgBar.Enabled = False
 
      If numLoaded& = 0 Then
         Screen.MousePointer = vbDefault
         MsgBox "Can't edit any of the records in the search report" & vbLf & _
                "since none of them derive from the scanned database!" & vbLf & vbLf & _
                "To edit these files, take note of their order numbers" & vbLf & _
                "and then press the Access program button/menu.", vbExclamation + vbOKOnly, "MapDigitizer Edit Scanned DB Error"
         Exit Sub
         End If
      
      If Not Previewing Then
         'go to beginning of search report list
         cmbOKEYNo.ListIndex = 0
         cmbSerialNo.ListIndex = 0
      Else
         'place combo boxes at current previewed record
         cmbOKEYNo.Text = OKeyVal$
         cmbSerialNo.Text = ONameVal$
         End If
      
      Screen.MousePointer = vbDefault
      
      If cmbOKEYNo.ListCount = 0 Then
         MsgBox "No scanned database records found in the search report!", _
                vbExclamation + vbOKOnly, "MapDigitizer"
         Exit Sub 'then no old db records in the search report
         End If
         
      modeEdit% = 3
      OKeyClick = False
      ONameClick = True 'load up info for first O_NAME entry
      LoadEditForm
      End If
      
End Sub

Private Sub optShekef_Click()
   foscat% = 2
  If foscat% <> foscat00% Then
     cmdUndoFos.Enabled = True
     cmdCancel.Enabled = True
     cmdSave.Enabled = True
     End If
End Sub

Private Sub optStepOKeyNo_Click()
   'step in O_KEY (id) number when next or previous buttons pressed
   StepDocNo = False
End Sub

Private Sub optStepSerialNo_Click()
   'step in Document Serial No when next or previous buttons pressed
   StepDocNo = True
End Sub

Private Sub optWells_Click()
   frmCore.Enabled = True
   optCore.Enabled = True
   optCuttings.Enabled = True
   oldSource% = 1
   If oldSource% <> oldSource00% Then
      cmdUndoSources.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub optWellscat_Click()
'load up Wellscat table of the active database
LoadPlaces% = 2
LoadWellscat
End Sub

Private Sub optWellSearch_Click()
   'load the Edit forms with well data
   If Not LoadingEditForm Then
      modeEdit% = 1
      minDigits = val(txtMinDigits)
      maxDigits = val(txtMaxDigits)
      OKeyClick = False
      ONameClick = False
      LoadEditForm
      End If
End Sub

Private Sub txtDepth_Change()
   If txtDepth.Text <> txtDepth00$ Then
     cmdUndoCoordinates.Enabled = True
     cmdCancel.Enabled = True
     cmdSave.Enabled = True
     End If
End Sub

Private Sub txtEarlyAge_Change()
   If txtEarlyAge.Text <> txtEarlyAge00$ Then
      cmdUndoAge.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtFormation_Change()
   If txtFormation.Text <> txtFormation00$ Then
      cmdUndoFormation.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtGL_Change()
   If txtGL.Text <> txtGL00$ Then
      cmdUndoCoordinates.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtITMx_Change()
   If txtITMx <> txtITMx00$ Then
      cmdUndoCoordinates.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtITMy_Change()
   If txtITMy <> txtITMy00$ Then
      cmdUndoCoordinates.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtLaterAge_Change()
   If txtLaterAge.Text <> txtLaterAge00$ Then
      cmdUndoAge.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtNames_Change()
   If txtNames.Text <> txtNames00$ Then
      cmdUndoCoordinates.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtPreE_Change()
   If txtPreE.Text <> txtPreE00$ Then
      cmdUndoAge.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub

Private Sub txtPreL_Change()
   If txtPreL.Text <> txtPreL00$ Then
      cmdUndoAge.Enabled = True
      cmdCancel.Enabled = True
      cmdSave.Enabled = True
      End If
End Sub
