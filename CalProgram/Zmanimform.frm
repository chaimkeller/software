VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Zmanimform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Z'manim Parameters"
   ClientHeight    =   7635
   ClientLeft      =   3540
   ClientTop       =   840
   ClientWidth     =   5370
   Icon            =   "Zmanimform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   5370
   Begin VB.Frame frmParshiot 
      Caption         =   "Parshot HaShavua"
      Height          =   535
      Left            =   80
      TabIndex        =   91
      Top             =   6480
      Width           =   5235
      Begin VB.OptionButton optDiasporaParshiot 
         Caption         =   "Sedra of Diaspora"
         Height          =   255
         Left            =   3480
         TabIndex        =   94
         Top             =   220
         Width           =   1695
      End
      Begin VB.OptionButton optEYParshiot 
         Caption         =   "Sedra of Eretz Yisroel"
         Height          =   195
         Left            =   1480
         TabIndex        =   93
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optNoParshiot 
         Caption         =   "No Sedra"
         Height          =   195
         Left            =   240
         TabIndex        =   92
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton savebut 
      Height          =   435
      Left            =   2400
      Picture         =   "Zmanimform.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Save the z'manim parameters as a template"
      Top             =   7080
      Width           =   555
   End
   Begin VB.CommandButton loadbut 
      Height          =   435
      Left            =   1860
      Picture         =   "Zmanimform.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Load stored z'manim parameters' template"
      Top             =   7080
      Width           =   555
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H8000000B&
      Caption         =   "Pick a subset of zemanim to be displayed and/ or reorder the list manually"
      Height          =   1035
      Left            =   60
      TabIndex        =   77
      Top             =   5400
      Width           =   5235
      Begin VB.CommandButton Command6 
         Caption         =   "Re&load"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4440
         TabIndex        =   80
         ToolTipText     =   "Load"
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Reorder"
         Height          =   615
         Left            =   120
         Picture         =   "Zmanimform.frx":1116
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "To enter, click desired item on Z'manim name list"
         Top             =   240
         Width           =   675
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         Height          =   645
         Left            =   840
         TabIndex        =   78
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.CommandButton calendarbut 
      Height          =   435
      Left            =   2880
      Picture         =   "Zmanimform.frx":1558
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Calculate these z'manim for this table"
      Top             =   7080
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Z'manim parameters"
      Height          =   4455
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   5235
      Begin MSComCtl2.UpDown UpDown12 
         Height          =   285
         Left            =   3720
         TabIndex        =   74
         Top             =   3720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text16"
         BuddyDispid     =   196621
         OrigLeft        =   4320
         OrigTop         =   3120
         OrigRight       =   4560
         OrigBottom      =   3375
         Max             =   60
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3180
         TabIndex        =   73
         Text            =   "60"
         Top             =   3740
         Width           =   555
      End
      Begin MSComCtl2.UpDown UpDown11 
         Height          =   285
         Left            =   3720
         TabIndex        =   72
         Top             =   3480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text15"
         BuddyDispid     =   196622
         OrigLeft        =   4260
         OrigTop         =   2820
         OrigRight       =   4500
         OrigBottom      =   3135
         Max             =   60
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3180
         TabIndex        =   71
         Text            =   "60"
         Top             =   3480
         Width           =   555
      End
      Begin VB.OptionButton Option18 
         BackColor       =   &H8000000B&
         Caption         =   "Round this time to nearest latter"
         Height          =   195
         Left            =   480
         TabIndex        =   70
         Top             =   3760
         Width           =   2595
      End
      Begin VB.OptionButton Option17 
         BackColor       =   &H8000000B&
         Caption         =   "Round this time to nearest earlier"
         Height          =   195
         Left            =   480
         TabIndex        =   69
         Top             =   3520
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add to Z'manim list"
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
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   4995
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Configure zemanim"
         Top             =   240
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   5530
         _Version        =   393216
         Tabs            =   7
         Tab             =   6
         TabsPerRow      =   4
         TabHeight       =   600
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Dawns"
         TabPicture(0)   =   "Zmanimform.frx":16E2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame3"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Twilights"
         TabPicture(1)   =   "Zmanimform.frx":16FE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Z'mane Hayom"
         TabPicture(2)   =   "Zmanimform.frx":171A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame5"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&Candel Lighting"
         TabPicture(3)   =   "Zmanimform.frx":1736
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame6"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "S&unrises"
         TabPicture(4)   =   "Zmanimform.frx":1752
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame7"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Su&nsets"
         TabPicture(5)   =   "Zmanimform.frx":176E
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame8"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "&Mishmarot"
         TabPicture(6)   =   "Zmanimform.frx":178A
         Tab(6).ControlEnabled=   -1  'True
         Tab(6).Control(0)=   "frmMishmarim"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).ControlCount=   1
         Begin VB.Frame frmMishmarim 
            Height          =   2175
            Left            =   120
            TabIndex        =   99
            Top             =   765
            Width           =   4575
            Begin VB.TextBox txtMishmarVis 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3840
               TabIndex        =   105
               Text            =   "33.33"
               ToolTipText     =   "Enter percentage (0-100)"
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox txtMishmarAst 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3840
               TabIndex        =   104
               Text            =   "33.33"
               ToolTipText     =   "Enter percentage (0-100)"
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txtMishmarMis 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3840
               TabIndex        =   103
               Text            =   "33.33"
               ToolTipText     =   "Enter percentage (0-100)"
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton optMishmarVis 
               Caption         =   "Percentage of the night from the visible sunset to the visible sunrise"
               Height          =   495
               Left            =   240
               TabIndex        =   102
               Top             =   1440
               Width           =   3615
            End
            Begin VB.OptionButton optMishmarAst 
               Caption         =   "Percentage of night from the astronomical sunset to the astronomical sunrise"
               Height          =   375
               Left            =   240
               TabIndex        =   101
               Top             =   840
               Width           =   3615
            End
            Begin VB.OptionButton optMishmarMis 
               Caption         =   "Percentage of night from the mishor sunset to the mishor sunrise"
               Height          =   375
               Left            =   240
               TabIndex        =   100
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.Frame Frame8 
            Height          =   2235
            Left            =   -74880
            TabIndex        =   62
            Top             =   765
            Width           =   4755
            Begin VB.OptionButton Option16 
               Caption         =   $"Zmanimform.frx":17A6
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
               Top             =   1260
               Width           =   4515
            End
            Begin VB.OptionButton Option15 
               Caption         =   $"Zmanimform.frx":18C2
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Left            =   120
               TabIndex        =   67
               Top             =   540
               Width           =   4575
            End
            Begin VB.OptionButton Option14 
               Caption         =   "Visible Sunset (lattest visible sunset for the entire city as seen from somewhere in the city).  Cannot be added now."
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
               Height          =   315
               Left            =   120
               TabIndex        =   66
               Top             =   180
               Width           =   4515
            End
         End
         Begin VB.Frame Frame7 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   61
            Top             =   705
            Width           =   4755
            Begin VB.OptionButton Option13 
               Caption         =   $"Zmanimform.frx":19B6
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   60
               TabIndex        =   65
               Top             =   1320
               Width           =   4635
            End
            Begin VB.OptionButton Option12 
               Caption         =   $"Zmanimform.frx":1AD6
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   60
               TabIndex        =   64
               Top             =   540
               Width           =   4635
            End
            Begin VB.OptionButton Option11 
               Caption         =   "Visible Sunrise (earliest visible sunrise for the entire city as seen from somewhere in or near the city).  Cannot be added here."
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
               Height          =   315
               Left            =   60
               TabIndex        =   63
               Top             =   180
               Width           =   4575
            End
         End
         Begin VB.Frame Frame6 
            Height          =   2235
            Left            =   -74880
            TabIndex        =   50
            Top             =   765
            Width           =   4755
            Begin VB.TextBox Text14 
               Height          =   285
               Left            =   840
               TabIndex        =   60
               Top             =   1500
               Width           =   3735
            End
            Begin MSComCtl2.UpDown UpDown10 
               Height          =   285
               Left            =   4140
               TabIndex        =   56
               Top             =   900
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text13"
               BuddyDispid     =   196636
               OrigLeft        =   3720
               OrigTop         =   840
               OrigRight       =   3960
               OrigBottom      =   1155
               Max             =   60
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text13 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3720
               TabIndex        =   55
               Text            =   "0"
               Top             =   900
               Width           =   435
            End
            Begin MSComCtl2.UpDown UpDown9 
               Height          =   315
               Left            =   4140
               TabIndex        =   54
               Top             =   360
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text12"
               BuddyDispid     =   196637
               OrigLeft        =   3720
               OrigTop         =   420
               OrigRight       =   3960
               OrigBottom      =   735
               Max             =   60
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text12 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   3720
               TabIndex        =   53
               Text            =   "0"
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Candel Lighting time defined by clock mintues before the visible sunset:"
               Height          =   375
               Left            =   480
               TabIndex        =   52
               Top             =   780
               Width           =   3015
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Candel Lighting time defined by clock minutes before the mishor sunset:"
               Height          =   375
               Left            =   480
               TabIndex        =   51
               Top             =   360
               Width           =   3015
            End
            Begin VB.Label Label14 
               Caption         =   "Name:"
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
               Left            =   240
               TabIndex        =   59
               Top             =   1500
               Width           =   555
            End
            Begin VB.Label Label13 
               Caption         =   "minutes"
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
               Height          =   135
               Left            =   3720
               TabIndex        =   58
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label12 
               Caption         =   "minutes"
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
               Height          =   135
               Left            =   3720
               TabIndex        =   57
               Top             =   180
               Width           =   495
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2175
            Left            =   -74880
            TabIndex        =   36
            Top             =   825
            Width           =   4755
            Begin VB.CheckBox chkSunset 
               Caption         =   "Before/After twi/sunset"
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
               Left            =   2040
               TabIndex        =   98
               Top             =   1560
               Width           =   1815
            End
            Begin MSComCtl2.UpDown updwnBeforeAfter 
               Height          =   285
               Left            =   4440
               TabIndex        =   97
               Top             =   1600
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtBeforeAfter"
               BuddyDispid     =   196645
               OrigLeft        =   4440
               OrigTop         =   1680
               OrigRight       =   4695
               OrigBottom      =   1815
               Max             =   59
               Min             =   -59
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtBeforeAfter 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3960
               TabIndex        =   96
               Text            =   "0"
               Top             =   1600
               Width           =   480
            End
            Begin VB.CheckBox chkSunrise 
               Caption         =   "Before/After dawn/sunrise"
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
               Left            =   80
               TabIndex        =   95
               ToolTipText     =   "used to define, e.g., 40 minutes before KS of the Groh"
               Top             =   1560
               Width           =   3735
            End
            Begin VB.TextBox Text11 
               Height          =   285
               Left            =   600
               TabIndex        =   49
               Top             =   1850
               Width           =   3315
            End
            Begin MSComCtl2.UpDown UpDown7 
               Height          =   285
               Left            =   4440
               TabIndex        =   46
               Top             =   1320
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text9"
               BuddyDispid     =   196648
               OrigLeft        =   3780
               OrigTop         =   1380
               OrigRight       =   4020
               OrigBottom      =   1695
               Max             =   12
               Min             =   -12
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text9 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3960
               TabIndex        =   45
               Text            =   "0"
               Top             =   1320
               Width           =   495
            End
            Begin VB.OptionButton Option8 
               Caption         =   "Clock hours after the selected dawn or sunrise:"
               Height          =   195
               Left            =   60
               TabIndex        =   44
               Top             =   1320
               Width           =   3675
            End
            Begin MSComCtl2.UpDown UpDown6 
               Height          =   285
               Left            =   4440
               TabIndex        =   43
               Top             =   1020
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text8"
               BuddyDispid     =   196650
               OrigLeft        =   3720
               OrigTop         =   1260
               OrigRight       =   3960
               OrigBottom      =   1575
               Max             =   12
               Min             =   -12
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3960
               TabIndex        =   42
               Text            =   "0"
               Top             =   1020
               Width           =   495
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Hours zemanios after the selected dawn or sunrise"
               Height          =   315
               Left            =   60
               TabIndex        =   41
               Top             =   1020
               Width           =   3915
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1740
               TabIndex        =   40
               Top             =   540
               Width           =   2955
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   1740
               TabIndex        =   38
               Top             =   180
               Width           =   2955
            End
            Begin VB.Label Label11 
               Caption         =   "Name:"
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
               Left            =   80
               TabIndex        =   48
               Top             =   1875
               Width           =   555
            End
            Begin VB.Label Label10 
               Caption         =   "minutes"
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
               Height          =   135
               Left            =   3960
               TabIndex        =   47
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label9 
               Caption         =   "Pick twilight or sunset:"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label8 
               Caption         =   "Pick dawn or sunrise:"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2235
            Left            =   -74880
            TabIndex        =   21
            Top             =   765
            Width           =   4755
            Begin VB.Frame Frame10 
               Height          =   365
               Left            =   300
               TabIndex        =   83
               Top             =   1440
               Width           =   3195
               Begin VB.OptionButton Option20 
                  Caption         =   "astronomical"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   1860
                  TabIndex        =   85
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.OptionButton Option19 
                  Caption         =   "mishor"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   84
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   795
               End
               Begin VB.Label Label18 
                  Caption         =   "sunrise/sunset:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   177
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   135
                  Left            =   60
                  TabIndex        =   86
                  Top             =   130
                  Width           =   975
               End
            End
            Begin VB.TextBox Text7 
               Height          =   285
               Left            =   840
               TabIndex        =   35
               Top             =   1860
               Width           =   3675
            End
            Begin MSComCtl2.UpDown UpDown5 
               Height          =   285
               Left            =   4200
               TabIndex        =   30
               Top             =   1260
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text6"
               BuddyDispid     =   196664
               OrigLeft        =   3660
               OrigTop         =   1260
               OrigRight       =   3900
               OrigBottom      =   1515
               Max             =   200
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   3720
               TabIndex        =   29
               Text            =   "0"
               Top             =   1260
               Width           =   495
            End
            Begin MSComCtl2.UpDown UpDown4 
               Height          =   285
               Left            =   4200
               TabIndex        =   28
               Top             =   780
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text5"
               BuddyDispid     =   196665
               OrigLeft        =   3720
               OrigTop         =   840
               OrigRight       =   3960
               OrigBottom      =   1155
               Max             =   200
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               ForeColor       =   &H00004000&
               Height          =   285
               Left            =   3720
               TabIndex        =   27
               Text            =   "0"
               Top             =   780
               Width           =   495
            End
            Begin MSComCtl2.UpDown UpDown3 
               Height          =   285
               Left            =   4200
               TabIndex        =   26
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text4"
               BuddyDispid     =   196666
               OrigLeft        =   3660
               OrigTop         =   360
               OrigRight       =   3900
               OrigBottom      =   675
               Max             =   45
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   3720
               TabIndex        =   25
               Text            =   "0"
               Top             =   300
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Twilight defined by clock minutes after the sunset---------------------------------------------------->"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   300
               TabIndex        =   24
               Top             =   1080
               Width           =   3195
            End
            Begin VB.OptionButton Option4 
               Caption         =   $"Zmanimform.frx":1BCD
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   300
               TabIndex        =   23
               Top             =   540
               Width           =   3375
            End
            Begin VB.OptionButton Option7 
               Caption         =   "Twilight defined by degrees of solar depression below the horizon ------------------------------------->"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   300
               TabIndex        =   22
               Top             =   120
               Width           =   3315
            End
            Begin VB.Line Line6 
               X1              =   300
               X2              =   120
               Y1              =   1620
               Y2              =   1620
            End
            Begin VB.Line Line5 
               X1              =   120
               X2              =   120
               Y1              =   1080
               Y2              =   1620
            End
            Begin VB.Line Line4 
               X1              =   180
               X2              =   120
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Line Line3 
               X1              =   420
               X2              =   180
               Y1              =   1380
               Y2              =   1380
            End
            Begin VB.Line Line2 
               X1              =   420
               X2              =   180
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Line Line1 
               X1              =   180
               X2              =   180
               Y1              =   540
               Y2              =   1380
            End
            Begin VB.Label Label7 
               Caption         =   "Name:"
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
               Left            =   240
               TabIndex        =   34
               Top             =   1860
               Width           =   555
            End
            Begin VB.Label Label6 
               Caption         =   "minutes"
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
               Height          =   135
               Left            =   3780
               TabIndex        =   33
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label3 
               Caption         =   "minutes"
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
               Height          =   135
               Left            =   3780
               TabIndex        =   32
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "degrees"
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
               Left            =   3780
               TabIndex        =   31
               Top             =   120
               Width           =   555
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000B&
            Height          =   2355
            Left            =   -74880
            TabIndex        =   4
            Top             =   705
            Width           =   4755
            Begin VB.Frame Frame11 
               BackColor       =   &H8000000B&
               Height          =   365
               Left            =   360
               TabIndex        =   87
               Top             =   1560
               Width           =   3075
               Begin VB.OptionButton Option22 
                  Caption         =   "mishor"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   89
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   795
               End
               Begin VB.OptionButton Option21 
                  Caption         =   "astronomical"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   1800
                  TabIndex        =   88
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.Label Label19 
                  Caption         =   "sunrise/sunset:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   177
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   135
                  Left            =   60
                  TabIndex        =   90
                  Top             =   130
                  Width           =   975
               End
            End
            Begin MSComCtl2.UpDown UpDown8 
               Height          =   285
               Left            =   4080
               TabIndex        =   19
               Top             =   1260
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text10"
               BuddyDispid     =   196685
               OrigLeft        =   3660
               OrigTop         =   1320
               OrigRight       =   3900
               OrigBottom      =   1575
               Max             =   200
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text10 
               Alignment       =   2  'Center
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   3540
               TabIndex        =   18
               Text            =   "0"
               Top             =   1260
               Width           =   555
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H8000000B&
               Caption         =   "Dawn defined by clock minutes before sunrise --------------------------------------------->"
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
               Height          =   375
               Left            =   420
               TabIndex        =   17
               Top             =   1140
               Width           =   2955
            End
            Begin VB.TextBox Text3 
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   780
               TabIndex        =   13
               ToolTipText     =   "Input the name for this z'manim parameter"
               Top             =   1980
               Width           =   3675
            End
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   255
               Left            =   4080
               TabIndex        =   12
               Top             =   780
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   450
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text2"
               BuddyDispid     =   196688
               OrigLeft        =   3600
               OrigTop         =   780
               OrigRight       =   3840
               OrigBottom      =   1095
               Max             =   200
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               ForeColor       =   &H00004000&
               Height          =   285
               Left            =   3540
               TabIndex        =   10
               Text            =   "0"
               Top             =   780
               Width           =   555
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   255
               Left            =   4080
               TabIndex        =   9
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   450
               _Version        =   393216
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text1"
               BuddyDispid     =   196689
               OrigLeft        =   2700
               OrigTop         =   840
               OrigRight       =   2940
               OrigBottom      =   1155
               Max             =   45
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   3540
               TabIndex        =   8
               Text            =   "0"
               Top             =   300
               Width           =   555
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H8000000B&
               Caption         =   "Dawn of minutes zemanios before mishor sunrise ( minutes zemanios is defined by 12 hrs. between sunrise and sunset)------------>"
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
               Height          =   675
               Left            =   420
               TabIndex        =   6
               Top             =   480
               Width           =   3135
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H8000000B&
               Caption         =   "Dawn defined by degrees of solar depression below the horizon ---------------->"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   420
               TabIndex        =   5
               Top             =   120
               Width           =   2955
            End
            Begin VB.Line Line12 
               X1              =   120
               X2              =   360
               Y1              =   1740
               Y2              =   1740
            End
            Begin VB.Line Line11 
               X1              =   120
               X2              =   120
               Y1              =   960
               Y2              =   1740
            End
            Begin VB.Line Line10 
               X1              =   240
               X2              =   120
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Line Line9 
               X1              =   240
               X2              =   240
               Y1              =   540
               Y2              =   1500
            End
            Begin VB.Line Line8 
               X1              =   420
               X2              =   240
               Y1              =   1500
               Y2              =   1500
            End
            Begin VB.Line Line7 
               X1              =   420
               X2              =   240
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Label Label15 
               BackColor       =   &H8000000B&
               Caption         =   "minutes"
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
               ForeColor       =   &H00000000&
               Height          =   135
               Left            =   3600
               TabIndex        =   20
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label5 
               BackColor       =   &H8000000B&
               Caption         =   "degrees"
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
               Height          =   195
               Left            =   3600
               TabIndex        =   16
               Top             =   120
               Width           =   555
            End
            Begin VB.Label Label4 
               BackColor       =   &H8000000B&
               Caption         =   "minutes"
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
               ForeColor       =   &H00000000&
               Height          =   135
               Left            =   3600
               TabIndex        =   15
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label2 
               BackColor       =   &H8000000B&
               Caption         =   "Name:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   7
               Top             =   2040
               Width           =   555
            End
         End
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000B&
         Caption         =   "seconds"
         Height          =   195
         Left            =   4020
         TabIndex        =   76
         Top             =   3760
         Width           =   675
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000B&
         Caption         =   "seconds"
         Height          =   195
         Left            =   4020
         TabIndex        =   75
         Top             =   3520
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Z'manim name list"
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   300
         TabIndex        =   3
         Top             =   300
         Width           =   4755
      End
   End
End
Attribute VB_Name = "Zmanimform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private savedit As Boolean, addbut As Boolean, newlistnum%
Private changes As Boolean, vis1%, mishornetznum%, mishorskiynum%, visiblenetzzman%
Private astnetznum%, astskiynum%
Private mis1%, ast1%, init As Boolean, radio As Boolean, sortnumber%
Private noon As Boolean, chazosnum%, tmpsortnum%(49), load As Boolean

Private Sub Command5_Click()
   myfile = Dir(drivjk$ + "zmanim.tmp")
   If myfile = sEmpty Then
     response = MsgBox("Can't find the zmanim list!", vbCritical + vbOKOnly, "Cal Program")
     Exit Sub
     End If
  
  'Close
  numsort% = -1
  reorder = True
  neworder = True 'False
  Command6.Enabled = True
  List1.Enabled = True
  List1.Clear
End Sub

Private Sub Command6_Click()
   'If Dir(drivjk$ + "zmanim.tmp") <> sEmpty And Dir(drivjk$ + "zmansort.tmp") <> sEmpty Then
   '   FileCopy drivjk$ + zmansort.tmp, drivjk$ + zmansort.out
   '   List1.Clear
   '   List1.Enabled = False
   '   Command6.Enabled = False
   '   reorder = False
   'Else
   '   response = MsgBox("Can't find the zmanim.tmp and/or zmansort.tmp file(s)!", vbCritical + vbOKOnly, "Cal Program")
   '   End If
   Close
   zmansortnum% = FreeFile
   Open drivjk$ + "zmansort.out" For Output As #zmansortnum%
   For inum% = 0 To numsort%
      Write #zmansortnum%, tmpsortnum%(inum%)
   Next inum%
   Close #zmansortnum%
   Close #zmannum%
   neworder = True
   reorder = True
   
End Sub


Private Sub optDiasporaParshiot_Click()
    parshiotEY = False
    parshiotdiaspora = True
End Sub

Private Sub optEYParshiot_Click()
    parshiotEY = True
    parshiotdiaspora = False
End Sub

Private Sub Option19_Click()
   optiontmish% = 0
End Sub

Private Sub Option20_Click()
   optiontmish% = 1
End Sub

Private Sub Option21_Click()
   optiondmish% = 1
End Sub

Private Sub Option22_Click()
   optiondmish% = 0
End Sub

Private Sub optMishmarAst_Click()
   txtMishmarMis.Enabled = False
   txtMishmarAst.Enabled = True
   txtMishmarVis.Enabled = False
   Option17.Value = True
   Text15.Text = "60"

End Sub

Private Sub optMishmarMis_Click()
   txtMishmarMis.Enabled = True
   txtMishmarAst.Enabled = False
   txtMishmarVis.Enabled = False
   Option17.Value = True
   Text15.Text = "60"
End Sub

Private Sub optMishmarVis_Click()
   txtMishmarMis.Enabled = False
   txtMishmarAst.Enabled = False
   txtMishmarVis.Enabled = True
   Option17.Value = True
   Text15.Text = "60"

End Sub

Private Sub optNoParshiot_Click()
    parshiotEY = False
    parshiotdiaspora = False
End Sub

Private Sub savebut_Click()
  On Error GoTo c3error
10: CommonDialog1.CancelError = True
  CommonDialog1.Filter = "zmanim.zma files (*.zma)|*.zma|"
  CommonDialog1.FilterIndex = 1
  CommonDialog1.FileName = drivjk$ + "*.zma"
  CommonDialog1.ShowSave
  filnam$ = CommonDialog1.FileName
  numfil% = FreeFile
  myfile = Dir(filnam$)
  If myfile <> sEmpty Then
     response = MsgBox("Overwrite the existing file with this name?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Cal Program")
     If response <> vbYes Then
        GoTo 10
        End If
     End If
  FileCopy drivjk$ + "zmanim.tmp", filnam$
  Close
  
  'now add descriptive title
  response = InputBox("Input a short description of the z'man template", "Z'manim header", sEmpty)
  ZmanTitle$ = response
     
  'add title/description
  'add sedra/holiday flag
  filnum% = FreeFile
  Open filnam$ For Append As filnum%
  Write #filnum%, "Title/Description", "-----", "-----", "-----", "-----", "-----", "-----", "-----"
  Print #filnum%, ZmanTitle$
  trflag% = 0
  If parshiotEY Then trflag% = 1
  If parshiotdiaspora Then trflag% = 2
  Write #filnum%, "sedra/holidays", "-----", "-----", "-----", "-----", "-----", "-----", "-----"
  Write #filnum%, trflag%
  
  If Dir(drivjk$ + "zmansort.out") <> sEmpty Then
     'add the sort list to the file
     Write #filnum%, "-----", "-----", "sort", "order", "-----", "-----", "-----", "-----"
     filsort% = FreeFile
     Open drivjk$ + "zmansort.out" For Input As filsort%
     Do Until EOF(filsort%)
        Input #filsort%, ISort%
        Write #filnum%, ISort%
     Loop
     End If
  Close
  
  changes = False
c3error:
  Exit Sub

End Sub

Private Sub loadbut_Click()
  On Error GoTo c3error
  neworder = False
  reorder = False
  load = True
  TufikZman = False
  
  If internet = True Then 'load the z'manim templates if flagged
     If zmanyes% = 1 Then
       If Not optionheb Then 'english zemanim templates
         If typezman% = 0 Then 'Mogen Avrohom 72 min
            CommonDialog1.FileName = drivcities$ + "MA72min.zma"
         ElseIf typezman% = 1 Then 'Mogen Avrohom 90 min
            CommonDialog1.FileName = drivcities$ + "MA90min.zma"
         ElseIf typezman% = 2 Then 'Groh mishor
            CommonDialog1.FileName = drivcities$ + "Gaonmishor.zma"
         ElseIf typezman% = 3 Then 'Groh astronomical
            CommonDialog1.FileName = drivcities$ + "Gaonastron.zma"
         ElseIf typezman% = 4 Then 'Ben Ish Chai
            CommonDialog1.FileName = drivcities$ + "benishchai.zma"
         ElseIf typezman% = 5 Then 'Baal Hatanya
            CommonDialog1.FileName = drivcities$ + "Chabad.zma"
         ElseIf typezman% = 6 Then 'zemanim for www.yeshiva.org.il
            CommonDialog1.FileName = drivcities$ + "Harav_Melamid.zma"
         ElseIf typezman% = 7 Then 'Yedidia's Astronomical Skiya with Torah cycle/holidays
            CommonDialog1.FileName = drivcities$ + "Yedidia.zma"
            End If
       Else 'hebrew zemanim templates
         If typezman% = 0 Then 'Mogen Avrohom 72 min
            CommonDialog1.FileName = drivcities$ + "MA72min_heb.zma"
         ElseIf typezman% = 1 Then 'Mogen Avrohom 90 min
            CommonDialog1.FileName = drivcities$ + "MA90min_heb.zma"
         ElseIf typezman% = 2 Then 'Groh mishor
            CommonDialog1.FileName = drivcities$ + "Gaonmishor_heb.zma"
         ElseIf typezman% = 3 Then 'Groh astronomical
            CommonDialog1.FileName = drivcities$ + "Gaonastron_heb.zma"
         ElseIf typezman% = 4 Then 'Ben Ish Chai
            CommonDialog1.FileName = drivcities$ + "benishchai_heb.zma"
         ElseIf typezman% = 5 Then 'Baal Hatanya
            CommonDialog1.FileName = drivcities$ + "Chabad_heb.zma"
         ElseIf typezman% = 6 Then 'zemanim for www.yeshiva.org.il
            CommonDialog1.FileName = drivcities$ + "Harav_Melamid_heb.zma"
         ElseIf typezman% = 7 Then 'Yedidia's Astronomical Skiya with Torah cycle/holidays
            CommonDialog1.FileName = drivcities$ + "Yedidia_heb.zma"
            End If
          End If
       GoTo 20
       End If
    End If

10: CommonDialog1.CancelError = True
  CommonDialog1.Filter = "zmanim.zma files (*.zma)|*.zma|"
  CommonDialog1.FilterIndex = 1
  CommonDialog1.FileName = drivjk$ + "*.zma"
  CommonDialog1.ShowOpen
20  filnam$ = CommonDialog1.FileName
  Combo1.Clear
  Combo2.Clear
  Combo3.Clear
  Close
  zmannum% = FreeFile
  Open filnam$ For Input As zmannum%
  TufikZman = False
  If InStr(LCase$(filnam$), "tufik") <> 0 Then TufikZman = True
'  zmannew% = FreeFile
'  Open drivjk$ + "zmanim.tmp" For Output As zmannew%
  ZmanTitle$ = sEmpty
  numtotatl% = -1
  Do Until EOF(zmannum%)
     Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
     If a$ = "sedra/holidays" Then
        Input #zmannum%, trflag%
        If trflag% = 0 Then
           optNoParshiot.Value = True
        ElseIf trflag% = 1 Then
           optNoParshiot.Value = True 'reset
           optEYParshiot.Value = True 'set flag
        ElseIf trflag% = 2 Then
           optNoParshiot.Value = True 'reset
           optDiasporaParshiot.Value = True 'set flag
           End If
        numtotal% = numtotal% - 1
        GoTo zm100
        End If
     If a$ = "Title/Description" Then
        Line Input #zmannum%, ZmanTitle$
        GoTo zm100
        End If
     If c$ = "sort" And D$ = "order" Then
        neworder = True
        reorder = True
        Exit Do
        End If
     numtotal% = numtotal% + 1
'     Print #zmannew%, a$, b$, c$, d$, E$, F$, GZ$, HZ$
     Combo1.AddItem a$
     If InStr(a$, "Dawn") Or InStr(a$, "Sunrise") Or InStr(a$, "Chazos") Then
        Combo2.AddItem a$
     ElseIf InStr(a$, "Twilight") Or InStr(a$, "Sunset") Then
        Combo3.AddItem a$
        End If
     Combo1.ListIndex = Combo1.ListCount - 1
     If Val(GZ$) = -1 Then
        Option17.Value = True
        Text15.Text = HZ$
     ElseIf Val(GZ$) = 1 Then
        Option18.Value = True
        Text16.Text = HZ$
        End If
zm100:
  Loop
  If neworder = True Then
     'read the rest of the input file and record the sorting order
     filsort% = FreeFile
     Open drivjk$ + "zmansort.out" For Output As filsort%
     numsort% = -1
     Do Until EOF(zmannum%)
        numsort% = numsort% + 1
        Input #zmannum%, ISort%
        Write #filsort%, ISort%
     Loop
     End If
  Close #zmannum%
'  Close #zmannew%
  Close #filsort%
  
  If neworder = True And numsort% < numtotal% Then 'add the sorted order to the resort list
                                                   'if numsort%< numtotal%
     List1.Clear
     filsort% = FreeFile
     Open drivjk$ + "zmansort.out" For Input As filsort%
     zmannum% = FreeFile
     Open filnam$ For Input As zmannum%
     Do Until EOF(filsort%)
        Input #filsort%, ISort%
        Seek #zmannum%, 1
        For inum% = 0 To ISort% - 1
           Line Input #zmannum%, doclin$
        Next inum%
        Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
        List1.AddItem a$
        List1.ListIndex = List1.ListCount - 1
     Loop
     Close #filsort%
     Close #zmannum%
     List1.Enabled = True
     End If
  
  'now copy the zma file to zmanim.tmp up to the sort parameters
  If Dir(drivjk$ + "zmanim.tmp") <> sEmpty Then
     Close
     Kill drivjk$ + "zmanim.tmp"
     End If
  zmannum% = FreeFile
  Open drivjk$ + "zmanim.tmp" For Output As #zmannum%
  filsort% = FreeFile
  Open filnam$ For Input As filsort%
  Do Until EOF(filsort%)
     Input #filsort%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
     If (a$ = "sedra/holidays") Or _
        (c$ = "sort" And D$ = "order") Or _
        (a$ = "Title/Description") Then
        Exit Do
        End If
     Write #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
  Loop
  Close #filsort%
  Close #zmannum%
  
c3error:
  Exit Sub


End Sub

Private Sub Option12_Click()
   optionsun1% = 2
   radio = True
   'chkSunset.Value = 0
   'Check3.Value = 0
End Sub

Private Sub Option13_Click()
   optionsun1% = 3
   radio = True
   Option21.Enabled = True
   'chkSunset.Value = 0
   'Check2.Value = 0
End Sub

Private Sub Option14_Click()
   optionsun2% = 1
   radio = True
   'Check5.Value = 0
   'Check6.Value = 0
End Sub

Private Sub Option15_Click()
   optionsun2% = 2
   radio = True
   'Check4.Value = 0
   'Check6.Value = 0
End Sub

Private Sub Option16_Click()
   optionsun2% = 3
   radio = True
   Option20.Enabled = True
   'Check4.Value = 0
   'Check5.Value = 0
End Sub

Private Sub Combo1_Click()
  On Error GoTo errhand
  
  Dim itsthere As Boolean
  
  If reorder = True Then
    Close
    tmpitem$ = Combo1.List(Combo1.ListIndex)
    zmannum% = FreeFile
    Open drivjk$ + "zmanim.tmp" For Input As #zmannum%
    'zmansortnum% = FreeFile
    'myfile = Dir(drivjk$ + "zmansort.tmp")
    'If myfile = sEmpty Then
    '   Open drivjk$ + "zmansort.tmp" For Output As #zmansortnum%
    'Else
    '   Open drivjk$ + "zmansort.tmp" For Append As #zmansortnum%
    '   End If
  
    searchnum% = -1
    Do Until EOF(zmannum%)
       searchnum% = searchnum% + 1
       Input #zmannum%, az$, bz$, cz$, dz$, Ez$, Fz$, GZ$, HZ$
       If InStr(az$, tmpitem$) Then
          'check for redundancies
          For j% = 0 To numsort% - 1
             If tmpsortnum%(j%) = searchnum% Then
                response = MsgBox("This z'man has already been recorded!", vbExclamation + vbOKOnly, "Cal Program")
                Exit Sub
                End If
          Next j%
          numsort% = numsort% + 1
          tmpsortnum%(numsort%) = searchnum%
          'Write #zmansortnum%, searchnum%
          List1.AddItem tmpitem$
          List1.ListIndex = List1.ListCount - 1
          Exit Do
          End If
    Loop
    Close #zmannum%
    'Close #zmansortnum%
    End If
  
  If addbut = True Then Exit Sub
  itsthere = False
  i% = Combo1.ListIndex
  tmptext$ = Combo1.List(i%)
  'find this in zmanim.tmp
  myfile = Dir(drivjk$ + "zmanim.tmp")
  If myfile <> sEmpty Then
     itsthere = True
     'Close
     zmannum% = FreeFile
     Open drivjk$ + "zmanim.tmp" For Input As #zmannum%
     End If
  
  On Error GoTo errorhand
  If InStr(tmptext$, "Dawn") Then
     SSTab1.Tab = 0
     If itsthere = True Then
        Do Until EOF(zmannum%)
           Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
           If a$ = tmptext$ Then
              optiond% = Val(c$)
              Select Case optiond%
                 Case 1
                    Option1.Value = True
                 Case 2, -2
                    Option2.Value = True
                    If optiond% = -2 Then Option21.Value = True
                 Case 3, -3
                    Option3.Value = True
                    If optiond% = -3 Then Option21.Value = True
                 Case Else
              End Select
              Text1.Text = D$
              Text2.Text = e$
              Text10.Text = f$
              Text3.Text = Mid$(a$, 7, Len(a$) - 6)
              If Val(GZ$) = -1 Then
                 Option17.Value = True
                 Text15.Text = HZ$
              ElseIf Val(GZ$) = 1 Then
                 Option18.Value = True
                 Text16.Text = HZ$
                 End If
              Exit Do
              End If
        Loop
'        Close #zmannum%
        End If
  ElseIf InStr(tmptext$, "Twilight") Then
     SSTab1.Tab = 1
     If itsthere = True Then
        Do Until EOF(zmannum%)
           Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
           If a$ = tmptext$ Then
              optiont% = Val(c$)
              Select Case optiont%
                 Case 1
                    Option7.Value = True
                 Case 2, -2
                    Option4.Value = True
                    If optiont% = -2 Then Option20.Value = True
                 Case 3, -3
                    Option5.Value = True
                    If optiont% = -3 Then Option20.Value = True
                 Case Else
              End Select
              Text4.Text = D$
              Text5.Text = e$
              Text6.Text = f$
              Text7.Text = Mid$(a$, 11, Len(a$) - 10)
              If Val(GZ$) = -1 Then
                 Option17.Value = True
                 Text15.Text = HZ$
              ElseIf Val(GZ$) = 1 Then
                 Option18.Value = True
                 Text16.Text = HZ$
                 End If
              Exit Do
              End If
        Loop
'        Close #zmannum%
        End If
  ElseIf InStr(tmptext$, "Zmanim") Then
     SSTab1.Tab = 2
     If itsthere = True Then
        Do Until EOF(zmannum%)
           Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
           If a$ = tmptext$ Then
           
              'search add/subtract clock minutes to this zman
              pos% = InStr(tmptext$, "||1;")
              If pos% > 0 Then
                 BeforeAddMinStr1$ = Mid$(tmptext$, pos% + 4, Len(tmptext$) - pos% - 3)
                 chkSunrise.Value = vbChecked
                 chkSunset.Value = vbUnchecked
                 txtBeforeAfter.Text = BeforeAddMinStr1$
              Else
                 chkSunrise.Value = vbUnchecked
                 txtBeforeAfter.Text = "0"
                 End If
              pos% = InStr(tmptext$, "||2;")
              If pos% > 0 Then
                 BeforeAddMinStr2$ = Mid$(tmptext$, pos% + 4, Len(tmptext$) - pos% - 3)
                 chkSunset.Value = vbChecked
                 chkSunrise.Value = vbUnchecked
                 txtBeforeAfter.Text = BeforeAddMinStr2$
              Else
                 chkSunset.Value = vbUnchecked
                 txtBeforeAfter.Text = "0"
                 End If
           
              Combo2.ListIndex = Val(c$)
              Combo3.ListIndex = Val(D$)
              optionz% = Val(b$)
              Select Case optionz%
                 Case 1
                    Option6.Value = True
                 Case 2
                    Option8.Value = True
                 Case Else
              End Select
              Text8.Text = e$
              Text9.Text = f$
              Text11.Text = Mid$(a$, 9, Len(a$) - 8)
              If Val(GZ$) = -1 Then
                 Option17.Value = True
                 Text15.Text = HZ$
              ElseIf Val(GZ$) = 1 Then
                 Option18.Value = True
                 Text16.Text = HZ$
                 End If
              Exit Do
              End If
        Loop
'        Close #zmannum%
        End If
  ElseIf InStr(tmptext$, "Candles") Then
     SSTab1.Tab = 3
     If itsthere = True Then
        Do Until EOF(zmannum%)
           Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
           If a$ = tmptext$ Then
              options% = Val(b$)
              Select Case options%
                 Case 1
                    Option9.Value = True
                 Case 2
                    Option10.Value = True
                 Case Else
              End Select
              Text12.Text = c$
              Text13.Text = D$
              Text14.Text = Mid$(a$, 9, Len(a$) - 9)
              If Val(GZ$) = -1 Then
                 Option17.Value = True
                 Text15.Text = HZ$
              ElseIf Val(GZ$) = 1 Then
                 Option18.Value = True
                 Text16.Text = HZ$
                 End If
              Exit Do
              End If
        Loop
'        Close #zmannum%
        End If
  ElseIf InStr(tmptext$, "Sunrise") Then
     SSTab1.Tab = 4
     If itsthere = True Then
        Do Until EOF(zmannum%)
           Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
           If a$ = tmptext$ Then
              optionsun1% = Val(b$)
              Select Case optionsun1%
                 Case 1
                    Option11.Value = True
                    'chkSunset.Value = 1
                    'Check2.Value = 0
                    'Check3.Value = 0
                 Case 2
                    Option12.Value = True
                    'chkSunset.Value = 0
                    'Check2.Value = 1
                    'Check3.Value = 0
                 Case 3
                    Option13.Value = True
                    'chkSunset.Value = 0
                    'Check2.Value = 0
                    'Check3.Value = 1
                 Case Else
              End Select
              If Val(GZ$) = -1 Then
                 Option17.Value = True
                 Text15.Text = HZ$
              ElseIf Val(GZ$) = 1 Then
                 Option18.Value = True
                 Text16.Text = HZ$
                 End If
              Exit Do
              End If
        Loop
'        Close #zmannum%
        End If
  ElseIf InStr(tmptext$, "Sunset") Then
     SSTab1.Tab = 5
     If itsthere = True Then
        Do Until EOF(zmannum%)
           Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
           If a$ = tmptext$ Then
              optionsun2% = Val(b$)
              Select Case optionsun2%
                 Case 1
                    Option14.Value = True
                    'Check4.Value = 1
                    'Check5.Value = 0
                    'Check6.Value = 0
                 Case 2
                    Option15.Value = True
                    'Check4.Value = 0
                    'Check5.Value = 1
                    'Check6.Value = 0
                 Case 3
                    Option16.Value = True
                    'Check4.Value = 0
                    'Check5.Value = 0
                    'Check6.Value = 1
                 Case Else
              End Select
              If Val(GZ$) = -1 Then
                 Option17.Value = True
                 Text15.Text = HZ$
              ElseIf Val(GZ$) = 1 Then
                 Option18.Value = True
                 Text16.Text = HZ$
                 End If
              Exit Do
              End If
        Loop
'        Close #zmannum%
        End If
  End If
  Close #zmannum%
  Exit Sub
  
errhand:
'   Resume Next
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Combo1_Click of Form Zmanimform" & vbCrLf & "Do you want to resume?", vbYesNo
    response = MsgBox("Zmanimform calendarbut encountered error number: " + Str(Err.Number) + ".  Do you want to resume?", vbYesNoCancel + vbCritical, "Cal Program")
    If response <> vbYes Then
       Resume Next
    Else
       Close
       End If
       
errorhand:
    Resume Next
   
End Sub

Private Sub Command1_Click()
   On Error GoTo errhand
   Dim AlreadyOpen As Boolean
   AlreadyOpen = False
   
   If zmantotal% = 50 Then
      response = MsgBox("You cannot exceed a maximum of 50 zemanim!", vbExclamation + vbOKOnly, "Cal Program")
      Exit Sub
      End If
   reorder = False
   neworder = False
   addbut = True
   myfile = Dir(drivjk$ + "zmanim.tmp")
   'If myfile <> sEmpty And zmanopen = True Then
   '   response = MsgBox("Do you wan't to record new set of parameters? (answer NO if you want to append these zmanim to previously recorded ones)", vbQuestion + vbYesNoCancel, "Cal Program")
   '   If response = vbNo Then 'load last recorded parameters
       'zmannum% = FreeFile
       'Open drivjk$ + "zmanim.tmp" For Append As zmannum%
   '      Do Until EOF(zmannum%)
   '         Input #zmannum%, a$, b$, c$, d$, E$, F$, GZ$, HZ$
   '         Combo1.AddItem a$
   '         If InStr(a$, "Dawn") Then
   '            Combo2.AddItem a$
   '         ElseIf InStr(a$, "Twilight") Then
   '            Combo3.AddItem a$
   '            End If
   '         Combo1.ListIndex = Combo1.ListCount - 1
   '         If Val(GZ$) = -1 Then
   '            Option17.Value = True
   '            Text15.Text = HZ$
   '         ElseIf Val(GZ$) = 1 Then
   '            Option18.Value = True
   '            Text16.Text = HZ$
   '            End If
   '      Loop
   '      Close #zmannum%
   '      zmannum% = FreeFile
   '      Open drivjk$ + "zmanim.tmp" For Append As #zmannum%
   '
   '   ElseIf response = vbYes Then
   '      zmannum% = FreeFile
   '      Open drivjk$ + "zmanim.tmp" For Output As #zmannum%
   '   ElseIf response = vbCancel Then
   '      Exit Sub
   '      End If
   'Close
   If Not AlreadyOpen Then
        If myfile = sEmpty Then 'And zmanopen = True Then
           zmannum% = FreeFile
           Open drivjk$ + "zmanim.tmp" For Output As #zmannum%
        ElseIf myfile <> sEmpty Then 'And zmanopen = False Then
           zmannum% = FreeFile
           Open drivjk$ + "zmanim.tmp" For Append As #zmannum%
           End If
        AlreadyOpen = True
        End If
      
'   If init And (vis Or mis Or ast Or noon Or _
'      parshiotEY Or parshiotdiaspora) Then
   If init And (vis Or mis Or ast Or noon) Then
      If noon = True Then
         tmpnam$ = "Chazos (day and night)"
         GoSub checknam
         Combo1.AddItem tmpnam$
         Combo1.ListIndex = Combo1.ListCount - 1
         Combo2.AddItem tmpnam$
         Combo2.ListIndex = Combo2.ListCount - 1
         stepss$ = "5"
         zmanopen = False
         Write #zmannum%, tmpnam$, "NA", Str$(Combo2.ListIndex), "NA", "NA", "NA", Str$(optionround%), stepss$
         GoTo 500
         End If
'      If parshiotEY = True Then
'         tmpnam$ = "Parshiot Eretz Yisroel"
'         GoSub checknam
'         Combo1.AddItem tmpnam$
'         Combo1.ListIndex = Combo1.ListCount - 1
'         Combo2.AddItem tmpnam$
'         Combo2.ListIndex = Combo2.ListCount - 1
'         zmanopen = False
'         Write #zmannum%, tmpnam$, "NA", Str$(Combo2.ListIndex), "NA", "NA", "NA", "NA", "NA"
'         GoTo 500
'         End If
'      If parshiotdiaspora = True Then
'         tmpnam$ = "Parshiot diaspora"
'         GoSub checknam
'         Combo1.AddItem tmpnam$
'         Combo1.ListIndex = Combo1.ListCount - 1
'         Combo2.AddItem tmpnam$
'         Combo2.ListIndex = Combo2.ListCount - 1
'         zmanopen = False
'         Write #zmannum%, tmpnam$, "NA", Str$(Combo2.ListIndex), "NA", "NA", "NA", "NA", "NA"
'         GoTo 500
'         End If
      If vis1% = 1 Or mis1% = 1 Or ast1% = 1 Then
         SSTab1.Tab = 4
      ElseIf vis1% = 2 Or mis1% = 2 Or ast1% = 2 Then
         SSTab1.Tab = 5
         End If
      End If

50:
   Select Case SSTab1.Tab
      Case 0 'Dawns
         If Text3.Text <> sEmpty Then
            'check for proper character imput
            If (Val(Text1.Text) <= 0 Or Val(Text1.Text) >= 45) And (optiond% = 1) Then
               response = MsgBox("Something is wrong with your solar depression value.  Please check it.", vbExclamation + vbOKOnly, "Cal Program")
               Exit Sub
               End If
            'check for unique name
            tmpnam$ = "Dawn: " & Text3.Text
            GoSub checknam
            Combo1.AddItem tmpnam$
            Combo1.ListIndex = Combo1.ListCount - 1
            Combo2.AddItem Text3.Text
            Combo2.ListIndex = Combo2.ListCount - 1
            zmanopen = False
         Else
            response = MsgBox("Name can't be blank", vbInformation + vbOKOnly, "Cal Program")
            Close #zmannum%
            Exit Sub
            End If
         If optionround% = -1 Then
           stepss$ = Text15.Text
         ElseIf optionround% = 1 Then
           stepss$ = Text16.Text
           End If
         If optiond% >= 2 And optiondmish% = 1 Then
            Write #zmannum%, tmpnam$, Str$(Combo2.ListIndex), Str$(-optiond%), Text1.Text, Text2.Text, Text10.Text, Str$(optionround%), stepss$
         Else
            Write #zmannum%, tmpnam$, Str$(Combo2.ListIndex), Str$(optiond%), Text1.Text, Text2.Text, Text10.Text, Str$(optionround%), stepss$
            End If
      Case 1 'Twilights
         If Text7.Text <> sEmpty Then
            If (Val(Text4.Text) <= 0 Or Val(Text4.Text) >= 45) And (optiont% = 1) Then
               response = MsgBox("Something is wrong with your solar depression value.  Please check it.", vbExclamation + vbOKOnly, "Cal Program")
               Exit Sub
               End If
            'check for unique name
            tmpnam$ = "Twilight: " & Text7.Text
            GoSub checknam
            Combo1.AddItem tmpnam$
            Combo1.ListIndex = Combo1.ListCount - 1
            Combo3.AddItem Text7.Text
            Combo3.ListIndex = Combo3.ListCount - 1
            zmanopen = False
         Else
            response = MsgBox("Name can't be blank", vbInformation + vbOKOnly, "Cal Program")
            Close #zmannum%
            Exit Sub
            End If
         If optionround% = -1 Then
            stepss$ = Text15.Text
         ElseIf optionround% = 1 Then
            stepss$ = Text16.Text
            End If
         If optiont% >= 2 And optiontmish% = 1 Then
            Write #zmannum%, tmpnam$, Str$(Combo3.ListIndex), Str$(-optiont%), Text4.Text, Text5.Text, Text6.Text, Str$(optionround%), stepss$
         Else
            Write #zmannum%, tmpnam$, Str$(Combo3.ListIndex), Str$(optiont%), Text4.Text, Text5.Text, Text6.Text, Str$(optionround%), stepss$
            End If
      Case 2 'Z'manim Hayom
         
         If Text11.Text <> sEmpty Then
            'check for unique name
            tmpnam$ = "Zmanim: " & Text11.Text
            If chkSunrise.Value = vbChecked Then
              'addding or subtracting clock minutes before this zman
              tmpnam$ = tmpnam$ + "||1;" + Trim$(txtBeforeAfter.Text)
              End If
             If chkSunset.Value = vbChecked Then
              'addding or subtracting clock minutes before this zman
              tmpnam$ = tmpnam$ + "||2;" + Trim$(txtBeforeAfter.Text)
              End If
           GoSub checknam
            Combo1.AddItem tmpnam$
            Combo1.ListIndex = Combo1.ListCount - 1
            zmanopen = False
         Else
            response = MsgBox("Name can't be blank", vbInformation + vbOKOnly, "Cal Program")
            Close #zmannum%
            Exit Sub
            End If
         If optionround% = -1 Then
            stepss$ = Text15.Text
         ElseIf optionround% = 1 Then
            stepss$ = Text16.Text
            End If
         Write #zmannum%, tmpnam$, Str$(optionz%), Str$(Combo2.ListIndex), Str$(Combo3.ListIndex), Text8.Text, Text9.Text, Str$(optionround%), stepss$
         
       Case 3 'Candle Lighting
          If Text14.Text <> sEmpty Then
            'check for unique name
            tmpnam$ = "Candles: " & Text14.Text
            GoSub checknam
            Combo1.AddItem tmpnam$
            Combo1.ListIndex = Combo1.ListCount - 1
            zmanopen = False
          Else
            response = MsgBox("Name can't be blank", vbInformation + vbOKOnly, "Cal Program")
            Close #zmannum%
            Exit Sub
            End If
          If optionround% = -1 Then
            stepss$ = Text15.Text
          ElseIf optionround% = 1 Then
            stepss$ = Text16.Text
            End If
          Write #zmannum%, tmpnam$, Str$(options%), Text12.Text, Text13.Text, "NA", "NA", Str$(optionround%), stepss$
       
       Case 4 'Sunrises and noons
       
          If vis = True And radio = False Then
             If vis1% = 1 Then optionsun1% = 1
             End If
          If mis = True And radio = False Then
             If mis1% = 1 Then optionsun1% = 2
             End If
          If ast = True And radio = False Then
             If ast1% = 1 Then optionsun1% = 3
             End If
          
          Select Case optionsun1%
             Case 1
                tmpnam$ = "Sunrise: Visible Sunrise"
                GoSub checknam
                Combo1.AddItem tmpnam$
                Combo1.ListIndex = Combo1.ListCount - 1
                'chkSunset.Value = 1
                'Check2.Value = 0
                'Check3.Value = 0
                Option11.Value = True
                Combo2.AddItem tmpnam$
                Combo2.ListIndex = Combo2.ListCount - 1
                zmanopen = False
             Case 2
                tmpnam$ = "Sunrise: Mishor Sunrise"
                GoSub checknam
                Combo1.AddItem tmpnam$
                Combo1.ListIndex = Combo1.ListCount - 1
                'chkSunset.Value = 0
                'Check2.Value = 1
                'Check3.Value = 0
                Option12.Value = True
                Combo2.AddItem tmpnam$
                Combo2.ListIndex = Combo2.ListCount - 1
                zmanopen = False
             Case 3
                tmpnam$ = "Sunrise: Astronomical Sunrise"
                GoSub checknam
                Combo1.AddItem tmpnam$
                Combo1.ListIndex = Combo1.ListCount - 1
                'chkSunset.Value = 0
                'Check2.Value = 0
                'Check3.Value = 1
                'Call Check3_Click
                Option13.Value = True
                Combo2.AddItem tmpnam$
                Combo2.ListIndex = Combo2.ListCount - 1
                zmanopen = False
             Case Else
          End Select
          If optionround% = -1 Then
            stepss$ = Text15.Text
          ElseIf optionround% = 1 Then
            stepss$ = Text16.Text
            End If
          Write #zmannum%, tmpnam$, Str$(optionsun1%), Str$(Combo2.ListIndex), "NA", "NA", "NA", Str$(optionround%), stepss$
       
       Case 5 'Sunsets
       
          If vis = True And radio = False Then
             If vis1% = 2 Then optionsun2% = 1
             End If
          If mis = True And radio = False Then
             If mis1% = 2 Then optionsun2% = 2
             End If
          If ast = True And radio = False Then
             If ast1% = 2 Then optionsun2% = 3
             End If
             
          Select Case optionsun2%
             Case 1
                tmpnam$ = "Sunset: Visible Sunset"
                GoSub checknam
                Combo1.AddItem tmpnam$
                Combo1.ListIndex = Combo1.ListCount - 1
                'Check4.Value = 1
                'Check5.Value = 0
                'Check6.Value = 0
                'Call Check4_Click
                Option14.Value = True
                Combo3.AddItem tmpnam$
                Combo3.ListIndex = Combo3.ListCount - 1
                zmanopen = False
             Case 2
                tmpnam$ = "Sunset: Mishor Sunset"
                GoSub checknam
                Combo1.AddItem tmpnam$
                Combo1.ListIndex = Combo1.ListCount - 1
                'Check4.Value = 0
                'Check5.Value = 1
                'Check6.Value = 0
                'Call Check5_Click
                Option15.Value = True
                Combo3.AddItem tmpnam$
                Combo3.ListIndex = Combo3.ListCount - 1
                zmanopen = False
             Case 3
                tmpnam$ = "Sunset: Astronomical Sunset"
                GoSub checknam
                Combo1.AddItem tmpnam$
                Combo1.ListIndex = Combo1.ListCount - 1
                'Check4.Value = 0
                'Check5.Value = 0
                'Check6.Value = 1
                'Call Check6_Click
                Option16.Value = True
                Combo3.AddItem tmpnam$
                Combo3.ListIndex = Combo3.ListCount - 1
                zmanopen = False
             Case Else
          End Select
          If optionround% = -1 Then
            stepss$ = Text15.Text
          ElseIf optionround% = 1 Then
            stepss$ = Text16.Text
            End If
          Write #zmannum%, tmpnam$, Str$(optionsun2%), Str$(Combo3.ListIndex), "NA", "NA", "NA", Str$(optionround%), stepss$
       Case 6 'Mishmarot
          If optMishmarMis Then
             tmpnam$ = "Mishmarot: mishor set/rise "
             optionmishmar% = 3
             FractionNight$ = Format(Str$(Val(txtMishmarMis.Text)), "0.0####")
             tmpnam$ = tmpnam$ + "night percent: " & FractionNight$
          ElseIf optMishmarAst Then
             tmpnam$ = "Mishmarot: astro set/rise"
             optionmishmar% = 4
             FractionNight$ = Format(Str$(Val(txtMishmarAst.Text)), "0.0####")
             tmpnam$ = tmpnam$ + "night percent: " & FractionNight$
          ElseIf optMishmarVis Then
             tmpnam$ = "Mishmarot: vis set/rise"
             FractionNight$ = Format(Str$(Val(txtMishmarVis.Text)), "0.0####")
             optionmishmar% = 5
             tmpnam$ = tmpnam$ + "night percent: " & FractionNight$
          Else
             Call MsgBox("Click on one of the mishmar options!", vbInformation, "Mshmar option")
             Exit Sub
             End If
          tmpnam$ = "Zmanim: " & tmpnam$ 'these fall under the calculation cateogry of zemanim
          GoSub checknam
          Combo1.AddItem tmpnam$
          Combo1.ListIndex = Combo1.ListCount - 1
          Combo3.AddItem tmpnam$
          Combo3.ListIndex = Combo3.ListCount - 1
          zmanopen = False
          
          If optionround% = -1 Then
            stepss$ = Text15.Text
          ElseIf optionround% = 1 Then
            stepss$ = Text16.Text
            End If
          Write #zmannum%, tmpnam$, Str$(optionmishmar%), Str$(Combo3.ListIndex), FractionNight$, "NA", "NA", Str$(optionround%), stepss$
       Case Else
   End Select
500:
  Close #zmannum%
  changes = True
  addbut = False
  zmantotal% = zmantotal% + 1
  radio = False
  Exit Sub
  
checknam:
   For i% = 0 To Combo1.ListCount - 1
      If tmpnam$ = Combo1.List(i%) Then
         If laod = False Then '(if loading .zma template then don't give error message)
            response = MsgBox("The item you have chosen is already recorded!", vbExclamation + vbOKOnly, "Cal Program")
            Close #zmannum%
            End If
         Exit Sub
         End If
   Next i%
   Return
   
Exit Sub

errhand:
   If Err.Number = 55 Then
      Resume Next
      End If
      
   MsgBox "Encountered Error Number: " & Str(Err.Number) + vbLf + _
          "Error description follows." & vbLf & _
          Err.Description, vbOKOnly, "Cal Program"
   
End Sub



Private Sub calendarbut_Click()
   Dim X As Double, mp As Double, mc As Double, ap As Double
   Dim ac As Double, ms As Double, ec As Double, aas As Double
   Dim e2c As Double, lr As Double, D As Double, fdy As Double
   'dim ch As Double
   'Dim pi As Double
   Dim Z As Double, yro As Integer, dyo As Integer, hgto As Single, t3sub99 As Integer
   Dim jdn As Double, td As Double, mday As Integer, mon As Integer, yl As Integer
   Dim ob As Double, rs As Double, dy As Integer, yr As Integer, RA As Double
   Dim vbweps(6) As Double, vdwref(6) As Double, ier As Integer
   Dim VDWSF As Double, VDWALT As Double, VDWEPSR As Double, VDWREFR As Double, lnhgt As Double, weather%
   Dim MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer, TK As Double
   Dim lt As Double, lg As Double, DiagnosticsTest As Boolean
   
   'Dim sumref(2, 2363), winref(2, 2363)
   Dim sumref(7) As Double, winref(7) As Double
   
   On Error GoTo generrhand
   
'//////////////////DST support for Israel, USA added 082921/////////////////////////////////////
    Dim stryrDST%, endyrDST%, strdaynum(1) As Integer, enddaynum(1) As Integer
    
    Dim MarchDate As Integer
    Dim OctoberDate As Integer
    Dim NovemberDate As Integer
    Dim YearLength As Integer
    Dim DSThour As Integer
    
    Dim DSTadd As Boolean
    
    Dim DSTPerpetualIsrael As Boolean
    Dim DSTPerpetualUSA As Boolean
    
    'set to true when these countries adopt universal DST
    DSTPerpetualIsrael = False
    DSTPerpetualUSA = False
    
    If CalMDIform.mnuDST.Checked = True Then
       DSTadd = True
       End If
    
    If Option2b Then
       If yrheb% < 1918 Then DSTadd = False
    Else
       If yrheb% < 5678 Then DSTadd = False
       End If
    
    If DSTadd Then
    
       If Not Option2b Then 'hebrew years
          stryrDST% = yrheb% + RefCivilYear% - RefHebYear% '(yrheb% - 5758) + 1997
          endyrDST% = yrheb% + RefCivilYear% - RefHebYear% + 1 '(yrheb% - 5758) + 1998
       Else
          stryrDST% = yrheb%
          endyrDST% = yrheb%
          End If
    
       'find beginning and ending day numbers for each civil year
       Select Case eroscountry$
    
          Case "Israel", "" 'EY eros or cities using 2017 DST rules
    
              MarchDate = (31 - (Fix(stryrDST% * 5 / 4) + 4) Mod 7) - 2 'starts on Friday = 2 days before EU start on Sunday
              OctoberDate = (31 - (Fix(stryrDST% * 5 / 4) + 1) Mod 7)
              YearLength% = DaysinYear(stryrDST%)
              strdaynum(0) = DayNumber(YearLength%, 3, MarchDate)
              enddaynum(0) = DayNumber(YearLength%, 10, OctoberDate)
    
              If DSTPerpetualIsrael Then
                 strdaynum(0) = 1
                 enddaynum(0) = YearLength%
                 End If
    
              MarchDate = (31 - (Fix(endyrDST% * 5 / 4) + 4) Mod 7) - 2 'starts on Friday = 2 days before EU start on Sunday
              OctoberDate = (31 - (Fix(endyrDST% * 5 / 4) + 1) Mod 7)
              YearLength% = DaysinYear(endyrDST%)
              strdaynum(1) = DayNumber(YearLength%, 3, MarchDate)
              enddaynum(1) = DayNumber(YearLength%, 10, OctoberDate)
    
              If DSTPerpetualIsrael Then
                 strdaynum(1) = 1
                 enddaynum(1) = YearLength%
                 End If
    
    
          Case "USA", "Canada" 'English {USA DST rules}
    
            'not all states in the US have DST
            If InStr(eroscity$, "Phoenix") Or InStr(eroscity$, "Honolulu") Or InStr(eroscity$, "Regina") Then
               DSTadd = False
            Else
    
              MarchDate = 14 - (Fix(1 + stryrDST% * 5 / 4) Mod 7)
              NovemberDate = 7 - (Fix(1 + stryrDST% * 5 / 4) Mod 7)
              YearLength% = DaysinYear(stryrDST%)
              strdaynum(0) = DayNumber(YearLength%, 3, MarchDate)
              enddaynum(0) = DayNumber(YearLength%, 11, NovemberDate)
    
              If DSTPerpetualUSA Then
                 strdaynum(0) = 1
                 enddaynum(0) = YearLength%
                 End If
    
              MarchDate = 14 - (Fix(1 + endyrDST% * 5 / 4) Mod 7)
              NovemberDate = 7 - (Fix(1 + endyrDST% * 5 / 4) Mod 7)
              YearLength% = DaysinYear(endyrDST%)
              strdaynum(1) = DayNumber(YearLength%, 3, MarchDate)
              enddaynum(1) = DayNumber(YearLength%, 11, NovemberDate)
    
              If DSTPerpetualUSA Then
                 strdaynum(1) = 1
                 enddaynum(1) = YearLength%
                 End If
    
              End If
    
          Case "England", "UK", "France", "Germany", "Netherlands", "Belgium", _
               "Northern_Ireland", "Yugoslavia", "Slovakia", "Romania", "Hungary", _
               "Denmark", "Ireland", "Switzerland", "Finland", "Ukraine", "Norway", _
               "France", "Czechoslovakia", "Sweden", "Italy", "Europe"
    
              MarchDate = (31 - (Fix(stryrDST% * 5 / 4) + 4) Mod 7) 'starts on Sunday, 2 days after EY
              OctoberDate = (31 - (Fix(stryrDST% * 5 / 4) + 1) Mod 7)
              YearLength% = DaysinYear(stryrDST%)
              strdaynum(0) = DayNumber(YearLength%, 3, MarchDate)
              enddaynum(0) = DayNumber(YearLength%, 10, OctoberDate)
    
              If DSTPerpetualIsrael Then
                 strdaynum(0) = 1
                 enddaynum(0) = YearLength%
                 End If
    
              MarchDate = (31 - (Fix(endyrDST% * 5 / 4) + 4) Mod 7) 'starts on Sunday = 2 days after EY
              OctoberDate = (31 - (Fix(endyrDST% * 5 / 4) + 1) Mod 7)
              YearLength% = DaysinYear(endyrDST%)
              strdaynum(1) = DayNumber(YearLength%, 3, MarchDate)
              enddaynum(1) = DayNumber(YearLength%, 10, OctoberDate)
    
              If DSTPerpetualIsrael Then
                 strdaynum(1) = 1
                 enddaynum(1) = YearLength%
                 End If
    
    
          Case Else 'not implemented yet for other countries
             DSTadd = False
    
       End Select
       End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////

  'define constants used for the analytic expression for the
  'refraction that replaces menatsum.ref and menatwin.ref data
  'the refraction terms, ref, eps have been fitted to the above
  'data acc. to the following expression: Y = exp(A + B*lnX + C*lnX*lnX + D*lnX*lnX*lnX)
  'where X = height in kilometers = hgt * 0.001.  The fit is extremely good.
  'The Chi squared is 0.9999 and the some of squared differences < 0.01
  'The expected deviation of the fit from the above data is on the order of +/- one second
  'for the entire range of heights for all the cities of the world
    sumrefo = 8.899
    sumref(0) = 2.791796282
    sumref(1) = 0.5032840405
    sumref(2) = 0.001353422287
    sumref(3) = 0.0007065245866
    sumref(4) = 1.050981251
    sumref(5) = 0.4931095603
    sumref(6) = -0.02078600882
    sumref(7) = -0.00315052518

    winrefo = 9.85
    winref(0) = 2.779751597
    winref(1) = 0.5040818795
    winref(2) = 0.001809029729
    winref(3) = 0.0007994475831
    winref(4) = 1.188723157
    winref(5) = 0.4911777019
    winref(6) = -0.0221410531
    winref(7) = -0.003454047139
    
    'now Van Der Werf constants
    'vdW dip angle vs height polynomial fit coefficients
    vbweps(1) = 2.77346593151086
    vbweps(2) = 0.497348466526589
    vbweps(3) = 2.53874620975453E-03
    vbweps(4) = 6.75587054940366E-04
    vbweps(5) = 3.94973974451576E-05
    
    'vdW atmospheric refraction vs height polynomial fit coefficients
    vdwref(1) = 1.16577538442405
    vdwref(2) = 0.468149166683532
    vdwref(3) = -0.019176833246687
    vdwref(4) = -4.8345814464145E-03
    vdwref(5) = -4.90660400743218E-04
    vdwref(6) = -1.60099622077352E-05
    
    weather% = 5 'use van der Werf ray tracing
'             = 3 'use Menat ray tracing
   
'   pi = 4 * Atn(1)
'   pi2 = 2 * pi
'   ch = 360# / (pi2 * 15)  '57.29578 / 15  'conv rad to hr
'   cd = pi / 180#  'conv deg to rad
   'hr = 60
   
   yro = 0
   dyo = 0
   hgto = 0

If Dir(drivjk$ + "zmansort.out") <> sEmpty Then
   reorder = False
   sortnum% = FreeFile
   Open drivjk$ + "zmansort.out" For Input As #sortnum%
   numsort% = -1
   Do Until EOF(sortnum%)
      Line Input #sortnum%, doclin$
      numsort% = numsort% + 1
   Loop
   Close #sortnum%
   If neworder = False And Combo1.ListCount - 1 = numsort% Then 'ask if wan't to use last sort
      response = MsgBox("Do you wan't to list and sort the times according to the last recorded sort list?", vbQuestion + vbYesNoCancel, "Cal Program")
      If response = vbYes Then reorder = True
      End If
'   sortnum% = FreeFile
'   Open drivjk$ + "zmansort.out" For Input As #sortnum%
'   sortnumber% = -1
'   'determine how many sorted files desired
'   Do Until EOF(sortnum%)
'      sortnumber% = sortnumber% + 1
'   Loop
'   Close #sortnum%
   If neworder = True Then
      reorder = True
      'neworder = False
      End If
   End If
 
Screen.MousePointer = vbHourglass
'read refraction constants
'filraf% = FreeFile
'Open drivjk$ & "menatsum.ref" For Input As #filraf%
'Input #filraf%, sumrefo
'For i% = 1 To 2363 '500
'   Input #filraf%, a, sumref(1, i%), b, c, sumref(2, i%)
'Next i%
'Close #filraf%
'filraf% = FreeFile
'Open drivjk$ & "menatwin.ref" For Input As #filraf%
'Input #filraf%, winrefo
'For i% = 1 To 2363 '500
'   Input #filraf%, a, winref(1, i%), b, c, winref(2, i%)
'Next i%
'Close #filraf%

'determine coordinates
GoSub coordnetz

''diagnostics
'DiagnosticsTest = True
'If DiagnosticsTest Then
'         fil995% = FreeFile
'         Open App.Path & "\1995-astron.txt" For Output As #fil995%
'         write1995 = True
'
'    For jday% = 1 To 366
'        hgt = avehgtnetz
'        yr = 1995 'Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
'        lg = -35.2385 'longitude at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
'        lt = 31.7499 'latitude at at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
'        td = 2  'time difference from Greenwich England
'        pi = 4# * Atn(1)  '3.1415927
'        pi2 = 2 * pi
'        ch = 57.29578 / 15#   'conv rad to hr
'        cd = pi / 180# '1.74532927777778E-02 'conv deg to rad
'        a1 = 0.833 'angle depression under horizon for sunset/sunrise
'        hgt = 754.9
'        ZA = -90
'        dy = jday%
'        GoSub cal
'        t3hr = Fix(t3)
'        t3min = Fix((t3 - t3hr) * 60)
'        t3sec = Fix((t3 - t3hr - t3min / 60) * 3600)
'        Print #fil995%, dy, t3, Trim$(Str(t3hr)) & ":" & Format(Trim$(Str(t3min)), "0#") & ":" & Format(Trim$(Str$(t3sec)), "0#")
'
'    Next jday%
'    Close #fil995%
'End If

'open temporary list file
newlistnum% = FreeFile
Open drivjk$ + "table.new" For Output As #newlistnum%
numday% = -1
For imonth% = 1 To endyr%
   If mmdate%(2, imonth%) > mmdate%(1, imonth%) Then
      kday% = 0
      For jday% = mmdate%(1, imonth%) To mmdate%(2, imonth%)
          numday% = numday% + 1
          kday% = kday% + 1
          Zmanimlistfm.List1.AddItem String$(50, "-")
          Print #newlistnum%, String$(50, "-")
          Zmanimlistfm.List1.AddItem stortim$(3, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(3, imonth% - 1, kday% - 1)
          Zmanimlistfm.List1.AddItem stortim$(2, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(2, imonth% - 1, kday% - 1)
          Zmanimlistfm.List1.AddItem stortim$(4, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(4, imonth% - 1, kday% - 1)
          If parshiotEY Then
             Zmanimlistfm.List1.AddItem stortim$(5, imonth% - 1, kday% - 1)
             Print #newlistnum%, stortim$(5, imonth% - 1, kday% - 1)
          ElseIf parshiotdiaspora Then
             Zmanimlistfm.List1.AddItem stortim$(5, imonth% - 1, kday% - 1)
             Print #newlistnum%, stortim$(5, imonth% - 1, kday% - 1)
             End If
          
          GoSub newzemanim
      Next jday%
   ElseIf mmdate%(2, imonth%) < mmdate%(1, imonth%) Then
      kday% = 0
      For jday% = mmdate%(1, imonth%) To yrend%(0)
          kday% = kday% + 1
          numday% = numday% + 1
          Zmanimlistfm.List1.AddItem String$(50, "-")
          Print #newlistnum%, String$(50, "-")
          Zmanimlistfm.List1.AddItem stortim$(3, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(3, imonth% - 1, kday% - 1)
          Zmanimlistfm.List1.AddItem stortim$(2, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(2, imonth% - 1, kday% - 1)
          Zmanimlistfm.List1.AddItem stortim$(4, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(4, imonth% - 1, kday% - 1)
          If parshiotEY Then
             Zmanimlistfm.List1.AddItem stortim$(5, imonth% - 1, kday% - 1)
             Print #newlistnum%, stortim$(5, imonth% - 1, kday% - 1)
          ElseIf parshiotdiaspora Then
             Zmanimlistfm.List1.AddItem stortim$(5, imonth% - 1, kday% - 1)
             Print #newlistnum%, stortim$(5, imonth% - 1, kday% - 1)
             End If
          
          GoSub newzemanim
      Next jday%
      yrn% = yrn% + 1
      For jday% = 1 To mmdate%(2, imonth%)
          kday% = kday% + 1
          numday% = numday% + 1
          Zmanimlistfm.List1.AddItem String$(50, "-")
          Print #newlistnum%, String$(50, "-")
          Zmanimlistfm.List1.AddItem stortim$(3, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(3, imonth% - 1, kday% - 1)
          Zmanimlistfm.List1.AddItem stortim$(2, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(2, imonth% - 1, kday% - 1)
          Zmanimlistfm.List1.AddItem stortim$(4, imonth% - 1, kday% - 1)
          Print #newlistnum%, stortim$(4, imonth% - 1, kday% - 1)
          If parshiotEY Then
             Zmanimlistfm.List1.AddItem stortim$(5, imonth% - 1, kday% - 1)
             Print #newlistnum%, stortim$(5, imonth% - 1, kday% - 1)
          ElseIf parshiotdiaspora Then
             Zmanimlistfm.List1.AddItem stortim$(5, imonth% - 1, kday% - 1)
             Print #newlistnum%, stortim$(5, imonth% - 1, kday% - 1)
             End If
          GoSub newzemanim
      Next jday%
      End If
Next imonth%
Close #newlistnum%

'cc = MaxHourZemanios

Zmanimlistfm.Visible = True
If neworder = True Then Zmanimlistfm.Toolbar1.Buttons(1).Enabled = False
'populate fontname combo box
With Zmanimlistfm.cmbFontName
   For numf% = 0 To Screen.FontCount - 1
      .AddItem Screen.Fonts(numf%)
   Next numf%
   .Text = "David"
End With
If Zmanimlistfm.MSFlexGrid1.Visible = False Then Zmanimlistfm.cmbFontName.Visible = False
Screen.MousePointer = vbDefault
Exit Sub
       


1360 ms = mp + mc * dy1
     If ms > pi2 Then ms = ms - pi2
     aas = ap + ac * dy1
     If aas > pi2 Then aas = aas - pi2
     es = ms + ec * Sin(aas) + e2c * Sin(2 * aas)
     If es > pi2 Then es = es - pi2
     D = FNarsin(Sin(ob) * Sin(es))
     Return

1500 t3hr = Fix(t3sub)
'     t3sec = Int((t3sub - t3hr - t3min / 60) * 3600 + 0.5)
     t3min = Fix((t3sub - t3hr) * 60)
     t3sec = Int((t3sub - t3hr - t3min / 60) * 3600 + 0.5)
     If t3sec = 60 Then t3min = t3min + 1: t3sec = 0
     If t3min = 60 Then t3hr = t3hr + 1: t3min = 0
     tshr$ = Trim$(Str$(t3hr))
     tsmin$ = Trim$(Str$(t3min))
     If Len(tsmin$) = 1 Then tsmin$ = "0" + tsmin$
     tssec$ = Trim$(Str$(t3sec))
     If Len(tssec$) = 1 Then tssec$ = "0" + tssec$
     t3subb$ = tshr$ + ":" + tsmin$ + ":" + tssec$
     Return



'************new round*************
round:
           sp% = 0: If Abs(setflag%) = 1 Then sp% = 1
           'If nadd1% = 1 Then sp% = 1
           If Len(RTrim$(LTrim$(t3subb$))) > 7 Then sp% = 1
           secmin = Val(Mid$(t3subb$, 6 + sp%, 2))
           minad = 0
           If secmin Mod steps = 0 Then GoTo rnd50
           If plusround% = 1 Then 'round up
              ssec = secmin / steps
              secmins = CInt(Fix((secmin / steps) * 10) / 10 + 0.499999)
              If ssec - Fix(ssec) + 0.000001 < 0.1 Then secmins = secmins + 1
              secmin = steps * secmins
              If secmin = 60 Then
                 secmin = 0
                 If Abs(setflag%) = 1 Then '***changes
                    minad = 1 - sp%
                 Else
                    minad = 1
                    End If
                 End If
           ElseIf plusround% = -1 Then 'round down
              secmin = steps * (Int(Fix((secmin / steps) * 10) / 10))
              End If
rnd50:     If secmin <> 0 Then
              If secmin < 10 Then
                 Mid$(t3subb$, 6 + sp%, 1) = "0"
                 Mid$(t3subb$, 7 + sp%, 1) = Trim$(Str$(secmin))
              Else
                 Mid$(t3subb$, 6 + sp%, 2) = Trim$(Str$(secmin))
                 End If
           Else
              Mid$(t3subb$, 6 + sp%, 2) = "00"
              End If
           minmin = Val(Mid$(t3subb$, 3 + sp%, 2)) + minad
           If minmin = 60 Then
              If Len(Trim$(t3subb$)) = 8 Then sp% = 1  '***changes
              If setflag% = 0 And sp% = 0 Then
                 If Val(Trim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1))) >= 10 Then
                    chtmp$ = Trim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1))
                    sp% = 1
                    Mid$(t3subb$, 1, 2) = chtmp$
                    Mid$(t3subb$, 3, 6) = ":00:0"
                    t3subb$ = t3subb$ + "0"
                 Else
                    Mid$(t3subb$, 1 + sp%, 1) = Trim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1))
                    End If
              Else
                 newhour% = Val(Mid$(t3subb$, 1, 1 + sp%)) + 1
                 t3subb$ = Trim$(Str$(Val(newhour%))) & Mid$(t3subb$, 2 + sp%, 6)
                 End If
              Mid$(t3subb$, 3 + sp%, 2) = "00"
           Else
              If minmin < 10 Then
                 Mid$(t3subb$, 3 + sp%, 1) = "0"
                 Mid$(t3subb$, 4 + sp%, 1) = Trim$(Str$(minmin))
              Else
                 Mid$(t3subb$, 3 + sp%, 2) = Trim$(Str$(minmin))
                 End If
              End If
           'If Mid$(t3subb$, 1, 2) = "24" Then Mid$(t3subb$, 1, 2) = "00"
Return
'**********************************

'round:     sp% = 0: IF ABS(setflag%) = 1 THEN sp% = 1
'           secmin = VAL(MID$(t3sub$, 6 + sp%, 2))
'           minad = 0
'           'IF secmin MOD steps = 0 THEN GOTO rnd50
'           IF plusround% = 1 THEN 'round up
'              ssec = secmin / steps
'              secmins = CINT(FIX((secmin / steps) * 10) / 10 + .499999)
'              IF ssec - FIX(ssec) + .000001 < .1 THEN secmins = secmins + 1
'              secmin = steps * secmins
'              IF secmin = 60 THEN
'                 secmin = 0
'                 minad = 1 '- sp%
'                 END IF
'              'IF secmin > 0 AND secmin <= 15 THEN
'              '   secmin = 15 * ABS(sp% - 1)
'              'ELSEIF secmin > 15 AND secmin <= 30 THEN
'              '   secmin = 15 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin > 30 AND secmin <= 45 THEN
'              '   secmin = 30 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin > 45 THEN
'              '   secmin = 45 * sp%
'              '   minad = 1 - sp%
'              '   END IF
'           ELSEIF plusround% = -1 THEN 'round down
'              secmin = steps * (INT(FIX((secmin / steps) * 10) / 10))
'              'IF secmin >= 0 AND secmin < 15 THEN
'              '   secmin = 15 * ABS(sp% - 1)
'              'ELSEIF secmin >= 15 AND secmin < 30 THEN
'              '   secmin = 15 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin >= 30 AND secmin < 45 THEN
'              '   secmin = 30 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin >= 45 THEN
'              '   secmin = 45 * sp%
'              '   minad = 1 - sp%
'              '   END IF
'              END IF
'rnd50:     IF secmin <> 0 THEN
'              IF secmin < 10 THEN
'                 MID$(t3sub$, 6 + sp%, 1) = "0"
'                 MID$(t3sub$, 7 + sp%, 1) = LTRIM$(RTRIM$(STR$(secmin)))
'              ELSE
'                 MID$(t3sub$, 6 + sp%, 2) = LTRIM$(RTRIM$(STR$(secmin)))
'                 END IF
'           ELSE
'              MID$(t3sub$, 6 + sp%, 2) = "00"
'              END IF
'           minmin = VAL(MID$(t3sub$, 3 + sp%, 2)) + minad
'           IF minmin = 60 THEN
'              MID$(t3sub$, 1 + sp%, 1) = LTRIM$(RTRIM$(STR$(VAL(MID$(t3sub$, 1 + sp%, 1)) + 1)))
'              MID$(t3sub$, 3 + sp%, 2) = "00"
'           ELSE
'              IF minmin < 10 THEN
'                 MID$(t3sub$, 3 + sp%, 1) = "0"
'                 MID$(t3sub$, 4 + sp%, 1) = LTRIM$(RTRIM$(STR$(minmin)))
'              ELSE
'                 MID$(t3sub$, 3 + sp%, 2) = LTRIM$(RTRIM$(STR$(minmin)))
'                 END IF
'              END IF
'          IF setflag% = -1 THEN MID$(t3sub$, 1, 2) = " " + LTRIM$(RTRIM$(STR$(VAL(MID$(t3sub$, 1, 2)) - 12)))
'
'RETURN



lglt:
        G1# = kmyo * 1000#
        G2# = kmxo * 1000#
        r# = 57.2957795131
        B2# = 0.03246816
        f1# = 206264.806247096
        s1# = 126763.49
        S2# = 114242.75
        e4# = 0.006803480836
        C1# = 0.0325600414007
        C2# = 2.55240717534E-09
        c3# = 0.032338519783
        X1# = 1170251.56
        yy1# = 1126867.91
        yy2# = G1#
'       GN & GE
        X2# = G2#
        If (X2# > 700000!) Then GoTo ca5
        X1# = X1# - 1000000#
ca5:    If (yy2# > 550000#) Then GoTo ca10
        yy1# = yy1# - 1000000#
ca10:   X1# = X2# - X1#
        yy1# = yy2# - yy1#
        D1# = yy1# * B2# / 2#
        O1# = S2# + D1#
        O2# = O1# + D1#
        A3# = O1# / f1#
        A4# = O2# / f1#
        B3# = 1# - e4# * Sin(A3#) ^ 2
        B4# = B3# * Sqr(B3#) * C1#
        C4# = 1# - e4# * Sin(A4#) ^ 2
        C5# = Tan(A4#) * C2# * C4# ^ 2
        C6# = C5# * X1# ^ 2
        D2# = yy1# * B4# - C6#
        C6# = C6# / 3#
'LAT
        l1# = (S2# + D2#) / f1#
        R3# = O2# - C6#
        R4# = R3# - C6#
        R2# = R4# / f1#
        A2# = 1# - e4# * Sin(l1#) ^ 2
        lt = r# * (l1#)
        A5# = Sqr(A2#) * c3#
        d3# = X1# * A5# / Cos(R2#)
' LON
        lg = r# * ((s1# + d3#) / f1#)
'       THIS IS THE EASTERN HEMISPHERE!
        lg = -lg
Return

coordnetz:
    If zmansetflag% > -4 And zmannetz = True Then
       'convert ITM to geo
       kmxo = avekmxnetz
       kmyo = avekmynetz
       hgt = avehgtnetz
       GoSub lglt
       td = 2
    ElseIf zmansetflag% > -4 And zmannetz = False Then
       kmxo = avekmxskiy
       kmyo = avekmyskiy
       hgt = avehgtskiy
       GoSub lglt
       td = 2
    ElseIf zmansetflag% <= -4 And zmannetz = True Then
       If astronfm = False Then
          lg = avekmxnetz
          lt = avekmynetz
          hgt = avehgtnetz
          td = geotz!
       Else
          lg = avekmynetz
          lt = avekmxnetz
          hgt = avehgtnetz
          td = geotz!
          End If
    ElseIf zmansetflag% <= -4 And zmannetz = False Then
       If astronfm = False Then
          lg = avekmxskiy
          lt = avekmyskiy
          hgt = avehgtskiy
          td = geotz!
       Else
          lg = avekmynetz
          lt = avekmxnetz
          hgt = avehgtnetz
          td = geotz!
          End If
       End If
       
    If weather% = 5 Then
        'load minimum and average temperatures for this place
        Call Temperatures(lt, -lg, MinT, AvgT, MaxT, ier)
        If ier = -1 Then
           MsgBox "Can't find termperature files!", vbCritical + vbOKOnly, "Missing Temperature Fiels"
           Exit Sub
           End If
        End If
   
   'If citnam$ = "ast" And Len(LTrim$(RTrim$(Str(Int(Abs(lt)))))) < 3 Then
   '   lt0 = lt
   '   lt = lg
   '   lg = lt0
   '   End If
       
Return
   
'not used anymore (now just using coordnetz)
'just left for reference
coordskiy:
    If zmansetflag% > -4 And zmanskiy = True Then
       'convert ITM to geo
       kmxo = avekmxskiy
       kmyo = avekmyskiy
       hgt = avehgtskiy
       GoSub lglt
       td = 2
    ElseIf zmansetflag% > -4 And zmanskiy = False Then
       kmxo = avekmxnetz
       kmyo = avekmynetz
       hgt = avehgtnetz
       GoSub lglt
       td = 2
    ElseIf zmansetflag% <= -4 And zmanskiy = True Then
       If astronfm = False Then
          lg = avekmxskiy
          lt = avekmyskiy
          hgt = avehgtskiy
          td = geotz!
       Else
          lg = avekmyskiy
          lt = avekmxskiy
          hgt = avehgtskiy
          td = geotz!
          End If
    ElseIf zmansetflag% <= -4 And zmanskiy = False Then
       If astronfm = False Then
          lg = avekmxnetz
          lt = avekmynetz
          hgt = avehgtnetz
          td = geotz!
       Else
          lg = avekmynetz
          lt = avekmxnetz
          hgt = avehgtnetz
          td = geotz!
          End If
       End If
       
    If weather% = 5 Then
        'load minimum and average temperatures for this place
        Call Temperatures(lt, -lg, MinT, AvgT, MaxT, ier)
        If ier = -1 Then
           MsgBox "Can't find termperature files!", vbCritical + vbOKOnly, "Missing Temperature Fiels"
           Exit Sub
           End If
        End If
       
   'If citnam$ = "ast" And Len(LTrim$(RTrim$(Str(Int(Abs(lt)))))) < 3 Then
   '   lt0 = lt
   '   lt = lg
   '   lg = lt0
   '   End If
      
Return

cal:
        If (yr <> yro) Then 'new civil year, reinitialze astronomical constants
            yro = yr
            
            
           'there is a bug in the program:  zemanim are calculated according to all the vantage points even if they were not selected
           'NEED TO FIX!!!  -- otherwise use the Astronomical calculator button for one place, or do the following as an example:
'lg = -35.2386092463325 '-35.239419 '-35.2385 'longitude at Armon HaNaziv
'lt = 31.752241715397 '31.751666 '31.7499 'latitude at Armon HaNaziv
'hgt = 796.8 '791 '754.9  'altitude of observer at Armon HaNaziv N.
           
            lr = cd * lt
            tf = td * 60 + 4 * lg
            
            If (CalcMethod% = 0) Then
                ac = cd * 0.9856003
                ap = cd * 357.528   'mean anomaly for Jan 0, 1996 12:00
                mc = cd * 0.9856474
                mp = cd * 280.461  'mean longitude of sun Jan 0, 1996 12:00
                ec = cd * 1.915
                e2c = cd * 0.02
                ob = cd * 23.439   'ecliptic angle for Jan 0, 1996 12:00
                ob1 = cd * (-0.0000004) 'change of ecliptic angle per day
                ap = ap - (td / 24) * ac 'compensate for our time zone
                mp = mp - (td / 24) * mc
                ob = ob - (td / 24) * ob1
                'calculate cumulative years since 1996
                yd = yr - 1996
                yf = 0
                If yd < 0 Then
                   For iyr% = 1995 To yr Step -1
                      yrtst% = iyr%
                      yltst% = 365
                      If yrtst% Mod 4 = 0 Then yltst% = 366
                      If yrtst% Mod 4 = 0 And yrtst% Mod 100 = 0 And yrtst% Mod 400 <> 0 Then yltst% = 365
                      yf = yf - yltst%
                   Next iyr%
                ElseIf yd >= 0 Then
                   For iyr% = 1996 To yr - 1 Step 1
                      yrtst% = iyr%
                      yltst% = 365
                      If yrtst% Mod 4 = 0 Then yltst% = 366
                      If yrtst% Mod 4 = 0 And yrtst% Mod 100 = 0 And yrtst% Mod 400 <> 0 Then yltst% = 365
                      yf = yf + yltst%
                   Next iyr%
                   End If
                yl = 365
                If yr Mod 4 = 0 Then yl = 366
                If yr Mod 4 = 0 And yr Mod 100 = 0 And yr Mod 400 <> 0 Then yl = 365
                yf = yf - 1462 'number of days from J2000.0 (called "n" in Almanac)
                ob = ob + ob1 * yf
                ob = ob + ob1 * (dy - 1) 'acculative daily change until yesterday's noon time
                mp = mp + yf * mc
600             If mp < 0 Then mp = mp + pi2
610             If mp < 0 Then GoTo 600
620             If mp > pi2 Then mp = mp - pi2
630             If mp > pi2 Then GoTo 620
640             ap = ap + yf * ac
650             If ap < 0 Then ap = ap + pi2
660             If ap < 0 Then GoTo 650
670             If ap > pi2 Then ap = ap - pi2
680             If ap > pi2 Then GoTo 670
            
                End If
            
            End If

        If dyo <> dy Or hgt <> hgto Then 'the following needs to be calculated only once per day
            dyo = dy
            hgto = hgt
            t3sub99 = 0
            
            If weather% = 3 Then
                '***********determine fractions of summer and winter atmoshperes***********
                '**************************************************************************
                'these parameters will change as global warming takes its toll
                If (Abs(lt) < 20#) Then
    '                  tropical region, so use tropical atmosphere which
    '                  is approximated well by the summer mid-latitude
    '                  atmosphere for the entire year
                    nweather = 0
    '                  define days of the year when winter and summer begin
                 ElseIf (Abs(lt) >= 20# And Abs(lt) < 30#) Then
                    ns1 = 55
                    ns2 = 320
                    'no adhoc fixes
                    ns3 = 0
                    ns4 = 366
                 ElseIf (Abs(lt) >= 30# And Abs(lt) < 40#) Then
    '                  weather similar to Eretz Israel */
                    ns1 = 85
                    ns2 = 290
    '               Times for ad-hoc fixes to the visible and astr. sunrise */
    '               (to fit observations of the winter netz in Neve Yaakov). */
    '               This should affect  sunrise and sunset equally. */
    '               However, sunset hasn't been observed, and since it would */
    '               make the sunset times later, it's best not to add it to */
    '               the sunset times as a chumrah. */
                    ns3 = 30
                    ns4 = 330
                 ElseIf (Abs(lt) >= 40# And Abs(lt) < 50#) Then
    '               the winter lasts two months longer than in Eretz Israel */
                    ns1 = 115
                    ns2 = 260
                    ns3 = 60
                    ns4 = 300
                ElseIf (Abs(lt) >= 50#) Then
    '               the winter lasts four months longer than in Eretz Israel */
                    ns1 = 145
                    ns2 = 230
                    ns3 = 90
                    ns4 = 270
                    End If
               
                If hgt > 0 Then lnhgt = Log(hgt * 0.001)
                If dy <= ns1 Or dy >= ns2 Then 'winter refraction
                   ref = 0: eps = 0
                   If hgt <= 0 Then GoTo 690
                   ref = Exp(winref(4) + winref(5) * lnhgt + _
                       winref(6) * lnhgt * lnhgt + winref(7) * lnhgt * lnhgt * lnhgt)
        '           ref = ((winref(2, n2%) - winref(2, n1%)) / 2) * (hgt - h1) + winref(2, n1%)
                   eps = Exp(winref(0) + winref(1) * lnhgt + _
                        winref(2) * lnhgt * lnhgt + winref(3) * lnhgt * lnhgt * lnhgt)
        '           eps = ((winref(1, n2%) - winref(1, n1%)) / 2) * (hgt - h1) + winref(1, n1%)
690                air = 90 * cd + (eps + ref + winrefo) / 1000
                ElseIf dy > ns1 And dy < ns2 Then
                   ref = 0: eps = 0
                   If hgt <= 0 Then GoTo 695
                   ref = Exp(sumref(4) + sumref(5) * lnhgt + _
                       sumref(6) * lnhgt * lnhgt + sumref(7) * lnhgt * lnhgt * lnhgt)
                   'ref = ((sumref(2, n2%) - sumref(2, n1%)) / 2) * (hgt - h1) + sumref(2, n1%)
                   eps = Exp(sumref(0) + sumref(1) * lnhgt + _
                        sumref(2) * lnhgt * lnhgt + sumref(3) * lnhgt * lnhgt * lnhgt)
        '           eps = ((sumref(1, n2%) - sumref(1, n1%)) / 2) * (hgt - h1) + sumref(1, n1%)
695                air = 90 * cd + (eps + ref + sumrefo) / 1000
                   End If
            ElseIf weather% = 5 Then
                'astronomical refraction calculations, hgt in meters
            
               'determine the minimum and average temperature for this day for current place
               'use Meeus's forumula p. 66 to convert daynumber to month,
               'no need to interpolate between temepratures -- that is overkill
               k% = 2
               If (yl = 366) Then k% = 1
               m% = Int(9 * (k% + dy) / 275 + 0.98)
               
               'refraction is determined by what part of the day
               'refraction is minimum near sunrise and maximum approximately at
               'TK = MT(m%) + 273.15 'mean minimum temperature for this month in degrees Kelvin
               TK = AvgT(m%) + 273.15 'use average temperatures for zemanim
               'TK = MaxT(m%) + 273.15 'use maximum temperature for zemanim, etc.
               'calculate van der Werf temperature scaling factor for refraction
               VDWSF = (288.15 / TK) ^ 1.687 ' 1.7081
               'calculate van der Werf scaling factor for view angles
               VDWALT = (288.15 / TK) ^ 0.69
               'calculate VDW scaling factor for eps
               VDWEPSR = (288.15 / TK) ^ -0.2
               'calculate VDW scaling for ref
               VDWREFR = (288.15 / TK) ^ 2.18
               
'static doublereal c_b175 = 1.687; //1.7081; //1.686701538132792; //1.7081;
'static doublereal c_b176 = .69;
'static doublereal c_b177 = 73.;
'static doublereal c_b178 = 9.56267125268496f; //9.572702286470884f; //9.56267125268496f;
'static doublereal c_b179 = 2.18; //exponent for ref  (need expln why it turned out different than c_b075 ????)
'                                 //but makes dif for Jerusalem astronomical is no more than 2 seconds.
'static doublereal c_b180 = 0.5; //reduced vdw exponent for astronomical altitudes
'static doublereal c_b181 = -0.2; //exponent for eps
'
'        d__2 = 288.15f / tk;
'        vdwsf = pow_dd(&d__2, &c_b175);
'/*          calculate van der Werf scaling factor for view angles */
'        d__2 = 288.15f / tk;
'        vdwalt = pow_dd(&d__2, &c_b176);
'        //scaling law for ref
'        d__2 = 288.15f / tk;
'        vdwrefr = pow_dd(&d__2, &c_b179);
'        //scaling for astronomical times
'        d__2 = 288.15f / tk;
'        vbwast = pow_dd(&d__2, &c_b180);
'        //scaling for local ray altitude (compliment of zenith angle)
'        d__2 = 288.15f / tk;
'        vbweps = pow_dd(&d__2, &c_b181);
'
'        air = cd * 90. + (vbweps * eps + vdwrefr * ref + vdwsf * refrac1) / 1e3;

                       
                ref = 0#
                eps = 0#
                If (hgt > 0) Then lnhgt = Log(hgt * 0.001)
                'calculate total atmospheric refraction from the observer's height
                'to the horizon and then to the end of the atmosphere
                'All refraction terms have units of mrad
                If (hgt <= 0#) Then GoTo 790
                ref = Exp(vdwref(1) + vdwref(2) * lnhgt + _
                     vdwref(3) * lnhgt * lnhgt + vdwref(4) * lnhgt * lnhgt * lnhgt + _
                     vdwref(5) * lnhgt * lnhgt * lnhgt * lnhgt + _
                     vdwref(6) * lnhgt * lnhgt * lnhgt * lnhgt * lnhgt)
                eps = Exp(vbweps(1) + vbweps(2) * lnhgt + _
                     vbweps(3) * lnhgt * lnhgt + vbweps(4) * lnhgt * lnhgt * lnhgt + _
                     vbweps(5) * lnhgt * lnhgt * lnhgt * lnhgt)
                 'now add the all the contributions together due to the observer's height
                 'along with the value of atm. ref. from hgt<=0, where view angle, a1 = 0
                 'for this calculation, leave the refraction in units of mrad
790:             a1 = 0
                 air = 90 * cd + (VDWEPSR * eps + VDWREFR * ref + VDWSF * 9.56267125268496) / 1000#
               End If
               
               
            dy1 = dy
            ob = ob + ob1 'change until this day's noon time
            
            If CalcMethod% = 1 Then
                'this is first call on this day to Decl3, so mode = 0
                D = Decl3(jdn, 12#, CInt(dy1), td, mday, mon, yr, yl, ms, aas, ob, rs, RA, 0)
            Else
                'change from noon of previous day to noon of today
                '-->not correct-->'ob = ob + ob1 'change from noon of previous day to noon of today
                'now incorporated inside gosub 1360
                GoSub 1360
                End If
                
            '**************check for polar circles**************
            If (pi - Sgn(lr) * (D + lr) <= air + 0.002) Or (Abs(lr - D) >= pi / 2 - 0.02) Then
               'sun doesn't set that day
               t3sub = -9999
               t3sub99 = -1
               Return
               End If
            '***************************************************
            If CalcMethod% = 0 Then
               RA = Atn(Cos(ob) * Tan(es))
               End If
               
            If RA < 0 Then RA = RA + pi
            df = ms - RA
            While Abs(df) > pi / 2
              df = df - Sgn(df) * pi
            Wend
            et = df / cd * 4
            t6 = 720 + tf - et
            
'            If CalcMethod% = 0 Then
'                dy1 = dy
'                GoSub 1360 'repeat improve accuracy of declination
'                End If
                
            If Abs((-Tan(D) * Tan(lr))) > 1# Then
               'sun doesn't set that day
               t3sub = -9999
               t3sub99 = -1
               Return
               End If
                
            fdy = FNarco(-Tan(D) * Tan(lr)) * ch / 24
            
            If CalcMethod% = 0 Then
                dy1 = dy - fdy
                GoSub 1360
            Else
                'second call to Decl3
                D = Decl3(jdn, 12# - fdy * 24#, dy, td, mday, mon, yr, yl, ms, aas, ob, rs, RA, 1)
                End If
                
            air = air + 0.2667 * (1 - 0.017 * Cos(aas)) * cd 'change of size of sun due to elliptical orbit of earth
            If Abs(((-Tan(lr) * Tan(D)) + (Cos(air) / Cos(lr) / Cos(D)))) > 1# Then
               'sun doesn't set that day
               t3sub = -9999
               t3sub99 = -1
               Return
              End If
              
            sr1 = FNarco((-Tan(lr) * Tan(D)) + (Cos(air) / Cos(lr) / Cos(D))) * ch
            sr = sr1 * hr
            t3 = t6 - sr
            t3 = t3 / hr: t6 = t6 / hr

'            air2 = (90 + a1 - 0.533 + ht) * cd 'total angle depression when sun fully above horizon 'this commented out section is not used since it is redundant>
'            If Abs(((-Tan(lr) * Tan(d)) + (Cos(air2) / Cos(lr) / Cos(d)))) > 1# Then
'               'sun doesn't set that day
'               t3sub = -9999
'               t3sub99 = -1
'               Return
'              End If
'            sr2 = FNarco((-Tan(lr) * Tan(d)) + (Cos(air2) / Cos(lr) / Cos(d))) * ch
         Else
            If t3sub99 = -1 Then
               t3sub = -9999
               Return 'no zemanim for this day since sun doesn't rise/set
               End If
            End If
            
        If ZA < -90 Then 'dawns
           If CalcMethod% = 0 Then
              dy1 = dy - fdy
              GoSub 1360 'approximate declination
           Else
              'already calculated declination for this hourangle
              End If
              
           ZA1 = Abs(ZA)
           air2 = ZA1 * cd 'zenith angle of dawn
           '********************check for no dawn*************
           If air2 > pi - Sgn(lr) * (D + lr) Or air2 < Abs(lr - D) Then
              'no dawn
              t3sub = -9999
              Return
              End If
           '**************************************************
           If Abs(((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D)))) > 1# Then
              'no dawn
              t3sub = -9999
              Return
             End If
           sr2 = FNarco((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D))) * ch
           
           If CalcMethod% = 0 Then
              dy1 = dy - sr2 / 24
              GoSub 1360  'calculate true declination
           Else
              D = Decl3(jdn, 12# - sr2, dy, td, mday, mon, yr, yl, ms, aas, ob, rs, RA, 1)
              End If
              
           If Abs(((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D)))) > 1# Then
              'no dawn
              t3sub = -9999
              Return
             End If
           sr2 = FNarco((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D))) * ch
           t3sub = t6 - sr2 'in fractions of hours
        ElseIf ZA = -90 Then
           'astronomical or mishor sunrise
           t3sub = t3
        ElseIf ZA = 0 Then 'true noon
           t3sub = t6
        ElseIf ZA = 90 Then
           'astronomical or mishor sunset
           If CalcMethod% = 0 Then
              dy1 = dy + fdy
              GoSub 1360 'approximate declination
           Else
              D = Decl3(jdn, 12# + fdy * 24#, dy, td, mday, mon, yr, yl, ms, aas, ob, rs, RA, 1)
              End If
           air2 = air 'zenith angle
           '********************check for no sunset (polar circle)*************
           If air2 > pi - Sgn(lr) * (D + lr) Or air2 < Abs(lr - D) Then
              'no dawn
              t3sub = -9999
              Return
              End If
           '**************************************************
           If Abs(((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D)))) > 1# Then
              'no dawn
              t3sub = -9999
              Return
             End If
           sr2 = FNarco((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D))) * ch
           
           If CalcMethod% = 0 Then
              dy1 = dy + sr2 / 24
              GoSub 1360  'calculate true declination
           Else
              D = Decl3(jdn, 12# + sr2, dy, td, mday, mon, yr, yl, ms, aas, ob, rs, RA, 1)
              End If
              
           If Abs(((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D)))) > 1# Then
              'no dawn
              t3sub = -9999
              Return
             End If
           sr2 = FNarco((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D))) * ch
           t3sub = t6 + sr2 'in fractions of hours
        ElseIf ZA > 90 Then 'twilights
        
           If CalcMethod% = 0 Then
              dy1 = dy + fdy
              GoSub 1360 'approximate declination
           Else
              D = Decl3(jdn, 12# + fdy * 24#, dy, td, mday, mon, yr, yl, ms, aas, ob, rs, RA, 1)
              End If
              
           air2 = ZA * cd 'zenith angle
           '********************check for no twilight*************
           If air2 > pi - Sgn(lr) * (D + lr) Or air2 < Abs(lr - D) Then
              'no dawn
              t3sub = -9999
              Return
              End If
           '**************************************************
           If Abs(((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D)))) > 1# Then
              'no dawn
              t3sub = -9999
              Return
              End If
           sr2 = FNarco((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D))) * ch
           
           If CalcMethod% = 0 Then
              dy1 = dy + sr2 / 24
              GoSub 1360 'calculate true declination
           Else
              D = Decl3(jdn, 12# + sr2, dy, td, mday, mon, yr, yl, ms, aas, ob, rs, RA, 1)
              End If
              
           If Abs(((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D)))) > 1# Then
              'no dawn
              t3sub = -9999
              Return
             End If
           sr2 = FNarco((-Tan(lr) * Tan(D)) + (Cos(air2) / Cos(lr) / Cos(D))) * ch
           t3sub = t6 + sr2 'in fractions of hours
           End If
           
          '//////////////added 082921--DST support//////////////
          If DSTadd And t3sub <> -9999 Then
             'add hour for DST
             j% = dy
             yrn% = yr
             If j% >= strdaynum(yrn% - stryrDST%) And j% < enddaynum(yrn% - stryrDST%) Then
                t3sub = t3sub + 1
                If t3sub > 24 Then t3sub = t3sub - 24
                End If
             End If
          '//////////////////////////////////////////////////


Return


newzemanim:

   myfile = Dir(drivjk$ + "zmanim.tmp")
   If myfile <> sEmpty Then
      zmannum% = FreeFile
      Open drivjk$ + "zmanim.tmp" For Input As #zmannum%
   ElseIf myfile = sEmpty Then
      response = MsgBox("Can't find the zmanim.tmp file!", vbCritical + vbOKOnly, "Cal Program")
      Exit Sub
      End If
      
   num% = -1
   Do Until EOF(zmannum%)
      Input #zmannum%, az$, bz$, cz$, dz$, Ez$, Fz$, GZ$, HZ$
      'calculate the zeman and record it according to its Combo2,Combo3 number
      If Mid$(az$, 1, 4) = "Dawn" Then
      
        num% = num% + 1
      
        Select Case Val(cz$)
           Case 1 'dawn defined by solar depression
             'GoSub coordnetz
             yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
             ZA = -90 - Val(dz$)
             dy = jday%
             GoSub cal
             zmannames(num%) = az$
             zmantimes(num%, numday%) = Str(t3sub)
             If t3sub = -9999 Then
                zmantimes(num%, numday%) = "none"
                End If
             zmannumber%(0, Val(bz$)) = num%
           Case 2 'dawn defined by minutes zemanios (as defined by mishor sunrise/sunset)
             t3subb$ = zmantimes(mishornetznum%, numday%) 'this is mishor sunrise
             If (t3subb$ <> "none") Then
                sunrise = Val(t3subb$)
                t3subb$ = zmantimes(mishorskiynum%, numday%) 'this is mishor sunset
                If (t3subb$ <> "none") Then
                   sunsets = Val(t3subb$)
                   hourszemanios = (sunsets - sunrise) / 12
                   'subtract minutes zemnanios
                   If Str(sunrise - (Val(Ez$) / 60) * hourszemanios) < 0 Then
                      t3subb$ = "none"
                   Else
                      t3subb$ = Str(sunrise - (Val(Ez$) / 60) * hourszemanios)
                      End If
                Else
                   t3subb$ = "none"
                   End If
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(0, Val(bz$)) = num%
           Case -2 'dawn defined by minutes zemanios (as defined by astron. sunrise/sunset)
             t3subb$ = zmantimes(astnetznum%, numday%) 'this is astronomical sunrise
             If (t3subb$ <> "none") Then
                sunrise = Val(t3subb$)
                t3subb$ = zmantimes(astskiynum%, numday%) 'this is mishor sunset
                If (t3subb$ <> "none") Then
                   sunsets = Val(t3subb$)
                   hourszemanios = (sunsets - sunrise) / 12
                   'subtract minutes zemnanios
                   If Str(sunrise - (Val(Ez$) / 60) * hourszemanios) < 0 Then
                      t3subb$ = "none"
                   Else
                      t3subb$ = Str(sunrise - (Val(Ez$) / 60) * hourszemanios)
                      End If
                Else
                   t3subb$ = "none"
                   End If
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(0, Val(bz$)) = num%
          Case 3
             t3subb$ = zmantimes(mishornetznum%, numday%) 'this is mishor sunrise
             If (t3subb$ <> "none") Then
                sunrise = Val(t3subb$)
                If Str(sunrise - Val(Fz$) / 60) < 0 Then
                   t3subb$ = "none"
                Else
                   t3subb$ = Str(sunrise - Val(Fz$) / 60) 'subtract fixed minutes
                   End If
             Else
                t3subb$ = "none"
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(0, Val(bz$)) = num%
          Case -3
             t3subb$ = zmantimes(astnetznum%, numday%) 'this is astron. sunrise
             If (t3subb$ <> "none") Then
                sunrise = Val(t3subb$)
                If Str(sunrise - Val(Fz$) / 60) < 0 Then
                   t3subb$ = "none"
                Else
                   t3subb$ = Str(sunrise - Val(Fz$) / 60) 'subtract fixed minutes
                   End If
             Else
                t3subb$ = "none"
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(0, Val(bz$)) = num%
           Case Else
        End Select
        
      ElseIf Mid$(az$, 1, 8) = "Twilight" Then
      
        num% = num% + 1
        
        Select Case Val(cz$)
           Case 1 'twilight defined by solar depression
             'GoSub coordnetz
             yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
             ZA = 90 + Val(dz$)
             dy = jday%
             GoSub cal
             zmannames(num%) = az$
             zmantimes(num%, numday%) = Str(t3sub)
             If t3sub = -9999 Then
                zmantimes(num%, numday%) = "none"
                End If
             zmannumber%(1, Val(bz$)) = num%
           Case 2 'twilight defined by minutes zemanios (as defined by mishor sunrise/sunset)
             t3subb$ = zmantimes(mishornetznum%, numday%) 'this is mishor sunrise
             If (t3subb$ <> "none") Then
                sunrise = Val(t3subb$)
                t3subb$ = zmantimes(mishorskiynum%, numday%) 'this is mishor sunset
                If (t3subb$ <> "none") Then
                   sunsets = Val(t3subb$)
                   hourszemanios = (sunsets - sunrise) / 12
                   If Str(sunsets + (Val(Ez$) / 60) * hourszemanios) < 0 Then
                      t3subb$ = "none"
                   Else
                      'add minutes zemnanios to sunset
                      t3subb$ = Str(sunsets + (Val(Ez$) / 60) * hourszemanios)
                      End If
                Else
                   t3subb$ = "none"
                   End If
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(1, Val(bz$)) = num%
           Case -2 'twilight defined by minutes zemanios (as defined by astron. sunrise/sunset)
             t3subb$ = zmantimes(astnetznum%, numday%) 'this is astron. sunrise
             If (t3subb$ <> "none") Then
                sunrise = Val(t3subb$)
                t3subb$ = zmantimes(astskiynum%, numday%) 'this is astron. sunset
                If (t3subb$ <> "none") Then
                   sunsets = Val(t3subb$)
                   hourszemanios = (sunsets - sunrise) / 12
                   If Str(sunsets + (Val(Ez$) / 60) * hourszemanios) < 0 Then
                      t3subb$ = "none"
                   Else
                      'add minutes zemnanios to sunset
                      t3subb$ = Str(sunsets + (Val(Ez$) / 60) * hourszemanios)
                      End If
                Else
                   t3subb$ = "none"
                   End If
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(1, Val(bz$)) = num%
          Case 3
             t3subb$ = zmantimes(mishorskiynum%, numday%) 'this is mishor sunset
             If (t3subb$ <> "none") Then
                sunsets = Val(t3subb$)
                If Str(sunsets + Val(Fz$) / 60) < 0 Then
                   t3subb$ = "none"
                Else
                   t3subb$ = Str(sunsets + Val(Fz$) / 60) 'add fixed minutes
                   End If
             Else
                t3subb$ = "none"
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(1, Val(bz$)) = num%
          Case -3
             t3subb$ = zmantimes(astskiynum%, numday%) 'this is astron. sunset
             If (t3subb$ <> "none") Then
                sunsets = Val(t3subb$)
                If Str(sunsets + Val(Fz$) / 60) < 0 Then
                   t3subb$ = "none"
                Else
                   t3subb$ = Str(sunsets + Val(Fz$) / 60) 'add fixed minutes
                   End If
             Else
                t3subb$ = "none"
                End If
             zmannames(num%) = az$
             zmantimes(num%, numday%) = t3subb$
             zmannumber%(1, Val(bz$)) = num%
           Case Else
        End Select
      ElseIf Mid$(az$, 1, 7) = "Sunrise" Then
      
          num% = num% + 1
          
          If vis = True And (InStr(az$, "Visible") Or _
             InStr(az$, heb3$(22)) Or InStr(az$, heb3$(23))) And zmannetz = True Then
             zmannames(num%) = "visible sunrise time: "
             If optionheb Then zmannames(num%) = heb3$(16) 'הנץ הנראה'
             timtmp$ = stortim$(0, imonth% - 1, kday% - 1)
             pos% = InStr(timtmp$, "_")
             If pos% <> 0 Then
                stortim$(0, imonth% - 1, kday% - 1) = Mid$(timtmp$, 1, pos% - 1)
                End If
             zmantimes(num%, numday%) = stortim$(0, imonth% - 1, kday% - 1)
             visiblenetzzman% = num%
          ElseIf ast = True And (InStr(az$, "Astronomical") Or _
             InStr(az$, heb3$(24)) Or InStr(az$, heb3$(25))) And zmannetz = True Then
             zmannames(num%) = "astronomical sunrise time: "
             If optionheb Then zmannames(num%) = heb3$(17) 'הנץ האסטרונומי'
             zmantimes(num%, numday%) = stortim$(0, imonth% - 1, kday% - 1)
             astnetznum% = num%
          ElseIf mis = True And (InStr(az$, "Mishor") Or _
             InStr(az$, heb3$(26)) Or InStr(az$, heb3$(27))) And zmannetz = True Then
             zmannames(num%) = "mishor sunrise time: "
             If optionheb Then zmannames(num%) = heb3$(18) 'הנץ המישורי'
             zmantimes(num%, numday%) = stortim$(0, imonth% - 1, kday% - 1)
             mishornetznum% = num%
             End If
          
          If InStr(az$, "Astronomical") Or _
             InStr(az$, heb3$(24)) Or InStr(az$, heb3$(25)) Then
             'GoSub coordnetz
             hgt = avehgtnetz
             yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
             ZA = -90
             dy = jday%
             GoSub cal
             zmannames(num%) = "astronomical sunrise time: "
             If optionheb Then zmannames(num%) = heb3$(17)   'הנץ האסטרונומי'
             zmantimes(num%, numday%) = Str(t3sub)
             If t3sub = -9999 Then
                zmantimes(num%, numday%) = "none"
                End If
             astnetznum% = num%
             
          ElseIf InStr(az$, "Mishor") Or _
             InStr(az$, heb3$(26)) Or InStr(az$, heb3$(27)) Then
             'GoSub coordnetz
             hgt = 0
             yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
             ZA = -90
             dy = jday%
             GoSub cal
             zmannames(num%) = "mishor sunrise time: "
             If optionheb Then zmannames(num%) = heb3$(18) 'הנץ המישורי'
             zmantimes(num%, numday%) = Str(t3sub)
             If t3sub = -9999 Then
                zmantimes(num%, numday%) = "none"
                End If
             mishornetznum% = num%
             End If
        
        zmannumber%(0, Val(cz$)) = num%
        
      ElseIf Mid$(az$, 1, 6) = "Sunset" Then
          num% = num% + 1
          
          If vis = True And (InStr(az$, "Visible") Or _
             InStr(az$, heb3$(22)) Or InStr(az$, heb3$(23))) And zmanskiy = True Then
             'convert to 24 hour clock
             timtmp$ = stortim$(1, imonth% - 1, kday% - 1)
             pos1% = InStr(timtmp$, "_")
             If pos1% <> 0 Then
                timtmp$ = Mid$(timtmp$, 1, pos1% - 1)
                End If
             pos2% = InStr(timtmp$, "*")
             stortim$(1, imonth% - 1, kday% - 1) = timtmp$
             addtim% = 0
             If pos2% <> 0 Then
                addtim% = 1
                End If
             lentime = Len(stortim$(1, imonth% - 1, kday% - 1)) - 6
             If lentime = 1 + addtim% Then
                tmptime$ = LTrim$(Str$(Val(Mid$(stortim$(1, imonth% - 1, kday% - 1), 1, lentime)) + 12) + Mid$(stortim$(1, imonth% - 1, kday% - 1), 2, 6))
                If pos2% <> 0 Then tmptime$ = tmptime$ + "*"
             Else
                tmptime$ = LTrim$(stortim$(1, imonth% - 1, kday% - 1))
                End If
             
             zmannames(num%) = "visible sunset time: "
             If optionheb Then zmannames(num%) = heb3$(19) 'השקיעה הנראית'
             zmantimes(num%, numday%) = tmptime$
          ElseIf ast = True And (InStr(az$, "Astronomical") Or _
             InStr(az$, heb3$(24)) Or InStr(az$, heb3$(25))) And zmanskiy = True Then
             zmannames(num%) = "astronomical sunset time: "
             If optionheb Then zmannames(num%) = heb3$(20) 'השקיעה האסטרונומית'
             zmantimes(num%, numday%) = stortim$(1, imonth% - 1, kday% - 1)
             astskiynum% = num%
          ElseIf mis = True And (InStr(az$, "Mishor") Or _
             InStr(az$, heb3$(26)) Or InStr(az$, heb3$(27))) And zmanskiy = True Then
             zmannames(num%) = "mishor sunset time: "
             If optionheb Then zmannames(num%) = heb3$(21) 'השקיעה המישורית'
             zmantimes(num%, numday%) = stortim$(1, imonth% - 1, kday% - 1)
             mishorskiynum% = num%  '<<<<<<<<<*******!!!!!!!!check
             End If
          
          If InStr(az$, "Astronomical") Or _
             InStr(az$, heb3$(24)) Or InStr(az$, heb3$(25)) Then
             'GoSub coordskiy
             hgt = avehgtskiy
             yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
             ZA = 90
             dy = jday%
             GoSub cal
             zmannames(num%) = "astronomical sunset time: "
             If optionheb Then zmannames(num%) = heb3$(20) 'השקיעה האסטרונומית'
             zmantimes(num%, numday%) = Str(t3sub)
             If t3sub = -9999 Then
                zmantimes(num%, numday%) = "none"
                End If
             astskiynum% = num%
          ElseIf InStr(az$, "Mishor") Or _
             InStr(az$, heb3$(26)) Or InStr(az$, heb3$(27)) Then
             'GoSub coordskiy
             hgt = 0
             yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
             ZA = 90
             dy = jday%
             GoSub cal
             zmannames(num%) = "mishor sunset time: "
             If optionheb Then zmannames(num%) = heb3$(21) 'השקיעה המישורית'
             zmantimes(num%, numday%) = Str(t3sub)
             If t3sub = -9999 Then
                zmantimes(num%, numday%) = "none"
                End If
             mishorskiynum% = num%
             End If
        
        zmannumber%(1, Val(cz$)) = num%
        
      ElseIf Mid$(az$, 1, 7) = "Candles" Then
      
        num% = num% + 1
           
        Select Case Val(bz$)
           Case 1 'minutes before mishor sunset
             t3sub = Val(zmantimes(mishorskiynum%, numday%)) 'this is mishor sunset
             t3sub = t3sub - Val(cz$) / 60
             zmannames(num%) = az$
             zmantimes(num%, numday%) = Str(t3sub)
           Case 2 'minutes before visible sunset (if applicable)
             'convert to 24 hour clock
             pos1% = InStr(stortim$(1, imonth% - 1, kday% - 1), "*")
             timtem$ = stortim$(1, imonth% - 1, kday% - 1)
             If pos1% <> 0 Then
                timtem$ = Mid$(timtem$, 1, pos1% - 1)
                End If
             lentime = Len(timtem$) - 6
             If lentime = 1 Then
                tmptime$ = LTrim$(Str$(Val(Mid$(timtem$, 1, lentime)) + 12) + Mid$(timtem$, 2, 6))
             Else
                tmptime$ = LTrim$(timtem$)
                End If
             t3sub = Val(Mid$(tmptime$, 1, 2)) + Val(Mid$(tmptime$, 4, 2)) / 60 + Val(Mid$(tmptime$, 7, 2)) / 3600
             t3sub = t3sub - Val(dz$) / 60
             zmannames(num%) = az$
             zmantimes(num%, numday%) = Str(t3sub)
           Case Else
        End Select

      ElseIf Mid$(az$, 1, 6) = "Chazos" Then
      
          num% = num% + 1
        
          'GoSub coordnetz
          hgt = 0
          yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
          ZA = 0
          dy = jday%
          GoSub cal
          zmannames(num%) = "Noon (exact): "
          zmantimes(num%, numday%) = Str(t3sub)
          chazosnum% = num%
          zmannumber%(0, Val(cz$)) = num%
          
      ElseIf Mid$(az$, 1, 6) = "Zmanim" Then
      
          num% = num% + 1 'reserve space for it
          
          End If
   Loop
   Close #zmannum%
   'now go back and do the zemanim
   zmannum% = FreeFile
   Open drivjk$ + "zmanim.tmp" For Input As #zmannum%
   numzman% = -1
   
   Do Until EOF(zmannum%)
      numzman% = numzman% + 1
      Input #zmannum%, az$, bz$, cz$, dz$, Ez$, Fz$, GZ$, HZ$
      If Mid$(az$, 1, 6) = "Zmanim" Then
      
         'look for added or subtracted time
         BeforeAddMin1 = 0
         pos% = InStr(az$, "||1;")
         If pos% > 0 Then
            BeforeAddMin1 = Val(Mid$(az$, pos% + 4, Len(az$) - pos% - 3))
            End If
         BeforeAddMin2 = 0
         pos% = InStr(az$, "||2;")
         If pos% > 0 Then
            BeforeAddMin2 = Val(Mid$(az$, pos% + 4, Len(az$) - pos% - 3))
            End If
      
         num% = numzman%
         
         Select Case Val(bz$) 'optionz%
            Case 1 'use hours zemanios
               dawn = zmantimes(zmannumber%(0, Val(cz$)), numday%)
               If InStr(dawn, ":") Then 'calculated times->convert to hours.fracofhours
                  dawn = Val(Mid$(dawn, 1, 1)) + Val(Mid$(dawn, 3, 2)) / 60 + Val(Mid$(dawn, 6, 2)) / 3600
                  End If
               twilight = zmantimes(zmannumber%(1, Val(dz$)), numday%)
               If InStr(twilight, ":") Then 'calculated times->convert to hours.fracofhours
                  twilight = Val(Mid$(twilight, 1, 2)) + Val(Mid$(twilight, 4, 2)) / 60 + Val(Mid$(twilight, 7, 2)) / 3600
                  End If
               If RTrim$(twilight) = "none" Or RTrim$(dawn) = "none" Then
                  t3sub = -9999
               Else
                  hourszemanios = (twilight - dawn) / 12
                  If BeforeAddMin1 <> 0 Then
                     t3sub = dawn + Val(Ez$) * hourszemanios + BeforeAddMin1 / 60#
                  ElseIf BeforeAddMin2 <> 0 Then
                     t3sub = twilight + Val(Ez$) * hourszemanios + BeforeAddMin2 / 60#
                  Else
                     t3sub = dawn + Val(Ez$) * hourszemanios
                     End If
                  If t3sub < 0 Or t3sub >= 24 Then
                     t3sub = -9999
                     End If
                  End If
            Case 2 'use clock hours after dawn
               dawn = zmantimes(zmannumber%(0, Val(cz$)), numday%)
               If InStr(dawn, ":") Then 'calculated times->convert to hours.fracofhours
                  dawn = Val(Mid$(dawn, 1, 1)) + Val(Mid$(dawn, 3, 2)) / 60 + Val(Mid$(dawn, 6, 2)) / 3600
                  End If
               twilight = zmantimes(zmannumber%(1, Val(dz$)), numday%)
               If InStr(twilight, ":") Then 'calculated times->convert to hours.fracofhours
                  twilight = Val(Mid$(twilight, 1, 2)) + Val(Mid$(twilight, 4, 2)) / 60 + Val(Mid$(twilight, 7, 2)) / 3600
                  End If
               If RTrim$(twilight) = "none" Or RTrim$(dawn) = "none" Then
                  t3sub = -9999
               Else
                  If BeforeAddMin1 <> 0 Then
                     t3sub = dawn + Val(Fz$) + BeforeAddMin1 / 60#
                  ElseIf BeforeAddMin2 <> 0 Then
                     t3sub = twilight + Val(Fz$) + BeforeAddMin2 / 60#
                  Else
                     t3sub = dawn + Val(Fz$)
                     End If
                  End If
               If t3sub < 0 Or t3sub >= 24 Then
                  t3sub = -9999
                  End If
            Case 3 'Mishmaroat using mishor sunrise and sunset
                'find hoursnightzemanios
                'sunset is for previous day, sunrise is for this day
                hgt = 0
                yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
                ZA = 90
                dy = jday% - 1
                GoSub cal
                sunset = t3sub
                If t3sub = -9999 Then
                   zmantimes(num%, numday%) = "none"
                Else 'calculate the mishor sunrise
                   hgt = 0
                   yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
                   ZA = -90
                   dy = jday%
                   GoSub cal
                   sunrise = t3sub
                   If t3sub = -9999 Then
                      zmantimes(num%, numday%) = "none"
                   Else
                      hoursnightzemanios = 24 - sunset + sunrise
                      portionofnight = hoursnightzemanios * Val(dz$) * 0.01
                      t3sub = sunset + portionofnight
                      If t3sub > 24 Then t3sub = t3sub - 24
                      End If
                   
                   End If
            Case 4 'Mishmarot using astronomical sunrise and sunset
                'find hoursnightzemanios
                'sunset is for previous day, sunrise is for this day
                hgt = avehgtskiy
                yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
                ZA = 90
                dy = jday% - 1
                GoSub cal
                sunset = t3sub
                If t3sub = -9999 Then
                   zmantimes(num%, numday%) = "none"
                Else 'calculate the mishor sunrise
                   hgt = avehgtskiy
                   yr = Val(Mid$(stortim$(2, imonth% - 1, kday% - 1), 8, 4))
                   ZA = -90
                   dy = jday%
                   GoSub cal
                   sunrise = t3sub
                   If t3sub = -9999 Then
                      zmantimes(num%, numday%) = "none"
                   Else
                      hoursnightzemanios = 24 - sunset + sunrise
                      portionofnight = hoursnightzemanios * Val(dz$) * 0.01
                      t3sub = sunset + portionofnight
                      If t3sub > 24 Then t3sub = t3sub - 24
                      End If
                   
                   End If
            Case 5 'Mishmarot using visible sunrise and sunset
                If kday% > 1 Then
                   sunset = stortim$(1, imonth% - 1, kday% - 2)
                Else
                   'use the sunset stored of Rosh Hashono, since the one for Erev RH wasn't calculated
                   sunset = stortim$(1, imonth% - 1, kday% - 1)
                   End If
                If sunset = -9999 Then
                Else
                   'both sunrise and sunset are now strings, so parse out the decimal time values
                   'remove "*" for near mountains indication
                   pos% = InStr(sunset, "*")
                   If pos% > 0 Then
                      sunset = Mid$(sunset, 1, pos% - 1)
                      End If
                   If Len(sunset) = 7 And InStr(sunset, ":") Then
                      sunset = Val(Mid$(sunset, 1, 1)) + Val(Mid$(sunset, 3, 2)) / 60 + Val(Mid$(sunset, 6, 2)) / 3600
                   ElseIf Len(sunset) = 8 And InStr(sunset, ":") Then
                      sunset = Val(Mid$(sunset, 1, 2)) + Val(Mid$(sunset, 4, 2)) / 60 + Val(Mid$(sunset, 7, 2)) / 3600
                   Else
                      sunset = Val(sunset)
                      End If
                   
                   sunrise = stortim$(0, imonth% - 1, kday% - 1)
                   pos% = InStr(sunrise, "*")
                   If pos% > 0 Then
                      sunrise = Mid$(sunrise, 1, pos% - 1)
                      End If
                   If Len(sunrise) = 7 And InStr(sunrise, ":") Then
                      sunrise = Val(Mid$(sunrise, 1, 1)) + Val(Mid$(sunrise, 3, 2)) / 60 + Val(Mid$(sunrise, 6, 2)) / 3600
                   ElseIf Len(sunset) = 8 And InStr(sunset, ":") Then
                      sunrise = Val(Mid$(sunrise, 1, 2)) + Val(Mid$(sunrise, 4, 2)) / 60 + Val(Mid$(sunrise, 7, 2)) / 3600
                   Else
                      sunrise = Val(sunrise)
                      End If
                   If sunrise = -9999 Then
                      t3sub = -9999
                   Else
                      hoursnightzemanios = 12 - sunset + sunrise
                      portionofnight = hoursnightzemanios * Val(dz$) * 0.01
                      t3sub = sunset + portionofnight
                      'convert to 24 hour clock
                      t3sub = t3sub + 12
                      If t3sub > 24 Then t3sub = t3sub - 24
                      End If
                   End If
            Case Else
         End Select
         
         zmantimes(num%, numday%) = Str(t3sub)
         If t3sub = -9999 Then
            zmantimes(num%, numday%) = "none"
            End If
         End If
   Loop
   Close #zmannum%
   
   'now add the zemanim to the List Box
   '************sort them here****************
   zmannum% = FreeFile
   Open drivjk$ + "zmanim.tmp" For Input As #zmannum%
   numend% = 0
   If reorder = True Then
      numend% = numsort% ' number of sorted zmanim to display
      sortnum% = FreeFile
      Open drivjk$ + "zmansort.out" For Input As #sortnum%
      End If
   nextnum% = -1
   For inum% = 0 To numend%
        If reorder = True Then
           Input #sortnum%, listnum%
           If inum% > 0 Then Seek #zmannum%, 1 'rewind the zmanim list file
           End If
        newnum% = -1
        Do Until EOF(zmannum%)
           newnum% = newnum% + 1
           If reorder = True Then
              If newnum% <> listnum% Then
                 Input #zmannum%, az$, bz$, cz$, dz$, Ez$, Fz$, GZ$, HZ$
                 GoTo 900
                 End If
              End If
           Input #zmannum%, az$, bz$, cz$, dz$, Ez$, Fz$, GZ$, HZ$
           nextnum% = nextnum% + 1
           'If LTrim$(zmantimes(newnum%, numday%)) = "none" Then
           '   t3subb$ = "none"
           If Mid$(az$, 1, 7) = "Candles" Then 'list erev shabbos candle lighting times
               If stortim$(4, imonth% - 1, kday% - 1) = heb4$(6) Then
                  t3sub = Val(zmantimes(newnum%, numday%))
                  GoSub 1500
                  plusround% = Val(GZ$)
                  steps = Val(HZ$)
                  GoSub round
                  If steps = 60 Then 'truncate seconds
                    t3subb$ = Mid$(t3subb$, 1, Len(t3subb$) - 3)
                    End If
               ElseIf stortim$(4, imonth% - 1, kday% - 1) <> heb4$(6) Then
                  'also list the candle lighting for erev yom tov
                  Select Case stortim$(3, imonth% - 1, kday% - 1)
                     'Case "כט-אלול", "ט-תשרי", "יד-תשרי", "כא-תשרי", "יד-ניסן", "כ-ניסן", "ה-סיון", "29-Elul", "9-Tishrey", "14-Tishrey", "21-Tishrey", "14-Nisan", "20-Nisan", "5-Sivan"
                     Case heb6$(1), heb6$(2), heb6$(3), heb6$(4), heb6$(5), heb6$(6), heb6$(7), _
                          "29-Elul", "9-Tishrey", "14-Tishrey", "21-Tishrey", "14-Nisan", "20-Nisan", "5-Sivan"
                         If stortim$(4, imonth% - 1, kday% - 1) = heb4$(7) Then
                            'this is erev yom-tov that falls on shabbos (NO CANDLE LIGHTING!)
                            newzmans(nextnum%) = "00:00:00"
                            If Val(HZ$) = 60 Then newzmans(nextnum%) = "00:00"
                            zmannames$(nextnum%) = az$
                            GoTo 900 'don't print out the candle lighting time
                            End If
                         t3sub = Val(zmantimes(newnum%, numday%))
                         GoSub 1500
                         plusround% = Val(GZ$)
                         steps = Val(HZ$)
                         GoSub round
                         If steps = 60 Then 'truncate seconds
                            t3subb$ = Mid$(t3subb$, 1, Len(t3subb$) - 3)
                            End If
                     Case Else
                        newzmans(nextnum%) = "00:00:00"
                        If Val(HZ$) = 60 Then newzmans(nextnum%) = "00:00"
                        zmannames$(nextnum%) = az$
                        GoTo 900 'don't print out the candle lighting time
                  End Select
                  End If
           ElseIf InStr(zmantimes(newnum%, numday%), ":") Then
              'time is already in clock format
              'these are the times appearing in the netz/skiya table
              t3subb$ = zmantimes(newnum%, numday%)
           Else
              If Trim$(zmantimes(newnum%, numday%)) = "none" Then
                 t3subb$ = "none"
                 GoTo 850
                 End If
              t3sub = Val(zmantimes(newnum%, numday%))
              GoSub 1500
              plusround% = Val(GZ$)
              steps = Val(HZ$)
              GoSub round
              If steps = 60 Then 'truncate seconds
                 t3subb$ = Mid$(t3subb$, 1, Len(t3subb$) - 3)
                 End If
              End If
850        Zmanimlistfm.List1.AddItem az$ + ": " + t3subb$
           Print #newlistnum%, az$ + ": " + t3subb$
           newzmans(nextnum%) = t3subb$ '  zmantimes(nextnum%, numday%) = t3subb$ 'use this for tables
           'remove the "Zmanim", etc from the name
           pos% = InStr(az$, ":")
           If pos% > 0 Then
              zmannames$(nextnum%) = Trim$(Mid$(az$, pos% + 1, Len(az$) - pos%))
           Else
              zmannames$(nextnum%) = az$
              End If
           If optionheb Then 'convert sunrises/sunsets into hebrew
              If InStr(zmannames$(nextnum%), "Mishor Sunrise") Then
                 zmannames$(nextnum%) = heb3$(18)
              ElseIf InStr(zmannames$(nextnum%), "Mishor Sunset") Then
                 zmannames$(nextnum%) = heb3$(21)
              ElseIf InStr(zmannames$(nextnum%), "Astronomical Sunrise") Then
                 zmannames$(nextnum%) = heb3$(17)
              ElseIf InStr(zmannames$(nextnum%), "Astronomical Sunset") Then
                 zmannames$(nextnum%) = heb3$(20)
              ElseIf InStr(zmannames$(nextnum%), "Visible Sunrise") Then
                 zmannames$(nextnum%) = heb3$(16)
              ElseIf InStr(zmannames$(nextnum%), "Visible Sunset") Then
                 zmannames$(nextnum%) = heb3$(19)
              ElseIf InStr(zmannames$(nextnum%), "Chazos") Then
                 zmannames$(nextnum%) = heb3$(28)
              ElseIf InStr(zmannames$(nextnum%), "Mishmarot") Then
                 'no hebrew string yet, so use the english one
                 End If
              End If
           If reorder = True Then GoTo 950
900      Loop
950   Next inum%
   Close #zmannum%
   If reorder = True Then
      Close #sortnum%
      newnum% = numsort%
      For mm% = 0 To nextnum%
         zmantimes(mm%, numday%) = newzmans(mm%)
      Next mm%
      End If
      
'      If hourszemanios > MaxHourZemanios Then MaxHourZemanios = hourszemanios
      
Return

generrhand:
     Screen.MousePointer = vbDefault
     If internet = True And Err.Number >= 0 Then   'exit the program
        'abort the program with a error messages
        errlog% = FreeFile
        Open drivjk$ + "Cal_zcbgeh.log" For Output As #errlog%
        Print #errlog%, "Cal Prog exited from zmanimform calendarbut with runtime error message " + Str(Err.Number)
        Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
        Close #errlog%
        Close
      
       'unload forms
        For i% = 0 To Forms.Count - 1
          Unload Forms(i%)
        Next i%
      
        myfile = Dir(drivfordtm$ + "busy.cal")
        If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      
        'kill the timer
        If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
        'bring program to abrupt end
        End
     
     Else
        Resume
        response = MsgBox("Zmanimform calendarbut encountered error number: " + Str(Err.Number) + ".  Do you want to abort?", vbYesNoCancel + vbCritical, "Cal Program")
        If response <> vbYes Then
           Close
           Exit Sub
        Else
           Close
           For i% = 0 To Forms.Count - 1
             Unload Forms(i%)
           Next i%
           End
           End If
        End If
        
End Sub

Private Sub Form_Load()
   'version: 04/08/2003
  
   On Error GoTo generrhand
   neworder = False
   reorder = False
   resortbutton = False
   'kill zmansort.out if it exists
   If Dir(drivjk$ + "zmansort.out") <> sEmpty Then
      Kill drivjk$ + "zmansort.out"
      End If
   If vis = False Or (vis = True And zmanskiy = False) Then
      Option10.Enabled = False
      End If
   init = True
   zmantotal% = 0
   vis1% = 0
   ast1% = 0
   mis1% = 0
   optiontmish% = 0
   optiondmish% = 0
   Combo1.Clear
   Combo2.Clear
   Combo3.Clear
   zmanopen = True
   savedit = False
   
   'flag adding sedra/holiday info
   If parshiotEY Then
      optNoParshiot.Value = True 'reset flag
      optEYParshiot.Value = True 'set flag
   ElseIf parshiotdiaspora Then
      optNoParshiot.Value = True 'reset flag
      optDiasporaParshiot.Value = True
   Else
      optNoParshiot.Value = True
      End If
   
   myfile = Dir(drivjk$ + "zmanim.tmp")
   If myfile <> sEmpty And zmanopen = True Then
      If internet = True Then Exit Sub
      response = MsgBox("Do you wan't to record new set of parameters? (answer NO if you want to append these zmanim to previously recorded ones)", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Cal Program")
      If response = vbNo Then 'load last recorded parameters
        init = False
        zmanopen = False
        zmannum% = FreeFile
        Open drivjk$ + "zmanim.tmp" For Input As zmannum%
        Do Until EOF(zmannum%)
           Input #zmannum%, a$, b$, c$, D$, e$, f$, GZ$, HZ$
           Combo1.AddItem a$
           If InStr(a$, "Dawn") Or InStr(a$, "Sunrise") Or InStr(a$, "Chazos") Then
              Combo2.AddItem a$
           ElseIf InStr(a$, "Twilight") Or InStr(a$, "Sunset") Then
              Combo3.AddItem a$
              End If
           Combo1.ListIndex = Combo1.ListCount - 1
           If Val(GZ$) = -1 Then
              Option17.Value = True
              Text15.Text = HZ$
           ElseIf Val(GZ$) = 1 Then
              Option18.Value = True
              Text16.Text = HZ$
              End If
        Loop

        Close #zmannum%
        Exit Sub
      ElseIf response <> vbNo Then 'kill old zmanim.tmp file
        Close
        On Error GoTo errhand
        Kill drivjk$ + "zmanim.tmp"
        neworder = False
        reorder = False
        End If
      End If
   
   Select Case optiond% 'dawns
      Case 0, 1
         Option1.Value = True
         optiond% = 1
      Case 2, -2
         Option2.Value = True
         If optiond% = -2 Then Option21.Value = True
      Case 3, -3
         Option3.Value = True
         If optiond% = -3 Then Option21.Value = True
      Case Else
   End Select
   Select Case optiont% 'twilights
      Case 0, 1
         Option7.Value = True
         optiont% = 1
      Case 2, -2
         Option4.Value = True
         If optiont% = -2 Then Option20.Value = True
      Case 3, -3
         Option5.Value = True
         If optiont% = -3 Then Option20.Value = True
      Case Else
   End Select
   Select Case optionz% 'zemanim
      Case 0, 1
         Option6.Value = True
         optionz% = 1
      Case 2
         Option8.Value = True
      Case Else
   End Select
   Select Case options% 'candle lighting
      Case 0, 1
         Option9.Value = True
         options% = 1
      Case 2
         Option10.Value = True
      Case Else
   End Select
   'now add a listing for the mishor sunrise and sunset
   misflg% = 0
   If mis = True Then misflg% = 1
   mis = True
   mis1% = 1
   Option18.Value = True
   Text16.Text = "5"
   Call Command1_Click
   Call Combo1_Click
   mis1% = 2
   Option17.Value = True
   Text15.Text = "5"
   Call Command1_Click
   Call Combo1_Click
   If misflg% = 0 Then mis = False
   'now add a listing for the visible sunrise or visible sunset or both
   'if they were calculated (this is determined by the value of nsetflag%)
   If vis = True Then
        Select Case vis0%
           Case 1
             vis1% = 1
             Option18.Value = True
             Text16.Text = "5"
             Call Command1_Click
             Call Combo1_Click
           Case 2
             vis1% = 2
             Option17.Value = True
             Text15.Text = "5"
             Call Command1_Click
             Call Combo1_Click
           Case 3
             vis1% = 1
             Option18.Value = True
             Text16.Text = "5"
             Call Command1_Click
             Call Combo1_Click
             vis1% = 2
             Option17.Value = True
             Text15.Text = "5"
             Call Command1_Click
             Call Combo1_Click
           Case Else
        End Select
        End If
   'now add a listing for the astronomical sunrise,sunset if calculated
   If ast = True Then
      Select Case ast0%
         Case 1
            ast = True
            ast1% = 1
            Option18.Value = True
            Text16.Text = "5"
            Call Command1_Click
            Call Combo1_Click
            Option20.Enabled = True
         Case 2
            ast = True
            ast1% = 2
            Option17.Value = True
            Text15.Text = "5"
            Call Command1_Click
            Call Combo1_Click
            Option21.Enabled = True
         Case 3
            ast = True
            ast1% = 1
            Option18.Value = True
            Text16.Text = "5"
            Call Command1_Click
            Call Combo1_Click
            ast1% = 2
            Option17.Value = True
            Text15.Text = "5"
            Call Command1_Click
            Call Combo1_Click
            Option20.Enabled = True
            Option21.Enabled = True
         Case Else
      End Select
      End If
    'now add a listing for exact noon
    noon = True
    Call Command1_Click
    Call Combo1_Click
    noon = False
    
    parshiotEY = False
    parshiotdiaspora = False
   ' 'now add a listing for parshiot names
   ' parshiotEY = True
   ' Call Command1_Click
   ' Call Combo1_click
   ' parshiotEY = False
   ' parshiotdiaspora = True
   ' Call Command1_Click
   ' Call Combo1_click
   ' parshiotdiaspora = False
    
    init = False
    
    Close
    Exit Sub
    
errhand:
   Resume Next
   
generrhand:
     If Err.Number >= 0 And internet = True Then 'exit the program
        'abort the program with a error messages
        errlog% = FreeFile
        Open drivjk$ + "Cal_zflgeh.log" For Output As #errlog%
        Print #errlog%, "Cal Prog exited from zmanimform Form_Load with runtime error message " + Str(Err.Number)
        Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
        Close #errlog%
        Close
      
       'unload forms
        For i% = 0 To Forms.Count - 1
          Unload Forms(i%)
        Next i%
      
        myfile = Dir(drivfordtm$ + "busy.cal")
        If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      
        'kill the timer
        If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
        'bring program to abrupt end
        End
     Else
        response = MsgBox("Zmanimform Form_Load encountered error number: " + Str(Err.Number) + ".  Do you want to abort?", vbYesNoCancel + vbCritical, "Cal Program")
        If response <> vbYes Then
           Close
           Exit Sub
        Else
           Close
           For i% = 0 To Forms.Count - 1
             Unload Forms(i%)
           Next i%
           End
           End If
        End If
   
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   zmanopen = False
   Close
   If internet = False And savedit = False And load = False And Combo1.ListCount >= 1 And changes = True Then
      response = MsgBox("Do you want to save these parameters to a permanent file?", vbQuestion + vbYesNo, "Cal Program")
      If response = vbYes Then
         Cancel = True
         Exit Sub
         End If
      End If
   load = False
   If Dir(drivjk$ + "zmansort.out") <> sEmpty Then
      Kill drivjk$ + "zmansort.out"
      End If
   If Dir(drivjk$ + "zmannew.tmp") <> sEmpty Then
      Kill drivjk$ + "zmannew.tmp"
      End If
   Unload Zmanimform
End Sub

Private Sub Option1_Click()
   optiond% = 1
   Text1.Text = "0"
   Text1.Enabled = True
   UpDown1.Enabled = True
   Text2.Text = "0"
   Text2.Enabled = False
   UpDown2.Enabled = False
   Text10.Text = "0"
   Text10.Enabled = False
   UpDown8.Enabled = False
   Text3.Text = sEmpty
End Sub

Private Sub Option10_Click()
   options% = 2
   Text12.Text = "0"
   Text12.Enabled = False
   UpDown9.Enabled = False
   Text13.Text = "0"
   Text13.Enabled = True
   UpDown10.Enabled = True
   Text14.Text = sEmpty
End Sub

Private Sub Option17_Click()
   optionround% = -1
   Text16.Text = sEmpty
End Sub

Private Sub Option18_Click()
   optionround% = 1
   Text15.Text = sEmpty
End Sub

Private Sub Option2_Click()
   optiond% = 2
   Text1.Text = "0"
   Text1.Enabled = False
   UpDown1.Enabled = False
   Text2.Text = "0"
   Text2.Enabled = True
   UpDown2.Enabled = True
   Text10.Text = "0"
   Text10.Enabled = False
   UpDown8.Enabled = False
   Text3.Text = sEmpty
End Sub

Private Sub Option3_Click()
   optiond% = 3
   Text1.Text = "0"
   Text1.Enabled = False
   UpDown1.Enabled = False
   Text2.Text = "0"
   Text2.Enabled = False
   UpDown2.Enabled = False
   Text10.Text = "0"
   Text10.Enabled = True
   UpDown8.Enabled = True
   Text3.Text = sEmpty
End Sub

Private Sub Option4_Click()
   optiont% = 2
   Text4.Text = "0"
   Text4.Enabled = False
   UpDown3.Enabled = False
   Text5.Text = "0"
   Text5.Enabled = True
   UpDown4.Enabled = True
   Text6.Text = "0"
   Text6.Enabled = False
   UpDown5.Enabled = False
   Text7.Text = sEmpty
End Sub

Private Sub Option5_Click()
   optiont% = 3
   Text4.Text = "0"
   Text4.Enabled = False
   UpDown3.Enabled = False
   Text5.Text = "0"
   Text5.Enabled = False
   UpDown4.Enabled = False
   Text6.Text = "0"
   Text6.Enabled = True
   UpDown5.Enabled = True
   Text7.Text = sEmpty
End Sub

Private Sub Option6_Click()
   optionz% = 1
   Text8.Text = "0"
   Text8.Enabled = True
   UpDown6.Enabled = True
   Text9.Text = "0"
   Text9.Enabled = False
   UpDown7.Enabled = False
   Text11.Text = sEmpty
End Sub

Private Sub Option7_Click()
   optiont% = 1
   Text4.Text = "0"
   Text4.Enabled = True
   UpDown3.Enabled = True
   Text5.Text = "0"
   Text5.Enabled = False
   UpDown4.Enabled = False
   Text6.Text = "0"
   Text6.Enabled = False
   UpDown5.Enabled = False
   Text7.Text = sEmpty
End Sub

Private Sub Option8_Click()
   optionz% = 2
   Text8.Text = "0"
   Text8.Enabled = False
   UpDown6.Enabled = False
   Text9.Text = "0"
   Text9.Enabled = True
   UpDown7.Enabled = True
   Text11.Text = sEmpty
End Sub

Private Sub Option9_Click()
   options% = 1
   Text12.Text = "0"
   Text12.Enabled = True
   UpDown9.Enabled = True
   Text13.Text = "0"
   Text13.Enabled = False
   UpDown10.Enabled = False
   Text14.Text = sEmpty
End Sub

