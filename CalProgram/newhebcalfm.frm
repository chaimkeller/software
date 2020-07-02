VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form newhebcalfm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cal Program (Enter Calendar Parameters)"
   ClientHeight    =   8130
   ClientLeft      =   1695
   ClientTop       =   330
   ClientWidth     =   8775
   FillColor       =   &H00C0C0C0&
   Icon            =   "newhebcalfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   8775
   Begin MSComDlg.CommonDialog CommonDialog6 
      Left            =   7800
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton newhebOpenbut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Picture         =   "newhebcalfm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   7320
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   360
      TabIndex        =   33
      Top             =   1440
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   5617386
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Top Calendar (usually sunrise)"
      TabPicture(0)   =   "newhebcalfm.frx":0974
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Bottom Calendar (usually sunset)"
      TabPicture(1)   =   "newhebcalfm.frx":0990
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame10 
         Height          =   2295
         Left            =   -68400
         TabIndex        =   67
         Top             =   360
         Width           =   1335
         Begin MSComCtl2.UpDown UpDown27 
            Height          =   285
            Left            =   960
            TabIndex        =   133
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text40"
            BuddyDispid     =   196611
            OrigLeft        =   960
            OrigTop         =   1560
            OrigRight       =   1200
            OrigBottom      =   1815
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text40 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   360
            TabIndex        =   132
            Text            =   "8200"
            Top             =   1560
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown25 
            Height          =   285
            Left            =   960
            TabIndex        =   85
            Top             =   1920
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text38"
            BuddyDispid     =   196612
            OrigLeft        =   840
            OrigTop         =   1800
            OrigRight       =   1080
            OrigBottom      =   2055
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text38 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   360
            TabIndex        =   84
            Text            =   "350"
            Top             =   1920
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown24 
            Height          =   285
            Left            =   960
            TabIndex        =   83
            Top             =   1320
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text37"
            BuddyDispid     =   196613
            OrigLeft        =   840
            OrigTop         =   1440
            OrigRight       =   1080
            OrigBottom      =   1695
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text37 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   360
            TabIndex        =   82
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown23 
            Height          =   285
            Left            =   960
            TabIndex        =   81
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text36"
            BuddyDispid     =   196614
            OrigLeft        =   840
            OrigTop         =   1320
            OrigRight       =   1080
            OrigBottom      =   1575
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text36 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   360
            TabIndex        =   80
            Text            =   "900"
            Top             =   1080
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown22 
            Height          =   285
            Left            =   960
            TabIndex        =   79
            Top             =   840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text35"
            BuddyDispid     =   196615
            OrigLeft        =   840
            OrigTop         =   1080
            OrigRight       =   1080
            OrigBottom      =   1335
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text35 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   360
            TabIndex        =   78
            Text            =   "700"
            Top             =   840
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown21 
            Height          =   285
            Left            =   960
            TabIndex        =   77
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text34"
            BuddyDispid     =   196616
            OrigLeft        =   840
            OrigTop         =   840
            OrigRight       =   1080
            OrigBottom      =   1095
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text34 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   360
            TabIndex        =   76
            Text            =   "300"
            Top             =   600
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown20 
            Height          =   285
            Left            =   960
            TabIndex        =   75
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text33"
            BuddyDispid     =   196617
            OrigLeft        =   840
            OrigTop         =   360
            OrigRight       =   1080
            OrigBottom      =   615
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text33 
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   360
            TabIndex        =   74
            Text            =   "8900"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Y5"
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
            Left            =   120
            TabIndex        =   134
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label38 
            Caption         =   "De"
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
            Left            =   120
            TabIndex        =   91
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label37 
            Caption         =   "Y4"
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
            TabIndex        =   90
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label36 
            Caption         =   "Y3"
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
            Left            =   120
            TabIndex        =   89
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label35 
            Caption         =   "Y2"
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
            Left            =   120
            TabIndex        =   88
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label34 
            Caption         =   "Y1"
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
            Left            =   120
            TabIndex        =   87
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label33 
            Caption         =   "Xc"
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
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00008000&
         Height          =   2295
         Left            =   -73440
         TabIndex        =   66
         Top             =   360
         Width           =   5055
         Begin VB.ComboBox Combo10 
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   120
            TabIndex        =   131
            Text            =   "הזמנים לפי שעון חורף, אין להשתמש בזמנים אלו כדי לקבוע שעות היום"
            Top             =   1800
            Width           =   4815
         End
         Begin VB.ComboBox Combo9 
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   120
            TabIndex        =   130
            Top             =   1440
            Width           =   4815
         End
         Begin VB.ComboBox Combo8 
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   120
            TabIndex        =   129
            Text            =   "בסיוע המודל טופוגרפי של ארץ ישראל"
            Top             =   1080
            Width           =   4815
         End
         Begin VB.ComboBox Combo7 
            ForeColor       =   &H00008000&
            Height          =   315
            Left            =   120
            TabIndex        =   128
            Text            =   "מבוסס על השקיעה המאוחרת ביותר הנראית בעיר כולה"
            Top             =   720
            Width           =   4815
         End
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "Narkisim"
               Size            =   14.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   405
            Left            =   120
            TabIndex        =   127
            Text            =   "לוח ""בכורי יוסף"" לשקיעת החמה ל"
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   1455
         Begin MSComCtl2.UpDown UpDown19 
            Height          =   495
            Left            =   960
            TabIndex        =   71
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text32"
            BuddyDispid     =   196632
            OrigLeft        =   840
            OrigTop         =   1440
            OrigRight       =   1080
            OrigBottom      =   1935
            Max             =   300
            Min             =   -300
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text32 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   240
            TabIndex        =   70
            Text            =   "-15"
            Top             =   1560
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDown18 
            Height          =   495
            Left            =   840
            TabIndex        =   69
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   393216
            Value           =   15
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text31"
            BuddyDispid     =   196633
            OrigLeft        =   960
            OrigTop         =   480
            OrigRight       =   1200
            OrigBottom      =   975
            Max             =   60
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text31 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   360
            TabIndex        =   68
            Text            =   "5"
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Add/Sub Sec."
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
            Left            =   120
            TabIndex        =   73
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Round to Sec"
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
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   2295
         Left            =   6600
         TabIndex        =   36
         Top             =   360
         Width           =   1335
         Begin MSComCtl2.UpDown UpDown26 
            Height          =   285
            Left            =   960
            TabIndex        =   125
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text39"
            BuddyDispid     =   196637
            OrigLeft        =   960
            OrigTop         =   1560
            OrigRight       =   1200
            OrigBottom      =   1815
            Max             =   99999
            Min             =   -99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text39 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            TabIndex        =   124
            Text            =   "8200"
            Top             =   1560
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown17 
            Height          =   285
            Left            =   960
            TabIndex        =   58
            Top             =   1920
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text30"
            BuddyDispid     =   196638
            OrigLeft        =   840
            OrigTop         =   1920
            OrigRight       =   1080
            OrigBottom      =   2175
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text30 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            TabIndex        =   57
            Text            =   "350"
            Top             =   1920
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown13 
            Height          =   285
            Left            =   960
            TabIndex        =   56
            Top             =   1320
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text26"
            BuddyDispid     =   196639
            OrigLeft        =   840
            OrigTop         =   1440
            OrigRight       =   1080
            OrigBottom      =   1695
            Max             =   99999
            Min             =   -99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text26 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            TabIndex        =   55
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown6 
            Height          =   285
            Left            =   960
            TabIndex        =   54
            Top             =   1080
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text19"
            BuddyDispid     =   196640
            OrigLeft        =   840
            OrigTop         =   1200
            OrigRight       =   1080
            OrigBottom      =   1455
            Max             =   99999
            Min             =   -99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text19 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            TabIndex        =   53
            Text            =   "900"
            Top             =   1080
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown5 
            Height          =   285
            Left            =   960
            TabIndex        =   52
            Top             =   840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text18"
            BuddyDispid     =   196641
            OrigLeft        =   840
            OrigTop         =   960
            OrigRight       =   1080
            OrigBottom      =   1215
            Max             =   99999
            Min             =   -99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text18 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            TabIndex        =   51
            Text            =   "700"
            Top             =   840
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown4 
            Height          =   285
            Left            =   960
            TabIndex        =   50
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text17"
            BuddyDispid     =   196642
            OrigLeft        =   840
            OrigTop         =   600
            OrigRight       =   1080
            OrigBottom      =   855
            Max             =   99999
            Min             =   -99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text17 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            TabIndex        =   49
            Text            =   "300"
            Top             =   600
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   285
            Left            =   960
            TabIndex        =   48
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text16"
            BuddyDispid     =   196643
            OrigLeft        =   840
            OrigTop         =   360
            OrigRight       =   1080
            OrigBottom      =   735
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text16 
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            TabIndex        =   47
            Text            =   "8900"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Y5"
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
            Left            =   120
            TabIndex        =   126
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "De"
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
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y4"
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
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y3"
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
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y2"
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
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y1"
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
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Xc"
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
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   2295
         Left            =   1560
         TabIndex        =   35
         Top             =   360
         Width           =   5055
         Begin VB.ComboBox Combo5 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   123
            Text            =   "הזמנים לפי שעון חורף-אין להשתמש בזמנים אלו כדי לקבוע שעות היום"
            Top             =   1800
            Width           =   4815
         End
         Begin VB.ComboBox Combo4 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   46
            Top             =   1440
            Width           =   4815
         End
         Begin VB.ComboBox Combo3 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   45
            Text            =   " בסיוע מודל טופוגרפי של ארץ ישראל"
            Top             =   1080
            Width           =   4815
         End
         Begin VB.ComboBox Combo2 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Text            =   " מבוסס על הזריחה המוקדמת ביותר הנראית בעיר כולה  מידי יום יום"
            Top             =   720
            Width           =   4815
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Narkisim"
               Size            =   14.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   120
            TabIndex        =   43
            Text            =   "לוח ""בכורי יוסף"" לנץ החמה ל"
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   2295
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1455
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   495
            Left            =   960
            TabIndex        =   40
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text2"
            BuddyDispid     =   196658
            OrigLeft        =   960
            OrigTop         =   1560
            OrigRight       =   1200
            OrigBottom      =   2055
            Max             =   300
            Min             =   -300
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   300
            TabIndex        =   39
            Text            =   "15"
            Top             =   1560
            Width           =   675
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   495
            Left            =   840
            TabIndex        =   38
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text1"
            BuddyDispid     =   196659
            OrigLeft        =   960
            OrigTop         =   600
            OrigRight       =   1200
            OrigBottom      =   1095
            Max             =   60
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   360
            TabIndex        =   37
            Text            =   "5"
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Add/Sub Sec."
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
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Round to Sec"
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
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2655
      Left            =   7080
      TabIndex        =   22
      Top             =   4320
      Width           =   1335
      Begin MSComCtl2.UpDown UpDown15 
         Height          =   285
         Left            =   960
         TabIndex        =   109
         Top             =   2160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text28"
         BuddyDispid     =   196663
         OrigLeft        =   960
         OrigTop         =   2160
         OrigRight       =   1200
         OrigBottom      =   2415
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   360
         TabIndex        =   108
         Text            =   "250"
         Top             =   2160
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown14 
         Height          =   285
         Left            =   960
         TabIndex        =   107
         Top             =   1920
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text27"
         BuddyDispid     =   196664
         OrigLeft        =   960
         OrigTop         =   1920
         OrigRight       =   1200
         OrigBottom      =   2175
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   360
         TabIndex        =   106
         Text            =   "250"
         Top             =   1920
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown12 
         Height          =   285
         Left            =   960
         TabIndex        =   105
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text25"
         BuddyDispid     =   196665
         OrigLeft        =   960
         OrigTop         =   1320
         OrigRight       =   1200
         OrigBottom      =   1575
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   360
         TabIndex        =   104
         Text            =   "212"
         Top             =   1320
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown11 
         Height          =   285
         Left            =   960
         TabIndex        =   103
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text24"
         BuddyDispid     =   196666
         OrigLeft        =   960
         OrigTop         =   1080
         OrigRight       =   1200
         OrigBottom      =   1335
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   360
         TabIndex        =   102
         Text            =   "1020"
         Top             =   1080
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown10 
         Height          =   285
         Left            =   960
         TabIndex        =   101
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text23"
         BuddyDispid     =   196667
         OrigLeft        =   960
         OrigTop         =   480
         OrigRight       =   1200
         OrigBottom      =   735
         Max             =   99999
         Min             =   -99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   360
         TabIndex        =   100
         Text            =   "1500"
         Top             =   480
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown9 
         Height          =   285
         Left            =   960
         TabIndex        =   99
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text22"
         BuddyDispid     =   196668
         OrigLeft        =   840
         OrigTop         =   240
         OrigRight       =   1080
         OrigBottom      =   495
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   360
         TabIndex        =   98
         Text            =   "14000"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C000&
         Caption         =   "Month Sep."
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
         Left            =   360
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C000&
         Caption         =   "DB"
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
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C000&
         Caption         =   "DT"
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
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C000&
         Caption         =   "Time Entries"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C000&
         Caption         =   "Xo"
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
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C000&
         Caption         =   "Yo"
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
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C000&
         Caption         =   "Cell Size"
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
         Left            =   360
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C000&
         Caption         =   "Dx"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C000&
         Caption         =   "Dy"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   1215
      Left            =   7080
      TabIndex        =   18
      Top             =   120
      Width           =   1335
      Begin MSComCtl2.UpDown UpDown16 
         Height          =   285
         Left            =   960
         TabIndex        =   97
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text29"
         BuddyDispid     =   196679
         OrigLeft        =   960
         OrigTop         =   840
         OrigRight       =   1200
         OrigBottom      =   1095
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   360
         TabIndex        =   96
         Text            =   "10000"
         Top             =   840
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown8 
         Height          =   285
         Left            =   960
         TabIndex        =   95
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text21"
         BuddyDispid     =   196680
         OrigLeft        =   960
         OrigTop         =   480
         OrigRight       =   1200
         OrigBottom      =   735
         Max             =   99999
         Min             =   -99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   360
         TabIndex        =   94
         Text            =   "1000"
         Top             =   480
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown7 
         Height          =   285
         Left            =   960
         TabIndex        =   93
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text20"
         BuddyDispid     =   196681
         OrigLeft        =   960
         OrigTop         =   240
         OrigRight       =   1200
         OrigBottom      =   495
         Max             =   99999
         Min             =   -99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   360
         TabIndex        =   92
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ys"
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
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Main Body"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Yo"
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
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Xo"
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
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   360
      TabIndex        =   16
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   4440
         TabIndex        =   144
         Text            =   "A4"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text42 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   285
         Left            =   2760
         TabIndex        =   143
         Text            =   "Paper: Standard; font: hebrew/regualr year"
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   2760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "newhebcalfm.frx":09AC
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Paper Format:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   225
         Left            =   2760
         TabIndex        =   145
         Top             =   860
         Width           =   1575
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Click on the desired orientaion"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   480
         TabIndex        =   122
         Top             =   0
         Width           =   2175
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00C00000&
         Height          =   735
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1800
         Picture         =   "newhebcalfm.frx":09CE
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "newhebcalfm.frx":1110
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2655
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   6495
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Change"
         Height          =   375
         Left            =   5400
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog5 
         Left            =   3720
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CommandButton newhebColorbut 
         BackColor       =   &H00C0C000&
         Caption         =   "Palette"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0C000&
         Caption         =   "Color Fill"
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
         Left            =   3360
         TabIndex        =   141
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Fill every 4 rows"
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
         Left            =   4680
         TabIndex        =   140
         Top             =   2280
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Fill every 3 rows"
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
         Left            =   4680
         TabIndex        =   139
         Top             =   2040
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C000&
         Caption         =   "מסמן שבת"
         BeginProperty Font 
            Name            =   "Narkisim"
            Size            =   11.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   138
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C000&
         Caption         =   "Grid Lines"
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
         Left            =   1800
         TabIndex        =   137
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C000&
         Caption         =   "Copyright"
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
         Left            =   480
         TabIndex        =   136
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "בס""ד"
         BeginProperty Font 
            Name            =   "Narkisim"
            Size            =   11.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   135
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   121
         Text            =   "David"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   120
         Text            =   "David"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text4 
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
         Left            =   1080
         TabIndex        =   119
         Text            =   "David"
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   118
         Text            =   "David"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Text10 
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
         Left            =   3600
         TabIndex        =   117
         Text            =   "Bold"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text9 
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
         Left            =   3600
         TabIndex        =   116
         Text            =   "Bold"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text8 
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
         Left            =   3600
         TabIndex        =   115
         Text            =   "Bold"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text7 
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
         Left            =   3600
         TabIndex        =   114
         Text            =   "Bold "
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         TabIndex        =   113
         Text            =   "6"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text13 
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
         Left            =   4800
         TabIndex        =   112
         Text            =   "16"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text12 
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
         Left            =   4800
         TabIndex        =   111
         Text            =   "8"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text11 
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
         Left            =   4800
         TabIndex        =   110
         Text            =   "8"
         Top             =   360
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CommonDialog4 
         Left            =   6240
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         FontStrikeThru  =   -1  'True
         FontUnderLine   =   -1  'True
      End
      Begin MSComDlg.CommonDialog CommonDialog3 
         Left            =   6240
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         FontStrikeThru  =   -1  'True
         FontUnderLine   =   -1  'True
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   6240
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         FontStrikeThru  =   -1  'True
         FontUnderLine   =   -1  'True
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Color           =   8438015
         FontStrikeThru  =   -1  'True
         FontUnderLine   =   -1  'True
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Change"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Change"
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   4560
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   4080
         Shape           =   3  'Circle
         Top             =   2280
         Width           =   15
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C000&
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C000&
         Caption         =   "Font Style"
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
         Left            =   3600
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         Caption         =   "Other Cap."
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
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         Caption         =   "Title Cap."
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
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "Months/#"
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
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Main Text"
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
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton newhebSavebut 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MaskColor       =   &H00FFFFC0&
      Picture         =   "newhebcalfm.frx":1852
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton newhebPrintbut 
      BackColor       =   &H00C0C0FF&
      Caption         =   "P&rint"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "newhebcalfm.frx":1D84
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton newhebPreviewbut 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      Picture         =   "newhebcalfm.frx":1E86
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton newhebExitbut 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Picture         =   "newhebcalfm.frx":1F88
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label41 
      BackColor       =   &H0080C0FF&
      Caption         =   "Captions/Sec."
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
      Left            =   5880
      TabIndex        =   147
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Shape Shape11 
      BorderWidth     =   2
      Height          =   855
      Left            =   5400
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   3
      Height          =   1215
      Left            =   7080
      Top             =   120
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      Height          =   2655
      Left            =   360
      Top             =   4320
      Width           =   6495
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   3
      Height          =   2655
      Left            =   7080
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   3
      Height          =   1215
      Left            =   360
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "newhebcalfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check5_Click()
   'enable or deenable color fill
   If Check5.Value = vbUnchecked Then
      Option1.Enabled = False
      Option2.Enabled = False
      newhebColorbut.Enabled = False
   ElseIf Check5.Value = vbChecked Then
      Option1.Enabled = True
      Option2.Enabled = True
      newhebColorbut.Enabled = True
      End If
End Sub


Private Sub Combo11_Click()
   If Combo11.ListIndex <> prespap% - 1 Then
        If Combo11.ListIndex + 1 <> prespap% Then
           prespap% = Combo11.ListIndex + 1
           Call readfont
           paperwidth = papersize(1, prespap%)
           paperheight = papersize(2, prespap%)
           leftmargin = margins(1, prespap%)
           rightmargin = margins(2, prespap%)
           topmargin = margins(3, prespap%)
           bottommargin = margins(4, prespap%)
           If hebcal = False Then
              newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: civil calendar"
           Else
              If hebleapyear = True Then
                 newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: " + "hebrew/leapyear"
              Else
                 newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: " + "hebrew/regular year"
                 End If
              End If
           End If
        End If
End Sub


Private Sub Command1_Click()
   'set cancel to true
   CommonDialog1.CancelError = True
   On Error GoTo errhandler1
   'set the flags property
   CommonDialog1.Flags = cdlCFBoth
   'display the font dialog box
   CommonDialog1.FontName = Text3.Text
   CommonDialog1.FontSize = Text11.Text
   If InStr(Text7.Text, "Bold") <> 0 Then
      CommonDialog1.FontBold = True
   Else
      CommonDialog1.FontBold = False
      End If
   If InStr(Text7.Text, "Italic") <> 0 Then
       CommonDialog1.FontItalic = True
   Else
       CommonDialog1.FontItalic = False
       End If
   CommonDialog1.ShowFont
   'set text properties according to user's seclections
   Text3.Text = CommonDialog1.FontName
   Text11.Text = CommonDialog1.FontSize
   If CommonDialog1.FontBold = False And CommonDialog1.FontItalic = False Then
      Text7.Text = "Regular"
   ElseIf CommonDialog1.FontBold = True And CommonDialog1.FontItalic = False Then
      Text7.Text = "Bold"
   ElseIf CommonDialog1.FontBold = False And CommonDialog1.FontItalic = True Then
      Text7.Text = "Italic"
   ElseIf CommonDialog1.FontBold = True And CommonDialog1.FontItalic = True Then
      Text7.Text = "Bold Italic"
      End If
   'save changes
   'Call savefont
   Exit Sub
errhandler1:
   'user pressed cancel button
   Exit Sub
End Sub

Private Sub Command2_Click()
      'set cancel to true
   CommonDialog2.CancelError = True
   On Error GoTo errhandler2
   'set the flags property
   CommonDialog2.Flags = cdlCFBoth
   'display the font dialog box
   CommonDialog2.FontName = Text4.Text
   CommonDialog2.FontSize = Text12.Text
   If InStr(Text8.Text, "Bold") <> 0 Then
      CommonDialog2.FontBold = True
   Else
      CommonDialog2.FontBold = False
      End If
   If InStr(Text8.Text, "Italic") <> 0 Then
       CommonDialog2.FontItalic = True
   Else
       CommonDialog2.FontItalic = False
       End If
   CommonDialog2.ShowFont
   'set text properties according to user's seclections
   Text4.Text = CommonDialog2.FontName
   Text12.Text = CommonDialog2.FontSize
   If CommonDialog2.FontBold = False And CommonDialog2.FontItalic = False Then
      Text8.Text = "Regular"
   ElseIf CommonDialog2.FontBold = True And CommonDialog2.FontItalic = False Then
      Text8.Text = "Bold"
   ElseIf CommonDialog2.FontBold = False And CommonDialog2.FontItalic = True Then
      Text8.Text = "Italic"
   ElseIf CommonDialog2.FontBold = True And CommonDialog2.FontItalic = True Then
      Text8.Text = "Bold Italic"
      End If
   'Call savefont
   Exit Sub
errhandler2:
   'user pressed cancel button
   Exit Sub
End Sub

Private Sub Command3_Click()
         'set cancel to true
   CommonDialog3.CancelError = True
   On Error GoTo errhandler3
   'set the flags property
   CommonDialog3.Flags = cdlCFBoth
   'display the font dialog box
   CommonDialog3.FontName = Text5.Text
   CommonDialog3.FontSize = Text13.Text
   If InStr(Text9.Text, "Bold") <> 0 Then
      CommonDialog3.FontBold = True
   Else
      CommonDialog3.FontBold = False
      End If
   If InStr(Text9.Text, "Italic") <> 0 Then
       CommonDialog3.FontItalic = True
   Else
       CommonDialog3.FontItalic = False
       End If
   CommonDialog3.ShowFont
   'set text properties according to user's seclections
   Text5.Text = CommonDialog3.FontName
   Text13.Text = CommonDialog3.FontSize
   If CommonDialog3.FontBold = False And CommonDialog3.FontItalic = False Then
      Text9.Text = "Regular"
   ElseIf CommonDialog3.FontBold = True And CommonDialog3.FontItalic = False Then
      Text9.Text = "Bold"
   ElseIf CommonDialog3.FontBold = False And CommonDialog3.FontItalic = True Then
      Text9.Text = "Italic"
   ElseIf CommonDialog3.FontBold = True And CommonDialog3.FontItalic = True Then
      Text9.Text = "Bold Italic"
      End If
   'Call savefont
   Exit Sub
errhandler3:
   'user pressed cancel button
   Exit Sub

End Sub

Private Sub Command4_Click()
      'set cancel to true
   CommonDialog4.CancelError = True
   On Error GoTo errhandler4
   'set the flags property
   CommonDialog4.Flags = cdlCFBoth
   'display the font dialog box
   CommonDialog4.FontName = Text6.Text
   CommonDialog4.FontSize = Text14.Text
   If InStr(Text10.Text, "Bold") <> 0 Then
      CommonDialog4.FontBold = True
   Else
      CommonDialog4.FontBold = False
      End If
   If InStr(Text10.Text, "Italic") <> 0 Then
       CommonDialog4.FontItalic = True
   Else
       CommonDialog4.FontItalic = False
       End If
   CommonDialog4.ShowFont
   'set text properties according to user's seclections
   Text6.Text = CommonDialog4.FontName
   Text11.Text = CommonDialog4.FontSize
   If CommonDialog4.FontBold = False And CommonDialog4.FontItalic = False Then
      Text10.Text = "Regular"
   ElseIf CommonDialog4.FontBold = True And CommonDialog4.FontItalic = False Then
      Text10.Text = "Bold"
   ElseIf CommonDialog4.FontBold = False And CommonDialog4.FontItalic = True Then
      Text10.Text = "Italic"
   ElseIf CommonDialog4.FontBold = True And CommonDialog4.FontItalic = True Then
      Text10.Text = "Bold Italic"
      End If
   'Call savefont
   Exit Sub
errhandler4:
   'user pressed cancel button
   Exit Sub

End Sub

Private Sub Form_Load()
   'version: 04/08/2003
  
   rescal = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    newhebExitbut_Click
End Sub

Private Sub Image1_Click()
   If portrait = False Then
      response = MsgBox("Save any table configuration parameters or font changes?", vbYesNo + vbQuestion + vbDefaultButton2, "Cal Program")
      If response = vbYes Then Call savefont
      portrait = True
      First = True
      End If
   Shape6.Left = Image1.Left - 115
   Text15.Text = sEmpty
   Text15.Text = "Portrait Orientation (Vertical)" '"Print two calendars one above the other"
   portrait = True
   Call readpaper
   Call readfont
   'newhebcalfm.SSTab1.Tab = 0
   'newhebcalfm.SSTab1.TabEnabled(1) = True
   Call readfont
End Sub

Private Sub Image2_Click()
   If portrait = True Then
      response = MsgBox("Save any table configuration parameters or font changes?", vbYesNo + vbQuestion + vbDefaultButton2, "Cal Program")
      If response = vbYes Then Call savefont
      First = True
      End If
   Shape6.Left = Image2.Left - 130
   Text15.Text = sEmpty
   Text15.Text = "Landscape Orientation (Horizontal)" '"Print one calendar lengthwise"
   portrait = False
   Call readpaper
   Call readfont
   'newhebcalfm.SSTab1.Tab = 0
   'newhebcalfm.SSTab1.TabEnabled(1) = False
   Call readfont
End Sub

Private Sub newhebColorbut_Click()
   newhebcalfm.CommonDialog5.CancelError = True
   On Error GoTo can5error
   newhebcalfm.CommonDialog5.Flags = cdlCCRGBInit
   newhebcalfm.CommonDialog5.ShowColor
   fillcol = newhebcalfm.CommonDialog5.Color
   Exit Sub
can5error:
  'user pressed Cancel button
  Exit Sub
End Sub

Private Sub newhebExitbut_Click()
  Close 'Close opened files
  geo = False
  eros = False
  astronplace = False
  astronfm = False
  If automatic = True Then
     newhebout = True
     runningscan = False
     newhebcalfm.Visible = False
     newhebout = True
     Exit Sub
     End If
  'save changes if any (i.e., compare stored values with present values)
  Changefont% = False: Changepaper% = False
  Call checkchangfont(Changefont%)
  Call checkpaperfont(Changepaper%)
  If Changefont% = True Or Changepaper% = True Then
     If internet = True Then
        response = vbYes
        GoTo 50
        End If
     response = MsgBox("Save any table configuration parameters or font changes?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Cal Program")
50   If response = vbYes Then
        Call savepaper
        Call savefont
     ElseIf response = vbCancel Then
        Exit Sub
        End If
     End If
  newhebout = True
  nearcolor = False
  nfind% = 0
  nfind1% = 0
  'clear out temporary files
 If Dir(drivfordtm$ + "netz\*.*") <> sEmpty Then Kill drivfordtm$ + "netz\*.*"
 If Dir(drivfordtm$ + "skiy\*.*") <> sEmpty Then Kill drivfordtm$ + "skiy\*.*"
 If Dir(drivcities$ + "ast\netz\*.*") <> sEmpty Then Kill drivcities$ + "ast\netz\*.*"
 If Dir(drivcities$ + "ast\skiy\*.*") <> sEmpty Then Kill drivcities$ + "ast\skiy\*.*"
 If Dir(drivjk$ + "netzskiy.*") <> sEmpty Then Kill drivjk$ + "netzskiy.*"
End Sub

Private Sub newhebOpenbut_Click()
5 On Error GoTo errorhandler
  CommonDialog6.Filter = "Caption Save Files (*.sav)|*.sav|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
  CommonDialog6.FilterIndex = 1 'default filter
  CommonDialog6.ShowOpen
  filnam$ = newhebcalfm.CommonDialog6.FileName
  If filnam$ = sEmpty Then Exit Sub
  filsav% = FreeFile
  On Error GoTo openerror
  trial% = 1
  'attempt to read it once to see if it has right format
  Open filnam$ For Input As #filsav%
  Do Until EOF(filsav%)
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    Line Input #filsav%, doclin$
    trialstep% = Val(doclin$)
    Line Input #filsav%, doclin$
    trialaccur% = Val(doclin$)
  Loop
  Close #filsav%
  'now read it for real
  trial% = 2
  tabnow% = newhebcalfm.SSTab1.Tab
  Open filnam$ For Input As #filsav%
  Do Until EOF(filsav%)
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo1.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo2.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo3.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo4.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo5.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Text1.Text = doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Text2.Text = doclin$
        
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo6.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo7.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo8.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo9.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Combo10.AddItem doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Text31.Text = doclin$
    Line Input #filsav%, doclin$
    If doclin$ <> "NA" Then newhebcalfm.Text32.Text = doclin$
  Loop
 Close #filsav%
 
 'set the lattest recorded captions as the default captions to be displayed
 newhebcalfm.Combo1.ListIndex = newhebcalfm.Combo1.ListCount - 1
 newhebcalfm.Combo2.ListIndex = newhebcalfm.Combo2.ListCount - 1
 newhebcalfm.Combo3.ListIndex = newhebcalfm.Combo3.ListCount - 1
 newhebcalfm.Combo4.ListIndex = newhebcalfm.Combo4.ListCount - 1
 newhebcalfm.Combo5.ListIndex = newhebcalfm.Combo5.ListCount - 1
 newhebcalfm.Combo6.ListIndex = newhebcalfm.Combo6.ListCount - 1
 newhebcalfm.Combo7.ListIndex = newhebcalfm.Combo7.ListCount - 1
 newhebcalfm.Combo8.ListIndex = newhebcalfm.Combo8.ListCount - 1
 newhebcalfm.Combo9.ListIndex = newhebcalfm.Combo9.ListCount - 1
 newhebcalfm.Combo10.ListIndex = newhebcalfm.Combo10.ListCount - 1
 
 'If tabnow% = 0 Then
 '  newhebcalfm.SSTab1.Tab = 1
 '  newhebcalfm.SSTab1.Tab = 0
 'Else
 '  newhebcalfm.SSTab1.Tab = 0
 '  newhebcalfm.SSTab1.Tab = 1
 '  End If
    
 'newhebcalfm.Combo1.Refresh
 'newhebcalfm.Combo2.Refresh
 'newhebcalfm.Combo3.Refresh
 'newhebcalfm.Combo4.Refresh
 'newhebcalfm.Combo5.Refresh
 'newhebcalfm.Combo6.Refresh
 'newhebcalfm.Combo7.Refresh
 'newhebcalfm.Combo8.Refresh
 'newhebcalfm.Combo9.Refresh
 'newhebcalfm.Combo10.Refresh
 'newhebcalfm.Text1.Refresh
 'newhebcalfm.Text2.Refresh
 'newhebcalfm.Text31.Refresh
 'newhebcalfm.Text32.Refresh
 Exit Sub
 
errorhandler:
'user pressed cancel button
   Exit Sub
openerror:
  MsgBox "The requested file has the wrong format!", vbExclamation + vbOKOnly, "Cal Program"
  GoTo 5
End Sub

Private Sub newhebPreviewbut_Click()
'generates print preview on previewfrm.previewpicture2

If automatic And autosave Then
   If Caldirectories.chkAutoRounding.Value = vbChecked Then
      'set rounding to what was set on the Caldirectories form
      newhebcalfm.Text1 = Caldirectories.cmbRounding.List(Caldirectories.cmbRounding.ListIndex)
      newhebcalfm.Text31 = Caldirectories.cmbRounding.List(Caldirectories.cmbRounding.ListIndex)
      End If
   End If
   
Dim PrinterFlag As Boolean
PrinterFlag = False

Call PrinttoDev(previewfm.previewpicture2, PrinterFlag)

End Sub



Private Sub newhebPrintbut_Click()

'print table without having first to preview

 Call PrinttoDev(Printer, True)
 
 End Sub

Private Sub newhebSavebut_Click()
  'save the parameters for this city under its own city name
  If InStr(currentdir$, "visual_tmp") Then
     Call MsgBox("Can't save titles for files written to the ""visual_tmp"" directory" _
                 & vbCrLf & "since this is a virtual directory shared by all tables for neighborhoods." _
                 , vbExclamation, "Save Error")
     Exit Sub
     End If
     
  If portrait = True Then
    suffix$ = "_port_w1255.sav"
  ElseIf portrait = False Then
    suffix$ = "_land.sav"
    End If

  myfile = Dir(currentdir + "\" + citnam$ + suffix$)
  savv% = 0
  If myfile <> sEmpty Then
      response = MsgBox("Do you want to add these captions to those already saved in the old SAVE file? (Answer No to open dialog box in order to store captions under any file name desired.)", vbQuestion + vbYesNoCancel, "Cal Program")
      If response = vbCancel Then
         GoTo cityappend
      ElseIf response = vbNo Then
         'open OPEN dialog box
         On Error GoTo errorhandler
         CommonDialog6.Filter = "Caption Save Files (*.sav)|*.sav"
         CommonDialog6.FilterIndex = 1
         CommonDialog6.ShowSave
         filnam$ = CommonDialog6.FileName
         savv% = 1
         GoTo cityappend
         End If
      End If
  filsav% = FreeFile
  If savv% = 0 Then
     Open currentdir + "\" + citnam$ + suffix$ For Append As #filsav%
  ElseIf savv% = 1 Then
     Open filnam$ For Output As #filsav%
     End If
  If Abs(nsetflag%) = 1 Or Abs(nsetflag%) = 3 Then
     Print #filsav%, newhebcalfm.Combo1.Text  'print 5 captions
     Print #filsav%, newhebcalfm.Combo2.Text
     Print #filsav%, newhebcalfm.Combo3.Text
     Print #filsav%, newhebcalfm.Combo4.Text
     Print #filsav%, newhebcalfm.Combo5.Text
     Print #filsav%, newhebcalfm.Text1.Text   'print sec to round to
     Print #filsav%, newhebcalfm.Text2.Text   'print accur
     'Print #filsav%, "Top"
     If Abs(nsetflag%) = 1 Then
        For i% = 1 To 5
           Print #filsav%, "NA"
        Next i%
        Print #filsav%, "NA"
        Print #filsav%, "NA"
        'Print #filsav%, "NA"
        End If
     End If
   If Abs(nsetflag%) = 2 Or Abs(nsetflag%) = 3 Then
     If Abs(nsetflag%) = 2 Then
        For i% = 1 To 5
           Print #filsav%, "NA"
        Next i%
        Print #filsav%, "NA"
        Print #filsav%, "NA"
        'Print #filsav%, "NA"
        End If
     Print #filsav%, newhebcalfm.Combo6.Text
     Print #filsav%, newhebcalfm.Combo7.Text
     Print #filsav%, newhebcalfm.Combo8.Text
     Print #filsav%, newhebcalfm.Combo9.Text
     Print #filsav%, newhebcalfm.Combo10.Text
     Print #filsav%, newhebcalfm.Text31.Text   'print sec to round to
     Print #filsav%, newhebcalfm.Text32.Text   'print accur
     'Print #filsav%, "Bottom"
     End If
   'Print #filsav%, papername$(prespap%)
   'If portrait = True Then
   '   Print #filsav%, "portrait"
   'Else
   '   Print #filsav%, "landscape"
   '   End If
   Close #filsav%
cityappend:
'now append directory name to list of analyze cities if portrait file and abs(nsetflag%)=3
If Abs(nsetflag%) = 3 And nearyesval = True Then 'tested for near mountains
   filist% = FreeFile
   myfile = Dir(drivcities$ + "citynams_w1255.lst")
   If myfile = sEmpty Then
      response = MsgBox("Append this city to the list of cities whose tables are to be published?", vbQuestion + vbYesNoCancel, "Cal Program")
      If response = vbYes Then
         Open drivcities$ + "citynams_w1255.lst" For Output As #filist%
         Write #filist%, hebcityname$, currentdir, tblmesag%, s1blk, s2blk
         Close #filist%
         End If
   Else
      Open drivcities$ + "citynams_w1255.lst" For Input As #filist%
      'check if already listed
      found% = 0
      Do Until EOF(filist%)
         Line Input #filist%, docline$
         If InStr(1, docline$, currentdir) <> 0 Then
            found% = 1
            Exit Do
            End If
      Loop
      Close #filist%
      If found% = 0 Then
         response = MsgBox("Append this city to the list of cities whose tables are to be published?", vbQuestion + vbYesNoCancel, "Cal Program")
         If response = vbYes Then
            Open drivcities$ + "citynams_w1255.lst" For Append As #filist%
            Write #filist%, hebcityname$, currentdir, tblmesag%, s1blk, s2blk
            End If
         Close #filist%
         End If
      End If
   End If
   Exit Sub
errorhandler:
   GoTo cityappend
End Sub
Private Sub checkchangfont(Changefont%) 'check if changes were made to the fonts
    If hebcal = True Then
      If hebleapyear = False Then
         ext$ = ".heb"
      ElseIf hebleapyear = True Then
         ext$ = ".hly"
         End If
    ElseIf hebcal = False Then
       ext$ = ".eng"
       End If
    If portrait = True Then
       suffix$ = "Potr"
    ElseIf portrait = False Then
       suffix$ = "Lnsc"
       End If
    Prefix$ = Trim$(Mid$(papername$(prespap%), 1, 4))
    formfilname$ = drivjk$ + sEmpty + Prefix$ + suffix$ + ext$
    filfont% = FreeFile
    If Dir(formfilname$) = sEmpty Then  'save present fonts
       Call savefont
       Changefont% = False
       Exit Sub
       End If
    Open formfilname$ For Input As #filfont%
    Line Input #filfont%, doclin$
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text3.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text4.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text5.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text6.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text7.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text8.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text9.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text10.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text11.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text12.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text13.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text14.Text Then
       Changefont% = True
       GoTo 900
       End If
    Line Input #filfont%, doclin$
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text20.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text21.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text29.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text22.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text23.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text24.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text25.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text27.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text28.Text Then
       Changefont% = True
       GoTo 900
       End If
    Line Input #filfont%, doclin$
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text16.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text17.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text18.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text19.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text26.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text39.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text30.Text Then
       Changefont% = True
       GoTo 900
       End If
    Line Input #filfont%, doclin$
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text33.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text34.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text35.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text36.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text37.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text40.Text Then
       Changefont% = True
       GoTo 900
       End If
    Input #filfont%, doclin$
    If doclin$ <> newhebcalfm.Text38.Text Then
       Changefont% = True
       GoTo 900
       End If
    Line Input #filfont%, doclin$
    Input #filfont%, nfillcol
    If nfillcol <> fillcol Then
       Changefont% = True
       GoTo 900
       End If
900 Close #filfont%
End Sub
Private Sub checkpaperfont(Changepaper%)
   Dim npapername$(20)
   Dim npapersize(2, 20), nmargins(4, 20)
       filpaper% = FreeFile
       If hebcal = True Then
          If hebleapyear = False Then
             ext$ = ".heb"
          ElseIf hebleapyear = True Then
             ext$ = ".hly"
             End If
       ElseIf hebcal = False Then
          ext$ = ".eng"
          End If
       If Dir(drivjk$ + "Calpaper" + ext$) = sEmpty Then
          Call savepaper
          Changepaper% = False
          Exit Sub
          End If
       Open drivjk$ + "Calpaper" + ext$ For Input As #filpaper%
       If portrait = True Then
          paperorien% = 1
       ElseIf portrait = False Then
          paperorien% = 2
          End If
       Input #filpaper%, nprespap%, npaperorien%
       If nprespap% <> prespap% Or npaperorien% <> paperorien% Then
          Changepaper% = True
          GoTo 900
          End If
       For i% = 1 To numpaper%
          Input #filpaper%, npapername$(i%)
          If npapername$(i%) <> papername$(i%) Then
             Changepaper% = True
             GoTo 900
             End If
          Input #filpaper%, npapersize(1, i%), npapersize(2, i%)
          If npapersize(1, i%) <> papersize(1, i%) Or npapersize(2, i%) <> papersize(2, i%) Then
             Changepaper% = True
             GoTo 900
             End If
          Input #filpaper%, nmargins(1, i%), nmargins(2, i%), nmargins(3, i%), nmargins(4, i%)
          For j% = 1 To 4
             If nmargins(j%, i%) <> margins(j%, i%) Then
                Changepaper% = True
                GoTo 900
                End If
           Next j%
       Next i%
900    Close #filpaper%
End Sub
