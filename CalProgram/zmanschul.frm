VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form zmanschul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Times for the shaliach zibur before netz"
   ClientHeight    =   7350
   ClientLeft      =   3390
   ClientTop       =   1080
   ClientWidth     =   6045
   Icon            =   "zmanschul.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbLanguage 
      Height          =   315
      Left            =   4600
      TabIndex        =   80
      Text            =   "cmbLanguage"
      ToolTipText     =   "Choose language of template/table"
      Top             =   5320
      Width           =   1000
   End
   Begin VB.Frame frmDST 
      Caption         =   "DST"
      Height          =   495
      Left            =   400
      TabIndex        =   78
      Top             =   5160
      Width           =   975
      Begin VB.CheckBox chkDST 
         Caption         =   "DST"
         Height          =   195
         Left            =   200
         TabIndex        =   79
         ToolTipText     =   "Check to add hour when DST"
         Top             =   210
         Width           =   735
      End
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   77
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton Command4 
      Height          =   435
      Left            =   3780
      Picture         =   "zmanschul.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "print out the z'manim"
      Top             =   6600
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      Height          =   435
      Left            =   3240
      Picture         =   "zmanschul.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "store the z'manim to disk"
      Top             =   6600
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Height          =   435
      Left            =   2320
      Picture         =   "zmanschul.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "save parameters aa a  template"
      Top             =   6600
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Left            =   1680
      Picture         =   "zmanschul.frx":1558
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "load stored template"
      Top             =   6600
      Width           =   675
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   15
      TabIndex        =   12
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Shul Times for Netz Minyan"
      TabPicture(0)   =   "zmanschul.frx":1BC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmShulNetz"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Minchah/Maariv Times"
      TabPicture(1)   =   "zmanschul.frx":1BDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmMM"
      Tab(1).ControlCount=   1
      Begin VB.Frame frmMM 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   55
         Top             =   360
         Width           =   5775
         Begin MSComCtl2.UpDown UpDown17 
            Height          =   375
            Left            =   5040
            TabIndex        =   71
            Top             =   2520
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            OrigLeft        =   5040
            OrigTop         =   2520
            OrigRight       =   5280
            OrigBottom      =   2895
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMaShab 
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
            Left            =   4560
            TabIndex        =   70
            Text            =   "0"
            Top             =   2520
            Width           =   495
         End
         Begin MSComCtl2.UpDown UpDown16 
            Height          =   375
            Left            =   4200
            TabIndex        =   69
            Top             =   2520
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            OrigLeft        =   4200
            OrigTop         =   2520
            OrigRight       =   4440
            OrigBottom      =   2895
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMaWeek 
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
            Left            =   3720
            TabIndex        =   68
            Text            =   "0"
            Top             =   2520
            Width           =   495
         End
         Begin MSComCtl2.UpDown UpDown15 
            Height          =   375
            Left            =   5040
            TabIndex        =   67
            Top             =   2040
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            OrigLeft        =   5040
            OrigTop         =   2040
            OrigRight       =   5280
            OrigBottom      =   2415
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMiShab 
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
            Left            =   4560
            TabIndex        =   66
            Text            =   "0"
            Top             =   2040
            Width           =   495
         End
         Begin MSComCtl2.UpDown UpDown14 
            Height          =   375
            Left            =   4200
            TabIndex        =   65
            Top             =   2040
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            OrigLeft        =   4200
            OrigTop         =   2040
            OrigRight       =   4440
            OrigBottom      =   2415
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMiWeed 
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
            Left            =   3720
            TabIndex        =   64
            Text            =   "0"
            Top             =   2040
            Width           =   495
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Check8"
            Height          =   195
            Left            =   240
            TabIndex        =   63
            Top             =   2640
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Check7"
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   2160
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox Text22 
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
            Left            =   840
            TabIndex        =   61
            Text            =   "מעריב"
            Top             =   2520
            Width           =   2535
         End
         Begin VB.TextBox txtMinchah 
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
            Left            =   840
            TabIndex        =   60
            Text            =   "מנחה"
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Frame frmSunset 
            Caption         =   "Reference Sunset"
            Height          =   1095
            Left            =   240
            TabIndex        =   56
            Top             =   240
            Width           =   5295
            Begin VB.OptionButton optAstron 
               Caption         =   "Astronomical Sunset"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1800
               TabIndex        =   59
               Top             =   720
               Width           =   1935
            End
            Begin VB.OptionButton optMishor 
               Caption         =   "Mishor Sunset"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1800
               TabIndex        =   58
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton optVisible 
               Caption         =   "Visible Sunset"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1800
               TabIndex        =   57
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "non-shabbos"
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
            Left            =   4560
            TabIndex        =   76
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "shabbos"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   3720
            TabIndex        =   75
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "minutes before sunrise"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   3600
            TabIndex        =   74
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Hebrew name of the zman as it will appear in the listing"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   73
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "activate"
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
            TabIndex        =   72
            Top             =   1800
            Width           =   555
         End
      End
      Begin VB.Frame frmShulNetz 
         Height          =   3735
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   5775
         Begin VB.TextBox Text18 
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
            Index           =   0
            Left            =   4680
            TabIndex        =   53
            Text            =   "0"
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox Text6 
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
            Index           =   0
            Left            =   3600
            TabIndex        =   51
            Text            =   "0"
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox Text17 
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
            Index           =   0
            Left            =   4680
            TabIndex        =   49
            Text            =   "0"
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox Text5 
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
            Index           =   0
            Left            =   3600
            TabIndex        =   47
            Text            =   "0"
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox Text16 
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
            Index           =   0
            Left            =   4680
            TabIndex        =   45
            Text            =   "0"
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox Text4 
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
            Index           =   0
            Left            =   3600
            TabIndex        =   43
            Text            =   "0"
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox Text15 
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
            Index           =   0
            Left            =   4680
            TabIndex        =   41
            Text            =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox Text3 
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
            Index           =   0
            Left            =   3600
            TabIndex        =   39
            Text            =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "David"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   38
            Text            =   "צור ישראל"
            Top             =   3120
            Width           =   2295
         End
         Begin VB.CheckBox Check6 
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   37
            Top             =   3120
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "David"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   36
            Text            =   "אמת"
            Top             =   2640
            Width           =   2295
         End
         Begin VB.CheckBox Check5 
            Height          =   435
            Index           =   0
            Left            =   360
            TabIndex        =   35
            Top             =   2640
            Value           =   1  'Checked
            Width           =   195
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "David"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   960
            TabIndex        =   34
            Text            =   "שמע"
            Top             =   2160
            Width           =   2295
         End
         Begin VB.CheckBox Check4 
            Height          =   435
            Index           =   0
            Left            =   360
            TabIndex        =   33
            Top             =   2160
            Value           =   1  'Checked
            Width           =   315
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "David"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   32
            Text            =   "ישתבח/שוכן עד"
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CheckBox Check3 
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   31
            Top             =   1680
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox Text14 
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
            Index           =   0
            Left            =   4680
            TabIndex        =   29
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox Text2 
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
            Index           =   0
            Left            =   3600
            TabIndex        =   27
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "David"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   26
            Text            =   "ברוך שאמר"
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox Check2 
            Height          =   435
            Index           =   0
            Left            =   360
            TabIndex        =   25
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox Text13 
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
            Index           =   0
            Left            =   4680
            TabIndex        =   23
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text1 
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
            Index           =   0
            Left            =   3600
            TabIndex        =   21
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "David"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   20
            Text            =   "מזמור שיר"
            Top             =   720
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   315
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   22
            Top             =   720
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text1(0)"
            BuddyDispid     =   196657
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   720
            OrigRight       =   4320
            OrigBottom      =   1095
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown7 
            Height          =   360
            Index           =   0
            Left            =   5160
            TabIndex        =   24
            Top             =   720
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   635
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text13(0)"
            BuddyDispid     =   196656
            BuddyIndex      =   0
            OrigLeft        =   5160
            OrigTop         =   720
            OrigRight       =   5400
            OrigBottom      =   1125
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   28
            Top             =   1200
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text2(0)"
            BuddyDispid     =   196653
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   1200
            OrigRight       =   4320
            OrigBottom      =   1575
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown8 
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   30
            Top             =   1200
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text14(0)"
            BuddyDispid     =   196652
            BuddyIndex      =   0
            OrigLeft        =   5160
            OrigTop         =   1200
            OrigRight       =   5400
            OrigBottom      =   1575
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   40
            Top             =   1680
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text3(0)"
            BuddyDispid     =   196643
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   1680
            OrigRight       =   4320
            OrigBottom      =   2055
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown9 
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   42
            Top             =   1680
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text15(0)"
            BuddyDispid     =   196642
            BuddyIndex      =   0
            OrigLeft        =   5160
            OrigTop         =   1680
            OrigRight       =   5400
            OrigBottom      =   2055
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown4 
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   44
            Top             =   2160
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text4(0)"
            BuddyDispid     =   196641
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   2160
            OrigRight       =   4320
            OrigBottom      =   2535
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown10 
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   46
            Top             =   2160
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text16(0)"
            BuddyDispid     =   196640
            BuddyIndex      =   0
            OrigLeft        =   5160
            OrigTop         =   2160
            OrigRight       =   5400
            OrigBottom      =   2535
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown5 
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   48
            Top             =   2640
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text5(0)"
            BuddyDispid     =   196639
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   2640
            OrigRight       =   4320
            OrigBottom      =   3015
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown11 
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   50
            Top             =   2640
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text17(0)"
            BuddyDispid     =   196638
            BuddyIndex      =   0
            OrigLeft        =   5160
            OrigTop         =   2640
            OrigRight       =   5400
            OrigBottom      =   3015
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown6 
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   52
            Top             =   3120
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text6(0)"
            BuddyDispid     =   196637
            BuddyIndex      =   0
            OrigLeft        =   4080
            OrigTop         =   3120
            OrigRight       =   4320
            OrigBottom      =   3495
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown12 
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   54
            Top             =   3120
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text18(0)"
            BuddyDispid     =   196636
            BuddyIndex      =   0
            OrigLeft        =   5160
            OrigTop         =   3120
            OrigRight       =   5400
            OrigBottom      =   3495
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "weekday"
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
            Left            =   4680
            TabIndex        =   18
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Shabbos"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   3600
            TabIndex        =   17
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "minutes before sunrise"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   3600
            TabIndex        =   16
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Hebrew name of the zman as it will appear in the listing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1040
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "activate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   555
         End
      End
   End
   Begin VB.Frame frmParshiot 
      Caption         =   "Parshot HaShavua"
      Height          =   535
      Left            =   380
      TabIndex        =   8
      Top             =   5760
      Width           =   5235
      Begin VB.OptionButton optNoParshiot 
         Caption         =   "No Sedra"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optEYParshiot 
         Caption         =   "Sedra of Eretz Yisroel"
         Height          =   195
         Left            =   1480
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optDiasporaParshiot 
         Caption         =   "Sedra of Diaspora"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   220
         Width           =   1695
      End
   End
   Begin MSComCtl2.UpDown UpDown13 
      Height          =   375
      Left            =   3280
      TabIndex        =   6
      Top             =   5280
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   5
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text19"
      BuddyDispid     =   196669
      OrigLeft        =   3780
      OrigTop         =   4140
      OrigRight       =   4020
      OrigBottom      =   4515
      Max             =   60
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text19 
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
      Left            =   2800
      TabIndex        =   5
      Text            =   "5"
      Top             =   5280
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Caption         =   "seconds"
      Height          =   195
      Left            =   3620
      TabIndex        =   7
      Top             =   5400
      Width           =   675
   End
   Begin VB.Label Label6 
      Caption         =   "Round times to "
      Height          =   255
      Left            =   1600
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
End
Attribute VB_Name = "zmanschul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkDST_Click()
   If chkDST.Value = vbChecked Then
      DSTcheck = True
   Else
      DSTcheck = False
      End If
End Sub

Private Sub cmbLanguage_click()
   Select Case cmbLanguage.ListIndex
      Case 0
         Label1.Caption = "Hebrew name of the zman as it will appear in the listing"
      Case 1
         Label1.Caption = "English name of the zman as it will appear in the listing"
   End Select
End Sub

Private Sub Command1_Click()
5:  On Error GoTo c3error
10:     CommonDialog1.CancelError = True
        CommonDialog1.Filter = "schul zmanim template files (*.sct)|*.sct|"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.FileName = drivjk$ + "*.sct"
        CommonDialog1.ShowOpen
        filnam$ = CommonDialog1.FileName
        
        If InStr(filnam$, "_Eng.sct") Then
           cmbLanguage.ListIndex = 1
        Else
           cmbLanguage.ListIndex = 0
           End If
           
        tmpnum% = FreeFile
        Open filnam$ For Input As #tmpnum%
        
        On Error GoTo errhand
        'Input #tmpnum%, Check1.Value, text7(0).Text, text1(0).Text, text13(0).Text
        Input #tmpnum%, at%, bt$, ct$, dt$
           If at% = 1 Then
              Check1(0).Value = vbChecked
           Else
              Check1(0).Value = vbUnchecked
              End If
           Text7(0).Text = bt$
           Text1(0).Text = ct$
           Text13(0).Text = dt$
'        Input #tmpnum%, check2(0).Value, text8(0).Text, text2(0).Text, text14(0).Text
        Input #tmpnum%, at%, bt$, ct$, dt$
           If at% = 1 Then
              Check2(0).Value = vbChecked
           Else
              Check2(0).Value = vbUnchecked
              End If
           Text8(0).Text = bt$
           Text2(0).Text = ct$
           Text14(0).Text = dt$
'        Input #tmpnum%, check3(0).Value, text9(0).Text, text3(0).Text, text15(0).Text
        Input #tmpnum%, at%, bt$, ct$, dt$
           If at% = 1 Then
              Check3(0).Value = vbChecked
           Else
              Check3(0).Value = vbUnchecked
              End If
           Text9(0).Text = bt$
           Text3(0).Text = ct$
           Text15(0).Text = dt$
'        Input #tmpnum%, check4(0).Value, text10(0).Text, text4(0).Text, text16(0).Text
        Input #tmpnum%, at%, bt$, ct$, dt$
           If at% = 1 Then
              Check4(0).Value = vbChecked
           Else
              Check4(0).Value = vbUnchecked
              End If
           Text10(0).Text = bt$
           Text4(0).Text = ct$
           Text16(0).Text = dt$
'        Input #tmpnum%, check5(0).Value, text11(0).Text, text5(0).Text, text17(0).Text
        Input #tmpnum%, at%, bt$, ct$, dt$
           If at% = 1 Then
              Check5(0).Value = vbChecked
           Else
              Check5(0).Value = vbUnchecked
              End If
           Text11(0).Text = bt$
           Text5(0).Text = ct$
           Text17(0).Text = dt$
'        Input #tmpnum%, check6(0).Value, text12(0).Text, text6(0).Text, text18(0).Text
        Input #tmpnum%, at%, bt$, ct$, dt$
           If at% = 1 Then
              Check6(0).Value = vbChecked
           Else
              Check6(0).Value = vbUnchecked
              End If
           Text12(0).Text = bt$
           Text6(0).Text = ct$
           Text18(0).Text = dt$
'        Input #tmpnum%, Text19.Text
        Input #tmpnum%, bt$
        Text19.Text = bt$
        Line Input #tmpnum%, txtTitlet$
        If txtTitlet$ = sEmpty Then txtTitlet$ = "לוח זמנים לש""ץ ל" & hebcityname$ & " לשנת " & yrcal$
        txtTitle = txtTitlet$
        parshiotinfo% = 0
        Input #tmpnum%, parshiotinfo% 'parshios info
        Select Case parshiotinfo%
           Case 0
              optNoParshiot = True
           Case 1
              optEYParshiot = True
           Case 2
              optDiasporaParshiot = True
        End Select

        Close #tmpnum%
           
c3error:
       Exit Sub
       
errhand:
      If Err.Number = 62 Then Resume Next
      response = MsgBox("Can't read the template file!  " & Str$(Err.Number) & "Do you want to try again?", vbExclamation + vby, "Cal Program")
      If response = vbYes Then GoTo 5
      
End Sub

Private Sub Command2_Click()
  On Error GoTo c3error
10:     CommonDialog1.CancelError = True
        CommonDialog1.Filter = "schul zmanim template files (*.sct)|*.sct|"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.FileName = drivjk$ + "*.sct"
        CommonDialog1.ShowSave
        filnam$ = CommonDialog1.FileName
        
        If cmbLanguage.ListIndex = 1 Then 'chose English template and table
           If InStr(filnam$, "_Eng.sct") Then
           Else
              'add suffix to name
              Dim pos%
              pos% = InStr(filnam$, ".")
              filnam$ = Mid$(filnam$, 1, pos% - 1) & "_Eng.sct"
              End If
           
           End If
           
        myfile = Dir(filnam$)
        If myfile <> sEmpty Then
           response = MsgBox("Overwrite the existing file with this name?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Cal Program")
           If response <> vbYes Then
              GoTo 10
              End If
           End If
           
        tmpnum% = FreeFile
        Open filnam$ For Output As #tmpnum%
        Write #tmpnum%, Check1(0).Value, Text7(0).Text, Text1(0).Text, Text13(0).Text
        Write #tmpnum%, Check2(0).Value, Text8(0).Text, Text2(0).Text, Text14(0).Text
        Write #tmpnum%, Check3(0).Value, Text9(0).Text, Text3(0).Text, Text15(0).Text
        Write #tmpnum%, Check4(0).Value, Text10(0).Text, Text4(0).Text, Text16(0).Text
        Write #tmpnum%, Check5(0).Value, Text11(0).Text, Text5(0).Text, Text17(0).Text
        Write #tmpnum%, Check6(0).Value, Text12(0).Text, Text6(0).Text, Text18(0).Text
        Write #tmpnum%, Text19.Text
        Print #tmpnum%, Trim$(txtTitle.Text)
        If Not parshiotEY And Not parshiotdiaspora Then
           Write #tmpnum%, 0
        ElseIf parshiotEY Then
           Write #tmpnum%, 1
        ElseIf parshiotdiaspora Then
           Write #tmpnum%, 2
           End If
        
        Close #tmpnum%
           
c3error:
End Sub

Private Sub Command3_Click()
   Dim ExcelApp As Excel.Application
   Dim ExcelBook As Excel.Workbook
   Dim ExcelSheet As Excel.Worksheet
   
   'variables for determining the onset and end of daylight saving time
   Dim stryrDST%, endyrDST%, strdaynum1%, enddaynum1%, strdaynum2%, enddaynum2%
   Dim MonStart%, MonEnd%, yl As Integer
   Dim DSTadd As Integer, DSThour As Integer
   
   Dim MarchDate As Integer
   Dim OctoberDate As Integer
   Dim NovemberDate As Integer
   Dim YearLength As Integer
   
   Dim RemovedUnderlined As Boolean

   Dim ier As Integer

  On Error GoTo c3error
  
  RemoveUnderline = True
  
  If hebcal Then fshabos% = fshabos0%
  lenyr1% = yrend%(0)
  
  If DSTcheck Then
  
     'open DST_EY.txt file and determine the daynumber of the beginning and end of DST in EY
     stryrDST% = yrheb% + RefCivilYear% - RefHebYear% '(yrheb% - 5758) + 1997
     endyrDST% = yrheb% + RefCivilYear% - RefHebYear% + 1 '(yrheb% - 5758) + 1998
     
     'find beginning and ending day numbers for each civil year
     Select Case eroscountry$
     
        Case "Israel", "" 'EY eros or cities using 2017 DST rules
        
            MarchDate = (31 - (Fix(stryrDST% * 5 / 4) + 4) Mod 7) - 2 'starts on Friday = 2 days before EU start on Sunday
            OctoberDate = (31 - (Fix(stryrDST% * 5 / 4) + 1) Mod 7)
            YearLength% = DaysinYear(stryrDST%)
            strdaynum1% = DayNumber(YearLength%, 3, MarchDate)
            enddaynum1% = DayNumber(YearLength%, 10, OctoberDate)
            
            MarchDate = (31 - (Fix(endyrDST% * 5 / 4) + 4) Mod 7) - 2 'starts on Friday = 2 days before EU start on Sunday
            OctoberDate = (31 - (Fix(endyrDST% * 5 / 4) + 1) Mod 7)
            YearLength% = DaysinYear(endyrDST%)
            strdaynum2% = DayNumber(YearLength%, 3, MarchDate)
            enddaynum2% = DayNumber(YearLength%, 10, OctoberDate)
        
        Case "USA" 'English {USA DST rules}
        
            MarchDate = 14 - (Fix(1 + stryrDST% * 5 / 4) Mod 7)
            NovemberDate = 7 - (Fix(1 + stryrDST% * 5 / 4) Mod 7)
            YearLength% = DaysinYear(stryrDST%)
            strdaynum1% = DayNumber(YearLength%, 3, MarchDate)
            enddaynum1% = DayNumber(YearLength%, 11, NovemberDate)
            
            MarchDate = 14 - (Fix(1 + endyrDST% * 5 / 4) Mod 7)
            NovemberDate = 7 - (Fix(1 + endyrDST% * 5 / 4) Mod 7)
            YearLength% = DaysinYear(endyrDST%)
            strdaynum2% = DayNumber(YearLength%, 3, MarchDate)
            enddaynum2% = DayNumber(YearLength%, 11, NovemberDate)
            
        Case Else 'not implemented yet for other countries
        
     End Select
     
'     MonStart% = 3
'     MonEnd% = 10
'
'     ier = DST_begend(stryrDST%, endyrDST%, strdaynum1%, enddaynum1%, strdaynum2%, enddaynum2%)
'
'     If ier < 0 Then
'        Select Case MsgBox("Can't find the beginning or ending civil dates for DST in this hebrew year." _
'                           & vbCrLf & "" _
'                           & vbCrLf & "Do you want to calculate the table without DST?" _
'                           , vbYesNoCancel Or vbExclamation Or vbDefaultButton1, "DST error")
'
'            Case vbYes
'
'            Case vbNo
'               DSTcheck = False
'               zmanschul.chkDST.Value = vbUnchecked
'
'            Case vbCancel
'               DSTcheck = False
'               zmanschul.chkDST.Value = vbUnchecked
'
'               Exit Sub
'
'
'        End Select
'
'     Else
'        'calculate the daynumbers
'        yl = DaysinYear(stryrDST%)
'        strdaynum1% = DayNumber(yl, MonStart%, strdaynum1%)
'        enddaynum1% = DayNumber(yl, MonEnd%, enddaynum1%)
'
'        yl = DaysinYear(endyrDST%)
'        strdaynum2% = DayNumber(yl, MonStart%, strdaynum2%)
'        enddaynum2% = DayNumber(yl, MonEnd%, enddaynum2%)
'
'        End If

     End If
  
10:     CommonDialog1.CancelError = True
        CommonDialog1.Filter = "schul davening times output (*.sch)|*.sch|schul davening times as excel file (*xls)|*.xls|"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.FileName = drivjk$ + "*.sch"
        CommonDialog1.ShowSave
        filnam$ = CommonDialog1.FileName
        pos% = InStr(filnam$, ".")
        ext$ = Mid(filnam$, pos% + 1, 3)
        myfile = Dir(filnam$)
        If myfile <> sEmpty Then
           response = MsgBox("Overwrite the existing file with this name?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Cal Program")
           If response <> vbYes Then
              GoTo 10
              End If
           End If
        
        Screen.MousePointer = vbHourglass
        plusround% = 1
        setflag% = 0
        sunrise% = 0
        steps = Val(Text19.Text)
        RemoveUnderline = True
        
        If ext$ = "sch" Then
            
            schulnum% = FreeFile
            Open filnam$ For Output As #schulnum%
            
    '       generate header
            outdoc$ = sEmpty
            If Check1(0).Value = vbChecked Then
               newdoc$ = Text7(0).Text
               For n% = 1 To Len(newdoc$) 'get rid of blanks
                  If Mid$(newdoc$, n%, 1) = Chr$(32) Then
                     Mid$(newdoc$, n%, 1) = "_"
                     End If
               Next n%
               outdoc$ = newdoc$
               End If
            If Check2(0).Value = vbChecked Then
               newdoc$ = Text8(0).Text
               For n% = 1 To Len(newdoc$)
                  If Mid$(newdoc$, n%, 1) = Chr$(32) Then
                     Mid$(newdoc$, n%, 1) = "_"
                     End If
               Next n%
               Select Case cmbLanguage.ListIndex
                  Case 0 'Hebrew rtl
                     outdoc$ = newdoc$ + "   " + outdoc$
                  Case 1 'English ltr
                     outdoc$ = outdoc$ + "   " + newdoc$
               End Select
'               outdoc$ = newdoc$ + "   " + outdoc$
               End If
            If Check3(0).Value = vbChecked Then
               newdoc$ = Text9(0).Text
               For n% = 1 To Len(newdoc$)
                  If Mid$(newdoc$, n%, 1) = Chr$(32) Then
                     Mid$(newdoc$, n%, 1) = "_"
                     End If
               Next n%
               Select Case cmbLanguage.ListIndex
                  Case 0 'Hebrew rtl
                     outdoc$ = newdoc$ + "   " + outdoc$
                  Case 1 'English ltr
                     outdoc$ = outdoc$ + "   " + newdoc$
               End Select
'               outdoc$ = newdoc$ + "   " + outdoc$
               End If
            If Check4(0).Value = vbChecked Then
               newdoc$ = Text10(0).Text
               For n% = 1 To Len(newdoc$)
                  If Mid$(newdoc$, n%, 1) = Chr$(32) Then
                     Mid$(newdoc$, n%, 1) = "_"
                     End If
               Next n%
               Select Case cmbLanguage.ListIndex
                  Case 0 'Hebrew rtl
                     outdoc$ = newdoc$ + "   " + outdoc$
                  Case 1 'English ltr
                     outdoc$ = outdoc$ + "   " + newdoc$
               End Select
'               outdoc$ = newdoc$ + "   " + outdoc$
               End If
            If Check5(0).Value = vbChecked Then
               newdoc$ = Text11(0).Text
               For n% = 1 To Len(newdoc$)
                  If Mid$(newdoc$, n%, 1) = Chr$(32) Then
                     Mid$(newdoc$, n%, 1) = "_"
                     End If
               Next n%
               Select Case cmbLanguage.ListIndex
                  Case 0 'Hebrew rtl
                     outdoc$ = newdoc$ + "   " + outdoc$
                  Case 1 'English ltr
                     outdoc$ = outdoc$ + "   " + newdoc$
               End Select
'               outdoc$ = newdoc$ + "   " + outdoc$
               End If
            If Check6(0).Value = vbChecked Then
               newdoc$ = Text12(0).Text
               For n% = 1 To Len(newdoc$)
                  If Mid$(newdoc$, n%, 1) = Chr$(32) Then
                     Mid$(newdoc$, n%, 1) = "_"
                     End If
               Next n%
               Select Case cmbLanguage.ListIndex
                  Case 0 'Hebrew rtl
                     outdoc$ = newdoc$ + "   " + outdoc$
                  Case 1 'English ltr
                     outdoc$ = outdoc$ + "   " + newdoc$
               End Select
'               outdoc$ = newdoc$ + "   " + outdoc$
               End If
               
            Select Case cmbLanguage.ListIndex
               Case 0 'Hebrew rtl
                  outdoc$ = "זריחה" + "   " + outdoc$
               Case 1 'English ltr
                  outdoc$ = outdoc$ + "   " + "sunrise"
            End Select
'            outdoc$ = "זריחה" + "   " + outdoc$
    
            Select Case cmbLanguage.ListIndex
               Case 0 'Hebrew rtl
                 outdoc$ = outdoc$ + "   " + "תאריך_לועזי" + "   " + "יום" + "   " + "תאריך_עברי"
               Case 1 'English ltr
                 outdoc$ = "Hebrew_date" + "   " + "day" + "   " + "civil_date" + "   " + outdoc$
            End Select
'            outdoc$ = outdoc$ + "   " + "תאריך_לועזי" + "   " + "יום" + "   " + "תאריך_עברי"
            Print #schulnum%, outdoc$
    
            numday% = -1
            For isc% = 1 To endyr%
               If mmdate%(2, isc%) > mmdate%(1, isc%) Then
                  ks% = 0
                  For j% = mmdate%(1, isc%) To mmdate%(2, isc%)
                      numday% = numday% + 1
                      ks% = ks% + 1
                      outdoc$ = sEmpty
                      netz = Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 1, 1)) + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 3, 2)) / 60 + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 6, 2)) / 3600
                      If Check1(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text1(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text13(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         outdoc$ = t3subb$
                         End If
                      If Check2(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text2(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text14(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check3(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text3(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text15(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check4(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text4(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text16(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check5(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text5(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text17(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check6(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text6(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text18(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      outdoc$ = Trim$(stortim$(0, isc% - 1, ks% - 1)) + "    " + outdoc$ + "   " + Trim$(stortim$(2, isc% - 1, ks% - 1)) + "   " + Trim$(stortim$(4, isc% - 1, ks% - 1)) + "       " + Trim$(stortim$(3, isc% - 1, ks% - 1))
                      If parshiotEY Or parshiotdiaspora Then
                        Call InsertHolidays(calday$, isc%, ks%)
                        outdoc$ = Trim$(stortim$(0, isc% - 1, ks% - 1)) + "    " + outdoc$ + "   " + Trim$(stortim$(2, isc% - 1, ks% - 1)) + "   " + calday$ + "       " + Trim$(stortim$(3, isc% - 1, ks% - 1))
                        End If
                      Print #schulnum%, outdoc$
                  Next j%
               ElseIf mmdate%(2, isc%) < mmdate%(1, isc%) Then
                  ks% = 0
                  For j% = mmdate%(1, isc%) To yrend%(0)
                      numday% = numday% + 1
                      ks% = ks% + 1
                      outdoc$ = sEmpty
                      netz = Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 1, 1)) + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 3, 2)) / 60 + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 6, 2)) / 3600
                      If Check1(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text1(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text13(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         outdoc$ = t3subb$
                         End If
                      If Check2(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text2(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text14(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check3(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text3(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text15(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check4(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text4(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text16(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check5(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text5(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text17(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check6(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text6(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text18(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      outdoc$ = Trim$(stortim$(0, isc% - 1, ks% - 1)) + "    " + outdoc$ + "   " + Trim$(stortim$(2, isc% - 1, ks% - 1)) + "   " + Trim$(stortim$(4, isc% - 1, ks% - 1)) + "       " + Trim$(stortim$(3, isc% - 1, ks% - 1))
                      If parshiotEY Or parshiotdiaspora Then
                        Call InsertHolidays(calday$, isc%, ks%)
                        outdoc$ = Trim$(stortim$(0, isc% - 1, ks% - 1)) + "    " + outdoc$ + "   " + Trim$(stortim$(2, isc% - 1, ks% - 1)) + "   " + calday$ + "       " + Trim$(stortim$(3, isc% - 1, ks% - 1))
                        End If
                      Print #schulnum%, outdoc$
                  Next j%
                  yrn% = yrn% + 1
                  For j% = 1 To mmdate%(2, isc%)
                      ks% = ks% + 1
                      numday% = numday% + 1
                      outdoc$ = sEmpty
                      netz = Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 1, 1)) + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 3, 2)) / 60 + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 6, 2)) / 3600
                      If Check1(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text1(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text13(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         outdoc$ = t3subb$
                         End If
                      If Check2(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text2(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text14(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check3(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text3(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text15(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check4(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text4(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text16(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check5(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text5(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text17(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      If Check6(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text6(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text18(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         Select Case cmbLanguage.ListIndex
                            Case 0 'Hebrew rtl
                               outdoc$ = t3subb$ + "   " + outdoc$
                            Case 1 'English ltr
                               outdoc$ = outdoc$ + "   " + t3subb$
                         End Select
'                         outdoc$ = t3subb$ + "   " + outdoc$
                         End If
                      outdoc$ = Trim$(stortim$(0, isc% - 1, ks% - 1)) + "    " + outdoc$ + "   " + Trim$(stortim$(2, isc% - 1, ks% - 1)) + "   " + Trim$(stortim$(4, isc% - 1, ks% - 1)) + "       " + Trim$(stortim$(3, isc% - 1, ks% - 1))
                      If parshiotEY Or parshiotdiaspora Then
                        Call InsertHolidays(calday$, isc%, ks%)
                        outdoc$ = Trim$(stortim$(0, isc% - 1, ks% - 1)) + "    " + outdoc$ + "   " + Trim$(stortim$(2, isc% - 1, ks% - 1)) + "   " + calday$ + "       " + Trim$(stortim$(3, isc% - 1, ks% - 1))
                        End If
                      Print #schulnum%, outdoc$
                  Next j%
                  End If
            Next isc%
            Close #schulnum%
        ElseIf ext$ = "xls" Then
            'count number of checked boxes
            numchecked% = 0
            If Check1(0).Value = vbChecked Then
               numchecked% = numchecked% + 1
               End If
            If Check2(0).Value = vbChecked Then
               numchecked% = numchecked% + 1
               End If
            If Check3(0).Value = vbChecked Then
               numchecked% = numchecked% + 1
               End If
            If Check4(0).Value = vbChecked Then
               numchecked% = numchecked% + 1
               End If
            If Check5(0).Value = vbChecked Then
               numchecked% = numchecked% + 1
               End If
            If Check6(0).Value = vbChecked Then
               numchecked% = numchecked% + 1
               End If
            numcol% = numchecked% + 5
            
            'calculate format letters
            begLet$ = "A"
            Select Case numchecked%
               Case 1
                  endLet$ = "F"
               Case 2
                  endLet$ = "G"
               Case 3
                  endLet$ = "H"
               Case 4
                  endLet$ = "I"
               Case 5
                  endLet$ = "J"
               Case 6
                  endLet$ = "K"
               Case 7
                  endLet$ = "L"
            End Select
            
            Screen.MousePointer = vbDefault
            
            If txtTitle = sEmpty Then
               Select Case cmbLanguage.ListIndex
                  Case 0 'Hebrew
                    sDefault$ = "לוח זמנים לש""ץ ל" & hebcityname$ & " לשנת " & yrcal$
                  Case 1 'English
                    sDefault$ = "Zemanim Table for the Prayer Leader in " & citnamp$ & " for the Hebrew year " & yrcal$
               End Select
            Else
               sDefault$ = Trim$(txtTitle)
               End If
            response = InputBox("Title of the schul excel file is: ", "Shul Times Table's Title", sDefault$)
            Screen.MousePointer = vbHourglass
            
            'Set ExcelSheet = CreateObject("Excel.Sheet")
            'ExcelSheet.Application.Visible = True
            Set ExcelApp = New Excel.Application
            Set ExcelBook = ExcelApp.Workbooks.Add
            Set ExcelSheet = ExcelBook.Worksheets.Add
            
            ExcelSheet.PageSetup.leftmargin = 15
            ExcelSheet.PageSetup.rightmargin = 5
            ExcelSheet.PageSetup.topmargin = 10
            ExcelSheet.PageSetup.bottommargin = 5
            ExcelSheet.PageSetup.HeaderMargin = 0
            ExcelSheet.PageSetup.FooterMargin = 0
            ExcelSheet.PageSetup.Orientation = xlLandscape
            
            'Set ExcelSheet = CreateObject("Excel.Sheet")
            Screen.MousePointer = vbDefault
            ExcelBook.Application.Visible = True
            ExcelBook.Windows(1).Visible = True
            
            ExcelSheet.Rows.RowHeight = 16.5
            
            rowhgt% = ExcelSheet.Rows.RowHeight 'default row height
            
            'format columns and rows
            Dim ColFor%(10, 3)
            ColFor%(0, 0) = 16 'column width
            ColFor%(0, 1) = 10 'row height
            ColFor%(0, 2) = 13 'font size
            ColFor%(1, 0) = 3  'column width
            ColFor%(1, 1) = 10 'row height
            ColFor%(1, 2) = 12 'font size
            ColFor%(2, 0) = 12 'column width
            ColFor%(2, 1) = 10 'row height
            ColFor%(2, 2) = 10 'font size
            ColFor%(3, 0) = 12 'column width
            ColFor%(3, 1) = 12 'row height
            ColFor%(3, 2) = 12 'font size
            ColFor%(4, 0) = 26 'column width
            ColFor%(4, 1) = 10 'row height
            ColFor%(4, 2) = 12 'font size
            With ExcelSheet
               For icol% = 1 To numcol% - 4 'zemanim columns
                  .Columns(icol%).ColumnWidth = ColFor%(0, 0)
                  .Columns(icol%).HorizontalAlignment = xlCenter
                  .Columns(icol%).VerticalAlignment = xlCenter
                  .Columns(icol%).Font.Size = ColFor%(0, 2)
                  .Columns(icol%).Font.Name = "Arial"
                  
               Next icol%
               For icol% = 3 To 0 Step -1 'date and spacer columns
                  .Columns(numcol% - icol%).ColumnWidth = ColFor%(4 - icol%, 0)
                  .Columns(numcol% - icol%).HorizontalAlignment = xlCenter
                  .Columns(numcol% - icol%).VerticalAlignment = xlCenter
                  .Columns(numcol% - icol%).Font.Size = ColFor%(4 - icol%, 2)
                  .Columns(numcol% - icol%).Font.Name = "Arial"
               Next icol%
            End With
            
            numday% = 0 '(if set numday% = 1 then have added a line to the first page for uniformity)
            numday% = numday% + 1
            
            'now write title
            ExcelSheet.Cells(numday%, Int(numcol% / 2 + 0.5)).Value = response 'the title
            'now meege the columns on this row
            ExcelSheet.Range(ExcelSheet.Cells(numday%, 1), ExcelSheet.Cells(numday%, numcol%)).Merge
            
            numday% = numday% + 2 'skip one line
            ColNum% = -1
            If Check1(0).Value = vbChecked Then
               newdoc$ = Text7(0).Text
               ColNum% = ColNum% + 1
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
               GoSub box
               End If
            If Check2(0).Value = vbChecked Then
               newdoc$ = Text8(0).Text
               ColNum% = ColNum% + 1
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
               GoSub box
               End If
            If Check3(0).Value = vbChecked Then
               newdoc$ = Text9(0).Text
               ColNum% = ColNum% + 1
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
               GoSub box
               End If
            If Check4(0).Value = vbChecked Then
               newdoc$ = Text10(0).Text
               ColNum% = ColNum% + 1
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
               GoSub box
               End If
            If Check5(0).Value = vbChecked Then
               newdoc$ = Text11(0).Text
               ColNum% = ColNum% + 1
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
               GoSub box
               End If
            If Check6(0).Value = vbChecked Then
               newdoc$ = Text12(0).Text
               ColNum% = ColNum% + 1
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
               ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
               GoSub box
               End If
               
            Select Case cmbLanguage.ListIndex
              Case 0 'Hebrew
               
                ColNum% = ColNum% + 1
                ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = "זריחה"
                ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                GoSub box
                ExcelSheet.Cells(numday%, numcol% - 3).Value = sEmpty
                ExcelSheet.Cells(numday%, numcol% - 3).Font.Bold = True
                ColNumber% = numcol% - 3
                GoSub box2
                ExcelSheet.Cells(numday%, numcol% - 2).Value = "תאריך לועזי"
                ExcelSheet.Cells(numday%, numcol% - 2).Font.Bold = True
                ColNumber% = numcol% - 2
                GoSub box2
                ExcelSheet.Cells(numday%, numcol% - 1).Value = "תאריך עברי"
                ExcelSheet.Cells(numday%, numcol% - 1).Font.Bold = True
                ColNumber% = numcol% - 1
                GoSub box2
                ExcelSheet.Cells(numday%, numcol%).Value = "יום"
                ExcelSheet.Cells(numday%, numcol%).Font.Bold = True
                ColNumber% = numcol%
                GoSub box2
                
              Case 1 'English
              
                ColNum% = ColNum% + 1
                ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = "Sunrise"
                ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                GoSub box
                ExcelSheet.Cells(numday%, numcol% - 3).Value = sEmpty
                ExcelSheet.Cells(numday%, numcol% - 3).Font.Bold = True
                ColNumber% = numcol% - 3
                GoSub box2
                ExcelSheet.Cells(numday%, numcol% - 2).Value = "Civil Date"
                ExcelSheet.Cells(numday%, numcol% - 2).Font.Bold = True
                ColNumber% = numcol% - 2
                GoSub box2
                ExcelSheet.Cells(numday%, numcol% - 1).Value = "Hebrew Date"
                ExcelSheet.Cells(numday%, numcol% - 1).Font.Bold = True
                ColNumber% = numcol% - 1
                GoSub box2
                ExcelSheet.Cells(numday%, numcol%).Value = "Day"
                ExcelSheet.Cells(numday%, numcol%).Font.Bold = True
                ColNumber% = numcol%
                GoSub box2
              
            End Select
            
            numday% = numday% + 1 'skip another line but make it smaller in size
            
            nshabos% = 0: newschulyr% = 0: dayweek% = rhday% - 1: addmon% = 0
            ntotshabos% = 0
            
            For isc% = 1 To endyr%
               
               'If isc% = 1 Then rowhgt% = ExcelSheet.Rows.RowHeight
               ExcelSheet.Rows(numday%).RowHeight = rowhgt% - 8
            
               If mmdate%(2, isc%) > mmdate%(1, isc%) Then 'Hebrew months that are within the first of the secular years that the Hebrew year spans
               
                  ks% = 0
                  For j% = mmdate%(1, isc%) To mmdate%(2, isc%)
                      numday% = numday% + 1
                      ks% = ks% + 1
                      outdoc$ = sEmpty
                      
                      'check for Shabbosim
                      changit% = 0
                      If newhebcalfm.Check4.Value = vbChecked And fshabos% + nshabos% * 7 = j% Then 'this is shabbos
                         nshabos% = nshabos% + 1 '<<<2--->>>
                         ntotshabos% = ntotshabos% + 1
                         changit% = 1
                         'check for end of year
                         If fshabos% + nshabos% * 7 > lenyr1% Then
                            newschulyr% = 1
                            fshabos% = 7 - (lenyr1% - (fshabos% + (nshabos% - 1) * 7))
                            nshabos% = 0
                            End If
                       End If
                      
                      RemovedUnderlined = False
                      If InStr(stortim$(0, isc% - 1, ks% - 1), "_") Or changit% = 1 Then
                         changit% = 0
                         'highlight the entire row
                         For icol% = 1 To numcol%
                            ExcelSheet.Cells(numday%, icol%).Interior.Color = RGB(216, 214, 214) 'grey highlighting
                         Next icol%
                         RemovedUnderlined = True
                         End If
                      
                      'remove underlines for shabbosim
                      stortim$(0, isc% - 1, ks% - 1) = Replace(stortim$(0, isc% - 1, ks% - 1), "_", " ")
                      
                      netz = Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 1, 1)) + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 3, 2)) / 60 + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 6, 2)) / 3600
                      
                      If DSTcheck Then
                         'add hour for DST
                         If j% >= strdaynum1% And j% < enddaynum1% Then
                            DSTadd = 1
                            netz = netz + 1
                         Else
                            DSTadd = 0
                            End If
                         End If
                         
                      ColNum% = -1
                      If Check1(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text1(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text13(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check2(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text2(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text14(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check3(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text3(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text15(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check4(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text4(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text16(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check5(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text5(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text17(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check6(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text6(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text18(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      ColNum% = ColNum% + 1
                      
                      'now print the netz time
                      t3sub = netz
                      GoSub 1500
                      GoSub round
                      ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                      
'                      If DSTadd = 1 Then
'                         'add hour for DST
'                         t3sub = Trim$(stortim$(0, isc% - 1, ks% - 1))
'                         DSThour = Mid(t3sub, 1, 1)
'                         Mid(t3sub, 1, 1) = Trim$(Str$(Val(DSThour) + 1))
'                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3sub
'                      Else
'                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = Trim$(stortim$(0, isc% - 1, ks% - 1))
'                         End If
                         
                      'GoSub underline
                      
                      'reformat civil date:
                      CivilDate$ = Replace(Trim$(stortim$(2, isc% - 1, ks% - 1)), "-", "/", 1, 2) 'replace the "-" with "/"
                      If parshiotEY Then 'convert from month in English to numbers
                         CivilDate$ = Format(CivilDate$, "dd.mm.yyyy")
                         End If
                      'ExcelSheet.Cells(numday%, numcol% - 2).Value = Trim$(stortim$(2, isc% - 1, ks% - 1)) 'civil date
                      ExcelSheet.Cells(numday%, numcol% - 2).Value = CivilDate$ 'civil date
                      
                      'ColNumber% = numcol% - 2
                      'GoSub underline2
                      ExcelSheet.Cells(numday%, numcol% - 1).Value = Trim$(stortim$(3, isc% - 1, ks% - 1)) 'Hebrew Calendar date
                      'ColNumber% = numcol% - 1
                      'GoSub underline2
                     If parshiotEY Or parshiotdiaspora Then
                         Call InsertHolidays(calday$, isc%, ks%)
                         'remove double shabbosim
                         If InStr(calday$, "שבת שבת") Then
                            calday$ = Replace(calday$, "שבת שבת", "שבת")
                            End If
                         ExcelSheet.Cells(numday%, numcol%).Value = calday$
                         'ColNumber% = numcol%
                         'GoSub underline2
                      Else
                         ExcelSheet.Cells(numday%, numcol%).Value = Trim$(stortim$(4, isc% - 1, ks% - 1))
                         'ColNumber% = numcol%
                         'GoSub underline2
                         End If
                      GoSub underline3
                      
                      'restore flag for shabbos
                      If RemovedUnderlined Then
                        stortim$(0, isc% - 1, ks% - 1) = stortim$(0, isc% - 1, ks% - 1) + "_"
                        RemovedUnderlined = False
                        End If
                      
                  Next j%
               ElseIf mmdate%(2, isc%) < mmdate%(1, isc%) Then  'do the last few days of the first secular year
                  ks% = 0
                  For j% = mmdate%(1, isc%) To yrend%(0)
                      numday% = numday% + 1
                      ks% = ks% + 1
                      outdoc$ = sEmpty
                      
                      'check for Shabbosim
                      changit% = 0
                      If newhebcalfm.Check4.Value = vbChecked And fshabos% + nshabos% * 7 = j% Then 'this is shabbos
                         nshabos% = nshabos% + 1
                         ntotshabos% = ntotshabos% + 1
                         changit% = 1
                         'check for end of year
                         If fshabos% + nshabos% * 7 > lenyr1% Then
                            newschulyr% = 1
                            fshabos% = 7 - (lenyr1% - (fshabos% + (nshabos% - 1) * 7))
                            nshabos% = 0
                            End If
                         End If
                      
                      RemovedUnderlined = False
                      If InStr(stortim$(0, isc% - 1, ks% - 1), "_") Or changit% = 1 Then
                         changit% = 0
                         'highlight the entire row
                         For icol% = 1 To numcol%
                            ExcelSheet.Cells(numday%, icol%).Interior.Color = RGB(216, 214, 214) 'grey highlighting
                         Next icol%
                         RemovedUnderlined = True
                         End If
                         
                      'remove underlines for shabbosim
                      stortim$(0, isc% - 1, ks% - 1) = Replace(stortim$(0, isc% - 1, ks% - 1), "_", " ")
                      
                      netz = Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 1, 1)) + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 3, 2)) / 60 + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 6, 2)) / 3600
                      
                      If DSTcheck Then
                         'add hour for DST
                         If j% >= strdaynum1% And j% < enddaynum1% Then
                            DSTadd = 1
                            netz = netz + 1
                         Else
                            DSTadd = 0
                            End If
                         End If
                      
                      ColNum% = -1
                      If Check1(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text1(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text13(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check2(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text2(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text14(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check3(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text3(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text15(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check4(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text4(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text16(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check5(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text5(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text17(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check6(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text6(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text18(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      ColNum% = ColNum% + 1
                      
                      'now print the netz time
                      t3sub = netz
                      GoSub 1500
                      GoSub round
                      ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                      
'                      If DSTadd = 1 Then
'                         'add hour for DST
'                         t3sub = Trim$(stortim$(0, isc% - 1, ks% - 1))
'                         DSThour = Mid(t3sub, 1, 1)
'                         Mid(t3sub, 1, 1) = Trim$(Str$(Val(DSThour) + 1))
'                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3sub
'                      Else
'                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = Trim$(stortim$(0, isc% - 1, ks% - 1))
'                         End If
                         
                      'GoSub underline
                      ExcelSheet.Cells(numday%, numcol% - 2).Value = Trim$(stortim$(2, isc% - 1, ks% - 1))
                      ColNumber% = numcol% - 2
                      'GoSub underline2
                      ExcelSheet.Cells(numday%, numcol% - 1).Value = Trim$(stortim$(3, isc% - 1, ks% - 1))
                      ColNumber% = numcol% - 1
                      'GoSub underline2
                      If parshiotEY Or parshiotdiaspora Then
                         Call InsertHolidays(calday$, isc%, ks%)
                         If InStr(calday$, "שבת שבת") Then
                            calday$ = Replace(calday$, "שבת שבת", "שבת")
                            End If
                         ExcelSheet.Cells(numday%, numcol%).Value = calday$
                         'ColNumber% = numcol%
                         'GoSub underline2
                      Else
                         ExcelSheet.Cells(numday%, numcol%).Value = Trim$(stortim$(4, isc% - 1, ks% - 1))
                         'ColNumber% = numcol%
                         'GoSub underline2
                         End If
                      GoSub underline3
                      
                      'restore flag for shabbos
                      If RemovedUnderlined Then
                        stortim$(0, isc% - 1, ks% - 1) = stortim$(0, isc% - 1, ks% - 1) + "_"
                        RemovedUnderlined = False
                        End If
                      
                  Next j%
                  yrn% = yrn% + 1
                  For j% = 1 To mmdate%(2, isc%) 'Hebrew months in the second secular year that the Hebrew year spans
                      ks% = ks% + 1
                      numday% = numday% + 1
                      outdoc$ = sEmpty
                      
                      'check for Shabbosim
                      If newhebcalfm.Check4.Value = vbChecked And fshabos% + nshabos% * 7 = j% Then 'this is shabbos
                         nshabos% = nshabos% + 1  '<<<--->>>
                         ntotshabos% = ntotshabos% + 1
                         changit% = 1
                         
                         'check for end of year
                         If fshabos% + nshabos% * 7 > lenyr1% Then
                            newschulyr% = 1
                            fshabos% = 7 - (lenyr1% - (fshabos% + (nshabos% - 1) * 7))
                            nshabos% = 0
                            End If
  
                         End If
                      
                      RemovedUnderlined = False
                      If InStr(stortim$(0, isc% - 1, ks% - 1), "_") Or changit% = 1 Then
                         changit% = 0
                         'highlight the entire row
                         For icol% = 1 To numcol%
                            ExcelSheet.Cells(numday%, icol%).Interior.Color = RGB(216, 214, 214) 'grey highlighting
                         Next icol%
                         RemovedUnderlined = True
                         End If
                      
                      'remove underlines for shabbosim
                      stortim$(0, isc% - 1, ks% - 1) = Replace(stortim$(0, isc% - 1, ks% - 1), "_", " ")
                      
                      netz = Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 1, 1)) + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 3, 2)) / 60 + Val(Mid$(stortim$(0, isc% - 1, ks% - 1), 6, 2)) / 3600
                      
                      If DSTcheck Then
                         'add hour for DST
                         If j% >= strdaynum2% And j% < enddaynum2% Then
                            DSTadd = 1
                            netz = netz + 1
                         Else
                            DSTadd = 0
                            End If
                         End If
                      
                      ColNum% = -1
                      If Check1(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text1(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text13(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check2(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text2(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text14(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check3(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text3(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text15(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check4(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text4(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text16(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check5(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text5(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text17(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      If Check6(0).Value = vbChecked Then
                         If Trim$(stortim$(4, isc% - 1, ks% - 1)) = heb4$(7) Then
                            t3sub = netz - Val(Text6(0).Text) / 60
                         Else
                            t3sub = netz - Val(Text18(0).Text) / 60
                            End If
                         GoSub 1500
                         GoSub round
                         ColNum% = ColNum% + 1
                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                         'GoSub underline
                         End If
                      ColNum% = ColNum% + 1
                      
                      'now print the netz time
                      t3sub = netz
                      GoSub 1500
                      GoSub round
                      ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3subb$
                      
'                      If DSTadd = 1 Then
'                         'add hour for DST in netz entry
'                         t3sub = Trim$(stortim$(0, isc% - 1, ks% - 1))
'                         DSThour = Mid(t3sub, 1, 1)
'                         Mid(t3sub, 1, 1) = Trim$(Str$(Val(DSThour) + 1))
'                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = t3sub
'                      Else
'                         ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = Trim$(stortim$(0, isc% - 1, ks% - 1))
'                         End If
                      
                      'GoSub underline
                      ExcelSheet.Cells(numday%, numcol% - 2).Value = Trim$(stortim$(2, isc% - 1, ks% - 1))
                      ColNumber% = numcol% - 2
                      'GoSub underline2
                      ExcelSheet.Cells(numday%, numcol% - 1).Value = Trim$(stortim$(3, isc% - 1, ks% - 1))
                      'ColNumber% = numcol% - 1
                      'GoSub underline2
                      If parshiotEY Or parshiotdiaspora Then
                         If InStr(calday$, "שבת שבת") Then
                            calday$ = Replace(calday$, "שבת שבת", "שבת")
                            End If
                         Call InsertHolidays(calday$, isc%, ks%)
                         ExcelSheet.Cells(numday%, numcol%).Value = calday$
                         'ColNumber% = numcol%
                         'GoSub underline2
                     Else
                         ExcelSheet.Cells(numday%, numcol%).Value = Trim$(stortim$(4, isc% - 1, ks% - 1))
                         'ColNumber% = numcol%
                         'GoSub underline2
                         End If
                     GoSub underline3
                      
                      'restore flag for shabbos
                      If RemovedUnderlined Then
                        stortim$(0, isc% - 1, ks% - 1) = stortim$(0, isc% - 1, ks% - 1) + "_"
                        RemovedUnderlined = False
                        End If
                        
                  Next j%
                  End If
                  
                'write headers after each month
                If isc% = endyr% Then GoTo 1200 'finished already, so don't write headers
                numday% = numday% + 1 'skip one line
                'now write title
                ExcelSheet.Cells(numday%, Int(numcol% / 2 + 0.5)).Value = response 'the title
                numday% = numday% + 2 'skip another line
                ColNum% = -1
                If Check1(0).Value = vbChecked Then
                   newdoc$ = Text7(0).Text
                   ColNum% = ColNum% + 1
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                   GoSub box
                   End If
                If Check2(0).Value = vbChecked Then
                   newdoc$ = Text8(0).Text
                   ColNum% = ColNum% + 1
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                   GoSub box
                   End If
                If Check3(0).Value = vbChecked Then
                   newdoc$ = Text9(0).Text
                   ColNum% = ColNum% + 1
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                   GoSub box
                   End If
                If Check4(0).Value = vbChecked Then
                   newdoc$ = Text10(0).Text
                   ColNum% = ColNum% + 1
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                   GoSub box
                   End If
                If Check5(0).Value = vbChecked Then
                   newdoc$ = Text11(0).Text
                   ColNum% = ColNum% + 1
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                   GoSub box
                   End If
                If Check6(0).Value = vbChecked Then
                   newdoc$ = Text12(0).Text
                   ColNum% = ColNum% + 1
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = newdoc$
                   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                   GoSub box
                   End If
                   
                Select Case cmbLanguage.ListIndex
                  Case 0 'Hebrew
                   
                    ColNum% = ColNum% + 1
                    ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = "זריחה"
                    ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                    GoSub box
                    ExcelSheet.Cells(numday%, numcol% - 3).Value = sEmpty
                    ExcelSheet.Cells(numday%, numcol% - 3).Font.Bold = True
                    ColNumber% = numcol% - 3
                    GoSub box2
                    ExcelSheet.Cells(numday%, numcol% - 2).Value = "תאריך לועזי"
                    ExcelSheet.Cells(numday%, numcol% - 2).Font.Bold = True
                    ColNumber% = numcol% - 2
                    GoSub box2
                    ExcelSheet.Cells(numday%, numcol% - 1).Value = "תאריך עברי"
                    ExcelSheet.Cells(numday%, numcol% - 1).Font.Bold = True
                    ColNumber% = numcol% - 1
                    GoSub box2
                    ExcelSheet.Cells(numday%, numcol%).Value = "יום"
                    ExcelSheet.Cells(numday%, numcol%).Font.Bold = True
                    ColNumber% = numcol%
                    GoSub box2
                    
                  Case 1 'English
                  
                    ColNum% = ColNum% + 1
                    ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = "Sunrise"
                    ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
                    GoSub box
                    ExcelSheet.Cells(numday%, numcol% - 3).Value = sEmpty
                    ExcelSheet.Cells(numday%, numcol% - 3).Font.Bold = True
                    ColNumber% = numcol% - 3
                    GoSub box2
                    ExcelSheet.Cells(numday%, numcol% - 2).Value = "Civil Date"
                    ExcelSheet.Cells(numday%, numcol% - 2).Font.Bold = True
                    ColNumber% = numcol% - 2
                    GoSub box2
                    ExcelSheet.Cells(numday%, numcol% - 1).Value = "Hebrew Date"
                    ExcelSheet.Cells(numday%, numcol% - 1).Font.Bold = True
                    ColNumber% = numcol% - 1
                    GoSub box2
                    ExcelSheet.Cells(numday%, numcol%).Value = "Day"
                    ExcelSheet.Cells(numday%, numcol%).Font.Bold = True
                    ColNumber% = numcol%
                    GoSub box2
                  
                End Select
                
'                ColNum% = ColNum% + 1
'                ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Value = "זריחה"
'                ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Font.Bold = True
'                GoSub box
'                ExcelSheet.Cells(numday%, numcol% - 3).Value = sEmpty
'                ExcelSheet.Cells(numday%, numcol% - 3).Font.Bold = True
'                ColNumber% = numcol% - 3
'                GoSub box2
'                ExcelSheet.Cells(numday%, numcol% - 2).Value = "תאריך לועזי"
'                ExcelSheet.Cells(numday%, numcol% - 2).Font.Bold = True
'                ColNumber% = numcol% - 2
'                GoSub box2
'                ExcelSheet.Cells(numday%, numcol% - 1).Value = "תאריך עברי"
'                ExcelSheet.Cells(numday%, numcol% - 1).Font.Bold = True
'                ColNumber% = numcol% - 1
'                GoSub box2
'                ExcelSheet.Cells(numday%, numcol%).Value = "יום"
'                ExcelSheet.Cells(numday%, numcol%).Font.Bold = True
'                ColNumber% = numcol%
'                GoSub box2
                
                
                numday% = numday% + 1 'skip another line
1200:
               'add page break
               If isc% <> endyr% Then
                  ExcelSheet.Cells(numday% - 3, 1).PageBreak = True
               ElseIf isc% = endyr% Then
                  ExcelSheet.Cells(numday% + 2, 1).PageBreak = True
                  End If
               
            Next isc%
            
            'excelsheet.Cells.Borderarounds
            
            'final page break
            'ExcelSheet.Rows(numday%).PageBreak

            ExcelSheet.SaveAs filnam$
            Screen.MousePointer = vbDefault
            zmanschul.SetFocus
            Select Case MsgBox("Do you want to close the EXCEL window?" _
                               & vbCrLf & "If you answer ""No"", then EXCEL will continue running, even after" _
                               & vbCrLf & "closing the Cal Program" _
                               , vbYesNo Or vbQuestion Or vbDefaultButton2, App.title)
            
               Case vbYes
                  
                  ExcelApp.Quit
                   
                  Set ExcelApp = Nothing
                  Set ExcelBook = Nothing
                  Set ExcelSheet = Nothing
            
               Case vbNo
            
            End Select
            
            End If
        
        Screen.MousePointer = vbDefault

c3error:
        Screen.MousePointer = vbDefault
        
        If Err.Number = 1004 Or Err.Number = 0 Then Exit Sub
        
        response = MsgBox("Error Num: " & Err.Number & vbCrLf & Err.Description, vbCritical & vbOKOnly, App.title)
        Exit Sub
        
        
1500 t3hr = Fix(t3sub): t3min = Fix((t3sub - t3hr) * 60)
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
                 'Mid$(t3subb$, 1 + sp%, 1) = Trim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1))
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
           If Mid$(t3subb$, 1, 2) = "24" Then Mid$(t3subb$, 1, 2) = "00"
Return

underline:  'underline border
   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Borders(xlEdgeBottom).Color = RGB(0, 0, 0) 'dark
   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Borders(xlEdgeLeft).Color = RGB(255, 255, 255) 'white
   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).Borders(xlEdgeRight).Color = RGB(255, 255, 255)
Return

underline2:  'underline border with column # = ColNumber%
   ExcelSheet.Cells(numday%, ColNumber%).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   ExcelSheet.Cells(numday%, ColNumber%).Borders(xlEdgeBottom).Color = RGB(0, 0, 0) 'dark
   ExcelSheet.Cells(numday%, ColNumber%).Borders(xlEdgeLeft).Color = RGB(255, 255, 255) 'white
   ExcelSheet.Cells(numday%, ColNumber%).Borders(xlEdgeRight).Color = RGB(255, 255, 255)
Return

underline3: 'underline one line at a time
   begFormat$ = begLet$ & LTrim$(Str$(numday%))
   endFormat$ = endLet$ & LTrim$(Str$(numday%))
   ExcelSheet.Range(begFormat$, endFormat$).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic 'can be xlThick, xlMedium, xlThin, xlHairline
   ExcelSheet.Range(begFormat$, endFormat$).Borders(xlEdgeBottom).Color = RGB(0, 0, 0) 'dark
   ExcelSheet.Range(begFormat$, endFormat$).Borders(xlEdgeLeft).Color = RGB(255, 255, 255) 'white
   ExcelSheet.Range(begFormat$, endFormat$).Borders(xlEdgeRight).Color = RGB(255, 255, 255) 'white
Return
            
box: 'box with column number = numcol - 4 - ColNum 'used to be xlThick, can also be xlHairline, xlThin
   ExcelSheet.Cells(numday%, numcol% - 4 - ColNum%).BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
Return
        
box2: 'box with column number = ColNumber%
   ExcelSheet.Cells(numday%, ColNumber%).BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
Return
        
      
End Sub


Private Sub Form_Load()
   If parshiotEY Then
      optEYParshiot.Value = True
   ElseIf parshiotdiaspora Then
      optDiasporaParshiot.Value = True
      End If
      
   SSTab1.Tab = 0 'show netz shul times as default
      
   'enable sunset reference types
   If vis Then
      optVisible.Enabled = True
      optVisible.Value = True
      End If
   If mis Then
      optMishor.Enabled = True
      optMishor.Value = True
      End If
   If ast Then
      optAstron.Enabled = True
      optAstron.Value = True
      End If
      
   With cmbLanguage
      .AddItem "Hebrew"
      .AddItem "English"
      .ListIndex = 0
   End With
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'version: 04/08/2003
  
   Unload Me
   Set zmanschul = Nothing
End Sub

Private Sub optDiasporaParshiot_Click()
   parshiotEY = False
   parshiotdiaspora = True
End Sub

Private Sub optEYParshiot_Click()
   parshiotEY = True
   parshiotdiaspora = False
End Sub

Private Sub optNoParshiot_Click()
   parshiotEY = False
   parshiotdiaspora = False
End Sub

