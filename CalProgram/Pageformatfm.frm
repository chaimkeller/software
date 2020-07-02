VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Pageformatfm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Paper Format"
   ClientHeight    =   4440
   ClientLeft      =   3675
   ClientTop       =   2190
   ClientWidth     =   4710
   FillStyle       =   2  'Horizontal Line
   Icon            =   "Pageformatfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton PageFormatCancelbut 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton PageFormatOKbut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Paper Size"
      TabPicture(0)   =   "Pageformatfm.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Margins"
      TabPicture(1)   =   "Pageformatfm.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   3975
         Begin MSComCtl2.UpDown UpDown4 
            Height          =   285
            Left            =   3360
            TabIndex        =   23
            Top             =   1800
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text6"
            BuddyDispid     =   196614
            OrigLeft        =   3360
            OrigTop         =   1800
            OrigRight       =   3600
            OrigBottom      =   2055
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   22
            Text            =   "10"
            Top             =   1800
            Width           =   1335
         End
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   285
            Left            =   3360
            TabIndex        =   21
            Top             =   1320
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text5"
            BuddyDispid     =   196616
            OrigLeft        =   3360
            OrigTop         =   1320
            OrigRight       =   3600
            OrigBottom      =   1575
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   20
            Text            =   "10"
            Top             =   1320
            Width           =   1455
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   3360
            TabIndex        =   19
            Top             =   840
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text4"
            BuddyDispid     =   196618
            OrigLeft        =   3360
            OrigTop         =   840
            OrigRight       =   3600
            OrigBottom      =   1095
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   18
            Text            =   "10"
            Top             =   840
            Width           =   1335
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   3360
            TabIndex        =   17
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text3"
            BuddyDispid     =   196620
            OrigLeft        =   3240
            OrigTop         =   360
            OrigRight       =   3480
            OrigBottom      =   615
            Max             =   99999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   16
            Text            =   "10"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Bottom Margin (mm):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Top Margin (mm):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Right Margin (mm):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Left Margin (mm):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74640
         TabIndex        =   7
         Top             =   1680
         Width           =   3735
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   4
            Text            =   "250"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   3
            Text            =   "200"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Paper Height (mm):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Paper Width (mm):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   -74640
         TabIndex        =   6
         Top             =   480
         Width           =   3735
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1680
            TabIndex        =   2
            Text            =   "A4"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Paper Type:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "Pageformatfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
   If Combo1.ListIndex <> prespap% - 1 Then
        If Combo1.ListIndex + 1 <> prespap% Then
           prespap% = Combo1.ListIndex + 1
           Text1.Text = papersize(1, prespap%)
           Text2.Text = papersize(2, prespap%)
           Text3.Text = margins(1, prespap%)
           Text4.Text = margins(2, prespap%)
           Text5.Text = margins(3, prespap%)
           Text6.Text = margins(4, prespap%)
           paperwidth = Text1.Text
           paperheight = Text2.Text
           leftmargin = Text3.Text
           rightmargin = Text4.Text
           topmargin = Text5.Text
           bottommargin = Text6.Text
           Call readfont
           If hebcal = False Then
              newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: civil calendar"
           Else
              If hebleapyear = True Then
                 newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: " + "hebrew/leapyear"
              Else
                 newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: " + "hebrew/regular year"
                 End If
              End If
           magnify = True
           newhebcalfm.newhebPreviewbut.Value = True
           End If
        End If
End Sub

Private Sub Form_Load()
   'version: 04/08/2003
  
   'Text1.Text = paperwidth
   'Text2.Text = paperheight
   'Text3.Text = leftmargin
   'Text4.Text = rightmargin
   'Text5.Text = topmargin
   'Text6.Text = bottommargin
   'For i% = 1 To numpaper%
   '   Combo1.AddItem papername$(i%)
   'Next i%
   'Combo1.ListIndex = prespap% - 1
   'Pageformatfm.Visible = False
End Sub

Private Sub PageFormatCancelbut_Click()
   Pageformatfm.Visible = False
   'Unload Me
End Sub

Private Sub PageFormatOKbut_Click()
    paperwidth = Val(Text1.Text)
    paperheight = Val(Text2.Text)
    leftmargin = Val(Text3.Text)
    rightmargin = Val(Text4.Text)
    topmargin = Val(Text5.Text)
    bottommargin = Val(Text6.Text)
    prespap% = Combo1.ListIndex + 1
    margins(1, prespap%) = leftmargin
    margins(2, prespap%) = rightmargin
    margins(3, prespap%) = topmargin
    margins(4, prespap%) = bottommargin
    
    response = MsgBox("Save any margin changes?", vbYesNoCancel + vbQuestion, "Cal Program")
    If response = vbYes Then
       Call savepaper
       End If
    
   If prespap% = numpaper% Then
      If papersize(1, prespap%) <> Text1.Text Or papersize(2, prespap%) <> Text2.Text Then
        papersize(1, prespap%) = Text1.Text
        papersize(2, prespap%) = Text2.Text
        magnify = True
        newhebcalfm.newhebPreviewbut.Value = True
        Exit Sub
        End If
      End If
      
   Pageformatfm.Visible = False
   'Unload Me
End Sub
