VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mapgraphfm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plot of horizon profile"
   ClientHeight    =   6750
   ClientLeft      =   3855
   ClientTop       =   1815
   ClientWidth     =   8130
   Icon            =   "mapgraphfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8130
   Visible         =   0   'False
   Begin VB.Frame frmObstructions 
      Caption         =   "obstructions"
      Height          =   615
      Left            =   6250
      TabIndex        =   38
      Top             =   5550
      Width           =   1560
      Begin VB.CheckBox chkObstruction 
         Caption         =   "Activate"
         Height          =   195
         Left            =   1200
         TabIndex        =   41
         ToolTipText     =   "Check to activate skipping when obstructions are numerous"
         Top             =   280
         Width           =   255
      End
      Begin MSComCtl2.UpDown UpDownObst 
         Height          =   285
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtObstructions"
         BuddyDispid     =   196611
         OrigLeft        =   840
         OrigTop         =   240
         OrigRight       =   1095
         OrigBottom      =   495
         Max             =   100
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtObstructions 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Text            =   "50"
         ToolTipText     =   "Percent of obstructions to reject"
         Top             =   240
         Width           =   480
      End
   End
   Begin MSComCtl2.UpDown updnDelay 
      Height          =   285
      Left            =   1560
      TabIndex        =   37
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtDelay"
      BuddyDispid     =   196612
      OrigLeft        =   1560
      OrigTop         =   5760
      OrigRight       =   1800
      OrigBottom      =   6015
      Max             =   60
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   36
      Text            =   "1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   195
      Left            =   7200
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   495
      Left            =   4380
      TabIndex        =   31
      ToolTipText     =   "Set obstruction distance up to which will be drawn in yellow"
      Top             =   5640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   873
      _Version        =   393216
      Value           =   5
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text3"
      BuddyDispid     =   196614
      OrigLeft        =   4380
      OrigTop         =   5640
      OrigRight       =   4620
      OrigBottom      =   6135
      Max             =   90
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   4080
      TabIndex        =   30
      Text            =   "5"
      Top             =   5640
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3600
      Picture         =   "mapgraphfm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "View sunrise/sunset horizon from -35 to 35 degrees"
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2880
      Picture         =   "mapgraphfm.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "start 3D Viewer"
      Top             =   5640
      Width           =   675
   End
   Begin VB.CommandButton restorelimitsbut 
      Height          =   495
      Left            =   4680
      Picture         =   "mapgraphfm.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Restore the limits"
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton TimeZonebut 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      Picture         =   "mapgraphfm.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Lookup the time zone"
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   60
      ScaleHeight     =   4815
      ScaleWidth      =   8055
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2640
         TabIndex        =   32
         Top             =   3900
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   4080
         TabIndex        =   22
         Top             =   4200
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   150
            Width           =   1815
         End
         Begin VB.CheckBox dirsavecheck 
            Height          =   300
            Left            =   120
            Picture         =   "mapgraphfm.frx":1816
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Save to cities directory"
            Top             =   160
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "Direc. Name"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   195
            Width           =   975
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   4320
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   390
         Picture         =   "mapgraphfm.frx":1918
         ScaleHeight     =   3255
         ScaleWidth      =   7395
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   7395
      End
      Begin VB.Label Label13 
         Caption         =   "Always ask if wan't to change Version Number"
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   3900
         Width           =   3855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Hebrew Name of Place"
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
         Left            =   600
         TabIndex        =   21
         Top             =   4320
         Width           =   1575
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   405
      Left            =   4920
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   714
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text1"
      BuddyDispid     =   196629
      OrigLeft        =   4440
      OrigTop         =   4920
      OrigRight       =   4680
      OrigBottom      =   5295
      Max             =   12
      Min             =   -12
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Text            =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6375
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
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
   Begin VB.CommandButton Calendarbut 
      Height          =   495
      Left            =   5640
      Picture         =   "mapgraphfm.frx":4F5B6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Make calendar"
      Top             =   5640
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   4920
      Width           =   5620
      _ExtentX        =   9922
      _ExtentY        =   873
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   12648447
      GridColor       =   8388608
      ScrollBars      =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^azimuth     |^view angle     |^longitude     |^latitude       |^distance     |^height        "
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Picture         =   "mapgraphfm.frx":4F740
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save the profile file"
      Top             =   5640
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   720
      ScaleHeight     =   3735
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      Begin VB.Shape Shape1 
         Height          =   210
         Left            =   1920
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "azi = -39.1, va = 1.23"
         Height          =   210
         Left            =   1920
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Shape shpDelay 
      Height          =   495
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDelay 
      Caption         =   "Auto Mode Delay Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   35
      Top             =   5740
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Zone Time"
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
      Left            =   3240
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "ymax ="
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
      Left            =   4560
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "ymin ="
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
      Left            =   2520
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   7560
      TabIndex        =   7
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label5 
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
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "80"
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
      Left            =   6960
      TabIndex        =   5
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "-80"
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
      Left            =   600
      TabIndex        =   4
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "View Angle"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Azimuth (degrees)"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   4080
      Width           =   5175
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   7320
      X2              =   7320
      Y1              =   3840
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   720
      X2              =   7320
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   720
      X2              =   720
      Y1              =   120
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   720
      X2              =   7320
      Y1              =   3840
      Y2              =   3840
   End
End
Attribute VB_Name = "mapgraphfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim azi As Single, va As Single, logi As Single, lati As Single
Dim dista As Single, hgti As Single, xcord As Single, ycord As Single
Dim xo As Single, yo As Single, calpress%, lResult As Long, lResult2 As Long
Dim erosfile$(1), kmxeros As Double, kmyeros As Double, hgteros As Single
Dim ae As Single, be As Single, ce As Single, de As Single, fe As Single
Dim dirfile$(1), myfile, lenfil%, tmpfil$, icity%, ipr%, iicity$, iipr$, batname$
Dim hebcityname$(1000), tdcities%(1000), batfile%, xminnew As Single, yminnew As Single
Dim xmaxnew As Single, ymaxnew As Single, nnpnt%, Mode%, checkver As Boolean
Dim xmino, ymino, xmaxo, ymaxo, cityfound%
Dim kmyeroscheck As Double, kmxeroscheck As Double

Private Sub Calendarbut_Click()
   On Error GoTo errhand
   
   If AutoScanlist Then mapgraphfm.txtDelay = Str$(IntOld2%)
   
   If world = True Then
     If sunmode% >= 1 Then
        Mode% = 1
     ElseIf sunmode% <= 0 Then
        Mode% = 0
        End If
   Else 'sunrise/sunset determined by Analyze's global variable setflag%
     If setflag% = 1 Then
        Mode% = 0
     ElseIf setflag% = 0 Then
        Mode% = 1
        End If
   End If
   If dirsavecheck.value = vbChecked And Combo2.Text = sEmpty And calpress% = 1 Then
      Beep
      mapgraphfm.StatusBar1.Panels(1) = "Enter a ENGLISH name for the city directory!"
      calpress% = 1
      Exit Sub
      End If
   calpress% = calpress% + 1
   If calpress% > 2 Then calpress% = 1
   If calpress% = 1 Then
      Screen.MousePointer = vbHourglass
      'gather information on all the cities(\eros) subdirectories if they exist
      ' Display the names of directories.

      If world = True Then
         mypath = drivcities$ + "eros\"  ' Set the path.
      Else
         GoTo l50 'this is output from Analyze, all directories
         'and their Hebrew names are stored in drivcities$ + "citynams.txt"
         'mypath = drivcities$  ' Set the path.
      End If
      myname = Dir(mypath, vbDirectory)   ' Retrieve the first entry.
      Do While myname <> sEmpty   ' Start the loop.
         ' Ignore the current directory and the encompassing directory.
         If myname <> "." And myname <> ".." And myname <> "netz" And myname <> "skiy" Then
            ' Use bitwise comparison to make sure MyName is a directory.
            If (GetAttr(mypath & myname) And vbDirectory) = vbDirectory Then
                Combo2.AddItem myname  ' Display entry only if it has a bat file
                End If                 ' it represents a directory.
            End If
         myname = Dir    ' Get next entry.
      Loop
      If Combo2.ListCount = 0 Then 'myname = sEmpty Then
         If Mode% = 1 Then
            myname = "visual_tmp"
            mynew = mypath + myname + "\NETZ\visu.bat"
            GoTo 160
         ElseIf Mode% = 0 Then
            myname = "visual_tmp"
            mynew = mypath + myname + "\SKIY\visu.bat"
            GoTo 160
            End If
         End If
      'now input hebrewnames, and tds into arrays
l50:
      If world = True Then
            scomb% = -1
            For i% = 0 To Combo2.ListCount - 1
               Combo2.ListIndex = i%
               myname = Combo2.Text
               If Mode% = 1 Then
                  mynew = Dir(mypath + myname + "\netz\*.bat")
               ElseIf Mode% = 0 Then
                  mynew = Dir(mypath + myname + "\skiy\*.bat")
                  End If
               If mynew <> sEmpty Then
                  batfil% = FreeFile
                  If Mode% = 1 Then
                     mynew = mypath + myname + "\netz\" + mynew
                  ElseIf Mode% = 0 Then
                     mynew = mypath + myname + "\skiy\" + mynew
                     End If
                  Open mynew For Input As #batfil%
                  Input #batfil%, hebcityname$(Combo2.ListIndex), tdcities%(Combo2.ListIndex)
                  If mapsearchfm.Combo2.Text = mapgraphfm.Combo2.Text Then
                     scomb% = i%
                     combotext$ = Combo2.Text
                     End If
                  End If
                  Close #batfil%
             Next i%
       ElseIf world = False Then
          myfile = Dir(drivcities$ + "citynams.txt")
          If myfile = sEmpty Then
             'can't find the citynams.txt file, so don't do anything
          Else
             On Error GoTo cb25
             cityfound% = 0
             filcit% = FreeFile
             Open drivcities$ + "citynams.txt" For Input As #filcit%
             Do Until EOF(filcit%)
                Input #filcit%, engcitynam$
                Input #filcit%, hebcitynam$
                Combo2.AddItem engcitynam$
                If engcitynam$ = FileViewDir$ Then
                   Text2.Text = hebcitynam$
                   cityfound% = 1
                   End If
             Loop
cb25:        Close #filcit%
             On Error GoTo errhand
             End If
          Combo2.Text = FileViewDir$
          End If
      'Combo2.ListIndex = Combo2.ListCount - 1
160:
      With mapgraphfm
        .Picture2.Visible = True
        .Picture2.Refresh
        .Picture3.Visible = True
        .Picture3.Refresh
        .MSFlexGrid1.Visible = False
        .restorelimitsbut.Enabled = False
        .Text3.Enabled = False
        .UpDown2.Enabled = False
        .Command2.Enabled = False
        .Command3.Enabled = False
        .Text1.Visible = True
        .Text1.Refresh
        .UpDown1.Visible = True
        .Text2.Visible = True
        .Text2.Refresh
        .Label10.Visible = True
        .Label10.Refresh
        .Frame1.Visible = True
        .Frame1.Refresh
        .Label12.Visible = True
        .Label12.Refresh
        .Calendarbut.Enabled = True
        .Calendarbut.Refresh
      
        If AutoProf Or AutoScanlist Then
           .shpDelay.Visible = True
           .shpDelay.Refresh
           .lblDelay.Visible = True
           .lblDelay.Refresh
           .txtDelay.Visible = True
           .txtDelay.Refresh
           .updnDelay.Visible = True
           End If
           
      End With
      
      mapgraphfm.StatusBar1.Panels(1) = "Input Zone time (>0 for East long.), enter city name, and press Calendar button again"
      If world = True Then
         TimeZonebut.Enabled = True
         Call Combo2_Click
         Screen.MousePointer = vbDefault
         If mapsearchfm.Visible = True Then
            dirsavecheck.value = vbChecked
            End If
         If scomb% = -1 Then
            mapgraphfm.Combo2.Text = sEmpty
            mapgraphfm.Text2 = sEmpty
            mapgraphfm.Text1 = "0"
         Else
            mapgraphfm.Combo2.Text = combotext$
            mapgraphfm.Text2 = hebcityname$(scomb%)
            mapgraphfm.Text1 = tdcities%(scomb%)
            End If
       Else
         Call Combo2_Click
         Screen.MousePointer = vbDefault
         dirsavecheck.value = vbChecked
         mapgraphfm.StatusBar1.Panels(1) = "Enter city's English (directory) name and Hebrew name"
         Text1.Text = 2
         Text1.Enabled = False
         UpDown1.Enabled = False
         Combo2.Enabled = True
         TimeZonebut.Enabled = False
         Exit Sub
      End If
   ElseIf calpress% = 2 Then
      If Text2.Text = sEmpty Then
         Text2.Enabled = True
         MsgBox "You must provide a Hebrew name of the city!", vbExclamation + vbOKOnly, "Maps & More"
         Screen.MousePointer = vbDefault
         calpress% = 1
         Exit Sub
         End If
      Screen.MousePointer = vbHourglass
      TimeZonebut.Enabled = False
      mapgraphfm.StatusBar1.Panels(1) = sEmpty
      tdcal% = Text1
      'now generate "pr1" file, and continue calendar calculations
      'If sunmode% = 1 Then 'sunrise
      'ElseIf sunrmode% = 0 Then 'sunset
      '   End If
'      ret = SetWindowPos(mapprogressfm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'  ret = SetWindowPos(mapprogressfm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'  ret = SetWindowPos(mapgraphfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      With mapgraphfm
        .Picture2.Visible = False
        .Picture3.Visible = False
        .MSFlexGrid1.Visible = True
        .Text1.Visible = False
        .UpDown1.Visible = False
        .Frame1.Visible = False
        .Label12.Visible = False
        .Text3.Enabled = True
        .UpDown2.Enabled = True
        .Calendarbut.Enabled = False 'so can't repeat the same operation
        .shpDelay.Visible = False
        .lblDelay.Visible = False
        .txtDelay.Visible = False
        .updnDelay.Visible = False
      End With
      If world = True Then mapgraphfm.Command2.Enabled = True
      If world = False Then
         TimeZonebut.Enabled = False
         End If
      mapgraphfm.Command3.Enabled = True
      mapgraphfm.frmObstructions.Enabled = True
      ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      'make save files (.pr* files, fnz,fsk files and directories)
      'if save to cities directories flagged then make directory under
      'that name and make netz/skiy subdirectories with bat files and
      'pr* files.  Also the Hebrew cityname, and directory name is added
      'to the citynames.txt file.  If save not flagged, then write to cities/erostmp file
      'and make bat file for it.  All EROS output has .bat file name eros.bat to
      'flag CAL PROGRAM to look for new value of the zone time.
      'If world = True Then
      '   plotfile$ = drivjk$ + "eros.tmp"
      '   'myfile = Dir(drivjk$ + "eros.tmp")
      'Else
      '   plotfile$ = drivjk$ + "EYisroel.tmp"
      '   End If
      If Dir(plotfile$) = sEmpty Then
         ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         response = MsgBox("Can't find the graph file: " & plotfile$, vbCritical + vbOKOnly, "Maps & More")
         ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         GoTo ca999
         End If
      'check that eros subdirectories on d:\prom and d:\prof and d:\cities already exist
      'if not, create them
      If world = False Then GoTo cb50
      myfile = Dir(drivcities$ + "eros", vbDirectory)
      If myfile = sEmpty Then MkDir (drivcities$ + "eros")
      If Mode% = 1 Then
         myfile = Dir(drivcities$ + "eros\netz", vbDirectory)
         If myfile = sEmpty Then MkDir (drivcities$ + "eros\netz")
      ElseIf Mode% = 0 Then
         myfile = Dir(drivcities$ + "eros\skiy", vbDirectory)
         If myfile = sEmpty Then MkDir (drivcities$ + "eros\skiy")
         End If
      'myfile = Dir(drivprof$ + "eros", vbDirectory)
      'If myfile = sEmpty Then MkDir (drivprof$ + "eros")
      'myfile = Dir(drivprom$ + "eros", vbDirectory)
      'If myfile = sEmpty Then MkDir (drivprom$ + "eros")
      If world = True Then
         drcities$ = drivcities$ + "eros\" 'drivprof$ + "eros\"
      Else
         drcities$ = drivcities$ 'drivprof$
         End If
cb50: If dirsavecheck.value = vbUnchecked Then
'--------------------world = true,false, no directory specified-----------
         'Copy plotfile$ to \cities as name eros.fnz/fsk
         'This will also apply for output from Analyze
         'since no directory was specified.
         'erosfile$(1) = drivprof$ + "eros\eros0001.fnz"
         'erosfile$(0) = drivprof$ + "eros\eros0001.fsk"
         erosfile$(1) = drivcities$ + "eros\netz\eros0001.pr1"
         erosfile$(0) = drivcities$ + "eros\skiy\eros0001.pr1"
         dirfile$(1) = drivcities$ + "eros\netz\"
         dirfile$(0) = drivcities$ + "eros\skiy\"
         erosfile$(1) = dirfile$(1) & "eros0001.pr1"
         erosfile$(0) = dirfile$(0) & "eros0001.pr1"
         erostmpfil% = FreeFile
         Open plotfile$ For Input As #erostmpfil%
         Line Input #erostmpfil%, doclin$
         Input #erostmpfil%, kmyeros, kmxeros, hgteros, ae, be, ce, de, fe
         kmxeros = -kmxeros
         FileCopy plotfile$, erosfile$(Mode%) 'copy eros.tmp to cities/eros/netz or skiy
         'If mode% = 1 Then
         '   FileCopy plotfile$, drivprom$ + "eros\eros0001.001"
         'ElseIf mode% = 0 Then
         '   FileCopy plotfile$, drivprom$ + "eros\eros0001.004"
         '   End If
         'newfile% = FreeFile
         'Open dirfile$(mode%) + "eros0001.pr1" For Output As #newfile%
         'Write #newfile%, "FILENAME, LAT, LOG, HGT: ", erosfile$(mode%), kmyeros, -kmxeros, hgteros
         'Print #newfile%, "  AZI  VIEWANG+REFRACT   FLGSUM   FLGWIN"
         batfile% = FreeFile
         Open dirfile$(Mode%) + "eros.bat" For Output As #batfile%
         Write #batfile%, Text2, Val(Text1)
         If Mode% = 1 Then
            Write #batfile%, drivfordtm$ + "netz\eros0001.pr1", _
                  Val(Format(Trim$(Str$(kmyeros)), "###0.0######")), _
                  Val(Format(Trim$(Str$(-kmxeros)), "###0.0######")), _
                  hgteros
         ElseIf Mode% = 0 Then
            Write #batfile%, drivfordtm$ + "skiy\eros0001.pr1", _
                  Val(Format(Trim$(Str$(kmyeros)), "###0.0######")), _
                  Val(Format(Trim$(Str$(-kmxeros)), "###0.0######")), _
                  hgteros
            End If
         Print #batfile%, "version"; ","; "1"; ","; Trim$(Str$(DTMflag)); ","; "0"
         Close #batfile%
         
         'Do Until EOF(erostmpfil%)
         '   Input #erostmpfil%, azi, va, ae, be, ce, de
         '   Print #newfile%, Format(Str(azi), "##0.0"); Tab(10); Format(Str(va), "#0.0000"); Tab(21); Format(Str(0#), "0.0000"); Tab(31); Format(Str(0#), "0.0000")
         'Loop
         'Close #erostmpfil%
         'Close #newfile%
      ElseIf dirsavecheck.value = vbChecked Then
'-------------------------directory specified-------------------------------------
       If world = False Then
'-----------------------------world=false, output from Analyze------------------------
            Mode% = sunmode%
            dirfile$(1) = drivcities$ + LTrim$(RTrim$(Combo2.Text)) + "\netz\"
            dirfile$(0) = drivcities$ + LTrim$(RTrim$(Combo2.Text)) + "\skiy\"
            'check if they exist, if not,create them
            If Dir(dirfile$(1), vbDirectory) = sEmpty And Dir(dirfile$(0), vbDirectory) = sEmpty Then
               MkDir drivcities$ + LTrim$(RTrim$(Combo2.Text))
               'update citynams.txt if necessary
               If cityfound% = 0 And Dir(drivcities$ + "citynams.txt") <> sEmpty Then
                  filcit% = FreeFile
                  Open drivcities$ + "citynams.txt" For Append As #filcit%
                  Print #filcit%, LTrim$(RTrim$(Combo2.Text))
                  Print #filcit%, LTrim$(RTrim$(Text2.Text))
                  Close #filcit%
                  End If
               End If
           
         If Dir(dirfile$(Mode%), vbDirectory) = sEmpty Then
            MkDir dirfile$(Mode%)
            End If
               
         newfile% = FreeFile
         'If world = True Then
         '   fnn$ = Mid$(fileo$, 1, 8)
         'Else
          'fnn$ = OutFile$ 'Mid$(fileo$, 1, 8)
          fnn$ = Mid$(OutFile$, 1, 8)
          iipr$ = Mid$(OutFile$, 9, 4)
         '    End If
         '-----------new procedures and formats----------
         'GoTo ca80
         GoTo ca81
         '-----------------------------------------------
         
         ipr% = 1
ca50:    If ipr% < 10 Then
            iipr$ = ".pr" + LTrim(Str$(ipr%))
         ElseIf ipr% >= 10 And ipr% < 100 Then
            iipr$ = ".p" + LTrim(Str$(ipr%))
         ElseIf ipr% >= 100 And ipr% < 999 Then
            iipr$ = "." + LTrim(Str$(ipr%))
         ElseIf ipr% > 999 Then
            'rename the root
           newRootNum = -1
ca60:      newRootNum = newRootNum + 1
            fnn$ = Mid$(fnn$, 1, 6) + Format(LTrim(Str$(newRootNum)), "00")
            ipr% = 1
ca70:      If ipr% < 10 Then
                iipr$ = ".pr" + LTrim(Str$(ipr%))
            ElseIf ipr% >= 10 And ipr% < 100 Then
               iipr$ = ".p" + LTrim(Str$(ipr%))
            ElseIf ipr% >= 100 And ipr% < 999 Then
               iipr$ = "." + LTrim(Str$(ipr%))
            ElseIf ipr% > 999 Then
               GoTo ca60
               End If
               
            tmpfil$ = dirfile$(Mode%) + fnn$ + iipr$
            'check if this file already exists
            myfile = Dir(tmpfil$)
            If myfile <> sEmpty Then
               ipr% = ipr% + 1
               GoTo ca70
            Else
               GoTo ca80
               End If
            End If
ca75:    tmpfil$ = dirfile$(Mode%) + fnn$ + iipr$
         'check if this file already exists
         myfile = Dir(tmpfil$)
         If myfile <> sEmpty Then
            ipr% = ipr% + 1
            GoTo ca50
            End If
ca80:    FileCopy fileo$, tmpfil$
ca81:
         'look for a preexisting bat file
         cityname$ = LTrim$(RTrim$(Mid$(Combo2.Text, 1, 4)))
         If Dir(dirfile$(Mode%) + "*.bat") = sEmpty Then '<<<
            batname$ = dirfile$(Mode%) + cityname$ + ".bat"
            batfile% = FreeFile
            Open dirfile$(Mode%) + cityname$ + ".bat" For Output As #batfile%
            If Mode% = 1 Then
               Write #batfile%, drivfordtm$ + "netz\" + fnn$ + iipr$, _
                     Val(Trim$(Format(CStr(coordAnalyze(0)), "####.0#####"))), _
                     Val(Trim$(Format(CStr(coordAnalyze(1)), "####.0#####"))), _
                     Val(Trim$(Format(CStr(coordAnalyze(2)), "####.00")))
            ElseIf Mode% = 0 Then
               Write #batfile%, drivfordtm$ + "skiy\" + fnn$ + iipr$, _
                     Val(Trim$(Format(CStr(coordAnalyze(0)), "####.0#####"))), _
                     Val(Trim$(Format(CStr(coordAnalyze(1)), "####.0#####"))), _
                     Val(Trim$(Format(CStr(coordAnalyze(2)), "####.00")))
               End If
            Print #batfile%, "version"; ","; "1"; ","; "9"; ","; "0"
            Close #batfile%
         Else 'just append if flagged '<<<<<<
           myfile = Dir(dirfile$(Mode%) + "*.bat")
           ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           If Not AutoProf Then
              response = MsgBox(myfile + " found. Do you wan't to append to this existing bat file?" + Chr(10), _
                      vbQuestion + vbYesNo, "Maps & More")
           Else
              response = vbYes 'automatically append
              End If
           ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           If response = vbYes Then '<<
               'find last version number and increment
               batfile% = FreeFile
               Open dirfile$(Mode%) + myfile For Input As #batfile%
               tmp2num% = FreeFile
               tmp2fil$ = dirfile$(Mode%) + "bat.tmp"
               Open tmp2fil$ For Output As #tmp2num%
               Line Input #batfile%, doclin$ 'read first line of documentation
               Print #tmp2num%, doclin$
               Do Until EOF(batfile%)
                  Line Input #batfile%, doclin$
                  If InStr(LCase$(doclin$), "version") = 0 Then
                     Print #tmp2num%, doclin$
                     'Print #tmp2num%, docbat$ & "," & _
                     'Trim$(Format(CStr(batlat), "####.0000")) & "," & _
                     'Trim$(Format(CStr(batlog), "####.0000")) & "," & _
                     'Trim$(Format(bathgt, "####.00"))
                  Else
                     If checkver = True Then
                        pos% = InStr(7, LCase$(doclin$), ",")
                        pos2% = InStr(pos% + 1, LCase$(doclin$), ",")
                        vernum$ = Mid$(doclin$, pos% + 1, pos2% - pos% - 1)
                        If Not AutoProf Then
                           response = MsgBox("The present version number is: " & vernum$ & ". Do you want to change the version number?", vbQuestion + vbYesNoCancel, "Maps & More")
                        Else
                           If AutoVer Then
                              response = vbYes 'automatically increment version number
                           Else
                              response = vbNo
                              End If
                           End If
                        If response = vbYes Then
                           versionnum = Val(vernum$) + 1
                        Else
                           versionnum = Val(vernum$)
                           End If
                        End If
                     End If
               Loop
               Close #batfile%
               'Open dirfile$(mode%) + cityname$ + ".bat" For Append As #batfile%
               'If mode% = 1 Then
               '   Write #batfile%, drivfordtm$ + "netz\" + fnn$ + iipr$, kmyeros, -kmxeros, hgteros
               'ElseIf mode% = 0 Then
               '   Write #batfile%, drivfordtm$ + "skiy\" + fnn$ + iipr$, kmyeros, -kmxeros, hgteros
               '   End If
               If Mode% = 1 Then
                  Write #tmp2num%, drivfordtm$ + "netz\" + fnn$, _
                  Val(Trim$(Format(CStr(coordAnalyze(0)), "####.0#####"))), _
                  Val(Trim$(Format(CStr(coordAnalyze(1)), "####.0#####"))), _
                  Val(Trim$(Format(CStr(coordAnalyze(2)), "####.00")))
               ElseIf Mode% = 0 Then
                  Write #tmp2num%, drivfordtm$ + "skiy\" + fnn$, _
                  Val(Trim$(Format(CStr(coordAnalyze(0)), "####.0#####"))), _
                  Val(Trim$(Format(CStr(coordAnalyze(1)), "####.0#####"))), _
                  Val(Trim$(Format(CStr(coordAnalyze(2)), "####.00")))
                  End If
               Print #tmp2num%, "Version" & "," & Trim$(CStr(versionnum)) & ",9,0"
               Close #tmp2num%
               Kill dirfile$(Mode%) + myfile
               Name tmp2fil$ As dirfile$(Mode%) + myfile
           Else
               ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               If Not AutoProf Then
                  response = MsgBox("This operation will append the old bat file to bat.tmp and write a new one in it's place!" + Chr(10) + _
                                  "You still want to proceed?", vbExclamation + vbYesNo, "Maps & More")
               Else
                  response = vbYes 'automatic append
                  End If
               ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               If response = vbNo Then
                  Close
                  GoTo ca999
               End If
               'append the current bat file to the contents of bat.tmp
               filbak% = FreeFile
               batfile% = FreeFile
               Open dirfile$(Mode%) + myfile For Input As #filbak%
               If Dir(dirfile$(Mode%) + "bat.tmp") = sEmpty Then
                  batfile% = FreeFile
                  Open dirfile$(Mode%) + "bat.tmp" For Output As #batfile%
                Else
                  Open dirfile$(Mode%) + "bat.tmp" For Append As #batfile%
                  End If
               Do Until EOF(filbak%)
                  Line Input #filbak%, doclin$
                  Print #batfile%, doclin$
               Loop
               Close #filbak%
               Close #batfile%
               'now write the new bat file
               batfile% = FreeFile
               Open dirfile$(Mode%) + myfile For Output As #batfile%
               If Mode% = 1 Then
                  'Write #batfile%, drivfordtm$ + "netz\" + fnn$ + iipr$, _
                  '   Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                  '   Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                  '   hgteros
                  Write #batfile%, drivfordtm$ + "netz\" + fnn$ + iipr$, _
                     Val(Trim$(Format(CStr(coordAnalyze(0)), "####.0#####"))), _
                     Val(Trim$(Format(CStr(coordAnalyze(1)), "####.0#####"))), _
                     Val(Trim$(Format(CStr(coordAnalyze(2)), "####.00")))
                  'If world = True Then Write #batfile%, Text2, Val(Text1)
                ElseIf Mode% = 0 Then
                  'Write #batfile%, drivfordtm$ + "skiy\" + fnn$ + iipr$, _
                  '      Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                  '      Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                  '      hgteros
                  Write #batfile%, drivfordtm$ + "skiy\" + fnn$ + iipr$, _
                  Val(Trim$(Format(CStr(coordAnalyze(0)), "####.0#####"))), _
                  Val(Trim$(Format(CStr(coordAnalyze(1)), "####.0#####"))), _
                  Val(Trim$(Format(CStr(coordAnalyze(2)), "####.00")))
               End If
               Print #batfile%, "version"; ","; "1"; ","; "9"; ","; "0"
               Close #batfile%
           End If '<<
           
           'now update the citynams.txt file if necessary
           myfile = Dir(drivcities$ + "citynams.txt")
           If myfile = sEmpty Then
             'can't find the citynams.txt file, so don't do anything
           Else
             On Error GoTo cb25
             cityfound% = 0
             filcit% = FreeFile
             Open drivcities$ + "citynams.txt" For Input As #filcit%
             Do Until EOF(filcit%)
                Input #filcit%, engcitynam$
                Input #filcit%, hebcitynam$
                Combo2.AddItem engcitynam$
                If engcitynam$ = FileViewDir$ Then
                   cityfound% = 1
                   GoTo ca280
                   End If
             Loop
             Close #filcit%
             filcit% = FreeFile
             Open drivcities$ + "citynams.txt" For Append As #filcit%
             Print #filcit%, Combo2.Text
             Print #filcit%, Text2.Text
             Close #filcit%
             End If
             
         End If '<<<
       'End If '<<<<< 'last revision (7-1-02) commented this out
             
       GoTo ca280
'-------------------------------world=true--------------------------------
    Else 'world=true, output from sunrisesunset
         
         'store as a permanent city file as a subdirectory of ther cities\eros directory
         'if world=true then filename will be derived from city name
         cityname$ = LTrim$(RTrim$(Mid$(Combo2.Text, 1, 4)))
'         lenfil% = Len(cityname$) 'if necessary, pad cityname to make it 4 characters long
'         If lenfil% < 4 Then cityname$ = cityname$ + String(4 - lenfil%, "0")
'         icity% = 1
'ca100:   If icity% < 10 Then
'            iicity$ = "000" + LTrim$(Str$(icity%))
'         ElseIf icity% >= 10 And icity% < 100 Then
'            iicity$ = "00" + LTrim$(Str$(icity%))
'         ElseIf icity% >= 100 And icity% < 1000 Then
'            iicity$ = "0" + LTrim$(Str$(icity%))
''        ElseIf icity% >= 1000 And icity% < 10000 Then
''           iicity$ = LTrim$(Str$(icity%))
'         Else
'            ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            response = MsgBox("Can only have up to 999 profiles for any city file!", vbCritical + vbOKOnly, "Maps & More")
'            ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            Screen.MousePointer = vbDefault
'            GoTo ca999
'            End If
'         erosfile$(1) = drivcities$ + "eros\" + cityname$ + iicity$ + ".fnz"
'         erosfile$(0) = drivcities$ + "eros\" + cityname$ + iicity$ + ".fsk"
'         myfile = Dir(erosfile$(mode%))
'         If myfile <> sEmpty Then
'            icity% = icity% + 1
'            GoTo ca100
'            End If
         dirfile$(1) = drivcities$ + "eros\" + LTrim$(RTrim$(Combo2.Text)) + "\netz\"
         dirfile$(0) = drivcities$ + "eros\" + LTrim$(RTrim$(Combo2.Text)) + "\skiy\"
         myfile = Dir(drivcities$ + "eros\" + LTrim$(RTrim$(Combo2.Text)), vbDirectory)
         If myfile = sEmpty Then
            MkDir (drivcities$ + "eros\" + LTrim$(RTrim$(Combo2.Text)))
            MkDir (dirfile$(Mode%))
         Else
            myfile = Dir(dirfile$(Mode%), vbDirectory) 'check if sub directory exists
            If myfile = sEmpty Then MkDir (dirfile$(Mode%))
            End If
         erostmpfil% = FreeFile
         Open plotfile$ For Input As #erostmpfil%
         Line Input #erostmpfil%, doclin$
         Input #erostmpfil%, kmyeros, kmxeros, hgteros, ae, be, ce, de, fe
         kmxeros = -kmxeros
         
         'check for place duplication
         batname$ = dirfile$(Mode%) + cityname$ + ".bat"
         myfile = Dir(batname$)
         If myfile <> sEmpty Then
            batfile% = FreeFile
            Open dirfile$(Mode%) + cityname$ + ".bat" For Input As #batfile%
            Line Input #batfile%, doclin$
            Do Until EOF(batfile%)
               Input #batfile%, doclin$, kmyeroscheck, kmxeroscheck, hgteroscheck
               If kmyeros = kmyeroscheck And kmxeros = -kmxeroscheck And hgteroscheck = hgteros Then
                  Close #batfile%
                  Screen.MousePointer = vbDefault
                  For i% = 0 To Forms.count - 1
                     ret = SetWindowPos(Forms(i%).hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                  Next i%
                  If Not AutoProf Then response = MsgBox("You have already recorded a place with these coordinates and elevations!", vbOKOnly + vbExclamation, "Maps & More")
                  For i% = 0 To Forms.count - 1
                    ret = SetWindowPos(Forms(i%).hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                  Next i%
                  GoTo ca999
                  End If
            Loop
            Close #batfile%
            End If
         
'         FileCopy plotfile$, erosfile$(mode%)
         'If mode% = 1 Then
         '   FileCopy plotfile$, drivprom$ + "eros\" + cityname$ + iicity$ + ".001"
         'ElseIf mode% = 0 Then
         '   FileCopy plotfile$, drivprom$ + "eros\" + cityname$ + iicity$ + ".004"
         '   End If
         'determine name of output file
         lenfil% = Len(LTrim$(RTrim$(Combo2.Text)))
         If lenfil% >= 8 Then 'if necessary, pad combo2.text to make it 8 characters long
            fnn$ = Mid$(LTrim$(RTrim$(Combo2.Text)), 1, 8)
         Else
            fnn$ = LTrim$(RTrim$(Combo2.Text)) + String$(8 - lenfil%, "0")
            End If

ca82:    newfile% = FreeFile
         If world = False Then fnn$ = Mid$(fileo$, 1, 8)
         ipr% = 1
ca84:    If ipr% < 10 Then
            iipr$ = ".pr" + LTrim(Str$(ipr%))
         ElseIf ipr% >= 10 And ipr% < 100 Then
            iipr$ = ".p" + LTrim(Str$(ipr%))
         ElseIf ipr% >= 100 And ipr% < 999 Then
            iipr$ = "." + LTrim(Str$(ipr%))
         ElseIf ipr% > 999 Then
             'rename the root
            newRootNum = -1
ca86:       newRootNum = newRootNum + 1
            fnn$ = Mid$(fnn$, 1, 6) + Format(LTrim(Str$(newRootNum)), "00")
            ipr% = 1
ca88:       If ipr% < 10 Then
                iipr$ = ".pr" + LTrim(Str$(ipr%))
            ElseIf ipr% >= 10 And ipr% < 100 Then
               iipr$ = ".p" + LTrim(Str$(ipr%))
            ElseIf ipr% >= 100 And ipr% < 999 Then
               iipr$ = "." + LTrim(Str$(ipr%))
            ElseIf ipr% > 999 Then
               GoTo ca86
               End If
               
            tmpfil$ = dirfile$(Mode%) + fnn$ + iipr$
            'check if this file already exists
            myfile = Dir(tmpfil$)
            If myfile <> sEmpty Then
               ipr% = ipr% + 1
               GoTo ca88
            Else
               GoTo ca90
               End If
'            ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            response = MsgBox("You have exceeded the maximum number of profile files allowed for any unique root name (defined by the first 4 letters of the city directory name)--pick a different name!", vbCritical + vbOKOnly, "Maps & More")
'            ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'            Screen.MousePointer = vbDefault
'            GoTo ca999
            End If
         tmpfil$ = dirfile$(Mode%) + fnn$ + iipr$
         'check if this file already exists
         myfile = Dir(tmpfil$)
         If myfile <> sEmpty Then
            ipr% = ipr% + 1
            GoTo ca84
            End If
         'Open tmpfil$ For Output As #newfile%
ca90:    FileCopy plotfile$, tmpfil$
         'Write #newfile%, "FILENAME, LAT, LOG, HGT: ", erosfile$(mode%), kmyeros, kmxeros, hgteros
         'Print #newfile%, "  AZI  VIEWANG+REFRACT   FLGSUM   FLGWIN"
         batname$ = dirfile$(Mode%) + cityname$ + ".bat"
         myfile = Dir(batname$)
         batfile% = FreeFile
         If myfile = sEmpty Then
            Open dirfile$(Mode%) + cityname$ + ".bat" For Output As #batfile%
            Write #batfile%, Text2, Val(Text1)
            If Mode% = 1 Then
               Write #batfile%, drivfordtm$ + "netz\" + fnn$ + iipr$, _
                      Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                      Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                      hgteros
            ElseIf Mode% = 0 Then
               Write #batfile%, drivfordtm$ + "skiy\" + fnn$ + iipr$, _
                       Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                       Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                       hgteros
               End If
            Print #batfile%, "version"; ","; "1"; ","; "0"; ","; "0"
            Close #batfile%
         Else 'just append if flagged
            If mapsearchfm.Visible = True Then
               response = vbYes
               GoTo ca250
               End If
            ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            If Not AutoProf Then
               response = MsgBox("Do you wan't to append this profile file to the existing list? (Answer NO in order to erase the old one)?", vbQuestion + vbYesNo, "Maps & More")
            Else
               response = vbYes 'automatically append
               End If
            ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
ca250:      If response = vbYes Then
               'find last version number and increment
               Open dirfile$(Mode%) + cityname$ + ".bat" For Input As #batfile%
               tmp2num% = FreeFile
               tmp2fil$ = dirfile$(Mode%) + "bat.tmp"
               Open tmp2fil$ For Output As #tmp2num%
               Line Input #batfile%, doclin$ 'read first line of documentation
               Print #tmp2num%, doclin$
               Do Until EOF(batfile%)
                  Input #batfile%, docbat$, batlat, batlog, bathgt
                  If LCase$(docbat$) <> "version" Then
                     Write #tmp2num%, docbat$, batlat, batlog, bathgt
                  Else
                     versionnum = batlat + 1 'default is to increment
                     If checkver = True Then
                        If Not AutoProf Then
                           response = MsgBox("The present version number is: " & Str(batlat) & ". Do you want to change the version number?", vbQuestion + vbYesNoCancel, "Maps & More")
                        Else
                           If AutoVer Then
                              response = vbYes 'automatically increment version number
                           Else
                              response = vbNo
                              End If
                           End If
                        If response = vbYes Then
                           versionnum = batlat + 1
                        Else
                           versionnum = batlat
                           End If
                        End If
                     End If
               Loop
               Close #batfile%
               'Open dirfile$(mode%) + cityname$ + ".bat" For Append As #batfile%
               'If mode% = 1 Then
               '   Write #batfile%, drivfordtm$ + "netz\" + fnn$ + iipr$, kmyeros, -kmxeros, hgteros
               'ElseIf mode% = 0 Then
               '   Write #batfile%, drivfordtm$ + "skiy\" + fnn$ + iipr$, kmyeros, -kmxeros, hgteros
               '   End If
               If Mode% = 1 Then
                  Write #tmp2num%, drivfordtm$ + "netz\" + fnn$ + iipr$, _
                        Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                        Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                        hgteros
               ElseIf Mode% = 0 Then
                  Write #tmp2num%, drivfordtm$ + "skiy\" + fnn$ + iipr$, _
                        Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                        Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                        hgteros
                  End If
               Print #tmp2num%, "Version"; ","; RTrim$(LTrim$(Str(versionnum))); ","; Trim$(Str$(DTMflag)); ","; "0"
               Close #tmp2num%
               Kill dirfile$(Mode%) + cityname$ + ".bat"
               Name tmp2fil$ As dirfile$(Mode%) + cityname$ + ".bat"
             Else
               ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               response = MsgBox("This operation will erase any existing profile list!  You still want to proceed?", vbExclamation + vbYesNo, "Maps & More")
               ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               If response = vbNo Then
                  Close
                  GoTo ca999
                  End If
               Open dirfile$(Mode%) + cityname$ + ".bat" For Output As #batfile%
               Write #batfile%, Text2, Val(Text1)
               If Mode% = 1 Then
                  Write #batfile%, drivfordtm$ + "netz\" + fnn$ + iipr$, _
                        Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                        Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                        hgteros
               ElseIf Mode% = 0 Then
                  Write #batfile%, drivfordtm$ + "skiy\" + fnn$ + iipr$, _
                        Val(Format(Trim$(Str$(kmyeros)), "###0.0#####")), _
                        Val(Format(Trim$(Str$(-kmxeros)), "###0.0#####")), _
                        hgteros
                  End If
               Print #batfile%, "version"; ","; "1"; ","; Trim$(Str$(DTMflag)); ","; "0"
               Close #batfile%
               End If
            End If
         'Do Until EOF(erostmpfil%)
         '   Input #erostmpfil%, azi, va, ae, be, ce, de
         '   Print #newfile%, Format(Str(azi), "##0.0"); Tab(10); Format(Str(va), "#0.0000"); Tab(21); Format(Str(0#), "0.0000"); Tab(31); Format(Str(0#), "0.0000")
         'Loop
         'Close #erostmpfil%
         'Close #newfile%
         
        End If
        End If '<<last revision (7-01-02) added this
        
        '==============END OF ROUTINES==============

ca280:  If AutoScanlist Or AutoProf Then
           'unload form and continue with automatic scans
           GoTo ca999
           End If
        
        mapgraphfm.restorelimitsbut.Enabled = True
        mapgraphfm.Refresh
        mapPictureform.Refresh
        ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
        ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
        response = MsgBox("Do you wan't to execute CAL PROGRAM now?", vbQuestion + vbYesNo, "Maps & More")
ca300:  If response = vbNo Then
            Screen.MousePointer = vbDefault
            ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            GoTo ca999
            End If
        ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
        ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
ca320:  resp$ = "c:\progra~1\devstu~1\vb\"
ca325:  If Dir(resp$ & "Cal Program.exe") = sEmpty Then
           'try another possiblility
           resp$ = "c:\devstu~1\vb\"
           If Dir(resp$ & "Cal Program.exe") <> sEmpty Then GoTo ca325
           Screen.MousePointer = vbDefault
           resp$ = InputBox("Can't find path to ""Cal Program.exe""!" & vbLf & vbLf & _
                         "Input the full path with all backslashes." & vbLf & _
                         "For example: ""c:\devstudio\""", "Path to Cal Program.exe", "c:\progra~1\devstu~1\vb\")
           If resp$ = sEmpty Then
              GoTo ca999
           Else
              GoTo ca325
              End If
        Else
           resp$ = resp$ & "Cal Program.exe"
           RetVal = Shell(resp$, 1)
           End If
'        waitime = Timer
'        Do Until Timer > waitime + 10
'        Loop
'        Screen.MousePointer = vbDefault
'        lResult = FindWindow("Cal Program", vbNullString)
'        Do Until lResult <> 0
'           waitime = Timer
'           lResult = FindWindow("Cal Program", vbNullString)
'           If killpicture = True Then Exit Do
'           DoEvents
'        Loop
'        ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
ca999:
      Screen.MousePointer = vbDefault
       
      If calpress% = 2 And AutoScanlist Or AutoProf Then
          'unload this form
          Call form_queryunload(0, 0)
          Exit Sub
          End If
      
      calpress% = 0
      Line1.Refresh
      TimeZonebut.Enabled = False
      If AutoScanlist Then graphwind = False
      If killpicture = True Then Exit Sub
      End If
      
  Exit Sub
  
errhand:
    Screen.MousePointer = vbDefault
    MsgBox "Encountered Error Number: " & Str(Err.Number) & Chr(10) & _
           "Error description: " & Err.Description & Chr(10) & _
           "Exiting calendarbut routine", vbCritical + vbOKOnly, "Maps & More"
    calpress% = calpress% - 1

End Sub

Private Sub Check1_Click()
   If Check1.value = vbChecked Then
      checkver = True
   Else
      checkver = False
      End If
End Sub

Private Sub chkObstruction_Click()
   If chkObstruction.value = vbChecked Then
      ObstructionCheck = True
   Else
      ObstructionCheck = False
      End If
End Sub

Private Sub cmdExit_Click()
   Call form_queryunload(0, 0)
End Sub

Private Sub Combo2_Click()
  'Combo2.Text = Combo2.ListIndex
  If dirsavecheck.value = vbChecked Then
     If Combo2.ListIndex > 0 Then
        Text2 = hebcityname$(Combo2.ListIndex)
        Text1 = tdcities%(Combo2.ListIndex)
        End If
     End If
End Sub

Private Sub Command1_Click()
   On Error GoTo sunerrhand
suns600:    Maps.CommonDialog1.CancelError = True
            Maps.CommonDialog1.Filter = "sunrise files (*.001)|*.001|sunset files (*.004)|*.004|"
            If sunmode% = 1 Then
               Maps.CommonDialog1.FilterIndex = 1
            Else
               Maps.CommonDialog1.FilterIndex = 0
               End If
            ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            Maps.CommonDialog1.ShowSave
            savfile$ = Maps.CommonDialog1.FileName
            myfile = Dir(savfile$)
            If myfile <> sEmpty Then
               response = MsgBox("File already exists.  Do you wan't to overwrite?", vbYesNo + vbExclamation, "Maps & More")
               If response = vbNo Then
                  GoTo suns600
                  End If
               End If
            FileCopy plotfile$, savfile$
sunerrhand:
     ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
     ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

Private Sub Command2_Click()
   'Call form_queryunload(i%, j%)
   Label1.Visible = False
   Label3.Visible = False
   Label4.Visible = False
   Label7.Visible = False
   Label8.Visible = False
   MSFlexGrid1.Top = 3900
   MSFlexGrid1.Refresh
   Call sunrisesunset(10) 'mode=10 = automatic DirectX press mode
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Command3_Click
' Author    : chaim
' Date      : 7/21/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Command3_Click()

   Dim percentObstruction As Double
   
   On Error GoTo Command3_Click_Error

      Screen.MousePointer = vbHourglass
      'erase old center line
      mapgraphfm.Picture1.Line (drag1x, drag1y)-(drag2x, drag2y), QBColor(15), B
      
      'determine new x,y mins and maximuns
      If world = True Then
        xminnew = xmino
        xmaxnew = xmaxo
      Else
        xminnew = xmin
        xmaxnew = xmax
        End If
      
      'determine new y limits
      filtmp% = FreeFile
      Open plotfile$ For Input As #filtmp%
      Line Input #filtmp%, doclin$
      Line Input #filtmp%, doclin$
      nnpnt% = 0
      Do Until EOF(filtmp%)
         Input #filtmp%, azi, va, logi, lati, dista, hgti
         If nnpnt% = 0 And azi >= xminnew Then
            yminnew = va
            ymaxnew = va
         ElseIf nnpnt% <> 0 And azi >= xminnew And azi <= xmaxnew Then
            If va < yminnew Then
               yminnew = va
               End If
            If va > ymaxnew Then
               ymaxnew = va
               End If
            End If
         nnpnt% = nnpnt% + 1
      Loop
      Close #filtmp%
      xmin = xminnew
      xmax = xmaxnew
      ymin = yminnew - (ymaxnew - yminnew) * 0.1
      ymax = ymaxnew + (ymaxnew - yminnew) * 0.1
      'refresh graph with new limits
        Screen.MousePointer = vbHourglass
        mapgraphfm.Picture1.Cls
        mapgraphfm.Picture1.DrawMode = 13
        mapgraphfm.Picture1.DrawStyle = vbDefault
        mapgraphfm.Picture1.DrawWidth = 2
        myfile = Dir(plotfile$)
        If myfile <> sEmpty Then
           filtmp% = FreeFile
           Open plotfile$ For Input As #filtmp%
           Line Input #filtmp%, doclin$
           Line Input #filtmp%, doclin$
           Label7.Caption = "ymin = " + Format(yminnew, "#0.0#")
           Label7.Visible = True
           Label8.Caption = "ymax = " + Format(ymaxnew, "#0.0#")
           Label8.Visible = True
           'ymax = CInt(ymax + 0.5)
           mapgraphfm.Label6.Caption = Format(yminnew, "#0.0#")
           mapgraphfm.Label5.Caption = Format(ymaxnew, "#0.0#")
           mapgraphfm.Label3.Caption = Format(xminnew, "##0.0")
           mapgraphfm.Label4.Caption = Format(xmaxnew, "##0.0")
           If crosssection = True Then
              mapgraphfm.Label4.Caption = Format(xmaxnew * 0.001, "###0.000")
              End If
           Seek #filtmp%, 1
           Line Input #filtmp%, doclin$
           Line Input #filtmp%, doclin$
           percentObstruction = 0
           For i% = 1 To nnpnt%
              Input #filtmp%, azi, va, logi, lati, dista, hgti
              xcord = ((azi - xmin) / (xmax - xmin)) * mapgraphfm.Picture1.Width
              ycord = mapgraphfm.Picture1.Height - ((va - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
              If dista <= Val(Text3.Text) Then
                 linecolor = QBColor(14)
                 percentObstruction = percentObstruction + 1
              Else
                 linecolor = QBColor(1)
                 End If
              If i% = 1 Then
                 mapgraphfm.Picture1.PSet (xcord, ycord), linecolor
                 xo = xcord
                 yo = ycord
              Else
                 mapgraphfm.Picture1.Line (xo, yo)-(xcord, ycord), linecolor
                 xo = xcord
                 yo = ycord
                 End If
           Next i%
           Close #filtmp%
           If yminnew < 0 Then
              xcord1 = 0
              ycord = mapgraphfm.Picture1.Height - ((0 - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
              mapgraphfm.Picture1.DrawStyle = vbDot
              mapgraphfm.Picture1.DrawWidth = 1
              mapgraphfm.Picture1.Line (xcord1, ycord)-(mapgraphfm.Picture1.Width, ycord), QBColor(0)
              End If
           End If
           
           percentObstruction = 100 * percentObstruction / nnpnt%
           If percentObstruction > Val(txtObstructions.Text) And chkObstruction.value = vbChecked Then
              'don't record this profile
              'wait a little bit before skipping
              waitime = Timer
              Do Until Timer > waitime + 0.5
                 DoEvents
              Loop
              cmdExit.value = True
              End If
           
        Screen.MousePointer = vbDefault
   'reprocess graph with new limits

   On Error GoTo 0
   Exit Sub

Command3_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command3_Click of Form mapgraphfm"
End Sub

Private Sub dirsavecheck_Click()
   If dirsavecheck.value = vbChecked Then
      Label11.Enabled = True
      Combo2.Enabled = True
      If world = True Then Combo2.ListIndex = Combo2.ListCount - 1
   ElseIf dirsavecheck.value = vbUnchecked Then
      Label11.Enabled = False
      Combo2.Enabled = False
      Combo2.Text = ""
      'Combo2.ListIndex = Combo2.ListCount - 1
      End If
End Sub

Private Sub form_load()

   If world Then
      Text3.Text = 6
      If ObstructionCheck Then
         chkObstruction.value = vbChecked
         End If
      End If
   
   If AutoScanlist Then
      mapgraphfm.txtDelay = Str$(IntOld2%)
   Else
      If Delay% = 0 Then Delay% = 45
      mapgraphfm.updnDelay.value = Delay%
      mapgraphfm.txtDelay = Val(Delay%)
      End If
   
   graphwind = True
   checkver = True
   calpress% = 0
   'Combo2.Text = sEmpty
   'For i% = 0 To Combo2.ListCount - 1
   '   Combo2.ListIndex = i%
   '   Combo2.Text = sEmpty
   'Next i%
   Combo2.Clear
   Screen.MousePointer = vbHourglass
   mapgraphfm.Picture1.DrawMode = 13
   mapgraphfm.Picture1.DrawWidth = 2
   If crosssection = False Then
        If sunmode% >= 1 Then
           mapgraphfm.Caption = "Sunrise horizon profile"
        ElseIf sunmode% <= 0 Then
           mapgraphfm.Caption = "Sunset horizon profile"
           End If
        plotfile$ = drivjk_c$ + "eros.tmp"
        If world = False Then
           plotfile$ = drivjk_c$ + "EYisroel.tmp"
           TimeZonebut.Enabled = False
           Command2.Enabled = False
           End If
   Else
       If world = True Then
          MSFlexGrid1.FormatString = "^azimuth     |^view angle     |^latitude      |^longitude      |^distance     |^height        "
       Else
          MSFlexGrid1.FormatString = "^azimuth     |^view angle     |^ITMx          |^ITMy           |^distance     |^height        "
          End If
       mapgraphfm.Caption = "Cross Section Plot"
       Command1.Enabled = False
       Command2.Enabled = False
       TimeZonebut.Enabled = False
       Command3.Enabled = False
       frmObstructions.Enabled = False
       Text3.Enabled = False
       UpDown2.Enabled = False
       restorelimitsbut.Enabled = True
       Calendarbut.Enabled = False
       Label1.Caption = "Distance from first point (km)"
       Label2.Caption = "Hgt (m)"
       Label3.Caption = "0"
       Label4.Caption = sEmpty
       plotfile$ = drivjk_c$ + "crossect.tmp"
       End If
ermk% = 1
On Error GoTo errhand
errmk% = 1
100
   myfile = Dir(plotfile$)
   If myfile <> sEmpty Then
      filtmp% = FreeFile
      Open plotfile$ For Input As #filtmp%
      Line Input #filtmp%, doclin$
      Line Input #filtmp%, doclin$
      Do Until EOF(filtmp%)
         Line Input #filtmp%, doclin$
         Combo1.AddItem doclin$
      Loop
      Seek #filtmp%, 1
      Line Input #filtmp%, doclin$
      Line Input #filtmp%, doclin$
      nnpnt% = 0
      Do Until EOF(filtmp%)
         Input #filtmp%, azi, va, logi, lati, dista, hgti
         If crosssection = True Then
            dista = dista * 0.001
            End If
         If nnpnt% = 0 Then
            xmin = azi
            xmax = azi
            ymin = va
            ymax = va
            If crosssection = True Then
               xmin = dista
               xmax = dista
               ymin = hgti
               ymax = hgti
               End If
         Else
            If crosssection = False Then
                If azi > xmax Then
                   xmax = azi
                   End If
                If va < ymin Then
                   ymin = va
                   End If
                If va > ymax Then
                   ymax = va
                   End If
            Else
                If dista > xmax Then
                   xmax = dista
                   End If
                If hgti < ymin Then
                   ymin = hgti
                   End If
                If hgti > ymax Then
                   ymax = hgti
                   End If
               End If
            End If
         nnpnt% = nnpnt% + 1
      Loop
      If ymin = ymax Then 'make sure that there is nonzero ranges
         ymin = ymin - 1
         ymax = ymax + 1
         End If
      If xmin = xmax Then
         response = MsgBox("The range of x values is zero!", vbCritical + vbOKOnly, "Maps & More")
         Close #filtmp%
         Exit Sub
         End If
      xmino = xmin
      xmaxo = xmax
      ymino = ymin
      ymaxo = ymax
      Label3.Caption = Format(xmin, "###0.000")
      Label4.Caption = Format(xmax, "###0.000")
      If crosssection = False And world = False Then
         'output from Analyze
         sectnumpnt& = (xmax - xmin) * 10 + 1
         End If
      Label7.Caption = "ymin = " + Format(ymin, "#0.0#")
      Label7.Visible = True
      ymin = CInt(ymin - 0.5)
      If ymin > 0 And crosssection = False Then ymin = 0
      Label8.Caption = "ymax = " + Format(ymax, "#0.0#")
      Label8.Visible = True
      ymax = CInt(ymax + 0.5)
      mapgraphfm.Label6.Caption = Format(ymin, "#0.0#")
      mapgraphfm.Label5.Caption = Format(ymax, "#0.0#")
      If crosssection = True Then
         Label7.Caption = "ymin = " + Format(ymin, "###0.0")
         Label8.Caption = "ymax = " + Format(ymax, "###0.0")
         mapgraphfm.Label6.Caption = Format(ymin, "###0.0")
         mapgraphfm.Label5.Caption = Format(ymax, "###0.0")
         End If
      Seek #filtmp%, 1
      Line Input #filtmp%, doclin$
      Line Input #filtmp%, doclin$
      xmino = xmin
      xmaxo = xmax
      ymino = ymin
      ymaxo = ymax
       
      On Error GoTo errhand
      ermk% = 1
150
      For i% = 1 To nnpnt%
         Input #filtmp%, azi, va, logi, lati, dista, hgti
         If crosssection = False Then
            xcord = ((azi - xmin) / (xmaxo - xmin)) * mapgraphfm.Picture1.Width
            ycord = mapgraphfm.Picture1.Height - ((va - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
         Else
            dista = dista * 0.001
            xcord = ((dista - xmin) / (xmaxo - xmin)) * mapgraphfm.Picture1.Width
            ycord = mapgraphfm.Picture1.Height - ((hgti - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
            End If
         If i% = 1 Then
            mapgraphfm.Picture1.PSet (xcord, ycord), QBColor(1)
            xo = xcord
            yo = ycord
         Else
            mapgraphfm.Picture1.Line (xo, yo)-(xcord, ycord), QBColor(1)
            xo = xcord
            yo = ycord
            End If
      Next i%
      Close #filtmp%
      If ymin < 0 Then
         xcord1 = 0
         ycord = mapgraphfm.Picture1.Height - ((0 - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
         mapgraphfm.Picture1.DrawStyle = vbDot
         mapgraphfm.Picture1.DrawWidth = 1
         mapgraphfm.Picture1.Line (xcord1, ycord)-(mapgraphfm.Picture1.Width, ycord), QBColor(0)
         End If
      mapgraphfm.Picture1.Refresh
      mapgraphfm.Refresh
      DoEvents
      End If
   Screen.MousePointer = vbDefault
   
Exit Sub

errhand:
    If ermk% = 1 Then
       Screen.MousePointer = vbDefault
       response = MsgBox("Error Number: " & Err.Number & " encountered." & vbLf & _
              Err.Description & vbLf & _
              "Resume read operation?", vbYesNo, "Maps&More")
       If response = vbYes Then
          Err.Clear
          GoTo 100
       Else
          Err.Clear
          Exit Sub
          End If
       End If
       
End Sub



Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
   Unload mapgraphfm
   Set mapgraphfm = Nothing
   graphwind = False
   killpicture = True
   If crosssection = True Then
      'erase the cross section line
      crosssection = False
      Call blitpictures
      mapCrossSection.Visible = True
      mapCrossSection.txtNumPoints.SetFocus
      End If
End Sub


Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      On Error GoTo errhand
      Label9.Visible = True
      Shape1.Visible = True
      Line1.Refresh
      MSFlexGrid1.Enabled = True
      Label9.Top = Y - 100
      Shape1.Top = Y - 100
      If X + 100 + Label9.Width < mapgraphfm.Picture1.Width Then
         Label9.Left = X + 200
         Shape1.Left = X + 200
      Else
         Label9.Left = X - 100 - Label9.Width
         Shape1.Left = X - 100 - Shape1.Width
         End If
      If crosssection = False Then
        Label9.Caption = "azi = " + Format(((X / mapgraphfm.Picture1.Width) * (xmax - xmin) + xmin), "##0.0") + _
                         ", va = " + Format(((Y / mapgraphfm.Picture1.Height) - 1) * (ymin - ymax) + ymin, "#0.0#")
        Combo1.ListIndex = ((X / mapgraphfm.Picture1.Width) * (xmax - xmin) + xmin - xmino) * 10
        'If world = True Then
        '   Combo1.ListIndex = ((x / mapgraphfm.Picture1.Width) * (xmax - xmin) + xmin + 80) * 10
        'Else
        '   Combo1.ListIndex = ((x / mapgraphfm.Picture1.Width) * (xmax - xmin) - xmino) * 10
        '   End If
      Else
        Label9.Width = 2400
        Shape1.Width = Label9.Width
        Label9.Caption = "dis = " + Format(((X / mapgraphfm.Picture1.Width) * (xmax - xmin) + xmin), "####0.000") + _
                         ", hgt = " + Format(((Y / mapgraphfm.Picture1.Height) - 1) * (ymin - ymax) + ymin, "###0.0")
        Combo1.ListIndex = ((X / mapgraphfm.Picture1.Width)) * ((xmax - xmin) / (xmaxo - xmino)) * sectnumpnt& + ((xmin - xmino) / (xmaxo - xmino)) * sectnumpnt&
        End If
      
      'If crosssection = False And world = True Then
      '  MSFlexGrid1.row = 1
      '  MSFlexGrid1.col = 0
      '  MSFlexGrid1.Text = RTrim$(LTrim$(Mid$(Combo1.Text, 1, 5)))
      '  MSFlexGrid1.col = 1
      '  MSFlexGrid1.Text = RTrim$(Trim$(Mid$(Combo1.Text, 7, 9)))
      '  MSFlexGrid1.col = 2
      '  MSFlexGrid1.Text = RTrim$(LTrim$(Mid$(Combo1.Text, 18, 9)))
      '  MSFlexGrid1.col = 3
      '  MSFlexGrid1.Text = RTrim$(LTrim$(Mid$(Combo1.Text, 31, 9)))
      '  MSFlexGrid1.col = 4
      '  MSFlexGrid1.Text = RTrim$(LTrim$(Mid$(Combo1.Text, 42, 9)))
      '  MSFlexGrid1.col = 5
      '  MSFlexGrid1.Text = RTrim$(LTrim$(Mid$(Combo1.Text, 51, 10)))
      'Else
        If Not crosssection Then conv = 1#
        If crosssection Then conv = 0.001
        MSFlexGrid1.row = 1
        MSFlexGrid1.col = 0
        posit1% = InStr(1, Combo1.Text, ",")
        MSFlexGrid1.Text = RTrim$(LTrim$(Mid$(Combo1.Text, 1, posit1% - 1)))
        MSFlexGrid1.col = 1
        posit2% = InStr(posit1% + 1, Combo1.Text, ",")
        MSFlexGrid1.Text = RTrim$(Trim$(Mid$(Combo1.Text, posit1% + 1, posit2% - posit1% - 1)))
        MSFlexGrid1.col = 2
        posit1% = posit2%
        posit2% = InStr(posit1% + 1, Combo1.Text, ",")
        MSFlexGrid1.Text = RTrim$(Trim$(Mid$(Combo1.Text, posit1% + 1, posit2% - posit1% - 1)))
        MSFlexGrid1.col = 3
        posit1% = posit2%
        posit2% = InStr(posit1% + 1, Combo1.Text, ",")
        MSFlexGrid1.Text = RTrim$(Trim$(Mid$(Combo1.Text, posit1% + 1, posit2% - posit1% - 1)))
        MSFlexGrid1.col = 4
        posit1% = posit2%
        posit2% = InStr(posit1% + 1, Combo1.Text, ",")
        MSFlexGrid1.Text = Str(Val(RTrim$(Trim$(Mid$(Combo1.Text, posit1% + 1, posit2% - posit1% - 1)))) * conv)
        MSFlexGrid1.col = 5
        MSFlexGrid1.Text = RTrim$(Trim$(Mid$(Combo1.Text, posit2% + 1, Len(Combo1.Text) - posit2%)))
      '  End If
      If Button = 1 And dragbegin = True Then
         mapgraphfm.Picture1.DrawMode = 7
         mapgraphfm.Picture1.DrawStyle = vbDot
         mapgraphfm.Picture1.DrawWidth = 1
         mapgraphfm.Picture1.Line (drag1x, drag1y)-(drag2x, drag2y), QBColor(15), B
         mapgraphfm.Picture1.Line (drag1x, drag1y)-(X, Y), QBColor(15), B
         drag2x = X
         drag2y = Y
         End If
         
Exit Sub
errhand:
   Resume Next
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label9.Visible = False
   Shape1.Visible = False
   Line1.Refresh
   Combo1.Text = sEmpty
    MSFlexGrid1.row = 1
    MSFlexGrid1.col = 0
    MSFlexGrid1.Text = sEmpty
    MSFlexGrid1.col = 1
    MSFlexGrid1.Text = sEmpty
    MSFlexGrid1.col = 2
    MSFlexGrid1.Text = sEmpty
    MSFlexGrid1.col = 3
    MSFlexGrid1.Text = sEmpty
    MSFlexGrid1.col = 4
    MSFlexGrid1.Text = sEmpty
    MSFlexGrid1.col = 5
    MSFlexGrid1.Text = sEmpty
   MSFlexGrid1.Enabled = False
End Sub


Private Sub restorelimitsbut_Click()

   On Error GoTo errhand:
   
   Screen.MousePointer = vbHourglass
   mapgraphfm.Picture1.DrawMode = 13
   mapgraphfm.Picture1.DrawWidth = 2

'   If crosssection = True Then
'       plotfile$ = drivjk$ + "crossect.tmp"
'   Else
'       openfile$ = drivjk$ + "eros.tmp"
'       If world = False Then
'          openfile$ = drivjk$ + "EYisroel.tmp"
'          End If
'       End If
   
   myfile = Dir(plotfile$)
   If myfile <> sEmpty Then
      xmin = xmino
      xmax = xmaxo
      ymin = ymino
      ymax = ymaxo
      filtmp% = FreeFile
      Label8.Caption = "ymax = " + Format(ymax, "#0.0#")
      Label8.Visible = True
      ymax = CInt(ymax + 0.5)
      mapgraphfm.Picture1.Cls
      mapgraphfm.Label6.Caption = Format(ymin, "#0.0#")
      mapgraphfm.Label5.Caption = Format(ymax, "#0.0#")
      mapgraphfm.Label3.Caption = Format(xmin, "#0.0")
      mapgraphfm.Label4.Caption = Format(xmax, "#0.0")
      If crosssection = True Then
         mapgraphfm.Label4.Caption = Format(xmax, "###0.000")
         End If
      Open plotfile$ For Input As #filtmp%
      Line Input #filtmp%, doclin$
      Line Input #filtmp%, doclin$
      For i% = 1 To nnpnt%
         Input #filtmp%, azi, va, logi, lati, dista, hgti
         If crosssection = False Then
            xcord = ((azi - xmin) / (xmax - xmin)) * mapgraphfm.Picture1.Width
            ycord = mapgraphfm.Picture1.Height - ((va - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
         Else
            xcord = ((dista * 0.001 - xmin) / (xmax - xmin)) * mapgraphfm.Picture1.Width
            ycord = mapgraphfm.Picture1.Height - ((hgti - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
            End If
         If i% = 1 Then
            mapgraphfm.Picture1.PSet (xcord, ycord), QBColor(1)
            xo = xcord
            yo = ycord
         Else
            mapgraphfm.Picture1.Line (xo, yo)-(xcord, ycord), QBColor(1)
            xo = xcord
            yo = ycord
            End If
      Next i%
      Close #filtmp%
      If ymin < 0 Then
         xcord1 = 0
         ycord = mapgraphfm.Picture1.Height - ((0 - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
         mapgraphfm.Picture1.DrawStyle = vbDot
         mapgraphfm.Picture1.DrawWidth = 1
         mapgraphfm.Picture1.Line (xcord1, ycord)-(mapgraphfm.Picture1.Width, ycord), QBColor(0)
         End If
      End If
   Screen.MousePointer = vbDefault
  Exit Sub
  
errhand:
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub TimeZonebut_Click()
   'ret = Shell("c:\windows\system\timedate.cpl", 1)
   ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   ret = SetWindowPos(mapgraphfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   Call mouse_event(MOUSEEVENTF_MOVE, 10000, 10000, 0, 0)  'move mouse to Location item
   waitime = Timer
   Do Until Timer > waitime + 1
   Loop
   Call mouse_event(MOUSEEVENTF_MOVE, -10, -10, 0, 0)  'move mouse to Location item
   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
   Call mouse_event(MOUSEEVENTF_MOVE, -100, 100, 0, 0)  'move mouse to Location item
'   Call mouse_event(VK_SHIFT + VK_TAB, 0, 0, 0, 0)
'   waitime = Timer
'   Do Until Timer > waitime + 1
'   Loop
'   Call mouse_event(VK_RIGHT, 0, 0, 0, 0)
End Sub
Private Sub picture1_mousedown(Button As Integer, _
   Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then 'maybe beginning of drag operation
      drag1x = X
      drag1y = Y
      dragbegin = True
      drag2x = drag1x
      drag2y = drag1y
      End If
   End Sub
Private Sub picture1_mouseup(Button As Integer, _
   Shift As Integer, X As Single, Y As Single)
   On Error GoTo errhand
   If Button = 1 And (drag1x <> drag2x Or drag1y <> drag2y) Then
      dragbegin = False
      drag2x = X
      drag2y = Y
      'erase final line and refresh graph with new limits
      mapgraphfm.Picture1.Line (drag1x, drag1y)-(drag2x, drag2y), QBColor(15), B
      
      'determine new x,y mins and maximuns
      xminnew = (drag1x / mapgraphfm.Picture1.Width) * (xmax - xmin) + xmin
      xmaxnew = (drag2x / mapgraphfm.Picture1.Width) * (xmax - xmin) + xmin
      yminnew = ((drag2y / mapgraphfm.Picture1.Height) - 1) * (ymin - ymax) + ymin
      ymaxnew = ((drag1y / mapgraphfm.Picture1.Height) - 1) * (ymin - ymax) + ymin
      If yminnew > ymaxnew Then
         yy = yminnew
         yminnew = ymaxnew
         ymaxnew = yy
         End If
      If xminnew > xmaxnew Then
         yy = xminnew
         xminnew = xmaxnew
         xmaxnew = yy
         End If
      'xmino = xmin
      'xmaxo = xmax
      'ymino = ymin
      'ymaxo = ymax
      
      xmin = xminnew
      xmax = xmaxnew
      ymin = yminnew
      ymax = ymaxnew
      'refresh graph with new limits
        Screen.MousePointer = vbHourglass
        mapgraphfm.Picture1.Cls
        mapgraphfm.Picture1.DrawMode = 13
        mapgraphfm.Picture1.DrawStyle = vbDefault
        mapgraphfm.Picture1.DrawWidth = 2
        'openfile$ = P
        'If crosssection = True Then
        '   plotfile$ = drivjk$ + "crossect.tmp"
        '   End If
        myfile = Dir(plotfile$)
        If myfile <> sEmpty Then
           filtmp% = FreeFile
           Open plotfile$ For Input As #filtmp%
           Line Input #filtmp%, doclin$
           Line Input #filtmp%, doclin$
           Label7.Caption = "ymin = " + Format(yminnew, "#0.0#")
           Label7.Visible = True
           Label8.Caption = "ymax = " + Format(ymaxnew, "#0.0#")
           Label8.Visible = True
           'ymax = CInt(ymax + 0.5)
           mapgraphfm.Label6.Caption = Format(yminnew, "#0.0#")
           mapgraphfm.Label5.Caption = Format(ymaxnew, "#0.0#")
           mapgraphfm.Label3.Caption = Format(xminnew, "##0.0")
           mapgraphfm.Label4.Caption = Format(xmaxnew, "##0.0")
           If crosssection = True Then
              mapgraphfm.Label3.Caption = Format(xminnew, "###0.000")
              mapgraphfm.Label4.Caption = Format(xmaxnew, "###0.000")
              End If
           Seek #filtmp%, 1
           Line Input #filtmp%, doclin$
           Line Input #filtmp%, doclin$
           For i% = 1 To nnpnt%
              Input #filtmp%, azi, va, logi, lati, dista, hgti
              If crosssection = False Then
                xcord = ((azi - xmin) / (xmax - xmin)) * mapgraphfm.Picture1.Width
                ycord = mapgraphfm.Picture1.Height - ((va - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
              Else
                xcord = ((dista * 0.001 - xmin) / (xmax - xmin)) * mapgraphfm.Picture1.Width
                ycord = mapgraphfm.Picture1.Height - ((hgti - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
                End If
              If i% = 1 Then
                 mapgraphfm.Picture1.PSet (xcord, ycord), QBColor(1)
                 xo = xcord
                 yo = ycord
               Else
                  mapgraphfm.Picture1.Line (xo, yo)-(xcord, ycord), QBColor(1)
                  xo = xcord
                  yo = ycord
                  End If
           Next i%
           Close #filtmp%
           If yminnew < 0 Then
              xcord1 = 0
              ycord = mapgraphfm.Picture1.Height - ((0 - ymin) / (ymax - ymin)) * mapgraphfm.Picture1.Height
              mapgraphfm.Picture1.DrawStyle = vbDot
              mapgraphfm.Picture1.DrawWidth = 1
              mapgraphfm.Picture1.Line (xcord1, ycord)-(mapgraphfm.Picture1.Width, ycord), QBColor(0)
              End If
           End If
        Screen.MousePointer = vbDefault
      End If
   'reprocess graph with new limits
   Exit Sub
errhand:
   Screen.MousePointer = vbDefault
End Sub


