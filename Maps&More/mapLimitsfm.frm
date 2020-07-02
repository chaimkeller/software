VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form mapLimitsfm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analysis Limits"
   ClientHeight    =   9435
   ClientLeft      =   6885
   ClientTop       =   1815
   ClientWidth     =   4725
   Icon            =   "mapLimitsfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmTemp 
      Caption         =   "Refraction-Temperature -Modeling"
      Height          =   515
      Left            =   120
      TabIndex        =   47
      Top             =   3480
      Width           =   4455
      Begin VB.ComboBox cmbModelTemp 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   48
         Text            =   "cmbModelTemp"
         ToolTipText     =   "Choose how to deal with terrestrial refraction"
         Top             =   180
         Width           =   3975
      End
   End
   Begin VB.CheckBox chkautorange 
      Caption         =   "Auto azi. range"
      Height          =   195
      Left            =   2520
      TabIndex        =   46
      ToolTipText     =   "check to automatically determine appropriate azimuth range"
      Top             =   1000
      Width           =   1815
   End
   Begin VB.Frame frmIgnoreTiles 
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   3000
      Width           =   4515
      Begin VB.CheckBox chkIgnoreTiles 
         Caption         =   "Ignore any missing tiles"
         Height          =   195
         Left            =   1320
         TabIndex        =   45
         ToolTipText     =   "Check to ignore any missing tiles"
         Top             =   200
         Width           =   2055
      End
   End
   Begin VB.Frame frmProfile 
      Caption         =   "View Angle vs. Azimuth Profiles"
      ForeColor       =   &H00400000&
      Height          =   1680
      Left            =   120
      TabIndex        =   35
      Top             =   7080
      Width           =   4575
      Begin VB.Frame frmrderos2 
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   4095
         Begin VB.OptionButton optrderos2_2 
            Caption         =   "Use rderos2"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2400
            TabIndex        =   43
            ToolTipText     =   "calculate profile using the program rderos2"
            Top             =   180
            Width           =   1335
         End
         Begin VB.OptionButton optrderos2_1 
            Caption         =   "Don't use rderos2"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   42
            ToolTipText     =   "calculate profile using only c++ simplified emulation of rderos2"
            Top             =   200
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin MSComCtl2.UpDown UpDown5 
         Height          =   375
         Left            =   3120
         TabIndex        =   39
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtAngRes"
         BuddyDispid     =   196618
         OrigLeft        =   3360
         OrigTop         =   600
         OrigRight       =   3600
         OrigBottom      =   975
         Max             =   50
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtAngRes 
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
         Left            =   2640
         TabIndex        =   37
         Text            =   "1"
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox chkProfile 
         Caption         =   "Calculate Profiles during DTM data extraction"
         Height          =   255
         Left            =   560
         TabIndex        =   36
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblAngResScroll 
         Caption         =   "x 0.01 degrees"
         Height          =   375
         Left            =   3480
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblAngRes 
         Caption         =   "Azimuth Step Size of View Angle vs. Azimuth Profile Scan"
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.Frame frmFileOnly 
      Height          =   735
      Left            =   3240
      TabIndex        =   33
      Top             =   2280
      Width           =   1400
      Begin VB.CheckBox chkFileOnly 
         Caption         =   "Just extract hieghts without analysis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.UpDown UpDown4 
      Height          =   375
      Left            =   3240
      TabIndex        =   24
      Top             =   5835
      Width           =   195
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   2
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text4"
      BuddyDispid     =   196633
      OrigLeft        =   3600
      OrigTop         =   5100
      OrigRight       =   3795
      OrigBottom      =   5475
      Max             =   100
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Set the extent"
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
      Left            =   960
      TabIndex        =   22
      ToolTipText     =   "Use user defined extent number (2-10)"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Text            =   "1.2"
      Top             =   5840
      Width           =   615
   End
   Begin VB.Frame frmRadar 
      Caption         =   "SRTM Radar Shadow"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   6440
      Width           =   4515
      Begin VB.CheckBox chkRadar 
         Caption         =   "Smooth SRTM radar shadows/voids"
         Enabled         =   0   'False
         Height          =   195
         Left            =   840
         TabIndex        =   32
         ToolTipText     =   "Use linear fit to remove SRTM voids"
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame frmDTM 
      Caption         =   "DTM source"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   35
      TabIndex        =   26
      Top             =   0
      Width           =   4735
      Begin VB.OptionButton optSRTM30 
         Caption         =   "SRTM30 (1 km)"
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
         Left            =   1000
         TabIndex        =   30
         ToolTipText     =   "SRTM/GTOPO30 30 arcsec DTM"
         Top             =   240
         Value           =   -1  'True
         Width           =   1300
      End
      Begin VB.OptionButton optSRTM2 
         Caption         =   "SRTM-1 (90m)"
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
         Left            =   2280
         TabIndex        =   29
         ToolTipText     =   "SRTM 3 arcsec DTM"
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optSRTM1 
         Caption         =   "SRTM-2 (30m)"
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
         Left            =   3480
         TabIndex        =   28
         ToolTipText     =   "SRTM 1 arcsec DTM"
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optGTOPO30 
         Caption         =   "GTOPO30"
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
         Left            =   80
         TabIndex        =   27
         ToolTipText     =   "GTOPO30 30 arcsec DTM"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   23
      Text            =   "2"
      Top             =   5880
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Frame Frame2 
      Caption         =   "3D Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2300
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   4515
      Begin VB.OptionButton Option6 
         Caption         =   "View BOTH horizons;  (1/extent) *full size"
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
         Left            =   180
         TabIndex        =   21
         Top             =   1320
         Value           =   -1  'True
         Width           =   4215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "View only ONE horizon with MINIMUM extent"
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
         Left            =   180
         TabIndex        =   20
         Top             =   1020
         Width           =   4275
      End
      Begin VB.OptionButton Option4 
         Caption         =   "View only ONE horizon with MEDIUM extent"
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
         Left            =   180
         TabIndex        =   19
         Top             =   660
         Width           =   4215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "View only ONE horizon with MAXIMUM extent"
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
         Left            =   180
         TabIndex        =   18
         Top             =   360
         Width           =   4275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "Extract DTM for both horizons"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Extract DTM for only one horizon"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Accept && &Record"
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
      Left            =   120
      TabIndex        =   13
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
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
      Left            =   2760
      TabIndex        =   12
      Top             =   9000
      Width           =   1815
   End
   Begin MSComCtl2.UpDown UpDown3 
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   1800
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   4
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text3"
      BuddyDispid     =   196644
      OrigLeft        =   3720
      OrigTop         =   1560
      OrigRight       =   3960
      OrigBottom      =   1935
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Left            =   2520
      TabIndex        =   8
      Text            =   "4"
      Top             =   1800
      Width           =   975
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   4
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text2"
      BuddyDispid     =   196645
      OrigLeft        =   3720
      OrigTop         =   960
      OrigRight       =   3960
      OrigBottom      =   1335
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   2520
      TabIndex        =   4
      Text            =   "4"
      Top             =   1320
      Width           =   975
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   550
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   80
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text1"
      BuddyDispid     =   196646
      OrigLeft        =   3000
      OrigTop         =   360
      OrigRight       =   3240
      OrigBottom      =   735
      Max             =   80
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   2520
      TabIndex        =   0
      Text            =   "80"
      Top             =   550
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4800
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   4800
      Y1              =   4070
      Y2              =   4070
   End
   Begin VB.Label Label6 
      Caption         =   "degrees"
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
      Left            =   3840
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Difference between beglat and endlat (for any one horizon)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "degrees"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Difference between beglog and endlog (for any one horizon)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "degrees"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Maximum half azmiuth range"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "mapLimitsfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim record As Boolean

Private Sub Check1_Click()
   If Check1.value = vbChecked Then
      Text4.Enabled = True
      Text5.Enabled = True
      UpDown4.Enabled = True
      If DTMflag <= 0 Then
         Text5 = modeval
      Else
         Text5 = modevals
         End If
   ElseIf Check1.value = vbUnchecked Then
      Text4.Enabled = False
      Text5.Enabled = False
      UpDown4.Enabled = False
      Text4 = 0
      End If
End Sub

Private Sub chkautorange_Click()
   If chkautorange.value = vbChecked Then
      autoazirange% = 1
   Else
      autoazirange% = 0
      End If
End Sub

Private Sub chkFileOnly_Click()
   If chkFileOnly.value = vbChecked Then
      OnlyExtractFile = True
   Else
      OnlyExtractFile = False
      End If
End Sub

Private Sub chkIgnoreTiles_Click()
    If chkIgnoreTiles.value = vbChecked Then
        IgnoreTiles% = 1
    Else
        IgnoreTiles% = 0
        End If
End Sub

Private Sub chkProfile_Click()
   If chkProfile.value = vbUnchecked Then
      CalculateProfile = 0
      frmrderos2.Enabled = False
      optrderos2_1.Enabled = False
      optrderos2_2.Enabled = False
   Else
      CalculateProfile = 1
      frmrderos2.Enabled = True
      optrderos2_1.Enabled = True
      optrderos2_2.Enabled = True
      End If
End Sub

Private Sub chkRadar_Click()
   If noVoidflag = 1 Then
      noVoidflag = 0
   ElseIf noVoidflag = 0 Then
      noVoidflag = 1
      End If
End Sub

Private Sub cmbModelTemp_Change()
   TemperatureModel% = cmbModelTemp.ListIndex - 1
End Sub

Private Sub cmbModelTemp_Click()
   TemperatureModel% = cmbModelTemp.ListIndex - 1
End Sub

'Private Sub chkTK_Click()
'   If chkTK.value = vbChecked Then
'      TemperatureModel% = 1
'   Else
'      TemperatureModel% = 0
'      End If
'End Sub
Private Sub Command1_Click()
   If DTMflag <= 0 Then 'GTOPO30
      If Val(Text1) <> 0 Then maxang% = Text1
      If Val(Text2) <> 0 Then diflog% = Text2
      If Val(Text3) <> 0 Then diflat% = Text3
      modeval = Val(Text5)
   Else
      If Val(Text1) <> 0 Then maxangs% = Text1
      If Val(Text2) <> 0 Then diflogs% = Text2
      If Val(Text3) <> 0 Then diflats% = Text3
      modevals = Val(Text5)
      End If
   If record = True Then
     TemperatureModel% = cmbModelTemp.ListIndex - 1
     maxangf% = maxang%
     diflogf% = diflog%
     diflatf% = diflat%
     fullrangef% = fullrange%
     viewmodef% = viewmode%
     modevalf = modeval
     maxangfs% = maxangs%
     diflogfs% = diflogs%
     diflatfs% = diflats%
     fullrangefs% = fullranges%
     viewmodefs% = viewmodes%
     modevalfs = modevals
     filnum% = FreeFile
     AziStep% = Val(txtAngRes)
     AziStepf% = AziStep%
     Open drivjk$ + "mapposition.sav" For Output As #filnum%
     'If hgtpos = sEmpty Then hgtpos = 0
     f1 = 1: If kmxsky > 1000 Then f1 = 0.001
     f2 = 0: f3 = 0: If kmysky > 1000 Then f2 = 1000000: f3 = 0.001
     Write #filnum%, kmxsky * f1, (kmysky - f2) * f3, hgtpos
     If Maps.Text7.Text = sEmpty Then
        hgtworld = 0
     Else
        hgtworld = Maps.Text7.Text
        End If
     Write #filnum%, lon, lat, hgtworld
     Write #filnum%, maxangf%, diflogf%, diflatf%, fullrangef%, viewmodef%, modevalf
     Write #filnum%, DTMflag
     Write #filnum%, maxangfs%, diflogfs%, diflatfs%, fullrangefs%, viewmodefs%, modevalfs
     Write #filnum%, CalculateProfile
     Write #filnum%, AziStepf%
     Write #filnum%, rderos2_use
     Write #filnum%, IgnoreTiles%
     Write #filnum%, autoazirange%
     Write #filnum%, TemperatureModel%
     Close #filnum%
     End If
     
   Call form_queryunload(i%, j%)
End Sub

Private Sub Command2_Click()
   record = True
   Command1_Click
End Sub

'maxang%, fullrange%, diflat%, diflog%
Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set mapLimitsfm = Nothing
End Sub
Private Sub form_load()
   record = False
   
   If DTMflag = -1 Then
      optGTOPO30.value = True
   ElseIf DTMflag = 0 Then
      optSRTM1.value = True
      optSRTM30.value = True
      frmRadar.Enabled = False
      chkRadar.Enabled = False
      chkRadar.value = vbUnchecked
      noVoidflag = 0
   ElseIf DTMflag = 1 Then
      optSRTM1.value = True
      frmRadar.Enabled = True
      chkRadar.Enabled = True
      If noVoidflag = 1 Then
         chkRadar.value = vbChecked
         noVoidflag = 1
         End If
   ElseIf DTMflag = 2 Then
      optSRTM2.value = True
      frmRadar.Enabled = True
      chkRadar.Enabled = True
      If noVoidflag = 1 Then
         chkRadar.value = vbChecked
         noVoidflag = 1
         End If
      End If
      
  If CalculateProfile Then
     chkProfile.value = vbChecked
  Else
     chkProfile.value = vbUnchecked
     End If
     
  textAngresol = AziStep
  
  If autoazirange% = 1 Then
     chkautorange.value = vbChecked
  Else
     chkautorange.value = vbUnchecked
     End If
  
  If rderos2_use Then
    optrderos2_2.value = True
  Else
    optrderos2_1.value = True
    End If
    
  txtAngRes.Text = AziStepf%
  
  If IgnoreTiles% = 1 Then
     chkIgnoreTiles.value = vbChecked
  ElseIf IgnoreTiles% = 0 Then
     chkIgnoreTiles.value = vbUnchecked
     End If
  
'  If TemperatureModel% = 1 Then
'     chkTK.value = vbChecked
'  ElseIf TemperatureModel% = 0 Then
'     chkTK.value = vbUnchecked
'     End If

   With cmbModelTemp
    .AddItem "Old terrestrial refraction model"
    .AddItem "No terrestrial refraction modeling"
    .AddItem "TR based on Avgerage Temp: don't remove from profile"
    .AddItem "TR based on Avgerage Temp: remove from profile"
   End With
   
   If TemperatureModel% = -1 Then
      cmbModelTemp.ListIndex = 0
   ElseIf TemperatureModel% = 0 Then
      cmbModelTemp.ListIndex = 1
   ElseIf TemperatureModel% = 1 Then
      cmbModelTemp.ListIndex = 2
   ElseIf TemperatureModel% = 2 Then
      cmbModelTemp.ListIndex = 3
   Else
      cmbModelTemp.ListIndex = 1
      End If
      
End Sub

Private Sub optGTOPO30_Click()
        
    DTMflag = -1
    If maxang% <> 0 Then Text1 = maxang%
    If diflog% <> 0 Then Text2 = diflog%
    If diflat% <> 0 Then Text3 = diflat%
    If fullrange% = 0 Then
       Option1.value = True
    ElseIf fullrange% = 1 Then
       Option2.value = True
       End If
    If viewmode% = 0 Then
       Option3.value = True
    ElseIf viewmode% = 1 Then
       Option4.value = True
    ElseIf viewmode% = 2 Then
       Option5.value = True
    ElseIf viewmode% = 3 Then
       Option6.value = True
       End If
    Text5.Text = modeval
    If modeval <> 0 Then
       Check1.value = vbChecked
       End If
    DTMflag = -1
    frmRadar.Enabled = False
    chkRadar.Enabled = False
    chkRadar.value = vbUnchecked
    noVoidflag = 0

End Sub

Private Sub Option1_Click()
   If DTMflag <= 0 Then 'GTOPO30/SRTM30
      fullrange% = 0
      If viewmode% = 3 Then
         Option4.value = True
         viewmode% = 1
         End If
      Option3.Enabled = True
      Option4.Enabled = True
      Option5.Enabled = True
      Option6.Enabled = False
      Check1.Enabled = True
    Else 'SRTM
      fullranges% = 0
      If viewmodes% = 3 Then
         Option4.value = True
         viewmodes% = 1
         End If
      Option3.Enabled = True
      Option4.Enabled = True
      Option5.Enabled = True
      Option6.Enabled = False
      Check1.Enabled = True
      End If
End Sub

Private Sub Option2_Click()
   If DTMflag <= 0 Then 'both horizons
     fullrange% = 1
     viewmode% = 3
     Option3.Enabled = False
     Option4.Enabled = False
     Option5.Enabled = False
     Option6.Enabled = True
     Option6.value = True
     Check1.Enabled = True
   Else 'SRTM
     fullranges% = 1
     viewmodes% = 3
     Option3.Enabled = False
     Option4.Enabled = False
     Option5.Enabled = False
     Option6.Enabled = True
     Option6.value = True
     Check1.Enabled = True
     End If
End Sub

Private Sub Option3_Click()
   If DTMflag <= 0 Then viewmode% = 0
   If DTMflag > 0 Then viewmodes% = 0
End Sub

Private Sub Option4_Click()
   If DTMflag <= 0 Then viewmode% = 1
   If DTMflag > 0 Then viewmodes% = 1
End Sub

Private Sub Option5_Click()
   If DTMflag <= 0 Then viewmode% = 2
   If DTMflag > 0 Then viewmodes% = 2
End Sub

Private Sub Option6_Click()
   If DTMflag <= 0 Then 'GTOPO30/SRTM30
      viewmode% = 3
      Option2.value = True
      fullrange% = 1
   Else 'SRTM
      viewmodes% = 3
      Option2.value = True
      fullranges% = 1
      End If
End Sub

Private Sub optrderos2_1_Click()
    rderos2_use = False
End Sub

Private Sub optrderos2_2_Click()
    rderos2_use = True
End Sub

Private Sub optSRTM1_Click()
        
    DTMflag = 1
    If maxangs% <> 0 Then Text1 = maxangs%
    If diflogs% <> 0 Then Text2 = diflogs%
    If diflats% <> 0 Then Text3 = diflats%
    If fullranges% = 0 Then
       Option1.value = True
    ElseIf fullranges% = 1 Then
       Option2.value = True
       End If
    If viewmodes% = 0 Then
       Option3.value = True
    ElseIf viewmodes% = 1 Then
       Option4.value = True
    ElseIf viewmodes% = 2 Then
       Option5.value = True
    ElseIf viewmodes% = 3 Then
       Option6.value = True
       End If
    Text5.Text = modevals
    If modevals <> 0 Then
       Check1.value = vbChecked
       End If
    DTMflag = 1
    frmRadar.Enabled = True
    chkRadar.Enabled = True

End Sub

Private Sub optSRTM2_Click()
        
    DTMflag = 2
    If maxangs% <> 0 Then Text1 = maxangs%
    If diflogs% <> 0 Then Text2 = diflogs%
    If diflats% <> 0 Then Text3 = diflats%
    If fullranges% = 0 Then
       Option1.value = True
    ElseIf fullranges% = 1 Then
       Option2.value = True
       End If
    If viewmodes% = 0 Then
       Option3.value = True
    ElseIf viewmodes% = 1 Then
       Option4.value = True
    ElseIf viewmodes% = 2 Then
       Option5.value = True
    ElseIf viewmodes% = 3 Then
       Option6.value = True
       End If
    Text5.Text = modevals
    If modevals <> 0 Then
       Check1.value = vbChecked
       End If
    DTMflag = 2
    frmRadar.Enabled = True
    chkRadar.Enabled = True

End Sub

Private Sub optSRTM30_Click()
    
    DTMflag = 0
    If maxang% <> 0 Then Text1 = maxang%
    If diflog% <> 0 Then Text2 = diflog%
    If diflat% <> 0 Then Text3 = diflat%
    If fullrange% = 0 Then
       Option1.value = True
    ElseIf fullrange% = 1 Then
       Option2.value = True
       End If
    If viewmode% = 0 Then
       Option3.value = True
    ElseIf viewmode% = 1 Then
       Option4.value = True
    ElseIf viewmode% = 2 Then
       Option5.value = True
    ElseIf viewmode% = 3 Then
       Option6.value = True
       End If
    Text5.Text = modeval
    If modeval <> 0 Then
       Check1.value = vbChecked
       End If
    DTMflag = 0
    frmRadar.Enabled = False
    chkRadar.Enabled = False
    chkRadar.value = vbUnchecked
    noVoidflag = 0

End Sub

Private Sub Text4_Change()
   Text5.Text = 1# + 0.2 * Val(Text4.Text)
End Sub
