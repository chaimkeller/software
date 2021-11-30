VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SunriseSunset 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sunrise/Sunset"
   ClientHeight    =   7965
   ClientLeft      =   3480
   ClientTop       =   540
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SunriseSunset.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   4815
   Begin VB.CommandButton OKbut0 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   240
      Picture         =   "SunriseSunset.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CheckBox chkObst 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add larger cushion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   28
      ToolTipText     =   "Add larger cushions depending on how close the obstruction is"
      Top             =   2740
      Width           =   1695
   End
   Begin VB.CheckBox chkOldCalcMethod 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   360
      TabIndex        =   27
      ToolTipText     =   "use oold (netzski6.exe) calculation method"
      Top             =   5280
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2460
      TabIndex        =   24
      Top             =   3100
      Width           =   1875
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "calendar in &english"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   26
         Top             =   240
         Width           =   1755
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "calendar in he&brew"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   25
         Top             =   0
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   240
      Left            =   1680
      TabIndex        =   21
      Top             =   3480
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      _Version        =   393216
      Value           =   5
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text2"
      BuddyDispid     =   196614
      OrigLeft        =   1680
      OrigTop         =   3360
      OrigRight       =   1920
      OrigBottom      =   3615
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   0   'False
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   20
      Text            =   "5"
      Top             =   3480
      Width           =   330
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   240
      Left            =   1680
      TabIndex        =   17
      Top             =   3120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      _Version        =   393216
      Value           =   5
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text1"
      BuddyDispid     =   196615
      OrigLeft        =   1680
      OrigTop         =   3120
      OrigRight       =   1920
      OrigBottom      =   3375
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   0   'False
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   16
      Text            =   "5"
      Top             =   3080
      Width           =   330
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mis. Su&nset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   2520
      TabIndex        =   15
      ToolTipText     =   "Ast sunset at hgt = 0"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Mis. Sunrise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   "Ast. sunrise at hgt = 0"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ast. S&unset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "Ast sunset incl. height"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ast. Sunrise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Ast. sunrise incl. height"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   2490
      TabIndex        =   9
      Top             =   1800
      Width           =   2000
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Hebrew Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   10
         Top             =   0
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Civil Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   11
         Top             =   280
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   540
      TabIndex        =   5
      Top             =   5880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton Cancelbut 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      Picture         =   "SunriseSunset.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Green fro Near Obs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "Print times from near obstructions in green"
      Top             =   2500
      Width           =   1755
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3960
      Top             =   5280
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "SunriseSunset.frx":0CC6
      Left            =   3000
      List            =   "SunriseSunset.frx":0CC8
      TabIndex        =   3
      Top             =   2600
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vis. &Sunset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "visible sunset"
      Top             =   2100
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Vis. Sunrise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "visible sunrise"
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   2280
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "%"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1980
      TabIndex        =   23
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Max allowed near obstruc."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   420
      TabIndex        =   22
      Top             =   3400
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "km"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2000
      TabIndex        =   19
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "min allowed dis to obstru."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   420
      TabIndex        =   18
      Top             =   3000
      Width           =   915
   End
   Begin VB.Shape Shape4 
      Height          =   675
      Left            =   2400
      Top             =   4000
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      Height          =   675
      Left            =   240
      Top             =   4000
      Width           =   2055
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   4440
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Shape Shape2 
      Height          =   2175
      Left            =   240
      Top             =   1740
      Width           =   2055
   End
   Begin VB.Line Line7 
      X1              =   2400
      X2              =   2400
      Y1              =   1920
      Y2              =   2880
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   2400
      Top             =   1740
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   240
      Picture         =   "SunriseSunset.frx":0CCA
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   4200
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   4560
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hebrew Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   2350
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"SunriseSunset.frx":2F6C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1620
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4350
   End
End
Attribute VB_Name = "SunriseSunset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancelbut_Click()
   Screen.MousePointer = vbDefault
   Close
   astronplace = False
   geo = False
   SunriseSunset.Visible = False
   If eroscityflag = True Then
      Unload SunriseSunset
      Exit Sub
      End If
   Caldirectories.Label1.Enabled = True
   Caldirectories.Drive1.Enabled = True
   Caldirectories.Dir1.Enabled = True
   'Caldirectories.List1.Enabled = True
   Caldirectories.Text1.Enabled = True
   Caldirectories.OKbutton.Enabled = True
   Caldirectories.ExitButton.Enabled = True
   Caldirectories.OKbutton.Enabled = True
   'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   eros = False
End Sub


Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      Check4.Value = vbUnchecked
      Check5.Value = vbUnchecked
      Check6.Value = vbUnchecked
      Check7.Value = vbUnchecked
      End If
   If Check1.Value = vbUnchecked And Check2.Value = vbUnchecked Then
      Check3.Value = vbUnchecked
      chkObst.Value = vbUnchecked
      End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      Check4.Value = vbUnchecked
      Check5.Value = vbUnchecked
      Check6.Value = vbUnchecked
      Check7.Value = vbUnchecked
      End If
   If Check1.Value = vbUnchecked And Check2.Value = vbUnchecked Then
      Check3.Value = vbUnchecked
      chkObst.Value = vbUnchecked
      End If
End Sub

Private Sub Check3_Click()
   If chkObst.Value = vbChecked Then
      chkObst.Value = vbUnchecked
      AddObsTime = 0
      End If
   If Check1.Value = vbUnchecked And Check2.Value = vbUnchecked Then
      Check3.Value = vbUnchecked
      End If
   If Check3.Value = vbChecked Then
      Label4.Enabled = True
      Label5.Enabled = True
      Label6.Enabled = True
      Label7.Enabled = True
      Text1.Enabled = True
      Text2.Enabled = True
      UpDown1.Enabled = True
      UpDown2.Enabled = True
      If eros = True Then
         If SRTMflag = 0 Then
            Text1 = obsdistlim(1) '30 'distant horizon = 30 km
            distlim = obsdistlim(1)
            outdistlim = obsdistlim(1)
            obscushion = cushion(1)
         ElseIf SRTMflag = 1 Then
            If InStr(eroscountry$, "USA") Then
               Text1 = obsdistlim(3) '6 'distant horizon = 6 km'
               distlim = obsdistlim(3)
               outdistlim = obsdistlim(3)
               obscushion = cushion(3)
               '<<<<<<<<<<<<<needs updating>>>>>>>>>>>>
               'really should be that for NED, then text1 = 6
               'and older SRTM files should have text1 = 10
            ElseIf InStr(eroscountry$, "Israel") Then
               Text1 = obsdistlim(0)
               distlim = obsdistlim(0)
               outdistlim = obsdistlim(0)
               obscushion = cushion(0)
            Else
               Text1 = 10 'SRTM is worst than NED
               distlim = 10
               outdistlim = 10
               obscushion = 30
               End If
         ElseIf SRTMflag = 2 Or geotz! <> 2 Then
            Text1 = obsdistlim(3) '6 '18 'distant horizon = 18 km
            distlim = obsdistlim(3)
            outdistlim = obsdistlim(3)
            obscushion = cushion(3)
         ElseIf SRTMflag = 9 And geotz! = 2 Then
            Text1 = obsdistlim(0) '5 'Eretz Yisroel DTM
            distlim = obsdistlim(0)
            outdistlim = obsdistlim(0)
            obscushion = cushion(0)
            End If
      Else
         SRTMflag = 9
         If eros = True And eroscountry$ = "Israel" Then
            geotz! = 2
            End If
         If geotz! = 2 Or eros = False Then
            Text1 = obsdistlim(0) '5 'Eretz Yisroel DTM
            distlim = obsdistlim(0)
            outdistlim = obsdistlim(0)
            obscushion = cushion(0)
         Else
            Text1 = obsdistlim(3) '6 '18 'distant horizon = 18 km
            distlim = obsdistlim(3)
            outdistlim = obsdistlim(3)
            obscushion = cushion(3)
            End If
         End If
    Else
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      Text1.Enabled = False
      Text2.Enabled = False
      UpDown1.Enabled = False
      UpDown2.Enabled = False
      End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = vbChecked Then
      Check1.Value = vbUnchecked
      Check2.Value = vbUnchecked
      Check3.Value = vbUnchecked
      chkObst.Value = vbUnchecked
      Check6.Value = vbUnchecked
      Check7.Value = vbUnchecked
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      Text1.Enabled = False
      Text2.Enabled = False
      UpDown1.Enabled = False
      UpDown2.Enabled = False
      End If
End Sub

Private Sub Check5_Click()
   If Check5.Value = vbChecked Then
      Check1.Value = vbUnchecked
      Check2.Value = vbUnchecked
      Check3.Value = vbUnchecked
      chkObst.Value = vbUnchecked
      Check6.Value = vbUnchecked
      Check7.Value = vbUnchecked
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      Text1.Enabled = False
      Text2.Enabled = False
      UpDown1.Enabled = False
      UpDown2.Enabled = False
      End If
End Sub
Private Sub Check6_Click()
   If Check6.Value = vbChecked Then
      Check1.Value = vbUnchecked
      Check2.Value = vbUnchecked
      Check3.Value = vbUnchecked
      chkObst.Value = vbUnchecked
      Check4.Value = vbUnchecked
      Check5.Value = vbUnchecked
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      Text1.Enabled = False
      Text2.Enabled = False
      UpDown1.Enabled = False
      UpDown2.Enabled = False
      End If
End Sub
Private Sub Check7_Click()
   If Check7.Value = vbChecked Then
      Check1.Value = vbUnchecked
      Check2.Value = vbUnchecked
      Check3.Value = vbUnchecked
      chkObst.Value = vbUnchecked
      Check4.Value = vbUnchecked
      Check5.Value = vbUnchecked
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      Text1.Enabled = False
      Text2.Enabled = False
      UpDown1.Enabled = False
      UpDown2.Enabled = False
      End If
End Sub

Private Sub chkObst_Click()
   If Check3.Value = vbChecked Then
      Check3.Value = vbUnchecked
      End If
   If Check1.Value = vbUnchecked And Check2.Value = vbUnchecked Then
      chkObst.Value = vbUnchecked
      End If
   If chkObst.Value = vbChecked Then
      Label4.Enabled = True
      Label5.Enabled = True
      Label6.Enabled = True
      Label7.Enabled = True
      Text1.Enabled = True
      Text2.Enabled = True
      UpDown1.Enabled = True
      UpDown2.Enabled = True
      If eros = True Then
         If SRTMflag = 0 Then
            Text1 = obsdistlim(1) '30 'distant horizon = 30 km
            distlim = obsdistlim(1)
            outdistlim = obsdistlim(1)
            obscushion = cushion(1)
         ElseIf SRTMflag = 1 Then
            If InStr(eroscountry$, "USA") Then
               Text1 = obsdistlim(3) '6 'distant horizon = 6 km'
               distlim = obsdistlim(3)
               outdistlim = obsdistlim(3)
               obscushion = cushion(3)
               '<<<<<<<<<<<<<needs updating>>>>>>>>>>>>
               'really should be that for NED, then text1 = 6
               'and older SRTM files should have text1 = 10
            ElseIf InStr(eroscountry$, "Israel") Then
               Text1 = obsdistlim(0)
               distlim = obsdistlim(0)
               outdistlim = obsdistlim(0)
               obscushion = cushion(0)
            Else
               Text1 = 10 'SRTM is worst than NED
               distlim = 10
               outdistlim = 10
               obscushion = 30
               End If
         ElseIf eros = True And eroscountry$ <> "Israel" Then
            Text1 = obsdistlim(3) '6 '18 'distant horizon = 18 km
            distlim = obsdistlim(3)
            outdistlim = obsdistlim(3)
            obscushion = cushion(3)
         ElseIf SRTMflag = 9 Or (eros = True And eroscountry$ = "Israel") Then
            Text1 = obsdistlim(0) '5 'Eretz Yisroel DTM
            distlim = obsdistlim(0)
            outdistlim = obsdistlim(0)
            obscushion = cushion(0)
            End If
      Else
         SRTMflag = 9
         If eros = True And eroscountry$ <> "Israel" Then
            Text1 = obsdistlim(3) '6 '18 'distant horizon = 18 km
            distlim = obsdistlim(3)
            outdistlim = obsdistlim(3)
            obscushion = cushion(3)
         ElseIf eros = True And eroscountry$ = "Israel" Then
            Text1 = obsdistlim(0) '5 'Eretz Yisroel DTM
            distlim = obsdistlim(0)
            outdistlim = obsdistlim(0)
            obscushion = cushion(0)
         Else
            Text1 = obsdistlim(0) 'Eretz Yisroel cushions, etc is default
            distlim = obsdistlim(0)
            outdistlim = obsdistlim(0)
            obscushion = cushion(0)
            End If
         End If
    Else
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      Text1.Enabled = False
      Text2.Enabled = False
      UpDown1.Enabled = False
      UpDown2.Enabled = False
      End If
      
   AddObsTime = 1
      
End Sub

Private Sub Form_Load()
   'version: 11/21/2008
    
    'visible times are default
    Check1.Value = vbChecked
    Check2.Value = vbChecked
    Check4.Value = vbUnchecked
    Check5.Value = vbUnchecked
    Check6.Value = vbUnchecked
    Check7.Value = vbUnchecked
    
    If Caldirectories.Runbutton.Enabled Then
       If visauto Then 'visible times
         'use defaults
       ElseIf mishorauto Then 'mishor times
         Check1.Value = vbUnchecked
         Check2.Value = vbUnchecked
         Check4.Value = vbUnchecked
         Check5.Value = vbUnchecked
         Check6.Value = vbChecked
         Check7.Value = vbChecked
       ElseIf astauto Then 'astronomical times
         Check1.Value = vbUnchecked
         Check2.Value = vbUnchecked
         Check4.Value = vbChecked
         Check5.Value = vbChecked
         Check6.Value = vbUnchecked
         Check7.Value = vbUnchecked
         End If
       End If
       
   If AddObsTime = 1 Then
     chkObst.Value = vbChecked
     Check3.Value = vbChecked
   Else
     Check1.Value = vbChecked
     chkObst.Value = vbUnchecked
     End If
    
   'Label3.Visible = False

   If Combo1.Text = sEmpty Then hebcal = True
   For i% = RefHebYear% To 6000
      Combo1.AddItem (Trim$(Str$(i%)))
   Next i%
   If yrheb% <> 0 And hebcal = True Then
      Combo1.ListIndex = yrheb% - RefHebYear% '5758 '1
   ElseIf hebcal = True Then
      'find current hebrew year
      finddate$ = Date$
      cha$ = sEmpty
      lenyr% = 0
      Do Until cha$ = "-"
        cha$ = Mid$(finddate$, Len(finddate$) - lenyr%, 1)
        lenyr% = lenyr% + 1
      Loop
      presyr% = Val(Mid$(finddate$, Len(finddate$) - lenyr% + 2, lenyr - 1))
      yrheb% = presyr% - RefCivilYear% + RefHebYear% '1997 + 5758
      Combo1.ListIndex = yrheb% - RefHebYear% '5758
      End If
   captmp$ = Label1.Caption
   hebcal = True
   Screen.MousePointer = vbDefault
End Sub
Private Sub OKbut0_Click()
   Dim erosfil%, myfile ',entrylin$(50)
   Dim netzlist$(2500), skiylist$(2500)
   Dim distnez(6, 601), distski(6, 601), distcheck(601)
   Dim dblEndTime As Double, CDpromcheck As Boolean
   Dim Coords() As String, kmxAT As Double, kmyAT As Double
   Dim ltAT As Double, lgAT As Double
   Dim MinTK(12) As Integer, AvgTK(12) As Integer, MaxTK(12) As Integer, ier As Integer
   
   On Error GoTo OKbuterhand
   
50:
   'read in hebrew strings
   myfile = Dir(drivjk$ + "calhebrew.txt")
   If myfile = sEmpty Then
      response = MsgBox("Can't find the hebrew strings: calhebrew.txt! You won't be able to obtain hebrew tables.", vbOKCancel + vbExclamation, "Cal Program")
      If response = vbCancel Then
         Exit Sub
      Else
         Option4.Value = True
         optionheb = False
         End If
   Else
      hebnumb% = FreeFile
      Open drivjk$ + "calhebrew.txt" For Input As #hebnumb%
      iheb% = 0
      Do Until EOF(hebnumb%)
        Line Input #hebnumb%, doclin$
        
        If InStr(doclin$, "[sunrisesunset]") <> 0 Then
           flgheb% = 0
        ElseIf InStr(doclin$, "[newhebcalfm]") <> 0 Then
           flgheb% = 1
           iheb% = 0
        ElseIf InStr(doclin$, "[zmanlistfm]") <> 0 Then
           flgheb% = 2
           iheb% = 0
        ElseIf InStr(doclin$, "[hebweek]") <> 0 Then
           flgheb% = 3
           iheb% = 0
        ElseIf InStr(doclin$, "[holiday_dates]") <> 0 Then
           flgheb% = 4
           iheb% = 0
        ElseIf InStr(doclin$, "[candle_lighting]") <> 0 Then
           flgheb% = 5
           iheb% = 0
           End If
        
        Select Case flgheb%
           Case 0
             heb1$(iheb%) = doclin$
           Case 1
             heb2$(iheb%) = doclin$
           Case 2
             heb3$(iheb%) = doclin$
           Case 3
             heb4$(iheb%) = doclin$
           Case 4
             heb5$(iheb%) = doclin$
           Case 5
             heb6$(iheb%) = doclin$
        End Select
        
        iheb% = iheb% + 1
      Loop
      Close #hebnumb%
      End If
      
    'now read sponsorship information
    SponsorLine$ = sEmpty
    myfile = Dir("c:\inetpub\webpub\SponsorLogo.txt")
    If myfile <> sEmpty Then
       filtmp% = FreeFile
       Open "c:\inetpub\webpub\SponsorLogo.txt" For Input As #filtmp%
       Do Until EOF(filtmp%)
          Input #filtmp%, doclin$
          doclin$ = Trim$(doclin$)
          
          If InStr(doclin$, "[English]") <> 0 And Not optionheb Then
             'english sponsor info
             Input #filtmp%, doclin$
             SponsorLine$ = doclin$
             End If
          
          If InStr(doclin$, "[titles]") <> 0 And Not optionheb Then
             'english tables' title info
             If Not eros And Not astronplace And Not ast Then 'Eretz Yisroel tables
                Input #filtmp%, doclin$
                TitleLine$ = doclin$
             Else 'world tables or astronomical calculations
                Input #filtmp%, doclin$
                Input #filtmp%, doclin$
                Input #filtmp%, doclin$
                TitleLine$ = doclin$
                End If
             End If
          
          If InStr(doclin$, "[Hebrew]") <> 0 And optionheb Then
             'hebrew sponsor info
             Input #filtmp%, doclin$
             SponsorLine$ = doclin$
             End If
       
          If InStr(doclin$, "[titles]") <> 0 And optionheb Then
             'hebrew tables' title info
             If Not eros And Not astronplace And Not ast Then 'Eretz Yisroel tables
                Input #filtmp%, doclin$
                Input #filtmp%, doclin$
             Else 'tables for the world or astronomical tables
                Input #filtmp%, doclin$
                Input #filtmp%, doclin$
                Input #filtmp%, doclin$
                Input #filtmp%, doclin$
                End If
             TitleLine$ = doclin$
             heb3$(1) = heb1$(1) & " " & Chr$(34) & TitleLine$ & Chr$(34) & " " & Mid$(heb3$(1), 1, 1)
             heb1$(4) = TitleLine$
             heb3$(12) = heb1$(1) & " " & Chr$(34) & TitleLine$ & Chr$(34) & " " & Mid$(heb3$(12), 10, 15)
             End If
       
       Loop
       Close #filtmp%
    Else 'sponsor file not found, use defaults
       If optionheb Then
          If eros Then
             TitleLine$ = heb1$(2)
          Else
             TitleLine$ = heb1$(4)
             End If
       ElseIf Not optionheb Then
          If eros Then
             TitleLine$ = "Chai"
          Else
             TitleLine$ = "Bikurei Yosef"
             End If
          End If
       
       End If
   
   nearnez = False
   nearski = False
   ast = False: vis = False: mis = False
   
   If Check6.Value = 1 And Check7.Value = 0 Then
      mis = True
      mis0% = 1
   ElseIf Check6.Value = 0 And Check7.Value = 1 Then
      mis = True
      mis0% = 2
   ElseIf Check6.Value = 1 And Check7.Value = 1 Then
      mis = True
      mis0% = 3
      End If
   If Check4.Value = 1 And Check5.Value = 0 Then
      ast = True
      ast0% = 1
   ElseIf Check4.Value = 0 And Check5.Value = 1 Then
      ast = True
      ast0% = 2
   ElseIf Check4.Value = 1 And Check5.Value = 1 Then
      ast = True
      ast0% = 3
      End If

   If Check1.Value = 1 And Check2.Value = 0 Then
      vis = True
      vis0% = 1
   ElseIf Check1.Value = 0 And Check2.Value = 1 Then
      vis = True
      vis0% = 2
   ElseIf Check1.Value = 1 And Check2.Value = 1 Then
      vis = True
      vis0% = 3
      End If
      
  ProgExec$ = "Netzski6"
  distlim = Val(SunriseSunset.Text1.Text)
  If Check3.Value = vbChecked Then
     distlim = Val(SunriseSunset.Text1.Text)
     nearnez = True 'check for near obstructions
     nearski = True
     nearcolor = True 'donote near obstructions with color
  Else
    nearnez = False
    nearski = False
    nearcolor = False
    End If
      
   If optionheb = True Then 'hebrew captions
     If eros = True Or astronplace = True Then 'And currentdir = "d:\cities\eros\visual_tmp" Then
         'title$ = "לוח " + Chr$(34) + "חי" + Chr$(34)
         'title$ = heb1$(1) + Chr$(34) + heb1$(2) + Chr$(34)
         title$ = heb1$(1) + Chr$(34) + TitleLine$ + Chr$(34)
     Else
         'title$ = "לוח " + Chr$(34) + "בכורי יוסף" + Chr$(34)
         'title$ = heb1$(3) + Chr$(34) + heb1$(4) + Chr$(34)
         title$ = heb1$(3) + Chr$(34) + TitleLine$ + Chr$(34)
         End If
   Else 'english captions
     If eros = True Or astronplace = True Then 'And currentdir = "d:\cities\eros\visual_tmp" Then
         title$ = TitleLine$ & " Tables" '"Chai Tables"
     Else
         title$ = TitleLine$ & " Tables" '"Bikurei Yosef Tables"
         End If
     End If
     
   CalcMethod% = 0 'for years between 1950-2050 use approximate ephemerels '<<<<restore to = 0 after debugging
   If Option1b = True And (Val(Combo1.Text) > 5810 Or Val(Combo1.Text) < 5710) Then
      CalcMethod% = 1 'must use SunCoo calculation of solar epehemerels
   ElseIf Option2b = True And ((Val(Combo1.Text) < 1950 And Val(Combo1.Text) >= 1600) Or Val(Combo1.Text) > 2050) Then
      CalcMethod% = 1 'must use SunCoo calculation of solar epehemerels
      End If
      
'   If Option1b = True And (Val(Combo1.Text) > 5810 Or Val(Combo1.Text) < 5710) Then
'      If internet = True Then
'         'write warning to calprog.log
'
'         lognum% = FreeFile
'         Open drivjk$ + "calprog.log" For Append As #lognum%
'         Print #lognum%, " "
'         Print #lognum%, "******************************************************************"
'         Print #lognum%, "Warning! User downloaded tables for year: " & Trim$(Combo1.Text)
'         Print #lognum%, "Astronomical constants only accurate to year 5810--update them!"
'         Print #lognum%, "*******************************************************************"
'         Print #lognum%, " "
'         Close #lognum%
'         response = vbYes
'         GoTo ssN:
'         End If
'
'      response = MsgBox("Astronomical constants are good to 6 seconds only for the years 1950-2050.  They must be updated for other years!  Do you still want to continue?", vbQuestion + vbYesNo, "Cal Program")
'ssN:  If response = vbNo Then
'         OKbut(0).Value = False
'         If yrheb% <> 0 Then 'convert from hebrew year to civil year
'            Combo1.ListIndex = yrheb% + RefCivilYear% - RefHebYear% - 1600 '1997 - 5758 - 1600
'         Else 'use current year
'            finddate$ = Date$
'            cha$ = sEmpty
'            lenyr% = 0
'            Do Until cha$ = "-"
'              cha$ = Mid$(finddate$, Len(finddate$) - lenyr%, 1)
'              lenyr% = lenyr% + 1
'            Loop
'            Combo1.ListIndex = Val(Mid$(finddate$, Len(finddate$) - lenyr% + 2, lenyr - 1)) - 1600
'            End If
'         End If
'   'ElseIf hebcal = True And Val(Combo1.Text) < 5758 Then
'   '   response = MsgBox("The permissible range of hebrew years is: 5758-6000", vbExclamation + vbOKOnly, "Cal Program")
'   '   Exit Sub
'   ElseIf Option2b = True And ((Val(Combo1.Text) < 1950 And Val(Combo1.Text) >= 1600) Or Val(Combo1.Text) > 2050) Then
'      response = MsgBox("Astronomical constants are good to 6 seconds only for the years 1950-2050.  They must be updated for other years!  Do you still want to continue?", vbQuestion + vbYesNo, "Cal Program")
'      If response = vbNo Then
'         OKbut(0).Value = False
'        If yrheb% <> 0 Then 'convert from hebrew year to civil year
'           Combo1.ListIndex = yrheb% + RefCivilYear% - RefHebYear% - 1600 ' 1997 - 5758 - 1600
'        Else 'use current year
'           finddate$ = Date$
'           cha$ = sEmpty
'           lenyr% = 0
'           Do Until cha$ = "-"
'             cha$ = Mid$(finddate$, Len(finddate$) - lenyr%, 1)
'             lenyr% = lenyr% + 1
'           Loop
'           Combo1.ListIndex = Val(Mid$(finddate$, Len(finddate$) - lenyr% + 2, lenyr - 1)) - 1600
'           End If
'         Exit Sub
'         End If
'   ElseIf Option2b = True And Val(Combo1.Text) < 1600 Then
   If Option2b = True And Val(Combo1.Text) < 1600 Then
      response = MsgBox("Permissable range of civil years: 1600 and onward", vbOKOnly + vbExclamation, "Cal Program")
      OKbut0.Value = False
      If yrheb% <> 0 Then 'convert from hebrew year to civil year
         Combo1.ListIndex = yrheb% + RefCivilYear% - RefHebYear% - 1600 '1997 - 5758 - 1600
      Else 'use current year
         finddate$ = Date$
         cha$ = sEmpty
         lenyr% = 0
         Do Until cha$ = "-"
           cha$ = Mid$(finddate$, Len(finddate$) - lenyr%, 1)
           lenyr% = lenyr% + 1
         Loop
         Combo1.ListIndex = Val(Mid$(finddate$, Len(finddate$) - lenyr% + 2, lenyr - 1)) - 1600
         End If
      Exit Sub
      End If
   nstat% = 0
   ntmp% = 0
   warn% = 0
   Check1.Enabled = False
   Check2.Enabled = False
   Check3.Enabled = False
   Check4.Enabled = False
   Check5.Enabled = False
   Check6.Enabled = False
   Check7.Enabled = False
   Combo1.Enabled = False
   Cancelbut.Enabled = False
   Label2.Enabled = False
   Label4.Enabled = False
   Label5.Enabled = False
   Label6.Enabled = False
   Label7.Enabled = False
   Text1.Enabled = False
   UpDown1.Enabled = False
   Text2.Enabled = False
   UpDown2.Enabled = False
   Label1.Caption = " Copying files...Please wait."
   Label1.Refresh
   Timer1.Enabled = False
   OKbut0.Enabled = False
   nearyesval = False
   If Option1b = True Then
      Option2.Enabled = False
      Option1.Enabled = False
   Else
      Option1.Enabled = False
      Option2.Enabled = False
      End If
   If optionheb = True Then
      Option4.Enabled = False
      Option3.Enabled = False
   Else
      Option3.Enabled = False
      Option4.Enabled = False
      End If
   If automatic = False Then
      yrheb% = Val(Combo1.Text)
   Else
      If Option1b = True Then Combo1.ListIndex = yrheb% - 1
      End If
   If Option1b = True Then
'      stryr% = (yrheb% - 5758) + 1997
'      endyr% = (yrheb% - 5758) + 1998
      GoSub lpyr   'determine if it is hebrew leap year
                   'and determine the beginning and ending daynumbers of the Hebrew year
      stryr% = yrheb% + RefCivilYear% - RefHebYear% '(yrheb% - 5758) + 1997
      endyr% = yrheb% + RefCivilYear% - RefHebYear% + 1 '(yrheb% - 5758) + 1998
      yrstrt%(1) = 1
   Else
      Option1.Value = False
      GoSub lpyrcivil 'determine if it is civil leap year
                      'and determine the beginning and ending daynumbers of the Civil year
      stryr% = yrheb%
      endyr% = yrheb%
      yrstrt%(0) = 1
      yrstrt%(1) = 1
      yrend%(0) = yl%
      yrend%(1) = yl%
      End If
   If Check1.Value = vbChecked And Check2.Value = vbUnchecked Then
      nsetflag% = 1
      portrait = False
      newhebcalfm.SSTab1.Tab = 0
      newhebcalfm.SSTab1.TabEnabled(0) = True
      newhebcalfm.SSTab1.TabEnabled(1) = False
   ElseIf Check1.Value = vbUnchecked And Check2.Value = vbChecked Then
      nsetflag% = 2
      portrait = False
      newhebcalfm.SSTab1.Tab = 1
      newhebcalfm.SSTab1.TabEnabled(1) = True
      newhebcalfm.SSTab1.TabEnabled(0) = False
   ElseIf Check1.Value = vbChecked And Check2.Value = vbChecked Then
      nsetflag% = 3
      portrait = True
      newhebcalfm.SSTab1.Tab = 0
      newhebcalfm.SSTab1.TabEnabled(0) = True
      newhebcalfm.SSTab1.TabEnabled(1) = True
      End If
   If (Check4.Value = vbChecked And Check5.Value = vbUnchecked) Or _
      (Check6.Value = vbChecked And Check7.Value = vbUnchecked) Then
      nsetflag% = -1
      portrait = False
      newhebcalfm.SSTab1.Tab = 0
      newhebcalfm.SSTab1.TabEnabled(0) = True
      newhebcalfm.SSTab1.TabEnabled(1) = False
   ElseIf (Check4.Value = vbUnchecked And Check5.Value = vbChecked) Or _
          (Check6.Value = vbUnchecked And Check7.Value = vbChecked) Then
      nsetflag% = -2
      portrait = False
      newhebcalfm.SSTab1.Tab = 1
      newhebcalfm.SSTab1.TabEnabled(1) = True
      newhebcalfm.SSTab1.TabEnabled(0) = False
   ElseIf (Check4.Value = vbChecked And Check5.Value = vbChecked) Or _
          (Check6.Value = vbChecked And Check7.Value = vbChecked) Then
      nsetflag% = -3
      portrait = True
      newhebcalfm.SSTab1.Tab = 0
      newhebcalfm.SSTab1.TabEnabled(0) = True
      newhebcalfm.SSTab1.TabEnabled(1) = True
      End If
   If portrait = True Then
     newhebcalfm.Shape6.Left = newhebcalfm.Image1.Left - 115
     newhebcalfm.Text15.Text = sEmpty
     newhebcalfm.Text15.Text = "Portrait Orientation (Vertical)" '"Print two calendars one above the other"
   ElseIf portrait = False Then
     newhebcalfm.Shape6.Left = newhebcalfm.Image2.Left - 130
     newhebcalfm.Text15.Text = sEmpty
     newhebcalfm.Text15.Text = "Landscape Orientation (Horizontal)" '"Print one calendar lengthwise"
     End If
   If Katz = True Then
     newhebcalfm.Check3.Value = vbChecked
     End If
   
   Call readpaper 'read in appropriate paper sizes
   Call readfont 'read in appropriate default fonts
   If Loadcombo% = 0 Then
      Loadcombo% = 1
      For i% = 1 To numpaper%
         Pageformatfm.Combo1.AddItem papername$(i%)
         newhebcalfm.Combo11.AddItem papername$(i%)
      Next i%
      End If
   Pageformatfm.Combo1.ListIndex = prespap% - 1
   newhebcalfm.Combo11.ListIndex = prespap% - 1
   If hebcal = False Then
      newhebcalfm.Check4.Enabled = False
      newhebcalfm.Check4.Value = vbUnchecked
      newhebcalfm.Option2.Value = True
      newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: civil calendar"
   Else
      newhebcalfm.Check4.Enabled = True
      newhebcalfm.Option1.Value = True
      If hebleapyear = True Then
         newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: " + "hebrew/leapyear"
      Else
         newhebcalfm.Text42.Text = "paper: " + papername$(prespap%) + "; font file: " + "hebrew/regular year"
         End If
      End If
   
   If astronplace = True Then 'use inputed name as city's hebrew name
      citnam$ = "ast"
      currentdir = drivcities$ + "ast"
      hebcityname$ = astname$
      'check for USA
'      If InStr(hebcityname$, "ארהב") <> 0 Then 'add " to USA's hebrew name
'         hebcityname$ = Mid$(astname$, 1, Len(astname$) - 4) + "ארה" + Chr$(34) + "ב"
'         End If
      If InStr(hebcityname$, heb1$(5)) <> 0 Then 'add " to USA's hebrew name
         hebcityname$ = Mid$(astname$, 1, Len(astname$) - 4) + heb1$(6) + Chr$(34) + heb1$(7)
         End If
      GoTo 100
      End If
   If eros = True Then 'read bat file to determine the hebcityname and zone time
      erosfil% = FreeFile
      If SunriseSunset.Check1.Value = vbChecked Or SunriseSunset.Check4.Value = vbChecked Or SunriseSunset.Check6.Value = vbChecked Then
         myfile = Dir(currentdir + "\netz\*.bat")
         If myfile <> sEmpty Then
            Open currentdir + "\netz\" + myfile For Input As #erosfil%
            Input #erosfil%, hebcityname$, geotz!
            Close #erosfil%
         Else
            If internet = True Then GoTo errinternet
            response = MsgBox("Can't find eros bat file!", vbCritical + vbOKOnly, "Cal Program")
            Cancelbut_Click
            Exit Sub
            End If
      ElseIf SunriseSunset.Check2.Value = vbChecked Or SunriseSunset.Check5.Value = vbChecked Or SunriseSunset.Check7.Value = vbChecked Then
         myfile = Dir(currentdir + "\skiy\*.bat")
         If myfile <> sEmpty Then
            Open currentdir + "\skiy\" + myfile For Input As #erosfil%
            Input #erosfil%, hebcityname$, geotz!
            Close #erosfil%
         Else
            If internet = True Then GoTo errinternet
            response = MsgBox("Can't find eros bat file!", vbCritical + vbOKOnly, "Cal Program")
            Cancelbut_Click
            Exit Sub
            End If
         End If
      End If
   
   If astronplace = True Or eros = True Then GoTo 75
   'find chosen city in city list and use it to define the default top caption
   lencit% = Len(currentdir)
   citnam$ = sEmpty
   For i% = lencit% To 1 Step -1
      cha$ = Mid$(currentdir, i%, 1)
      If cha$ <> "\" Then
         citnam$ = cha$ + citnam$
      Else
         Exit For
         End If
   Next i%
   For i% = 1 To numcities%
      'If i% = 371 Then
      '   cc = 1
      '   End If

      If LCase$(citnam$) = LCase$(citynames$(i%)) Then
         citnamp$ = ""
         nextflg% = 0
         For jlen% = 1 To Len(citnam$)
           cha$ = Mid(citnam$, jlen%, 1)
           If jlen% = 1 Then
              cha$ = UCase(cha$)
              End If
           If nextflg% = 1 Then
              cha$ = UCase(cha$)
              nextflg% = 0
              End If
           If cha$ = "_" Then
              cha$ = " "
              nextflg% = 1
              End If
           If cha$ = "-" Then nextflg% = 1
           citnamp$ = citnamp$ & cha$
         Next jlen%
         If optionheb = True Then
            hebcityname$ = cityhebnames$(i%)
'            newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה הנראה ב" + cityhebnames$(i%)
'            compare1$ = title$ + " לזמני הנץ החמה הנראה ב" + cityhebnames$(i%)
'            newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה הנראית ב" + cityhebnames$(i%)
'            compare2$ = title$ + " לזמני שקיעת החמה הנראית ב" + cityhebnames$(i%)
            newhebcalfm.Combo1.AddItem title$ + heb1$(8) + cityhebnames$(i%)
            compare1$ = title$ + heb1$(8) + cityhebnames$(i%)
            newhebcalfm.Combo6.AddItem title$ + heb1$(9) + cityhebnames$(i%)
            compare2$ = title$ + heb1$(9) + cityhebnames$(i%)
         Else
            newhebcalfm.Combo1.AddItem title$ + " of the Visible Sunrise Times for " + citnamp$ 'UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
            compare1$ = title$ + " of the Visible Sunrise Times for " + citnamp$ 'UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
            newhebcalfm.Combo6.AddItem title$ + " of the Visible Sunset Times for " + citnamp$ '+ UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
            compare2$ = title$ + " of the Visible Sunset Times for " + citnamp$ 'UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
            End If
         Exit For
         End If
   Next i%

75 If eros = True And currentdir$ <> drivcities$ & "eros\visual_tmp" Then
'      newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה הנראה ב" + cityhebnames$(i%)
'      compare1$ = title$ + " לזמני הנץ החמה הנראה ב" + hebcityname$
'      newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה הנראית ב" + cityhebnames$(i%)
'      compare2$ = title$ + " לזמני שקיעת החמה הנראית ב" + hebcityname$

      newhebcalfm.Combo1.AddItem title$ + heb1$(8) + cityhebnames$(i%)
      compare1$ = title$ + heb1$(8) + hebcityname$
      newhebcalfm.Combo6.AddItem title$ + heb1$(9) + cityhebnames$(i%)
      compare2$ = title$ + heb1$(9) + hebcityname$
      
       If optionheb = False Then
         newhebcalfm.Combo1.AddItem title$ + " of the Visible Sunrise Times for " + citnamp$ 'UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
         compare1$ = title$ + " of the Visible Sunrise Times for " + citnamp$ 'UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
         newhebcalfm.Combo6.AddItem title$ + " of the Visible Sunset Times for " + citnamp$ ' UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
         compare2$ = title$ + " of the Visible Sunset Times for " + citnamp$ 'UCase(Mid(citynames$(i%), 1, 1)) + Mid(citynames$(i%), 2, Len(citynames$(i%)) - 1)
         End If
     ' If SunriseSunset.Check4.Value = vbChecked Then
     '    newhebcalfm.Combo1.AddItem title$ + heb1$(41) + cityhebnames$(i%)
     '    compare1$ = title$ + heb1$(41) + hebcityname$
     '    End If
     ' If SunriseSunset.Check5.Value = vbChecked Then
     '    newhebcalfm.Combo6.AddItem title$ + heb1$(42) + cityhebnames$(i%)
     '    compare2$ = title$ + heb1$(42) + hebcityname$
     '    End If
     ' If SunriseSunset.Check6.Value = vbChecked Then
     '    newhebcalfm.Combo1.AddItem title$ + heb1$(43) + cityhebnames$(i%)
     '    compare1$ = title$ + heb1$(43) + hebcityname$
     '    End If
     ' If SunriseSunset.Check7.Value = vbChecked Then
     '    newhebcalfm.Combo6.AddItem title$ + heb1$(44) + cityhebnames$(i%)
     '    compare2$ = title$ + heb1$(44) + hebcityname$
     '    End If
   ElseIf eros = True And currentdir$ = drivcities$ & "eros\visual_tmp" Then
'      newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה הנראה ב" + cityhebnames$(i%)
'      compare1$ = title$ + " לזמני הנץ החמה הנראה באיזור " + hebcityname$
'      newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה הנראית ב" + cityhebnames$(i%)
'      compare2$ = title$ + " לזמני שקיעת החמה הנראית באיזור " + hebcityname$
      newhebcalfm.Combo1.AddItem title$ + heb1$(10) + cityhebnames$(i%)
      compare1$ = title$ + heb1$(10) + hebcityname$
      newhebcalfm.Combo6.AddItem title$ + heb1$(11) + cityhebnames$(i%)
      compare2$ = title$ + heb1$(11) + hebcityname$
      
      If IsraelNeighborhood And optionheb Then
        hebcityname$ = eroshebcity$
        compare1$ = title$ + heb1$(10) + eroshebcity$
        compare2$ = title$ + heb1$(11) + eroshebcity$
        End If
      
      If optionheb = False Then
         'parse the english name
         citnamp$ = erosareabat
         'remove the country's name
         For ij% = Len(citnamp$) To 1 Step -1
             If Mid$(citnamp$, ij%, 1) = "_" Then
                citnamp$ = Mid$(citnamp$, 1, ij% - 1)
                Exit For
                End If
         Next ij%
         For ij% = 1 To Len(citnamp$)
            If Mid(citnamp$, ij%, 1) = "_" Then
               Mid(citnamp$, ij%, 1) = " "
               End If
         Next ij%
         'remove the word "area"
         pos% = InStr(citnamp$, "area")
         If pos% <> 0 Then
            citnamp$ = Mid$(citnamp$, 1, pos% - 1) & Mid$(citnamp$, pos% + 4, Len(citnamp$) - pos% - 3)
            Mid(citnamp$, pos% - 1, 1) = ","
            End If
         newhebcalfm.Combo1.AddItem title$ + " of the Visible Sunrise for" + citnamp$ 'cityhebnames$(i%)
         compare1$ = title$ + " of the Visible Sunrise for the region of " + citnamp$
         newhebcalfm.Combo6.AddItem title$ + " of the Visible Sunset for" + citnamp$ 'cityhebnames$(i%)
         compare2$ = title$ + " of the Visible Sunset for the region of " + citnamp$
         End If
      End If
  'now load in other captions if there is save file

100 If portrait = True Then
      suffix$ = "_port_w1255.sav"
   ElseIf portrait = False Then
      suffix$ = "_land.sav"
      End If
   myname = Dir(currentdir + "\" + citnam$ + suffix$)
   If myname <> sEmpty And optionheb = True Then
      filsav% = FreeFile
      Open currentdir + "\" + citnam$ + suffix$ For Input As #filsav%
      Do Until EOF(filsav%)
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo1.AddItem doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo2.AddItem doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo3.AddItem doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo4.AddItem doclin$
         'If portrait = False And doclin$ <> "רצוי לוודא זמנים אלה ע" + Chr$(34) + "י תצפיות" Then
         '   newhebcalfm.Combo4.AddItem "רצוי לוודא זמנים אלה ע" + Chr$(34) + "י תצפיות"
         '   End If
         If portrait = False And doclin$ <> heb1$(14) + Chr$(34) + heb1$(15) Then
            newhebcalfm.Combo4.AddItem heb1$(14) + Chr$(34) + heb1$(15)
            End If
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo5.AddItem doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Text1.Text = doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Text2.Text = doclin$
         
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo6.AddItem doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then 'first check for old spelling bug in ".sav" files
            'If doclin$ = "מבוסס על השקיעה המאוחרת ביותר הנראה מעל האופק המערבי האמיתי, מנקודה כלשהי ביישוב - מידי יום ביומו" Then
            '   doclin$ = "מבוסס על השקיעה המאוחרת ביותר הנראית מעל האופק המערבי האמיתי, מנקודה כלשהי ביישוב - מידי יום ביומו"
            '   End If
            If doclin$ = heb1$(12) Then
               doclin$ = heb1$(13)
               End If
            newhebcalfm.Combo7.AddItem doclin$
            End If
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo8.AddItem doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo9.AddItem doclin$
         'If portrait = False And doclin$ <> "רצוי לוודא זמנים אלה ע" + Chr$(34) + "י תצפיות" Then
         '   newhebcalfm.Combo9.AddItem "רצוי לוודא זמנים אלה ע" + Chr$(34) + "י תצפיות"
         '   End If
         If portrait = False And doclin$ <> heb1$(14) + Chr$(34) + heb1$(15) Then
            newhebcalfm.Combo9.AddItem heb1$(14) + Chr$(34) + heb1$(15)
            End If
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Combo10.AddItem doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Text31.Text = doclin$
         Line Input #filsav%, doclin$
         If doclin$ <> "NA" Then newhebcalfm.Text32.Text = doclin$
      Loop
      Close #filsav%
   Else 'input defaults
      'newhebcalfm.Combo2.AddItem "מבוסס על הזריחה המוקדמת ביותר הנראית מעל האופק המזרחי האמיתי, מנקודה כלשהי ביישוב - מידי יום ביומו"
      newhebcalfm.Combo2.AddItem heb1$(16)
      If optionheb = False Then
         newhebcalfm.Combo2.AddItem "Based on the earliest sunrise that is seen anywhere within this place on any day"
         End If
      If eros = True And currentdir = drivcities$ + "eros\visual_tmp" Then
         'newhebcalfm.Combo2.AddItem " קו רוחב: " + Str$(eroslatitude) + ", קו אורך: " + Str$(eroslongitude) + ")" + " - מידי יום ביומו" + ") " + eroscity$ + "  מבוסס על הזריחה המוקדמת ביותר הנראית מעל האופק המזרחי האמיתי, מנקודה כלשהי מסביב הישוב  "
         'newhebcalfm.Combo2.AddItem heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + ")" + heb1$(19) + ") " + eroscity$ + heb1$(20)
         If SRTMflag = 0 Then
            newhebcalfm.Combo2.AddItem "(" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) " + eroscity$ + heb1$(49) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
         ElseIf SRTMflag = 1 Then
            newhebcalfm.Combo2.AddItem "(" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) " + eroscity$ + heb1$(51) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
         ElseIf SRTMflag = 2 Then
            newhebcalfm.Combo2.AddItem "(" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) " + eroscity$ + heb1$(53) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
         ElseIf SRTMflag = 9 Then
            newhebcalfm.Combo2.AddItem "(" & heb1$(17) + Str$(eroslongitude) + heb1$(18) + Str$(eroslatitude) + " ) " + eroscity$ + heb1$(58) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
            End If
            
         If IsraelNeighborhood Then
            newhebcalfm.Combo2.AddItem heb1$(58) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61) & eroshebcity$ & " (" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) "
            End If
            
         If optionheb = False Then
            If SRTMflag = 0 Then
               'newhebcalfm.Combo2.AddItem "Based on the earliest sunrise that is seen on any day anywhere within the search radius around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
               newhebcalfm.Combo2.AddItem "GTOPO30 DTM based calculations of the earliest sunrise within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
            ElseIf SRTMflag = 1 Then
               newhebcalfm.Combo2.AddItem "SRTM-2 DTM based calculations of the earliest sunrise within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
               If eroscountry$ = "USA" Then
                  newhebcalfm.Combo2.AddItem "30m NED DTM based calculations of the earliest sunrise within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
                  End If
            ElseIf SRTMflag = 2 Then
               newhebcalfm.Combo2.AddItem "SRTM-1 DTM based calculations of the earliest sunrise within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
            ElseIf SRTMflag = 9 Then
               newhebcalfm.Combo2.AddItem "Israel DTM based calculations of the earliest sunrise within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; ITMy: " & Str(eroslatitude) & ", ITMx: " & Str(eroslongitude)
               End If
            End If
         End If
      If eros = False And astronplace = False Then
         'newhebcalfm.Combo3.AddItem "בסיוע מודל טופוגרפי ממוחשב של ארץ ישראל"
         newhebcalfm.Combo3.AddItem heb1$(21)
         If optionheb = False Then
            newhebcalfm.Combo3.AddItem "Terrain model: 25 meter DTM of Eretz Yisroel"
            End If
      ElseIf eros = True Then 'And currentdir <> drivcities$ + "eros\visual_tmp" Then
         'newhebcalfm.Combo3.AddItem "בסיוע מודל טופוגרפי ממוחשב של העולם"
         'newhebcalfm.Combo3.AddItem heb1$(22)
         If SRTMflag = 0 Then
            newhebcalfm.Combo3.AddItem heb1$(55)
         ElseIf SRTMflag = 1 Then
            newhebcalfm.Combo3.AddItem heb1$(56)
         ElseIf SRTMflag = 2 Then
            newhebcalfm.Combo3.AddItem heb1$(57)
         ElseIf SRTMflag = 9 Then
            newhebcalfm.Combo3.AddItem heb1$(60)
            End If
            
         If IsraelNeighborhood Then
            newhebcalfm.Combo3.AddItem heb1$(21)
            End If
         
         If optionheb = False Then
            'newhebcalfm.Combo3.AddItem "Terrain model: USGS GTOPO30"
            If SRTMflag = 0 Then
               newhebcalfm.Combo3.AddItem "Terrain model: USGS GTOPO30"
            ElseIf SRTMflag = 1 Then
               newhebcalfm.Combo3.AddItem "Terrain model: SRTM-2"
            ElseIf SRTMflag = 2 Then
               newhebcalfm.Combo3.AddItem "Terrain model: SRTM-1"
            ElseIf SRTMflag = 9 Then
               newhebcalfm.Combo3.AddItem "Terrain model: 25m DTM of Eretz Yisroel"
               End If
            End If
            
         End If
      If portrait = True Then newhebcalfm.Combo4.AddItem sEmpty
      If portrait = False And astronplace = False Then
         'newhebcalfm.Combo4.AddItem "רצוי לוודא זמנים אלה ע" + Chr$(34) + "י תצפיות"
         newhebcalfm.Combo4.AddItem heb1$(14) + Chr$(34) + heb1$(15)
         If optionheb = False Then
            newhebcalfm.Combo4.AddItem "It is advisable to check these times with observations"
            End If
         End If
      If eros = True And currentdir = drivcities$ + "eros\visual_tmp" Then
         'newhebcalfm.Combo5.AddItem " קו רוחב: " + Str$(eroslatitude) + ", קו אורך: " + Str$(eroslongitude) + ")" + " - מידי יום ביומו" + ") " + eroscity$ + "  מבוסס על השקיעה המאוחרת ביותר הנראית מעל האופק המערבי האמיתי, מנקודה כלשהי מסביב הישוב  "
         'newhebcalfm.Combo5.AddItem heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + ")" + heb1$(19) + ") " + eroscity$ + heb1$(23)
         If SRTMflag = 0 Then
            newhebcalfm.Combo7.AddItem "(" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) " + eroscity$ + heb1$(50) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
         ElseIf SRTMflag = 1 Then
            newhebcalfm.Combo7.AddItem "(" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) " + eroscity$ + heb1$(52) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
         ElseIf SRTMflag = 2 Then
            newhebcalfm.Combo7.AddItem "(" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) " + eroscity$ + heb1$(54) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
         ElseIf SRTMflag = 9 Then
            newhebcalfm.Combo7.AddItem "(" & heb1$(17) + Str$(eroslongitude) + heb1$(18) + Str$(eroslatitude) + " ) " + eroscity$ + heb1$(59) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61)
            End If
            
         If IsraelNeighborhood Then
            newhebcalfm.Combo7.AddItem heb1$(59) & " " & Trim$(Str$(searchradius)) & " " & heb1$(61) & eroshebcity$ & " (" & heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ) "
            End If
            
         If optionheb = False Then
            If SRTMflag = 0 Then
               'newhebcalfm.Combo5.AddItem "Based on the lattest sunset that is seen on any day anywhere within the search radius around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
               newhebcalfm.Combo7.AddItem "GTOPO30 DTM based calculations of the lattest sunset within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
            ElseIf SRTMflag = 1 Then
               newhebcalfm.Combo7.AddItem "SRTM-2 DTM based calculations of the lattest sunset within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
                If eroscountry$ = "USA" Then
                  newhebcalfm.Combo7.AddItem "30m NED DTM based calculations of the earliest sunrise within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
                  End If
            ElseIf SRTMflag = 2 Then
               newhebcalfm.Combo7.AddItem "SRTM-1 DTM based calculations of the lattest sunset within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
            ElseIf SRTMflag = 9 Then
               newhebcalfm.Combo7.AddItem "Israel DTM based calculations of the lattest sunset within " & Trim$(Str$(searchradius)) & "km. around " & eroscity$ & "; ITMy: " & Str(eroslatitude) & ", ITMx: " & Str(eroslongitude)
               End If
            End If
         End If
      'newhebcalfm.Combo5.AddItem "כל הזמנים לפי שעון חורף.  כדי להשתמש בזמנים אלו לקביעת שעות היום, יש לעשות שאלת חכם."
      
      '////////////////added DST support on 082921////////////////////////////////////
      If optionheb Then
         If CalMDIform.mnuDST.Checked Then
            newhebcalfm.Combo5.AddItem sEmpty
         Else
            newhebcalfm.Combo5.AddItem heb1$(24)
            End If
      Else
         If CalMDIform.mnuDST.Checked Then
            newhebcalfm.Combo5.AddItem sEmpty
         Else
            newhebcalfm.Combo5.AddItem "All times are according to Standard Time."
            End If
         End If
      '/////////////////////////////////////////////////////////////////
      
      '////////////////////fixed 082921 -- PathZones cushions overides any other stored setting//////////////////
      newhebcalfm.Text2.Text = obscushion
      newhebcalfm.Text32.Text = -obscushion
      '////////////////////////////////////////////////////////////////////////////////////////////////////////
      
      'newhebcalfm.Combo7.AddItem "מבוסס על השקיעה המאוחרת ביותר הנראית מעל האופק המערבי האמיתי, מנקודה כלשהי ביישוב - מידי יום ביומו"
      If eros = False Then
         newhebcalfm.Combo7.AddItem heb1$(13)
         If optionheb = False Then
            newhebcalfm.Combo7.AddItem "Based on the lattest sunset that is seen anywhere in this place on any day"
            End If
         End If
      If eros = False And astronplace = False Then
'         newhebcalfm.Combo8.AddItem "בסיוע מודל טופוגרפי ממוחשב של ארץ ישראל"
         newhebcalfm.Combo8.AddItem heb1$(21)
         'newhebcalfm.Combo8.AddItem "מדרש בכורי יוסף"
         If optionheb = False Then
            newhebcalfm.Combo8.AddItem "Terrain model: 25 meter DTM of Eretz Yisorel"
            End If
      ElseIf eros = True Then
         'newhebcalfm.Combo8.AddItem "בסיוע מודל טופוגרפי ממוחשב של העולם"
         'newhebcalfm.Combo8.AddItem heb1$(22)
         If SRTMflag = 0 Then
            newhebcalfm.Combo8.AddItem heb1$(55)
         ElseIf SRTMflag = 1 Then
            newhebcalfm.Combo8.AddItem heb1$(56)
         ElseIf SRTMflag = 2 Then
            newhebcalfm.Combo8.AddItem heb1$(57)
         ElseIf SRTMflag = 9 Then
            newhebcalfm.Combo8.AddItem heb1$(60)
            End If
            
         If IsraelNeighborhood Then
            newhebcalfm.Combo8.AddItem heb1$(21)
            End If
            
         If optionheb = False Then
            'newhebcalfm.Combo8.AddItem "Terrain model: USGS GTOPO30"
            If SRTMflag = 0 Then
               newhebcalfm.Combo8.AddItem "Terrain model: USGS GTOPO30"
            ElseIf SRTMflag = 1 Then
               newhebcalfm.Combo8.AddItem "Terrain model: SRTM-2"
            ElseIf SRTMflag = 2 Then
               newhebcalfm.Combo8.AddItem "Terrain model: SRTM-1"
            ElseIf SRTMflag = 9 Then
               newhebcalfm.Combo8.AddItem "Terrain model: 25m DTM of Eretz Israel"
               End If
            End If
         End If
      If portrait = True Then newhebcalfm.Combo9.AddItem sEmpty
      If portrait = False Then
         'newhebcalfm.Combo9.AddItem "רצוי לוודא זמנים אלה ע" + Chr$(34) + "י תצפיות"
         newhebcalfm.Combo9.AddItem heb1$(14) + Chr$(34) + heb1$(15)
         If optionheb = False Then
            newhebcalfm.Combo9.AddItem "It is advisable to check these times with observations"
            End If
         End If
      'newhebcalfm.Combo10.AddItem "כל הזמנים לפי שעון חורף.  כדי להשתמש בזמנים אלו לקביעת שעות היום, יש לעשות שאלת חכם."
'      newhebcalfm.Combo10.AddItem heb1$(24)
'      If optionheb = False Then
'         newhebcalfm.Combo10.AddItem "All times are according to Standard Time."
'         End If
         
      '////////////////added DST support on 082921////////////////////////////////////
      If optionheb Then
        If CalMDIform.mnuDST.Checked Then
           newhebcalfm.Combo10.AddItem sEmpty
        Else
           newhebcalfm.Combo10.AddItem heb1$(24)
           End If
      Else
         If CalMDIform.mnuDST.Checked Then
            newhebcalfm.Combo10.AddItem sEmpty
         Else
            newhebcalfm.Combo10.AddItem "All times are according to Standard Time."
            End If
         End If
      '/////////////////////////////////////////////////////////////////
         
         
      End If
   
   If Check4.Value = vbChecked Then 'add captions for astronomical sunrise
      If astronplace = True Then 'use inputed name as city's hebrew name
         'newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה האסטרונומי ב" + hebcityname$
         newhebcalfm.Combo1.AddItem title$ + heb1$(25) + hebcityname$
         'If internet = True Then
         '   newhebcalfm.Combo1.AddItem hebcityname$ + title$ + heb1$(25)
         '   End If
         If optionheb = False Then
            newhebcalfm.Combo1.AddItem title$ + " of the Astronomical Sunrise for " + hebcityname$
            End If
      ElseIf eros = True Then
         'newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה האסטרונומי ב" + hebcityname$
         'newhebcalfm.Combo1.AddItem title$ + heb1$(25) + hebcityname$
         newhebcalfm.Combo1.AddItem title$ + heb1$(41) + hebcityname$
        
         If optionheb = False Then
            newhebcalfm.Combo1.AddItem title$ + " of the Astronomical Sunrise for the region of " + citnamp$ 'eroscity$
            End If
         'newhebcalfm.Combo3.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         newhebcalfm.Combo3.AddItem heb1$(26)
      Else
         'newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה האסטרונומי ב" + cityhebnames$(i%)
         newhebcalfm.Combo1.AddItem title$ + heb1$(25) + cityhebnames$(i%)
         If optionheb = False Then
            newhebcalfm.Combo1.AddItem title$ + " of the Astronomical Sunrise for " + citnamp$ ' UCase$(Mid$(citnam$, 1, 1)) + Mid$(citnam$, 2, Len(citnam$) - 1)
            End If
         If astronplace = False Then
            'newhebcalfm.Combo3.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה של ארץ ישראל"
            newhebcalfm.Combo3.AddItem heb1$(27)
            End If
         End If
      If astronplace = False Then
         'newhebcalfm.Combo2.AddItem "מבוסס על הזריחה האסטרונומי המוקדמת ביותר, מנקודה כלשהי ביישוב - מידי יום ביומו"
         newhebcalfm.Combo2.AddItem heb1$(28) '<----fix here!!!!!!!!!!!!
         If optionheb = False Then
            newhebcalfm.Combo2.AddItem "Based on calculation of the earliest astronomical sunrise for this place"
            End If
         If eros = True Then
            newhebcalfm.Combo2.AddItem heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ," + eroscity$ + heb1$(45)
            If optionheb = False Then
               newhebcalfm.Combo2.AddItem "Based on the astronomical sunrise for " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
               End If
            End If
      Else
         'newhebcalfm.Combo2.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         If internet = False Then
            newhebcalfm.Combo2.AddItem heb1$(26)
            newhebcalfm.Combo2.AddItem heb1$(28)
            newhebcalfm.Combo3.AddItem sEmpty
            If optionheb = False Then
               newhebcalfm.Combo2.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
               End If
         Else
            If nettype$ = "Astr" Then
               newhebcalfm.Combo2.AddItem heb1$(26)
               If optionheb = False Then
                  newhebcalfm.Combo2.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
                  End If
            Else
               newhebcalfm.Combo2.AddItem heb1$(28)
               If optionheb = False Then
                  newhebcalfm.Combo2.AddItem "Based on the earliest astronomical sunrise for this place"
                  End If
               End If
            End If
         End If
'      newhebcalfm.Combo4.AddItem "רצוי לוודא ע" + Chr$(34) + "י תצפיות כשיש אופק מזרחי פנוי מהסתרים"
      newhebcalfm.Combo4.AddItem heb1$(29) + Chr$(34) + heb1$(30)
      End If
   If Check5.Value = vbChecked Then 'add captions for astronomical sunset
      If astronplace = True Then  'use inputed name as city's hebrew name
         'newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה האסטרונומית ב" + hebcityname$
         newhebcalfm.Combo6.AddItem title$ + heb1$(31) + hebcityname$
         'If internet = True Then
         '   newhebcalfm.Combo1.AddItem hebcityname$ + title$ + heb1$(31)
         '   End If
         If optionheb = False Then
            newhebcalfm.Combo6.AddItem title$ + " of the Astronomical Sunset for " + hebcityname$
            End If
      ElseIf eros = True Then
         'newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה האסטרונומית ב" + hebcityname$
         newhebcalfm.Combo6.AddItem title$ + heb1$(42) + hebcityname$
         If optionheb = False Then
            newhebcalfm.Combo6.AddItem title$ + " of the Astronomical Sunset for the region of " + citnamp$ 'eroscity$
            End If
         'newhebcalfm.Combo8.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         newhebcalfm.Combo8.AddItem heb1$(26)
      Else
         'newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה האסטרונומית ב" + cityhebnames$(i%)
         newhebcalfm.Combo6.AddItem title$ + heb1$(31) + cityhebnames$(i%)
         If optionheb = False Then
            newhebcalfm.Combo6.AddItem title$ + " of the Astronomical Sunset for " + citnamp$  'UCase$(Mid$(citnam$, 1, 1)) + Mid$(citnam$, 2, Len(citnam$) - 1)
            End If
         If astronplace = False Then
            'newhebcalfm.Combo8.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה של ארץ ישראל"
            newhebcalfm.Combo8.AddItem heb1$(27)
            End If
         End If
      If astronplace = False Then
         'newhebcalfm.Combo7.AddItem "מבוסס על השקיעה האסטרונומית המאוחרת ביותר, מנקודה כלשהי ביישוב - מידי יום ביומו"
         newhebcalfm.Combo7.AddItem heb1$(32)
         newhebcalfm.Combo8.AddItem sEmpty
         If optionheb = False Then
            newhebcalfm.Combo7.AddItem "Based on the lattest Astronomical Sunrise for this place"
            End If
         If eros = True Then
            newhebcalfm.Combo7.AddItem heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ," + eroscity$ + heb1$(46)
            If optionheb = False Then
               newhebcalfm.Combo7.AddItem "Based on the astronomical sunset for " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
               End If
            End If
      Else
         'newhebcalfm.Combo7.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         If internet = False Then
            newhebcalfm.Combo7.AddItem heb1$(26)
            newhebcalfm.Combo7.AddItem heb1$(32)
            If optionheb = False Then
               newhebcalfm.Combo7.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
               End If
         Else
            If nettype$ = "Astr" Then
               newhebcalfm.Combo7.AddItem heb1$(26)
               If optionheb = False Then
                  newhebcalfm.Combo7.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
                  End If
            Else
               newhebcalfm.Combo7.AddItem heb1$(32)
               If optionheb = False Then
                  newhebcalfm.Combo7.AddItem "Based on latest astronomical sunset for this place"
                  End If
               End If
            End If
           
         End If
      'newhebcalfm.Combo9.AddItem "רצוי לוודא ע" + Chr$(34) + "י תצפיות כשיש אופק מערבי פנוי מהסתרים"
      newhebcalfm.Combo9.AddItem heb1$(29) + Chr$(34) + heb1$(33)
      End If
   If Check6.Value = vbChecked Then 'add captions for mishor sunrise
      If astronplace = True Then  'use inputed name as city's hebrew name
         If Katz = False Then
             'newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה המישורי ב" + hebcityname$
             newhebcalfm.Combo1.AddItem title$ + heb1$(34) + hebcityname$
             'If internet = True Then
             '   newhebcalfm.Combo1.AddItem hebcityname$ + title$ + heb1$(34)
             '   End If
             If optionheb = False Then
                newhebcalfm.Combo1.AddItem title$ + " of the Mishor Sunrise of " + hebcityname$
                End If
         Else
             'newhebcalfm.Combo1.AddItem "זריחה"
             newhebcalfm.Combo1.AddItem heb1$(36)
             End If
      ElseIf eros = True Then
         'newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה המישורי ב" + hebcityname$
         'newhebcalfm.Combo1.AddItem title$ + heb1$(34) + hebcityname$
         'If optionheb = False Then
         '   newhebcalfm.Combo1.AddItem title$ + " of the Mishor Sunrise of " + eroscity$
         '   End If
         newhebcalfm.Combo1.AddItem title$ + heb1$(43) + hebcityname$
        
         If optionheb = False Then
            newhebcalfm.Combo1.AddItem title$ + " of the Mishor Sunrise for the region of " + citnamp$ 'eroscity$
            End If
         'newhebcalfm.Combo3.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         newhebcalfm.Combo3.AddItem heb1$(26)
      Else
         'newhebcalfm.Combo1.AddItem title$ + " לזמני הנץ החמה המישורי ב" + cityhebnames$(i%)
         newhebcalfm.Combo1.AddItem title$ + heb1$(34) + cityhebnames$(i%)
         If optionheb = False Then
            newhebcalfm.Combo1.AddItem title$ + " of the Mishor Sunrise of " + citnamp$ '+ UCase$(Mid$(citnam$, 1, 1)) + Mid$(citnam$, 2, Len(citnam$) - 1)
            End If
         'newhebcalfm.Combo3.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה של ארץ ישראל"
         newhebcalfm.Combo3.AddItem heb1$(27)
         End If
      If astronplace = False Then
         'newhebcalfm.Combo2.AddItem "מבוסס על הזריחה המישורי המוקדמת ביותר, מנקודה כלשהי ביישוב - מידי יום ביומו"
         newhebcalfm.Combo2.AddItem heb1$(26)
         newhebcalfm.Combo2.AddItem heb1$(35)
         newhebcalfm.Combo3.AddItem sEmpty
         newhebcalfm.Combo4.AddItem sEmpty
         
         If optionheb = False Then
            newhebcalfm.Combo2.AddItem "Based on the earliest mishor sunrise for this place"
            End If
         If eros = True Then
            newhebcalfm.Combo2.AddItem heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ," + eroscity$ + heb1$(47)
            If optionheb = False Then
               newhebcalfm.Combo2.AddItem "Based on the mishor sunrise for " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
               End If
            End If
            
      Else
         'newhebcalfm.Combo2.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         If internet = False Then
            newhebcalfm.Combo2.AddItem heb1$(26)
            newhebcalfm.Combo2.AddItem heb1$(35)
            If optionheb = False Then
               newhebcalfm.Combo2.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
               End If
         Else
            If nettype$ = "Astr" Then
               newhebcalfm.Combo2.AddItem heb1$(26)
               If optionheb = False Then
                  newhebcalfm.Combo2.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
                  End If
            Else
               newhebcalfm.Combo2.AddItem heb1$(35)
               If optionheb = False Then
                  newhebcalfm.Combo2.AddItem "Based on earliest mishor sunrise for this place"
                  End If
               End If
            End If
               
         End If
      newhebcalfm.Combo3.AddItem sEmpty
      End If
   If Check7.Value = vbChecked Then 'add captions for mishor sunset
      If astronplace = True Then 'use inputed name as city's hebrew name
         If Katz = False Then
            'newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה המישורית ב" + hebcityname$
            newhebcalfm.Combo6.AddItem title$ + heb1$(37) + hebcityname$
         'If internet = True Then
         '   newhebcalfm.Combo1.AddItem hebcityname$ + title$ + heb1$(37)
         '   End If
         If optionheb = False Then
            newhebcalfm.Combo6.AddItem title$ + " of the Mishor Sunset for " + hebcityname$
            End If
         Else
            'newhebcalfm.Combo6.AddItem "שקיעה"
            newhebcalfm.Combo6.AddItem heb1$(38)
            End If
      ElseIf eros = True Then
         'newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה המישורית ב" + hebcityname$
         'newhebcalfm.Combo6.AddItem title$ + heb1$(37) + hebcityname$
         'If optionheb = False Then
         '   newhebcalfm.Combo6.AddItem title$ + " of the Mishor Sunset for " + eroscity$
         '   End If
         newhebcalfm.Combo6.AddItem title$ + heb1$(44) + hebcityname$
         If optionheb = False Then
            newhebcalfm.Combo6.AddItem title$ + " of the Mishor Sunset for the region of " + citnamp$ 'eroscity$
            End If
         
         'newhebcalfm.Combo8.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         newhebcalfm.Combo8.AddItem heb1$(26)
      Else
         'newhebcalfm.Combo6.AddItem title$ + " לזמני שקיעת החמה המישורית ב" + cityhebnames$(i%)
         newhebcalfm.Combo6.AddItem title$ + heb1$(37) + cityhebnames$(i%)
         If optionheb = False Then
           newhebcalfm.Combo6.AddItem title$ + " of the Mishor Sunset for " + citnamp$ 'UCase$(Mid$(citnam$, 1, 1)) + Mid$(citnam$, 2, Len(citnam$) - 1)
           End If
         If astronplace = False Then
            'newhebcalfm.Combo8.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה של ארץ ישראל"
            newhebcalfm.Combo8.AddItem heb1$(27)
            End If
         End If
      If astronplace = False Then
         'newhebcalfm.Combo7.AddItem "מבוסס על השקיעה האסטרונומית המאוחרת ביותר, מנקודה כלשהי ביישוב - מידי יום ביומו"
         newhebcalfm.Combo7.AddItem heb1$(26)
         newhebcalfm.Combo7.AddItem heb1$(39)
         If optionheb = False Then
            newhebcalfm.Combo7.AddItem "Based on the lattest mishor sunset for this place"
            End If
         If eros = True Then
            newhebcalfm.Combo7.AddItem heb1$(17) + Str$(eroslatitude) + heb1$(18) + Str$(eroslongitude) + " ," + eroscity$ + heb1$(48)
            If optionheb = False Then
               newhebcalfm.Combo7.AddItem "Based on the mishor sunset for " & eroscity$ & "; longitude: " & Str(eroslongitude) & ", latitude: " & Str(eroslatitude)
               End If
            End If
            
      Else
         'newhebcalfm.Combo7.AddItem "מבוסס על חישובים משוכללים של שבירת קרני השמש באטמוספירה"
         If internet = False Then
            newhebcalfm.Combo7.AddItem heb1$(26)
            newhebcalfm.Combo7.AddItem heb1$(39)
            newhebcalfm.Combo8.AddItem sEmpty
            If optionheb = False Then
               newhebcalfm.Combo7.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
               End If
         Else
            If nettype$ = "Astr" Then
               newhebcalfm.Combo2.AddItem heb1$(26)
               If optionheb = False Then
                  newhebcalfm.Combo7.AddItem "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres"
                  End If
            Else
               newhebcalfm.Combo2.AddItem heb1$(39)
               If optionheb = False Then
                  newhebcalfm.Combo7.AddItem "Based on the latest mishor sunset for this place"
                  End If
               End If
            End If
         End If
      newhebcalfm.Combo9.AddItem sEmpty
      End If
   If eros = False Or internet Then 'add sponsor line
      If optionheb = True Then
         'newhebcalfm.Combo4.AddItem "מדרש בכורי יוסף, ירושלים"
         newhebcalfm.Combo4.AddItem SponsorLine$ 'heb1$(40)
         'newhebcalfm.Combo9.AddItem "מדרש בכורי יוסף, ירושלים"
         newhebcalfm.Combo9.AddItem SponsorLine$ 'heb1$(40)
      ElseIf optionheb = False Then
         newhebcalfm.Combo4.AddItem SponsorLine$ '"Midrash Bikurei Yosef, Jerusalem"
         newhebcalfm.Combo9.AddItem SponsorLine$ '"Midrash Bikurei Yosef, Jerusalem"
         End If
      End If
   
   'if dstcheck then add blank to list in order to remove mention of standard time
   If CalMDIform.mnuDST.Checked Then newhebcalfm.Combo5.AddItem sEmpty
   If CalMDIform.mnuDST.Checked Then newhebcalfm.Combo10.AddItem sEmpty
   
   'add dedications to Seide
   If optionheb Then
      'add dedication to Seide
      newhebcalfm.Combo3.AddItem heb2$(16)
      newhebcalfm.Combo8.AddItem heb2$(16)
   Else
      newhebcalfm.Combo3.AddItem "In loving memory of Avrohom Yitzhak ben Zvi z''l"
      newhebcalfm.Combo8.AddItem "In loving memory of Avrohom Yitzhak ben Zvi z''l"
      End If
      
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
   
   If (newhebcalfm.Combo1.Text <> compare1$ Or newhebcalfm.Combo6.Text <> compare2$) And _
      (Check4.Value = vbUnchecked And Check5.Value = vbUnchecked And Check6.Value = vbUnchecked And Check7.Value = vbUnchecked) Then
      'check for previous titles
      title2$ = "לוח " + Chr$(34) + "בכורי יוסף" + Chr$(34)
      If eros = True And currentdir$ = drivcities$ & "eros\visual_tmp" Then
         title2 = "לוח " + Chr$(34) + "חי" + Chr$(34)
         End If
      compare3$ = title2$ + " לזמני הנץ החמה הנראה ב" + cityhebnames$(i%)
      compare4$ = title2$ + " לזמני שקיעת החמה הנראית ב" + cityhebnames$(i%)
      If (newhebcalfm.Combo1.Text <> compare3$ Or newhebcalfm.Combo6.Text <> compare4$) And _
         (Check4.Value = vbUnchecked And Check5.Value = vbUnchecked And Check6.Value = vbUnchecked And Check7.Value = vbUnchecked And eros = False) Then
         If internet = True Then
            'theoretically should abort this scan, but there is
            'a problem with the Windows 2000 hebrew, so let it run
            ''abort this scan
            'Close
            'myfile = Dir(drivfordtm$ + "busy.cal")
            'If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
            'For i% = 0 To Forms.Count - 1
            '  Unload Forms(i%)
            'Next i%
            ''Unload CalMDIform
            newhebcalfm.Combo1.AddItem compare1$
            newhebcalfm.Combo6.AddItem compare2$
            newhebcalfm.Combo1.ListIndex = newhebcalfm.Combo1.ListCount - 1
            newhebcalfm.Combo6.ListIndex = newhebcalfm.Combo6.ListCount - 1
            GoTo i500
            End If
         If automatic = True And Not autosave Then
            response = MsgBox("There seems to be a discrepancy between the city's hebrew name and the stored name in the .SAV file!...will ABORT automatic run for next iteration.", vbExclamation + vbOKCancel, "Cal Program")
            If response = vbCancel Then
               autocancel = True
               End If
         ElseIf Not autosave Then
            MsgBox "There seems to be a discrepancy between the city's hebrew name and the stored name in the .SAV file...so be SURE you fix it!", vbExclamation + vbOKOnly, "Cal Program"
            End If
      Else
         newhebcalfm.Combo1.AddItem compare1$
         newhebcalfm.Combo6.AddItem compare2$
         newhebcalfm.Combo1.ListIndex = newhebcalfm.Combo1.ListCount - 1
         newhebcalfm.Combo6.ListIndex = newhebcalfm.Combo6.ListCount - 1
         End If
      End If
i500:
   lognum% = FreeFile
   Open drivjk$ + "calprog.log" For Append As #lognum%
   Print #lognum%, "Step #4: SunriseSunset finished determining captions"
   Close #lognum%

   If Abs(nsetflag%) = 1 Or Abs(nsetflag%) = 3 Then
      nset$ = "netz"
   ElseIf Abs(nsetflag%) = 2 Then
      nset$ = "skiy"
      End If
'   outnam1$ = currentdrive + "\" + currentdir + "\" + nset$ + "\*.bat"
   outnam1$ = currentdir + "\" + nset$ + "\*.bat"
   batnam1$ = Dir(outnam1$)
'   bat1$ = currentdrive + "\" + currentdir + "\" + nset$ + "\" + LTrim$(batnam1$)
   bat1$ = currentdir + "\" + nset$ + "\" + LTrim$(batnam1$)
   myfile = Dir(bat1$)
   If myfile = sEmpty Then
      If internet = True And Err.Number >= 0 Then   'fatal fault, abort
         Close
         myfile = Dir(drivfordtm$ + "busy.cal")
         If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
         
         lognum% = FreeFile
         Open drivjk$ + "calprog.log" For Append As #lognum%
         Print #lognum%, "Fatal Error:  Bat file not found. Abort program"
         Close #lognum%
         
         For i% = 0 To Forms.Count - 1
           Unload Forms(i%)
         Next i%
          
          'kill timer
          If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)

          'end program abruptly
          End
         End If
      If nset$ = "netz" Then
         If automatic Then 'uncheck it
            SunriseSunset.Check1.Value = False
            GoTo 50
            End If
         response = MsgBox("Can't find file: " + bat1$ + "! Check if there is a netz directory.", vbCritical + vbOKOnly, "Cal Program")
      ElseIf nset$ = "skiy" Then
         If automatic Then 'uncheck it
            SunriseSunset.Check2.Value = False
            GoTo 50
            End If
         response = MsgBox("Can't find file: " + bat1$ + "! Check if there is a skiy directory.", vbCritical + vbOKOnly, "Cal Program")
         End If
      Cancelbut_Click
      Exit Sub
      End If
   newfilbat% = FreeFile
   Open bat1$ For Input As #newfilbat%
   nfil% = 0
   If eros = True Then Line Input #newfilbat%, doclin$
   Do Until EOF(newfilbat%)
      Line Input #newfilbat%, doclin$
      If LCase$(Mid$(doclin$, 1, 7)) = "version" Then
         Exit Do
         End If
      nfil% = nfil% + 1
   Loop
   Close #newfilbat%
   If Abs(nsetflag%) = 3 Then
'      outnam2$ = currentdrive + "\" + currentdir + "\skiy\*.bat"
      outnam2$ = currentdir + "\skiy\*.bat"
      batnam2$ = Dir(outnam2$)
'      bat2$ = currentdrive + "\" + currentdir + "\skiy\" + LTrim$(batnam2$)
      bat2$ = currentdir + "\skiy\" + LTrim$(batnam2$)
      myfile = Dir(bat2$)
      If myfile = sEmpty Then
         response = MsgBox("Can't find file: " + bat2$ + "!  Check if there is a skiy directory.", vbCritical + vbOKOnly, "Cal Program")
         Cancelbut_Click
         Exit Sub
         End If
      newfilbat% = FreeFile
      Open bat2$ For Input As #newfilbat%
      If eros = True Then Line Input #newfilbat%, doclin$ 'extra line of documentation
      Do Until EOF(newfilbat%)
         Line Input #newfilbat%, doclin$
         If InStr(doclin$, "version") <> 0 Then Exit Do
         nfil% = nfil% + 1
      Loop
      Close #newfilbat%
      End If
      
   '////113021 fixed bug that reset eros flag and lost geotz info when redoing eros luchos///
   If InStr(currentdir, "eros") <> 0 Then
      eros = True
      geo = True
      If eroscountry$ = "Israel" Then geotz! = 2
      End If
   '/////////////////////////////////////
      
   filtm3num% = FreeFile
   Open drivjk$ + "netzskiy.tm4" For Output As #filtm3num%
   Write #filtm3num%, yrheb%
   If geo = False Or eroscountry$ = "Israel" Then 'using ITM coordinates
      Write #filtm3num%, nsetflag%, geotz!
   ElseIf geo = True And eros = False Then
      Write #filtm3num%, nsetflag% - 3, geotz!
   ElseIf eros = True Then
      If nsetflag% < 0 Then
         Write #filtm3num%, nsetflag% - 3, geotz!
      Else
         Write #filtm3num%, -nsetflag% - 3, geotz!
         End If
      End If
   Write #filtm3num%, nfil%
   Print #filtm3num%, Mid$(batnam1$, 1, 4)
   Print #filtm3num%, currentdir
'
   batnam$ = bat1$
   checklst% = 0
   nchknez% = 0
   nchkski% = 0
   nfind% = 0
   oldnset% = 0
   maxang% = 30
   nckpronez% = 0
   nckproski% = 0
   sumkmxo = 0
   sumkmyo = 0
   sumhgt = 0
   filbatnum% = FreeFile
300   Open batnam$ For Input As #filbatnum%
      If eros = True Then Line Input #filbatnum%, doclin$
      nbat% = 0
      Do Until EOF(filbatnum%)
         Input #filbatnum%, doclin$, kmxo, kmyo, hgt
         pos% = InStr(LCase$(doclin$), "version")
         If pos% <> 0 Then
            'read version number and exit
            'datavernum = Val(Mid(doclin$, pos% + 8, 1))
            datavernum = kmxo
            If eros Then
               If kmyo = 0 Then
                  'GTOPO30
                  SRTMflag = 0
               ElseIf kmyo = 1 Then
                  'SRTM level 1 (30 meter) DTM
                  SRTMflag = 1
                  'SunriseSunset.Check3.Value = vbChecked 'check for near mountains
                  'nearski = True
                  'nearnez = True
                  'nearcolor = True
               ElseIf kmyo = 2 Then
                  'SRTM level 2 (90 meter) DTM
                  SRTMflag = 2
               ElseIf kmyo = 9 Or kmyo = 0 Then
                  'Israel DTM and Jerusalem neighborhoods
                  SRTMflag = 9
                  End If
            Else
               SRTMflag = 9 'Israel 25m DTM
               End If
            GoTo 475
            End If
         nbat% = nbat% + 1
         sumkmxo = sumkmxo + kmxo
         sumkmyo = sumkmyo + kmyo
         sumhgt = sumhgt + hgt
         For i% = Len(doclin$) To 1 Step -1
            If Mid$(doclin$, i%, 1) = "\" Then
               fileo$ = Mid$(doclin$, i% + 1, Len(doclin$) - i%)
               Exit For
               End If
         Next i%
         direco$ = drivfordtm$ + nset$
         filen$ = direco$ & "\" & fileo$
         If Check6.Value = vbChecked Or Check7.Value = vbChecked Then hgt = 0
         If nsetflag% = 1 Or nsetflag% = 3 Then '<---check here
           'Print #filtm3num%, filen$; Tab(31); Format(kmxo, "###.000"); Tab(42); Format(kmyo, "####.000"); Tab(53); Format(hgt, "####.0"); Tab(60); stryr%; Tab(67); endyr%; Tab(74); 0
           Print #filtm3num%, filen$
           fildoc$ = Str$(kmxo) & "," & Str$(kmyo) & "," & Str$(hgt) & "," & Str$(stryr%) & "," & Str$(yrstrt%(0)) & "," & Str$(yrend%(0)) & "," & Str$(endyr%) & ", 1," & Str$(yrend%(1)) & ", 0"
           Print #filtm3num%, fildoc$
         ElseIf nsetflag% = 2 Then
           'Print #filtm3num%, filen$; Tab(31); Format(kmxo, "###.000"); Tab(42); Format(kmyo, "####.000"); Tab(53); Format(hgt, "####.0"); Tab(60); stryr%; Tab(67); endyr%; Tab(74); 1
           Print #filtm3num%, filen$
           fildoc$ = Str$(kmxo) & "," & Str$(kmyo) & "," & Str$(hgt) & "," & Str$(stryr%) & "," & Str$(yrstrt%(0)) & "," & Str$(yrend%(0)) & "," & Str$(endyr%) & ", 1," & Str$(yrend%(1)) & ", 1"
           Print #filtm3num%, fildoc$
         ElseIf nsetflag% = -1 Or nsetflag% = -3 Then
           'Print #filtm3num%, filen$; Tab(31); Format(kmxo, "###.000"); Tab(42); Format(kmyo, "####.000"); Tab(53); Format(hgt, "####.0"); Tab(60); stryr%; Tab(67); endyr%; Tab(74); 2
           Print #filtm3num%, filen$
           fildoc$ = Str$(kmxo) & "," & Str$(kmyo) & "," & Str$(hgt) & "," & Str$(stryr%) & "," & Str$(yrstrt%(0)) & "," & Str$(yrend%(0)) & "," & Str$(endyr%) & ", 1," & Str$(yrend%(1)) & ", 2"
           Print #filtm3num%, fildoc$
         ElseIf nsetflag% = -2 Then
           'Print #filtm3num%, filen$; Tab(31); Format(kmxo, "###.000"); Tab(42); Format(kmyo, "####.000"); Tab(53); Format(hgt, "####.0"); Tab(60); stryr%; Tab(67); endyr%; Tab(74); 3
           Print #filtm3num%, filen$
           fildoc$ = Str$(kmxo) & "," & Str$(kmyo) & "," & Str$(hgt) & "," & Str$(stryr%) & "," & Str$(yrstrt%(0)) & "," & Str$(yrend%(0)) & "," & Str$(endyr%) & ", 1," & Str$(yrend%(1)) & ", 3"
           Print #filtm3num%, fildoc$
           End If
                  
200      SourceFile = currentdir + "\" + nset$ + "\" + fileo$
         DestinationFile = direco$ + "\" + fileo$
         FileCopy SourceFile, DestinationFile    ' Copy source to target.
475
      Loop
    If Abs(nsetflag%) = 3 Then
       oldnset% = 3
       Close #filbatnum%
       filbatnum% = FreeFile
       If nsetflag% = -3 Then
          avekmxnetz = sumkmxo / nbat%
          avekmynetz = sumkmyo / nbat%
          avehgtnetz = sumhgt / nbat%
          sumkmxo = 0
          sumkmyo = 0
          sumhgt = 0
          nsetflag% = -2
       Else
          avekmxnetz = sumkmxo / nbat%
          avekmynetz = sumkmyo / nbat%
          avehgtnetz = sumhgt / nbat%
          sumkmxo = 0
          sumkmyo = 0
          sumhgt = 0
          nsetflag% = 2
          End If
       batnam$ = bat2$
       nset$ = "skiy"
       GoTo 300
    Else
       avekmxnetz = sumkmxo / nbat%
       avekmynetz = sumkmyo / nbat%
       avehgtnetz = sumhgt / nbat%
       avekmxskiy = sumkmxo / nbat%
       avekmyskiy = sumkmyo / nbat%
       avehgtskiy = sumhgt / nbat%
       End If
       
    If aveusa = True Then 'invert the coordinates
       avetmp = avekmynetz
       avekmynetz = avekmxnetz
       avekmxnetz = avetmp
       avetmp = avekmyskiy
       avekmyskiy = avekmxskiy
       avekmxskiy = avetmp
       'aveusa = False
       End If
       
    Close #filbatnum%
    Close #filtm3num%
    
    If oldnset% = 3 Then
       If aveusa = False Then
          avekmxskiy = sumkmxo / nbat%
          avekmyskiy = sumkmyo / nbat%
       Else
          avekmxskiy = sumkmyo / nbat%
          avekmyskiy = sumkmxo / nbat%
          End If
       avehgtskiy = sumhgt / nbat%
       aveusa = False
       nsetflag% = 3
       End If
    If Check3.Value = vbChecked Then
       distlimnum = Val(SunriseSunset.Text2) 'Maximum acceptable percentage of entries that are inaccurate
       If eros = False Then
          distlim = SunriseSunset.Text1 'If obstructions are within this distance, then
                      'time entries have questionable accuracies
       Else
          distlim = SunriseSunset.Text1 'If obstructions are within this distance, then
                       'time entries have questionable accuracies
          End If
       End If
    
    'for new format, skip old type of obstruction checking
 
'500  If checklst% = 1 Then

   'present list of pro files and require user to check/uncheck those they he desires
   If automatic = False And internet = False Then
      Label1.Caption = " Please check/uncheck the desired places and enter OK"
   Else
      Label1.Caption = " Below are the list of profile files to be processed "
      End If
   Label1.Refresh
   filnetz4% = FreeFile
   Open drivjk$ + "netzskiy.tm4" For Input As #filnetz4%
   nn4% = 0
   netzskiyfm.Netzskiylist(2).Clear
   Do Until EOF(filnetz4%)
      nn4% = nn4% + 1
      Line Input #filnetz4%, doclin1$
      If nn4% > 5 Then
         Line Input #filnetz4%, doclin2$
         doclin$ = doclin1$ & "," & doclin2$
         netzskiyfm.Netzskiylist(2).AddItem doclin$
         If nearyesval = True Or (automatic = True And nearauto = False) Or internet = True Then
            If tblmesag% = 0 Then
              netzskiyfm.Netzskiylist(2).Selected(nn4% - 6) = True
            ElseIf tblmesag% = 1 And InStr(1, doclin$, drivfordtm$ + "netz") <> 0 Then
              netzskiyfm.Netzskiylist(2).Selected(nn4% - 6) = False
            ElseIf tblmesag% = 1 And InStr(1, doclin$, drivfordtm$ + "netz") = 0 Then
              netzskiyfm.Netzskiylist(2).Selected(nn4% - 6) = True
            ElseIf tblmesag% = 2 And InStr(1, doclin$, drivfordtm$ + "skiy") <> 0 Then
              netzskiyfm.Netzskiylist(2).Selected(nn4% - 6) = False
            ElseIf tblmesag% = 2 And InStr(1, doclin$, drivfordtm$ + "skiy") = 0 Then
              netzskiyfm.Netzskiylist(2).Selected(nn4% - 6) = True
              End If
         ElseIf nearyesval = False Then 'never checked
           netzskiyfm.Netzskiylist(2).Selected(nn4% - 6) = True
           End If
         End If
      netzski$(0, nn4%) = doclin1$
      netzski$(1, nn4%) = doclin2$
   Loop
   Close #filnetz4%
   netzskiyok = False
   netzskiyfm.Visible = True
   'ret = SetWindowPos(netzskiyfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   If automatic = True Or internet = True Then
      netzskiyfm.NetzSkiyOkbut0.Value = True
      End If
      
750   dblEndTime = Timer + 2#
      Do While dblEndTime > Timer
        If netzskiyok = True Then GoTo 800
        DoEvents
      Loop
      GoTo 750
800
   lwin = FindWindow(vbNullString, "Netzski3")
   If lwin <> 0 Then GoTo 815
   If internet = True Then 'check that last scan was finished
                           'this protects against simultaneous user
      lognum% = FreeFile
      Open drivjk$ + "calprog.log" For Append As #lognum%
      Print #lognum%, "Step #8: Netzskiyfm closed successfully"
      Close #lognum%
                           
      myfile = Dir(drivfordtm$ + "netz\netzskiy.tm2")
      If myfile <> sEmpty Then
         lognum% = FreeFile
         Open drivjk$ + "calprog.log" For Append As #lognum%
         Print #lognum%, "Warning: ...fordtm\netz\netzskiy.tm2 found, continuing nevertheless"
         Close #lognum%
         End If

      
     myfile = Dir(drivfordtm$ + "skiy\netzskiy.tm2")
     'tim1 = Timer
     If myfile <> sEmpty Then
         lognum% = FreeFile
         Open drivjk$ + "calprog.log" For Append As #lognum%
         Print #lognum%, "Warning: ...fordtm\skiy\netzskiy.tm2 found, continuing nevertheless"
         Close #lognum%
         End If

      End If
      
   filnez3% = FreeFile
   Open drivjk$ + "netzskiy.tm3" For Output As #filnez3%
   For i% = 1 To nn4%
      If i% <= 2 Or (i% > 3 And i% <= 5) Then
         Print #filnez3%, netzski$(0, i%) 'Hebrew Year, type of table, city abbreviaton
      ElseIf i% = 3 Then
         Write #filnez3%, numchecked% 'number checked
      ElseIf i% > 5 And nchecked%(i% - 5) = 1 Then
         Print #filnez3%, netzski$(0, i%) 'filename of profile file
         Print #filnez3%, netzski$(1, i%) 'coordinates,hgt,year, etc
         'determine minimum and average temperatures for these coordinates
         If chkOldCalcMethod.Value = vbUnchecked Then GoSub AddTemps
         'now add distlim
         Write #filnez3%, AddObsTime, CInt(outdistlim), obscushion 'distlim
         End If
   Next i%
   Close #filnez3%
   myfile = Dir(drivjk$ + "netzskiy.tm4")
   If myfil <> sEmpty Then Kill drivjk$ + "netzskiy.tm4"
   If internet = True Then
      lognum% = FreeFile
      Open drivjk$ + "calprog.log" For Append As #lognum%
      Print #lognum%, "Step #9: Finished writing file: netzskiy.tm3"
      Close #lognum%
      End If
815:
   suntop% = SunriseSunset.Top
   'SunriseSunset.Top = 3200
   ''erase old stat files if they exist
   'For i% = 1 To ntmp% + 10
   '   filstat$ = drivjk$+":\jk\stat"
   '   If i% <= 9 Then
   '     filstat$ = filstat$ + "00" + LTrim$(CStr(i%)) + ".tmp"
   '   ElseIf i >= 10 And i% < 100 Then
   '     filstat$ = filstat$ + "0" + LTrim$(CStr(i%)) + ".tmp"
   '   ElseIf i% >= 100 Then
   '     filstat$ = filstat$ + LTrim$(CStr(i%)) + ".tmp"
   '     End If
   '   myfile = Dir(filstat$)
   '   If myfile <> sEmpty Then Kill filstat$
   'Next i%
   Label1.Caption = " Calculating...Please Wait."
   Label1.Refresh
   'Label1.Enabled = False
   nstat% = 0
   ProgressBar1.Enabled = True
   ProgressBar1.Visible = True
   ProgressBar1.Value = 0
   'Label3.Visible = True
   'Label3.Enabled = True
   'If calnode.Visible = True Then GoTo 825
   If eroscityflag = True Then GoTo 825

   Caldirectories.Visible = False
825: If startedscan = True Then Exit Sub
   'RetVal = WinExec(drivjk$ + "Netzski3.exe", 6) ' Run netzski3 as DOS shell (VB Shell function can't be run twice in succession--a bug?)
   'RetVal = Shell(drivjk$ + "Netzski3.exe", 6) ' Run netzski3 as DOS shell
   If internet = True Then 'check if there is an almost simultaneous client
      'don't start new scan until fordtm/netz and fordtm/skiy are empty
      lwin = FindWindow(vbNullString, ProgExec$)
      begintim = Timer
      Do Until lwin = 0
         DoEvents
         lwin = FindWindow(vbNullString, ProgExec$)
         diftim = Timer - begintim
         If diftim > 300 Then
            lognum% = FreeFile
            Open drivjk$ + "calprog.log" For Append As #lognum%
            Print #lognum%, ProgExec$ & ".exe from different job is running. Abort this program!"
            Close #lognum%
              
            myfile = Dir(drivfordtm$ + "busy.cal")
            If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
            For i% = 0 To Forms.Count - 1
              Unload Forms(i%)
            Next i%
      
            myfile = Dir(drivfordtm$ + "busy.cal")
            If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
          
            'kill timer
            If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)

            'end program abruptly
            End
            End If
      Loop
      'now wait until netzend is erased--meaning that earlier client has finished
      'this assumes that two completely simultaneous users is highly unlikely
      myfile = Dir(drivjk$ + "netzend.tmp")
      Do Until myfile = sEmpty
         DoEvents
         myfile = Dir(drivjk$ + "netzend.tmp")
      Loop
      
      End If
      
   Screen.MousePointer = vbDefault
   If Katz = True Then
      If chkOldCalcMethod.Value = vbChecked Then
        RetVal = Shell(drivjk$ & ProgExec$ & "_XP" & ".exe", 6) ' Run netzski3 as DOS shell
      Else
        RetVal = Shell(drivjk$ & ProgExec$ & ".exe", 6) ' Run netzski3 as DOS shell
        End If
   Else
      If internet = False Then
         If chkOldCalcMethod.Value = vbChecked Then
            RetVal = Shell(drivjk$ + ProgExec$ & "_XP" & ".exe", 6) ' Run netzski3 as DOS shell
         Else
            RetVal = Shell(drivjk$ + ProgExec$ & ".exe", 6) ' Run netzski3 as DOS shell
            End If
      Else
         lognum% = FreeFile
         Open drivjk$ + "calprog.log" For Append As #lognum%
         Print #lognum%, "Step #10: Executing Netzski4/5/6.exe as DOS shell"
         Close #lognum%

         If chkOldCalcMethod.Value = vbChecked Then
            RetVal = Shell(drivjk$ + ProgExec$ & "_XP" & ".exe", 6) ' Run netzski3 as DOS shell
         Else
            RetVal = Shell(drivjk$ + ProgExec$ & ".exe", 6) ' Run netzski3 as DOS shell
            End If
         End If
      
      Do Until RetVal <> 0
         DoEvents
      Loop
      End If
   startedscan = True
   ntmp% = 0
   timerwait = Timer + 1#
   Do While timerwait > Timer 'wait 1 sec before enabling timer to avoid file use conflicts
      'Timer1.Enabled = True
      DoEvents
   Loop
850 Timer1.Enabled = True
GoTo 900
   
errinternet:
      errlog% = FreeFile
      Open drivjk$ + "Cal_OKbh.log" For Output As errlog%
      Print #errlog%, "Cal Prog exited from SunriseSunset: Can't find the city directory"
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


OKbuterhand:
   geo = False
   Close
   If internet = True And Err.Number >= 0 Then
      'abort the program with a error messages
      errlog% = FreeFile
      Open drivjk$ + "Cal_OKbh.log" For Output As errlog%
      Print #errlog%, "Cal Prog exited from SunriseSunset with runtime error message " + Str(Err.Number)
      Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
      Close #errlog%
      Close
      
     'unload forms
      For i% = 0 To Forms.Count - 1
        Unload Forms(i%)
      Next i%
      
      'kill the timer
      ktimer% = 0
oke5: If lngTimerID <> 0 Then
         lngTimerID = KillTimer(0, lngTimerID)
         If lngTimerID <> 0 Then
            errlog% = FreeFile
            ktimer% = ktimer% + 1
            Open drivjk$ + "Cal_OKbh.log" For Append As errlog%
            Print #errlog%, sEmpty
            Print #errlog%, "Can't kill the timer, lngTimerID = ", lngTimerID, "trial #: ", ktimer%
            Close #errlog%
            If ktimer% < 6 Then GoTo oke5 'try again
            End If
         End If
      
      'bring program to abrupt end
      End
      End If
   
   If Err.Number = 70 Then
     lwin = FindWindow(vbNullString, "Netzski3")
     If lwin <> 0 Then
        MsgBox "Please wait until netzski3 window is closed", vbOKOnly + vbExclamation, "Cal Programs"
        Do Until lwin = 0
           DoEvents
           lwin = FindWindow(vbNullString, "Netzski3")
        Loop
     Else
        MsgBox "Please wait until required file is available for opening", vbOKOnly + vbExclamation, "Cal Programs"
        End If
     Resume
   ElseIf Err.Number = 13 And tstfil% = -5 Then
     'nearest approach number out of scale due to old RDHALBAT.FOR error
     apprn = 10
     Resume Next
   ElseIf (Err.Number >= 68 And Err.Number < 70) Or Err.Number = 71 Then
      MsgBox "SunriseSunset encountered error " + CStr(Err.Number) + " while trying to read the CD-ROM drive, make sure that the CD-ROM containing the PROM direc. is loaded properly.", vbExclamation, "Cal Program"
      GoTo 200
   ElseIf (Err.Number = 68) Then
      MsgBox "SunriseSunset could not find the PROM directory, check the CD-ROM disk", vbExclamation, "Cal Program"
      GoTo 200
   Else 'other error display error message and start again
      If Err.Number = 52 Then
         MsgBox "SunriseSunset can't read the CD-ROM!, start from the beginning!", vbExclamation, "Cal Program"
      ElseIf Err.Number = 53 Then
         MsgBox "SunriseSunset couldn't find the file: " & filchk$ & " listed in the .BAT file! Start from the beginning.", vbExclamation, "Cal Program"
         Close
         SunriseSunset.Visible = False
         Caldirectories.Label1.Enabled = True
         Caldirectories.Drive1.Enabled = True
         Caldirectories.Dir1.Enabled = True
         'Caldirectories.List1.Enabled = True
         Caldirectories.Text1.Enabled = True
         Caldirectories.OKbutton.Enabled = True
         Caldirectories.ExitButton.Enabled = True
         Caldirectories.OKbutton.Enabled = True
         'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         astronplace = False
         eros = False
      Else
         MsgBox "SunriseSunset encountered unexpected error number" + Str$(Err.Number) + " encountered, start from the beginning.", vbCritical, "Cal Program"
         End If
      SunriseSunset.OKbut0.Value = False
      SunriseSunset.Cancelbut.Value = True
      Timer1.Enabled = False
      SunriseSunset.Visible = False
      Caldirectories.Visible = True
      Caldirectories.Label1.Enabled = True
      Caldirectories.Drive1.Enabled = True
      Caldirectories.Dir1.Enabled = True
      'Caldirectories.List1.Enabled = True
      Caldirectories.Text1.Enabled = True
      Caldirectories.OKbutton.Enabled = True
      Caldirectories.ExitButton.Enabled = True
      Caldirectories.OKbutton.Enabled = True
      'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      astronplace = False
      eros = False
      Exit Sub
      End If
900   'Do '    Exit Sub
      '  DoEvents
      'Loop
      GoTo 999
      
'-------------------inline gosubs----------------------------------
      
'lpyr: 'determine if it is Hebrew Leap Year
'    difdec% = 0
'    stdyy% = 299    'English date paramters of Rosh Hashanoh year 1
'    'stdyy% = 275   'English date paramters of Rosh Hashanoh 5758
'    styr% = -3760
'    'styr% = 1997
'    yl% = 366
'    'yl% = 365
'    rh2% = 2       'Rosh Hashanoh year 1 is on day 2 (Monday)
'    'rh2% = 5      'Rosh Hashanoh 5758 is on day 5 (Thursday)
'    yrr% = 2     'year 1 and year 5758 is a regular kesidrah year of 354 days
'    leapyear% = 0
'    ncal1% = 2: ncal2% = 5: ncal3% = 204  'molad of Tishri 1 year 1-- day;hour;chelakim
'    'ncal1% = 5: ncal2% = 4: ncal3% = 129  'molad of Tishri 1 5758 day;hour;chelakim
'    nt1% = ncal1%: nt2% = ncal2%: nt3% = ncal3%: n1rhoo% = nt1%
'    leapyr2% = leapyear%
'    n1yreg% = 4: n2yreg% = 8: n3yreg% = 876 'change in molad after 12 lunations of reg. year
'    n1ylp% = 5: n2ylp% = 21: n3ylp% = 589 'change in molad after 13 lunations of leap year
'    n1mon% = 1: n2mon% = 12: n3mon% = 793 'monthly change in molad after 1 lunation
'    n11% = ncal1%: n22% = ncal2%: n33% = ncal3%  'initialize molad
'
'    'chosen year to calculate monthly moladim
'    yrstep% = yrheb% - 1
'    'yrstep% = yrheb% - 5758
'    nyear% = 0: flag% = 0
'    For kyr% = 1 To yrstep%
'       nyear% = nyear% + 1
'       nnew% = 1
'       GoSub newdate
'       n1rhooo% = n1rho%
'    Next kyr%
'    'now calculate molad of Tishri 1 of next year in order to
'    'determine if the desired year is choser,kesidrah, or sholem
'    leapyr2% = leapyear%
'    nyear% = nyear% + 1
'    flag% = 1
'    nnew% = 0
'    GoSub newdate
'    'now calculate english date and molad of each rosh chodesh of desired year, yr%
'    n1rh% = n1rhoo%: n2rh% = nt2%: n3rh% = nt3%
'
'    constdif% = n1rhooo% - rh2%
'    rhday% = rh2%: If rhday% = 0 Then rhday% = 7
'
'    'record year datum to be used for parshiot and holidays
'    dayRoshHashono% = rhday%
'    yeartype% = yrr%
'    hebleapyear = False
'    If leapyear% = 1 Then hebleapyear = True
'
'    'load up names of shabbos torah reading
'    LoadParshiotNames
'
'Return
'
'newdate:
'    If nyear% = 20 Then nyear% = 1
'    Select Case nyear%
'       Case 3, 6, 8, 11, 14, 17, 19
'          leapyear% = 1
'          n111% = n1ylp%: n222% = n2ylp%: n333% = n3ylp%
'       Case Else
'          leapyear% = 0
'          n111% = n1yreg%: n222% = n2yreg%: n333% = n3yreg%
'    End Select
'    n33% = n33% + n333%
'    cal3 = n33% / 1080
'    ncal3% = CInt((cal3 - Fix(cal3)) * 1080)
'    n22% = n22% + n222%
'    cal2 = (n22% + Fix(cal3)) / 24
'    ncal2% = CInt((cal2 - Fix(cal2)) * 24)
'    n11% = n11% + n111%
'    cal1 = (n11% + Fix(cal2)) / 7
'    ncal1% = CInt((cal1 - Fix(cal1)) * 7)
'    n11% = ncal1%
'    n22% = ncal2%
'    n33% = ncal3%   'molad of Tishri 1 of this iteration
'    'now use dechiyos to determine which day of week Rosh Hashanoh falls on
'    n1rh% = n11%: n2rh% = n22%: n3rh% = n33%
'    'difd% = 0
'    n1rho% = n1rh%
'    Select Case nyear%
'       Case 3, 6, 8, 11, 14, 17, 19
'          If n11% = 2 And n22% + n33% / 1080 > 15.545 Then
'             n1rh% = n1rh% + 1
'             GoTo nd500
'             End If
'    End Select
'    If n2rh% >= 18 Then
'       n1rh% = n1rh% + 1
'       If n1rh% = 8 Then n1rh% = 1
'       If n1rh% = 1 Or n1rh% = 4 Or n1rh% = 6 Then
'          n1rh% = n1rh% + 1
'          End If
'       GoTo nd500
'       End If
'    If n1rh% = 1 Or n1rh% = 4 Or n1rh% = 6 Then
'       n1rh% = n1rh% + 1
'       End If
'    If (flag% = 0 Or flag% = 1 Or kyr% = yrstep%) And n1rh% = 3 And n2rh% + n3rh% / 1080 > 9.188 Then
'       Select Case nyear% + 1
'          Case 20, 2, 4, 5, 7, 9, 12, 13, 15, 16, 18
'             n1rh% = 5
'       End Select
'       End If
'    If n1rh% = 0 Then n1rh% = 7
'
'nd500: If rh2% >= n1rh% Then difrh% = 7 - rh2% + n1rh%
'       If rh2% < n1rh% Then difrh% = n1rh% - rh2%
'       If nnew% = 1 Then n1rhoo% = n1rh%
'       If (leapyear% = 0 And difrh% = 3) Or (leapyear% = 1 And difrh% = 5) Then
'          yrr% = 1
'       ElseIf (leapyear% = 0 And difrh% = 4) Or (leapyear% = 1 And difrh% = 6) Then
'          yrr% = 2
'       ElseIf (leapyear% = 0 And difrh% = 5) Or (leapyear% = 1 And difrh% = 7) Then
'          yrr% = 3
'          End If
'       If leapyear% = 0 Then
'          If yrr% = 1 Then difdyy% = 353
'          If yrr% = 2 Then difdyy% = 354
'          If yrr% = 3 Then difdyy% = 355
'       ElseIf leapyear% = 1 Then
'          If yrr% = 1 Then difdyy% = 383
'          If yrr% = 2 Then difdyy% = 384
'          If yrr% = 3 Then difdyy% = 385
'          End If
'       If flag% <> 1 Then
'          dyy% = stdyy% + difdyy% - yl% '- difdec%
'          stdyy% = dyy%
'          styr% = styr% + 1
'          yd% = styr% - 1988
'          yl% = 365
'          If yd% Mod 4 = 0 Then yl% = 366
'          If yd% Mod 4 = 0 And styr% Mod 100 = 0 And styr% Mod 400 <> 0 Then yl% = 365
'          rh2% = n1rh%
'          leapyr2% = leapyear%
'          nt1% = n11%
'          nt2% = n22%
'          nt3% = n33%
'          End If
'Return

lpyr:

monthe$(1) = "Jan-"
monthe$(2) = "Feb-"
monthe$(3) = "Mar-"
monthe$(4) = "Apr-"
monthe$(5) = "May-"
monthe$(6) = "Jun-"
monthe$(7) = "Jul-"
monthe$(8) = "Aug-"
monthe$(9) = "Sep-"
monthe$(10) = "Oct-"
monthe$(11) = "Nov-"
monthe$(12) = "Dec-"
monthh$(1, 1) = "Tishrey"
monthh$(1, 2) = "Chesvan"
monthh$(1, 3) = "Kislev"
monthh$(1, 4) = "Teves"
monthh$(1, 5) = "Shvat"
monthh$(1, 6) = "Adar I"
monthh$(1, 7) = "Adar II"
monthh$(1, 8) = "Nisan"
monthh$(1, 9) = "Iyar"
monthh$(1, 10) = "Sivan"
monthh$(1, 11) = "Tamuz"
monthh$(1, 12) = "Av"
monthh$(1, 13) = "Elul"
monthh$(1, 14) = "Adar"

If internet Then
   '------------------------using year 5758 (1997) as reference-------------
   sthebyr% = 5758
   'Use Hebrew Year 5758 as reference year
   difdec% = 0
   stdyy% = 275   'English date paramters of Rosh Hashanoh 5758
   styr% = 1997
   yl% = 365
   rh2% = 5      'Rosh Hashanoh 5758 is on day 5 (Thursday)
   yrr% = 2     '5758 is a regular kesidrah year of 354 days
   leapyear% = 0
   ncal1% = 5: ncal2% = 4: ncal3% = 129  'molad of Tishri 1 5758 day;hour;chelakim
   '--------------using year 1 (-3760) as reference---------------
ElseIf Not internet Then
   sthebyr% = 1
   'Use Hebrew year 1 as the reference year
   difdec% = 0
   'English date paramters of Rosh Hashanoh year 1 (see E.S. 12.31) are:
   stdyy% = 299 'Rosh Hashonah of Hebrew year 1 occured October 7 = day number 299
   styr% = -3760 'Rosh Hashonah of Hebrew year 1 occured on the year -3760 of their counting
   yl% = 366 'that year was a leap year in the English calendar
   rh2% = 2       'Rosh Hashanoh year 1 is on day 2 (Monday)
   yrr% = 2     'year 1 is a regular kesidrah year of 354 days
   leapyear% = 0
   ncal1% = 2: ncal2% = 5: ncal3% = 204  'molad of Tishri 1 year 1-- day;hour;chelakim
   '--------------------------------------------------------------------
   End If

nt1% = ncal1%: nt2% = ncal2%: nt3% = ncal3%: n1rhoo% = nt1%
leapyr2% = leapyear%
n1yreg% = 4: n2yreg% = 8: n3yreg% = 876 'change in molad after 12 lunations of reg. year
n1ylp% = 5: n2ylp% = 21: n3ylp% = 589 'change in molad after 13 lunations of leap year
n1mon% = 1: n2mon% = 12: n3mon% = 793 'monthly change in molad after 1 lunation
n11% = ncal1%: n22% = ncal2%: n33% = ncal3%  'initialize molad

yr% = yrheb
'Cls
'chosen year to calculate monthly moladim
'yrstep% = yrheb% - 1
yrstep% = yr% - sthebyr%

nyear% = 0: flag% = 0
For kyr% = 1 To yrstep%
   nyear% = nyear% + 1
   nnew% = 1
   GoSub newdate
   n1rhooo% = n1rho%
Next kyr%
'now calculate molad of Tishri 1 of next year in order to
'determine if the desired year is choser,kesidrah, or sholem
leapyr2% = leapyear%
nyear% = nyear% + 1
flag% = 1
nnew% = 0
GoSub newdate
'now calculate english date and molad of each rosh chodesh of desired year, yr%
n1rh% = n1rhoo%: n2rh% = nt2%: n3rh% = nt3%

constdif% = n1rhooo% - rh2%
rhday% = rh2%: If rhday% = 0 Then rhday% = 7

    'record year datum to be used for parshiot and holidays
    dayRoshHashono% = rhday%
    yeartype% = yrr%
    hebleapyear = False
    If leapyear% = 1 Then hebleapyear = True
    
    'load up names of shabbos torah reading
    LoadParshiotNames


GoSub dmh

dyy% = stdyy% '- difdec%
hdryr$ = "-" + Trim$(Str$(styr%))
newschulyr% = 0
GoSub engdate
iheb% = 1
mdates$(1, 1) = monthh$(iheb%, 1): mdates$(2, 1) = dates$: mmdate%(1, 1) = dyy%

'start and end day numbers of the first year
yrstrt%(0) = dyy%
yrend%(0) = 365: If leapyr% = 1 Then yrend%(0) = 366

'If newhebcalfm.Check4.Value = vbChecked Then
   'calculate dyy% for first shabbos
   fshabos0% = 7 - rhday% + dyy%
'   End If

'now calculate other molados and their english date
endyr% = 12: If leapyear% = 1 Then endyr% = 13
'If magnify = True Then GoTo 250
For k% = 2 To endyr%
   
   If k% < 6 Then
      monthhh$ = monthh$(1, k%)
   ElseIf k% = 6 Then
      If leapyear% = 0 Then
         monthhh$ = monthh$(1, 14)
      Else
         monthhh$ = monthh$(1, 6)
         End If
   ElseIf k% >= 7 Then
      If leapyear% = 0 Then
         monthhh$ = monthh$(1, k% + 1)
      Else
         monthhh$ = monthh$(1, k%)
         End If
      End If
         
   n33% = n3rh% + n3mon%
   cal3 = n33% / 1080
   ncal3% = CInt((cal3 - Fix(cal3)) * 1080)
   n22% = n2rh% + n2mon%
   cal2 = (n22% + Fix(cal3)) / 24
   ncal2% = CInt((cal2 - Fix(cal2)) * 24)
   n11% = n1rh% + n1mon%
   cal1 = (n11% + Fix(cal2)) / 7
   ncal1% = CInt((cal1 - Fix(cal1)) * 7)
   n1rh% = ncal1%
   n2rh% = ncal2%
   n3rh% = ncal3%
   GoSub dmh
   n1day% = n1rh%: If n1day% = 0 Then n1day% = 7
    If k% = 2 Then
      dyy% = dyy% + 30
   ElseIf k% = 3 Then
      If yrr% <> 3 Then dyy% = dyy% + 29
      If yrr% = 3 Then dyy% = dyy% + 30
   ElseIf k% = 4 Then
      If yrr% = 1 Then dyy% = dyy% + 29
      If yrr% <> 1 Then dyy% = dyy% + 30
   ElseIf k% = 5 Then
      dyy% = dyy% + 29
   ElseIf k% = 6 Then
      dyy% = dyy% + 30
   ElseIf k% >= 7 And leapyear% = 0 Then
      If k% = 7 Then dyy% = dyy% + 29
      If k% = 8 Then dyy% = dyy% + 30
      If k% = 9 Then dyy% = dyy% + 29
      If k% = 10 Then dyy% = dyy% + 30
      If k% = 11 Then dyy% = dyy% + 29
      If k% = 12 Then dyy% = dyy% + 30
   ElseIf k% >= 7 And leapyear% = 1 Then
      If k% = 7 Then dyy% = dyy% + 30
      If k% = 8 Then dyy% = dyy% + 29
      If k% = 9 Then dyy% = dyy% + 30
      If k% = 10 Then dyy% = dyy% + 29
      If k% = 11 Then dyy% = dyy% + 30
      If k% = 12 Then dyy% = dyy% + 29
      If k% = 13 Then dyy% = dyy% + 30
      End If
   hdryr$ = "-" + Trim$(Str$(styr%))
   dyy% = dyy% - 1
   GoSub engdate
   mdates$(2, k% - 1) = dates$: mmdate%(2, k% - 1) = dyy%
   dyy% = dyy% + 1
   GoSub engdate
   mdates$(1, k%) = monthhh$: mdates$(2, k%) = dates$: mmdate%(1, k%) = dyy%
 Next k%
250 If styr% = 1997 Then styr% = 1998
dyy% = dyy% + 28: hdryr$ = "-" + Trim$(Str$(styr%))
GoSub engdate ': LOCATE 13, 5: Print "end of year: "; dates$
mdates$(2, endyr%) = dates$: mmdate%(2, endyr%) = dyy%
yrend%(1) = dyy%
Return

newdate:
    If nyear% = 20 Then nyear% = 1
    Select Case nyear%
       Case 3, 6, 8, 11, 14, 17, 19
          leapyear% = 1
          n111% = n1ylp%: n222% = n2ylp%: n333% = n3ylp%
       Case Else
          leapyear% = 0
          n111% = n1yreg%: n222% = n2yreg%: n333% = n3yreg%
    End Select
    n33% = n33% + n333%
    cal3 = n33% / 1080
    ncal3% = CInt((cal3 - Fix(cal3)) * 1080)
    n22% = n22% + n222%
    cal2 = (n22% + Fix(cal3)) / 24
    ncal2% = CInt((cal2 - Fix(cal2)) * 24)
    n11% = n11% + n111%
    cal1 = (n11% + Fix(cal2)) / 7
    ncal1% = CInt((cal1 - Fix(cal1)) * 7)
    n11% = ncal1%
    n22% = ncal2%
    n33% = ncal3%   'molad of Tishri 1 of this iteration
    'now use dechiyos to determine which day of week Rosh Hashanoh falls on
    n1rh% = n11%: n2rh% = n22%: n3rh% = n33%
    'difd% = 0
    n1rho% = n1rh%
    Select Case nyear%
       Case 3, 6, 8, 11, 14, 17, 19
          If n11% = 2 And n22% + n33% / 1080 > 15.545 Then
             n1rh% = n1rh% + 1
             'difd% = 1
             GoTo 500
             End If
    End Select
    If n2rh% >= 18 Then
       n1rh% = n1rh% + 1
       'difd% = 1
       If n1rh% = 8 Then n1rh% = 1
       If n1rh% = 1 Or n1rh% = 4 Or n1rh% = 6 Then
          n1rh% = n1rh% + 1
          End If
       GoTo 500
       End If
    If n1rh% = 1 Or n1rh% = 4 Or n1rh% = 6 Then
       n1rh% = n1rh% + 1
       'difd% = 1
       End If
    'GOTO 500

'    IF (leapyear% <> 1 AND flag% = 0) AND n1rh% = 3 AND n2rh% + n3rh% / 1080 > 9.188 THEN
'       n1rh% = 5
'       END IF
    If (flag% = 0 Or flag% = 1 Or kyr% = yrstep%) And n1rh% = 3 And n2rh% + n3rh% / 1080 > 9.188 Then
       Select Case nyear% + 1
          Case 20, 2, 4, 5, 7, 9, 12, 13, 15, 16, 18
             n1rh% = 5
       End Select
       End If
    If n1rh% = 0 Then n1rh% = 7

500    If rh2% >= n1rh% Then difrh% = 7 - rh2% + n1rh%
       If rh2% < n1rh% Then difrh% = n1rh% - rh2%
       If nnew% = 1 Then n1rhoo% = n1rh%
       If (leapyear% = 0 And difrh% = 3) Or (leapyear% = 1 And difrh% = 5) Then
          yrr% = 1
       ElseIf (leapyear% = 0 And difrh% = 4) Or (leapyear% = 1 And difrh% = 6) Then
          yrr% = 2
       ElseIf (leapyear% = 0 And difrh% = 5) Or (leapyear% = 1 And difrh% = 7) Then
          yrr% = 3
          End If
       If leapyear% = 0 Then
          If yrr% = 1 Then difdyy% = 353
          If yrr% = 2 Then difdyy% = 354
          If yrr% = 3 Then difdyy% = 355
       ElseIf leapyear% = 1 Then
          If yrr% = 1 Then difdyy% = 383
          If yrr% = 2 Then difdyy% = 384
          If yrr% = 3 Then difdyy% = 385
          'INPUT "cr", crr$
          End If
          
'////////////////////////end calculation type of year occurance////////////////
'       If yrr% = 1 Then
'          countyear% = countyear% + 1
'         Select Case leapyear%
'            Case "0" 'Hebrew calendar non-leapyear
'               Iparsha$ = "[0_"
'            Case "1" 'Hebrew calendar leapyear
'               Iparsha$ = "[1_"
'         End Select
'         Select Case rh2% 'day of the week of RoshHashono
'            Case 2 'Monday
'               Iparsha$ = Iparsha$ & "2_"
'            Case 3 'Tuesday
'               Iparsha$ = Iparsha$ & "3_"
'            Case 5 'Thursday
'               Iparsha$ = Iparsha$ & "5_"
'            Case 7 'Shabbos
'               Iparsha$ = Iparsha$ & "7_"
'         End Select
'         Select Case yrr% 'chaser, kesidrah, shalem
'            Case 1 'chaser
'               Iparsha$ = Iparsha$ & "1"
'            Case 2 'kesidrah
'               Iparsha$ = Iparsha$ & "2"
'            Case 3 'shalem
'               Iparsha$ = Iparsha$ & "3"
'         End Select
'
'          If Iparsha$ = "[0_7_1" Then
'             freefil% = FreeFile
'             Open "c:\jk_c\countyears.txt" For Append As #freefil%
'             Print #freefil%, styr% + 3761, styr% + 1
'             Close #freefil%
'             End If
'
'          End If
'////////////////////////end calculation type of year occurance////////////////
             
       If flag% <> 1 Then
          dyy% = stdyy% + difdyy% - yl% '- difdec%
          stdyy% = dyy%
          styr% = styr% + 1
          yd% = styr% - 1988
          yl% = 365
          'LOCATE 18, 1: PRINT "styr%,dy;difdyy%="; styr%; dyy%; difdyy%
          'LOCATE 19, 1: PRINT "rh2%,n1rh%;leapyear;leapyr2%="; rh2%; n1rh%; leapyear%; leapyr2%
          'LOCATE 20, 1: PRINT "leapyear%;leapyr2%;year%;kyr%"; leapyear%; leapyr2%; kyr% - 1 + 5758
          'INPUT "cr", crr$
          If yd% Mod 4 = 0 Then yl% = 366
          If yd% Mod 4 = 0 And styr% Mod 100 = 0 And styr% Mod 400 <> 0 Then yl% = 365
          rh2% = n1rh% 'rh2% is the day of the week (1-7) of the Rosh Hashonoh
          leapyr2% = leapyear%
          nt1% = n11%
          nt2% = n22%
          nt3% = n33%
          End If
Return

lpyrcivil:
      yd% = yrheb% - 1988
      yl% = 365
      If yd% Mod 4 = 0 Then yl% = 366
      If yd% Mod 4 = 0 And styr% Mod 100 = 0 And styr% Mod 400 <> 0 Then yl% = 365
Return

engdate:
   newyear% = 0
   ydeng% = styr% - 1988
   yreng% = styr%
   yleng% = 365
   If ydeng% Mod 4 = 0 Then yleng% = 366
   If ydeng% Mod 4 = 0 And yreng% Mod 100 = 0 And yreng% Mod 400 <> 0 Then yleng% = 365
   If dyy% > yleng% Or newschulyr% = 1 Then
      newschulyr% = 0
      myear% = yreng%
      myear0% = myear%
      yreng% = yreng% + 1
      hdryr$ = "-" + Trim$(Str$(yreng%))
      dyy% = dyy% - yleng%
      ydeng% = Abs(Val(hdryr$)) - 1988
      yleng% = 365
      If ydeng% Mod 4 = 0 Then yleng% = 366
      If ydeng% Mod 4 = 0 And yreng% Mod 100 = 0 And yreng% Mod 400 <> 0 Then yleng% = 365
      styr% = yreng%: yl% = yleng%
      newyear% = 1
      End If
   leapyr% = 0
   If yl% = 366 Then leapyr% = 1 'leap years
   If dyy% >= 1 And dyy% < 32 Then dates$ = monthe$(1) + Trim$(Str$(dyy%)) + hdryr$
   If dyy% >= 32 And dyy% < 60 + leapyr% Then dates$ = monthe$(2) + Trim$(Str$(dyy% - 31)) + hdryr$
   If dyy% >= 60 + leapyr% And dyy% < 91 + leapyr% Then dates$ = monthe$(3) + Trim$(Str$(dyy% - 59 - leapyr%)) + hdryr$
   If dyy% >= 91 + leapyr% And dyy% < 121 + leapyr% Then dates$ = monthe$(4) + Trim$(Str$(dyy% - 90 - leapyr%)) + hdryr$
   If dyy% >= 121 + leapyr% And dyy% < 152 + leapyr% Then dates$ = monthe$(5) + Trim$(Str$(dyy% - 120 - leapyr%)) + hdryr$
   If dyy% >= 152 + leapyr% And dyy% < 182 + leapyr% Then dates$ = monthe$(6) + Trim$(Str$(dyy% - 151 - leapyr%)) + hdryr$
   If dyy% >= 182 + leapyr% And dyy% < 213 + leapyr% Then dates$ = monthe$(7) + Trim$(Str$(dyy% - 181 - leapyr%)) + hdryr$
   If dyy% >= 213 + leapyr% And dyy% < 244 + leapyr% Then dates$ = monthe$(8) + Trim$(Str$(dyy% - 212 - leapyr%)) + hdryr$
   If dyy% >= 244 + leapyr% And dyy% < 274 + leapyr% Then dates$ = monthe$(9) + Trim$(Str$(dyy% - 243 - leapyr%)) + hdryr$
   If dyy% >= 274 + leapyr% And dyy% < 305 + leapyr% Then dates$ = monthe$(10) + Trim$(Str$(dyy% - 273 - leapyr%)) + hdryr$
   If dyy% >= 305 + leapyr% And dyy% < 335 + leapyr% Then dates$ = monthe$(11) + Trim$(Str$(dyy% - 304 - leapyr%)) + hdryr$
   If dyy% >= 335 + leapyr% And dyy% < 365 + leapyr% Then dates$ = monthe$(12) + Trim$(Str$(dyy% - 334 - leapyr%)) + hdryr$
   'IF newyear% = 1 AND yl% = 366 THEN dyy% = dyy% - 1
Return

dmh:
   Hourr = n2rh% + 6
   If Hourr < 12 Then
      tm$ = " PM night"
   ElseIf Hourr >= 12 Then
      Hourr = Hourr - 12
      If Hourr < 12 Then
         If Hourr = 0 Then Hourr = 12
         tm$ = " AM"
      ElseIf Hourr >= 12 Then
         Hourr = Hourr - 12
         If Hourr = 0 Then Hourr = 12
         tm$ = " PM afternoon"
         End If
      End If
   minc = n3rh% * (60 / 1080)
   Min = Fix(minc)
   hel = CInt((minc - Min) * 18)
Return

AddTemps:
    'parse coordinate line, extract coordinates, and write corresponding temperatures
    Coords = Split(netzski$(1, i%), ",")
    If UBound(Coords) > 3 Then
        kmxAT = Val(Coords(0))
        kmyAT = Val(Coords(1))
        If geo And eroscountry <> "Israel" Then 'geo coordinates
           lgAT = kmxAT
           ltAT = kmyAT
        Else 'EY old ITM coordinates
           Call casgeo(kmxAT, kmyAT, lgAT, ltAT)
           lgAT = -lgAT 'this is convention for WorldClim
           End If
        If (eros Or geo) And LCase(eroscountry$) <> "israel" Then
            Call Temperatures(lgAT, -ltAT, MinTK, AvgTK, MaxTK, ier)
        Else
            Call Temperatures(ltAT, lgAT, MinTK, AvgTK, MaxTK, ier)
            End If
        If ier = 0 Then
           Write #filnez3%, MinTK(1), MinTK(2), MinTK(3), MinTK(4), MinTK(5), MinTK(6), MinTK(7), MinTK(8), MinTK(9), MinTK(10), MinTK(11), MinTK(12)
           Write #filnez3%, AvgTK(1), AvgTK(2), AvgTK(3), AvgTK(4), AvgTK(5), AvgTK(6), AvgTK(7), AvgTK(8), AvgTK(9), AvgTK(10), AvgTK(11), AvgTK(12)
           Write #filnez3%, MaxTK(1), MaxTK(2), MaxTK(3), MaxTK(4), MaxTK(5), MaxTK(6), MaxTK(7), MaxTK(8), MaxTK(9), MaxTK(10), MaxTK(11), MaxTK(12)
        ElseIf ier < 0 Then
           'abort
           Close
           Call MsgBox("Couldn't complete writing of temperatures to netzskiy.tm3 file" _
                       & vbCrLf & "Aborting....." _
                       , vbCritical, "Error detected")
           
           Exit Sub
           End If
    Else
        Call MsgBox("Couldn't complete writing of temperatures to netzskiy.tm3 file" _
                    & vbCrLf & "Too few coordinates in coordinate lines" _
                    & vbCrLf & "Aborting....." _
                    , vbCritical, "Error detected")
        
        Exit Sub
        End If
Return

999 End Sub

Private Sub Option1_Click()
   'If Option1.Value = True Then
      Label2.Caption = "Hebrew Year"
      Label2.Refresh
      Combo1.Clear
      For i% = 1 To 6000
         Combo1.AddItem (Trim$(Str$(i%)))
      Next i%
      If yrheb% <> 0 And Option1b = True Then
         Combo1.ListIndex = yrheb% - 1
      Else
         'find current hebrew year
         finddate$ = Date$
         cha$ = sEmpty
         lenyr% = 0
         Do Until cha$ = "-"
           cha$ = Mid$(finddate$, Len(finddate$) - lenyr%, 1)
           lenyr% = lenyr% + 1
         Loop
         presyr% = Val(Mid$(finddate$, Len(finddate$) - lenyr% + 2, lenyr - 1))
         yrheb% = presyr% + RefCivilYear% - RefHebYear% '- 1997 + 5758
         If yrheb% - 1 >= 0 Then
            Combo1.ListIndex = yrheb% - 1
            End If
         End If
      hebcal = True
      Option1b = True
      Option2b = False
   '   End If
End Sub

Private Sub Option2_Click()
  'If Option1.Value = False Then
      Label2.Caption = "Civil Year"
      Label2.Refresh
      Combo1.Clear
      For i% = 1600 To 2240
         Combo1.AddItem (Trim$(Str$(i%)))
      Next i%
      If yrheb% <> 0 And Option1b = True And Katz = False Then 'convert from hebrew year to civil year
         listyr% = yrheb% + RefCivilYear% - RefHebYear% - 1600 '+ 1997 - 5758 - 1600
         If listyr% >= 0 Then
            Combo1.ListIndex = yrheb% + RefCivilYear% - RefHebYear% - 1600 '+ 1997 - 5758 - 1600
         Else
            Combo1.Text = yrheb%
            End If
      Else 'use current year
         If Katz = True Then
            Combo1.ListIndex = 2002 - 1600
            GoTo 50
            End If
         finddate$ = Date$
         cha$ = sEmpty
         lenyr% = 0
         Do Until cha$ = "-"
           cha$ = Mid$(finddate$, Len(finddate$) - lenyr%, 1)
           lenyr% = lenyr% + 1
         Loop
         Combo1.ListIndex = Val(Mid$(finddate$, Len(finddate$) - lenyr% + 2, lenyr - 1)) - 1600
         End If
50:   Option2b = True
      Option1b = False
      hebcal = False
  '    End If
End Sub

Private Sub Option3_Click()
   optionheb = True
End Sub

Private Sub Option4_Click()
   optionheb = False
End Sub

Private Sub Timer1_Timer()
   On Error GoTo errorhandel
   CalMDIform.Visible = False
   Caldirectories.Visible = False
   BringWindowToTop (SunriseSunset.hwnd)
   '*********to speed up the program-activate the following line
   'nstat% = 100: GoTo t80
   '**************************************
   If internet = True Then 'don't read in the stat files
      nstat% = 100
      GoTo t90
      End If
t10:
   myfile = Dir(drivjk$ + "netzend.tmp")
  If myfile <> sEmpty Then
     SunriseSunset.CurrentX = Label3.Left + 100
     SunriseSunset.CurrentY = Label3.Top - 50
     SunriseSunset.ForeColor = Label3.BackColor
     SunriseSunset.FontBold = True
     SunriseSunset.Print "    "
     nstat% = 100
     GoTo t80
     End If

   ntmp% = ntmp% + 1
   filstat$ = drivjk$ + "stat"
   If ntmp% <= 9 Then
     filstat$ = filstat$ + "000" + LTrim$(CStr(ntmp%)) + ".tmp"
   ElseIf ntmp% >= 10 And ntmp% < 100 Then
     filstat$ = filstat$ + "00" + LTrim$(CStr(ntmp%)) + ".tmp"
   ElseIf ntmp% >= 100 And ntmp% < 1000 Then
     filstat$ = filstat$ + "0" + LTrim$(CStr(ntmp%)) + ".tmp"
   ElseIf ntmp% >= 1000 And ntmp% < 10000 Then
     filstat$ = filstat$ + LTrim$(CStr(ntmp%)) + ".tmp"
     End If
   myfile = sEmpty
   myfile = Dir(filstat$)
   If myfile <> sEmpty Then
      If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 15
      If Timer1.Interval > 500 Then Timer1.Interval = Timer1.Interval - 500
      filstatnum% = FreeFile
T50:  'Label3.Visible = True
      Open filstat$ For Input As #filstatnum%
      ntmp% = ntmp% - 1
      SunriseSunset.CurrentX = Label3.Left + 100
      SunriseSunset.CurrentY = Label3.Top - 50
      SunriseSunset.ForeColor = Label3.BackColor
      SunriseSunset.FontBold = True
      SunriseSunset.Print LTrim$(CStr(nstat%) + "%")
      'SunriseSunset.FontBold = False
      Do Until EOF(filstatnum%)
         Input #filstatnum%, stat
         nstat% = Int(stat)
         If hebcal = False Or Option2b = True Then nstat% = 2 * nstat%
         ntmp% = ntmp% + 1
         nstato% = nstat%
      Loop
      Close #filstatnum%
      'Label3.Caption = LTrim$(CStr(nstat%) + "%")
      'Label3.Refresh
'      CurrentX = Label3.Left + 100
'      CurrentY = Label3.Top + 20
'      oldfold& = SunriseSunset.ForeColor
'      If ntmp% > 1 Then
'        'SunriseSunset.DrawMode = 7
'        SunriseSunset.CurrentX = Label3.Left + 100
'        SunriseSunset.CurrentY = Label3.Top + 20
'        SunriseSunset.ForeColor = Label3.BackColor
'        SunriseSunset.Print LTrim$(CStr(nstato%) + "%")
        'SunriseSunset.DrawMode = 1
'        End If
t80:
      SunriseSunset.ForeColor = QBColor(14)
      SunriseSunset.CurrentX = Label3.Left + 100
      SunriseSunset.CurrentY = Label3.Top - 50
      SunriseSunset.Print LTrim$(CStr(nstat%) + "%")
      SunriseSunset.ForeColor = oldfold&
      ProgressBar1.Value = nstat%
t90:  If nstat% = 100 Then
        If internet = False Then
           'erase any old *.tmp files on startup
           mypath = drivjk$ & "*.tmp" ' Set the path.
           myname = LCase(Dir(mypath, vbNormal))   ' Retrieve the first entry.
           Do While myname <> sEmpty   ' Start the loop.
              DoEvents
              If InStr(LCase(myname), "stat") <> 0 Then
                 Kill drivjk$ & myname
                 End If
              myname = Dir
           Loop

            'For i% = 1 To ntmp% 'erase stat#.tmp files
            '  filstat$ = drivjk$ + "stat"
            ' If ntmp% <= 9 Then
            '   filstat$ = filstat$ + "000" + LTrim$(CStr(ntmp%)) + ".tmp"
            ' ElseIf ntmp% >= 10 And ntmp% < 100 Then
            '   filstat$ = filstat$ + "00" + LTrim$(CStr(ntmp%)) + ".tmp"
            ' ElseIf ntmp% >= 100 And ntmp% < 1000 Then
            '   filstat$ = filstat$ + "0" + LTrim$(CStr(ntmp%)) + ".tmp"
            ' ElseIf ntmp% >= 1000 And ntmp% < 10000 Then
            '   filstat$ = filstat$ + LTrim$(CStr(ntmp%)) + ".tmp"
            '   End If
            '   myfile = Dir(filstat$)
            '   If myfile <> sEmpty Then Kill filstat$
            'Next i%
           End If
'        'now look for c:\jk\netzend.tmp which signifies that
'        'NETZKI3 has completed all calculations
'        Waitfm.Visible = True
'        Waitfm.Refresh
'        Waitfm.Frame1.Refresh
'        Waitfm.Label1.Refresh
'        If Dir(drivjk$+":\jk\netzend.tmp") = sEmpty Then
'85         waittim = Timer + 1#
'           Do Until Timer > waittim
'           Loop
'           mydir = Dir(drivjk$+":\jk\netzend.tmp")
'           If mydir <> sEmpty Then GoTo 90
'           GoTo 85
'           End If
        'waittim = Timer + 5#
        'Do Until Timer > waittim
        'Loop
        'Do While Dir(drivjk$+":\jk\netzend.tmp") = sEmpty
        ''   DoEvents
        'Loop
'90      Waitfm.Visible = False
        SunriseSunset.Label1.Caption = captmp$
        SunriseSunset.Refresh
        SunriseSunset.Label1.Enabled = False
        SunriseSunset.Visible = False
        SunriseSunset.Timer1.Enabled = False
        astronplace = False
        If Not calnodevis Then eros = False
        If Not calnodevis Then geo = False
        If Katz = True Then
           If katznum% = 0 Then
              katznum% = 1
              Unload SunriseSunset
              Set SunriseSunset = Nothing
              'erase old stat files if they exist
              'erase any old *.tmp files on startup
              mypath = drivjk$ & "*.tmp" ' Set the path.
              myname = LCase(Dir(mypath, vbNormal))   ' Retrieve the first entry.
              Do While myname <> sEmpty   ' Start the loop.
                 DoEvents
                 If InStr(LCase(myname), "stat") <> 0 Then
                    Kill drivjk$ & myname
                    End If
                 myname = Dir
              Loop
              
              'For i% = 1 To ntmp% + 10
              '  filstat$ = drivjk$ + "stat"
              '  If ntmp% <= 9 Then
              '    filstat$ = filstat$ + "000" + LTrim$(CStr(ntmp%)) + ".tmp"
              '  ElseIf ntmp% >= 10 And ntmp% < 100 Then
              '    filstat$ = filstat$ + "00" + LTrim$(CStr(ntmp%)) + ".tmp"
              '  ElseIf ntmp% >= 100 And ntmp% < 1000 Then
              '    filstat$ = filstat$ + "0" + LTrim$(CStr(ntmp%)) + ".tmp"
              '  ElseIf ntmp% >= 1000 And ntmp% < 10000 Then
              '    filstat$ = filstat$ + LTrim$(CStr(ntmp%)) + ".tmp"
              '    End If
              '  myfile = Dir(filstat$)
              '  If myfile <> sEmpty Then Kill filstat$
              'Next i%
              myfile = Dir(drivjk$ + "netzend.tmp")
              If myfile <> sEmpty Then Kill drivjk$ + "netzend.tmp"
              'change name of files
              myfile = Dir(drivfordtm$ + "netz\*.pl1")
              netzfile$ = drivfordtm$ & "netz\" & myfile
              If Dir(drivfordtm$ & "netz\" & Mid$(myfile, 1, Len(myfile) - 3) & "pl0") <> sEmpty Then
                 Kill drivfordtm$ & "netz\" & Mid$(myfile, 1, Len(myfile) - 3) & "pl0"
                 End If
              Name netzfile$ As drivfordtm$ & "netz\" & Mid$(myfile, 1, Len(myfile) - 3) & "pl0"
              myfile = Dir(drivfordtm$ + "skiy\*.pl1")
              skiyfile$ = drivfordtm$ & "skiy\" & myfile
              If Dir(drivfordtm$ & "skiy\" & Mid$(myfile, 1, Len(myfile) - 3) & "pl0") <> sEmpty Then
                 Kill drivfordtm$ & "skiy\" & Mid$(myfile, 1, Len(myfile) - 3) & "pl0"
                 End If
              Name skiyfile$ As drivfordtm$ & "skiy\" & Mid$(myfile, 1, Len(myfile) - 3) & "pl0"
              katzhebnam$ = hebcityname$
              'ask for next file
              hebcal = False
              AstronForm.Visible = True
              If Katz = True And katztotal% < AstronForm.Combo1.ListCount - 1 Then
                 katztotal% = katztotal% + 1
                 AstronForm.Combo1.ListIndex = katztotal%
                 AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
                 If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
                    AstronForm.Option2.Value = True
                 ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
                    AstronForm.Option1.Value = True
                    End If
              ElseIf Katz = True And katztotal% > AstronForm.Combo1.ListCount - 1 Then
                 katztotal% = 0
                 AstronForm.Combo1.ListIndex = 0
                 AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
                 AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
                 If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
                    AstronForm.Option2.Value = True
                 ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
                    AstronForm.Option1.Value = True
                    End If
                 End If
              
              Exit Sub
           ElseIf katznum% = 1 Then
              katztotal% = katztotal% + 1
              katznum% = 0
              'proceed
              End If
           End If
        newhebcalfm.Visible = True
        'ret = SetWindowPos(newhebcalfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
        If automatic = True Or internet = True Then
           'wait a bit more to make sure netzski3.for finished
           'waittime = Timer + 5#
           'Do While waittime > Timer
           '   DoEvents
           'Loop
           newhebcalfm.Check2.Value = vbChecked
           newhebcalfm.Check3.Value = vbChecked
           newhebcalfm.Check4.Value = vbChecked
           'newhebcalfm.Check5.Value = vbChecked
           newhebcalfm.Check5.Value = vbUnchecked
           If viseros = True Then
              If SRTMflag = 0 Then 'GTOPO30
                 newhebcalfm.Text2 = "45" '45 sec. sunrise cushion' "30"
                 newhebcalfm.Text32 = "-45" 'sunset cushion
              ElseIf SRTMflag = 1 Then 'SRTM-2
                 newhebcalfm.Text2 = "20" '20 second sunrise cushion'
                 newhebcalfm.Text32 = "-20" 'sunset cushion
              ElseIf SRTMflag = 2 Then 'SRTM-1
                 newhebcalfm.Text2 = "35" '35 second sunrise cushion "30" '30 second cushion'
                 newhebcalfm.Text32 = "-35" 'sunset cushion
              ElseIf SRTMflag = 9 And geotz! = 2 Then '<--EY: Eretz Yisroel DTM
                 newhebcalfm.Text2 = "15" '15 sunrise second cushion (default)
                 newhebcalfm.Text32 = "-15" 'sunset cushion
              ElseIf SRTMflag = 9 And geotz! <> 2 Then
                 newhebcalfm.Text2 = "20"
                 newhebcalfm.Text32 = "-20"
                 End If
              End If
              
           If RoundSeconds% <> 0 Then
              newhebcalfm.Text1 = RoundSeconds%
              newhebcalfm.Text31 = RoundSeconds%
              End If
           
           lognum% = FreeFile
           Open drivjk$ + "calprog.log" For Append As #lognum%
           Print #lognum%, "Step #11: Activate Previewfm"
           Close #lognum%
              
           newhebcalfm.newhebPreviewbut.Value = True
           End If
           
        If internet = True Then
           lognum% = FreeFile
           Open drivjk$ + "calprog.log" For Append As #lognum%
           Print #lognum%, "Step #12: Unloading SunriseSunset Form"
           Close #lognum%
           End If
        
        Unload SunriseSunset
        Set SunriseSunset = Nothing
           
        If internet = True Then
           lognum% = FreeFile
           Open drivjk$ + "calprog.log" For Append As #lognum%
           Print #lognum%, "Step #13: SunriseSunset Form unloaded successfully"
           Close #lognum%
           End If
        
        hebtimtot = 0
100     hebtim = Timer + 2#
        Do While hebtim > Timer
           If newhebout = True Then GoTo 150
           DoEvents
        Loop
        hebtimtot = hebtimto + 2
        If internet = True Then
           If hebtimtot < 300 Then
              GoTo 100
           Else
              'terminate this process with an error message
              On Error GoTo unloaderr2
              
              lognum% = FreeFile
              Open drivjk$ + "calprog.log" For Append As #lognum%
              Print #lognum%, "Waited for the scan to finish for at least 5 minutes. Abort this process!"
              Close #lognum%
              
              'check to see if Netzski4/6.exe is hungup
              lwin = FindWindow(vbNullString, ProgExec$)
              If lwin <> sEmpty Then 'Netzski4/6 is hungup
                 lognum% = FreeFile
                 Open drivjk$ + "calprog.log" For Append As #lognum%
                 Print #lognum%, "Netzski4/6 is not advancing!"
                 Close #lognum%
                 'attempt to close the Netzski4 shell
                 'this only works if Netzski4 is a maximized window
                 trial = 0
str10:           lResult = PostMessage(lwin, WM_CANCELMODE, 0, 0&)
                 lResult = PostMessage(lwin, WM_CLOSE, 0, 0)
                 For itr% = 1 To 2
                     ret = SetWindowPos(lwin, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE + SWP_SHOWWINDOW)
                     DoEvents
                 Next itr%
                 'the squeal box is now on top, so answer yes by a tab and a return
                 waitime = Timer
                 Do Until Timer > waitime + 0.5
                    DoEvents
                 Loop
                 Call keybd_event(VK_TAB, 0, 0, 0)
                 Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
                 Call keybd_event(VK_RETURN, 0, 0, 0)
                 Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
                 lwin = FindWindow(vbNullString, ProgExec$)
                 If lwin <> 0 Then
                    lognum% = FreeFile
                    Open drivjk$ + "calprog.log" For Append As #lognum%
                    Print #lognum%, "Tried to kill Netzski4/6 without success!"
                    Close #lognum%
                    End If
                 End If
              
              myfile = Dir(drivfordtm$ + "busy.cal")
              If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
              For i% = 0 To Forms.Count - 1
                Unload Forms(i%)
              Next i%
          
              'kill timer
              If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
    
              'end program abruptly
              End
              End If
        Else
           GoTo 100
           End If
           
150     newhebout = False
        newhebcalfm.Visible = False
        If internet = True Then
           lognum% = FreeFile
           Open drivjk$ + "calprog.log" For Append As #lognum%
           Print #lognum%, "Step #14: Newhebcalfm signaled that it finsihed, unload it now."
           Close #lognum%
           End If
        
        'If calnode.Visible = True Then GoTo 175 'don't want Caldirectories to appear
                                                'rather want calnode to reappear
        If eroscityflag = True Then GoTo 175
                                                
        'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
175:    eros = False
        'erase old stat files if they exist
        If internet = False Then
        
            'erase any old *.tmp files on startup
             mypath = drivjk$ & "*.tmp" ' Set the path.
             myname = LCase(Dir(mypath, vbNormal))   ' Retrieve the first entry.
             Do While myname <> sEmpty   ' Start the loop.
                DoEvents
                If InStr(LCase(myname), "stat") <> 0 Then
                   Kill drivjk$ & myname
                   End If
                myname = Dir
             Loop
        
'            For i% = 1 To ntmp% + 10
'              filstat$ = drivjk$ + "stat"
'              If ntmp% <= 9 Then
'                filstat$ = filstat$ + "000" + LTrim$(CStr(ntmp%)) + ".tmp"
'              ElseIf ntmp% >= 10 And ntmp% < 100 Then
'                filstat$ = filstat$ + "00" + LTrim$(CStr(ntmp%)) + ".tmp"
'              ElseIf ntmp% >= 100 And ntmp% < 1000 Then
'                filstat$ = filstat$ + "0" + LTrim$(CStr(ntmp%)) + ".tmp"
'              ElseIf ntmp% >= 1000 And ntmp% < 10000 Then
'                filstat$ = filstat$ + LTrim$(CStr(ntmp%)) + ".tmp"
'                End If
'                myfile = Dir(filstat$)
'                If myfile <> sEmpty Then Kill filstat$
'            Next i%
            End If
        myfile = Dir(drivjk$ + "netzend.tmp")
        If myfile <> sEmpty Then Kill drivjk$ + "netzend.tmp"
        
        If internet = True Then
           'erase old sunrise/sunset tables
            pos% = InStr(1, servnam$, ".ser")
            servMaxNum& = Val(Mid$(servnam$, 1, pos% - 1))
            mypath = dirint$ & "\"
            myname = Dir(mypath, vbNormal)
            found% = 0
            Do While myname <> sEmpty
               If InStr(1, myname, ".html") <> 0 Then
                  pos1% = InStr(1, myname, ".html")
                  htmlNum& = Val(Mid$(myname, 1, pos1% - 1))
                  If htmlNum& <= servMaxNum& - 50 Then
                     Kill dirint$ & "\" & myname
                     found% = 1
                     End If
                  End If
               myname = Dir
            Loop
            'record action in log file
            lognum% = FreeFile
            Open drivjk$ + "calprog.log" For Append As #lognum%
            If found% = 1 Then
               Print #lognum%, "Step #14.32: html files deleted successfully"
            Else
               Print #lognum%, "Step #14.32: No html files found for deletion"
               End If
            Close #lognum%
           
           'now erase old z'manim tables
           If zmantype% = 0 Then
              zmanext$ = ".csv"
           ElseIf zmantype% = 1 Then
              zmanext$ = ".zip"
           ElseIf zmantype% = 2 Then
              zmanext$ = ".xml"
              End If
           goneback% = 0
ssu500:    mypath = dirint$ & "\"
           myname = Dir(mypath, vbNormal)
           found% = 0
           Do While myname <> sEmpty
               If InStr(1, myname, zmanext$) <> 0 Then
                  pos1% = InStr(1, myname, zmanext$)
                  zmannum& = Val(Mid$(myname, 1, pos1% - 1))
                  If zmannum& <= servMaxNum& - 50 Then
                     Kill dirint$ & "\" & myname
                     found% = 1
                     End If
                  End If
               myname = Dir
            Loop
            'record action in log file
            lognum% = FreeFile
            Open drivjk$ + "calprog.log" For Append As #lognum%
            If found% = 1 Then
               Print #lognum%, "Step #14.31: " & zmanext$ & " files deleted successfully"
            Else
               Print #lognum%, "Step #14.31: No " & zmanext$ & " files found for deletion"
               End If
            Close #lognum%

           If zmantype% = 2 And goneback% = 0 Then 'go back to also erase xsl files
              zmanext$ = ".xsl"
              goneback% = 1
              GoTo ssu500
              End If
              
           On Error GoTo unloaderr
           myfile = Dir(drivfordtm$ + "busy.cal")
           If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
           
           lognum% = FreeFile
           Open drivjk$ + "calprog.log" For Append As #lognum%
           Print #lognum%, "Success! Cal program terminated normally."
           Close #lognum%
           
           For i% = 0 To Forms.Count - 1
             Unload Forms(i%)
           Next i%
           
           'if timer was set, then stop it
           If lngTimerID <> 0 Then
              lngTimerID = KillTimer(0, lngTimerID)
              End If

           'Unload CalMDIform
           End
           End If
        If eroscityflag = True Then Exit Sub
        Caldirectories.Visible = True
        Caldirectories.Label1.Enabled = True
        Caldirectories.Drive1.Enabled = True
        Caldirectories.Dir1.Enabled = True
        'Caldirectories.List1.Enabled = True
        Caldirectories.Text1.Enabled = True
        Caldirectories.OKbutton.Enabled = True
        Caldirectories.ExitButton.Enabled = True
        Caldirectories.OKbutton.Enabled = True
        End If
        If automatic = True Then
           Caldirectories.Runbutton.Value = True
           End If
      Else
         ''ProgressBar1.Value = nstat% 'use old value
         ''Label3.Caption = CStr(nstat%) + "%"
         'SunriseSunset.Label3.Visible = False
         ntmp% = ntmp% - 1
         Timer1.Interval = Timer1.Interval + 80
         If Timer1.Interval >= 1000 Then
            If internet = True Then
               'fatal error, abort Cal Program
               'this probably is caused by two clients trying to
               'obtain tables almost simultaneously.
               Close
               On Error GoTo unloaderr
               myfile = Dir(drivfordtm$ + "busy.cal")
               If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
               
               lognum% = FreeFile
               Open drivjk$ + "calprog.log" For Append As #lognum%
               Print #lognum%, "Netzski4/6 does not seem to be progressing. Abort this process."
               Close #lognum%
               
               For i% = 0 To Forms.Count - 1
                 Unload Forms(i%)
               Next i%
               
               'kill timer
               If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
               
               'end program abruptly
               End
               Exit Sub
               
unloaderr:
               lognum% = FreeFile
               Open drivjk$ + "calprog.log" For Append As #lognum%
               Print #lognum%, "While unloading forms, encountered error number: " & Str(Err.Number)
               Print #lognum%, "Cal Program was therefore terminated prematurely"
               Close #lognum%
               End
               
unloaderr2:
               lognum% = FreeFile
               Open drivjk$ + "calprog.log" For Append As #lognum%
               Print #lognum%, "While unloading forms, encountered error number: " & Str(Err.Number)
               Print #lognum%, "Cal Program was therefore terminated prematurely"
               Close #lognum%
               End
               'Unload CalMDIform
               End If
            response = MsgBox("Netzski6 doesn't seem to be advancing." & vbLf & vbLf & _
                              "This may mean that the azimuthal range was" & vbLf & _
                              "insufficient to calculate the solar ephemerals." & vbLf & _
                              "You can test this by running c:\jk\netzski3/6.exe and looking for error messages." & vbLf & _
                              "In the meantime do you want to abort?", vbCritical + vbOKCancel, "Cal Program")
            If response = vbOK Then
               Close
               If automatic = True Then
                  Caldirectories.AutoCancelbut.Value = True
                  Caldirectories.Text2.Text = newpagenum% + autonum% + 1
                  End If
               Timer1.Enabled = False
               SunriseSunset.Cancelbut.Value = True
                SunriseSunset.Visible = False
                Caldirectories.Visible = True
                Caldirectories.Label1.Enabled = True
                Caldirectories.Drive1.Enabled = True
                Caldirectories.Dir1.Enabled = True
                Caldirectories.Text1.Enabled = True
                Caldirectories.OKbutton.Enabled = True
                Caldirectories.ExitButton.Enabled = True
                Caldirectories.OKbutton.Enabled = True
                astronplace = False
                eros = False
             Else 'give it another try, or see if the end was signaled
               Timer1.Interval = Timer1.Interval - 500
               End If
            End If
         End If
      
      GoTo 200

errorhandel:
            If internet = True Then
               lognum% = FreeFile
               Open drivjk$ + "calprog.log" For Append As #lognum%
               Print #lognum%, "SunriseSunset in errorhandel: encountered error number: " & Str(Err.Number)
               Print #lognum%, "Aborting program"
               Close #lognum%
               
               For i% = 0 To Forms.Count - 1
                 Unload Forms(i%)
               Next i%
               'Unload CalMDIform
      
               myfile = Dir(drivfordtm$ + "busy.cal")
               If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
               
               'kill timer
               If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
               
               'end program
               End
               End If
            If (Err.Number >= 52 And Err.Number <= 63) Or Err.Number = 75 Then
               Close #filstatnum%
               ntmp% = ntmp% - 1
               GoTo t10
            Else
                MsgBox "SunriseSunset(Timer 1) encountered undetermined error: " + CStr(Err.Number) + ", start from the beginning!", vbExclamation, "Cal Program"
                SunriseSunset.OKbut0.Value = False
                SunriseSunset.Cancelbut.Value = True
                Timer1.Enabled = False
                SunriseSunset.Visible = False
                Caldirectories.Visible = True
                Caldirectories.Label1.Enabled = True
                Caldirectories.Drive1.Enabled = True
                Caldirectories.Dir1.Enabled = True
                'Caldirectories.List1.Enabled = True
                Caldirectories.Text1.Enabled = True
                Caldirectories.OKbutton.Enabled = True
                Caldirectories.ExitButton.Enabled = True
                Caldirectories.OKbutton.Enabled = True
                'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                eros = False
                Close
                Exit Sub
                End If

serverhan:
       'response = MsgBox("Inside the Error Handler: serverhan", vbOKOnly + vbExclamation, "Cal Debug")
       If internet = True Then
          lognum% = FreeFile
          Open drivjk$ + "calprog.log" For Append As #lognum%
          Print #lognum%, "While erasing tables and server files, encountered error number: " & Str(Err.Number)
          Print #lognum%, "Aborting program"
          Close #lognum%
          
          For i% = 0 To Forms.Count - 1
            Unload Forms(i%)
          Next i%
      
          myfile = Dir(drivfordtm$ + "busy.cal")
          If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
          
          'kill timer
          If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)

          'end program
          End
      Else
          lognum% = FreeFile
          Open drivjk$ + "calprog.log" For Append As #lognum%
          Print #lognum%, "While erasing tables and server files, encountered error number: " & Str(Err.Number)
          Print #lognum%, "Cal Program skipped this step and went on to next one."
          Close #lognum%
          Resume Next
          End If
             
200
 '  If Check1.Enabled = False Then Check1.Enabled = True
 '  If Check2.Enabled = False Then Check2.Enabled = True
 '  If Combo1.Enabled = False Then Combo1.Enabled = True
 '  If Cancelbut.Enabled = False Then Cancelbut.Enabled = True
 '  If Label3.Enabled = False Then Label3.Enabled = True
 '  If Label3.Visible = False Then
 '     Label3.Caption = sEmpty
 '     Label3.Visible = True
 '     End If
   'Label1.Caption = captmp$
   'Label1.Refresh
 '   If OKbut(0).Enabled = False Then OKbut(0).Enabled = True
End Sub
