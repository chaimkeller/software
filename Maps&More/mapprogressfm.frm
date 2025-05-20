VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mapprogressfm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1605
   ClientLeft      =   2820
   ClientTop       =   5760
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "mapprogressfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1200
      TabIndex        =   15
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame frmDTM 
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
      Height          =   980
      Left            =   4860
      TabIndex        =   11
      Top             =   250
      Visible         =   0   'False
      Width           =   1095
      Begin VB.OptionButton optALOS 
         Caption         =   "AlOS (30m)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   720
         Width           =   900
      End
      Begin VB.OptionButton optSRTM2 
         Caption         =   "SRTM-90 m"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   350
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.OptionButton optSRTM1 
         Caption         =   "SRTM-30 m"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   560
         Width           =   1000
      End
      Begin VB.OptionButton optGTOPO30 
         Caption         =   "1 km DTM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   140
         Width           =   915
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   4140
      Picture         =   "mapprogressfm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "view in 3D"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5790
      Picture         =   "mapprogressfm.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   220
   End
   Begin VB.CommandButton Acceptbut 
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
      Left            =   3540
      Picture         =   "mapprogressfm.frx":0B02
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Calculate 2D Profile using inputed nearest approach (km)"
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   -240
      ScaleHeight     =   255
      ScaleWidth      =   6015
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nearest Approach/3D Viewer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   0
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   200
         Left            =   0
         Picture         =   "mapprogressfm.frx":0E0C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   250
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   555
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   979
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text1"
      BuddyDispid     =   196621
      OrigLeft        =   3720
      OrigTop         =   240
      OrigRight       =   3960
      OrigBottom      =   735
      Max             =   200
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1200
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   10000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1230
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "Extracting the relavant portion of the DTM"
            TextSave        =   "Extracting the relavant portion of the DTM"
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
   Begin VB.Label Label3 
      Caption         =   "km"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "mapprogressfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acceptbut_Click()
   On Error GoTo errhand
   
   accept = True
   apprn = Val(mapprogressfm.Text2.Text)
   If RepairMode Then apprn = 0.5
   'create eros.tm6 file
   dtmfile% = FreeFile
   Open ramdrive & ":\eros.tm6" For Output As #dtmfile%
   Select Case DTMflag
      Case 0, -1 'GTOPO30
         outdrive$ = worlddtm
      Case 1, 2 'SRTM
         outdrive$ = srtmdtm
      Case 3 'ALOS
         outdrive$ = alosdtm
   End Select
   Print #dtmfile%, outdrive$; ","; DTMflag
   Close #dtmfile%
   
   Call Form_QueryUnload(i%, j%)
   Exit Sub
   
errhand:
   Call MsgBox("Encountered error number: " & Str$(Err.Number) & vbLf & _
               Err.Description, vbCritical, "DTM file error")
   
End Sub

Private Sub Command1_Click()
   Dim myfile
   If Picture1.Visible = True Then
      accept = False
   Else 'aborted reading the DTM, erase the DTM files
      myfile = Dir(ramdrive + ":\*.bin")
      If myfile <> sEmpty Then
         Close
         Kill (ramdrive + ":\" + myfile)
         End If
      myfile = Dir(ramdrive + ":\*.bi1")
      If myfile <> sEmpty Then
         Close
         Kill (ramdrive + ":\" + myfile)
         End If
      myfile = Dir(drivjk_c$ + "eros.tm3")
      If myfile <> sEmpty Then Kill (drivjk_c$ + "eros.tm3")
      mapprogressfm.Visible = False
      abortDTM = True
      End If
   Call Form_QueryUnload(i%, j%)
End Sub

Private Sub Command2_Click()
   On Error GoTo errhand
   
   viewer3D = True
   Call Form_QueryUnload(i%, j%)
   
   'create eros.tm6 file
   dtmfile% = FreeFile
   Open ramdrive & ":\eros.tm6" For Output As #dtmfile%
   Select Case DTMflag
      Case 0 'GTOPO30
         outdrive$ = worlddtm
      Case 1, 2 'SRTM
         outdrive$ = srtmdtm
      Case 3 'ALOS
         outdrive$ = alosdtm
   End Select
   Print #dtmfile%, outdrive$; ","; DTMflag
   Close #dtmfile%
   
   'ret = Shell("c:\samples\vc98\sdk\graphics\directx\egg\debug\egg.exe", vbNormalFocus)
   Exit Sub
   
errhand:
   Call MsgBox("Encountered error number: " & Str$(Err.Number) & vbLf & _
               Err.Description, vbCritical, "DTM file error")

End Sub

Private Sub form_load()
   'read last eros.tm6 file to determine last DTM used
'   If Dir(ramdrive & ":\eros.tm6") <> sEmpty Then
'        dtmfile% = FreeFile
'        Open ramdrive & ":\eros.tm6" For Input As #dtmfile%
'        Input #dtmfile%, outdrive$, DTMflag
        Select Case DTMflag
           Case 0, -1 'GTOPO30 / SRTM30 (1 km)
              optGTOPO30.value = True
           Case 1 'SRTM-1 (30 meter)
              optSRTM1.value = True
           Case 2 'SRTM-3 / MERIT (90 meter)
              optSRTM2.value = True
           Case 3 'ALOS (30 meters)
              optALOS.value = True
        End Select
        
'        Close #dtmfile%
'   Else
'       DTMflag = 0 'default is SRTM30
'       optGTOPO30.value = True
'       End If

   
   'DTMflag = 2 'default is SRTM layer 2 DEM (90 meter)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If viewer3D = False Then BringWindowToTop (mapPictureform.hwnd) 'ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
   Unload Me
End Sub

Private Sub optALOS_Click()
   DTMflag = 3
End Sub

Private Sub optGTOPO30_Click()
   DTMflag = 0
End Sub

Private Sub optSRTM1_Click()
   DTMflag = 1
End Sub

Private Sub optSRTM2_Click()
   DTMflag = 2
End Sub

Private Sub Text2_Change()
   If UpDown1.value <> 2 * Val(Text2.Text) Then UpDown1.value = 2 * Val(Text2.Text)
End Sub

Private Sub UpDown1_Change()
   Text2.Text = 0.5 * Val(Text1.Text)
End Sub
