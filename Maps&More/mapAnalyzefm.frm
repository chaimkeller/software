VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form mapAnalyzefm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analysis Wizard Form"
   ClientHeight    =   5670
   ClientLeft      =   7080
   ClientTop       =   1365
   ClientWidth     =   4665
   Icon            =   "mapAnalyzefm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4665
   Begin VB.CheckBox chkAutomatic 
      Caption         =   "&Automatic Reset at Analysis Completion"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      ToolTipText     =   "Check for reset at completion of analyizing any one file"
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   4800
      Width           =   915
   End
   Begin MSComctlLib.StatusBar StatusBarAnalyze 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5295
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1860
      TabIndex        =   9
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   4800
      Width           =   915
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Accept or Edit the Output File's Name"
      Enabled         =   0   'False
      Height          =   975
      Left            =   60
      TabIndex        =   6
      Top             =   3420
      Width           =   4515
      Begin VB.TextBox txtOutput 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frameCities 
      Caption         =   "Accept or Edit the city's directory name"
      Enabled         =   0   'False
      Height          =   1155
      Left            =   60
      TabIndex        =   2
      Top             =   2220
      Width           =   4515
      Begin VB.TextBox txtCities 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   2835
      End
   End
   Begin VB.Frame frameExtensions 
      Caption         =   "Pick an Extension"
      Enabled         =   0   'False
      Height          =   1155
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   4515
      Begin VB.ComboBox cmbExtensions 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Frame frameRoots 
      Caption         =   "Pick a Place for Name for Analysis"
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      Begin VB.ComboBox cmbRoots 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   2475
      End
   End
End
Attribute VB_Name = "mapAnalyzefm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WizNumber%, Extensions$, PrelimOutFile$

Private Sub chkAutomatic_Click()
   If chkAutomatic.value = vbChecked Then
      AutoScanlist = True
   Else
      AutoScanlist = False
      End If
End Sub

Private Sub cmdBack_Click()
   BringWindowToTop (mapAnalyzefm.hWnd)
   WizNumber% = WizNumber% - 1
   Select Case WizNumber%
      Case 0
        cmdBack.Enabled = False
        cmdNext.Enabled = True
        cmbRoots.Enabled = True
        frameRoots.Enabled = True
        frameExtensions.Enabled = False
        cmbExtensions.Enabled = False
        StatusBarAnalyze.Panels(1).Text = "Choose a place for analysis"
      Case 1
        cmbExtensions.Enabled = True
        frameExtensions.Enabled = True
        frameCities.Enabled = False
        txtCities.Enabled = False
        StatusBarAnalyze.Panels(1).Text = "Choose a file extensions(s) to analyze"
     Case 2
        frameOutput.Enabled = False
        txtOutput.Enabled = False
        frameCities.Enabled = True
        txtCities.Enabled = True
         StatusBarAnalyze.Panels(1).Text = "Enter the city's name"
      Case 3
        frameOutput.Enabled = True
        txtOutput.Enabled = True
        StatusBarAnalyze.Panels(1).Text = "Enter the name of the profile file"
        cmdNext.Enabled = True
   End Select
End Sub

Private Sub cmdCancel_Click()
  Call form_queryunload(0, 0)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdNext_Click
' DateTime  : 9/30/2003 07:57
' Author    : Chaim Keller
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdNext_Click()
   On Error GoTo cmdNext_Click_Error

   BringWindowToTop (mapAnalyzefm.hWnd)
   WizNumber% = WizNumber% + 1
cmdn5:
   Select Case WizNumber%
      Case 1
        cmdBack.Enabled = True
        PrelimOutFile$ = cmbRoots.Text
        frameRoots.Enabled = False
        cmbRoots.Enabled = False
        frameExtensions.Enabled = True
        cmbExtensions.Enabled = True
        StatusBarAnalyze.Panels(1).Text = "Choose a file extensions(s) to analyze"
        'populate extension combo box
        FileVT% = cmbRoots.ListIndex
        If FileVT% = -1 Then
           response = MsgBox("The name you requested is not listed!" & vbLf & _
                  "Do you wish to use the inputed name?", _
                  vbYesNoCancel + vbExclamation, "Maps & More")
           If response = vbYes Then
              FileVT% = 0
           ElseIf response = vbNo Then
              GoTo cmdn5
           Else
              Exit Sub
              End If
           End If
        Select Case FileViewFileType(FileVT%)
           Case -4
              cmbExtensions.AddItem "004"
           Case -3
              cmbExtensions.AddItem "004"
              cmbExtensions.AddItem "005"
              cmbExtensions.AddItem "004+005"
           Case -2
              cmbExtensions.AddItem "004"
              cmbExtensions.AddItem "005"
              cmbExtensions.AddItem "006"
              cmbExtensions.AddItem "004+005"
              cmbExtensions.AddItem "004+005+006"
           Case 1
              cmbExtensions.AddItem "001"
           Case 2
              cmbExtensions.AddItem "001"
              cmbExtensions.AddItem "002"
              cmbExtensions.AddItem "001+002"
           Case 3
              cmbExtensions.AddItem "001"
              cmbExtensions.AddItem "002"
              cmbExtensions.AddItem "003"
              cmbExtensions.AddItem "001+002"
              cmbExtensions.AddItem "002+003"
              cmbExtensions.AddItem "001+002+003"
           Case 4
              cmbExtensions.AddItem "001"
              cmbExtensions.AddItem "002"
              cmbExtensions.AddItem "003"
              cmbExtensions.AddItem "001+002"
              cmbExtensions.AddItem "002+003"
              cmbExtensions.AddItem "001+002+003"
              cmbExtensions.AddItem "004"
              cmbExtensions.AddItem "001+002+003 | 004"
           Case 5
              cmbExtensions.AddItem "001"
              cmbExtensions.AddItem "002"
              cmbExtensions.AddItem "003"
              cmbExtensions.AddItem "001+002"
              cmbExtensions.AddItem "002+003"
              cmbExtensions.AddItem "001+002+003"
              cmbExtensions.AddItem "004"
              cmbExtensions.AddItem "005"
              cmbExtensions.AddItem "004+005"
              cmbExtensions.AddItem "001+002+003 | 004+005"
           Case 6
              cmbExtensions.AddItem "001"
              cmbExtensions.AddItem "002"
              cmbExtensions.AddItem "003"
              cmbExtensions.AddItem "001+002"
              cmbExtensions.AddItem "002+003"
              cmbExtensions.AddItem "001+002+003"
              cmbExtensions.AddItem "004"
              cmbExtensions.AddItem "005"
              cmbExtensions.AddItem "006"
              cmbExtensions.AddItem "004+005"
              cmbExtensions.AddItem "004+005+006"
              cmbExtensions.AddItem "001+002+003 | 004+005+006"
           Case Else
        End Select
        cmbExtensions.ListIndex = cmbExtensions.ListCount - 1 'place on last item
      Case 2
         Extensions$ = cmbExtensions.List(cmbExtensions.ListIndex)
         If Trim$(cmbExtensions.Text) <> cmbExtensions.List(cmbExtensions.ListIndex) And cmbExtensions.Text <> sEmpty Then
            Extensions$ = Trim$(cmbExtensions.Text)
            End If
         cmbExtensions.Enabled = False
         frameExtensions.Enabled = False
         frameCities.Enabled = True
         txtCities.Enabled = True
         If InStr(FileViewDir$, drivcities$) <> 0 Then
            txtCities.Text = Mid$(FileViewDir$, Len(drivcities$) + 1, Len(FileViewDir$) - Len(drivcities$) + 1)
         Else
            txtCities.Text = FileViewDir$
            End If
         StatusBarAnalyze.Panels(1).Text = "Enter the city's name"
      Case 3
         If txtCities.Text <> FileViewDir$ Then
            FileViewDir$ = txtCities.Text
            End If
         frameCities.Enabled = False
         txtCities.Enabled = False
         txtOutput.Enabled = True
         frameOutput.Enabled = True
         txtOutput.Text = Mid$(PrelimOutFile$, Len(drivprom$) + 1, Len(PrelimOutFile$) - Len(drivprom$) + 1)
         StatusBarAnalyze.Panels(1).Text = "Enter the name of the profile file"
         If AutoScanlist Then
            StatusBarAnalyze.Panels(1).Text = "Press the next button to obtain a graph"
            End If
      Case 4
         StatusBarAnalyze.Panels(1).Text = "Press ""Back"" to analyze more files"
         cmdNext.Enabled = False
         frameOutput.Enabled = False
         txtOutput.Enabled = False
         OutFile$ = txtOutput.Text
         'FileViewDir$ = Mid$(FileViewDir$, Len(drivcities$) + 1, Len(FileViewDir$) - Len(drivcities$) + 1)
         Analyze
         
         'analyze the place and plot it
         If chkAutomatic.value = vbChecked Then
            'automatic reset and make next file appear
            If cmbRoots.ListIndex < UniqueRoots% Then
               cmbRoots.ListIndex = cmbRoots.ListIndex + 1
               cmdBack_Click
               cmdBack_Click
               cmdBack_Click
               cmdBack_Click
               cmdNext_Click
               cmdNext_Click
               cmdNext_Click
            Else
               'finished
               AutoProf = False
               AutoScanlist = False
               chkAutomatic.value = vbUnchecked
               Exit Sub
               End If
            End If
            
            Do Until graphwind = False
               DoEvents
            Loop
            waitime = Timer
a500:       If Timer > waitime + 5 Then
               'push the button automatically after a minute
               AutoProf = True
               cmdNext_Click
            Else
               If Int(5 - Timer + waitime) <> IntOld% Then
                  StatusBarAnalyze.Panels(1).Text = "Auto Mode..." & Str$(Int(5 - Timer + waitime)) & " sec left."
                  StatusBarAnalyze.Refresh
                  IntOld% = Int(5 - Timer + waitime)
                  End If
                DoEvents
                GoTo a500
                End If
            
   End Select

   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:
    If Err.Number = 364 Then
        Err.Clear
        Exit Sub
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdNext_Click of Form mapAnalyzefm"
        End If
End Sub

Private Sub form_load()
   On Error GoTo form_load_Error

   If UniqueRoots% > 1 Then
      chkAutomatic.value = vbChecked
      'AutoScanlist = True
      End If
   WizNumber% = 0
   'display list of uniuqe root names
   For i% = 0 To UniqueRoots%
      If Trim$(FileViewFileName(0, i%)) <> sEmpty Then
         cmbRoots.AddItem FileViewFileName(0, i%)
      Else 'don't count this blank
         UniqueRoots% = UniqueRoots% - 1
         End If
   Next i%
   cmbRoots.ListIndex = 0 'show first place
   StatusBarAnalyze.Panels(1).Text = "Choose a place for analysis"

   On Error GoTo 0
   Exit Sub

form_load_Error:
    
    Select Case Err.Number
       Case 380 'trying to exit
           Call form_queryunload(0, 0)
       Case Else
          MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure form_load of Form mapAnalyzefm"
    End Select
End Sub

Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set mapAnalyzefm = Nothing
   AutoProf = False
   AutoScanlist = False
   'zero dynamic arrays
   UniqueRoots% = 0
   ReDim FileViewFileName(7, UniqueRoots%)
   ReDim FileViewFileType(UniqueRoots%)
   
End Sub
Private Sub Analyze()
'Windows edition: 1-17-2002
'PROGRAM ANALYZE finds earlist/latest netz/skiya of .001-.003/.004-.006 files
'and then determines refraction of mountains and adds that refraction to the
'view angles (using the refraction tables HGTREFS.SUM/WIN), and then finds if
'there are holes in the output (i.e., where ref=-9.999 = menat4 didn't converge)
'and interpolates between those points, and finally writes an output file
' ".fnz" for netz (sunrise), ".fsk" for skiya (sunset).
'NOTE: The .001/.004 file can have a larger angular range than the other files
'but the other files must have the same anglular range

'FIRST COMBINE OUTPUT FILES FROM RDHALBAT
'AND WRITE OUTPUT TO d:\prof

'   mode%(setflag%) =
'                   0 = merges .001-.003 for netz/.004-.006 for skiya
'                   1 = don't merge, rather only use file ext
'                   2 = skip this entry altogether
'                   3 = merge only .004-.005 for skiya
'                   4 = use named file for the .002,.003 file
'                   5 = merge only .001-.002 for netz
'                   6 = merge only .002-.003 for netz
'
'   modref%(setflag%) = 0, compares rdhalbat output without adding
'                          refraction, determines maximum view angle
'                          and only then adds the refraction
'                     = 1, first adds refraction to each individual file
'                          and then determines maximum view angle
'                          will not keep track if tables not complete
'                     = 2, same as 1, but uses the analytic formulas for
'                          refraction to determine the total refraction
'                     = -2, uses analytic formulas for refraction, but
'                           does not first add refraction to individual files
'
'   ext$(1,1) = ".001" 'must give extension for mode%=1
'                ext$(setflag%)

Dim Mode%(1), ext$(1), filout As String
'Form$ = "+##.#     +##.####         #       #"
'form1$ = "+##.#   +##.####   ###.####   ###.####   ###.###  +####.#"
'form2$ = "###.###   ####.###   ####.#   ###.###   ####.###   ###.   ###.   #.#"
'form3$ = "+##.#   +##.####   ##.####   ##.####"
'form5$ = "&   ###.###   ####.###   ####.#   ####   ####   #"
'--------------------------user input------------------------
'named$ = "1266c-te" <<<this option not supported

'OutFile stripped of the drivcities$ name
On Error GoTo errhand

'---------rdhalba3 already took care of refraction-------------
'If cmbExtensions.List(cmbExtensions.ListIndex) = "001" Then
If cmbExtensions.Text = "001" Then
   netzskiy$ = "\netz\"
   sunmode% = 1
Else
   netzskiy$ = "\skiy\"
   sunmode% = 0
   End If
'open file and read coordinates

tmpfil% = FreeFile
'Open drivcities$ & txtCities.Text & netzskiy$ & _
'        cmbRoots.List(cmbRoots.ListIndex) For Input As #tmpfil%
Open drivcities$ & txtCities.Text & netzskiy$ & _
         cmbRoots.Text For Input As #tmpfil%
Line Input #tmpfil%, doclin$
Input #tmpfil%, kmxoa, kmyoa, hgta, begkmx, endkmx, dX, dY, apprn
coordAnalyze(0) = kmxoa
coordAnalyze(1) = kmyoa
coordAnalyze(2) = hgta
Close #tmpfil%

If mapPictureform.Visible = True Then
   'move to map to that point
   Maps.Text5.Text = kmxoa
   Maps.Text6.Text = kmyoa
   Maps.Text7.Text = hgta
   Call goto_click
   End If

'FileCopy drivcities$ & txtCities.Text & netzskiy$ & _
'        cmbRoots.List(cmbRoots.ListIndex), drivjk$ + "EYisroel.tmp"
FileCopy drivcities$ & txtCities.Text & netzskiy$ & _
         cmbRoots.Text, drivjk_c$ + "EYisroel.tmp"
'GoTo a800
'--------------------------------------------------------------
'old file methods (no longer used)
'
'infile$ = PrelimOutFile$
'outfil$ = OutFile$
'
''determine mode%
''first sunrise
'mode%(0) = 2
'If InStr(Extensions$, "001") <> 0 Then
'   mode%(0) = 1
'   ext$(0) = ".001"
'   If InStr(Extensions$, "002") <> 0 Then
'      mode%(0) = 3
'      If InStr(Extensions$, "003") <> 0 Then
'         mode%(0) = 0
'      Else
'         mode%(0) = 5
'      End If
'   End If
'ElseIf InStr(Extensions$, "002") <> 0 Then
'   mode%(0) = 1
'   ext$(0) = ".002"
'   If InStr(Extensions$, "003") <> 0 Then
'      mode%(0) = 6
'   End If
'ElseIf InStr(Extensions$, "003") <> 0 Then
'   mode%(0) = 1
'   ext$(0) = ".003"
'End If
'
''now sunset
'mode%(1) = 2
'If InStr(Extensions$, "004") <> 0 Then
'   mode%(1) = 1
'   ext$(1) = ".004"
'   If InStr(Extensions$, "005") <> 0 Then
'      mode%(1) = 3
'      If InStr(Extensions$, "006") <> 0 Then
'         mode%(1) = 0
'      End If
'   End If
'ElseIf InStr(Extensions$, "005") <> 0 And InStr(Extensions$, "006") = 0 Then
'   mode%(1) = 1
'   ext$(1) = ".005"
'ElseIf InStr(Extensions$, "006") <> 0 And InStr(Extensions$, "005") = 0 Then
'   mode%(1) = 1
'   ext$(1) = ".006"
'End If
'
'
'For setflag% = 0 To 1
'    Screen.MousePointer = vbHourglass
'
'    If mode%(setflag%) = 2 Then GoTo 850 'skip this entry
'
'    If mode%(setflag%) = 1 Then 'don't merge
'       fil1% = FreeFile
'       f1$ = infile$ + ext$(setflag%)
'       Open f1$ For Input As #fil1%
'       Line Input #fil1%, doclin$
'       Input #fil1%, kmxoa, kmyoa, hgta, begkmx, endkmx, dx, dy, apprn
'       coordAnalyze(0) = kmxoa
'       coordAnalyze(1) = kmyoa
'       coordAnalyze(2) = hgta
'       nentry% = 0
'       Do Until EOF(fil1%)
'          Input #fil1%, X1, Y1, A1, B1, C1, D1
'          nentry% = nentry% + 1
'          If nentry% = 1 Then begang = X1
'       Loop
'       Close #fil1%
'       If setflag% = 0 Then
'          f4$ = drivprof$ + outfil$ + ".fnz"
'       ElseIf setflag% = 1 Then
'          f4$ = drivprof$ + outfil$ + ".fsk"
'          End If
'       FileCopy f1$, f4$
'       GoTo 50
'       End If
'
'    If setflag% = 0 Then
'       f1$ = infile$ + ".001"
'       If mode%(0) = 4 Then 'not supported as of now
'         f2$ = drivprom$ + named$ + ".002"
'         f3$ = drivprom$ + named$ + ".003"
'       Else
'          f2$ = infile$ + ".002"
'          f3$ = infile$ + ".003"
'          If mode%(0) = 5 Then f3$ = infile$ + ".002"
'          If mode%(0) = 6 Then f1$ = infile$ + ".002"
'          End If
'       f4$ = drivprof$ + outfil$ + ".fnz"
'    ElseIf setflag% = 1 Then
'       f1$ = infile$ + ".004"
'       f2$ = infile$ + ".005"
'       f3$ = infile$ + ".006"
'       If mode%(1) = 3 Then f3$ = infile$ + ".005"
'       f4$ = drivprof$ + outfil$ + ".fsk"
'       End If
'
'50:
'   fil1% = FreeFile
'   Open f1$ For Input As #fil1%
'   If mode%(setflag%) <> 1 Then
'     fil2% = FreeFile
'     Open f2$ For Input As #fil2%
'     fil3% = FreeFile
'     Open f3$ For Input As #fil3%
'     End If
'   fil4% = FreeFile
'   Open f4$ For Output As #fil4%
'   filout = drivjk$ + outfil$ + ".tmp"
'   fil5% = FreeFile
'   Open filout For Output As #fil5%
'
'   Line Input #fil1%, doclin$
'   Print #fil4%, doclin$
'   If mode%(setflag%) <> 1 Then
'      Line Input #fil2%, doclin$
'      Line Input #fil3%, doclin$
'      End If
'
'   Input #fil1%, kmxoa, kmyoa, hgta, begkmx, endkmx, dx, dy, apprn
'   coordAnalyze(0) = kmxoa
'   coordAnalyze(1) = kmyoa
'   coordAnalyze(2) = hgta
'   If setflag% = 0 Then
'      minkmx = begkmx
'   ElseIf setflag% = 1 Then
'      maxkmx = endkmx
'   End If
'
'   If mode%(setflag%) <> 1 Then
'     Input #fil2%, kmxoa, kmyoa, hgta, begkmx, endkmx, dx, dy, apprn
'
'     Input #fil3%, kmxoa, kmyoa, hgta, begkmx, endkmx, dx, dy, apprn
'     End If
'
'   If setflag% = 0 Then
'     maxkmx = endkmx
'   ElseIf setflag% = 1 Then
'     minkmx = begkmx
'     End If
'
'   Write #fil5%, "FILENAME, KMX, KMY, HGT: ", f4$, kmxoa, kmyoa, hgta
'   Print #fil5%, "  AZI  VIEWANG+REFRACT   FLGSUM   FLGWIN"
'
'   Write #fil4%, kmxoa, kmyoa, hgta, minkmx, maxkmx, dx, dy, apprn
'   nentry% = 0
'   Start% = 0
'   Do Until EOF(fil1%)
'      nentry% = nentry% + 1
'      Input #fil1%, X1, Y1, A1, B1, C1, D1
'
'      va = Y1: dispk = C1: hgt2 = D1
'      deltd = hgta - hgt2
'      distd = dispk
'      Y1 = va + AVREF(deltd, distd)
'
'      ymax = Y1
'      amax = A1
'      bmax = B1
'      cmax = C1
'      DMax = D1
'
'      If mode%(setflag%) = 1 Then GoTo a500
'
'      If nentry% = 1 Then
'         Input #fil2%, X2, Y2, A2, B2, C2, D2
'         va = Y2: dispk = C2: hgt2 = D2
'         deltd = hgta - hgt2
'         distd = dispk
'         Y2 = va + AVREF(deltd, distd)
'
'         Input #fil3%, x3, y3, A3, B3, c3, d3
'         va = y3: dispk = c3: hgt2 = d3
'         deltd = hgta - hgt2
'         distd = dispk
'         y3 = va + AVREF(deltd, distd)
'
'         begang = X1
'         begang2 = X2
'         endang2 = -X2
'         Start% = 1
'         End If
'
'      If X1 >= begang2 And X1 <= endang2 Then
'         If Start% <> 1 Then
'            Input #fil2%, X2, Y2, A2, B2, C2, D2
'            va = Y2: dispk = C2: hgt2 = D2
'            deltd = hgta - hgt2
'            distd = dispk
'            Y2 = va + AVREF(deltd, distd)
'
'            Input #fil3%, x3, y3, A3, B3, c3, d3
'            va = y3: dispk = c3: hgt2 = d3
'            deltd = hgta - hgt2
'            distd = dispk
'            y3 = va + AVREF(deltd, distd)
'            End If
'
'         Start% = 0
'
'         If Y2 > ymax Then
'            ymax = Y2
'            amax = A2
'            bmax = B2
'            cmax = C2
'            DMax = D2
'            End If
'         If y3 > ymax Then
'            ymax = y3
'            amax = A3
'            bmax = B3
'            cmax = c3
'            DMax = d3
'            End If
'         End If
'
'a500:
'      'Print #fil5%, USING; Form$; X1; ymax; nflgsum%; nflgwin%
'      'Try both
'      'Print #fil5%, Format(Str(X1), "##0.0"); Spc(5); Format(Str(ymax), "#0.0000"); Spc(5); Format(Str(0#), "0.0000"); Spc(5); Format(Str(0#), "0.0000")
'      Print #fil5%, Format(Str(X1), "##0.0"); Tab(10); Format(Str(ymax), "#0.0000"); Tab(21); Format(Str(0#), "0.0000"); Tab(31); Format(Str(0#), "0.0000")
'
'      'Print #fil4%, Format(Str(X1), "###0.0"); Spc(5); Format(Str(ymax), "#0.0000"); Spc(5); Format(Str(amax), "###.000"); Spc(5); Format(Str(bmax), "###.000"); Spc(5); Format(Str(cmax), "##0.000"); Spc(5); Format(Str(dmax), "####0.0")
'      Print #fil4%, Format(Str(X1), "##0.0"); Tab(10); Format(Str(ymax), "#0.0000"); Tab(20); Format(Str(amax), "###.000"); Tab(31); Format(Str(bmax), "###.000"); Tab(43); Format(Str(cmax), "##0.000"); Tab(52); Format(Str(DMax), "####0.0")
'   Loop
'
'   Close #fil1%
'   If mode%(setflag%) <> 1 Then
'      Close #fil2%
'      Close #fil3%
'      End If
'   Close #fil4%
'   Close #fil5%
'
'   'write plotting file
'   FileCopy f4$, drivjk$ + "EYisroel.tmp"
'
'    If setflag% = 0 Then
'       fileo$ = drivfordtm$ + "netz\" + outfil$ + ".pr0"
'    ElseIf setflag% = 1 Then
'       fileo$ = drivfordtm$ + "skiy\" + outfil$ + ".pr0"
'       End If
'    FileCopy filout, fileo$
'    On Error Resume Next
'    Kill filout
    
a800:
     Screen.MousePointer = vbDefault
    'now plot the temporary plot file
     mapgraphfm.Visible = True
     ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
     
     cal1% = 0
     If IntOld2% = 0 Then
       IntOld2% = 20 'start with 20 sec countdown
       mapgraphfm.txtDelay = 20
       End If
       waitime = Timer
     
     Do While killpicture = False
        DoEvents
         
        If AutoScanlist = True Then 'automatic operation
        '**************************
a250:       If Timer > waitime + 2 And cal1% = 0 Then
               'push the obstruction button
               cal1% = 1
               mapgraphfm.Command3.value = True
            ElseIf Timer > waitime + 8 And cal1% = 1 Then
               'push the restore limits button
               cal1% = 2
               mapgraphfm.restorelimitsbut.value = True
            ElseIf Timer > waitime + IntOld2% And cal1% = 2 Then
               'push the button automatically after a IntOld2% seconds
               cal1% = 3
               mapgraphfm.Calendarbut.value = True
               'reset to push this button again after IntOld2% seconds
               cal1% = 2
               waitime = Timer
            Else
               If Int(IntOld2% - Timer + waitime) <> IntOld% Then
                  mapgraphfm.StatusBar1.Panels(1).Text = "Auto Mode...Calc. starting in" & Str$(Int(IntOld2% - Timer + waitime)) & " sec."
                  mapgraphfm.StatusBar1.Refresh
                  IntOld% = Int(IntOld2% - Timer + waitime)
                  IntOld2tmp% = Val(mapgraphfm.txtDelay)
                  If IntOld2tmp% < 10 Then
                     mapgraphfm.txtDelay = 10
                     IntOld2tmp% = 10
                     End If
                  IntOld2% = IntOld2tmp%
                  End If
                DoEvents
                GoTo a250
                End If
                       
                       
        '***********************
               End If
         
        
     Loop
     killpicture = False
     
'850:
'Next setflag%

Screen.MousePointer = vbDefault
Exit Sub

errhand:
   Screen.MousePointer = vbDefault
   Close
   MsgBox "Error Number: " & Str(Err.Number) & " encountered." + Chr(10) & _
          "Error Description: " & Err.Description + Chr(10) & _
          "Aborting!", vbCritical + vbOKOnly, "Maps & More"
End Sub
