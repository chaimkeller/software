VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mapTempfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WorldClim Temperature Model Ver. 2"
   ClientHeight    =   10290
   ClientLeft      =   6570
   ClientTop       =   2430
   ClientWidth     =   6765
   Icon            =   "mapTempfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBarTemp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   9915
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmComparison 
      Caption         =   "Elevations"
      Height          =   1935
      Left            =   240
      TabIndex        =   8
      Top             =   7920
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid flxgrdCompare 
         Height          =   1280
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   2249
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CheckBox chkShowTestDTM 
      Caption         =   "Make Compare DTMs visible"
      Height          =   195
      Left            =   2280
      TabIndex        =   7
      Top             =   80
      Width           =   2415
   End
   Begin VB.Frame frmCompareDTMs 
      Caption         =   "Coordinates and predicted heights, compare DTM, compare files"
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   6600
      Width           =   6255
      Begin VB.CheckBox chkHere 
         Caption         =   "use map loc."
         Height          =   195
         Left            =   2000
         TabIndex        =   18
         ToolTipText     =   "Use last clicked map location"
         Top             =   790
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.TextBox txtElevation 
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
         Left            =   2520
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "Map elevation (trig point or contour line("
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton cmdPaste 
         Height          =   375
         Left            =   5040
         Picture         =   "mapTempfrm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Paste Map coordinates"
         Top             =   300
         Width           =   375
      End
      Begin MSComDlg.CommonDialog comdlgCompare 
         Left            =   5760
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   375
         Left            =   5520
         Picture         =   "mapTempfrm.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "browse for compare file"
         Top             =   300
         Width           =   375
      End
      Begin VB.CheckBox chkRecord 
         Caption         =   "Record to ..\ Compare.txt"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Compare DTMs"
         Height          =   495
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtNorthing 
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
         Left            =   850
         TabIndex        =   6
         Text            =   "-2.0870"
         ToolTipText     =   "ITMy xxx.xxx"
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox txtEasting 
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
         Left            =   850
         TabIndex        =   5
         Text            =   "138.4821"
         ToolTipText     =   "ITMx xxx.xxx"
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label lblElevation 
         Caption         =   "hgt (m)"
         Height          =   375
         Left            =   2000
         TabIndex        =   17
         Top             =   350
         Width           =   615
      End
      Begin VB.Label lblNorthing 
         Caption         =   "Northing"
         Height          =   255
         Left            =   200
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbEasting 
         Caption         =   "Easting"
         Height          =   255
         Left            =   200
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame TKfrm 
      Caption         =   "Termperatures"
      Height          =   5220
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid msFlxGrdTK 
         Height          =   4620
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   8149
         _Version        =   393216
         Rows            =   13
         Cols            =   4
         BackColor       =   -2147483624
         BackColorFixed  =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Loadfrm 
      Caption         =   "Load Temperature Data"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.CommandButton cmdLoadTK 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Load Temps for current coordinates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "mapTempfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileIn$

Private Sub chkShowTestDTM_Click()
   With chkShowTestDTM
      If .value = vbChecked Then
         mapTempfrm.Height = 10740
      ElseIf .value = vbUnchecked Then
         mapTempfrm.Height = 7000
         End If
   End With
End Sub

Private Sub cmdCompare_Click()
   Dim latitude
   Dim longitude
   Dim ITMx
   Dim ITMy
   Dim hgt(3)
   Dim IsraelDTMsource0%
   Dim RMS(3)
   
   ITMx = Val(txtEasting.Text) * 1000
   ITMy = 1000000 + Val(txtNorthing.Text) * 1000
   kmx = ITMx
   kmy = ITMy
   Call heights(kmx, kmy, hgt(0)) 'JKH DTM height
   flxgrdCompare.TextMatrix(1, 1) = hgt(0)
   
   'convert coordinates to geo
   If ggpscorrection = True Then 'apply conversion from Clark geoid to WGS84
       Dim N As Long
       Dim E As Long
       Dim lat As Double
       Dim lon As Double
       N = ITMy
       E = ITMx
       Call ics2wgs84(N, E, lat, lon)
       lgh = lon
       lth = lat
       'Call casgeo(kmx, kmy, lgh, lth)
'       ggpscorrection = False
    Else
       
       Call casgeo(kmx, kmy, lgh, lth)
       End If
       
    IsraelDTMsource0% = IsraelDTMsource%
    
    For IsraelDTMsource% = 1 To 3
        Call worldheights(lgh, lth, hgt(IsraelDTMsource%))
        flxgrdCompare.TextMatrix(1, IsraelDTMsource% + 1) = hgt(IsraelDTMsource%)
    Next IsraelDTMsource%
    
    If chkRecord.value = vbChecked Then
       If Dir(App.Path & "\Compare.txt") = sEmpty Then
          filrec% = FreeFile
          Open App.Path & "\Compare.txt" For Append As #filrec%
          Print #filrec%, "Easting", "Northing", "JKH DTM", "SRTM1", "MERIT", "ALOS", "TRIG POINT"
          If mapTempfrm.chkHere Then
             Print #filrec%, Val(txtEasting), Val(txtNorthing), hgt(0), hgt(1), hgt(2), hgt(3), Val(mapTempfrm.txtElevation)
          Else
             Print #filrec%, Val(txtEasting), Val(txtNorthing), hgt(0), hgt(1), hgt(2), hgt(3), -9999
             End If
          Close #filrec%
       Else
          filrec% = FreeFile
          Open App.Path & "\Compare.txt" For Append As #filrec%
          If mapTempfrm.chkHere Then
             Print #filrec%, Val(txtEasting), Val(txtNorthing), hgt(0), hgt(1), hgt(2), hgt(3), Val(mapTempfrm.txtElevation)
          Else
             Print #filrec%, Val(txtEasting), Val(txtNorthing), hgt(0), hgt(1), hgt(2), hgt(3), -9999
             End If
          Close #filrec%
          End If
       End If
       
    'calculate running rms
    For i% = 0 To 3
       RMS(i%) = 0
    Next i%
    
    filrms% = FreeFile
    Open App.Path & "\Compare.txt" For Input As #filrms%
    Line Input #filrms%, doclin$ 'skip the doc line
    numrms% = 0
    Do Until EOF(filrms%)
        Input #filrms%, Es, Ns, hgt(0), hgt(1), hgt(2), hgt(3), trighgt
        For i% = 0 To 3
            RMS(i%) = RMS(i%) + (trighgt - hgt(i%)) ^ 2
        Next i%
        numrms% = numrms% + 1
    Loop
    Close #filrms%
    
    For i% = 0 To 3
       RMS(i%) = Sqr(RMS(i%) / numrms%)
       mapTempfrm.flxgrdCompare.TextMatrix(2, i% + 1) = Val(Format(RMS(i%), "#0.0"))
    Next i%
    
    StatusBarTemp.Panels(1).Text = Str$(numrms%) & " points"
    
       
    IsraelDTMsource% = IsraelDTMsource0%
    
   'convert ITM to geo coordinates
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdLoadTK_Click
' Author    : Dr-John-K-Hall
' Date      : 11/6/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdLoadTK_Click()
    Dim lt1 As Double, lg1 As Double, i As Integer
    Dim MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer, ier As Integer
    
   On Error GoTo cmdLoadTK_Click_Error

    lg1 = Maps.Text5.Text
    lt1 = Maps.Text6.Text
    
    If Not world Then
       'EY ITM, convert to geo coordinates
       Call casgeo(lg1, lt1, lg, lt)
       lg1 = -lg
       lt1 = lt
    Else
'       tmplt = lt1
'       lt1 = lg1
'       lg1 = -tmplt
       End If
       
    Call Temperatures(lt1, lg1, MinT, AvgT, MaxT, ier)
    
    With msFlxGrdTK
       .ColAlignment(1) = 4
       .ColAlignment(2) = 4
       .ColAlignment(3) = 4
       For i = 1 To 12
         .TextMatrix(i, 1) = MinT(i)
         .TextMatrix(i, 2) = AvgT(i)
         .TextMatrix(i, 3) = MaxT(i)
       Next i
    End With

   On Error GoTo 0
   Exit Sub

cmdLoadTK_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdLoadTK_Click of Form mapTempfrm"
    
End Sub

Private Sub cmdOpen_Click()
   On Error GoTo cmdOpen_Click_Error
   
   Dim hgt(3), ITMx, ITMy, Height, RMS(3)
   Dim N As Long
   Dim E As Long
   Dim lat As Double
   Dim lon As Double

   With comdlgCompare
      .CancelError = True
      .Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
      .ShowOpen
       FileIn$ = .FileName
   End With
   
   filin% = FreeFile
   Open FileIn$ For Input As #filin%
   filout% = FreeFile
   Open App.Path & "\NewCompare.txt" For Output As #filout%
   Print #filout%, "Eastling", "Northing", "Hgt (m)", "JKH DTM", "SRTM1", "MERIT", "ALOS"
   
   numpnt% = 0
   For i% = 0 To 3
      RMS(i%) = 0
   Next i%
   
   Do Until EOF(filin%)
      Input #filin%, ITMx, ITMy, Height
      numpnt% = numpnt% + 1
      
      ITMx = ITMx * 1000
      ITMy = 1000000 + ITMy * 1000
      kmx = ITMx
      kmy = ITMy
      Call heights(kmx, kmy, hgt(0)) 'JKH DTM height
   
       'convert coordinates to geo
       If ggpscorrection = True Then 'apply conversion from Clark geoid to WGS84
           N = ITMy
           E = ITMx
           Call ics2wgs84(N, E, lat, lon)
           lgh = lon
           lth = lat
           'Call casgeo(kmx, kmy, lgh, lth)
    '       ggpscorrection = False
        Else
           
           Call casgeo(kmx, kmy, lgh, lth)
           End If
           
        IsraelDTMsource0% = IsraelDTMsource%
    
        For IsraelDTMsource% = 1 To 3
            Call worldheights(lgh, lth, hgt(IsraelDTMsource%))
        Next IsraelDTMsource%
      
        IsraelDTMsource% = IsraelDTMsource0%
        
        Print #filout%, ITMx * 0.001, (ITMy - 1000000) * 0.001, Height, hgt(0), hgt(1), hgt(2), hgt(3)
        
        For i% = 0 To 3
           RMS(i%) = RMS(i%) + (hgt(i%) - Height) ^ 2
        Next i%
        
    Loop
    
    For i% = 0 To 3
       RMS(i%) = Sqr(RMS(i%) / numpnt%)
    Next i%
    
    Print #filout%, "RMS JKH", "RMS SRTM1", "RMS MERIT", "RMS ALOS"
    Print #filout%, RMS(0), RMS(1), RMS(2), RMS(3)

    Close #filout%
    Close #filin%
   On Error GoTo 0
   Exit Sub

cmdOpen_Click_Error:
   If filin% > 0 Then Close #filin%
   If filout% > 0 Then Close #filout%
End Sub

Private Sub cmdPaste_Click()
  If map400 Or map50 Then
     txtEasting = Val(Maps.Text5) * 0.001
     txtNorthing = (Val(Maps.Text6) - 1000000) * 0.001
     cmdCompare = True
     End If
     
End Sub

Private Sub form_load()
   
   If Not map50 And Not map400 Then
      chkShowTestDTM.Visible = False
      End If
      
   mapTempfrm.Height = 7000
   
   With msFlxGrdTK
      .ColAlignment(0) = 1
      .TextMatrix(1, 0) = "January"
      .TextMatrix(2, 0) = "February"
      .TextMatrix(3, 0) = "March"
      .TextMatrix(4, 0) = "April"
      .TextMatrix(5, 0) = "May"
      .TextMatrix(6, 0) = "June"
      .TextMatrix(7, 0) = "July"
      .TextMatrix(8, 0) = "August"
      .TextMatrix(9, 0) = "September"
      .TextMatrix(10, 0) = "October"
      .TextMatrix(11, 0) = "November"
      .TextMatrix(12, 0) = "December"
      .TextMatrix(0, 1) = "Min. Temp."
      .TextMatrix(0, 2) = "Avg. Temp."
      .TextMatrix(0, 3) = "Max. Temp."
   End With
   
   With flxgrdCompare
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .TextMatrix(1, 0) = "hgt (m)"
      .TextMatrix(2, 0) = "RMS"
      .TextMatrix(0, 1) = "JK DTM"
      .TextMatrix(0, 2) = "SRTM1"
      .TextMatrix(0, 3) = "MERIT"
      .TextMatrix(0, 4) = "ALOS"
      .TextMatrix(1, 1) = sEmpty
   End With
   
   cmdLoadTK_Click
   TempFormVis = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

   TempFormVis = False
   tblbuttons(29) = 0
   Maps.Toolbar1.Buttons(29).value = tbrUnpressed
   Set mapTempfrm = Nothing

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:
   Resume Next
End Sub
