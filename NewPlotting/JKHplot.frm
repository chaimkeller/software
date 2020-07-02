VERSION 5.00
Begin VB.Form JKHplot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversion Program"
   ClientHeight    =   3195
   ClientLeft      =   3255
   ClientTop       =   3255
   ClientWidth     =   8550
   Icon            =   "JKHplot.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8550
   Begin VB.OptionButton optSVC 
      Caption         =   "Dump calculated sound velocity (SVC) data to output ""rel"" file"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1860
      TabIndex        =   3
      Top             =   2400
      Width           =   4815
   End
   Begin VB.OptionButton optSV 
      Caption         =   "Dump sound Velocity (SV) data to output ""rel"" file"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1860
      TabIndex        =   2
      Top             =   2160
      Width           =   4575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   8415
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      Height          =   555
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   2100
      Width           =   5175
   End
End
Attribute VB_Name = "JKHplot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SVcheck%, dumpmode%
Private Sub cmdConvert_Click()
    'this program converts csv Multi-beam raw data to plot ready
    'data. It takes the place of CTD717.for
    'there are two modes determined by the flag dumpmode%
    'dumpmode% = 0 'plot mode for evaluation
    '          = 1 'dump either SV or SVC to rel file
    'the evaluation mode produces three files:
    '(1)depth,Temperature = nameTemp.tmp
    '(2)depth,SV = nameSV.tmp
    '(3)depth,SVC = nameSVC.tmp
    'in the output mode, the user decides whether he wants
    'to produce either:
    '(1)depth, Temperature, SV = name.rel
    '(2)depth, Temperature, SVC = name.rel
    
    On Error GoTo errhand
    
    Dim FileIn As String, FileOut As String
    Dim FileOut1 As String, FileOut2 As String, FileOut3 As String
    Dim P() As Single, T() As Single, C() As Single
    Dim SV() As Single, B() As Single, s() As Single
    Dim d() As Single, SVC() As Single
    
    Dim yr As Integer, mo As Integer, dy As Integer, Ext$
    Dim hr As Integer, Min As Integer, Sec As Integer
    Dim JDD As Long, PD As Single, TD As Single, SD As Single
    Dim SVCC As Single, DD As Single, SVD As Single, ALPHA As Single

    Screen.MousePointer = vbHourglass
    
    FileIn = PlotInfofrm.lblFileName.Caption
    If dumpmode% = 0 Then
       Ext$ = "tmp"
    Else
       Ext$ = "rel"
       End If
    If dumpmode% = 0 Then
       FileOut1 = Mid$(FileIn, 1, Len(FileIn) - 4) & "Tmp." & Ext$
       FileOut2 = Mid$(FileIn, 1, Len(FileIn) - 4) & "SVM." & Ext$
       FileOut3 = Mid$(FileIn, 1, Len(FileIn) - 4) & "SVC." & Ext$
    Else
       If InStr(FileIn, "Tmp") <> 0 Or InStr(FileIn, "SVC") <> 0 Or InStr(FileIn, "SVM") <> 0 Then
          FileIn = Mid$(FileIn, 1, Len(FileIn) - 7) & ".csv"
          FileOut = Mid$(FileIn, 1, Len(FileIn) - 3) & Ext$
       Else
          FileOut = Mid$(FileIn, 1, Len(FileIn) - 3) & Ext$
          End If
       End If
    DMAX = 0#
    TX = 0#
    TN = 100#
    SVX = 0#
    SVN = 2000#
    ABSAV = 0#
    DIFFAV = 0#
    filin% = FreeFile
    Open FileIn For Input As #filin%
    If dumpmode% = 1 Then
       filout% = FreeFile
       Open FileOut For Output As #filout%
       End If
    For I% = 1 To 4
       Line Input #filin%, doclin$ 'skip document lines
    Next I%
    'start reading in data
    numpnts% = 0
    WroteHeader% = 0
    List1.Clear
    Do Until EOF(filin%)
       Input #filin%, DateTimeIn$, PD, TD, CD, SVD, BD, SD, DD
       If numpnts% = 0 And WroteHeader% = 0 And dumpmode% = 1 Then
          yr = Year(DateTimeIn$)
          mo = Month(DateTimeIn$)
          dy = Day(DateTimeIn$)
          hr = Hour(DateTimeIn$)
          Min = Minute(DateTimeIn$)
          Sec = Second(DateTimeIn$)
          
          'convert to Julian Date
          Call JULDAY(dy, mo, yr, JDD)
          'determine nearest time
          timeshift% = 0
          If InStr(DateTimeIn$, "PM") Then timeshift% = 12
          If Sec > 30 Then Min = Min + 1
          If Min > 60 Then
             Min = 0
             hr = hr + 1
             End If
          If hr > 24 Then
             hr = 0
             JDD = JDD + 1
             End If
         'formulate time string
          DateString$ = Trim$(Str(mo)) & "-" + Trim$(Str(dy)) & "-" & Trim$(Str(yr - 2000))
          'formulate time string
          hrString$ = Trim$(Str(hr))
          If hr < 10 Then hrString$ = "0" & Trim$(Str(hr))
          MinString$ = Trim$(Str(Min))
          If Min < 10 Then MinString$ = "0" & Trim$(Str(Min))
          TimeString$ = hrString$ & MinString$
          
          'write output file header
          Print #filout%, "CALC,717," & DateString$ & ",-1,meters"
          
          Print #filout%, "AML SOUND VELOCITY PROFILER S/N:717"
          Print #filout%, "DATE:" & Trim$(Str(JDD)) & " TIME:" & TimeString$
          Print #filout%, "DEPTH OFFSET (M): 0.0"
          Print #filout%, "DEPTH (M) VELOCITY (M/S) TEMP (C)"
          WroteHeader% = 1
        End If
'C     Decide if data is valuable
       If (SVD < 1490# Or SVD > 1600#) Then GoTo 100
       If (SD < 30# Or SD > 45#) Then GoTo 100
       If (TD < 10# Or TD > 45#) Then GoTo 100
       
       ReDim Preserve P(numpnts%)
       ReDim Preserve T(numpnts%)
       ReDim Preserve C(numpnts%)
       ReDim Preserve SV(numpnts%)
       ReDim Preserve B(numpnts%)
       ReDim Preserve s(numpnts%)
       ReDim Preserve d(numpnts%)
       ReDim Preserve SVC(numpnts%)

'      Don't keep measurements less than 20 cm apart
       If numpnts% <> 0 Then
          If Abs(PD - P(numpnts% - 1)) < 0.2 Then
             GoTo 100
             End If
          End If
       If (TD > TX) Then TX = TD
       If (TD < TN) Then TN = TD
       If (SVD > SVX) Then SVX = SVD
       If (SVD < SVN) Then SVN = SVD
       
       P(numpnts%) = PD
       T(numpnts%) = TD
       C(numpnts%) = CD
       SV(numpnts%) = SVD
       B(numpnts%) = BD
       s(numpnts%) = SD
       d(numpnts%) = DD
       
       If (PD > DMAX) Then NMAX = numpnts%
       If (PD > DMAX) Then DMAX = PD
       If (numpnts% >= 2048) Then Exit Do
       Call VELCALC(PD, TD, SD, SVCC)
       SVC(numpnts%) = SVCC
'C     Include these in the Max/min in case the SV Probe is stuck
       If (SVC(numpnts%) > SVX) Then SVX = SVC(numpnts%)
       If (SVC(numpnts%) < SVN) Then SVN = SVC(numpnts%)
       DSV = SV(numpnts%) - SVC(numpnts%)
       DIFFAV = DIFFAV + DSV
       Call ABSCOEF(DD, SVD, TD, SD, ALPHA)
       ABSAV = ABSAV + ALPHA
       List1.AddItem Str$(numpnts%) & "," & Str$(P(numpnts%)) & "," & _
                     Str$(T(numpnts%)) & "," & Str$(C(numpnts%)) & "," & _
                     Str$(SV(numpnts%)) & "," & Str(SVC(numpnts%)) & _
                     "," & Str(DSV) & "," & Str$(B(numpnts%)) & _
                     "," & Str(s(numpnts%)) & "," & Str$(d(numpnts%)) & _
                     " ," & Str$(ALPHA)
      numpnts% = numpnts% + 1
100 Loop
   Close #filin%
   
   ABSAV = ABSAV / numpnts%
   DIFFAV = DIFFAV / numpnts%
   List1.AddItem "Dmax =" & Str(DMAX) & " m"
   List1.AddItem "SVmin =" & Str(SVN) & " m/s"
   List1.AddItem "SVmax =" & Str(SVX) & " m/s"
   List1.AddItem "Average SVmeas-SVcalc =" & Str(DIFFAV) & " m/s"
   List1.AddItem "Tmin =" & Str(TN) & " ø"
   List1.AddItem "Tmax =" & Str(TX) & " ø"
   List1.AddItem "SVmin=" & Str(SVN) & " dB/km"
   List1.AddItem "Average AbsorptionCoefficient = " & Str(ABSAV)
   
   
   If dumpmode% = 1 Then
        'just include the downgoing data
        For I% = 1 To NMAX
           If SVcheck% = 1 Then
              Print #filout%, Trim$(Str(Format(P(I%), "##0.0#"))) & Str(Format(SV(I%), "##0.0#")) & Str(Format(T(I%), "##0.0##"))
           Else
              Print #filout%, Trim$(Str(Format(P(I%), "##0.0#"))) & Str(Format(SVC(I%), "##0.0#")) & Str(Format(T(I%), "##0.0##"))
           End If
        Next I%
        Print #filout%, "  0  0  0"
        Close #filout%
   Else 'output the three evaluation files that include all the data
        filout1% = FreeFile
        Open FileOut1 For Output As #filout1%
        filout2% = FreeFile
        Open FileOut2 For Output As #filout2%
        filout3% = FreeFile
        Open FileOut3 For Output As #filout3%
        
        For I% = 0 To numpnts% - 1
           Write #filout1%, P(I%), T(I%)
           Write #filout2%, P(I%), SV(I%)
           Write #filout3%, P(I%), SVC(I%)
        Next I%
        Close #filout1%
        Close #filout2%
        Close #filout3%
      End If
      
      List1.AddItem "Points in the Sound Speed Profile: " & Str$(NMAX)
      List1.ListIndex = List1.ListCount - 1
  
   Close #filout%
   
   'release the memory
    ReDim P(0)
    ReDim T(0)
    ReDim C(0)
    ReDim SV(0)
    ReDim B(0)
    ReDim s(0)
    ReDim d(0)
    ReDim SVC(0)
    
    If dumpmode% = 1 Then 'exit
       dumpmode% = 0
       Screen.MousePointer = vbDefault
       MsgBox "Output ""rel"" file: " & FileOut & " written successfully.", vbInformation + vbOKOnly, "Plot"
       Unload JKHplot
       Set JKHplot = Nothing
       JKHplotVis = False
       'clear plot buffer
       'frmSetCond.cmdClear.Value = True
       Exit Sub
       End If
   
   'find short path name
    Dim MaxDirLen As Integer, sShortPath As String
    Dim sPath As String, numOldSelected%

    MaxDirLen = Int(frmSetCond.flxlstFiles.Width / 70) - 30
    'find path
    NumFileOut% = 3
    For I% = 1 To NumFileOut%
        If I% = 1 Then
           FileOut = FileOut1
        ElseIf I% = 2 Then
           FileOut = FileOut2
        ElseIf I% = 3 Then
           FileOut = FileOut3
           End If
        found% = 0
        For J% = Len(FileOut) To 1 Step -1
           If Mid$(FileOut, J%, 1) = "\" Then
              RootName$ = Mid$(FileOut, J% + 1, Len(FileOut) - J% - 4)
              sPath = Mid$(FileOut, 1, Len(FileOut) - Len(RootName$) - 4)
              found% = 1
              Exit For
              End If
        Next J%
        If found% = 1 Then
           'shorten this name to fit into plot buffer List Box
           Call ShortPath(sPath, MaxDirLen, sShortPath)
        Else
           sPath = ""
           End If
        
        If I% = 1 Then
            'set up the Plotinfo for the CSV file
             PlotInfo(0, numSelected%) = "10"
             PlotInfo(1, numSelected%) = "1" 'Points
             PlotInfo(2, numSelected%) = "2" 'Blue
             PlotInfo(3, numSelected%) = "1.0" 'txtXA
             PlotInfo(4, numSelected%) = "0.0" 'txtXB
             PlotInfo(5, numSelected%) = "-1.0" 'txtYA
             PlotInfo(6, numSelected%) = "0.0" 'txtYB
             PlotInfo(7, numSelected%) = PlotInfofrm.lblFileName
             PlotInfo(8, numSelected%) = ""
             PlotInfo(9, numSelected%) = "1"
             PlotInfofrm.cmdCancel.Value = True 'close the PlotInfo Form
             End If
       'place the three evaluation files in the buffer
       'don't normalize
        
        'add the file to plot buffer
        'frmSetCond.lstFiles.AddItem sShortPath & RootName$
        frmSetCond.flxlstFiles.AddItem sShortPath & RootName$
        flxlstFiles.Refresh
        numfiles% = numfiles% + 1
        ReDim Preserve Files(numfiles%)
        Files(numfiles% - 1) = FileOut
        
        ''highlight (select) this file
        'frmSetCond.lstFiles.Selected(frmSetCond.lstFiles.ListCount - 1) = True
        ReDim Preserve PlotInfo(9, numfiles%)
 '       numOldSelected% = numSelected%
 '       numSelected% = frmSetCond.lstFiles.ListCount - 1
 '       If numSelected% >= numPlotInfo% Then
 '          ReDim Preserve PlotInfo(7, numSelected%)
 '          numPlotInfo% = numSelected%
 '          End If
        
        If I% = 1 Then
           'frmSetCond.lstFiles.Selected(numOldSelected%) = False
           'Temp vs Depth
            PlotInfo(0, numfiles% - 1) = "9"
            PlotInfo(1, numfiles% - 1) = "1" 'Points
            PlotInfo(2, numfiles% - 1) = "1" 'Black
            PlotInfo(3, numfiles% - 1) = "1.0" 'txtXA
            PlotInfo(4, numfiles% - 1) = "0.0" 'txtXB
            PlotInfo(5, numfiles% - 1) = "-1.0" 'txtYA
            PlotInfo(6, numfiles% - 1) = "0.0" 'txtYB
            PlotInfo(7, numfiles% - 1) = FileOut1
            PlotInfo(8, numfiles% - 1) = sEmpty
            PlotInfo(9, numfiles% - 1) = "1"
        ElseIf I% = 2 Then
           'measured sound velocity vs depth
            PlotInfo(0, numfiles% - 1) = "9"
            PlotInfo(1, numfiles% - 1) = "1" 'Points
            PlotInfo(2, numfiles% - 1) = "3" 'Green
            PlotInfo(3, numfiles% - 1) = "1.0" 'txtXA
            PlotInfo(4, numfiles% - 1) = "0.0" 'txtXB
            PlotInfo(5, numfiles% - 1) = "-1.0" 'txtYA
            PlotInfo(6, numfiles% - 1) = "0.0" 'txtYB
            PlotInfo(7, numfiles% - 1) = FileOut2
            PlotInfo(8, numfiles% - 1) = sEmpty
            PlotInfo(9, numfiles% - 1) = "1"
        ElseIf I% = 3 Then
           'calculated sound velocity vs depth
            PlotInfo(0, numfiles% - 1) = "9"
            PlotInfo(1, numfiles% - 1) = "1" 'Points
            PlotInfo(2, numfiles% - 1) = "5  'Red"
            PlotInfo(3, numfiles% - 1) = "1.0" 'txtXA
            PlotInfo(4, numfiles% - 1) = "0.0" 'txtXB
            PlotInfo(5, numfiles% - 1) = "-1.0" 'txtYA
            PlotInfo(6, numfiles% - 1) = "0.0" 'txtYB
            PlotInfo(7, numfiles% - 1) = FileOut3
            PlotInfo(8, numfiles% - 1) = sEmpty
            PlotInfo(9, numfiles% - 1) = "1"
           End If
        
'        PlotInfofrm.optLine.Value = True
        'keep track of position of file in plot buffer
'        numSelected% = frmSetCond.lstFiles.ListCount - 1
'        PlotInfofrm.cmdAccept.Value = True
'        PlotInfo(7, numSelected%) = FileOut
   Next I%
   
   'change the titles
   frmSetCond.txtXTitle.Text = "depth(m)"
   frmSetCond.txtYTitle.Text = "T,SV,SVC"
   frmSetCond.txtTitle.Text = "JKH bathyemtry plot"

   'reshape the conversion program to be a data summary sheet
   'JKHplot.cmdConvert.Visible = False
   JKHplot.cmdConvert.Caption = "Write ""rel"" file"
   'JKHplot.Height = JKHplot.List1.Height + 450
   JKHplot.Caption = "Data summary"
   JKHplot.Top = frmDraw.Top + frmDraw.Height - JKHplot.Height
   JKHplot.optSV.Enabled = True
   JKHplot.optSVC.Enabled = True
   dumpmode% = 1
   
   Screen.MousePointer = vbDefault
   Exit Sub
   
errhand:
    Screen.MousePointer = vbDefault
    MsgBox "Encountered error number: " & Err.Number & vbLf & _
           Err.Description & vbLf & _
           "Stopping the wizard and closing the files.", vbCritical + vbOKOnly, "Plot"
           Close
End Sub

Sub JULDAY(ID As Integer, IM As Integer, IYR As Integer, JD As Long)
      Dim MD(13) As Integer
      'determine day number
      MD(1) = 0
      MD(2) = 31
      MD(3) = 59
      MD(4) = 90
      MD(5) = 120
      MD(6) = 151
      MD(7) = 181
      MD(8) = 212
      MD(9) = 243
      MD(10) = 273
      MD(11) = 304
      MD(12) = 334
      MD(13) = 365
      KLY = 0
      DJ = CDbl(IYR - 1900) * 365.25
      JD = Fix(DJ + 0.75)
      
      'IF(DJ-JD)10,5,10
      If (DJ - JD) < 0 Then GoTo 10
      If (DJ - JD) = 0 Then GoTo 5
      If (DJ - JD) > 0 Then GoTo 10
      
5     If (IM >= 3) Then KLY = 1
10    JD = JD + Fix(MD(IM) + ID - 1 + KLY)
End Sub

      Sub VELCALC(d As Single, TC As Single, s As Single, SV As Single)
'C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - AUGUST 3 2001
'C     VELCALC uses the ideas in the EM1002 Operation Manual pg. 383, altered to
'C      work in the Red Sea, Mediterranean, and Persian Gulf.
'      REAL*4 D,SV,TC,S,Z,Z2,Z3,T,T2,T3,SN,C0ST,C2ZST
      Z = 0.001 * d
      Z2 = Z * Z
      T = 0.1 * TC
      T2 = T * T
      T3 = T2 * T
      SN = s - 35#
'C     Coppens, A. B., 1981, J. Acoust. Soc. Am., 69(3), pp. 862-863. Eqn.4.
      C0ST = 1449.05 + 45.7 * T - 5.21 * T2 + 0.23 * T3 + _
               (1.333 - 0.126 * T + 0.009 * T2) * SN
'C     Coppens, A. B., 1981, J. Acoust. Soc. Am., 69(3), pp. 862-863. Eqn. 7.
      C2ZST = C0ST + (16.23 + 0.253 * T) * Z + (0.213 - 0.1 * T) * Z2 + _
               (0.016 + 0.0002 * SN) * SN * T * Z
      SV = C2ZST
 End Sub
     
      Sub ABSCOEF(d As Single, SV As Single, T As Single, s As Single, ALPHA As Single)
'C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - AUGUST 3 2001
'C     ABSCOEF calculates the Absorption Coefficient using the equations
'C      in the EM1002 Operation Manual, pg. 384.
'C     pH can range from 7.6 to 8.2
'      REAL*4 D,SV,T,S,ALPHA,PH,Z,SF,SFSQ,A1,A2,P2,P3,SF1,SF2,
'     *       TERM1,TERM2,TERM3
      PH = 7.9
      Z = d * 0.001
      SF = 95#
      SFSQ = SF * SF
      A1 = (8.86 - 10# ^ (0.78 * PH - 5#)) / SV
      A2 = (21.44 * s * (1# + 0.025 * T)) / SV
      If (T <= 20#) Then A3 = 0.0004937 - T * (0.0000259 - T * (0.000000911 - T * 0.000000015))
      If (T >= 20#) Then A3 = 0.0003964 - T * (0.00001146 - T * (0.000000145 - T * 0.00000000065))
      P2 = 1# - Z * (0.137 - 0.0062 * Z)
      P3 = 1# - Z * (0.0383 - Z * 0.00049)
      SF1 = 2.8 * Sqr(s * 10# ^ (4# - 1245# / (T + 273#)) / 35#)
      SF2 = (8.17 * 10# ^ (8# - 1990# / (T + 273#))) / (1# + 0.0018 * (s - 35#))
      TERM1 = A1 * SF1 * SFSQ / (SFSQ + SF1 * SF1)
      TERM2 = A2 * P2 * SF2 * SFSQ / (SFSQ + SF2 * SF2)
      TERM3 = A3 * P3 * SFSQ
      ALPHA = TERM1 + TERM2 + TERM3
End Sub



Private Sub Form_Load()
   'JKHplot.chkVel.Enabled = True
   JKHplot.optSV.Enabled = False
   JKHplot.optSVC.Enabled = False
   JKHplot.cmdConvert.Enabled = True
   JKHplotVis = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set JKHplot = Nothing
   JKHplotVis = False
End Sub


Private Sub optSV_Click()
   SVcheck% = 1
End Sub

Private Sub optSVC_Click()
   SVcheck% = 0
End Sub
