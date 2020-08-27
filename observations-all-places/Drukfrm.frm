VERSION 5.00
Begin VB.Form Drukfrm 
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame progressfrm 
      Height          =   735
      Left            =   360
      TabIndex        =   21
      Top             =   8400
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox picProgBar 
         Height          =   340
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   3555
         TabIndex        =   22
         Top             =   220
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdRunNetzski6 
      Caption         =   "Run Netzski6"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame frmInv 
      Caption         =   "Inversions"
      Height          =   1095
      Left            =   360
      TabIndex        =   13
      Top             =   7200
      Width           =   3855
      Begin VB.CommandButton cmdAveSonde 
         Caption         =   "Avg Sondes"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlotFiles 
         Caption         =   "Plot File\"
         Height          =   315
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdFix2 
         Caption         =   "Fix 2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdFix 
         Caption         =   "Fix-4"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdInv 
         Caption         =   "Inversion Search"
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frmObserver 
      Caption         =   "Observation set"
      Height          =   1755
      Left            =   2040
      TabIndex        =   10
      Top             =   900
      Width           =   2295
      Begin VB.OptionButton optZuriel 
         Caption         =   "Rav Zuriel (places in Bnei Brak)"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optGolan 
         Caption         =   "R' Golan (MA shul N'vei Ya'akov)"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optDruk 
         Caption         =   "Rabbi Druk (Armon Hanatziv)"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmTK 
      Caption         =   "Temperature modeling"
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton optAveTK 
         Caption         =   "Use WordClim Average Termparatures"
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   3015
      End
      Begin VB.OptionButton optMinTK 
         Caption         =   "Use WordClim Minimum Temperatures"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
   End
   Begin VB.Frame extractfrm 
      Caption         =   "Extract refrac. vs temp. for specific hgt"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   6000
      Width           =   3855
      Begin VB.CommandButton cmdLoop 
         Caption         =   "Loop"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cmbHeight 
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Text            =   "cmbHeight"
         ToolTipText     =   "Choose a height"
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdVisTimes 
      Caption         =   "Prepare Vis-ast for plotting"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Do final buble sort and output as csv file"
      Top             =   5400
      Width           =   2535
   End
   Begin VB.ListBox lstSort 
      Height          =   2400
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2760
      Width           =   4095
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join files into one sorted file"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Join files and initial sort by ascending day number"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Write renormalized druke files"
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Drukfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub HeightLoopRun()
   Dim fileeps As Integer, fileref As Integer, filein As Integer, HeightVal As Integer
   Dim DataCols() As String, a(4) As Double, b(4) As Double, numRow As Integer
   Dim EpsVal As Double, RefVal As Double, TimeStr As String
   
   DoEvents
   cmbHeight.Refresh
   
   If Initialize And Not HeightLoop Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   cmbHeight.Enabled = False
   
   If HeightLoop Then
      HeightVal = 0
      cmbHeight.Text = Trim(Str(HeightVal))
      End If
      
100:
   If HeightLoop Then
      HeightVal = HeightVal + 200
      cmbHeight.Text = Trim(Str(HeightVal))
      If HeightVal > 3000 Then
         HeightLoop = False
         Exit Sub
         End If
   Else
      HeightVal = Val(cmbHeight.Text)
      End If

   fileeps = FreeFile
   Open App.Path & "\EPSvsTemp-hgt" & Trim(cmbHeight.Text) & ".csv" For Output As #fileeps
   fileref = FreeFile
   Open App.Path & "\REFvsTemp-hgt" & Trim(cmbHeight.Text) & ".csv" For Output As #fileref
   HeightVal = Val(cmbHeight.Text)
   Dim directory As String
   Dim FileRoot As String
   Dim FullFileName As String
   Dim Temp As Integer
   directory = "e:\AtmRef\"
   FileRoot = directory & "TR_VDW_"
   For Temp = 260 To 320 Step 3
      FullFileName = FileRoot & Trim(Str(Temp)) & "_0_32.dat"
      If Dir(FullFileName) <> "" Then
         'extract eps and ref at this requested height and add to file
         filein = FreeFile
         numRow = 0
         Open FullFileName For Input As #filein
         Do Until EOF(filein)
            Input #filein, a(0), a(1), a(2), a(3), a(4)
            DoEvents
            If a(2) = 0# Then
                numRow = numRow + 1
                If numRow = 1 Then
                   b(0) = a(0)
                   b(1) = a(1)
                   b(2) = a(2)
                   b(3) = a(3)
                   b(4) = a(4)
                Else
                    If b(1) <= HeightVal And a(1) > HeightVal Then
                       'interpolate the zero view angle refraction vs height profile
                       EpsVal = (HeightVal - b(1)) * ((a(3) - b(3)) / (a(1) - b(1))) + b(3)
                       RefVal = (HeightVal - b(1)) * ((a(4) - b(4)) / (a(1) - b(1))) + b(4)
                       Write #fileeps, Temp, EpsVal
                       Write #fileref, Temp, RefVal
                       Exit Do
                       End If
                   End If
               End If
               DoEvents
         Loop
         Close #filein
         End If
   Next Temp
   Close #fileeps
   Close #fileref
   
   Screen.MousePointer = vbDefault
   cmbHeight.Enabled = True
   
   If HeightLoop Then GoTo 100
   
End Sub

Private Sub cmbHeight_Change()
   If Not HeightLoop Then Call HeightLoopRun
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAveSonde_Click
' Author    : Dr-John-K-Hall
' Date      : 5/6/2020
' Purpose   : create average summer and winter sonde from the list of used sondes
'---------------------------------------------------------------------------------------
'
Private Sub cmdAveSonde_Click()
   Dim FileNameIn As String, SplitDoc() As String, doclin$, filein%, DateSonde$
   Dim SumPlotfile As String, WinPlotfile As String, Dif As Double
   Dim SondeFileName As String, AveTemp(2, 701) As Double, AvePress(2, 701) As Double, filesonde As Integer
   Dim ho As Double, Temp0 As Double, h1 As Double, Temp1 As Double, p0 As Double, p1 As Double
   Dim numTmp As Integer, hgtStart As Double, numSum As Integer, numWin As Integer
   Dim numSeason As Integer, numWinTot As Integer, numSumTot As Integer, numWinFiles As Integer, numSumFiles As Integer
   ', DifThresholdW As Double, DifThresholdS As Double
'   Dim UseThresholds As Boolean
   
   On Error GoTo cmdAveSonde_Click_Error

'   DifThresholdW = -0.3
'   DifThresholdS = -0.1
'   UseThresholds = False
   
   SumPlotfile = App.Path & "\Sum-Ave-sonde.txt"
   WinPlotfile = App.Path & "\Win-Ave-sonde.txt"
   
   filesum = FreeFile
   Open SumPlotfile For Output As #filesum
   filewin = FreeFile
   Open WinPlotfile For Output As #filewin
   
   'average sondes files will be extrapolated from all the sum/win sondes in the Druk-all-dates file
   'the average sonde will be generated by interpolating each sonde to a 50 meter grind from 0 to 11,000 meters (end of troposphere)
'   Print #filesum, "This file is used by Plot. Don't erase it!"
'   Print #filesum, "X-values"
'   Print #filesum, "Y-values"
'   Print #filesum, """"""
'   Print #filewin, "This file is used by Plot. Don't erase it!"
'   Print #filewin, "X-values"
'   Print #filewin, "Y-values"
'   Print #filewin, """"""
   
   FileNameIn = App.Path & "\Druk-all-dates.csv"
   filein% = FreeFile
   numFiles = 0
   Open FileNameIn For Input As #filein%
   Do Until EOF(filein%)
      Line Input #filein%, doclin$
      SplitDoc = Split(doclin$, ",")
      DateSonde$ = Mid$(SplitDoc(0), 2, Len(SplitDoc(0)) - 2)
'      Dif = Val(SplitDoc(7))
      
      SondeFileName = App.Path & "\" & DateSonde$ & "-sondes.txt"
      If InStr(DateSonde$, "Jan") Or InStr(DateSonde$, "Feb") Or InStr(DateSonde$, "Nov") Or InStr(datemode$, "Dec") Then
         numSeason = 0 'winter
         numWin = 0
         numWinFiles = numWinFiles + 1
      Else
         numSeason = 1 'summer
         numSum = 0
         numSumFiles = numSumFiles + 1
         End If
         
      filesonde = FreeFile
      'zero the counters

      Open SondeFileName For Input As #filesonde
      numTmp = 0
      Do Until EOF(filesonde)
         Input #filesonde, h0, Temp0, p0
50
         If EOF(filesonde) Then Exit Do
         Input #filesonde, h1, Temp1, p1
         'determine ground temperature
         If numTmp = 0 Then

            hgtStart = 0
         
            If numSeason = 0 Then
               numTmp = numWin
            ElseIf numSeason = 1 Then
               numTmp = numSum
               End If
            AveTemp(numSeason, numTmp) = AveTemp(numSeason, numTmp) - h0 * (Temp1 - Temp0) / (h1 - h0) + Temp0
            AvePress(numSeason, numTmp) = AvePress(numSeason, numTmp) - h0 * (p1 - p0) / (h1 - h0) + p0
            numTmp = numTmp + 1
            hgtStart = hgtStart + 50
            If hgtStart > 35000 Then Exit Do
            If numSeason = 0 Then
               numWin = numWin + 1
            Else
               numSum = numSum + 1
               End If
            
            If h0 <= hgtStart And hgtStart < h1 Then
100:
                If numSeason = 0 Then
                   numTmp = numWin
                ElseIf numSeason = 1 Then
                   numTmp = numSum
                   End If
                   
               AveTemp(numSeason, numTmp) = AveTemp(numSeason, numTmp) + (hgtStart - h0) * (Temp1 - Temp0) / (h1 - h0) + Temp0
               AvePress(numSeason, numTmp) = AvePress(numSeason, numTmp) + (hgtStart - h0) * (p1 - p0) / (h1 - h0) + p0

               hgtStart = hgtStart + 50
               If hgtStart > 35000 Then Exit Do

               If numSeason = 0 Then
                  numWin = numWin + 1
               Else
                  numSum = numSum + 1
                  End If
                  
               If h0 <= hgtStart And hgtStart < h1 Then
                  GoTo 100
                  End If
                  
               h0 = h1
               Temp0 = Temp1
               p0 = p1
               GoTo 50
            ElseIf h0 < hgtStart And h1 < hgtStart Then
               GoTo 50
               End If
               
            h0 = h1
            p0 = p1
            Temp0 = Temp1
               
            GoTo 50
            End If
         
         If h0 <= hgtStart And hgtStart < h1 Then
200:
            If numSeason = 0 Then
               numTmp = numWin
            ElseIf numSeason = 1 Then
               numTmp = numSum
               End If
               
            AveTemp(numSeason, numTmp) = AveTemp(numSeason, numTmp) + (hgtStart - h0) * (Temp1 - Temp0) / (h1 - h0) + Temp0
            AvePress(numSeason, numTmp) = AvePress(numSeason, numTmp) + (hgtStart - h0) * (p1 - p0) / (h1 - h0) + p0
            
            hgtStart = hgtStart + 50
            
            If hgtStart > 35000 Then Exit Do
            numTmp = numTmp + 1
            If numSeason = 0 Then
               numWin = numWin + 1
            Else
               numSum = numSum + 1
               End If
               
           If ho <= hgtStart And hgtStart < h1 Then
              GoTo 200
              End If
            h0 = h1
            p0 = p1
            Temp0 = Temp1
            GoTo 50
         ElseIf h0 < hgtStart And h1 < hgtStart Then
            GoTo 50
            End If

      Loop
      Close #filesonde
      numWinTot = numWin
      numSumTot = numSum
      'zero the counters
   
'      If InStr(DateSonde$, "Jan") Or InStr(DateSonde$, "Feb") Or InStr(DateSonde$, "Nov") Or InStr(datemode$, "Dec") Then
'         If Not UseThresholds Or Dif <= DifThresholdW Then
'            Print #filewin, """ 2"","" 0"","" 0"",""1"",""0"",""1"",""0"",""" & SondeFileName & """,""none:none"",""3"""
'            End If
'      ElseIf InStr(DateSonde$, "May") Or InStr(DateSonde$, "Jun") Or InStr(DateSonde$, "Jul") Then
'         If Not UseThresholds Or (Dif > -0.2 And Dif <= DifThresholdS) Then
'            Print #filesum, """ 2"","" 0"","" 0"",""1"",""0"",""1"",""0"",""" & SondeFileName & """,""none:none"",""3"""
'            End If
'         End If
         
   Loop
   Close #filein%
   
   'set ground pressure to pressure at the surface
   AvePress(0, 1) = AvePress(0, 0)
   AvePress(1, 1) = AvePress(1, 0)
   
    For i = 1 To numWinTot
        AveTemp(0, i - 1) = AveTemp(0, i - 1) / numWinFiles
        AvePress(0, i - 1) = AvePress(0, i - 1) / numWinFiles
        Write #filewin, i * 50, AveTemp(0, i - 1), AvePress(0, i - 1)
      Next i
      For i = 1 To numSumTot
        AveTemp(1, i - 1) = AveTemp(1, i - 1) / numSumFiles
        AvePress(1, i - 1) = AvePress(1, i - 1) / numSumFiles
        Write #filesum, i * 50, AveTemp(1, i - 1), AvePress(1, i - 1)
    Next i
    
    Close #filewin
    Close #filesum
    Close #filesum
    Close #filewin
   
   

   On Error GoTo 0
   Exit Sub

cmdAveSonde_Click_Error:
    Close
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAveSonde_Click of Form Drukfrm"

End Sub

Private Sub cmdFix_Click()
   Dim free4 As Long, free3 As Long, free5 As Long
   Dim Split3() As String, Split4() As String, Date4 As String, Date3 As String, SplitDate() As String, Split5(19) As String
   Dim time4 As String, time5 As String, yr3 As Integer, yr4 As Integer, day3 As Integer, day4 As Integer, mon3 As Integer, mon4 As Integer
   Dim dayNum3 As Integer, dayNum4 As Integer, daysYear3 As Integer, daysYear4 As Integer
   Dim hr3 As Double, hr4 As Double, Hour3() As String, Hour4() As String
   
   free3 = FreeFile
   Open App.Path & "\Jerusalem-sondes-Druk\" & "ims_data - Jerusalem 2009-2013-3.csv" For Input As #free3
   free4 = FreeFile
   Open App.Path & "\Jerusalem-sondes-Druk\" & "ims_data - Jerusalem 2009-2013-4.csv" For Input As #free4
   free5 = FreeFile
   Open App.Path & "\Jerusalem-sondes-Druk\" & "ims_data - Jerusalem 2009-2013-5.csv" For Output As #free5
   
   Line Input #free3, doclin3$
   Line Input #free4, doclin4$  'header line
   Print #free5, doclin3$
   
   'fill in missing data in free4 from free3
fix500:
      If Not EOF(free3) Then
        Line Input #free3, doclin3$
        Split3 = Split(doclin3$, ",")
        'determine date and time
        Date3 = Split3(0)
        SplitDate = Split(Date3, "/")
        day3 = Val(SplitDate(0))
        mon3 = Val(SplitDate(1))
        yr3 = Val(SplitDate(2))
        time3 = Split3(1)
        Hour3 = Split(time3, ":")
        hr3 = Val(Hour3(0)) + Val(Hour3(1)) / 60#
        daysYear3 = DaysinYear(yr3)
        dayNum3 = DayNumber(daysYear3, mon3, day3)
      Else
        GoTo fix900
        End If
      
      If Not EOF(free4) Then
        Line Input #free4, doclin4$
        Split4 = Split(doclin4$, ",")
        'determine date and time
        Date4 = Split4(0)
        SplitDate = Split(Date4, "/")
        day4 = Val(SplitDate(0))
        mon4 = Val(SplitDate(1))
        yr4 = Val(SplitDate(2))
        time4 = Split4(1)
        Hour4 = Split(time4, ":")
        If (Hour4(0) = 23) Then
           ccc = 1
           End If
        hr4 = Val(Hour4(0)) + Val(Hour4(1)) / 60#
        daysYear4 = DaysinYear(yr4)
        dayNum4 = DayNumber(daysYear4, mon4, day4)
      Else
        GoTo fix900
        End If
      
fix700:
      If dayNum3 = dayNum4 And yr3 = yr4 And hr3 = hr4 Then
         'fill in the missing info and write to new file
         Split5(0) = Split3(0)
         Split5(1) = Split3(1)
         For i = 2 To 18
            If Split3(i) = "-" And Split4(i) <> "-" Then
               Split5(i) = Split4(i)
            ElseIf Split3(i) <> "-" And Split4(i) = "-" Then
               Split5(i) = Split3(i)
            Else
               Split5(i) = "-"
               End If
         Next i
         Write #free5, Split5(0), Split5(1), Split5(2), Split5(3), Split5(4), Split5(5), _
                       Split5(6), Split5(7), Split5(8), Split5(9), Split5(10), Split5(11), _
                       Split5(12), Split5(13), Split5(14), Split5(15), Split5(16), Split5(17), _
                       Split5(18)
          GoTo fix500
     ElseIf dayNum3 = dayNum4 And yr3 = yr4 And hr4 < hr3 Then
        'increment file 4
        If Not EOF(free4) Then
          Line Input #free4, doclin4$
          Split4 = Split(doclin4$, ",")
          'determine date and time
          Date4 = Split4(0)
          SplitDate = Split(Date4, "/")
          day4 = Val(SplitDate(0))
          mon4 = Val(SplitDate(1))
          yr4 = Val(SplitDate(2))
          time4 = Split4(1)
          Hour4 = Split(time4, ":")
          hr4 = Val(Hour4(0)) + Val(Hour4(1)) / 60#
          daysYear4 = DaysinYear(yr4)
          dayNum4 = DayNumber(daysYear4, mon4, day4)
          GoTo fix700
        Else
          GoTo fix900
          End If
     ElseIf dayNum3 = dayNum4 And yr3 = yr4 And hr4 > hr3 Then
        'increment file3
        If Not EOF(free3) Then
          Line Input #free3, doclin3$
          Split3 = Split(doclin3$, ",")
          'determine date and time
          Date3 = Split3(0)
          SplitDate = Split(Date3, "/")
          day3 = Val(SplitDate(0))
          mon3 = Val(SplitDate(1))
          yr3 = Val(SplitDate(2))
          time3 = Split3(1)
          Hour3 = Split(time3, ":")
          hr3 = Val(Hour3(0)) + Val(Hour3(1)) / 60#
          daysYear3 = DaysinYear(yr3)
          dayNum3 = DayNumber(daysYear3, mon3, day3)
          GoTo fix700
        Else
          GoTo fix900
          End If
     ElseIf dayNum3 > dayNum4 And yr3 = yr4 Then
        'increment file 4
        If Not EOF(free4) Then
          Line Input #free4, doclin4$
          Split4 = Split(doclin4$, ",")
          'determine date and time
          Date4 = Split4(0)
          SplitDate = Split(Date4, "/")
          day4 = Val(SplitDate(0))
          mon4 = Val(SplitDate(1))
          yr4 = Val(SplitDate(2))
          time4 = Split4(1)
          Hour4 = Split(time4, ":")
          hr4 = Val(Hour4(0)) + Val(Hour4(1)) / 60#
          daysYear4 = DaysinYear(yr4)
          dayNum4 = DayNumber(daysYear4, mon4, day4)
          GoTo fix700
        Else
          GoTo fix900
          End If
     ElseIf dayNum3 < dayNum4 And yr3 = yr4 Then
        'increment file3
        If Not EOF(free3) Then
          Line Input #free3, doclin3$
          Split3 = Split(doclin3$, ",")
          'determine date and time
          Date3 = Split3(0)
          SplitDate = Split(Date3, "/")
          day3 = Val(SplitDate(0))
          mon3 = Val(SplitDate(1))
          yr3 = Val(SplitDate(2))
          time3 = Split3(1)
          Hour3 = Split(time3, ":")
          hr3 = Val(Hour3(0)) + Val(Hour3(1)) / 60#
          daysYear3 = DaysinYear(yr3)
          dayNum3 = DayNumber(daysYear3, mon3, day3)
          GoTo fix700
        Else
          GoTo fix900
          End If
     ElseIf yr3 < yr4 Then
        'increment free3
        If Not EOF(free3) Then
          Line Input #free3, doclin3$
          Split3 = Split(doclin3$, ",")
          'determine date and time
          Date3 = Split3(0)
          SplitDate = Split(Date3, "/")
          day3 = Val(SplitDate(0))
          mon3 = Val(SplitDate(1))
          yr3 = Val(SplitDate(2))
          time3 = Split3(1)
          Hour3 = Split(time3, ":")
          hr3 = Val(Hour3(0)) + Val(Hour3(1)) / 60#
          daysYear3 = DaysinYear(yr3)
          dayNum3 = DayNumber(daysYear3, mon3, day3)
          GoTo fix700
        Else
          GoTo fix900
          End If
    ElseIf yr3 > yr4 Then 'increment free4
        'increment file 4
        If Not EOF(free4) Then
          Line Input #free4, doclin4$
          Split4 = Split(doclin4$, ",")
          'determine date and time
          Date4 = Split4(0)
          SplitDate = Split(Date4, "/")
          day4 = Val(SplitDate(0))
          mon4 = Val(SplitDate(1))
          yr4 = Val(SplitDate(2))
          time4 = Split4(1)
          Hour4 = Split(time4, ":")
          hr4 = Val(Hour4(0)) + Val(Hour4(1)) / 60#
          daysYear4 = DaysinYear(yr4)
          dayNum4 = DayNumber(daysYear4, mon4, day4)
          GoTo fix700
        Else
          GoTo fix900
          End If
  Else
      ccc = 1
      End If
        
fix900:

    Close #file3
    Close #file4
    Close #file5
   
End Sub

Private Sub cmdFix2_Click()
   Dim filein As Integer, FileInName As String, SplitLine() As String, doclin$
   Dim fileout As Integer, FileOutName As String, NewDate() As String, daysintheyear As Integer, daysnum As Integer
   Dim daynew As Integer, monnew As Integer, TotalNum As Long, yrnew As Integer, StartDayNum As Long
   
   FileInName = App.Path & "\Compare-refr.csv"
   filein = FreeFile
   Open FileInName For Input As #filein
   fileout = FreeFile
   FileOutName = App.Path & "\Compare-refr-more.csv"
   Open FileOutName For Output As #fileout
   
   Do Until EOF(filein)
      Line Input #filein, doclin$
      SplitLine = Split(doclin$, ",")
      'convert date to UXA format, and then determine daynumber and add daynumber column
      NewDate = Split(SplitLine(0), "/")
      daynew = Val(NewDate(0))
      monnew = Val(NewDate(1))
      yrnew = Val(NewDate(2))
      If yrnew < 100 Then yrnew = 1900 + yrnew
      USformat$ = NewDate(1) + "/" + NewDate(0) + "/" + NewDate(2)
      daysintheyear = DaysinYear(yrnew)
      daysnum = DayNumber(daysintheyear, monnew, daynew)
      If yrnew = 1985 Then
         StartDayNum = 0
      ElseIf yrnew = 1986 Then
         StartDayNum = DaysinYear(1985) + 1
      ElseIf yrnew = 1987 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + 1
      ElseIf yrnew = 1988 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + 1
      ElseIf yrnew = 1989 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + 1
      ElseIf yrnew = 1990 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + DaysinYear(1989) + 1
      ElseIf yrnew = 1991 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + DaysinYear(1989) + DaysinYear(1990) + 1
      ElseIf yrnew = 1992 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + DaysinYear(1989) + DaysinYear(1990) + DaysinYear(1991) + 1
      ElseIf yrnew = 1993 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + DaysinYear(1989) + DaysinYear(1990) + DaysinYear(1991) + DaysinYear(1992) + 1
      ElseIf yrnew = 1994 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + DaysinYear(1989) + DaysinYear(1990) + DaysinYear(1991) + DaysinYear(1992) + DaysinYear(1993) + 1
      ElseIf yrnew = 1995 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + DaysinYear(1989) + DaysinYear(1990) + DaysinYear(1991) + DaysinYear(1992) + DaysinYear(1993) + DaysinYear(1994) + 1
      ElseIf yrnew = 1996 Then
         StartDayNum = DaysinYear(1985) + DaysinYear(1986) + DaysinYear(1987) + DaysinYear(1988) + DaysinYear(1989) + DaysinYear(1990) + DaysinYear(1991) + DaysinYear(1992) + DaysinYear(1993) + DaysinYear(1994) + DaysinYear(1995) + 1
         End If
      TotalNum = StartDayNum + daysnum
      Write #fileout, TotalNum, USformat$, Val(SplitLine(1)), Val(SplitLine(2)), Val(SplitLine(3)), Val(SplitLine(4)), Val(SplitLine(5))
   Loop
   Close #filein
   Close #fileout
End Sub

Private Sub cmdInv_Click()

Dim FileInNameDruk(12) As String, filein As Integer, filecalc As Integer
Dim FileInNameGolan(5) As String
Dim FileCalcName As String
Dim FileWeather As String
Dim FileWeatherResults As String
Dim DifYear(366) As Double
Dim InputSplit() As String, SplitWeather() As String, SplitDate() As String
Dim dysplit As Double, difsplit As Double, dat$, difCompare As Double
Dim Nexday As Integer, SpeedWind As Single, CloudCover As Integer
Dim TimeDay As String, nofilter As Boolean, yr As Integer, yl As Integer
Dim aaday As Integer, aamon As Integer, aayear As Integer
Dim bbday As Integer, bbmon As Integer, bbyear As Integer

nofilter = True  '= true 'no filtering, but list in separte column the wind speed (mph) and the cloud cover
                 '= 'false for filtering days based on wind speed (mph) and cloud cover


  'find days of low wind speed and no clouds and see if it correlates to days of early sunrises
  
filein = FreeFile
If optDruk.Value = True Then
    FileWeather = App.Path & "\Jerusalem-sondes-Druk\" & "ims_data - Jerusalem airport 1985-1996-2.csv" 'ims_data - Jerusalem 1985-1996-2.csv"

    If optMinTK.Value = True Then
        FileInNameDruk(0) = "drukvdw_compare.001"
        FileInNameDruk(1) = "drukvdw_compare.002"
        FileInNameDruk(2) = "drukvdw_compare.003"
        FileInNameDruk(3) = "drukvdw_compare.004"
        FileInNameDruk(4) = "drukvdw_compare.005"
        FileInNameDruk(5) = "drukvdw_compare.006"
        FileInNameDruk(6) = "drukvdw_compare.007"
        FileInNameDruk(7) = "drukvdw_compare.008"
        FileInNameDruk(8) = "drukvdw_compare.009"
        FileInNameDruk(9) = "drukvdw_compare.010"
        FileInNameDruk(10) = "drukvdw_compare.011"
        FileInNameDruk(11) = "drukvdw_compare.012"
        FileCalcName = App.Path & "\RavD_NO_mt_1995.csv" 'no adhoc sunrise fix of subtracting 15 seconds
        
        filecalcin = FreeFile
        Open FileCalcName For Input As #filecalcin
        numdy = 0
        Do Until EOF(filecalcin)
            Input #filecalcin, dy, DifYear(numdy)
            numdy = numdy + 1
        Loop
        Close #filecalcin
        
        fileweatherin = FreeFile
        Open FileWeather For Input As #fileweatherin
        Line Input #fileweatherin, doclin$ 'read doc line
        
        
        FileWeatherResults = App.Path & "\Druk-mt-weather-compare-6.csv"
        fileresults = FreeFile
        Open FileWeatherResults For Output As #fileresults
        
        For i = 0 To 11
           yr = 1985 + i
           'determine if it is leap year
           yl = DaysinYear(yr)
           
           filedrukin = FreeFile
           Open App.Path & "\" & FileInNameDruk(i) For Input As #filedrukin
           Do Until EOF(filedrukin)
comp250:
              Line Input #filedrukin, doclin$
              InputSplit = Split(doclin$, ",")
              dat$ = InputSplit(0)
              If InStr(dat$, """") Then dat$ = Mid$(dat$, 2, 10)
              dysplit = Val(InputSplit(1))
              difsplit = Val(InputSplit(2))
              'interpolate dif in compare file via the dysplit and determine if refraction is greater
              'if refraction is greater than expected, then look for corresponding date in weather file
              'and determine the wind and weather info and record to the results file
              nextday = Fix(dysplit) + 1
              If nextday >= 2 Then
                 difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(nextday - 2)) + DifYear(nextday - 2)
              Else 'daysplit is lest then 1, so wrap to dec 31
                 difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(364)) + DifYear(364)
                 End If
                 
              'find corresponding date in weather file and
comp500:        If EOF(fileweatherin) Then
                   Exit Do
                   End If
                Line Input #fileweatherin, doclin$
                SplitWeather = Split(doclin$, ",")
                TimeDay = SplitWeather(1)
                
                bb$ = SplitWeather(0) 'Format(SplitWeather(0), "dd/mm/yyyy")

                SplitDate = Split(bb$, "/")
                bbday = Val(SplitDate(0))
                bbmon = Val(SplitDate(1))
                bbyear = Val(SplitDate(2))
                
'                bbday$ = Mid$(bb$, 4, 2)
'                bbmon$ = Mid$(bb$, 1, 2)
'                bbyear$ = Mid$(bb$, 7, 4)
                
'                aa$ = Format(dat$, "dd/mm/yyyy")
'                If dat$ = "02/06/1988" Then
'                   ccc = 1
'                   End If
'
'                If bb$ = "1/6/1988" Then
'                   ccc = 1
'                   End If
comp750:
                SplitDate = Split(dat$, "/")
                aaday = Val(SplitDate(0))
                aamon = Val(SplitDate(1))
                aayear = Val(SplitDate(2))
                
'                If (bbyear = 1989) Then
'                   ccc = 1
'                   End If
         
'                aaday$ = Mid$(aa$, 4, 2)
'                aamon$ = Mid$(aa$, 1, 2)
'                aayear$ = Mid$(aa$, 7, 4)
'                If aaday$ = bbday$ And aamon$ = bbmon$ And aayear$ = bbyear$ Then
                If (aaday = bbday And aamon = bbmon And aayear = bbyear) And (TimeDay = "5:00" Or TimeDay = "8:00") Then
                      '-5 is days where there are no weather data
                      '-6 is earlier sunrises with relaxed conditions wind speed <=5 and cloudcover <= 1
                      '-7 is above with later sunrises
                      '-7 analyses all days that have weather associated with them by printing their wind speed (mph), and cloud cover
                      '-8 strict zero wind and cloud conditions but no filtering of whether above or below calculated values
                      '-9 fix(wind) <= 1, fix(clouds) <=1 and ""
                      '-10 fix(wind) <= 2, fix(clouds) <=1 and ""
                      '-11 fix(wind) <= 3, fix(clouds) <=1 and ""
                      '-12 fix(wind) <= 4, fix(clouds) <=1 and ""
                      '-13 fix(wind) <= 5, fix(clouds) <=1 and ""
                      '-14 fix(wind) <= 5, fix(clouds) <=3 and ""
                      '-15 fix(wind) anythin, fix(clouds) = 0 and ""
                
                    If difsplit < difCompare Or difsplit >= difCompare Then
                      'look for corresponding weather data for this date
                
                       'record to results if wind is less then 5 mph and sky is mostly clear, i.e., total cloud cover <=1
                       If SplitWeather(6) <> "-" And SplitWeather(10) <> "-" Then
                            SpeedWind = Val(SplitWeather(6)) * 2.23694 'mph
                            CloudCover = Val(SplitWeather(10))
                            
                            If nofilter Then
                                Print #fileresults, dat$, difsplit - difCompare, dysplit, difsplit, SpeedWind, CloudCover
                            Else
                                If Fix(SpeedWind) <= 0 And Fix(CloudCover) <= 0 Then 'record this data
                                   Print #fileresults, dat$, difsplit - difCompare, dysplit, difsplit, SpeedWind, CloudCover
                                   End If
                                End If
                       Else
                            'no data is in -5 file
'                           Print #fileresults, dat$, difsplit - difCompare, dysplit, difsplit, SpeedWind, CloudCover
                           End If
                           
                       End If
                       
              Else 'weather file or druk file is not in sync, so skip
              
                'determine if got here because of hole in weather data
                daynuma = DayNumber(yl, aamon, aaday)
                daynumb = DayNumber(yl, bbmon, bbday)
                
                If daynuma < daynumb And aayear = bbyear Then 'weather file is missing some entries, so iterate the druk compare file to match the hole
                  'increment druk compare file's entry
                  If EOF(filedrukin) Then
                     Exit Do
                     End If
                  Line Input #filedrukin, doclin$
                  InputSplit = Split(doclin$, ",")
                  dat$ = InputSplit(0)
                  If InStr(dat$, """") Then dat$ = Mid$(dat$, 2, 10)
                  dysplit = Val(InputSplit(1))
                  difsplit = Val(InputSplit(2))
                  'interpolate dif in compare file via the dysplit and determine if refraction is greater
                  'if refraction is greater than expected, then look for corresponding date in weather file
                  'and determine the wind and weather info and record to the results file
                  nextday = Fix(dysplit) + 1
                  If nextday >= 2 Then
                     difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(nextday - 2)) + DifYear(nextday - 2)
                  Else 'daysplit is lest then 1, so wrap to dec 31
                     difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(364)) + DifYear(364)
                     End If
                
                  'finish comparison to weather entry
                  GoTo comp750
                       
                Else
                  'increment to next line in weather file
                  GoTo comp500
                  End If
               End If
               
           Loop
comp900:
           Close #filedrukin
       Next i
       
       Close #fileresults
       Close #fileweatherin
    ElseIf optAveTK.Value = True Then
'        FileInName = App.Path & "\druk-at-combined-sorted.csv"
        End If
'    Open FileInName For Input As #filein
'    filecalc = FreeFile
'    Open FileCalcName For Input As #filecalc
'    Do Until EOF(filein)
'        Input #filein, dy, dif
'        'find corresponding calculated difference, and determine the difference between them
'        'if the calculated value is higher, then take the difference,
'        'then open the weather data, search for the date and determine if there is wind data and cloud data
'        'only include those days where the wind is less than 5 mph and there are no clouds at 8:00 and 20:00 the night before
'    Loop
'    Close #filein
ElseIf optGolan.Value = True Then
    FileWeather = App.Path & "\Jerusalem-sondes-Druk\" & "ims_data - Jerusalem 2009-2013-6.csv"
    If optMinTK.Value = True Then
        FileInNameGolan(0) = "NeveYaakov_compare.001"
        FileInNameGolan(1) = "NeveYaakov_compare.002"
        FileInNameGolan(2) = "NeveYaakov_compare.003"
        FileInNameGolan(3) = "NeveYaakov_compare.004"
        FileInNameGolan(4) = "NeveYaakov_compare.005"
        FileCalcName = App.Path & "\Golan_NO_mt_2009.csv" 'no adhoc sunrise fix of subtracting 15 seconds
        
        filecalcin = FreeFile
        Open FileCalcName For Input As #filecalcin
        numdy = 0
        Do Until EOF(filecalcin)
            Input #filecalcin, dy, DifYear(numdy)
            numdy = numdy + 1
        Loop
        Close #filecalcin
        
        fileweatherin = FreeFile
        Open FileWeather For Input As #fileweatherin
        Line Input #fileweatherin, doclin$ 'read doc line
        
        
        FileWeatherResults = App.Path & "\Golan-mt-weather-compare-16.csv"
        fileresults = FreeFile
        Open FileWeatherResults For Output As #fileresults
        
        For i = 0 To 4
           yr = 2009 + i
           'determine if it is leap year
           yl = DaysinYear(yr)
           
           filedrukin = FreeFile
           Open App.Path & "\" & FileInNameGolan(i) For Input As #filedrukin
           Do Until EOF(filedrukin)
comp1250:
              Line Input #filedrukin, doclin$
              InputSplit = Split(doclin$, ",")
              dat$ = InputSplit(0)
              If InStr(dat$, """") Then dat$ = Mid$(dat$, 2, 10)
              dysplit = Val(InputSplit(1))
              difsplit = Val(InputSplit(2))
              'interpolate dif in compare file via the dysplit and determine if refraction is greater
              'if refraction is greater than expected, then look for corresponding date in weather file
              'and determine the wind and weather info and record to the results file
              nextday = Fix(dysplit) + 1
              If nextday >= 2 Then
                 difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(nextday - 2)) + DifYear(nextday - 2)
              Else 'daysplit is lest then 1, so wrap to dec 31
                 difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(364)) + DifYear(364)
                 End If
                 
              'find corresponding date in weather file and
comp1500:        If EOF(fileweatherin) Then
                   Exit Do
                   End If
                Line Input #fileweatherin, doclin$
                SplitWeather = Split(doclin$, ",")
                TimeDay = SplitWeather(1)
                If InStr(TimeDay, """") Then TimeDay = Mid$(TimeDay, 2, Len(TimeDay) - 1)
                
                bb$ = SplitWeather(0) 'Format(SplitWeather(0), "dd/mm/yyyy")
                If InStr(bb, """") Then bb$ = Mid$(bb$, 2, 10)

                SplitDate = Split(bb$, "/")
                bbday = Val(SplitDate(0))
                bbmon = Val(SplitDate(1))
                bbyear = Val(SplitDate(2))
                
                If bbday = 12 And bbmon = 5 Then
                   ccc = 1
                   End If
                
'                bbday$ = Mid$(bb$, 4, 2)
'                bbmon$ = Mid$(bb$, 1, 2)
'                bbyear$ = Mid$(bb$, 7, 4)
                
'                aa$ = Format(dat$, "dd/mm/yyyy")
'                If dat$ = "02/06/1988" Then
'                   ccc = 1
'                   End If
'
'                If bb$ = "1/6/1988" Then
'                   ccc = 1
'                   End If
comp1750:
                SplitDate = Split(dat$, "/")
                aaday = Val(SplitDate(0))
                aamon = Val(SplitDate(1))
                aayear = Val(SplitDate(2))
                
'                If (bbyear = 1989) Then
'                   ccc = 1
'                   End If
         
'                aaday$ = Mid$(aa$, 4, 2)
'                aamon$ = Mid$(aa$, 1, 2)
'                aayear$ = Mid$(aa$, 7, 4)
'                If aaday$ = bbday$ And aamon$ = bbmon$ And aayear$ = bbyear$ Then
                If (aaday = bbday And aamon = bbmon And aayear = bbyear) And (InStr(TimeDay, "5:00") Or InStr(TimeDay, "8:00")) Then
                      '-5 is days where there are no weather data
                      '-6 is earlier sunrises with relaxed conditions wind speed <=5 and cloudcover <= 1
                      '-7 is above with later sunrises
                      '-7 analyses all days that have weather associated with them by printing their wind speed (mph), and cloud cover
                      '-8 strict zero wind and cloud conditions but no filtering of whether above or below calculated values
                      '-9 fix(wind) <= 1, fix(clouds) <=1 and ""
                      '-10 fix(wind) <= 2, fix(clouds) <=1 and ""
                      '-11 fix(wind) <= 3, fix(clouds) <=1 and ""
                      '-12 fix(wind) <= 4, fix(clouds) <=1 and ""
                      '-13 fix(wind) <= 5, fix(clouds) <=1 and ""
                      '-14 fix(wind) <= 5, fix(clouds) <=3 and ""
                      '-15 fix(wind) <= 10 (basically no filter), fix(clouds) <= 0 and ""
                      '-16 any wind, fix(clouds) <= 1 and ""
                
                    If difsplit < difCompare Or difsplit >= difCompare Then
                      'look for corresponding weather data for this date
                
                       'record to results if wind is less then 5 mph and sky is mostly clear, i.e., total cloud cover <=1
                       If SplitWeather(6) <> "-" And SplitWeather(10) <> "-" Then
                            SpeedWind = Val(SplitWeather(6)) * 2.23694 'mph
                            CloudCover = Val(SplitWeather(10))
                            
                            If nofilter Then
                                Print #fileresults, dat$, difsplit - difCompare, dysplit, difsplit, SpeedWind, CloudCover
                            Else
                                If Fix(CloudCover) <= 0 Then 'Fix(SpeedWind) <= 10 And Fix(CloudCover) <= 1 Then 'record this data
                                   Print #fileresults, dat$, difsplit - difCompare, dysplit, difsplit, SpeedWind, CloudCover
                                   End If
                                End If
                       Else
                            'no data is in -5 file
'                           Print #fileresults, dat$, difsplit - difCompare, dysplit, difsplit, SpeedWind, CloudCover
                           End If
                           
                       End If
                       
              Else 'weather file or druk file is not in sync, so skip
              
                'determine if got here because of hole in weather data
                daynuma = DayNumber(yl, aamon, aaday)
                daynumb = DayNumber(yl, bbmon, bbday)
                
                If daynuma < daynumb And aayear = bbyear Then 'weather file is missing some entries, so iterate the druk compare file to match the hole
                  'increment druk compare file's entry
                  If EOF(filedrukin) Then
                     Exit Do
                     End If
                  Line Input #filedrukin, doclin$
                  InputSplit = Split(doclin$, ",")
                  dat$ = InputSplit(0)
                  If InStr(dat$, """") Then dat$ = Mid$(dat$, 2, 10)
                  dysplit = Val(InputSplit(1))
                  difsplit = Val(InputSplit(2))
                  'interpolate dif in compare file via the dysplit and determine if refraction is greater
                  'if refraction is greater than expected, then look for corresponding date in weather file
                  'and determine the wind and weather info and record to the results file
                  nextday = Fix(dysplit) + 1
                  If nextday >= 2 Then
                     difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(nextday - 2)) + DifYear(nextday - 2)
                  Else 'daysplit is lest then 1, so wrap to dec 31
                     difCompare = (dysplit - Fix(dysplit)) * (DifYear(nextday - 1) - DifYear(364)) + DifYear(364)
                     End If
                
                  'finish comparison to weather entry
                  GoTo comp1750
                       
                Else
                  'increment to next line in weather file
                  GoTo comp1500
                  End If
               End If
               
           Loop
comp1900:
           Close #filedrukin
       Next i
       
       Close #fileresults
       Close #fileweatherin
    ElseIf optAveTK.Value = True Then
'        FileInName = App.Path & "\golan-at-combined-sorted.csv"
        End If

   End If
   
End Sub

Private Sub cmdJoin_Click()
Dim FileOutName As String, fileout As Integer, filein As Integer
Dim FileInName As String, list() As Double, ListItems() As String
Dim NameOut$(13)

   On Error GoTo cmdJoin_Click_Error
   
   lstSort.Clear

weather% = 6          '=0 for Menat's summer weather
                      '=1 for Menat's winter weather
                      '=2 for Almanac's weather
                      '=3 for mix of weather
                      '=4 for navigational sunrise
                      '=5 for using van der Werf atmospheres
                      '=6 for using vdw calculations of the visible sunrise for each place and
                      '   extrapolating all the necessary information from those files

fileout = FreeFile
If optDruk.Value = True Then
    FileOutName = App.Path & "\druk-combined"
    If weather% = 0 Then
        FileOutName = FileOutName & "-0.csv"
        name1$ = "druksum.001"
        name2$ = "druksum.002"
        name3$ = "druksum.003"
        name4$ = "druksum.004"
        name5$ = "druksum.005"
        name6$ = "druksum.006"
        name7$ = "druksum.007"
        name8$ = "druksum.008"
        name9$ = "druksum.009"
        name10$ = "druksum.010"
        name11$ = "druksum.011"
        name12$ = "druksum.012"
    ElseIf weather% = 1 Then
        FileOutName = FileOutName & "-1.csv"
        name1$ = "drukwin.001"
        name2$ = "drukwin.002"
        name3$ = "drukwin.003"
        name4$ = "drukwin.004"
        name5$ = "drukwin.005"
        name6$ = "drukwin.006"
        name7$ = "drukwin.007"
        name8$ = "drukwin.008"
        name9$ = "drukwin.009"
        name10$ = "drukwin.010"
        name11$ = "drukwin.011"
        name12$ = "drukwin.012"
    ElseIf weather% = 3 Then
        FileOutName = FileOutName & "-3.csv"
        name1$ = "drukmix.001"
        name2$ = "drukmix.002"
        name3$ = "drukmix.003"
        name4$ = "drukmix.004"
        name5$ = "drukmix.005"
        name6$ = "drukmix.006"
        name7$ = "drukmix.007"
        name8$ = "drukmix.008"
        name9$ = "drukmix.009"
        name10$ = "drukmix.010"
        name11$ = "drukmix.011"
        name12$ = "drukmix.012"
    ElseIf weather% = 5 Then
        FileOutName = FileOutName & "-5.csv"
        name1$ = "drukvdw.001"
        name2$ = "drukvdw.002"
        name3$ = "drukvdw.003"
        name4$ = "drukvdw.004"
        name5$ = "drukvdw.005"
        name6$ = "drukvdw.006"
        name7$ = "drukvdw.007"
        name8$ = "drukvdw.008"
        name9$ = "drukvdw.009"
        name10$ = "drukvdw.010"
        name11$ = "drukvdw.011"
        name12$ = "drukvdw.012"
     ElseIf weather% = 6 Then
        FileOutName = FileOutName & "-6.csv"
        name1$ = "drukvdw.085"
        name2$ = "drukvdw.086"
        name3$ = "drukvdw.087"
        name4$ = "drukvdw.088"
        name5$ = "drukvdw.089"
        name6$ = "drukvdw.090"
        name7$ = "drukvdw.091"
        name8$ = "drukvdw.092"
        name9$ = "drukvdw.093"
        name10$ = "drukvdw.094"
        name11$ = "drukvdw.095"
        name12$ = "drukvdw.096"
        End If
        
     Open FileOutName For Append As #fileout
     filein = FreeFile
     Open App.Path & "\" & name1$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     filein = FreeFile
     Open App.Path & "\" & name2$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     filein = FreeFile
     Open App.Path & "\" & name3$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     filein = FreeFile
     Open App.Path & "\" & name4$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name5$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name6$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name7$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name8$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name9$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name10$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name11$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
      filein = FreeFile
     Open App.Path & "\" & name12$ For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     Close #fileout
     
 ElseIf optGolan.Value = True Then
    If weather = 5 Then
        FileOutName = App.Path & "\Golan-combined-5.csv"
        NameOut$(0) = "NeveYaakov.001"
        NameOut$(1) = "NeveYaakov.002"
        NameOut$(2) = "NeveYaakov.003"
        NameOut$(3) = "NeveYaakov.004"
        NameOut$(4) = "NeveYaakov.005"
    ElseIf weather = 6 Then
        FileOutName = App.Path & "\Golan-combined-6.csv"
        NameOut$(0) = "NeveYaakov.009"
        NameOut$(1) = "NeveYaakov.010"
        NameOut$(2) = "NeveYaakov.011"
        NameOut$(3) = "NeveYaakov.012"
        NameOut$(4) = "NeveYaakov.013"
       End If
    
    fileout = FreeFile
     Open FileOutName For Append As #fileout
     filein = FreeFile
     Open App.Path & "\" & NameOut$(0) For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     Open App.Path & "\" & NameOut$(1) For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     Open App.Path & "\" & NameOut$(2) For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     Open App.Path & "\" & NameOut$(3) For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     Open App.Path & "\" & NameOut$(4) For Input As filein
     Do Until EOF(filein)
        Line Input #filein, doclin$
        Print #fileout, doclin$
        lstSort.AddItem doclin$
     Loop
     Close #filein
     Close #fileout
     
 ElseIf optZuriel.Value = True Then
    If weather = 5 Then
        FileOutName = App.Path & "\Zuriel-combined-5.csv"
        NameOut$(0) = "ZurielBB.061"
        NameOut$(1) = "ZurielBB.062"
        NameOut$(2) = "ZurielBB.064"
        NameOut$(3) = "ZurielBB.065"
        NameOut$(4) = "ZurielBB.066"
        NameOut$(5) = "ZurielBB.077"
        NameOut$(6) = "ZurielBB.079"
        NameOut$(7) = "ZurielBB.080"
        NameOut$(8) = "ZurielBB.083"
        NameOut$(9) = "ZurielBB.088"
        NameOut$(10) = "ZurielBB.089"
        NameOut$(11) = "ZurielBB.090"
        NameOut$(12) = "ZurielBB.091"
        NameOut$(13) = "ZurielBB.092"
    ElseIf weather = 6 Then
        FileOutName = App.Path & "\Zuriel-combined-6.csv"
        NameOut$(0) = "ZurielBB.061"
        NameOut$(1) = "ZurielBB.062"
        NameOut$(2) = "ZurielBB.064"
        NameOut$(3) = "ZurielBB.065"
        NameOut$(4) = "ZurielBB.066"
        NameOut$(5) = "ZurielBB.077"
        NameOut$(6) = "ZurielBB.079"
        NameOut$(7) = "ZurielBB.080"
        NameOut$(8) = "ZurielBB.083"
        NameOut$(9) = "ZurielBB.088"
        NameOut$(10) = "ZurielBB.089"
        NameOut$(11) = "ZurielBB.090"
        NameOut$(12) = "ZurielBB.091"
        NameOut$(13) = "ZurielBB.092"
       End If
    
    For i = 0 To 13
        fileout = FreeFile
         Open FileOutName For Append As #fileout
         filein = FreeFile
         Open App.Path & "\" & NameOut$(i) For Input As filein
         Do Until EOF(filein)
            Line Input #filein, doclin$
            Print #fileout, doclin$
            lstSort.AddItem doclin$
         Loop
         Close #filein
         Close #fileout
     Next i
     
    Close
    
    End If
 
 ReDim Preserve list(1, lstSort.ListCount - 1)
 
 'now add the almost sorted list to an array, and use a bubble sort routine to finish the job
 For i = 1 To lstSort.ListCount
    ListItems = Split(lstSort.list(i - 1), ",")
    list(0, i - 1) = Val(ListItems(0))
    list(1, i - 1) = Val(ListItems(1))
 Next i
 
 'now sort the file and output to the sorted file
Screen.MousePointer = vbHourglass
Call BubbleSort(list, 0, lstSort.ListCount - 1)

filein = FreeFile
If optDruk.Value = True Then
    If optMinTK.Value = True Then
        FileInName = App.Path & "\druk-mt-combined-sorted-new.csv"
    ElseIf optAveTK.Value = True Then
        FileInName = App.Path & "\druk-at-combined-sorted-new.csv"
        End If
    Open FileInName For Output As #filein
    For i = 1 To lstSort.ListCount
       Write #filein, list(0, i - 1), list(1, i - 1) 'lstSort.list(i - 1)
    Next i
    Close #filein
ElseIf optGolan.Value = True Then
    If optMinTK.Value = True Then
        FileInName = App.Path & "\golan-mt-combined-sorted-new-2.csv"
    ElseIf optAveTK.Value = True Then
        FileInName = App.Path & "\golan-at-combined-sorted-new.csv"
        End If
    Open FileInName For Output As #filein
    For i = 1 To lstSort.ListCount
       Write #filein, list(0, i - 1), list(1, i - 1) 'lstSort.list(i - 1)
    Next i
    Close #filein
ElseIf optZuriel.Value = True Then
    If optMinTK.Value = True Then
        FileInName = App.Path & "\zuriel-mt-combined-sorted-new.csv"
    ElseIf optAveTK.Value = True Then
        FileInName = App.Path & "\zuriel-at-combined-sorted-new.csv"
        End If
    Open FileInName For Output As #filein
    For i = 1 To lstSort.ListCount
       Write #filein, list(0, i - 1), list(1, i - 1) 'lstSort.list(i - 1)
    Next i
    Close #filein
    End If
Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Sub

cmdJoin_Click_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdJoin_Click of Form Drukfrm"
End Sub

Private Sub cmdLoop_Click()
   HeightLoop = True
   Initialize = False
   Call HeightLoopRun
End Sub

Private Sub cmdPlotFiles_Click()
   'create Plot save file, one for all the winter dates, another for the summer dates
   'Druk-all-dates.csv
   Dim FileNameIn As String, SplitDoc() As String, doclin$, filein%, DateSonde$
   Dim SumPlotfile As String, WinPlotfile As String, Dif As Double
   Dim SondeFileName As String, DifThresholdW As Double, DifThresholdS As Double
   Dim UseThresholds As Boolean
   
   DifThresholdW = -0.3
   DifThresholdS = -0.1
   UseThresholds = False
   
   SumPlotfile = App.Path & "\Sum-PlotFiles.txt"
   WinPlotfile = App.Path & "\Win-PlotFiles.txt"
   
   filesum = FreeFile
   Open SumPlotfile For Output As #filesum
   filewin = FreeFile
   Open WinPlotfile For Output As #filewin
   Print #filesum, "This file is used by Plot. Don't erase it!"
   Print #filesum, "X-values"
   Print #filesum, "Y-values"
   Print #filesum, """"""
   Print #filewin, "This file is used by Plot. Don't erase it!"
   Print #filewin, "X-values"
   Print #filewin, "Y-values"
   Print #filewin, """"""
   
   FileNameIn = App.Path & "\Druk-all-dates.csv"
   filein% = FreeFile
   Open FileNameIn For Input As #filein%
   Do Until EOF(filein%)
      Line Input #filein%, doclin$
      SplitDoc = Split(doclin$, ",")
      DateSonde$ = Mid$(SplitDoc(0), 2, Len(SplitDoc(0)) - 2)
      Dif = Val(SplitDoc(7))
      
      SondeFileName = App.Path & "\" & DateSonde$ & "-sondes.txt"
      
      If InStr(DateSonde$, "Jan") Or InStr(DateSonde$, "Feb") Or InStr(DateSonde$, "Nov") Or InStr(datemode$, "Dec") Then
         If Not UseThresholds Or Dif <= DifThresholdW Then
            Print #filewin, """ 2"","" 0"","" 0"",""1"",""0"",""1"",""0"",""" & SondeFileName & """,""none:none"",""3"""
            End If
      ElseIf InStr(DateSonde$, "May") Or InStr(DateSonde$, "Jun") Or InStr(DateSonde$, "Jul") Then
         If Not UseThresholds Or (Dif > -0.2 And Dif <= DifThresholdS) Then
            Print #filesum, """ 2"","" 0"","" 0"",""1"",""0"",""1"",""0"",""" & SondeFileName & """,""none:none"",""3"""
            End If
         End If
         
   Loop
   Close #filein%
   Close #filesum
   Close #filewin

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdRun_Click
' Author    : Dr-John-K-Hall
' Date      : 4/21/2020
' Purpose   : In this version, no calculations other than extrapolations:
'           azimuth values for the pl1 files corresponding to that year
'           and difference from astronomical from the calculated visible times and diferences to ast from the _year fiels
'---------------------------------------------------------------------------------------
'
Private Sub cmdRun_Click()
'this file converts druks file into a sequential file of day,time
'it is meant to average together several years of observation
'so the program calculates the dy of each days noon starting from
'the year after a leap year
'
Dim sumref(2, 500), winref(2, 500)
Dim air As Double, dy1 As Double, lt As Double, lg As Double
Dim MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer, TK As Double
Dim vbweps(6) As Double, vdwref(6) As Double, ier As Integer
Dim VDWSF As Double, VDWALT As Double, lnhgt As Double, nyr_increment As Integer
Dim DrukTables%, SplitStr() As String, comparemode As Integer
Dim NameOut$(13), DateStr As String, TimeStr As String
Dim NameZuriel As String, yl As Integer, dy As Integer, yr As Integer
Dim FindHgtSondeRefCalc$, FindZeroSondeRefCalc$
Dim HgtSonde As Boolean, ZeroHgtSonde As Boolean
Dim RefHgtSonde As Double, RefZeroHgtSonde As Double
Dim va1 As Double, ref1 As Double, va0 As Double, ref0 As Double
Dim azi0 As Double, azi1 As Double, viewang As Double
Dim RefVDWValHgt As Double, RefVDWValZeroHgt As Double
'Dim FileInCom As String, FileAziInBnei(4) As String
'Dim filazibnei(4) As Integer, filecom As Integer, pos%, FileNamePl$, docaziZ$(4), NumAzi As Integer

   On Error GoTo cmdRun_Click_Error
   
pi = 4# * Atn(1)  '3.1415927
pi2 = 2 * pi
ch = 57.29578 / 15#   'conv rad to hr
cd = pi / 180# '1.74532927777778E-02 'conv deg to rad
hr = 60


weather% = 6          '=0 for Menat's summer weather
                      '=1 for Menat's winter weather
                      '=2 for Almanac's weather
                      '=3 for mix of weather
                      '=4 for navigational sunrise
                      '=5 for using van der Werf atmospheres
                      '=6 for using vdw calculations of the visible sunrise for each place and
                      '   extrapolating all the necessary information from those files
DrukTables% = 1       '=0 for old ast.constants (mistake?)
                      '=1 based on 1996 Ast Almanac constants used by all the other progarms
                      
comparemode = 0       '=0 for not searching for dependence on wind and clouds
                      '=1 for searching for above
                      
qualityobservation = 0 '=0 for not excluding observations marked by B,M,K
                       '=1 to exclude those observations
                      
If weather <> 5 And weather <> 6 Then
    Open "c:\for\menatsum.ref" For Input As #1
    Input #1, sumrefo
    For i% = 1 To 500
       Input #1, a, sumref(1, i%), b, c, sumref(2, i%)
    Next i%
    Close #1
    Open "c:\for\menatwin.ref" For Input As #1
    Input #1, winrefo
    For i% = 1 To 500
       Input #1, a, winref(1, i%), b, c, winref(2, i%)
    Next i%
    Close #1
    
ElseIf weather = 5 Then

    'vdW dip angle vs height polynomial fit coefficients
    vbweps(1) = 2.77346593151086
    vbweps(2) = 0.497348466526589
    vbweps(3) = 2.53874620975453E-03
    vbweps(4) = 6.75587054940366E-04
    vbweps(5) = 3.94973974451576E-05
    
    'vdW atmospheric refraction vs height polynomial fit coefficients
    vdwref(1) = 1.16577538442405
    vdwref(2) = 0.468149166683532
    vdwref(3) = -0.019176833246687
    vdwref(4) = -4.8345814464145E-03
    vdwref(5) = -4.90660400743218E-04
    vdwref(6) = -1.60099622077352E-05
    
    End If

Screen.MousePointer = vbHourglass
If optDruk Then
    Name0$ = "druk.001"
    names$ = "druk9396.dat"
ElseIf optGolan Then
    Dim namein$(5)
    namein$(0) = App.Path & "\" & "Golan_2009.csv"
    namein$(1) = App.Path & "\" & "Golan_2010.csv"
    namein$(2) = App.Path & "\" & "Golan_2011.csv"
    namein$(3) = App.Path & "\" & "Golan_2012.csv"
    namein$(4) = App.Path & "\" & "Golan_2013.csv"
ElseIf optZuriel Then
    NameZuriel = "Zuriel_Bnei_Brqk_61-92.csv"
    End If
    
Dim tim$(12)

If weather% = 0 Then '1
    name1$ = "druksum.001"
    name2$ = "druksum.002"
    name3$ = "druksum.003"
    name4$ = "druksum.004"
    name5$ = "druksum.005"
    name6$ = "druksum.006"
    name7$ = "druksum.007"
    name8$ = "druksum.008"
    name9$ = "druksum.009"
    name10$ = "druksum.010"
    name11$ = "druksum.011"
    name12$ = "druksum.012"
ElseIf weather% = 1 Then
    name1$ = "drukwin.001"
    name2$ = "drukwin.002"
    name3$ = "drukwin.003"
    name4$ = "drukwin.004"
    name5$ = "drukwin.005"
    name6$ = "drukwin.006"
    name7$ = "drukwin.007"
    name8$ = "drukwin.008"
    name9$ = "drukwin.009"
    name10$ = "drukwin.010"
    name11$ = "drukwin.011"
    name12$ = "drukwin.012"
ElseIf weather% = 3 Then
    name1$ = "drukmix.001"
    name2$ = "drukmix.002"
    name3$ = "drukmix.003"
    name4$ = "drukmix.004"
    name5$ = "drukmix.005"
    name6$ = "drukmix.006"
    name7$ = "drukmix.007"
    name8$ = "drukmix.008"
    name9$ = "drukmix.009"
    name10$ = "drukmix.010"
    name11$ = "drukmix.011"
    name12$ = "drukmix.012"
ElseIf weather% = 5 Then '1
    If optDruk.Value = True Then '2
        If comparemode = 0 Then '3
            name1$ = "drukvdw.001"
            name2$ = "drukvdw.002"
            name3$ = "drukvdw.003"
            name4$ = "drukvdw.004"
            name5$ = "drukvdw.005"
            name6$ = "drukvdw.006"
            name7$ = "drukvdw.007"
            name8$ = "drukvdw.008"
            name9$ = "drukvdw.009"
            name10$ = "drukvdw.010"
            name11$ = "drukvdw.011"
            name12$ = "drukvdw.012"
         ElseIf comparemode = 1 Then '3
            name1$ = "drukvdw_compare.001"
            name2$ = "drukvdw_compare.002"
            name3$ = "drukvdw_compare.003"
            name4$ = "drukvdw_compare.004"
            name5$ = "drukvdw_compare.005"
            name6$ = "drukvdw_compare.006"
            name7$ = "drukvdw_compare.007"
            name8$ = "drukvdw_compare.008"
            name9$ = "drukvdw_compare.009"
            name10$ = "drukvdw_compare.010"
            name11$ = "drukvdw_compare.011"
            name12$ = "drukvdw_compare.012"
            End If '3
    ElseIf optGolan.Value = True Then '2
'       Dim nameout$(5)
       If comparemode = 0 Then '3
            NameOut$(0) = App.Path & "\" & "NeveYaakov.001"
            NameOut$(1) = App.Path & "\" & "NeveYaakov.002"
            NameOut$(2) = App.Path & "\" & "NeveYaakov.003"
            NameOut$(3) = App.Path & "\" & "NeveYaakov.004"
            NameOut$(4) = App.Path & "\" & "NeveYaakov.005"
       ElseIf comparemode = 1 Then '3
            NameOut$(0) = App.Path & "\" & "NeveYaakov_compare.001"
            NameOut$(1) = App.Path & "\" & "NeveYaakov_compare.002"
            NameOut$(2) = App.Path & "\" & "NeveYaakov_compare.003"
            NameOut$(3) = App.Path & "\" & "NeveYaakov_compare.004"
            NameOut$(4) = App.Path & "\" & "NeveYaakov_compare.005"
          End If '3
    ElseIf optZuriel.Value = True Then '2
'       Dim nameout$(10)
       If comparemode = 0 Then '3
          NameOut$(0) = App.Path & "\" & "ZurielBB.061"
          NameOut$(1) = App.Path & "\" & "ZurielBB.062"
          NameOut$(2) = App.Path & "\" & "ZurielBB.064"
          NameOut$(3) = App.Path & "\" & "ZurielBB.065"
          NameOut$(4) = App.Path & "\" & "ZurielBB.066"
          NameOut$(5) = App.Path & "\" & "ZurielBB.077"
          NameOut$(6) = App.Path & "\" & "ZurielBB.079"
          NameOut$(7) = App.Path & "\" & "ZurielBB.080"
          NameOut$(8) = App.Path & "\" & "ZurielBB.083"
          NameOut$(9) = App.Path & "\" & "ZurielBB.088"
          NameOut$(10) = App.Path & "\" & "ZurielBB.089"
          NameOut$(11) = App.Path & "\" & "ZurielBB.090"
          NameOut$(12) = App.Path & "\" & "ZurielBB.091"
          NameOut$(13) = App.Path & "\" & "ZurielBB.092"

       ElseIf comparemode = 1 Then '3
          NameOut$(0) = App.Path & "\" & "ZurielBB_compare.061"
          NameOut$(1) = App.Path & "\" & "ZurielBB_compare.062"
          NameOut$(2) = App.Path & "\" & "ZurielBB_compare.064"
          NameOut$(3) = App.Path & "\" & "ZurielBB_compare.065"
          NameOut$(4) = App.Path & "\" & "ZurielBB_compare.066"
          NameOut$(5) = App.Path & "\" & "ZurielBB_compare.077"
          NameOut$(6) = App.Path & "\" & "ZurielBB_compare.079"
          NameOut$(7) = App.Path & "\" & "ZurielBB_compare.080"
          NameOut$(8) = App.Path & "\" & "ZurielBB_compare.083"
          NameOut$(9) = App.Path & "\" & "ZurielBB_compare.088"
          NameOut$(10) = App.Path & "\" & "ZurielBB_compare.089"
          NameOut$(11) = App.Path & "\" & "ZurielBB_compare.090"
          NameOut$(12) = App.Path & "\" & "ZurielBB_compare.091"
          NameOut$(13) = App.Path & "\" & "ZurielBB_compare.092"
          End If '3
       End If '2
ElseIf weather = 6 Then '1
    If optDruk.Value = True Then
'        Dim nameout$(12)
        If comparemode = 0 Then
            NameOut$(0) = "drukvdw.085"
            NameOut$(1) = "drukvdw.086"
            NameOut$(2) = "drukvdw.087"
            NameOut$(3) = "drukvdw.088"
            NameOut$(4) = "drukvdw.089"
            NameOut$(5) = "drukvdw.090"
            NameOut$(6) = "drukvdw.091"
            NameOut$(7) = "drukvdw.092"
            NameOut$(8) = "drukvdw.093"
            NameOut$(9) = "drukvdw.094"
            NameOut$(10) = "drukvdw.095"
            NameOut$(11) = "drukvdw.096"
         ElseIf comparemode = 1 Then
            NameOut$(0) = "drukvdw_compare.085"
            NameOut$(1) = "drukvdw_compare.086"
            NameOut$(2) = "drukvdw_compare.087"
            NameOut$(3) = "drukvdw_compare.088"
            NameOut$(4) = "drukvdw_compare.089"
            NameOut$(5) = "drukvdw_compare.090"
            NameOut$(6) = "drukvdw_compare.091"
            NameOut$(7) = "drukvdw_compare.092"
            NameOut$(8) = "drukvdw_compare.093"
            NameOut$(9) = "drukvdw_compare.094"
            NameOut$(10) = "drukvdw_compare.095"
            NameOut$(11) = "drukvdw_compare.096"
            End If
    ElseIf optGolan.Value = True Then
'       Dim nameout$(5)
       If comparemode = 0 Then
            NameOut$(0) = App.Path & "\" & "NeveYaakov.009"
            NameOut$(1) = App.Path & "\" & "NeveYaakov.010"
            NameOut$(2) = App.Path & "\" & "NeveYaakov.011"
            NameOut$(3) = App.Path & "\" & "NeveYaakov.012"
            NameOut$(4) = App.Path & "\" & "NeveYaakov.013"
       ElseIf comparemode = 1 Then
            NameOut$(0) = App.Path & "\" & "NeveYaakov_compare.009"
            NameOut$(1) = App.Path & "\" & "NeveYaakov_compare.010"
            NameOut$(2) = App.Path & "\" & "NeveYaakov_compare.011"
            NameOut$(3) = App.Path & "\" & "NeveYaakov_compare.012"
            NameOut$(4) = App.Path & "\" & "NeveYaakov_compare.013"
          End If
    ElseIf optZuriel.Value = True Then
'       Dim nameout$(12)
       If comparemode = 0 Then
          NameOut$(0) = App.Path & "\" & "ZurielBB.061"
          NameOut$(1) = App.Path & "\" & "ZurielBB.062"
          NameOut$(2) = App.Path & "\" & "ZurielBB.064"
          NameOut$(3) = App.Path & "\" & "ZurielBB.065"
          NameOut$(4) = App.Path & "\" & "ZurielBB.066"
          NameOut$(5) = App.Path & "\" & "ZurielBB.077"
          NameOut$(6) = App.Path & "\" & "ZurielBB.079"
          NameOut$(7) = App.Path & "\" & "ZurielBB.080"
          NameOut$(8) = App.Path & "\" & "ZurielBB.083"
          NameOut$(9) = App.Path & "\" & "ZurielBB.088"
          NameOut$(10) = App.Path & "\" & "ZurielBB.089"
          NameOut$(11) = App.Path & "\" & "ZurielBB.090"
          NameOut$(12) = App.Path & "\" & "ZurielBB.091"
          NameOut$(13) = App.Path & "\" & "ZurielBB.092"

       ElseIf comparemode = 1 Then
          NameOut$(0) = App.Path & "\" & "ZurielBB_compare.061"
          NameOut$(1) = App.Path & "\" & "ZurielBB_compare.062"
          NameOut$(2) = App.Path & "\" & "ZurielBB_compare.064"
          NameOut$(3) = App.Path & "\" & "ZurielBB_compare.065"
          NameOut$(4) = App.Path & "\" & "ZurielBB_compare.066"
          NameOut$(5) = App.Path & "\" & "ZurielBB_compare.077"
          NameOut$(6) = App.Path & "\" & "ZurielBB_compare.079"
          NameOut$(7) = App.Path & "\" & "ZurielBB_compare.080"
          NameOut$(8) = App.Path & "\" & "ZurielBB_compare.083"
          NameOut$(9) = App.Path & "\" & "ZurielBB_compare.088"
          NameOut$(10) = App.Path & "\" & "ZurielBB_compare.089"
          NameOut$(11) = App.Path & "\" & "ZurielBB_compare.090"
          NameOut$(12) = App.Path & "\" & "ZurielBB_compare.091"
          NameOut$(13) = App.Path & "\" & "ZurielBB_compare.092"
          End If
       End If
    End If
    
If weather = 6 Then GoTo NewCalc
                      
P = 1013
T = 27
If optDruk.Value = True Then
    lg = -35.237435642287 '81333572129 '-35.238456 '5 'longitude at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
    lt = 31.748552568177 '8959288296 '31.749942 'latitude at at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
    hgt = 756.7 '754.9  'altitude of observer at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
    nyr_increment = 11
ElseIf optGolan.Value = True Then 'R' Israel Golan's observation point, i.e., the Magen Avrohom shul of N'vei Ya'akov, Jerusalem
    lg = -35.2472030237111 '173.510
    lt = 31.8424065056182 '138.840
    hgt = 684.3
    nyr_increment = 4
ElseIf optZuriel.Value = True Then 'R' Yosef Zuriel z"l's collection of observations in Bnei Brak during 1961 - 1992
    'fictitious observation point created from average of three places near the Ponevitch Yeshiva
    lt = 32.084319333
    lg = -34.83215883
    hgt = 52
    nyr_increment = 30
    End If
td = 2  'time difference from Greenwich England
A1 = 0.833 'angle depression under horizon for sunset/sunrise
ht = 0.0348 * Sqr(hgt)
lr = cd * lt
tf = td * 60 + 4 * lg
ac = cd * 0.9856003
mc = cd * 0.9856474
ec = cd * 1.915
e2c = cd * 0.02
ob1 = cd * (-0.00000036) 'change of ecliptic angle per day

'load minimum and average temperatures for this place
Call Temperatures(lt, -lg, MinT, AvgT, MaxT, ier)
If ier = -1 Then
   Screen.MousePointer = vbDefault
   Exit Sub
   End If
'
'DEF FNarsin(x) = Atn(x / Sqr(-x * x + 1))
'DEF FNarco(x) = -Atn(x / Sqr(-x * x + 1)) + pi / 2
'DEF FNms(x) = mp + mc * x
'DEF FNaas(x) = ap + ac * x
'DEF FNes = ms + ec * Sin(aas) + e2c * Sin(2 * aas)
'DEF FNha(x) = FNarco((-Tan(lr) * Tan(d)) + (Cos(x) / Cos(lr) / Cos(d))) * ch
'DEF FNfrsum(x) = (P / (T + 273)) * (0.1419 - 0.0073 * x + 0.00005 * x * x) / (1 + 0.3083 * x + 0.01011 * x * x)
'DEF FNfrwin(x) = (P / (T + 273)) * (0.1561 - 0.0082 * x + 0.00006 * x * x) / (1 + 0.3254 * x + 0.01086 * x * x)
'DEF FNref(x) = (P / (T + 273)) * (0.1594 + 0.0196 * x + 0.00002 * x * x) / (1 + 0.505 * x + 0.0845 * x * x)
'
'IF weather% = 0 THEN
'   'x = .-ht
'   'GOSUB refsum
'   'a1 = frsum
'   GOSUB refsum1
'   'a1 = FNfrsum(-ht)
'   'a1 = .5114545
'   'air = (90 + .2667 + a1 + ht) * cd
'   x = -.032 * SQR(hgt)
'   GOSUB refsum1
'   a1 = frsum
'   air = (90 + .2667 + a1 - x) * cd
'ELSEIF weather% = 1 THEN
'   'x = -ht
'   'GOSUB refwin
'   'a1 = frwin
'   'a1 = .5650201
'   'air = (90 + .2667 + a1 + ht) * cd
'   x = -.032 * SQR(hgt)
'   GOSUB refwin1
'   a1 = frwin
'   air = (90 + .2667 + a1 - x) * cd
n1% = Fix((hgt - 1) / 2 + 1)
n2% = n1% + 1
h1 = (n1% - 1) * 2 + 1

'If weather <> 5 Then 'Menat atmospheres
'    If weather% = 1 Then 'winter refraction
'       ref = ((winref(2, n2%) - winref(2, n1%)) / 2) * (hgt - h1) + winref(2, n1%)
'       eps = ((winref(1, n2%) - winref(1, n1%)) / 2) * (hgt - h1) + winref(1, n1%)
'       air = (90 + 0.2667) * cd + (eps + ref + winrefo) / 1000
'    ElseIf weather% = 0 Then 'summer refraction
'       ref = ((sumref(2, n2%) - sumref(2, n1%)) / 2) * (hgt - h1) + sumref(2, n1%)
'       eps = ((sumref(1, n2%) - sumref(1, n1%)) / 2) * (hgt - h1) + sumref(1, n1%)
'       air = (90 + 0.2667) * cd + (eps + ref + sumrefo) / 1000
'
'    ElseIf weather% = 2 Then
'       a1 = FNref(-ht)
'       air = (90 + 0.2667 + a1 + ht) * cd
'    ElseIf weather% = 4 Then
'       a1 = 0.833
'       air = (90 + a1 + ht) * cd  'total angle depression
'       End If
'ElseIf weather = 5 Then 'van der Werf atmosphere
'    'astronomical refraction calculations, hgt in meters
'    ref = 0#
'    eps = 0#
'    If (hgt > 0) Then lnhgt = Log(hgt * 0.001)
'    'calculate total atmospheric refraction from the observer's height
'    'to the horizon and then to the end of the atmosphere
'    'All refraction terms have units of mrad
'    If (hgt <= 0#) Then GoTo 690
'    ref = Exp(vdwref(1) + vdwref(2) * lnhgt + _
'         vdwref(3) * lnhgt * lnhgt + vdwref(4) * lnhgt * lnhgt * lnhgt + _
'         vdwref(5) * lnhgt * lnhgt * lnhgt * lnhgt + _
'         vdwref(6) * lnhgt * lnhgt * lnhgt * lnhgt * lnhgt)
'    eps = Exp(vbweps(1) + vbweps(2) * lnhgt + _
'         vbweps(3) * lnhgt * lnhgt + vbweps(4) * lnhgt * lnhgt * lnhgt + _
'         vbweps(5) * lnhgt * lnhgt * lnhgt * lnhgt)
'     'now add the all the contributions together due to the observer's height
'     'along with the value of atm. ref. from hgt<=0, where view angle, a1 = 0
'     'for this calculation, leave the refraction in units of mrad
'690: a1 = 0
'     air = 90# * cd + (eps + VDWSF * (ref + 9.56267125268496)) / 1000#
'     a1 = (ref + 9.56267125268496) / 1000#
'    'leave a1 in radians
'    a1 = Atn(Tan(a1) * VDWALT)
'    End If
       
 nyr = 0
    If optDruk.Value = True Then
        sumet = -0.2422 'debit of days at beginning of 1985 'i.e., zero started on 1984, each year
        '(then on fourth year of cycle one day is added to approx. make up the debit = Julean calendar)
        'i.e., length of sidiereal year is 365.2422 days (Koserver, p. 59)
    ElseIf optGolan.Value = True Then
        sumet = 0.2422
        End If
        
    Do Until nyr > nyr_increment
    nfound = 0
    If optDruk.Value = True Then
      yr = 1985 + nyr 'year for calculation (starts in 1985)
      If yr = 1985 Then Open App.Path & "\" & name1$ For Output As #2
      If yr = 1986 Then Open App.Path & "\" & name2$ For Output As #2
      If yr = 1987 Then Open App.Path & "\" & name3$ For Output As #2
      If yr = 1988 Then Open App.Path & "\" & name4$ For Output As #2
      If yr = 1989 Then Open App.Path & "\" & name5$ For Output As #2
      If yr = 1990 Then Open App.Path & "\" & name6$ For Output As #2
      If yr = 1991 Then Open App.Path & "\" & name7$ For Output As #2
      If yr = 1992 Then Open App.Path & "\" & name8$ For Output As #2
      If yr = 1993 Then Open App.Path & "\" & name9$ For Output As #2
      If yr = 1994 Then Open App.Path & "\" & name10$ For Output As #2
      If yr = 1995 Then Open App.Path & "\" & name11$ For Output As #2
      If yr = 1996 Then Open App.Path & "\" & name12$ For Output As #2
      
        If nyr <= 7 Then Open App.Path & "\" & Name0$ For Input As #1  'open druk.001 file
        If nyr >= 8 Then Open App.Path & "\" & names$ For Input As #1 'open druk9396.dat file
        If nyr <= 7 Then
           n% = 1
           Do Until n% > 9
              Line Input #1, chrr$   'skipping header lines
              n% = n% + 1
           Loop
        ElseIf nyr >= 8 Then
           For n% = 1 To 3
              Line Input #1, doclin$
           Next n%
           End If
           
      ElseIf optGolan.Value = True Then
         yr = 2009 + nyr
         freein% = 1
         freeout% = 2
         If yr = 2009 Then
            Open namein$(0) For Input As #freein%
            Open NameOut$(0) For Output As #freeout%
         ElseIf yr = 2010 Then
            Open namein$(1) For Input As #freein%
            Open NameOut$(1) For Output As #freeout%
         ElseIf yr = 2011 Then
            Open namein$(2) For Input As #freein%
            Open NameOut$(2) For Output As #freeout%
         ElseIf yr = 2012 Then
            Open namein$(3) For Input As #freein%
            Open NameOut$(3) For Output As #freeout%
         ElseIf yr = 2013 Then
            Open namein$(4) For Input As #freein%
            Open NameOut$(4) For Output As #freeout%
            End If
     ElseIf optZuriel.Value = True Then
         yr = 1961 + nyr
         freein% = 1
         freeout% = 2
         End If
      
      If DrukTables% = 0 Then
      
          ap = cd * 357.528 '356.65629 'mean anomaly for Jan 0, 1988 12:00
          mp = cd * 280.46  '279.3828 'mean longitude of sun Jan 0, 1988 12:00
          ob = cd * 23.440852 'ecliptic angle for Jan 0, 1988 12:00
          ap = ap + lg / 360 * ac
          mp = mp + lg / 360 * mc
          
          yd = yr - 1988
          If yd > 0 Then yf1 = 1 Else yf1 = 0
          yf = yd * 365 + yf1
          yf = yf + Fix((yd - yf1) / 4)
          yl = 365
          If yd Mod 4 = 0 Then yl = 366
          ob = ob + ob1 * yf
          mp = mp + yf * mc
500       If mp < 0 Then mp = mp + 2 * pi
510       If mp < 0 Then GoTo 500
520       If mp > 2 * pi Then mp = mp - 2 * pi
530       If mp > 2 * pi Then GoTo 520
540       ap = ap + yf * ac
    
550       If ap < 0 Then ap = ap + 2 * pi
560       If ap < 0 Then GoTo 550
570       If ap > 2 * pi Then ap = ap - 2 * pi
580       If ap > 2 * pi Then GoTo 570

      ElseIf DrukTables% = 1 Then
      
'diagnostics
'If yr = 1995 Then
'   ccc = 1
'   End If
      
            ac = cd * 0.9856003
            ap = cd * 357.528   'mean anomaly for Jan 0, 1996 12:00
            mc = cd * 0.9856474
            mp = cd * 280.461  'mean longitude of sun Jan 0, 1996 12:00
            ec = cd * 1.915
            e2c = cd * 0.02
            ob = cd * 23.439   'ecliptic angle for Jan 0, 1996 12:00
            ob1 = cd * (-0.0000004) 'change of ecliptic angle per day
            ap = ap - (td / 24) * ac 'compensate for our time zone
            mp = mp - (td / 24) * mc
            ob = ob - (td / 24) * ob1
            'calculate cumulative years since 1996
            yd = yr - 1996
            yf = 0
            If yd < 0 Then
               For iyr% = 1995 To yr Step -1
                  yrtst% = iyr%
                  yltst% = 365
                  If yrtst% Mod 4 = 0 Then yltst% = 366
                  If yrtst% Mod 4 = 0 And yrtst% Mod 100 = 0 And yrtst% Mod 400 <> 0 Then yltst% = 365
                  yf = yf - yltst%
               Next iyr%
            ElseIf yd >= 0 Then
               For iyr% = 1996 To yr - 1 Step 1
                  yrtst% = iyr%
                  yltst% = 365
                  If yrtst% Mod 4 = 0 Then yltst% = 366
                  If yrtst% Mod 4 = 0 And yrtst% Mod 100 = 0 And yrtst% Mod 400 <> 0 Then yltst% = 365
                  yf = yf + yltst%
               Next iyr%
               End If
            yl = 365
            If yr Mod 4 = 0 Then yl = 366
            If yr Mod 4 = 0 And yr Mod 100 = 0 And yr Mod 400 <> 0 Then yl = 365
            yf = yf - 1462 'number of days from J2000.0 (called "n" in Almanac)
            ob = ob + ob1 * yf
            mp = mp + yf * mc
600         If mp < 0 Then mp = mp + pi2
610         If mp < 0 Then GoTo 600
620         If mp > pi2 Then mp = mp - pi2
630         If mp > pi2 Then GoTo 620
640         ap = ap + yf * ac
650         If ap < 0 Then ap = ap + pi2
660         If ap < 0 Then GoTo 650
670         If ap > pi2 Then ap = ap - pi2
680         If ap > pi2 Then GoTo 670

           End If
           
      If yr = 1995 Then
         fil995% = FreeFile
         Open App.Path & "\1995-astron.txt" For Output As #fil995%
         write1995 = True
      Else
         write1995 = False
         End If
         
      dy = 1
      Do While dy < yl + 1
      
         If weather = 5 Then 'van der Werf atmosphere
            'determine the minimum and average temperature for this day for current place
            'use Meeus's forumula p. 66 to convert daynumber to month,
            'no need to interpolate between temepratures -- that is overkill
            k% = 2
            If (yl = 366) Then k% = 1
            M% = Int(9 * (k% + dy) / 275 + 0.98)
            TK = MaxT(M%) + 273.15
'            If optMinTK.Value = True Then
'                TK = MT(M%) + 273.15 'mean minimum temperature for this month in degrees Kelvin
'            ElseIf optAveTK.Value = True Then
'                TK = AT(M%) + 273.15
                End If
            
            'calculate van der Werf temperature scaling factor for refraction
            VDWSF = (288.15 / TK) ^ 1.7081
            'calculate van der Werf scaling factor for view angles
            VDWALT = (288.15 / TK) ^ 0.69
            
'            End If
         
         dy1 = dy
         ob = ob + ob1  'ecliptic angle changes by one day
         GoSub 1360
         ra = Atn(Cos(ob) * Tan(es))
         If ra < 0 Then ra = ra + pi
         df = ms - ra
         While Abs(df) > pi / 2
            df = df - Sgn(df) * pi
         Wend
         et = (df / cd) / 360  'equation of time (fract of day)
         If dy = 60 And yl = 366 Then
            sumet = -0.00066 + sumet + et + 0.99934 'add one whole day
         Else
            sumet = -0.00066 + sumet + et
            End If
         If weather% = 3 Then
           If dy <= 85 Or dy >= 297 Then 'winter refraction
             Ref = ((winref(2, n2%) - winref(2, n1%)) / 2) * (hgt - h1) + winref(2, n1%)
             eps = ((winref(1, n2%) - winref(1, n1%)) / 2) * (hgt - h1) + winref(1, n1%)
             air = (90 + 0.2667) * cd + (eps + Ref + winrefo) / 1000
           ElseIf dy > 85 And dy < 297 Then
             Ref = ((sumref(2, n2%) - sumref(2, n1%)) / 2) * (hgt - h1) + sumref(2, n1%)
             eps = ((sumref(1, n2%) - sumref(1, n1%)) / 2) * (hgt - h1) + sumref(1, n1%)
             air = 90 * cd + (eps + Ref + sumrefo) / 1000
             End If
         ElseIf weather = 5 Then
            'astronomical refraction calculations, hgt in meters
            Ref = 0#
            eps = 0#
            If (hgt > 0) Then lnhgt = Log(hgt * 0.001)
            'calculate total atmospheric refraction from the observer's height
            'to the horizon and then to the end of the atmosphere
            'All refraction terms have units of mrad
            If (hgt <= 0#) Then GoTo 790
            Ref = Exp(vdwref(1) + vdwref(2) * lnhgt + _
                 vdwref(3) * lnhgt * lnhgt + vdwref(4) * lnhgt * lnhgt * lnhgt + _
                 vdwref(5) * lnhgt * lnhgt * lnhgt * lnhgt + _
                 vdwref(6) * lnhgt * lnhgt * lnhgt * lnhgt * lnhgt)
            eps = Exp(vbweps(1) + vbweps(2) * lnhgt + _
                 vbweps(3) * lnhgt * lnhgt + vbweps(4) * lnhgt * lnhgt * lnhgt + _
                 vbweps(5) * lnhgt * lnhgt * lnhgt * lnhgt)
             'now add the all the contributions together due to the observer's height
             'along with the value of atm. ref. from hgt<=0, where view angle, a1 = 0
             'for this calculation, leave the refraction in units of mrad
790:         A1 = 0
             air = 90 * cd + (eps + VDWSF * (Ref + 9.56267125268496)) / 1000#
             'air = 90 * cd + (eps + ref + VDWSF * 9.56267125268496) / 1000#
             A1 = VDWSF * (Ref + 9.56267125268496) / 1000#
            'leave a1 in radians
             A1 = Atn(Tan(A1) * VDWALT)
             End If

'            x = -ht
'            IF dy <= 85 OR dy >= 297 THEN
'               GOSUB refwin2
'               a1 = frwin
'            ELSEIF dy > 85 AND dy < 297 THEN
'               GOSUB refsum2
'               a1 = frsum
'               END IF
'            air = (90 + .2667 + a1 + ht) * cd
'            End If
         et = df / cd * 4  'now calculate sunrise
         t6 = 720 + tf - et
         fdy! = FNarco(-Tan(d) * Tan(lr)) * ch / 24
         dy1 = dy - fdy!
         GoSub 1360
         air = air + 0.2667 * (1 - 0.017 * Cos(aas)) * cd 'change of size of sun due to elliptical orbit of earth
         sr1 = FNha(air)
         sr = sr1 * hr
         t3 = (t6 - sr) / hr
         If write1995 Then
            'convert t3 to time
            t3hr = Fix(t3)
            t3min = Fix((t3 - t3hr) * 60)
            t3sec = Fix((t3 - t3hr - t3min / 60) * 3600)
            Print #fil995%, dy, t3, Trim$(Str(t3hr)) & ":" & Format(Trim$(Str(t3min)), "0#") & ":" & Format(Trim$(Str$(t3sec)), "0#")
            End If
         If optDruk.Value = True Then
             If yr >= 1993 Then
                Line Input #1, doclin$
                tim$(9) = Mid$(doclin$, 25, 24)
                If qualityobservation And (InStr(tim$(9), "B") Or InStr(tim$(9), "M") Or InStr(tim$(9), "K")) Then GoTo 999 'skip questional dates
                tim$(10) = Mid$(doclin$, 49, 24)
                If qualityobservation And (InStr(tim$(10), "B") Or InStr(tim$(10), "M") Or InStr(tim$(10), "K")) Then GoTo 999  'skip questional dates
                tim$(11) = Mid$(doclin$, 73, 24)
                If qualityobservation And (InStr(tim$(11), "B") Or InStr(tim$(11), "M") Or InStr(tim$(11), "K")) Then GoTo 999 'skip questional dates
                tim$(12) = Mid$(doclin$, 97, 17)
                If qualityobservation And (InStr(tim$(12), "B") Or InStr(tim$(12), "M") Or InStr(tim$(12), "K")) Then GoTo 999 'skip questional dates
                If yr = 1996 Then GoTo 700
                If dy = 60 Then
                   Line Input #1, doclin$
                   tim$(9) = Mid$(doclin$, 25, 24)
                   If qualityobservation And (InStr(tim$(9), "B") Or InStr(tim$(9), "M") Or InStr(tim$(9), "K")) Then GoTo 999 'skip questional dates
                   tim$(10) = Mid$(doclin$, 49, 24)
                   If qualityobservation And (InStr(tim$(10), "B") Or InStr(tim$(10), "M") Or InStr(tim$(10), "K")) Then GoTo 999 'skip questional dates
                   tim$(11) = Mid$(doclin$, 73, 24)
                   If qualityobservation And (InStr(tim$(11), "B") Or InStr(tim$(11), "M") Or InStr(tim$(11), "K")) Then GoTo 999 'skip questional dates
                   tim$(12) = Mid$(doclin$, 97, 17)
                   If qualityobservation And (InStr(tim$(12), "B") Or InStr(tim$(12), "M") Or InStr(tim$(12), "K")) Then GoTo 999 'skip questional dates
                   End If
             Else
                Input #1, Days, dat$
                ntim = 1
                Do While ntim < 9
                   Input #1, tim$(ntim)
                   ntim = ntim + 1
                Loop
                If yr = 1988 Or yr = 1992 Then GoTo 700
                If dy = 60 Then  'then reread in order to skip Feb 29 entry
                   Input #1, Days, dat$
                   ntim = 1
                   Do While ntim < 9
                      Input #1, tim$(ntim)
                      ntim = ntim + 1
                   Loop
                   End If
                End If
700          ntim = yr - 1985 + 1
            If yr <= 1992 Then If Mid$(tim$(ntim), 1, 1) = "#" Then GoTo 999
            If yr >= 1993 Then If Mid$(tim$(ntim), 1, 1) = "," Then GoTo 999
            If qualityobservation And (InStr(tim$(ntim), "B") Or InStr(tim$(ntim), "M") Or InStr(tim$(ntim), "K")) Then GoTo 999 'skip questional dates
            nfound = nfound + 1
            'special search for certain date
'            If yr = 1987 And nfound = 7 Then
'               fileout = FreeFile
'               Open App.Path & "\compare-for-clouds-for-inversion-layers.txt" For Append As #fileout
'               Print #fileout, dat$ & "-1987-" & tim$(ntim)
'               Close #fileout
'               End If
'            If yr = 1988 And nfound = 9 Then
'               fileout = FreeFile
'               Open App.Path & "\compare-for-clouds-for-inversion-layers.txt" For Append As #fileout
'               Print #fileout, dat$ & "-1988-" & tim$(ntim)
'               Close #fileout
'               End If
'            If yr = 1992 And nfound = 7 Then
'               fileout = FreeFile
'               Open App.Path & "\compare-for-clouds-for-inversion-layers.txt" For Append As #fileout
'               Print #fileout, dat$ & "-1992-" & tim$(ntim)
'               Close #fileout
'               End If
'            If yr = 1996 And nfound = 5 Then
'               fileout = FreeFile
'               Open App.Path & "\compare-for-clouds-for-inversion-layers.txt" For Append As #fileout
'               Print #fileout, dat$ & "-1995-" & tim$(12)
'               Close #fileout
'               End If
            s1 = Val(Mid$(tim$(ntim), 1, 2))
            s2 = Val(Mid$(tim$(ntim), 4, 5))
            s3 = Val(Mid$(tim$(ntim), 7, 8))
            s1 = s1 + (s2 + s3 / 60) / 60
            'WRITE #2, dy + sumet, s1, t3, s1 - t3
            If (s1 - t3) * 60 < 7 Then
               If comparemode = 0 Then
                  Write #2, dy + sumet, (s1 - t3) * 60
               ElseIf comparemode = 1 Then
                  dat$ = DateAdd("d", Fix(dy), "31/Dec/" & Trim$(Str$(yr - 1))) 'date for weather comparisons goes according that year's calendar
                  Write #2, Format(dat$, "dd/mm/yyyy"), dy + sumet, (s1 - t3) * 60
                  End If
               End If
         ElseIf optGolan.Value = True Then
            Line Input #freein%, doclin$
            If doclin$ <> "" Then
                SplitStr = Split(doclin$, ",")
                ab = UBound(SplitStr)
                TimeStr = SplitStr(2)
                If Len(TimeStr) > 1 Then
                   TimeStr = SplitStr(2)
                   s1 = Val(Mid$(TimeStr, 1, 1))
                   s2 = Val(Mid$(TimeStr, 3, 4))
                   s3 = Val(Mid$(TimeStr, 6, 7))
                   s1 = s1 + (s2 + s3 / 60) / 60
                   If (s1 - t3) * 60 < 7 Then
                      If comparemode = 0 Then
                         Write #freeout%, dy + sumet, (s1 - t3) * 60
                      ElseIf comparemode = 1 Then
                         'Write #freeout%, SplitStr(1), dy + sumet, (s1 - t3) * 60
                          dat$ = DateAdd("d", Fix(dy), "31/Dec/" & Trim$(Str$(yr - 1))) 'date for weather comparisons goes according that year's calendar
                          Write #2, Format(dat$, "dd/mm/yyyy"), dy + sumet, (s1 - t3) * 60
                         End If
                       End If
                   End If
                Else
                   MsgBox "Blank line at line number: " & dy & " of file: " & namein$(nyr), vbOKOnly + vbCritical, "Error"
                   Exit Sub
                End If
            End If

999      dy = dy + 1
         Loop
         If optDruk.Value = True Then
            Close #1
            Close #2
            If write1995 Then Close #fil995%
         ElseIf optGolan.Value = True Then
            Close #freein%
            Close #freeout%
            End If
         nyr = nyr + 1   'go to next year of calculation
    Loop
    GoTo 9999
    
NewCalc:

If optDruk.Value = True Then
   myfile = Dir(App.Path & "\Druk-all-dates.csv")
   If myfile <> vbNullString Then
      Kill App.Path & "\Druk-all-dates.csv"
      End If
    End If
    
For i = 0 To UBound(NameOut$)
    'extract year from naemout file and determine pl1 and _year file names
    pos% = InStr(NameOut$(i), ".")
    yrext$ = Mid$(NameOut$(i), pos% + 2, 2)
    If optDruk.Value = True Then '1
    
       If i > 11 Then
          Close
          Screen.MousePointer = vbDefault
          Exit Sub
          End If
       
       yr = 1900 + Val(yrext$)
       
       yl = DaysinYear(Int(yr))
       
       nyr = i
       
       Open App.Path & "\" & NameOut$(i) For Output As #2
      
       If nyr <= 7 Then Open App.Path & "\" & Name0$ For Input As #1  'open druk.001 file
       If nyr >= 8 Then Open App.Path & "\" & names$ For Input As #1 'open druk9396.dat file
    
       If fileall = 0 Then
          fileall = FreeFile
          Open App.Path & "\Druk-all-dates.csv" For Append As #fileall
          End If
       
       If nyr <= 7 Then '2
          n% = 1
          Do Until n% > 9
             Line Input #1, chrr$   'skipping header lines
             n% = n% + 1
          Loop
       ElseIf nyr >= 8 Then '2
          For n% = 1 To 3
             Line Input #1, doclin$
          Next n%
          End If '2
          
       'now open the azimuth and the sunrise files for that year
'       RavD1996.pl1 'azi file
'       RavD1995 'zemanim file
        FileAziIn$ = App.Path & "\" & "RavD" & Trim$(Str$(yr)) & ".pl1"
        FileZmaIn$ = App.Path & "\" & "RavD" & Trim$(Str$(yr))
        fileazi = FreeFile
        Open FileAziIn$ For Input As #fileazi
        filezma = FreeFile
        Open FileZmaIn$ For Input As #filezma
       
       dy = 1
       Do While dy < yl + 1
        
        If yr >= 1993 Then '2
           Line Input #1, doclin$
           tim$(9) = Mid$(doclin$, 25, 24)
           If qualityobservation And (InStr(tim$(9), "B") Or InStr(tim$(9), "M") Or InStr(tim$(9), "K")) Then GoTo 1999 'skip questional dates
           tim$(10) = Mid$(doclin$, 49, 24)
           If qualityobservation And (InStr(tim$(10), "B") Or InStr(tim$(10), "M") Or InStr(tim$(10), "K")) Then GoTo 1999  'skip questional dates
           tim$(11) = Mid$(doclin$, 73, 24)
           If qualityobservation And (InStr(tim$(11), "B") Or InStr(tim$(11), "M") Or InStr(tim$(11), "K")) Then GoTo 1999 'skip questional dates
           tim$(12) = Mid$(doclin$, 97, 17)
           If qualityobservation And (InStr(tim$(12), "B") Or InStr(tim$(12), "M") Or InStr(tim$(12), "K")) Then GoTo 1999 'skip questional dates
           If yr = 1996 Then GoTo 1700
           If dy = 60 Then '3
              Line Input #1, doclin$
              tim$(9) = Mid$(doclin$, 25, 24)
              If qualityobservation And (InStr(tim$(9), "B") Or InStr(tim$(9), "M") Or InStr(tim$(9), "K")) Then GoTo 1999 'skip questional dates
              tim$(10) = Mid$(doclin$, 49, 24)
              If qualityobservation And (InStr(tim$(10), "B") Or InStr(tim$(10), "M") Or InStr(tim$(10), "K")) Then GoTo 1999 'skip questional dates
              tim$(11) = Mid$(doclin$, 73, 24)
              If qualityobservation And (InStr(tim$(11), "B") Or InStr(tim$(11), "M") Or InStr(tim$(11), "K")) Then GoTo 1999 'skip questional dates
              tim$(12) = Mid$(doclin$, 97, 17)
              If qualityobservation And (InStr(tim$(12), "B") Or InStr(tim$(12), "M") Or InStr(tim$(12), "K")) Then GoTo 1999 'skip questional dates
              End If '3
        Else '2
           Input #1, Days, dat$
           ntim = 1
           Do While ntim < 9
              Input #1, tim$(ntim)
              ntim = ntim + 1
           Loop
           If yr = 1988 Or yr = 1992 Then GoTo 1700
           If dy = 60 Then  'then reread in order to skip Feb 29 entry
              Input #1, Days, dat$
              ntim = 1
              Do While ntim < 9
                 Input #1, tim$(ntim)
                 ntim = ntim + 1
              Loop
              End If
           End If '2
           
1700       ntim = yr - 1985 + 1
            
            Line Input #fileazi, docazi$
            Line Input #filezma, doczma$
            'extract the azimuth and the vis sunrise time and the difference of time to the astr
            azimuth = Val(Mid$(docazi$, 40, 7))

            'find view angle corresponding to this azimuth
            fileprofile% = FreeFile
            Open App.Path & "\RavDrkTR.pr1" For Input As #fileprofile%
            Line Input #fileprofile%, docprof$ 'first header line
            Line Input #fileprofile%, docprof$ '2nd header line

            Do Until EOF(fileprofile%)
                Line Input #fileprofile%, docprof$
                SplitStr = Split(docprof$, ",")
                azi0 = Val(SplitStr(0))
                va0 = Val(SplitStr(1))
                viewang = -99999 'this is flag if search was unsuccessful
1720:
                If EOF(fileprofile%) Then Exit Do
                Line Input #fileprofile%, docprof$
                SplitStr = Split(docprof$, ",")
                azi1 = Val(SplitStr(0))
                va1 = Val(SplitStr(1))
                
                If azimuth >= azi0 And azimuth < azi1 Then
                   viewang = (azimuth - azi0) * (va1 - va0) / (azi1 - azi0) + va0
                   Exit Do
                ElseIf azimuth >= azi1 And azimuth < azi0 Then
                   viewang = (azimuth - azi1) * (va0 - va1) / (azi0 - azi1) + va1
                   Exit Do
                Else
                   va0 = va1
                   azi0 = azi1
                   GoTo 1720
                   End If
                   
            Loop
            Close #fileprofile%
            
            'calculate the Gregorian date of this daynumber
             GregDate$ = DayNumToDate(yl, dy, yr, 1)

            'now look for a sondes refraction calculation for this date and viewangle and record the calculated refraction
            'example of vdw refraction calculation using the actual height: 31-May-1986-sondes-tc-3-VDW.dat
            HgtSonde = False
            FindHgtSondeRefCalc$ = App.Path & "\" & GregDate$ & "-sondes-tc-3-VDW.dat"
            If Dir(FindHgtSondeRefCalc$) <> vbNullString Then
               HgtSonde = True
               End If
               
            If HgtSonde Then 'find the calculated refraction value for this viewang
               filesonde% = FreeFile
               Open FindHgtSondeRefCalc$ For Input As #filesonde%
               Input #filesonde%, NumRef
               found% = 0
               Do Until EOF(filesonde%)
                  Input #filesonde%, va0, ref0
                  va0 = va0 / 60# 'convert to degrees
1800:
                  If EOF(filesonde%) Then Exit Do
                  Input #filesonde%, va1, ref1
                  va1 = va1 / 60# 'convert to degrees
                  If viewang >= va0 And viewang < va1 Then
                     RefHgtSonde = (viewang - va0) * (ref1 - ref0) / (va1 - va0) + ref0
                     found% = 1
                     Exit Do
                  ElseIf viewang >= va1 And viewang < va0 Then
                     RefHgtSonde = (viewang - va1) * (ref0 - ref1) / (va0 - va1) + ref1
                     found% = 1
                     Exit Do
                  Else
                     va0 = va1
                     ref0 = ref1
                     GoTo 1800
                     End If
               Loop
               If found% = 0 Then HgtSonde = False 'failed search
               If found% = 1 And RefHgtSonde = 0 Then
                  HgtSonde = False
                  End If
               Close #filesonde%
               End If
               
            
            'also look for a zero angle refraction calculation for this date and record the calculated refraction
            'example of vdw refraction calculation using zero ground height: '31-May-1990-sondes-no-tc-3-VDW.dat
            
            ZeroHgtSonde = False
            FindZeroSondeRefCalc$ = App.Path & "\" & GregDate$ & "-sondes-no-tc-VDW.dat"
            If Dir(FindZeroSondeRefCalc$) <> vbNullString Then
               ZeroHgtSonde = True
               End If
               
            If ZeroHgtSonde Then 'find the calculated refraction value for viewangle
               'if viewangle < 0, then use the ref at zero,
               'if viewangle > 0, then determine the refraction at that viewangle
               If viewang < 0 Then
                  viewang2 = 0#
               Else
                  viewang2 = viewang
                  End If
                  
               filesonde% = FreeFile
               Open FindZeroSondeRefCalc$ For Input As #filesonde%
               Input #filesonde%, NumRef
               found% = 0
               Do Until EOF(filesonde%)
                  Input #filesonde%, va0, ref0
                  va0 = va0 / 60#  'convert to degrees
1850:
                  If EOF(filesonde%) Then Exit Do
                  Input #filesonde%, va1, ref1
                  va1 = va1 / 60# 'convert to degrees
                  If viewang2 >= va0 And viewang < va1 Then
                     RefZeroHgtSonde = (viewang2 - va0) * (ref1 - ref0) / (va1 - va0) + ref0
                     found% = 1
                     Exit Do
                  ElseIf viewang2 >= va1 And viewang < va0 Then
                     RefZeroHgtSonde = (viewang2 - va1) * (ref0 - ref1) / (va0 - va1) + ref1
                     found% = 1
                     Exit Do
                  Else
                     va0 = va1
                     ref0 = ref1
                     GoTo 1850
                     End If
               Loop
               If found% = 0 Then ZeroHgtSonde = False 'failed search
               If found% = 1 And RefZeroHgtSonde = 0 Then
                  ZeroHgtSonde = False
                  End If
               Close #filesonde%
               End If
               
            'finally calculate the vdw refraction for the worldclim temperature and vdw atmosphere used in the ray tracing that determined the refraction coefs for the netzki6 calculations
            'and take the two differences
            
            sunrise$ = Mid$(doczma$, 16, 7)
            'this is the calculated visible sunrise time over the hills
            sunrisetime = Val(Mid$(sunrise$, 1, 1)) + Val(Mid$(sunrise$, 3, 2)) / 60# + Val(Mid$(sunrise$, 6, 2)) / 3600#
            difast$ = Mid$(doczma$, 39, 7)
            'this is the difference of the calculated visible sunrise and the astronomical sunrise = vis sunrise - ast sunrise
            difasttime = Val(Mid$(difast$, 1, 1)) + Val(Mid$(difast$, 3, 2)) / 60# + Val(Mid$(difast$, 6, 2)) / 3600#
            
            'only process the above data if there was an observation recorded for this day
            If yr <= 1992 Then If Mid$(tim$(ntim), 1, 1) = "#" Then GoTo 1999 'no observation recorded
            If yr >= 1993 Then If Mid$(tim$(ntim), 1, 1) = "," Then GoTo 1999 'no observation recorded
            If qualityobservation And (InStr(tim$(ntim), "B") Or InStr(tim$(ntim), "M") Or InStr(tim$(ntim), "K")) Then GoTo 1999 'skip questional dates
            nfound = nfound + 1

            s1 = Val(Mid$(tim$(ntim), 1, 2))
            s2 = Val(Mid$(tim$(ntim), 4, 5))
            s3 = Val(Mid$(tim$(ntim), 7, 8))
            s1 = s1 + (s2 + s3 / 60) / 60  'this is extracted sunrise observation in fractional hours
            
            'diference of observerd visible sunrise from the calculated vis sunrise =
            difvissun = s1 - sunrisetime 'if observed sunrise is later than calculated. then difference is defined positive
            diffromast = difvissun + difasttime 'if difvissun is positive than make the difference also more positive than for the calculated vis sunrise
            
            If yr = 1996 Then '2
               refdy = dy
               Write #2, refdy, diffromast * 60#
               If refdy > 284 And diffromast * 60 > 5 Then
                  ccc = 1
                  End If
            Else '2
                First% = 0
                'determine which daynumber this azimuth corresponds on the 1995 year
950:
                fileref = FreeFile
                FileRefIn$ = App.Path & "\" & "RavD1996.pl1"
                Open FileRefIn$ For Input As #fileref
                Line Input #fileref, docref$
                aziref0 = Val(Mid$(docref$, 40, 7))
                aziref00 = aziref0
                NumDay = 1
1000:
                Line Input #fileref, docref$
                NumDay = NumDay + 1
                aziref1 = Val(Mid$(docref$, 40, 7))
                If azimuth >= aziref0 And azimuth < aziref1 And Abs(dy - NumDay + 1) < 2 Then '3
                   refdy = (azimuth - aziref1) / (aziref1 - aziref0) + NumDay - 1
                   Close #fileref
                   Write #2, refdy, diffromast * 60#
                   If refdy > 284 And diffromast * 60 > 5 Then
                      ccc = 1
                      End If
                   
                   If HgtSonde And ZeroHgtSonde Then
                   
                        'calculate the vdw value of the refraction
                        RefVDWValHgt = CalcVDWRef(0, 0, 756.7, dy, yr, viewang) * 0.001 / cd  'calculate vdw refraction value and convert to degrees from mrad
                        RefVDWValZeroHgt = CalcVDWRef(0, 0, 0, dy, yr, viewang) * 0.001 / cd
                        
                        If dy > 42 And dy < 306 Then
                           Write #fileall, GregDate$, refdy, azimuth, diffromast * 60#, difvissun * 60#, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                        Else
                           'remove the 15 second adhoc fix for the fileall tabulation
'                           Write #fileall, GregDate$, refdy, azimuth, diffromast * 60# - 0.25, difvissun * 60# - 0.25, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                           'visible sunrise now calculated without any adhoc fix
                           Write #fileall, GregDate$, refdy, azimuth, diffromast * 60#, difvissun * 60#, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                           End If
                           
                        End If
                        
                   GoTo 1999
                ElseIf azimuth >= aziref1 And azimuth < aziref0 And Abs(dy - NumDay + 1) < 2 Then '3
                   refdy = (azimuth - aziref0) / (aziref0 - aziref1) + NumDay
                   Close #fileref
                   Write #2, refdy, diffromast * 60#
                   If refdy > 284 And diffromast * 60 > 5 Then
                      ccc = 1
                      End If
                   
                   If HgtSonde And ZeroHgtSonde Then
                   
                        'calculate the vdw value of the refraction
                        RefVDWValHgt = CalcVDWRef(0, 0, 756.7, dy, yr, viewang) * 0.001 / cd  'calculate vdw refraction value and convert to degrees from mrad
                        RefVDWValZeroHgt = CalcVDWRef(0, 0, 0, dy, yr, viewang) * 0.001 / cd
                        
                        If dy > 42 And dy < 306 Then
                           Write #fileall, GregDate$, refdy, azimuth, diffromast * 60#, difvissun * 60#, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                        Else
                           'remove the 15 second adhoc fix for the fileall tabulation
'                           Write #fileall, GregDate$, refdy, azimuth, diffromast * 60# - 0.25, difvissun * 60# - 0.25, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                           'visible sunrise now calculated without any adhoc fix
                           Write #fileall, GregDate$, refdy, azimuth, diffromast * 60#, difvissun * 60#, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                           End If
                           
                        End If
                        
                   GoTo 1999
                Else '3
                   If NumDay <= 365 Then '4
                      aziref0 = aziref1
                      GoTo 1000
                   Else '4
                      'daynumber must be less than 1, so open azimuth file from previous year
                      Close #fileref
                      If First% = 0 Then
                         First% = 1
                         GoTo 950
                       Else 'do search for nearest azimuth and if close enough, use it
                            fileref = FreeFile
                            FileRefIn$ = App.Path & "\" & "RavD1996.pl1"
                            Open FileRefIn$ For Input As #fileref
                            For j = 1 To 366
                               Line Input #fileref, docref$
                               aziref1 = Val(Mid$(docref$, 40, 7))
                               If Abs(aziref1 - azimuth) < 0.002 Then
                                  refdy = j
                                  Write #2, refdy, diffromast * 60#
                                    If refdy > 284 And diffromast * 60 > 5 Then
                                       ccc = 1
                                       End If
                   
                                  If HgtSonde And ZeroHgtSonde Then
                                
                                     'calculate the vdw value of the refraction
                                     RefVDWValHgt = CalcVDWRef(0, 0, 756.7, dy, yr, viewang) * 0.001 / cd  'calculate vdw refraction value and convert to degrees from mrad
                                     RefVDWValZeroHgt = CalcVDWRef(0, 0, 0, dy, yr, viewang) * 0.001 / cd
                                     
                                     If dy > 42 And dy < 306 Then
                                        Write #fileall, GregDate$, refdy, azimuth, diffromast * 60#, difvissun * 60#, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                                     Else
                                        'remove the 15 second adhoc fix for the fileall tabulation
'                                        Write #fileall, GregDate$, refdy, azimuth, diffromast * 60# - 0.25, difvissun * 60# - 0.25, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                                        'visible sunrise now calculated without any adhoc fix
                                        Write #fileall, GregDate$, refdy, azimuth, diffromast * 60#, difvissun * 60#, viewang, RefHgtSonde / 60# + RefVDWValHgt, RefZeroHgtSonde / 60# + RefVDWValZeroHgt
                                        End If
                                        
                                     End If
                                     
                                  GoTo 1999
                                  Exit For
                                  First% = 0
                                  End If
                            Next j
                            Close #fileref
                            End If
                      End If '4
                      
                   End If '3
                   
               End If '2
            
1999:       dy = dy + 1
         Loop
         Close #fileazi
         Close #filezma
         Close #1
         Close #2
         Close #fileall
         fileall = 0
           
    ElseIf optGolan.Value = True Then '1
       If i > 4 Then
          Close
          Screen.MousePointer = vbDefault
          Exit Sub
          End If
          
       yr = 2000 + Val(yrext$)
              
       yl = DaysinYear(Int(yr))
        
        nyr = i
'         yr = 2009 + nyr
         freein% = 1
         freeout% = 2
         
        Open namein$(i) For Input As #freein%
        Open NameOut$(i) For Output As #freeout%

        FileAziIn$ = App.Path & "\" & "599_" & Trim$(Str$(yr)) & ".pl1"
        FileZmaIn$ = App.Path & "\" & "599_" & Trim$(Str$(yr))
        fileazi = FreeFile
        Open FileAziIn$ For Input As #fileazi
        filezma = FreeFile
        Open FileZmaIn$ For Input As #filezma
        
       
       dy = 1
       Do While dy < yl + 1
        
            Line Input #fileazi, docazi$
            Line Input #filezma, doczma$
            'extract the azimuth and the vis sunrise time and the difference of time to the astr
            azimuth = Val(Mid$(docazi$, 40, 7))
            sunrise$ = Mid$(doczma$, 16, 7)
            sunrisetime = Val(Mid$(sunrise$, 1, 1)) + Val(Mid$(sunrise$, 3, 2)) / 60# + Val(Mid$(sunrise$, 6, 2)) / 3600#
            difast$ = Mid$(doczma$, 39, 7)
            difasttime = Val(Mid$(difast$, 1, 1)) + Val(Mid$(difast$, 3, 2)) / 60# + Val(Mid$(difast$, 6, 2)) / 3600#
            
            Line Input #freein%, doclin$
            SplitStr = Split(doclin$, ",")
            TimeStr = SplitStr(2)
            
            If Len(TimeStr) > 1 Then '1
                TimeStr = SplitStr(2)
                s1 = Val(Mid$(TimeStr, 1, 1))
                s2 = Val(Mid$(TimeStr, 3, 2))
                s3 = Val(Mid$(TimeStr, 6, 2))
                s1 = s1 + (s2 + s3 / 60) / 60
            
                'diference from vis sunrise =
                difvissun = s1 - sunrisetime
                diffromast = difvissun + difasttime
            
                If yr = 2012 Then '2
                   refdy = dy
                   Write #2, refdy, diffromast * 60#
                Else '2
                    First% = 0
                    'determine which daynumber this azimuth corresponds on the 1995 year
1950:
                    fileref = FreeFile
                    FileRefIn$ = App.Path & "\" & "599_2012.pl1"
                    Open FileRefIn$ For Input As #fileref
                    Line Input #fileref, docref$
                    aziref0 = Val(Mid$(docref$, 40, 7))
                    aziref00 = aziref0
                    NumDay = 1
2000:
                    Line Input #fileref, docref$
                    NumDay = NumDay + 1
                    aziref1 = Val(Mid$(docref$, 40, 7))
                    If azimuth >= aziref0 And azimuth < aziref1 And Abs(dy - NumDay + 1) < 2 Then '3
                       refdy = (azimuth - aziref1) / (aziref1 - aziref0) + NumDay - 1
                       Close #fileref
                       Write #2, refdy, diffromast * 60#
                       GoTo 2999
                    ElseIf azimuth >= aziref1 And azimuth < aziref0 And Abs(dy - NumDay + 1) < 2 Then '3
                       refdy = (azimuth - aziref0) / (aziref0 - aziref1) + NumDay
                       Close #fileref
                       Write #2, refdy, diffromast * 60#
                       GoTo 2999
                    Else '3
                       If NumDay <= 365 Then '4
                          aziref0 = aziref1
                          GoTo 2000
                       Else '4
                          'daynumber must be less than 1, so open azimuth file from previous year
                          Close #fileref
                          If First% = 0 Then '5
                             First% = 1
                             GoTo 1950
                           Else 'do search for nearest azimuth and if close enough, use it '5
                                fileref = FreeFile
                                FileRefIn$ = App.Path & "\" & "599_2012.pl1"
                                Open FileRefIn$ For Input As #fileref
                                For j = 1 To 366
                                   Line Input #fileref, docref$
                                   aziref1 = Val(Mid$(docref$, 40, 7))
                                   If Abs(aziref1 - azimuth) < 0.002 Then
                                      refdy = j
                                      Write #2, refdy, diffromast * 60#
                                      GoTo 2999
                                      Exit For
                                      First% = 0
                                      End If
                                Next j
                                Close #fileref
                                End If '5
                                
                          End If '4
                          
                       End If '3
                       
                   End If '2
                   
                End If '1
                
2999:           dy = dy + 1
             Loop
             
             Close #fileazi
             Close #filezma
             Close #1
             Close #2
            
        Screen.MousePointer = vbDefault

    ElseIf optZuriel.Value = True Then
     
'Zuri1991.pl1
'Zuri1991

       If i > 13 Then
          Close
          Screen.MousePointer = vbDefault
          Exit Sub
          End If
       
       yr = 1900 + Val(yrext$)
       
       If yr = 1980 Then
          ccc = 1
          End If
       
       yl = DaysinYear(Int(yr))
       
'        FileAziIn$ = App.Path & "\" & "Zuri" & Trim$(Str$(yr)) & ".pl1"
'        FileZmaIn$ = App.Path & "\" & "Zuri" & Trim$(Str$(yr))
'        FileAziIn$ = App.Path & "\" & "213-" & Trim$(Str$(yr)) & ".pl1"
'        FileZmaIn$ = App.Path & "\" & "213-" & Trim$(Str$(yr))
        nyr = i
'         yr = 2009 + nyr
         freein% = 1
         freeout% = 2
         
        Open App.Path & "\" & NameZuriel For Input As #freein%
        'advance in the file until it is at the position of the current year
2925:
        If EOF(freein%) Then GoTo 5000
        Line Input #freein%, doclin$
        SplitStr = Split(doclin$, ",")
        TimeStr = SplitStr(0)
        DateStr = SplitStr(1)
        
        mondy = Val(Mid$(DateStr, 1, 2))
        daydy = Val(Mid$(DateStr, 4, 2))
        yrdy = 1900 + Val(Mid$(DateStr, 7, 2))
        
        daynumbdy = DayNumber(Int(yl), Int(mondy), Int(daydy))
        
        If yrdy <> yr Then GoTo 2925
           
        Open NameOut$(i) For Output As #freeout%
        
'        FileInCom = App.Path & "\" & "bnei" & Trim$(Str$(yr)) & ".com"
'        filecom = FreeFile
'        Open FileInCom For Input As #filecom

'        FileAziIn$ = App.Path & "\" & "Zuri" & Trim$(Str$(yr)) & ".pl1"
        FileAziIn$ = App.Path & "\" & "bnei" & Trim$(Str$(yr)) & ".com"
'        FileZmaIn$ = App.Path & "\" & "Zuri" & Trim$(Str$(yr))
        FileZmaIn$ = App.Path & "\" & "213-" & Trim$(Str$(yr))
        
'        For i = 0 To 3
'            FileAziInBnei(i) = App.Path & "\213-" & Trim$(Str$(yr)) & ".pl" & Trim$(Str$(i + 1))
'            filazibnei(i) = FreeFile
'            Open FileAziInBnei(i) For Input As #filazibnei(i)
'        Next i

        fileazi = FreeFile
        Open FileAziIn$ For Input As #fileazi
        filezma = FreeFile
        Open FileZmaIn$ For Input As #filezma
        
       found% = 0
       dy = 1
       Do While dy < yl + 1
        
            Line Input #fileazi, docazi$
            Line Input #filezma, doczma$
            
'            For i = 0 To 3
'               Line Input #filazibnei(i), docaziZ$(i)
'            Next i
            
            'determine which pl file to use for the viewangle
'            Line Input #filecom, doccom$
            pos% = InStr(docazi$, "213-bnei.pr")
            FileNamePr$ = App.Path & "\" & Mid$(docazi$, pos%, 12)
            myfile = Dir(FileNamePr$)
            If myfile = vbNullString Then
               Call MsgBox("Can't find the following ""pr"" file:" _
                           & vbCrLf & vbCrLf & FileNamePr$ _
                           & vbCrLf & "Aborting this loop" _
                           , vbInformation, "Missing file")
               GoTo 3999
'            Else
''               FileAziIn$ = FileNamePl$
'               'find which file this is
'               pos% = InStr(docazi$, ".pr")
'               NumAzi = Val(Mid$(fileazin$, pos% + 3, 1)) - 1
               End If
               
            'extract the azimuth and the vis sunrise time and the difference of time to the astr
            azimuth = Val(Mid$(docazi$, 44, 7))
            sunrise$ = Mid$(doczma$, 16, 7)
            sunrisetime = Val(Mid$(sunrise$, 1, 1)) + Val(Mid$(sunrise$, 3, 2)) / 60# + Val(Mid$(sunrise$, 6, 2)) / 3600#
            difast$ = Mid$(doczma$, 39, 7)
            difasttime = Val(Mid$(difast$, 1, 1)) + Val(Mid$(difast$, 3, 2)) / 60# + Val(Mid$(difast$, 6, 2)) / 3600#
            
            If found% = 1 Then 'read next line in observation file
                If EOF(freein%) Then GoTo 5000
                Line Input #freein%, doclin$
                SplitStr = Split(doclin$, ",")
                TimeStr = SplitStr(0)
                DateStr = SplitStr(1)
                
                mondy = Val(Mid$(DateStr, 1, 2))
                daydy = Val(Mid$(DateStr, 4, 2))
                yrdy = 1900 + Val(Mid$(DateStr, 7, 2))
                
                If yrdy <> yr Then
                   Close
                   found% = 0
                   GoTo 5000
                   End If
                
                daynumbdy = DayNumber(Int(yl), Int(mondy), Int(daydy))
                End If
                
            If daynumbdy <> dy Then
                found% = 0
                GoTo 3999
            Else
                found% = 1
                End If
            
            If Len(TimeStr) > 1 Then '1
                s1 = Val(Mid$(TimeStr, 1, 2))
                s2 = Val(Mid$(TimeStr, 4, 2))
                s3 = Val(Mid$(TimeStr, 7, 2))
                s1 = s1 + (s2 + s3 / 60) / 60
            
                'diference from vis sunrise =
                difvissun = s1 - sunrisetime
                diffromast = difvissun + difasttime
            
                If yr = 1992 Then '2
                   refdy = dy
                   Write #2, refdy, diffromast * 60#

                Else '2
                    First% = 0
                    'determine which daynumber this azimuth corresponds on the 1992 year
2950:
                    fileref = FreeFile
                    FileRefIn$ = App.Path & "\" & "bnei1992.com"
                    Open FileRefIn$ For Input As #fileref
                    Line Input #fileref, docref$
                    aziref0 = Val(Mid$(docref$, 44, 7))
                    aziref00 = aziref0
                    NumDay = 1
3000:
                    Line Input #fileref, docref$
                    NumDay = NumDay + 1
                    aziref1 = Val(Mid$(docref$, 44, 7))
                    If azimuth >= aziref0 And azimuth < aziref1 And Abs(dy - NumDay + 1) < 2 Then '3
                       refdy = (azimuth - aziref1) / (aziref1 - aziref0) + NumDay - 1
                       Close #fileref
                       Write #2, refdy, diffromast * 60#
                       GoTo 3999
                    ElseIf azimuth >= aziref1 And azimuth < aziref0 And Abs(dy - NumDay + 1) < 2 Then '3
                       refdy = (azimuth - aziref0) / (aziref0 - aziref1) + NumDay
                       Close #fileref
                       Write #2, refdy, diffromast * 60#
                       GoTo 3999
                    Else '3
                       If NumDay <= 365 Then '4
                          aziref0 = aziref1
                          GoTo 3000
                       Else '4
                          'daynumber must be less than 1, so open azimuth file from previous year
                          Close #fileref
                          If First% = 0 Then '5
                             First% = 1
                             GoTo 2950
                           Else 'do search for nearest azimuth and if close enough, use it '5
                                fileref = FreeFile
                                FileRefIn$ = App.Path & "\" & "bnei1992.com"
                                Open FileRefIn$ For Input As #fileref
                                For j = 1 To 366
                                   Line Input #fileref, docref$
                                   aziref1 = Val(Mid$(docref$, 44, 7))
                                   If Abs(aziref1 - azimuth) < 0.002 Then
                                      refdy = j
                                      Write #2, refdy, diffromast * 60#
                                      GoTo 3999
                                      Exit For
                                      First% = 0
                                      End If
                                Next j
                                Close #fileref
                                End If '5
                                
                          End If '4
                          
                       End If '3
                       
                   End If '2
                   
                End If '1
                
3999:           dy = dy + 1
             Loop
             
             Close #fileazi
'             For i = 0 To 3
'                Close #filazibnei(0)
'             Next i
'             Close #filecom
             Close #filezma
             Close #1
             Close #2
       
       End If
       
5000:
Next i
8000:
    Screen.MousePointer = vbDefault
'
9999  End

1360 ms = FNms(dy1)
     If ms > pi2 Then ms = ms - pi2
     aas = FNaas(dy1)
     If aas > pi2 Then aas = aas - pi2
     es = FNes(aas)
     If es > pi2 Then es = es - pi2
     d = FNarsin(Sin(ob) * Sin(es))
     Return

refwin2:
     X1 = (-0.000014745) * (hgt - 750) + 0.156272
     X2 = 0.000001085 * (hgt - 800) - 0.0180992
     frwin = (P / (T + 273)) * (X1 + X2 * X + 0.004479 * X * X)
     frwin = frwin / (1 + 0.230189 * X + 0.045977 * X * X)
Return

refsum2:
     X1 = (-0.000012378) * (hgt - 760) + 0.14206
     X2 = 0.0000025467 * (hgt - 760) - 0.0178073
     frsum = (P / (T + 273)) * (X1 + X2 * X + 0.00437 * X * X)
     frsum = frsum / (1 + 0.20879 * X + 0.04276 * X * X)
Return
refwin1:
     X1 = (-0.000015) * (hgt - 770) + 0.1565775
     frwin = (P / (T + 273)) * (X1 - 0.0082 * X + 0.0006 * X * X)
     frwin = frwin / (1 + 0.3256 * X + 0.01084 * X * X)
Return

refsum1:
     X1 = (-0.0000124) * (hgt - 759.5828) + 0.1423908
     frsum = (P / (T + 273)) * (X1 - 0.0073 * X + 0.0005 * X * X)
     frsum = frsum / (1 + 0.3085 * X + 0.0101 * X * X)
Return


   On Error GoTo 0
   Exit Sub

cmdRun_Click_Error:
    Close
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRun_Click of Form Drukfrm"
'    Resume 'diagnostics

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdRunNetzski6_Click
' Author    : Dr-John-K-Hall
' Date      : 5/12/2020
' Purpose   : Run latest version of netzski6 to create files used for observation plotting
'---------------------------------------------------------------------------------------
'
Private Sub cmdRunNetzski6_Click()

   On Error GoTo cmdRunNetzski6_Click_Error
   
   Dim yr As Integer, yl As Integer, waitime As Long
   Dim yrStart As Integer, yrEnd As Integer
   
     '------------------progress bar initialization
    With Drukfrm
      '------fancy progress bar settings---------
      .progressfrm.Visible = True
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
    Call UpdateStatus(Drukfrm, picProgBar, 0, 0) 'reset

   If optDruk.Value = True Then
   
      yrStart = 1985
      yrEnd = 1996
   
      For yr = yrStart - 1 To yrEnd Step 1
      
        fileout = FreeFile
        Open "c:/jk/netzskiy.tm3" For Output As #fileout
      
        yl = DaysinYear(yr)
      
        Write #fileout, yr  '1996
        Write #fileout, 1, 0
        Write #fileout, 1
        Print #fileout, "Druk"
        Print #fileout, "c:\cities\jerusalem_other_neighborhoods"
        Print #fileout, "c:\fordtm\netz\RavDrkTR.pr3"
'         172.654, 128.471, 756.7, 1988, 1, 366, 1988, 1, 366, 0
        Print #fileout, " 172.654, 128.471, 756.7, " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", 0"
        Print #fileout, "2,3,5,10,13,16,19,20,18,16,10,5"
        Print #fileout, "8,9,11,16,20,22,24,24,23,20,15,10"
        Print #fileout, "12,13,16,21,25,28,29,29,28,25,19,13"
        Close #fileout
        
        waitime = Timer
        Do Until Timer > waitime + 1
           DoEvents
        Loop
        
'        RetVal = Shell("c:/jk/netzski6.exe", 6) ' Run netzski3 as DOS shell
'        myfile = Dir("c:/MyProjects (CHAIM-PIV)/netzski6_c_2/netzski6/Release/netzski6.exe")
        RetVal = Shell("c:/MyProjects (CHAIM-PIV)/netzski6_c_2/netzski6/Release/netzski6.exe", 6) ' Run netzski3 as DOS shell
        
      
       Call UpdateStatus(Drukfrm, picProgBar, 1, CLng(100# * (yr - yrStart) / (yrEnd - yrStart)))
        
      Next yr
      Call UpdateStatus(Drukfrm, picProgBar, 0, 0) 'reset
      progressfrm.Visible = False
   
   ElseIf optGolan.Value = True Then
   
      yrStart = 2009
      yrEnd = 2013
      
      For yr = yrStart - 1 To yrEnd Step 1
      
        fileout = FreeFile
        Open "c:/jk/netzskiy.tm3" For Output As #fileout
      
        yl = DaysinYear(yr)
      
        Write #fileout, yr  '1996
        Write #fileout, 1, 2
        Write #fileout, 1
        Print #fileout, "visu"
        Print #fileout, "c:\cities\eros\visual_tmp"
        Print #fileout, "c:\fordtm\netz\599_jeru.pr1"
'         172.654, 128.471, 756.7, 1988, 1, 366, 1988, 1, 366, 0
        Print #fileout, " 173.51, 138.84, 684.3, " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", 0"
        Print #fileout, "2,3,5,10,13,16,19,19,18,15,10,6"
        Print #fileout, "9,9,12,16,20,22,24,24,23,20,15,10"
        Print #fileout, "12,13,16,21,25,28,29,29,28,25,19,13"
        Close #fileout
        
        waitime = Timer
        Do Until Timer > waitime + 1
           DoEvents
        Loop
        
'        RetVal = Shell("c:/jk/netzski6.exe", 6) ' Run netzski3 as DOS shell
'        myfile = Dir("c:/MyProjects (CHAIM-PIV)/netzski6_c_2/netzski6/Release/netzski6.exe")
        RetVal = Shell("c:/MyProjects (CHAIM-PIV)/netzski6_c_2/netzski6/Release/netzski6.exe", 6) ' Run netzski3 as DOS shell
        
      
       Call UpdateStatus(Drukfrm, picProgBar, 1, CLng(100# * (yr - yrStart + 1) / (yrEnd - yrStart + 1)))
        
      Next yr
      Call UpdateStatus(Drukfrm, picProgBar, 0, 0) 'reset
      progressfrm.Visible = False
   
   ElseIf optZuriel.Value = True Then
   
     yrStart = 1961
      yrEnd = 1992
      
      For yr = yrStart - 1 To yrEnd Step 1
      
        fileout = FreeFile
        Open "c:/jk/netzskiy.tm3" For Output As #fileout
      
        yl = DaysinYear(yr)
      
        Write #fileout, yr  '1996
        Write #fileout, 1, 0
        Write #fileout, 3 '4
        Print #fileout, "bnei"
        Print #fileout, "c:\cities\bnei_brak_Zuriel"
        Print #fileout, "c:\fordtm\netz\213-bnei.pr1"
'         172.654, 128.471, 756.7, 1988, 1, 366, 1988, 1, 366, 0
        Print #fileout, " 134.175, 166.35, 52.2, " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", 0"
        Print #fileout, "9,9,11,14,17,20,21,22,21,18,14,10"
        Print #fileout, "13,13,15,18,21,24,26,26,25,22,18,15"
        Print #fileout, "17,18,20,24,26,28,30,30,29,27,23,19"
        Print #fileout, "c:\fordtm\netz\213-bnei.pr2"
'         172.654, 128.471, 756.7, 1988, 1, 366, 1988, 1, 366, 0
        Print #fileout, " 133.906, 165.787, 53.2, " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", 0"
        Print #fileout, "9,9,11,14,17,20,21,22,21,18,14,10"
        Print #fileout, "13,13,15,18,21,24,26,26,25,22,18,15"
        Print #fileout, "17,18,20,24,26,28,30,30,29,27,23,19"
        Print #fileout, "c:\fordtm\netz\213-bnei.pr3"
'         172.654, 128.471, 756.7, 1988, 1, 366, 1988, 1, 366, 0
        Print #fileout, " 133.974, 166.364, 52.2, " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", 0"
        Print #fileout, "9,9,11,14,17,20,21,22,21,18,14,10"
        Print #fileout, "13,13,15,18,21,24,26,26,25,22,18,15"
        Print #fileout, "17,18,20,24,26,28,30,30,29,27,23,19"
        Print #fileout, "c:\fordtm\netz\213-bnei.pr5"
'         172.654, 128.471, 756.7, 1988, 1, 366, 1988, 1, 366, 0
        Print #fileout, " 134.413, 164.965, 61.8, " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", " & Trim$(Str$(yr)) & ", 1, " & Trim$(Str$(yl)) & ", 0"
        Print #fileout, "9,9,11,14,17,20,21,22,21,18,14,10"
        Print #fileout, "13,13,15,18,21,24,26,26,25,22,18,15"
        Print #fileout, "17,18,20,24,26,28,30,30,29,27,23,19"
        Close #fileout
        
        waitime = Timer
        Do Until Timer > waitime + 1
           DoEvents
        Loop
        
'        RetVal = Shell("c:/jk/netzski6.exe", 6) ' Run netzski3 as DOS shell
'        myfile = Dir("c:/MyProjects (CHAIM-PIV)/netzski6_c_2/netzski6/Release/netzski6.exe")
        RetVal = Shell("c:/MyProjects (CHAIM-PIV)/netzski6_c_2/netzski6/Release/netzski6.exe", 6) ' Run netzski3 as DOS shell
        
      
       Call UpdateStatus(Drukfrm, picProgBar, 1, CLng(100# * (yr - yrStart + 1) / (yrEnd - yrStart + 1)))
        
      Next yr
      Call UpdateStatus(Drukfrm, picProgBar, 0, 0) 'reset
      progressfrm.Visible = False
      
       End If

   On Error GoTo 0
   Exit Sub

cmdRunNetzski6_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRunNetzski6_Click of Form Drukfrm"
End Sub

'Private Sub cmdStart_Click()
'   Initialize = False
'   HeightLoop = True
'   cmdheight_Click() = True
'End Sub

Private Sub cmdVisTimes_Click()
   'convert netzski6 output for either with or without adhoc sunrise fix to daynumber, time difference files
   Dim filein As Integer, fileout As Integer, FileInName As String, FileOutName As String
   Dim TimeString() As String, Hours As String, mins As Double, NumDay As Integer
   
   If optDruk.Value = True Then
    FileInName = "c:\fordtm\netz" & "\RavD1996"
'    FileInName = App.Path & "\RavD1996"

'   FileOutName = App.Path & "\Druk-vis-sunrise-1995.csv"
'   FileOutName = App.Path & "\RavD_TR_1995.csv"
   If optMinTK.Value = True Then
'        FileOutName = App.Path & "\RavD_NO_mt_1996.csv"
        FileOutName = App.Path & "\RavD_TR_mt_1996.csv" '"\RavD_NO_mt_1995.csv"
     ElseIf optAveTK.Value = True Then
        FileOutName = App.Path & "\RavD_TR_at_1996.csv"
        End If
   ElseIf optGolan.Value = True Then
'     FileInName = App.Path & "\599_2012"
     FileInName = "c:\fordtm\netz" & "\599_2012"
     If optMinTK.Value = True Then
'        FileOutName = App.Path & "\Golan_NO_mt_2012.csv"
        FileOutName = App.Path & "\Golan_TR_mt_2012.csv"
     ElseIf optAveTK.Value = True Then
        FileOutName = App.Path & "\Golan_TR_at_2012.csv"
        End If
   ElseIf optZuriel.Value = True Then
'     FileInName = App.Path & "\213-1992"
'     FileInName = "c:\fordtm\netz" & "\Zuri1992"
'     FileInName = "c:\fordtm\netz" & "\Zuri1988"
'     FileInName = App.Path & "\Zuri1992"
     FileInName = "c:\fordtm\netz" & "\213-1992"
     If optMinTK.Value = True Then
        FileOutName = App.Path & "\Zuriel_NO_mt_1992.csv"
'        FileOutName = App.Path & "\Zuriel_TR_mt_1992.csv"
     ElseIf optAveTK.Value = True Then
        FileOutName = App.Path & "\Zuriel_TR_at_1992.csv"
        End If
     End If
      
   filein = FreeFile
   Open FileInName For Input As #filein
   fileout = FreeFile
   Open FileOutName For Output As #fileout
   NumDay = 0
   Do Until EOF(filein)
      NumDay = NumDay + 1
      Line Input #filein, doclin$
      Hours = Mid$(doclin$, Len(doclin$) - 6, 7)
      TimeString = Split(Hours, ":")
      mins = Val(TimeString(1)) + Val(TimeString(2)) / 60#
      Print #fileout, NumDay, mins
   Loop
   Close #fileout
   Close #filein
   
End Sub

Private Sub Form_Load()
   Initialize = True
   With cmbHeight
      For i = 0 To 3000 Step 100
         .AddItem i
      Next i
      .ListIndex = 0
   End With
   Initialize = False
End Sub
