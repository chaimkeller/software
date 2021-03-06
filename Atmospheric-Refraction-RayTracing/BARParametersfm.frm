VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form BARParametersfm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculation mode"
   ClientHeight    =   13335
   ClientLeft      =   7395
   ClientTop       =   2880
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13335
   ScaleWidth      =   5985
   Begin VB.Frame frmSeason 
      Caption         =   "Season"
      Height          =   475
      Left            =   240
      TabIndex        =   30
      Top             =   4500
      Width           =   5535
      Begin VB.OptionButton optAllOrigPress 
         Caption         =   "All orig pressures"
         Height          =   195
         Left            =   3720
         TabIndex        =   35
         ToolTipText     =   "calculate refraction for all seasons using the recorded pressures"
         Top             =   180
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optAllSeasons 
         Caption         =   "All seasons"
         Height          =   195
         Left            =   2520
         TabIndex        =   34
         ToolTipText     =   "calculate the sondes for all seasons using vdw ciddor pressures"
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton optSummer 
         Caption         =   "Summer"
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         ToolTipText     =   "Calculate atmospheres and refraction for May-Jul"
         Top             =   160
         Width           =   1335
      End
      Begin VB.OptionButton optWinter 
         Caption         =   "Winter"
         Height          =   195
         Left            =   500
         TabIndex        =   31
         ToolTipText     =   "Calculate atmosphrees for Jan-Feb, Nov-Dec"
         Top             =   160
         Width           =   855
      End
   End
   Begin VB.Frame frmPurpose 
      Caption         =   "Choose option"
      Height          =   840
      Left            =   240
      TabIndex        =   23
      Top             =   3680
      Width           =   5535
      Begin VB.OptionButton optfit1 
         Caption         =   "eps and ref fitting"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton optCalculate 
         Caption         =   "Calculate ray trracing for atm files"
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   220
         Width           =   2655
      End
      Begin VB.OptionButton optConvert 
         Caption         =   "Convert sondes to atm. files"
         Height          =   195
         Left            =   220
         TabIndex        =   24
         Top             =   220
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame frmsondes 
      Caption         =   "Convert Sondes to atmosphere csv files"
      Height          =   8175
      Left            =   240
      TabIndex        =   16
      Top             =   5040
      Width           =   5535
      Begin VB.CommandButton cmdFitFiles 
         Caption         =   "Fit Files"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1500
         TabIndex        =   37
         Top             =   7840
         Width           =   855
      End
      Begin VB.CommandButton cmdAutoBrowse 
         Caption         =   "Auto Browse"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         ToolTipText     =   "Auto browse directory for file type"
         Top             =   7840
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddRef 
         Caption         =   "Add Ref"
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   7840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddVA 
         Caption         =   "Add VAs"
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   7840
         Width           =   1575
      End
      Begin VB.CommandButton cmdCalcAtms 
         Caption         =   "Calculate"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   7560
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   7560
         Width           =   495
      End
      Begin VB.CommandButton cmdUnselect 
         Caption         =   "Unselect all"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   7560
         Width           =   975
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Check All"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   7560
         Width           =   855
      End
      Begin VB.CommandButton cmdConvertSonde 
         Caption         =   "Convert"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         ToolTipText     =   "convert to atmosphere file: elevation,temperature,pressure"
         Top             =   7560
         Width           =   735
      End
      Begin VB.ListBox lstSondes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7080
         Left            =   240
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton cmdBrowseSonde 
         Caption         =   "Browse"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "browse for sondes to convert or for atmospheric files to calculate ray tracing"
         Top             =   7560
         Width           =   855
      End
   End
   Begin VB.Frame frmCompareCalc 
      Caption         =   "Compare TR Caculations"
      Height          =   3135
      Left            =   2400
      TabIndex        =   10
      Top             =   520
      Width           =   3375
      Begin VB.CommandButton cmdCompareTR 
         Caption         =   "Compare TR calculations"
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "Check"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdRedo 
         Caption         =   "Redo"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   615
      End
      Begin VB.PictureBox picProgBar 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   2835
         TabIndex        =   13
         Top             =   2280
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse for profile fiile"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtFileName 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Text            =   "txtFileName"
         Top             =   360
         Width           =   2895
      End
      Begin MSComDlg.CommonDialog comdlgCompare 
         Left            =   2520
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame frmFindPressure 
      Caption         =   "Ciddor Pressure"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   360
      TabIndex        =   2
      Top             =   520
      Visible         =   0   'False
      Width           =   1815
      Begin VB.TextBox txtCiddorHgt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCiddor 
         Caption         =   "Calculate"
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtCiddorDry 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtCiddorWet 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblCiddorHeight 
         Caption         =   "Obsv. hgt (m)"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblCiddorDry 
         Caption         =   "Dry Pressure (mb)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblCiddorWet 
         Caption         =   "Vapor Pressure (mb)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkPlotMode 
      Caption         =   "Plot suns on main form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   280
      Width           =   2655
   End
   Begin VB.CheckBox chkCalcMode 
      Caption         =   "Use newerer interface for calculation"
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
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "BARParametersfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkCalcMode_Click()
   CalcMode = 0
   If chkCalcMode.Value = vbChecked Then
      CalcMode = 1
      End If
End Sub

Private Sub chkPlotMode_Click()
   SunPlotMode = 0
   If chkPlotMode.Value = vbChecked Then
      SunPlotMode = 1
      End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAddRef_Click
' Author    : Dr-John-K-Hall
' Date      : 3/18/2020
' Purpose   : determines the refraction value corresponding the the view angle, and finds the difference from the standard vdW value
'---------------------------------------------------------------------------------------
'
Private Sub cmdAddRef_Click()

   On Error GoTo cmdAddRef_Click_Error

   'open the c:/jk/Druk-Vangeld-data/Druk-sondes-and-view-angles.csv file
   'read the sondes name, and the view angle, and then find the corresponding tc file
   'then find the tc corresponding to this sondes and view angle, and then compare to the standard vdW for the calculation height of 800.5 meters
   'tc file: tc_VDW_800point5 meters-standardvdWatms.dat, write difference in degrees so it can be plotted in minutes by multiplying by a number
   
   Dim FileNameIn As String, filein As Integer
   Dim FileNameOut As String, fileout As Integer
   Dim FileNameRef As String, fileref As Integer
   Dim FileNamevdWRef As String, filevdw As Integer
   Dim doclin$, DocSplit() As String, VA As Double, RefSondes As Double, DifRef As Double
   Dim VA1 As Double, VA2 As Double, ref1 As Double, ref2 As Double, numRef As Integer, numiter As Integer
   Dim RefVDW(73) As Double, RefVDWVal As Double, lenDirect As Integer, DateRef As String
   Dim daysonde As Integer, monsonde As Integer, yrsonde As Integer, yl As Integer
   Dim RefVDW1 As Double, RefVDW2 As Double, DayNumberSonde As Integer, height As Double
   Dim pi As Double, cd As Double
   
'   ZeroRefTesting = False 'set to true to only compare the zero viewanlge refraction values
'   If prjAtmRefMainfm.OptionSelby.Value = True And prjAtmRefMainfm.chkHgtProfile.Value = vbUnchecked Then ZeroRefTesting = True
'
'   ZeroRefTesting = True
   
   'store the reference tc values
   'use instead the formulas for the vdW refraction at a certain height and temperature and view angle
   '<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   FileNamevdWRef = "c:/jk/Druk-Vangeld-data/tc_VDW_800point5 meters-standardvdWatms.dat"
'   filevdw = FreeFile
'   Open FileNamevdWRef For Input As #filevdw
'   Input #filevdw, numRef
'   For i = 0 To numRef
'      Input #filevdw, VA1, ref1
'      RefVDW(i) = ref1
'   Next i
'   Close #filevdw

   pi = 4 * Atn(1)
   cd = pi / 180#  'conv deg to rad

    height = Val(prjAtmRefMainfm.txtHOBS.Text) 'DistModel(0)
    
'    ZeroRefTesting = True 'diagnostics

   If DirectOut$ = "" Then DirectOut$ = "c:/jk/Druk-Vangeld-data/"
   lenDirect = Len(DirectOut$)
   
   If optWinter.Value = True Then
      FileNameIn = DirectOut$ & "Druk-sondes-and-view-angles.csv"
   ElseIf optSummer.Value = True Then
      FileNameIn = DirectOut$ & "Druk-sondes-and-view-angles-sum.csv"
   ElseIf optAllSeasons.Value = True Then
      FileNameIn = DirectOut$ & "Druk-sondes-and-view-angles-all-2.csv"
      If ZeroRefTesting Then
         FileNameIn = DirectOut$ & "Druk-sondes-and-view-angles-no-all-2.csv"
         End If
   ElseIf optAllOrigPress.Value = True Then
      FileNameIn = DirectOut$ & "Druk-sondes-and-view-angles-all-3.csv"
      If ZeroRefTesting Then
         FileNameIn = DirectOut$ & "Druk-sondes-and-view-angles-no-all-3.csv"
         End If
     End If
      
   filein = FreeFile
   Open FileNameIn For Input As #filein
   
   If optWinter.Value = True Then
      FileNameOut = DirectOut$ & "Druk-final-refraction-results.csv"
   ElseIf optSummer.Value = True Then
      FileNameOut = DirectOut$ & "Druk-final-refraction-results-sum.csv"
   ElseIf optAllSeasons.Value = True Then
      FileNameOut = DirectOut$ & "Druk-final-refraction-results-all-2.csv"
      If ZeroRefTesting Then
         FileNameOut = DirectOut$ & "Druk-final-refraction-results-no-all-2.csv"
         End If
   ElseIf optAllOrigPress.Value = True Then
      FileNameOut = DirectOut$ & "Druk-final-refraction-results-all-3.csv"
      If ZeroRefTesting Then
         FileNameOut = DirectOut$ & "Druk-final-refraction-results-no-all-3.csv"
         End If
         
     End If
      
   fileout = FreeFile
   Open FileNameOut For Output As #fileout
   
   Do Until EOF(filein)
      found% = 0
      Line Input #filein, doclin$
      DocSplit = Split(doclin$, ",")
      VA = Val(DocSplit(UBound(DocSplit))) 'in degrees '* 60 'convert view angle from degrees to arcminutes
      'determine name of tc listing corresponding to this sonde
      'e.g.: 24-Jan-1993-sondes-tc-VDW.dat
      If Not ZeroRefTesting And (optWinter.Value = True Or optSummer.Value = True Or optAllSeasons.Value = True) Then
        FileNameRef = Mid$(DocSplit(0), 1, Len(DocSplit(0)) - 4) & "-tc-2-VDW.dat"
      ElseIf Not ZeroRefTesting And optAllOrigPress.Value = True Then
        FileNameRef = Mid$(DocSplit(0), 1, Len(DocSplit(0)) - 4) & "-tc-3-VDW.dat"
      ElseIf optAllSeasons.Value = True And ZeroRefTesting Then
        FileNameRef = Mid$(DocSplit(0), 1, Len(DocSplit(0)) - 4) & "-no-tc-VDW.dat"
      ElseIf optAllOrigPress.Value = True And ZeroRefTesting Then
        FileNameRef = Mid$(DocSplit(0), 1, Len(DocSplit(0)) - 4) & "-no-tc-3-VDW.dat"
        End If
      'extract date
      DateRef = Mid$(DocSplit(0), lenDirect + 1, 11)
      yrsonde = Val(Mid$(DateRef, 8, 4))
      daysonde = Val(Mid$(DateRef, 1, 2))
      MonName$ = Mid$(DateRef, 4, 3)
      Select Case MonName$
         Case "Jan"
            monsonde = 1
         Case "Feb"
            monsonde = 2
         Case "Mar"
            monsonde = 3
         Case "Apr"
            monsonde = 4
         Case "May"
            monsonde = 5
         Case "Jun"
            monsonde = 6
         Case "Jul"
            monsonde = 7
         Case "Aug"
            monsonde = 8
         Case "Sep"
            monsonde = 9
         Case "Oct"
            monsonde = 10
         Case "Nov"
            monsonde = 11
         Case "Dec"
            monsonde = 12
       End Select
       
       'now determine the daynumber
       yl = DaysinYear(yrsonde)
       DayNumberSonde = DayNumber(yl, monsonde, daysonde)
      
'      If FileNameRef = "C:\jk\Druk-Vangeld-data\26-Jan-1996-sondes-tc-VDW.dat" Then
'         ccc = 1
'         End If
      If Dir(FileNameRef) <> vbNullString Then
         fileref = FreeFile
         found% = 1
         Open FileNameRef For Input As #fileref
         Input #fileref, numRef
         Input #fileref, VA1, ref1
         numiter = 0
         
100:
         numiter = numiter + 1
         Input #fileref, VA2, ref2
         
         If VA >= VA2 And VA < VA1 And Not ZeroRefTesting And (optSummer.Value = True Or optWinter.Value = True Or optAllSeasons.Value = True Or optAllOrigPress.Value = True) Then
            RefSondes = ((ref1 - ref2) / (VA1 - VA2)) * (VA - VA2) + ref2
           
            'now determine the reference vdw value
                   
'            RefVDW1 = CalcVDWRef(0, 0, height, DayNumberSonde, yrsonde, VA1) 'use Beit Dagan's coordinates to determine the ground temperature
'            RefVDW2 = CalcVDWRef(0, 0, height, DayNumberSonde, yrsonde, VA2)
'            RefVDWVal = ((RefVDW1 - RefVDW2) / (VA1 - VA2)) * (VA - VA2) + RefVDW2
            RefVDWVal = CalcVDWRef(0, 0, height, DayNumberSonde, yrsonde, VA) * 0.001 / cd 'calculate vdw refraction value and convert to degrees from mrad
            RefVDWVal = -RefVDWVal 'refraction is less then zero by definition
'            RefVDWVal = ((RefVDW(numiter - 1) - RefVDW(numiter)) / (VA1 - VA2)) * (VA - VA2) + RefVDW(numiter)
            DifRef = -RefSondes + RefVDWVal * 60#
            'now record these results in degrees
            Print #fileout, doclin$ & "," & Trim$(Str$(-RefVDWVal / 60#)) & "," & Trim$(Str$(DifRef / 60#))
            found% = 1
            Close #fileref
            GoTo 200 'read next sondes name
         ElseIf VA2 = 0 And (optAllSeasons.Value = True Or optAllOrigPress.Value = True) And ZeroRefTesting Then
            RefSondes = ref2
            'due to low height, there is no further view angles in the sondes output file, so use the VA = 0 value
            height = 0 'calculating for ground height of zero
            RefVDWVal = CalcVDWRef(0, 0, height, DayNumberSonde, yrsonde, 0#)
            RefVDWVal = RefVDWVal * 0.001 / cd 'calculate vdw refraction value and convert to degrees from mrad
            RefVDWVal = -RefVDWVal * 60# 'refraction is less then zero by definition (convert to arcminutes)
'            RefVDWVal = ((RefVDW(numiter - 1) - RefVDW(numiter)) / (VA1 - VA2)) * (VA - VA2) + RefVDW(numiter)
            DifRef = RefSondes - RefVDWVal
            'now record these results in degrees
            Print #fileout, doclin$ & "," & Trim$(Str$(-RefVDWVal / 60#)) & "," & Trim$(Str$(DifRef / 60#))
            found% = 1
            Close #fileref
            GoTo 200 'read next sondes name
         Else
            If VA2 = numRef * 0.25 Then Exit Do
            VA1 = VA2
            ref1 = ref2
            GoTo 100
            End If
         Close #fileref
      Else
         found% = 0
         Exit Do
         End If
200:

   Loop
   Close #filein
   Close #fileout
    
   If found% = 0 Then
      Call MsgBox("Can't find corresponding tc file for the following entry:" & vbCrLf & vbCrLf & FileNameRef, vbCritical, "Search failed")
      End If

   On Error GoTo 0
   Exit Sub

cmdAddRef_Click_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdAddRef_Click of Form BARParametersfm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAddVA_Click
' Author    : Dr-John-K-Hall
' Date      : 3/18/2020
' Purpose   : Add view angle information to the sondes-found-and-results.csv file
'---------------------------------------------------------------------------------------
'
Private Sub cmdAddVA_Click()

   Dim FileNameSondes As String, filein As Integer, fileout As Integer
   Dim FileNameAzimuth As String, fileazi As Integer
   Dim FileNameAziOut As String
   Dim FileDrukProfile As String, filedruk As Integer
   Dim doclin$, DocSplit() As String
   Dim DayNumber As Double, daynum1 As Double, daynum2 As Double
   Dim az1 As Double, az2 As Double, VA1 As Double, VA2 As Double
   Dim AziOut As Double, VAOut As Double, doclinIn$
   Dim azi1 As Double, azi2 As Double
   Dim FileNameRecord As String, filerec As Integer
   Dim FileResults As String, fileres As Integer
   Dim DateRes$, ref As Double
   
'   ZeroRefTesting = True 'diagnostics

   If DirectOut$ = "" Then DirectOut$ = "c:/jk/Druk-Vangeld-data/"
   
   If optWinter.Value = True Then
      FileNameAziOut = DirectOut$ & "Druk-sondes-and-view-angles.csv"
      FileNameSondes = DirectOut$ & "Druk-sondes-found-and-results.csv"

   ElseIf optSummer.Value = True Then
      FileNameAziOut = DirectOut$ & "Druk-sondes-and-view-angles-sum.csv"
      FileNameSondes = DirectOut$ & "Druk-sondes-found-and-results-sum.csv"
      
   ElseIf optAllSeasons.Value = True Then
      FileNameAziOut = DirectOut$ & "Druk-sondes-and-view-angles-all-2.csv"
      FileNameSondes = DirectOut$ & "Druk-sondes-found-and-results-all-2.csv"
      If ZeroRefTesting Then
        FileNameAziOut = DirectOut$ & "Druk-sondes-and-view-angles-no-all-2.csv"
        FileNameSondes = DirectOut$ & "Druk-sondes-found-and-results-no-all-2.csv"
        End If
        
   ElseIf optAllOrigPress.Value = True Then
      FileNameAziOut = DirectOut$ & "Druk-sondes-and-view-angles-all-3.csv"
      FileNameSondes = DirectOut$ & "Druk-sondes-found-and-results-all-3.csv"
      If ZeroRefTesting Then
        FileNameAziOut = DirectOut$ & "Druk-sondes-and-view-angles-no-all-3.csv"
        FileNameSondes = DirectOut$ & "Druk-sondes-found-and-results-no-all-3.csv"
        End If
      End If
   
   If Dir(FileNameAziOut) <> vbNullString Then
      Select Case MsgBox("Output file already exists!" _
                         & vbCrLf & "" _
                         & vbCrLf & "Do you want to overwrite?" _
                         , vbYesNo Or vbQuestion Or vbDefaultButton2, "Overwrite protect")
      
        Case vbYes
      
        Case vbNo
            Exit Sub
      End Select
      End If
   
'   FileNameSondes = DirectOut$ & Druk-sondes-found-and-results.csv"
   If Dir(FileNameSondes) = vbNullString Then 'write the file
      If optWinter.Value = True Then
         FileNameRecord = DirectOut$ & "Druk-sondes-found.csv"
         FileResults = DirectOut$ & "Druk-sondes-results.csv"
      ElseIf optSummer.Value = True Then
         FileNameRecord = DirectOut$ & "Druk-sondes-found-sum.csv"
         FileResults = DirectOut$ & "Druk-sondes-results-sum.csv"
      ElseIf optAllSeasons.Value = True Then
         FileNameRecord = DirectOut$ & "Druk-sondes-found-all-2.csv"
         FileResults = DirectOut$ & "Druk-sondes-results-all-2.csv"
         If ZeroRefTesting Then
            FileNameRecord = DirectOut$ & "Druk-sondes-found-no-all-2.csv"
            FileResults = DirectOut$ & "Druk-sondes-results-no-all-2.csv"
            End If
      ElseIf optAllOrigPress.Value = True Then
         FileNameRecord = DirectOut$ & "Druk-sondes-found-all-3.csv"
         FileResults = DirectOut$ & "Druk-sondes-results-all-3.csv"
         If ZeroRefTesting Then
            FileNameRecord = DirectOut$ & "Druk-sondes-found-no-all-3.csv"
            FileResults = DirectOut$ & "Druk-sondes-results-no-all-3.csv"
            End If
         End If
         
      If Dir(FileNameRecord) <> vbNullString And Dir(FileResults) <> vbNullString Then
      Else
         Call MsgBox("Can't find the sondes-found file or the sondes-results file: " _
         & vbCrLf & vbCrLf & FileNameRecord, vbInformation, "Can't find file")
         Exit Sub
         End If
           
      fileout = FreeFile
      Open FileNameSondes For Output As #fileout
      filerec = FreeFile
      Open FileNameRecord For Input As #filerec
      
      
      Do Until EOF(filerec)
         Line Input #filerec, doclin$
         DocSplit = Split(doclin$, ",")
         'determine the date
         FileSonde$ = DocSplit(0)
         FileDate$ = Mid$(FileSonde$, Len(DirectOut$) + 1, 11)
         fileres = FreeFile
         Open FileResults For Input As #fileres
         Do Until EOF(fileres)
            Line Input #fileres, doclin2$ 'DateRes$, Ref
            DateRes$ = Mid$(doclin2$, 1, 11)
            ref = Val(Mid$(doclin2$, 12, Len(doclin2$) - 12))
            If DateRes$ = FileDate And ref <> 0 Then
               Print #fileout, doclin$ & "," & Trim$(Str$(ref))
               Close #fileres
               Exit Do
               End If
         Loop
         Close #fileres
      Loop
      Close #filerec
      Close #fileout
        
      End If
      
   filein = FreeFile
   Open FileNameSondes For Input As #filein
   
   FileNameAzimuth = DirectOut$ & "RavD1995-azimuths.csv"
   
'   FileNameAziOut = "c:/jk/Druk-Vangeld-data/Druk-sondes-and-view-angles.csv"
   fileout = FreeFile
   Open FileNameAziOut For Output As #fileout
   
   FileDrukProfile = DirectOut$ & "RavDrkTR.pr1"
   
   On Error GoTo cmdAddVA_Click_Error
   
50:
   If EOF(filein) Then GoTo 900
   Line Input #filein, doclinIn$
   DocSplit = Split(doclinIn$, ",")
   'extract the daynumber
   DayNumber = Val(DocSplit(2))
   
   'now use the azimuth file to determine what azimuth corresponds to the sunrise on this daynumber
   fileazi = FreeFile
   Open FileNameAzimuth For Input As #fileazi
   
100:
    found% = 0
    Input #fileazi, daynum1, azi1
    If daynum1 = 365 Then GoTo 900
    
150:
    Input #fileazi, daynum2, azi2
    If DayNumber >= daynum1 And DayNumber < daynum2 Then
       'interpolate to determine the azimuth
       AziOut = ((azi2 - azi1) / (daynum2 - daynum1)) * (DayNumber - daynum1) + azi1
       'now open the profile file and determine the corresponding view angle at this azimuth
       filedruk = FreeFile
       Open FileDrukProfile For Input As #filedruk
       Line Input #filedruk, doclin$ 'skip two lines of documentation
       Line Input #filedruk, doclin$

       Input #filedruk, az1, VA1, bb, cc, dd, ee
       If az1 = 45# Then GoTo 900
170:
       Input #filedruk, az2, VA2, bb, cc, dd, ee
       If AziOut >= az1 And AziOut < az2 Then
          VAOut = ((VA2 - VA1) / (az2 - az1)) * (AziOut - az1) + VA1
          Print #fileout, doclinIn$ & "," & Trim$(Str$(AziOut)) & "," & Trim$(Str$(VAOut))
          found% = 1
       Else
          If az2 = 45# Then GoTo 900
          az1 = az2
          VA1 = VA2
          found% = 0
          GoTo 170 'read next line in the profile file
          End If
       
       Close #filedruk
       Close #fileazi
       GoTo 50 'loop to next sondes-found entry
    
    Else
       If daynum2 = 365 Then GoTo 900
       daynum1 = daynum2
       azi1 = azi2 'read next daynumber, azimuth
       found% = 0
       GoTo 150
       End If

900:

If found% = 0 And Not EOF(filein) Then
    Call MsgBox("Search failed for following line of input:" & vbCrLf & vbCrLf & doclinIn$, vbCritical, "Search failed")
    End If

Close

   On Error GoTo 0
   Exit Sub

cmdAddVA_Click_Error:
    Close
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdAddVA_Click of Form BARParametersfm"
End Sub

Private Sub cmdAll_Click()
   For i = 1 To lstSondes.ListCount
      lstSondes.Selected(i - 1) = True
   Next i
End Sub

Private Sub cmdAutoBrowse_Click()
    Dim dirname As String
    Dim files As Collection
    Dim filename As Variant
    Dim txtPatterns As String
    
        lstSondes.Clear
        If optfit1.Value = True Then 'auto load the tr_vdw files from c:\devstudio\vb
            dirname = "c:\devstudio\vb\"
            txtPatterns = "*_0_32.dat"
            DirectOut$ = dirname
            Set files = FindFiles(dirname, txtPatterns)
            For Each filename In files
               If InStr(filename, "LAYERS") Or InStr(filename, "INV") Then
               Else
                  lstSondes.AddItem dirname & filename
                  End If
            Next filename
            If lstSondes.ListCount > 0 Then cmdFitFiles.Enabled = True
            Exit Sub
            End If
           
        dirname = InputBox("Enter the full name of the directory to search", "Search direc.", "c:\jk\Druk-Vangeld-data")

'        dirname = txtDirectory.Text
        If optConvert.Value = True Then
           txtPatterns = InputBox("Enter the search pattern", "Search parameters", "*sondes.txt")
        ElseIf optCalculate.Value = True Then
           txtPatterns = InputBox("Enter the search pattern", "Search parameters", "*-sondes.txt")
           End If
           
'        txtPatterns = "*-sondes.txt"
        If Right$(dirname, 1) <> "\" Then dirname = dirname & "\"
        DirectOut$ = dirname
        
        Set files = FindFiles(dirname, txtPatterns)
        For Each filename In files
            If optConvert.Value = True Then 'only search for the txt files dowloaded from the sondes archive
                If InStr(filename, "-") <> 0 Then GoTo 500
            ElseIf optCalculate.Value = True Then
                If InStr(filename, "-") = 0 Then GoTo 500
                End If
                
            If InStr(filename, "beit-dagan") <> 0 Then GoTo 500 'skip these files, since they are the raw sondes.
            
            If optWinter.Value = True Then
               If InStr(filename, "Jan") Or InStr(filename, "Feb") Or InStr(filename, "Nov") Or InStr(filename, "Dec") Then
                  lstSondes.AddItem dirname & filename
                  End If
            ElseIf optSummer.Value = True Then
               If InStr(filename, "May") Or InStr(filename, "Jun") Or InStr(filename, "Jul") Then
                  lstSondes.AddItem dirname & filename
                  End If
            ElseIf optAllSeasons.Value = True Or optAllOrigPress.Value = True Then
               If InStr(filename, "Jan") Or InStr(filename, "Feb") Or InStr(filename, "Nov") Or InStr(filename, "Dec") _
                  Or InStr(filename, "May") Or InStr(filename, "Jun") Or InStr(filename, "Jul") Then
                  lstSondes.AddItem dirname & filename
                  End If
               End If
500:
        Next filename
End Sub

Private Sub cmdBrowse_Click()

  On Error GoTo errhand
  
  With comdlgCompare
    .CancelError = True
    .Filter = "pr files (*.p*)|*.p*|All files (*.*)|*.*"
    .ShowOpen
    txtFileName = .filename
  End With
  Exit Sub
  
errhand:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdBrowse_Click of Form BARParametersfm"

End Sub

Private Sub cmdBrowseSonde_Click()
    Dim strTemp() As String

   On Error GoTo errhand
   
   With comdlgCompare
        .CancelError = True
        .filename = "*.txt" & Space$(2048) & vbNullChar & vbNullChar
        .MaxFileSize = Len(.filename)
        .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        .Flags = cdlOFNAllowMultiselect Or cdlOFNLongNames Or cdlOFNExplorer
        .ShowOpen
        strTemp = Split(.filename, vbNullChar)
   End With
   
   DirectOut$ = strTemp(0)
   
   For i = 1 To UBound(strTemp)
      If InStr(strTemp(i), "*.txt") Then 'some sort of bug, don't add this to the list
      Else
        lstSondes.AddItem DirectOut$ & "\" & strTemp(i)
        End If
   Next i
   lstSondes.ListIndex = lstSondes.ListCount - 1
   lstSondes.Refresh
   Exit Sub
   
errhand:
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCalculate_Click
' Author    : Dr-John-K-Hall
' Date      : 10/30/2019
' Purpose   : compare TR calculations
'---------------------------------------------------------------------------------------
'
Private Sub cmdCalculate_Click()

    'open chosen file and read coordinates, etc, and compare the TR vdW calculation to the modified Wikipedia formula for TR
    
   On Error GoTo cmdCalculate_Click_Error
   
   Dim lg1 As Double, lt1 As Double, kmxo As Double, kmyo As Double, H11 As Double
   Dim lg2 As Double, lt2 As Double, H21 As Double, lg As Double, lt As Double
   Dim MT(12) As Integer, AT(12) As Integer, ier As Integer, FileNameOut As String
   Dim AveMinTmp As Double, AveAvgTmp As Double, azi As Double, VA As Double, kmx As Double, kmy As Double
   Dim distd As Double, deltd As Double, defm As Double, defb As Double, avref As Double
   Dim PATHLENGTH As Double, Press0 As Double, j As Integer, NNN As Integer
   Dim FileMode As Integer, HMAXT As Double, RELHUM As Double, StartAng As Double, EndAng As Double
   Dim WAVELN As Double, OBSLAT As Double, NSTEPS As Long, HUMID As Double, HOBS As Double
   Dim StepSize As Integer, RecordTLoop As Boolean, ier2 As Long, LastVA As Double, NAngles As Long
   Dim DistTolerance As Double, D1 As Double, viewangle As Double, TRRayTrace As Double
   Dim X1 As Double, X2 As Double, Y1 As Double, Y2 As Double, Exponent As Double
   Dim z1 As Double, z2 As Double, re1 As Double, re2 As Double
   Dim dist1 As Double, dist2 As Double, ANGLE As Double, hgtDTM
   Dim MinAzimuth As Double, MaxAzimuth As Double, geo As Boolean
   Dim hgtworld As Double, kcurve As Double ', Rcurve As Double
   
   Rearth = 6356766#
   RE = Rearth
   
    pi = 4# * Atn(1#) '3.141592654
    CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
    cd = pi / 180# 'conversion of degrees to radians
   
'   Dim MinAzimuth As Double, MaxAzimuth As Double

    filein% = FreeFile
    Open txtFileName.Text For Input As #filein%
    Line Input #filein%, doclin$
    If InStr(doclin$, "kmxo") Then
       geo = False
    ElseIf InStr(doclin$, "Lati") Then
       geo = True
    Else
       Call MsgBox("Can't determine if this a geo file or not from the header" _
                   & vbCrLf & "" _
                   & vbCrLf & "Aborting....." _
                   , vbExclamation, "geo coordinates")
       Close
       picProgBar.Visible = False
       Screen.MousePointer = vbDefault
       Exit Sub
       End If
       
    Input #filein%, lg1, lt1, H11, startkmx, sofkmx, dkmx, dkmy, APPRNR
    If Not geo Then
       'EY ITM, convert to geo coordinates
       Call casgeo(lg1, lt1, lg, lt)
       lg1 = -lg
       lt1 = lt
    ElseIf geo Then
       tmplt = lt1
       lt1 = lg1
       lg1 = -tmplt
       End If
       
    'now load up minimum and average world temperatures
    Call Temperatures(lt1, lg1, MT, AT, ier)
    
    'determine solar azimuth range for this latitude
    'at sunirse, sunet, cos(azimuth) = sin(decl)/cos(latitude)
    'declination varies from -23.5 to 23.5 degrees therefore
    MinAzimuth = -DASIN(Sin(23.5 * cd) / Cos(lt1 * cd)) / cd
    MaxAzimuth = -MinAzimuth
    'MaxAzimuth at June 21, Minimum azimuth at Dec 21, zero at Mar 21 and Sep 21 but temperature
    'very different during March through April than from June through October
    'find average mean temperature over the year and use that value
    AveMinTmp = 0
    For i = 1 To 12
       AveMinTmp = MT(i) + AveMinTmp
    Next i
    AveMinTmp = AveMinTmp / 12 + 273.15
    
    If Dir(App.Path & "\CompareTR-7.txt") <> sEmpty Then
       Select Case MsgBox("File: output file ""CompareTR-7.txt "" already exists!" _
                          & vbCrLf & "" _
                          & vbCrLf & "Do you want to copy it to a backup before proceeding?" _
                          , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "File Overwrite protection")
       
        Case vbYes
            FileCopy App.Path & "\CompareTR-7.txt", App.Path & "\CompareTR-7-old.txt"
            
            Call MsgBox("""CompareTR-7.txt"" has been copied to ""CompareTR-7-old.txt""" _
                        & vbCrLf & "" _
                        , vbInformation Or vbDefaultButton1, "File Overwrite protection")
            
        Case vbNo
       
        Case vbCancel
       
       End Select
       End If
    
    fileout% = FreeFile
    Open App.Path & "\CompareTR-7.txt" For Output As #fileout%
    Print #fileout%, "Expected VA (deg.), Old; TR(degrees), Wikipedia; TR(degrees), RayTracing; TR(degrees)"
   
   Screen.MousePointer = vbHourglass
   
    '-------------------------------------------------
    With BARParametersfm
      '------fancy progress bar settings---------
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
        
    Call UpdateStatus(BARParametersfm, picProgBar, 1, 0)
    
    j = 0
    NNN = CInt(2# * Abs(MinAzimuth) / 0.1) + 1
    
    'use this value for comparisons
    Do Until EOF(filein%)
        Input #filein%, azi, VA, kmx, kmy, distd, H21
        If azi < MinAzimuth Or azi > MaxAzimuth Then GoTo 1000
        
        If Not geo Then
           Call heights(kmx, kmy, hgtDTM)
           H21 = hgtDTM
           'now convert to geo coordinates
           Call casgeo(kmx, kmy, lg, lt)
        Else
           Call worldheights(kmx, kmy, hgtworld)
           H21 = hgtworld
           lg = -kmx
           lt = kmy
           End If
           
        D1 = Rearth * DistTrav(lt1, -lg1, lt, lg, 1)
        
        'first caclulate old terrestrial refraction
        deltd = H11 - H21
        If (deltd <= 0#) Then
            defm = 0.000782 - deltd * 0.000000311
            defb = deltd * 0.000034 - 0.0141
        ElseIf (deltd > 0#) Then
            defm = deltd * 0.000000309 + 0.000764
            defb = -0.00915 - deltd * 0.0000269
            End If
        avref = defm * distd + defb
        If (avref < 0#) Then
            avref = 0#
            End If
            
        FilePath = App.Path
        StepSie = 1
        RecordTLoop = False
        FileMode = 1 'mode used for determination of terrestrial refraction using the dll
        
        With prjAtmRefMainfm
            Press0 = Val(.txtPress0)
            HMAXT = Val(.txtHMAXT)
            RELHUM = Val(.txtRELHUM)
            StartAng = Val(.txtBETAHI) * 60# 'convert to arc minutes
            EndAng = Val(.txtBETALO) * 60#
            StepAng = Val(.txtBETAST) * 60#
            WAVELN = Val(.txtKmin) * 0.001 'Val(.txtWAVELN)
            OBSLAT = lt1
            NSTEPS = Val(.txtNSTEPS)
        End With
        
        If NSTEPS < 5000 Then NSTEPS = 5000
        HUMID = RELHUM
        HOBS = H11
        StepSize = Val(prjAtmRefMainfm.txtHeightStepSize.Text)
        NAngles = 2 * StartAng / StepAng + 1
        LastVA = 9999999 'insure proper temperature progression, which should be proportional to the inverse square of the temperature
        DistTolerance = 1
    
        'now calculate estimate of TR using Reijs's formula using the average minimum temperature
'        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rearth * 0.001)) ^ 2#)
'        TR = (0.0083 * PATHLENGTH * Press0) / (AveMinTmp * AveMinTmp)
        
        'use Wikipedia expression instead
        'https://en.wikipedia.org/wiki/Atmospheric_refraction#Terrestrial_refraction
        
'        '//////////////////////begin old version////////////////////////////////
'        'curvature of rays is according to Wikipedia article
'        'https://en.wikipedia.org/wiki/Atmospheric_refraction
'        lR = -0.0065  'K/m  'lapse rate of US standard atmosphere
'        'kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
'        Rcurve = Rearth '/ (1 - kcurve)
'        'use parabolic path length instead of distd
'        'approximate the path length as the ratio of the curvatures
''        PATHLENGTH = distd / (1 - kcurve)
'        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
'        Exponent = 1#  '0.9975
'        TR = 8.15 * (PATHLENGTH ^ Exponent) * 1000 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp) 'arcseconds
'        TR = TR / 3600 'degrees
'        'TR = TR / 1.3195 '/ 1.52 '1.3195 'fudge factor
'        '//////////////////////////////end old version/////////////////////////////////////
        
        lR = -0.0065  'K/m
        'curvature of rays is according to Wikipedia article
        'https://en.wikipedia.org/wiki/Atmospheric_refraction
'        kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
'        Rcurve = Rearth / (1 - kcurve)
        'use parabolic path length instead of distd
'        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
        
        lR = -0.0065  'K/m  'lapse rate of US standard atmosphere
'        kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
        Rcurve = Rearth ' / (1 - kcurve)
        'use parabolic path length instead of distd
        'approximate the path length as the ratio of the curvatures
'        PATHLENGTH = distd / (1 - kcurve)
        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
        PATHLENGTH = PATHLENGTH * 1000 'convert to meters
'        PATHLENGTH = Sqr(2# * Rcurve * Abs(H21 - H11) + (H21 - H11) ^ 2#) 'path length in meters
        If (H21 - H11) > 1000 Then
           Exponent = 0.99 '0.9975  '0.9975
        Else
           Exponent = 0.9965 '1 '0.9945
           End If
        '0.0342 is the lapse rate of an uniformaly dense atmosphere at hydrostatic equilibrium
        'i.e., determines how much would have to decrease the temperature vs height to keep the density constant
        'thereofore, rays wouldn't bend if LR = 0
        TR = 8.15 * (PATHLENGTH ^ Exponent) * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp) 'arcseconds
        TR = TR / 3600 'degrees
        
        'calculate expected viewangle in radians
        GoSub VAsub
        
        BARParametersfm.Refresh
        
        ier2 = RayTracing(StartAng, EndAng, StepAng, LastVA, NAngles, _
                         D1, viewangle, H21, DistTolerance, FileMode, _
                         H11, AveMinTmp, HMAXT, FilePath, StepSize, _
                         Press0, WAVELN, HUMID, OBSLAT, NSTEPS, _
                         RecordTLoop, AveMinTmp, AveMinTmp, AddressOf MyCallback)
        If ier2 = 0 Then
            TRRayTrace = (LastVA - viewangle) / cd 'calculated TR in degrees
        ElseIf ier2 < 0 Then 'didn't converge
            TRRayTrace = 0#
            End If
        
'        Print #fileout%, j, viewangle, avref, TR / 3600, TRRayTrace
        DoEvents
        
        j = j + 1

        Print #fileout%, j, Format(Str(viewangle), "#0.0#####"), Format(Str(avref), "#0.0#####"), Format(Str(TR), "#0.0#####"), Format(Str(TRRayTrace), "#0.0#####")
        
        If j = 15 Then
           ccc = 1
           End If
           
        Call UpdateStatus(BARParametersfm, picProgBar, 1, CLng(100# * j / NNN))
        
1000:
    
    Loop
    
    Close #filein%
    Close #fileout%
    
    Screen.MousePointer = vbDefault
    
    Call UpdateStatus(BARParametersfm, picProgBar, 1, 0)
    
    picProgBar.Visible = False
    
    'now plot the results as function of line number

   On Error GoTo 0
   Exit Sub
   
VAsub:
'    RE = Rearth
    hgt1 = H11
    hgt2 = H21
    X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
    X2 = Cos(lt * cd) * Cos(-lg * cd)
    Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
    Y2 = Cos(lt * cd) * Sin(-lg * cd)
    z1 = Sin(lt1 * cd)
    z2 = Sin(lt * cd)
'    Rearth = 6371315#
    re1 = (hgt1 + RE)
    re2 = (hgt2 + RE)
    X1 = re1 * X1
    Y1 = re1 * Y1
    z1 = re1 * z1
    X2 = re2 * X2
    Y2 = re2 * Y2
    z2 = re2 * z2
    dist1 = re1
    dist2 = re2
    ANGLE = DACOS((X1 * X2 + Y1 * Y2 + z1 * z2) / (dist1 * dist2))
    viewangle = Atn((-re1 + re2 * Cos(ANGLE)) / (re2 * Sin(ANGLE)))
Return

cmdCalculate_Click_Error:
    Close
    picProgBar.Visible = False
    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdCalculate_Click of Form BARParametersfm"
End Sub

Private Sub cmdCalculateRef_Click(Index As Integer)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdCalcAtms_Click
' Author    : Dr-John-K-Hall
' Date      : 3/10/2020
' Purpose   : loops through the radiosondes based atmospheres and computes the vdW ray tracing for them
'---------------------------------------------------------------------------------------
'
Private Sub cmdCalcAtms_Click()

   Select Case MsgBox("Do you want the atmoshpere to follow the terrain?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Hill Hugging")
    Case vbYes
        HillHugging = True
    Case vbNo
        HillHugging = False
   End Select
   
   If Not HillHugging Then
      Select Case MsgBox("Do you want to only compare observed to calculated refraction at viewangle = 0 ?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Zero Ref Testing")
       Case vbYes
         ZeroRefTesting = True
       Case vbNo
         ZeroRefTesting = False
      End Select
      
      Select Case MsgBox("Do you want renormalize the sonde initial elevation to zero?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Renorm sondes starting elevation")
       Case vbYes
         ReNormHeight = True
       Case vbNo
         ReNormHeight = False
       End Select
       
       End If
       
'///////////////////convert Druk csv dates to same format as radiosondes dates/////////////////
'///////////////////STEP 1/////////////////////////
'  Dim DocSplit() As String, OldDate$, NewDate$
'  Dim DateSplit() As String
'  'first convert dates in weather file Druk-mt-weather-compare-8-2.csv to dd-mmm-yyyy format
'  filein = FreeFile
'  Open "c:\jk\Druk-Vangeld-data\Druk-mt-weather-compare-8-2.csv" For Input As #filein
'  fileout = FreeFile
'  Open "c:\jk\Druk-Vangeld-data\Druk-mt-weather-compare-8-2-d.csv" For Output As #fileout
'  Do Until EOF(filein)
'     Line Input #filein, doclin$
'     DocSplit = Split(doclin$, ",")
'     'convert date
'     OldDate$ = DocSplit(0)
'     DateSplit = Split(OldDate$, "/")
'     Select Case DateSplit(1)
'        Case "01"
'           NewDate$ = DateSplit(0) & "-Jan-" & DateSplit(2)
'        Case "02"
'           NewDate$ = DateSplit(0) & "-Feb-" & DateSplit(2)
'        Case "03"
'           NewDate$ = DateSplit(0) & "-Mar-" & DateSplit(2)
'        Case "04"
'           NewDate$ = DateSplit(0) & "-Apr-" & DateSplit(2)
'        Case "05"
'           NewDate$ = DateSplit(0) & "-May-" & DateSplit(2)
'        Case "06"
'           NewDate$ = DateSplit(0) & "-Jun-" & DateSplit(2)
'        Case "07"
'           NewDate$ = DateSplit(0) & "-Jul-" & DateSplit(2)
'        Case "08"
'           NewDate$ = DateSplit(0) & "-Aug-" & DateSplit(2)
'        Case "09"
'           NewDate$ = DateSplit(0) & "-Sep-" & DateSplit(2)
'        Case "10"
'           NewDate$ = DateSplit(0) & "-Oct-" & DateSplit(2)
'        Case "11"
'           NewDate$ = DateSplit(0) & "-Nov-" & DateSplit(2)
'        Case "12"
'           NewDate$ = DateSplit(0) & "-Dec-" & DateSplit(2)
'        Case Else
'     End Select
'     Print #fileout, NewDate$ & "," & Trim$(DocSplit(1)) & "," & Trim$(DocSplit(2)) & "," & Trim$(DocSplit(3)) & "," & Trim$(DocSplit(4)) & "," & Trim$(DocSplit(5))
'  Loop
'  Close #filein
'  Close #fileout


'//////////////////////find the radiosondes that have corresponding good weather and Druk sunrise observations//////////////
'///////////////////STEP 2////////////////////////////////
   Dim FileNameIn As String, filein As Integer, DateName As String
   Dim DocSplit() As String, FileNameInRoot As String
   Dim filedruk As Integer, FileNameDruk As String
   Dim filerec As Integer, FileNameRecord As String, FileNameSave As String
   Dim filesav As Integer, waitime As Long

   On Error GoTo cmdCalcAtms_Click_Error
   
   If DirectOut$ = "" Then DirectOut$ = "c:/jk/Druk-Vangeld-data/"

   If optWinter.Value = True Then
      FileNameRecord = DirectOut$ & "Druk-sondes-found.csv"
   ElseIf optSummer.Value = True Then
      FileNameRecord = DirectOut$ & "Druk-sondes-found-sum.csv"
   ElseIf optAllSeasons.Value = True Then
      FileNameRecord = DirectOut$ & "Druk-sondes-found-all-2.csv"
   ElseIf optAllOrigPress.Value = True Then
      FileNameRecord = DirectOut$ & "Druk-sondes-found-all-3.csv"
      If ZeroRefTesting Then
        FileNameRecord = DirectOut$ & "Druk-sondes-found-no-all-3.csv"
        End If
     End If
      
   filerec = FreeFile
   Open FileNameRecord For Output As #filerec
   
   If optWinter.Value = True Then
      FileNameSave = DirectOut$ & "Druk-sondes-found.sav"
   ElseIf optSummer.Value = True Then
      FileNameSave = DirectOut$ & "Druk-sondes-found-sum.sav"
   ElseIf optAllSeasons.Value = True Then
      FileNameSave = DirectOut$ & "Druk-sondes-found-all-2.sav"
   ElseIf optAllOrigPress.Value = True Then
      FileNameSave = DirectOut$ & "Druk-sondes-found-all-3.sav"
      If ZeroRefTesting Then
        FileNameSave = DirectOut$ & "Druk-sondes-found-no-all-3.sav"
        End If
     End If
   
   filesav = FreeFile
      
   Open FileNameSave For Output As #filesav

   FileNameDruk = DirectOut$ & "Druk-mt-weather-compare-8-2-d.csv"

   For i = 1 To lstSondes.ListCount
      If lstSondes.Selected(i - 1) = True And lstSondes.List(i - 1) <> "*.txt" Then
         'compare date on sondes file to determine if it is one of the clear morning nights in the
         'Druk-mt-weather-compare-8-2-d.csv file, if so do the ray tracing
         FileNameIn = lstSondes.List(i - 1)
'         If Not InStr(FileNameIn, "\") Then GoTo 100 'not a sondes file for sure
50:
         If InStr(FileNameIn, "\") = 0 Then GoTo 100 'not a sondes file for sure
         DocSplit = Split(FileNameIn, "\")
         FileNameInRoot = DocSplit(UBound(DocSplit))
         If InStr(FileNameInRoot, "-sondes.txt") Then
            DateName = Mid$(FileNameInRoot, 1, 11)
            found% = 0
            filedruk = FreeFile
            Open FileNameDruk For Input As #filedruk
            Do Until EOF(filedruk)
                Line Input #filedruk, doclin$
                DocSplit = Split(doclin$, ",")
                If DocSplit(0) = DateName Then
                   'do raytracing and record 90 degrees zenith angle refraction
                   
                   'display name of currently active sondes file in the stutus bar
                   MDIAtmRef.StatusBar.Panels(2).Text = FileNameIn

                   found% = 1
                   Exit Do
                   End If
            Loop
            Close #filedruk
            End If

         If found% = 1 Then
            Print #filerec, FileNameIn & "," & DocSplit(1) & "," & DocSplit(2) & "," & DocSplit(3)
            Print #filesav, FileNameIn
            found% = 0
            'do ray tracing of this date

'            filein = FreeFile
'            Open lstSondes.List(i - 1) For Input As #filein
'            Do Until EOF(filein)
            End If
        End If
100:
'    GoTo 50
   Next i
   Close #filerec
   Close #filesav
'
'   On Error GoTo 0
'   Exit Sub

'////////////////restore list to lstsondes and calculate the refraction up to +/- 0.3 degrees
'/////////////////////step 3///////////////////
'Dim FileNameIn As String, filein As Integer
'Dim FileNameAtmOut As String, fileoutatm As Integer
'Dim DocSplit() As String
'Dim FileNameInRoot As String

lstSondes.Clear
If optWinter.Value = True Then
   FileNameIn = DirectOut$ & "Druk-sondes-found.sav"
ElseIf optSummer.Value = True Then
   FileNameIn = DirectOut$ & "Druk-sondes-found-sum.sav"
ElseIf optAllSeasons.Value = True Then
   FileNameIn = DirectOut$ & "Druk-sondes-found-all-2.sav"
ElseIf optAllOrigPress.Value = True Then
   FileNameIn = DirectOut$ & "Druk-sondes-found-all-3.sav"
   If ZeroRefTesting Then
      FileNameIn = DirectOut$ & "Druk-sondes-found-no-all-3.sav"
      End If
   End If
   
filein = FreeFile
Open FileNameIn For Input As #filein
Do Until EOF(filein)
   Line Input #filein, doclin$
   lstSondes.AddItem doclin$
Loop
Close #filein

'loop through listted sondes files and calculate their raytracing, and record the zero angle refraction on the output file
'select all the files
For i = 1 To lstSondes.ListCount
   lstSondes.Selected(i - 1) = True
Next i

If optWinter.Value = True Then
   FileNameAtmOut = DirectOut$ & "Druk-sondes-results.csv"
ElseIf optSummer.Value = True Then
   FileNameAtmOut = DirectOut$ & "Druk-sondes-results-sum.csv"
ElseIf optAllSeasons.Value = True Then
   FileNameAtmOut = DirectOut$ & "Druk-sondes-results-all-2.csv"
ElseIf optAllOrigPress.Value = True Then
   FileNameAtmOut = DirectOut$ & "Druk-sondes-results-all-3.csv"
   If ZeroRefTesting Then
      FileNameAtmOut = DirectOut$ & "Druk-sondes-results-no-all-3.csv"
      End If
   End If

NumStart = 1
If Dir(FileNameAtmOut) <> "" Then
   'ask user if want to skip dates that have already been recorded in the results file
   Select Case MsgBox("Do you wish to skip the dates that have already been calculated?", vbYesNo Or vbQuestion Or vbDefaultButton1, "Repeat calculations")
   
    Case vbYes
        SkipRepeats = True
    Case vbNo
        SkipRepeats = False
   End Select
   
   End If
   
   
'   'check how many have been done already
'    fileoutatm = FreeFile
'    Open FileNameAtmOut For Input As #fileoutatm
'    NumDone = 0
'    Do Until EOF(fileoutatm)
'       Line Input #fileoutatm, doclin$
'       NumDone = NumDone + 1
'    Loop
'    Close #fileoutatm
'    If NumDone > 0 Then
'       Select Case MsgBox(Str(NumDone) & " results have already been recorded." _
'                          & vbCrLf & "" _
'                          & vbCrLf & "Do you want to start after the last recorded result?" _
'                          & vbCrLf & "" _
'                          & vbCrLf & "(Answer ""No"" if you want to start from the beginning." _
'                          & vbCrLf & "In either case, the result file will not be erased rather appended to)" _
'                          , vbYesNo Or vbQuestion Or vbDefaultButton1, "Increment start of calculation")
'
'        Case vbYes
'            NumStart = NumDone + 1
'        Case vbNo
'            NumStart = 1
'       End Select
'       End If
'   End If
   
   
    If optWinter.Value = True Then
       FileNameAtmOut = DirectOut$ & "Druk-sondes-results.csv"
    ElseIf optSummer.Value = True Then
       FileNameAtmOut = DirectOut$ & "Druk-sondes-results-sum.csv"
    ElseIf optAllSeasons.Value = True Then
       FileNameAtmOut = DirectOut$ & "Druk-sondes-results-all-2.csv"
    ElseIf optAllOrigPress.Value = True Then
       FileNameAtmOut = DirectOut$ & "Druk-sondes-results-all-3.csv"
       If ZeroRefTesting Then
          FileNameAtmOut = DirectOut$ & "Druk-sondes-results-no-all-3.csv"
          End If
       End If

   For i = NumStart To lstSondes.ListCount
      If lstSondes.Selected(i - 1) = True Then
      
         If SkipRepeats Then
            'extract the date
            DocSplit = Split(lstSondes.List(i - 1), "\")
            FileNameInRoot = DocSplit(UBound(DocSplit))
            DateNameAtm = Mid$(FileNameInRoot, 1, 11)
            
            If Not ZeroRefTesting Then
                If optAllSeasons.Value = True Then
                   myfile = Dir(DirectOut$ & DateNameAtm & "-sondes-tc-2-VDW.dat")
                ElseIf optAllOrigPress.Value = True Then
                   myfile = Dir(DirectOut$ & DateNameAtm & "-sondes-tc-3-VDW.dat")
                   End If
            ElseIf ZeroRefTesting Then
                If optAllSeasons.Value = True Then
                   myfile = Dir(DirectOut$ & DateNameAtm & "-sondes-no-tc-VDW.dat")
                ElseIf optAllOrigPress.Value = True Then
                   myfile = Dir(DirectOut$ & DateNameAtm & "-sondes-no-tc-3-VDW.dat")
                   End If
                End If
                
            If myfile <> vbNullString Then
               LoopingAtmTracing = False
               FinishedTracing = True
               GoTo 200 'don't repeat the calculation
               End If
            End If
            
         With prjAtmRefMainfm
            .OptionSelby.Value = True
            .opt10.Value = True
            .txtOther.Text = lstSondes.List(i - 1)

            'extract the date
            DocSplit = Split(lstSondes.List(i - 1), "\")
            FileNameInRoot = DocSplit(UBound(DocSplit))
            DateNameAtm = Mid$(FileNameInRoot, 1, 11)

            .chkMeters.Value = vbChecked
            .OptionSelby.Value = True
            If ReNormHeight Then .chkReNorm.Value = vbChecked
            .WindowState = vbMaximized
            If HillHugging Then
                .chkHgtProfile.Value = vbChecked
                End If
            .txtYSize.Text = "0.3"
            .TabRef.Tab = 0

            LoopingAtmTracing = True

            FinishedTracing = False

            .cmdVDW.Value = True
            'wait here until calculation completes
            
            .SetFocus
            .Refresh
            waitime = Timer 'if scan doesn't finish in 30 minutes = 30 * 60 seconds, then notify if want to abort
            Do Until FinishedTracing = True
                DoEvents
                If cmdVDW_error = -1 Then
                   cmdVDW_error = 0 'error detected, skip iteration, reset cmdVDW_error flag
                   FinishedTracing = True
                   End If
                If Timer - waitime > 1800 Then
                   Select Case MsgBox("The ray tracing for this file is taking more than 30 minutes" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "The sondes file is: " & DateNameAtm _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you want to abort" _
                                      , vbYesNo Or vbQuestion Or vbDefaultButton2, "Scan not advancing")
                   
                    Case vbYes
                       FinishedTracing = True
                       'advance to next sondes
                    Case vbNo
                   
                   End Select
                   End If
            Loop
         End With
         End If
         
      BARParametersfm.SetFocus
      BARParametersfm.Refresh
200:
   Next i

   LoopingAtmTracing = False
   Close #fileoutatm

''/////////////////////////add azimuth information to Druk-sondes-found-and-results.csv file//////////
'//////////////////////////STEP 4//////////////////////////
''////////////////////////////also add the view angle derived from the RavD_NO_mt_1995.csv (no adhocrise fix horizon file)/////////
'Dim FileNameIn As String, filein As Integer
'Dim FileNameOut As String, fileout As Integer
'Dim FileAzimuth As String, fileazi As Integer
'Dim FileDrukHorizon As String, filedruk As Integer
'Dim DocSplit() As String, doclin$, daynumber As Integer
'
'FileNameIn = "c:/jk/Druk-Vangeld-data/Druk-sondes-found-and-results.csv"
'filein = FreeFile
'Open FileNameIn For Input As #filein
'
''open azimuth file now since the above file is sorted by daynumber
'fileazi = FreeFile
'FileAzimuth = "c:/jk/Druk-Vangeld-data/RavD1995-azimuths.csv"
'Open FileAzimuth For Input As #fileazi
'
'Do Until EOF(filein)
'   Line Input #filein, doclin$
'   DocSplit = Split(doclin$, ",")
'   'extract the daynumber
'   daynumber = Val(DocSplit(2))
'   'now use the azimuth file to determine what azimuth corresponds to the sunrise on this daynumber
'100:
'    found% = 0
'    Input #fileazi, daynum1, azi1
'    If daynum1 < 365 Then
'       Input #fileazi, daynum2, azi2
'       If daynumber >= daynum1 And daynumber < daynum2 Then
'          found% = 1
'          GoTo 200 'loop to next sondes-found entry
'          End If
'    Else
'       'reached end of file
'       found% = 0
'       Exit Do
'    Else
'       GoTo 100
'       End If
'200:
'Loop
'Close #filein
'Close #fileazi

   On Error GoTo 0
   Exit Sub
   
cmdCalcAtms_Click_Error:
    LoopingAtmTracing = False
    Close
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdCalcAtms_Click of Form BARParametersfm"
End Sub

Private Sub cmdCheck_Click()

   Dim lg1 As Double, lt1 As Double, kmxo As Double, kmyo As Double, H11 As Double
   Dim lg2 As Double, lt2 As Double, H21 As Double, lg As Double, lt As Double
   Dim MT(12) As Integer, AT(12) As Integer, ier As Integer, FileNameOut As String
   Dim AveMinTmp As Double, AveAvgTmp As Double, azi As Double, VA As Double, kmx As Double, kmy As Double
   Dim distd As Double, deltd As Double, defm As Double, defb As Double, avref As Double
   Dim PATHLENGTH As Double, Press0 As Double, j As Integer, NNN As Integer
   Dim FileMode As Integer, HMAXT As Double, RELHUM As Double, StartAng As Double, EndAng As Double
   Dim WAVELN As Double, OBSLAT As Double, NSTEPS As Long, HUMID As Double, HOBS As Double
   Dim StepSize As Integer, RecordTLoop As Boolean, ier2 As Long, LastVA As Double, NAngles As Long
   Dim DistTolerance As Double, D1 As Double, viewangle As Double, TRRayTrace As Double
   Dim X1 As Double, X2 As Double, Y1 As Double, Y2 As Double, Exponent As Double
   Dim z1 As Double, z2 As Double, re1 As Double, re2 As Double
   Dim dist1 As Double, dist2 As Double, ANGLE As Double, hgtDTM
   Dim MinAzimuth As Double, MaxAzimuth As Double, geo As Boolean
   Dim hgtworld As Double, kcurve As Double ', Rcurve As Double
   
   Rearth = 6356766#
   RE = Rearth
   
    pi = 4# * Atn(1#) '3.141592654
    CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
    cd = pi / 180# 'conversion of degrees to radians
    
'18.944050,99.202150,1585,99.3021,98.3021,0.000277778,0.000277778,0.00440527
    H11 = 1585
    H21 = 1749
    lt1 = 18.94405
    lg1 = -99.220215
    lt = 18.931734
    lg = 99.110833
    GoSub VAsub
    viewangle = viewangle / cd
    
    
'8.1,0.927812,-99.110833,18.931734,9.70194,1749
'8,0.89678,-99.110833,18.931888,9.69954,1749
'8,0.902617,-99.110833,18.931888,9.69954,1749

Exit Sub
VAsub:
'    RE = Rearth
    hgt1 = H11
    hgt2 = H21
    X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
    X2 = Cos(lt * cd) * Cos(-lg * cd)
    Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
    Y2 = Cos(lt * cd) * Sin(-lg * cd)
    z1 = Sin(lt1 * cd)
    z2 = Sin(lt * cd)
'    Rearth = 6371315#
    re1 = (hgt1 + RE)
    re2 = (hgt2 + RE)
    X1 = re1 * X1
    Y1 = re1 * Y1
    z1 = re1 * z1
    X2 = re2 * X2
    Y2 = re2 * Y2
    z2 = re2 * z2
    dist1 = re1
    dist2 = re2
    ANGLE = DACOS((X1 * X2 + Y1 * Y2 + z1 * z2) / (dist1 * dist2))
    viewangle = Atn((-re1 + re2 * Cos(ANGLE)) / (re2 * Sin(ANGLE)))
Return
End Sub

Private Sub cmdCiddor_Click()
  Dim H As Double
  Dim PDRY As Double
  Dim PVAP As Double
  Dim NumLayers As Long

    RELHUM = Val(txtRELHUM) 'relative humidity
    RELH = RELHUM / 100
    For i = 2 To 50
       If HL(i) = 0 Then
          NumLayers = i - 1
          Exit For
          End If
    Next i

    H = H * 1000#  'convert to meters

'    PDRY = fFNDPD1(H, PRESSD1, Dist, NumLayers) 'to get to work, need to reference this function globally in a module
'    PVAP = RELH * fVAPOR(H, Dist, NumLayers)
    
    txtCiddorDry.Text = " "
    txtCiddorDry.Text = PDRY
    txtCiddorWet.Text = " "
    txtCiddorWet.Text = PVAP
End Sub

Private Sub cmdClear_Click()
   If lstSondes.ListCount > 0 Then
      Select Case MsgBox("This will clear the entire file buffer!" _
                         & vbCrLf & "" _
                         & vbCrLf & "Proceed?" _
                         , vbOKCancel Or vbQuestion Or vbDefaultButton1, "Clear file buffer")
      
        Case vbOK
            lstSondes.Clear
        Case vbCancel
        
      End Select
      End If
End Sub

Private Sub cmdCompareTR_Click()

    'open chosen file and read coordinates, etc, and compare the TR vdW calculation to the modified Wikipedia formula for TR
    
   On Error GoTo cmdCompareTR_Click_Error
   
   Dim lg1 As Double, lt1 As Double, kmxo As Double, kmyo As Double, H11 As Double
   Dim lg2 As Double, lt2 As Double, H21 As Double, lg As Double, lt As Double
   Dim MT(12) As Integer, AT(12) As Integer, ier As Integer, FileNameOut As String
   Dim AveMinTmp As Double, AveAvgTmp As Double, azi As Double, VA As Double, kmx As Double, kmy As Double
   Dim distd As Double, deltd As Double, defm As Double, defb As Double, avref As Double
   Dim PATHLENGTH As Double, Press0 As Double, j As Integer, NNN As Integer
   Dim FileMode As Integer, HMAXT As Double, RELHUM As Double, StartAng As Double, EndAng As Double
   Dim WAVELN As Double, OBSLAT As Double, NSTEPS As Long, HUMID As Double, HOBS As Double
   Dim StepSize As Integer, RecordTLoop As Boolean, ier2 As Long, LastVA As Double, NAngles As Long
   Dim DistTolerance As Double, D1 As Double, viewangle As Double, TRRayTrace As Double
   Dim X1 As Double, X2 As Double, Y1 As Double, Y2 As Double, Exponent As Double
   Dim z1 As Double, z2 As Double, re1 As Double, re2 As Double
   Dim dist1 As Double, dist2 As Double, ANGLE As Double, hgtDTM
   Dim MinAzimuth As Double, MaxAzimuth As Double, geo As Boolean
   Dim hgtworld As Double, kcurve As Double ', Rcurve As Double
   
   Rearth = 6356766#
   RE = Rearth
   
    pi = 4# * Atn(1#) '3.141592654
    CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
    cd = pi / 180# 'conversion of degrees to radians
   
'   Dim MinAzimuth As Double, MaxAzimuth As Double

    filein% = FreeFile
    Open txtFileName.Text For Input As #filein%
    Line Input #filein%, doclin$
    If InStr(doclin$, "kmxo") Then
       geo = False
    ElseIf InStr(doclin$, "Lati") Then
       geo = True
    Else
       Call MsgBox("Can't determine if this a geo file or not from the header" _
                   & vbCrLf & "" _
                   & vbCrLf & "Aborting....." _
                   , vbExclamation, "geo coordinates")
       Close
       picProgBar.Visible = False
       Screen.MousePointer = vbDefault
       Exit Sub
       End If
       
    Input #filein%, lg1, lt1, H11, startkmx, sofkmx, dkmx, dkmy, APPRNR
    If Not geo Then
       'EY ITM, convert to geo coordinates
       Call casgeo(lg1, lt1, lg, lt)
       lg1 = -lg
       lt1 = lt
    ElseIf geo Then
       tmplt = lt1
       lt1 = lg1
       lg1 = -tmplt
       End If
       
    'now load up minimum and average world temperatures
    Call Temperatures(lt1, lg1, MT, AT, ier)
    
    'determine solar azimuth range for this latitude
    'at sunirse, sunet, cos(azimuth) = sin(decl)/cos(latitude)
    'declination varies from -23.5 to 23.5 degrees therefore
    MinAzimuth = -DASIN(Sin(23.5 * cd) / Cos(lt1 * cd)) / cd
    MaxAzimuth = -MinAzimuth
    'MaxAzimuth at June 21, Minimum azimuth at Dec 21, zero at Mar 21 and Sep 21 but temperature
    'very different during March through April than from June through October
    'find average mean temperature over the year and use that value
    AveMinTmp = 0
    For i = 1 To 12
       AveMinTmp = MT(i) + AveMinTmp
    Next i
    AveMinTmp = AveMinTmp / 12 + 273.15
    
    If Dir(App.Path & "\CompareTR-7.txt") <> sEmpty Then
       Select Case MsgBox("File: output file ""CompareTR-7.txt "" already exists!" _
                          & vbCrLf & "" _
                          & vbCrLf & "Do you want to copy it to a backup before proceeding?" _
                          , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "File Overwrite protection")
       
        Case vbYes
            FileCopy App.Path & "\CompareTR-7.txt", App.Path & "\CompareTR-7-old.txt"
            
            Call MsgBox("""CompareTR-7.txt"" has been copied to ""CompareTR-7-old.txt""" _
                        & vbCrLf & "" _
                        , vbInformation Or vbDefaultButton1, "File Overwrite protection")
            
        Case vbNo
       
        Case vbCancel
       
       End Select
       End If
    
    fileout% = FreeFile
    Open App.Path & "\CompareTR-7.txt" For Output As #fileout%
    Print #fileout%, "Expected VA (deg.), Old; TR(degrees), Wikipedia; TR(degrees), RayTracing; TR(degrees)"
   
   Screen.MousePointer = vbHourglass
   
    '-------------------------------------------------
    With BARParametersfm
      '------fancy progress bar settings---------
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
        
    Call UpdateStatus(BARParametersfm, picProgBar, 1, 0)
    
    j = 0
    NNN = CInt(2# * Abs(MinAzimuth) / 0.1) + 1
    
    'use this value for comparisons
    Do Until EOF(filein%)
        Input #filein%, azi, VA, kmx, kmy, distd, H21
        If azi < MinAzimuth Or azi > MaxAzimuth Then GoTo 1000
        
        If Not geo Then
           Call heights(kmx, kmy, hgtDTM)
           H21 = hgtDTM
           'now convert to geo coordinates
           Call casgeo(kmx, kmy, lg, lt)
        Else
           Call worldheights(kmx, kmy, hgtworld)
           H21 = hgtworld
           lg = -kmx
           lt = kmy
           End If
           
        D1 = Rearth * DistTrav(lt1, -lg1, lt, lg, 1)
        
        'first caclulate old terrestrial refraction
        deltd = H11 - H21
        If (deltd <= 0#) Then
            defm = 0.000782 - deltd * 0.000000311
            defb = deltd * 0.000034 - 0.0141
        ElseIf (deltd > 0#) Then
            defm = deltd * 0.000000309 + 0.000764
            defb = -0.00915 - deltd * 0.0000269
            End If
        avref = defm * distd + defb
        If (avref < 0#) Then
            avref = 0#
            End If
            
        FilePath = App.Path
        StepSie = 1
        RecordTLoop = False
        FileMode = 1 'mode used for determination of terrestrial refraction using the dll
        
        With prjAtmRefMainfm
            Press0 = Val(.txtPress0)
            HMAXT = Val(.txtHMAXT)
            RELHUM = Val(.txtRELHUM)
            StartAng = Val(.txtBETAHI) * 60# 'convert to arc minutes
            EndAng = Val(.txtBETALO) * 60#
            StepAng = Val(.txtBETAST) * 60#
            WAVELN = Val(.txtKmin) * 0.001 'Val(.txtWAVELN)
            OBSLAT = lt1
            NSTEPS = Val(.txtNSTEPS)
        End With
        
        If NSTEPS < 5000 Then NSTEPS = 5000
        HUMID = RELHUM
        HOBS = H11
        StepSize = Val(prjAtmRefMainfm.txtHeightStepSize.Text)
        NAngles = 2 * StartAng / StepAng + 1
        LastVA = 9999999 'insure proper temperature progression, which should be proportional to the inverse square of the temperature
        DistTolerance = 1
    
        'now calculate estimate of TR using Reijs's formula using the average minimum temperature
'        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rearth * 0.001)) ^ 2#)
'        TR = (0.0083 * PATHLENGTH * Press0) / (AveMinTmp * AveMinTmp)
        
        'use Wikipedia expression instead
        'https://en.wikipedia.org/wiki/Atmospheric_refraction#Terrestrial_refraction
        
'        '//////////////////////begin old version////////////////////////////////
'        'curvature of rays is according to Wikipedia article
'        'https://en.wikipedia.org/wiki/Atmospheric_refraction
'        lR = -0.0065  'K/m  'lapse rate of US standard atmosphere
'        'kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
'        Rcurve = Rearth '/ (1 - kcurve)
'        'use parabolic path length instead of distd
'        'approximate the path length as the ratio of the curvatures
''        PATHLENGTH = distd / (1 - kcurve)
'        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
'        Exponent = 1#  '0.9975
'        TR = 8.15 * (PATHLENGTH ^ Exponent) * 1000 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp) 'arcseconds
'        TR = TR / 3600 'degrees
'        'TR = TR / 1.3195 '/ 1.52 '1.3195 'fudge factor
'        '//////////////////////////////end old version/////////////////////////////////////
        
        lR = -0.0065  'K/m
        'curvature of rays is according to Wikipedia article
        'https://en.wikipedia.org/wiki/Atmospheric_refraction
'        kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
'        Rcurve = Rearth / (1 - kcurve)
        'use parabolic path length instead of distd
'        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
        
        lR = -0.0065  'K/m  'lapse rate of US standard atmosphere
'        kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
        Rcurve = Rearth ' / (1 - kcurve)
        'use parabolic path length instead of distd
        'approximate the path length as the ratio of the curvatures
'        PATHLENGTH = distd / (1 - kcurve)
        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
        PATHLENGTH = PATHLENGTH * 1000 'convert to meters
'        PATHLENGTH = Sqr(2# * Rcurve * Abs(H21 - H11) + (H21 - H11) ^ 2#) 'path length in meters
        If (H21 - H11) > 1000 Then
           Exponent = 0.99 '0.9975  '0.9975
        Else
           Exponent = 0.9965 '1 '0.9945
           End If
        '0.0342 is the lapse rate of an uniformaly dense atmosphere at hydrostatic equilibrium
        'i.e., determines how much would have to decrease the temperature vs height to keep the density constant
        'thereofore, rays wouldn't bend if LR = 0
        TR = 8.15 * (PATHLENGTH ^ Exponent) * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp) 'arcseconds
        TR = TR / 3600 'degrees
        
        'calculate expected viewangle in radians
        GoSub VAsub
        
        BARParametersfm.Refresh
        
        ier2 = RayTracing(StartAng, EndAng, StepAng, LastVA, NAngles, _
                         D1, viewangle, H21, DistTolerance, FileMode, _
                         H11, AveMinTmp, HMAXT, FilePath, StepSize, _
                         Press0, WAVELN, HUMID, OBSLAT, NSTEPS, _
                         RecordTLoop, AveMinTmp, AveMinTmp, AddressOf MyCallback)
        If ier2 = 0 Then
            TRRayTrace = (LastVA - viewangle) / cd 'calculated TR in degrees
        ElseIf ier2 < 0 Then 'didn't converge
            TRRayTrace = 0#
            End If
        
'        Print #fileout%, j, viewangle, avref, TR / 3600, TRRayTrace
        DoEvents
        
        j = j + 1

        Print #fileout%, j, Format(Str(viewangle), "#0.0#####"), Format(Str(avref), "#0.0#####"), Format(Str(TR), "#0.0#####"), Format(Str(TRRayTrace), "#0.0#####")
        
        If j = 15 Then
           ccc = 1
           End If
           
        Call UpdateStatus(BARParametersfm, picProgBar, 1, CLng(100# * j / NNN))
        
1000:
    
    Loop
    
    Close #filein%
    Close #fileout%
    
    Screen.MousePointer = vbDefault
    
    Call UpdateStatus(BARParametersfm, picProgBar, 1, 0)
    
    picProgBar.Visible = False
    
    'now plot the results as function of line number

   On Error GoTo 0
   Exit Sub
   
VAsub:
'    RE = Rearth
    hgt1 = H11
    hgt2 = H21
    X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
    X2 = Cos(lt * cd) * Cos(-lg * cd)
    Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
    Y2 = Cos(lt * cd) * Sin(-lg * cd)
    z1 = Sin(lt1 * cd)
    z2 = Sin(lt * cd)
'    Rearth = 6371315#
    re1 = (hgt1 + RE)
    re2 = (hgt2 + RE)
    X1 = re1 * X1
    Y1 = re1 * Y1
    z1 = re1 * z1
    X2 = re2 * X2
    Y2 = re2 * Y2
    z2 = re2 * z2
    dist1 = re1
    dist2 = re2
    ANGLE = DACOS((X1 * X2 + Y1 * Y2 + z1 * z2) / (dist1 * dist2))
    viewangle = Atn((-re1 + re2 * Cos(ANGLE)) / (re2 * Sin(ANGLE)))
Return

cmdCompareTR_Click_Error:
    Close
    picProgBar.Visible = False
    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdCalculate_Click of Form BARParametersfm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdConvertSonde_Click
' Author    : Dr-John-K-Hall
' Date      : 3/3/2020
' Purpose   : Convert radiosonde text files from University of Wyoming for station 40179 (Beit Dagan)
'             Radiosondes downloaded from following website: http://weather.uwyo.edu/upperair/sounding.html
'---------------------------------------------------------------------------------------
'
Private Sub cmdConvertSonde_Click()

    Dim FileDate As String, StrIdentifierWin As String, StrIdentifierSum As String, StrIdentifier
    Dim lenIS As Integer
    Dim FileOutName As String
    Dim i, jdoc As Integer
    Dim hgt As Double, Temp As Double, Pressure As Double
    
    StrIdentifierSum = "40179 Bet Dagan Observations at 00Z"
    StrIdentifierWin = "40179 Bet Dagan Observations at 06Z"
    lenIS = Len(StrIdentifierSum)

   On Error GoTo cmdConvertSonde_Click_Error
   
   Screen.MousePointer = vbHourglass

   For i = 1 To lstSondes.ListCount
      If lstSondes.Selected(i - 1) = True Then
         'open the file and create atmosphere file for all the 06Z sondes for any date.
         'look for line starting with: 40179 Bet Dagan Observations at 06Z, then read date: 40179 Bet Dagan Observations at 06Z 01 Feb 1996
         filein = FreeFile
         Open lstSondes.List(i - 1) For Input As #filein
         If InStr(lstSondes.List(i - 1), "Jan") Or InStr(lstSondes.List(i - 1), "Feb") Or InStr(lstSondes.List(i - 1), "Nov") Or InStr(lstSondes.List(i - 1), "Dec") Then
            StrIdentifier = StrIdentifierWin
         ElseIf InStr(lstSondes.List(i - 1), "May") Or InStr(lstSondes.List(i - 1), "Jun") Or InStr(lstSondes.List(i - 1), "Jul") Then
            StrIdentifier = StrIdentifierSum
            End If
         Do Until EOF(filein)
            Line Input #filein, doclin$
            If InStr(doclin$, StrIdentifier) Then
               'record date and convert file into atmosphere file
               FileDate = Mid$(doclin$, lenIS + 1, Len(doclin$) - lenIS)
               FileDate = Replace(FileDate, " ", "-") 'fill in spaces within date
               FileDate = Mid$(FileDate, 2, Len(FileDate) - 1)
               'open output file
               FileOutName = DirectOut$ & "\" & FileDate & "-2-sondes.txt"
               fileout = FreeFile
               Open FileOutName For Output As #fileout
               
               'skip 4 header lines
               For jdoc = 1 To 4
                  Line Input #filein, doclin$
               Next jdoc
               
               'start splitting data line
rdline:
               Line Input #filein, doclin$
               
               If Trim$(doclin$) = "" Or InStr(doclin$, "Station information and sounding indices") Then
                  Close #fileout
                  fileout = 0
               Else
                  'process this data line by reading and recording the height, temperature, and pressure
                  If Trim$(Mid$(doclin$, 8, 7)) = vbNullString Or Trim$(Mid$(doclin$, 15, 7)) = vbNullString Then
                     'missing hgt and/or temp data, so skip this sonde
                     Close #fileout
                     Kill FileOutName 'delete this file
                     Exit Do
                     End If
                  Pressure = Val(Mid$(doclin$, 1, 7))
                  hgt = Val(Mid$(doclin$, 8, 7))
                  Temp = Val(Mid$(doclin$, 15, 7))
                  Write #fileout, hgt, Temp, Pressure
                  
                  'read next data line
                  GoTo rdline
                  End If
                  
               End If
         Loop
         End If
      
   Next i
   If fileout > 0 Then Close #fileout
   Close #filein
   Screen.MousePointer = vbDefault
   
   On Error GoTo 0
   Exit Sub

cmdConvertSonde_Click_Error:
    Screen.MousePointer = vbDefault
    Close 'close all the open files
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdConvertSonde_Click of Form BARParametersfm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdFitFiles_Click
' Author    : chaim
' Date      : 5/24/2020
' Purpose   : wite files containing )1) eps and (2) ref values corresponding to every 30 meters in height from 0 to 3000 meters as function of temperature
'             to be used to determine temperature scaling relationshiop
'---------------------------------------------------------------------------------------
'
Private Sub cmdFitFiles_Click()

   On Error GoTo cmdFitFiles_Click_Error

   Dim EPS(21, 101) As Double, ref(21, 101) As Double, Temp As Integer
   Dim eps1 As Double, eps2 As Double, ref1 As Double, ref2 As Double
   Dim hgt1 As Double, hgt2 As Double, hgt As Double, hgtNum As Integer
   Dim NumTemp As Integer, NewHgt As Double
   Dim EPS0(101) As Double, ref0(101) As Double, RefTemp As Boolean
   
   
   For i = 1 To lstSondes.ListCount
   
      If lstSondes.Selected(i - 1) = True Then
         'determine temperature from the name
         
         pos% = InStr(lstSondes.List(i - 1), "TR_")
         Temp = Val(Mid$(lstSondes.List(i - 1), pos% + 7, 3))
         If Temp = 288 Then
            RefTemp = True
         Else
            NumTemp = (Temp - 260) / 3
            RefTemp = False
            End If
         
         'now read the eps and ref data and interpolate the values every 30 m
         filein = FreeFile
         hgt = 0
         hgtNum = 0
         NewHgt = 0
         Open lstSondes.List(i - 1) For Input As #filein
         Line Input #filein, doclin$
         hgt1 = Val(Mid$(doclin$, 29, 11))
         eps1 = Val(Mid$(doclin$, 57, 9))
         ref1 = Val(Mid$(doclin$, 70, 8))
         If Not RefTemp Then
            EPS(NumTemp, hgtNum) = eps1
            ref(NumTemp, hgtNum) = ref1
         Else
            EPS0(hgtNum) = eps1
            ref0(hgtNum) = ref1
            End If
         NewHgt = NewHgt + 30
         
50:
         If EOF(filein) Then GoTo 900
         Line Input #filein, doclin$
         hgt2 = Val(Mid$(doclin$, 29, 11))
         eps2 = Val(Mid$(doclin$, 57, 9))
         ref2 = Val(Mid$(doclin$, 70, 8))
         
         If NewHgt >= hgt1 And NewHgt < hgt2 Then
'            If hgtNum + 1 = 99 And RefTemp Then
'               ccc = 1
'               End If
            If Not RefTemp Then
               EPS(NumTemp, hgtNum + 1) = (NewHgt - hgt1) * (eps2 - eps1) / (hgt2 - hgt1) + eps1
               ref(NumTemp, hgtNum + 1) = (NewHgt - hgt1) * (ref2 - ref1) / (hgt2 - hgt1) + ref1
            Else
               EPS0(hgtNum + 1) = (NewHgt - hgt1) * (eps2 - eps1) / (hgt2 - hgt1) + eps1
               ref0(hgtNum + 1) = (NewHgt - hgt1) * (ref2 - ref1) / (hgt2 - hgt1) + ref1
               End If
            hgtNum = hgtNum + 1
            NewHgt = NewHgt + 30
            hgt1 = hgt2
            eps1 = eps2
            ref1 = ref2
            If NewHgt > 3000 Then GoTo 900 '3000 meters is maximum height to record
            End If
            
         GoTo 50
         
900:
         Close #filein
         End If
   
   Next i
   
   For j = 0 To hgtNum
        k = 30 * j
        fileout = FreeFile
        Open "c:/jk/Druk-Vangeld-Data/RefData_" & Format(Trim$(Str$(k)), "0000") & ".txt" For Output As #fileout
        For i = 0 To NumTemp
'           If j = 0 Then
'              Write #fileout, Log((260 + i * 3) / 288.15), Log(EPS(i, j) / EPS0(j)), 0#
'           Else
'              Write #fileout, Log((260 + i * 3) / 288.15), Log(EPS(i, j) / EPS0(j)), Log(ref(i, j) / ref0(j))
'              End If
           If j = 0 Then
              Write #fileout, 288.15 / (260 + i * 3), EPS(i, j) / EPS0(j), 0#
           Else
              Write #fileout, 288.15 / (260 + i * 3), EPS(i, j) / EPS0(j), ref(i, j) / ref0(j)
              End If
        Next i
        Close #fileout
   Next j
   
   On Error GoTo 0
   Exit Sub

cmdFitFiles_Click_Error:
    Close
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdFitFiles_Click of Form BARParametersfm"
'    Resume
End Sub

Private Sub cmdRedo_Click()
   'redo the comparison now using the final version of the modified Wikipedia TR formula
   
   Dim lg1 As Double, lt1 As Double, kmxo As Double, kmyo As Double, H11 As Double
   Dim lg2 As Double, lt2 As Double, H21 As Double, lg As Double, lt As Double
   Dim MT(12) As Integer, AT(12) As Integer, ier As Integer, FileNameOut As String
   Dim AveMinTmp As Double, azi As Double, VA As Double, kmx As Double, kmy As Double
   Dim distd As Double, deltd As Double, defm As Double, defb As Double, avref As Double
   Dim PATHLENGTH As Double, Press0 As Double, j As Integer, NNN As Integer
   Dim FileMode As Integer, HMAXT As Double, RELHUM As Double, StartAng As Double, EndAng As Double
   Dim WAVELN As Double, OBSLAT As Double, NSTEPS As Long, HUMID As Double, HOBS As Double
   Dim StepSize As Integer, RecordTLoop As Boolean, ier2 As Long, LastVA As Double, NAngles As Long
   Dim DistTolerance As Double, D1 As Double, viewangle As Double, TRRayTrace As Double
   Dim X1 As Double, X2 As Double, Y1 As Double, Y2 As Double
   Dim z1 As Double, z2 As Double, re1 As Double, re2 As Double
   Dim dist1 As Double, dist2 As Double, ANGLE As Double, geo As Boolean
   Dim MinAzimuth As Double, MaxAzimuth As Double
   Dim Rcurve As Double, kcurve As Double, Exponent As Double
   
   Rearth = 6356766#
   RE = Rearth
   geo = False
   
    pi = 4# * Atn(1#) '3.141592654
    CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
    cd = pi / 180# 'conversion of degrees to radians
   
'   Dim MinAzimuth As Double, MaxAzimuth As Double
   
    '-------------------------------------------------
    With BARParametersfm
      '------fancy progress bar settings---------
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      .picProgBar.Visible = True
    End With
    pbScaleWidth = 100
    '-------------------------------------------------
    
    Call UpdateStatus(BARParametersfm, picProgBar, 1, 0)

    filein% = FreeFile
    Open txtFileName.Text For Input As #filein%
    Line Input #filein%, doclin$
    If InStr(doclin$, "kmxo") Then
       geo = False
    ElseIf InStr(doclin$, "Lati") Then
       geo = True
    Else
       Call MsgBox("Can't determine if this a geo file or not from the header" _
                   & vbCrLf & "" _
                   & vbCrLf & "Aborting....." _
                   , vbExclamation, "geo coordinates")
       Close
       picProgBar.Visible = False
       Screen.MousePointer = vbDefault
       Exit Sub
       End If
    
    Input #filein%, lg1, lt1, H11, startkmx, sofkmx, dkmx, dkmy, APPRNR
       
    If Not geo Then
       'EY ITM, convert to geo coordinates
       Call casgeo(lg1, lt1, lg, lt)
       lg1 = -lg
       lt1 = lt
    ElseIf geo Then 'apparently latitude and longitude are switched from the header designation
       tmplt = lt1
       lt1 = lg1
       lg1 = -tmplt
       End If
       
    'now load up minimum and average world temperatures
    Call Temperatures(lt1, lg1, MT, AT, ier)
    
    'determine solar azimuth range for this latitude
    'at sunirse, sunet, cos(azimuth) = sin(decl)/cos(latitude)
    'declination varies from -23.5 to 23.5 degrees therefore
    MinAzimuth = -DASIN(Sin(23.5 * cd) / Cos(lt1 * cd)) / cd
    MaxAzimuth = -MinAzimuth
    'MaxAzimuth at June 21, Minimum azimuth at Dec 21, zero at Mar 21 and Sep 21 but temperature
    'very different during March through April than from June through October
    'find average mean temperature over the year and use that value
    AveMinTmp = 0
    For i = 1 To 12
       AveMinTmp = AT(i) + AveMinTmp
    Next i
    AveMinTmp = AveMinTmp / 12 + 273.15
    
    'compareTR-2 is 599-jure.pr6
    'compareTR-4 is c:/cities/erors/netz/netz/Cuernavaca-NETZ0000.pr1
    'compareTR-5 is c:/cities/eros/netz/skiy/NETZ0000.pr1 'about 80 km east of the Rockies and east of Boulder
    'compareTR-6 is c:/cities/givat_zeev_agan_haayalot_moav/277-MOAV.PR1
    'compareTR-7 is Rav Druk's observation spot at Armon Hanatziv
    
    filein2% = FreeFile
    Open App.Path & "\CompareTR-6.txt" For Input As #filein2% 'must be output file of comCalculate TR raytracing on the corresponding txtFileName file
'    Open App.Path & "\CompareTR - Copy.txt" For Input As #filein2%
    Line Input #filein2%, doclin$

    fileout% = FreeFile
    Open App.Path & "\CompareTR-6-SL-2.txt" For Output As #fileout%
'    Open App.Path & "\CompareTR-Jeru.txt" For Output As #fileout%
  
    Print #fileout%, doclin$
    
    j = 0
    NNN = CInt(2# * Abs(MinAzimuth) / 0.1) + 1
    
    Press0 = Val(prjAtmRefMainfm.txtPress0)
    
    Do Until EOF(filein%)
        Input #filein%, azi, VA, kmx, kmy, distd, H21
        If azi < MinAzimuth Or azi > MaxAzimuth Then GoTo 1000
        
        If Not geo Then
           Call heights(kmx, kmy, hgtDTM)
           H21 = hgtDTM
           'now convert to geo coordinates
        Else
           Call worldheights(kmx, kmy, hgtworld)
           H21 = hgtworld
           End If
        
        Input #filein2%, i, X1, X2, Y1, Y2
        
        lR = -0.0065  'K/m
        'curvature of rays is according to Wikipedia article
        'https://en.wikipedia.org/wiki/Atmospheric_refraction
'        kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
'        Rcurve = Rearth / (1 - kcurve)
        'use parabolic path length instead of distd
'        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
        
        lR = -0.0065  'K/m  'lapse rate of US standard atmosphere
'        kcurve = 503 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp)
        Rcurve = Rearth ' / (1 - kcurve)
        'use parabolic path length instead of distd
        'approximate the path length as the ratio of the curvatures
'        PATHLENGTH = distd / (1 - kcurve)
        PATHLENGTH = Sqr(distd ^ 2# + ((H21 - H11) * 0.001 - 0.5 * (distd ^ 2#) / (Rcurve * 0.001)) ^ 2#)
        PATHLENGTH = PATHLENGTH * 1000 'convert to meters
'        PATHLENGTH = Sqr(2# * Rcurve * Abs(H21 - H11) + (H21 - H11) ^ 2#) 'path length in meters
        If (H21 - H11) > 1000 Then
           Exponent = 0.99 '0.9975  '0.9975
        Else
           Exponent = 0.9965 '1 '0.9945
           End If
        TR = 8.15 * (PATHLENGTH ^ Exponent) * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp) 'arcseconds
        TR = TR / 3600 'degrees
        'TR = TR / 1.3195 '/ 1.52 '1.3195 'fudge factor

'        TR = 8.15 * PATHLENGTH * 1000 * Press0 * (0.0342 + lR) / (AveMinTmp * AveMinTmp) 'arcseconds
'        TR = TR / 3600 'degrees
'        TR = TR / 1.3195 'fudge factor
        
        j = j + 1
        
        Print #fileout%, j, Format(Str(X1), "#0.0#####"), Format(Str(X2), "#0.0#####"), Format(Str(TR), "#0.0#####"), Format(Str(Y2), "#0.0#####")
        
        Call UpdateStatus(BARParametersfm, picProgBar, 1, CLng(100# * j / NNN))
1000:
    Loop
    
    Close
    picProgBar.Visible = False
    
End Sub

Private Sub cmdUnselect_Click()
   For i = 1 To lstSondes.ListCount
      lstSondes.Selected(i - 1) = False
   Next i
End Sub

Private Sub Form_Load()
    Dim dtmdir As String, n%, i&, j&, filnum%
    
    dtmdir = "c:\dtm"
    
    CHMNEO = "XX"
    filnum% = FreeFile
    Open dtmdir & "\dtm-map.loc" For Input As #filnum%
    For i& = 1 To 3
       Line Input #filnum%, doclin$
    Next i&
    n% = 0
    For i& = 4 To 54
       Line Input #filnum%, doclin$
       If i& Mod 2 = 0 Then
          n% = n% + 1
          For j& = 1 To 14
             CHMAP(j&, n%) = Mid$(doclin$, 6 + (j& - 1) * 5, 2)
          Next j&
          End If
    Next i&
    Close #filnum%

    sEmpty = ""
   
   'load in CD # for USGS EROS DEM (tiles are numbered from
   'left to right, top to bottom - see Cds.gif file)
   worldcd%(1) = 1
   worldcd%(2) = 1
   worldcd%(3) = 1
   worldcd%(4) = 1
   worldcd%(5) = 3
   worldcd%(6) = 3
   worldcd%(7) = 3
   worldcd%(8) = 3
   worldcd%(9) = 3
   worldcd%(10) = 1
   worldcd%(11) = 1
   worldcd%(12) = 1
   worldcd%(13) = 2
   worldcd%(14) = 2
   worldcd%(15) = 2
   worldcd%(16) = 3
   worldcd%(17) = 3
   worldcd%(18) = 4
   worldcd%(19) = 4
   worldcd%(20) = 4
   worldcd%(21) = 2
   worldcd%(22) = 2
   worldcd%(23) = 2
   worldcd%(24) = 2
   worldcd%(25) = 4
   worldcd%(26) = 4
   worldcd%(27) = 4
   worldcd%(28) = 5
End Sub

Private Sub optCalculate_Click()
   If optCalculate.Value = True Then
'      cmdConvertSonde.Enabled = False
'      cmdCalcAtms.Enabled = True
'      End If
        With BARParametersfm
           .cmdAddRef.Enabled = False
           .cmdAddVA.Enabled = False
           .cmdCalcAtms.Enabled = True
           .cmdConvertSonde.Enabled = False
           .cmdFitFiles.Enabled = False
        End With
        End If
End Sub

Private Sub optConvert_Click()
   If optConvert.Value = True Then
'      cmdConvertSonde.Enabled = True
'      cmdCalcAtms.Enabled = False
'      End If
        With BARParametersfm
           .cmdAddRef.Enabled = False
           .cmdAddVA.Enabled = False
           .cmdCalcAtms.Enabled = False
           .cmdConvertSonde.Enabled = True
           .cmdFitFiles.Enabled = False
        End With
        End If
End Sub

Private Sub optfit1_Click()
   If optfit1.Value = True Then
        With BARParametersfm
           .cmdAddRef.Enabled = False
           .cmdAddVA.Enabled = False
           .cmdCalcAtms.Enabled = False
           .cmdConvertSonde.Enabled = False
        End With
        End If
End Sub
