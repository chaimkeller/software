VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{825967DA-1756-11D3-B695-ED78B587442C}#30.0#0"; "FlexListBox.ocx"
Begin VB.Form frmSpline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spline and Polynomial Fits"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmSpline.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDetails 
      Caption         =   "Details"
      Height          =   975
      Left            =   1800
      TabIndex        =   25
      Top             =   2640
      Width           =   1695
      Begin VB.ComboBox cmbSpline 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbDeg 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Text            =   "Pick/enter deg."
         ToolTipText     =   "Polynomial or spline degree"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmFitType 
      Caption         =   "Type of Fit"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
      Begin VB.CheckBox chkCurvature 
         Caption         =   "Curvature"
         Height          =   195
         Left            =   440
         TabIndex        =   28
         ToolTipText     =   "Calculate the curvature at x=0 for polyn deg >= 2"
         Top             =   460
         Width           =   1095
      End
      Begin VB.OptionButton optSpline 
         Caption         =   "Spline fit"
         Height          =   195
         Left            =   160
         TabIndex        =   13
         ToolTipText     =   "Fit data to spline"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optPoly 
         Caption         =   "Polynomial Fit"
         Height          =   255
         Left            =   160
         TabIndex        =   12
         ToolTipText     =   "Polynomial least square fit"
         Top             =   220
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame frmSaveFit 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   8160
      Width           =   3495
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Select all the files in the list"
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkWizard 
         Caption         =   "Fit Wizard"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Click to activate the fit wizard"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdFit 
         Caption         =   "Fit & Plot"
         Height          =   375
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add to Plot list"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   610
         Width           =   1335
      End
      Begin VB.CommandButton cmdSavetoDisk 
         Caption         =   "&Save to computer"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         ToolTipText     =   "Save fit file using displayed coeficients"
         Top             =   610
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog comdlgFit 
         Left            =   3000
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame frmFitParam 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3495
      Begin VB.CheckBox chkAuto 
         Caption         =   "Auto Record"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         ToolTipText     =   "Automatically record every fit"
         Top             =   4280
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "Record Fit Details"
         Height          =   260
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Record the fit details"
         Top             =   4240
         Width           =   1695
      End
      Begin VB.Frame frmFitResults 
         Caption         =   "Fit Results"
         Height          =   1895
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   3255
         Begin MSFlexGridLib.MSFlexGrid flxGridFit 
            Height          =   1575
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2778
            _Version        =   393216
            BackColor       =   -2147483624
            BackColorFixed  =   12640511
            GridColor       =   -2147483624
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frmNumPnts 
         Caption         =   "Number points in trend line"
         Height          =   580
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3255
         Begin VB.TextBox txtNumFitPnts 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Text            =   "200"
            Top             =   210
            Width           =   855
         End
      End
      Begin VB.Frame frmType 
         Caption         =   "Fit line plot type and color"
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3255
         Begin VB.ComboBox cmbPlotColor 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Pick a plotting color for the fit"
            Top             =   280
            Width           =   1335
         End
         Begin VB.ComboBox cmbPlotType 
            Height          =   315
            Left            =   165
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Choose a plotting line type forthe fit"
            Top             =   280
            Width           =   1455
         End
      End
      Begin VB.Frame frmXmax 
         Caption         =   "Upper Bound (x)"
         Height          =   855
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   1575
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   320
            Left            =   120
            TabIndex        =   10
            Text            =   "0"
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmXmin 
         Caption         =   "Lower Bound (x)"
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1575
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   320
            Left            =   120
            TabIndex        =   9
            Text            =   "0"
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.Frame frmFile 
      Caption         =   "Select file to fit"
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin FlexList.FlexListBox flxlstFiles 
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3836
         ForeColorUnselected=   0
         BackColorUnselected=   -2147483624
         BackColorSelected=   16384
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSpline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
' Description.....: Programme de visualisation d'interpolation Spline
' Systeme.........: Visual Basic 6.0 sous Windows NT.
' Auteur original.: F. Languasco ®
' Note ...........: Modifié par Cuq pour une visualisation en 3D
'===========================================================================
'
'Option Explicit
'
'Dim NPI&            ' N. de Points dans la courbe.
'Dim Pi() As P_Type  ' Coordonnees des Points de l'interpolation.
'Dim NPC&            ' N. de point approximant la courbe.
'Dim Pc() As P_Type  ' Coordonnees des points pour l'approximation.
'Dim NK&             ' Degree pour la B-Spline.
'Dim VZ&             ' Tension de la courbe T-Spline.
'
'Dim TypeC$          ' Type d'interpolation activée
'
'Dim ShOx!, ShOy!    ' Offset pour le centre de l'indicateur du point ( dépend de l'échelle)
                    
'Dim PSel&           ' Indice du point selectionné
'
'Dim Xmin!, Xmax!    ' Coordonnees minimum et maximum
'Dim Ymin!, Ymax!    ' du quadrillage.
'Dim Zmin!, Zmax!    '
'
'Dim Vue             ' Vue actuelle de visualisation
'                    ' 0 Vue XY
'                    ' 1 Vue XZ
'                    ' 2 Vue YZ
'Dim ResteValue     ' Sauvegarde de la valeur de l'axe 3D non traité selon la vue
'                   ' après Modif dans la grille
''
'Dim DirNome$        ' Repertoire des fichiers Splines.
'Const PExt$ = "dat" ' Extension des fichiers de point
''
''
'Const BZ$ = ""
'Const CS$ = ""
'Const BS$ = "&Degree" & vbNewLine & "2 <= NK <= NPI"
'Const TS$ = "Te&nsion" & vbNewLine & "1 <= VZ <= 100"
''
'Dim RS1&                ' Position pour l'edition
'Dim CS1&                ' de la coordonnees des points
'Dim RS1_O&              ' dans la table
'Dim CS1_O&              '
''
'Dim GrillePoint_Left&      '
'Dim GrillePoint_Top&       '
''
'Dim NoPaint As Boolean  ' Evite de redessiner la courbe si il n'y a pas de modification
''
'Const PCHL& = &HC0FFFF  ' Couleur  de fond pour la valeur actuelle de la position  du curseur.
'
'--- GetLocale: ----------------------------------------------------------------
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
(ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String _
, ByVal cchData As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
'
Private Const LOCALE_SDECIMAL& = &HE
Private Const LOCALE_STHOUSAND& = &HF
Private Const LOCALE_SDATE& = &H1D
Private Const LOCALE_STIME& = &H1E
Private Sub chkAll_Click()
   If chkAll.Value = vbChecked Then 'select all the files in the list
       For I = 0 To frmSpline.flxlstFiles.list.Count - 1
          frmSpline.flxlstFiles.list.item(I + 1).Selected = True
       Next I
   Else 'unselect all the files in the list
       For I = 0 To frmSpline.flxlstFiles.list.Count - 1
          frmSpline.flxlstFiles.list.item(I + 1).Selected = False
       Next I
      End If
   frmSpline.flxlstFiles.Refresh
   DoEvents
End Sub

Private Sub chkWizard_Click()
   If chkWizard.Value = vbChecked Then
      Call MsgBox("1. Select the files in the flex list that you want to fit." _
                  & vbCrLf & "2. Choose which type of fit" _
                  & vbCrLf & "3. Set the plot parameters" _
                  & vbCrLf & "4. Set the degree of the fit" _
                  & vbCrLf & "5. Press the ""Fit"" button" _
                  & vbCrLf & "" _
                  & vbCrLf & "(Hint: If you want to save the fit coeficients, make sure to" _
                  & vbCrLf & "check the automatic save fit checkbox.)" _
                  & vbCrLf & "" _
                  , vbInformation, "Fit wizard")
      frmSpline.flxlstFiles.MultiSelect = True
  Else
      Call MsgBox("The fit wizard has been deactivated." _
                  & vbCrLf & "" _
                  , vbInformation, "Fit wizard")
      
      frmSpline.flxlstFiles.MultiSelect = False
      End If
End Sub

Private Sub cmbDeg_Change()
 If IsNumeric(cmbDeg.Text) Then
    If optSpline.Value = True Then
       'check values
       Select Case SplineType%
          Case 2 'B-splines
          
             If Int(Val(cmbDeg.Text)) < 1 Then 'Or Val(cmbDeg.Text) > numrecords& Then
                Call MsgBox("The smoothing parameter for B-splines must be:" _
                            & vbCrLf & "" _
                            & vbCrLf & "1. greater or equal to 2" _
                            & vbCrLf & "2. less than or equal to the num. of entries." _
                            & vbCrLf & "" _
                            & vbCrLf & "(Hint: enter a new value, and try again.)" _
                            , vbInformation, "Input error")
                            
              cmbDeg.Text = sEmpty
              Exit Sub
              End If
              
           Case 4 'T-splines
          
             If Int(Val(cmbDeg.Text)) < 1 Or Val(cmbDeg.Text) > 100 Then
                Call MsgBox("The smoothing parameter for B-splines must be:" _
                            & vbCrLf & "" _
                            & vbCrLf & "1. greater or equal to 1" _
                            & vbCrLf & "2. less than or equal to 100" _
                            & vbCrLf & "" _
                            & vbCrLf & "(Hint: enter a new value, and try again.)" _
                            , vbInformation, "Input error")
                            
              cmbDeg.Text = sEmpty
              Exit Sub
              End If
              
       End Select
       
       End If
 Else
    Call MsgBox("Only enter numbers", vbInformation, "Input error")
    cmbDeg.Text = sEmpty
    End If
 
End Sub

Private Sub cmbPlotColor_Change()
   FitPlotColor% = cmbPlotColor.ListIndex - 1
End Sub

Private Sub cmbPlotType_Change()
   FitPlotType% = cmbPlotType.ListIndex - 1
End Sub

Private Sub cmbSpline_click()
   SplineType% = cmbSpline.ListIndex
   Select Case SplineType%
      Case 1 'Bezier
         cmbDeg.Visible = False
      Case 2 'B-spline
         cmbDeg.Visible = True
         If SplineDeg% = 0 Then SplineDeg% = 1
         cmbDeg.Text = SplineDeg%
      Case 3 'C-spline
         cmbDeg.Visible = False
      Case 4 'T-spline
         cmbDeg.Visible = True
         If SplineDeg% = 0 Then SplineDeg% = 1
         cmbDeg.Text = SplineDeg%
   End Select
   If SplineDeg% = 0 Then SplineDeg% = 1
   cmbDeg.Text = SplineDeg%
End Sub

Private Sub cmdAdd_Click()

    Dim FileSave$, sPath As String, sTemp As String, MaxDirLen As Integer
    Dim sShortPath As String

    On Error GoTo errhand
    
    With comdlgFit
    
       .CancelError = True
       .FileName = App.Path & "\*.fit"
       .Filter = "fit files (*.fit)|*.fit|text files (.txt)|*.txt|all files (*.*)|*.*"
       .ShowOpen
       FileSave$ = .FileName
   
    End With
    
    If Dir(FileSave$) <> sEmpty Then
    
       NumAddSave = NumAddSave + 1
    
       ReDim Preserve FileAddSave(NumAddSave - 1)
           
       FileAddSave(NumAddSave - 1) = FileSave$
       End If
    
    
'    If Dir(FileSave$) <> sEmpty Then
'
'       ReDim Preserve Files(UBound(Files) + 1)
'       Files(UBound(Files)) = FileSave$
'
'       sTemp = FileRoot(FileSave$)
'
'       sPath = Mid$(FileSave$, 1, Len(FileSave$) - Len(sTemp))
'
'       MaxDirLen = Int(flxlstFiles.Width / 70) - 30
'
'       Call ShortPath(sPath, MaxDirLen, sShortPath)
'
'       ReDim Preserve flxFileBuffer(UBound(flxFileBuffer) + 1)
'       flxFileBuffer(UBound(flxFileBuffer)) = sShortPath & sTemp
''       frmSetCond.flxlstFiles.AddItem sShortPath & sTemp
''       frmSetCond.Refresh
''       DoEvents
'
'       'don't add plot information
'       numfiles% = numfiles% + 1
'       ReDim Preserve Files(numfiles%)
'       Files(numfiles% - 1) = FitFileName
'       ReDim Preserve PlotInfo(7, numfiles%)
'
'       'add plot information, i.e., two columns X,Y no headers = format #2
'       PlotInfo(0, numfiles% - 1) = 1
'       PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
'       PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex 'Str$(chkColor%)
'       PlotInfo(3, numfiles% - 1) = "1" 'txtXA
'       If Val(PlotInfo(3, numfiles% - 1)) = 0 Then
'           PlotInfo(3, numfiles% - 1) = "1.0"
'           End If
'       PlotInfo(4, numfiles% - 1) = "0" 'txtXB
'       PlotInfo(5, numfiles% - 1) = "1" 'txtYA
'       If Val(PlotInfo(5, numfiles% - 1)) = 0 Then
'           PlotInfo(5, numfiles% - 1) = "1.0"
'           End If
'       PlotInfo(6, numfiles% - 1) = "0"
'       PlotInfo(7, numfiles% - 1) = FitFileName 'PlotInfofrm.lblFileName
'
'       frmSetCond.flxlstFiles.Refresh
'       DoEvents
'
'       End If
    
Exit Sub
errhand:

End Sub

Private Sub cmdFit_Click()
  Dim Xvalue As Double, Yvalue As Double
  Dim FitXStep As Double, XFit As Double, YFit As Double, NumFitSteps As Long, K As Integer
  
  Dim pos As Long
  Dim buff As String
  Dim sLongname As String
  Dim sShortname As String
  
  Dim sTemp As String, sPath As String, sShortPath As String
  Dim numlist%, pos1%, MultiSelectPath As Boolean, sPath0 As String
  Dim sDriveLetter As String, MaxDirLen As Integer, PlotFileName$
  Dim numToFit As Integer, ier As Integer
  
'  On Error GoTo cmdFit_Click_Error

  numlist% = 0
  MaxDirLen = Int(flxlstFiles.Width / 70) - 30
  
  If flxlstFiles.MultiSelect Then GoTo WizardSection
  
  found% = 0
  For I = 0 To flxlstFiles.list.Count - 1
      If flxlstFiles.list.item(I + 1).Selected Then
           found% = 1
           If optPoly.Value = True Then
           
                 If PlotInfo(3, I) = "" Then
                 
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before fitting)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   
                   Exit Sub
                   End If
                   
                 testfile% = 1
                 If UBound(dPlot, 3) = 0 Then
                   testfile% = 0
                   Select Case MsgBox("Can't find the chosen file's data in the data buffer!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before fitting)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to plot this file at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   
                   Exit Sub
                   End If
                     
                testfile% = 0
                HasSolution = False
                Set PtX = New Collection
                Set PtY = New Collection
                     
                PlotFileName$ = FileRoot(flxlstFiles.list.item(I + 1).Text)
                
                'select corresponding file in the frmsetcond flxlist
                found1% = 0
                For J = 0 To frmSpline.flxlstFiles.list.Count - 1
                    If frmSpline.flxlstFiles.list.item(J + 1).Selected Then
                       If FileRoot(PlotInfo(7, J)) = PlotFileName$ And found1% = 0 Then
                          frmSetCond.flxlstFiles.list.item(J + 1).Selected = True
                          found1% = 1
                          CurrentFitFileIndex = J + 1
                          Exit For
                          End If
                      End If
                Next J
                
                If found1% = 0 Then
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before fitting)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   Exit Sub
                Else
                   numSelected% = 1
                   frmSetCond.flxlstFiles.Refresh
                   DoEvents
                   End If
                   
                Fitting = True
                  
                freefil% = FreeFile
                Open PlotInfo(7, CurrentFitFileIndex - 1) For Input As #freefil%
                'skip the header lines
                For idoc = 1 To FilForm(0, PlotInfo(0, I))
                   Line Input #freefil%, doclin$
                Next idoc
        
                numRows% = 0
                MaxX = -999999
                MinX = -MaxX
                numToFit = 0
                Do Until EOF(freefil%)
                   Line Input #freefil%, doclin$
                   numRows% = numRows% + 1
                   Xvalue = dPlot(I, 0, numRows% - 1)
                   Yvalue = dPlot(I, 1, numRows% - 1)
                   If Xvalue >= Val(Text1.Text) And Xvalue <= Val(Text2.Text) Then
                      If Xvalue > MaxX Then MaxX = Xvalue
                      If Xvalue < MinX Then MinX = Xvalue
                      PtX.Add Xvalue
                      PtY.Add Yvalue
                      numToFit = numToFit + 1
                      End If
                Loop
                
                Text1 = MinX
                Text2 = MaxX
                If Val(txtNumFitPnts.Text) = 0 Then txtNumFitPnts.Text = numToFit 'numRows%
                
                Close #freefil%
               
                 ' Find a good fit.
                If IsNumeric(cmbDeg.Text) Then
                    degree = Val(cmbDeg.Text) 'cmbDeg.ListIndex + 1
                Else
                   cmbDeg.Text = "1"
                   degree = 1
                   End If
                   
                If degree = 0 Then degree = 1
                PolyDeg = degree
                Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree, ier)
                If ier = -1 Then Exit Sub
            
                With flxGridFit
                    .Rows = PolyDeg + 2
                    If cmbDeg.Text >= 2 And chkCurvature.Value = vbChecked Then .Rows = PolyDeg + 3
                    .Cols = 2
                    .ColAlignment(0) = 1
                    .ColAlignment(1) = 0
                    .Visible = True
                    
                    .TextMatrix(0, 0) = "Degree"
                    .TextMatrix(0, 1) = "Coeficients"
                    
                    .TextMatrix(0, 0) = "Degree"
                     For J = 0 To PolyDeg
                       .TextMatrix(J + 1, 0) = Str(J) & "th degree"
                       .TextMatrix(J + 1, 1) = sEmpty
                    Next J
                    
                    For J = 0 To BestCoeffs.Count - 1
                       .TextMatrix(J + 1, 1) = Format(BestCoeffs.item(J + 1), "0.00000E+00")
                    Next J
                    
                End With
                
                Call AutosizeGridColumns(flxGridFit, degree + 2, 100)
                
                'now overplot the file and the fit
                'create temperorary two column fit file with the requested number of points
                'in case calcualting curvature, determine the peak position of the fit
                Dim YFitMax As Double
                Dim XFitMax As Double
                YFitMax = -9999
                NumFitSteps = Val(txtNumFitPnts.Text)
                FitXStep = (MaxX - MinX) / (NumFitSteps - 1)
                filetmp% = FreeFile
                FitFileName = App.Path & "\tmp-fit.txt"
                Open FitFileName For Output As #filetmp%
                For J = 1 To NumFitSteps
                   XFit = MinX + (J - 1) * FitXStep
                   YFit = 0#
                   For K = 0 To BestCoeffs.Count - 1
                       YFit = YFit + BestCoeffs.item(K + 1) * XFit ^ K
                   Next K
                   Print #filetmp%, XFit, YFit
                   'keep track of XFit position of maximum YFit
                   If YFit > YFitMax Then
                      YFitMax = YFit
                      XFitMax = XFit
                      End If
                Next J
                Close #filetmp%
                
                If cmbDeg.Text >= 2 And chkCurvature.Value = vbChecked Then
                   With flxGridFit
                       .TextMatrix(PolyDeg + 2, 0) = "Curva(X:" & Format(Str$(XFitMax), "0.00000E+00") & ")"
                       .TextMatrix(PolyDeg + 2, 1) = sEmpty
                        'calculate second derivate at XFitMax, 0th and 1st order terms don't contribute
                        YFit = 0#
                        For K = 2 To BestCoeffs.Count - 1
                           YFit = YFit + K * (K - 1) * BestCoeffs.item(K + 1) * XFitMax ^ (K - 2)
                        Next K
                        .TextMatrix(PolyDeg + 2, 1) = Format(Str$(YFit), "0.00000E+00")
                   End With
                   End If
                
                If chkAuto.Value = vbChecked Then
                   'automatically record the fit results
                   cmdRecord.Value = True
                   End If
                
                'now overplot the file and its fit file
                'first add this file to the flxlstFiles
                'then add plot information in the plot arrays, then give command to make the overplot
                
                'Show the members of the returned sFile string
                'It path is larger than 3 characters, only show
                'most inner path (but record the rest in the file buffer)
      
                sTemp = FileRoot(FitFileName)
                sPath = App.Path
      
                'don't change last recorded plot directory, since this is just a tmp file
      
                'determine short form of this path consisting of
                'the innermost directory
                
                If notAlreadyFitted Then
                  Call ShortPath(sPath, MaxDirLen, sShortPath)
        
                  frmSetCond.flxlstFiles.AddItem sShortPath & sTemp
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
                  numfiles% = numfiles% + 1
                  ReDim Preserve Files(numfiles%)
                  Files(numfiles% - 1) = FitFileName
                  'highlight it
                  frmSetCond.flxlstFiles.list.item(numfiles%).Selected = True
                  notAlreadyFitted = False 'flag so won't repeat adding fit file to plot buffer
                  
                  'add plot information, it is always two column numbers separated by spaces and no headers (format 2)
                    
                  PlotInfo(0, numfiles% - 1) = 1
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  PlotInfo(3, numfiles% - 1) = "1" 'txtXA
                  If Val(PlotInfo(3, numfiles% - 1)) = 0 Then
                     PlotInfo(3, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(4, numfiles% - 1) = "0" 'txtXB
                  PlotInfo(5, numfiles% - 1) = "1" 'txtYA
                  If Val(PlotInfo(5, numfiles% - 1)) = 0 Then
                     PlotInfo(5, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(6, numfiles% - 1) = "0"
                  PlotInfo(7, numfiles% - 1) = FitFileName 'PlotInfofrm.lblFileName
                  PlotInfo(8, numfiles% - 1) = sEmpty
                  PlotInfo(9, numfiles% - 1) = "1"
                    
                  numFilesToPlot% = 2
                  numSelected% = 2
                      
                  frmSpline.Refresh 'return focus to the spline form
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
                Else
                  'detect changes to plot type or plot color
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  End If
                
                Load frmDraw
                frmDraw.Show vbModeless
                frmDraw.Width = Screen.Width - frmSetCond.Width
                frmDraw.ScaleHeight = frmSetCond.ScaleHeight
                frmDraw.Left = frmSetCond.Width
                If frmSetCond.txtXTitle.Text = "" Then frmSetCond.txtXTitle.Text = "X-values"
                If frmSetCond.txtYTitle.Text = "" Then frmSetCond.txtYTitle.Text = "Y-values"
'                If frmSetCond.txtTitle.Text = "" Then frmSetCond.txtTitle.Text = "Chart Title (blank)"
                frmSetCond.cmdOK = True
                PlotAll = False 'reset plt all flag
                
                Exit For
                
                
            ElseIf optSpline.Value = True Then
            
                 If PlotInfo(3, I) = "" Then
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before fitting)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   Exit Sub
                   End If
                   
                 testfile% = 1
                 If UBound(dPlot, 3) = 0 Then
                   testfile% = 0
                   Select Case MsgBox("Can't find the chosen file's data in the data buffer!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before fitting)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to plot this file at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   
                   Exit Sub
                   End If
                     
                textfile% = 0
                PlotFileName$ = PlotInfo(7, I) 'FileRoot(flxlstFiles.List.item(I + 1).Text)
                
                'select corresponding file in the frmsetcond flxlist
                found1% = 0
                For J = 0 To frmSpline.flxlstFiles.list.Count - 1
                    If frmSpline.flxlstFiles.list.item(J + 1).Selected Then
                       If PlotInfo(7, J) = PlotFileName$ Then
                          frmSetCond.flxlstFiles.list.item(J + 1).Selected = True
                          found1% = 1
                          CurrentFitFileIndex = J + 1
                          Exit For
                          End If
                      End If
                Next J
                
                If found1% = 0 Then
                      Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                       & vbCrLf & "(Apparently you haven't plotted it yet before fitting)" _
                                       & vbCrLf & "" _
                                       & vbCrLf & "Do you wish to add format information to it at this time?" _
                                       , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                    
                       Case vbOK
                          Unload Me
                       Case vbCancel
                          Exit Sub
                      End Select
                      Exit Sub
                   Else
                   numSelected% = 1
                   frmSetCond.flxlstFiles.Refresh
                   DoEvents
                   End If
                
                Screen.MousePointer = vbHourglass
                
                Fitting = True
                  
                freefil% = FreeFile
                Open PlotFileName$ For Input As #freefil%
                'skip the header lines
                For idoc = 1 To FilForm(0, PlotInfo(0, I))
                   Line Input #freefil%, doclin$
                Next idoc
        
                numRows% = 0
                MaxX = -999999
                MinX = -MaxX
                MaxY = MaxX
                MinY = MinX
                Do Until EOF(freefil%)
                   Line Input #freefil%, doclin$
                   numRows% = numRows% + 1
                   ReDim Preserve Pi(0 To numRows% - 1)
                   Xvalue = dPlot(I, 0, numRows% - 1)
                   Yvalue = dPlot(I, 1, numRows% - 1)
                   Pi(numRows% - 1).X = Xvalue
                   Pi(numRows% - 1).Y = Yvalue
                   Pi(numRows% - 1).Z = 0#
                   If Xvalue > MaxX Then MaxX = Xvalue
                   If Xvalue < MinX Then MinX = Xvalue
                   If Yvalue > MaxY Then MaxY = Yvalue
                   If Yvalue < MinY Then MinY = Yvalue
                Loop
                
                Text1 = MinX
                Text2 = MaxX
                NPI = numRows%
                If Trim(txtNumFitPnts.Text) = sEmpty Then
                   NPC = NPI
                   txtNumFitPnts.Text = NPC
                Else
                   NPC = Val(txtNumFitPnts.Text)
                   End If
                NumFitSteps = NPC
                   
                Close #freefil%
               
                ' Find a good fit.
                'execute spline fit
                'results reside in array Pc(0 to NPC - 1)
            '
                ReDim Pc(0 To NPC - 1) ' As P_Type
                
                Select Case SplineType%
                   Case 1 'Bezier
                      If NPI > 1029 Then
                      
                         Call MsgBox("Bezier spline limited to 1029 points," _
                                     & vbCrLf & "but your file has" & Str(NPC) & " points" _
                                     & vbCrLf & "" _
                                     , vbInformation, "File too big")
                         
                         Screen.MousePointer = vbDefault
                         Exit Sub
                         End If
                      Call Bezier(Pi(), Pc())
                   Case 2 'B-spline
                      NK = CLng(Val(cmbDeg.Text))
                      Call B_Spline(Pi(), NK, Pc())
                   Case 3 'C-spline
                      Call C_Spline(Pi(), Pc())
                   Case 4 'T-spline
                      VZ = CLng(Val(cmbDeg.Text))
                      Call T_Spline(Pi(), VZ, Pc())
                End Select
                
                'now overplot the file and the fit
                'create temperorary two column fit file with the requested number of points

                filetmp% = FreeFile
                FitFileName = App.Path & "\tmp-fit.txt"
                Open FitFileName For Output As #filetmp%
                For J = 1 To UBound(Pc) + 1
                   Print #filetmp%, Pc(J - 1).X, Pc(J - 1).Y
                Next J
                Close #filetmp%
                
                'now overplot the file and its fit file
                'first add this file to the flxlstFiles
                'then add plot information in the plot arrays, then give command to make the overplot
                
                'Show the members of the returned sFile string
                'It path is larger than 3 characters, only show
                'most inner path (but record the rest in the file buffer)
      
                sTemp = FileRoot(FitFileName)
                sPath = App.Path
      
                'don't change last recorded plot directory, since this is just a tmp file
      
                'determine short form of this path consisting of
                'the innermost directory
                
                If notAlreadyFitted Then
                  Call ShortPath(sPath, MaxDirLen, sShortPath)
        
                  frmSetCond.flxlstFiles.AddItem sShortPath & sTemp
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
                  numfiles% = numfiles% + 1
                  ReDim Preserve Files(numfiles%)
                  Files(numfiles% - 1) = FitFileName
                  'highlight it
                  frmSetCond.flxlstFiles.list.item(numfiles%).Selected = True
                  notAlreadyFitted = False 'flag so won't repeat adding fit file to plot buffer
                  
                  'add plot information, it is always two column numbers separated by spaces and no headers (format 2)
                    
                  PlotInfo(0, numfiles% - 1) = 1
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  PlotInfo(3, numfiles% - 1) = "1" 'txtXA
                  If Val(PlotInfo(3, numfiles% - 1)) = 0 Then
                     PlotInfo(3, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(4, numfiles% - 1) = "0" 'txtXB
                  PlotInfo(5, numfiles% - 1) = "1" 'txtYA
                  If Val(PlotInfo(5, numfiles% - 1)) = 0 Then
                     PlotInfo(5, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(6, numfiles% - 1) = "0"
                  PlotInfo(7, numfiles% - 1) = FitFileName 'PlotInfofrm.lblFileName
                  PlotInfo(8, numfiles% - 1) = sEmpty
                  PlotInfo(9, numfiles% - 1) = "1"
                    
                  numFilesToPlot% = 2
                  numSelected% = 2
                      
                  frmSpline.Refresh 'return focus to the spline form
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
                Else
                  'detect changes to plot type or plot color
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  End If
                  
                Load frmDraw
                frmDraw.Show vbModeless
                frmDraw.Width = Screen.Width - frmSetCond.Width
                frmDraw.ScaleHeight = frmSetCond.ScaleHeight
                frmDraw.Left = frmSetCond.Width
                If frmSetCond.txtXTitle.Text = "" Then frmSetCond.txtXTitle.Text = "X-values"
                If frmSetCond.txtYTitle.Text = "" Then frmSetCond.txtYTitle.Text = "Y-values"
'                If frmSetCond.txtTitle.Text = "" Then frmSetCond.txtTitle.Text = "Chart Title (blank)"
                frmSetCond.cmdOK = True
                PlotAll = False 'reset plt all flag
                
                Screen.MousePointer = vbDefault
                Exit For
            
            Else
            
               Call MsgBox("You need to choose a fit option!", vbInformation, "Fit option")
               Exit Sub
            
            End If
            
            Exit For
            
        End If
        
   Next I
   
   If found% = 0 Then
      Call MsgBox("Please select a fit file in the list", vbInformation, "No fit file selected")
      End If
      
   Fitting = False
      
Exit Sub

'/////////////////////////////FIT WIZARD SECTION//////////////////////////////////////////////////////

WizardSection:

  If flxlstFiles.list.Count = 0 Then
     MsgBox "Plot buffer is empty!", vbExclamation + vbOKOnly, "Plot"
     Exit Sub
     End If
     
  flxlstFiles.MultiSelect = True
  
  FitWizard = True
  
'  maxFilesToPlot% = MaxNumOverplotFiles
'  numFilesToPlot% = 0
  
  If flxlstFiles.list.Count > 15 Then
    'scroll the flex list box to the beginning
    For I = 1 To flxlstFiles.list.Count
       flxlstFiles.SetFocus
       waitime = Timer
       Do Until Timer > waitime + 0.001
          DoEvents
       Loop
       Call keybd_event(VK_UP, 0, 0, 0)
       Call keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)
    Next I
    End If
           
  numSelected% = -1
  For I = 0 To flxlstFiles.list.Count - 1
  
        '==============do cleanup of file list============================
        If InStr(frmSetCond.flxlstFiles.list.item(numfiles%).Text, "tmp-fit.txt") Then
            
            If FitFileName <> "" And Dir(FitFileName) <> sEmpty Then
               Kill FitFileName
               End If
               
            numfiles% = numfiles% - 1
            ReDim Preserve Files(numfiles%)
               
            ReDim Preserve PlotInfo(9, numfiles%)
            
            notAlreadyFitted = True
    
            End If
            
        frmSetCond.flxlstFiles.Clear
        
        'restore list
        For J = 0 To UBound(flxFileBuffer)
            frmSetCond.flxlstFiles.AddItem flxFileBuffer(J)
        Next J
        
        frmSetCond.flxlstFiles.Refresh
        DoEvents
                
        'restore plot buffer number
        numFilesToPlot% = OriginalNumPlotFiles
        '====================================================================
    
        'scroll down one
        flxlstFiles.SetFocus
        waitime = Timer
        Do Until Timer > waitime + 0.001
           DoEvents
        Loop
        flxlstFiles.SetFocus
        Call keybd_event(VK_DOWN, 0, 0, 0)
        Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
        
        If flxlstFiles.list.item(I + 1).Selected Then
            'scroll down one item
            flxlstFiles.SetFocus
            waitime = Timer
            Do Until Timer > waitime + 0.001
               DoEvents
            Loop
            flxlstFiles.SetFocus
            Call keybd_event(VK_DOWN, 0, 0, 0)
            Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
            
'           numSelected% = I
'           numFilesToPlot% = 2 'numFilesToPlot% + 1 'so far software can only overplot the fitting and fit files
'           If I > maxFilesToPlot% Then
'              MsgBox "You are allowed to overplot a maximum of " & _
'              Str(maxFilesToPlot%) & "files!", vbExclamation + vbOKOnly, "Plot"
'              Exit For
'              End If
           tmpBackColor& = flxlstFiles.list.item(I + 1).ItemBackColor
           tmpForeColor& = flxlstFiles.list.item(I + 1).ItemForeColor
           Call keybd_event(VK_DOWN, 0, 0, 0)
           Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
           flxlstFiles.list.item(I + 1).ItemBackColor = &HFF&
           flxlstFiles.list.item(I + 1).ItemForeColor = QBColor(0)
           flxlstFiles.ListIndex = I
           flxlstFiles.Refresh
           DoEvents
           
           If optPoly.Value = True Then
           
                 If PlotInfo(3, I) = "" Then
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before entering splines)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   Exit Sub
                   End If
                   
                 testfile% = 1
                 If UBound(dPlot, 3) = 0 Then
                   testfile% = 0
                   Select Case MsgBox("Can't find the chosen file's data in the data buffer!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before entering splines)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to plot this file at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   
                   Exit Sub
                   End If
                     
                textfile% = 0
                     
                HasSolution = False
                Set PtX = New Collection
                Set PtY = New Collection
                     
                PlotFileName$ = FileRoot(flxlstFiles.list.item(I + 1).Text)
                
                'select corresponding file in the frmsetcond flxlist
                found1% = 0
                For J = 0 To frmSpline.flxlstFiles.list.Count - 1
                    If frmSpline.flxlstFiles.list.item(J + 1).Selected Then
                       If FileRoot(PlotInfo(7, J)) = PlotFileName$ And found1% = 0 Then
                          frmSetCond.flxlstFiles.list.item(J + 1).Selected = True
                          found1% = 1
                          CurrentFitFileIndex = J + 1
                          Exit For
                          End If
                      End If
                Next J
                
                If found1% = 0 Then
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before entering splines)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   Exit Sub
                Else
                   numSelected% = 1
                   frmSetCond.flxlstFiles.Refresh
                   DoEvents
                   End If
                   
                Fitting = True
                  
                freefil% = FreeFile
                Open PlotInfo(7, CurrentFitFileIndex - 1) For Input As #freefil%
                'skip the header lines
                For idoc = 1 To FilForm(0, PlotInfo(0, I))
                   Line Input #freefil%, doclin$
                Next idoc
        
                numRows% = 0
                MaxX = -999999
                MinX = -MaxX
                Do Until EOF(freefil%)
                   Line Input #freefil%, doclin$
                   numRows% = numRows% + 1
                   Xvalue = dPlot(I, 0, numRows% - 1)
                   Yvalue = dPlot(I, 1, numRows% - 1)
                   If Xvalue > MaxX Then MaxX = Xvalue
                   If Xvalue < MinX Then MinX = Xvalue
                   PtX.Add Xvalue
                   PtY.Add Yvalue
                Loop
                
                Text1 = MinX
                Text2 = MaxX
'                If Trim(txtNumFitPnts.Text) = sEmpty Then txtNumFitPnts.Text = numRows%
                txtNumFitPnts.Text = numRows%
                
                Close #freefil%
               
                 ' Find a good fit.
                If IsNumeric(cmbDeg.Text) Then
                    degree = Val(cmbDeg.Text) 'cmbDeg.ListIndex + 1
                Else
                    cmbDeg.Text = "1"
                    degree = 1
                    End If
                    
                If degree = 0 Then degree = 1
                PolyDeg = degree
                Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree, ier)
                If ier = -1 Then Exit Sub
            
                With flxGridFit
                    .Rows = PolyDeg + 2
                    .Cols = 2
                    .ColAlignment(0) = 1
                    .ColAlignment(1) = 0
                    .Visible = True
                    
                    .TextMatrix(0, 0) = "Degree"
                    .TextMatrix(0, 1) = "Coeficients"
                    
                    .TextMatrix(0, 0) = "Degree"
                     For J = 0 To PolyDeg
                       .TextMatrix(J + 1, 0) = Str(J) & "th degree"
                       .TextMatrix(J + 1, 1) = sEmpty
                    Next J
                    
                    For J = 0 To BestCoeffs.Count - 1
                       .TextMatrix(J + 1, 1) = Format(BestCoeffs.item(J + 1), "0.00000E+00")
                    Next J
                    
                End With
                
                Call AutosizeGridColumns(flxGridFit, degree + 2, 100)
                
                'now overplot the file and the fit
                'create temperorary two column fit file with the requested number of points
                NumFitSteps = Val(txtNumFitPnts.Text)
                FitXStep = (MaxX - MinX) / (NumFitSteps - 1)
                filetmp% = FreeFile
                FitFileName = App.Path & "\tmp-fit.txt"
                Open FitFileName For Output As #filetmp%
                For J = 1 To NumFitSteps
                   XFit = MinX + (J - 1) * FitXStep
                   YFit = 0#
                   For K = 0 To BestCoeffs.Count - 1
                       YFit = YFit + BestCoeffs.item(K + 1) * XFit ^ K
                   Next K
                   Print #filetmp%, XFit, YFit
                Next J
                Close #filetmp%
                
                If chkAuto.Value = vbChecked Then
                   'automatically record the fit results
                   cmdRecord.Value = True
                   End If
                
                'now overplot the file and its fit file
                'first add this file to the flxlstFiles
                'then add plot information in the plot arrays, then give command to make the overplot
                
                'Show the members of the returned sFile string
                'It path is larger than 3 characters, only show
                'most inner path (but record the rest in the file buffer)
      
                sTemp = FileRoot(FitFileName)
                sPath = App.Path
      
                'don't change last recorded plot directory, since this is just a tmp file
      
                'determine short form of this path consisting of
                'the innermost directory
                
                If notAlreadyFitted Then
                  Call ShortPath(sPath, MaxDirLen, sShortPath)
        
                  frmSetCond.flxlstFiles.AddItem sShortPath & sTemp
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
                  numfiles% = numfiles% + 1
                  ReDim Preserve Files(numfiles%)
                  Files(numfiles% - 1) = FitFileName
                  'highlight it
                  frmSetCond.flxlstFiles.list.item(numfiles%).Selected = True
                  notAlreadyFitted = False 'flag so won't repeat adding fit file to plot buffer
                  
                  'add plot information, it is always two column numbers separated by spaces and no headers (format 2)
                    
                  PlotInfo(0, numfiles% - 1) = 1
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  PlotInfo(3, numfiles% - 1) = "1" 'txtXA
                  If Val(PlotInfo(3, numfiles% - 1)) = 0 Then
                     PlotInfo(3, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(4, numfiles% - 1) = "0" 'txtXB
                  PlotInfo(5, numfiles% - 1) = "1" 'txtYA
                  If Val(PlotInfo(5, numfiles% - 1)) = 0 Then
                     PlotInfo(5, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(6, numfiles% - 1) = "0"
                  PlotInfo(7, numfiles% - 1) = FitFileName 'PlotInfofrm.lblFileName
                  PlotInfo(8, numfiles% - 1) = sEmpty
                  PlotInfo(9, numfiles% - 1) = "1"
                    
                  numFilesToPlot% = 2
                  numSelected% = 2
                      
                  frmSpline.Refresh 'return focus to the spline form
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
               Else
                  'detect changes to plot type and plot color
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  
'                  numFilesToPlot% = 2
'                  numSelected% = 2
                 
                  End If
                
                Load frmDraw
                frmDraw.Show vbModeless
                frmDraw.Width = Screen.Width - frmSetCond.Width
                frmDraw.ScaleHeight = frmSetCond.ScaleHeight
                frmDraw.Left = frmSetCond.Width
                If frmSetCond.txtXTitle.Text = "" Then frmSetCond.txtXTitle.Text = "X-values"
                If frmSetCond.txtYTitle.Text = "" Then frmSetCond.txtYTitle.Text = "Y-values"
'                If frmSetCond.txtTitle.Text = "" Then frmSetCond.txtTitle.Text = "Chart Title (blank)"
                frmSetCond.cmdOK = True
                PlotAll = False 'reset plt all flag
                
                'now wait a bit to show the fit before going on
                
                waittime = Timer
                Do Until Timer > waittime + 1
                   DoEvents
                Loop
                
            ElseIf optSpline.Value = True Then
            
           
                 If PlotInfo(3, I) = "" Then
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before entering splines)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   Exit Sub
                   End If
                   
                 testfile% = 1
                 If UBound(dPlot, 3) = 0 Then
                   testfile% = 0
                   Select Case MsgBox("Can't find the chosen file's data in the data buffer!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before entering splines)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to plot this file at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   
                   Exit Sub
                   End If
                     
                textfile% = 0
                     
                PlotFileName$ = FileRoot(flxlstFiles.list.item(I + 1).Text)
                
                'select corresponding file in the frmsetcond flxlist
                found1% = 0
                For J = 0 To frmSpline.flxlstFiles.list.Count - 1
                    If frmSpline.flxlstFiles.list.item(J + 1).Selected Then
                       If FileRoot(PlotInfo(7, J)) = PlotFileName$ Then
                          frmSetCond.flxlstFiles.list.item(J + 1).Selected = True
                          found1% = 1
                          CurrentFitFileIndex = J + 1
                          Exit For
                          End If
                      End If
                Next J
                
                If found1% = 0 Then
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before entering splines)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   Exit Sub
                Else
                   numSelected% = 1
                   frmSetCond.flxlstFiles.Refresh
                   DoEvents
                   End If
                   
                Screen.MousePointer = vbHourglass
                
                Fitting = True
                  
                freefil% = FreeFile
                Open PlotFileName$ For Input As #freefil%
                'skip the header lines
                For idoc = 1 To FilForm(0, PlotInfo(0, I))
                   Line Input #freefil%, doclin$
                Next idoc
        
                numRows% = 0
                MaxX = -999999
                MinX = -MaxX
                MaxY = MaxX
                MinY = MinX
                Do Until EOF(freefil%)
                   Line Input #freefil%, doclin$
                   numRows% = numRows% + 1
                   ReDim Preserve Pi(0 To numRows% - 1)
                   Xvalue = dPlot(I, 0, numRows% - 1)
                   Yvalue = dPlot(I, 1, numRows% - 1)
                   Pi(numRows% - 1).X = Xvalue
                   Pi(numRows% - 1).Y = Yvalue
                   Pi(numRows% - 1).Z = 0#
                   If Xvalue > MaxX Then MaxX = Xvalue
                   If Xvalue < MinX Then MinX = Xvalue
                   If Yvalue > MaxY Then MaxY = Yvalue
                   If Yvalue < MinY Then MinY = Yvalue
                Loop
                
                Text1 = MinX
                Text2 = MaxX
                If Trim(txtNumFitPnts.Text) = sEmpty Then txtNumFitPnts.Text = numRows%
                NPI = numRows%
                If Trim(txtNumFitPnts.Text) = sEmpty Then
                   NPC = NPI
                   txtNumFitPnts.Text = NPC
                Else
                   NPC = Val(txtNumFitPnts.Text)
                   End If
                NumFitSteps = NPC
                
                Close #freefil%
               
                ' Find a good fit.
                'execute spline fit
                'results reside in array Pc(0 to NPC - 1)
            '
                ReDim Pc(0 To NPC - 1) ' As P_Type
                
                Select Case SplineType%
                   Case 1 'Bezier
                      If NPI > 1029 Then
                      
                         Call MsgBox("Bezier spline limited to 1029 points," _
                                     & vbCrLf & "but your file has" & Str(NPC) & " points" _
                                     & vbCrLf & "" _
                                     , vbInformation, "File too big")
                         
                         Screen.MousePointer = vbDefault
                         Exit Sub
                         End If
                      Call Bezier(Pi(), Pc())
                   Case 2 'B-spline
                      NK = CLng(Val(cmbDeg.Text))
                      Call B_Spline(Pi(), NK, Pc())
                   Case 3 'C-spline
                      Call C_Spline(Pi(), Pc())
                   Case 4 'T-spline
                      VZ = CLng(Val(cmbDeg.Text))
                      Call T_Spline(Pi(), VZ, Pc())
                End Select
                
                'now overplot the file and the fit
                'create temperorary two column fit file with the requested number of points

                filetmp% = FreeFile
                FitFileName = App.Path & "\tmp-fit.txt"
                Open FitFileName For Output As #filetmp%
                For J = 1 To UBound(Pc) + 1
                   Print #filetmp%, Pc(J - 1).X, Pc(J - 1).Y
                Next J
                Close #filetmp%
                
                'now overplot the file and its fit file
                'first add this file to the flxlstFiles
                'then add plot information in the plot arrays, then give command to make the overplot
                
                'Show the members of the returned sFile string
                'It path is larger than 3 characters, only show
                'most inner path (but record the rest in the file buffer)
      
                sTemp = FileRoot(FitFileName)
                sPath = App.Path
      
                'don't change last recorded plot directory, since this is just a tmp file
      
                'determine short form of this path consisting of
                'the innermost directory
                
                If notAlreadyFitted Then
                  Call ShortPath(sPath, MaxDirLen, sShortPath)
        
                  frmSetCond.flxlstFiles.AddItem sShortPath & sTemp
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
                  numfiles% = numfiles% + 1
                  ReDim Preserve Files(numfiles%)
                  Files(numfiles% - 1) = FitFileName
                  'highlight it
                  frmSetCond.flxlstFiles.list.item(numfiles%).Selected = True
                  notAlreadyFitted = False 'flag so won't repeat adding fit file to plot buffer
                  
                  'add plot information, it is always two column numbers separated by spaces and no headers (format 2)
                    
                  PlotInfo(0, numfiles% - 1) = 1
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  PlotInfo(3, numfiles% - 1) = "1" 'txtXA
                  If Val(PlotInfo(3, numfiles% - 1)) = 0 Then
                     PlotInfo(3, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(4, numfiles% - 1) = "0" 'txtXB
                  PlotInfo(5, numfiles% - 1) = "1" 'txtYA
                  If Val(PlotInfo(5, numfiles% - 1)) = 0 Then
                     PlotInfo(5, numfiles% - 1) = "1.0"
                     End If
                  PlotInfo(6, numfiles% - 1) = "0"
                  PlotInfo(7, numfiles% - 1) = FitFileName 'PlotInfofrm.lblFileName
                  PlotInfo(8, numfiles% - 1) = sEmpty
                  PlotInfo(9, numfiles% - 1) = "1"
                    
                  numFilesToPlot% = 2
                  numSelected% = 2
                      
                  frmSpline.Refresh 'return focus to the spline form
                  frmSetCond.flxlstFiles.Refresh
                  DoEvents
                  
               Else
                  'detect changes to plot type and plot color
                  PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
                  PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
                  
'                  numFilesToPlot% = 2
'                  numSelected% = 2
                 
                  End If
                
                Load frmDraw
                frmDraw.Show vbModeless
                frmDraw.Width = Screen.Width - frmSetCond.Width
                frmDraw.ScaleHeight = frmSetCond.ScaleHeight
                frmDraw.Left = frmSetCond.Width
                If frmSetCond.txtXTitle.Text = "" Then frmSetCond.txtXTitle.Text = "X-values"
                If frmSetCond.txtYTitle.Text = "" Then frmSetCond.txtYTitle.Text = "Y-values"
'                If frmSetCond.txtTitle.Text = "" Then frmSetCond.txtTitle.Text = "Chart Title (blank)"
                frmSetCond.cmdOK = True
                PlotAll = False 'reset plt all flag
                Screen.MousePointer = vbDefault
                'now wait a bit to show the fit before going on
                
                waittime = Timer
                Do Until Timer > waittime + 1
                   DoEvents
                Loop
                
            
            Else
            
               Call MsgBox("You need to choose a fit option!", vbInformation, "Fit option")
               Exit Sub
            
            End If
            
            'restore colors
            flxlstFiles.list.item(I + 1).ItemBackColor = tmpBackColor&
            flxlstFiles.list.item(I + 1).ItemForeColor = tmpForeColor&
            flxlstFiles.Refresh
            DoEvents
           
        End If
9999:
   Next I
   
   Fitting = False
   FitWizard = False
   
   If flxlstFiles.MultiSelect Then Exit Sub
  
   If found% = 0 Then
      Call MsgBox("Please select a fit file in the list", vbInformation, "No fit file selected")
      End If
      
Exit Sub

   On Error GoTo 0
   Exit Sub

cmdFit_Click_Error:

    If Err.Number = 9 And testfile% = 1 Then Resume Next
    
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFit_Click of Form frmSpline"
    Fitting = False
    Close
    
End Sub

Private Sub cmdRecord_Click()

   On Error GoTo cmdRecord_Click_Error
   
   If CurrentFitFileIndex = 0 Then
      Call MsgBox("No fit has been done yet.", vbExclamation, "Fit error")
      Exit Sub
                                  
      End If

   fileoutcoef% = FreeFile
   Open App.Path & "\Fit-coefficients.txt" For Append As #fileoutcoef%
   Print #fileoutcoef%, String(40, "=") 'demarcation
   Print #fileoutcoef%, FileRoot(PlotInfo(7, CurrentFitFileIndex - 1))
   Print #fileoutcoef%, "Polynomical coeficients of Plot program's LS fit to" & Str(cmbDeg.ListIndex + 1) & "th degree polynomial."
   
    For J = 0 To BestCoeffs.Count - 1
       doclin$ = Str(J) & ",    " & BestCoeffs.item(J + 1)
       Print #fileoutcoef%, doclin$
    Next J
    
    Close #fileoutcoef%

   On Error GoTo 0
   Exit Sub

cmdRecord_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRecord_Click of Form frmSpline"
    
End Sub

Private Sub cmdWizard_Click()

  Dim ier As Integer

  If flxlstFiles.list.Count = 0 Then
     MsgBox "Plot buffer is empty!", vbExclamation + vbOKOnly, "Plot"
     Exit Sub
     End If
     
  flxlstFiles.MultiSelect = True
  
  maxFilesToPlot% = MaxNumOverplotFiles
  numFilesToPlot% = 0
  
'  cmdWizard.Enabled = False
'  cmdClear.Enabled = False
'  mnuFormat.Enabled = False
'  mnuOpen.Enabled = False
  
  Dim I As Long, J As Long, idoc As Integer, degree As Integer
  
'  For i% = 0 To flxlstFiles.List.Count - 1
'      If flxlstFiles.List.item(i% + 1).Selected Then
  
  If flxlstFiles.list.Count > 15 Then
    'scroll the flex list box to the beginning
    For I = 1 To flxlstFiles.list.Count
       flxlstFiles.SetFocus
       waitime = Timer
       Do Until Timer > waitime + 0.001
          DoEvents
       Loop
       Call keybd_event(VK_UP, 0, 0, 0)
       Call keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)
    Next I
    End If
           
  numSelected% = -1
  For I = 0 To flxlstFiles.list.Count - 1
        'scroll down one
        flxlstFiles.SetFocus
        waitime = Timer
        Do Until Timer > waitime + 0.001
           DoEvents
        Loop
        flxlstFiles.SetFocus
        Call keybd_event(VK_DOWN, 0, 0, 0)
        Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
        
        If flxlstFiles.list.item(I + 1).Selected Then
            'scroll down one item
            flxlstFiles.SetFocus
            waitime = Timer
            Do Until Timer > waitime + 0.001
               DoEvents
            Loop
            flxlstFiles.SetFocus
            Call keybd_event(VK_DOWN, 0, 0, 0)
            Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
            
           numSelected% = I
           numFilesToPlot% = numFilesToPlot% + 1
           If I > maxFilesToPlot% Then
              MsgBox "You are allowed to overplot a maximum of " & _
              Str(maxFilesToPlot%) & "files!", vbExclamation + vbOKOnly, "Plot"
              Exit For
              End If
           tmpBackColor& = flxlstFiles.list.item(I + 1).ItemBackColor
           tmpForeColor& = flxlstFiles.list.item(I + 1).ItemForeColor
           Call keybd_event(VK_DOWN, 0, 0, 0)
           Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
           flxlstFiles.list.item(I + 1).ItemBackColor = &HFF&
           flxlstFiles.list.item(I + 1).ItemForeColor = QBColor(0)
           flxlstFiles.Refresh
           
           If optPoly.Value = True Then
           
              If PlotInfo(3, I) = 0 Then
                   Select Case MsgBox("Can't fit since you haven't yet defined this file's plot format!" _
                                      & vbCrLf & "(Apparently you haven't plotted it yet before entering splines)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "Do you wish to add format information to it at this time?" _
                                      , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                   
                    Case vbOK
                        Unload Me
                    Case vbCancel
                        Exit Sub
                   End Select
                   Exit Sub
                 End If
           
                HasSolution = False
                Set PtX = New Collection
                Set PtY = New Collection
                     
                PlotFileName$ = PlotInfo(7, I)
                freefil% = FreeFile
                Open PlotFileName$ For Input As #freefil%
                'skip the header lines
                For idoc = 1 To FilForm(0, PlotInfo(0, I))
                   Line Input #freefil%, doclin$
                Next idoc
        
                numRows% = 0
                Do Until EOF(freefil%)
                   Line Input #freefil%, doclin$
                   numRows% = numRows% + 1
                   PtX.Add dPlot(I, 0, numRows% - 1)
                   PtY.Add dPlot(I, 1, numRows% - 1)
                Loop
                                
                Close #freefil%
                
                
                 ' Find a good fit.
                degree = cmbDeg.ListIndex + 1
                If degree = 0 Then degree = 1
                PolyDeg = degree
                Set BestCoeffs = FindPolynomialLeastSquaresFit(PtX, PtY, degree, ier)
                If ier = -1 Then Exit Sub
            
                With flxGridFit
                    .Rows = PolyDeg + 2
                    .Cols = 2
                    .ColAlignment(0) = 1
                    .ColAlignment(1) = 0
                    .Visible = True
                    
                    .TextMatrix(0, 0) = "Degree"
                    .TextMatrix(0, 1) = "Coeficients"
                    
                    .TextMatrix(0, 0) = "Degree"
                     For J = 0 To PolyDeg
                       .TextMatrix(J + 1, 0) = Str(J) & "th degree"
                       .TextMatrix(J + 1, 1) = sEmpty
                    Next J
                    
                    For J = 0 To BestCoeffs.Count - 1
                       .TextMatrix(J + 1, 1) = Format(BestCoeffs.item(J + 1), "0.00000E+00")
                    Next J
                    
                End With
                
                Call AutosizeGridColumns(flxGridFit, degree + 2, 100)


            ElseIf optSpline.Value = True Then
            
            Else
                Call MsgBox("You need to choose a fit option!", vbInformation, "Fit option")
                
                Exit Sub

            End If

        End If
  Next I
  flxlstFiles.Refresh
'  cmdWizard.Enabled = True
'  cmdClear.Enabled = True
'  mnuFormat.Enabled = True
'  mnuOpen.Enabled = True
'
'  If numSelected% = -1 Then
'     MsgBox "No file was selected for plotting!", vbExclamation + vbOKOnly, "Plot"
'     Exit Sub
'     End If
'
'Load frmDraw
'frmDraw.Show vbModeless
'frmDraw.Width = Screen.Width - frmSetCond.Width
'frmDraw.ScaleHeight = frmSetCond.ScaleHeight
'frmDraw.Left = frmSetCond.Width
'If txtXTitle.Text = "" Then txtXTitle.Text = "X-values"
'If txtYTitle.Text = "" Then txtYTitle.Text = "Y-values"
'cmdOK_Click
'PlotAll = False 'reset plt all flag

End Sub

Private Sub cmdSavetoDisk_Click()

   Dim FileSave$, NumFitSteps As Integer
   Dim FitXStep As Double, XFit As Double, YFit As Double

On Error GoTo errhand
100:
    With comdlgFit
    
       .CancelError = True
       If optPoly.Value = True Then
          .FileName = App.Path & "\FitFile-" & FileRoot(FitFileName) '& "-deg:" & Trim$(Str(cmbDeg.ListIndex + 1)) & ".fit"
       Else
          .FileName = App.Path & "\FitFile-" & FileRoot(FitFileName) '& "-spline:" & Trim$(Str(cmbDeg.ListIndex + 1)) & ".fit"
          End If
          
       .Filter = "fit files (*.fit)|*.fit|text files (.txt)|*.txt|all files (*.*)|*.*"
       .ShowSave
       FileSave$ = .FileName
    
    End With
    
    If Dir(FileSave$) <> sEmpty Then
    
        Select Case MsgBox("File already exists!" _
                           & vbCrLf & "" _
                           & vbCrLf & "Overwrite?" _
                           , vbYesNoCancel Or vbExclamation Or vbDefaultButton1, "File Overwrite Protection")
        
            Case vbYes
                FileCopy FitFileName, FileSave$
                cmdAdd.Enabled = True
                Exit Sub
            Case vbNo
                GoTo 100
            Case vbCancel
                Exit Sub
        End Select
        End If
    
    FileCopy FitFileName, FileSave$
    
    cmdAdd.Enabled = True
      
    Exit Sub
      
errhand:
    Close
    MsgBox "Error detected, error number: " & Str(Err.Number) & vbCrLf & "Error description: " & Err.Description, vbCritical + vbOKOnly, "File error"

End Sub

Private Sub flxlstFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If flxlstFiles.MultiSelect Then Exit Sub 'wizard mode, clicks are just adding more files

    If InStr(frmSetCond.flxlstFiles.list.item(numfiles%).Text, "tmp-fit.txt") Then
        
        If FitFileName <> "" And Dir(FitFileName) <> sEmpty Then
           Kill FitFileName
           End If
           
        numfiles% = numfiles% - 1
        ReDim Preserve Files(numfiles%)
           
        ReDim Preserve PlotInfo(9, numfiles%)
        
        notAlreadyFitted = True

        End If
        
    frmSetCond.flxlstFiles.Clear
    
    'restore list
    For I = 0 To UBound(flxFileBuffer)
        frmSetCond.flxlstFiles.AddItem flxFileBuffer(I)
    Next I
    
    frmSetCond.flxlstFiles.Refresh
    DoEvents
            
    'restore plot buffer number
    numFilesToPlot% = OriginalNumPlotFiles

   
End Sub

Private Sub Form_Load()

  If NumAddSave > 0 Then
     NumAddSave = 0
     ReDim FileAddSave(NumAddSave)
     End If

  With frmSpline
     .Left = frmSetCond.Left
     .Top = frmSetCond.Top
  End With
  
  MaxPolyDeg = 12
  
  notAlreadyFitted = True

  For I = 1 To MaxPolyDeg
    cmbDeg.AddItem I
  Next I
  
  cmbDeg.Visible = True
  
  If PolyDeg% = 0 Then PolyDeg% = 1
  cmbDeg.Text = Trim$(Str$(PolyDeg%)) 'cmbDeg.ListIndex = StoredPolyDeg%
  
  'copy all the files listed in the flex filelist of the parent form to the spline's flex filelist
  If numfiles% > 0 Then
     numSelected% = 0
     For I = 1 To numfiles%
        'add to this list
         frmSpline.flxlstFiles.AddItem frmSetCond.flxlstFiles.list.item(I).Text
         ReDim Preserve flxFileBuffer(I - 1)
         flxFileBuffer(I - 1) = frmSetCond.flxlstFiles.list.item(I).Text
         numSelected% = numSelected% + 1
         ReDim Preserve SelectedFileNum(numSelected% - 1)
         SelectedFileNum(numSelected% - 1) = I
     Next I
     
     OriginalNumPlotFiles = numFilesToPlot% 'record how many files were in the plot buffer before fitting
     
    'now unselect all the files in the frmsetcond flexlst
    Text1.Text = frmSetCond.txtValueX0
    Text2.Text = frmSetCond.txtValueX1
    
     frmSpline.flxGridFit.Refresh
     For I = 1 To numfiles%
        frmSetCond.flxlstFiles.list.item(I).Selected = False
     Next I
     frmSetCond.flxlstFiles.Refresh
     frmSetCond.Refresh
     DoEvents
     End If
     
 With cmbPlotType
    .AddItem "Pick a plot type"
    .AddItem "Line" 'DRAWN_AS.AS_CONLINE
    .AddItem "Point" 'DRAWN_AS.AS_POINT
    .AddItem "Bar" 'DRAWN_AS.AS_BAR '"Point"
    .AddItem "Dash" 'DRAWN_AS.AS_DASH
    .AddItem "Dot" 'DRAWN_AS.AS_DOT
    .AddItem "Dashdot" 'DRAWN_AS.AS_DASHDOT
    .AddItem "Dashdotdot" 'DRAWN_AS.AS_DASHDOTDOT
    .AddItem "Circle" 'DRAWN_AS.AS_CIRCLE
    .AddItem "Filled Circle" 'DRAWN_AS.AS_FILLEDCIRCLE
    .ListIndex = 1
 End With
 
 If SplineType% <> 0 Then
    cmbPlotType.ListIndex = SplineType% - 1
    End If
 
 If FitPlotType% > 0 Then cmbPlotType.ListIndex = FitPlotType%
 
 With cmbPlotColor
    .AddItem "Pick a color"
    .AddItem "Automatic"
    .AddItem "black"
    .AddItem "blue"
    .AddItem "green"
    .AddItem "cyan"
    .AddItem "red"
    .AddItem "magenta"
    .AddItem "yellow"
    .AddItem "gray"
    .AddItem "light blue"
    .ListIndex = 1
 End With
 
 If FitPlotColor% > 0 Then cmbPlotColor.ListIndex = FitPlotColor%
 
 With cmbSpline
    .AddItem "Spline type"
    .AddItem "Bezier"
    .AddItem "B-Spline"
    .AddItem "C-Spline"
    .AddItem "T-Spline"
    .ListIndex = 1
 End With
 
 If SplineType% > 0 Then cmbSpline.ListIndex = SplineType%
 If SplineType% = 0 Then SplineType% = 1
 
 If NumFitPoints% > 0 Then
    txtNumFitPnts.Text = NumFitPoints%
 Else
    txtNumFitPnts.Text = "200"
    End If

    With flxGridFit
        .AllowUserResizing = flexResizeBoth
   End With
   
   Select Case FitMethod%
   
      Case 1 'polynomial least square fitting
          optPoly.Value = True
          cmbDeg.Visible = True
          cmbSpline.Visible = False
          cmdRecord.Visible = True
          chkAuto.Visible = True
          frmFitResults.Visible = True
          If PolyDeg = 0 Then PolyDeg = 1
          cmbDeg.Text = PolyDeg
          
      Case 2 'spline fitting
          optSpline.Value = True
          cmbSpline.Visible = True
          frmFitResults.Visible = False
          cmdRecord.Visible = False
          chkAuto.Visible = False
          If SplineType% = 0 Then SplineType% = 2
          Select Case SplineType%
             Case 1 'Bezier
                cmbDeg.Visible = False
             Case 2 'B-spline
                cmbDeg.Visible = True
                If SplineDeg% = 0 Then SplineDeg% = 3
                cmbDeg.Text = SplineDeg%
             Case 3 'C-spline
                cmbDeg.Visible = False
             Case 4 'T-spline
                cmbDeg.Visible = True
                If SplineDeg% = 0 Then SplineDeg% = 30
                cmbDeg.Text = SplineDeg%
          End Select
          
      Case Else
         FitMethod% = 1
         optPoly.Value = True
         cmbSpline.Visible = False
         frmFitParam.Visible = True
         frmFitResults.Visible = False
         cmdRecord.Visible = True
         chkAuto.Visible = True
         If PolyDeg = 0 Then PolyDeg = 1
         cmbDeg.Text = PolyDeg
             
   End Select
         
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'erase the temporary fit file
    'highlight it
    
    Dim FileSave$, sPath As String, sTemp As String, MaxDirLen As Integer, sShortPath As String
    
   On Error GoTo Form_Unload_Error

    MaxDirLen = Int(flxlstFiles.Width / 70) - 30
       
    If InStr(frmSetCond.flxlstFiles.list.item(numfiles%).Text, "tmp-fit.txt") Then
        
        If FitFileName <> "" And Dir(FitFileName) <> sEmpty Then
           Kill FitFileName
           End If
           
        numfiles% = numfiles% - 1
        ReDim Preserve Files(numfiles%)
           
        ReDim Preserve PlotInfo(9, numfiles%)
        
        notAlreadyFitted = True

        End If
        
    frmSetCond.flxlstFiles.Clear
    
    'restore list
    For I = 0 To UBound(flxFileBuffer)
        frmSetCond.flxlstFiles.AddItem flxFileBuffer(I)
    Next I
            
    'restore plot buffer number
    numFilesToPlot% = OriginalNumPlotFiles
        
    'restore original highlighted, i.e., selected files
    numSelected% = 0
    For J = 0 To UBound(SelectedFileNum)
      frmSetCond.flxlstFiles.list.item(J + 1).Selected = True
      numSelected% = numSelected% + 1
    Next J
    
    ReDim SelectedFileNum(0)
    numFilesToPlot% = OriginalNumPlotFiles
    frmSetCond.flxlstFiles.Refresh
    DoEvents
    
    'now add any saved add files
    If NumAddSave > 0 Then
       For I = 1 To NumAddSave
         numfiles% = numfiles% + 1
         ReDim Preserve Files(numfiles% - 1)
         
         FileSave$ = FileAddSave(I - 1)
         
         Files(numfiles% - 1) = FileSave$
         
         sTemp = FileRoot(FileSave$)
        
         sPath = Mid$(FileSave$, 1, Len(FileSave$) - Len(sTemp))
        
         MaxDirLen = Int(flxlstFiles.Width / 70) - 30
        
         Call ShortPath(sPath, MaxDirLen, sShortPath)
         
         ReDim Preserve flxFileBuffer(UBound(flxFileBuffer) + 1)
         flxFileBuffer(UBound(flxFileBuffer)) = sShortPath & sTemp
         frmSetCond.flxlstFiles.AddItem sShortPath & sTemp
         frmSetCond.Refresh
         DoEvents
        
        'add plot information
        
        ReDim Preserve PlotInfo(9, numfiles%)
        
        'add plot information, i.e., two columns X,Y no headers = format #2
        PlotInfo(0, numfiles% - 1) = 1
        PlotInfo(1, numfiles% - 1) = cmbPlotType.ListIndex - 1 'Str$(chkPlotType%)
        PlotInfo(2, numfiles% - 1) = cmbPlotColor.ListIndex - 1 'Str$(chkColor%)
        PlotInfo(3, numfiles% - 1) = "1" 'txtXA
        If Val(PlotInfo(3, numfiles% - 1)) = 0 Then
            PlotInfo(3, numfiles% - 1) = "1.0"
            End If
        PlotInfo(4, numfiles% - 1) = "0" 'txtXB
        PlotInfo(5, numfiles% - 1) = "1" 'txtYA
        If Val(PlotInfo(5, numfiles% - 1)) = 0 Then
            PlotInfo(5, numfiles% - 1) = "1.0"
            End If
        PlotInfo(6, numfiles% - 1) = "0"
        PlotInfo(7, numfiles% - 1) = FileSave$
        PlotInfo(8, numfiles% - 1) = sEmpty
        PlotInfo(9, numfiles% - 1) = "1"
        
        frmSetCond.flxlstFiles.Refresh
        DoEvents
    Next I
    End If
       
     'record plot info
    'record this information
    filplt% = FreeFile
    Open App.Path & "\PlotDirec.txt" For Output As #filplt%
    Write #filplt%, "This file is used by Plot. Don't erase it!"
    Write #filplt%, direct$, directPlot$, dirWordpad
    Write #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
    Close #filplt%
    
   Unload Me
   Set frmSpline = Nothing

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    
End Sub

Private Sub optPoly_Click()
   If optPoly.Value = True Then
      frmFitParam.Visible = True
      cmbDeg.Visible = True
      frmFitResults.Visible = True
      cmdRecord.Visible = True
      chkAuto.Visible = True
      cmbSpline.Visible = False
      FitMethod% = 1
      chkCurvature.Visible = True
      End If
End Sub

Private Sub optSpline_Click()
   If optSpline.Value = True Then
      frmFitResults.Visible = False
      cmdRecord.Visible = False
      chkAuto.Visible = False
      cmbSpline.Visible = True
      FitMethod% = 2
      cmbSpline.ListIndex = 1
      cmbDeg.Visible = False
      chkCurvature.Visible = False
      End If
End Sub
