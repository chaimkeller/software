VERSION 5.00
Begin VB.Form frmStat 
   Caption         =   "Statistics"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   6810
   Icon            =   "frmStat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmAdj 
      Caption         =   "Units for variation"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   6960
      Width           =   6615
      Begin VB.TextBox txtUnits 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Text            =   "Seconds"
         ToolTipText     =   "String value: units of variance"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtMult 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Text            =   "60.0"
         ToolTipText     =   "Multiply variance by"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblUnits 
         Caption         =   "Units:"
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
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblMult 
         Caption         =   "Multiply by:"
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
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frmHelp 
      Caption         =   "Instructions"
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   6495
      Begin VB.Label Label1 
         Caption         =   $"frmStat.frx":0442
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdStat 
      Caption         =   "Calculate the error from the fit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   6495
   End
   Begin VB.Frame frmStats 
      Caption         =   "Statistics"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   6615
      Begin VB.Label lblVariance 
         Alignment       =   2  'Center
         Caption         =   "Square root of the Variance: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Squared variance"
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         Caption         =   "Mean Difference: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Mean difference"
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame frmFit 
      Caption         =   "Fit"
      Height          =   2680
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   6615
      Begin VB.ListBox lstFit 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Data to fit"
      Height          =   2640
      Left            =   120
      TabIndex        =   0
      Top             =   950
      Width           =   6615
      Begin VB.ListBox lstData 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   180
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure : cmdStat_Click
' Author    : chaim
' Date      : 1/1/2021
' Purpose   : Calculates the normalized absolute variance between chosen data and fit files
'---------------------------------------------------------------------------------------
'
Private Sub cmdStat_Click()
  Dim SelectedData As Integer
  Dim SelectedFit As Integer
  Dim DataSelected As Boolean
  Dim FitSelected As Boolean
  Dim DataValues As COORDINATE
  Dim FitValues As COORDINATE
  Dim Xo As Double, Yo As Double
  Dim X1 As Double, Y1 As Double
  Dim YInterpolate As Double
  Dim freefitfile%, freedatafile%, numRowsData%, numRowsFit%
  Dim SumOfAbsVariance As Double
  Dim NumSumVariance As Long
  Dim MeanDiff As Double
  Dim StatVariance As Double
  
  'interpolate each point of the data to the "fit" value and take the absolute difference and add to sum
  
  'find selected data index
   On Error GoTo cmdStat_Click_Error
   
   StatVariance = 0#
   MeanDiff = 0#

  For I = 1 To lstData.ListCount
    If lstData.Selected(I - 1) = True Then
       SelectedData = I - 1
       DataSelected = True
       Exit For
       End If
  Next I
  
  'find selected fit index
  For I = 1 To lstFit.ListCount
    If lstFit.Selected(I - 1) = True Then
       SelectedFit = I - 1
       FitSelected = True
       Exit For
       End If
  Next I
  
  If DataSelected = False Or FitSelected = False Then
     Call MsgBox("Please select one data and one fit file", vbInformation, "Missing selection")
     Exit Sub
     End If
     
  If SelectedData = SelectedFit Then
     Call MsgBox("The data and fit files must be different", vbInformation, "Bad selection")
     Exit Sub
     End If
     
  If PlotInfo(3, SelectedData) = "" Then
    
      Select Case MsgBox("Can't calculate since you haven't yet defined the data file's plot format!" _
                         & vbCrLf & "(Apparently you haven't plotted it yet before entering statistics)" _
                         & vbCrLf & "" _
                         & vbCrLf & "Do you wish to add format information to it at this time?" _
                         , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                         
        Case vbOK
          Unload Me
        Case vbCancel
          Exit Sub
      End Select
      
      End If
      
  If PlotInfo(3, SelectedFit) = "" Then
    
      Select Case MsgBox("Can't calculate since you haven't yet defined the fit file's plot format!" _
                         & vbCrLf & "(Apparently you haven't plotted it yet before entering statistics)" _
                         & vbCrLf & "" _
                         & vbCrLf & "Do you wish to add format information to it at this time?" _
                         , vbOKCancel Or vbInformation Or vbDefaultButton1, "Missing format information")
                         
        Case vbOK
          Unload Me
        Case vbCancel
          Exit Sub
      End Select
      
      End If
      
    Screen.MousePointer = vbHourglass
      
   'open both files and begin interpolation
    freedatafil% = FreeFile
    Open PlotInfo(7, SelectedData) For Input As #freedatafil%
    'skip the header lines
    For idoc = 1 To FilForm(0, PlotInfo(0, SelectedData))
       Line Input #freedatafil%, doclin$
    Next idoc
    
    freefitfil% = FreeFile
    Open PlotInfo(7, SelectedFit) For Input As #freefitfil%
    'skip the header lines
    For idoc = 1 To FilForm(0, PlotInfo(0, SelectedFit))
       Line Input #freefitfil%, doclin$
    Next idoc

    numRowsData% = 0
    Do Until EOF(freedatafil%)
       Line Input #freedatafil%, doclin$
       numRowsData% = numRowsData% + 1
       DataValues.X = dPlot(SelectedData, 0, numRowsData% - 1)
       DataValues.Y = dPlot(SelectedData, 1, numRowsData% - 1)
       Seek (freefitfil%), 1 'set at beginning of fit file
       numRowsFit% = 0
       Do Until EOF(freefitfil%)
          Line Input #freefitfil%, doclin$
          numRowsFit% = numRowsFit% + 1
          FitValues.X = dPlot(SelectedFit, 0, numRowsFit% - 1)
          FitValues.Y = dPlot(SelectedFit, 1, numRowsFit% - 1)
          X0 = FitValues.X
          Y0 = FitValues.Y
          If EOF(freefitfil%) Then Exit Do
          Line Input #freefitfil%, doclin$
          numRowsFit% = numRowsFit% + 1
          FitValues.X = dPlot(SelectedFit, 0, numRowsFit% - 1)
          FitValues.Y = dPlot(SelectedFit, 1, numRowsFit% - 1)
          X1 = FitValues.X
          Y1 = FitValues.Y
          If X0 < DataValues.X And X1 < DataValues.X Then
             'keep on looping through fit file
          ElseIf X0 <= DataValues.X And X1 > DataValues.X Then
             'found x match, so interpolate
             YInterpolate = (DataValues.X - X0) * (Y1 - Y0) / (X1 - X0) + Y0
             MeanDiff = MeanDiff + Abs(DataValues.Y - YInterpolate)
             StatVariance = StatVariance + (DataValues.Y - YInterpolate) * (DataValues.Y - YInterpolate)
             NumSumVariance = NumSumVariance + 1
             'now add to sum
          ElseIf X0 > DataValues.X And X1 > DataValues.X Then
             'couldn't find any match, so skip this point
             Exit Do
             End If
       Loop
    Loop
    
    Close #freedatafil%
    Close #freefitfil%
    freedatafil% = 0
    freefitfil% = 0
    
    If NumSumVariance > 0 Then
        MeanDiff = Val(txtMult.Text) * MeanDiff / NumSumVariance
        StatVariance = Val(txtMult.Text) * Sqr(StatVariance / NumSumVariance)
        lblStat.Caption = "Mean Absolute Difference: " & Format(Str$(MeanDiff), "#####0.0###") & " " & Trim$(txtUnits.Text)
        lblVariance.Caption = "Square Root Variance: " & Format(Str$(StatVariance), "#####0.0###") & " " & Trim$(txtUnits.Text)
    Else
        lblStat.Caption = "Normalized absolute difference undetermined!"
        lblVariance.Caption = "Square Root Variance undetermined!"
        End If
        
   Screen.MousePointer = vbDefault
     
  'begin interpolation

   On Error GoTo 0
   Exit Sub

cmdStat_Click_Error:
    Screen.MousePointer = vbDefault
    If freedatafil% > 0 Then Close #freedatafil%
    If freefitfil% > 0 Then Close #freefitfil%
    If Err.Number = 9 Then
       Call MsgBox("You must first plot the files before running a statistical analysis!", vbInformation, "Statistical analysis error")
    Else
       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdStat_Click of Form frmStat"
       End If
  
End Sub

Private Sub Form_Load()
  If numfiles% > 0 Then
     
     For I = 1 To numfiles%
        'add to this list
         lstData.AddItem frmSetCond.flxlstFiles.list.item(I).Text
         lstFit.AddItem frmSetCond.flxlstFiles.list.item(I).Text
     Next I

     End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload Me
  Set frmStat = Nothing
End Sub

Private Sub lstData_Click()
   'make sure that only one record is checked
   numchecked% = 0
   For I = 0 To lstData.ListCount - 1
       If lstData.Selected(I) = True Then
          numchecked% = numchecked% + 1
          End If
   Next I
   If numchecked% > 1 Then
      Call MsgBox("Check only one data file!" & vbCrLf & vbCrLf & "(The first checked file is used)", vbInformation, "Too many selected")
      End If
   
End Sub

Private Sub lstFit_Click()
   'make sure that only one record is checked
   numchecked% = 0
   For I = 0 To lstFit.ListCount - 1
       If lstFit.Selected(I) = True Then
          numchecked% = numchecked% + 1
          End If
   Next I
   If numchecked% > 1 Then
      Call MsgBox("Check only one fit file!" & vbCrLf & vbCrLf & "(The first checked file is used)", vbInformation, "Too many selected")
      End If
End Sub
