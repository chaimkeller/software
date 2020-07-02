VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GDReportfrm 
   Caption         =   "Search Report"
   ClientHeight    =   3390
   ClientLeft      =   375
   ClientTop       =   990
   ClientWidth     =   9195
   Icon            =   "GDReportfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   9195
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancelFilter 
      BackColor       =   &H008080FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   175
      Left            =   4300
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancel the filter operation"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtZmax 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4275
      TabIndex        =   10
      ToolTipText     =   "Maximum height to be selected  (meters)"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtZmin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   4275
      TabIndex        =   8
      ToolTipText     =   "Minimum height to be selected (meters)"
      Top             =   30
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdZfilter 
      Caption         =   "&Height filter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Activate a height filter for the records"
      Top             =   30
      Width           =   1365
   End
   Begin VB.CheckBox chkGL 
      Caption         =   "Replace zero ground levels with the DTM heights"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Check the box to replace zero ground levels with the DTM heights"
      Top             =   460
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "&Select All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Select all the search results and enter them to the plot buffer"
      Top             =   30
      Width           =   1215
   End
   Begin MSComCtl2.Animation anmReport 
      Height          =   375
      Left            =   8580
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   661
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   30
      FullHeight      =   25
   End
   Begin VB.PictureBox picAnimation 
      Height          =   495
      Left            =   8520
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":0442
            Key             =   "cono"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":059C
            Key             =   "diatom"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":06F6
            Key             =   "foram"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":0850
            Key             =   "mega"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":09AA
            Key             =   "multi"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":0B04
            Key             =   "nano"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":0C5E
            Key             =   "ostra"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":0DB8
            Key             =   "paly"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":0F12
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":106C
            Key             =   "point"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":14BE
            Key             =   "line"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":1910
            Key             =   "Contour"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":1D62
            Key             =   "Eraser"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDReportfrm.frx":1EBC
            Key             =   "point2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwReport 
      Height          =   2595
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Unselect All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1330
      TabIndex        =   2
      ToolTipText     =   "Unselect all the selected ponts and clear the plot buffer, and erase the plotted points from the map"
      Top             =   30
      Width           =   1215
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "&Plot selected records"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Plot the selected records on the map"
      Top             =   30
      Width           =   1995
   End
   Begin VB.Label lblZMax 
      Caption         =   "max:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3960
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   280
   End
   Begin VB.Label lblZMin 
      Caption         =   "min:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3960
      TabIndex        =   9
      Top             =   40
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "GDReportfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iITMx&, iITMy&, iMD&, iGL&, iZ&, PreviouslySelected As Boolean

'---------------------------------------------------------------------------------------
' Procedure : ReplaceGLs
' DateTime  : 5/26/2004 22:54
' Author    : Chaim Keller
' Purpose   : Replaces and Unreplaces Ground Levels with the DTM heights
'---------------------------------------------------------------------------------------
'
Private Sub ReplaceGLs(GLmode%)

   On Error GoTo ReplaceGLs_Error
   
   Screen.MousePointer = vbHourglass
   
   If GLmode% = 0 Then 'replace zero GLs with DTM heights
   
      'first store search results in a temporary file, and then replace GL's
'      Close
      filtm1& = FreeFile
      If filtm1& = filnumg% Then filtm1& = filtm1& + 1 '(avoids conflicts with the "Close" statement in sub DTMheights)
      Open direct$ & "\temp.sav" For Output As #filtm1&
      'write identifying information
      Print #filtm1&, "MapDigitizer Search Results, Date/Time: " & Now()
      Print #filtm1&, sEmpty
      'write number of columns and rows:
      Print #filtm1&, "[Number of Columns and Rows]"
      Write #filtm1&, GDReportfrm.lvwReport.ColumnHeaders.count, numReport&
      Print #filtm1&, sEmpty
      'write column headers
      Print #filtm1&, "[Column Headers]"
      Dim sColum$
      sColum$ = sEmpty
      For i& = 1 To GDReportfrm.lvwReport.ColumnHeaders.count
         If i& = 1 Then
            sColum$ = Chr(34) & GDReportfrm.lvwReport.ColumnHeaders(i&) & Chr( _
                34)
         Else
            sColum$ = sColum$ & "," & Chr( _
                34) & GDReportfrm.lvwReport.ColumnHeaders(i&) & Chr(34)
            If InStr(Mid$(LCase(GDReportfrm.lvwReport.ColumnHeaders(i&)), 1, 3), "itm") <> 0 Then
               iITMx& = i& 'ITMx coordinate
               End If
            If InStr(Mid$(LCase(GDReportfrm.lvwReport.ColumnHeaders(i&)), 1, 3), "itm") <> 0 Then
               iITMy& = i& 'ITMx coordinate
               End If
            If InStr(GDReportfrm.lvwReport.ColumnHeaders(i&), "Mean Depth") <> 0 Then
               iMD& = i& 'colum contains mean depth
               End If
            If InStr(GDReportfrm.lvwReport.ColumnHeaders(i&), "Gnd Level") <> 0 Then
               iGL& = i& 'this is colum that contains GL info
               End If
            If InStr(GDReportfrm.lvwReport.ColumnHeaders(i&), "Z (meter)") <> 0 Then
               iZ& = i& 'this is colum that contains depth w.r.t sea level
               End If
            End If
      Next i&
      Print #filtm1&, sColum$
      Print #filtm1&, sEmpty
      Print #filtm1&, "[Search Results]"
      
      For j& = 1 To numReport&
        
        sColum$ = sEmpty
        madechange% = 0 'flag that flags changes
        For i& = 1 To GDReportfrm.lvwReport.ColumnHeaders.count
           If i& = 1 Then
              sColum$ = Chr(34) & GDReportfrm.lvwReport.ListItems( _
                  j&).Text & Chr(34)
           Else
              sColum$ = sColum$ & "," & Chr( _
                  34) & GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1) & Chr( _
                  34)
              If i& = iITMx& Then
                 ITMx = val(GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1))
                 'change height to DTM height
                 End If
              If i& = iITMy& Then
                 ITMy = val(GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1))
                 'change height to DTM height
                 End If
              If i& = iMD& Then
                 MD = val(GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1))
                 'change height to DTM height
                 End If
              If i& = iGL& And val(ITMx) <> 0 And val(ITMy) <> 0 And _
                 val(GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1)) = 0 Then
                 'change height to DTM height
                 kmx = ITMx
                 kmy = ITMy
                 'Call DTMheight(kmx, kmy, hgt)
                 Dim hgt As Integer
                 Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
                 If hgt <> -9999 Then
                    GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1) = Trim$(str$(hgt))
                    madechange% = 1
                    End If
                 End If
              If i& = iZ& And madechange% = 1 Then
                 'change height w.r.t. sea level
                 GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1) = Trim$(str$(hgt - MD))
                 End If
                 
              End If
        Next i&
        Print #filtm1&, sColum$
100   Next j&
      Close #filtm1&
     
   
   ElseIf GLmode% = 1 Then 'restore original GLs and Z's
   
      filtm1& = FreeFile
      If filtm1& = filnumg% Then filtm1& = filtm1& + 1 '(avoids conflicts with the "Close" statement in sub DTMheights)
      If Dir(direct$ & "\temp.sav") = sEmpty Then
         Screen.MousePointer = vbDefault
         MsgBox "Ground levels can't be restored!" & vbLf & _
                "Run another search to restore them.", vbExclamation + vbOKOnly, App.Title
         Exit Sub
         End If
                 
      Open direct$ & "\temp.sav" For Input As #filtm1&
      Line Input #filtm1&, doclin$
      Line Input #filtm1&, doclin$
      Line Input #filtm1&, doclin$
      'read in number of columns and rows
      Input #filtm1&, NumCol&, numRow&
      Line Input #filtm1&, doclin$
      Line Input #filtm1&, doclin$
      
     'column headers
      For i& = 1 To NumCol&
         Input #filtm1&, doclin$
      Next i&
       
       Line Input #filtm1&, doclin$
       Line Input #filtm1&, doclin$
       
       'read in and load up search results
       For j& = 1 To numRow&
          For i& = 1 To NumCol&
              Input #filtm1&, doclin$
              If i& = iITMx& Then
                 ITMx = val(doclin$)
                 End If
              If i& = iITMy& Then
                 ITMy = val(doclin$)
                 End If
              If i& = iGL& And val(ITMx) <> 0 And val(ITMy) <> 0 And _
                 val(doclin$) = 0 Then
                 'restore the height to search result values
                 GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1) = Trim$(doclin$)
                 End If
              If i& = iZ& Then
                 'restore depth w.r.t. sea level
                 GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1) = Trim$(doclin$)
                 End If
          Next i&
       Next j&
    
       Close #filtm1&
       numReport& = numRow&
   
      End If
      
      
   Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Sub

ReplaceGLs_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReplaceGLs of Form GDReportfrm", vbCritical + vbOKOnly
End Sub

Private Sub RestoreHighlighted()
    'clicking unselects old selected points, so
    'make sure that they remain selected under "Clear All"
    'button is used
    
    On Error GoTo errhand
    
    NumReportPnts& = GDReportfrm.lvwReport.ListItems.count
    
    If NumReportPnts& > 1 And numHighlighted& > 1 Then
        'left clicking erased the highlighted points, so restore them
        'now that the mouse is released
        For i& = 1 To numReport& 'search over all the search results
            If Highlighted(i& - 1) = 1 Then
               GDReportfrm.lvwReport.ListItems(i&).Selected = True
               GDReportfrm.lvwReport.ListItems(NewHighlighted&).EnsureVisible
            Else
               GDReportfrm.lvwReport.ListItems(i&).Selected = False
               GDReportfrm.lvwReport.ListItems(NewHighlighted&).EnsureVisible
               End If
            DoEvents  'yield to windows messaging
        Next i&

        'add delay to rectify asynchronous processing of Window's messages
         If numReport& >= 100 Then
            addtime = 50 / numReport&
         Else
            addtime = 0.5
            End If
        waitime = Timer
        Do Until Timer > waitime + addtime
          DoEvents
        Loop
        
        GDReportfrm.lvwReport.ListItems(NewHighlighted&).EnsureVisible
        End If
         
        Exit Sub
        
errhand:
   Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGL_Click
' DateTime  : 5/26/2004 22:47
' Author    : Chaim Keller
' Purpose   : Replaces zero Ground Levels (GL) with DTM heights
'             This is important when lacking info. on the Gnd. level of a well
'---------------------------------------------------------------------------------------
'
Private Sub chkGL_Click()
   On Error GoTo chkGL_Click_Error

   If chkGL.value = vbChecked And _
      chkGL.Caption = "Replace zero ground levels with the DTM heights" Then
      
      Select Case MsgBox("Replace any zero ground levels (GL=0) with the DTM heights." _
                         & vbCrLf & "Replace them?" _
                         , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, App.Title)
      
        Case vbYes 'replace the zero ground levels
          Call ReplaceGLs(0)
          chkGL.Caption = "Restore ground levels"
          chkGL.ToolTipText = "Check the box to restore ground levels"
          chkGL.value = vbUnchecked
        Case vbNo
          chkGL.value = vbUnchecked
          'don't do anything more
        Case vbCancel
          chkGL.value = vbUnchecked
          'don't do anything more
      End Select
   
   ElseIf chkGL.value = vbChecked And _
      chkGL.Caption = "Restore ground levels" Then
      
      response = MsgBox("Undo the changes in the Ground Levels?", vbQuestion + vbYesNoCancel, App.Title)
      If response = vbYes Then
         Call ReplaceGLs(1)
         chkGL.Caption = "Replace zero ground levels with the DTM heights"
         chkGL.value = vbUnchecked
      ElseIf response = vbCancel Or response = vbNo Then
         chkGL.value = vbUnchecked
         End If
      End If

   On Error GoTo 0
   Exit Sub

chkGL_Click_Error:
   
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGL_Click of Form GDReportfrm"
End Sub

Private Sub cmdAll_Click()
    'highligh all the records in the list view of GDReportfrm
    
    Screen.MousePointer = vbHourglass
    
    For i& = 1 To numReport&
        GDReportfrm.lvwReport.ListItems(i&).Selected = True
    Next i&
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdCancelFilter_Click()
   'cancel the filter operation
   With GDReportfrm

     'restore the filter button to its original appearence
     .cmdZfilter.Caption = "&Height filter"
     .cmdZfilter.ToolTipText = "Activate a height filter for the records"
     .cmdZfilter.BackColor = &H8000000F
     .cmdCancelFilter.Visible = False
    
     'deactivate the height filter text boxes
     .lblZMin.Visible = False
     .lblZMax.Visible = False
     .txtZmin.Visible = False
     .txtZmax.Visible = False

   End With
End Sub

Private Sub cmdClear_Click()
  
    'erase old plotted search points if any
    
    Screen.MousePointer = vbHourglass

    EraseOldSearchPoints
  
    NumReportPnts& = 0 'clear plot point buffer
    ReDim ReportPnts(1, 0) 'clear plot point memory
    numHighlighted& = 0 'clear highlighted point buffer
    ReDim Highlighted(0) 'clear the highlighted point memory
   
    'clear highlights
    ClearHighlightedPoints
    
    Screen.MousePointer = vbDefault
    
    GDReportfrm.lvwReport.Refresh
    
    If GeoMap Or TopoMap Then
       'reset blinking
       GDMDIform.CenterPointTimer.Enabled = True
       ce& = 1
       End If
       
    Ret = ShowWindow(GDReportfrm.hwnd, SW_MAXIMIZE)
    
End Sub

Private Sub cmdLocate_Click()

    Call LocatePointOnMap(0)
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdZfilter_Click
' DateTime  : 2/25/2005 15:30
' Author    : Chaim Keller
' Purpose   : Enables record selection based on height
'---------------------------------------------------------------------------------------
'
Private Sub cmdZfilter_Click()

   On Error GoTo cmdZfilter_Click_Error

   With GDReportfrm
   
      Select Case .cmdZfilter.Caption
         Case "&Height filter"
            .cmdZfilter.Caption = "&Apply the filter"
            .cmdZfilter.ToolTipText = "Select records between the inputed Z minimum and maximum"
            .cmdZfilter.BackColor = &HC0E0FF
            .cmdCancelFilter.Visible = True
            .lblZMin.Visible = True
            .lblZMax.Visible = True
            .txtZmin.Visible = True
            .txtZmax.Visible = True
         Case "&Apply the filter"
            'apply the filter, but first check if the selected range makes sense
            
            If val(.txtZmin) > val(.txtZmax) Then
               Call MsgBox("Z maximum must be equal or larger to Z minimum!" _
                           & vbCrLf & "Check your inputs in the Z minimum and maximum text boxes." _
                           , vbExclamation, App.Title)
               Exit Sub
               End If
               
            'check if some records have already been selected
            Screen.MousePointer = vbHourglass
            PreviouslySelected = False
            For i& = 1 To numReport&
                If GDReportfrm.lvwReport.ListItems(i&).Selected = True Then
                   PreviouslySelected = True
                   Exit For
                   End If
            Next i&
            Screen.MousePointer = vbDefault
               
            'now apply filter
            If PreviouslySelected Then 'detected records that have been previously selected
                Select Case MsgBox("Records have been previously selected!" _
                                   & vbCrLf & "" _
                                   & vbCrLf & "Do you wish that those records remain selected after the filter operation?" _
                                   & vbCrLf & "" _
                                   & vbCrLf & "Click ""Yes"": To retain them even after the filter operation." _
                                   & vbCrLf & "Click  ""No"": To unselect (and unplot) them before applying the filter operation." _
                                   , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, App.Title)
                
                    Case vbYes
                         GoTo s500 'overplot filtered points over old points
                    Case vbNo
                         .cmdClear.value = 1 'erase then filter
                         GoTo s500
                    Case vbCancel 'cancel the last operation
                
                End Select
             Else 'filter
                GoTo s500
                End If
                
      End Select
      
   End With

   On Error GoTo 0
   Exit Sub
   
'-------------------select records based on Z filter-------------------
s500:
    'select points based on the z filter
    
    Screen.MousePointer = vbHourglass
    
    With GDReportfrm
    
        'find with column contains the Z data
        For i& = 1 To GDReportfrm.lvwReport.ColumnHeaders.count
            If InStr(GDReportfrm.lvwReport.ColumnHeaders(i&), "elevation") <> 0 Then
               iZ& = i& 'this is colum that contains depth w.r.t sea level
               Exit For
               End If
        Next i&
        
        'now select the records in the desired z range
        
        For i& = 1 To numReport&
            If val(GDReportfrm.lvwReport.ListItems(i&).ListSubItems(iZ& - 1)) >= val(txtZmin) _
               And val(GDReportfrm.lvwReport.ListItems(i&).ListSubItems(iZ& - 1)) <= val(txtZmax) Then
               
               'this record is within the range so select it
               GDReportfrm.lvwReport.ListItems(i&).Selected = True
               End If
        Next i&
        
     'restore the button to its original appearence
     .cmdZfilter.Caption = "&Height filter"
     .cmdZfilter.ToolTipText = "Activate a height filter for the records"
     .cmdZfilter.BackColor = &H8000000F
     .cmdCancelFilter.Visible = False
     
     'deactivate the height filter text boxes
     .lblZMin.Visible = False
     .lblZMax.Visible = False
     .txtZmin.Visible = False
     .txtZmax.Visible = False
        
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
'----------------------------------------------------------------


'error handler
cmdZfilter_Click_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdZfilter_Click of Form GDReportfrm", vbCritical + vbOKOnly
End Sub

Private Sub Form_Load()
   NumReportPnts& = 0
   With GDMDIform
     .Toolbar1.Buttons(28).ToolTipText = "Show summary of search results"
     'enable some buttons
     .Toolbar1.Buttons(29).Enabled = True 'print
     .mnuPrintReport.Enabled = True
     .Toolbar1.Buttons(30).Enabled = True 'save
     .mnuSave.Enabled = True
     If lblX = "lon." And LblY = "lat." And Trim$(googledir) <> sEmpty Then .Toolbar1.Buttons(33).Enabled = True 'Google
     .mnuGoogle.Enabled = True
     If heights Then chkGL.Enabled = True
     
     Call WheelHook(Me.hwnd)

'     If EditDBVis Then 'actativate ability to dump search results to editform of scanned db
'        GDEditScannedDBfrm.optReport.Enabled = True
'        End If
   End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If Searching Then
      Cancel = True
      Exit Sub
      End If
         
   If TopoMap Then
      GDMDIform.CenterPointTimer.Enabled = True
      ce& = 1
      End If
      
'   erase old plotted search points if any
'   If SearchDB Then EraseOldSearchPoints
   
   NewHighlighted& = 0 'clear current highlighted point
   NumReportPnts& = 0 'clear plot point buffer
   ReDim ReportPnts(1, 0) 'clear plot point memory
   numHighlighted& = 0 'clear highlighted point buffer
   ReDim Highlighted(0) 'clear the highlighted point memory
   GDMDIform.StatusBar1.Panels(2) = sEmpty 'clear current record status
   
   PicSum = False
   
   GDMDIform.Toolbar1.Buttons(28).value = tbrUnpressed
   buttonstate&(28) = 0
   GDMDIform.Toolbar1.Refresh
   
   If TopoMap Then
      GDMDIform.CenterPointTimer.Enabled = True
      ce& = 1
      End If

   If ShowDetails Then
     'A detailed report is visible,
     'so unload it also.
      Unload GDDetailReportfrm
      End If
      
   With GDMDIform
     'disenable some buttons
     GDMDIform.Toolbar1.Buttons(29).Enabled = False
     .mnuPrintReport.Enabled = False
     GDMDIform.Toolbar1.Buttons(30).Enabled = False
     .mnuSave.Enabled = False
     GDMDIform.Toolbar1.Buttons(33).Enabled = False
     .mnuGoogle.Enabled = False
        
     GDMDIform.Toolbar1.Buttons(28).ToolTipText = "Preview database record"
   End With
'
'   'disenable Report Form option in GDeditscanneddbfrm
'   If EditDBVis Then GDEditScannedDBfrm.optReport.Enabled = False
   Call WheelUnHook(Me.hwnd)
   
   Unload Me
   Set GDReportfrm = Nothing
   
End Sub

Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

   On Error Resume Next
   If GDReportfrm.WindowState = vbMinimized Then Exit Sub
   lvwReport.Width = GDReportfrm.Width - 120
   lvwReport.Height = GDReportfrm.Height - 1100 '500
   
   If GDReportfrm.WindowState = vbMaximized Then
      RestoreHighlighted
   Else
      RecordHighlighted
      End If

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form GDReportfrm"
End Sub

Private Sub lvwReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   On Error GoTo lvwReport_ColumnClick_Error

   'sort search results according to the column clicked
   GDReportfrm.lvwReport.Sorted = True
   GDReportfrm.lvwReport.SortKey = ColumnHeader.Index - 1
   GDReportfrm.lvwReport.SortOrder = lvwAscending

   On Error GoTo 0
   Exit Sub

lvwReport_ColumnClick_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lvwReport_ColumnClick of Form GDReportfrm"
End Sub

Private Sub lvwReport_DblClick()

   On Error GoTo lvwReport_DblClick_Error

   'present all the fields of this record in the PrintPreview Form
   '(i.e., have it ready for printing)
   If Not Previewing Then 'no previews minmized
      PrintPreview.Visible = True
      Ret = ShowWindow(PrintPreview.hwnd, SW_MAXIMIZE)
   Else 'refresh the print preview with this one
      FillPrintCombo
      If Not LoadInit Then
         PreviewPrint
         End If
      If EditDBVis Then
         PrintPreview.cmdEditScannedDB.Visible = False
      Else
         PrintPreview.cmdEditScannedDB.Visible = True
         End If
      Ret = ShowWindow(PrintPreview.hwnd, SW_MAXIMIZE)
      End If
   
   
   
   waitime = Timer
   Do Until Timer > waitime + 0.1
      DoEvents
   Loop
   
   'restore the highlights to the highlighted records
   Screen.MousePointer = vbHourglass
   RestoreHighlighted
   Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Sub

lvwReport_DblClick_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lvwReport_DblClick of Form GDReportfrm"

End Sub

Private Sub lvwReport_ItemClick(ByVal Item As MSComctlLib.ListItem)
   On Error GoTo lvwReport_ItemClick_Error

   'keep record of the recorded point

    NewHighlighted& = Item.Index
   
   'display record number on status bar
    GDMDIform.StatusBar1.Panels(2) = sEmpty
    GDMDIform.StatusBar1.Panels(2) = "Result #: " & Trim$(str$(NewHighlighted&))
      
   '(This is a note for me)
   'Other examples of how to use the ItemClick Event
   'item1 = Item.SubItems(1)
   'itemtilte = Item
   'itemnum& = Item.SubItems(4)

   On Error GoTo 0
   Exit Sub

lvwReport_ItemClick_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lvwReport_ItemClick of Form GDReportfrm"
   
End Sub
Private Sub lvwReport_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   On Error GoTo lvwReport_MouseUp_Error

    'User clicked on the GDReportfrm.  This adds a point if left clicked, or moves
    'to the clicked record's position if right clicked.
    
    Screen.MousePointer = vbHourglass
    
    'restore the highlights to the previously selected records
    '(they are unselected by the new click)
    RestoreHighlighted
    
    'restore the highlighting to the clicked point
    GDReportfrm.lvwReport.ListItems(NewHighlighted&).Selected = True
      
    '(re)record ALL the highlighted point
    RecordHighlighted
    
    Select Case Button
       Case 1
          'left button clicked--nothing more to do
           
       Case 2
          'right button clicked--still need to plot record's position
          
          Call LocatePointOnMap(1)
          
       Case Else
    End Select
    
    Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Sub

lvwReport_MouseUp_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lvwReport_MouseUp of Form GDReportfrm"

End Sub

Private Sub Form_Activate()
   'If GDReportfrm.WindowState = vbMaximized Then
     'make button stay pressed
     GDMDIform.Toolbar1.Buttons(28).value = tbrPressed
     buttonstate&(28) = 1
    'End If
   'If Not Searching Then ret = ShowWindow(GDReportfrm.hWnd, SW_MAXIMIZE)
End Sub

Private Sub Form_Deactivate()
   'unpress button to encourage user to press it again inorder to activate form
   GDMDIform.Toolbar1.Buttons(28).value = tbrUnpressed
   buttonstate&(28) = 0
   If Not Searching Then Ret = ShowWindow(GDReportfrm.hwnd, SW_MINIMIZE)
End Sub
' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' source : wheel mouse hook: http://www.vbforums.com/showthread.php?388222-VB6-MouseWheel-with-Any-Control-%28originally-just-MSFlexGrid-Scrolling%29
'          two files examples: WheelHook-AllControls.zip, WheelHook-NestedControls.zip

' ===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
 
  'original WheelWheel code for interacting with very controls on the form is below
  Dim ctl As Control, cContainerCtl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
'  Dim cc

  For Each ctl In Controls
    ' Is the mouse over the control
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hwnd, Xpos, Ypos))
    On Error GoTo 0

    If bOver Then
      ' If so, respond accordingly
      bHandled = True
      Select Case True

'        Case TypeOf ctl Is MSFlexGrid
'          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
'
'        Case TypeOf ctl Is PictureBox, TypeOf ctl Is Frame
'          Set cContainerCtl = ctl
'          bHandled = False
'
'        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
'          ' These controls already handle the mousewheel themselves, so allow them to:
'          If ctl.Enabled Then ctl.SetFocus
          
        Case TypeOf ctl Is ListView
           ListViewScroll ctl, MouseKeys, Rotation, Xpos, Ypos

        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
'    Debug.Print ctl.Name
  Next ctl

'  If Not cContainerCtl Is Nothing Then
'    If TypeOf cContainerCtl Is PictureBox Then PictureBoxZoom GDform1.Picture2, MouseKeys, Rotation, Xpos, Ypos, 0
'  Else
'    ' Scroll was not handled by any controls, so treat as a general message send to the form
'    GDMDIform.StatusBar1.Panels(1) = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
'  End If
End Sub
