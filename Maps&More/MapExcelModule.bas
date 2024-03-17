Attribute VB_Name = "MapExcelModule"
'---------------------------------------------------------------------------------------
' Module    : MapExcelModule
' DateTime  : 8/4/2003 19:14
' Author    : Chaim Keller
' Purpose   : Contains all the routines that use EXCEL
'---------------------------------------------------------------------------------------
'exports 3D topo data to Excel for plotting
Sub ExportToExcel(drag1x, drag1y, drag2x, drag2y)

   Dim KMXStep As Double, KMYStep As Double
   Dim invstepx As Double, invstepy As Double
   Dim kmxsteps As Double, kmysteps As Double
   
   On Error GoTo errhand
   
   'convert screen coordinates to geo coordinates
   Call ScreenToGeo(drag1x, drag1y, kmxDrag, kmyDrag, 1, ier%)
   If kmxDrag = 0 And kmyDrag = 0 Then
      MsgBox "Coordinate system not supported", vbExclamation + vbOKOnly, "Maps & More"
      Exit Sub 'coordinate system not supported
      End If
   kmxUL = kmxDrag
   kmyUL = kmyDrag
   Call ScreenToGeo(drag2x, drag2y, kmxDrag, kmyDrag, 1, ier%)
   If kmxDrag = 0 And kmyDrag = 0 Then
      MsgBox "Coordinate system not supported", vbExclamation + vbOKOnly, "Maps & More"
      Exit Sub 'coordinate system not supported
      End If
   kmxLR = kmxDrag
   kmyLR = kmyDrag
   
  
   If Not world Then
      KMXStep = 25
      KMYStep = 25
   Else
      KMXStep = XDIM
      KMYStep = YDIM
      End If
      
   'convert to appropriate grid
   invstepx = 1# / KMXStep
   invstepy = 1# / KMYStep
   kmxUL = CLng(kmxUL * invstepx) * KMXStep
   kmxLR = CLng(kmxLR * invstepx) * KMXStep
   kmyUL = CLng(kmyUL * invstepy) * KMYStep
   kmyLR = CLng(kmyLR * invstepy) * KMYStep
   
      
      
    Select Case MsgBox("Do you want to export to EXCEL file?" _
                       & vbCrLf & "" _
                       & vbCrLf & "(Answer ""No"" to export to xyz file)" _
                       , vbYesNoCancel Or vbExclamation Or vbDefaultButton1, "Export")
    
      Case vbYes
         'do nothing (handler below after select)
    
      Case vbNo
         'export to xyz file and then exit
          Select Case MsgBox("Export to xyz height file (ASCII text file with x,y,z columns)?", vbOKCancel Or vbQuestion Or vbDefaultButton1, "Export to xyz file")
         
            Case vbOK
               noExcel = True
               'backup to xyz file
               Maps.CommonDialog2.CancelError = True
               Maps.CommonDialog2.Filter = "xyz files (*.xyz)|*.xyz|"
               Maps.CommonDialog2.FilterIndex = 1
               Maps.CommonDialog2.ShowSave

               FileName = Maps.CommonDialog2.FileName
               filnum% = FreeFile
               
               Dim YColumn As Boolean
               frmMsgBox.MsgCstm "Inside loop for export", "Y or X?", mbQuestion, 1, False, _
                                 "Y columns are inside loop", "X rows are inside loop", "Cancel"
               Select Case frmMsgBox.g_lBtnClicked
                   Case 1 'the 1st button in your list was clicked
                        YColumn = True
            
                   Case 2 'the 2nd button in your list was clicked
                        YColumn = False
                    
                  Case 0, 3 'cancel.
                        Close
                        Exit Sub
               End Select
               
               Screen.MousePointer = vbHourglass
               
               With mapprogressfm
                    .Visible = True
                    .Text1.Visible = False
                    .Text2.Visible = False
                    .Acceptbut.Visible = False
                    .Command2.Visible = False
                    .frmDTM.Visible = False
                    .Caption = "Exporting xyz file, 0%"
                    .StatusBar1.Panels(1).Text = "Please wait..."
                    .ProgressBar1.Min = 0
                    .ProgressBar1.Max = CLng((kmxLR - kmxUL) / KMXStep) + 1
                    .ProgressBar1.value = 0
                    .Left = Maps.Left + 0.5 * Maps.Width - 0.5 * .Width
                    .Top = Maps.Top + 0.5 * Maps.Height - 0.5 * .Height
               End With
               
               ProgressView = True

               Open FileName For Output As #filnum%
               
                  stepNum& = 0
                  
                  If YColumn Then
                                   
                      For kmxsteps = kmxUL To kmxLR Step KMXStep 'kmystep = kmyLR To kmyUL Step 25
                            
                          For kmysteps = kmyLR To kmyUL Step KMYStep 'kmxstep = kmxUL To kmxLR Step 25
                              
                              'determine heights at those coordinates
                              If Not world Then
                                If noheights = False Then
                                   kmxpoint = kmxsteps
                                   kmypoint = kmysteps
                                   Call heights(kmxpoint, kmypoint, hgt2)
                                ElseIf noheights = True Then
                                   hgt2 = 0#
                                   End If
                              Else
                                 
                                If noheights = False Then
                                   kmxpoint = kmxsteps
                                   kmypoint = kmysteps
                                   Call worldheights(kmxpoint, kmypoint, hgt2)
                                   If hgt2 = -9999 Then hgt2 = 0
                                Else
                                   hgt2 = 0
                                   End If
                                 
                                 End If
                                 
                                 Write #filnum%, kmxsteps, kmysteps, hgt2
                              
                          Next kmysteps 'kmxstep
                          
                          numSteps& = numSteps& + 1
                          mapprogressfm.ProgressBar1.value = numSteps&
                          mapprogressfm.Caption = "Exporting xyz file, " + Str$(CLng(100 * numSteps& / mapprogressfm.ProgressBar1.Max)) + "%"
                          mapprogressfm.Label1.Caption = Str$(CLng(100 * numSteps& / mapprogressfm.ProgressBar1.Max)) + "%"
                          mapprogressfm.Refresh
    
                        Next kmxsteps 'kmystep
              
                    Else
                            
                      For kmysteps = kmyLR To kmyUL Step KMYStep 'kmxstep = kmxUL To kmxLR Step 25
                          
                          For kmxsteps = kmxUL To kmxLR Step KMXStep
                              
                              'determine heights at those coordinates
                              If Not world Then
                                If noheights = False Then
                                   kmxpoint = kmxsteps
                                   kmypoint = kmysteps
                                   Call heights(kmxpoint, kmypoint, hgt2)
                                ElseIf noheights = True Then
                                   hgt2 = 0#
                                   End If
                              Else
                                 
                                If noheights = False Then
                                   kmxpoint = kmxsteps
                                   kmypoint = kmysteps
                                   Call worldheights(kmxpoint, kmypoint, hgt2)
                                   If hgt2 = -9999 Then hgt2 = 0
                                Else
                                   hgt2 = 0
                                   End If
                                 
                                 End If
                                 
                                 Write #filnum%, kmysteps, kmxsteps, hgt2
                              
                          Next kmxsteps 'kmxstep
                          
                          numSteps& = numSteps& + 1
                          mapprogressfm.ProgressBar1.value = numSteps&
                          mapprogressfm.Caption = "Exporting xyz file, " + Str$(CLng(100 * numSteps& / mapprogressfm.ProgressBar1.Max)) + "%"
                          mapprogressfm.Label1.Caption = Str$(CLng(100 * numSteps& / mapprogressfm.ProgressBar1.Max)) + "%"
                          mapprogressfm.Refresh
    
                        Next kmysteps 'kmystep
                    
                       End If
                       
               Close #filnum%
               Unload mapprogressfm
               ProgressView = False
               Screen.MousePointer = vbDefault
              
             Case vbCancel
         
          End Select
          
         Exit Sub
    
      Case vbCancel
         Exit Sub
    
    End Select
    
'''''''''''''''''''''''''''''''''comment out if don't have EXCEL.exe registered as a component'''''''''''''''''''''''''
     
    Screen.MousePointer = vbHourglass
    
'    Use the Excel (OLE) Object library
    
    Dim ExcelApp As Excel.Application
    Dim ExcelBook As Excel.Workbook
    Dim ExcelSheet As Excel.Worksheet

    Set ExcelApp = New Excel.Application
    Set ExcelBook = ExcelApp.Workbooks.Add
    Set ExcelSheet = ExcelBook.Worksheets.Add

    Screen.MousePointer = vbDefault
    ExcelBook.Application.Visible = True
    ExcelBook.Windows(1).Visible = True
    middle& = CLng(Abs(kmxUL - kmxLR) * 0.5 / KMXStep)
    ExcelSheet.Cells( _
        1, middle&).value = "Maps&More export file, Date/Time: " & Now()

    'Headers (X coordinate)
    j& = 0
    For kmxExcel = kmxUL To kmxLR Step KMXStep
        j& = j& + 1
        ExcelSheet.Cells(3, j& + 1) = kmxExcel
    Next kmxExcel


    numRow& = 3 'starting row for data is numRow&+1

    'Y coordinates and heights
    For kmyExcel = kmyUL To kmyLR Step -KMYStep

        numRow& = numRow& + 1
        ExcelSheet.Cells(numRow&, 1) = kmyExcel 'Y Coordinate

        i& = 0
        For kmxExcel = kmxUL To kmxLR Step KMXStep
            i& = i& + 1

            'determine heights at those coordinates
            If Not world Then
              If noheights = False Then
                 kmxExcel0 = kmxExcel
                 kmyExcel0 = kmyExcel
                 Call heights(kmxExcel0, kmyExcel0, hgt2)
              ElseIf noheights = True Then
                 hgt2 = 0#
                 End If
            Else

              If noheights = False Then
                 kmxExcel0 = kmxExcel
                 kmyExcel0 = kmyExcel
                 Call worldheights(kmxExcel0, kmyExcel0, hgt2)
                 If hgt2 = -9999 Then hgt2 = 0
              Else
                 hgt2 = 0
                 End If

               End If

            ExcelSheet.Cells(numRow&, i& + 1) = hgt2
        Next kmxExcel
     Next kmyExcel

    ExcelSheet.SaveAs drivjk$ & "dtmpiec.xls"
    Screen.MousePointer = vbDefault

    response = MsgBox( _
        "Do you wan't to close the EXCEL window? " + _
        "(If you answer No, then EXCEL will continue running, even after " + _
        "closing Maps & More.)", vbQuestion + vbYesNo + vbDefaultButton2, _
        "Maps&More")
    If response = vbYes Then
       ExcelApp.Quit

       Set ExcelApp = Nothing
       Set ExcelBook = Nothing
       Set ExcelSheet = Nothing
       End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub

errhand:

    Screen.MousePointer = vbDefault
    
    If ProgressView Then
        ProgressView = False
        Unload mapprogressfm
        Close #filnum%
        End If
    
    If noExcel Then Exit Sub

    MsgBox "Error Number: " & Str$(Err.Number) & " encountered!" & vbLf & _
           Err.Description & vbLf & _
           "Excel window will be closed!", vbCritical + vbOKOnly, "Maps&More"
    ExcelApp.Quit

    Set ExcelApp = Nothing
    Set ExcelBook = Nothing
    Set ExcelSheet = Nothing
           
   outfil$ = drivjk$ & "dtmpiec2.out"
   filnum% = FreeFile
   Open outfil$ For Output As #filnum%
   KMXStep = 25
   KMYStep = 25
   doclin2$ = ""
   nn% = 0
   For kmxExcel = kmxUL To kmxLR Step KMXStep
      If nn% = 0 Then
         doclin2$ = "---," + Str$(kmxExcel)
      Else
         doclin2$ = doclin2$ + "," + Str$(kmxExcel)
         End If
      nn% = nn% + 1
   Next kmxExcel
   Print #filnum%, doclin2$

   For kmyExcel = kmyLR To kmyUL Step KMYStep
      kmyn = kmyExcel
      nn% = 0
      doclin2$ = ""
      For kmxExcel = kmxUL To kmxLR Step KMXStep
          'determine heights at those coordinates

          If Not world Then
            If noheights = False Then
               kmxExcel0 = kmxExcel
               kmyExcel0 = kmyExcel
               Call heights(kmxExcel0, kmyExcel0, hgt2)
            ElseIf noheights = True Then
               hgt2 = 0#
               End If
          Else

            If noheights = False Then
               kmxExcel0 = kmxExcel
               kmyExcel0 = kmyExcel
               Call worldheights(kmxExcel0, kmyExcel0, hgt2)
               If hgt2 = -9999 Then hgt2 = 0
            Else
               hgt2 = 0
               End If

             End If

          If nn% = 0 Then
             doclin2$ = Str$(kmyn) + "," + Str$(hgt2)
          Else
             doclin2$ = doclin2$ + "," + Str$(hgt2)
             End If
          nn% = nn% + 1
       Next kmxExcel
       Print #filnum%, doclin2$
   Next kmyExcel
   Close #filnum%
   
End Sub

Sub TrigPointAdjust(drag1x, drag1y, drag2x, drag2y)
   'Adjust the DTM for heights at trig points.
   'Do the adjustment in the rectangular UL-LR: (drag1x,drag2x)-(drag2y,drag2y).
   'ITM coordinates of trig points are kmxTrig, kmyTrig.
   
   'The mountains are assumed to have some fractal shape with
   'a characteristic exponent L.  So for a distance D/25 m from the trig
   'point with DTM height Ho, it's new height will be modeled as
   'H = Ho + (Htrig - Ho) * L ^ -D.  First try L = e
   
   On Error GoTo errhand
   
   Dim exponent As Single
   
   'determine exponent
   response = InputBox("Input the base of exponent." & vbLf & _
                       "If the base is zero, then all the heights" & vbLf & _
                       "in the drag region will be equal to the" & vbLf & _
                       "to theheight of the trig point." & vbLf & vbLf & _
                       "The base is:", "Exponent", "1.4", 6450)
   If response = sEmpty Then
      Exit Sub
   Else
      exponent = Val(response)
      End If
   
   'Convert drag coordinates to kmx
   Call ScreenToGeo(drag1x, drag1y, kmxDrag, kmyDrag, 1, ier%)
   KmxStart = kmxDrag
   KmyEnd = kmyDrag

   Call ScreenToGeo(drag2x, drag2y, kmxDrag, kmyDrag, 1, ier%)
   KmxEnd = kmxDrag
   KmyStart = kmyDrag
   
   'Now determine end and start with respect to actual DTM
   'points:
   KmxStart = CLng(KmxStart * 0.04) * 25#
   KmxEnd = CLng(KmxEnd * 0.04) * 25#
   KmyStart = CLng(KmyStart * 0.04) * 25#
   KmyEnd = CLng(KmyEnd * 0.04) * 25#
   
   '---------------------change to xyz------------------------------
   
'   '------------dump before and after pictures to Excel--------------------
'   'dump to Excel the original height profile, and then
'   'the changed height profile
'
'    response = MsgBox("Dump before and after 3D profiles to Excel?", _
'                    vbQuestion + vbYesNoCancel + vbDefaultButton2, "Maps&More")
'    If response <> vbYes Then GoTo tp500
'
'    Screen.MousePointer = vbHourglass
'
'    'Use the Excel (OLE) Object library
'    Dim ExcelApp As Excel.Application
'    Dim ExcelBook As Excel.Workbook
'    Dim ExcelSheet As Excel.Worksheet
'
'    Set ExcelApp = New Excel.Application
'    Set ExcelBook = ExcelApp.Workbooks.Add
'    Set ExcelSheet = ExcelBook.Worksheets.Add
'
'    Screen.MousePointer = vbHourglass
'
'    ExcelBook.Application.Visible = True
'    ExcelBook.Windows(1).Visible = True
'    middle& = CLng(Abs(kmxEnd - kmxStart) * 0.002)
'    ExcelSheet.Cells( _
'        1, middle&).Value = "Maps&More Trig Point export file, Date/Time: " & Now()
'
'    'Headers (X coordinate)
'    j& = 0
'    For kmxExcel = kmxStart To kmxEnd Step 25
'        j& = j& + 1
'        ExcelSheet.Cells(3, j& + 1) = kmxExcel
'    Next kmxExcel
'
'    numRow& = 3 'starting row for data is numRow&+1
'
'    'Y coordinates and heights
'    For kmyExcel = kmyStart To kmyEnd Step 25
'
'        numRow& = numRow& + 1
'        ExcelSheet.Cells(numRow&, 1) = kmyExcel 'Y Coordinate
'
'        i& = 0
'        For kmxExcel = kmxStart To kmxEnd Step 25
'            i& = i& + 1
'
'            'determine heights at those coordinates
'            If Not world Then
'              If noheights = False Then
'                 kmxExcel0 = kmxExcel
'                 kmyExcel0 = kmyExcel
'                 Call heights(kmxExcel0, kmyExcel0, hgt2)
'              ElseIf noheights = True Then
'                 hgt2 = 0#
'                 End If
'            Else
'
'              If noheights = False Then
'                 kmxExcel0 = kmxExcel
'                 kmyExcel0 = kmyExcel
'                 Call worldheights(kmxExcel0, kmyExcel0, hgt2)
'                 If hgt2 = -9999 Then hgt2 = 0
'              Else
'                 hgt2 = 0
'                 End If
'
'               End If
'
'            ExcelSheet.Cells(numRow&, i& + 1) = hgt2
'        Next kmxExcel
'     Next kmyExcel
'
'     'now record altered heights on the same Excel sheet
'     'Headers (X coordinate)
'     numRow0& = numRow&
'     j& = 0
'     For kmxExcel = kmxStart To kmxEnd Step 25
'        j& = j& + 1
'        ExcelSheet.Cells(numRow0& + 3, j& + 1) = kmxExcel
'     Next kmxExcel
'
'   numRow& = numRow0& + 3
'
'   For kmy = kmyStart To kmyEnd Step 25
'
'      numRow& = numRow& + 1
'      ExcelSheet.Cells(numRow&, 1) = kmy 'Y Coordinate
'
'      i& = 0
'      For kmx = kmxStart To kmxEnd Step 25
'         i& = i& + 1
'
'         'old DTM height as this point
'         If noheights = False Then
'            kmx1 = kmx
'            kmy1 = kmy
'            Call heights(kmx1, kmy1, hgt)
'         ElseIf noheights = True Then
'            hgt = 0#
'            End If
'
'         If exponent <> 0 Then
'            D = Sqr((kmx - kmxTrig) ^ 2 + (kmy - kmyTrig) ^ 2)
'            'divide by typical scaling distance of 25 meter
'            D = D * 0.04
'            hgtNew = hgt + (hgtTrig - hgt) * exponent ^ (-D)
'         Else
'            hgtNew = hgtTrig
'            End If
'
'         ExcelSheet.Cells(numRow&, i& + 1) = hgtNew
'
'      Next kmx
'   Next kmy
'
'
'   ExcelSheet.SaveAs drivjk$ & "trigpiec.xls"
'   Screen.MousePointer = vbDefault
'
'    response = MsgBox( _
'        "Do you wan't to close the EXCEL window? " + _
'        "(If you answer No, then EXCEL will continue running, even after " + _
'        "closing Maps & More.)", vbQuestion + vbYesNo + vbDefaultButton2, _
'        "Maps&More")
'    If response = vbYes Then
'       ExcelApp.Quit
'
'       Set ExcelApp = Nothing
'       Set ExcelBook = Nothing
'       Set ExcelSheet = Nothing
'       End If
       
'--------------------------------uncomment if have excel.ocx------------------------------
       
       
tp500: '-----------save the points to the DTM tile----------------------
       response = MsgBox("Backup DTM files before saving changes?" & vbLf & _
                     "(The date will be added as a suffix to the backup tiles)", _
                     vbQuestion + vbYesNoCancel, "Maps&More")
       If response = vbYes Then
          backup% = 1
          End If
               
          
        'determine which tile(s) are being used and back them up
        CHFind$ = sEmpty
        For kmy = KmyStart To KmyEnd Step 25
           For kmx = KmxStart To KmxEnd Step 25
              kmxDTM = kmx * 0.001
              kmyDTM = (kmy - 1000000) * 0.001
              IKMX& = Int((kmxDTM + 20!) * 40!) + 1
              IKMY& = Int((380! - kmyDTM) * 40!) + 1
              NROW% = IKMY&: NCOL% = IKMX&

              'FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
              Jg% = 1 + Int((NROW% - 2) / 800)
              Ig% = 1 + Int((NCOL% - 2) / 800)
              
             'Since roundoff errors in converting from coord to
             'integer indexes, just count columns and rows assuming
             'that the first one has no roundoff error
              If kmx = KmxStart And kmy = KmyStart Then
                 IR% = NROW% - (Jg% - 1) * 800
                 IC% = NCOL% - (Ig% - 1) * 800
                 IR0% = IR%
                 IC0% = IC%
              Else
                 IC% = CInt((kmx - KmxStart) * 0.04) + IC0%
                 IR% = IR0% - CInt((kmy - KmyStart) * 0.04)
                 End If
              
              IFN& = (IR% - 1) * 801! + IC%
              
              CHFindTmp$ = CHMAP(Ig%, Jg%)
tp250:        If CHFindTmp$ <> CHFind$ Then
                 newtile% = 1
                 CHFind$ = CHFindTmp$
                 
                 If backup% = 1 Then
                    'back it up if not already backed up
                    FirstTry% = 0
                    If Dir(israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)) = sEmpty Then
                       FileCopy israeldtm + ":\dtm\" & CHFind$, israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                    Else 'warn before overwriting last backup
                       response = MsgBox("File with backup name already exists!" & vbLf & _
                              "Do you want to overwrite it?", vbExclamation + vbYesNoCancel, "Maps&More")
                       If response = vbYes Then
                          FileCopy israeldtm + ":\dtm\" & CHFind$, israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                       Else
                          response = InputBox("Enter the name of the backup:", _
                                     "New backup tile name", israeldtm + ":\dtm\" & CHFind$ & _
                                     "_" & Month(Date) & Day(Date) & Year(Date), 6450)
                          If response = sEmpty Then
                             Exit Sub
                          Else
                             If Dir(response) = sEmpty Then
                                FileCopy israeldtm + ":\dtm\" & CHFind$, israeldtm + ":\dtm\" & response
                             Else
                                GoTo tp250
                                End If
                             End If
                          End If
                       End If
                    End If
                    
                    Close 'close any open files
                    CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
                    'open the tile for writing
                    filn% = FreeFile
                    Open israeldtm + ":\dtm\" & CHFind$ For Random As #filn% Len = 2
                    
                 End If
                 
               'old DTM height as this point
               If noheights = False Then
                  kmx1 = kmx
                  kmy1 = kmy
                  Call heights(kmx1, kmy1, hgt)
               ElseIf noheights = True Then
                  hgt = 0#
                  End If
    
               If exponent <> 0 Then
                  'calculate new height at this point
                  D = Sqr((kmx - kmxTrig) ^ 2 + (kmy - kmyTrig) ^ 2)
                  'divide by typical scaling distance of 25 meter
                  D = D * 0.04
                  hgtNew = hgt + (hgtTrig - hgt) * exponent ^ (-D)
               Else
                  hgtNew = hgtTrig
                  End If
             
               'write the changes to the DTM tile
               'Since roundoff errors in converting from coord to
               'integer indexes, just count columns and rows assuming
               'that the first one has no roundoff error
                If (kmx = KmxStart And kmy = KmyStart) Or newtile% = 1 Then
                   IR% = NROW% - (Jg% - 1) * 800
                   IC% = NCOL% - (Ig% - 1) * 800
                   IR0% = IR%
                   IC0% = IC%
                   newtile% = 0
                Else
                   IC% = CInt((kmx - KmxStart) * 0.04) + IC0%
                   IR% = IR0% - CInt((kmy - KmyStart) * 0.04)
                   End If
              
                IFN& = (IR% - 1) * 801! + IC%
                Put #filn%, IFN&, CInt(hgtNew * 10)
              
           Next kmx
        Next kmy
          
       Close #filn%
       CHMNEO = sEmpty
          
    Exit Sub
    
errhand:

    Screen.MousePointer = vbDefault
    cc = Err.Number
    
    Close 'close any open files
    Select Case Err.Number
       Case 55 'can't save file to this directory
          If FirstTry% = 0 Then 'try once again
             FirstTry% = 1
             Resume
             End If
          response = MsgBox("Can't save the tile to the dtm directory!" & vbLf & _
                 "Do you want to save it to a different directory?", _
                 vbExclamation + vbYesNoCancel, "Maps & More")
          If response = vbYes Then
             res = InputBox("Input directory name." & vbLf & _
                 "Add a final backslashes '\'" & vbLf & _
                 "Example: ""c:\windows\temp\""", "Backup Tile", _
                 israeldtm + ":\dtm\", 6450)
             If res = sEmpty Then
                Exit Sub
             Else
                FileCopy israeldtm + ":\dtm\" & CHFind$, res & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                Resume Next
                End If
             End If
       Case 75 'can't save to tile
          MsgBox "The dtm tile is read only and can't be edited!", vbCritical + vbOKOnly, "Maps&More"
          Exit Sub
       Case Else
            Close
            MsgBox "Error Number: " & Str$(Err.Number) & " encountered!" & vbLf & _
                   Err.Description & vbLf & _
                   "Current procedure will be aborted!", vbCritical + vbOKOnly, "Maps&More"
            
            'ExcelApp.Quit
            
            'Set ExcelApp = Nothing
            'Set ExcelBook = Nothing
            'Set ExcelSheet = Nothing
      End Select
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GlitchFix
' DateTime  : 12/26/2013 00:28
' Author    : Chaim Keller
' Purpose   : Fixes Column (mode$ = 0) or Row (mode% = 1) glitches within the drag area
'---------------------------------------------------------------------------------------
'
Sub GlitchFix(drag1x, drag1y, drag2x, drag2y, Coord As Long, Mode%)
   'Adjust the DTM for heights at trig points.
   'Do the adjustment in the rectangular UL-LR: (drag1x,drag2x)-(drag2y,drag2y).
   'ITM coordinates of trig points are kmxTrig, kmyTrig.
   
   'The mountains are assumed to have some fractal shape with
   'a characteristic exponent L.  So for a distance D/25 m from the trig
   'point with DTM height Ho, it's new height will be modeled as
   'H = Ho + (Htrig - Ho) * L ^ -D.  First try L = e
   Dim KmxStart, KmxEnd
   Dim KmyStart, KmyEnd
   
   On Error GoTo GlitchFix_Error
   
   'Convert drag coordinates to kmx
   Call ScreenToGeo(drag1x, drag1y, kmxDrag, kmyDrag, 1, ier%)
   KmxStart = kmxDrag
   KmyEnd = kmyDrag

   Call ScreenToGeo(drag2x, drag2y, kmxDrag, kmyDrag, 1, ier%)
   KmxEnd = kmxDrag
   KmyStart = kmyDrag
   
   'Now determine end and start with respect to actual DTM
   'points:
   KmxStart = CLng(KmxStart * 0.04) * 25#
   KmxEnd = CLng(KmxEnd * 0.04) * 25#
   KmyStart = CLng(KmyStart * 0.04) * 25#
   KmyEnd = CLng(KmyEnd * 0.04) * 25#
   
'   KmxStart = 163500
'   KmxEnd = 175000
   
   '---------------------change to xyz-----------------------------
   
   '------------------enable if have EXCEL.ocx----------------------------
   
'   '------------dump before and after pictures to Excel--------------------
'   'dump to Excel the original height profile, and then
'   'the changed height profile
'
'    response = MsgBox("Dump before and after 3D profiles to Excel?", _
'                    vbQuestion + vbYesNoCancel + vbDefaultButton2, "Maps&More")
'    If response <> vbYes Then GoTo tp500
'
'    Screen.MousePointer = vbHourglass
'
'    'Use the Excel (OLE) Object library
'    Dim ExcelApp As Excel.Application
'    Dim ExcelBook As Excel.Workbook
'    Dim ExcelSheet As Excel.Worksheet
'
'    Set ExcelApp = New Excel.Application
'    Set ExcelBook = ExcelApp.Workbooks.Add
'    Set ExcelSheet = ExcelBook.Worksheets.Add
'
'    Screen.MousePointer = vbHourglass
'
'    ExcelBook.Application.Visible = True
'    ExcelBook.Windows(1).Visible = True
'    middle& = CLng(Abs(kmxEnd - kmxStart) * 0.002)
'    ExcelSheet.Cells( _
'        1, middle&).Value = "Maps&More Trig Point export file, Date/Time: " & Now()
'
'    'Headers (X coordinate)
'    j& = 0
'    For kmxExcel = kmxStart To kmxEnd Step 25
'        j& = j& + 1
'        ExcelSheet.Cells(3, j& + 1) = kmxExcel
'    Next kmxExcel
'
'    numRow& = 3 'starting row for data is numRow&+1
'
'    'Y coordinates and heights
'    For kmyExcel = kmyStart To kmyEnd Step 25
'
'        numRow& = numRow& + 1
'        ExcelSheet.Cells(numRow&, 1) = kmyExcel 'Y Coordinate
'
'        i& = 0
'        For kmxExcel = kmxStart To kmxEnd Step 25
'            i& = i& + 1
'
'            'determine heights at those coordinates
'            If Not world Then
'              If noheights = False Then
'                 kmxExcel0 = kmxExcel
'                 kmyExcel0 = kmyExcel
'                 Call heights(kmxExcel0, kmyExcel0, hgt2)
'              ElseIf noheights = True Then
'                 hgt2 = 0#
'                 End If
'            Else
'
'              If noheights = False Then
'                 kmxExcel0 = kmxExcel
'                 kmyExcel0 = kmyExcel
'                 Call worldheights(kmxExcel0, kmyExcel0, hgt2)
'                 If hgt2 = -9999 Then hgt2 = 0
'              Else
'                 hgt2 = 0
'                 End If
'
'               End If
'
'            ExcelSheet.Cells(numRow&, i& + 1) = hgt2
'        Next kmxExcel
'     Next kmyExcel
'
'     'now record altered heights on the same Excel sheet
'     'Headers (X coordinate)
'     numRow0& = numRow&
'     j& = 0
'     For kmxExcel = kmxStart To kmxEnd Step 25
'        j& = j& + 1
'        ExcelSheet.Cells(numRow0& + 3, j& + 1) = kmxExcel
'     Next kmxExcel
'
'   numRow& = numRow0& + 3
'
'   For kmy = kmyStart To kmyEnd Step 25
'
'      numRow& = numRow& + 1
'      ExcelSheet.Cells(numRow&, 1) = kmy 'Y Coordinate
'
'      i& = 0
'      For kmx = kmxStart To kmxEnd Step 25
'         i& = i& + 1
'
'         'old DTM height as this point
'         If Mode% = 0 And kmx = Coord Then
'            If noheights = False Then
'               kmx1 = kmx - 25
'               kmy1 = kmy
'               Call heights(kmx1, kmy1, hgt1)
'               kmx2 = kmx + 25
'               kmy2 = kmy
'               Call heights(kmx2, kmy2, hgt2)
'               hgtNew = 0.5 * (hgt2 - hgt1) + hgt1
'            ElseIf noheights = True Then
'               hgtNew = 0#
'               End If
'         ElseIf Mode% = 1 And kmy = Coord Then
'            If noheights = False Then
'               kmx1 = kmx
'               kmy1 = kmy - 25
'               Call heights(kmx1, kmy1, hgt1)
'               kmx2 = kmx
'               kmy2 = kmy + 25
'               Call heights(kmx2, kmy2, hgt2)
'               hgtNew = 0.5 * (hgt2 - hgt1) + hgt1
'            ElseIf noheights = True Then
'               hgtNew = 0#
'               End If
'          Else 'use original coordinates
'            If noheights = False Then
'               kmx1 = kmx
'               kmy1 = kmy
'               Call heights(kmx1, kmy1, hgt1)
'               hgtNew = hgt1
'            Else
'               hgtNew = 0#
'               End If
'            End If
'
'         ExcelSheet.Cells(numRow&, i& + 1) = hgtNew
'
'      Next kmx
'   Next kmy
'
'
'   ExcelSheet.SaveAs drivjk$ & "trigpiec.xls"
'   Screen.MousePointer = vbDefault
'
'    response = MsgBox( _
'        "Do you wan't to close the EXCEL window? " + _
'        "(If you answer No, then EXCEL will continue running, even after " + _
'        "closing Maps & More.)", vbQuestion + vbYesNo + vbDefaultButton2, _
'        "Maps&More")
'    If response = vbYes Then
'       ExcelApp.Quit
'
'       Set ExcelApp = Nothing
'       Set ExcelBook = Nothing
'       Set ExcelSheet = Nothing
'       End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
tp500: '-----------save the points to the DTM tile----------------------
       response = MsgBox("Backup DTM files before saving changes?" & vbLf & _
                     "(The date will be added as a suffix to the backup tiles)", _
                     vbQuestion + vbYesNoCancel, "Maps&More")
       If response = vbYes Then
          backup% = 1
          End If
               
          
        'determine which tile(s) are being used and back them up
        CHFind$ = sEmpty
        For kmy = KmyStart To KmyEnd Step 25
           For kmx = KmxStart To KmxEnd Step 25
              kmxDTM = kmx * 0.001
              kmyDTM = (kmy - 1000000) * 0.001
              IKMX& = Int((kmxDTM + 20!) * 40!) + 1
              IKMY& = Int((380! - kmyDTM) * 40!) + 1
              NROW% = IKMY&: NCOL% = IKMX&

              'FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
              Jg% = 1 + Int((NROW% - 2) / 800)
              Ig% = 1 + Int((NCOL% - 2) / 800)
              
             'Since roundoff errors in converting from coord to
             'integer indexes, just count columns and rows assuming
             'that the first one has no roundoff error
              If kmx = KmxStart And kmy = KmyStart Then
                 IR% = NROW% - (Jg% - 1) * 800
                 IC% = NCOL% - (Ig% - 1) * 800
                 IR0% = IR%
                 IC0% = IC%
              Else
                 IC% = CInt((kmx - KmxStart) * 0.04) + IC0%
                 IR% = IR0% - CInt((kmy - KmyStart) * 0.04)
                 End If
              
              IFN& = (IR% - 1) * 801! + IC%
              
              CHFindTmp$ = CHMAP(Ig%, Jg%)
tp250:        If CHFindTmp$ <> CHFind$ Then
                 CHFind$ = CHFindTmp$
                 
                 If backup% = 1 Then
                    'back it up if not already backed up
                    If Dir(israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)) = sEmpty Then
                       FileCopy israeldtm + ":\dtm\" & CHFind$, israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                       End If
                    End If
                    
                 CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
                 Close 'close any opened files (for example by sub heights)
                 filn% = FreeFile
                 Open israeldtm + ":\dtm\" & CHFind$ For Random As #filn% Len = 2
                    
                 End If
               
                 
               'old DTM height as this point
               If Mode% = 0 And kmx = Coord Then
                  If noheights = False Then
                     kmx1 = kmx - 25
                     kmy1 = kmy
                     Call heights(kmx1, kmy1, hgt1)
                     kmx2 = kmx + 25
                     kmy2 = kmy
                     Call heights(kmx2, kmy2, hgt2)
                     hgtNew = 0.5 * (hgt2 - hgt1) + hgt1
                  ElseIf noheights = True Then
                     hgtNew = 0#
                     End If
               ElseIf Mode% = 1 And kmy = Coord Then
                  If noheights = False Then
                     kmx1 = kmx
                     kmy1 = kmy - 25
                     Call heights(kmx1, kmy1, hgt1)
                     kmx2 = kmx
                     kmy2 = kmy + 25
                     Call heights(kmx2, kmy2, hgt2)
                     hgtNew = 0.5 * (hgt2 - hgt1) + hgt1
                  ElseIf noheights = True Then
                     hgtNew = 0#
                     End If
               Else 'use original coordinates
                 If noheights = False Then
                    kmx1 = kmx
                    kmy1 = kmy
                    Call heights(kmx1, kmy1, hgt1)
                    hgtNew = hgt1
                 Else
                    hgtNew = 0#
                    End If
                  End If
             
               'write the changes to the DTM tile
               'Since roundoff errors in converting from coord to
               'integer indexes, just count columns and rows assuming
               'that the first one has no roundoff error
               IR% = NROW% - (Jg% - 1) * 800
               IC% = NCOL% - (Ig% - 1) * 800
               IR0% = IR%
               IC0% = IC%
              
               IFN& = (IR% - 1) * 801! + IC%
               Put #filn%, IFN&, CInt(hgtNew * 10)
              
           Next kmx
        Next kmy
          
       If filn% > 0 Then Close #filn%
       CHMNEO = sEmpty
          
    Exit Sub
    
GlitchFix_Error:

    Screen.MousePointer = vbDefault
    cc = Err.Number
    
    Close 'close any open files
    Select Case Err.Number
       Case 55 'can't save file to this directory
          If FirstTry% = 0 Then 'try once again
             FirstTry% = 1
             Resume
             End If
          response = MsgBox("Can't save the tile to the dtm directory!" & vbLf & _
                 "Do you want to save it to a different directory?", _
                 vbExclamation + vbYesNoCancel, "Maps & More")
          If response = vbYes Then
             res = InputBox("Input directory name." & vbLf & _
                 "Add a final backslashes '\'" & vbLf & _
                 "Example: ""c:\windows\temp\""", "Backup Tile", _
                 israeldtm + ":\dtm\", 6450)
             If res = sEmpty Then
                Exit Sub
             Else
                FileCopy israeldtm + ":\dtm\" & CHFind$, res & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                Resume Next
                End If
             End If
       Case 75 'can't save to tile
          MsgBox "The dtm tile is read only and can't be edited!", vbCritical + vbOKOnly, "Maps&More"
          Exit Sub
       Case Else
            Close
            MsgBox "Error Number: " & Str$(Err.Number) & " encountered!" & vbLf & _
                   Err.Description & vbLf & _
                   "Current procedure will be aborted!", vbCritical + vbOKOnly, "Maps&More"
            
            'ExcelApp.Quit
            
            'Set ExcelApp = Nothing
            'Set ExcelBook = Nothing
            'Set ExcelSheet = Nothing
      End Select

   On Error GoTo 0
   Exit Sub


End Sub

