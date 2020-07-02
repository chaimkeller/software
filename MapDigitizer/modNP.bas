Attribute VB_Name = "modNP"
'Sub ShortPath(sPath As String, MaxDirLen As Integer, sShortPath As String)
'   'This routine finds abreviated path names to fit in
'   'the plot buffer list box
'
'      sPath = Mid$(sPath, 1, Len(sPath) - 1) 'first trim off final "\"
'      If Len(sPath) > MaxDirLen Then 'try to find short version
'         pos1% = InStr(sPath, "\")
'         'find drive letter
'         sDriveLetter = Mid$(sPath, 1, pos1%)
'         'Now find abbreviated form of the path for
'         'displaying in the list box.
'         'Abbreviated version just contains most inner directory
'         If pos1% <> 0 Then
'            pos1% = InStr(pos1% + 1, sPath, "\")
'            If pos1% <> 0 Then
'               For i% = Len(sPath) To 1 Step -1
'                  If Mid$(sPath, i%, 1) = "\" Then
'                     sShortPath = sDriveLetter & "...\" & Mid$(sPath, i% + 1, Len(sPath) - i%) & "\"
'                     Exit For
'                     End If
'               Next i%
'            Else
'               sShortPath = sDriveLetter & "...\"
'               End If
'            End If
'      Else
'         sShortPath = sPath & "\" 'put back final "\"
'         End If
'
'End Sub

Sub OpenRead(numfil%)

    'This routine opens the plot files and reads in the data
    'Data format is determined by the stored values of the
    'different formats in FilForm.  There are 11 formats.
    'The current one is number PlotInfo(0, numfil%)
    
    On Error GoTo errhand
    
    Dim Data() As Double, Xvalue As Double, Yvalue As Double
    Dim doclin$
    
    Screen.MousePointer = vbHourglass
    PlotFileName$ = PlotInfo(7, numfil%)
    freefil% = FreeFile
    Open PlotFileName$ For Input As #freefil%
    'skip the header lines
'    For i% = 1 To FilForm(0, PlotInfo(0, numfil%))
'       Line Input #freefil%, doclin$
'    Next i%
    
    Dim numRows%, pos1%, pos2%, sData$, foundX%, foundY%
    numRows% = 0
    
'    Select Case FilForm(1, val(PlotInfo(0, numfil%)))
'       Case 0 'common separated mixed strings and numbers
'          Do Until EOF(freefil%)
'             numRows% = numRows% + 1
'             pos2% = 1
'             foundX% = 0
'             foundY% = 0
'             Line Input #freefil%, doclin$
'             For j% = 1 To FilForm(2, val(PlotInfo(0, numfil%))) - 1
'                pos1% = InStr(pos2%, doclin$, ",")
'                sData$ = Mid$(doclin$, pos2%, pos1% - pos2%)
'                pos2% = pos1% + 1
'                If j% = FilForm(3, val(PlotInfo(0, numfil%))) Then
'                   Xvalue = val(sData$)
'                   foundX% = 1
'                   If foundX% = 1 And foundY% = 1 Then Exit For
'                   End If
'                If j% = FilForm(4, val(PlotInfo(0, numfil%))) Then
'                   Yvalue = val(sData$)
'                   foundY% = 1
'                   If foundX% = 1 And foundY% = 1 Then Exit For
'                   End If
'             Next j%
'             'parse last bit of the line if haven't yet
'             'read the X,Y values
'             If foundX% <> 1 Or foundY% <> 1 Then
'                pos2% = pos1% + 1
'                sData$ = Mid$(doclin$, pos2%, Len(doclin$) - pos2% + 1)
'                If j% = FilForm(3, val(PlotInfo(0, numfil%))) Then
'                   Xvalue = val(sData$)
'                   foundX% = 1
'                   End If
'                If j% = FilForm(4, val(PlotInfo(0, numfil%))) Then
'                   Yvalue = val(sData$)
'                   foundY% = 1
'                   End If
'                End If
'
'                If numRows% > numRowsToNow% Then
'                   ReDim Preserve dPlot(maxFilesToPlot%, 1, numRows% - 1)
'                   numRowsToNow% = numRows%
'                   End If
'             dPlot(numfil%, 0, numRows% - 1) = Xvalue / val(PlotInfo(3, numfil%)) + val(PlotInfo(4, numfil%))
'             dPlot(numfil%, 1, numRows% - 1) = Yvalue / val(PlotInfo(5, numfil%)) + val(PlotInfo(6, numfil%))
'
'          Loop
'          Close (freefil%)
'
'       Case 1 'delimited row of numbers
'          'read one row having PlotInfo(2, numfil%) columns
'          Do Until EOF(freefil%)
'             For j% = 1 To FilForm(2, val(PlotInfo(0, numfil%)))
'                 ReDim Preserve Data(j% - 1)
'                 Input #freefil%, Data(j% - 1)
'             Next j%
'             Xvalue = Data(FilForm(3, val(PlotInfo(0, numfil%))) - 1)
'             Yvalue = Data(FilForm(4, val(PlotInfo(0, numfil%))) - 1)
'             numRows% = numRows% + 1
'             If numRows% > numRowsToNow% Then
'                ReDim Preserve dPlot(maxFilesToPlot%, 1, numRows% - 1)
'                numRowsToNow% = numRows%
'                End If
'             dPlot(numfil%, 0, numRows% - 1) = Xvalue / val(PlotInfo(3, numfil%)) + val(PlotInfo(4, numfil%))
'             dPlot(numfil%, 1, numRows% - 1) = Yvalue / val(PlotInfo(5, numfil%)) + val(PlotInfo(6, numfil%))
'          Loop
'          Close (freefil%)
'          ReDim Data(0) 'reclaim memory
'    End Select
    Screen.MousePointer = vbDefault

Exit Sub

errhand:
    Screen.MousePointer = vbDefault
    MsgBox "Encountered error number: " & Err.Number & vbLf & _
           "while reading the file.  The error description is:" & vbLf & _
           Err.Description & vbLf & _
           "Closing the file " & PlotFileName$ & vbLf & _
           "Now resuming plotting.", vbExclamation + vbOKOnly, "Plot"
           Close (freefil%)
           ReDim Data(0)
           
End Sub

Sub DblClickForm()

'restore plotting defaults changed during drag
'or changed while replotting

Screen.MousePointer = vbHourglass

'If drm% <> 0 Then
'   frmDraw.DrawMode = drm%
'   frmDraw.DrawStyle = drs%
'   frmDraw.DrawWidth = drw%
'   End If
'
'If numFilesToPlot% <= 0 Then
'   Screen.MousePointer = vbDefault
'   MsgBox "Sorry, you need to rerun the plot wizard!", vbExclamation + vbOKOnly, "Plot"
'   Exit Sub
'   End If

'Call frmSetCond.DefineLayout
'Plot frmDraw, dPlot, udtMyGraphLayout

Screen.MousePointer = vbDefault

End Sub
