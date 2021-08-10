Attribute VB_Name = "NPmodule"
Sub ShortPath(sPath As String, MaxDirLen As Integer, sShortPath As String)
   'This routine finds abreviated path names to fit in
   'the plot buffer list box
      
      sPath = Mid$(sPath, 1, Len(sPath) - 1) 'first trim off final "\"
      If Len(sPath) > MaxDirLen Then 'try to find short version
         pos1% = InStr(sPath, "\")
         'find drive letter
         sDriveLetter = Mid$(sPath, 1, pos1%)
         'Now find abbreviated form of the path for
         'displaying in the list box.
         'Abbreviated version just contains most inner directory
         If pos1% <> 0 Then
            pos1% = InStr(pos1% + 1, sPath, "\")
            If pos1% <> 0 Then
               For I% = Len(sPath) To 1 Step -1
                  If Mid$(sPath, I%, 1) = "\" Then
                     sShortPath = sDriveLetter & "...\" & Mid$(sPath, I% + 1, Len(sPath) - I%) & "\"
                     Exit For
                     End If
               Next I%
            Else
               sShortPath = sDriveLetter & "...\"
               End If
            End If
      Else
         sShortPath = sPath & "\" 'put back final "\"
         End If

End Sub

Sub OpenRead(numfil%, nFile As Integer)

    'This routine opens the plot files and reads in the data
    'Data format is determined by the stored values of the
    'different formats in FilForm.  There are 11 formats.
    'The current one is number PlotInfo(0, numfil%)
    
    On Error GoTo errhand
    
    Dim Data() As Double, Xvalue As Double, Yvalue As Double
    Dim doclin$, FuncX As String, FuncY As String, pos%
    'declare sorting variables
    Dim SoundWarning As Boolean
    Dim isort As Integer, list() As Double, ListItems() As String
    
    Screen.MousePointer = vbHourglass
    PlotFileName$ = PlotInfo(7, numfil%)
    freefil% = FreeFile
    Open PlotFileName$ For Input As #freefil%
    'skip the header lines
    For I% = 1 To FilForm(0, PlotInfo(0, numfil%))
       Line Input #freefil%, doclin$
    Next I%
    
    Dim numRows%, pos1%, pos2%, sData$, foundX%, foundY%
    numRows% = 0
    
    Select Case FilForm(1, Val(PlotInfo(0, numfil%)))
       Case 0 'common separated mixed strings and numbers
          Do Until EOF(freefil%)
             numRows% = numRows% + 1
             pos2% = 1
             foundX% = 0
             foundY% = 0
10:
             If numRows% > 1 Then
                doclin0$ = doclin$
                End If
                
             Line Input #freefil%, doclin$
             
             'skip blank lines
             If Trim$(doclin$) = sEmpty Then GoTo 10
             For J% = 1 To FilForm(2, Val(PlotInfo(0, numfil%))) - 1
                pos1% = InStr(pos2%, doclin$, ",")
                sData$ = Mid$(doclin$, pos2%, pos1% - pos2%)
                pos2% = pos1% + 1
                If J% = FilForm(3, Val(PlotInfo(0, numfil%))) Then
                   Xvalue = Val(sData$)
                   foundX% = 1
                   If foundX% = 1 And foundY% = 1 Then Exit For
                   End If
                If J% = FilForm(4, Val(PlotInfo(0, numfil%))) Then
                   Yvalue = Val(sData$)
                   foundY% = 1
                   If foundX% = 1 And foundY% = 1 Then Exit For
                   End If
             Next J%
             'parse last bit of the line if haven't yet
             'read the X,Y values
             If foundX% <> 1 Or foundY% <> 1 Then
                pos2% = pos1% + 1
                sData$ = Mid$(doclin$, pos2%, Len(doclin$) - pos2% + 1)
                If J% = FilForm(3, Val(PlotInfo(0, numfil%))) Then
                   Xvalue = Val(sData$)
                   foundX% = 1
                   End If
                If J% = FilForm(4, Val(PlotInfo(0, numfil%))) Then
                   Yvalue = Val(sData$)
                   foundY% = 1
                   End If
                End If
                
                If numRows% > numRowsToNow% Then
                   ReDim Preserve dPlot(maxFilesToPlot%, 1, numRows% - 1)
                   numRowsToNow% = numRows%
                   End If
             dPlot(numfil%, 0, numRows% - 1) = Xvalue / Val(PlotInfo(3, numfil%)) + Val(PlotInfo(4, numfil%))
             dPlot(numfil%, 1, numRows% - 1) = Yvalue / Val(PlotInfo(5, numfil%)) + Val(PlotInfo(6, numfil%))
             GoSub Wrapper
             
             'check for negative progression of x values
             If numRows% - 1 >= 1 And Not SoundWarning And Not Fitting Then 'And PlotInfo(1, numfil%) = 0 Then
                If dPlot(numfil%, 0, numRows% - 1) < dPlot(numfil%, 0, numRows% - 2) Then
                   
                   'sound warning only once
                   Select Case MsgBox("Some or all of the X values of the following file are not sorted from smallest to largest." _
                                      & vbCrLf & vbCrLf & "" _
                                      & PlotInfo(7, numfil%) _
                                      & vbCrLf & vbCrLf & "This is what was found: " _
                                      & vbCrLf & "x value at row: " & Str$(numRows% - 1) & " was < than x value at row: " & Str$(numRows% - 2) _
                                      & vbCrLf & Str$(dPlot(numfil%, 0, numRows% - 1)) & " < " & Str$(dPlot(numfil%, 0, numRows% - 2)) _
                                      & vbCrLf & "The two lines in the file are: " _
                                      & vbCrLf & "first: " & doclin0$ & " and then: " & doclin$ _
                                      & vbCrLf & "" _
                                      & vbCrLf & "This will limit your plotting and fitting options." _
                                      & vbCrLf & "Do you want it sorted? (Choose ""Cancel"" to end warnings.)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "(Hint: another option is to  permanently reverse the ordeer" _
                                      & vbCrLf & "by checking the ""Reverse ""checkbox in the plot option dialog)" _
                                      , vbYesNoCancel Or vbInformation Or vbDefaultButton1, "X values not sorted")
                   
                    Case vbYes
                        SoundWarning = True
                    Case vbNo
                   
                    Case vbCancel
                        SoundWarning = True
                   End Select
                   
                   End If
                 End If
                   
          Loop
          Close (freefil%)
          ReDim Data(0) 'reclaim memory
          
          If SoundWarning Then 'sort the x,y, first roughly using the sorted ListSort list box
                               'and then finish the sorting using a bubblesort
             GoSub ReSort
             End If
          
       Case 1 'delimited row of numbers
          'read one row having PlotInfo(2, numfil%) columns
           Do Until EOF(freefil%)
          
            If numRows% > 1 Then
               Xvalue0 = Xvalue
               End If
             
             For J% = 1 To FilForm(2, Val(PlotInfo(0, numfil%)))
                 ReDim Preserve Data(J% - 1)
                 Input #freefil%, Data(J% - 1)
             Next J%
             Xvalue = Data(FilForm(3, Val(PlotInfo(0, numfil%))) - 1)
             Yvalue = Data(FilForm(4, Val(PlotInfo(0, numfil%))) - 1)
             numRows% = numRows% + 1
             If numRows% > numRowsToNow% Then
                ReDim Preserve dPlot(maxFilesToPlot%, 1, numRows% - 1)
                numRowsToNow% = numRows%
                End If
             dPlot(numfil%, 0, numRows% - 1) = Xvalue / Val(PlotInfo(3, numfil%)) + Val(PlotInfo(4, numfil%))
             dPlot(numfil%, 1, numRows% - 1) = Yvalue / Val(PlotInfo(5, numfil%)) + Val(PlotInfo(6, numfil%))
             GoSub Wrapper
             
             'check for negative progression of x values
             If numRows% - 1 >= 1 And Not SoundWarning And Not Fitting Then ' And PlotInfo(1, numfil%) = 0 Then
                If dPlot(numfil%, 0, numRows% - 1) < dPlot(numfil%, 0, numRows% - 2) Then
                   
                   'sound warning only once
                   Select Case MsgBox("Some or all of the X values of the following file are not sorted from smallest to largest." _
                                      & vbCrLf & vbCrLf & "" _
                                      & PlotInfo(7, numfil%) _
                                      & vbCrLf & vbCrLf & "This is what was found: " _
                                      & vbCrLf & "x value at row: " & Str$(numRows%) & " was < than x value at row: " & Str$(numRows% - 1) _
                                      & vbCrLf & vbCrLf _
                                      & vbCrLf & "The two x values read in are: " _
                                      & vbCrLf & Str$(dPlot(numfil%, 0, numRows% - 1)) & " < " & Str$(dPlot(numfil%, 0, numRows% - 2)) _
                                      & vbCrLf & "" _
                                      & vbCrLf & "This will limit your plotting and fitting options." _
                                      & vbCrLf & "Do you want it sorted? (Choose ""Cancel"" to end warnings.)" _
                                      & vbCrLf & "" _
                                      & vbCrLf & "(Hint: another option is to  permanently reverse the ordeer" _
                                      & vbCrLf & "by checking the ""Reverse ""checkbox in the plot option dialog)" _
                                      , vbYesNoCancel Or vbInformation Or vbDefaultButton1, "X values not sorted")
                   
                    Case vbYes
                        SoundWarning = True
                    Case vbNo
                   
                    Case vbCancel
                   
                   End Select
                       SoundWarning = True
                   End If
                 End If
          Loop
          Close (freefil%)
          ReDim Data(0) 'reclaim memory
          
          If SoundWarning Then 'sort the x,y, first roughly using the sorted ListSort list box
                               'and then finish the sorting using a bubblesort
             GoSub ReSort
             End If
             
    End Select
    
    '//////////////////fix on 02/04/20 added buffer to store the recordsize of the files being plotted////////////
    ReDim Preserve RecordSize(nFile)
    RecordSize(nFile - 1) = numRows%
    '//////////////////////////////////////////////////////////
    
    Screen.MousePointer = vbDefault

Exit Sub

Wrapper: 'handle special wrapper functions on x and y
             
    If PlotInfo(8, numfil%) <> "" Then
       pos% = InStr(PlotInfo(8, numfil%), ":")
       If pos% > 0 Then
          FuncX = Mid$(PlotInfo(8, numfil%), 1, pos% - 1)
          FuncY = Mid$(PlotInfo(8, numfil%), pos% + 1, Len(PlotInfo(8, numfil%)) - pos%)
          
          Select Case FuncX
          
             Case "none"
             Case "log"
                 If (dPlot(numfil%, 0, numRows% - 1)) > 0 Then
                    dPlot(numfil%, 0, numRows% - 1) = Log(dPlot(numfil%, 0, numRows% - 1))
                 Else
                    Call MsgBox("X value is zero or negative, and its log is undefined!" _
                                & vbCrLf & "" _
                                & vbCrLf & "You should remove this data point before fitting" _
                                & vbCrLf & "(It is on line:" & Str(numRows%) & " )" _
                                , vbInformation, "log of zero")
                    End If
             Case "exp"
                 dPlot(numfil%, 0, numRows% - 1) = Exp(dPlot(numfil%, 0, numRows% - 1))
             Case "cos"
                 dPlot(numfil%, 0, numRows% - 1) = Cos(dPlot(numfil%, 0, numRows% - 1))
             Case "sin"
                 dPlot(numfil%, 0, numRows% - 1) = Sin(dPlot(numfil%, 0, numRows% - 1))
             Case "tan"
                 dPlot(numfil%, 0, numRows% - 1) = Tan(dPlot(numfil%, 0, numRows% - 1))
             Case Else
                 
          End Select
          
          Select Case FuncY
          
             Case "none"
             Case "log"
                 If (dPlot(numfil%, 1, numRows% - 1)) > 0 Then
                    dPlot(numfil%, 1, numRows% - 1) = Log(dPlot(numfil%, 1, numRows% - 1))
                 Else
                    Call MsgBox("Y value is zero or negative, and its log is undefined!" _
                                & vbCrLf & "" _
                                & vbCrLf & "You should remove this data point before fitting" _
                                & vbCrLf & "(It is on line:" & Str(numRows%) & " )" _
                                , vbInformation, "log of zero")
                    End If
             Case "exp"
                 dPlot(numfil%, 1, numRows% - 1) = Exp(dPlot(numfil%, 1, numRows% - 1))
             Case "cos"
                 dPlot(numfil%, 1, numRows% - 1) = Cos(dPlot(numfil%, 1, numRows% - 1))
             Case "sin"
                 dPlot(numfil%, 1, numRows% - 1) = Sin(dPlot(numfil%, 1, numRows% - 1))
             Case "tan"
                 dPlot(numfil%, 1, numRows% - 1) = Tan(dPlot(numfil%, 1, numRows% - 1))
             Case Else
             
          End Select
          
          End If
       End If
       
Return

ReSort: 'resort the plotting array so that the x values are sorted from smallest to largest values
             
        Screen.MousePointer = vbHourglass
                          
        frmSetCond.ListSort.Clear
        For isort = 0 To numRows% - 1
           frmSetCond.ListSort.AddItem dPlot(numfil%, 0, isort) & "," & dPlot(numfil%, 1, isort)
        Next isort
        
        ReDim Preserve list(1, frmSetCond.ListSort.ListCount - 1)

        'now add the almost sorted list to an array, and use a bubble sort routine to finish the job
        For I = 1 To frmSetCond.ListSort.ListCount
           ListItems = Split(frmSetCond.ListSort.list(I - 1), ",")
           list(0, I - 1) = Val(ListItems(0))
           list(1, I - 1) = Val(ListItems(1))
        Next I
        
        'now sort the file and output to the sorted file
        Call BubbleSort(list, 0, frmSetCond.ListSort.ListCount - 1)
        
        'now update the plot buffer with the sorted array
        For isort = 0 To numRows% - 1
           dPlot(numfil%, 0, isort) = list(0, isort)
           dPlot(numfil%, 1, isort) = list(1, isort)
        Next isort
        
        Screen.MousePointer = vbDefault
        
        SoundWarning = False
             
Return

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

If drm% <> 0 Then
   frmDraw.DrawMode = drm%
   frmDraw.DrawStyle = drs%
   frmDraw.DrawWidth = drw%
   End If

If numFilesToPlot% <= 0 Then
   Screen.MousePointer = vbDefault
   MsgBox "Sorry, you need to rerun the plot wizard!", vbExclamation + vbOKOnly, "Plot"
   Exit Sub
   End If

Call frmSetCond.DefineLayout
Plot frmDraw, dPlot, udtMyGraphLayout

Screen.MousePointer = vbDefault

End Sub
Public Function BreakDown(ByVal Full$, Optional ByRef PName$ _
    , Optional ByRef FName$, Optional ByRef Ext$) As Boolean
'
'   Décompose un nom de fichier en différente sous partie:
'   Full$  = Nom Complet du fichier.
'   PName$ = Chemin du fichier.
'   FName$ = Nom du fichier avec son extension.
'   Ext$   = .extension du fichier.
'
'   Si le fichier n'existe pas retourne une valeur False.
'
    Dim Sloc&, Dot&
'
    BreakDown = Len(Dir$(Full$))
'
    If InStr(Full$, "\") Then
        FName$ = Full$
        PName$ = ""
        Sloc = InStr(FName$, "\")
        Do While Sloc <> 0
            PName$ = PName$ & Left$(FName$, Sloc)
            FName$ = Mid$(FName$, Sloc + 1)
            Sloc = InStr(FName$, "\")
        Loop
    Else
        PName$ = ""
        FName$ = Full$
    End If
'
    Dot = InStr(Full$, ".")
    If Dot <> 0 Then
        Ext$ = Mid$(Full$, Dot)
    Else
        Ext$ = ""
    End If
'
'
'
End Function

