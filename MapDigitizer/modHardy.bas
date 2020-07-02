Attribute VB_Name = "modHardy"
Public Matrix_A() As Double
Public Operations_Matrix() As Double 'Matrix where the calculations are done
Public Inverse_Matrix() As Double  'Matrix with the Inverse of [A]
Public GaussMethod As Boolean  'Matrix inversion method:  true to use Gaussian Elimination, false to use Cramer's rule (JKH's original code's method)
Public System_DIM As Long
Public C_vector() As Double
Public XminC As Double, YminC As Double, XmaxC As Double, YmaxC As Double

Public Declare Function SolveEquationsC Lib "MapDigitizer.dll" (first_X As Double, first_Y As Double, first_Z As Double, _
                                                     firstC_vector As Double, _
                                                     first_Xcoord As Double, first_Ycoord As Double, _
                                                     first_ht As Double, first_htf As Single, first_hts As Integer, Precision As Integer, _
                                                     xmin As Double, StepX As Double, ymin As Double, StepY As Double, _
                                                     np As Long, ByVal pFunc As Long) As Long 'this dll needs to be installed in system32
'Public Declare Function SolveEquationsC Lib "MapDigitizer.dll" (first_X As Double, first_Y As Double, first_Z As Double, _
'                                                     first_arr As Double, firstC_vector As Double, _
'                                                     first_Xcoord As Double, first_Ycoord As Double, first_ht As Double, _
'                                                     xmin As Double, StepX As Double, ymin As Double, StepY As Double, _
'                                                     np As Long, ByVal pFunc As Long) As Long 'this dll needs to be installed in system32
                                                     
'used for creating DTMs
'Public Declare Function SolveEquationsDTM Lib "MapDigitizer.dll" (np1 As Long, np2 As Long, np3 As Long, np4 As Long, np5 As Long, np6 As Long, np7 As Long)

Public Declare Function SolveEquationsDTM Lib "MapDigitizer.dll" (np As Long, num_rows As Long, num_cols As Long, _
                                                                  first_X As Double, first_Y As Double, firstC_vector As Double, _
                                                                  first_Xcoord As Double, first_Ycoord As Double, _
                                                                  first_Zcoord As Double, first_Zcoordf As Single, first_Zcoors As Integer, Precision As Integer, _
                                                                  ByVal pFunc As Long) As Long
                                                                  
'version without any file writings
'Public Declare Function Profiles Lib "MapDigitizer.dll" (xo As Double, yo As Double, zo As Double, _
'                                                         first_xx As Double, first_yy As Double, first_zz As Double, _
'                                                         num_array&, CoordMode%, HorizMode%, _
'                                                         first_va As Double, first_azi As Double, _
'                                                         first_Xazi As Double, first_Yazi As Double, first_Zazi As Double, first_distazi As Double, _
'                                                         numazi&, StepSizeAzi As Double, HalfAziRange As Double, _
'                                                         ByVal pFunc As Long) As Long
                                                         
Public Declare Function Profiles Lib "MapDigitizer.dll" (ByVal File_In As String, ByVal File_Out As String, _
                                                         xo As Double, yo As Double, zo As Double, _
                                                         nrows As Long, ncols As Long, CoordMode%, HorizMode%, Save_xyz%, _
                                                         first_va As Double, first_azi As Double, _
                                                         first_Xazi As Double, first_Yazi As Double, first_Zazi As Double, first_distazi As Double, _
                                                         numazi&, StepSizeAzi As Double, HalfAziRange As Double, Apprn As Double, _
                                                         ByVal pFunc As Long) As Long
                                                         
Public Declare Function Profiles2 Lib "MapDigitizer.dll" (ByVal File_In As String, ByVal File_Out As String, _
                                                         xo As Double, yo As Double, zo As Double, _
                                                         numpoints As Long, CoordMode%, HorizMode%, Save_xyz%, _
                                                         first_va As Double, first_azi As Double, _
                                                         first_Xazi As Double, first_Yazi As Double, first_Zazi As Double, first_distazi As Double, _
                                                         numazi&, StepSizeAzi As Double, HalfAziRange As Double, Apprn As Double, _
                                                         ByVal pFunc As Long) As Long
                                                         
                                                     
'functions used for raising the priority of the thread
'Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const NORMAL_PRIORITY_CLASS = &H20
Declare Function GetCurrentThread Lib "kernel32" () As Long
Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_PRIORITY_NORMAL = 0

Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
' Return the smallest parameter value.
Private Function min(ParamArray values() As Variant) As Variant
Dim i As Long
Dim minvalue As Variant

    minvalue = values(LBound(values))
    For i = LBound(values) + 1 To UBound(values)
        If minvalue > values(i) Then minvalue = values(i)
    Next i

    min = minvalue
End Function
' Return the largest parameter value.
Private Function Max(ParamArray values() As Variant) As Variant
Dim i As Long
Dim maxvalue As Variant

    maxvalue = values(LBound(values))
    For i = LBound(values) + 1 To UBound(values)
        If maxvalue < values(i) Then maxvalue = values(i)
    Next i

    Max = maxvalue
End Function
'---------------------------------------------------------------------------------------
' Procedure : FindPointsHardy
' Author    : Dr-John-K-Hall
' Date      : 2/16/2015
' Purpose   : Adds the digitized points to the Hardy analysis array
'           mode = 0 the draw over clean canvas
'                = 1 to draw over canvas with other Hardy points
'---------------------------------------------------------------------------------------
'
Public Function FindPointsHardy(RectCoord() As POINTAPI, mode As Integer, FileName As String) As Integer

   Dim i As Long, j As Long, K As Long
   Dim xPoint As Long, yPoint As Long, zPoint As Double
   Dim Slope As Double, ier As Integer
   Dim R As Long, r2 As Long, Test_Color As couleur
   Dim X1, Y1, X2, Y2
   Dim XGeo As Double, YGeo As Double
   Dim numSpaceLine As Integer
   Dim color_line As Long
   Dim Height_Color As couleur
   
   X1 = RectCoord(0).x
   Y1 = RectCoord(0).Y
   X2 = RectCoord(1).x
   Y2 = RectCoord(1).Y
   
   'create byte array 0 for
   
   TraceColor = ContourColor& 'QBColor(12)
   TracingColor = recupcouleur(TraceColor)
   numSpaceLine = numDistLines '5 'minimum distance between line vertices to save

   On Error GoTo FindPointsHardy_Error
   
   '------------------progress bar initialization
   With GDMDIform
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

   'check all the temporary buffers for points and then subtract those that were erased
   
   If mode = 0 Then numDigiHardyPoints = 0
   
'   '------diagnostics-------------------
'   'read in diagnostics file
'    filnum% = FreeFile
'    Open App.Path & "\topo-out.gsc" For Input As #filnum%
'    Do Until EOF(filnum%)
'       Input #filnum%, xPoint, yPoint, XGeo, YGeo, zPoint
'       If numDigiHardyPoints = 0 Then
'          ReDim DigiHardyPoints(numDigiHardyPoints)
'       Else
'          ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
'          End If
'       DigiHardyPoints(numDigiHardyPoints).X = xPoint
'       DigiHardyPoints(numDigiHardyPoints).Y = yPoint
'       DigiHardyPoints(numDigiHardyPoints).z = zPoint
'       numDigiHardyPoints = numDigiHardyPoints + 1
'    Loop
'    Close #filnum%
'    ier = 0
'    FindPointsHardy = ier
'    Exit Function
'    '--------------------------------------------------
   
   Call UpdateStatus(GDMDIform, 1, 0)
   GDMDIform.StatusBar1.Panels(1).Text = "Please wait, finding relevant digitized data points for Hardy quadratic suface analysis"
   
   For i = 0 To numDigiPoints - 1
   
      If DigiPoints(i).x >= X1 And DigiPoints(i).x <= X2 And _
         DigiPoints(i).Y >= Y1 And DigiPoints(i).Y <= Y2 Then
         
            If numDigiHardyPoints = 0 Then
               ReDim DigiHardyPoints(numDigiHardyPoints)
               DigiHardyPoints(numDigiHardyPoints).x = DigiPoints(i).x
               DigiHardyPoints(numDigiHardyPoints).Y = DigiPoints(i).Y
               DigiHardyPoints(numDigiHardyPoints).Z = DigiPoints(i).Z
               numDigiHardyPoints = numDigiHardyPoints + 1
            Else
               'check for duplicate points
               If DigiPoints(i).x <> DigiHardyPoints(numDigiHardyPoints - 1).x _
                  And DigiPoints(i).Y <> DigiHardyPoints(numDigiHardyPoints - 1).Y Then
                  ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
                  DigiHardyPoints(numDigiHardyPoints).x = DigiPoints(i).x
                  DigiHardyPoints(numDigiHardyPoints).Y = DigiPoints(i).Y
                  DigiHardyPoints(numDigiHardyPoints).Z = DigiPoints(i).Z
                  numDigiHardyPoints = numDigiHardyPoints + 1
                  End If
               End If
          End If
      
      Call UpdateStatus(GDMDIform, 1, CLng(100 * (i + 1) / numDigiPoints))
         
   Next i
   
   'now check the contour buffer
   Call UpdateStatus(GDMDIform, 1, 0)
   GDMDIform.StatusBar1.Panels(1).Text = "Please wait, finding relevant digitized contours for Hardy quadratic suface analysis"
         
   For i = 0 To numDigiContours - 1
   
      If DigiContours(i).x >= X1 And DigiContours(i).x <= X2 And _
         DigiContours(i).Y >= Y1 And DigiContours(i).Y <= Y2 Then
         
'         If InitDigiGraph Then 'contours drawn in TraceColor color
         
            'only accept those contour points that haven't been erased
            'first check for countours drawn in TracingColor
            R = GDform1.Picture2.Point(CLng(DigiContours(i).x * DigiZoom.LastZoom), CLng(DigiContours(i).Y * DigiZoom.LastZoom))
            If R <> -1 Then
               Test_Color = recupcouleur(R)
               If Test_Color.R = TracingColor.R And _
                  Test_Color.V = TracingColor.V And _
                  Test_Color.b = TracingColor.b Then
                  'this is on a contour, so accept it if it is not a repeat
                  If numDigiHardyPoints = 0 Then
                     ReDim DigiHardyPoints(numDigiHardyPoints)
                     DigiHardyPoints(numDigiHardyPoints).x = DigiContours(i).x
                     DigiHardyPoints(numDigiHardyPoints).Y = DigiContours(i).Y
                     DigiHardyPoints(numDigiHardyPoints).Z = DigiContours(i).Z
                     numDigiHardyPoints = numDigiHardyPoints + 1
                  Else
                     'don't record duplicate points
                     If DigiHardyPoints(numDigiHardyPoints - 1).x <> DigiContours(i).x _
                        And DigiHardyPoints(numDigiHardyPoints - 1).Y <> DigiContours(i).Y Then
                        ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
                        DigiHardyPoints(numDigiHardyPoints).x = DigiContours(i).x
                        DigiHardyPoints(numDigiHardyPoints).Y = DigiContours(i).Y
                        DigiHardyPoints(numDigiHardyPoints).Z = DigiContours(i).Z
                        numDigiHardyPoints = numDigiHardyPoints + 1
                        End If
                     End If
                     GoTo fp100
                  End If
               End If
            
'        ElseIf Not InitDigiGraph And LineElevColors& = 1 And numcpt > 0 Then 'contours drawn in rainbow colors dependent on height
            If LineElevColors& <> 1 Or numcpt = 0 Then GoTo fp100
            'check for color
            colornum% = ((DigiContours(i).Z - MinColorHeight) / (MaxColorHeight - MinColorHeight)) * UBound(cpt, 2) + 1
            color_line = RGB(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))
            
            R = GDform1.Picture2.Point(CLng(DigiContours(i).x * DigiZoom.LastZoom), CLng(DigiContours(i).Y * DigiZoom.LastZoom))
            If R <> -1 Then
               Test_Color = recupcouleur(R)
               Height_Color = recupcouleur(color_line)
               If Test_Color.R = Height_Color.R And _
                  Test_Color.V = Height_Color.V And _
                  Test_Color.b = Height_Color.b Then
                  'this is on a contour, so accept it if it is not a repeat
                  If numDigiHardyPoints = 0 Then
                     ReDim DigiHardyPoints(numDigiHardyPoints)
                     DigiHardyPoints(numDigiHardyPoints).x = DigiContours(i).x
                     DigiHardyPoints(numDigiHardyPoints).Y = DigiContours(i).Y
                     DigiHardyPoints(numDigiHardyPoints).Z = DigiContours(i).Z
                     numDigiHardyPoints = numDigiHardyPoints + 1
                  Else
                     'don't record duplicate points
                     If DigiHardyPoints(numDigiHardyPoints - 1).x <> DigiContours(i).x _
                        And DigiHardyPoints(numDigiHardyPoints - 1).Y <> DigiContours(i).Y Then
                        ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
                        DigiHardyPoints(numDigiHardyPoints).x = DigiContours(i).x
                        DigiHardyPoints(numDigiHardyPoints).Y = DigiContours(i).Y
                        DigiHardyPoints(numDigiHardyPoints).Z = DigiContours(i).Z
                        numDigiHardyPoints = numDigiHardyPoints + 1
                        End If
                     End If
                  End If
               End If
             End If
               
   
'        End If
fp100:
      Call UpdateStatus(GDMDIform, 1, CLng(100 * (i + 1) / numDigiContours))
         
   Next i
   
   'now check the line buffer
   Call UpdateStatus(GDMDIform, 1, 0)
   GDMDIform.StatusBar1.Panels(1).Text = "Please wait, finding relevant digitized lines for Hardy quadratic suface analysis"
   
   For i = 0 To numDigiLines - 1
   
      If DigiLines(0, i).x >= X1 And DigiLines(0, i).x <= X2 And _
         DigiLines(0, i).Y >= Y1 And DigiLines(0, i).Y <= Y2 And _
         DigiLines(1, i).x >= X1 And DigiLines(1, i).x <= X2 And _
         DigiLines(1, i).Y >= Y1 And DigiLines(1, i).Y <= Y2 Then
         
        'convert lines into points along
        If DigiLines(1, i).x <> DigiLines(0, i).x Then
        
           Slope = (DigiLines(1, i).Y - DigiLines(0, i).Y) / (DigiLines(1, i).x - DigiLines(0, i).x)
           
           For xPoint = DigiLines(0, i).x To DigiLines(1, i).x Step (DigiLines(1, i).x - DigiLines(0, i).x) / Abs((DigiLines(1, i).x - DigiLines(0, i).x)) * numSpaceLine
           
               yPoint = Slope * (xPoint - DigiLines(0, i).x) + DigiLines(0, i).Y
               
                If numDigiHardyPoints = 0 Then
                   ReDim DigiHardyPoints(numDigiHardyPoints)
                   DigiHardyPoints(numDigiHardyPoints).x = xPoint
                   DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                   DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                   numDigiHardyPoints = numDigiHardyPoints + 1
                Else
                  If DigiHardyPoints(numDigiHardyPoints - 1).x <> xPoint _
                     And DigiHardyPoints(numDigiHardyPoints - 1).Y <> yPoint Then
                        'not a duplicate, so record it
                        ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
                        DigiHardyPoints(numDigiHardyPoints).x = xPoint
                        DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                        DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                        numDigiHardyPoints = numDigiHardyPoints + 1
                        End If
                   End If
              
           Next xPoint
           If xPoint <> DigiLines(1, i).x Then 'end vertex wasn't recorded, so record it
                xPoint = DigiLines(1, i).x
                yPoint = DigiLines(1, i).Y
                If numDigiHardyPoints = 0 Then
                   DigiHardyPoints(numDigiHardyPoints).x = xPoint
                   DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                   DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                   numDigiHardyPoints = numDigiHardyPoints + 1
                   ReDim DigiHardyPoints(numDigiHardyPoints)
                Else
                  If DigiHardyPoints(numDigiHardyPoints - 1).x <> xPoint _
                     And DigiHardyPoints(numDigiHardyPoints - 1).Y <> yPoint Then
                        'not a duplicate, so record it
                        ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
                        DigiHardyPoints(numDigiHardyPoints).x = xPoint
                        DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                        DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                        numDigiHardyPoints = numDigiHardyPoints + 1
                        End If
                   End If
              End If
       Else 'vertical lines
           If DigiLines(1, i).Y <> DigiLines(0, i).Y Then
                For yPoint = DigiLines(0, i).Y To DigiLines(1, i).Y Step (DigiLines(1, i).Y - DigiLines(0, i).Y) / Abs(DigiLines(1, i).Y - DigiLines(0, i).Y) * numSpaceLine
                     xPoint = DigiLines(1, i).x
                     If numDigiHardyPoints = 0 Then
                        ReDim DigiHardyPoints(numDigiHardyPoints)
                        DigiHardyPoints(numDigiHardyPoints).x = xPoint
                        DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                        DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                        numDigiHardyPoints = numDigiHardyPoints + 1
                     Else
                       If DigiHardyPoints(numDigiHardyPoints - 1).x <> xPoint _
                          And DigiHardyPoints(numDigiHardyPoints - 1).Y <> yPoint Then
                             'not a duplicate, so record it
                             ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
                             DigiHardyPoints(numDigiHardyPoints).x = xPoint
                             DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                             DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                             numDigiHardyPoints = numDigiHardyPoints + 1
                             End If
                        End If
                Next yPoint
                End If
           End If
           If yPoint <> DigiLines(1, i).Y Then 'end vertex was recorded, so record it
                xPoint = DigiLines(1, i).x
                yPoint = DigiLines(1, i).Y
                If numDigiHardyPoints = 0 Then
                   ReDim DigiHardyPoints(numDigiHardyPoints)
                   DigiHardyPoints(numDigiHardyPoints).x = xPoint
                   DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                   DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                   numDigiHardyPoints = numDigiHardyPoints + 1
                Else
                  If DigiHardyPoints(numDigiHardyPoints - 1).x <> xPoint _
                     And DigiHardyPoints(numDigiHardyPoints - 1).Y <> yPoint Then
                        'not a duplicate, so record it
                        ReDim Preserve DigiHardyPoints(numDigiHardyPoints)
                        DigiHardyPoints(numDigiHardyPoints).x = xPoint
                        DigiHardyPoints(numDigiHardyPoints).Y = yPoint
                        DigiHardyPoints(numDigiHardyPoints).Z = DigiLines(0, i).Z
                        numDigiHardyPoints = numDigiHardyPoints + 1
                        End If
                   End If
              End If
           
       End If
      
      Call UpdateStatus(GDMDIform, 1, CLng(100 * (i + 1) / numDigiLines))
         
   Next i
   
   Call UpdateStatus(GDMDIform, 1, 0)
   
   If FileName = sEmpty Then
   
      'find root name of map
      Dim RootName() As String
      RootName = Split(picnam$, "\")
      ln% = UBound(RootName)
      RootFileName$ = RootName(ln%)
      pos% = InStr(RootFileName$, ".")
      RootFileName$ = Mid$(RootFileName$, 1, pos% - 1) & ".dxf"
      
   Else
   
      RootFileName$ = FileName
      
      End If
   
   GDMDIform.StatusBar1.Panels(1).Text = "Finished finding the relevant points, finding bounds and writing temporay file ''topo-out.GSC''."
   
   If numDigiHardyPoints > 0 Then
      GDMDIform.Toolbar1.Buttons(30).Enabled = True 'enable save button
   Else
      GDMDIform.Toolbar1.Buttons(30).Enabled = False
      ier = -1
      FindPointsHardy = ier
      GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
      GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
      GDMDIform.picProgBar.Visible = False
      Exit Function
      End If
   
   'now write temporary file containing the data, and convert pixels into coordinates
   Call UpdateStatus(GDMDIform, 1, 0)
   filnum% = FreeFile
   
   Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
   xmin = INIT_VALUE
   xmax = -INIT_VALUE
   ymin = INIT_VALUE
   ymax = -INIT_VALUE
   zmin = INIT_VALUE
   zmax = -INIT_VALUE
   
   Open App.Path & "\topo-out.GSC" For Output As #filnum%
   For i = 0 To numDigiHardyPoints - 1
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(DigiHardyPoints(i).x), CDbl(DigiHardyPoints(i).Y), XGeo, YGeo)
        ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(DigiHardyPoints(i).x), CDbl(DigiHardyPoints(i).Y), XGeo, YGeo)
        ElseIf RSMethod0 Then
           ier = Simple_pixel_to_coord(CDbl(DigiHardyPoints(i).x), CDbl(DigiHardyPoints(i).Y), XGeo, YGeo)
           End If
        Write #filnum%, DigiHardyPoints(i).x, DigiHardyPoints(i).Y, XGeo, YGeo, DigiHardyPoints(i).Z
        
        xmin = min(xmin, XGeo)
        xmax = Max(xmax, XGeo)
        ymin = min(ymin, YGeo)
        ymax = Max(ymax, YGeo)
        zmin = min(zmin, DigiHardyPoints(i).Z)
        zmax = Max(zmax, DigiHardyPoints(i).Z)
        
        Call UpdateStatus(GDMDIform, 1, (i + 1) * 100 / numDigiHardyPoints)
   Next i
   Close #filnum%
   
   GDMDIform.StatusBar1.Panels(1).Text = "Writing dxf file: ''" & RootFileName$ & "."
   Call UpdateStatus(GDMDIform, 1, 0)
   GDMDIform.picProgBar.Visible = False
   filnum% = FreeFile
   Open App.Path & "\" & RootFileName$ For Output As #filnum%
   Print #filnum%, "  0"
   Print #filnum%, "SECTION"
   Print #filnum%, "  2"
   Print #filnum%, "HEADER"
   Print #filnum%, "  9"
   Print #filnum%, "$ACADVER"
   Print #filnum%, "  1"
   Print #filnum%, "AC1009"
   Print #filnum%, "  9"
   Print #filnum%, "$LUNITS"
   Print #filnum%, " 70"
   Print #filnum%, "2"
   Print #filnum%, "  9"
   Print #filnum%, "$LIMMIN"
   Print #filnum%, "  10"
   Print #filnum%, Trim$(str$(xmin))
   Print #filnum%, "  20"
   Print #filnum%, Trim$(str$(ymin))
   Print #filnum%, "  9"
   Print #filnum%, "$LIMMAX"
   Print #filnum%, "  10"
   Print #filnum%, Trim$(str$(xmax))
   Print #filnum%, "  20"
   Print #filnum%, Trim$(str$(ymax))
   Print #filnum%, "  9"
   Print #filnum%, "$EXTMIN"
   Print #filnum%, "  10"
   Print #filnum%, Trim$(str$(xmin))
   Print #filnum%, "  20"
   Print #filnum%, Trim$(str$(ymin))
   Print #filnum%, "  30"
   Print #filnum%, Trim$(str$(zmin))
   Print #filnum%, "  9"
   Print #filnum%, "$EXTMAX"
   Print #filnum%, "  10"
   Print #filnum%, Trim$(str$(xmax))
   Print #filnum%, "  20"
   Print #filnum%, Trim$(str$(ymax))
   Print #filnum%, "  30"
   Print #filnum%, Trim$(str$(zmax))
   Print #filnum%, "  999"
   Print #filnum%, "Created by MapDigitizer"
   Print #filnum%, "  999"
   Print #filnum%, "Projection: " & lblX & ", " & LblY
   Print #filnum%, "  999"
   Print #filnum%, "Datum: "
   Print #filnum%, "  999"
   Print #filnum%, "Ground Units: meters"
   Print #filnum%, "  999"
   Print #filnum%, "ZONE: "
   Print #filnum%, "  0"
   Print #filnum%, "ENDSEC"
   Print #filnum%, "  0"
   Print #filnum%, "SECTION"
   Print #filnum%, "  2"
   Print #filnum%, "ENTITIES"
   
   filgsc% = FreeFile
   Open App.Path & "\topo-out.GSC" For Input As #filgsc%
   i = 0
   Do Until EOF(filgsc%)
      Input #filgsc%, XX1, YY1, XGeo, YGeo, ZGeo
      
      Print #filnum%, "  0"
      Print #filnum%, "POINT"
      Print #filnum%, "  8"
      Print #filnum%, "3D_PNT"
      Print #filnum%, "  10"
      Print #filnum%, Trim$(str$(XGeo))
      Print #filnum%, "  20"
      Print #filnum%, Trim$(str$(YGeo))
      Print #filnum%, "  30"
      Print #filnum%, Trim$(str$(ZGeo))
      
      Call UpdateStatus(GDMDIform, 1, (i + 1) * 100 / numDigiHardyPoints)
      i = i + 1
   Loop
   
   Print #filnum%, "  0"
   Print #filnum%, "ENDSEC"
   Print #filnum%, "  0"
   Print #filnum%, "EOF"
   
   Close #filgsc%
   Close #filnum%
   
   GDMDIform.picProgBar.Visible = False
   GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
   
   ier = 0
   FindPointsHardy = ier

   On Error GoTo 0
   Exit Function

FindPointsHardy_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FindPointsHardy of Module modHardy"
'    GDMDIform.picProgBar.Visible = False
'    GDMDIform.StatusBar1.Panels(2).Text = gsEmpty

'    ier = -1
'    FindPointsHardy = ier
    Resume Next
End Function
'---------------------------------------------------------------------------------------
' Procedure : HardyQuadraticSurfaces
' Author    : Dr-John-K-Hall
' Date      : 2/16/2015
' Purpose   : Calculates the Hardy Quadratic Surfaces
'---------------------------------------------------------------------------------------
'
Public Function HardyQuadraticSurfaces(Pic As PictureBox, RectCoord() As POINTAPI) As Integer

    Dim a() As Double
    Dim x() As Double, Y() As Double, Z() As Double
    Dim Xcoord() As Double, Ycoord() As Double
    Dim ht() As Double, ht2() As Double
    Dim htf() As Single, ht2f() As Single
    Dim hts() As Integer, ht2s() As Integer
    Dim np As Long, ndep As Long
    Dim xmin As Double, xmax As Double
    Dim ymin As Double, ymax As Double
    Dim zmin As Double, zmax As Double
    Dim i As Long, j As Long, K As Long
    Dim ixp As Long, iyp As Long
    Dim StepX As Double, StepY As Double
    Dim nc As Integer
    Dim contour() As Double
    Dim X1, Y1, X2, Y2
    Dim XGeo As Double, YGeo As Double
    Dim EscapeKeyPressed As Boolean
    Dim JKHmethod As Boolean
    Dim DLLMethod As Boolean
    Dim tryingtokill As Boolean

    Dim Reduce_Dim As Long 'reduce the resolution by this factor in order to speed up calculation of the output data and contours
    Reduce_Dim = 1
    
    Dim ier As Integer
    
    HardyCoordinateOutput = True 'output coordinates
    JKHmethod = False 'don't use slow method of JKH called lineqem
    DLLMethod = True
    
    ier = 0

    Dim ContourInterval As Integer
    ContourInterval = val(GDMDIform.combContour.Text) '2 '10 '5 '10 '100 'contour intervals in height units
    
    zmin = INIT_VALUE
    zmax = -INIT_VALUE
    xmin = INIT_VALUE 'new
    xmax = -INIT_VALUE 'new
    ymin = INIT_VALUE 'new
    ymax = -INIT_VALUE 'new
   
    X1 = RectCoord(0).x
    Y1 = RectCoord(0).Y
    X2 = RectCoord(1).x
    Y2 = RectCoord(1).Y
    
    XminC = X1
    YminC = Y1
    XmaxC = X2
    YmaxC = Y2
    
   '------------------progress bar initialization
   With GDMDIform
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

   On Error GoTo HardyQuadraticSurfaces_Error
   
    If Not DLLMethod Then
    
        ReDim x(1 To numDigiHardyPoints)
        ReDim Y(1 To numDigiHardyPoints)
        ReDim Z(1 To numDigiHardyPoints)
        ReDim Xcoord(0 To numDigiHardyPoints - 1)
        ReDim Ycoord(0 To numDigiHardyPoints - 1)
        If HeightPrecision = 0 Then
            ReDim hts(0 To numDigiHardyPoints - 1, 0 To numDigiHardyPoints - 1)
            ReDim ht(0)
            ReDim htf(0)
        ElseIf HeightPrecision = 1 Then
            ReDim htf(0 To numDigiHardyPoints - 1, 0 To numDigiHardyPoints - 1)
            ReDim ht(0)
            ReDim hts(0)
        ElseIf HeightPrecision = 2 Then
            ReDim ht(0 To numDigiHardyPoints - 1, 0 To numDigiHardyPoints - 1)
            ReDim hts(0)
            ReDim htf(0)
            End If
    
        GDMDIform.StatusBar1.Panels(1).Text = "Stuffing data into arrays, please wait...."
        Call UpdateStatus(GDMDIform, 1, 0)
        
        For i = 1 To numDigiHardyPoints
           x(i) = DigiHardyPoints(i - 1).x
           If x(i) < xmin Then xmin = x(i)
           If x(i) > xmax Then xmax = x(i)
           Y(i) = DigiHardyPoints(i - 1).Y
           If Y(i) < ymin Then ymin = Y(i)
           If Y(i) > ymax Then ymax = Y(i)
           Z(i) = DigiHardyPoints(i - 1).Z
           If Z(i) < zmin Then zmin = Z(i)
           If Z(i) > zmax Then zmax = Z(i)
           Call UpdateStatus(GDMDIform, 1, 100 * (i + 1) / numDigiHardyPoints)
        Next i
    
        np = numDigiHardyPoints
        
        End If
   
    If Not GaussMethod And JKHmethod Then 'use JKH's "lineqem" (some sort of gaussian inversion method to solve simultaneous equations, but works slowly)

        ReDim a(1 To numDigiHardyPoints, 1 To numDigiHardyPoints + 1)

        '-----------------topo------------------------
        Call UpdateStatus(GDMDIform, 1, 0)
        GDMDIform.StatusBar1.Panels(1).Text = "Beginning ''topo'' stage of calculating Hardy quadratic surfaces."
L5:
       
        For i = 1 To np
           For j = 1 To i
               If (i <> j) Then GoTo L10
               a(i, j) = 0#
               GoTo L15
L10:
               dX = x(j) - x(i)
               dy = Y(j) - Y(i)
               If (dX = 0# And dy = 0#) Then GoTo L25
               a(i, j) = Sqr(dX ^ 2 + dy ^ 2)
               a(j, i) = a(i, j)
L15:
           Next j
L20:
           a(i, np + 1) = Z(i)
           DoEvents
        Next i
        GoTo L40
    '     two points are the same, remove one and repeat the process
L25:
        If (i = np) Then GoTo L35
        npm1 = np - 1
        For K = i To npm1
            l = K + 1
            x(K) = x(l)
            Y(K) = Y(l)
L30:
            Z(K) = Z(l)
            DoEvents
        Next K
L35:
        np = np - 1
        DoEvents
        If npm1 > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * (numDigiHardyPoints - np + 2) / npm1)
        GoTo L5
L40:
       '----------------lineqem---------------------
   
       GDMDIform.StatusBar1.Panels(1).Text = "Matrix inversion (using Cramer's rule) stage...press ""ESC"" to abort."
       Call UpdateStatus(GDMDIform, 1, 0)
       
        Dim ts As Double
        Dim tm As Double
    
        izdt = 1
        m = np + 1
        nm = np - 1
        For i = 1 To np
            l = i + 1
            If (a(i, i) <> 0#) Then GoTo L200
            For j = 1 To np
                If (a(j, i) = 0#) Then GoTo L100
                If (a(i, j) = 0#) Then GoTo L100
                For K = i To m
                    ts = a(i, K)
                    a(i, K) = a(j, K)
                    a(j, K) = ts
l50:
                Next K
                GoTo L200
L100:
            Next j
            
            izdt = 0
            'if got here, detected zero diagonal term, so can't invert matrices
            ier = -1
            HardyQuadraticSurfaces = ier
            GDMDIform.StatusBar1.Panels(1).Text = sEmpty
            GDMDIform.picProgBar.Visible = False
            GDMDIform.StatusBar1.Panels(2).Text = sEmpty
            Exit Function
L200:
            td = 1# / a(i, i)
            For j = 1 To m
                a(i, j) = a(i, j) * td
L250:
            Next j
            If (i = np) Then GoTo L400
            For j = l To np
                tm = a(j, i)
                a(j, i) = 0#
                For K = l To m
                    a(j, K) = a(j, K) - tm * a(i, K)
L300:
                Next K
L350:
            Next j
L400:
            DoEvents
            If np > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * (i + 1) / np)
        Next i
        
        Call UpdateStatus(GDMDIform, 1, 0)
        For i = 1 To nm
            j = np - i
            l = j + 1
            For K = 1 To j
                a(K, m) = a(K, m) - a(K, l) * a(l, m)
                a(K, l) = 0#
L450:
            Next K
L500:
            Call UpdateStatus(GDMDIform, 1, 100 * i / nm)
            
            DoEvents
            
            '---------------------break on ESC key-------------------------------
            If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
               EscapeKeyPressed = True
               GoTo L5000
               End If
            '---------------------------------------------------------------------
            
        Next i
    
        If Save_xyz% = 1 Then
            '-------------------generate temporary xyz file---------------------
             tryingtokill = True
             
             If Dir(App.Path & "\topo_coord.xyz") <> gsEmpty Then Kill App.Path & "\topo_coord.xyz"
             If Dir(App.Path & "\topo_pixel.xyz") <> gsEmpty Then Kill App.Path & "\topo_pixel.xyz"
             
             tryingtokill = False
            
             GDMDIform.StatusBar1.Panels(1).Text = "Packing height array and writing xyz data, please wait....(press ''ESC'' key to abort)"
             filnum% = FreeFile
             Open App.Path & "\topo_coord.xyz" For Output As #filnum%
             filtopo% = FreeFile
             Open App.Path & "\topo_pixel.xyz" For Output As #filtopo%
             
        Else
        
             GDMDIform.StatusBar1.Panels(1).Text = "Packing height array, please wait....(press ''ESC'' key to abort)"
             End If
        
        xmax = X2
        xmin = X1
        ymax = Y2
        ymin = Y1
            
        StepX = (xmax - xmin) / (np - 1)
        StepY = (ymax - ymin) / (np - 1)
        
        Call UpdateStatus(GDMDIform, 1, 0)
        For ixp = 1 To np
            xi = xmin + (ixp - 1) * StepX
            Xcoord(ixp - 1) = xi
            For iyp = 1 To np
                yi = ymin + (iyp - 1) * StepY
                m = np + 1
                zi = 0#
                For j = 1 To np
                    dX = xi - x(j)
                    dy = yi - Y(j)
L2500:
                    zi = zi + a(j, m) * Sqr(dX * dX + dy * dy)
                    DoEvents
                Next j
                
                If HeightPrecision = 0 Then
                    hts(ixp - 1, iyp - 1) = zi '* 10#
                ElseIf HeightPrecision = 1 Then
                    htf(ixp - 1, iyp - 1) = zi '* 10#
                ElseIf HeightPrecision = 2 Then
                    ht(ixp - 1, iyp - 1) = zi '* 10#
                    End If
                    
                Ycoord(iyp - 1) = yi
                
                If Save_xyz% Then
                    If HardyCoordinateOutput Then
                        'convert pixel coordinates to geographic coordinates  'use this to output final values
                        If RSMethod1 Then
                           ier = RS_pixel_to_coord2(CDbl(xi), CDbl(yi), XGeo, YGeo)
                        ElseIf RSMethod2 Then
                           ier = RS_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                        ElseIf RSMethod0 Then
                           ier = Simple_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                           End If
                    Else
                       'output pixel coordinates
                       XGeo = xi
                       YGeo = yi
                       End If
                       
                    If HardyCoordinateOutput Then
                        If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                           XGeo = val(Format(str$(XGeo), "#####0.0##")) 'CLng(XGeo)
                           YGeo = val(Format(str$(YGeo), "######0.0##")) 'CLng(YGeo)
                        ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                           XGeo = val(Format(str$(XGeo), "####0.0######"))
                           YGeo = val(Format(str$(YGeo), "####0.0######"))
                        Else
                           XGeo = CLng(XGeo)
                           YGeo = CLng(YGeo)
                           End If
                    Else
                       XGeo = CLng(XGeo)
                       YGeo = CLng(YGeo)
                       End If
                       
                    Write #filnum%, XGeo, YGeo, zi '* 10#
                    Write #filtopo%, Nint(Xcoord(ixp - 1)), Nint(Ycoord(iyp - 1)), zi '* 10#
                    End If
                
            Next iyp
L3500:

            If numDigiHardyPoints > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * ixp / np) 'numDigiHardyPoints)
            DoEvents
            
            '---------------------break on ESC key-------------------------------
            If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
               EscapeKeyPressed = True
               GoTo L4000
               End If
            '---------------------------------------------------------------------

        Next ixp
L4000:
        Close #filnum%
        Close #filtopo%
        
        g_nrows = np
        g_ncols = np
        
L5000:
        'reclaim memory
        ReDim a(1 To 1, 1 To 1)
        
   ElseIf Not GaussMethod And Not JKHmethod Then 'use Gaussian Elimination method to calculate matrix inverse (faster than JKH's but still slow)
   
       GDMDIform.StatusBar1.Panels(1).Text = "Beginning matrix inversion (using Gaussian elimination method) stage of calculating Hardy quadratic surfaces."
       Call UpdateStatus(GDMDIform, 1, 0)
       
       'redimension arrays
       ReDim Matrix_A(1 To np, 1 To np)
       ReDim Operations_Matrix(1 To np, 1 To 2 * np)  'Matrix where the calculations are done
       ReDim Inverse_Matrix(1 To np, 1 To np)  'Matrix with the Inverse of [A]
       ReDim C_vector(1 To np)
       
        '-----------------pack the array (replaces "topo")------------------------
        Call UpdateStatus(GDMDIform, 1, 0)
        GDMDIform.StatusBar1.Panels(1).Text = "Beginning calculating Hardy quadratic surfaces."
LL5:
           
        For i = 1 To np
           For j = 1 To i
               If (i <> j) Then GoTo LL10
               Matrix_A(i, j) = 0#
               GoTo LL15
LL10:
               dX = x(j) - x(i)
               dy = Y(j) - Y(i)
               If (dX = 0# And dy = 0#) Then GoTo LL25
               Matrix_A(i, j) = Sqr(dX ^ 2 + dy ^ 2)
               Matrix_A(j, i) = Matrix_A(i, j)
LL15:
           Next j
LL20:
'           DoEvents
        Next i
        GoTo LL40
    '     two points are the same, remove one and repeat the process
LL25:
        If (i = np) Then GoTo LL35
        npm1 = np - 1
        For K = i To npm1
            l = K + 1
            x(K) = x(l)
            Y(K) = Y(l)
LL30:
            Z(K) = Z(l)
'            DoEvents
        Next K
LL35:
        np = np - 1
        DoEvents
        If npm1 > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * (numDigiHardyPoints - np + 2) / npm1)
        GoTo LL5
LL40:

       System_DIM = np

       Call UpdateStatus(GDMDIform, 1, 0)
       GDMDIform.StatusBar1.Panels(1).Text = "Solving simultaneous equations using Gaussian elimination...press ""ESC"" to abort..."

       ier = Calculate_Inverse(EscapeKeyPressed)
       If ier < 0 Then GoTo LLL50
       
       'multiply inverse by z to get the coeficients, c
       
       GDMDIform.StatusBar1.Panels(1).Text = "Calculating Hardy quadratic surfaces coeficients..."
       GDMDIform.picProgBar.Visible = True
       Call UpdateStatus(GDMDIform, 1, 0)
       
       For i = 1 To np
          C_vector(i) = 0#
          For j = 1 To np
             C_vector(i) = C_vector(i) + Inverse_Matrix(i, j) * Z(j)
          Next j
          DoEvents
          Call UpdateStatus(GDMDIform, 1, 100 * i / np)
       Next i
      
       If Save_xyz% = 1 Then
             '-------------------generate temporary xyz file---------------------
             tryingtokill = True
             
             If Dir(App.Path & "\topo_coord.xyz") <> gsEmpty Then Kill App.Path & "\topo_coord.xyz"
             If Dir(App.Path & "\topo_pixel.xyz") <> gsEmpty Then Kill App.Path & "\topo_pixel.xyz"
             
             tryingtokill = False
            
             Call UpdateStatus(GDMDIform, 1, 0)
             GDMDIform.StatusBar1.Panels(1).Text = "Packing height array and writing xyz data, please wait....(press ''ESC'' key to abort)"
             
             filnum% = FreeFile
             Open App.Path & "\topo_coord.xyz" For Output As #filnum%
             filtopo% = FreeFile
             Open App.Path & "\topo_pixel.xyz" For Output As #filtopo%
             
        Else
             GDMDIform.StatusBar1.Panels(1).Text = "Packing height array, please wait....(press ''ESC'' key to abort)"
             End If
        
        xmax = X2
        xmin = X1
        ymax = Y2
        ymin = Y1
            
        StepX = (xmax - xmin) / (np - 1)
        StepY = (ymax - ymin) / (np - 1)
        
        For ixp = 1 To np
            xi = xmin + (ixp - 1) * StepX
            Xcoord(ixp - 1) = xi
            For iyp = 1 To np
                yi = ymin + (iyp - 1) * StepY
                m = np + 1
                zi = 0#
                For j = 1 To np
                    dX = xi - x(j)
                    dy = yi - Y(j)
LLL25:
                    zi = zi + C_vector(j) * Sqr(dX * dX + dy * dy)
'                    DoEvents
                Next j
                
                If HeightPrecision = 0 Then
                    hts(ixp - 1, iyp - 1) = zi
                ElseIf HeightPrecision = 1 Then
                    htf(ixp - 1, iyp - 1) = zi
                ElseIf HeightPrecision = 2 Then
                    ht(ixp - 1, iyp - 1) = zi
                    End If
                    
                Ycoord(iyp - 1) = yi
                
                If Save_xyz% = 1 Then
                    If HardyCoordinateOutput Then
                        'convert pixel coordinates to geographic coordinates  'use this to output final values
                        If RSMethod1 Then
                           ier = RS_pixel_to_coord2(CDbl(xi), CDbl(yi), XGeo, YGeo)
                        ElseIf RSMethod2 Then
                           ier = RS_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                        ElseIf RSMethod0 Then
                           ier = Simple_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                           End If
                    Else
                       'output pixel coordinates
                       XGeo = xi
                       YGeo = yi
                       End If
                       
                    If HardyCoordinateOutput Then
                        If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                           XGeo = val(Format(str$(XGeo), "#####0.0##")) 'CLng(XGeo)
                           YGeo = val(Format(str$(YGeo), "######0.0##")) 'CLng(YGeo)
                        ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                           XGeo = val(Format(str$(XGeo), "####0.0######"))
                           YGeo = val(Format(str$(YGeo), "####0.0######"))
                        Else
                           XGeo = CLng(XGeo)
                           YGeo = CLng(YGeo)
                           End If
                    Else
                       XGeo = CLng(XGeo)
                       YGeo = CLng(YGeo)
                       End If
                       
                    Write #filnum%, XGeo, YGeo, zi '* 10#
                    Write #filtopo%, Nint(Xcoord(ixp - 1)), Nint(Ycoord(iyp - 1)), zi '* 10#
                    End If
                
            Next iyp
LLL35:

            If numDigiHardyPoints > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * ixp / np) 'numDigiHardyPoints)
            DoEvents
            
            '---------------------break on ESC key-------------------------------
            If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
               EscapeKeyPressed = True
               GoTo LLL40
               End If
            '---------------------------------------------------------------------

        Next ixp
LLL40:
        Close #filnum%
        Close #filtopo%
        
        g_nrows = np
        g_ncols = np
       
LLL50:
       'reclaim memory allocated to Matrix arrays
       ReDim Matrix_A(1 To 1, 1 To 1)
       ReDim Operations_Matrix(1 To 1, 1 To 2)   'Matrix where the calculations are done
       ReDim Inverse_Matrix(1 To 1, 1 To 1)  'Matrix with the Inverse of [A]
'       ReDim C_vector(0)
       
   ElseIf GaussMethod And Not DLLMethod Then 'fastest method, doesn't try to find inverse rather solves simultaneous equations to find the Hardy coeficients
   
'       ReDim x(1 To numDigiHardyPoints)
'       ReDim Y(1 To numDigiHardyPoints)
'       ReDim Z(1 To numDigiHardyPoints)
'       ReDim Xcoord(0 To numDigiHardyPoints - 1)
'       ReDim Ycoord(0 To numDigiHardyPoints - 1)
'       ReDim ht(0 To numDigiHardyPoints - 1, 0 To numDigiHardyPoints - 1)
   
       GDMDIform.StatusBar1.Panels(1).Text = "Beginning matrix inversion (using Gaussian elimination method) stage of calculating Hardy quadratic surfaces."
       Call UpdateStatus(GDMDIform, 1, 0)
       
       'redimension arrays
       ReDim Matrix_A(1 To np, 1 To np + 1)
       ReDim C_vector(1 To np)
       
        '-----------------pack the array (replaces "topo")------------------------
        Call UpdateStatus(GDMDIform, 1, 0)
        GDMDIform.StatusBar1.Panels(1).Text = "Beginning calculating Hardy quadratic surfaces."
LG5:
           
        For i = 1 To np
           For j = 1 To i
               If (i <> j) Then GoTo LG10
               Matrix_A(i, j) = 0#
               GoTo LG15
LG10:
               dX = x(j) - x(i)
               dy = Y(j) - Y(i)
               If (dX = 0# And dy = 0#) Then GoTo LG25
'               Matrix_A(i, j) = Sqr(dX ^ 2 + dy ^ 2)
               Matrix_A(i, j) = Sqr(dX * dX + dy * dy)
               Matrix_A(j, i) = Matrix_A(i, j)
LG15:
           Next j
LG20:
           Matrix_A(i, np + 1) = Z(i)
'           DoEvents
        Next i
        GoTo LG40
    '     two points are the same, remove one and repeat the process
LG25:
        If (i = np) Then GoTo LG35
        npm1 = np - 1
        For K = i To npm1
            l = K + 1
            x(K) = x(l)
            Y(K) = Y(l)
LG30:
            Z(K) = Z(l)
'            DoEvents
        Next K
LG35:
        np = np - 1
        DoEvents
        If npm1 > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * (numDigiHardyPoints - np + 2) / npm1)
        GoTo LG5
LG40:

       System_DIM = np

       Call UpdateStatus(GDMDIform, 1, 0)
       GDMDIform.StatusBar1.Panels(1).Text = "Inverting coeficient matrix using Gaussian Elimination...press ""ESC"" to abort..."

       ier = SolveEquations(EscapeKeyPressed, Z)
       If ier < 0 Then GoTo LGL50
       
       'multiply inverse by z to get the coeficients, c
       
       GDMDIform.StatusBar1.Panels(1).Text = "Calculating Hardy quadratic surfaces coeficients..."
       GDMDIform.picProgBar.Visible = True
       Call UpdateStatus(GDMDIform, 1, 0)
      
       If Save_xyz% = 1 Then
       
            '-------------------generate temporary xyz file---------------------
             tryingtokill = True
             
             If Dir(App.Path & "\topo_coord.xyz") <> gsEmpty Then Kill App.Path & "\topo_coord.xyz"
             If Dir(App.Path & "\topo_pixel.xyz") <> gsEmpty Then Kill App.Path & "\topo_pixel.xyz"
             
             tryingtokill = False
            
             Call UpdateStatus(GDMDIform, 1, 0)
             GDMDIform.StatusBar1.Panels(1).Text = "Packing height array and writing xyz data, please wait....(press ""ESC"" key to abort)"
             
             filnum% = FreeFile
             Open App.Path & "\topo_coord.xyz" For Output As #filnum%
             filtopo% = FreeFile
             Open App.Path & "\topo_pixel.xyz" For Output As #filtopo%
             
        Else
            GDMDIform.StatusBar1.Panels(1).Text = "Packing height array, please wait....(press ""ESC"" key to abort)"
            End If

        
        xmax = X2
        xmin = X1
        ymax = Y2
        ymin = Y1
        
        Reduce_Dim = 2
        
        np = CLng(np / Reduce_Dim)
            
        StepX = (xmax - xmin) / (np - 1)
        StepY = (ymax - ymin) / (np - 1)
        
        For ixp = 1 To np
            xi = xmin + (ixp - 1) * StepX
            Xcoord(ixp - 1) = xi
            For iyp = 1 To np
                yi = ymin + (iyp - 1) * StepY
                m = np + 1
                zi = 0#
                For j = 1 To System_DIM 'np
                    dX = xi - x(j)
                    dy = yi - Y(j)
LGL25:
                    zi = zi + C_vector(j) * Sqr(dX * dX + dy * dy)
'                    DoEvents
                Next j
                
                If HeightPrecision = 0 Then
                    hts(ixp - 1, iyp - 1) = zi
                ElseIf HeightPrecision = 1 Then
                    htf(ixp - 1, iyp - 1) = zi
                ElseIf HeightPrecision = 2 Then
                    ht(ixp - 1, iyp - 1) = zi
                    End If
                
                Ycoord(iyp - 1) = yi
                
                If Save_xyz% = 1 Then
                    If HardyCoordinateOutput Then
                        'convert pixel coordinates to geographic coordinates  'use this to output final values
                        If RSMethod1 Then
                           ier = RS_pixel_to_coord2(CDbl(xi), CDbl(yi), XGeo, YGeo)
                        ElseIf RSMethod2 Then
                           ier = RS_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                        ElseIf RSMethod0 Then
                           ier = Simple_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                           End If
                    Else
                       'output pixel coordinates
                       XGeo = xi
                       YGeo = yi
                       End If
                       
                    If HardyCoordinateOutput Then
                        If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                           XGeo = val(Format(str$(XGeo), "#####0.0##")) 'CLng(XGeo)
                           YGeo = val(Format(str$(YGeo), "######0.0##")) 'CLng(YGeo)
                        ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                           XGeo = val(Format(str$(XGeo), "####0.0######"))
                           YGeo = val(Format(str$(YGeo), "####0.0######"))
                        Else
                           XGeo = CLng(XGeo)
                           YGeo = CLng(YGeo)
                           End If
                    Else
                       XGeo = CLng(XGeo)
                       YGeo = CLng(YGeo)
                       End If
                       
                    If Save_xyz% = 1 Then
                        Write #filnum%, XGeo, YGeo, zi '* 10#
                        Write #filtopo%, Nint(Xcoord(ixp - 1)), Nint(Ycoord(iyp - 1)), zi '* 10#
                        End If
                    
                    End If
                
            Next iyp
LGL35:

            If numDigiHardyPoints > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * ixp / np) 'numDigiHardyPoints)
            DoEvents
            
            '---------------------break on ESC key-------------------------------
            If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
               EscapeKeyPressed = True
               GoTo LGL40
               End If
            '---------------------------------------------------------------------

        Next ixp
        
LGL40:
        If Save_xyz% = 1 Then
            Close #filnum%
            Close #filtopo%
            g_nrows = np
            g_ncols = np
            End If
       
LGL50:
       'reclaim memory allocated to Matrix arrays
       ReDim Matrix_A(1 To 1, 1 To 1)
       ReDim Operations_Matrix(1 To 1, 1 To 2 * (1))  'Matrix where the calculations are done
'       ReDim C_vector(0)
       
   ElseIf GaussMethod And DLLMethod Then 'use dll to cacluate everything
        
        np = numDigiHardyPoints
        
        Dim np2 As Long
        'np2 = (np + 1) * (np + 1)
        np2 = np * np
        
        'Make the C dll arrays start to start non zero content from index of 1 (in order to keep all the code the same).
        ReDim x(np)
        ReDim Y(np)
        ReDim Z(np)
        ReDim Xcoord(np - 1)
        ReDim Ycoord(np - 1)
        
        Select Case (HeightPrecision)
        
           Case 0 'integer precision
           
            ReDim ht2s(np2)
            ReDim hts(np - 1, np - 1)
            ReDim ht2f(0)
            ReDim ht2(0)
            ReDim htf(0)
            ReDim ht(0)
           
           Case 1 'single precision
           
            ReDim ht2f(np2)
            ReDim htf(np - 1, np - 1)
            ReDim ht2s(0)
            ReDim hts(0)
            ReDim ht2(0)
            ReDim ht(0)
           
           Case 2 'double precision
        
            ReDim ht2(np2)
            ReDim ht(np - 1, np - 1)
            ReDim ht2f(0)
            ReDim htf(0)
            ReDim hts(0)
            ReDim ht2s(0)
            
        End Select
        
        ReDim C_vector(np)
'        ReDim Matrix_A(np + 1, np + 1)
    
        GDMDIform.StatusBar1.Panels(1).Text = "Stuffing data into arrays, please wait...."
        Call UpdateStatus(GDMDIform, 1, 0)
        
        For i = 1 To np
           x(i) = DigiHardyPoints(i - 1).x
           If x(i) < xmin Then xmin = x(i)
           If x(i) > xmax Then xmax = x(i)
           Y(i) = DigiHardyPoints(i - 1).Y
           If Y(i) < ymin Then ymin = Y(i)
           If Y(i) > ymax Then ymax = Y(i)
           Z(i) = DigiHardyPoints(i - 1).Z
           If Z(i) < zmin Then zmin = Z(i)
           If Z(i) > zmax Then zmax = Z(i)
           Call UpdateStatus(GDMDIform, 1, 100 * i / numDigiHardyPoints)
        Next i
           
        GDMDIform.picProgBar.Visible = False
    
        np = numDigiHardyPoints
   
        xmax = X2
        xmin = X1
        ymax = Y2
        ymin = Y1
            
        StepX = (xmax - xmin) / (np - 1)
        StepY = (ymax - ymin) / (np - 1)
        
'        '============debugging==============
'        fileout% = FreeFile
'        Open App.Path & "\Map-dump.txt" For Output As #fileout%
'        Write #fileout%, np, xmin, xmax, ymin, ymax
'        For i = 1 To np
'           Write #fileout%, x(i), Y(i), Z(i)
'        Next i
'        Close #fileout%
''        For i = 1 To np
''            Matrix_A(i, i) = 2 * i
''        Next i
'        '===================================
        GDMDIform.StatusBar1.Panels(1).Text = "Calculating Hardy quadratic sufraces using Gaussian elimination," & np & " x " & np & " points....please wait..."
                
'        If Dir(App.Path & "\arrorw-around-globe.avi") <> gsEmpty Then
'           GDMDIform.Picture4.Refresh
'           GDMDIform.ani_prg.Visible = True
'           GDMDIform.ani_prg.Open App.Path & "\arrorw-around-globe.avi"
'           GDMDIform.ani_prg.Play
'           End If
        
        Call UpdateStatus(GDMDIform, 1, 0)
        
        'update the progress bar using the Callback routine MyCallback
        On Error Resume Next
        
        'boost the priority of this class and thread
        'source: http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_10139257.html
'        Call SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)
'        Call SetThreadPriority(GetCurrentThread, THREAD_BASE_PRIORITY_MAX)
        
''        ier = SolveEquationsC(x(0), Y(0), Z(0), Matrix_A(0, 0), C_vector(0), Xcoord(0), Ycoord(0), ht2(0), xmin, StepX, ymin, StepY, np, AddressOf MyCallback)

'        ier = SolveEquationsC(X(0), Y(0), Z(0), C_vector(0), Xcoord(0), Ycoord(0), ht2(0), xmin, StepX, ymin, StepY, np, AddressOf MyCallback)
        
        Dim Precision As Integer
        Precision = HeightPrecision
        ier = SolveEquationsC(x(0), Y(0), Z(0), C_vector(0), Xcoord(0), Ycoord(0), ht2(0), ht2f(0), ht2s(0), Precision, xmin, StepX, ymin, StepY, np, AddressOf MyCallback)
       
        'restore priority of this thread to normal
'        Call SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_NORMAL)
'        Call SetPriorityClass(GetCurrentProcess, NORMAL_PRIORITY_CLASS)
        
'        If Not FreeLibrary(GetModuleHandle("MapDigitizer.dll")) Then
'           'freeing memory failed
'           'give message
'           End If
        
'        If Dir(App.Path & "\arrorw-around-globe.avi") <> gsEmpty Then
'           GDMDIform.ani_prg.Stop
'           GDMDIform.ani_prg.Visible = False
'           GDMDIform.Picture4.Refresh
'
'           GDMDIform.CenterPointTimer.Enabled = True
'           End If
       
'       ReDim Matrix_A(0, 0)
'       ReDim C_vector(0)
       
       'recalculate the Step Size and refill the height array
        On Error GoTo HardyQuadraticSurfaces_Error
       
        StepX = (xmax - xmin) / (np - 1)
        StepY = (ymax - ymin) / (np - 1)
       
'        Call UpdateStatus(GDMDIform, 1, 0)
        If Save_xyz% = 1 Then
           GDMDIform.StatusBar1.Panels(1).Text = "Packing height array and writing xyz data, please wait....(press ""ESC"" key to abort)"
        Else
           GDMDIform.StatusBar1.Panels(1).Text = "Packing height array, please wait....(press ""ESC"" key to abort)"
           End If
        
        If Save_xyz% = 1 Then
            filnum% = FreeFile
            Open App.Path & "\topo_coord.xyz" For Output As #filnum%
            filtopo% = FreeFile
            Open App.Path & "\topo_pixel.xyz" For Output As #filtopo%
            End If
        
        For ixp = 1 To np
            xi = xmin + (ixp - 1) * StepX
            For iyp = 1 To np
                yi = ymin + (iyp - 1) * StepY
                'step from South to North, or from maximum y pixels to minium y pixels
                yii = ymax - (iyp - 1) * StepY
                
                If HeightPrecision = 0 Then
                    zi = ht2s(ixp - 1 + (iyp - 1) * np)
                    zii = ht2s(ixp - 1 + (np - iyp) * np)
                    hts(ixp - 1, iyp - 1) = zi
                ElseIf HeightPrecision = 1 Then
                    zi = ht2f(ixp - 1 + (iyp - 1) * np)
                    zii = ht2f(ixp - 1 + (np - iyp) * np)
                    htf(ixp - 1, iyp - 1) = zi
                ElseIf HeightPrecision = 2 Then
                    zi = ht2(ixp - 1 + (iyp - 1) * np)
                    zii = ht2(ixp - 1 + (np - iyp) * np)
                    ht(ixp - 1, iyp - 1) = zi
                    End If
                
                If Save_xyz% = 1 Then
                    If HardyCoordinateOutput Then
                        'convert pixel coordinates to geographic coordinates  'use this to output final values
'                        If RSMethod1 Then
'                           ier = RS_pixel_to_coord2(CDbl(xi), CDbl(yi), XGeo, YGeo)
'                        ElseIf RSMethod2 Then
'                           ier = RS_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
'                        ElseIf RSMethod0 Then
'                           ier = Simple_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
'                           End If
                           
                        If RSMethod1 Then
                           ier = RS_pixel_to_coord2(CDbl(xi), CDbl(yii), XGeo, YGeo)
                        ElseIf RSMethod2 Then
                           ier = RS_pixel_to_coord(CDbl(xi), CDbl(yii), XGeo, YGeo)
                        ElseIf RSMethod0 Then
                           ier = Simple_pixel_to_coord(CDbl(xi), CDbl(yii), XGeo, YGeo)
                           End If
                          
                    Else
                       'output pixel coordinates
                       XGeo = xi
                       YGeo = yi
                       End If
                       
                    If HardyCoordinateOutput Then
                        If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                           XGeo = val(Format(str$(XGeo), "#####0.0##")) 'CLng(XGeo)
                           YGeo = val(Format(str$(YGeo), "######0.0##")) 'CLng(YGeo)
                        ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                           XGeo = val(Format(str$(XGeo), "####0.0######"))
                           YGeo = val(Format(str$(YGeo), "####0.0######"))
                        Else
                           XGeo = CLng(XGeo)
                           YGeo = CLng(YGeo)
                           End If
                    Else
                       XGeo = CLng(XGeo)
                       YGeo = CLng(YGeo)
                       End If
                       
'                    Write #filnum%, XGeo, YGeo, zi '* 10#
                    Write #filnum%, XGeo, YGeo, zii
'                    Write #filtopo%, Nint(Xcoord(ixp - 1)), Nint(Ycoord(iyp - 1)), zi '* 10#
                    Write #filtopo%, xi, yi, zi
                    End If
               
            Next iyp
            
            If numDigiHardyPoints > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * ixp / np) 'numDigiHardyPoints)
            DoEvents
            
        Next ixp
        
        If Save_xyz% = 1 Then
           Close #filnum%
           Close #filtopo%
           g_nrows = np
           g_ncols = np
           End If
       
       'reclaim memory allocated to Matrix arrays
'       ReDim Matrix_A(1 To 1, 1 To 1)
'       ReDim Operations_Matrix(1 To 1, 1 To 2 * (1))  'Matrix where the calculations are done
'       ReDim C_vector(0)

       If ier < 0 Then 'no solution
          GDMDIform.picProgBar.Visible = False
          HardyQuadraticSurfaces = -1
          GDMDIform.StatusBar1.Panels(1).Text = "No solution found...."
          GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
          Exit Function
          End If
   
       End If
    
LLL60:
    Call UpdateStatus(GDMDIform, 1, 0)
    GDMDIform.picProgBar.Visible = False
    GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
 
    If EscapeKeyPressed Then 'calculation was aborted gently
       GDMDIform.StatusBar1.Panels(1) = sEmpty
       GDMDIform.StatusBar1.Panels(2) = sEmpty
       
       'reclaim memory
       ReDim x(1 To 1)
       ReDim Y(1 To 1)
       ReDim Z(1 To 1)
       ReDim Xcoord(0 To 0)
       ReDim Ycoord(0 To 0)
       ReDim ht2(0 To 0, 0 To 0)
       ReDim ht2s(0 To 0, 0 To 0)
       ReDim ht2f(0 To 0, 0 To 0)
       
       HardyQuadraticSurfaces = ier
       Exit Function
       End If
    
    '-------------------generate contours----------------------------
    GDMDIform.StatusBar1.Panels(1).Text = "Generating and plotting contours, please wait......"
    numContourPoints = 0 'zero contour lines array
    ReDim ContourPoints(numContourPoints)
    ReDim contour(0) 'zero contour color array
    
    nc = Int((zmax - zmin) / ContourInterval) + 1

    For i = 1 To nc
       If i > 0 Then
          ReDim Preserve contour(i)
          End If
       contour(i - 1) = zmin + (i - 1) * ContourInterval
    Next i
    
    ier = ReDrawMap(0)
    If Not InitDigiGraph Then
       InputDigiLogFile 'load up saved digitizing data for the current map sheet
    Else
       ier = RedrawDigiLog
       End If
       
'----------------------diagnositics----------------------------------
'    filetestout% = FreeFile
'    Open App.Path & "\test.txt" For Output As #filetestout%
'    For i = 0 To np - 1
'       For j = 0 To np - 1
'          Write #filetestout%, Xcoord(i), Ycoord(j), hts(i, j)
'       Next j
'    Next i
'    Close #filetestout%
'--------------------------------------------------------------------
    
    ier = conrec(Pic, ht, htf, hts, Xcoord, Ycoord, nc, contour, 0, np - 1, 0, np - 1, xmin, ymin, xmax, ymax, 0)
    If ier = -1 Then
       Call MsgBox("Palette file: rainbow.cpt is missing in the program directory." _
                  & vbCrLf & "" _
                  & vbCrLf & "Contours won't be drawn" _
                  , vbExclamation, "Hardy contours")
       End If
       
'   'enable create DTM button 'reserve warnings for CreateDTM function in case the corner coordinates are still not defined
    GDMDIform.Toolbar1.Buttons(45).Enabled = True
   
   'enable profiling
   If Save_xyz% = 1 And heights Then GDMDIform.Toolbar1.Buttons(51).Enabled = True
    
    '-------------------------finished----------------------------------
   HardyQuadraticSurfaces = ier
   
   On Error GoTo 0
   Exit Function

HardyQuadraticSurfaces_Error:
    If tryingtokill Then Resume Next
    
    HardyQuadraticSurfaces = -1
    If Err.Number = 7 Then HardyQuadraticSurfaces = -2
    GDMDIform.picProgBar.Visible = False
    GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
    GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HardyQuadraticSurfaces of Module modHardy"
End Function
'---------------------------------------------------------------------------------------
' Procedure : generateDTM
' Author    : Dr-John-K-Hall
' Date      : 12/27/2015
' Purpose   : button manager for generating DTMs
'---------------------------------------------------------------------------------------
'
Public Sub generateDTM()

   Dim ier As Integer
   Dim CurrentX As Double, CurrentY As Double
   Dim ShiftX As Double, ShiftY As Double
   
   Dim Tolerance As Double
   On Error GoTo generateDTM_Error

   Tolerance = 0.000015
   
    If buttonstate&(45) = 0 Then
       buttonstate&(45) = 1
       GDMDIform.Toolbar1.Buttons(45).value = tbrPressed
       DTMcreating = True
       
      'check if DTM base file exists, which is the entire map using the default DTM at the grid spacing defined in the Options menu
      'if it doesn't exist, create it now
      DTMfile$ = dirNewDTM & "\" & RootName(picnam$) & ".grd"
      DTMhdrfile$ = dirNewDTM & "\" & RootName(picnam$) & ".hdr"
      If Dir(DTMfile$) = gsEmpty Or Dir(DTMhdrfile$) = gsEmpty Then
         Select Case MsgBox("The initial phase of DTM creation requires the construction of a background DTM encomposssing the entire map using the elevations of the default DTM that is available." _
                            & vbCrLf & "" _
                            & vbCrLf & "Proceed?" _
                            , vbOKCancel Or vbInformation Or vbDefaultButton1, "DTM creation")
         
            Case vbOK
                ier = CreateDTMBackground
                If ier < 0 Then 'exited function after error
                    buttonstate&(45) = 0
                    GDMDIform.Toolbar1.Buttons(45).value = tbrUnpressed
                    buttonstate&(45) = 0
                    DTMcreating = False
                    BasisDTMheights = False
                    If basedtm% > 0 Then Close #basedtm%
                Else
                    BasisDTMheights = True 'flag that can use this DTM for displaying heights
                    DTMfile$ = dirNewDTM & "\" & RootName(picnam$) & ".grd"
                    basedtm% = FreeFile
                    Open DTMfile$ For Binary As #basedtm%
                    
                    'show height text boxes
                    
                    End If
                    
                If (heights Or BasisDTMheights) And (RSMethod1 Or RSMethod2 Or RSMethod0) And DigiRubberSheeting Then
            
                    'enable search height button
                    GDMDIform.Toolbar1.Buttons(50).Enabled = True
                    'enable contour generation
                    GDMDIform.Toolbar1.Buttons(51).Enabled = True
                    
                    GDMDIform.Label1 = lblX
                    GDMDIform.Label5 = lblX
                    GDMDIform.Label2 = LblY
                    GDMDIform.Label6 = LblY
                    
                    GDMDIform.Text3.Visible = True
                    GDMDIform.Label3.Visible = True
                    GDMDIform.Text7.Visible = True
                    GDMDIform.Label7.Visible = True
                    
                    GDMDIform.Text4.Visible = True
                    GDMDIform.Label4.Visible = True
                        
                    End If
                    
            Case vbCancel
                buttonstate&(45) = 0
                GDMDIform.Toolbar1.Buttons(45).value = tbrUnpressed
                buttonstate&(45) = 0
                DTMcreating = False
         
         End Select
      Else
         'read the header files and plot filled boxes wherever the DTM was already merged into the background DTM
        If LRGeoX <> ULGeoX And ULGeoY <> LRGeoY Then
            Dim GeoToPixelX As Double, GeoToPixelY As Double
            Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
            Dim X3 As Double, Y3 As Double, X4 As Double, Y4 As Double
            Dim GeoX As Double, GeoY As Double
            Dim XGeo As Double, YGeo As Double
            Dim XDif As Double, YDif As Double
       
            GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
            GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
             
             
             filhdr% = FreeFile
             Open DTMhdrfile$ For Input As #filhdr%
             Input #filhdr%, nRowLL
             Input #filhdr%, nColLL
             Input #filhdr%, xLL
             Input #filhdr%, yLL
             Input #filhdr%, XStepLL
             Input #filhdr%, YStepLL
             Input #filhdr%, zminLL
             Input #filhdr%, zmaxLL
             Input #filhdr%, AngLL
             Input #filhdr%, blank_LL
             
             'now read the coordinates of the mergers and overplot them on the map as a colored box
             Dim xin(0 To 10)
             gdfs = GDform1.Picture2.FillStyle
             gdco = GDform1.Picture2.FillColor
             gdwi = GDform1.Picture2.DrawWidth
             gdds = GDform1.Picture2.DrawStyle
             
gdtm50:
             ninput% = 0
             Do Until EOF(filhdr%)
                Input #filhdr%, xin(ninput%)
                ninput% = ninput% + 1
                If ninput% = 7 Then
                   If xin(ninput%) < zminLL Then
                      zminLL = xin(ninput%)
                      End If
                ElseIf ninput% = 8 Then
                   If xin(ninput%) > zmaxLL Then
                      zmaxLL = xin(ninput%)
                      End If
                   End If
                If ninput% = 10 Then Exit Do
             Loop
             
             If ninput% = 10 Then 'succcessfully read a merge region's boundaries
                'plot the merged regions
                                                 
                GeoX = xin(2)
                GeoY = xin(3)
                GoSub GeotoCoord
                X1 = CurrentX * DigiZoom.LastZoom
                Y1 = CurrentY * DigiZoom.LastZoom
                GeoX = xin(2) + xin(1) * xin(4)
                GoSub GeotoCoord
                X2 = CurrentX * DigiZoom.LastZoom
                Y2 = CurrentY * DigiZoom.LastZoom
                GeoY = xin(3) + xin(0) * xin(5)
                GoSub GeotoCoord
                X3 = CurrentX * DigiZoom.LastZoom
                Y3 = CurrentY * DigiZoom.LastZoom
                GeoX = xin(2)
                GoSub GeotoCoord
                X4 = CurrentX * DigiZoom.LastZoom
                Y4 = CurrentY * DigiZoom.LastZoom
                
                If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                   'draw box
                   GDform1.Picture2.FillStyle = 4
                   GDform1.Picture2.DrawMode = 9 '8 '5 '2 '3 '5
                   GDform1.Picture2.Line (X1, Y1)-(X3, Y3), , BF
                Else
'                    GDform1.Picture2.DrawMode = 13
'                    GDform1.Picture2.DrawWidth = 2
                    GDform1.Picture2.FillStyle = 4
                    GDform1.Picture2.DrawMode = 9
'                    GDform1.Picture2.DrawStyle = vbDot
                    GDform1.Picture2.FillColor = QBColor(14)
'                    xmax = Max(X2, X3)
'                    xmin = min(X1, X4)
'                    ymax = Max(Y1, Y2)
'                    ymin = min(Y3, Y4)
'                    GDform1.Picture2.Line (xmin, ymin)-(xmax, ymax), , BF
                    GDform1.Picture2.Line (X1, Y1)-(X2, Y2) ', , BF
                    GDform1.Picture2.Line (X2, Y2)-(X3, Y3) ', , BF
                    GDform1.Picture2.Line (X3, Y3)-(X4, Y4) ', , BF
                    GDform1.Picture2.Line (X4, Y4)-(X1, Y1) ', , BF
                    End If
                GoTo gdtm50
             Else
                Close #filhdr%
                GDform1.Picture2.FillStyle = gdfs
                GDform1.Picture2.FillColor = gdco
                GDform1.Picture2.DrawWidth = gdwi
                GDform1.Picture2.DrawStyle = gdds
                End If
                
             'also allow for smoothing
             GDMDIform.Toolbar1.Buttons(46).Enabled = True
                
            Else
                Call MsgBox("Previously merged regions cannot be plotted until you define the" _
                            & vbCrLf & "pixel coordinates of the map's corners.  " _
                            & vbCrLf & "" _
                            & vbCrLf & "Use the Option menu to define them" _
                            , vbInformation Or vbDefaultButton1, "DTM merging")
            
               End If
            
         End If
         
         'enable smothing
         GDMDIform.Toolbar1.Buttons(46).Enabled = True
       
    Else
       buttonstate&(45) = 0
       GDMDIform.Toolbar1.Buttons(45).value = tbrUnpressed
       buttonstate&(45) = 0
       DTMcreating = False
       'disable smoothing
       GDMDIform.Toolbar1.Buttons(46).Enabled = False
       Belgier_Smoothing = False
       buttonstate&(46) = 0
       GDMDIform.Toolbar1.Buttons(46).value = tbrUnpressed
       
       'disenable smothing
       GDMDIform.Toolbar1.Buttons(46).Enabled = False
       
       'remove the boxes showing the merged regions
       ier = ReDrawMap(0)
       Call PictureBoxZoom(GDform1.Picture2, 0, 0, 0, 0, 0)
       End If
       
Exit Sub

GeotoCoord:
    
    CurrentX = ((GeoX - ULGeoX) * GeoToPixelX) + ULPixX
    CurrentY = ((ULGeoY - GeoY) * GeoToPixelY) + ULPixY
    
    If RSMethod1 Or RSMethod2 Then
       
       If RSMethod1 Then
          ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
       ElseIf RSMethod2 Then
          ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
          End If
          
        Dim DifX As Double, DifY As Double
        DifX = Abs(GeoX - XGeo)
        DifY = Abs(GeoY - YGeo)
       
        ShiftX = CurrentX - (((XGeo - ULGeoX) * GeoToPixelX) + ULPixX)
        ShiftY = CurrentY - (((ULGeoY - YGeo) * GeoToPixelY) + ULPixY)
        
        CurrentX = CurrentX + ShiftX
        CurrentY = CurrentY + ShiftY
        
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
         ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
           End If

        If Abs(GeoX - XGeo) > DifX Then
           CurrentX = CurrentX - ShiftX
           End If
           
        If Abs(GeoY - YGeo) > DifY Then
           CurrentY = CurrentY - ShiftY
           End If

'        If Abs(GeoX - XGeo) > DifX And Abs(GeoY - YGeo) > DifY Then
''        If Abs(GeoX - XGeo) > Tolerance Or Abs(GeoY - YGeo) > Tolerance Then
'                Call MsgBox("Inverse coordinate transformation unsuccessful" _
'                        & vbCrLf & "Coordinate grid rotation too large for first approx." _
'                        & vbCrLf & vbCrLf & "(Redo using a less-rotated grid as reference...)" _
'                        , vbInformation, "Generate DTM Error")
'              Screen.MousePointer = vbDefault
'              GDMDIform.picProgBar.Visible = False
'              GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
'              GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
'              Exit Sub
'              End If
              
   Else
        'cuurentx, currenty are the pixel coordinates
        End If
Return
   
'   hidden graphic planes
'   'merge the files, or recalculate?
'   GDform1.PictureBlit.Picture = LoadPicture(picnam$)
   
   'now blit it half size to the picture2
'   DigiZoom.LastZoom = 0.5
'   ier = ReDrawMap(0)
   
'   GDform1.Picture2.Cls
'   GDform1.Picture2.Refresh
'
'   GDform1.Picture2.Width = 0.5 * GDform1.Picture2.Width
'   GDform1.Picture2.Height = 0.5 * GDform1.Picture2.Height
'
'   GDform1.PictureBlit.PaintPicture GDform1.Picture2, -1000, 0, GDform1.Picture2.ScaleWidth, GDform1.Picture2.ScaleHeight
   

   On Error GoTo 0
   Exit Sub

generateDTM_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generateDTM of Module modHardy"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CreateDTM
' Author    : Dr-John-K-Hall
' Date      : 8/18/2015
' Purpose   : Creates DTM from Hardy results or repairs exisiting DTM's from those results
'---------------------------------------------------------------------------------------
'
Public Function CreateDTM(Pic As PictureBox, RectCoord() As POINTAPI) As Integer

   On Error GoTo CreateDTM_Error
   
   Dim XStep As Double, YStep As Double
   Dim SizeX As Double, SizeY As Double
   
   Dim SearchGeoCoord(1) As POINTGEO
   Dim i As Long, j As Long
   Dim XGeo As Double
   Dim YGeo As Double
   Dim numXsteps&, numYsteps&
   Dim np As Long
   
   Dim Xcoord() As Double, Ycoord() As Double, Zcoord() As Double
   Dim x() As Double
   Dim Y() As Double
   
   Dim X1, Y1, X2, Y2
   
    Dim zmin As Double, zmax As Double
    zmin = INIT_VALUE
    zmax = -INIT_VALUE
           
   Dim hgtNews As Integer
   
   Dim ier As Integer
   
   Dim Tolerance As Double
   Tolerance = 0.00001
   
   Dim VarS As Integer, VarL As Long, VarF As Single, VarD As Double
   Dim BytePosit As Long
   Dim DLLMethod As Boolean
   
   Dim GeoToPixelX As Double, GeoToPixelY As Double
   Dim CurrentX As Double, CurrentY As Double
   Dim Pix(1) As POINTAPI
   
   GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
   GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
   
   DLLMethod = True
   
   ier = 0

   'create or fix old dtm
   
    If LRGeoX = ULGeoX Or ULGeoY = LRGeoY Then
       'use rubber sheeting to determine them
       MsgBox "Corner grid coordinates undefined." & vbCrLf & vbCrLf & "(Hint: use options menu)", vbInformation + vbOKOnly, "DTM creation error"
       GDMDIform.Toolbar1.Buttons(45).Enabled = False
       buttonstate&(45) = 0
       GDMDIform.Toolbar1.Buttons(45).value = tbrUnpressed
       ier = -1
       CreateDTM = ier
       Exit Function
       End If
   
   If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
    
       X1 = RectCoord(0).x
       Y1 = RectCoord(0).Y
       X2 = RectCoord(1).x
       Y2 = RectCoord(1).Y
       
       If X1 < XminC Or X2 > XmaxC Or Y1 < YminC Or Y2 > YmaxC Then
          Call MsgBox("DTM boundaries are outside the Hardy countours!" _
                      & vbCrLf & "" _
                      & vbCrLf & "(Hint: drag only within the contoured area)" _
                      , vbQuestion, App.Title)
          
          ier = -1
          CreateDTM = ier
          Exit Function
          End If
       
       ier = 0
    
'       If mode% <= 3 Then
         frmMsgBox.MsgCstm "Please press one of the following buttons..." _
                          & vbCrLf & "", _
                          "DEM repair and DTM merging", mbinformation, 2, False, _
                          "Repair DEM", "Splice to base DTM", "Cancel"
        
         Select Case frmMsgBox.g_lBtnClicked
        
            Case 1
            
                If heights = False Then
                   Call MsgBox("This option is not valid since there is only a base DTM." _
                               & vbCrLf & "So start again and pick option 2." _
                               , vbInformation, "Editing DTM")
                    ier = -1
                    CreateDTM = ier
                    Exit Function
                   
                    End If
                   
                'repair the current DTM
                'determine boundaries that fit within the selected region
                
                If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                    XStep = 25
                    YStep = 25
                ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                    XStep = 30
                    YStep = 30
                ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                    XStep = 8.33333333333333E-04 / 3#
                    YStep = 8.33333333333333E-04 / 3#
                   End If
                   
            Case 2
                'merging DTM on the basis DTM
                
               DTMfile$ = dirNewDTM & "\" & RootName(picnam$) & ".grd"
               If Dir(DTMfile$) <> gsEmpty Then
                  
                 If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                    XStep = XStepITM
                    YStep = YStepITM
                 ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                    XStep = XStepITM
                    YStep = YStepITM
                 ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                    XStep = XStepDTM * 8.33333333333333E-04 / 3#
                    YStep = YStepDTM * 8.33333333333333E-04 / 3#
                 Else
                    XStep = XStepITM
                    YStep = YStepITM
                    End If
                    
              Else
                 
                 Call MsgBox("The basis DTM not found!", vbExclamation Or vbDefaultButton1, "DTM merging")
                 ier = -1
                 CreateDTM = ier
                 Exit Function
                 End If
                
            Case 0, 3
                ier = -1
                CreateDTM = ier
                Exit Function
                
          End Select

         'record elevations in the selected region
       
          Screen.MousePointer = vbHourglass
       
          'determine DTM spacing in pixels
             
          '------------------progress bar initialization
          With GDMDIform
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
          
          Call UpdateStatus(GDMDIform, 1, 0)
          GDMDIform.StatusBar1.Panels(1).Text = "Determining bounds, and screen coordinates of merge region, please wait...."
        
    '      iprogress& = 0
    '
          'convert the pixel search coordinates to geo coordinates
          For i = 0 To 1
    
             If RSMethod1 Then
                ier = RS_pixel_to_coord2(CDbl(RectCoord(i).x), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
             ElseIf RSMethod2 Then
                ier = RS_pixel_to_coord(CDbl(RectCoord(i).x), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
             ElseIf RSMethod0 Then
                ier = Simple_pixel_to_coord(CDbl(RectCoord(i).x), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
                End If
    
            'now determine new boundaries so that boundaries are on data point
            'determine lat,lon of first element of this tile
    
            If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
    
                If DTMtype = 2 Then
                    lg1 = Int(SearchGeoCoord(i).XGeo)
                    If lg1 < 0 And lg1 > lg2 Then lg1 = lg1 - 1
                    lt1 = Int(SearchGeoCoord(i).YGeo) 'SRTM tiles are named by SW corner
                    If lt1 < 0 And lt1 > lt Then lt1 = lt1 - 1
    
                    lt1 = lt1 + 1 'the first record in SRTM tiles in the NW corner
                    SearchGeoCoord(i).XGeo = Int((SearchGeoCoord(i).XGeo - lg1) / XStep) * XStep + lg1
                    SearchGeoCoord(i).YGeo = lt1 - Int((lt1 - SearchGeoCoord(i).YGeo) / YStep) * YStep
    
                ElseIf DTMtype = 1 Then 'ASTER
                    lg1 = Int(SearchGeoCoord(i).XGeo)
                    If lg1 < 0 And lg1 > lg2 Then lg1 = lg1 - 1
                    lt1 = Int(SearchGeoCoord(i).YGeo) 'ASTER tiles are named by SW corner
                    If lt1 < 0 And lt1 > lt Then lt1 = lt1 - 1
    
                    'first data point is in SW corner
                    SearchGeoCoord(i).XGeo = Int((SearchGeoCoord(i).XGeo - lg1) / XStep) * XStep + lg1
                    SearchGeoCoord(i).YGeo = Int((SearchGeoCoord(i).YGeo - lt1) / YStep) * YStep + lt1
    
                    End If
    
             ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
    
                If JKHDTM Then
                   '25 meter spacing
                   SearchGeoCoord(i).XGeo = Int(SearchGeoCoord(i).XGeo / 25) * 25
                   SearchGeoCoord(i).YGeo = Int(SearchGeoCoord(i).YGeo / 25) * 25
                Else
                   'approximate 30m spacing for latitudes of Eretz Yisroel
                   SearchGeoCoord(i).XGeo = Int(SearchGeoCoord(i).XGeo / 30) * 30
                   SearchGeoCoord(i).YGeo = Int(SearchGeoCoord(i).YGeo / 30) * 30
                   End If
                   
            Else
            
                SearchGeoCoord(i).XGeo = Int(SearchGeoCoord(i).XGeo / XStep) * XStep
                SearchGeoCoord(i).YGeo = Int(SearchGeoCoord(i).YGeo / YStep) * YStep
    
                End If
    
         Next i
         
         'convert Geo coordinates back into pixels
         
         For i = 0 To 1
         
            GeoX = SearchGeoCoord(i).XGeo
            GeoY = SearchGeoCoord(i).YGeo
            
            GoSub GeotoCoord

            Pix(i).x = CurrentX
            Pix(i).Y = CurrentY
         
         Next i
         
         'now determine numXsteps&
         
         numXsteps& = Int((SearchGeoCoord(1).XGeo - SearchGeoCoord(0).XGeo) / XStep) + 1
         numYsteps& = Int((SearchGeoCoord(0).YGeo - SearchGeoCoord(1).YGeo) / YStep) + 1
         
         ReDim Xcoord(0 To numXsteps& - 1)
         ReDim Ycoord(0 To numYsteps& - 1)
         ReDim Zcoord(0 To numXsteps& - 1, 0 To numYsteps& - 1)
         
         For i = 0 To numXsteps& - 1
            GeoX = SearchGeoCoord(0).XGeo + (i - 1) * XStep
            For j = 0 To numYsteps& - 1
                GeoY = SearchGeoCoord(1).YGeo + (j - 1) * YStep
                GoSub GeotoCoord
                Xcoord(i) = CurrentX
                Ycoord(j) = CurrentY
            Next j
            DoEvents
            Call UpdateStatus(GDMDIform, 1, CLng(100 * i / (numXsteps& - 1)))
         Next i
         
        ReDim x(0 To numDigiHardyPoints - 1)
        ReDim Y(0 To numDigiHardyPoints - 1)
       
        GDMDIform.StatusBar1.Panels(1).Text = "Stuffing data into arrays, please wait...."
        Call UpdateStatus(GDMDIform, 1, 0)
        
        For i = 0 To numDigiHardyPoints - 1
           x(i) = DigiHardyPoints(i).x
           Y(i) = DigiHardyPoints(i).Y
           Call UpdateStatus(GDMDIform, 1, 100 * i / numDigiHardyPoints)
        Next i
    
        np = numDigiHardyPoints
        
        If Not DLLMethod Then
        '==============alternative way of calculating the heights using VB6 instead of the dll==============

            'diagnostics======================================================
'            newfilnum% = FreeFile
'            Open App.Path & "\testcode.txt" For Output As #newfilnum%
'            Print #newfilnum%, "xi, X[k], dx, yi, Y[k], dy, C_vector[k+1], zi"
            '=================================================================
    
            Dim xi As Double, yi As Double, zi As Double
            Dim K As Long, dX As Double, dy As Double
            Dim l As Long, npm1 As Long
            
            Call UpdateStatus(GDMDIform, 1, 0)
            
LG5:
            For i = 1 To np
               For j = 1 To i
                   If j <> i And (x(j - 1) = x(i - 1) And Y(j - 1) = Y(i - 1)) Then GoTo LG25
               Next j
            Next i
            GoTo LG40
           'two points are the same, remove one and repeat the process
LG25:
            If (i = np) Then GoTo LG35
            npm1 = np - 1
            For K = i To npm1
                l = K + 1
                x(K - 1) = x(l - 1)
                Y(K - 1) = Y(l - 1)
LG30:
            Next K
LG35:
            np = np - 1
            DoEvents
            If npm1 > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * (numDigiHardyPoints - np + 2) / npm1)
            GoTo LG5
LG40:
        
            Call UpdateStatus(GDMDIform, 1, 0)
            For i = 0 To numXsteps& - 1
                xi = Xcoord(i)
                For j = 0 To numYsteps& - 1
                    yi = Ycoord(j)
                    zi = 0#
                    For K = 0 To np - 1
                        dX = xi - x(K)
                        dy = yi - Y(K)
                        If dX <> 0 And dy <> 0 Then
                            zi = zi + C_vector(K + 1) * Sqr(dX * dX + dy * dy)
                            End If
    
                        'diagnostics===================================================
'                        Write #newfilnum%, xi, x(k), dX, yi, Y(k), dy, C_vector(k + 1), zi
                        '================================================================
    
                    Next K
                    Zcoord(i, j) = zi
                    zmax = Max(zmax, Zcoord(i, j))
                    zmin = min(zmin, Zcoord(i, j))
'                    If zmax > zmaxh Or zmin < zminh Then
'                       cc = 1
'                       End If
                    'diagnostics===================================================
'                    Write #newfilnum%, zi
                    'diagnostics==================================================
    
                Next j
    
                Call UpdateStatus(GDMDIform, 1, i * 100 / (numXsteps& - 1))
    
            Next i
            Call UpdateStatus(GDMDIform, 1, 0)
            GDMDIform.picProgBar.Visible = False

'            'diagnostics=========================================
'            Close #newfilnum%
'            '=====================================================
         
        Else '//////////////////////use dll method////////////////////////////
            Dim ht2() As Double
            Dim ht2s() As Integer
            Dim ht2f() As Single
    
            If HeightPrecision = 0 Then
               ReDim ht2s(0 To numXsteps& * numYsteps& - 1)
               ReDim ht2f(0)
               ReDim ht2(0)
            ElseIf HeightPrecision = 1 Then
               ReDim ht2f(0 To numXsteps& * numYsteps& - 1)
               ReDim ht2s(0)
               ReDim ht2(0)
            ElseIf HeightPrecision = 2 Then
               ReDim ht2(0 To numXsteps& * numYsteps& - 1)
               ReDim ht2s(0)
               ReDim ht2f(0)
               End If
    
            Dim npp&
    
            npp& = np
           
            GDMDIform.StatusBar1.Panels(1).Text = "Calculating DTM points, please wait...."

    '        Call SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)
    '        Call SetThreadPriority(GetCurrentThread, THREAD_BASE_PRIORITY_MAX)
        
    
    '        ier = SolveEquationsDTM(npp&, numYsteps&, numXsteps&, _
    '                                X(0), Y(0), C_vector(0), _
    '                                Xcoord(0), Ycoord(0), ht2(0), _
    '                                AddressOf MyCallback)
    
            Dim Precision As Integer
            Precision = HeightPrecision
            ier = SolveEquationsDTM(npp&, numYsteps&, numXsteps&, _
                                    x(0), Y(0), C_vector(0), _
                                    Xcoord(0), Ycoord(0), _
                                    ht2(0), ht2f(0), ht2s(0), Precision, _
                                    AddressOf MyCallback)
    
            Call UpdateStatus(GDMDIform, 1, 0)
            GDMDIform.picProgBar.Visible = False
            GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
                                    
'            If Not FreeLibrary(GetModuleHandle("MapDigitizer.dll")) Then
'               'freeing memory failed
'               'give message
'               End If
    
            Call UpdateStatus(GDMDIform, 1, 0)
            If HeightPrecision = 0 Then
                For i = 0 To numXsteps& - 1
                   For j = 0 To numYsteps& - 1
                      Zcoord(i, j) = ht2s(i + j * numXsteps&)
                      zmax = Max(zmax, Zcoord(i, j))
                      zmin = min(zmin, Zcoord(i, j))
                   Next j
                   If numXsteps& > 1 Then Call UpdateStatus(GDMDIform, 1, CLng(100 * i / (numXsteps& - 1)))
                Next i
            ElseIf HeightPrecision = 1 Then
                For i = 0 To numXsteps& - 1
                   For j = 0 To numYsteps& - 1
                      Zcoord(i, j) = ht2f(i + j * numXsteps&)
                      zmax = Max(zmax, Zcoord(i, j))
                      zmin = min(zmin, Zcoord(i, j))
                   Next j
                   If numXsteps& > 1 Then Call UpdateStatus(GDMDIform, 1, CLng(100 * i / (numXsteps& - 1)))
                Next i
            ElseIf HeightPrecision = 2 Then
                For i = 0 To numXsteps& - 1
                   For j = 0 To numYsteps& - 1
                      Zcoord(i, j) = ht2(i + j * numXsteps&)
                      zmax = Max(zmax, Zcoord(i, j))
                      zmin = min(zmin, Zcoord(i, j))
                   Next j
                   If numXsteps& > 1 Then Call UpdateStatus(GDMDIform, 1, CLng(100 * i / (numXsteps& - 1)))
                Next i
                End If
                
            'restore priority of this thread to normal
    '        Call SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_NORMAL)
    '        Call SetPriorityClass(GetCurrentProcess, NORMAL_PRIORITY_CLASS)
            
            GDMDIform.prbSearch.Visible = False
            Screen.MousePointer = vbDefault
            
            End If
       
        'output is the Zcoord()
         'step in grid, determine height at this point, write to DTM file
         
         Select Case frmMsgBox.g_lBtnClicked
         
            Case 1
               'repair DTM
               'first make backup then replace
               If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
               
                  If DTMtype = 1 Then 'ASTER (1 arc second)
                     GoSub ModifyAster
                     
                  ElseIf DTMtype = 2 Then '1 arc sec SRTM
                     GoSub ModifySRTM
                     
                     End If
               
               ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                  
                  If JKHDTM Then
                     GoSub ModifyJKHDTM
                  Else
                     MsgBox "This operation is not yet supported for 30m DTMs of Eretz Yisroel"
                     End If
                     
                  End If
            
            Case 2
               'merge this section onto the basis DTM
               'find the BytePosit where overwrite begins
               
               
               'also update the hdr file with the edited region
               'each data point at coordinates (kmx,kmy) is written to byte position
                'BytePosit = 101 + 8 * ((kmx - xLL) / XStep) + 8 * nColLL * ((kmy - yLL) / YStep)
                
                'first check that the merged region is compatible with the map
                If Abs(XStep - XStepLL) > XStepLL * 0.0001 Or Abs(YStep - YStepLL) > YStepLL * 0.001 Or _
                   SearchGeoCoord(0).XGeo < xLL Or SearchGeoCoord(1).XGeo > xLL + nColLL * XStepLL Or _
                   SearchGeoCoord(1).YGeo < yLL Or SearchGeoCoord(0).YGeo > yLL + nRowLL * YStepLL Then
                   Call MsgBox("The basis map's DTM is incompatible with the created DTM" _
                               & vbCrLf & "and cannot be merged!" _
                               , vbInformation Or vbDefaultButton1, "DTM merging")
                   
                   ier = -1
                   CreateDTM = ier
                   GDMDIform.picProgBar.Visible = False
                   GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
                   Exit Function
                   End If
               
                'the first data point
                GDMDIform.StatusBar1.Panels(1).Text = "Merging onto basis DTM file, please wait...."
                
                'first close the binary grd file if it is being used for heights
                If UseNewDTM% = 1 And basedtm% > 0 Then
                   ier = OpenCloseBaseDTM(1)
                   End If
                  
                filedtm% = FreeFile
                Open DTMfile$ For Binary As #filedtm%
                
                filehdr% = FreeFile
                Open dirNewDTM & "\" & RootName(picnam$) & ".hdr" For Append As #filehdr%
                
'                BytePosit = 1
'                VarL = &H42525344
'                Put #filedtm%, BytePosit, VarL 'header tag
'                BytePosit = 5
'                VarL = 4
'                Put #filedtm%, BytePosit, VarL
'                BytePosit = 9
'                VarL = 1
'                Put #filedtm%, BytePosit, VarL
'                BytePosit = 13
'                VarL = &H44495247 'tag for grid section
'                Put #filedtm%, BytePosit, VarL
'                BytePosit = 17
'                VarL = 72
'                Put #filedtm%, BytePosit, VarL
'                BytePosit = 21
                VarL = numYsteps& 'nRow
'                Put #filedtm%, BytePosit, VarL
                Print #filehdr%, VarL
'                BytePosit = 25
                VarL = numXsteps& 'nCol
'                Put #filedtm%, BytePosit, VarL
                Print #filehdr%, VarL
'                BytePosit = 29
                VarD = SearchGeoCoord(0).XGeo 'xLL = min GeoX
'                Put #filedtm%, BytePosit, VarD
                Print #filehdr%, VarD
'                BytePosit = 37
                VarD = SearchGeoCoord(1).YGeo 'yLL = min GeoY
'                Put #filedtm%, BytePosit, VarD
                Print #filehdr%, VarD
'                BytePosit = 45
                VarD = XStep
'                Put #filedtm%, BytePosit, VarD
                Print #filehdr%, VarD
'                BytePosit = 53
                VarD = YStep
'                Put #filedtm%, BytePosit, VarD
                Print #filehdr%, VarD
                BytePosit = 61
                VarD = zmin
                If zmin < zminLL Then 'edit the zmin
                   zminLL = zmin
                   Put #filedtm%, BytePosit, VarD
                   End If
                Print #filehdr%, VarD
                BytePosit = 69
                VarD = zmax
                If zmax > zmaxLL Then 'edit the zmax
                   zmaxLL = zmax
                   Put #filedtm%, BytePosit, VarD
                   End If
                Print #filehdr%, VarD
'                BytePosit = 77
                VarD = ANG
'                Put #filedtm%, BytePosit, VarD
                Print #filehdr%, VarD
'                BytePosit = 85
                VarD = blank_value  'flags blank data value
'                Put #filedtm%, BytePosit, VarD
                Print #filehdr%, VarD
'                BytePosit = 93
'                VarL = &H41544144 'tag for data section
'                Put #filedtm%, BytePosit, VarL
'                BytePosit = 97
'                VarL = numXsteps& * numYsteps& * 8 'byte size of data section
'                Put #filedtm%, BytePosit, VarL
                
'                BytePosit = 101
                Call UpdateStatus(GDMDIform, 1, 0)
                For j = 0 To numYsteps& - 1
                   kmy = SearchGeoCoord(1).YGeo + j * YStep
                   For i = 0 To numXsteps& - 1
                      kmx = SearchGeoCoord(0).XGeo + i * XStep
                      BytePosit = 101 + 8 * ((kmx - xLL) / XStep) + 8 * nColLL * ((kmy - yLL) / YStep)
                      VarD = CDbl(Zcoord(i, j)) * 1# * DigiConvertToMeters
                      Put #filedtm%, BytePosit, VarD
                      BytePosit = BytePosit + 8
                   Next i
                   DoEvents
                   If numYsteps& > 1 Then Call UpdateStatus(GDMDIform, 1, CLng(100 * j / (numYsteps& - 1)))
                Next j
                
                Close #filedtm%
                Close #filehdr%
                
                GDMDIform.picProgBar.Visible = False
                GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
                
                'add box for this merge
                GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
                GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
             
                gdfs = GDform1.Picture2.FillStyle
                gdco = GDform1.Picture2.FillColor
                gdwi = GDform1.Picture2.DrawWidth
                gdds = GDform1.Picture2.DrawStyle
                
                'first erase guide line defining region to be merged
                GDform1.Picture2.DrawMode = 7
                GDform1.Picture2.DrawWidth = Max(5, CInt(DigiZoom.LastZoom))
                GDform1.Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                                                 
                'now highlight the region that was merged
                GeoX = SearchGeoCoord(0).XGeo
                GeoY = SearchGeoCoord(1).YGeo
                GoSub GeotoCoord
                X1 = CurrentX * DigiZoom.LastZoom
                Y1 = CurrentY * DigiZoom.LastZoom
                GeoX = SearchGeoCoord(0).XGeo + numXsteps& * XStep
                GoSub GeotoCoord
                X2 = CurrentX * DigiZoom.LastZoom
                Y2 = CurrentY * DigiZoom.LastZoom
                GeoY = SearchGeoCoord(1).YGeo + numYsteps& * YStep
                GoSub GeotoCoord
                X3 = CurrentX * DigiZoom.LastZoom
                Y3 = CurrentY * DigiZoom.LastZoom
                GeoX = SearchGeoCoord(0).XGeo
                GoSub GeotoCoord
                X4 = CurrentX * DigiZoom.LastZoom
                Y4 = CurrentY * DigiZoom.LastZoom
                
                'stop blinking search points for 1:50000 maps
                 GDMDIform.CenterPointTimer.Enabled = False
                 ce& = 0 'reset flag that draws blinking cursor
                
                If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                   'draw box
                   GDform1.Picture2.FillStyle = 4
                   GDform1.Picture2.DrawMode = 9 '8 '5 '2 '3 '5
                   GDform1.Picture2.Line (X1, Y1)-(X3, Y3), , BF
                Else
'                    GDform1.Picture2.DrawMode = 13
                    GDform1.Picture2.DrawWidth = Max(5, CInt(DigiZoom.LastZoom))
'                    GDform1.Picture2.DrawStyle = vbDot
                    GDform1.Picture2.FillStyle = 4
                    GDform1.Picture2.DrawMode = 9
'                    GDform1.Picture2.DrawStyle = vbDot
                    GDform1.Picture2.FillColor = QBColor(14)
                    GDform1.Picture2.Line (X1, Y1)-(X2, Y2)
                    GDform1.Picture2.Line (X2, Y2)-(X3, Y3)
                    GDform1.Picture2.Line (X3, Y3)-(X4, Y4)
                    GDform1.Picture2.Line (X4, Y4)-(X1, Y1)
                    End If
                
                GDform1.Picture2.FillStyle = gdfs
                GDform1.Picture2.FillColor = gdco
                GDform1.Picture2.DrawWidth = gdwi
                GDform1.Picture2.DrawStyle = gdds
                
                'renable blinking
                GDMDIform.CenterPointTimer.Enabled = True
                ce& = 1


         End Select
     
      End If
      
    'reopen baseDTM for reading heights
    If UseNewDTM% = 1 Then
       ier = OpenCloseBaseDTM(0)
       End If
     
   CreateDTM = ier
        
   Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Function
   
'========================================GOSUBS===================================================
   
ModifyJKHDTM: '-----------save the points to the JKH's DTM tile----------------------
       backup% = 0
       response = MsgBox("Backup DTM files before saving changes?" & vbLf & _
                     "(The date will be added as a suffix to the backup tiles)", _
                     vbQuestion + vbYesNoCancel, "Maps&More")
       If response = vbYes Then
          backup% = 1
          End If
               
        GDMDIform.picProgBar.Visible = True
        Call UpdateStatus(GDMDIform, 1, 0)
        GDMDIform.StatusBar1.Panels(1).Text = "Writing new values into the 25m DTM tiles, please wait...."
        
        'determine which tile(s) are being used and back them up

        CHFind$ = sEmpty
        For j = 0 To numYsteps& - 1
            kmy = SearchGeoCoord(1).YGeo + (j - 1) * YStep
            For i = 0 To numXsteps& - 1
              kmx = SearchGeoCoord(0).XGeo + (i - 1) * XStep
        
              kmxDTM = kmx * 0.001
              kmyDTM = (kmy - 1000000) * 0.001
              IKMX& = Int((kmxDTM + 20!) * 40!) + 1
              IKMY& = Int((380! - kmyDTM) * 40!) + 1
              NRow% = IKMY&: NCol% = IKMX&

              'FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
              Jg% = 1 + Int((NRow% - 2) / 800)
              Ig% = 1 + Int((NCol% - 2) / 800)
              
             'Since roundoff errors in converting from coord to
             'integer indexes, just count columns and rows assuming
             'that the first one has no roundoff error
              If kmx = SearchGeoCoord(0).XGeo And kmy = SearchGeoCoord(0).YGeo Then
                 IR% = NRow% - (Jg% - 1) * 800
                 IC% = NCol% - (Ig% - 1) * 800
                 IR0% = IR%
                 IC0% = IC%
              Else
                 IC% = CInt((kmx - SearchGeoCoord(0).XGeo) * 0.04) + IC0%
                 IR% = IR0% - CInt((kmy - SearchGeoCoord(0).YGeo) * 0.04)
                 End If
              
              IFN& = (IR% - 1) * 801! + IC%
              
              CHFindTmp$ = CHMAP(Ig%, Jg%)
tp250:        If CHFindTmp$ <> CHFind$ Then
                 newtile% = 1
                 CHFind$ = CHFindTmp$
                 
                 If backup% = 1 Then
                    'back it up if not already backed up
                    FirstTry% = 0
                    If Dir(dtmdir + "\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)) = sEmpty Then
                       FileCopy dtmdir + "\" & CHFind$, dtmdir + "\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                    Else 'warn before overwriting last backup
                       response = MsgBox("File with backup name already exists!" & vbLf & _
                              "Do you want to overwrite it?", vbExclamation + vbYesNoCancel, "Maps&More")
                       If response = vbYes Then
                          FileCopy dtmdir + "\" & CHFind$, dtmdir + "\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                       Else
                          response = InputBox("Enter the name of the backup:", _
                                     "New backup tile name", dtmdir + "\" & CHFind$ & _
                                     "_" & Month(Date) & Day(Date) & Year(Date), 6450)
                          If response = sEmpty Then
                             Exit Function
                          Else
                             If Dir(response) = sEmpty Then
                                FileCopy dtmdir + "\" & CHFind$, dtmdir + "\" & response
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
                    Open dtmdir + "\" & CHFind$ For Random As #filn% Len = 2
                    
                 End If
                 
               hgtNews = Zcoord(i, j)
             
               'write the changes to the DTM tile
               'Since roundoff errors in converting from coord to
               'integer indexes, just count columns and rows assuming
               'that the first one has no roundoff error
                If (kmx = SearchGeoCoord(0).XGeo And kmy = SearchGeoCoord(1).YGeo) Or newtile% = 1 Then
                   IR% = NRow% - (Jg% - 1) * 800
                   IC% = NCol% - (Ig% - 1) * 800
                   IR0% = IR%
                   IC0% = IC%
                   newtile% = 0
                Else
                   IC% = CInt((kmx - SearchGeoCoord(0).XGeo) * 0.04) + IC0%
                   IR% = IR0% - CInt((kmy - SearchGeoCoord(1).YGeo) * 0.04)
                   End If
              
                IFN& = (IR% - 1) * 801! + IC%
                Put #filn%, IFN&, CInt(hgtNews * 10)
              
           Next i
           
           Call UpdateStatus(GDMDIform, 1, Int(100 * j / (numYsteps& - 1)))
           
        Next j
          
       Close #filn%
       CHMNEO = sEmpty
       
       Call UpdateStatus(GDMDIform, 1, 0)
       GDMDIform.picProgBar.Visible = False
       GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
          
Return

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
GeotoCoord:
    
    CurrentX = ((GeoX - ULGeoX) * GeoToPixelX) + ULPixX
    CurrentY = ((ULGeoY - GeoY) * GeoToPixelY) + ULPixY
    
    If RSMethod1 Or RSMethod2 Then
       
       If RSMethod1 Then
          ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
       ElseIf RSMethod2 Then
          ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
          End If
          
        Dim DifX As Double, DifY As Double
        DifX = Abs(GeoX - XGeo)
        DifY = Abs(GeoY - YGeo)
       
        ShiftX = CurrentX - (((XGeo - ULGeoX) * GeoToPixelX) + ULPixX)
        ShiftY = CurrentY - (((ULGeoY - YGeo) * GeoToPixelY) + ULPixY)
        
        CurrentX = CurrentX + ShiftX
        CurrentY = CurrentY + ShiftY
        
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
         ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
           End If

        If Abs(GeoX - XGeo) > DifX Then
           CurrentX = CurrentX - ShiftX
           End If
           
        If Abs(GeoY - YGeo) > DifY Then
           CurrentY = CurrentY - ShiftY
           End If
           
'        If Abs(GeoX - XGeo) > DifX And Abs(GeoY - YGeo) > DifY Then
''        If Abs(GeoX - XGeo) > Tolerance Or Abs(GeoY - YGeo) > Tolerance Then
'                Call MsgBox("Inverse coordinate transformation unsuccessful" _
'                        & vbCrLf & "Coordinate grid rotation too large for first approx." _
'                        & vbCrLf & vbCrLf & "(Redo using a less-rotated grid as reference...)" _
'                        , vbInformation, "Create DTM Error")
'              ier = -1
'              CreateDTM = ier
'              Screen.MousePointer = vbDefault
'              GDMDIform.picProgBar.Visible = False
'              GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
'              GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
'              Exit Function
'              End If
        
   Else
        'cuurentx, currenty are the pixel coordinates
        End If
Return

'-----------------------------ModifyAster------------------------------------
ModifyAster:

       backup% = 0
       response = MsgBox("Backup ASTER tiles before saving changes?" & vbLf & _
                     "(The date will be added as a suffix to the backup tiles)", _
                     vbQuestion + vbYesNoCancel, "ASTER DEM editing")
       If response = vbYes Then
          backup% = 1
          End If

    GDMDIform.picProgBar.Visible = True
    Call UpdateStatus(GDMDIform, 1, 0)
    GDMDIform.StatusBar1.Panels(1).Text = "Writing new values into the relevant ASTER tiles, please wait...."

    For j = 0 To numYsteps& - 1
        degy = SearchGeoCoord(1).YGeo + (j - 1) * YStep
        For i = 0 To numXsteps& - 1
          degx = SearchGeoCoord(0).XGeo + (i - 1) * XStep

            If ASTERbilOpen Then
               If degx >= ASTEREast And degx <= ASTEREast + 1 Then
                  If degy >= ASTERNorth And degy <= ASTERNorth + 1 Then
                     'the height resides on the same tile
                     GoTo g10
                     End If
                  End If
               End If
    
        '-----------------------------------------------
             'Else need to open new tile
             'determine root name from bottom left coordinates
             If ASTERfilnum% <> 0 Then Close #ASTERfilnum% 'close last file and open new one
             ASTERbilOpen = False
             
             ASTERNorth = Int(degy) '-35.3 -> -36, 35.3 -> 35
             If lt >= 0 Then 'North latitude
                ltch$ = "N"
             ElseIf lt < 0 Then 'South latitude
                ltch$ = "S"
                End If
             ASTEREast = Int(degx) '35.3 ->35, -118.3 -> -119
             If degx >= 0 Then 'east longitue
                lgch$ = "E"
             ElseIf degx < 0 Then 'West longitude
                lgch$ = "W"
                End If
             
             ASTERfilename = ltch$ & Format(Trim$(str$(Abs(ASTERNorth))), "00") & lgch$ & Format(Trim$(str$(Abs(ASTEREast))), "000")
           
             'If ASTERfilnum% <> 0 Then Close #ASTERfilnum% 'close last file and open new one
             'ASTERbilOpen = False
              
             filin1% = FreeFile
             If Dir(ASTERdir & "\" & ASTERfilename & ".bil") <> sEmpty And Dir(ASTERdir & "\" & ASTERfilename & ".hdr") <> sEmpty Then
                filin1% = FreeFile
                Open ASTERdir + "\" + ASTERfilename + ".hdr" For Input As #filin1%
                ASTERbilOpen = True
                Do Until EOF(filin1%)
                   
                   Line Input #filin1%, doclin$
                   doclin$ = Trim$(doclin$)
                   If InStr(doclin$, "NROWS") <> 0 Then
                      pos1% = InStr(doclin$, "NROWS")
                      ASTERNrows% = val(Mid$(doclin$, pos1% + 5, Len(doclin$)))
                   ElseIf InStr(doclin$, "NCOLS") <> 0 Then
                      pos2% = InStr(doclin$, "NCOLS")
                      ASTERNcols% = val(Mid$(doclin$, pos2% + 5, Len(doclin$)))
                   ElseIf InStr(doclin$, "XDIM") <> 0 Then
                      pos3% = InStr(doclin$, "XDIM")
                      ASTERxdim = val(Mid$(doclin$, pos3% + 4, Len(doclin$)))
                   ElseIf InStr(doclin$, "YDIM") <> 0 Then
                      pos4% = InStr(doclin$, "YDIM")
                      ASTERydim = val(Mid$(doclin$, pos4% + 4, Len(doclin$)))
                      Exit Do 'found everything
                      End If
                      
                Loop
                Close #filin1%
                
                'backup the tile if flagged
g5:
                 If backup% = 1 Then
                    'back it up if not already backed up
                    FirstTry% = 0
                    If Dir(ASTERdir + "\" + ASTERfilename & "_" & Month(Date) & Day(Date) & Year(Date) & ".bil") = sEmpty Then
                       FileCopy ASTERdir + "\" + ASTERfilename + ".bil", ASTERdir + "\" + ASTERfilename & "_" & Month(Date) & Day(Date) & Year(Date) & ".bil"
                    Else 'warn before overwriting last backup
                       response = MsgBox("File with backup name already exists!" & vbLf & _
                              "Do you want to overwrite it?", vbExclamation + vbYesNoCancel, "Overwrite ASTER tile")
                       If response = vbYes Then
                          FileCopy ASTERdir + "\" + ASTERfilename + ".bil", ASTERdir + "\" + ASTERfilename & "_" & Month(Date) & Day(Date) & Year(Date) & ".bil"
                       Else
                          response = InputBox("Enter the suffix (usually a date) of the backup only (e.g., 102015):", _
                                     "New backup tile suffix", _
                                     "_" & Month(Date) & Day(Date) & Year(Date), 6450)
                          If response = sEmpty Then
                             Return
                          Else
                             If Dir(response) = sEmpty Then
                                FileCopy ASTERdir + "\" + ASTERfilename + ".bil", ASTERdir + "\" + ASTERfilename & response & ".bil"
                             Else
                                GoTo g5
                                End If
                             End If
                          End If
                       End If
                    End If
                
                'now open the new bil file
                ASTERfilnum% = FreeFile
                Open ASTERdir + "\" + ASTERfilename + ".bil" For Binary As #ASTERfilnum%
                ASTERbilOpen = True
                   
              Else
                Select Case MsgBox("The following ASTER tile is missing:" _
                                   & vbCrLf & ASTERfilename _
                                   & vbCrLf & "" _
                                   & vbCrLf & "Abort the edit?" _
                                   , vbYesNo Or vbExclamation Or vbDefaultButton1, "ASTER tile missing")
                
                    Case vbYes
                       Return
                    Case vbNo
                       GoTo g20
                End Select
                Return
                End If
            
g10:
            'write a height
            hgtNews = Zcoord(i, j)
            
            ASTERIKMY% = CInt((degy - ASTERNorth) / ASTERydim) + 1
            ASTERIKMX% = CInt((degx - ASTEREast) / ASTERxdim) + 1
            tncols& = ASTERNcols%
            tnrows& = ASTERNrows%
            numrec& = (tnrows& - ASTERIKMY%) * tncols& + ASTERIKMX%
            Put #ASTERfilnum%, (numrec& - 1) * 2 + 1, hgtNews
g20:
        Next i
        
        Call UpdateStatus(GDMDIform, 1, Int(100 * j / (numYsteps& - 1)))
        
  Next j
  
    Call UpdateStatus(GDMDIform, 1, 0)
    GDMDIform.picProgBar.Visible = False
    GDMDIform.StatusBar1.Panels(1).Text = gsEmpty

Return

'----------------------------ModifySRTM-----------------------------------------
ModifySRTM:

    backup% = 0
    response = MsgBox("Backup ASTER tiles before saving changes?" & vbLf & _
                  "(The date will be added as a suffix to the backup tiles)", _
                  vbQuestion + vbYesNoCancel, "ASTER DEM editing")
    If response = vbYes Then
       backup% = 1
       End If

    GDMDIform.picProgBar.Visible = True
    Call UpdateStatus(GDMDIform, 1, 0)
    GDMDIform.StatusBar1.Panels(1).Text = "Writing new values into the relevant ASTER tiles, please wait...."
    
    Dim EWch$, NSch$, lgsrtm As Integer, ltsrtm As Integer, lg1ch$, lt1ch$

    For j = 0 To numYsteps& - 1
        degy = SearchGeoCoord(1).YGeo + (j - 1) * YStep
        For i = 0 To numXsteps& - 1
          degx = SearchGeoCoord(0).XGeo + (i - 1) * XStep
          
            If SRTMfileOpen Then
               If degx >= SRTMEast And degx <= SRTMEast + 1 Then
                  If degy >= SRTMNorth And degy <= SRTMNorth + 1 Then
                     'the height resides on the same tile
                     GoTo g50
                     End If
                  End If
               End If
               
        'Else need to open new tile
        'determine root name from bottom left coordinates
        If SRTMfilnum% <> 0 Then Close #SRTMfilnum% 'close last file and open new one
        SRTMfileOpen = False
          
        'determine tile name
        lgsrtm = Int(degx)
        If lgsrtm < 0 And lgsrtm > degx Then lgsrtm = lgsrtm - 1
        If lgsrtm < 0 Then EWch$ = "W" Else EWch$ = "E"
        If Abs(lgsrtm) < 10 Then
           lg1ch$ = "00" & Trim$(str$(Abs(lgsrtm)))
        ElseIf Abs(lgsrtm) >= 10 And Abs(lgsrtm) < 100 Then
           lg1ch$ = "0" & Trim$(str$(Abs(lgsrtm)))
        ElseIf Abs(lgsrtm) >= 100 Then
           lg1ch$ = Trim$(str$(Abs(lgsrtm)))
           End If
        ltsrtm = Int(degy) 'SRTM tiles are named by SW corner
        If ltsrtm < 0 And ltsrtm > lt Then ltsrtm = ltsrtm - 1
        
        SRTMEast = lgsrtm
        SRTMNorth = ltsrtm
        
        If ltsrtm < 0 Then NSch$ = "S" Else NSch$ = "N"
        If Abs(ltsrtm) < 10 Then
           lt1ch$ = "0" & Trim$(str$(Abs(ltsrtm)))
        ElseIf Abs(ltsrtm) >= 10 Then
           lt1ch$ = Trim$(str$(Abs(ltsrtm)))
           End If
        ltsrtm = ltsrtm + 1 'the first record in SRTM tiles in the NW corner
        SRTMfil$ = NEDdir & "\" & NSch$ & lt1ch$ & EWch$ & lg1ch$
        If Dir(SRTMfil$ & ".hgt") = gsEmpty Then
           Select Case MsgBox("The following SRTM tile is missing:" _
                              & vbCrLf & SRTMfil$ & ".hgt" _
                              & vbCrLf & "" _
                              & vbCrLf & "Abort the edit?" _
                              , vbYesNo Or vbExclamation Or vbDefaultButton1, "SRTM tile missing")
           
            Case vbYes
               Return
            Case vbNo
               GoTo g100
           End Select
           End If
           
g40:
            'backup the tile if flagged
             If backup% = 1 Then
                'back it up if not already backed up
                FirstTry% = 0
                If Dir(NEDdir + "\" + SRTMfil$ & "_" & Month(Date) & Day(Date) & Year(Date) & ".bil") = sEmpty Then
                   FileCopy NEDdir + "\" + SRTMfil$ & ".hgt", NEDdir + "\" + SRTMfil$ & "_" & Month(Date) & Day(Date) & Year(Date) & ".hgt"
                Else 'warn before overwriting last backup
                   response = MsgBox("File with backup name already exists!" & vbLf & _
                          "Do you want to overwrite it?", vbExclamation + vbYesNoCancel, "Overwrite SRTM tile")
                   If response = vbYes Then
                      FileCopy NEDdir + "\" + SRTMfil$ + ".hgt", NEDdir + "\" + SRTMfil$ & "_" & Month(Date) & Day(Date) & Year(Date) & ".hgt"
                   Else
                      response = InputBox("Enter the suffix (usually a date) of the backup only (e.g., 102015):", _
                                 "New backup tile suffix", _
                                 "_" & Month(Date) & Day(Date) & Year(Date), 6450)
                      If response = sEmpty Then
                         Return
                      Else
                         If Dir(response) = sEmpty Then
                            FileCopy NEDdir + "\" + SRTMfil$ + ".hgt", NEDdir + "\" + SRTMfil$ & response & ".hgt"
                         Else
                            GoTo g40
                            End If
                         End If
                      End If
                   End If
                End If
           
        SRTMfilnum% = FreeFile
        Open SRTMfil$ & ".hgt" For Binary As #SRTMfilnum%
        SRTMfileOpen = True
       
g50:
        'write a height, first invert into BIG Endian
        hgtNews = Zcoord(i, j)
        
        If hgtNews <> 0 Then
            '===================InvertBytes==================================
            A0$ = LTrim$(RTrim$(Hex$(hgtNews)))
            AA$ = ""
            'swap the two bytes using their hex representation
            'e.g., ABCD --> CDAB, etc.
            If Len(A0$) = 4 Then
               A1$ = Mid$(A0$, 1, 2)
               A2$ = Mid$(A0$, 3, 2)
               If Mid$(A2$, 3, 1) = "0" And Mid$(A2$, 3, 2) <> "0" Then
                  A2$ = Mid$(A0$, 4, 1)
               ElseIf Mid$(A2$, 3, 1) = "0" And Mid$(A2$, 3, 2) = "0" Then
                  A2$ = ""
                  End If
               AA$ = A2$ + A1$
            ElseIf Len(A0$) = 3 Then
               A1$ = "0" + Mid$(A0$, 1, 1)
               A2$ = Mid$(A0$, 2, 2)
               If Mid$(A0$, 2, 1) = "0" Then A2$ = Mid$(A0$, 3, 1)
               AA$ = A2$ + A1$
            ElseIf Len(A0$) = 2 Or Len(A0$) = 1 Then
               A1$ = "00"
               A2$ = A0$
               AA$ = A2$ + A1$
               End If
            
             'convert swaped hexadecimel to an integer value
             leng% = Len(LTrim$(RTrim$(AA$)))
             integ1& = 0
             For jj% = leng% To 1 Step -1
                 V$ = Mid$(LTrim$(RTrim$(AA$)), jj%, 1)
                 If InStr("ABCDEF", V$) <> 0 Then
                    If V$ = "A" Then
                       NO& = 10
                    ElseIf V$ = "B" Then
                       NO& = 11
                    ElseIf V$ = "C" Then
                       NO& = 12
                    ElseIf V$ = "D" Then
                       NO& = 13
                    ElseIf V$ = "E" Then
                       NO& = 14
                    ElseIf V$ = "F" Then
                       NO& = 15
                       End If
                 Else
                    NO& = val(V$)
                   End If
                If jj% = leng% - 3 Then
                   integ1& = integ1& + 4096 * NO&
                ElseIf jj% = leng% - 2 Then
                   integ1& = integ1& + 256 * NO&
                ElseIf jj% = leng% - 1 Then
                   integ1& = integ1& + 16 * NO&
                ElseIf jj% = leng% Then
                   integ1& = integ1& + NO&
                   End If
             Next jj%
             'positive 2 byte integers are stored as numbers 1 to 32767.
             'negative 2 byte integers are stored as numbers
             'greater than 32767 (since 2 byte, i.e.,  8 bits encompass
             'the integer range -32768 to 32767), where -1 is 65535 and
             '-2 is 65534, etc up to -32768 which is represented
             'as 32768, i.e.,
             If integ1& > 32767 Then integ1& = integ1& - 65536
             hgtNews = integ1&
             End If
        
        'determine record number of height in SRTM file
        SRTMIKMY% = CInt(((ltsrtm + 1!) - degy) / YStep) + 1
        SRTMIKMX% = CInt((lgsrtm - degx) / XStep) + 1
        tncols& = 3601
        numrec& = (SRTMIKMY% - 1) * tncols& + SRTMIKMX%
        'replace the SRTM height with the NED height
        Put #SRTMfilnum%, (numrec& - 1) * 2 + 1, hgtNews

g100:
        Next i
        
        Call UpdateStatus(GDMDIform, 1, Int(100 * j / (numYsteps& - 1)))
        
  Next j
  
    Call UpdateStatus(GDMDIform, 1, 0)
    GDMDIform.picProgBar.Visible = False
    GDMDIform.StatusBar1.Panels(1).Text = gsEmpty

  
Return

'========================================end GoSubs==============================================================

CreateDTM_Error:
    GDMDIform.picProgBar.Visible = False
    Screen.MousePointer = vbDefault

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateDTM of Module modHardy"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Calculate_Inverse
' Author    : Dr-John-K-Hall
' Date      : 8/30/2015
' Purpose   : Use Gauss elimination method for calculating the inverse matrix
' source    : http://www.freevbcode.com/ShowCode.asp?ID=6221
'---------------------------------------------------------------------------------------
'
Private Function Calculate_Inverse(EscapeKeyPressed As Boolean) As Integer
'Uses Gauss elimination method in order to calculate the inverse matrix [A]-1
'Method: Puts matrix [A] at the left and the singular matrix [I] at the right:
'[ a11 a12 a13 | 1 0 0 ]
'[ a21 a22 a23 | 0 1 0 ]
'[ a31 a32 a33 | 0 0 1 ]
'Then using line operations, we try to build the singular matrix [I] at the left.
'After we have finished, the inverse matrix [A]-1 (bij) will be at the right:
'[ 1 0 0 | b11 b12 b13 ]
'[ 0 1 0 | b21 b22 b23 ]
'[ 0 0 1 | b31 b32 b33 ]

Dim ier As Integer

ier = 0

On Error GoTo ErrHandler 'In case the inverse cannot be found (Determinant = 0)

Solution_Problem = False

'Assign values from matrix [A] at the left
Call UpdateStatus(GDMDIform, 1, 0)
For N = 1 To System_DIM
    For m = 1 To System_DIM
        Operations_Matrix(m, N) = Matrix_A(m, N)
    Next
    DoEvents
    Call UpdateStatus(GDMDIform, 1, 100 * N / System_DIM)
Next

'Assign values from singular matrix [I] at the right
Call UpdateStatus(GDMDIform, 1, 0)
For N = 1 To System_DIM
    For m = 1 To System_DIM
        If N = m Then
            Operations_Matrix(m, N + System_DIM) = 1
        Else
            Operations_Matrix(m, N + System_DIM) = 0
        End If
    Next
    DoEvents
    Call UpdateStatus(GDMDIform, 1, 100 * N / System_DIM)
Next

'Build the Singular matrix [I] at the left
Call UpdateStatus(GDMDIform, 1, 0)
For K = 1 To System_DIM
   'Bring a non-zero element first by changes lines if necessary
   If Operations_Matrix(K, K) = 0 Then
      For N = K To System_DIM
        If Operations_Matrix(N, K) <> 0 Then line_1 = N: Exit For 'Finds line_1 with non-zero element
      Next N
      'Change line k with line_1
      For m = K To System_DIM * 2
         temporary_1 = Operations_Matrix(K, m)
         Operations_Matrix(K, m) = Operations_Matrix(line_1, m)
         Operations_Matrix(line_1, m) = temporary_1
      Next m
   End If
   
    elem1 = Operations_Matrix(K, K)
   For N = K To 2 * System_DIM
    Operations_Matrix(K, N) = Operations_Matrix(K, N) / elem1
   Next N
   
   'For other lines, make a zero element by using:
   'Ai1=Aij-A11*(Aij/A11)
   'and change all the line using the same formula for other elements
   For N = 1 To System_DIM
        If N = K And N = System_DIM Then Exit For 'Finished
        If N = K And N < System_DIM Then N = N + 1 'Do not change that element (already equals to 1), go for next one
      If Operations_Matrix(N, K) <> 0 Then 'if it is zero, stays as it is
         multiplier_1 = Operations_Matrix(N, K) / Operations_Matrix(K, K)
         For m = K To 2 * System_DIM
            Operations_Matrix(N, m) = Operations_Matrix(N, m) - Operations_Matrix(K, m) * multiplier_1
         Next m
      End If
   Next N
   
   DoEvents
   Call UpdateStatus(GDMDIform, 1, 100 * K / System_DIM)
   
    '---------------------break on ESC key-------------------------------
    If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
       EscapeKeyPressed = True
       ier = -1
       Calculate_Inverse = ier
       Exit Function
       End If
    '---------------------------------------------------------------------

Next K

'Assign the right part to the Inverse_Matrix
Call UpdateStatus(GDMDIform, 1, 0)
For N = 1 To System_DIM
    For K = 1 To System_DIM
        Inverse_Matrix(N, K) = Operations_Matrix(N, System_DIM + K)
    Next K
    DoEvents
    Call UpdateStatus(GDMDIform, 1, 100 * N / System_DIM)
Next N

EscapeKeyPressed = False
Calculate_Inverse = ier

Exit Function

ErrHandler:
message$ = "An error occured during the calculation process. Determinant of Matrix [A] is probably equal to zero."
response = MsgBox(message$, vbCritical)
Solution_Problem = True
ier = -1
Calculate_Inverse = ier

End Function
'---------------------------------------------------------------------------------------
' Procedure : SolveEquations
' Author    : Dr-John-K-Hall
' Date      : 8/31/2015
' Purpose   : Solve a system of equations with Gaussian elimination
' source    : http://www.vb-helper.com/howto_gaussian_elimination.html
'---------------------------------------------------------------------------------------
'
Public Function SolveEquations(EscapeKeyPressed As Boolean, Z() As Double) As Integer
Const TINY As Double = 0.00001
Dim num_rows As Long
Dim num_cols As Long
Dim R As Long
Dim c As Long
Dim r2 As Long
Dim tmp As Double
Dim factor As Double
Dim arr() As Double
Dim Txt As String
Dim ier As Integer

ier = 0

   On Error GoTo SolveEquations_Error
   
   num_rows = System_DIM
   num_cols = System_DIM
   
   ReDim arr(1 To num_rows, 1 To num_cols + 2)

    ' Build the augmented matrix.
    Call UpdateStatus(GDMDIform, 1, 0)
    For m = 1 To num_rows
        For N = 1 To num_cols
            arr(m, N) = Matrix_A(m, N)
        Next
        arr(m, num_cols + 1) = Matrix_A(m, num_cols + 1) 'Z(m)
        DoEvents
        Call UpdateStatus(GDMDIform, 1, 100 * m / num_rows)
    Next

    ' Start solving.
    Call UpdateStatus(GDMDIform, 1, 0)
    For R = 1 To num_rows - 1
        ' Zero out all entries in column r after this row.
        ' See if this row has a non-zero entry in column r.
        If Abs(arr(R, R)) < TINY Then
            ' Not a non-zero value. Try to swap with a later row.
            For r2 = R + 1 To num_rows
                If Abs(arr(r2, R)) > TINY Then
                    ' This row will work. Swap them.
                    For c = 1 To num_cols + 1
                        tmp = arr(R, c)
                        arr(R, c) = arr(r2, c)
                        arr(r2, c) = tmp
                    Next c
                    Exit For
                End If
            Next r2
        End If

        ' If this row has a non-zero entry in column r, skip this column.
        If Abs(arr(R, R)) > TINY Then
            ' Zero out this column in later rows.
            For r2 = R + 1 To num_rows
                factor = -arr(r2, R) / arr(R, R)
                For c = R To num_cols + 1
                    arr(r2, c) = arr(r2, c) + factor * arr(R, c)
                Next c
            Next r2
        End If
        
        DoEvents
        Call UpdateStatus(GDMDIform, 1, 100 * (R / (num_rows - 1)))

    Next R


    ' See if we have a solution.
    If arr(num_rows, num_cols) = 0 Then
        ' We have no solution.
        ier = -1
        SolveEquations = ier
        Exit Function
    Else
        ' Back solve.
        Call UpdateStatus(GDMDIform, 1, 0)
        For R = num_rows To 1 Step -1
            tmp = arr(R, num_cols + 1)
            For r2 = R + 1 To num_rows
                tmp = tmp - arr(R, r2) * arr(r2, num_cols + 2)
            Next r2
            arr(R, num_cols + 2) = tmp / arr(R, R)
            C_vector(R) = arr(R, num_cols + 2)
            
            DoEvents
            Call UpdateStatus(GDMDIform, 1, 100 * ((num_rows - R + 1) / num_rows))

        Next R

    End If

   SolveEquations = ier
   
   On Error GoTo 0
   Exit Function

SolveEquations_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SolveEquations of Module modHardy"
    ier = -1
    SolveEquations = ier
End Function
'---------------------------------------------------------------------------------------
' Procedure : ShowProfioe
' Author    : Dr-John-K-Hall
' Date      : 11/9/2015
' Purpose   : Calculates view angle vs azimuth
' input:  HorizMode% = 1 for eastern
'                    = 2 for western
'---------------------------------------------------------------------------------------
'
Public Function ShowProfile(HorizMode%) As Integer

   Dim xo As Double
   Dim yo As Double
   Dim zo As Double
   Dim va() As Double
   Dim azi() As Double
   Dim Xazi() As Double
   Dim Yazi() As Double
   Dim Zazi() As Double
   Dim distazi() As Double
   Dim numazi&, HalfAziRange As Double, StepSizeAzi As Double
   
   Dim ier As Long

   On Error GoTo ShowProfile_Error
   
   numazi& = 2 * Abs(HalfAzi) / Abs(StepAzi) + 1 'number of azimuth points for viewing in horizon profile
   StepSizeAzi = StepAzi
   HalfAziRange = HalfAzi
   
   ReDim va(numazi&) As Double
   ReDim azi(numazi&) As Double
   ReDim Xazi(numazi&) As Double
   ReDim Yazi(numazi&) As Double
   ReDim Zazi(numazi&) As Double
   ReDim distazi(numazi&) As Double
   
    If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
        CoordMode% = 1
    ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
        CoordMode% = 2
        End If
        
    xo = val(GDMDIform.Text5)
    yo = val(GDMDIform.Text6)
    zo = val(GDMDIform.Text7)

    If Dir(App.Path & "\topo_coord.xyz") <> gsEmpty Then
    
        Screen.MousePointer = vbHourglass
        
'//////////////////for version that doesn't read and write files/////////////////////////////
'        GDMDIform.StatusBar1.Panels(1).Text = "Collecting the data, please wait...."
'        Fname$ = App.Path & "\topo_coord.xyz"
'        ff = FreeFile
'        Open Fname$ For Binary As #ff
'        Raw$ = String$(LOF(ff), 32)
'        Get #ff, 1, Raw$
'        Close #ff
'        Dim Txt$()
'        Txt$ = Split(Raw$, vbCrLf)
'        num_array& = UBound(Txt$)
'
'        'check if g_np = sqr(num_array&)
'        If g_np <> Sqr(num_array&) Then
'           Call MsgBox("Topo_coord.txt file is truncated..." _
'                       & vbCrLf & "" _
'                       & vbCrLf & "Aborting...." _
'                       , vbExclamation, "Data Error")
'           ier = -1
'           ShowProfile = ier
'           Exit Function
'           End If
'
'        Dim xx() As Double
'        Dim yy() As Double
'        Dim zz() As Double
'
'        GDMDIform.picProgBar.Visible = True
'        Call UpdateStatus(GDMDIform, 1, 0)
'
'        For i = LBound(Txt$) To num_array&
'           Dim Coords$()
'           Coords$ = Split(Txt$(i), ",")
'           If UBound(Coords$) = 2 Then
'              ReDim Preserve xx(i) As Double
'              ReDim Preserve yy(i) As Double
'              ReDim Preserve zz(i) As Double
'              xx(i) = Coords$(0)
'              yy(i) = Coords$(1)
'              zz(i) = Coords$(2)
'              End If
'
'          Call UpdateStatus(GDMDIform, 1, 100 * i / (num_array& + LBound(Txt$) - 1))
'
'        Next i
'//////////////////////////////////////////////////////////////////////////////
        
'        Call SetPriorityClass(GetCurrentProcess, HIGH_PRIORITY_CLASS)
'        Call SetThreadPriority(GetCurrentThread, THREAD_BASE_PRIORITY_MAX)
        
        Call UpdateStatus(GDMDIform, 1, 0)
        GDMDIform.StatusBar1.Panels(1).Text = "Calculating profile, please wait..."
        
        'shut down blinkers
        ce& = 0 'reset blinker flag
        If GDMDIform.CenterPointTimer.Enabled = True Then
           ce& = 1 'flag that timer has been shut down during drag
           GDMDIform.CenterPointTimer.Enabled = False
           End If
        
        Dim nrows As Long, ncols As Long
        nrows = g_nrows
        ncols = g_ncols
        
        Dim NearApprn As Double
        NearApprn = Apprn 'nearest approach, i.e., amount to shave...
        
        'version without file reading and writing
'        ier = Profiles(xo, yo, zo, _
'                        xx(0), yy(0), zz(0), _
'                        npp&, CoordMode%, HorizMode%, _
'                        va(0), azi(0), _
'                        Xazi(0), Yazi(0), Zazi(0), distazi(0), _
'                        numazi&, StepSizeAzi, HalfAziRange, _
'                        AddressOf MyCallback)
                        
        If Not RotatedGrid Then
            'use standard profile calculation assuming data arranged as xyz files with rows or x and columns of y
            ier = Profiles(App.Path & "\topo_coord.xyz", App.Path & "\profile.txt", _
                            xo, yo, zo, _
                            nrows, ncols, CoordMode%, HorizMode%, Save_xyz%, _
                            va(0), azi(0), _
                            Xazi(0), Yazi(0), Zazi(0), distazi(0), _
                            numazi&, StepSizeAzi, HalfAziRange, NearApprn, _
                            AddressOf MyCallback)
        Else
           'calculate the profile by merging into azimuth grid without assuming xyz format, i.e., random lines of x,y,z is sufficient
            ier = Profiles2(App.Path & "\topo_coord.xyz", App.Path & "\profile.txt", _
                            xo, yo, zo, _
                            numXYZpoints, CoordMode%, HorizMode%, Save_xyz%, _
                            va(0), azi(0), _
                            Xazi(0), Yazi(0), Zazi(0), distazi(0), _
                            numazi&, StepSizeAzi, HalfAziRange, NearApprn, _
                            AddressOf MyCallback)
           
           End If
                        
        
        Call UpdateStatus(GDMDIform, 1, 0)
        GDMDIform.picProgBar.Visible = False
        GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
        
        'renable blinking
        GDMDIform.CenterPointTimer.Enabled = True
        ce& = 1
                                
'        If Not FreeLibrary(GetModuleHandle("MapDigitizer.dll")) Then
'           'freeing memory failed
'           'give message
'           End If

        'restore priority of this thread to normal
'        Call SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_NORMAL)
'        Call SetPriorityClass(GetCurrentProcess, NORMAL_PRIORITY_CLASS)
        
        Call GDMDIform.mnuProfile_Click
        
        If ier = -1 Then
            Call MsgBox("Profile generation failed." _
                        & vbCrLf & "" _
                        & vbCrLf & "Two possible caues of this error are:" _
                        & vbCrLf & "1. Writing xyz data to files in the Options dialog. is not enabled." _
                        & vbCrLf & "2. The topo_coord.xyz file can't be read due to a permission error." _
                        , vbExclamation, "Error in profile generation")
        ElseIf ier = -2 Then
           Call MsgBox("Your search area is too small." _
                       & vbCrLf & "" _
                       & vbCrLf & "Try again by expanding the selected area." _
                       , vbInformation, "Profile Error")
           
        Else 'calculation completed successfully, so plot the horizon profile
        
mm500:      mapgraphfm.Visible = True
            Dim Ret As Long
            Ret = SetWindowPos(mapgraphfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            Do While MapGraphVis = True
                DoEvents
            Loop
            
            End If
        
        Screen.MousePointer = vbDefault
        
    Else
    
       ier = -1
       
       GDMDIform.picProgBar.Visible = False
       ShowProfile = ier
        
       End If

   On Error GoTo 0
   Exit Function

ShowProfile_Error:
 
    GDMDIform.picProgBar.Visible = False
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowProfile of Module modHardy"
        
End Function


'---------------------------------------------------------------------------------------
' Procedure : MyCallback
' Author    : Dr-John-K-Hall
' Date      : 9/11/2015
' Purpose   : Callback function from dll that does the Gaussian Elimination used to move the progrress bar
'---------------------------------------------------------------------------------------
'
Private Sub MyCallback(ByVal parm As Long)

   On Error GoTo MyCallback_Error

   Call UpdateStatus(GDMDIform, 1, parm)
   
   DoEvents

   On Error GoTo 0
   Exit Sub

MyCallback_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyCallback of Module modHardy"
End Sub
