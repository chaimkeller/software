Attribute VB_Name = "modContours"
'Public xmin As Double, xmax As Double
'Public ymin As Double, ymax As Double
'Public zmin As Double, zmax As Double
'Public cpt() As Integer 'gmt rainbow palette


'---------------------------------------------------------------------------------------
' Procedure : conrec
' Author    : Dr-John-K-Hall
' Date      : 2/22/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function conrec(Pic As PictureBox, Z() As Double, Zf() As Single, Zs() As Integer, _
                         x() As Double, Y() As Double, ByVal nc As Integer, _
                         contour() As Double, ByVal ilb As Long, ByVal iub As Long, _
                         ByVal jlb As Long, ByVal jub As Long, _
                         xmin, ymin, xmax, ymax, mode%) As Integer
'=============================================================================
'     CONREC is a contouring subroutine for rectangularily spaced data.
'
'     It emits calls to a line drawing subroutine supplied by the user
'     which draws a contour map corresponding to real*4 (double) data
'     on a randomly spaced rectangular grid.
'     The coordinates emitted are in the same units given in the X() and Y() arrays.
'     Any number of contour levels may be specified but they must be
'     in order of increasing value.
'
'     adapted from the fortran-77 routine CONREC.F developed by Paul D. Bourke
'=============================================================================

' Z(#,#)          ' matrix of data to contour
' ilb,iub,jlb,jub ' index bounds of data matrix (x-lower,x-upper,y-lower,y-upper)
' X(#)            ' data matrix column coordinates
' Y(#)            ' data matrix row coordinates
' xmin,xmax,ymin,ymax are the bounding pixels of the picturebox Pic that the contours are drawn on
' nc              ' number of contour levels
' contour(#)      ' contour levels in increasing order

'additions: ENK 081615: mode% = 0 don't filter area
'                       mode% = 1 filter area

Dim m1, m2, m3, case_value As Integer
Dim dmin, dmax As Double
Dim X1, X2, Y1, Y2 As Double
Dim i, j, K, m As Integer
Dim h(5) As Double
Dim sh(5) As Integer
Dim xh(5), yh(5) As Double
Dim color As Long
Dim ier As Integer
Dim DrawingWidth As Long
'Dim cpt() As Integer

Dim Nullcode As Double

On Error GoTo conrec_Error

Pic.DrawMode = 13

DrawingWidth = Max(2, CInt(DigiZoom.LastZoom))

Nullcode = 1E+37

Dim im(4), jm(4) As Integer
im(0) = 0
im(1) = 1
im(2) = 1
im(3) = 0
jm(0) = 0
jm(1) = 0
jm(2) = 1
jm(3) = 1

Dim castab(3, 3, 3) As Integer
castab(0, 0, 0) = 0
castab(0, 0, 1) = 0
castab(0, 0, 2) = 8 '
castab(0, 1, 0) = 0
castab(0, 1, 1) = 2
castab(0, 1, 2) = 5 '
castab(0, 2, 0) = 7
castab(0, 2, 1) = 6
castab(0, 2, 2) = 9 '
castab(1, 0, 0) = 0
castab(1, 0, 1) = 3
castab(1, 0, 2) = 4 '
castab(1, 1, 0) = 1
castab(1, 1, 1) = 3
castab(1, 1, 2) = 1 '
castab(1, 2, 0) = 4
castab(1, 2, 1) = 3
castab(1, 2, 2) = 0 '
castab(2, 0, 0) = 9
castab(2, 0, 1) = 6
castab(2, 0, 2) = 7 '
castab(2, 1, 0) = 5
castab(2, 1, 1) = 2
castab(2, 1, 2) = 0 '
castab(2, 2, 0) = 8
castab(2, 2, 1) = 0
castab(2, 2, 2) = 0 '


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

   If numcpt = 0 Then
        
        myfile = Dir(App.Path & "\rainbow.cpt")
        If myfile = sEmpty Then
           ier = -1
           conrec = ier
           Exit Function
           End If
        
        '-----------------------load color palette--------------------------
        numpercent = -1
        numloop% = 0
        nowread = True
        num% = 0
   
        ier = 0
        
        ReDim cpt(3, 0)
        
        filenum% = FreeFile
        Open App.Path & "\rainbow.cpt" For Input As #filenum%
        
        Do Until EOF(filenum%)
           Line Input #filenum%, doclin$
           colorattributes = Split(doclin$, " ")
           For i = 0 To 10
             cc$ = colorattributes(i)
             If Trim$(cc$) <> vbNullString Then
                If numloop% = 0 Then
                   If val(cc$) >= numpercent Then
                       num% = val(cc$)
                       
                       If num% - 1 > UBound(cpt, 2) Then
                          ReDim Preserve cpt(3, UBound(cpt, 2) + 1)
                          End If
                       
                       cpt(0, num% - 1) = val(cc$)
                       numloop% = 1
                       numpercent = val(cc$)
                       nowread = True
                   Else
                       nowread = False
                       End If
                ElseIf numloop% = 1 Then
                   If nowread Then cpt(1, num% - 1) = val(cc$)
                   numloop% = 2
                ElseIf numloop% = 2 Then
                   If nowread Then cpt(2, num% - 1) = val(cc$)
                   numloop% = 3
                ElseIf numloop% = 3 Then
                   If nowread Then cpt(3, num% - 1) = val(cc$)
                   numloop% = 0
                   nowread = False
                   Exit For
                   End If
                End If
           Next i
        Loop
        Close #filenum%
        End If

Call UpdateStatus(GDMDIform, 1, 0)
GDMDIform.StatusBar1.Panels(1).Text = "Drawing countours, please wait...."

initporgress% = 0

If nc <> 0 Then
For j = jub - 1 To jlb Step -1
  For i = ilb To iub - 1
       Dim temp1, temp2 As Double
       
       If HeightPrecision = 0 Then
            temp1 = min(Zs(i, j), Zs(i, j + 1))
            temp2 = min(Zs(i + 1, j), Zs(i + 1, j + 1))
            dmin = min(temp1, temp2)
            temp1 = Max(Zs(i, j), Zs(i, j + 1))
            temp2 = Max(Zs(i + 1, j), Zs(i + 1, j + 1))
            dmax = Max(temp1, temp2)
       ElseIf HeightPrecision = 1 Then
            temp1 = min(Zf(i, j), Zf(i, j + 1))
            temp2 = min(Zf(i + 1, j), Zf(i + 1, j + 1))
            dmin = min(temp1, temp2)
            temp1 = Max(Zf(i, j), Zf(i, j + 1))
            temp2 = Max(Zf(i + 1, j), Zf(i + 1, j + 1))
            dmax = Max(temp1, temp2)
       ElseIf HeightPrecision = 2 Then
            temp1 = min(Z(i, j), Z(i, j + 1))
            temp2 = min(Z(i + 1, j), Z(i + 1, j + 1))
            dmin = min(temp1, temp2)
            temp1 = Max(Z(i, j), Z(i, j + 1))
            temp2 = Max(Z(i + 1, j), Z(i + 1, j + 1))
            dmax = Max(temp1, temp2)
            End If
      
'-------------------------------------------------------------------------
       'extra conditional added here to insure that large values are not plotted
       'if an area should not be contoured, values above nullcode should be entered in
       'the matrix Z
      
'------------------------------------------------------------------------
       If dmax >= contour(0) And dmin <= contour(nc - 1) And dmax < Nullcode Then
         For K = 0 To nc - 1
         
           If contour(K) >= dmin And contour(K) < dmax Then
             For m = 4 To 0 Step -1
               If (m > 0) Then
                 If HeightPrecision = 0 Then
                    h(m) = Zs(i + im(m - 1), j + jm(m - 1)) - contour(K)
                 ElseIf HeightPrecision = 1 Then
                    h(m) = Zf(i + im(m - 1), j + jm(m - 1)) - contour(K)
                 ElseIf HeightPrecision = 2 Then
                    h(m) = Z(i + im(m - 1), j + jm(m - 1)) - contour(K)
                    End If
                 xh(m) = x(i + im(m - 1))
                 yh(m) = Y(j + jm(m - 1))
               Else:
                 h(0) = 0.25 * (h(1) + h(2) + h(3) + h(4))
                 xh(0) = 0.5 * (x(i) + x(i + 1))
                 yh(0) = 0.5 * (Y(j) + Y(j + 1))
               End If
              If (h(m) > 0#) Then
                sh(m) = 1
              ElseIf h(m) < 0# Then
                sh(m) = -1
              Else:
                sh(m) = 0
              End If
            Next m
           
'=================================================================
            '
            ' Note: at this stage the relative heights of the corners and the
            ' centre are in the h array, and the corresponding coordinates are
            ' in the xh and yh arrays. The centre of the box is indexed by 0
            ' and the 4 corners by 1 to 4 as shown below.
            ' Each triangle is then indexed by the parameter m, and the 3
            ' vertices of each triangle are indexed by parameters m1,m2,and m3.
            ' It is assumed that the centre of the box is always vertex 2
            ' though this isimportant only when all 3 vertices lie exactly on
            ' the same contour level, in which case only the side of the box
            ' is drawn.
            '
            '
            '      vertex 4 +-------------------+ vertex 3
            '               | \               / |
            '               |   \    m-3    /   |
            '               |     \       /     |
            '               |       \   /       |
            '               |  m=2    X   m=2   |       the centre is vertex 0
            '               |       /   \       |
            '               |     /       \     |
            '               |   /    m=1    \   |
            '               | /               \ |
            '      vertex 1 +-------------------+ vertex 2
            '
            '
            '
            '               Scan each triangle in the box
            '
           
'=================================================================
             For m = 1 To 4
               m1 = m
               m2 = 0
               If (m <> 4) Then
                 m3 = m + 1
               Else:
                 m3 = 1
               End If
               case_value = castab(sh(m1) + 1, sh(m2) + 1, sh(m3) + 1)
               If case_value <> 0 Then
                 Select Case case_value
                 
'===========================================================
                  '     Case 1 - Line between vertices 1 and 2
                 
'===========================================================
                Case 1
                   X1 = xh(m1)
                   Y1 = yh(m1)
                   X2 = xh(m2)
                   Y2 = yh(m2)
                 
'===========================================================
                  '     Case 2 - Line between vertices 2 and 3
                 
'===========================================================
                 Case 2
                   X1 = xh(m2)
                   Y1 = yh(m2)
                   X2 = xh(m3)
                   Y2 = yh(m3)
                 
'===========================================================
                  '     Case 3 - Line between vertices 3 and 1
                 
'===========================================================
                 Case 3
                   X1 = xh(m3)
                   Y1 = yh(m3)
                   X2 = xh(m1)
                   Y2 = yh(m1)
                 
'===========================================================
                  '     Case 4 - Line between vertex 1 and side 2-3
                 
'===========================================================
                 Case 4
                   X1 = xh(m1)
                   Y1 = yh(m1)
                   X2 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   Y2 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   
                 
'===========================================================
                  '     Case 5 - Line between vertex 2 and side 3-1
                 
'===========================================================
                 Case 5
                   X1 = xh(m2)
                   Y1 = yh(m2)
                   X2 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   Y2 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                 
'===========================================================
                  '     Case 6 - Line between vertex 3 and side 1-2
                 
'===========================================================
                 Case 6
                   X1 = xh(m3)
                   Y1 = yh(m3)
                   X2 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   Y2 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                 
'===========================================================
                  '     Case 7 - Line between sides 1-2 and 2-3
                 
'===========================================================
                 Case 7
                   X1 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   Y1 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                   X2 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   Y2 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   
                 
'===========================================================
                  '     Case 8 - Line between sides 2-3 and 3-1
                 
'===========================================================
                 Case 8
                   X1 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   Y1 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   X2 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   Y2 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                 
'===========================================================
                  '     Case 9 - Line between sides 3-1 and 1-2
                 
'===========================================================
                 Case 9
                   X1 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   Y1 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                   X2 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   Y2 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                End Select
        '--------------------------------------------------------------
                'this is where the program specific drawing routine comes in.
                'This specific command will work well for a properly dimensioned
                'vb picture box or vb form (where "ContoursFrm" is the name of the form)
               
'-------------------------------------------------------------------
'                'make it fit into form size
'                scale_x = Pic.ScaleWidth / (xmax - xmin)
'                scale_y = Pic.ScaleHeight / (ymax - ymin)
'                'color is that of contour(k)
'                colornum% = (k + 1) * (UBound(cpt, 2) + 1) / nc
'                scale_x = 1#
'                scale_y = 1#
'                color = rgb(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))
'                Pic.Line (CInt((X1 - xmin) * scale_x), CInt((ymax - Y1) * scale_y))-(CInt((X2 - xmin) * scale_x), CInt((ymax - Y2) * scale_y)), color
                'color is that of contour(k)
                colornum% = (K + 1) * (UBound(cpt, 2) + 1) / nc
                color = RGB(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))

                'only plot the contour if its within the boundary
                X1 = min(X1, xmax)
                X1 = Max(X1, xmin)
                Y1 = min(Y1, ymax)
                Y1 = Max(Y1, ymin)
                X2 = min(X2, xmax)
                X2 = Max(X2, xmin)
                Y2 = min(Y2, ymax)
                Y2 = Max(Y2, ymin)
                If X1 <> X2 And Y1 <> Y2 Or mode% = 0 Then
                
                   Pic.DrawWidth = DrawingWidth
               
                   Pic.Line (X1 * DigiZoom.LastZoom, Y1 * DigiZoom.LastZoom)-(X2 * DigiZoom.LastZoom, Y2 * DigiZoom.LastZoom), color
                    
                    'store endpoints and colors for fast replotting, e.g., when zooming in and out
                    If numContourPoints > 0 Then
                       ReDim Preserve ContourPoints(numContourPoints)
                       End If
                       
                    ContourPoints(numContourPoints).X1 = X1
                    ContourPoints(numContourPoints).Y1 = Y1
                    ContourPoints(numContourPoints).X2 = X2
                    ContourPoints(numContourPoints).Y2 = Y2
                    ContourPoints(numContourPoints).color = color
                    numContourPoints = numContourPoints + 1
                    
                    End If
'-------------------------------------------------------------------
               End If
             Next m
          End If
        Next K
      End If
     Next i
  
'--------------------------------------------------------------------------------------
   'used to refresh the drawing surface after each row is contoured (for impatient users)
'   Pic.Refresh
   DoEvents
   newprogress& = Int(100 * ((jub - j) / (jub - jlb)))
   Call UpdateStatus(GDMDIform, 1, newprogress&)
   DoEvents
'-------------------------------------------------------------------------------------
   Next j
End If

GDMDIform.StatusBar1.Panels(1).Text = sEmpty
GDMDIform.StatusBar1.Panels(2).Text = sEmpty
Call UpdateStatus(GDMDIform, 1, 0)
GDMDIform.picProgBar.Visible = False

''-------------------diagnostics------------------------------
'filnum% = FreeFile
'Open App.Path & "\contour-out.txt" For Output As #filnum%
'Write #filnum%, "inputed boundaries: ", xmin, xmax, ymin, ymax
'Write #filnum%, "calculated boundaries: ", XMin1, XMax1, YMin1, YMax1
'Close #filnum%
''------------------------------------------------------------

'ContoursFrm.Caption = "Contours"
conrec = ier

   On Error GoTo 0
   Exit Function

conrec_Error:

''-------------------diagnostics------------------------------
'    MsgBox "Error at ncols, nrows, i,j,k,m = " & iub & ", " & jub & ", " & i & ", " & j & ", " & k & ", " & m
''------------------------------------------------------------
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure conrec of Module modContours"

End Function

Public Function min(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  min = v1
Else: min = v2
End If
End Function

Public Function Max(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Max = v2
Else: Max = v1
End If
End Function
'---------------------------------------------------------------------------------------
' Procedure : ReDrawContours
' Author    : Dr-John-K-Hall
' Date      : 3/18/2015
' Purpose   : redraws contours generated from Hardy qudratic surfaces
'---------------------------------------------------------------------------------------
'
Public Function ReDrawContours(Pic As PictureBox) As Integer
   
   'redraw the contours
   Dim i As Long
   Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
   Dim iprogress&
   Dim EscapeKeyPressed As Boolean
   DigiReDrawContours = True
   
   '------------------progress bar initialization
   On Error GoTo ReDrawContours_Error
   
   If Not DigiReDrawContours Or numContourPoints = 0 Then Exit Function
   
   GDMDIform.StatusBar1.Panels(1).Text = "Plotting contours, please wait...(press ''ESC'' key to abort)..."

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
    
    iprogress& = 0
    
    Pic.DrawMode = 13
    Pic.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
    
    For i = 0 To numContourPoints - 1
        
        X1 = ContourPoints(i).X1
        Y1 = ContourPoints(i).Y1
        X2 = ContourPoints(i).X2
        Y2 = ContourPoints(i).Y2
        
        Pic.Line (X1 * DigiZoom.LastZoom, Y1 * DigiZoom.LastZoom)-(X2 * DigiZoom.LastZoom, Y2 * DigiZoom.LastZoom), ContourPoints(i).color
        
        newprogress& = CInt(100 * i / numContourPoints)
        If iprogress& <> newprogress& Then
           iprogress& = newprogress&
           Call UpdateStatus(GDMDIform, 1, iprogress&)
           End If
           
        DigiReDrawContours = True
          
'        DoEvents
'
'        '---------------------break on ESC key-------------------------------
'        If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
'           EscapeKeyPressed = True
'           Exit For
'           End If
        '---------------------------------------------------------------------
        
    
    Next i
        
    GDMDIform.StatusBar1.Panels(1).Text = sEmpty
    GDMDIform.StatusBar1.Panels(2).Text = sEmpty
    Call UpdateStatus(GDMDIform, 1, 0)
    GDMDIform.picProgBar.Visible = False
       
    DigiReDrawContours = False
    
'-------------------------------------------------------------------------------------
    ReDrawContours = 0

   On Error GoTo 0
   Exit Function

ReDrawContours_Error:

    GDMDIform.StatusBar1.Panels(1).Text = sEmpty
    GDMDIform.StatusBar1.Panels(2).Text = sEmpty
    GDMDIform.picProgBar.Visible = False
    
    ReDrawContours = -1
   
End Function
