Attribute VB_Name = "modContours"
Public Xmin As Double, Xmax As Double
Public Ymin As Double, Ymax As Double
Public Zmin As Double, Zmax As Double
Public cpt() As Integer 'gmt rainbow palette


Public Function conrec(Z() As Double, X() As Double, Y() As Double, ByVal nc As Integer, _
                         contour() As Double, ByVal ilb As Integer, ByVal iub As Integer, _
                         ByVal jlb As Integer, ByVal jub As Integer, ByVal color As Long)

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
' nc              ' number of contour levels
' contour(#)      ' contour levels in increasing order

Dim m1, m2, m3, case_value As Integer
Dim dmin, dmax As Double
Dim x1, x2, y1, y2 As Double
Dim i, j, k, m As Integer
Dim h(5) As Double
Dim sh(5) As Integer
Dim xh(5), yh(5) As Double

Dim Nullcode As Double
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

ContoursFrm.progbar.Value = 0
ContoursFrm.txtProgressX.Text = vbNullString
ContoursFrm.txtProgressY.Text = vbNullString
initporgress% = 0

If nc <> 0 Then
For j = jub - 1 To jlb Step -1
  For i = ilb To iub - 1
       Dim temp1, temp2 As Double
       temp1 = Min(Z(i, j), Z(i, j + 1))
       temp2 = Min(Z(i + 1, j), Z(i + 1, j + 1))
       dmin = Min(temp1, temp2)
       temp1 = Max(Z(i, j), Z(i, j + 1))
       temp2 = Max(Z(i + 1, j), Z(i + 1, j + 1))
       dmax = Max(temp1, temp2)
      
'-------------------------------------------------------------------------
       'extra conditional added here to insure that large values are not plotted
       'if an area should not be contoured, values above nullcode should be entered in
       'the matrix Z
      
'------------------------------------------------------------------------
       If dmax >= contour(0) And dmin <= contour(nc - 1) And dmax < Nullcode Then
         For k = 0 To nc - 1
           If contour(k) >= dmin And contour(k) < dmax Then
             For m = 4 To 0 Step -1
               If (m > 0) Then
                 h(m) = Z(i + im(m - 1), j + jm(m - 1)) - contour(k)
                 xh(m) = X(i + im(m - 1))
                 yh(m) = Y(j + jm(m - 1))
               Else:
                 h(0) = 0.25 * (h(1) + h(2) + h(3) + h(4))
                 xh(0) = 0.5 * (X(i) + X(i + 1))
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
                   x1 = xh(m1)
                   y1 = yh(m1)
                   x2 = xh(m2)
                   y2 = yh(m2)
                 
'===========================================================
                  '     Case 2 - Line between vertices 2 and 3
                 
'===========================================================
                 Case 2
                   x1 = xh(m2)
                   y1 = yh(m2)
                   x2 = xh(m3)
                   y2 = yh(m3)
                 
'===========================================================
                  '     Case 3 - Line between vertices 3 and 1
                 
'===========================================================
                 Case 3
                   x1 = xh(m3)
                   y1 = yh(m3)
                   x2 = xh(m1)
                   y2 = yh(m1)
                 
'===========================================================
                  '     Case 4 - Line between vertex 1 and side 2-3
                 
'===========================================================
                 Case 4
                   x1 = xh(m1)
                   y1 = yh(m1)
                   x2 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   y2 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   
                 
'===========================================================
                  '     Case 5 - Line between vertex 2 and side 3-1
                 
'===========================================================
                 Case 5
                   x1 = xh(m2)
                   y1 = yh(m2)
                   x2 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   y2 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                 
'===========================================================
                  '     Case 6 - Line between vertex 3 and side 1-2
                 
'===========================================================
                 Case 6
                   x1 = xh(m3)
                   y1 = yh(m3)
                   x2 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   y2 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                 
'===========================================================
                  '     Case 7 - Line between sides 1-2 and 2-3
                 
'===========================================================
                 Case 7
                   x1 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   y1 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                   x2 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   y2 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   
                 
'===========================================================
                  '     Case 8 - Line between sides 2-3 and 3-1
                 
'===========================================================
                 Case 8
                   x1 = (h(m3) * xh(m2) - h(m2) * xh(m3)) / (h(m3) - h(m2))
                   y1 = (h(m3) * yh(m2) - h(m2) * yh(m3)) / (h(m3) - h(m2))
                   x2 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   y2 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                 
'===========================================================
                  '     Case 9 - Line between sides 3-1 and 1-2
                 
'===========================================================
                 Case 9
                   x1 = (h(m1) * xh(m3) - h(m3) * xh(m1)) / (h(m1) - h(m3))
                   y1 = (h(m1) * yh(m3) - h(m3) * yh(m1)) / (h(m1) - h(m3))
                   x2 = (h(m2) * xh(m1) - h(m1) * xh(m2)) / (h(m2) - h(m1))
                   y2 = (h(m2) * yh(m1) - h(m1) * yh(m2)) / (h(m2) - h(m1))
                End Select
        '--------------------------------------------------------------
                'this is where the program specific drawing routine comes in.
                'This specific command will work well for a properly dimensioned
                'vb picture box or vb form (where "ContoursFrm" is the name of the form)
               
'-------------------------------------------------------------------
                'make it fit into form size
                scale_x = ContoursFrm.Picture1.ScaleWidth / (Xmax - Xmin)
                scale_y = ContoursFrm.Picture1.ScaleHeight / (Ymax - Ymin)
                'color is that of contour(k)
                colornum% = (k + 1) * (UBound(cpt, 2) + 1) / nc
                color = RGB(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))
                ContoursFrm.Picture1.Line (CInt((x1 - Xmin) * scale_x), CInt((Ymax - y1) * scale_y))-(CInt((x2 - Xmin) * scale_x), CInt((Ymax - y2) * scale_y)), color
'-------------------------------------------------------------------
               End If
             Next m
          End If
        Next k
      End If
     Next i
  
'--------------------------------------------------------------------------------------
   'used to refresh the drawing surface after each row is contoured (for impatient users)
   ContoursFrm.Picture1.Refresh
   ContoursFrm.txtProgressY.Text = j
   newprogress% = Int(100 * ((jub - j) / (jub - jlb)))
   If initprogress% <> newprogress% Then
      initprogress% = newprogress%
      ContoursFrm.txtProgressX.Text = Str$(initprogress%) + "%"
      ContoursFrm.progbar.Value = initprogress%
      End If
   DoEvents
  
'-------------------------------------------------------------------------------------
   Next j
End If

ContoursFrm.Caption = "Contours"

End Function

Public Function Min(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Min = v1
Else: Min = v2
End If
End Function

Public Function Max(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Max = v2
Else: Max = v1
End If
End Function


