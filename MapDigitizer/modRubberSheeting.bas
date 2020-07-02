Attribute VB_Name = "modRubberSheeting"
Option Explicit

DefDbl A-H, O-Z
DefLng I-N

Public A2_CALST1, B0_CALST1, B1_CALST1, B2_CALST1
Public SX0_CALST1, SY0_CALST1, GX0_CALST1, GY0_CALST1, A0_CALST1, A1_CALST1
Public SXPIX_CALST3, SYPIX_CALST3, GXPIX_CALST3, GYPIX_CALST3
Public SA_CALST4(4), SB_CALST4(4), ICAL_CALST4, JCAL_CALST4
Public DX_CALST2(), DY_CALST2(), XBAR_CALST2(), YBAR_CALST2()

Public ANG As Double 'mean angle that original coordinate system is rotated by
Public SX0 As Double, SY0 As Double 'origin of rotation backwards

Public GeoPerPixelX As Double, GeoPerPixelY As Double
Public BeginPixelX As Long, BeginPixelY As Long

Public StatusColor As Long
Public RotateOut As Boolean

Private Declare Function beep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : Step1toStep2
' Author    : Chaim Keller
' Date      : 2/5/2015
' Purpose   :  runs Step1 and Step 2 of the rubber sheeting
'---------------------------------------------------------------------------------------
'
Public Function Step1toStep2() As Integer
  'uses John Hall's rubber sheeting converted fortran routines to convert from
  'screen coordinates to geographic map coordinates taking map sheet scanning distortions into consideration
  
  'Ref: "Gridded Affine Transformation and Rubber Sheeting Algorithm with Fortran
  'Program for Calibrating Scanned Hydrographic Survey Maps"
  'Y. Doythser, John K. Hall, Computers and Geosciences, Vol. 23, No. 7, pp. 785-791, 1997
  'inputed coordinates to transofrm to geo coordinates
  
'    DefDbl A-H, O-Z
'    DefLng I-N
     Dim AB As Double, ac As Double, BB As Double, bc As Double, cc As Double
     Dim alx As Double, blx As Double, clx As Double, aly As Double, bly As Double
     Dim cly As Double, AA As Double, aai As Double
     Dim aa1 As Double, BB1 As Double, bb1i As Double, bc1 As Double, cc1 As Double
     Dim blx1 As Double, clx1 As Double, bly1 As Double, cly1 As Double, cc2 As Double
     Dim clx2 As Double, cly2 As Double, cc2i As Double
     Dim dsx As Double, dsy As Double, dgx As Double, dgy As Double
     
     Dim ier As Integer, i As Long, j As Long, K As Long
     Dim Xcoord As Double, Ycoord As Double
     
     Dim SX As Double, SY As Double
     
     Dim RoundOffX As Double
     Dim RoundOffY As Double
     
     RotateOut = False
   
   On Error GoTo Step1toStep2_Error
   
    ier = 0
    
    If NX_CALDAT > 1 And NY_CALDAT > 1 Then
      XGridSteps = (LRGridX - ULGridX) / (NX_CALDAT - 1)
      YGridSteps = (ULGridY - LRGridY) / (NY_CALDAT - 1)
      End If
      
    If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60#)
    If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60#)
    
    DigiRubberSheeting = False
   
    If NX_CALDAT = 0 Or NY_CALDAT = 0 Then
       'error detected
       ier = -1
       Step1toStep2 = ier
       Exit Function
    ElseIf numRS <> NX_CALDAT * NY_CALDAT Then
       'not all the grid points were digitized
       Call MsgBox("Rubber sheeting doesn't seem to have been completed!" _
                   & vbCrLf & "" _
                   & vbCrLf & "Check if all the grid intersections are marked." _
                   , vbInformation Or vbDefaultButton1, "Rubber Sheeting")
       ier = -1
       Step1toStep2 = ier
       Exit Function
       End If
    
    'first zero the arrays
    ReDim DX_CALST2(0, 0)
    ReDim DY_CALST2(0, 0)
    ReDim XBAR_CALST2(0)
    ReDim YBAR_CALST2(0)
    
    ReDim SX_CALDAT(0, 0)
    ReDim SY_CALDAT(0, 0)
    ReDim GX_CALDAT(0, 0)
    ReDim GY_CALDAT(0, 0)
    
    'redimension the common arrays
    ReDim DX_CALST2(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim DY_CALST2(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim XBAR_CALST2(1 To NX_CALDAT)
    ReDim YBAR_CALST2(1 To NY_CALDAT)
    
    ReDim SX_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim SY_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim GX_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim GY_CALDAT(1 To NX_CALDAT, 1 To NY_CALDAT)

    'initialize the sx,sy,gx,gy arrays
    For i = 1 To NX_CALDAT
       For j = 1 To NY_CALDAT
          SX_CALDAT(i, j) = INIT_VALUE
          SY_CALDAT(i, j) = INIT_VALUE
          GX_CALDAT(i, j) = INIT_VALUE
          GY_CALDAT(i, j) = INIT_VALUE
       Next j
    Next i

    Dim numCheck
    numCheck = 0
    For j = 1 To NY_CALDAT
       For i = 1 To NX_CALDAT
          Xcoord = ULGridX + (i - 1) * XGridSteps
          Ycoord = LRGridY + (j - 1) * YGridSteps 'the latitude arrays are filled opposite than the wizard, i.e., from South to North
          For K = 0 To numRS - 1
'              If RS(k).XGeo = Xcoord And RS(k).YGeo = Ycoord Then
              If RS(K).XGeo >= Xcoord - RoundOffX And RS(K).XGeo <= Xcoord + RoundOffX And _
                 RS(K).YGeo >= Ycoord - RoundOffY And RS(K).YGeo <= Ycoord + RoundOffY Then
                 SX_CALDAT(i, j) = RS(K).xScreen
                 SY_CALDAT(i, j) = -RS(K).yScreen 'minus sign is in keeping with JKH's convention in his paper that y screen coordinate increases from bottom to top
                 GX_CALDAT(i, j) = RS(K).XGeo
                 GY_CALDAT(i, j) = RS(K).YGeo
                 numCheck = numCheck + 1
                 End If
              DoEvents
          Next K
        Next i
        Call UpdateStatus(GDRSfrm, 1, CLng(100 * j / NX_CALDAT))
        GDRSfrm.Refresh
    Next j
    If numCheck <> NX_CALDAT * NY_CALDAT Then
       Call MsgBox("Some coordinates are erroneous, or are missing" _
                   & vbCrLf & "" _
                   & vbCrLf & "The RS file may be an old format type, so check it." _
                   & vbCrLf & "(There should be a 0 or 1 on the first line.)" _
                   & vbCrLf & "" _
                   & vbCrLf & "In the meantime, digitizing functions are disabled." _
                   , vbExclamation Or vbDefaultButton1, "Rubber Sheeting Error")
       Err.Raise vbObjectError + 513 + 1, "Step1toStep2", "Missing geographic coordinates"
   Else
       'wite to diagnostic log file
'       Dim filout%
'       filout% = FreeFile
'       Open App.Path & "\Rs-sorted-out.txt" For Output As #filout%
'       For j = 1 To NY_CALDAT
'          For i = 1 To NX_CALDAT
'             Write #filout%, SX_CALDAT(i, j), SY_CALDAT(i, j), GX_CALDAT(i, j), GY_CALDAT(i, j)
'          Next i
'       Next j
'       Close #filout%
       End If
       
    'now determine if there is significant rotation of the grid lines, and if so, rotate them back if flagged
    'this method should be able to handle any rotation without rotating back, but just in case keep this code....
    'find mean rotation angle
    ANG = 0#
    For i = 1 To NY_CALDAT
       ANG = ANG + Atan2(SX_CALDAT(NX_CALDAT, i) - SX_CALDAT(1, i), SY_CALDAT(1, i) - SY_CALDAT(NX_CALDAT, i))
    Next i
        
    ANG = ANG / NY_CALDAT 'this is mean angle
        
    '    If ANG > 0.0017 Then
    '
    '        Select Case MsgBox("Warning: the grid lines are rotated by:" & str$(ANG / cd) & " degrees." _
    '                           & vbCrLf & "" _
    '                           & vbCrLf & "You need to rotate the map using a photo editor by that amount" _
    '                           & vbCrLf & "until the grid lines are approximately parallel." _
    '                           & vbCrLf & "Otherwise, the rubber sheeting will fail" _
    '                           & vbCrLf & "" _
    '                           & vbCrLf & "Do you want to stop the rubber sheeting (recommended)?" _
    '                           & vbCrLf & "" _
    '                           , vbYesNo Or vbExclamation Or vbDefaultButton1, "Grid lines are rotated warning")
    '
    '            Case vbYes
    '
    '               ier = -1
    '               Step1toStep2 = ier
    '               Exit Function
    '
    '            Case vbNo
    '
    '        End Select
    '        End If
    If RotateOut Then
        'pick origin of rotation of entire frame
        SX0 = SX_CALDAT(1, 1)
        SY0 = SY_CALDAT(1, 1)
    '
    ''    -----------------diagnostics-----------------------------------
    ''    Dim filout%
    ''    filout% = FreeFile
    ''    Open App.Path & "\Rotated-grid.txt" For Output As #filout%
    ''    ------------------------------------------------------------
    '
        'rotate backwards (clockwise) for positive ANG
        For j = 1 To NY_CALDAT
           For i = 1 To NX_CALDAT
               SX = SX_CALDAT(i, j) - SX0
               SY = SY0 - SY_CALDAT(i, j)
               SX_CALDAT(i, j) = SX * Cos(ANG) + SY * Sin(ANG) + SX0
               SY_CALDAT(i, j) = SY0 + SX * Sin(ANG) - SY * Cos(ANG)
    '          '---------------diagnostics-------------------------------------
    '           Write #filout%, SX_CALDAT(i, j), SY_CALDAT(i, j), GX_CALDAT(i, j), GY_CALDAT(i, j)
    '          '---------------------------------------------------------------
           Next i
        Next j
    End If
''    ----------diagnostics-------------------------
''    Close #filout%
''    ------------------------------------------------
    
    'finished initialization and checks and now begin actual calculation of the affine transformation
    '-----------------------------Step 1------------------------------------------
      SX0_CALST1 = SX_CALDAT(1, 1)
      SY0_CALST1 = SY_CALDAT(1, 1)
      GX0_CALST1 = GX_CALDAT(1, 1)
      GY0_CALST1 = GY_CALDAT(1, 1)
      AB = 0#
      ac = 0#
      BB = 0#
      bc = 0#
      cc = 0#
      alx = 0#
      blx = 0#
      clx = 0#
      aly = 0#
      bly = 0#
      cly = 0#
      AA = NX_CALDAT * NY_CALDAT
      
      Call UpdateStatus(GDRSfrm, 1, 0)
      
      For i = 1 To NX_CALDAT
          For j = 1 To NY_CALDAT
              dsx = SX_CALDAT(i, j) - SX0_CALST1
              dsy = SY_CALDAT(i, j) - SY0_CALST1
              dgx = GX_CALDAT(i, j) - GX0_CALST1
              dgy = GY_CALDAT(i, j) - GY0_CALST1
              AB = AB + dsx
              ac = ac + dsy
              BB = BB + dsx * dsx
              bc = bc + dsx * dsy
              cc = cc + dsy * dsy
              alx = alx + dgx
              blx = blx + dgx * dsx
              clx = clx + dgx * dsy
              aly = aly + dgy
              bly = bly + dgy * dsx
              cly = cly + dgy * dsy
              DoEvents
          Next j
          Call UpdateStatus(GDRSfrm, 1, CLng(100 * i / NX_CALDAT))
          GDRSfrm.Refresh
      Next i
      
      '     compute inverses to speed calculation
      aai = 1# / AA
      BB1 = BB - aai * AB * AB
      bb1i = 1# / BB1
      bc1 = bc - aai * ac * AB
      cc1 = cc - aai * ac * ac
      blx1 = blx - aai * AB * alx
      clx1 = clx - aai * ac * alx
      bly1 = bly - aai * AB * aly
      cly1 = cly - aai * ac * aly
      cc2 = cc1 - bb1i * bc1 * bc1
      clx2 = clx1 - bb1i * bc1 * blx1
      cly2 = cly1 - bb1i * bc1 * bly1
      '     compute inverse to speed calculation
      cc2i = 1# / cc2
      ' '     now calculate coefficients of the affine transformation
      A2_CALST1 = cc2i * clx2
      A1_CALST1 = bb1i * (blx1 - bc1 * A2_CALST1)
      A0_CALST1 = aai * (alx - ac * A2_CALST1 - AB * A1_CALST1)
      B2_CALST1 = cc2i * cly2
      B1_CALST1 = bb1i * (bly1 - bc1 * B2_CALST1)
      B0_CALST1 = aai * (aly - ac * B2_CALST1 - AB * B1_CALST1)
      '----------------------------end of step1---------------------------------
      
      
      '----------------------------Step 2-----------------------------------------
      
      Call UpdateStatus(GDRSfrm, 1, 0)
      
      For j = 1 To NY_CALDAT
          For i = 1 To NX_CALDAT
              DX_CALST2(i, j) = GX_CALDAT(i, j) - GX0_CALST1 - (A0_CALST1 + A1_CALST1 * (SX_CALDAT(i, j) - SX0_CALST1) + A2_CALST1 * (SY_CALDAT(i, j) - SY0_CALST1))
              DY_CALST2(i, j) = GY_CALDAT(i, j) - GY0_CALST1 - (B0_CALST1 + B1_CALST1 * (SX_CALDAT(i, j) - SX0_CALST1) + B2_CALST1 * (SY_CALDAT(i, j) - SY0_CALST1))
              DoEvents
         Next i
         Call UpdateStatus(GDRSfrm, 1, CLng(100 * i / NY_CALDAT))
         GDRSfrm.Refresh
      Next j
      '     compute the vectors xbar(nx) , ybar(ny) , the average values of the
      '      transformed (affine)
      
      Call UpdateStatus(GDRSfrm, 1, 0)
      
      For i = 1 To NX_CALDAT
          XBAR_CALST2(i) = 0#
          For j = 1 To NY_CALDAT
              XBAR_CALST2(i) = XBAR_CALST2(i) + (A0_CALST1 + A1_CALST1 * (SX_CALDAT(i, j) - SX0_CALST1) + A2_CALST1 * (SY_CALDAT(i, j) - SY0_CALST1))
              DoEvents
          Next j
          XBAR_CALST2(i) = GX0_CALST1 + XBAR_CALST2(i) / CDbl(NY_CALDAT)
          Call UpdateStatus(GDRSfrm, 1, CLng(100 * i / NX_CALDAT))
          GDRSfrm.Refresh
      Next i
      
      Call UpdateStatus(GDRSfrm, 1, 0)
      
      For j = 1 To NY_CALDAT
          YBAR_CALST2(j) = 0#
          For i = 1 To NX_CALDAT
              YBAR_CALST2(j) = YBAR_CALST2(j) + (B0_CALST1 + B1_CALST1 * (SX_CALDAT(i, j) - SX0_CALST1) + B2_CALST1 * (SY_CALDAT(i, j) - SY0_CALST1))
              DoEvents
          Next i
          YBAR_CALST2(j) = GY0_CALST1 + YBAR_CALST2(j) / CDbl(NX_CALDAT)
          Call UpdateStatus(GDRSfrm, 1, CLng(100 * i / NY_CALDAT))
          GDRSfrm.Refresh
      Next j
      
'      '--------------diagnostics------------------------
'      Dim S2filnum%
'      S2filnum% = FreeFile
'      Open App.Path & "\Step2-out.txt" For Output As #S2filnum%
'      For j = NY_CALDAT To 1 Step -1
'         For i = 1 To NX_CALDAT
'            Write #S2filnum%, j, i, DX_CALST2(i, j), DY_CALST2(i, j)
'         Next i
'      Next j
'      Close #S2filnum%
'      '----------------------
      '----------------------------------end of step2------------------------------------------------
      
   If ier = 0 Then DigiRubberSheeting = True
   Step1toStep2 = ier

   On Error GoTo 0
   Exit Function

Step1toStep2_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Step1toStep2 of Module modRubberSheeting"
    ier = -1 'Err.Number
    Step1toStep2 = ier
   
End Function


'---------------------------------------------------------------------------------------
' Procedure : RS_pixel_to_coord
' Author    : Chaim Keller
' Date      : 2/5/2015
' Purpose   : runs actual conversion, i.e., Steps 3-5
'---------------------------------------------------------------------------------------
'
Public Function RS_pixel_to_coord(Xpix As Double, Ypix As Double, XGeo As Double, YGeo As Double) As Integer
  
  'uses John Hall's rubber sheeting converted fortran routines to convert from
  'screen coordinates to geographic map coordinates taking map sheet scanning distortions into consideration
  
  'Ref: "Gridded Affine Transformation and Rubber Sheeting Algorithm with Fortran
  'Program for Calibrating Scanned Hydrographic Survey Maps"
  'Y. Doythser, John K. Hall, Computers and Geosciences, Vol. 23, No. 7, pp. 785-791, 1997
  'inputed coordinates to transofrm to geo coordinates
  
'    DefDbl A-H, O-Z
'    DefLng I-N

    Dim su As Double, sv As Double
    Dim ier As Integer, inout As Integer, icheck As Integer
    Dim i As Long, j As Long, gxpixf As Double, gypixf As Double
    Dim sxa As Double, sya As Double, sxb As Double, syb As Double
    Dim sdx As Double, sdy As Double, sdd As Double, sdxx As Double
    Dim sdyy As Double, Desc$
    
    Dim SXPIX As Double, SYPIX As Double
  
    On Error GoTo RS_pixel_to_coord_Error
    
    ier = 0
    
'-------------------diagnostics----------------------------
'    filnum% = FreeFile
'    Open App.Path & "\r-41-19_20-RS.txt" For Output As #filnum%
'    For j = 1 To NY_CALDAT
'       For i = 1 To NX_CALDAT
'          Write #filnum%, SX_CALDAT(i, j), SY_CALDAT(i, j), GX_CALDAT(i, j), GY_CALDAT(i, j)
'       Next i
'    Next j
'    Close #filnum%
'---------------------------------------------------
    
    GDMDIform.Text1.ForeColor = QBColor(0)
    GDMDIform.Text2.ForeColor = QBColor(0)
    
    If RotateOut Then
        'first rotate backwards (clockwise) by mean tilt (ANG) of grid lines
        SXPIX = Xpix - SX0
        SYPIX = SY0 - Ypix
        SXPIX_CALST3 = SXPIX * Cos(ANG) + SYPIX * Sin(ANG) + SX0
        SYPIX_CALST3 = -(SY0 + SXPIX * Sin(ANG) - SYPIX * Cos(ANG)) 'minus sign is in keeping with calculation's convention that y screen coordinate increases from bottom to top
    Else
        SXPIX_CALST3 = Xpix
        SYPIX_CALST3 = -Ypix 'minus sign is in keeping with calculation's convention that y screen coordinate increases from bottom to top
        End If
        
    '--------------------------Step 3-----------------------------------------
    GXPIX_CALST3 = A0_CALST1 + A1_CALST1 * (SXPIX_CALST3 - SX0_CALST1) + A2_CALST1 * (SYPIX_CALST3 - SY0_CALST1) + GX0_CALST1
    GYPIX_CALST3 = B0_CALST1 + B1_CALST1 * (SXPIX_CALST3 - SX0_CALST1) + B2_CALST1 * (SYPIX_CALST3 - SY0_CALST1) + GY0_CALST1
    '---------------------------end of Step 3---------------------------------------------
    
    '----------------------------------Step 4---------------------------------------------------
    inout = 0
    'might have to be lenient with this since some coordinates are outside this range
    If (GXPIX_CALST3 < XBAR_CALST2(1)) Then GoTo 900
    If (GXPIX_CALST3 > XBAR_CALST2(NX_CALDAT)) Then GoTo 900
    If (GYPIX_CALST3 < YBAR_CALST2(1)) Then GoTo 900
    If (GYPIX_CALST3 > YBAR_CALST2(NY_CALDAT)) Then GoTo 900
    inout = 1
    For i = 1 To NX_CALDAT - 1
        If (GXPIX_CALST3 >= XBAR_CALST2(i) And GXPIX_CALST3 <= XBAR_CALST2(i + 1)) Then GoTo S4_L10
    Next i
'   at this point in the program, i = NX_CALDAT-1
S4_L10:
    For j = 1 To NY_CALDAT - 1
        If (GYPIX_CALST3 >= YBAR_CALST2(j) And GYPIX_CALST3 <= YBAR_CALST2(j + 1)) Then GoTo S4_L20
S4_L15:
    Next j
'   at this point in the program, j = NY_CALDAT - 1
S4_L20:
    icheck = 1: GoSub substep4
    If (SB_CALST4(1) < 0# And i > 1) Then GoTo S4_L25
    GoTo S4_L30
S4_L25:
    i = i - 1
    GoTo S4_L20
S4_L30:
    icheck = 2: GoSub substep4
    If (SB_CALST4(2) < 0# And j < NY_CALDAT - 2) Then GoTo S4_L35
    GoTo S4_L40
S4_L35:
    j = j + 1
    GoTo S4_L30
S4_L40:
    icheck = 3: GoSub substep4
    If (SB_CALST4(3) < 0# And i < NX_CALDAT - 2) Then GoTo S4_L45
    GoTo S4_L50
S4_L45:
    i = i + 1
    GoTo S4_L40
S4_L50:
    icheck = 4: GoSub substep4
    If (SB_CALST4(4) < 0# And j > 1) Then GoTo S4_L55
    GoTo S4_L60
S4_L55:
    j = j - 1
    GoTo S4_L50
    '     if we arrive here then sb(1) =>0; sb(2) =>0; sb(3) =>0; sb(4) =>0;
S4_L60:
    icheck = 1: GoSub substep4
    
    icheck = 2: GoSub substep4
    
    icheck = 3: GoSub substep4
    
    icheck = 4: GoSub substep4
    
    ICAL_CALST4 = i
    JCAL_CALST4 = j
    '----------------------------------------end of Step 4---------------------

    '-----------------------------Step 5-----------------------------------------
    If (inout = 0) Then GoTo 900
    
    i = ICAL_CALST4
    j = JCAL_CALST4
    su = (SA_CALST4(2) * SB_CALST4(4) - SB_CALST4(2) * SA_CALST4(4) + SB_CALST4(2)) / (SB_CALST4(2) + SB_CALST4(4))
    sv = (SA_CALST4(1) * SB_CALST4(3) - SB_CALST4(1) * SA_CALST4(3) + SB_CALST4(1)) / (SB_CALST4(1) + SB_CALST4(3))
    gxpixf = GXPIX_CALST3 + DX_CALST2(i, j) + su * (DX_CALST2(i + 1, j) - DX_CALST2(i, j)) + sv * (DX_CALST2(i, j + 1) - DX_CALST2(i, j)) + su * sv * (DX_CALST2(i, j) + DX_CALST2(i + 1, j + 1) - DX_CALST2(i, j + 1) - DX_CALST2(i + 1, j))
    gypixf = GYPIX_CALST3 + DY_CALST2(i, j) + su * (DY_CALST2(i + 1, j) - DY_CALST2(i, j)) + sv * (DY_CALST2(i, j + 1) - DY_CALST2(i, j)) + su * sv * (DY_CALST2(i, j) + DY_CALST2(i + 1, j + 1) - DY_CALST2(i, j + 1) - DY_CALST2(i + 1, j))
    
    XGeo = gxpixf 'conversion in geographic coordinates
    YGeo = gypixf
    
    RS_pixel_to_coord = ier
    
    Exit Function
    
900:
    'Call MsgBox("Conversion wasn't successful for these coordinates!", vbExclamation, "Rubber Sheeting")
    If RSMethodBoth Then
       ier = RS_pixel_to_coord2(CDbl(Xpix), CDbl(Ypix), XGeo, YGeo)
       If ier = 0 Then
          ier = 0
          RS_pixel_to_coord = ier
          Exit Function
          End If
    Else
      GDMDIform.StatusBar1.Panels(1).Text = "Warning: This portion of the map has not been rubber sheeted!" '"Error: " & Desc$
      GDMDIform.Text1.ForeColor = QBColor(12)
      GDMDIform.Text2.ForeColor = QBColor(12)
        
      ier = -1
      RS_pixel_to_coord = ier
      End If
         

Exit Function

'-----------------------Gosub substep4----------------------------------------
substep4:

    If icheck = 1 Then
       GoTo S44_L5
    ElseIf icheck = 2 Then
       GoTo S44_L10
    ElseIf icheck = 3 Then
       GoTo S44_L15
    ElseIf icheck = 4 Then
       GoTo S44_L20
       End If
S44_L5:
    sxa = SX_CALDAT(i, j)
    sya = SY_CALDAT(i, j)
    sxb = SX_CALDAT(i, j + 1)
    syb = SY_CALDAT(i, j + 1)
    GoTo S44_L25
S44_L10:
    sxa = SX_CALDAT(i, j + 1)
    sya = SY_CALDAT(i, j + 1)
    sxb = SX_CALDAT(i + 1, j + 1)
    syb = SY_CALDAT(i + 1, j + 1)
    GoTo S44_L25
S44_L15:
    sxa = SX_CALDAT(i + 1, j + 1)
    sya = SY_CALDAT(i + 1, j + 1)
    sxb = SX_CALDAT(i + 1, j)
    syb = SY_CALDAT(i + 1, j)
    GoTo S44_L25
S44_L20:
    sxa = SX_CALDAT(i + 1, j)
    sya = SY_CALDAT(i + 1, j)
    sxb = SX_CALDAT(i, j)
    syb = SY_CALDAT(i, j)
S44_L25:
    sdx = sxb - sxa
    sdy = syb - sya
    sdd = Sqr(sdx * sdx + sdy * sdy)
    '     note that sxc,syx is the same as sxpix,sypix
    sdxx = SXPIX_CALST3 - sxa
    sdyy = SYPIX_CALST3 - sya
    SA_CALST4(icheck) = (sdxx * sdx + sdyy * sdy) / (sdd * sdd)
    SB_CALST4(icheck) = (sdxx * sdy - sdyy * sdx) / (sdd * sdd)
    Return

   On Error GoTo 0
   Exit Function

RS_pixel_to_coord_Error:

    If RSMethodBoth Then
       ier = RS_pixel_to_coord2(CDbl(Xpix), CDbl(Ypix), XGeo, YGeo)
       Exit Function
    Else

        'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RS_pixel_to_coord of Module modRubberSheeting"
        ier = Err.Number
        Desc$ = Err.Description
    '    Beep 1000, 100
        GDMDIform.StatusBar1.Panels(1).Text = "Warning: This portion of the map has not been rubber sheeted!" '"Error: " & Desc$
        GDMDIform.Text1.ForeColor = QBColor(12)
        GDMDIform.Text2.ForeColor = QBColor(12)
        RS_pixel_to_coord = ier
        End If

End Function

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
' Return the smallest parameter value.
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
' Procedure : ReadRSfile
' Author    : Chaim Keller
' Date      : 2/6/2015
' Purpose   : read stored Rubber Sheeting file
'---------------------------------------------------------------------------------------
'
Function ReadRSfile() As Integer

   'read stored Rubber Sheeting file
   
   Dim pos%, picext$, RSfilnam$
   Dim gddm As Long, gddw As Long, Ret As Long, ier As Integer
   Dim tmpRS As Long, doclin$, Coords() As String

   On Error GoTo ReadRSfile_Error
   
      If GDRSfrmVis Then tmpRS = numRS

'      pos% = InStr(picnam$, ".")
'      picext$ = Mid$(picnam$, pos% + 1, 3)
      RSfilnam$ = App.Path & "\" & RootName(picnam$) & "-RS" & ".txt"

      If Dir(RSfilnam$) <> sEmpty Then

         If RSopenedfile Then
            RSopenedfile = False
            Close #RSfilnum%
            End If

         RSfilnum% = FreeFile
         Open RSfilnam$ For Input As #RSfilnum%
         RSopenedfile = True
         
         numRS = 0
         ReDim RS(numRS)
         
         gddm = GDform1.Picture2.DrawMode
         gddw = GDform1.Picture2.DrawWidth
         GDform1.Picture2.DrawMode = 13
         GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
         
         Line Input #RSfilnum%, doclin$
         If InStr(doclin$, ",") Then
            Call MsgBox("The RS file seems to be corrupted or of an older type." _
                        & vbCrLf & "(The first line must consist of a boolean flag equal to 0 or 1)" _
                        & vbCrLf & "" _
                        & vbCrLf & "Please check the RS file, and edit it if necessary." _
                        , vbExclamation, "RS file import error")
            ReadRSfile = -1
            Exit Function
         Else
            DigiRSStepType = val(doclin$)
            End If
         
         Do Until EOF(RSfilnum%)
            'use safe read
            Line Input #RSfilnum%, doclin$
            Coords = Split(doclin$, ",")
            If UBound(Coords) = 3 Then
            
               ReDim Preserve RS(numRS)
               
               RS(numRS).xScreen = val(Coords(0))
               RS(numRS).yScreen = val(Coords(1))
               RS(numRS).XGeo = val(Coords(2))
               RS(numRS).YGeo = val(Coords(3))
               End If
               
'            Input #RSfilnum%, RS(numRS).xScreen, RS(numRS).yScreen, RS(numRS).XGeo, RS(numRS).YGeo

            'mark the points already done on the map
            GDform1.Picture2.Line (RS(numRS).xScreen * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom), RS(numRS).yScreen * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom))-(RS(numRS).xScreen * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom), RS(numRS).yScreen * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)), RSColor& 'QBColor(14)
            GDform1.Picture2.Line (RS(numRS).xScreen * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom), RS(numRS).yScreen * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom))-(RS(numRS).xScreen * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom), RS(numRS).yScreen * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom)), RSColor& 'QBColor(14)
            
            numRS = numRS + 1
            If GDRSfrmVis And numRS > tmpRS Then 'some points erased already
               numRS = tmpRS
               Exit Do
               End If
            
         Loop
         
         GDform1.Picture2.DrawMode = gddm
         GDform1.Picture2.DrawWidth = gddw

         Close #RSfilnum%
         
'         If numRS > 0 Then
'            Call ShiftMap(CSng(RS(numRS - 1).xScreen * DigiZoom.LastZoom), CSng(RS(numRS - 1).yScreen * DigiZoom.LastZoom)) 'move map to last point
'            End If
         
         
'         RSopenedfile = False
          
'        If numRS = NX_CALDAT * NY_CALDAT Then
        
'           Select Case MsgBox("Rubber sheeting seems to be complete for this map!" _
'                              & vbCrLf & "" _
'                              & vbCrLf & "(Hint: Press the ''Activate Calculation Method'' button to activate it.)" _
'                              , vbOKOnly + vbInformation, "Rubber Sheeting")
           
'            Case vbOK
'                'the rubber sheeting is complete, so run it
'                Close #RSfilnum%  'no need to keep it open
'                RSopenedfile = False
'
''                'reload map without the x's
''                ier = ReDrawMap(0)
'
'                GDRSfrm.Visible = True
'                BringWindowToTop (GDRSfrm.hWnd)
'                GDRSfrm.cmdConvert.Enabled = True
                
'                'run the rubbersheeting and then close the dialog
'                GDRSfrm.cmdConvert = True
                'move map to last RS point done
'                If GDRSfrmVis And (GeoMap Or TopoMap) Then
'                  GDRSfrm.Visible = True
'                  ret = BringWindowToTop(GDRSfrm.hWnd)
'                  GDRSfrm.txtX = RS(numRS - 1).xScreen
'                  GDRSfrm.txtY = RS(numRS - 1).yScreen
'                  GDRSfrm.txtGeoX = RS(numRS - 1).XGeo
'                  GDRSfrm.txtGeoY = RS(numRS - 1).YGeo
'                  Call ShiftMap(CSng(RS(numRS - 1).xScreen * DigiZoom.LastZoom), CSng(RS(numRS - 1).yScreen * DigiZoom.LastZoom)) 'move map to last point
'                  End If
          
'           End Select
'          End If
'         Else
      
'            'move map to last RS point done
'            If GDRSfrmVis And (GeoMap Or TopoMap) And numRS - 1 > 0 Then
'              GDRSfrm.Visible = True
'              ret = BringWindowToTop(GDRSfrm.hWnd)
'              GDRSfrm.txtX = RS(numRS - 1).xScreen
'              GDRSfrm.txtY = RS(numRS - 1).yScreen
'              GDRSfrm.txtGeoX = RS(numRS - 1).XGeo
'              GDRSfrm.txtGeoY = RS(numRS - 1).YGeo
'              Call ShiftMap(CSng(RS(numRS - 1).xScreen * DigiZoom.LastZoom), CSng(RS(numRS - 1).yScreen * DigiZoom.LastZoom)) 'move map to last point
'              End If
         
'            End If
            
         End If

   On Error GoTo 0
   ReadRSfile = 0
   Exit Function

ReadRSfile_Error:
    If Err.Number = 62 Then Resume Next
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadRSfile of Module modRubberSheeting"
    ReadRSfile = -1
        
End Function
'---------------------------------------------------------------------------------------
' Procedure : RS_convert_init
' Author    : Chaim Keller
' Date      : 2/11/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function RS_convert_init() As Integer

'   initialize arrays for nearest neighbor calculation

   On Error GoTo RS_convert_init_Error
   
    Dim ier As Integer
    Dim i As Long, j As Long, K As Long

    ReDim SX_CALDAT_2(0, 0)
    ReDim SY_CALDAT_2(0, 0)
    ReDim GX_CALDAT_2(0, 0)
    ReDim GY_CALDAT_2(0, 0)
    
    ReDim SX_CALDAT_2(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim SY_CALDAT_2(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim GX_CALDAT_2(1 To NX_CALDAT, 1 To NY_CALDAT)
    ReDim GY_CALDAT_2(1 To NX_CALDAT, 1 To NY_CALDAT)
    
'    ReDim SX_ROT(0, 0)
'    ReDim SY_ROT(0, 0)
'
'    ReDim SX_ROT(1 To NX_CALDAT, 1 To NY_CALDAT)
'    ReDim SY_ROT(1 To NX_CALDAT, 1 To NY_CALDAT)
    
    Dim Xcoord As Double, Ycoord As Double
     
    Dim SX As Double, SY As Double
    Dim RoundOffX As Double, RoundOffY As Double
        
    RoundOffX = 0
    RoundOffY = 0
    
    If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60#)
    If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60#)

    Dim numCheck
    numCheck = 0
    For j = 1 To NY_CALDAT
       Ycoord = LRGridY + (j - 1) * Abs(YGridSteps) 'the latitude arrays are filled opposite than the wizard, i.e., from South to North
       For i = 1 To NX_CALDAT
          Xcoord = ULGridX + (i - 1) * Abs(XGridSteps)
          For K = 0 To numRS - 1
'              If RS(k).XGeo = Xcoord And RS(k).YGeo = Ycoord Then
              If RS(K).XGeo >= Xcoord - RoundOffX And RS(K).XGeo <= Xcoord + RoundOffX And _
                 RS(K).YGeo >= Ycoord - RoundOffY And RS(K).YGeo <= Ycoord + RoundOffY Then
                 SX_CALDAT_2(i, j) = RS(K).xScreen
                 SY_CALDAT_2(i, j) = RS(K).yScreen 'use screen coordinate convention where larger y is on bottom of screen
                 GX_CALDAT_2(i, j) = RS(K).XGeo
                 GY_CALDAT_2(i, j) = RS(K).YGeo
                 numCheck = numCheck + 1
                 Exit For
                 End If
              DoEvents
          Next K
        Next i
        DoEvents
        Call UpdateStatus(GDRSfrm, 1, CLng(100 * j / NY_CALDAT))
    Next j
    
    GDRSfrm.Refresh

    If numCheck <> NX_CALDAT * NY_CALDAT Then
       Call UpdateStatus(GDRSfrm, 1, 0)
       GDRSfrm.picProgBar.Visible = False
       Call MsgBox("Some coordinates are erroneous, or are missing" _
                   & vbCrLf & "" _
                   & vbCrLf & "The RS file may be an old format type, so check it." _
                   & vbCrLf & "(There should be a 0 or 1 on the first line.)" _
                   & vbCrLf & "" _
                   & vbCrLf & "In the meantime, digitizing functions are disabled." _
                   , vbExclamation Or vbDefaultButton1, "Rubber Sheeting Error")
       Err.Raise vbObjectError + 513 + 1, "RS_convert_init", "Missing geographic coordinates"
       End If
    
    'find mean rotation angle
    ANG = 0#
    For i = 1 To NY_CALDAT
       ANG = ANG + Atan2(SX_CALDAT_2(NX_CALDAT, i) - SX_CALDAT_2(1, i), SY_CALDAT_2(1, i) - SY_CALDAT_2(NX_CALDAT, i))
    Next i
    
    ANG = ANG / NY_CALDAT 'this is mean angle
   
    'pick origin of rotation of entire frame
    SX0 = SX_CALDAT_2(1, 1)
    SY0 = SY_CALDAT_2(1, 1)
'
''    -----------------diagnostics-----------------------------------
''    Dim filout%
''    filout% = FreeFile
''    Open App.Path & "\Rotated-grid.txt" For Output As #filout%
''    ------------------------------------------------------------
    Call UpdateStatus(GDRSfrm, 1, 0)
'
    'rotate backwards (clockwise) for positive ANG
    For j = 1 To NY_CALDAT
       For i = 1 To NX_CALDAT
           SX = SX_CALDAT_2(i, j) - SX0
           SY = SY0 - SY_CALDAT_2(i, j)
           SX_CALDAT_2(i, j) = SX * Cos(ANG) + SY * Sin(ANG) + SX0
           SY_CALDAT_2(i, j) = SY0 + SX * Sin(ANG) - SY * Cos(ANG)
'          '---------------diagnostics-------------------------------------
'           Write #filout%, SX_CALDAT(i, j), SY_CALDAT(i, j), GX_CALDAT(i, j), GY_CALDAT(i, j)
'          '---------------------------------------------------------------
       Next i
        DoEvents
        Call UpdateStatus(GDRSfrm, 1, CLng(100 * j / NY_CALDAT))
    Next j
    
    Call UpdateStatus(GDRSfrm, 1, 0)
    GDRSfrm.picProgBar.Visible = False
    GDRSfrm.Refresh
    
    'find average beginning pixel in X, Y
    BeginPixelX = 0
    For j = 1 To NY_CALDAT
       BeginPixelX = BeginPixelX + SX_CALDAT_2(1, j)
    Next j
    BeginPixelX = BeginPixelX / NY_CALDAT
    
    BeginPixelY = 0
    For i = 1 To NX_CALDAT
       BeginPixelY = BeginPixelY + SY_CALDAT_2(i, 1)
    Next i
    BeginPixelY = BeginPixelY / NX_CALDAT
    
    'find average pixel spacing along x and y direction
    GeoPerPixelX = 0
    For j = 1 To NY_CALDAT
       GeoPerPixelX = GeoPerPixelX + SX_CALDAT_2(NX_CALDAT, j) - SX_CALDAT_2(1, j)
    Next j
    GeoPerPixelX = GeoPerPixelX / NY_CALDAT
    'now convert to geo units per pixel in x direction
    GeoPerPixelX = (LRGridX - ULGridX) / GeoPerPixelX
    
    GeoPerPixelY = 0
    For i = 1 To NX_CALDAT
       GeoPerPixelY = GeoPerPixelY + SY_CALDAT_2(i, 1) - SY_CALDAT_2(i, NY_CALDAT)
    Next i
    GeoPerPixelY = GeoPerPixelY / NX_CALDAT
    'now convert to geo units per pixel in y direction
    GeoPerPixelY = (ULGridY - LRGridY) / GeoPerPixelY
    
       
'   'now rotate the geographic coordinates of the vertices to a coordinate system parallel with the screen coordinates
'   For j = 1 To NY_CALDAT
'      For i = 1 To NX_CALDAT - 1
'         ANG = Atan2(SX_CALDAT(i + 1, j) - SX_CALDAT(i, j), SY_CALDAT(i, j) - SY_CALDAT(i + 1, j))
'         'rotate the coordinates clockwise to be parallel to the screen
'         SX_ROT(i, j) = GX_CALDAT(i, j) * Cos(ANG) - GY_CALDAT(i, j) * Sin(ANG)
'         SY_ROT(i, j) = GX_CALDAT(i, j) * Sin(ANG) + GY_CALDAT(i, j) * Cos(ANG)
'      Next i
'      If (i = NX_CALDAT) Then
'         'use last angle calculated
'         SX_ROT(i, j) = GX_CALDAT(i, j) * Cos(ANG) - GY_CALDAT(i, j) * Sin(ANG)
'         SY_ROT(i, j) = GX_CALDAT(i, j) * Sin(ANG) + GY_CALDAT(i, j) * Cos(ANG)
'         End If
'     Call UpdateStatus(GDRSfrm, CLng(100 * j / NY_CALDAT))
'   Next j
   
   RS_convert_init = ier
   
   If ier = 0 Then DigiRubberSheeting = True

   On Error GoTo 0
   Exit Function

RS_convert_init_Error:

    ier = -1
    RS_convert_init = ier
    Call UpdateStatus(GDRSfrm, 1, 0)
    GDRSfrm.picProgBar.Visible = False
    GDMDIform.StatusBar1.Panels(1).Text = "Error " & Err.Number & " (" & Err.Description & ") in procedure RS_pixel_to_coord2_init of Module modRubberSheeting"

End Function
'---------------------------------------------------------------------------------------
' Procedure : RS_pixel_to_coord2
' Author    : Chaim Keller
' Date      : 2/12/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function RS_pixel_to_coord2(Xpix As Double, Ypix As Double, XGeo As Double, YGeo As Double) As Integer
    'use alternative method to do rubber sheeting
    'that will work for UTM grids that are rotated and maybe contorted
    Dim Corner1 As Boolean
    Dim Corner2 As Boolean
    Dim Corner3 As Boolean
    Dim Corner4 As Boolean
    
    Dim ier As Integer
    Dim ix1 As Long, ix2 As Long, iy1 As Long, iy2 As Long
    Dim XGeo1 As Double, YGeo1 As Double
    Dim XGeo2 As Double, YGeo2 As Double
    Dim i As Long, j As Long
    
    Dim SX As Double, SY As Double
    Dim SXPIX As Double, SYPIX As Double
    
    Dim inout As Integer, icheck As Integer
    
    Dim sxa As Double, sya As Double, sxb As Double, syb As Double
    Dim sdx As Double, sdy As Double, sdd As Double

    Dim sdxx As Double, sdyy As Double
    
    Dim GeoPerPixelXA As Double, GeoPerPixelYA As Double
    Dim BeginGeoXA As Double, BeginGeoYA As Double
    
    Dim XGeo0, YGeo0
    
    Dim SXPIX_CALST3_2, SYPIX_CALST3_2, GXPIX_CALST3_2, GYPIX_CALST3_2
    Dim SA_CALST4_2(4) As Double, SB_CALST4_2(4) As Double
    Dim ICAL_CALST4_2 As Long, JCAL_CALST4_2 As Long
    
    On Error GoTo RS_pixel_to_coord2_Error
    
    ier = 0
    GDMDIform.Text1.ForeColor = QBColor(0)
    GDMDIform.Text2.ForeColor = QBColor(0)
    
    'first rotate backwards (clockwise) by mean tilt (ANG) of grid lines
    'rotation matrix |cos(ang)  - sin ang]|
    '                |sin(ang)    cos(ang)|
    
'    '---diagnostics--------------
'    Xpix = 3976 'corresponds to 11448
'    Ypix = 2122 'corresponds to 7736
'    Xpix = 4093 'right one up one
'    Ypix = 1998
'    Xpix = 3857 'left one up one
'    Ypix = 2006
'    Xpix = 4093 'right one down one
'    Ypix = 2225
'    Xpix = 3876 'left one down one
'    Ypix = 2248
'    '-----------------------
    SXPIX = Xpix - SX0
    SYPIX = SY0 - Ypix
    SXPIX_CALST3_2 = SXPIX * Cos(ANG) + SYPIX * Sin(ANG) + SX0
    SYPIX_CALST3_2 = SY0 + SXPIX * Sin(ANG) - SYPIX * Cos(ANG)
   
    'now convert pixels to coordinates in lowest accuracy based on mean step sizes defined by the four corners of the map
    XGeo = (SXPIX_CALST3_2 - BeginPixelX) * GeoPerPixelX + ULGridX
    YGeo = (BeginPixelY - SYPIX_CALST3_2) * GeoPerPixelY + LRGridY
    XGeo0 = XGeo
    YGeo0 = YGeo
    
'   improve accuracy by using the step sizes based on the nearest vertices to the desired point
    'however, have to skip this improved accuracy calc. for those coordinates that are outside the grids
    If (SXPIX_CALST3_2 < 0.5 * (SX_CALDAT_2(1, 1) + SX_CALDAT_2(1, NY_CALDAT))) Then GoTo 900
    If (SXPIX_CALST3_2 > 0.5 * (SX_CALDAT_2(NX_CALDAT, 1) + SX_CALDAT_2(NX_CALDAT, NY_CALDAT))) Then GoTo 900
    If (SYPIX_CALST3_2 > 0.5 * (SY_CALDAT_2(1, 1) + SY_CALDAT_2(NX_CALDAT, 1))) Then GoTo 900 'opposite sign convention then rubber sheeting
    If (SYPIX_CALST3_2 < 0.5 * (SY_CALDAT_2(1, NY_CALDAT) + SY_CALDAT_2(NX_CALDAT, NY_CALDAT))) Then GoTo 900 'opposite sign convention then rubber sheeting
    
    inout = 1
    For i = 1 To NX_CALDAT - 1
        If (SXPIX_CALST3_2 >= SX_CALDAT_2(i, 1) And SXPIX_CALST3_2 <= SX_CALDAT_2(i + 1, 1)) Then GoTo S4_L10
    Next i
'   at this point in the program, i = NX_CALDAT-1
S4_L10:
    For j = 1 To NY_CALDAT - 1
        If (SYPIX_CALST3_2 >= SY_CALDAT_2(i, j + 1) And SYPIX_CALST3_2 <= SY_CALDAT_2(i, j)) Then GoTo S4_L20
S4_L15:
    Next j
'   at this point in the program, j = NY_CALDAT - 1
S4_L20:
    icheck = 1: GoSub substep4
    If (SB_CALST4_2(1) < 0# And i > 1) Then GoTo S4_L25
    GoTo S4_L30
S4_L25:
    i = i - 1
    GoTo S4_L20
S4_L30:
    icheck = 2: GoSub substep4
    If (SB_CALST4_2(2) < 0# And j < NY_CALDAT - 2) Then GoTo S4_L35
    GoTo S4_L40
S4_L35:
    j = j + 1
    GoTo S4_L30
S4_L40:
    icheck = 3: GoSub substep4
    If (SB_CALST4_2(3) < 0# And i < NX_CALDAT - 2) Then GoTo S4_L45
    GoTo S4_L50
S4_L45:
    i = i + 1
    GoTo S4_L40
S4_L50:
    icheck = 4: GoSub substep4
    If (SB_CALST4_2(4) < 0# And j > 1) Then GoTo S4_L55
    GoTo S4_L60
S4_L55:
    j = j - 1
    GoTo S4_L50
    '     if we arrive here then sb(1) =>0; sb(2) =>0; sb(3) =>0; sb(4) =>0;
S4_L60:
    icheck = 1: GoSub substep4
    
    icheck = 2: GoSub substep4
    
    icheck = 3: GoSub substep4
    
    icheck = 4: GoSub substep4
    
    ICAL_CALST4 = i
    JCAL_CALST4 = j
    
    GeoPerPixelXA = XGridSteps / ((SX_CALDAT_2(ICAL_CALST4 + 1, JCAL_CALST4) - SX_CALDAT_2(ICAL_CALST4, JCAL_CALST4)))
    GeoPerPixelYA = YGridSteps / ((SY_CALDAT_2(ICAL_CALST4, JCAL_CALST4) - SY_CALDAT_2(ICAL_CALST4, JCAL_CALST4 + 1)))
    BeginGeoXA = GX_CALDAT_2(ICAL_CALST4, JCAL_CALST4)
    BeginGeoYA = GY_CALDAT_2(ICAL_CALST4, JCAL_CALST4)
    XGeo = (SXPIX_CALST3_2 - SX_CALDAT_2(ICAL_CALST4, JCAL_CALST4)) * GeoPerPixelXA + BeginGeoXA
    YGeo = -(SYPIX_CALST3_2 - SY_CALDAT_2(ICAL_CALST4, JCAL_CALST4)) * GeoPerPixelYA + BeginGeoYA
    
    'now use these vertices as the reference vertices

900:
   
   RS_pixel_to_coord2 = ier

   On Error GoTo 0
   Exit Function
   
'-----------------------Gosub substep4----------------------------------------
substep4:

    If icheck = 1 Then
       GoTo S44_L5
    ElseIf icheck = 2 Then
       GoTo S44_L10
    ElseIf icheck = 3 Then
       GoTo S44_L15
    ElseIf icheck = 4 Then
       GoTo S44_L20
       End If
S44_L5:
    sxa = SX_CALDAT(i, j)
    sya = SY_CALDAT(i, j)
    sxb = SX_CALDAT(i, j + 1)
    syb = SY_CALDAT(i, j + 1)
    GoTo S44_L25
S44_L10:
    sxa = SX_CALDAT(i, j + 1)
    sya = SY_CALDAT(i, j + 1)
    sxb = SX_CALDAT(i + 1, j + 1)
    syb = SY_CALDAT(i + 1, j + 1)
    GoTo S44_L25
S44_L15:
    sxa = SX_CALDAT(i + 1, j + 1)
    sya = SY_CALDAT(i + 1, j + 1)
    sxb = SX_CALDAT(i + 1, j)
    syb = SY_CALDAT(i + 1, j)
    GoTo S44_L25
S44_L20:
    sxa = SX_CALDAT(i + 1, j)
    sya = SY_CALDAT(i + 1, j)
    sxb = SX_CALDAT(i, j)
    syb = SY_CALDAT(i, j)
S44_L25:
    sdx = sxb - sxa
    sdy = -(syb - sya) 'opposite site than in rubber sheeting
    sdd = Sqr(sdx * sdx + sdy * sdy)
    '     note that sxc,syx is the same as sxpix,sypix
    sdxx = SXPIX_CALST3_2 - sxa
    sdyy = -(SYPIX_CALST3_2 - sya) 'opposite sign than in rubber sheeting
    SA_CALST4_2(icheck) = (sdxx * sdx + sdyy * sdy) / (sdd * sdd)
    SB_CALST4_2(icheck) = (sdxx * sdy - sdyy * sdx) / (sdd * sdd)
    Return
   

RS_pixel_to_coord2_Error:

   If Err.Number = 9 Then
      'outside of 4 vertices, so just use simplified vertices
      XGeo = XGeo0
      YGeo = YGeo0
      ier = 0
      RS_pixel_to_coord2 = ier
      Exit Function
   Else
      ier = -1
      RS_pixel_to_coord2 = ier
      GDMDIform.StatusBar1.Panels(1).Text = "Error " & Err.Number & " (" & Err.Description & ") in procedure RS_pixel_to_coord2 of Module modRubberSheeting"
      End If
        
End Function

'---------------------------------------------------------------------------------------
' Procedure : Simple_pixel_to_coord
' Author    : Dr-John-K-Hall
' Date      : 3/8/2015
' Purpose   : Uses corner coordinates (pixel and geo) to calculate conversion from pixels to coordinates
'---------------------------------------------------------------------------------------
'
Public Function Simple_pixel_to_coord(Xpix As Double, Ypix As Double, XGeo As Double, YGeo As Double) As Integer

   On Error GoTo Simple_pixel_to_coord_Error
   
   'determine geo coordinates of upper left corner of picture

   XGeo = PixToCoordX * (Xpix - ULPixX) + CDbl(ULGeoX)
   YGeo = PixToCoordY * (Ypix - ULPixY) + CDbl(ULGeoY)
   
   Simple_pixel_to_coord = 0

   On Error GoTo 0
   Exit Function

Simple_pixel_to_coord_Error:
    Simple_pixel_to_coord = -1
End Function
