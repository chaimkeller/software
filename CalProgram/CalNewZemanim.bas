Attribute VB_Name = "CalNewZemanim"
Public CalcMethod%
'Public Const pi As Double = 3.14159265358979
'Public Const pi2 As Double = 2 * pi 'twice pi
'Public Const cd As Double = pi / 180# 'conversion from degrees to radians
'Public Const ch As Double = 1 / (cd * 15#)  'conv of deg to rad
Public pi As Double
Public pi2 As Double
Public cd As Double
Public ch As Double
Public Const hr As Double = 60#
Public TufikZman As Boolean 'flags if ouputing xml zemanim for Rav Tufik, if so write xml files in out format that the tufik program can read
   


Public Function fnrev(X As Double) As Double
    'places x into one of the four angular quadrants
    fnrev = X - Int(X / 360#) * 360#
End Function
Public Function DASIN(XX As Double) As Double
   If XX >= 1# Then
      DASIN = 90# * cd
   ElseIf XX <= -1# Then
      DASIN = 270# * cd
   Else
      DASIN = Atn(XX / Sqr(-XX * XX + 1#))
      End If
End Function
Public Function DACOS(XX As Double) As Double
   If XX >= 1# Then
      DACOS = 0#
   ElseIf XX <= -1# Then
      DACOS = 180# * cd
   Else
      DACOS = -Atn(XX / Sqr(-XX * XX + 1#)) + pi / 2
      End If
End Function
Public Function atan2(ByVal Y As Double, ByVal X As Double) _
    As Double
    'keeps angle within -pi to pi (returns angle in radians)
  Dim theta As Double

  If (Abs(X) < 0.0000001) Then
    If (Abs(Y) < 0.0000001) Then
      theta = 0#
    ElseIf (Y > 0#) Then
      theta = 1.5707963267949
    Else
      theta = -1.5707963267949
    End If
  Else
    theta = Atn(Y / X)
  
    If (X < 0) Then
      If (Y >= 0#) Then
        theta = 3.14159265358979 + theta
      Else
        theta = theta - 3.14159265358979
      End If
    End If
  End If
    
  atan2 = theta

End Function
'---------------------------------------------------------------------------------------
' Procedure : obliqeq
' Author    : Chaim Keller
' Date      : 2/22/2009
' Purpose   :'  OBLIQEQ  --  Calculate the obliquity of the ecliptic for a given Julian
'                 date.  This uses Laskar's tenth-degree polynomial fit
'                 (J. Laskar, Astronomy and Astrophysics, Vol. 157, page 68 [1986])
'                 which is accurate to within 0.01 arc second between AD 1000
'                 and AD 3000, and within a few seconds of arc for +/-10000
'                 years around AD 2000.  If we're outside the range in which
'                 this fit is valid (deep time) we simply return the J2000 value
'                 of the obliquity, which happens to be almost precisely the mean.  */
'                 taken from Home Planet
'---------------------------------------------------------------------------------------
'
Function obliqeq(jd As Double) As Double

    Dim oterms(10) As Double, eps As Double, U As Double, v As Double
    Dim i As Integer
    
        oterms(0) = Asec(-4680.93)
        oterms(1) = Asec(-1.55)
        oterms(2) = Asec(1999.25)
        oterms(3) = Asec(-51.38)
        oterms(4) = Asec(-249.67)
        oterms(5) = Asec(-39.05)
        oterms(6) = Asec(7.12)
        oterms(7) = Asec(27.87)
        oterms(8) = Asec(5.79)
        oterms(9) = Asec(2.45)

    eps = 23 + (26 / 60#) + (21.448 / 3600#)

    U = (jd - 2451545#) / 3652500#
    v = U

    If (Abs(U) < 1#) Then
       For i = 0 To 9
           eps = eps + oterms(i) * v
           v = v * U
       Next i
       End If

    obliqeq = eps

End Function
'---------------------------------------------------------------------------------------
' Procedure : kepler
' Author    : Chaim Keller
' Date      : 2/22/2009
' Purpose   : Solve the equation of Kepler (taken from Home Planet)
'---------------------------------------------------------------------------------------
'
Function kepler(m As Double, ecc As Double) As Double

    Dim e As Double, delta As Double, M_rad As Double
    Const EPSILON = 0.000001
    
    M_rad = m * cd

    e = M_rad
    delta = e - ecc * Sin(e) - M_rad

    While (Abs(delta) > EPSILON)
        e = e - delta / (1 - ecc * Cos(e))
        delta = e - ecc * Sin(e) - M_rad
    Wend
    
    kepler = e
    
End Function
Function dtr(X As Double) As Double
   dtr = X * cd 'Degrees to Radians
End Function
Function rtd(X As Double) As Double
   rtd = X / cd 'Radians to degrees
End Function
Function Asec(X As Double) As Double
    Asec = X / 3600#
End Function

'---------------------------------------------------------------------------------------
' Procedure : Decl3
' Author    : Chaim Keller
' Date      : 6/7/2013
' Purpose   : Calculates declination using low precison ephemerel of sun valid for any year of the Gregorian calendar
'---------------------------------------------------------------------------------------
'
Public Function Decl3(jdn As Double, hrs As Double, dy As Integer, td As Double, mday As Integer, mon As Integer, _
                      yrjd As Integer, yl As Integer, ms As Double, aas As Double, ob As Double, rv As Double, _
                      rs As Double, mode As Integer)
'///////////////////////////Decl3/////////////////////////////////////////////////////////////
'// calculate the solar declination to low accuracy valid for any year after the reformation of the Gregorian calendar.
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'// input variables: hrs = hours from midnight
'//                  td = time zone (negative for west of greenwhich)
'//                  yrjd = the Gregorian year
'//  output variables: ms  = Geometric mean longitude of the Sun, referred to the
'//                          mean equinox of the date.
'//                    aas = Sun's mean anomaly.
'//                    ob  = Obliquity of the ecliptic.
'//                    rs  = Sun's right accension in radians
'//                    rv  = Sun's radius in astronomical units
'//                    mday = day (1-366)
'//                    mon = month (1-12)
'//                    jdn = portion of Julain day number at 0:00:00 not referred to any equinox
'//  mode              = 0 to calculate the month and day and julian day and other astron. cosntanst
'//                    = 1 to skip everything that was already calculated
'//  parameters        apparent = should be TRUE_ if  the  apparent  position
'//                    (corrected  for  nutation  and aberration) is desired (defuult).
'//  dstflag           daylight saving time flag = 1 for DST, = 0 for Standard time
'//  routine returns the Sun's declination in radians
'////////////////////////////////////////////////////////////////////////////////////////////////
'{
'
'    //first calculate the julain day number
    Dim Y As Integer, m As Integer, a As Integer, b As Integer
    Dim UT As Double, jd As Double
    Dim t As Double, t2 As Double, t3 As Double, l As Double, ma As Double
    Dim e As Double, ea As Double, v As Double, theta As Double
    Dim omega As Double, eps As Double
    Dim apparent As Boolean
    Dim dstflag As Integer
    Dim month As Integer
    Dim day As Integer
    'month (beginning and ending days for normal and leap years
    Dim monthsN(13)
    Dim monthsL(13)
    
   On Error GoTo Decl3_Error

    monthsN(0) = 0
    monthsN(1) = 31
    monthsN(2) = 59
    monthsN(3) = 90
    monthsN(4) = 120
    monthsN(5) = 151
    monthsN(6) = 181
    monthsN(7) = 212
    monthsN(8) = 243
    monthsN(9) = 273
    monthsN(10) = 304
    monthsN(11) = 334
    monthsN(12) = 365
    
    monthsL(0) = 0
    monthsL(1) = 31
    monthsL(2) = 60
    monthsL(3) = 91
    monthsL(4) = 121
    monthsL(5) = 152
    monthsL(6) = 182
    monthsL(7) = 213
    monthsL(8) = 244
    monthsL(9) = 274
    monthsL(10) = 305
    monthsL(11) = 335
    monthsL(12) = 366
    
    apparent = True
'
'    /////////////////month and day///////////////////////
'
    If mode = 0 Then 'calculate the month and day
'        //determine if leap year
        yl = 365
        If yrjd Mod 4 = 0 Then
            yl = 366
            End If
        If yrjd Mod 100 = 0 And yrjd Mod 400 <> 0 Then
            yl = 365
            End If

        If yl = 365 Then
            While monthsN(month) < dy
                'Figure out the correct month.
                month = month + 1
            Wend

'            // Get the day thanks to the months array
            day = dy - monthsN(month - 1)
        Else
            While monthsL(month) < dy
                'Figure out the correct month.
                month = month + 1
            Wend

'            // Get the day thanks to the months array
            day = dy - monthsL(month - 1)
            End If

        mon = month
        mday = day

'    //////////////////////////////////////////////////////////////////////////
'
'        //calculate the julain date
'        //use Meeus formula 7.1 //modified from home planet code function: ucttoj::Astro.c

        m = mon
        Y = yrjd

        If (m <= 2) Then
            Y = Y - 1
            m = m + 12
            End If

        'Determine whether date is in Julian or Gregorian calendar based on
        'canonical date of calendar reform.

        If ((yrjd < 1582) Or ((yrjd = 1582) And ((mon < 9) Or (mon = 9 And mday < 5)))) Then
           b = 0
        Else
           a = (Y \ 100)
           b = 2 - a + (a \ 4)
           End If

        jdn = Fix(365.25 * (Y + 4716)) + Fix(30.6001 * (m + 1)) + _
                       mday + b - 1524.5  'Julian date, "JD", "jd", etc
                       
        End If
                        
    UT = hrs - td - dstflag '//universal time
    jd = jdn + UT / 24#
                        
'    ///////////////////////SunPos (modified from home planet source code: SunPos: Astro.c, Ofek Mathlab routines: SunCoo)////////////
'
'    Dim t As Double, t2 As Double, t3 As Double, l As Double, m As Double
'    Dim e As Double, ea As Double, v As Double, theta As Double, omega As Double, eps As Double
'
'    /* Time, in Julian centuries of 36525 ephemeris days,
'       measured from the epoch 1900 January 0.5 ET. */

    t = (jd - 2451545#) / 36525#
    t2 = t * t
    t3 = t2 * t
    
    ''Meeus expression 21.2 for the obliquity that is accurate to 1 second up to year 3000!
    'oblecl = ((23 + 26# / 60# + 21.448 / 3600#) - 46.815 * T / 360# - 0.0059 * t2 / 3600# + 0.001813 * t3 / 3600#) * cd

    'The geometric mean longitude of the Sun refered to the mean equinox of the date is:
    l = fnrev(280.46645 + 36000.76983 * t + 0.0003032 * t2)
    
    ms = l * cd
    
    'Sun's mean anomaly.
    ma = fnrev(357.5291 + 35999.0503 * t - 0.0001559 * t2 - 0.00000048 * t3)
    
    aas = ma * cd
    
    'Eccentricity of the Earth's orbit.
    e = 0.016708617 - 0.000042037 * t - 0.0000001236 * t2
    
    'Eccentricity anomaly
    ea = kepler(ma, e)
    
    '/* True anomaly */
    
    v = fnrev(2 * (Atn(Sqr((1 + e) / (1 - e)) * Tan(ea / 2))) / cd)
    
    '/* Sun's true longitude. */
    
    theta = fnrev(l + v - ma)
    
    '/* Corrections for Sun's apparent longitude, if desired. */
    
    '/* Obliquity of the ecliptic. */
    eps = obliqeq(jd)
    
    ob = eps * cd
    
    If (apparent) Then
       omega = fnrev(259.18 - 1934.142 * t)
       theta = theta - 0.00569 - 0.00479 * Sin(omega * cd)
       eps = eps + 0.00256 * Cos(omega * cd)
       End If
    
    '/* Return Sun's longitude and radius vector */
    
    SL = theta
    
    SLrad = SL * cd
        
    rv = (1.0000001018 * (1 - e * e)) / (1 + e * Cos(v * cd))
    
    '/* Determine solar co-ordinates (RA,Decl) in radians. */
    
    'right ascension
    rs = (fnrev(atan2(Cos(eps * cd) * Sin(theta * cd), Cos(theta * cd)) / cd)) * cd
    'declination
    Decl3 = DASIN(Sin(eps * cd) * Sin(theta * cd))
    

   On Error GoTo 0
   Exit Function

Decl3_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Decl3 of Module CalNewZemanim"

End Function

'---------------------------------------------------------------------------------------
' Procedure : DST_begend
' Author    : Dr-John-K-Hall
' Date      : 9/9/2016
' Purpose   : Extract beginning and ending daynumbers of DST for the relevant Hebrew year
'---------------------------------------------------------------------------------------
'
Function DST_begend(stryrDST%, endyrDST%, strdaynum1%, enddaynum1%, strdaynum2%, enddaynum2%) As Integer

    Dim myfile As String
    Dim pos1%, pos2%
        
   On Error GoTo DST_begend_Error

    strdaynum1% = 0
    enddaynum1% = 0
    strdaynum2% = 0
    enddaynum2% = 0
    
    DST_begend = 0

    'determine daynumber of beginning and ending civil year within the hebrew year
    myfile = Dir(App.Path & "\DST_EY.txt")
    If myfile <> sEmpty Then
       filin% = FreeFile
       Open App.Path & "\DST_EY.txt" For Input As #filin%
       Do Until EOF(filin%)
          'search for starting year
          Line Input #filin%, doclin$
          
          If InStr(doclin$, Trim$(Str$(stryrDST%))) <> 0 Then
             'parse the date of the onset of DST
             'each line looks like: 2015    Friday, March 27, 02:00 Sunday, October 25, 02:00
             pos1% = InStr(1, doclin$, "March")
             pos2% = InStr(pos1% + 1, doclin$, ",")
             If pos1% = 0 Or pos2% = 0 Then
                DST_begend = -1
                Exit Function
             Else
                strdaynum1% = Val(Mid$(doclin$, pos2% - 2, 2))
                End If
                
             pos1% = InStr(1, doclin$, "October")
             pos2% = InStr(pos1% + 1, doclin$, ",")
             If pos1% = 0 Or pos2% = 0 Then
                DST_begend = -1
                Exit Function
             Else
                enddaynum1% = Val(Mid$(doclin$, pos2% - 2, 2))
                End If
                
             End If
             
          If InStr(doclin$, Trim$(Str$(endyrDST%))) <> 0 Then
             'parse the date of the onset of DST
             'each line looks like: 2015    Friday, March 27, 02:00 Sunday, October 25, 02:00
             'parse the date of the onset of DST
             'each line looks like: 2015    Friday, March 27, 02:00 Sunday, October 25, 02:00
             pos1% = InStr(1, doclin$, "March")
             pos2% = InStr(pos1% + 1, doclin$, ",")
             If pos1% = 0 Or pos2% = 0 Then
                DST_begend = -1
                Exit Function
             Else
                strdaynum2% = Val(Mid$(doclin$, pos2% - 2, 2))
                End If
                
             pos1% = InStr(1, doclin$, "October")
             pos2% = InStr(pos1% + 1, doclin$, ",")
             If pos1% = 0 Or pos2% = 0 Then
                DST_begend = -1
                Exit Function
             Else
                enddaynum2% = Val(Mid$(doclin$, pos2% - 2, 2))
                End If

             End If
       
       Loop
       
       If strdaynum1% = 0 Or enddaynum1% = 0 Or _
          strdaynum2% = 0 Or enddaynum2% = 0 Then
          DST_begend = -1
          End If
    Else
    
       DST_begend = -1
       Exit Function
       End If
      
    

   On Error GoTo 0
   Exit Function

DST_begend_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DST_begend of Module CalNewZemanim"

End Function

Function DaysinYear(yrdy As Integer) As Integer

    'function calculates number of day in the civil year, yrdy
    
    Dim yd As Integer
    
    'determine if it is a leap year
    yd = yrdy - 1996
    DaysinYear = 365
    If yd Mod 4 = 0 Then DaysinYear = 366 'its a leap year
    'exclude century years that are not multiple of 400
    If yd Mod 4 = 0 And yrdy Mod 100 = 0 And yrdy Mod 400 <> 0 Then DaysinYear = 365
    
End Function

Function DayNumber(yljd As Integer, mon%, mday%) As Integer
   
   'determines day number for any month = mon%, day = mday%
   'yljd = 365 for regular year, 366 for leap year
   'based on Meeus' formula, p. 65
   
    Dim KK%
   
    KK% = 2
    If yljd = 366 Then KK% = 1
    DayNumber = (275 * mon%) \ 9 - KK * ((mon% + 9) \ 12) + mday% - 30
   
  
End Function
Public Sub casgeo(kmx, kmy, lg, lt)

        If kmy > 9999 Then
            G1# = kmy - 1000000
        Else
            G1# = kmy * 1000#
            End If
            
        If kmx < 9999 Then
            G2# = kmx * 1000#
        Else
            G2# = kmx
            End If
            
        r# = 57.2957795131
        B2# = 0.03246816
        f1# = 206264.806247096
        s1# = 126763.49
        S2# = 114242.75
        e4# = 0.006803480836
        C1# = 0.0325600414007
        C2# = 2.55240717534E-09
        c3# = 0.032338519783
        X1# = 1170251.56
        yy1# = 1126867.91
        yy2# = G1#
'       GN & GE
        X2# = G2#
        If (X2# > 700000!) Then GoTo ca5
        X1# = X1# - 1000000#
ca5:    If (yy2# > 550000#) Then GoTo ca10
        yy1# = yy1# - 1000000#
ca10:   X1# = X2# - X1#
        yy1# = yy2# - yy1#
        D1# = yy1# * B2# / 2#
        O1# = S2# + D1#
        O2# = O1# + D1#
        A3# = O1# / f1#
        A4# = O2# / f1#
        B3# = 1# - e4# * Sin(A3#) ^ 2
        B4# = B3# * Sqr(B3#) * C1#
        C4# = 1# - e4# * Sin(A4#) ^ 2
        C5# = Tan(A4#) * C2# * C4# ^ 2
        C6# = C5# * X1# ^ 2
        D2# = yy1# * B4# - C6#
        C6# = C6# / 3#
'LAT
        l1# = (S2# + D2#) / f1#
        R3# = O2# - C6#
        R4# = R3# - C6#
        R2# = R4# / f1#
        A2# = 1# - e4# * Sin(l1#) ^ 2
        lt = r# * (l1#)
        A5# = Sqr(A2#) * c3#
        d3# = X1# * A5# / Cos(R2#)
' LON
        lg = r# * ((s1# + d3#) / f1#)
'       THIS IS THE EASTERN HEMISPHERE!
        lg = -lg

End Sub

