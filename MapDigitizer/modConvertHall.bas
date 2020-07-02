Attribute VB_Name = "modConvertHall"
Public Enum eDatum
    eWGS84 = 0
    eGRS80 = 1
    eCLARK80M = 2
End Enum

Public Type datum
    a As Double   '// a  Equatorial earth radius
    b As Double   '// b  Polar earth radius
    f As Double   '// f= (a-b)/a  Flatenning
    esq As Double '// esq = 1-(b*b)/(a*a)  Eccentricity Squared
    E As Double   '// sqrt(esq)  Eccentricity
    '// deltas to WGS84
    dX As Double
    dy As Double
    dZ As Double
End Type

Public Enum e_grid
    gICS = 0
    gITM = 1
End Enum

Public Type grid
    lon0 As Double
    lat0 As Double
    k0 As Double
    false_e As Double
    false_n As Double
End Type

Public g_datum As datum
Public g_grid As grid

Public Function sin2(X As Double) As Double
   sin2 = Sin(X) * Sin(X)
End Function

Public Function cos2(X As Double) As Double
   cos2 = Cos(X) * Cos(X)
End Function

Public Function tan2(X As Double) As Double
   tan2 = Tan(X) * Tan(X)
End Function

Public Function tan4(X As Double) As Double
   tan4 = tan2(X) * tan2(X)
End Function

'static DATUM Datum[3] = {
'
'    // WGS84 data
'    {
'        6378137.0,              // a
'        6356752.3142,           // b
'        0.00335281066474748,    // f = 1/298.257223563
'        0.006694380004260807,   // esq
'        0.0818191909289062,     // e
'        // deltas to WGS84
'        0,
'        0,
'0
'    },
'
'    // GRS80 data
'    {
'        6378137.0,              // a
'        6356752.3141,           // b
'        0.0033528106811823,     // f = 1/298.257222101
'        0.00669438002290272,    // esq
'        0.0818191910428276,     // e
'        // deltas to WGS84
'        -48,
'        55,
'52
'    },
'
'    // Clark 1880 Modified data
'    {
'        6378300.789,            // a
'        6356566.4116309,        // b
'        0.003407549767264,      // f = 1/293.466
'        0.006803488139112318,   // esq
'        0.08248325975076590,    // e
'        // deltas to WGS84
'        -235,
'        -85,
'264
'    }
'};


    
'static GRID Grid[2] = {
'
'    // ICS data
'    {
'        0.6145667421719,            // lon0 = central meridian in radians of 35.12'43.490"
'        0.55386447682762762,        // lat0 = central latitude in radians of 31.44'02.749"
'        1.00000,                    // k0 = scale factor
'        170251.555,                 // false_easting
'        2385259.0                   // false_northing
'    },
'
'    // ITM data
'    {
'        0.61443473225468920,        // lon0 = central meridian in radians 35.12'16.261"
'        0.55386965463774187,        // lat0 = central latitude in radians 31.44'03.817"
'        1.0000067,                  // k0 = scale factor
'        219529.584,                 // false_easting
'        2885516.9488                // false_northing = 3512424.3388-626907.390
'                                    // MAPI says the false northing is 626907.390, and in another place
'                                    // that the meridional arc at the central latitude is 3512424.3388
'    }
'};


'//=================================================
'// Israel New Grid (ITM) to WGS84 conversion
'//=================================================
Public Sub itm2wgs84(N As Long, E As Long, lat As Double, lon As Double)
    '// 1. Local Grid (ITM) -> GRS80
    Dim lat80 As Double, lon80 As Double
    Call Grid2LatLon(N, E, lat80, lon80, gITM, eGRS80)

    '// 2. Molodensky GRS80->WGS84
    Dim lat84 As Double, lon84 As Double
    Call Molodensky(lat80, lon80, lat84, lon84, eGRS80, eWGS84)

    '// final results
    lat = lat84 * 180 / PI
    lon = lon84 * 180 / PI
End Sub


'//=================================================
'// WGS84 to Israel New Grid (ITM) conversion
'//=================================================
Public Sub wgs842itm(lat As Double, lon As Double, N As Long, E As Long)
    Dim latr As Double, lonr As Double
    latr = lat * PI / 180
    lonr = lon * PI / 180

    '// 1. Molodensky WGS84 -> GRS80
    Dim lat80 As Double, lon80 As Double
    Call Molodensky(latr, lonr, lat80, lon80, eWGS84, eGRS80)

    '// 2. Lat/Lon (GRS80) -> Local Grid (ITM)
    Call LatLon2Grid(lat80, lon80, N, E, eGRS80, gITM)

End Sub


'//=================================================
'// Israel Old Grid (ICS) to WGS84 conversion
'//=================================================
Public Sub ics2wgs84(N As Long, E As Long, lat As Double, lon As Double)
    '// 1. Local Grid (ICS) -> Clark_1880_modified
    Dim lat80 As Double, lon80 As Double
    '//printf("inside N, E = %d, %d", N, E);
    '//pause();
    Call Grid2LatLon(N, E, lat80, lon80, gICS, eCLARK80M)
    '//printf("grid lat, lon = %f, %f", lat80*180/pi(), lon80*180/pi());
    '//pause();

    '// 2. Molodensky Clark_1880_modified -> WGS84
    Dim lat84 As Double, lon84 As Double
    Call Molodensky(lat80, lon80, lat84, lon84, eCLARK80M, eWGS84)

    '// final results
    lat = lat84 * 180 / PI
    lon = lon84 * 180 / PI
    '//printf("final lat, lon = %f, %f", lat, lon);
    '//pause();
End Sub


'//=================================================
'// WGS84 to Israel Old Grid (ICS) conversion
'//=================================================
Public Sub wgs842ics(lat As Double, lon As Double, N As Long, E As Long)
    Dim latr As Double, lonr As Double
    latr = lat * PI / 180
    lonr = lon * PI / 180

    '// 1. Molodensky WGS84 -> Clark_1880_modified
    Dim lat80 As Double, lon80 As Double
    Call Molodensky(latr, lonr, lat80, lon80, eWGS84, eCLARK80M)

    '// 2. Lat/Lon (Clark_1880_modified) -> Local Grid (ICS)
    Call LatLon2Grid(lat80, lon80, N, E, eCLARK80M, gICS)
End Sub


'//====================================
'// Local Grid to Lat/Lon conversion
'//====================================
Public Sub Grid2LatLon(North As Long, East As Long, lat As Double, lon As Double, gGridfrom As e_grid, eDatumto As eDatum)

'    //================
'    // GRID -> Lat/Lon
'    //================
'    //printf("gridfrom.falsen = %f", Grid[from].false_n);
'    //printf("datumto.a = %f", Datum[to].a);
'    //pause();

    Call LoadGrid(gGridfrom, eDatumto)
    
    Dim Y As Double, X As Double, m As Double
    Y = North + g_grid.false_n
    X = East - g_grid.false_e
    m = Y / g_grid.k0

    Dim a As Double, b As Double, E As Double, esq As Double
    a = g_datum.a
    b = g_datum.b
    E = g_datum.E
    esq = g_datum.esq

    Dim mu As Double
    mu = m / (a * (1 - E * E / 4 - 3 * (E ^ 4) / 64# - 5 * (E ^ 6) / 256#))

    Dim ee As Double, e1 As Double, j1 As Double, j2 As Double, j3 As Double, j4 As Double
    ee = Sqr(1 - esq)
    e1 = (1 - ee) / (1 + ee)
    j1 = 3 * e1 / 2 - 27 * e1 * e1 * e1 / 32
    j2 = 21 * e1 * e1 / 16 - 55 * e1 * e1 * e1 * e1 / 32
    j3 = 151 * e1 * e1 * e1 / 96
    j4 = 1097 * e1 * e1 * e1 * e1 / 512

    '// Footprint Latitude
    Dim fp As Double
    fp = mu + j1 * Sin(2 * mu) + j2 * Sin(4 * mu) + j3 * Sin(6 * mu) + j4 * Sin(8 * mu)

    Dim sinfp As Double, cosfp As Double, tanfp As Double, eg As Double, eg2 As Double
    Dim C1 As Double, T1 As Double, R1 As Double, N1 As Double, D As Double
    sinfp = Sin(fp)
    cosfp = Cos(fp)
    tanfp = sinfp / cosfp
    eg = (E * a / b)
    eg2 = eg * eg
    C1 = eg2 * cosfp * cosfp
    T1 = tanfp * tanfp
    R1 = a * (1 - E * E) / ((1 - (E * sinfp) * (E * sinfp)) ^ 1.5)
    N1 = a / Sqr(1 - (E * sinfp) * (E * sinfp))
    D = X / (N1 * g_grid.k0)

    Dim Q1 As Double, Q2 As Double, Q3 As Double, Q4 As Double
    Q1 = N1 * tanfp / R1
    Q2 = D * D / 2
    Q3 = (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * eg2 * eg2) * (D * D * D * D) / 24
    Q4 = (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 3 * C1 * C1 - 252 * eg2 * eg2) * (D * D * D * D * D * D) / 720
    '// result lat
    lat = fp - Q1 * (Q2 - Q3 + Q4)

    Dim Q5 As Double, Q6 As Double, Q7 As Double
    Q5 = D
    Q6 = (1 + 2 * T1 + C1) * (D * D * D) / 6
    Q7 = (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * eg2 * eg2 + 24 * T1 * T1) * (D * D * D * D * D) / 120
    '// result lon
    lon = g_grid.lon0 + (Q5 - Q6 + Q7) / cosfp

End Sub


'//====================================
'// Lat/Lon to Local Grid conversion
'//====================================
Public Sub LatLon2Grid(lat As Double, lon As Double, North As Long, East As Long, eDatumfrom As eDatum, gGridto As e_grid)
    
    Call LoadGrid(gGridto, eDatumfrom)
    
    '// Datum data for Lat/Lon to TM conversion
    Dim a As Double, E As Double, b As Double
    a = g_datum.a
    E = g_datum.E   '// sqrt(esq);
    b = g_datum.b

'    //===============
'    // Lat/Lon -> TM
'    //===============
    Dim slat1 As Double, clat1 As Double, clat1sq As Double, tanlat1sq As Double
    Dim e2 As Double, e4 As Double, e6 As Double, eg As Double, eg2 As Double
    slat1 = Sin(lat)
    clat1 = Cos(lat)
    clat1sq = clat1 * clat1
    tanlat1sq = slat1 * slat1 / clat1sq
    e2 = E * E
    e4 = e2 * e2
    e6 = e4 * e2
    eg = (E * a / b)
    eg2 = eg * eg

    Dim l1 As Double, l2 As Double, l3 As Double, l4 As Double
    l1 = 1 - e2 / 4 - 3 * e4 / 64 - 5 * e6 / 256
    l2 = 3 * e2 / 8 + 3 * e4 / 32 + 45 * e6 / 1024
    l3 = 15 * e4 / 256 + 45 * e6 / 1024
    l4 = 35 * e6 / 3072
    m = a * (l1 * lat - l2 * Sin(2 * lat) + l3 * Sin(4 * lat) - l4 * Sin(6 * lat))
    '//double rho = a*(1-e2) / pow((1-(e*slat1)*(e*slat1)),1.5);
    Dim nu As Double, p As Double, k0 As Double
    nu = a / Sqr(1 - (E * slat1) * (E * slat1))
    p = lon - g_grid.lon0
    k0 = g_grid.k0
    '// y = northing = K1 + K2p2 + K3p4, where
    Dim K1 As Double, K2 As Double, K3 As Double
    K1 = m * k0
    K2 = k0 * nu * slat1 * clat1 / 2
    K3 = (k0 * nu * slat1 * clat1 * clat1sq / 24) * (5 - tanlat1sq + 9 * eg2 * clat1sq + 4 * eg2 * eg2 * clat1sq * clat1sq)
    '// ING north
    Dim Y As Double
    Y = K1 + K2 * p * p + K3 * p * p * p * p - g_grid.false_n

    '// x = easting = K4p + K5p3, where
    Dim K4 As Double, K5 As Double
    K4 = k0 * nu * clat1
    K5 = (k0 * nu * clat1 * clat1sq / 6) * (1 - tanlat1sq + eg2 * clat1 * clat1)
    '// ING east
    Dim X As Double
    X = K4 * p + K5 * p * p * p + g_grid.false_e

    '// final rounded results
    East = CLng(X + 0.5)
    North = CLng(Y + 0.5)

End Sub


'//======================================================
'// Abridged Molodensky transformation between 2 datums
'//======================================================
Public Sub Molodensky(ilat As Double, ilon As Double, olat As Double, olon As Double, eDatumfrom As eDatum, eDatumto As eDatum)

    '// from->WGS84 - to->WGS84 = from->WGS84 + WGS84->to = from->to
    Dim dX As Double, dy As Double, dZ As Double
    Call LoadGrid(-1, eDatumfrom)
    dX = g_datum.dX
    dy = g_datum.dy
    dZ = g_datum.dZ
    Call LoadGrid(-1, eDatumto)
    dX = dX - g_datum.dX
    dy = dy - g_datum.dy
    dZ = dZ - g_datum.dZ
    
    Dim slat As Double, clat As Double, slon As Double, clon As Double, ssqlat As Double
    slat = Sin(ilat)
    clat = Cos(ilat)
    slon = Sin(ilon)
    clon = Cos(ilon)
    ssqlat = slat * slat

    '//dlat = ((-dx * slat * clon - dy * slat * slon + dz * clat)
    '//        + (da * rn * from_esq * slat * clat / from_a)
    '//        + (df * (rm * adb + rn / adb )* slat * clat))
    '//       / (rm + from.h);

    Dim from_f As Double, df As Double, from_a As Double, da As Double, from_esq As Double, adb, rn As Double, rm As Double, from_h As Double
    Call LoadGrid(-1, eDatumfrom)
    from_f = g_datum.f
    Call LoadGrid(-1, eDatumto)
    df = g_datum.f - from_f
    Call LoadGrid(-1, eDatumfrom)
    from_a = g_datum.a
    Call LoadGrid(-1, eDatumto)
    da = g_datum.a - from_a
    Call LoadGrid(-1, eDatumfrom)
    from_esq = g_datum.esq
    adb = 1# / (1# - from_f)
    rn = from_a / Sqr(1 - from_esq * ssqlat)
    rm = from_a * (1 - from_esq) / ((1 - from_esq * ssqlat) ^ 1.5)
    from_h = 0#  '; // we're flat!

    Dim dlat As Double
    dlat = (-dX * slat * clon - dy * slat * slon + dZ * clat _
                   + da * rn * from_esq * slat * clat / from_a + _
                   df * (rm * adb + rn / adb) * slat * clat) / (rm + from_h)

    '// result lat (radians)
    olat = ilat + dlat

    '// dlon = (-dx * slon + dy * clon) / ((rn + from.h) * clat);
    Dim dlon As Double
    dlon = (-dX * slon + dy * clon) / ((rn + from_h) * clat)
    '// result lon (radians)
    olon = ilon + dlon

End Sub


Public Sub LoadGrid(igrid As e_grid, idatum As eDatum)

    Select Case idatum
        Case 0
        '// WGS84 data
            g_datum.a = 6378137# ',              // a
            g_datum.b = 6356752.3142 ',           // b
            g_datum.f = 3.35281066474748E-03 ',    // f = 1/298.257223563
            g_datum.esq = 6.69438000426081E-03 ',   // esq
            g_datum.E = 8.18191909289062E-02 ',     // e
            '// deltas to WGS84
            g_datum.dX = 0 ',
            g_datum.dy = 0 ',
            g_datum.dZ = 0 '
        
        Case 1
        '// GRS80 data
            g_datum.a = 6378137# ',              // a
            g_datum.b = 6356752.3141 ',           // b
            g_datum.f = 3.3528106811823E-03 ',     // f = 1/298.257222101
            g_datum.esq = 6.69438002290272E-03 ',    // esq
            g_datum.E = 8.18191910428276E-02 ',     // e
            '// deltas to WGS84
            g_datum.dX = -48 ',
            g_datum.dy = 55 ',
            g_datum.dZ = 52
        Case 2
        '// Clark 1880 Modified data
            g_datum.a = 6378300.789 ',            // a
            g_datum.b = 6356566.4116309 ',        // b
            g_datum.f = 0.003407549767264 ',      // f = 1/293.466
            g_datum.esq = 6.80348813911232E-03 ',   // esq
            g_datum.E = 8.24832597507659E-02 ',    // e
            '// deltas to WGS84
            g_datum.dX = -235 ',
            g_datum.dy = -85 ',
            g_datum.dZ = 264
         Case Else
    End Select

    Select Case igrid
        Case 0
        '// ICS data
            g_grid.lon0 = 0.6145667421719 ',            // lon0 = central meridian in radians of 35.12'43.490"
            g_grid.lat0 = 0.553864476827628  ',        // lat0 = central latitude in radians of 31.44'02.749"
            g_grid.k0 = 1#     ',                    // k0 = scale factor
            g_grid.false_e = 170251.555 ',                 // false_easting
            g_grid.false_n = 2385259# '                   // false_northing
        Case 1
        '// ITM data
            g_grid.lon0 = 0.614434732254689  ',        // lon0 = central meridian in radians 35.12'16.261"
            g_grid.lat0 = 0.553869654637742  ',        // lat0 = central latitude in radians 31.44'03.817"
            g_grid.k0 = 1.0000067 ',                  // k0 = scale factor
            g_grid.false_e = 219529.584 ',                 // false_easting
            g_grid.false_n = 2885516.9488 '                // false_northing = 3512424.3388-626907.390
                                        '// MAPI says the false northing is 626907.390, and in another place
                                        '// that the meridional arc at the central latitude is 3512424.3388
        Case Else
        
    End Select
    

End Sub
Public Sub DTMheight2(kmx As Double, kmy As Double, hgt2 As Integer)
'This routine reads the DTM at a coordinate (kmx=ITMx,kym=ITMy)
'and returns the height, hgt2, at that point
Dim lg2 As Double
Dim lt As Double

   If DTMtype = 1 Then
      'convert itm coorinates to wgs34 geo
      Call ics2wgs84(CLng(kmy), CLng(kmx), lt, lg2)
      
      'determine if bil file is already open
      If ASTERbilOpen Then
         If lg2 >= ASTEREast And lg2 <= ASTEREast + 1 Then
            If lt >= ASTERNorth And lt <= ASTERNorth + 1 Then
               'the height resides on the same tile
               GoTo g10
               End If
            End If
         End If
         
 '-----------------------------------------------
      'Else need to open new tile
      'determine root name from bottom left coordinates
      ASTERNorth = Int(lt) '-35.3 -> -36, 35.3 -> 35
      If lt >= 0 Then 'North latitude
         ltch$ = "N"
      ElseIf lt < 0 Then 'South latitude
         ltch$ = "S"
         End If
      ASTEREast = Int(lg2) '35.3 ->35, -118.3 -> -119
      If lg2 >= 0 Then 'east longitue
         lgch$ = "E"
      ElseIf lg2 < 0 Then 'West longitude
         lgch$ = "W"
         End If
      
      ASTERfilename = ltch$ & Format(Trim$(str$(Abs(ASTERNorth))), "00") & lgch$ & Format(Trim$(str$(Abs(ASTEREast))), "000")
    
      If ASTERfilnum% <> 0 Then Close #ASTERfilnum% 'close last file and open new one
      ASTERbilOpen = False
       
      filin1% = FreeFile
      If Dir(ASTERdir & "\" & ASTERfilename & ".bil") <> sEmpty And Dir(ASTERdir & "\" & ASTERfilename & ".hdr") <> sEmpty Then
         filin1% = FreeFile
         Open ASTERdir + "\" + ASTERfilename + ".hdr" For Input As #filin1%
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
            
       Else
         hgt2 = -9999
         Exit Sub
         End If
   
      'now open the new bil file
      ASTERfilnum% = FreeFile
      Open ASTERdir + "\" + ASTERfilename + ".bil" For Binary As #ASTERfilnum%
      ASTERbilOpen = True
        
g10:
      'read a height
      IKMY% = CInt((lt - ASTERNorth) / ASTERydim) + 1
      IKMX% = CInt((lg2 - ASTEREast) / ASTERxdim) + 1
      tncols& = ASTERNcols%
      tnrows& = ASTERNrows%
      numrec& = (tnrows& - IKMY%) * tncols& + IKMX%
      Get #ASTERfilnum%, (numrec& - 1) * 2 + 1, IO%
      hgt2 = IO%
        
 '-------------------------------------------------
      
   ElseIf DTMtype = 2 Then
   
        On Error GoTo g35
        
        kmx = kmx * 0.001
        kmy = (kmy - 1000000) * 0.001
        IKMX% = Int((kmx + 20!) * 40!) + 1
        IKMY% = Int((380! - kmy) * 40!) + 1
        NRow% = IKMY%: NCol% = IKMX%

'       GETZ FINDS THE HEIGHT OF A POINT AT THE NORW AND NCOL FROM 380N
'       AND -20E WHERE 1,1 IS THAT CORNER POINT
'       FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
g15:    Jg% = 1 + Int((NRow% - 2) / 800)
        Ig% = 1 + Int((NCol% - 2) / 800)
        CHMNE = CHMAP(Ig%, Jg%)
        If CHMNE = "  " Then GoTo g35
        If CHMNE = CHMNEO Then GoTo g21
        jj% = filnumg%
        Close #filnumg%
        SF = CHMNE
        Sffnam$ = dtmdir + "\" + SF
        filnumg% = FreeFile
        Open Sffnam$ For Random As #filnumg% Len = 2
        CHMNEO = CHMNE
'       CONVERT TO GRID LOCATION IN .SUM FILE
g21:    IR% = NRow% - (Jg% - 1) * 800
        IC% = NCol% - (Ig% - 1) * 800
        IFN& = (IR% - 1) * 801! + IC%
        Get #filnumg%, IFN&, IO%
        hgt2 = IO% * 0.1
        If hgt2 < -1000 Then hgt2 = -9999
        GoTo g99
g35:    er% = Err.Number
        hgt2 = -9999 'MsgBox " ERROR IN GETZ ", vbCritical + vbOKOnly, "SkyLight"
g99:
        End If

End Sub

'4320
'2.314814815 e-4
Public Sub ASTERheight(lt As Double, lg2 As Double, hgt2 As Integer)

      'determine if bil file is already open
      If ASTERbilOpen Then
         If lg2 >= ASTEREast And lg2 <= ASTEREast + 1 Then
            If lt >= ASTERNorth And lt <= ASTERNorth + 1 Then
               'the height resides on the same tile
               GoTo g10
               End If
            End If
         End If
         
 '-----------------------------------------------
      'Else need to open new tile
      'determine root name from bottom left coordinates
      ASTERNorth = Int(lt) '-35.3 -> -36, 35.3 -> 35
      If lt >= 0 Then 'North latitude
         ltch$ = "N"
      ElseIf lt < 0 Then 'South latitude
         ltch$ = "S"
         End If
      ASTEREast = Int(lg2) '35.3 ->35, -118.3 -> -119
      If lg2 >= 0 Then 'east longitue
         lgch$ = "E"
      ElseIf lg2 < 0 Then 'West longitude
         lgch$ = "W"
         End If
      
      ASTERfilename = ltch$ & Format(Trim$(str$(Abs(ASTERNorth))), "00") & lgch$ & Format(Trim$(str$(Abs(ASTEREast))), "000")
    
      If ASTERfilnum% <> 0 Then Close #ASTERfilnum% 'close last file and open new one
      ASTERbilOpen = False
       
      filin1% = FreeFile
      If Dir(ASTERdir & "\" & ASTERfilename & ".bil") <> sEmpty And Dir(ASTERdir & "\" & ASTERfilename & ".hdr") <> sEmpty Then
         filin1% = FreeFile
         Open ASTERdir + "\" + ASTERfilename + ".hdr" For Input As #filin1%
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
            
       Else
         hgt2 = -9999
         Exit Sub
         End If
   
      'now open the new bil file
      ASTERfilnum% = FreeFile
      Open ASTERdir + "\" + ASTERfilename + ".bil" For Binary As #ASTERfilnum%
      ASTERbilOpen = True
        
g10:
      'read a height
      IKMY% = CInt((lt - ASTERNorth) / ASTERydim) + 1
      IKMX% = CInt((lg2 - ASTEREast) / ASTERxdim) + 1
      tncols& = ASTERNcols%
      tnrows& = ASTERNrows%
      numrec& = (tnrows& - IKMY%) * tncols& + IKMX%
      Get #ASTERfilnum%, (numrec& - 1) * 2 + 1, IO%
      hgt2 = IO%

End Sub

Public Sub worldheights(lg As Double, lt As Double, hgt As Integer)
   Dim leros As Long, lmag As Long
   On Error GoTo worlderror
   
   Dim nrows As Integer, ncols As Integer
   
   If lt > 90 Or lt < -90 Or lg < -180 Or lg > 180 Then Exit Sub
   
      
   'check if have correct CD in the drive, if not present error message
'   If (world = False And IsraelDTMsource% = 1) Or (DTMflag > 0 And (lt >= -60 And lt <= 61)) Then 'SRTM
'
'      If world = False And IsraelDTMsource% = 1 Then
'         'use 90-m SRTM of Eretz Yisroel
'         xdim = 8.33333333333333E-04
'         ydim = 8.33333333333333E-04
'         lg = -lg
'         DEMfile$ = israeldtm + ":\dtm\"
'         nrows = 1201
'         ncols = 1201
'         GoTo wh50
'         End If
'
'      If DTMflag = 1 And Dir(srtmdtm & ":/USA/", vbDirectory) <> sEmpty Then
         xdim = 8.33333333333333E-04 / 3#
         ydim = 8.33333333333333E-04 / 3#
         DEMfile$ = NEDdir & "\"
         nrows = 3601
         ncols = 3601
'      ElseIf DTMflag = 2 And Dir(srtmdtm & ":/3AS/", vbDirectory) <> sEmpty Then
'         xdim = 8.33333333333333E-04
'         ydim = 8.33333333333333E-04
'         DEMfile$ = srtmdtm & ":/3AS/"
'         nrows = 1201
'         ncols = 1201
'         End If
wh50:
      'determine tile name
      lg1 = Int(lg)
      If lg1 < 0 And lg1 > lg Then lg1 = lg1 - 1
      If lg1 < 0 Then EWch$ = "W" Else EWch$ = "E"
      If Abs(lg1) < 10 Then
         lg1ch$ = "00" & Trim$(str$(Abs(lg1)))
      ElseIf Abs(lg1) >= 10 And Abs(lg1) < 100 Then
         lg1ch$ = "0" & Trim$(str$(Abs(lg1)))
      ElseIf Abs(lg1) >= 100 Then
         lg1ch$ = Trim$(str$(Abs(lg1)))
         End If
      lt1 = Int(lt) 'SRTM tiles are named by SW corner
      If lt1 < 0 And lt1 > lt Then lt1 = lt1 - 1
      If lt1 < 0 Then NSch$ = "S" Else NSch$ = "N"
      If Abs(lt1) < 10 Then
         lt1ch$ = "0" & Trim$(str$(Abs(lt1)))
      ElseIf Abs(lt1) >= 10 Then
         lt1ch$ = Trim$(str$(Abs(lt1)))
         End If
      lt1 = lt1 + 1 'the first record in SRTM tiles in the NW corner
      DEMfile$ = DEMfile$ & NSch$ & lt1ch$ & EWch$ & lg1ch$ & ".hgt"
      If Dir(DEMfile$) = sEmpty Then
         GoTo gtopo
         'mapEROSDTMwarn.Visible = True
         'ret = SetWindowPos(mapEROSDTMwarn.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         'mapEROSDTMwarn.Label3.Caption = sEmpty
         'mapEROSDTMwarn.Label2.Caption = DEMfile$
         'leros = FindWindow(vbNullString, "         USGS EROS DEM CD not found!")
         'If leros > 0 Then
         '   ret = BringWindowToTop(leros) 'bring message to top
         '   End If
      Else
      
         worldfnum% = FreeFile
         Open DEMfile$ For Binary As #worldfnum%
         GoSub Eroshgt
         Close #worldfnum%
         worldfnum% = 0
         hgt = integ2%
         If hgt = -32768 Then hgt = -9999 'void
         Exit Sub
         End If
'       End If

gtopo:
   hgt = -9999
   Exit Sub
   
'   xdim = 8.33333333333333E-03
'   ydim = 8.33333333333333E-03
'   If lt > -60 Then
'      nx% = Fix((lg + 180) * 0.025)
'      lg1 = -180 + nx% * 40
'      If Abs(lg1) >= 100 Then
'         lg1ch$ = RTrim$(LTrim$(str$(Abs(lg1))))
'      Else
'         lg1ch$ = "0" + RTrim$(LTrim$(str$(Abs(lg1))))
'         End If
'      If lg1 < 0 Then
'         EW$ = "W"
'      Else
'         EW$ = "E"
'         End If
'      ny% = Int((90 - lt) * 0.02)
'      lt1 = 90 - 50 * ny%
'      lt1ch$ = LTrim$(RTrim$(str$(Abs(lt1))))
'      If lt1 > 0 Then
'         ns$ = "N"
'      Else
'         ns$ = "S"
'         End If
'      DEMfile0$ = EW$ + lg1ch$ + ns$ + lt1ch$
'      DEMfile1$ = worlddtm + ":\" + DEMfile0$ + "\" + DEMfile0$
'      DEMfile$ = DEMfile1$ + ".dem"
'      nrows = 6000
'      ncols = 4800
'      numCD% = worldCD%(ny% * 9 + nx% + 1)
'   Else 'Antartic - Cd #5
'      nx% = Fix((lg + 180) / 60)
'      lg1 = -180 + nx% * 60
'      If Abs(lg1) >= 100 Then
'         lg1ch$ = LTrim$(RTrim$(str$(Abs(lg1))))
'      ElseIf Abs(lg1) < 100 And Abs(lg1) <> 0 Then
'         lg1ch$ = "0" + RTrim$(LTrim$(str$(Abs(lg1))))
'      ElseIf Abs(lg1) = 0 Then
'         lg1ch$ = "000"
'         End If
'      If lg1 <= 0 Then
'         EW$ = "W"
'      Else
'         EW$ = "E"
'         End If
'      ns$ = "S"
'      lt1 = -60
'      lt1ch$ = "60"
'      DEMfile0$ = EW$ + lg1ch$ + ns$ + lt1ch$
'      DEMfile1$ = worlddtm + ":\" + DEMfile0$ + "\" + DEMfile0$
'      DEMfile$ = DEMfile1$ + ".dem"
'      nrows = 3600
'      ncols = 7200
'      numCD% = 5
'      End If
'   If worldfil$ <> DEMfile1$ Then
'      myfile = Dir(DEMfile$)
'      If myfile = sEmpty Then
'         mapEROSDTMwarn.Visible = True
'         ret = SetWindowPos(mapEROSDTMwarn.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'         mapEROSDTMwarn.Label3.Caption = numCD%
'         leros = FindWindow(vbNullString, "         USGS EROS DEM CD not found!")
'         If leros > 0 Then
'            ret = BringWindowToTop(leros) 'bring message to top
'            End If
'      Else
'         If mapEROSDTMwarn.Visible = True Then
'            Unload mapEROSDTMwarn
'            Set skyerosdtwarn = Nothing
'            If magbox = True Then
'               lmag = FindWindow(vbNullString, mapMAGfm.Caption)
'               If lmag > 0 Then
'                  ret = BringWindowToTop(lmag) 'bring mapMAGfm back to top of Z order
'                  'ret = ShowWindow(lmag, SW_RESTORE) 'redisplay mapMAGfm
'                  End If
'               End If
'            End If
'         If worldfnum% <> 0 Then Close #worldfnum%
'         '******set as constants
'         'worldfnum% = FreeFile
'         'worldfil$ = DEMfile1$
'         'Open DEMfile1$ + ".STX" For Input As #worldfnum%
'         'Input #worldfnum%, A, elevmin%, elevmax%, D, E
'         'Close #worldfnum%
'         'Open DEMfile1$ + ".HDR" For Input As #worldfnum%
'         'npos% = 0
'         'Do Until EOF(worldfnum%)
'         '  npos% = npos% + 1
'         '  Line Input #worldfnum%, doclin$
'         '  If npos% = 3 Then
'         '     nrows% = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
'         '  ElseIf npos% = 4 Then
'         '     ncols% = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
'         '  ElseIf npos% = 13 Then
'         '     xdim = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
'         '  ElseIf npos% = 14 Then
'         '     ydim = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
'         '     End If
'        'Loop
'        'Close #worldfnum%
'        worldfnum% = FreeFile
'        Open DEMfile$ For Binary As #worldfnum%
'        GoSub Eroshgt
'        Close #worldfnum%
'        worldfnum% = 0
'        hgt = integ2%
'        End If
'    Else
'       If mapEROSDTMwarn.Visible = True Then
'          Unload mapEROSDTMwarn
'          Set skyerosdtwarn = Nothing
'          End If
'       If magbox = True Then
'          lmag = FindWindow(vbNullString, mapMAGfm.Caption)
'          If lmag > 0 Then
'             ret = BringWindowToTop(lmag) 'bring mapMAGfm back to top of Z order
''             ret = ShowWindow(lmag, SW_RESTORE) 'redisplay mapMAGfm
'             End If
'          End If
'       'continue reading
'        GoSub Eroshgt
'        hgt = integ2%
'        End If
'    Exit Sub

Eroshgt:
'   IKMY& = CInt(((lt1 - ydim * 0.5) - lt) / ydim) + 1
'   IKMX& = CInt((lg - (lg1 + xdim * 0.5)) / xdim) + 1
   IKMY& = CLng((lt1 - lt) / ydim) + 1
   IKMX& = CLng((lg - lg1) / xdim) + 1
   
'   GDMDIform.StatusBar1.Panels(1).Text = "lt1,lt,lg,lg1,IKMY,IKMX = " & lt1 & ", " & lt & ", " & lg1 & ", " & lg & ", " & IKMY& & ", " & IKMX&
   
   tncols = ncols
   c% = worldfnum%
   numrec& = (IKMY& - 1) * tncols + IKMX&
   Get #worldfnum%, (numrec& - 1) * 2 + 1, IO%
'   A$ = sEmpty
'   A$ = Hex$(io%)
   'first attempt to swap bytes the fattest way--i.e.,
   'by modular division by 256 (= 100) (since the first byte, i.e.,
   'the first two bits, represent integers in the range 0 to 255)
   '(this fails for negative integers due to the way negative integers
   'are represented, as detailed later).
    If IO% < 0 Then GoTo mer130 'then modular division failed, use HEX swap
    T1 = IO% Mod 256
    T2 = Int(IO% / 256)
    tr = T1 * 256 + T2
    integ1& = tr
mer130:
    If IO% < 0 Or integ1& > elevmax% Then 'modular division failed use HEX swap
       A0$ = LTrim$(RTrim$(Hex$(IO%)))
       AA$ = sEmpty
       'swap the two bytes using their hex representation
       'e.g., ABCD --> CDAB, etc.
       If Len(A0$) = 4 Then
          A1$ = Mid$(A0$, 1, 2)
          A2$ = Mid$(A0$, 3, 2)
          If Mid$(A2$, 3, 1) = "0" And Mid$(A2$, 3, 2) <> "0" Then
             A2$ = Mid$(A0$, 4, 1)
          ElseIf Mid$(A2$, 3, 1) = "0" And Mid$(A2$, 3, 2) = "0" Then
             A2$ = sEmpty
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
        For j% = leng% To 1 Step -1
            V$ = Mid$(LTrim$(RTrim$(AA$)), j%, 1)
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
           If j% = leng% - 3 Then
              integ1& = integ1& + 4096 * NO&
           ElseIf j% = leng% - 2 Then
              integ1& = integ1& + 256 * NO&
           ElseIf j% = leng% - 1 Then
              integ1& = integ1& + 16 * NO&
           ElseIf j% = leng% Then
              integ1& = integ1& + NO&
              End If
        Next j%
        'positive 2 byte integers are stored as numbers 1 to 32767.
        'negative 2 byte integers are stored as numbers
        'greater than 32767 (since 2 byte, i.e.,  8 bits encompass
        'the integer range -32768 to 32767), where -1 is 65535 and
        '-2 is 65534, etc up to -32768 which is represented
        'as 32768, i.e.,
        If integ1& > 32767 Then integ1& = integ1& - 65536
    End If
    integ2% = integ1&
Return

worlderror:
   If routeload = True Or travelmode = True Then
      hgt = 0
      Exit Sub
      End If
   Ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   response = MsgBox("An error in reading the CD has occured! Do you wish to try again?", vbCritical + vbRetryCancel, "Maps & More")
   Ret = SetWindowPos(mapPictureform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   If response = vbCancel Then Exit Sub
   Resume
End Sub
Public Sub DTMheight(kmx, kmy, hgt2)
'This routine reads the DTM at a coordinate (kmx=ITMx,kym=ITMy)
'and returns the height, hgt2, at that point
Dim lg2 As Double

   If DTMtype = 1 Then
      'convert itm coorinates to wgs34 geo
      ll1% = 0
      If Not GpsCorrection Then ll1% = 1 'flag to return to Clark geoid after height determination
      GpsCorrection = True 'always use WGS84 for ASTER geo coordinates
      Call casgeo(kmx, kmy, lg, lt)
      lg2 = -lg 'convert East longitue to positive, West longitude to negative
      If ll1% = 1 Then GpsCorrection = False 'return to Clark geoid
      
      'determine if bil file is already open
      If ASTERbilOpen Then
         If lg2 >= ASTEREast And lg2 <= ASTEREast + 1 Then
            If lt >= ASTERNorth And lt <= ASTERNorth + 1 Then
               'the height resides on the same tile
               GoTo g10
               End If
            End If
         End If
         
 '-----------------------------------------------
      'Else need to open new tile
      'determine root name from bottom left coordinates
      ASTERNorth = Int(lt) '-35.3 -> -36, 35.3 -> 35
      If lt >= 0 Then 'North latitude
         ltch$ = "N"
      ElseIf lt < 0 Then 'South latitude
         ltch$ = "S"
         End If
      ASTEREast = Int(lg2) '35.3 ->35, -118.3 -> -119
      If lg2 >= 0 Then 'east longitue
         lgch$ = "E"
      ElseIf lg2 < 0 Then 'West longitude
         lgch$ = "W"
         End If
      
      ASTERfilename = ltch$ & Format(Trim$(str$(Abs(ASTERNorth))), "00") & lgch$ & Format(Trim$(str$(Abs(ASTEREast))), "000")
    
      If ASTERfilnum% <> 0 Then Close #ASTERfilnum% 'close last file and open new one
      ASTERbilOpen = False
       
      filin1% = FreeFile
      If Dir(ASTERdir & "\" & ASTERfilename & ".bil") <> sEmpty And Dir(ASTERdir & "\" & ASTERfilename & ".hdr") <> sEmpty Then
         filin1% = FreeFile
         Open ASTERdir + "\" + ASTERfilename + ".hdr" For Input As #filin1%
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
            
       Else
         hgt2 = -9999
         Exit Sub
         End If
   
      'now open the new bil file
      ASTERfilnum% = FreeFile
      Open ASTERdir + "\" + ASTERfilename + ".bil" For Binary As #ASTERfilnum%
      ASTERbilOpen = True
        
g10:
      'read a height
      IKMY% = CInt((lt - ASTERNorth) / ASTERydim) + 1
      IKMX% = CInt((lg2 - ASTEREast) / ASTERxdim) + 1
      tncols& = ASTERNcols%
      tnrows& = ASTERNrows%
      numrec& = (tnrows& - IKMY%) * tncols& + IKMX%
      Get #ASTERfilnum%, (numrec& - 1) * 2 + 1, IO%
      hgt2 = IO%
        
 '-------------------------------------------------
      
   ElseIf DTMtype = 2 Then
   
        On Error GoTo g35
        
        kmx = kmx * 0.001
        kmy = (kmy - 1000000) * 0.001
        IKMX% = Int((kmx + 20!) * 40!) + 1
        IKMY% = Int((380! - kmy) * 40!) + 1
        NRow% = IKMY%: NCol% = IKMX%

'       GETZ FINDS THE HEIGHT OF A POINT AT THE NORW AND NCOL FROM 380N
'       AND -20E WHERE 1,1 IS THAT CORNER POINT
'       FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
g15:    Jg% = 1 + Int((NRow% - 2) / 800)
        Ig% = 1 + Int((NCol% - 2) / 800)
        CHMNE = CHMAP(Ig%, Jg%)
        If CHMNE = "  " Then GoTo g35
        If CHMNE = CHMNEO Then GoTo g21
        jj% = filnumg%
        Close #filnumg%
        SF = CHMNE
        Sffnam$ = dtmdir + "\" + SF
        filnumg% = FreeFile
        Open Sffnam$ For Random As #filnumg% Len = 2
        CHMNEO = CHMNE
'       CONVERT TO GRID LOCATION IN .SUM FILE
g21:    IR% = NRow% - (Jg% - 1) * 800
        IC% = NCol% - (Ig% - 1) * 800
        IFN& = (IR% - 1) * 801! + IC%
        Get #filnumg%, IFN&, IO%
        hgt2 = IO% * 0.1
        If hgt2 < -1000 Then hgt2 = -9999
        GoTo g99
g35:    er% = Err.Number
        hgt2 = -9999 'MsgBox " ERROR IN GETZ ", vbCritical + vbOKOnly, "SkyLight"
g99:
        End If

End Sub
Public Sub InitializeDTM()
'This routine is used to initialize the reading of the Israel 25 meter
'DTM by loading up the names of the different tiles.

        CHMNEO = "XX"
        filnum% = FreeFile
        Open dtmdir & "\dtm-map.loc" For Input As #filnum%
        For i& = 1 To 3
           Line Input #filnum%, doclin$
        Next i&
        N% = 0
        For i& = 4 To 54
           Line Input #filnum%, doclin$
           If i& Mod 2 = 0 Then
              N% = N% + 1
              For j& = 1 To 14
                 CHMAP(j&, N%) = Mid$(doclin$, 6 + (j& - 1) * 5, 2)
              Next j&
              End If
        Next i&
        Close #filnum%
        JKHDTM = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SearchMaxHeights
' Author    : Dr-John-K-Hall
' Date      : 8/13/2015
' Purpose   : Search area on map for the maximum height
'---------------------------------------------------------------------------------------
'
Public Function SearchMaxHeights(Pic As PictureBox, RectCoord() As POINTAPI) As Integer

   Dim SearchGeoCoord(3) As POINTGEO
   Dim i As Integer, lat As Double, lon As Double, hgt2 As Integer
   Dim pixStepX As Double, pixStepY As Double '<<<<< added these two variables
   Dim lg2 As Double, lt2 As Double
   Dim XGeo As Double, GeoX As Double '<<<<< added GeoX
   Dim YGeo As Double, GeoY As Double
   Dim XStep As Double, YStep As Double '<<<next lines
   Dim MaxHeight As Double, MaxXGeo As Double, MaxYGeo As Double
   Dim iprogress&, newprogress&
   Dim numXsteps&, numYsteps&
   Dim num_x&, num_y&
   Dim lg1 As Integer, lt1 As Integer
   Dim RoundOffX As Double, RoundOffY As Double '<<<<<added both
   
   RoundOffX = 0.00001 '<<<<<
   RoundOffY = 0.00001 '<<<<<
   
   Dim ier As Integer
   
   ier = 0

   On Error GoTo SearchMaxHeights_Error
   
   Dim xmin As Double, xmax As Double '<<<next three lines
   Dim ymin As Double, ymax As Double
   Dim zmin As Double, zmax As Double
   
   Dim X1, Y1, X2, Y2  '<<<<<added 4 lines
   Dim XPixStep As Double, YPixStep As Double
   Dim VarD As Double
   Dim BytePosit As Long
   
   Dim ConstantXGeo As Boolean '<<<<added all lines up to MaxHeight =
   Dim XGeoConstant As Double
   Dim YGeoConstant As Double
   
   Dim CalculateRotatedGrid As Boolean
   
   Dim CurrentX As Double, CurrentY As Double
   Dim ShiftX As Double, ShiftY As Double
   
   Dim Tolerance As Double, ANG1 As Double, ANG2 As Double
   Tolerance = 0.00001
   
   ConstantXGeo = True
   
   RotateddGrid = False
   
   MaxHeight = -INIT_VALUE
   MaxXGeo = 0
   MaxYGeo = 0
   
   X1 = RectCoord(0).X '<<<<next 8 lines
   Y1 = RectCoord(0).Y
   X2 = RectCoord(1).X
   Y2 = RectCoord(1).Y
   
   xmax = X2
   xmin = X1
   ymax = Y2
   ymin = Y1
   
   If Not (heights Or (UseNewDTM% And BasisDTMheights)) And Not (RSMethod0 Or RSMethod1 Or RSMethod2) Then '<<<<
       
      ier = -1
      SearchMaxHeights = ier
      Exit Function
      
      End If
      
  Screen.MousePointer = vbHourglass '<<<<
  
  'determine DTM spacing in pixels
   
   If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
      XStep = 25
      YStep = 25
   ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
      XStep = 30
      YStep = 30
   ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
      XStep = 8.33333333333333E-04 / 3#
      YStep = 8.33333333333333E-04 / 3#
   Else
      XStep = XStepDTM
      YStep = YStepDTM
      End If

   'search the selected region for the highest point
   
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
     
'       iprogress& = 0
   
     'convert the pixel search coordinates to geo coordinates
  For i = 0 To 1 '--2

    If RSMethod1 Then '--3
       ier = RS_pixel_to_coord2(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
    ElseIf RSMethod2 Then
       ier = RS_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
    ElseIf RSMethod0 Then
       ier = Simple_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
       End If '--3
           
    'now determine new boundaries so that boundaries are on data point
    'determine lat,lon of first element of this tile
       
    If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then '--3
        
        If DTMtype = 2 Then '--4
            lg1 = Int(SearchGeoCoord(i).XGeo)
            If lg1 < 0 And lg1 > lg2 Then lg1 = lg1 - 1
            lt1 = Int(SearchGeoCoord(i).YGeo) 'SRTM tiles are named by SW corner
            If lt1 < 0 And lt1 > lt Then lt1 = lt1 - 1
            
            lt1 = lt1 + 1 'the first record in SRTM tiles in the NW corner
            SearchGeoCoord(i).XGeo = Int((SearchGeoCoord(i).XGeo - lg1) / XStep) * XStep + lg1
            SearchGeoCoord(i).YGeo = lt1 - Int((lt1 - SearchGeoCoord(i).YGeo) / YStep) * YStep
            
        ElseIf DTMtype = 1 Then 'ASTER '--4
            lg1 = Int(SearchGeoCoord(i).XGeo)
            If lg1 < 0 And lg1 > lg2 Then lg1 = lg1 - 1
            lt1 = Int(SearchGeoCoord(i).YGeo) 'ASTER tiles are named by SW corner
            If lt1 < 0 And lt1 > lt Then lt1 = lt1 - 1
            
            'first data point is in SW corner
            SearchGeoCoord(i).XGeo = Int((SearchGeoCoord(i).XGeo - lg1) / XStep) * XStep + lg1
            SearchGeoCoord(i).YGeo = Int((SearchGeoCoord(i).YGeo - lt1) / YStep) * YStep + lt1
            
            End If '--4
        
     ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then '--3
        
        If JKHDTM Then '--4
           '25 meter spacing
           SearchGeoCoord(i).XGeo = Int(SearchGeoCoord(i).XGeo / 25) * 25
           SearchGeoCoord(i).YGeo = Int(SearchGeoCoord(i).YGeo / 25) * 25
        Else
           'approximate 30m spacing for latitudes of Eretz Yisroel
           SearchGeoCoord(i).XGeo = Int(SearchGeoCoord(i).XGeo / 30) * 30
           SearchGeoCoord(i).YGeo = Int(SearchGeoCoord(i).YGeo / 30) * 30
           End If '--4
           
     Else '--3
     
        SearchGeoCoord(i).XGeo = Int(SearchGeoCoord(i).XGeo / XStep) * XStep
        SearchGeoCoord(i).YGeo = Int(SearchGeoCoord(i).YGeo / YStep) * YStep
            
        End If '--3
           
  Next i '--2
     
  Dim XX As Long, YY As Long
 
  If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
       'now calculate the other corners to see how square the grid is
        'convert the pixel search coordinates to geo coordinates
       For i = 2 To 3

        If i = 2 Then 'SW corner
           XX = RectCoord(0).X
           YY = RectCoord(1).Y
        ElseIf i = 3 Then 'NE corner
           XX = RectCoord(1).X
           YY = RectCoord(0).Y
           End If
       
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod0 Then
           ier = Simple_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
           End If

        'now determine new boundaries so that boundaries are on data point
        'determine lat,lon of first element of this tile

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

     Next i
 
     If SearchGeoCoord(0).XGeo <> SearchGeoCoord(2).XGeo Or SearchGeoCoord(1).XGeo <> SearchGeoCoord(3).XGeo Or _
        SearchGeoCoord(0).YGeo <> SearchGeoCoord(3).YGeo Or SearchGeoCoord(1).YGeo <> SearchGeoCoord(2).YGeo Then
    
         RotatedGrid = True
    
         Select Case MsgBox("The geographic grid is rotated with respect to the pixel grid." _
                            & vbCrLf & "" _
                            & vbCrLf & "Using the pixel grid to search may skip points" _
                            & vbCrLf & "while searching for the highest elevation." _
                            & vbCrLf & "" _
                            & vbCrLf & "It is recommmended to use the rotated geographic grid instead" _
                            & vbCrLf & vbCrLf & "Proceed?" _
                            , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "Maximum Elevation Search")
         
             Case vbYes
                 CalculateRotatedGrid = True
                 RotatedGrid = False
             Case vbNo, vbCancel
                 CalculateRotatedGrid = False
         
         End Select
         End If
       
  ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
      CalculateRotatedGrid = False
  Else
    
    For i = 2 To 3
    
        If i = 2 Then 'SW corner
           XX = RectCoord(0).X
           YY = RectCoord(1).Y
        ElseIf i = 3 Then 'NE corner
           XX = RectCoord(1).X
           YY = RectCoord(0).Y
           End If
          
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod0 Then
           ier = Simple_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
           End If
    
        'now determine new boundaries so that boundaries are on data point
        'determine lat,lon of first element of this tile
        SearchGeoCoord(i).XGeo = Int((SearchGeoCoord(i).XGeo) / XStep) * XStep
        SearchGeoCoord(i).YGeo = Int((SearchGeoCoord(i).YGeo) / YStep) * YStep
    
    Next i
 
    If SearchGeoCoord(0).XGeo <> SearchGeoCoord(2).XGeo Or SearchGeoCoord(1).XGeo <> SearchGeoCoord(3).XGeo Or _
       SearchGeoCoord(0).YGeo <> SearchGeoCoord(3).YGeo Or SearchGeoCoord(1).YGeo <> SearchGeoCoord(2).YGeo Then
           
           RotatedGrid = True
           
           Select Case MsgBox("The geographic grid is rotated with respect to the pixel grid." _
                              & vbCrLf & "" _
                              & vbCrLf & "Using the pixel grid to search may skip points" _
                              & vbCrLf & "when searching for the highest elevation." _
                              & vbCrLf & "" _
                              & vbCrLf & "It is recommmended to use the rotated geographic grid instead" _
                              & vbCrLf & vbCrLf & "Proceed?" _
                              , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "Maximum Elevation Search")
           
               Case vbYes
                   CalculateRotatedGrid = True
                   RotatedGrid = False
               Case vbNo, vbCancel
                   CalculateRotatedGrid = False
           
           End Select
     Else
        CalculateRotatedGrid = False
        End If
        
     End If
    
If CalculateRotatedGrid Then '==========================================================================================================
     
    'define square grid using the geo coordinates, and determine equivalent pixel coordinates for drawing contours
    
    If LRGeoX = ULGeoX Or ULGeoY = LRGeoY Then
       'use rubber sheeting to determine them
       MsgBox "Corner grid coordinates undefined." & vbCrLf & vbCrLf & "(Hint: use options menu)", vbInformation + vbOKOnly, "DTM creation error"
       GDMDIform.Toolbar1.Buttons(45).Enabled = False
       ier = -1
       SearchMaxHeights = ier
       Exit Function
    Else
    
        GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
        GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
   
       End If
       
    'determine maximum x and y ranges
    Dim XmaxRange As Double, YmaxRange As Double, XminRange As Double, YminRange As Double
    XmaxRange = Max(CDbl(SearchGeoCoord(1).XGeo), CDbl(SearchGeoCoord(3).XGeo))
    YmaxRange = Max(CDbl(SearchGeoCoord(0).YGeo), CDbl(SearchGeoCoord(3).YGeo))
    XminRange = min(CDbl(SearchGeoCoord(0).XGeo), CDbl(SearchGeoCoord(2).XGeo))
    YminRange = min(CDbl(SearchGeoCoord(1).YGeo), CDbl(SearchGeoCoord(2).YGeo))
       
'     numXsteps& = Int((SearchGeoCoord(1).XGeo - SearchGeoCoord(0).XGeo) / XStep) + 1 'add one for roundoff
'     numYsteps& = Int((SearchGeoCoord(0).YGeo - SearchGeoCoord(1).YGeo) / YStep) + 1
     numXsteps& = Int((XmaxRange - XminRange) / XStep) + 1 '(X2 - X1 + 1)
     numYsteps& = Int((YmaxRange - YminRange) / YStep) + 1 '(Y2 - Y1 + 1)
     
     GDMDIform.StatusBar1.Panels(1).Text = "Searching for highest point, please wait..."
     
     num_x& = 0

     For i = 1 To numXsteps& 'loop from west from east
     
         GeoX = XminRange + XStep * (i - 1)
     
         For j = 1 To numYsteps& 'loop from south to north
         
            GeoY = YminRange + YStep * (j - 1)
            
            'determine estimate of the corresponding Pixel coordinates
         
            GoSub GeotoCoord
            
            'now determine the heights at these coordinates
            XGeo = GeoX
            YGeo = GeoY
            
            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                 'convert from ITM to WGS84
                 kmx = XGeo
                 kmy = YGeo
                 Call ics2wgs84(CLng(kmy), CLng(kmx), lt2, lg2)
            Else
                 lg2 = XGeo
                 lt2 = YGeo
                 End If
                 
            If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
            
                If BasisDTMheights And UseNewDTM% Then
                   'use background dtm as height reference
                   kmx = XGeo
                   kmy = YGeo
                   
                   If XGeo >= xLL And YGeo >= yLL Then
                        BytePosit = 101 + 8 * Nint((XGeo - xLL) / XStepLL) + 8 * nColLL * Nint((YGeo - yLL) / YStepLL)
                        If BytePosit < 0 Then
                           VarD = 0
                        Else
                           Get #basedtm%, BytePosit, VarD
                           End If
                        
                        If VarD = blank_value Then
                           VarD = -9999
                        ElseIf VarD < -100000 Or VarD > 100000 Then
                           VarD = -9999 'flag unreadible height
                           End If
                        
                        hgt2 = VarD / (DigiConvertToMeters * MapUnits)
                   Else
                       hgt2 = -9999
                       End If
                   
                Else 'use stored dtm's
                 
                    If DTMtype = 1 Then
                       'use ASTER
                       Call ASTERheight(lg2, lt2, hgt2)
                    ElseIf DTMtype = 2 Then
                       'use JKH's DTM if ITM coordinates, else use NED, SRTM
                       If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                          Call DTMheight2(lg2, lt2, hgt2)
                       Else
                          Call worldheights(lg2, lt2, hgt2)
                          End If
                       End If
                       
                    End If
                    
                End If
               
            If hgt2 > MaxHeight Then
               MaxHeight = hgt2
               MaxXGeo = lg2
               MaxYGeo = lt2
               End If
               
         Next j
         
         num_x& = num_x& + 1
         
         DoEvents
         
         Call UpdateStatus(GDMDIform, 1, 100 * num_x& / numXsteps&)
         
     Next i
   
     Screen.MousePointer = vbDefault
     
     GDMDIform.StatusBar1.Panels(1).Text = sEmpty
     GDMDIform.StatusBar1.Panels(2).Text = sEmpty
     Call UpdateStatus(GDMDIform, 1, 0)
     GDMDIform.picProgBar.Visible = False
       
ElseIf Not CalculateRotatedGrid Then '=============================================================================================
 
     'grid is square
     numXsteps& = Int((SearchGeoCoord(1).XGeo - SearchGeoCoord(0).XGeo) / XStep) + 1  '(X2 - X1 + 1)
     numYsteps& = Int((SearchGeoCoord(0).YGeo - SearchGeoCoord(1).YGeo) / YStep) + 1  '(Y2 - Y1 + 1)
     
     XPixStep = (X2 - X1 + 1) / numXsteps& '1
     YPixStep = (Y2 - Y1 + 1) / numYsteps& '1
     
     GDMDIform.StatusBar1.Panels(1).Text = "Searching for the highest point, please wait..."
     
     num_x& = 0
     
     For pixStepX = X1 To X2 + RoundOffX Step XPixStep
         
         num_y& = 0
     
         For pixStepY = Y2 To Y1 - RoundOffY Step -YPixStep
            
            'convert pixel coordinate to geographic cooridinate
          
            If RSMethod1 Then
               ier = RS_pixel_to_coord2(pixStepX, pixStepY, XGeo, YGeo)
            ElseIf RSMethod2 Then
               ier = RS_pixel_to_coord(pixStepX, pixStepY, XGeo, YGeo)
            ElseIf RSMethod0 Then
               ier = Simple_pixel_to_coord(pixStepX, pixStepY, XGeo, YGeo)
               End If
     
            If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" And ConstantXGeo And RSMethod2 Then
                'rotate plot to rectangular grid based on the upper-left corner value of the geographic coordinates
                If ConstantXGeo And RSMethod2 And (pixStepX = CDbl(X1) And pixStepY = CDbl(Y2)) Then
                   XGeoConstant = XGeo
                   YGeoConstant = YGeo
                ElseIf ConstantXGeo And RSMethod2 And Not (pixStepX = CDbl(X1) And pixStepY = CDbl(Y2)) Then
                   'ignore rubber sheeting coordinates, rather use square grid.
                   XGeo = XGeoConstant + ((pixStepX - CDbl(X1)) / XPixStep) * XStep
                   YGeo = YGeoConstant + ((CDbl(Y2) - pixStepY) / YPixStep) * YStep
                   End If
                End If
                
            If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
            
                If BasisDTMheights And UseNewDTM% Then
                   'use background dtm as height reference
                   kmx = XGeo
                   kmy = YGeo
                    
                   If XGeo >= xLL And YGeo >= yLL Then
                        BytePosit = 101 + 8 * Nint((XGeo - xLL) / XStepLL) + 8 * nColLL * Nint((YGeo - yLL) / YStepLL)
                        If BytePosit < 0 Then
                           VarD = 0
                        Else
                           Get #basedtm%, BytePosit, VarD
                           End If
                        
                        If VarD = blank_value Then
                           VarD = -9999
                        ElseIf VarD < -100000 Or VarD > 100000 Then
                           VarD = -9999 'flag unreadible height
                           End If
                        
                        hgt2 = VarD / (DigiConvertToMeters * MapUnits)
                    Else
                        hgt2 = -9999
                        End If
                   
                Else 'use stored dtm's
                
                    lg2 = XGeo
                    lt2 = YGeo
                 
                    If DTMtype = 1 Then
                       'use ASTER
                       Call ASTERheight(lg2, lt2, hgt2)
                    ElseIf DTMtype = 2 Then
                       'use JKH's DTM if ITM coordinates, else use NED, SRTM
                       If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                          Call DTMheight2(lg2, lt2, hgt2)
                       Else
                          Call worldheights(lg2, lt2, hgt2)
                          End If
                       End If
                       
                    End If
                    
                End If
                 
            If num_y& > numYsteps& - 1 Then Exit For
            
            If hgt2 > MaxHeight Then
               MaxHeight = hgt2
               MaxXGeo = lg2
               MaxYGeo = lt2
               End If
               
            num_y& = num_y& + 1
               
         Next pixStepY
         
         If num_x& > numXsteps& - 1 Then Exit For
         
         num_x& = num_x& + 1
         
         DoEvents
         
         Call UpdateStatus(GDMDIform, 1, 100 * num_x& / numXsteps&)
         
     Next pixStepX
   
     Screen.MousePointer = vbDefault
     
     GDMDIform.StatusBar1.Panels(1).Text = sEmpty
     GDMDIform.StatusBar1.Panels(2).Text = sEmpty
     Call UpdateStatus(GDMDIform, 1, 0)
     GDMDIform.picProgBar.Visible = False
     
     End If
    
'     numsteps& = 0
'
'     For XGeo = SearchGeoCoord(0).XGeo To SearchGeoCoord(1).XGeo + XStep Step XStep
'
'         numsteps& = numsteps& + 1
'
'         For YGeo = SearchGeoCoord(1).YGeo To SearchGeoCoord(0).YGeo + YStep Step YStep
'
'            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
'                 'convert from ITM to WGS84
'                 kmx = XGeo
'                 kmy = YGeo
'                 Call ics2wgs84(CLng(kmy), CLng(kmx), lt2, lg2)
'            Else
'                 lg2 = XGeo
'                 lt2 = YGeo
'                 End If
'
'              If DTMtype = 1 Then
'                 'use ASTER
'                 Call ASTERheight(lg2, lt2, hgt2)
'              ElseIf DTMtype = 2 Then
'                 'use JKH's DTM if ITM coordinates, else use NED, SRTM
'                 If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
'                    kmx = lg2
'                    kmy = lt2
'                    Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt2)
'                 Else
'                    Call worldheights(lg2, lt2, hgt2)
'                    End If
'                 End If
'
'              If hgt2 > MaxHeight Then
'                 MaxHeight = hgt2
'                 MaxXGeo = lg2
'                 MaxYGeo = lt2
'                 End If
'
'         Next YGeo
'
'        newprogress& = CInt(100 * numsteps& / numXsteps&)
'        If iprogress& <> newprogress& Then
'           iprogress& = newprogress&
'           Call UpdateStatus(GDMDIform, 1, iprogress&)
'           End If
'
'     Next XGeo
'
'  Else
'
'     ier = -1
'     SearchMaxHeights = ier
'     Exit Function
'
'     End If
'
'   Screen.MousePointer = vbDefault
'   GDMDIform.StatusBar1.Panels(1).Text = sEmpty
'   GDMDIform.StatusBar1.Panels(2).Text = sEmpty
''
'   Call UpdateStatus(GDMDIform, 1, 0)
'   GDMDIform.picProgBar.Visible = False
     
   GDMDIform.Text5.Text = Format(str$(MaxXGeo), "#######.####0")
   GDMDIform.Text6.Text = Format(str$(MaxYGeo), "#######.####0")
   
   DoEvents 'pauss to process the message queue
   
   GDMDIform.Text5.Refresh
   GDMDIform.Text6.Refresh
     
   Call gotocoord
    
   SearchMaxHeights = ier

   On Error GoTo 0
   Exit Function
   
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

         End If
              
Return
   

SearchMaxHeights_Error:

   Select Case Err.Number
      Case 52
         'problem with base dtm's file number
         'close it and reopen
         ier = OpenCloseBaseDTM(0)
         Resume
      Case 63
         'bad record number caused by being off the map sheet
         'return the blank height value
         'exit sub with error
     End Select

    Screen.MousePointer = vbDefault
    GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SearchMaxHeights of Module modConvertHall"
    SearchMaxHeights = -1
    
End Function
Public Function DistTrav(lat_0 As Double, lon_0 As Double, lat_1 As Double, lon_1 As Double, mode%) As Double
   
   'calclates angular distance (central angle) on spherical earth from (lat0,lon1) to (lat1,lon1)
   'to convert to kilometers multiply by Rearthkm
   'to convert to meters, multiply by Reathkm * 1000.0
   
   'mode% = 0 'passed coordinates are in radians
   'mode% = 1 'passed coordinaes are in degrees
   'mode% = 2 'passed coordinates are in radians, use Vincenty formula source: http://en.wikipedia.org/wiki/Great-circle_distance
   'mode% = 3 'passed coordinates are in degrees, use Vincenty formula
   
   Dim AA As Double, BB As Double, cc As Double
   Dim lat0 As Double, lat1 As Double, lon0 As Double, lon1 As Double
   Dim DistTrav2 As Double, DifDist As Double, ccc
   
   On Error GoTo DistTrav_Error

   If mode% = 0 Or mode% = 2 Then
      lat0 = lat_0
      lat1 = lat_1
      lon0 = lon_0
      lon1 = lon_1
   Else
      lat0 = lat_0 * cd
      lat1 = lat_1 * cd
      lon0 = lon_0 * cd
      lon1 = lon_1 * cd
      End If

   If mode% <= 1 Then
      AA = Sin((lat0 - lat1) * 0.5)
      BB = Sin((lon0 - lon1) * 0.5)
      DistTrav = 2# * DASIN(Sqr(AA * AA + Cos(lat0) * Cos(lat1) * BB * BB)) 'central angle in radians
      
   ElseIf mode% > 1 Then 'use Vincenty formula (more accurate for small distances and for 32 bit calculations)
      AA = Cos(lat1) * Sin(lon1 - lon0)
      BB = Cos(lat0) * Sin(lat1) - Sin(lat0) * Cos(lat1) * Cos(lon1 - lon0)
      cc = Sin(lat0) * Sin(lat1) + Cos(lat0) * Cos(lat1) * Cos(lon1 - lon0)
      DistTrav = Atan2(cc, Sqr(AA * AA + BB * BB)) 'central angle in radians
      
      End If

   On Error GoTo 0
   Exit Function

DistTrav_Error:

    Exit Function
   
End Function

'---------------------------------------------------------------------------------------
' Procedure : Contours
' Author    : Dr-John-K-Hall
' Date      : 8/16/2015
' Purpose   : Generate contours for selected region on map
'---------------------------------------------------------------------------------------
'
Public Function Contours(Pic As PictureBox, RectCoord() As POINTAPI) As Integer

   Dim SearchGeoCoord(3) As POINTGEO
   Dim i As Long, j As Long, lat As Double, lon As Double, hgt2 As Integer
   Dim pixStepX As Double, pixStepY As Double
   Dim lg2 As Double, lt2 As Double
   Dim XGeo As Double, GeoX As Double
   Dim YGeo As Double, GeoY As Double
   Dim XStep As Double, YStep As Double
   Dim MaxHeight As Double, MaxXGeo As Double, MaxYGeo As Double
   Dim iprogress&, newprogress&
   Dim numXsteps&, numYsteps&, numContourPnts&
   Dim num_x&, num_y&
   Dim lg1 As Integer, lt1 As Integer
   Dim RoundOffX As Double, RoundOffY As Double
   
   RoundOffX = 0.00001
   RoundOffY = 0.00001
   
   Dim Xcoord() As Double, Ycoord() As Double
   Dim ht() As Double, htf() As Single, hts() As Integer
   Dim xmin As Double, xmax As Double
   Dim ymin As Double, ymax As Double
   Dim zmin As Double, zmax As Double
   
   Dim nc As Integer
   Dim contour() As Double
   Dim X1, Y1, X2, Y2
   Dim XPixStep As Double, YPixStep As Double
   Dim VarD As Double
   Dim BytePosit As Long
   
   Dim ConstantXGeo As Boolean
   Dim XGeoConstant As Double
   Dim YGeoConstant As Double
   
   Dim CalculateRotatedGrid As Boolean
   
   Dim CurrentX As Double, CurrentY As Double
   Dim ShiftX As Double, ShiftY As Double
   
   Dim ier As Integer
   Dim ContourInterval As Integer
   
   Dim Tolerance As Double, ANG1 As Double, ANG2 As Double
   Tolerance = 0.00001
   
   ConstantXGeo = True
   
   RotateddGrid = False
   
   On Error GoTo Contours_Error
   
   ContourInterval = val(GDMDIform.combContour.Text) '10 '5 '10 '100 'contour intervals in height units
    
   zmin = INIT_VALUE
   zmax = -INIT_VALUE
   
   X1 = RectCoord(0).X
   Y1 = RectCoord(0).Y
   X2 = RectCoord(1).X
   Y2 = RectCoord(1).Y
   
   xmax = X2
   xmin = X1
   ymax = Y2
   ymin = Y1
   
   ier = 0

   'record elevations in the selected region
   
   If Not (heights Or (UseNewDTM% And BasisDTMheights)) And Not (RSMethod0 Or RSMethod1 Or RSMethod2) Then
       
      ier = -1
      Contours = ier
      Exit Function
      
      End If
   
  Screen.MousePointer = vbHourglass
   
  'determine DTM spacing in pixels

  If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
     XStep = 25
     YStep = 25
  ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
     XStep = 30
     YStep = 30
  ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" And (DTMtype = 1 Or DTMtype = 2) Then
     XStep = 8.33333333333333E-04 / 3#
     YStep = 8.33333333333333E-04 / 3#
  Else
     XStep = XStepDTM
     YStep = YStepDTM
     End If
     
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

'      iprogress& = 0
'      find coordinates of the coordinates to see if the grid is square
'      if it is not square, then have to use iteration to convert the geo coordinates back to aproximate pixels,
'      and then use the rubbersheeting to find the exact geo coordinates, and then only give spacing to the accuracy of the dtm grid
'
  'convert the pixel search coordinates to geo coordinates
  For i = 0 To 1

     If RSMethod1 Then
        ier = RS_pixel_to_coord2(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
     ElseIf RSMethod2 Then
        ier = RS_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
     ElseIf RSMethod0 Then
        ier = Simple_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
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
     
 Dim XX As Long, YY As Long
 
 If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
    'now calculate the other corners to see how square the grid is
     'convert the pixel search coordinates to geo coordinates
     For i = 2 To 3
     
        If i = 2 Then 'SW corner
           XX = RectCoord(0).X
           YY = RectCoord(1).Y
        ElseIf i = 3 Then 'NE corner
           XX = RectCoord(1).X
           YY = RectCoord(0).Y
           End If
           
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod0 Then
           ier = Simple_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
           End If

        'now determine new boundaries so that boundaries are on data point
        'determine lat,lon of first element of this tile
    
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
    
     Next i
 
    If SearchGeoCoord(0).XGeo <> SearchGeoCoord(2).XGeo Or SearchGeoCoord(1).XGeo <> SearchGeoCoord(3).XGeo Or _
       SearchGeoCoord(0).YGeo <> SearchGeoCoord(3).YGeo Or SearchGeoCoord(1).YGeo <> SearchGeoCoord(2).YGeo Then
       
       RotatedGrid = True
       
       Select Case MsgBox("The geographic grid is rotated with respect to the pixel grid." _
                          & vbCrLf & "" _
                          & vbCrLf & "Using the pixel grid to produce xyz files may cause unforseen" _
                          & vbCrLf & "problems when calculating profies." _
                          & vbCrLf & "" _
                          & vbCrLf & "It is recommmended to use the rotated geographic grid instead" _
                          & vbCrLf & vbCrLf & "Proceed?" _
                          , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "Contour generation")
       
           Case vbYes
               CalculateRotatedGrid = True
               RotatedGrid = False
           Case vbNo, vbCancel
               CalculateRotatedGrid = False
       
       End Select
       End If
       
 ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
     CalculateRotatedGrid = False
 Else
    
    For i = 2 To 3
    
       If i = 2 Then 'SW corner
          XX = RectCoord(0).X
          YY = RectCoord(1).Y
       ElseIf i = 3 Then 'NE corner
          XX = RectCoord(1).X
          YY = RectCoord(0).Y
          End If
          
       If RSMethod1 Then
          ier = RS_pixel_to_coord2(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
       ElseIf RSMethod2 Then
          ier = RS_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
       ElseIf RSMethod0 Then
          ier = Simple_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
          End If
    
       'now determine new boundaries so that boundaries are on data point
       'determine lat,lon of first element of this tile
       SearchGeoCoord(i).XGeo = Int((SearchGeoCoord(i).XGeo) / XStep) * XStep
       SearchGeoCoord(i).YGeo = Int((SearchGeoCoord(i).YGeo) / YStep) * YStep
    
    Next i
 
    If SearchGeoCoord(0).XGeo <> SearchGeoCoord(2).XGeo Or SearchGeoCoord(1).XGeo <> SearchGeoCoord(3).XGeo Or _
       SearchGeoCoord(0).YGeo <> SearchGeoCoord(3).YGeo Or SearchGeoCoord(1).YGeo <> SearchGeoCoord(2).YGeo Then
           
           RotatedGrid = True
           
           Select Case MsgBox("The geographic grid is rotated with respect to the pixel grid." _
                              & vbCrLf & "" _
                              & vbCrLf & "Using the pixel grid to produce xyz files may cause unforseen" _
                              & vbCrLf & "problems when calculating profies." _
                              & vbCrLf & "" _
                              & vbCrLf & "It is recommmended to use the rotated geographic grid instead" _
                              & vbCrLf & vbCrLf & "Proceed?" _
                              , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "Contour generation")
           
               Case vbYes
                   CalculateRotatedGrid = True
                   RotatedGrid = False
               Case vbNo, vbCancel
                   CalculateRotatedGrid = False
           
           End Select
     Else
        CalculateRotatedGrid = False
        End If
        
     End If
    
If CalculateRotatedGrid Then '==========================================================================================================
    
    'define square grid using the geo coordinates, and determine equivalent pixel coordinates for drawing contours
    
    If LRGeoX = ULGeoX Or ULGeoY = LRGeoY Then
       'use rubber sheeting to determine them
       MsgBox "Corner grid coordinates undefined." & vbCrLf & vbCrLf & "(Hint: use options menu)", vbInformation + vbOKOnly, "DTM creation error"
       GDMDIform.Toolbar1.Buttons(45).Enabled = False
       ier = -1
       Contours = ier
       Exit Function
    Else
    
        GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
        GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
   
       End If
    
    'determine maximum x and y ranges
    Dim XmaxRange As Double, YmaxRange As Double, XminRange As Double, YminRange As Double
    XmaxRange = Max(CDbl(SearchGeoCoord(1).XGeo), CDbl(SearchGeoCoord(3).XGeo))
    YmaxRange = Max(CDbl(SearchGeoCoord(0).YGeo), CDbl(SearchGeoCoord(3).YGeo))
    XminRange = min(CDbl(SearchGeoCoord(0).XGeo), CDbl(SearchGeoCoord(2).XGeo))
    YminRange = min(CDbl(SearchGeoCoord(1).YGeo), CDbl(SearchGeoCoord(2).YGeo))
    
     numXsteps& = Int((XmaxRange - XminRange) / XStep) + 1 '(X2 - X1 + 1)
     numYsteps& = Int((YmaxRange - YminRange) / YStep) + 1 '(Y2 - Y1 + 1)
     
     g_nrows = numYsteps&
     g_ncols = numXsteps&
     numXYZpoints = g_nrows * g_ncols
    
     ReDim Xcoord(0 To numXsteps& - 1)
     ReDim Ycoord(0 To numYsteps& - 1)
     
     If HeightPrecision = 0 Then
         ReDim hts(0 To numXsteps& - 1, 0 To numYsteps& - 1)
         ReDim ht(0)
         ReDim htf(0)
     ElseIf HeightPrecision = 1 Then
         ReDim htf(0 To numXsteps& - 1, 0 To numYsteps& - 1)
         ReDim hts(0)
         ReDim ht(0)
     ElseIf HeightPrecision = 2 Then
         ReDim ht(0 To numXsteps& - 1, 0 To numYsteps& - 1)
         ReDim hts(0)
         ReDim htf(0)
         End If
     
     numContourPnts& = numXsteps& * numYsteps&
     
     GDMDIform.StatusBar1.Panels(1).Text = "Generating contours, please wait..."
     
     If Save_xyz% = 1 Then
         
        Dim tryingtokill As Boolean
        tryingtokill = True
         
        If Dir(App.Path & "\topo_coord.xyz") <> gsEmpty Then Kill App.Path & "\topo_coord.xyz"
        If Dir(App.Path & "\topo_pixel.xyz") <> gsEmpty Then Kill App.Path & "\topo_pixel.xyz"
        
        tryingtokill = False
        
        filnum% = FreeFile
        Open App.Path & "\topo_coord.xyz" For Output As #filnum%
        filtopo% = FreeFile
        Open App.Path & "\topo_pixel.xyz" For Output As #filtopo%
        End If
     
     num_x& = 0
     
     For i = 1 To numXsteps& 'loop from west from east
     
         GeoX = XminRange + XStep * (i - 1)
         
         num_y& = 0
     
         For j = 1 To numYsteps& 'loop from south to north
         
            GeoY = YminRange + YStep * (j - 1)
            
            'determine estimate of the corresponding Pixel coordinates
         
            GoSub GeotoCoord
            
            Xcoord(num_x&) = CurrentX
            Ycoord(num_y&) = CurrentY
            
            'now determine the heights at these coordinates
            XGeo = GeoX
            YGeo = GeoY
            
            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                 'convert from ITM to WGS84
                 kmx = XGeo
                 kmy = YGeo
                 Call ics2wgs84(CLng(kmy), CLng(kmx), lt2, lg2)
            Else
                 lg2 = XGeo
                 lt2 = YGeo
                 End If
                 
            If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
            
                If BasisDTMheights And UseNewDTM% Then
                   'use background dtm as height reference
                   kmx = XGeo
                   kmy = YGeo
                   
                   If XGeo >= xLL And YGeo >= yLL Then
                        BytePosit = 101 + 8 * Nint((XGeo - xLL) / XStepLL) + 8 * nColLL * Nint((YGeo - yLL) / YStepLL)
                        If BytePosit < 0 Then
                           VarD = 0
                        Else
                           Get #basedtm%, BytePosit, VarD
                           End If
                        
                        If VarD = blank_value Then
                           VarD = -9999
                        ElseIf VarD < -100000 Or VarD > 100000 Then
                           VarD = -9999 'flag unreadible height
                           End If
                        
                        hgt2 = VarD / (DigiConvertToMeters * MapUnits)
                   Else
                       hgt2 = -9999
                       End If
                   
                Else 'use stored dtm's
                 
                    If DTMtype = 1 Then
                       'use ASTER
                       Call ASTERheight(lg2, lt2, hgt2)
                    ElseIf DTMtype = 2 Then
                       'use JKH's DTM if ITM coordinates, else use NED, SRTM
                       If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                          Call DTMheight2(lg2, lt2, hgt2)
                       Else
                          Call worldheights(lg2, lt2, hgt2)
                          End If
                       End If
                       
                    End If
                    
                End If
               
            If hgt2 > zmax Then zmax = hgt2
            If hgt2 < zmin Then zmin = hgt2
                 
            If HeightPrecision = 0 Then
                hts(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 1 Then
                htf(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 2 Then
                ht(num_x&, num_y&) = hgt2
                End If
            
            num_y& = num_y& + 1
            
'            numsteps& = numsteps& + 1
            
            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                XGeo = val(Format(str$(XGeo), "#####0.0##")) 'CLng(XGeo)
                YGeo = val(Format(str$(YGeo), "######0.0##")) 'CLng(YGeo)
            ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
               XGeo = val(Format(str$(XGeo), "####0.0######"))
               YGeo = val(Format(str$(YGeo), "####0.0######"))
               End If
         
            If Save_xyz% = 1 Then
                Write #filtopo%, CurrentX, CurrentY, hgt2
                Write #filnum%, XGeo, YGeo, hgt2
                End If
         
         Next j
         
         num_x& = num_x& + 1
         
         DoEvents
         
         Call UpdateStatus(GDMDIform, 1, 100 * num_x& / numXsteps&)
         
     Next i
   
     Screen.MousePointer = vbDefault
     
     If Save_xyz% = 1 Then
        Close #filnum%
        Close #filtopo%
        End If
     
     filnum% = 0
     filtopo% = 0
     
     GDMDIform.StatusBar1.Panels(1).Text = sEmpty
     GDMDIform.StatusBar1.Panels(2).Text = sEmpty
     Call UpdateStatus(GDMDIform, 1, 0)
     GDMDIform.picProgBar.Visible = False
       
ElseIf Not CalculateRotatedGrid Then '=============================================================================================
 
     'grid is square
     numXsteps& = Int((SearchGeoCoord(1).XGeo - SearchGeoCoord(0).XGeo) / XStep) + 1  '(X2 - X1 + 1)
     numYsteps& = Int((SearchGeoCoord(0).YGeo - SearchGeoCoord(1).YGeo) / YStep) + 1  '(Y2 - Y1 + 1)
     
     XPixStep = (X2 - X1 + 1) / numXsteps& '1
     YPixStep = (Y2 - Y1 + 1) / numYsteps& '1
     
'     numXsteps& = CInt((X2 - X1 + 1) / XPixStep)
'     numYsteps& = CInt((Y2 - Y1 + 1) / YPixStep)
     
     g_nrows = numYsteps&
     g_ncols = numXsteps&
     
     ReDim Xcoord(0 To numXsteps& - 1)
     ReDim Ycoord(0 To numYsteps& - 1)
     
     If HeightPrecision = 0 Then
        ReDim hts(0 To numXsteps& - 1, 0 To numYsteps& - 1)
        ReDim ht(0)
        ReDim htf(0)
     ElseIf HeightPrecision = 1 Then
        ReDim htf(0 To numXsteps& - 1, 0 To numYsteps& - 1)
        ReDim ht(0)
        ReDim hts(0)
     ElseIf HeightPrecision = 2 Then
        ReDim ht(0 To numXsteps& - 1, 0 To numYsteps& - 1)
        ReDim hts(0)
        ReDim htf(0)
        End If
     
     numContourPnts& = numXsteps& * numYsteps&
     
     GDMDIform.StatusBar1.Panels(1).Text = "Generating contours, please wait..."
     
     If Save_xyz% = 1 Then
         
        tryingtokill = True
         
        If Dir(App.Path & "\topo_coord.xyz") <> gsEmpty Then Kill App.Path & "\topo_coord.xyz"
        If Dir(App.Path & "\topo_pixel.xyz") <> gsEmpty Then Kill App.Path & "\topo_pixel.xyz"
        
        tryingtokill = False
        
        filnum% = FreeFile
        Open App.Path & "\topo_coord.xyz" For Output As #filnum%
        filtopo% = FreeFile
        Open App.Path & "\topo_pixel.xyz" For Output As #filtopo%
        End If
     
     num_x& = 0
     
     For pixStepX = X1 To X2 + RoundOffX Step XPixStep
            
         Xcoord(num_x&) = pixStepX
         
         num_y& = 0
     
         For pixStepY = Y2 To Y1 - RoundOffY Step -YPixStep
         
            Ycoord(num_y&) = pixStepY
            
            'convert pixel coordinate to geographic cooridinate
          
            If RSMethod1 Then
               ier = RS_pixel_to_coord2(pixStepX, pixStepY, XGeo, YGeo)
            ElseIf RSMethod2 Then
               ier = RS_pixel_to_coord(pixStepX, pixStepY, XGeo, YGeo)
            ElseIf RSMethod0 Then
               ier = Simple_pixel_to_coord(pixStepX, pixStepY, XGeo, YGeo)
               End If
               
            If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" And ConstantXGeo And RSMethod2 Then
                'rotate plot to rectangular grid based on the upper-left corner value of the geographic coordinates
                If ConstantXGeo And RSMethod2 And (pixStepX = CDbl(X1) And pixStepY = CDbl(Y2)) Then
                   XGeoConstant = XGeo
                   YGeoConstant = YGeo
                ElseIf ConstantXGeo And RSMethod2 And Not (pixStepX = CDbl(X1) And pixStepY = CDbl(Y2)) Then
                   'ignore rubber sheeting coordinates, rather use square grid.
                   XGeo = XGeoConstant + ((pixStepX - CDbl(X1)) / XPixStep) * XStep
                   YGeo = YGeoConstant + ((CDbl(Y2) - pixStepY) / YPixStep) * YStep
                   End If
                End If
    
            If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
            
                If BasisDTMheights And UseNewDTM% Then
                   'use background dtm as height reference
                   kmx = XGeo
                   kmy = YGeo
                   
                   If XGeo >= xLL And YGeo >= yLL Then
                        BytePosit = 101 + 8 * Nint((XGeo - xLL) / XStepLL) + 8 * nColLL * Nint((YGeo - yLL) / YStepLL)
                        If BytePosit < 0 Then
                           VarD = 0
                        Else
                           Get #basedtm%, BytePosit, VarD
                           End If
                        
                        If VarD = blank_value Then
                           VarD = -9999
                        ElseIf VarD < -100000 Or VarD > 100000 Then
                           VarD = -9999 'flag unreadible height
                           End If
                        
                        hgt2 = VarD / (DigiConvertToMeters * MapUnits)
                    Else
                        hgt2 = -9999
                        End If
                   
                Else 'use stored dtm's
                
                    lg2 = XGeo
                    lt2 = YGeo
                 
                    If DTMtype = 1 Then
                       'use ASTER
                       Call ASTERheight(lg2, lt2, hgt2)
                    ElseIf DTMtype = 2 Then
                       'use JKH's DTM if ITM coordinates, else use NED, SRTM
                       If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                          Call DTMheight2(lg2, lt2, hgt2)
                       Else
                          Call worldheights(lg2, lt2, hgt2)
                          End If
                       End If
                       
                    End If
                    
                End If
                 
            If num_y& > numYsteps& - 1 Then Exit For
            
            If hgt2 > zmax Then zmax = hgt2
            If hgt2 < zmin Then zmin = hgt2
                 
            If HeightPrecision = 0 Then
                hts(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 1 Then
                htf(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 2 Then
                ht(num_x&, num_y&) = hgt2
                End If
            
            num_y& = num_y& + 1
            
'            numsteps& = numsteps& + 1
            
            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                XGeo = val(Format(str$(XGeo), "#####0.0##")) 'CLng(XGeo)
                YGeo = val(Format(str$(YGeo), "######0.0##")) 'CLng(YGeo)
            ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
               XGeo = val(Format(str$(XGeo), "####0.0######"))
               YGeo = val(Format(str$(YGeo), "####0.0######"))
'               XGeo = val(Format(str$(XGeo), "####0.0###"))
'               YGeo = val(Format(str$(YGeo), "####0.0###"))
               End If
         
            If Save_xyz% = 1 Then
                Write #filtopo%, i, Y1 + Int((Y2 - j) / YPixStep) * YPixStep, hgt2
                Write #filnum%, XGeo, YGeo, hgt2
                End If
         
            'Call UpdateStatus(GDMDIform, 1, 100 * CInt(numsteps& / numContourPnts&))
         
         Next pixStepY
         
         If num_x& > numXsteps& - 1 Then Exit For
         
         num_x& = num_x& + 1
         
         DoEvents
         
         Call UpdateStatus(GDMDIform, 1, 100 * num_x& / numXsteps&)
         
     Next pixStepX
   
     Screen.MousePointer = vbDefault
     
     If Save_xyz% = 1 Then
        Close #filnum%
        Close #filtopo%
        End If
     
     filnum% = 0
     filtopo% = 0
     
     GDMDIform.StatusBar1.Panels(1).Text = sEmpty
     GDMDIform.StatusBar1.Panels(2).Text = sEmpty
     Call UpdateStatus(GDMDIform, 1, 0)
     GDMDIform.picProgBar.Visible = False
     
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
    
    'clear last canvas before drawing new one
    ier = ReDrawMap(0)
    
   'also clear old clutter
        
    If DigitizeOn Then
       If Not InitDigiGraph Then
          InputDigiLogFile 'load up saved digitizing data for the current map sheet
       Else
          ier = RedrawDigiLog
          End If
       End If
    
    
'    ier = conrec(Pic, ht, Xcoord, Ycoord, nc, contour, 0, numXsteps& - 1, 0, numYsteps& - 1, X1, Y1, X2, Y2)
    ier = conrec(Pic, ht, htf, hts, Xcoord, Ycoord, nc, contour, 0, num_x& - 1, 0, num_y& - 1, xmin, ymin, xmax, ymax, 0)

    If ier = -1 Then
       Call MsgBox("Palette file: rainbow.cpt is missing in the program directory." _
                  & vbCrLf & "" _
                  & vbCrLf & "Contours won't be drawn" _
                  , vbExclamation, "Contours")
       End If
    
    '-------------------------finished----------------------------------
   Contours = ier
   

   On Error GoTo 0
   Exit Function
   
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
'
'            Call MsgBox("Inverse coordinate transformation unsuccessful" _
'                        & vbCrLf & "Coordinate grid rotation too large for first approx." _
'                        & vbCrLf & vbCrLf & "(Redo using a less-rotated grid as reference...)" _
'                        , vbInformation, "Contours Error")
'            ier = -1
'            Contours = ier
'            Screen.MousePointer = vbDefault
'            GDMDIform.picProgBar.Visible = False
'            GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
'            GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
'            Exit Function
'            End If
         End If
              
Return

Contours_Error:

   Select Case Err.Number
      Case 52
         'problem with base dtm's file number
         'close it and reopen
         ier = OpenCloseBaseDTM(0)
         Resume
      Case 63
         'bad record number caused by being off the map sheet
         'return the blank height value
         'exit sub with error
     End Select

   If tryingtokill Then Resume Next

   Screen.MousePointer = vbDefault
   GDMDIform.StatusBar1.Panels(1).Text = sEmpty
   GDMDIform.StatusBar1.Panels(2).Text = sEmpty
    
   GDMDIform.picProgBar.Visible = False
   
   If Save_xyz% = 1 Then
      If filnum% > 0 Then Close #filnum%
      If filtopo% > 0 Then Close #filtopo%
      End If
   
   GDMDIform.StatusBar1.Panels(1).Text = "Error Number: " & Err.Number & ", " & Err.Description
   
   ier = -1
   Contours = ier

End Function
'---------------------------------------------------------------------------------------
' Procedure : CreateDTMBackground
' Author    : Dr-John-K-Hall
' Date      : 12/20/2015
' Purpose   : Creates background DTM used for merging onto
'---------------------------------------------------------------------------------------
'
Public Function CreateDTMBackground() As Integer

   Dim SearchGeoCoord(3) As POINTGEO
   Dim RectCoord(1) As POINTAPI
   Dim i As Long, j As Long, lat As Double, lon As Double, hgt2 As Integer
   Dim pixStepX As Double, pixStepY As Double
   Dim lg2 As Double, lt2 As Double
   Dim XGeo As Double, GeoX As Double
   Dim YGeo As Double, GeoY As Double
   Dim XStep As Double, YStep As Double
   Dim MaxHeight As Double, MaxXGeo As Double, MaxYGeo As Double
   Dim iprogress&, newprogress&
   Dim numXsteps&, numYsteps&, numContourPnts&
   Dim num_x&, num_y&
   Dim lg1 As Integer, lt1 As Integer
   Dim RoundOffX As Double, RoundOffY As Double
   
   RoundOffX = 0.00001
   RoundOffY = 0.00001
   
   Dim VarS As Integer
   Dim VarL As Long
   Dim VarF As Single
   Dim VarD As Double
   
   Dim ht() As Double, htf() As Single, hts() As Integer
   Dim zmin As Double, zmax As Double
   
   Dim X1, Y1, X2, Y2
   Dim XPixStep As Double, YPixStep As Double
   
   Dim ConstantXGeo As Boolean
   Dim XGeoConstant As Double
   Dim YGeoConstant As Double
   
   Dim CalculateRotatedGrid As Boolean
   
   Dim CurrentX As Double, CurrentY As Double
   Dim ShiftX As Double, ShiftY As Double
   
   Dim ier As Integer
   Dim ContourInterval As Integer
   
   Dim Tolerance As Double, ANG1 As Double, ANG2 As Double
   On Error GoTo CreateDTMBackground_Error

   Tolerance = 0.00001
   
   ConstantXGeo = True
   
   RotateddGrid = False
    
   zmin = INIT_VALUE
   zmax = -INIT_VALUE
   
   If ULPixX = LRPixX Or ULPixY = LRPixY Then
        Call MsgBox("The pixel coordinatges of the maps upper left and/or lower right corners are undefined." _
                    & vbCrLf & "" _
                    & vbCrLf & "Use the Option menu to define them" _
                    , vbExclamation Or vbDefaultButton1, "DTM creation")
      ier = -1
      GDMDIform.Toolbar1.Buttons(45).Enabled = False
      CreateDTMBackground = ier
      Exit Function
      End If
   
   If ULPixX = LRPixX Or ULPixY = LRPixX Then
      Call MsgBox("You must first define the pixel coordinates of the four corners of the map!" _
                  & vbCrLf & "" _
                  & vbCrLf & "Use the Options dialog to define them." _
                  , vbInformation, "Creating background DTM")
      ier = -1
      CreateDTMBackground = ier
      Exit Function
      End If
      
   RectCoord(0).X = ULPixX
   RectCoord(0).Y = ULPixY
   RectCoord(1).X = LRPixX
   RectCoord(1).Y = LRPixY
   
   X1 = RectCoord(0).X
   Y1 = RectCoord(0).Y
   X2 = RectCoord(1).X
   Y2 = RectCoord(1).Y
   
   ier = 0

   'record elevations in the selected region
   
   If Not heights And Not (RSMethod0 Or RSMethod1 Or RSMethod2) Then
         
      ier = -1
      CreateDTMBackground = ier
      Exit Function
      
      End If
   
    Screen.MousePointer = vbHourglass
   
  'determine DTM spacing in pixels

  If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
     XStep = 25
     YStep = 25
  ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
     XStep = XStepITM
     YStep = YStepITM
  ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" And (DTMtype = 1 Or DTMtype = 2) Then
     XStep = XStepDTM * 8.33333333333333E-04 / 3#
     YStep = YStepDTM * 8.33333333333333E-04 / 3#
  Else
     XStep = XStepITM
     YStep = YStepITM
     End If
     
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

'      iprogress& = 0
'      find coordinates of the coordinates to see if the grid is square
'      if it is not square, then have to use iteration to convert the geo coordinates back to aproximate pixels,
'      and then use the rubbersheeting to find the exact geo coordinates, and then only give spacing to the accuracy of the dtm grid
'
  'convert the pixel search coordinates to geo coordinates
  For i = 0 To 1

     If RSMethod1 Then
        ier = RS_pixel_to_coord2(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
     ElseIf RSMethod2 Then
        ier = RS_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
     ElseIf RSMethod0 Then
        ier = Simple_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
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
 
 Dim XX As Long, YY As Long

 If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
    'now calculate the other corners to see how square the grid is
     'convert the pixel search coordinates to geo coordinates
     For i = 2 To 3
     
        If i = 2 Then 'SW corner
           XX = RectCoord(0).X
           YY = RectCoord(1).Y
        ElseIf i = 3 Then 'NE corner
           XX = RectCoord(1).X
           YY = RectCoord(0).Y
           End If
           
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod0 Then
           ier = Simple_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
           End If

        'now determine new boundaries so that boundaries are on data point
        'determine lat,lon of first element of this tile
    
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
    
     Next i
 
    If SearchGeoCoord(0).XGeo <> SearchGeoCoord(2).XGeo Or SearchGeoCoord(1).XGeo <> SearchGeoCoord(3).XGeo Or _
       SearchGeoCoord(0).YGeo <> SearchGeoCoord(3).YGeo Or SearchGeoCoord(1).YGeo <> SearchGeoCoord(2).YGeo Then
       
       CalculateRotatedGrid = True
      
    Else
       CalculateRotatedGrid = False
       End If
       
 ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
    CalculateRotatedGrid = False 'by definition, ITM is not rotated
    
 Else
     'now calculate the other corners to see how square the grid is
     'convert the pixel search coordinates to geo coordinates
     For i = 2 To 3
     
        If i = 2 Then 'SW corner
           XX = RectCoord(0).X
           YY = RectCoord(1).Y
        ElseIf i = 3 Then 'NE corner
           XX = RectCoord(1).X
           YY = RectCoord(0).Y
           End If
           
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod0 Then
           ier = Simple_pixel_to_coord(CDbl(XX), CDbl(YY), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
           End If

        'now determine new boundaries so that boundaries are on data point
        'determine lat,lon of first element of this tile
        SearchGeoCoord(i).XGeo = Int(SearchGeoCoord(i).XGeo / XStep) * XStep
        SearchGeoCoord(i).YGeo = Int(SearchGeoCoord(i).YGeo / YStep) * YStep
    
     Next i
 
    If SearchGeoCoord(0).XGeo <> SearchGeoCoord(2).XGeo Or SearchGeoCoord(1).XGeo <> SearchGeoCoord(3).XGeo Or _
       SearchGeoCoord(0).YGeo <> SearchGeoCoord(3).YGeo Or SearchGeoCoord(1).YGeo <> SearchGeoCoord(2).YGeo Then
       
       CalculateRotatedGrid = True
      
    Else
       CalculateRotatedGrid = False
       End If

    End If
    
If CalculateRotatedGrid Then '==========================================================================================================
    
    'define square grid using the geo coordinates, and determine equivalent pixel coordinates for drawing contours
    
    If LRGeoX = ULGeoX Or ULGeoY = LRGeoY Then
       'use rubber sheeting to determine them
       MsgBox "Corner grid coordinates undefined." & vbCrLf & vbCrLf & "(Hint: use options menu)", vbInformation + vbOKOnly, "DTM creation error"
       GDMDIform.Toolbar1.Buttons(45).Enabled = False
       ier = -1
       CreateDTMBackground = ier
       Exit Function
    Else
    
        GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
        GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
   
       End If
    
    'determine maximum x and y ranges
    Dim XmaxRange As Double, YmaxRange As Double, XminRange As Double, YminRange As Double
    XmaxRange = Max(CDbl(SearchGeoCoord(1).XGeo), CDbl(SearchGeoCoord(3).XGeo))
    YmaxRange = Max(CDbl(SearchGeoCoord(0).YGeo), CDbl(SearchGeoCoord(3).YGeo))
    XminRange = min(CDbl(SearchGeoCoord(0).XGeo), CDbl(SearchGeoCoord(2).XGeo))
    YminRange = min(CDbl(SearchGeoCoord(1).YGeo), CDbl(SearchGeoCoord(2).YGeo))
    
     numXsteps& = Int((XmaxRange - XminRange) / XStep) + 1 '(X2 - X1 + 1)
     numYsteps& = Int((YmaxRange - YminRange) / YStep) + 1 '(Y2 - Y1 + 1)
     
     If HeightPrecision = 0 Then
         ReDim hts(0 To numXsteps& - 1, 0 To numYsteps& - 1)
         ReDim ht(0)
         ReDim htf(0)
     ElseIf HeightPrecision = 1 Then
         ReDim htf(0 To numXsteps& - 1, 0 To numYsteps& - 1)
         ReDim hts(0)
         ReDim ht(0)
     ElseIf HeightPrecision = 2 Then
         ReDim ht(0 To numXsteps& - 1, 0 To numYsteps& - 1)
         ReDim hts(0)
         ReDim htf(0)
         End If
     
     numContourPnts& = numXsteps& * numYsteps&
     
     GDMDIform.StatusBar1.Panels(1).Text = "Generating basis DTM heights, please wait..."
     
     num_x& = 0
     
     For i = 1 To numXsteps& 'loop from west from east
     
         GeoX = XminRange + XStep * (i - 1)
         
         num_y& = 0
     
         For j = 1 To numYsteps& 'loop from south to north
         
            GeoY = YminRange + YStep * (j - 1)
            
            'determine estimate of the corresponding Pixel coordinates
         
            GoSub GeotoCoord
            
            'check that the pixel coordinates are within the boundaries of the map
            'if the coordinate grid is rotated significantly w.r.t. pixel grid, the corners can be rotated outside the map range
            If CurrentX < 0 Or CurrentX > pixwi Then
               numXsteps& = numXsteps& - 1
               num_x& = num_x& - 1
               If CurrentX > pixwi Then
                  SearchGeoCoord(1).XGeo = SearchGeoCoord(1).XGeo - XStep
               ElseIf CurrentX < 0 Then
                  SearchGeoCoord(0).XGeo = SearchGeoCoord(0).XGeo + XStep
                  End If
               Exit For
               End If
               
            If CurrentY < 0 Or CurrentY > pixhi Then
               numYsteps& = numYsteps& - 1
               If CurrentY > pixhi Then
                  SearchGeoCoord(1).YGeo = SearchGeoCoord(1).YGeo + YStep
               ElseIf CurrentY < 0 Then
                  SearchGeoCoord(0).YGeo = SearchGeoCoord(0).YGeo - YStep
                  End If
               Exit For
               End If
            
            'now determine the heights at these coordinates
            XGeo = GeoX
            YGeo = GeoY
            
            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                'convert from ITM to WGS84
                kmx = XGeo
                kmy = YGeo
                Call ics2wgs84(CLng(kmy), CLng(kmx), lt2, lg2)
            Else
                lg2 = XGeo
                lt2 = YGeo
                End If
                 
            If heights Then
                If DTMtype = 1 Then
                   'use ASTER
                   Call ASTERheight(lg2, lt2, hgt2)
                ElseIf DTMtype = 2 Then
                   'use JKH's DTM if ITM coordinates, else use NED, SRTM
                   If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                      Call DTMheight2(lg2, lt2, hgt2)
                   Else
                      Call worldheights(lg2, lt2, hgt2)
                      End If
                   End If
            Else
               hgt2 = -9999 'just record blanks
               End If
               
            If hgt2 <> -9999 Then 'not a void, so use it to determine zmin, zmax
                If hgt2 > zmax Then zmax = hgt2
                If hgt2 < zmin Then zmin = hgt2
                End If
                 
            If HeightPrecision = 0 Then
                hts(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 1 Then
                htf(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 2 Then
                ht(num_x&, num_y&) = hgt2
                End If
            
            num_y& = num_y& + 1
            
         Next j
         
         num_x& = num_x& + 1
         
         DoEvents
         
         Call UpdateStatus(GDMDIform, 1, 100 * num_x& / numXsteps&)
         
     Next i
   
     Screen.MousePointer = vbDefault
     
     GDMDIform.StatusBar1.Panels(1).Text = sEmpty
     GDMDIform.StatusBar1.Panels(2).Text = sEmpty
     Call UpdateStatus(GDMDIform, 1, 0)
     GDMDIform.picProgBar.Visible = False
       
ElseIf Not CalculateRotatedGrid Then '=============================================================================================
 
     'grid is square
     numXsteps& = Nint((SearchGeoCoord(1).XGeo - SearchGeoCoord(0).XGeo) / XStep) + 1   '(X2 - X1 + 1)
     numYsteps& = Nint((SearchGeoCoord(0).YGeo - SearchGeoCoord(1).YGeo) / YStep) + 1   '(Y2 - Y1 + 1)
     
     XPixStep = CDbl(X2 - X1) / numXsteps&  '1
     YPixStep = CDbl(Y2 - Y1) / numYsteps&  '1
     
'     numXsteps& = Int((X2 - X1 + 1) / XPixStep)
'     numYsteps& = Int((Y2 - Y1 + 1) / YPixStep)
     
     If HeightPrecision = 0 Then
        ReDim hts(0 To numXsteps& - 1, 0 To numYsteps& - 1)
        ReDim ht(0)
        ReDim htf(0)
     ElseIf HeightPrecision = 1 Then
        ReDim htf(0 To numXsteps& - 1, 0 To numYsteps& - 1)
        ReDim ht(0)
        ReDim hts(0)
     ElseIf HeightPrecision = 2 Then
        ReDim ht(0 To numXsteps& - 1, 0 To numYsteps& - 1)
        ReDim hts(0)
        ReDim htf(0)
        End If
     
     GDMDIform.StatusBar1.Panels(1).Text = "Generating basis DTM heights, please wait..."
     
     num_x& = 0
     
     For pixStepX = X1 To X2 + RoundOffX Step XPixStep
         
         num_y& = 0
     
         For pixStepY = Y2 To Y1 - RoundOffY Step -YPixStep
            
            'convert pixel coordinate to geographic cooridinate
          
            If RSMethod1 Then
               ier = RS_pixel_to_coord2(pixStepX, pixStepY, XGeo, YGeo)
            ElseIf RSMethod2 Then
               ier = RS_pixel_to_coord(pixStepX, pixStepY, XGeo, YGeo)
            ElseIf RSMethod0 Then
               ier = Simple_pixel_to_coord(pixStepX, pixStepY, XGeo, YGeo)
               End If
               
            If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" And ConstantXGeo And RSMethod2 Then
                'rotate plot to rectangular grid based on the upper-left corner value of the geographic coordinates
                If ConstantXGeo And RSMethod2 And (pixStepX = X1 And pixStepY = Y2) Then
                   XGeoConstant = XGeo
                   YGeoConstant = YGeo
                ElseIf ConstantXGeo And RSMethod2 And Not (pixStepX = CDbl(X1) And pixStepY = CDbl(Y2)) Then
                   'ignore rubber sheeting coordinates, rather use square grid.
                   XGeo = XGeoConstant + ((pixStepX - CDbl(X1)) / XPixStep) * XStep
                   YGeo = YGeoConstant + ((CDbl(Y2) - pixStepY) / YPixStep) * YStep
                   End If
                End If
    
            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                'convert from ITM to WGS84
                kmx = XGeo
                kmy = YGeo
                Call ics2wgs84(CLng(kmy), CLng(kmx), lt2, lg2)
            Else
                lg2 = XGeo
                lt2 = YGeo
                End If
                 
            If heights Then
                If DTMtype = 1 Then
                   'use ASTER
                   Call ASTERheight(lg2, lt2, hgt2)
                ElseIf DTMtype = 2 Then
                   'use JKH's DTM if ITM coordinates, else use NED, SRTM
                   If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                      Call DTMheight2(lg2, lt2, hgt2)
                   Else 'use NED (in SRTM format) or SRTM 1 arcsec
                      Call worldheights(lg2, lt2, hgt2)
                      End If
                   End If
            Else
               'use blank heights
               hgt2 = -9999
               End If
               
            If hgt2 <> -9999 Then 'not a void, use it for determine elevation bounds
                If hgt2 > zmax Then zmax = hgt2
                If hgt2 < zmin Then zmin = hgt2
                End If
                 
            If num_y& > numYsteps& - 1 Then Exit For
            
            If HeightPrecision = 0 Then
                hts(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 1 Then
                htf(num_x&, num_y&) = hgt2
            ElseIf HeightPrecision = 2 Then
                ht(num_x&, num_y&) = hgt2
                End If
            
            num_y& = num_y& + 1
            
         Next pixStepY
         
         num_x& = num_x& + 1
         
         If num_x& > numXsteps& - 1 Then Exit For
         
         DoEvents
         
         Call UpdateStatus(GDMDIform, 1, 100 * num_x& / (numXsteps& - 1))
         
     Next pixStepX
   
     Screen.MousePointer = vbDefault
     
     GDMDIform.StatusBar1.Panels(1).Text = sEmpty
     GDMDIform.StatusBar1.Panels(2).Text = sEmpty
     Call UpdateStatus(GDMDIform, 1, 0)
     GDMDIform.picProgBar.Visible = False
     
     End If
    
   '-------------------generate actual basis DTM----------------------------
    GDMDIform.StatusBar1.Panels(1).Text = "Generating basis DTM, please wait......"
   
    'create DTM file into dirNewDTM
    'write picnam_Boundaries.txt file
    
    'create DTM with the Surfer ver 7 binary grid file format
    'the DTM consists of the old DTM heights for grid points outside the edited region and
    'uses the edited DTM heights for places within the edited region
       
    GDMDIform.StatusBar1.Panels(1).Text = "Creating DTM file, please wait...."
    DTMfile$ = dirNewDTM & "\" & RootName(picnam$) & ".grd"
       
    filedtm% = FreeFile
    Open DTMfile$ For Binary As #filedtm%
     
    filehdr% = FreeFile
    Open dirNewDTM & "\" & RootName(picnam$) & ".hdr" For Output As #filehdr%
     
    BytePosit = 1
    VarL = &H42525344
    Put #filedtm%, BytePosit, VarL 'header tag
    BytePosit = 5
    VarL = 4
    Put #filedtm%, BytePosit, VarL
    BytePosit = 9
    VarL = 1
    Put #filedtm%, BytePosit, VarL
    BytePosit = 13
    VarL = &H44495247 'tag for grid section
    Put #filedtm%, BytePosit, VarL
    BytePosit = 17
    VarL = 72
    Put #filedtm%, BytePosit, VarL
    BytePosit = 21
    VarL = numYsteps& 'nRowLL
    nRowLL = VarL
    Put #filedtm%, BytePosit, VarL
    Print #filehdr%, VarL
    BytePosit = 25
    VarL = numXsteps& 'nColLL
    nColLL = VarL
    Put #filedtm%, BytePosit, VarL
    Print #filehdr%, VarL
    BytePosit = 29
    VarD = SearchGeoCoord(0).XGeo 'xLL = min GeoX
    xLL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 37
    VarD = SearchGeoCoord(1).YGeo 'yLL = min GeoY
    yLL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 45
    VarD = XStep
    XStepLL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 53
    VarD = YStep
    YStepLL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 61
    VarD = zmin
    zminLL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 69
    VarD = zmax
    zmaxLL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 77
    VarD = ANG
    AngLL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 85
    VarD = blank_value  'flags blank data value
    blank_LL = VarD
    Put #filedtm%, BytePosit, VarD
    Print #filehdr%, VarD
    BytePosit = 93
    VarL = &H41544144 'tag for data section
    Put #filedtm%, BytePosit, VarL
    BytePosit = 97
    VarL = numXsteps& * numYsteps& * 8 'byte size of data section
    Put #filedtm%, BytePosit, VarL
     
     BytePosit = 101
     Call UpdateStatus(GDMDIform, 1, 0)
     For j = 0 To numYsteps& - 1
        For i = 0 To numXsteps& - 1
           
           If HeightPrecision = 0 Then
              VarD = hts(i, j) * 1#
           ElseIf HeightPrecision = 1 Then
              VarD = htf(i, j) * 1#
           ElseIf HeightPrecision = 2 Then
              VarD = ht(i, j)
              End If
              
           If VarD = -9999 Then
              VarD = blank_value 'this is a void data
           Else
              VarD = VarD * DigiConvertToMeters
              End If
           
           Put #filedtm%, BytePosit, VarD
           BytePosit = BytePosit + 8
        Next i
        DoEvents
        If numYsteps& > 1 Then Call UpdateStatus(GDMDIform, 1, CLng(100 * j / numYsteps&))
     Next j
     
     Close #filedtm%
     Close #filehdr%
    
                
     GDMDIform.picProgBar.Visible = False
     GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
     
    '-------------------------finished----------------------------------
   CreateDTMBackground = ier

   On Error GoTo 0
   Exit Function
   
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
'                        , vbInformation, "Creaate DTM background error")
'              ier = -1
'              CreateDTMBackground = ier
'              Screen.MousePointer = vbDefault
'              GDMDIform.picProgBar.Visible = False
'              GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
'              GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
'              Exit Function
'              End If
              
           End If
Return

CreateDTMBackground_Error:

   Screen.MousePointer = vbDefault
   GDMDIform.StatusBar1.Panels(1).Text = sEmpty
   GDMDIform.StatusBar1.Panels(2).Text = sEmpty
    
   GDMDIform.picProgBar.Visible = False

   ier = -1
   CreateDTMBackground = ier
   
End Function
'---------------------------------------------------------------------------------------
' Procedure : OpenCloseBaseDTM
' Author    : Dr-John-K-Hall
' Date      : 12/23/2015
' Purpose   : Opens and closes the base DTM
'---------------------------------------------------------------------------------------
'
Public Function OpenCloseBaseDTM(mode%) As Integer
   'opens the base DTM for reading if mode% = 0
   'close the base DTM when mode% = 1
   
   On Error GoTo OpenCloseBaseDTM_Error
   
   Dim EditDTM As Boolean

   Select Case mode%
   
     Case 0 'open it
        If basedtm% > 0 Then
           Close #basedtm%
           basedtm% = 0
           End If
           
        DTMfile$ = dirNewDTM & "\" & RootName(picnam$) & ".grd"
        DTMhdrfile$ = dirNewDTM & "\" & RootName(picnam$) & ".hdr"
        If Dir(DTMfile$) <> gsEmpty And Dir(DTMhdrfile$) <> gsEmpty Then
            'read header information
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
            
            'read merge header info to determine current zmin, zmax
            numLines = 0
            Do Until EOF(filhdr%)
               Input #filhdr%, XX
               numLines = numLines + 1
               If numLines = 7 Then
                  If XX < zminLL And XX <> blank_value Then
                     zminLL = XX
                     EditDTM = True
                     End If
               ElseIf numLines = 8 Then
                   If XX > zmaxLL And XX <> blank_value Then
                     zmaxLL = XX
                     EditDTM = True
                     End If
               ElseIf numLines = 11 Then
                   numLines = 1
                   End If
            Loop
            
            Close #filhdr%
           
            basedtm% = FreeFile
            Open DTMfile$ For Binary As #basedtm%
            BasisDTMheights = True
            
            'now update the zmin, zmax inside the file itself
            If EditDTM Then
                Dim BytePosit As Long
                Dim VarD As Double
                
                BytePosit = 61
                Get #basedtm%, BytePosit, VarD
                If zminLL < VarD Then
                   VarD = zminLL
                   Put #basedtm%, BytePosit, VarD
                   End If
                BytePosit = 69
                Get #basedtm%, BytePosit, VarD
                If zmaxLL > VarD Then
                   VarD = zmaxLL
                   Put #basedtm%, BytePosit, VarD
                   End If
                End If
            
            End If
     
     Case 1 'close it
        If basedtm% > 0 Then
           Close #basedtm%
           BasisDTMheights = False
           basedtm% = 0
           End If
     
  End Select

   On Error GoTo 0
   Exit Function

OpenCloseBaseDTM_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenCloseBaseDTM of Module modConvertHall"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Smoothing
' Author    : Dr-John-K-Hall
' Date      : 12/30/2015
' Purpose   : Uses Belgier leas square spline to smooth merged DTM perpendicular to the long axis
' Source    : Spline code: http://files.codes-sources.com/fichier.aspx?id=18366&f=Splines.bas
'---------------------------------------------------------------------------------------
'
Public Function Smoothing(Pic As PictureBox, RectCoord() As POINTAPI) As Integer
   
   Dim XStep As Double, YStep As Double
   Dim SizeX As Double, SizeY As Double
   
   Dim SearchGeoCoord(1) As POINTGEO
   Dim i As Long, j As Long
   Dim XGeo As Double
   Dim YGeo As Double
   Dim GeoX As Double
   Dim GeoY As Double
   Dim numXsteps&, numYsteps&
   
   Dim X1, Y1, X2, Y2
   
   Dim zmin As Double, zmax As Double
   
   Dim ier As Integer
   
   Dim VarD As Double
   Dim BytePosit As Long
   
   Dim GeoToPixelX As Double, GeoToPixelY As Double
   Dim CurrentX As Double, CurrentY As Double
   
   On Error GoTo Smoothing_Error
      
   GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
   GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
   
   DLLMethod = True
   
   ier = 0

   'create or fix old dtm

    X1 = RectCoord(0).X
    Y1 = RectCoord(0).Y
    X2 = RectCoord(1).X
    Y2 = RectCoord(1).Y

    ier = 0

    'smoothing basis DTM on the basis DTM

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
        MsgBox "Background DTM" & DTMfile$ & " is missing..."
        Smoothing = -1
        Exit Function
        End If


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
    GDMDIform.StatusBar1.Panels(1).Text = "Determining bounds, and screen coordinates of smoothing region, please wait...."

'     'convert the pixel search coordinates to geo coordinates
     For i = 0 To 1

        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
        ElseIf RSMethod0 Then
           ier = Simple_pixel_to_coord(CDbl(RectCoord(i).X), CDbl(RectCoord(i).Y), SearchGeoCoord(i).XGeo, SearchGeoCoord(i).YGeo)
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

     'now determine number os steps

     numXsteps& = Int((SearchGeoCoord(1).XGeo - SearchGeoCoord(0).XGeo) / XStep) + 1
     numYsteps& = Int((SearchGeoCoord(0).YGeo - SearchGeoCoord(1).YGeo) / YStep) + 1

     Dim hgt() As Double, hgt2 As Double
     Dim CB_Tav#()
     Dim NPC_1 As Long
     Dim NPI_1 As Long
     Dim NPI_1_O As Long
     Dim u As Double
     Dim ue As Double
     Dim u_1 As Double
     Dim K As Long, N As Long
     
'     '/////////////////diagnositics///////////////////
'     filediag% = FreeFile
'     Open App.Path & "\Diagnostics.txt" For Append As #filediag%
'     '//////////////////////////////////////////////////

     If UseNewDTM% = 1 Then 'reopen database for editing
        ier = OpenCloseBaseDTM(0)
        End If

     If numXsteps& > numYsteps& Then

        'anneal in the y direction
        NPC_1 = numYsteps& - 1 'number of vertices to fit
        NPI_1 = NPC_1 'number of knots NPI_1 is initially set to NPC_1, in general NPI_1 <= NPC_1
        ReDim hgt(0 To NPC_1)

       'prepare spline binomial coeficients
        If NPI_1 <> NPI_1_O Then
            GDMDIform.StatusBar1.Panels(1).Text = "Preparing binomial coeficients for the spline, please wait...."
            ReDim CB_Tav#(0 To NPI_1)
            Call UpdateStatus(GDMDIform, 1, 0)
            For K = 0 To NPI_1
                CB_Tav(K) = rncr(NPI_1, K)
                Call UpdateStatus(GDMDIform, 1, CLng(100 * K / NPI_1))
            Next K
            NPI_1_O = NPI_1
            End If

        Call UpdateStatus(GDMDIform, 1, 0)

        For i = 0 To numXsteps& - 1
'           '/////////////diagnositcs///////////////
'           Print #filediag%, "==================original heights - annealing in y direc=========================="
'           '///////////////////////////////////////////////////////////////////
           GeoX = SearchGeoCoord(0).XGeo + (i - 1) * XStep
           K = 0
           NPI_1 = NPC_1 'if there are no voids, then the number of vertices equal to the number of steps
           For j = 0 To NPC_1
               GeoY = SearchGeoCoord(1).YGeo + (j - 1) * YStep
               'extract heights and stuff them into spline array
               'skip voids
               BytePosit = 101 + 8 * ((GeoX - xLL) / XStep) + 8 * nColLL * ((GeoY - yLL) / YStep)
               Get #basedtm%, BytePosit, VarD
               If VarD = blank_value Then
                  'this is void, so skip it
                  NPI_1 = NPI_1 - 1
               Else
                  hgt(K) = VarD
'                  '///////////////diagnostics////////////////////////////////
'                  Write #filediag%, GeoX, GeoY, hgt(K)
'                  '/////////////////////////////////////////////////////////
                  K = K + 1
                  End If
           Next j
           
           If NPI_1 <> NPI_1_O Then 'redo spline coeficients
              ReDim CB_Tav#(0 To NPI_1)
              For K = 0 To NPI_1
                  CB_Tav(K) = rncr(NPI_1, K)
              Next K
              NPI_1_O = NPI_1
              End If

           'do spline and replace previous heights and voids
           'ignore small shifts in x,y and only calculate changes in height
           'heights at endpoints retain their values
'           '/////////////diagnositcs///////////////
'           Print #filediag%, "==================fitted heights=========================="
'           '///////////////////////////////////////////////////////////////////
            For j = 1 To NPC_1 - 1
                u = CDbl(j) / CDbl(NPC_1)
                ue = 1#
                u_1 = 1# - u
                hgt2 = 0#
                For K = 0 To NPI_1
                    BF = CB_Tav(K) * ue * u_1 ^ (NPI_1 - K)
                    hgt2 = hgt2 + hgt(K) * BF
                    ue = ue * u
                Next K

                'now overwrite height with smoothed height
                If HeightPrecision = 0 Then
                   VarD = Nint(hgt2)
                ElseIf HeightPrecision = 1 Then
                   VarD = CSng(hgt2)
                ElseIf HeightPrecision = 2 Then
                   VarD = hgt2
                   End If
                   
                GeoY = SearchGeoCoord(1).YGeo + (j - 1) * YStep

'                '///////////////diagnostics////////////////////////////////
'                Write #filediag%, GeoX, GeoY, hgt2
'                '/////////////////////////////////////////////////////////
                
                BytePosit = 101 + 8 * ((GeoX - xLL) / XStep) + 8 * nColLL * ((GeoY - yLL) / YStep)
                Put #basedtm%, BytePosit, VarD

                'adjust zminLL, zmaxLL
                zmin = min(zminLL, VarD)
                zmax = Max(zmaxLL, VarD)

            Next j
            DoEvents
            Call UpdateStatus(GDMDIform, 1, CLng(100 * i / (numXsteps& - 1)))
        Next i
     Else
        'anneal in the x direction
        NPC_1 = numXsteps& - 1
        NPI_1 = NPC_1 'number of knots NPI_1 is initially set to NPC_1, in general NPI_1 <= NPC_1
        ReDim hgt(0 To NPC_1)

       'prepare spline binomial coeficients
        If NPI_1 <> NPI_1_O Then
            GDMDIform.StatusBar1.Panels(1).Text = "Preparing binomial coeficients for the spline, please wait...."
            ReDim CB_Tav#(0 To NPI_1)
            Call UpdateStatus(GDMDIform, 1, 0)
            For K = 0 To NPI_1
                CB_Tav(K) = rncr(NPI_1, K)
                Call UpdateStatus(GDMDIform, 1, CLng(100 * K / NPI_1))
            Next K
            NPI_1_O = NPI_1
            End If

        Call UpdateStatus(GDMDIform, 1, 0)
        
        For i = 0 To numYsteps& - 1
'           '/////////////diagnositcs///////////////
'           Print #filediag%, "==================original heights - annealing in x direc=========================="
'           '///////////////////////////////////////////////////////////////////
           GeoY = SearchGeoCoord(1).YGeo + (i - 1) * YStep
           K = 0
           NPI_1 = NPC_1 'if there are no voids, then the number of vertices equal to the number of steps
           For j = 0 To NPC_1
               GeoX = SearchGeoCoord(0).XGeo + (j - 1) * XStep
               'extract heights and stuff them into spline array
               'skip voids
               BytePosit = 101 + 8 * ((GeoX - xLL) / XStep) + 8 * nColLL * ((GeoY - yLL) / YStep)
               Get #basedtm%, BytePosit, VarD
               If VarD = blank_value Then
                  'this is void, so skip it
                  NPI_1 = NPI_1 - 1
               Else
                  hgt(K) = VarD
'                  '///////////////diagnostics////////////////////////////////
'                  Write #filediag%, GeoX, GeoY, hgt(K)
'                  '/////////////////////////////////////////////////////////
                  K = K + 1
                  End If
           Next j
           
           If NPI_1 <> NPI_1_O Then 'redo spline coeficients
              ReDim CB_Tav#(0 To NPI_1)
              For K = 0 To NPI_1
                  CB_Tav(K) = rncr(NPI_1, K)
              Next K
              NPI_1_O = NPI_1
              End If

           'do spline and replace heights
           'ignore small shifts in x,y and only calculate changes in height
           'heights at endpoints retain their values
'           '/////////////diagnositcs///////////////
'           Print #filediag%, "==================fitted heights=========================="
'           '///////////////////////////////////////////////////////////////////
            For j = 1 To NPC_1 - 1
                u = CDbl(j) / CDbl(NPC_1)
                ue = 1#
                u_1 = 1# - u
                hgt2 = 0#
                For K = 0 To NPI_1
                    BF = CB_Tav(K) * ue * u_1 ^ (NPI_1 - K)
                    hgt2 = hgt2 + hgt(K) * BF
                    ue = ue * u
                Next K

                'now overwrite height with smoothed height
                If HeightPrecision = 0 Then
                   VarD = Nint(hgt2)
                ElseIf HeightPrecision = 1 Then
                   VarD = CSng(hgt2)
                ElseIf HeightPrecision = 2 Then
                   VarD = hgt2
                   End If
                   
                GeoX = SearchGeoCoord(0).XGeo + (j - 1) * XStep

'                '///////////////diagnostics////////////////////////////////
'                Write #filediag%, GeoX, GeoY, hgt2
'                '/////////////////////////////////////////////////////////
                
                BytePosit = 101 + 8 * ((GeoX - xLL) / XStep) + 8 * nColLL * ((GeoY - yLL) / YStep)
                Put #basedtm%, BytePosit, VarD

                'adjust zminLL, zmaxLL
                zmin = min(zminLL, VarD)
                zmax = Max(zmaxLL, VarD)

            Next j
            DoEvents
            Call UpdateStatus(GDMDIform, 1, CLng(100 * i / (numYsteps& - 1)))

        Next i
        End If
        
'    '/////////////////diagnositcs//////////////////
'    Close #filediag%
'    '/////////////////////////////////////////////

    'determine new zmin, zmax over entire base dtm to see if it changed
    
    Dim ReWrite As Boolean
    'now update the zmin, zmax if necessary
    If zmin < zminLL Then
        ReWrite = True
        BytePosit = 61
        VarD = zmin
        zminLL = zmin
        Put #basedtm%, BytePosit, VarD
        End If
    If zmax > zmaxLL Then
        ReWrite = True
        BytePosit = 69
        VarD = zmax
        zmaxLL = zmax
        Put #basedtm%, BytePosit, VarD
        End If

    'now add updated header info containing the new zmin, zmax if necessary
    DTMhdrfile$ = dirNewDTM & "\" & RootName(picnam$) & ".hdr"
    If Dir(DTMhdrfile$) <> gsEmpty And ReWrite Then
        filehdr% = FreeFile
        Open DTMhdrfile$ For Append As #filehdr%
        Write #filhdr%, nRowLL
        Write #filhdr%, nColLL
        Write #filhdr%, xLL
        Write #filhdr%, yLL
        Write #filhdr%, XStepLL
        Write #filhdr%, YStepLL
        Write #filhdr%, zminLL
        Write #filhdr%, zmaxLL
        Write #filhdr%, AngLL
        Write #filhdr%, blank_LL
        Close #filehdr%
        End If

    GDMDIform.prbSearch.Visible = False
    GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
    GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
    Screen.MousePointer = vbDefault
    GDMDIform.picProgBar.Visible = False

    'reopen baseDTM for reading heights
'    If UseNewDTM% = 1 Then
'       ier = OpenCloseBaseDTM(0)
'       End If

   Smoothing = ier

   Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Function

Smoothing_Error:
   Select Case Err.Number
      Case 52
         'problem with base dtm's file number
         'close it and reopen
         ier = OpenCloseBaseDTM(0)
         Resume
      Case 63
         'bad record number caused by being off the map sheet
         'return the blank height value
         VarD = blank_value
         Resume Next
    End Select
    
    GDMDIform.picProgBar.Visible = False
    Screen.MousePointer = vbDefault
    ier = -1
    Smoothing = ier
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Smoothing of Module modConvertHall"

End Function
Private Function rncr(ByVal N&, ByVal K&) As Double

'Adapted from Gli algorithms, P. Bourke, D. Cholasuk http://files.codes-sources.com/fichier.aspx?id=18366&f=Splines.bas
'Compute the binomial coefficients Cn, k as:
'Rncr = N!/(K! * (N - K)!)
'
'Note: this function only makes sense for 0 < N, K <= N
'and 0 <= K.  No error trapping, so beware

    Dim i&, rncr_T#
'
    If ((N < 1) Or (K < 1) Or (N = K)) Then
        rncr = 1#
'
    Else
        rncr_T = 1#
        For i = 1 To N - K
            rncr_T = rncr_T * (1# + CDbl(K) / CDbl(i))
        Next i
'
        rncr = rncr_T
    End If
End Function
