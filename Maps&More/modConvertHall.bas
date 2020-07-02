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
    dY As Double
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
Public dtmdir As String
Public ASTERdir As String

'**********DTM variables**************
Public CHMAP(14, 26) As String * 2, filnumg%
Public CHMNE As String * 2, CHMNEO As String * 2, SF As String * 2
Public stepx As Double, stepy As Double
Public DTMtype As Integer
Public ASTERbilOpen As Boolean
Public ASTERNorth As Integer
Public ASTEREast As Integer
Public ASTERfilename As String
Public ASTERNrows%
Public ASTERNcols%
Public ASTERxdim As Double
Public ASTERydim As Double
Public ASTERfilnum%
Public tiledir As String


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
    lat = lat84 * 180 / pi
    lon = lon84 * 180 / pi
End Sub


'//=================================================
'// WGS84 to Israel New Grid (ITM) conversion
'//=================================================
Public Sub wgs842itm(lat As Double, lon As Double, N As Long, E As Long)
    Dim latr As Double, lonr As Double
    latr = lat * pi / 180
    lonr = lon * pi / 180

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
    lat = lat84 * 180 / pi
    lon = lon84 * 180 / pi
    '//printf("final lat, lon = %f, %f", lat, lon);
    '//pause();
End Sub


'//=================================================
'// WGS84 to Israel Old Grid (ICS) conversion
'//=================================================
Public Sub wgs842ics(lat As Double, lon As Double, N As Long, E As Long)
    Dim latr As Double, lonr As Double
    latr = lat * pi / 180
    lonr = lon * pi / 180

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
    
    Dim Y As Double, X As Double, M As Double
    Y = North + g_grid.false_n
    X = East - g_grid.false_e
    M = Y / g_grid.k0

    Dim a As Double, b As Double, E As Double, esq As Double
    a = g_datum.a
    b = g_datum.b
    E = g_datum.E
    esq = g_datum.esq

    Dim mu As Double
    mu = M / (a * (1 - E * E / 4 - 3 * (E ^ 4) / 64# - 5 * (E ^ 6) / 256#))

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
    M = a * (l1 * lat - l2 * Sin(2 * lat) + l3 * Sin(4 * lat) - l4 * Sin(6 * lat))
    '//double rho = a*(1-e2) / pow((1-(e*slat1)*(e*slat1)),1.5);
    Dim nu As Double, p As Double, k0 As Double
    nu = a / Sqr(1 - (E * slat1) * (E * slat1))
    p = lon - g_grid.lon0
    k0 = g_grid.k0
    '// y = northing = K1 + K2p2 + K3p4, where
    Dim K1 As Double, K2 As Double, K3 As Double
    K1 = M * k0
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
    Dim dX As Double, dY As Double, dZ As Double
    Call LoadGrid(-1, eDatumfrom)
    dX = g_datum.dX
    dY = g_datum.dY
    dZ = g_datum.dZ
    Call LoadGrid(-1, eDatumto)
    dX = dX - g_datum.dX
    dY = dY - g_datum.dY
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
    dlat = (-dX * slat * clon - dY * slat * slon + dZ * clat _
                   + da * rn * from_esq * slat * clat / from_a + _
                   df * (rm * adb + rn / adb) * slat * clat) / (rm + from_h)

    '// result lat (radians)
    olat = ilat + dlat

    '// dlon = (-dx * slon + dy * clon) / ((rn + from.h) * clat);
    Dim dlon As Double
    dlon = (-dX * slon + dY * clon) / ((rn + from_h) * clat)
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
            g_datum.dY = 0 ',
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
            g_datum.dY = 55 ',
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
            g_datum.dY = -85 ',
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
Public Sub DTMheight(kmx As Double, kmy As Double, hgt2 As Integer)
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
      
      ASTERfilename = ltch$ & Format(Trim$(Str$(Abs(ASTERNorth))), "00") & lgch$ & Format(Trim$(Str$(Abs(ASTEREast))), "000")
    
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
               ASTERNrows% = Val(Mid$(doclin$, pos1% + 5, Len(doclin$)))
            ElseIf InStr(doclin$, "NCOLS") <> 0 Then
               pos2% = InStr(doclin$, "NCOLS")
               ASTERNcols% = Val(Mid$(doclin$, pos2% + 5, Len(doclin$)))
            ElseIf InStr(doclin$, "XDIM") <> 0 Then
               pos3% = InStr(doclin$, "XDIM")
               ASTERxdim = Val(Mid$(doclin$, pos3% + 4, Len(doclin$)))
            ElseIf InStr(doclin$, "YDIM") <> 0 Then
               pos4% = InStr(doclin$, "YDIM")
               ASTERydim = Val(Mid$(doclin$, pos4% + 4, Len(doclin$)))
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
        NROW% = IKMY%: NCOL% = IKMX%

'       GETZ FINDS THE HEIGHT OF A POINT AT THE NORW AND NCOL FROM 380N
'       AND -20E WHERE 1,1 IS THAT CORNER POINT
'       FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
g15:    Jg% = 1 + Int((NROW% - 2) / 800)
        Ig% = 1 + Int((NCOL% - 2) / 800)
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
g21:    IR% = NROW% - (Jg% - 1) * 800
        IC% = NCOL% - (Ig% - 1) * 800
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
Sub InitializeDTM()
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
End Sub
'4320
'2.314814815 e-4
Sub ASTERheight(lt As Double, lg2 As Double, hgt2 As Integer)

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
      
      ASTERfilename = ltch$ & Format(Trim$(Str$(Abs(ASTERNorth))), "00") & lgch$ & Format(Trim$(Str$(Abs(ASTEREast))), "000")
    
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
               ASTERNrows% = Val(Mid$(doclin$, pos1% + 5, Len(doclin$)))
            ElseIf InStr(doclin$, "NCOLS") <> 0 Then
               pos2% = InStr(doclin$, "NCOLS")
               ASTERNcols% = Val(Mid$(doclin$, pos2% + 5, Len(doclin$)))
            ElseIf InStr(doclin$, "XDIM") <> 0 Then
               pos3% = InStr(doclin$, "XDIM")
               ASTERxdim = Val(Mid$(doclin$, pos3% + 4, Len(doclin$)))
            ElseIf InStr(doclin$, "YDIM") <> 0 Then
               pos4% = InStr(doclin$, "YDIM")
               ASTERydim = Val(Mid$(doclin$, pos4% + 4, Len(doclin$)))
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
