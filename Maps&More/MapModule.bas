Attribute VB_Name = "MapModule"
'*****************Windows API functions, subroutines and constants*********
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_GETTEXT = &HD
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SWP_SHOWWINDOW = &H40
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const WM_COMMAND = &H111
Public Const WM_SETFOCUS = &H7
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetVersion Lib "Kernel32" () As Long
Declare Function WinExec Lib "Kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
'Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWNORMAL = 1
'Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function EmptyClipboard Lib "user32" () As Long
'Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Public Const GW_CHILD = 5
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_HIDEWINDOW = &H80
'Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
'Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOZORDER = &H4
'Public Const WM_GETTEXT = &HD
Public Const KEYEVENTF_KEYUP = &H2
'Public Const RIGHT_ALT_PRESSED = &H1     '  the right alt key is pressed.
'Public Const LEFT_ALT_PRESSED = &H2     '  the left alt key is pressed.
'Public Const CF_TEXT = 1
Public Const SW_SHOW = 5
'Public Const SW_MAXIMIZE = 3
'Public Const SW_SHOWMAXIMIZED = 3
'Public Const SW_RESTORE = 9
'Public Const KF_ALTDOWN = &H2000
Public Const VK_SNAPSHOT = &H2C
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
'Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
'Public Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
'Public Const VK_ADD = &H6B
'Public Const VK_ATTN = &HF6
'Public Const VK_BACK = &H8
'Public Const VK_CANCEL = &H3
'Public Const VK_CAPITAL = &H14
'Public Const VK_CLEAR = &HC
Public Const VK_CONTROL = &H11
'Public Const VK_CRSEL = &HF7
'Public Const VK_DECIMAL = &H6E
Public Const VK_DELETE = &H2E
'Public Const VK_DIVIDE = &H6F
Public Const VK_DOWN = &H28
Public Const VK_END = &H23
'Public Const VK_EREOF = &HF9
Public Const VK_ESCAPE = &H1B
'Public Const VK_EXECUTE = &H2B
'Public Const VK_EXSEL = &HF8
Public Const VK_F1 = &H70
Public Const VK_F10 = &H79
'Public Const VK_F11 = &H7A
'Public Const VK_F12 = &H7B
'Public Const VK_F13 = &H7C
'Public Const VK_F14 = &H7D
'Public Const VK_F15 = &H7E
'Public Const VK_F16 = &H7F
'Public Const VK_F17 = &H80
'Public Const VK_F18 = &H81
'Public Const VK_F19 = &H82
'Public Const VK_F2 = &H71
'Public Const VK_F20 = &H83
'Public Const VK_F21 = &H84
'Public Const VK_F22 = &H85
'Public Const VK_F23 = &H86
'Public Const VK_F24 = &H87
'Public Const VK_F3 = &H72
'Public Const VK_F4 = &H73
'Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
'Public Const VK_F9 = &H78
'Public Const VK_HELP = &H2F
Public Const VK_HOME = &H24
Public Const VK_INSERT = &H2D
'Public Const VK_LBUTTON = &H1
Public Const VK_LCONTROL = &HA2
'Public Const VK_LEFT = &H25
Public Const VK_LMENU = &HA4
'Public Const VK_LSHIFT = &HA0
'Public Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON
'Public Const VK_MENU = &H12
'Public Const VK_MULTIPLY = &H6A
'Public Const VK_NEXT = &H22
'Public Const VK_NONAME = &HFC
'Public Const VK_NUMLOCK = &H90
'Public Const VK_NUMPAD0 = &H60
'Public Const VK_NUMPAD1 = &H61
'Public Const VK_NUMPAD2 = &H62
'Public Const VK_NUMPAD3 = &H63
'Public Const VK_NUMPAD4 = &H64
'Public Const VK_NUMPAD5 = &H65
'Public Const VK_NUMPAD6 = &H66
'Public Const VK_NUMPAD7 = &H67
'Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
'Public Const VK_PAUSE = &H13
'Public Const VK_PLAY = &HFA
'Public Const VK_RBUTTON = &H2
'Public Const VK_PROCESSKEY = &HE5
'Public Const VK_RCONTROL = &HA3
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
'Public Const VK_RMENU = &HA5
'Public Const VK_RSHIFT = &HA1
'Public Const VK_SCROLL = &H91
'Public Const VK_SELECT = &H29
'Public Const VK_SEPARATOR = &H6C
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_SUBTRACT = &H6D
Public Const VK_TAB = &H9
Public Const VK_UP = &H26
'Public Const VK_ZOOM = &HFB
'Public Const VOS__BASE = &H0&
Public Const sEmpty = ""

Public Const hgtobs As Double = 1.8

'*********************program constants and parameters*************
Public Const cd As Double = 1.74532927777778E-02 'conv deg to rad
Public Const pi As Double = 3.14159265358979
Public Const defaultmapwidth = 4206.93 'default width of map form
Public Const defaultmapheight = 7086.3 'default height of map form
Public Const defaultmaptop = 1480 'default top of map form
'Public Const travelmax% = 3000 'maximum number of travel points
Public Const traveltimerinteral = 3000

'*******************public variables and arrays************************************************
Public km400x, km400y, km50x, km50y, kmwx, kmwy, sizex, sizey, sizewx, sizewy
Public leftim%, pic1$(9), pic2(9) As String * 2, skyleftjump As Boolean, TempFormVis As Boolean
Public map400 As Boolean, map50 As Boolean, sunmode%, hChild As Long, Skynum%
'Public n400%(4, 7), n50%(19, 45), terwt As Boolean, cir50 As Boolean, cir400 As Boolean
Public n400x%(9), n400y%(9), n50x%(9), n50y%(9), Skycoord%, OverhWnd As Long
Public kmxoo, kmyoo, filnumg%, hgt, hgtpos, picf$, sgnfudx, xpix As Integer, ypix As Integer
'Public CHMAP(14, 26) As String * 2, CHMNE As String * 2, CHMNEO As String * 2, SF As String * 2
Public coordmode%, dragx, dragy, kmxc, kmyc, nplac%, Tdxname As String, MapLatCenter As Double, MapLonCenter As Double
Public X50c As Single, Y50c As Single, kmx50c, kmy50c, hgt50c, TdxhWnd As Long, SavedAll As Boolean
Public X400c As Single, Y400c As Single, kmx400c, kmy400c, hgt400c, bn%(4)
Public kmxorigin, kmyorigin, picold$(9), bufx(2, 4), bufy(2, 4), bufwi(2, 4), bufhi(2, 4)
Public maphi, mapwi, maphi2, mapwi2, nhwnd As Long, tblbuttons%(30), FileEdit As Boolean
Public gotojump As Boolean, coordmode2%, mapxdif, mapydif, kmxcd, kmycd, newRootNum As Integer
Public scrolling2 As Boolean, noheights As Boolean, dragbox As Boolean, MapFormatVis As Boolean
Public drag1x As Single, drag1y As Single, drag2x As Single, drag2y As Single
Public drag3x As Single, drag3y As Single, magx As Single, magy As Single, mag As Single
Public dragbegin As Boolean, drawbox As Boolean, magclose As Boolean, Delay%
Public magbox As Boolean, ht1, wt1, placdblclk As Boolean ', xrel, yrel
Public obsfile$, obsnum%, obstflag As Boolean, printing As Boolean, obs() As Single ',obsfilnum%
Public world As Boolean, worldCD%(28), worldfil$, worldfnum%, kmx400origin, kmy400origin
Public NROWS%, NCOLS%, Xworld As Single, Yworld As Single ', xdim As Double, ydim As Double
Public hgtworld As Single, cirworld As Boolean, lplac%, Xcoord, Ycoord, gotobutton As Boolean
Public lon, lat, lono, lato, worldxorigin As Double, worldyorigin As Double, kmxobs, kmyobs, jumpworld As Boolean
Public printeroffset, skyx, skyy, skymove As Boolean, kmxsky, kmysky, killpicture As Boolean
Public travelmode As Boolean, travel() As Single, travelnum%, routeload As Boolean, lonobs, latobs
Public routenum%, routnum%, routeX, routeY, terranam$, speed, speedmodify As Boolean
Public showroute As Boolean, newblit As Boolean, worldmove As Boolean ', openfile$, openfilnum%
Public exit1 As Boolean, exit2 As Boolean, exit3 As Boolean, resourcenum%, SR
Public reboot As Boolean, resizes As Boolean, mapcapold$, apprn As Single ', init As Boolean
Public accept As Boolean, abortDTM As Boolean, graphwind As Boolean, appendtravel As Boolean
Public maxang%, fullrange%, diflat%, diflog%, maxangf%, viewmode%, fullrangef%, diflatf%, diflogf%, AziStep%, AziStepf%
Public maxangs%, fullranges%, diflats%, diflogs%, maxangfs%, viewmodes%, fullrangefs%, diflatfs%, diflogfs%
Public viewmodef%, modeval, modevalf, viewer3D As Boolean, checkdtm As Boolean, modevals, modevalfs
Public xmin As Single, xmax As Single, ymin As Single, ymax As Single, insiderouteform As Boolean
Public xmino As Single, xmaxo As Single, ymino As Single, ymaxo As Single, appendfile$
Public dojump As Boolean, Ccontinue As Boolean, C10 As String, C20 As String
Public israeldtm As String * 1, israeldtmcd As Boolean, topotype%, AutoNum&, IntOld2%, cal1%
Public israeldtmf As String * 1, israeldtmcdf As Boolean, mapSearchVis As Boolean
Public worlddtm As String * 1, worlddtmcd As Boolean, ExplorerDir As String, plotfile$
Public worlddtmf As String * 1, worlddtmcdf As Boolean, searchhgt As Single, treehgtStored As Double
Public srtmdtm As String * 1, srtmdtmcd As Boolean, srtmdtmf As String * 1, srtmdtmcdf As Boolean
Public ramdrive As String * 1, ramdrivef As String * 1, viewsearch As Boolean
Public terradir$, terradirf$, impcenter As Boolean, resetorigin As Boolean, ObstructionCheck As Boolean
Public mapimport As Boolean, mapfile$, deglat As Double, deglog As Double, AutoVer As Boolean
Public blank$, woxorigin As Double, woyorigin As Double, fudx, fudy, AutoProf As Boolean
Public adx1, bdy1, adx1f, bdy1f, drivjk$, drivjk_c$, drivfordtm$, drivprom$, drivprof$, drivcities$, drivdtm$
Public crosssectionpnt(1, 1), crosssectionhgt(1), crosssection As Boolean, sectnumpnt&
Public greatcircle As Boolean, ObsHeight As Boolean, GoCrossSection As Boolean, RdHalYes As Boolean
Public ggpscorrection As Boolean, FileView As Boolean, FileViewName As String, FileViewError As Boolean
Public FileViewFileName() As String, AbrevDir$, AutoScanlist As Boolean, MapOn As Boolean
Public FileViewDir$, FileViewFileType() As Integer, UniqueRoots%, fileo$, setflag%, coordAnalyze(2) As Double
Public OutFile$, bAirPath As Boolean, kmxTrig, kmyTrig, hgtTrig, DTMflag As Integer, GoogleMapVis As Boolean
Public SearchVis As Boolean, WinVer As Long, XDIM As Double, YDIM As Double, noVoidflag As Integer
Public AutoPress As Boolean, IsraelDTMsource%, OnlyExtractFile As Boolean, EastOnly As Boolean, WestOnly As Boolean
Public CalculateProfile As Integer, SearchCrossSection As Boolean, SearchCrossObstruct As Boolean
Public rderos2_use As Boolean, IgnoreTiles%, autoazirange%, NoCDWarning As Boolean, TemperatureModel%
Public MainDir$, Turbo2cdDir$, USADir$, GEOTOPO30Dir$, D3ASDir$, SamplesDir$, D3dExplorerDir$, ErosCitiesDir$

'----------------GPS global constants------------------------
Public Const MAX_PORT = 15 'maximum number of com ports to search
Public Const MAX_GPS_WAIT = 18 'maximum number of timer intervals to wait for GPS satellite signal
Public Const MAX_WAIT_NUMBER = 300 'maximum number of GPS timer intervels to test for zero velocity
Public Const MAX_REPEAT_DISTANCE_TEST = 5 'maximum number of cycles to test spurious coordinates
Public Const MAX_NO_SIGNAL = 300 'maximum number of cycles to allow before showing "no signal" message box
Public GPSconnected As Boolean
Public ComPort% 'com port that gps is connected to
Public GPS_timer_trials As Integer
Public GPS_signal_lost As Boolean
Public Const GPS_Dist_Resolution = 5#  'km
Public GPSConnectString As String
Public GPSConnectString0 As String
Public GPSenabled As Boolean 'flag to determine if gps connection was ever established
Public GPSSetupVis As Boolean
'Public distkmTraveld As Double 'elapsed distance log
Public numTestGPSpnts% 'gps test points
Public GPSNow As Date
Public GPS_no_message As Boolean
Public GPS_latitude As Double
Public GPS_longitude As Double
Public GPS_speed As Double
Public GPS_date As String
Public GPS_time As String
Public GPS_altitude As Double
Public GPS_bearing As Double
Public GPS_ModeReceived As String
Public GPS_altitudeunits As String
Public GPS_warning_number As Long
Public GPS_no_signal_number As Long
Public GPS_test_loaded As Boolean
Public waitimeGPS As Long
Public waitimeGPSvis As Long

'browse for directory
Private Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal hMem As Long)

Private Declare Function lstrcat Lib "Kernel32" _
   Alias "lstrcatA" (ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long
   
Private Declare Function SHBrowseForFolder Lib "shell32" _
   (lpBI As BrowseInfo) As Long
   
Private Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long
   
'variables that define the imported map picture file
Private Type MapInfos
   name As String 'root name of the map picture file (i.e., without the path and extension)
   type As Integer  'bmp = 0, gif = 1, jpg = 2
   xsize As Integer 'pixel x size of map
   ysize As Integer 'pixel y size of map
   pixcx As Integer 'pixel value of cosen center in x direction (chosen center is usually an intersection of lines of latitude and longitude)
   pixcy As Integer 'pixel value of chosen center in y direction
   pixkm As Double 'pixels per kilmoters
   pixlon As Integer 'pixels for one degree of longitude
   pixlat As Integer 'pixels for one degree of latitude
   loncenter As Double 'longitude of the chosen center of the picture (e.g., the intersection of lines of lat and lon.)
   latcenter As Double 'latitude of the chosen center of the picture
End Type

Public MapInfo As MapInfos
   

Public Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String

    On Error GoTo ehBrowseForFolder 'Trap for errors

    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo

    'Set API properties (housed in a UDT)
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lstrcat(strPrompt, sEmpty)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Display the browse folder...
    lngIDList = SHBrowseForFolder(udtBI)

    If lngIDList <> 0 Then
        'Create string of nulls so it will fill in with the path
        strPath = String(MAX_PATH, 0)

        'Retrieves the path selected, places in the null
         'character filled string
        lngResult = SHGetPathFromIDList(lngIDList, strPath)

        'Frees memory
        Call CoTaskMemFree(lngIDList)

        'Find the first instance of a null character,
         'so we can get just the path
        intNull = InStr(strPath, vbNullChar)
        'Greater than 0 means the path exists...
        If intNull > 0 Then
            'Set the value
            strPath = Left(strPath, intNull - 1)
        End If
    End If

    'Return the path name
    BrowseForFolder = strPath
    Exit Function 'Abort

ehBrowseForFolder:

    'Return no value
    BrowseForFolder = Empty

End Function


Public Sub heights(kmx, kmy, hgt2)
      On Error GoTo g35
      
      If IsraelDTMsource% = 1 Then 'convert to long,lat and use SRTM extraction
         'Call casgeo(kmx, kmy, lgh, lth)

         If ggpscorrection = True Then 'apply conversion from Clark geoid to WGS84
            Dim N As Long
            Dim E As Long
            Dim lat As Double
            Dim lon As Double
            N = kmy
            E = kmx
            Call ics2wgs84(N, E, lat, lon)
            lgh = lon
            lth = lat
            'Call casgeo(kmx, kmy, lgh, lth)
            ggpscorrection = False
         Else
            Call casgeo(kmx, kmy, lgh, lth)
            End If
         
         Call worldheights(lgh, lth, hgt2)
         GoTo g99
         End If
      
      kmx = kmx * 0.001
      kmy = (kmy - 1000000) * 0.001
      IKMX& = Int((kmx + 20!) * 40!) + 1
      IKMY& = Int((380! - kmy) * 40!) + 1
      NROW% = IKMY&: NCOL% = IKMX&

'       GETZ FINDS THE HEIGHT OF A POINT AT THE NORW AND NCOL FROM 380N
'       AND -20E WHERE 1,1 IS THAT CORNER POINT
'       FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
g15:    Jg% = 1 + Int((NROW% - 2) / 800)
        Ig% = 1 + Int((NCOL% - 2) / 800)
        CHMNE = CHMAP(Ig%, Jg%)
        If CHMNE = "  " Then GoTo g35
        If CHMNE = CHMNEO Then GoTo g21
        Close #filnumg%
        SF = CHMNE
        Sffnam$ = israeldtm + ":\dtm\" + SF
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
g35:    hgt = -9999 'MsgBox " ERROR IN GETZ ", vbCritical + vbOKOnly, "SkyLight"
        Close #filnumg%
        CHMNEO = "  "
g99:

End Sub
Public Sub casgeo(kmx, kmy, lg, lt)
'converts ITM to geo using Clark geoid
        G1# = kmy - 1000000
        G2# = kmx
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
        Y1# = 1126867.91
        Y2# = G1#
'       GN & GE
        X2# = G2#
        If (X2# > 700000#) Then GoTo ca5
        X1# = X1# - 1000000#
ca5:    If (Y2# > 550000#) Then GoTo ca10
        Y1# = Y1# - 1000000#
ca10:   X1# = X2# - X1#
        Y1# = Y2# - Y1#
        D1# = Y1# * B2# / 2#
        O1# = S2# + D1#
        O2# = O1# + D1#
        A3# = O1# / f1#
        A4# = O2# / f1#
        B3# = 1# - e4# * Sin(A3#) ^ 2#
        B4# = B3# * Sqr(B3#) * C1#
        C4# = 1# - e4# * Sin(A4#) ^ 2#
        C5# = Tan(A4#) * C2# * C4# ^ 2#
        C6# = C5# * X1# ^ 2#
        D2# = Y1# * B4# - C6#
        C6# = C6# / 3#
'LAT
        l1# = (S2# + D2#) / f1#
        R3# = O2# - C6#
        R4# = R3# - C6#
        R2# = R4# / f1#
        A2# = 1# - e4# * Sin(l1#) ^ 2#
        lt = r# * (l1#)
        A5# = Sqr(A2#) * c3#
        d3# = X1# * A5# / Cos(R2#)
' LON
        lg = r# * ((s1# + d3#) / f1#)
'       THIS IS THE EASTERN HEMISPHERE!
        lg = -lg
        If ggpscorrection = True Then
           'Use the approximate correction factor
           'in order to agree with GPS.
           lg = lg - 0.0013
           lt = lt + 0.0013
           End If

End Sub

Public Sub GEOUTM(L11, L22, Z%, G1, G2)
'      INTRINSIC SIN, COS, SQR
       Dim a As Double, A0 As Double, A1 As Double, A2 As Double
       Dim A3 As Double, A4 As Double, A5 As Double, l2 As Double
       Dim b As Double, B0 As Double, B1 As Double, c As Double
       Dim C1 As Double, D As Double, E As Double, l1 As Double
       Dim L0 As Double, p As Double, r As Double, s1 As Double
       Dim T1 As Double, T2 As Double, T3 As Double, X1 As Double

      l1 = L11
      l2 = -L22
      a = 6375836.645
      b = 6354369.181
      p = 0.00672267019391
      E = 0.00676817037114
      A0 = 1.00507398896
      B0 = 0.00508468605159
      c = 0.000010718137317
      D = 2.10868032448E-08
      E0 = 3.99957158294E-11
      f = 7.25992839731E-14
      r = 57.2957795131
'    UTM Zones begin at the International Date Line (180 West and proceed
'     east in 6 increments.
'     The Central Meridian = -180+(Z%-1)*6+3
      Z% = Int((l2 + 180#) / 6#) + 1
5     L0 = l2 - (-180 + (Z% - 1) * 6 + 3)
      s1 = Sin(l1 / r)
      C1 = Cos(l1 / r)
      T1 = s1 / C1
      T2 = E * C1 ^ 2
      B1 = A0 * (l1 / r) - B0 / 2# * Sin(2# * l1 / r) + c / 4# * Sin(4# * l1 / r) - D / 6# * Sin(6# * l1 _
      / r) + E0 / 8# * Sin(8# * l1 / r) - f / 10# * Sin(10# * l1 / r)
      B1 = B1 * a * (1# - p)
      X1 = a / Sqr(1# - p * s1 ^ 2)
      A1 = 1# / r * X1 * C1
      T3 = T1 ^ 2
      A2 = 1# / 2# * X1 * C1 ^ 2 * T1 / r ^ 2
      A3 = 1# / 6# * X1 * C1 ^ 3 * (1# - T3 + T2) / r ^ 3
      A4 = 1# / 24# * X1 * C1 ^ 4 * T1 * (5# - T3 + 9# * T2 + 4 * T2 ^ 2) / r ^ 4
      A5 = 1# / 120# * X1 * C1 ^ 5 * (5# - 18# * T3 + T1 ^ 4 + 14# * T2 - 58# * T3 * T2) / r ^ 5
      G1 = B1 + A2 * L0 ^ 2 + A4 * L0 ^ 4
      G2 = 500000# + A1 * L0 + A3 * L0 ^ 3 + A5 * L0 ^ 5
End Sub


Public Sub GEOCASC(L11, L22, G11, G22)
      'convert from geo, (lg,lt) = (l11,l22) to itm (kmx,kmy) = (G11,G22)
      Dim D1 As Double, D2 As Double, D5 As Double, E3 As Double
      Dim G1 As Double, G2 As Double, G3 As Double, D4 As Double
      Dim l1 As Double, l2 As Double, l3 As Double, l4 As Double
      Dim s1 As Double, S2 As Double, AL As Double
      Dim M1 As Integer, lResult As Long, GpsCorrOff As Boolean
      If ggpscorrection Then
         'Use the approximate correction factor
         'in order to agree with GPS.
         L22 = L22 - 0.0013
         L11 = L11 - 0.0013
         GpsCorrOff = True
         ggpscorrection = False
         End If
      l1 = L11
      l2 = L22
      G1 = G11
      G2 = G22
      M1 = 0
      G1 = 100000#
      G2 = 100000#
      s1 = 31.4896370431
      S2 = 34.4727144086
      E3 = 0.0001
      l3 = l1
      l4 = l2
5     AL = (s1 + l3) / (2# * 57.2957795131)
      G3 = FNM(AL)
      G4 = FNP(AL)
      D1 = (l3 - s1) * G3
      D2 = (l4 - S2) * G4
      G1 = G1 + D1
      G2 = G2 + D2
      M1 = M1 + 1
      D5 = Sqr(D1 ^ 2 + D2 ^ 2)
      If (D5 < E3) Then GoTo 10
      If (M1 > 10) Then GoTo 15
      G111# = G1 + 1000000# '<<<<<
      Call casgeo(G2, G111#, l2, l1) '<<<<
      l2 = -l2 '<<<<
      s1 = l1
      S2 = l2
      GoTo 5
10    l1 = l3
      l2 = l4
      If (G1 < 0) Then G1 = G1 + 1000000#
      G11 = G1
      G22 = G2
      If GpsCorrOff Then ggpscorrection = True
      Exit Sub
15    If world = False Then
         lResult = FindWindow(vbNullString, terranam$)
         If lResult > 0 And terranam$ <> "" Then
'            ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            BringWindowToTop (lResult)
            End If
         End If
      Screen.MousePointer = vbDefault
      response = MsgBox("Routine GEOCASC failed to Converge", vbCritical + vbOKOnly, "Maps&More")
      If world = False And lResult > 0 And terranam$ <> "" Then
'         ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (lResult)
         End If
      If GpsCorrOff Then ggpscorrection = True
End Sub


Public Sub UTMGEO(G11, G22, Z11, L11, L22)
      Dim a As Double, A0 As Double, b As Double, B1 As Double, B2 As Double
      Dim B3 As Double, B4 As Double, B5 As Double, c As Double, C1 As Double
      Dim D As Double, D1 As Double, E As Double, E0 As Double, EP As Double
      Dim f As Double, G0 As Double, G1 As Double, G2 As Double, l1 As Double
      Dim l2 As Double, p As Double, r As Double, s1 As Double, T1 As Double
      Dim T2 As Double, T3 As Double, X1 As Double
      Dim Z As Integer, M As Integer
      Z = Z11
      G1 = G11
      G2 = G22
      a = 6375836.645
      b = 6354369.181
      p = 0.00672267019391
      E = 0.00676817037114
      A0 = 1.00507398896
      B0 = 0.00508468605159
      c = 0.000010718137317
      D = 2.10868032448E-08
      E0 = 3.99957158294E-11
      f = 7.25992839731E-14
      r = 57.2957795131
      EP = 0.001
      G0 = G2 - 500000#
      M = 0
      l1 = (2# * G1) / (a + b)
'    ROUTINE POLY
5     s1 = Sin(l1)
      C1 = Cos(l1)
      T1 = s1 / C1
      T2 = E * C1 ^ 2
      B1 = A0 * (l1) - B0 / 2# * Sin(2# * l1) + c / 4# * Sin(4# * l1) - D / 6# * Sin(6# * l1) + _
           E0 / 8# * Sin(8# * l1) - f / 10# * Sin(10# * l1)
      B1 = B1 * a * (1# - p)
      X1 = a / Sqr(1# - p * s1 ^ 2)
'     END ROUTINE POLY
      D1 = G1 - B1
      If (M = 0) Then GoTo 10
      If (Abs(D1) < EP) Then GoTo 15
10    l1 = l1 + (2# * D1 / (a + b))
      M = M + 1
      GoTo 5
15    B1 = 1# / (X1 * C1)
      T3 = T1 ^ 2
      B2 = -(T1 * (1# + T2) / (2# * X1 ^ 2))
      B3 = -(1# / (6# * X1 ^ 3 * C1) * (1# + 2 * T3 + T2))
      B4 = T1 / (24# * X1 ^ 4) * (5# - 3# * T3 + 7# * T2)
      B5 = 1# / (120# * X1 ^ 5 * C1) * (5# + 28# * T3 + 24# * T1 ^ 4)
      l1 = r * (l1 + (B2 * G0 ^ 2 + B4 * G0 ^ 4))
      l2 = r * (B1 * G0 + B3 * G0 ^ 3 + B5 * G0 ^ 5) + Z
      L11 = l1
      L22 = l2
End Sub

Public Function FNP(p As Double) As Double
'  Length of a degree parallel: P must be latitude in radians
   'Dim P As Double
   FNP = 111415.13 * Cos(p) - 94.55 * Cos(3# * p) + 0.012 * Cos(5# * p)
End Function

Public Function FNM(p As Double) As Double
'  Length of a degree meridian: P must be latitude in radians
   'Dim P As Double
   FNM = 111132.09 - 566.05 * Cos(2# * p) + 1.2 * Cos(4# * p) - 0.002 * Cos(6# * p)
End Function
Public Sub ITMSKY(kmxn, kmyn, T1, T2, Mode%)
   If Mode% = 1 Then 'ITM to SKY
       T1 = kmxn
       T2 = kmyn - 1000000
       'T1 = Fix(0.5 + kmxn + 49988.91922 + 5.081)
       'T2 = Fix(0.5 + kmyn - 500630.57728236 - 57.423 + 9.5)
   ElseIf Mode% = 2 Then 'SKY to ITM
       kmxn = T1
       kmyn = T2 + 1000000
       'kmxn = Fix(0.5 + T1 - 49988.91922 - 5.081)
       'kmyn = Fix(0.5 + T2 + 500630.5772826 + 57.423 - 9.5)
      End If
End Sub
Public Function EnumFunc(ByVal hWndChild As Long, ByVal lParam As Long) As Boolean
   Dim size As Long
   Dim s As String
   s = String(255, 0)
   'size = DefWindowProc(hWndChild, WM_GETTEXT, 256, s1&)
   size = GetWindowText(hWndChild, s, Len(s))
   's = CStr(s1&)
   s = Left$(s, size)
   Skynum% = Skynum% + 1
   If Skynum% = 44 Then
      hChild = hWndChild
      EnumFunc = False 'signal callback routine to stop searching
      Exit Function
      End If
   'If Skycoord% = 0 Then nhwnd = hWndChild
   'winx = 0
   'winy = 50
   'winw = 515 '515
   'winh = 1000 '475
   'winp = True
   'Skycoord% = Skycoord% + 1
   ' ret = MoveWindow(nhwnd, winx, winy, winw, winh, winp)

   ''*************experimental*****************
   'If lParam = 1 Then
   '   If Skycoord% = 1 And Val(s) = 0 Then
   '      ret = SetWindowText(hWndChild, C1p)
   '      Skycoord% = 2
   '   ElseIf Skycoord% = 2 And Val(s) = 0 Then
   '      ret = SetWindowText(hWndChild, C2p)
   '      EnumFunc = False
   '      Exit Function
   '      End If
   '   End If
   'If routeload = True Then
   '   sext$ = sEmpty
   '   If Len(s) > 4 Then sext$ = Mid$(s, Len(s) - 2, 3)
   '   If sext$ = "trf" Then
   '      If s = routename$ Then
   '         EnumFunc = False 'signal callback routine to stop searching
   '         Exit Function
   '      Else 'keep on looking for .trf files
   '         lOpen = FindWindow(vbNullString, "Open")
   '         Call BringWindowToTop(lOpen)
   '         Call keybd_event(VK_DOWN, 0, 0, 0)
   '         Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
   '         End If
   '      End If
   'Else

    '******this sections is no longer valid so skip it*****
    GoTo e900
    '*******************************
      If Val(s) > "10000" And lParam = 0 Then
         If Skycoord% = 1 Then
            skyx = Val(s)
            'Maps.Text5.Text = s
            Skycoord% = 2
         ElseIf Skycoord% = 2 And lParam = 0 Then
            skyy = Val(s)
            'Maps.Text6.Text = s
            EnumFunc = False 'signal callback routine to stop searching
            Exit Function
            End If
         End If
   '   End If
e900: EnumFunc = True 'continue searching
End Function
Public Function EnumFunc2(ByVal hWndChild As Long, ByVal lParam As Long) As Boolean
   Dim size As Long
   Dim s As String
   s = String(255, 0)
   size = GetWindowText(hWndChild, s, Len(s))
   s = Left$(s, size)
   posit% = InStr(s, "%")
   If posit% <> 0 Then
      Source = Val(Mid$(s, posit% - 2, 2))
      If Source <= 15 Then
        Maps.StatusBar1.Font.size = 8
        Maps.StatusBar1.Font.Bold = True
        Maps.StatusBar1.Panels(3) = "Warning LOW Sys.Res."
        Maps.StatusBar1.Panels(2) = "You have used up the system resources--you'll need to reboot!"
        'close the terraviewer and shut down the video timers
        Maps.Timer1.Enabled = False
        Maps.Timer2.Enabled = False
        lResult = FindWindow(vbNullString, terranam$)
        If lResult > 0 Then
           For i% = 18 To 19
              Maps.Toolbar1.Buttons(i%).Enabled = False
              Maps.Toolbar1.Buttons(i%).value = tbrUnpressed
              tblbuttons(i%) = 0
           Next i%
           Maps.Toolbar1.Buttons(23).value = tbrUnpressed
           Maps.Toolbar1.Buttons(23).Enabled = False
           tblbuttons(23) = 0
           Maps.Toolbar1.Buttons(24).value = tbrUnpressed
           Maps.Toolbar1.Buttons(24).Enabled = False
           tblbuttons(24) = 0
           Maps.Toolbar1.Buttons(25).value = tbrUnpressed
           Maps.Toolbar1.Buttons(25).Enabled = False
           tblbuttons(25) = 0
           showroute = False
           Maps.Timer2.Enabled = False
           skyleftjump = False
           skymove = False
           'close the terraviewer
           ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
           Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
           Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
           'waitime = Timer
           'Do Until Timer > waitime + 0.01
           '   DoEvents
           'Loop
           Call keybd_event(Asc("F"), 0, 0, 0)  'goes into Settings menu
           Call keybd_event(Asc("F"), 0, KEYEVENTF_KEYUP, 0)
           Call keybd_event(Asc("E"), 0, 0, 0)  'goes into Settings menu
           Call keybd_event(Asc("E"), 0, KEYEVENTF_KEYUP, 0)
           Maps.Toolbar1.Buttons(17).value = tbrUnpressed
           Maps.Loadfm.Enabled = False
           Maps.recoverroutefm.Enabled = False
           mapPictureform.Visible = False
           For i% = 6 To 15
              Maps.Toolbar1.Buttons(i%).value = tbrUnpressed
              tblbuttons(i%) = 0
           Next i%
           response = MsgBox("System resources are almost exhausted!. Do you wan't to restart the computer now?", vbQuestion + vbYesNo, "Maps & More")
           If response = vbYes Then reboot = True
           End If
         EnumFunc2 = False
         Exit Function
      Else
         If Source <= 30 Then
           Maps.Timer3.Interval = 10000
         ElseIf Source <= 25 Then
           Maps.Timer3.Interval = 5000
         ElseIf Source <= 20 Then
           Maps.Timer3.Interval = 1000
           End If
         resourcenum% = resourcenum% + 1
         SR = SR + Source / 2
         If resourcenum% = 2 Then
            Maps.StatusBar1.Panels(3) = "Average SysRes: " + LTrim$(RTrim$(Str(Format(SR, "#0.0#")))) + "%"
            resourcenum% = 0
            SR = 0
            EnumFunc2 = False
            Exit Function
            End If
         End If
      End If
   EnumFunc2 = True 'continue searching
End Function
Function EnumWndProc(ByVal hwnd As Long, lParam As Long) As Long    ' Increment count    lParam = lParam + 1    ' Get window title and insert into ListBox    Dim s As String    s = WindowTextFromWnd(hWnd)    If s <> sEmpty Then        lstEnumRef.AddItem s
    ' Get 3D Explorer window title
    Dim s As String
    s = String(255, 0)
    size = GetWindowText(hwnd, s, Len(s))
    s = Left$(s, size)
    If Mid$(s, 1, 6) = "3DXUSA" Then
       Tdxname = s
       TdxhWnd = hwnd
       EnumWndProc = False 'Return False to stop enumerating
       Exit Function
    Else
       End If    ' Return True to keep enumerating
    EnumWndProc = True
End Function

Public Sub worldheights(lg, lt, hgt)
   Dim leros As Long, lmag As Long
   On Error GoTo worlderror
   
   If lt > 90 Or lt < -90 Or lg < -180 Or lg > 180 Then Exit Sub
   
      
   'check if have correct CD in the drive, if not present error message
   '//changes 061222 - added northern range to N70 for Alaska DEM and EU-dem files
   If (world = False And IsraelDTMsource% = 1) Or (DTMflag > 0 And (lt >= -60 And lt <= 70)) Then 'SRTM
      
      If world = False And IsraelDTMsource% = 1 Then
         'use 90-m SRTM of Eretz Yisroel
         XDIM = 8.33333333333333E-04
         YDIM = 8.33333333333333E-04
         lg = -lg
         DEMfile$ = israeldtm + ":\dtm\"
         NROWS = 1201
         NCOLS = 1201
         GoTo wh50
         End If
         
      If DTMflag = 1 And Dir(srtmdtm & ":/USA/", vbDirectory) <> sEmpty Then
         XDIM = 8.33333333333333E-04 / 3#
         YDIM = 8.33333333333333E-04 / 3#
         DEMfile$ = srtmdtm & ":/USA/"
         NROWS = 3601
         NCOLS = 3601
      ElseIf DTMflag = 2 And Dir(srtmdtm & ":/3AS/", vbDirectory) <> sEmpty Then
         XDIM = 8.33333333333333E-04
         YDIM = 8.33333333333333E-04
         DEMfile$ = srtmdtm & ":/3AS/"
         NROWS = 1201
         NCOLS = 1201
         End If
wh50:
      'determine tile name
      lg1 = Int(lg)
      If lg1 < 0 And lg1 > lg Then lg1 = lg1 - 1
      If lg1 < 0 Then EWch$ = "W" Else EWch$ = "E"
      If Abs(lg1) < 10 Then
         lg1ch$ = "00" & Trim$(Str$(Abs(lg1)))
      ElseIf Abs(lg1) >= 10 And Abs(lg1) < 100 Then
         lg1ch$ = "0" & Trim$(Str$(Abs(lg1)))
      ElseIf Abs(lg1) >= 100 Then
         lg1ch$ = Trim$(Str$(Abs(lg1)))
         End If
      lt1 = Int(lt) 'SRTM tiles are named by SW corner
      If lt1 < 0 And lt1 > lt Then lt1 = lt1 - 1
      If lt1 < 0 Then NSch$ = "S" Else NSch$ = "N"
      If Abs(lt1) < 10 Then
         lt1ch$ = "0" & Trim$(Str$(Abs(lt1)))
      ElseIf Abs(lt1) >= 10 Then
         lt1ch$ = Trim$(Str$(Abs(lt1)))
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
         If mapEROSDTMwarn.Visible = True Then
            Unload mapEROSDTMwarn
            Set mapEROSDTMwarn = Nothing
            If magbox = True Then
               lmag = FindWindow(vbNullString, mapMAGfm.Caption)
               If lmag > 0 Then
                  ret = BringWindowToTop(lmag) 'bring mapMAGfm back to top of Z order
                  End If
               End If
            End If
      
         worldfnum% = FreeFile
         Open DEMfile$ For Binary As #worldfnum%
         GoSub Eroshgt
         Close #worldfnum%
         worldfnum% = 0
         hgt = integ2%
         If hgt = -32768 Then hgt = 0 'void
         Exit Sub
         End If
       End If

gtopo:

   XDIM = 8.33333333333333E-03
   YDIM = 8.33333333333333E-03
   If lt > -60 Then
      nx% = Fix((lg + 180) * 0.025)
      lg1 = -180 + nx% * 40
      If Abs(lg1) >= 100 Then
         lg1ch$ = RTrim$(LTrim$(Str$(Abs(lg1))))
      Else
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
         End If
      If lg1 < 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ny% = Int((90 - lt) * 0.02)
      lt1 = 90 - 50 * ny%
      lt1ch$ = LTrim$(RTrim$(Str$(Abs(lt1))))
      If lt1 > 0 Then
         ns$ = "N"
      Else
         ns$ = "S"
         End If
      DEMfile0$ = EW$ + lg1ch$ + ns$ + lt1ch$
      DEMfile1$ = worlddtm + ":\" + DEMfile0$ + "\" + DEMfile0$
      DEMfile$ = DEMfile1$ + ".dem"
      NROWS = 6000
      NCOLS = 4800
      numCD% = worldCD%(ny% * 9 + nx% + 1)
   Else 'Antartic - Cd #5
      nx% = Fix((lg + 180) / 60)
      lg1 = -180 + nx% * 60
      If Abs(lg1) >= 100 Then
         lg1ch$ = LTrim$(RTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) < 100 And Abs(lg1) <> 0 Then
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) = 0 Then
         lg1ch$ = "000"
         End If
      If lg1 <= 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ns$ = "S"
      lt1 = -60
      lt1ch$ = "60"
      DEMfile0$ = EW$ + lg1ch$ + ns$ + lt1ch$
      DEMfile1$ = worlddtm + ":\" + DEMfile0$ + "\" + DEMfile0$
      DEMfile$ = DEMfile1$ + ".dem"
      NROWS = 3600
      NCOLS = 7200
      numCD% = 5
      End If
   If worldfil$ <> DEMfile1$ Then
      myfile = Dir(DEMfile$)
      If myfile = sEmpty Then
         mapEROSDTMwarn.Visible = True
'         ret = SetWindowPos(mapEROSDTMwarn.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (mapEROSDTMwarn.hwnd)
         mapEROSDTMwarn.Label3.Caption = numCD%
         leros = FindWindow(vbNullString, "         USGS EROS DEM CD not found!")
         If leros > 0 Then
            ret = BringWindowToTop(leros) 'bring message to top
            End If
      Else
         If mapEROSDTMwarn.Visible = True Then
            Unload mapEROSDTMwarn
            Set skyerosdtwarn = Nothing
            If magbox = True Then
               lmag = FindWindow(vbNullString, mapMAGfm.Caption)
               If lmag > 0 Then
                  ret = BringWindowToTop(lmag) 'bring mapMAGfm back to top of Z order
                  'ret = ShowWindow(lmag, SW_RESTORE) 'redisplay mapMAGfm
                  End If
               End If
            End If
         If worldfnum% <> 0 Then Close #worldfnum%
         '******set as constants
         'worldfnum% = FreeFile
         'worldfil$ = DEMfile1$
         'Open DEMfile1$ + ".STX" For Input As #worldfnum%
         'Input #worldfnum%, A, elevmin%, elevmax%, D, E
         'Close #worldfnum%
         'Open DEMfile1$ + ".HDR" For Input As #worldfnum%
         'npos% = 0
         'Do Until EOF(worldfnum%)
         '  npos% = npos% + 1
         '  Line Input #worldfnum%, doclin$
         '  If npos% = 3 Then
         '     nrows% = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
         '  ElseIf npos% = 4 Then
         '     ncols% = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
         '  ElseIf npos% = 13 Then
         '     xdim = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
         '  ElseIf npos% = 14 Then
         '     ydim = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
         '     End If
        'Loop
        'Close #worldfnum%
        worldfnum% = FreeFile
        Open DEMfile$ For Binary As #worldfnum%
        GoSub Eroshgt
        Close #worldfnum%
        worldfnum% = 0
        hgt = integ2%
        End If
    Else
       If mapEROSDTMwarn.Visible = True Then
          Unload mapEROSDTMwarn
          Set skyerosdtwarn = Nothing
          End If
       If magbox = True Then
          lmag = FindWindow(vbNullString, mapMAGfm.Caption)
          If lmag > 0 Then
             ret = BringWindowToTop(lmag) 'bring mapMAGfm back to top of Z order
'             ret = ShowWindow(lmag, SW_RESTORE) 'redisplay mapMAGfm
             End If
          End If
       'continue reading
        GoSub Eroshgt
        hgt = integ2%
        End If
    Exit Sub

Eroshgt:
'   IKMY& = CInt(((lt1 - ydim * 0.5) - lt) / ydim) + 1
'   IKMX& = CInt((lg - (lg1 + xdim * 0.5)) / xdim) + 1
   IKMY& = CLng((lt1 - lt) / YDIM) + 1
   IKMX& = CLng((lg - lg1) / XDIM) + 1
   tncols = NCOLS%
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
       aa$ = sEmpty
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
          aa$ = A2$ + A1$
       ElseIf Len(A0$) = 3 Then
          A1$ = "0" + Mid$(A0$, 1, 1)
          A2$ = Mid$(A0$, 2, 2)
          If Mid$(A0$, 2, 1) = "0" Then A2$ = Mid$(A0$, 3, 1)
          aa$ = A2$ + A1$
       ElseIf Len(A0$) = 2 Or Len(A0$) = 1 Then
          A1$ = "00"
          A2$ = A0$
          aa$ = A2$ + A1$
          End If
    
        'convert swaped hexadecimel to an integer value
        leng% = Len(LTrim$(RTrim$(aa$)))
        integ1& = 0
        For j% = leng% To 1 Step -1
            v$ = Mid$(LTrim$(RTrim$(aa$)), j%, 1)
            If InStr("ABCDEF", v$) <> 0 Then
               If v$ = "A" Then
                  NO& = 10
               ElseIf v$ = "B" Then
                  NO& = 11
               ElseIf v$ = "C" Then
                  NO& = 12
               ElseIf v$ = "D" Then
                  NO& = 13
               ElseIf v$ = "E" Then
                  NO& = 14
               ElseIf v$ = "F" Then
                  NO& = 15
                  End If
            Else
               NO& = Val(v$)
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
   ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   response = MsgBox("An error in reading the CD has occured! Do you wish to try again?", vbCritical + vbRetryCancel, "Maps & More")
'   ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (mapPictureform.hwnd)
   If response = vbCancel Then Exit Sub
   Resume
End Sub
Public Function DASIN(XX As Double) As Double
   If XX = 1 Then
      DASIN = 90# * cd
   ElseIf XX = -1 Then
      DASIN = 270# * cd
   Else
      DASIN = Atn(XX / Sqr(-XX * XX + 1))
      End If
End Function
Public Function DACOS(XX As Double) As Double
   If XX = 1# Then
      DACOS = 0#
   ElseIf XX = -1# Then
      DACOS = 180# * cd
   Else
      DACOS = -Atn(XX / Sqr(-XX * XX + 1)) + pi / 2
      End If
End Function
Public Function atan2(ByVal y As Double, ByVal x As Double) _
    As Double
    'keeps angle within -180 to 180
  Dim theta As Double

  If (Abs(x) < 0.0000001) Then
    If (Abs(y) < 0.0000001) Then
      theta = 0#
    ElseIf (y > 0#) Then
      theta = 1.5707963267949
    Else
      theta = -1.5707963267949
    End If
  Else
    theta = Atn(y / x)
  
    If (x < 0) Then
      If (y >= 0#) Then
        theta = 3.14159265358979 + theta
      Else
        theta = theta - 3.14159265358979
      End If
    End If
  End If
    
  atan2 = theta

End Function
Public Function datan2(ByVal y As Double, ByVal x As Double) _
    As Double
  Dim theta As Double

  If (Abs(x) < 0.0000001) Then
    If (Abs(y) < 0.0000001) Then
      theta = 0#
    ElseIf (y > 0#) Then
      theta = 1.5707963267949
    Else
      theta = -1.5707963267949
    End If
  Else
    theta = Atn(y / x)
  
    If (x < 0) Then
      If (y >= 0#) Then
        theta = 3.14159265358979 + theta
      Else
        theta = theta - 3.14159265358979
      End If
    End If
  End If
    
  datan2 = theta / cd
End Function
Public Sub dipcoord()
    Dim cosang As Double
    If world = True Then
       If noheights = False Then
         lg = lono '-180# + X * 360# / SkyLightfm.mapPictureform.mapPicture.Width
         lt = lato '90# - Y * 180# / SkyLightfm.mapPictureform.mapPicture.Height
         Call worldheights(lg, lt, hgt2)
         If hgt2 = -9999 Then hgt2 = 0
       Else
         hgt2 = 0#
         End If
       lt1 = Maps.Text6.Text
       lg1 = -Maps.Text5.Text
       hgt1 = 0: If Maps.Text7.Text <> sEmpty Then hgt1 = Maps.Text7.Text
       lg2 = -lono '(-180# + X * 360# / mapPictureform.mapPicture.Width)
       lt2 = lato '90# - Y * 180# / mapPictureform.mapPicture.Height
       X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
       X2 = Cos(lt2 * cd) * Cos(lg2 * cd)
       Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
       Y2 = Cos(lt2 * cd) * Sin(lg2 * cd)
       Z1 = Sin(lt1 * cd)
       Z2 = Sin(lt2 * cd)
       'distance is Re * Angle between vectors
       'cos(Angle between unit vectors) = Dot product of unit vectors
       cosang = X1 * X2 + Y1 * Y2 + Z1 * Z2
       distkm = 6371.315 * DACOS(cosang)
       'distkm = 6371.315 * Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2 + (Z1 - Z2) ^ 2)
    Else
       kmxcd = kmxsky
       kmycd = kmysky
'       If coordmode2% = 1 Then 'goto coord in ITM
'          kmxcd = Maps.Text5.Text
'          kmycd = Maps.Text6.Text
'       ElseIf coordmode2% = 2 Then 'goto coord in GEO
'          'cannot convert this format at this point use defaults
'          kmxcd = kmxc
'          kmycd = kmyc
'       ElseIf coordmode2% = 3 Then 'goto coord in UTM
'          G1 = Maps.Text5.Text
'          G2 = Maps.Text6.Text
'          Z = 33
'          Call UTMGEO(G1, G2, Z, L1, L2)
'          Call GEOCASC(L1, L2, kmxg, kmyg)
'          ITM1 = Fix(0.5 + kmyg)
'          If kmxg < 870000 Then
'             ITM2 = Fix(0.5 + kmxg) + 1000000
'          Else
'             ITM2 = Fix(0.5 + kmxg)
'             End If
'          kmxcd = ITM1
'          kmycd = ITM2
'       ElseIf coordmode2% = 4 Then 'goto coord in SKY
'          T1 = Maps.Text5.Text
'          T2 = Maps.Text6.Text
'          Call ITMSKY(ITM1, ITM2, T1, T2, 2)
'          kmxcd = ITM1
'          kmycd = ITM2
'          End If
       distkm = Sqr((kmxoo - kmxcd) ^ 2 + (kmyoo - kmycd) ^ 2) * 0.001
       End If
    If distkm <= 0.005 Then
       Maps.Text4.Text = "0"
       Maps.Text2.Text = "0"
       Maps.Text3.Text = hgt2
       Maps.Text1.Text = LTrim$(Format(distkm, "###,##0.000"))
       Exit Sub
       End If

    If map400 = True And world = False Then
       hgt2 = hgt
       hgt1 = 0: If Maps.Text7.Text <> sEmpty Then hgt1 = Maps.Text7.Text
       Call casgeo(kmxoo, kmyoo, lg, lt)
       lg2 = lg
       lt2 = lt
       Call casgeo(kmxcd, kmycd, lg, lt)
       lg1 = lg
       lt1 = lt
    ElseIf map50 = True And world = False Then
       hgt2 = hgt
       hgt1 = 0: If Maps.Text7.Text <> sEmpty Then hgt1 = Maps.Text7.Text
       Call casgeo(kmxoo, kmyoo, lg, lt)
       lg2 = lg
       lt2 = lt
       Call casgeo(kmxcd, kmycd, lg, lt)
       lg1 = lg
       lt1 = lt
       End If
     X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
     X2 = Cos(lt2 * cd) * Cos(lg2 * cd)
     Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
     Y2 = Cos(lt2 * cd) * Sin(lg2 * cd)
     Z1 = Sin(lt1 * cd)
     Z2 = Sin(lt2 * cd)
     Re = 6371315#
     re1 = (hgt1 + Re)
     re2 = (hgt2 + Re)
     X1 = re1 * X1
     Y1 = re1 * Y1
     Z1 = re1 * Z1
     X2 = re2 * X2
     Y2 = re2 * Y2
     Z2 = re2 * Z2
     dist1 = re1
     dist2 = re2
     Angle = DACOS((X1 * X2 + Y1 * Y2 + Z1 * Z2) / (dist1 * dist2))
     viewang = Atn((-re1 + re2 * Cos(Angle)) / (re2 * Sin(Angle)))
     D = (dist1 - dist2 * Cos(Angle)) / dist1
     x1d = X1 * (1 - D) - X2
     y1d = Y1 * (1 - D) - Y2
     z1d = Z1 * (1 - D) - Z2
     'x1p = -Y1
     'y1p = X1
     'z1p = 0
     'azicos = (x1p * x1d + y1p * y1d) / Sqr(X1 ^ 2 + Y1 ^ 2)
     x1p = -Sin(lg1 * cd)
     y1p = Cos(lg1 * cd)
     z1p = 0
     azicos = (x1p * x1d + y1p * y1d)
     x1s = -Cos(lg1 * cd) * Sin(lt1 * cd)
     y1s = -Sin(lg1 * cd) * Sin(lt1 * cd)
     z1s = Cos(lt1 * cd)
     azisin = (x1s * x1d + y1s * y1d + z1s * z1d)
     azi = Atn(azisin / azicos)
     If world = True Then distkm = Angle * Re * 0.001
     Maps.Text1.Text = LTrim$(Format(distkm, "###,##0.000"))
     Maps.Text4.Text = LTrim$(Format(viewang / cd, "##0.000"))
     Maps.Text2.Text = LTrim$(Format(azi / cd, "##0.000"))
     Maps.Text3.Text = hgt2
End Sub

Public Sub loadpictures() 'this routine loads up the PictureClip buffers

   On Error GoTo loadpictures_Error

If mapwi = mapPictureform.Width Then
   If mapwi2 > mapwi - mapwi * 0.1 Then
      mapwi2 = mapwi
   Else
      mapwi2 = defaultmapwidth '0.469 * mapwi
   End If
Else
  mapwi2 = mapPictureform.Width
  End If
If maphi = mapPictureform.Height Or printing = True Then
   If maphi2 > maphi - maphi * 0.1 Then
      maphi2 = maphi
   Else
      maphi2 = defaultmapheight '0.79 * maphi
      End If
Else
  maphi2 = mapPictureform.Height
  End If

'If Abs(mapxdif) > 100 Then
'   mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
'   End If
'If Abs(mapydif) > 100 Then
'   mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
'   End If


If world = False Then
    If mag > 1 Then
       kmxcc = kmxc
       kmycc = kmyc
    Else
       If map50 = True Then
          kmxcc = kmxc + (km50x) * (mapwi - mapwi2 + mapxdif) / 2
          kmycc = kmyc - (km50y) * (maphi - maphi2 + mapydif) / 2
       ElseIf map400 Then
          kmxcc = kmxc + (km400x) * (mapwi - mapwi2 + mapxdif) / 2
          kmycc = kmyc - (km400y) * (maphi - maphi2 + mapydif) / 2
          End If
        End If
ElseIf world = True Then
   '(first check for busy signal from egg.exe)
    If Maps.Timer2.Enabled = True Then
       myfile = Dir(ramdrive + ":\wait.x")
       If myfile <> sEmpty Then
           waitime = Timer
           Do Until Timer > waitime + 0.5
              DoEvents
           Loop
           Exit Sub
           End If
       End If

    If mapimport = False Then
       deglog = 180
       deglat = 180
       End If
    If mag > 1 Then
       lonc = lon '+ fudx / mag
       latc = lat '+ fudy / mag
    Else
       'lonc = lon + (180 / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
       'latc = lat - (180 / sizewy) * (maphi - maphi2 + mapydif) / 2
       lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx    '+ 0.166
       latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy   '+ 0.204
       End If
    End If

If map50 = True Then
   If topotype% = 0 Then 'old format 50000 scale maps
       'each 1:50 tile is 5km x 5km, and is known as (n50x%(1),n50y%(1))
       'the South-West corner of each tile has ITM coord: (kmxorigin, kmyorigin)
       'can't display more than 1 tile worth of Eretz Israel at a time
       'but this can be any combination of the 4 tiles stored in the PictureClip boxes
       'kmxc,kmyc is always in the middle of the picture
       'assuming that the mapPicturefm as width Wi, and height Hi (in twips)
       'then each tile is 8850 twips wide and 8850 twips high corresponding to
       '5000 meters by 5000 meters.
       'The nine PictureClip1s are organized as follows:
       '
       '  1  2  3                          2  3  4
       '  8  0  4   they contain pic1's:   9  1  5
       '  7  6  5                          8  7  6
       '
       'if kmxc, kmyc are the coordinates of the middle of mapPictureform, then
       'then find the map file that contains it and load it
       'into pictureclip1(0).
       
       If kmxcc = 0 Then kmxcc = 100000
       If kmycc = 0 Then kmycc = 1100000

       n50x%(1) = Int((kmxcc - 70000) / 10000) + 1
       s11 = n50x%(1) - 1
       kmxorigin = 70000 + s11 * 10000
       dxx% = Int((kmxcc - kmxorigin) / 5000) 'Mod 2
       s111 = dxx%
       kmxorigin = 70000 + s11 * 10000 + 5000 * s111
       If dxx% = 0 Then 'also find other tiles' coordinates
          Mid$(pic2(1), 2, 1) = "l"
          Mid$(pic2(2), 2, 1) = "r"
          n50x%(2) = n50x%(1) - 1
          Mid$(pic2(3), 2, 1) = "l"
          n50x%(3) = n50x%(1)
          Mid$(pic2(4), 2, 1) = "r"
          n50x%(4) = n50x%(1)
          Mid$(pic2(5), 2, 1) = "r"
          n50x%(5) = n50x%(1)
          Mid$(pic2(6), 2, 1) = "r"
          n50x%(6) = n50x%(1)
          Mid$(pic2(7), 2, 1) = "l"
          n50x%(7) = n50x%(1)
          Mid$(pic2(8), 2, 1) = "r"
          n50x%(8) = n50x%(1) - 1
          Mid$(pic2(9), 2, 1) = "r"
          n50x%(9) = n50x%(1) - 1
       Else
          Mid$(pic2(1), 2, 1) = "r"
          Mid$(pic2(2), 2, 1) = "l"
          n50x%(2) = n50x%(1)
          Mid$(pic2(3), 2, 1) = "r"
          n50x%(3) = n50x%(1)
          Mid$(pic2(4), 2, 1) = "l"
          n50x%(4) = n50x%(1) + 1
          Mid$(pic2(5), 2, 1) = "l"
          n50x%(5) = n50x%(1) + 1
          Mid$(pic2(6), 2, 1) = "l"
          n50x%(6) = n50x%(1) + 1
          Mid$(pic2(7), 2, 1) = "r"
          n50x%(7) = n50x%(1)
          Mid$(pic2(8), 2, 1) = "l"
          n50x%(8) = n50x%(1)
          Mid$(pic2(9), 2, 1) = "l"
          n50x%(9) = n50x%(1)
          End If
       n50y%(1) = Int((kmycc - 870000) / 10000) + 1
       s22 = n50y%(1) - 1
       kmyorigin = 870000 + s22 * 10000
       dyy% = Int((kmycc - kmyorigin) / 5000) 'Mod 2
       s33 = dyy%
       kmyorigin = 870000 + s22 * 10000 + 5000 * s33
       'so if kmxc,kmyc are in X=Wi/2,Y=Hi/2 then kmxorigin,kmyorigin
       'are at Wi/2-(kmxc-kmxorigin)*8850/5000,Hi/2+(kmyc-kmyorigin)*8850/5000
       'find name of first picture
       If dyy% = 0 Then
          Mid$(pic2(1), 1, 1) = "d"
          Mid$(pic2(2), 1, 1) = "u"
          n50y%(2) = n50y%(1)
          Mid$(pic2(3), 1, 1) = "u"
          n50y%(3) = n50y%(1)
          Mid$(pic2(4), 1, 1) = "u"
          n50y%(4) = n50y%(1)
          Mid$(pic2(5), 1, 1) = "d"
          n50y%(5) = n50y%(1)
          Mid$(pic2(6), 1, 1) = "u"
          n50y%(6) = n50y%(1) - 1
          Mid$(pic2(7), 1, 1) = "u"
          n50y%(7) = n50y%(1) - 1
          Mid$(pic2(8), 1, 1) = "u"
          n50y%(8) = n50y%(1) - 1
          Mid$(pic2(9), 1, 1) = "d"
          n50y%(9) = n50y%(1)
       Else
          Mid$(pic2(1), 1, 1) = "u"
          Mid$(pic2(2), 1, 1) = "d"
          n50y%(2) = n50y%(1) + 1
          Mid$(pic2(3), 1, 1) = "d"
          n50y%(3) = n50y%(1) + 1
          Mid$(pic2(4), 1, 1) = "d"
          n50y%(4) = n50y%(1) + 1
          Mid$(pic2(5), 1, 1) = "u"
          n50y%(5) = n50y%(1)
          Mid$(pic2(6), 1, 1) = "d"
          n50y%(6) = n50y%(1)
          Mid$(pic2(7), 1, 1) = "d"
          n50y%(7) = n50y%(1)
          Mid$(pic2(8), 1, 1) = "d"
          n50y%(8) = n50y%(1)
          Mid$(pic2(9), 1, 1) = "u"
          n50y%(9) = n50y%(1)
          End If

       'now load the buffers
        For i% = 1 To 9
           If n50y%(i%) <= 12 Then
              pic1$(i%) = 788 + (n50x%(i%) - 1) * 100 + (n50y%(i%) - 1)
           Else
              pic1$(i%) = 700 + (n50x%(i%) - 1) * 100 + (n50y%(i%) - 13)
              End If
           If n50x%(i%) >= 4 Then
              picf$ = Turbo2cdDir$ & "Itmv2s\MAP50\CLI" + pic1$(i%) + "." + pic2(i%)
              If picold$(i%) = picf$ Then GoTo ld200
              myfile = Dir(picf$)
              If myfile <> sEmpty Then
                 Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
              Else 'load blank bitmap
                 picf$ = Turbo2cdDir$ & "Itmv2s\MAP50\CLI0707.dl"
                 If picold$(i%) = picf$ Then GoTo ld200
                 Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
                 End If
              picold$(i%) = picf$
           Else
              picf$ = Turbo2cdDir$ & "Itmv2s\MAP50\CLI0" + pic1$(i%) + "." + pic2(i%)
              If picold$(i%) = picf$ Then GoTo ld200
              myfile = Dir(picf$)
              If myfile <> sEmpty Then
                 Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
              Else 'load blank bitmap
                 picf$ = Turbo2cdDir$ & "Itmv2s\MAP50\CLI0707.dl"
                 If picold$(i%) = picf$ Then GoTo ld200
                 Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
                 End If
              picold$(i%) = picf$
              End If
ld200:
           'map tile names loaded
           'Maps.Caption = Maps.Caption & " " & picf$

        Next i%
     ElseIf topotype% = 1 Then 'new 50000 maps
       kmxorigin = Int(kmxcc / 10000) * 100
       If kmycc > 1000000 Then
          kmyorigin = Int((kmycc - 1000000) / 10000)
       Else
          kmyorigin = Int(kmycc / 10000)
          End If

       centpic& = kmxorigin + kmyorigin + 1
       pic1$(1) = Str$(centpic&)
       If centpic& < 1000 Then pic1$(1) = "0" & pic1$(1)
       pic1$(2) = Str$(centpic& - 99)
       If centpic& < 1000 Then pic1$(2) = "0" & pic1$(2)
       pic1$(3) = Str$(centpic& + 1)
       If centpic& < 1000 Then pic1$(3) = "0" & pic1$(3)
       pic1$(4) = Str$(centpic& + 101)
       If centpic& < 1000 Then pic1$(4) = "0" & pic1$(4)
       pic1$(5) = Str$(centpic& + 100)
       If centpic& < 1000 Then pic1$(5) = "0" & pic1$(5)
       pic1$(6) = Str$(centpic& + 99)
       If centpic& < 1000 Then pic1$(6) = "0" & pic1$(6)
       pic1$(7) = Str$(centpic& - 1)
       If centpic& < 1000 Then pic1$(7) = "0" & pic1$(7)
       pic1$(8) = Str$(centpic& - 101)
       If centpic& < 1000 Then pic1$(8) = "0" & pic1$(8)
       pic1$(9) = Str$(centpic& - 100)
       If centpic& < 1000 Then pic1$(9) = "0" & pic1$(9)

       kmxorigin = Int(kmxcc / 10000) * 10000
       kmyorigin = Int((kmycc - 1000000) / 10000) * 10000 + 1000000

       'now load the buffers
        For i% = 1 To 9
            picf$ = Turbo2cdDir$ & "Itmv3.1\map50\cli" + LTrim$(pic1$(i%) & ".bmp")
            If picold$(i%) = picf$ Then GoTo ld300
            myfile = Dir(picf$)
            If myfile <> sEmpty Then
               Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
            Else 'load blank bitmap
               picf$ = Turbo2cdDir$ & "Itmv3.1\map50\cli0707.bmp"
               If picold$(i%) = picf$ Then GoTo ld300
               Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
               End If
           picold$(i%) = picf$
ld300:
        Next i%


       End If

ElseIf map400 = True Then

    n400x%(1) = Int((kmxcc - 80000) / 80000) + 1
    s11 = n400x%(1) - 1
    kmx400origin = 80000 + s11 * 80000
    dxx% = Int((kmxcc - kmx400origin) / 40000) 'Mod 2
    s111 = dxx%
    kmx400origin = 80000 + s11 * 80000 + 40000 * s111
    If dxx% = 0 Then
      Mid$(pic2(1), 2, 1) = "l"
      Mid$(pic2(2), 2, 1) = "r"
      n400x%(2) = n400x%(1) - 1
      Mid$(pic2(3), 2, 1) = "l"
      n400x%(3) = n400x%(1)
      Mid$(pic2(4), 2, 1) = "r"
      n400x%(4) = n400x%(1)
      Mid$(pic2(5), 2, 1) = "r"
      n400x%(5) = n400x%(1)
      Mid$(pic2(6), 2, 1) = "r"
      n400x%(6) = n400x%(1)
      Mid$(pic2(7), 2, 1) = "l"
      n400x%(7) = n400x%(1)
      Mid$(pic2(8), 2, 1) = "r"
      n400x%(8) = n400x%(1) - 1
      Mid$(pic2(9), 2, 1) = "r"
      n400x%(9) = n400x%(1) - 1
    Else
      Mid$(pic2(1), 2, 1) = "r"
      Mid$(pic2(2), 2, 1) = "l"
      n400x%(2) = n400x%(1)
      Mid$(pic2(3), 2, 1) = "r"
      n400x%(3) = n400x%(1)
      Mid$(pic2(4), 2, 1) = "l"
      n400x%(4) = n400x%(1) + 1
      Mid$(pic2(5), 2, 1) = "l"
      n400x%(5) = n400x%(1) + 1
      Mid$(pic2(6), 2, 1) = "l"
      n400x%(6) = n400x%(1) + 1
      Mid$(pic2(7), 2, 1) = "r"
      n400x%(7) = n400x%(1)
      Mid$(pic2(8), 2, 1) = "l"
      n400x%(8) = n400x%(1)
      Mid$(pic2(9), 2, 1) = "l"
      n400x%(9) = n400x%(1)
      End If


   n400y%(1) = Int((kmycc - 840000) / 80000) + 1
   s22 = n400y%(1) - 1
   kmy400origin = 840000 + s22 * 80000
   dyy% = Int((kmycc - kmy400origin) / 40000) 'Mod 2
   s33 = dyy%
   kmy400origin = 840000 + s22 * 80000 + 40000 * s33
   If dyy% = 0 Then
      Mid$(pic2(1), 1, 1) = "d"
      Mid$(pic2(2), 1, 1) = "u"
      n400y%(2) = n400y%(1)
      Mid$(pic2(3), 1, 1) = "u"
      n400y%(3) = n400y%(1)
      Mid$(pic2(4), 1, 1) = "u"
      n400y%(4) = n400y%(1)
      Mid$(pic2(5), 1, 1) = "d"
      n400y%(5) = n400y%(1)
      Mid$(pic2(6), 1, 1) = "u"
      n400y%(6) = n400y%(1) - 1
      Mid$(pic2(7), 1, 1) = "u"
      n400y%(7) = n400y%(1) - 1
      Mid$(pic2(8), 1, 1) = "u"
      n400y%(8) = n400y%(1) - 1
      Mid$(pic2(9), 1, 1) = "d"
      n400y%(9) = n400y%(1)
    Else
      Mid$(pic2(1), 1, 1) = "u"
      Mid$(pic2(2), 1, 1) = "d"
      n400y%(2) = n400y%(1) + 1
      Mid$(pic2(3), 1, 1) = "d"
      n400y%(3) = n400y%(1) + 1
      Mid$(pic2(4), 1, 1) = "d"
      n400y%(4) = n400y%(1) + 1
      Mid$(pic2(5), 1, 1) = "u"
      n400y%(5) = n400y%(1)
      Mid$(pic2(6), 1, 1) = "d"
      n400y%(6) = n400y%(1)
      Mid$(pic2(7), 1, 1) = "d"
      n400y%(7) = n400y%(1)
      Mid$(pic2(8), 1, 1) = "d"
      n400y%(8) = n400y%(1)
      Mid$(pic2(9), 1, 1) = "u"
      n400y%(9) = n400y%(1)
      End If

   'now load the buffers
   For i% = 1 To 9
       'DoEvents
       pic1$(i%) = 792 + 8 * (n400y%(i%) - 1) + 800 * (n400x%(i%) - 1)
       If n400x%(i%) >= 2 Then
          picf$ = Turbo2cdDir$ & "Itmv2s\MAP400\C4u" + pic1$(i%) + "." + pic2(i%)
          If picold$(i%) = picf$ Then GoTo ld400
          myfile = Dir(picf$)
          If myfile <> sEmpty Then
             Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
          Else 'load blank bitmap
             picf$ = Turbo2cdDir$ & "Itmv2s\MAP50\CLI0707.dl"
             If picold$(i%) = picf$ Then GoTo ld400
             Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
             End If
          picold$(i%) = picf$
       Else
          picf$ = Turbo2cdDir$ & "Itmv2s\MAP400\C4u0" + pic1$(i%) + "." + pic2(i%)
          If picold$(i%) = picf$ Then GoTo ld400
          myfile = Dir(picf$)
          If myfile <> sEmpty Then
             Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
          Else 'load blank bitmap
             picf$ = Turbo2cdDir$ & "Itmv2s\MAP50\CLI0707.dl"
             If picold$(i%) = picf$ Then GoTo ld400
             Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
             End If
          picold$(i%) = picf$
          End If
ld400:
   Next i%
   
ElseIf world = True Then '''''''<<<<<<<<<<<world tiles and imported maps>>>>>>>>>>>>>>>>>>>>>>

   If mapimport = False Then
      n400x%(1) = Int((lonc + 900) / 360) + 1
      s11 = n400x%(1) - 1
      worldxorigin = -900 + s11 * 360
      dxx% = Int((lonc - worldxorigin) / 180)
      s111 = dxx%
      worldxorigin = -900 + s11 * 360 + 180 * s111
   Else
      'If Val(lon) >= woxorigin And Val(lon) < woxorigin + deglog Then
      '   n400x%(1) = 1
      'Else
      '   n400x%(1) = 0
      '   End If
      'dxx% = Int((Val(lonc) - woxorigin) / deglog) + 1
      'worldxorigin = woxorigin + deglog * Int((lonc - woxorigin) / deglog + 0.5)
      n400x%(1) = Int((lonc - woxorigin) / (2 * deglog)) + 1
      s11 = n400x%(1) - 1
      worldxorigin = woxorigin + s11 * 2# * deglog
      dxx% = Int((lonc - worldxorigin) / deglog)
      s111 = dxx%
      worldxorigin = woxorigin + s11 * 2# * deglog + s111 * deglog
      End If

    If dxx% = 0 Then
      Mid$(pic2(1), 2, 1) = "l"
      Mid$(pic2(2), 2, 1) = "r"
      n400x%(2) = n400x%(1) - 1
      Mid$(pic2(3), 2, 1) = "l"
      n400x%(3) = n400x%(1)
      Mid$(pic2(4), 2, 1) = "r"
      n400x%(4) = n400x%(1)
      Mid$(pic2(5), 2, 1) = "r"
      n400x%(5) = n400x%(1)
      Mid$(pic2(6), 2, 1) = "r"
      n400x%(6) = n400x%(1)
      Mid$(pic2(7), 2, 1) = "l"
      n400x%(7) = n400x%(1)
      Mid$(pic2(8), 2, 1) = "r"
      n400x%(8) = n400x%(1) - 1
      Mid$(pic2(9), 2, 1) = "r"
      n400x%(9) = n400x%(1) - 1
    Else
      Mid$(pic2(1), 2, 1) = "r"
      Mid$(pic2(2), 2, 1) = "l"
      n400x%(2) = n400x%(1)
      Mid$(pic2(3), 2, 1) = "r"
      n400x%(3) = n400x%(1)
      Mid$(pic2(4), 2, 1) = "l"
      n400x%(4) = n400x%(1) + 1
      Mid$(pic2(5), 2, 1) = "l"
      n400x%(5) = n400x%(1) + 1
      Mid$(pic2(6), 2, 1) = "l"
      n400x%(6) = n400x%(1) + 1
      Mid$(pic2(7), 2, 1) = "r"
      n400x%(7) = n400x%(1)
      Mid$(pic2(8), 2, 1) = "l"
      n400x%(8) = n400x%(1)
      Mid$(pic2(9), 2, 1) = "l"
      n400x%(9) = n400x%(1)
      End If

   If mapimport = False Then
      n400y%(1) = Int((latc + 810) / 360) + 1
      s22 = n400y%(1) - 1
      worldyorigin = -810 + s22 * 360
      dyy% = Int((latc - worldyorigin) / 180)
      s33 = dyy%
      worldyorigin = -810 + s22 * 360 + 180 * s33
   Else
      'If Val(lat) >= woyorigin And Val(lat) < woyorigin + deglat Then
      '   n400y%(1) = 1
      'Else
      '   n400y%(1) = 0
      '   End If
      'dyy% = Int((Val(latc) - woyorigin) / deglat)
      'worldyorigin = woyorigin + deglat * Int((latc - woyorigin) / deglat + 0.5)
      n400y%(1) = Int((latc - woyorigin) / (2 * deglat)) + 1
      s22 = n400y%(1) - 1
      worldyorigin = woyorigin + s22 * 2 * deglat
      dyy% = Int((latc - worldyorigin) / deglat)
      s33 = dyy%
      worldyorigin = woyorigin + s22 * 2 * deglat + s33 * deglat
      End If

   If dyy% = 0 Then
      Mid$(pic2(1), 1, 1) = "d"
      Mid$(pic2(2), 1, 1) = "u"
      n400y%(2) = n400y%(1)
      Mid$(pic2(3), 1, 1) = "u"
      n400y%(3) = n400y%(1)
      Mid$(pic2(4), 1, 1) = "u"
      n400y%(4) = n400y%(1)
      Mid$(pic2(5), 1, 1) = "d"
      n400y%(5) = n400y%(1)
      Mid$(pic2(6), 1, 1) = "u"
      n400y%(6) = n400y%(1) - 1
      Mid$(pic2(7), 1, 1) = "u"
      n400y%(7) = n400y%(1) - 1
      Mid$(pic2(8), 1, 1) = "u"
      n400y%(8) = n400y%(1) - 1
      Mid$(pic2(9), 1, 1) = "d"
      n400y%(9) = n400y%(1)
    Else
      Mid$(pic2(1), 1, 1) = "u"
      Mid$(pic2(2), 1, 1) = "d"
      n400y%(2) = n400y%(1) + 1
      Mid$(pic2(3), 1, 1) = "d"
      n400y%(3) = n400y%(1) + 1
      Mid$(pic2(4), 1, 1) = "d"
      n400y%(4) = n400y%(1) + 1
      Mid$(pic2(5), 1, 1) = "u"
      n400y%(5) = n400y%(1)
      Mid$(pic2(6), 1, 1) = "d"
      n400y%(6) = n400y%(1)
      Mid$(pic2(7), 1, 1) = "d"
      n400y%(7) = n400y%(1)
      Mid$(pic2(8), 1, 1) = "d"
      n400y%(8) = n400y%(1)
      Mid$(pic2(9), 1, 1) = "u"
      n400y%(9) = n400y%(1)
      End If

   'now load the buffers (first check for busy signal from egg.exe)
    If Maps.Timer2.Enabled = True Then
       myfile = Dir(ramdrive + ":\wait.x")
       If myfile <> sEmpty Then
           waitime = Timer
           Do Until Timer > waitime + 0.5
              DoEvents
           Loop
           Exit Sub
           End If
       End If

   For i% = 1 To 9
       'DoEvents
       pic1$(i%) = LTrim$(RTrim$(Str$(n400x%(i%)))) + LTrim$(RTrim$(Str$(n400y%(i%))))
       If mapimport = False Then
            picf$ = Turbo2cdDir$ & "worldl" + pic1$(i%) + "." + pic2(i%)
            If picold$(i%) = picf$ Then GoTo ld600
            myfile = Dir(picf$)
            If myfile <> sEmpty Then
               Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
            Else 'load blank bitmap
               picf$ = Turbo2cdDir$ & "worldl33.ur"
               If picold$(i%) = picf$ Then GoTo ld600
               Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
               End If
            picold$(i%) = picf$
ld600:
      Else
         If pic1$(i%) = "11" And pic2$(i%) = "dl" Then
            picf$ = mapfile$
            myfile = Dir(picf$)
            If myfile <> sEmpty Then
               Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
               End If
         Else
            picf$ = blank$
            Maps.PictureClip1(i% - 1).Picture = LoadPicture(picf$)
            End If
         End If
   Next i%
   Exit Sub

'ld700: 'imported picture is always loaded into picture clip 0
'   'load the blank files into the other picture clips
'   For i% = 1 To 9
'      If i% = 1 Then
'         Maps.PictureClip1(i% - 1).Picture = LoadPicture(mapfile$)
'      Else
'         Maps.PictureClip1(i% - 1).Picture = LoadPicture(blank$)
'         End If
'   Next i%
   End If

   On Error GoTo 0
   Exit Sub

loadpictures_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadpictures of Module MapModule"

End Sub
Public Sub blitpictures()
   Dim lResult As Long
   On Error GoTo errblit
'   ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)

   If mapwi = mapPictureform.Width Then
      'ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
      If mapwi2 > mapwi - mapwi * 0.1 Then
         mapwi2 = mapwi
      Else
         mapwi2 = defaultmapwidth '0.469 * mapwi
         End If
   Else
     mapwi2 = mapPictureform.Width
     End If
   If maphi = mapPictureform.Height Or printing = True Then
      If maphi2 > maphi - maphi * 0.1 Then
         maphi2 = maphi
      Else
         maphi2 = defaultmapheight '0.79 * maphi
         End If
   Else
     maphi2 = mapPictureform.Height
     End If

'If Abs(mapxdif) > 100 Then
'   mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
'    End If
'If Abs(mapydif) > 100 Then
'   mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
'   End If

If world = False Then
    'keep EY topo maps within EY boundaries
    If map400 Then
       If kmxc > 250000 Then kmxc = 73000
       If kmxc < 73000 Then kmxc = 250000
       If kmyc > 1310000 Then kmyc = 83000
       If kmyc < 83000 Then kmyc = 1310000
    ElseIf map50 Then
       If kmxc > 250000 Then kmxc = 73000
       If kmxc < 73000 Then kmxc = 250000
       If kmyc > 1320000 Then kmyc = 83000
       If kmyc < 83000 Then kmyc = 1320000
       End If
    
    If mag > 1 Then
       kmxcc = kmxc
       kmycc = kmyc
    Else
       If map50 = True Then
          kmxcc = kmxc + (km50x) * (mapwi - mapwi2 + mapxdif) / 2
          kmycc = kmyc - (km50y) * (maphi - maphi2 + mapydif) / 2
       ElseIf map400 Then
          kmxcc = kmxc + (km400x) * (mapwi - mapwi2 + mapxdif) / 2
          kmycc = kmyc - (km400y) * (maphi - maphi2 + mapydif) / 2
          End If
        End If
ElseIf world = True Then

   '(first check for busy signal from egg.exe)
    If Maps.Timer2.Enabled = True Then
       myfile = Dir(ramdrive + ":\wait.x")
       If myfile <> sEmpty Then
           waitime = Timer
           Do Until Timer > waitime + 0.5
              DoEvents
           Loop
           Exit Sub
           End If
       End If

    If mapimport = False Then
       deglog = 180
       deglat = 180
       End If
    If mag > 1 Then
       If impcenter Then
            lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx  '+ 0.166
            latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy  '+ 0.204
       Else
          If lonc = lon And latc = lat Then
             lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx  '+ 0.166
             latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy  '+ 0.204
         Else
             lonc = lon '+ fudx / mag
             latc = lat '+ fudy / mag
             End If
          End If
       If Not mapimport Then
            If lon > 180 Then lon = lon - 360
            If lon < -180 Then lon = lon + 360
            If lat > 90 Then lat = lat - 180
            If lat < -90 Then lat = lat + 180
       Else
            If lon > woxorigin + deglog Then lon = lon - deglog
            If lon < woxorigin Then lon = lon + woxorigin
            If lat > woyorigin + deglat Then lat = lat - deglat
            If lat < woyorigin Then lat = lat + deglat
            End If
    Else
       'lonc = lon + (180 / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
       'latc = lat - (180 / sizewy) * (maphi - maphi2 + mapydif) / 2
       'keep lon and lat aways in boundaries of world
       If Not mapimport Then
            If lon > 180 Then lon = lon - 360
            If lon < -180 Then lon = lon + 360
            If lat > 90 Then lat = lat - 180
            If lat < -90 Then lat = lat + 180
       Else
            If lon > woxorigin + deglog Then lon = lon - deglog
            If lon < woxorigin Then lon = lon + deglog
            If lat > woyorigin + deglat Then lat = lat - deglat
            If lat < woyorigin Then lat = lat + deglat
            End If
       lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx  '+ 0.166
       latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy  '+ 0.204
       End If
    End If


   If world = False Then
       sgny = 0
       'determine the clipping regions
       'Defining Wi=sizex,Hi=sizey, then
       'kmxc,kmyc are located at X=Wi/2,Y=Hi/2 when kmxorigin,kmyorigin
       'are at Wi/2-(kmxc-kmxorigin)*/km50x,Hi/2+(kmyc-kmyorigin)/km50y
bl10:   If map50 = True Then
           xorigin = sizex / 2 - (kmxcc - kmxorigin) / km50x ' (8820*mag) / 5000
           yorigin = sizey / 2 + (kmycc - kmyorigin) / km50y ' * (8820*mag) / 5000
           'If topotype% = 1 Then
           '   xorigin = xorigin - 5800
           '   yorigin = yorigin - 10650
           '   kmxc = kmxc + 5800 * km50x
           '   kmyc = kmyc - 10650 * km50y
           '   End If
           kmm = km50x 'sizex
        ElseIf map400 = True Then
           xorigin = sizex / 2 - (kmxcc - kmx400origin) / km400x ' * (8820*mag) / 40000
           yorigin = sizey / 2 + (kmycc - kmy400origin) / km400y ' * (8820*mag) / 40000
           kmm = km400x 'sizex
           End If
       If xorigin <= -sizex Or xorigin >= sizex Or yorigin <= 0 Or yorigin >= 2 * sizey Then
           Call loadpictures  'load appropriate map tiles into off-screen buffers
           GoTo bl10
           End If
       If xorigin >= 0 And yorigin <= sizey Then
           'blit pieces of back buffers 1,7,8,9 to the screen (front buffer)
           If map50 = True Then
              kmxn = kmxorigin
              kmyn = kmyorigin
           ElseIf map400 = True Then
              kmxn = kmx400origin
              kmyn = kmy400origin
              End If
           'sgny = 0
           bn%(1) = 0 'NE
           bn%(2) = 8  'NW
           bn%(3) = 6  'SE
           bn%(4) = 7  'SW
       ElseIf xorigin >= 0 And yorigin > sizey Then
          'blit pieces of back buffers 1,2,3,9 to the screen (front buffer)
           yorigin = yorigin - sizey
           If map50 = True Then
              kmxn = kmxorigin
              kmyn = kmyorigin + 5000 * (1 + topotype%)
           ElseIf map400 = True Then
              kmyn = kmy400origin + 40000
              kmxn = kmx400origin
              End If
           'sgny = 1
           bn%(1) = 2
           bn%(2) = 1
           bn%(3) = 0
           bn%(4) = 8
       ElseIf xorigin < 0 And yorigin > sizey Then
          'blit pieces of back buffers 1,3,4,5 to the screen (front buffer)
           yorigin = yorigin - sizey
           xorigin = xorigin + sizex
           If map50 = True Then
              kmxn = kmxorigin + 5000 * (1 + topotype%)
              kmyn = kmyorigin
           ElseIf map400 = True Then
              kmxn = kmx400origin + 40000
              kmyn = kmy400origin
              End If
           sgny = 1
           bn%(1) = 3
           bn%(2) = 2
           bn%(3) = 4
           bn%(4) = 0
      ElseIf xorigin < 0 And yorigin <= sizey Then
          'blit pieces of back buffers 1,5,6,7 to the screen (front buffer)
           xorigin = xorigin + sizex
           If map50 = True Then
              kmxn = kmxorigin + 5000 * (1 + topotype%)
              kmyn = kmyorigin + 5000 * (1 + topotype%)
           ElseIf map400 = True Then
              kmxn = kmx400origin + 40000
              kmyn = kmy400origin + 40000
              End If
           sgny = -1
           bn%(1) = 4
           bn%(2) = 0
           bn%(3) = 5
           bn%(4) = 6
           End If

        bufwi(1, 1) = (sizex - xorigin) * mag
        bufwi(2, 1) = bufwi(1, 1) / mag
        bufhi(1, 1) = yorigin * mag
        bufhi(2, 1) = bufhi(1, 1) / mag
        bufx(1, 1) = xorigin - (mag - 1) * (kmxcc - kmxn) / kmm 'origin to blit to on screen
        bufx(2, 1) = 0 'origin of blit in buffer
        bufy(1, 1) = yorigin * (1 - mag) + (mag - 1) * (kmycc - kmyn) / kmm - sgny * sizey * (mag - 1)
        bufy(2, 1) = sizey - yorigin

        bufhi(1, 2) = bufhi(1, 1)
        bufhi(2, 2) = bufhi(1, 1) / mag
        bufwi(1, 2) = xorigin * mag
        bufwi(2, 2) = bufwi(1, 2) / mag
        bufx(1, 2) = xorigin * (1 - mag) - (mag - 1) * (kmxcc - kmxn) / kmm
        bufx(2, 2) = sizex - xorigin
        bufy(1, 2) = bufy(1, 1)
        bufy(2, 2) = bufy(2, 1)

        bufx(1, 3) = bufx(1, 1)
        bufx(2, 3) = bufx(2, 1)
        bufwi(1, 3) = bufwi(1, 1) * mag
        bufwi(2, 3) = bufwi(1, 3) / mag
        bufy(1, 3) = yorigin + (mag - 1) * (kmycc - kmyn) / kmm - sgny * sizey * (mag - 1)
        bufy(2, 3) = 0
        bufhi(1, 3) = (sizey - yorigin) * mag
        bufhi(2, 3) = bufhi(1, 3) / mag

        bufx(1, 4) = bufx(1, 2)
        bufx(2, 4) = bufx(2, 2)
        bufwi(1, 4) = bufwi(1, 2)
        bufwi(2, 4) = bufwi(1, 2) / mag
        bufy(1, 4) = bufy(1, 3)
        bufy(2, 4) = 0
        bufhi(1, 4) = bufhi(1, 3)
        bufhi(2, 4) = bufhi(1, 3) / mag
        
ElseIf world = True Then

       sgny = 0
       'determine the clipping regions
       'Defining Wi=sizex,Hi=sizey, then
       'kmxc,kmyc are located at X=Wi/2,Y=Hi/2 when kmxorigin,kmyorigin
       'are at Wi/2-(kmxc-kmxorigin)*/km50x,Hi/2+(kmyc-kmyorigin)/km50y

bl100: xorigin = sizewx / 2# - (lonc + (mag - 1) * fudx / mag - worldxorigin) / (deglog / sizewx) '<??must stay constant
       yorigin = sizewy / 2# + (latc + (mag - 1) * fudy / mag - worldyorigin) / (deglat / sizewy)
       kmm = (deglat / sizewy)
       
'       kmm = (180 / sizewy)
       If xorigin <= -sizewx Or xorigin >= sizewx Or yorigin <= 0 Or yorigin >= 2 * sizewy Then
           lResult = FindWindow(vbNullString, "Extracting relevant portion of the DTM")
           If lResult > 0 Then Exit Sub

           '(first check for busy signal from egg.exe)
           If Maps.Timer2.Enabled = True Then
              myfile = Dir(ramdrive + ":\wait.x")
              If myfile <> sEmpty Then
                 waitime = Timer
                 Do Until Timer > waitime + 0.5
                    DoEvents
                 Loop
                 Exit Sub
                 End If
              End If
              

           Call loadpictures  'load appropriate map tiles into off-screen buffers
           GoTo bl100
           End If
       
       If xorigin >= 0 And yorigin <= sizewy / mag Then
           'blit pieces of back buffers 1,7,8,9 to the screen (front buffer)
           kmxn = worldxorigin
           kmyn = worldyorigin
           bn%(1) = 0 'NE
           bn%(2) = 8  'NW
           bn%(3) = 6  'SE
           bn%(4) = 7  'SW
       ElseIf xorigin >= 0 And yorigin > sizewy / mag Then
          'blit pieces of back buffers 1,2,3,9 to the screen (front buffer)
           If Not mapimport Then
                yorigin = yorigin - sizewy
                kmxn = worldxorigin
                kmyn = worldyorigin + deglog
                If mapimport = True Then
                   End If
                bn%(1) = 2
                bn%(2) = 1
                bn%(3) = 0
                bn%(4) = 8
            Else
                yorigin = yorigin - sizewy
                xorigin = xorigin + sizewx
                kmxn = worldxorigin + deglog
                kmyn = worldyorigin
                sgny = 1
                bn%(1) = 3
                bn%(2) = 2
                bn%(3) = 4
                bn%(4) = 0
                End If
       ElseIf xorigin < 0 And yorigin > sizewy / mag Then
          'blit pieces of back buffers 1,3,4,5 to the screen (front buffer)
           yorigin = yorigin - sizewy
           xorigin = xorigin + sizewx
           kmxn = worldxorigin + deglog
           kmyn = worldyorigin
           sgny = 1
           bn%(1) = 3
           bn%(2) = 2
           bn%(3) = 4
           bn%(4) = 0
      ElseIf xorigin < 0 And yorigin <= sizewy / mag Then
          'blit pieces of back buffers 1,5,6,7 to the screen (front buffer)
           xorigin = xorigin + sizewx
           kmxn = worldxorigin + deglog
           kmyn = worldyorigin + deglat
           sgny = -1
           bn%(1) = 4
           bn%(2) = 0
           bn%(3) = 5
           bn%(4) = 6
           End If

        bufwi(1, 1) = (sizewx - xorigin) * mag
        bufwi(2, 1) = bufwi(1, 1) / mag
        bufhi(1, 1) = yorigin * mag
        bufhi(2, 1) = bufhi(1, 1) / mag
        bufx(1, 1) = xorigin - (mag - 1) * (lonc - kmxn) / kmm 'origin to blit to on screen
        bufx(2, 1) = 0 'origin of blit in buffer
        bufy(1, 1) = yorigin * (1 - mag) + (mag - 1) * (latc - kmyn) / kmm - sgny * sizewy * (mag - 1)
        bufy(2, 1) = sizewy - yorigin

        bufhi(1, 2) = bufhi(1, 1)
        bufhi(2, 2) = bufhi(1, 1) / mag
        bufwi(1, 2) = xorigin * mag
        bufwi(2, 2) = bufwi(1, 2) / mag
        bufx(1, 2) = xorigin * (1 - mag) - (mag - 1) * (lonc - kmxn) / kmm
        bufx(2, 2) = sizewx - xorigin
        bufy(1, 2) = bufy(1, 1)
        bufy(2, 2) = bufy(2, 1)

        bufx(1, 3) = bufx(1, 1)
        bufx(2, 3) = bufx(2, 1)
        bufwi(1, 3) = bufwi(1, 1) * mag
        bufwi(2, 3) = bufwi(1, 3) / mag
        bufy(1, 3) = yorigin + (mag - 1) * (latc - kmyn) / kmm - sgny * sizewy * (mag - 1)
        bufy(2, 3) = 0
        bufhi(1, 3) = (sizewy - yorigin) * mag
        bufhi(2, 3) = bufhi(1, 3) / mag

        bufx(1, 4) = bufx(1, 2)
        bufx(2, 4) = bufx(2, 2)
        bufwi(1, 4) = bufwi(1, 2)
        bufwi(2, 4) = bufwi(1, 2) / mag
        bufy(1, 4) = bufy(1, 3)
        bufy(2, 4) = 0
        bufhi(1, 4) = bufhi(1, 3)
        bufhi(2, 4) = bufhi(1, 3) / mag
        End If

    If mag > 1 Then 'Or world = True Then
       For i% = 1 To 4
             bufx(1, i%) = bufx(1, i%) - (mapwi - mapwi2 + mapxdif) / 2
             bufy(1, i%) = bufy(1, i%) - (maphi - maphi2 + mapydif) / 2
       Next i%
       End If

    If mapPictureform.Visible = False And ((mapwi2 = 0 Or (mapwi = mapwi2)) And (maphi2 = 0 Or (maphi = maphi2))) Then
       mapPictureform.Top = defaultmaptop
       mapPictureform.Height = defaultmapheight '0.79 * maphi
       mapPictureform.Width = defaultmapwidth '0.469 * mapwi
       mapPictureform.Visible = True
'       ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
       BringWindowToTop (mapPictureform.hwnd)
    ElseIf mapPictureform.Visible = False Then
       mapPictureform.Top = defaultmaptop
       mapPictureform.Height = maphi2
       mapPictureform.Width = mapwi2
       mapPictureform.Visible = True
'       ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
       BringWindowToTop (mapPictureform.hwnd)
       End If

    If world = True Then
       lResult = FindWindow(vbNullString, "Extracting relevant portion of the DTM")
       If lResult > 0 Then Exit Sub
       If Maps.Timer2.Enabled = True Then
           myfile = Dir(ramdrive + ":\wait.x")
           If myfile <> sEmpty Then
               waitime = Timer
               Do Until Timer > waitime + 0.5
                  DoEvents
               Loop
               Exit Sub
               End If
           End If
       End If

    mapPictureform.mapPicture.Cls
       For i% = 1 To 4
         'check for nonsense widths,heights that cause program to bomb
         If bufwi(1, i%) <= 0 Then GoTo bl500
         If bufhi(1, i%) <= 0 Then GoTo bl500
         If bufwi(2, i%) <= 0 Then GoTo bl500
         If bufhi(2, i%) <= 0 Then GoTo bl500
         'DoEvents
         mapPictureform.mapPicture.PaintPicture Maps.PictureClip1(bn%(i%)).Picture, bufx(1, i%), bufy(1, i%), bufwi(1, i%), bufhi(1, i%), bufx(2, i%), bufy(2, i%), bufwi(2, i%), bufhi(2, i%)
bl500: Next i%
       newblit = True
'       If topotype% = 0 Then
'          mapxdifn = mapxdif
'          mapydifn = mapydif
'       ElseIf topotype% = 1 Then
'          mapxdifn = 60
'          mapydifn = 70
'          End If

          mapxdifn = 0
          mapydifn = 0

       If obstflag = False Then
          mapPictureform.mapPicture.DrawWidth = 1 '1 * mag
          mapPictureform.mapPicture.Circle (CSng(mapPictureform.Width / 2 - mapxdifn), CSng(mapPictureform.Height / 2 - mapydifn)), 100, 255 '100 * mag, 255
          mapPictureform.mapPicture.DrawWidth = 2 '2 * mag
          mapPictureform.mapPicture.Circle (CSng(mapPictureform.Width / 2 - mapxdifn), CSng(mapPictureform.Height / 2 - mapydifn)), 20, 255 '20 * mag, 255
          mapPictureform.mapPicture.DrawWidth = 1 '1 * mag
       Else
          mapPictureform.mapPicture.DrawWidth = 1 '1 * mag
          mapPictureform.mapPicture.Circle (CSng(mapPictureform.Width / 2 - mapxdifn), CSng(mapPictureform.Height / 2 - mapydifn)), 100, QBColor(14) '100 * mag, QBColor(14)
          mapPictureform.mapPicture.DrawWidth = 2 '2 * mag
          mapPictureform.mapPicture.Circle (CSng(mapPictureform.Width / 2 - mapxdifn), CSng(mapPictureform.Height / 2 - mapydifn)), 20, QBColor(14) '20 * mag, QBColor(14)
          mapPictureform.mapPicture.DrawWidth = 1 '1 * mag
          Call obstructions(mapPictureform.mapPicture)
          End If
       If showroute = True Or crosssection = True Then
          Call showtheroute(mapPictureform.mapPicture)
          End If
          
       If SearchVis Then 'replot search results
          mapsearchfm.cmdPlotSearchPnts.value = True
          End If
          
       Exit Sub
errblit:
   Resume Next
   'ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
   'response = MsgBox("The blitting operation from the map buffers encountered unexpected error #" + Str(Err.Number) + " ." + _
   '           "The values of the coordinates are: " + Str(i%) + "," + Str(bn(i%)) + "," + Str(bufx(1, i%)) + "," + Str(bufy(1, i%)) + "," + _
   '           Str(bufwi(1, i%)) + "," + Str(bufhi(1, i%)) + "," + Str(bufx(2, i%)) + "," + Str(bufy(2, i%)) + "," + Str(bufwi(2, i%)) + _
   '           "," + Str(bufhi(2, i%)) + "                                                       " + _
   '           "Do you want to cancel the translation and return to the origin?", vbYesNo + vbCritical, "Maps & More")
   'If response = vbYes Then
   '   For i% = 12 To 15
   '      tblbuttons(i%) = 0
   '      Maps.Toolbar1.Buttons(i%).Value = tbrUnpressed
   '   Next i%
   '   kmxc = 172352
   '   kmyc = 1131700
   '   mag = 1
   '   Maps.Combo1.Text = 100
   '   Call blitpictures
   '   End If
   'Exit Sub
End Sub
Public Sub skyTERRAgoto()
   Dim lResult As Long, xwin As Long, ywin As Long, winw As Long, winh As Long, winp As Long
   Dim dX As Long, dY As Long, bVk As Byte, ljump As Long, lerror As Long, lResult2 As Long
   'lResult = FindWindow(vbNullString, terranam$)
   If routeload = True Or Maps.Timer2.Enabled = True Then Exit Sub
skyt5: lResult = FindWindow(vbNullString, terranam$)
   If terranam$ = "TerraExplorer - " And lResult = 0 Then
      lResult2 = FindWindow(vbNullString, "TerraExplorer - " + terradir$ + "\Israel9.teh")
      If lResult2 > 0 Then terranam$ = "TerraExplorer - " + terradir$ + "\Israel9.teh"
      End If
   If lResult = 0 And lResult2 = 0 Then
      waitime = Timer
      Do Until Timer > waitime + 1
      Loop
      GoTo skyt5
      End If
   'convert to Sky coordinates if necessary
   erreturn% = 0
skyt10:
   If Maps.Label5.Caption = "SKYx" Then
      C1$ = Maps.Text5.Text
      C2$ = Maps.Text6.Text
      'C1p = C1$: C2p = T2
   ElseIf Maps.Label5.Caption = "ITMx" Then
      'convert to SKY coordinates
      ITM1 = Maps.Text5.Text
      ITM2 = Maps.Text6.Text
      Call ITMSKY(ITM1, ITM2, T1, T2, 1)
      C1$ = T1
      C2$ = T2
      'C1p = T1: C2p = T2
   ElseIf Maps.Label4.Caption = "UTMx" Then
      'convert from UTM to ITM and then to SKY
      G1 = Maps.Text5.Text
      G2 = Maps.Text6.Text
      
      'first turn off gps correction if flagged
      If ggpscorrection Then
        GpsCorrOff = True
        ggpscorrection = False
        End If
      Z = 33
      Call UTMGEO(G1, G2, Z, l1, l2)
      Call GEOCASC(l1, l2, kmxg, kmyg)
      If GpsCorrOff Then
         ggpscorrection = True
         GpsCorrOff = False
         End If

      ITM1 = Fix(0.5 + kmyg)
      If kmxg < 870000 Then
         ITM2 = Fix(0.5 + kmxg) + 1000000
      Else
         ITM2 = Fix(0.5 + kmxg)
         End If
      Call ITMSKY(ITM1, ITM2, T1, T2, 1)
      C1$ = T1
      C2$ = T2
      'C1p = T1: C2p = T2
      End If

   'C3$ = "-500"
   'in new version need to jump to X or Y boxes and click
   Clipboard.Clear
   Clipboard.SetText C1$
'   Clipboard.SetText SkyLightfm.Text4.Text
   'C1$ = SkyLightfm.Text4.Text
   ret = CloseClipboard()
   If lResult > 0 Or lResult2 > 0 Then
       'first check that there is no error message from last jump
       lerror = FindWindow(vbNullString, "TerraExplorer Error")
'       lerror = FindWindow(vbNullString, "TerraViewer Error")
       If lerror > 0 Then 'bring error message to top of z, and then cancel it
          ret = BringWindowToTop(lerror)
          Call keybd_event(VK_RETURN, 0, 0, 0) 'enters return
          Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
          'For i% = 1 To 2 'warn user that there was error by beeping twice
          '  Beep
          'Next i%
          End If
       Screen.MousePointer = vbHourglass
       ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
'       timewait = Timer + 0.05 'wait for the window to appear
'       Do Until Timer > timewait
'         DoEvents
'       Loop
       Screen.MousePointer = vbDefault
       'Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
       'Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
       'Call keybd_event(Asc("L"), 0, 0, 0)  'goes into location menu
       'Call keybd_event(Asc("L"), 0, KEYEVENTF_KEYUP, 0)
       'Call keybd_event(Asc("J"), 0, 0, 0)    'brings Jump to Location menu to top
       'Call keybd_event(Asc("J"), 0, KEYEVENTF_KEYUP, 0)
       'that was old setup, now move cursor to X or Y window and click

       If erreturn% <> 0 Then
          'try imputing the numbers again
          Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'drop-down Location menu
          Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
          GoTo sky50
          End If

       If skyleftjump = False Then
          dx1 = -80
          If WinVer = 5 Then dx1 = -70
          
          If gotobutton = True Then dx1 = 0 'coming from goto button
          dy1 = 221 '23 '33
          If WinVer = 5 Then dy1 = 220 '200
          
          If placdblclk = True Then 'coming from dblclick of place window
             dx1 = -245 'position cursor over goto button '<<<!!!>>> fix for XP
             dy1 = 25
             If WinVer = 5 Then
                dx1 = 0.9 * dx1
                dy1 = 0.9 * dy1
                End If
             Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
             dx1 = 5 '-80
             dy1 = 221
             If WinVer = 5 Then
                dx1 = 0.9 * dx1
                dy1 = 0.9 * dy1
                End If
             placdblclk = False
             End If
          Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
          Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'drop-down Location menu
          Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
       Else
          dxo = (mapwi2 - Xcoord) / 30 + 30
          dyo = (maphi2 - Ycoord) / 30 - 40
          If WinVer = 5 Then '<<<!!!>>>
             dxo = 0.9 * dxo
             dyo = 1.025 * dyo '0.9 * dyo
             End If
          dx1 = dxo
          dy1 = dyo
          Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
          Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'drop-down Location menu
          Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
          End If
       timewait = Timer + 0.05 'wait for location window to appear
       Do Until Timer > timewait
         DoEvents
       Loop
       '*********experimental*************
       'timewait = Timer + 0.05 'wait for the window to appear
       'Do Until Timer > timewait
       '  DoEvents
       'Loop
       'ljump = FindWindow(vbNullString, "Jump To Location")
       'If ljump > 0 Then
       '   Skycoord% = 1
       '   bRtn = EnumChildWindows(ljump, AddressOf EnumFunc, 1) 'read captions
       '   End If


'      dx1 = 12
'      dy1 = -261
'      Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item
'      Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'      Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'     'depress menu item "Jump To Locations"
'      timerwait = Timer + 0.05 'wait for the window to drop down
'      Do Until Timer > timerwait
'         DoEvents
'      Loop
'      dx1 = 0
'      dy1 = 33 '23 '33
'      Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0) 'move mouse to Location item
'      Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'drop-down Location menu
'      Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
'
'      timerwait = Timer + 0.05 'wait for the window to drop down
'      Do Until Timer > timerwait
'         DoEvents
'      Loop
'

sky50: Call keybd_event(VK_SHIFT, 0, 0, 0) 'enter SKYx
       Call keybd_event(VK_INSERT, 0, 0, 0)
       Call keybd_event(VK_INSERT, 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
       timerwait = Timer + 0.1
       Do Until Timer > timerwait
         DoEvents
       Loop
       Call keybd_event(VK_TAB, 0, 0, 0)
       Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)

       Clipboard.Clear
       Clipboard.SetText C2$
       'Clipboard.SetText SkyLightfm.Text1.Text
       ret = CloseClipboard()
       Call keybd_event(VK_SHIFT, 0, 0, 0) 'enters SKYy
       Call keybd_event(VK_INSERT, 0, 0, 0)
       Call keybd_event(VK_INSERT, 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(VK_RETURN, 0, 0, 0) 'enters return and goes to desired position

       Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0) '<<<add this if want to be able to rise up

       'check first for error message
       waitime = Timer
       Do Until Timer > waitime + 0.1
          DoEvents
       Loop
       lerror = FindWindow(vbNullString, "TerraExplorer Error")
'       lerror = FindWindow(vbNullString, "TerraViewer Error")
       If lerror > 0 Then 'bring error message to top of z, and then cancel it
          ret = BringWindowToTop(lerror)
          Call keybd_event(VK_RETURN, 0, 0, 0) 'enters return
          Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
          erreturn% = erreturn% + 1
          If erreturn% = 2 Then
             'reset coordinates after 2nd attempt
             Maps.Text5.Text = kmxsky
             Maps.Text6.Text = kmysky
             coordmode2% = 1
             Maps.Label5.Caption = "ITMx"
             Maps.Label6.Caption = "ITMy"
          ElseIf erreturn% > 6 Then
             'display error message and exit routine after 6th attempt
             ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             response = MsgBox("Something is wrong with the goto coordinates!", vbOKOnly + vbCritical, "Maps & more")
'             ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'             ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             BringWindowToTop (lResult)
             BringWindowToTop (mapPictureform.hwnd)
             Exit Sub
             End If
          GoTo skyt10
          End If

       'immediately jump out of picture area
'       dx1 = -12 '-34
'       dy1 = 153  '160
'       Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0) 'move mouse to Location item
'         dxo = (mapwi2 - Xcoord) / 30 + 50
'         dyo = (maphi2 - Ycoord) / 30 - 30
'         shifx = mapwi2 / 60
'         shify = maphi2 / 60
'         dx1 = dxo
'         dy1 = dyo
'         Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item
'         dx1 = -(mapwi2 / 60) - 50 - mapxdif / 30
'         dy1 = -(maphi2 / 60) + 30 - mapydif / 30
'         Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item



'        ***old program
'         waitime = Timer
'         Do Until Timer > waitime + 0.1
'            DoEvents
'         Loop
'         'go to minimum ground elevation
'         'move mouse to active region for TerraViewer
'         dxo = (mapwi2 - Xcoord) / 30 + 50
'         dyo = (maphi2 - Ycoord) / 30 - 30
'         shifx = mapwi2 / 60
'         shify = maphi2 / 60
'         dx1 = dxo
'         dy1 = dyo
'         Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item
'         'Call BringWindowToTop(lResult)
'         For i% = 1 To 20 'hold down the keys for a bit
'            Call keybd_event(VK_SHIFT, 0, 0, 0)
'            Call keybd_event(Asc("X"), 0, 0, 0)  'goes into Settings menu
'         Next i%
'         Call keybd_event(Asc("X"), 0, KEYEVENTF_KEYUP, 0)
'         Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
'         ''move pointer back to original position
'         'dx1 = -dxo
'         'dy1 = -dyo
'
'         'return pointer to middle of picture
'         dx1 = -(mapwi2 / 60) - 50 - mapxdif / 30
'         dy1 = -(maphi2 / 60) + 30 - mapydif / 30
'         Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item
'         'Call BringWindowToTop(mapPictureform.hwnd)

         'If terwt = True Then
         '  'wait some time to activate keyboard controls for TerraViewer
         '   'waittime = Timer + 1 '0.1
         '   'Do Until Timer > waittime
         '      Call keybd_event(88, 0, 0, 0) 'activate these two keys inorder to allow for individual presses
         '      Call keybd_event(88, 0, KEYEVENTF_KEYUP, 0)
         '   '   DoEvents
         '   'Loop
         ' '  For i% = 1 To 40 'go to minimum altitude
         '''   'Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'drop-down Location menu
         '''   'Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
         '      'Call keybd_event(VK_SHIFT, 0, 0, 0) 'enters SKYy
         ''<<<  Call keybd_event(88, 0, 0, 0) 'activate these two keys inorder to allow for individual presses
         ''<<<  Call keybd_event(88, 0, KEYEVENTF_KEYUP, 0)
         '      'Call keybd_event(88, 0, KEYEVENTF_KEYUP, 0)
         '      'Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0) 'enters SKYx
         ' '  Next i%
         ' '  Call keybd_event(88, 0, KEYEVENTF_KEYUP, 0)
         ' '  Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0) 'enters SKYx
         '  'wait some time to lower height to ground, then jump back
         '   'waittime = Timer + 0.05
         '   'Do Until Timer > waittime
         '   '   DoEvents
         '   'Loop
         '   terwt = False
         '   End If
'         dx1 = 0 '22
'         dy1 = 75 '68
'         Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0) 'move mouse to Location item

       'goto to altitude edit box
        dx1 = 10
        dy1 = -10
        If WinVer = 5 Then
           dx1 = 12
           dy1 = -13
           End If
        
        dxo = dxo + dx1
        dyo = dyo + dy1
        Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
        'now wait to get to new location
        waitime = Timer
        Do Until Timer > waitime + 1
          DoEvents
        Loop
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'activate altitude edit box
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        Call keybd_event(VK_HOME, 0, 0, 0) 'go to beginning of numbers
        Call keybd_event(VK_HOME, 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_SHIFT, 0, 0, 0) 'select last entry for erasure
        Call keybd_event(VK_END, 0, 0, 0)
        Call keybd_event(VK_END, 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_DELETE, 0, 0, 0)
        Call keybd_event(VK_DELETE, 0, KEYEVENTF_KEYUP, 0)
        'Clipboard.Clear
        'Clipboard.SetText C3$
        'ret = CloseClipboard()
        Call keybd_event(VK_SHIFT, 0, 0, 0) 'enter minimum altitude
        Call keybd_event(VK_INSERT, 0, 0, 0)
        Call keybd_event(VK_INSERT, 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_SUBTRACT, 0, 0, 0) 'enter minimum possible height
        Call keybd_event(VK_SUBTRACT, 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(Asc("5"), 0, 0, 0)
        Call keybd_event(Asc("5"), 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(Asc("0"), 0, 0, 0) 'enter minimum height
        Call keybd_event(Asc("0"), 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(Asc("0"), 0, 0, 0) 'enter minimum height
        Call keybd_event(Asc("0"), 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_RETURN, 0, 0, 0) 'enters return
        Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)


       If skyleftjump = True Then
          'return pointer to center of map
          dx1 = -(mapwi2 / 60) - 40 - mapxdif / 30
          dy1 = -(maphi2 / 60) + 50 - mapydif / 30
          If WinVer = 5 Then
             dx1 = -67 '-dxo + 38
             dy1 = -38 '-dyo + 22
             End If
          Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
          End If
       End If
End Sub
Public Sub goto_click()
  On Error GoTo gotoerror
'        If Maps.text5.Text = sEmpty Or Maps.text6.Text = sEmpty Then Exit Sub
'        If Maps.text5.Text < 999 And Maps.text6.Text < 999 Then
'           kmxc = Maps.text5.Text * 1000
'           kmyc = Maps.text6.Text * 1000 + 1000000
'        Else
'           kmxc = Maps.text5.Text
'           kmyc = Maps.text6.Text
'           End If
'        Call blitpictures
'        If mapPictureform.Visible = True Then
'           ret = SetWindowPos(mapPictureform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
'           End If

        If world = True Then
           'check the coordinates
           If Abs(Val(Maps.Text5.Text)) > 180 Or Abs(Val(Maps.Text6.Text)) > 90 Then
              response = MsgBox("The inputed coordinates are not correct!", vbCritical + vbOKOnly, "Maps & More")
              Screen.MousePointer = vbDefault
              Exit Sub
              End If
           GoTo go10
           End If
        If Maps.Text5.Text <> sEmpty And Maps.Text6.Text <> sEmpty Then
        If (Maps.Text5.Text < 70000 Or Maps.Text5.Text > 240000 Or Maps.Text6.Text < 870000 Or Maps.Text6.Text > 1310000) And Maps.Label5.Caption = "ITMx" Then
           If (Maps.Text5.Text < 250 And Maps.Text5.Text > 80) And (Maps.Text6.Text < 300 And Maps.Text6.Text > -125) Then '>80 -- old limit
              Maps.Text5.Text = Maps.Text5.Text * 1000
              Maps.Text6.Text = 1000000 + Maps.Text6.Text * 1000
              GoTo go10
              End If
           lResult = FindWindow(vbNullString, terranam$)
           If terranam$ = sEmpty Then lResult = 0
           If lResult > 0 Then
'               ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               BringWindowToTop (lResult)
               End If
           response = MsgBox("The coordinates are not in the appropriate format, try again!", vbExclamation + vbOKOnly, "Maps&More")
'           ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           BringWindowToTop (lResult)
           Exit Sub
           End If
go10:   If Maps.Label5.Caption <> "ITMx" Or skymove = True Or worldmove = True Or jumpworld = True Then
           'convert to ITM before going there
           If Maps.Label5.Caption = "SKYx" Or Maps.Label5.Caption = "UTMx" Or Maps.Label5.Caption = "long." Or skymove = True Or worldmove = True Or jumpworld = True Then
              If Skycoord% <> 2 And world = False Then
                lResult = FindWindow(vbNullString, terranam$)
                If terranam$ = sEmpty Then lResult = 0
                'If lResult > 0 Then
                '    ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                '    End If
                 'response = MsgBox("You have not specified the GOTO (RED) coordinates in ITM. " + _
                 '                "SkyLight will have to " + _
                 '                "convert to ITM in order to move to the desired location. " + _
                 '                "Do you want the GOTO coordinates converted to ITM? " + _
                 '                "(not recommended if the ITM coordinates are already " + _
                 '                "known to the program due to a sllight loss of accuracy during the " + _
                 '                "inversion process). If ITM was imputed, then use the PgUp key to " + _
                 '                "to reconvert to ITM.", vbQuestion + vbYesNoCancel, "SkyLight")
                 response = vbYes
                 If lResult > 0 Then
'                    ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                    BringWindowToTop (lResult)
                    End If
              Else
                 Skycoord% = 0
                 response = vbYes 'jumped into this routine from GotoSky button
                 End If
              If response = vbYes Or world = True Then
                 If Maps.Label5.Caption = "SKYx" Or skymove = True Then
                    If skymove = False Then
                       T1 = Maps.Text5.Text
                       T2 = Maps.Text6.Text
                    ElseIf skymove = True Then
                       T1 = skyx
                       T2 = skyy
                       End If
                    If Maps.Label5.Caption = "SKYx" Or skymove = True Then
                       Mode% = 2 'inverse transform from SKY to ITM
                       Call ITMSKY(G11, G22, T1, T2, Mode%)
                       kmxc = G11: kmyc = G22
                       If mapsearchfm.Visible = True Then
                          kmxc = Val(T1)
                          kmyc = Val(T2)
                          End If
                       End If
                 ElseIf Maps.Label5.Caption = "UTMx" Then
                    G1 = Maps.Text5.Text
                    G2 = Maps.Text6.Text
                    'now convert UTM to GEO and then to Israel grid
                    
                    'first turn off gps correction if flagged
                    If ggpscorrection Then
                      GpsCorrOff = True
                      ggpscorrection = False
                      End If
                    Z = 33
                    Call UTMGEO(G1, G2, Z, l1, l2)
                    Call GEOCASC(l1, l2, kmxg, kmyg)
                    If GpsCorrOff Then
                       ggpscorrection = True
                       GpsCorrOff = False
                       End If

                    kmxc = Fix(0.5 + kmyg)
                    If kmxg < 870000 Then
                       kmyc = Fix(0.5 + kmxg) + 1000000
                    Else
                       kmyc = Fix(0.5 + kmxg)
                       End If
                 ElseIf Maps.Label5.Caption = "long." Or jumpworld = True Then
                    l1 = Maps.Text6.Text 'latitude
                    l2 = Maps.Text5.Text 'longitude
                    If worldmove = True Or jumpworld = True Then
                       l1 = lat
                       l2 = lon
                       End If
                    If world = False Or jumpworld = True Then
                       If l2 <= 0 Then l2 = -l2 'East longitude is positive for GEOCASC routine
                       Call GEOCASC(l1, l2, kmxg, kmyg)
                       kmxc = Fix(0.5 + kmyg)
                       If kmxg < 870000 Then
                          kmyc = Fix(0.5 + kmxg) + 1000000
                       Else
                          kmyc = Fix(0.5 + kmxg)
                          End If
                     Else
                        'determine height at that point if not inputed already
                        If noheights = False Then 'And Maps.Text5.Text = sEmpty Or Maps.Text5.Text = "0" Then
                           lg = l2: lt = l1
                           Call worldheights(lg, lt, hgt)
                           If hgt = -9999 Then hgt = 0
                           If worldmove = True Then
                              Maps.Text3.Text = hgt
                           Else
                              Maps.Text7.Text = hgt
                              End If
                           End If
                        If placdblclk = True Then placdblclk = False
                        If worldmove = False Then
                           lon = Maps.Text5.Text
                           lat = Maps.Text6.Text
                           Call blitpictures
                           Exit Sub
                           End If
                        End If
                    End If
                 End If
              End If
           If tblbuttons(18) = 0 And world = False And worldmove = False And jumpworld = False Then
              Maps.Label5.Caption = "ITMx"
              Maps.Label6.Caption = "ITMy"
              Maps.Text5.Text = kmxc
              Maps.Text6.Text = kmyc
              kmxsky = kmxc: kmysky = kmyc
           Else 'determine the coordinates at that point
                If worldmove = False Then
                   kmxcc = kmxc: kmycc = kmyc
                   End If
                Select Case coordmode%
                  Case 1 'ITM
                     Maps.Text1.Text = kmxcc
                     Maps.Text2.Text = kmycc
                  Case 2 'GEO
                     If worldmove = False Then
                        Call casgeo(kmxcc, kmycc, lg, lt)
                     Else
                        lg = lon
                        lt = lat
                        End If
                     lgdeg = Fix(lg)
                     lgmin = Abs(Fix((lg - Fix(lg)) * 60))
                     lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
                     ltdeg = Fix(lt)
                     ltmin = Abs(Fix((lt - Fix(lt)) * 60))
                     ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
                     If ltdeg = 0 And lt < 0 Then
                        Maps.Text2.Text = "-" + Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
                     Else
                        Maps.Text2.Text = Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
                        End If
                     If lgdeg = 0 And lg < 0 Then
                        Maps.Text1.Text = "-" + Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
                     Else
                        Maps.Text1.Text = Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
                        End If
                  Case 3 'UTM
                     Call casgeo(kmxcc, kmycc, lg, lt)
                     ZZ% = 0
                     Call GEOUTM(lt, lg, ZZ%, G1, G2)
                     Maps.Text1.Text = Fix(G1)
                     Maps.Text2.Text = Fix(G2)
                  Case 4 'SKYLINE UTM
                     Mode% = 1
                     Call ITMSKY(kmxcc, kmycc, T1, T2, Mode%)
                     Maps.Text1.Text = T1
                     Maps.Text2.Text = T2
                  Case 5 'distance, viewangle, azimuth
                     kmxoo = kmxcc: kmyoo = kmycc
                    If world = True Then
                       If mag > 1 Then
                         lonc = lon '+ fudx / mag
                         latc = lat '+ fudy / mag
                         'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                         'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                         'lono = xo + Xcoord * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
                         'lato = yo - Ycoord * (180# / (sizewy * mag))
                         xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                         yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                         lono = xo + Xcoord * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
                         lato = yo - Ycoord * (deglat / (sizewy * mag))
                       Else
                         'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
                         'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
                         'xo = lonc - 90#
                         'yo = latc + 90#
                         'lono = xo + Xcoord * (180 / sizewx)
                         'lato = yo - Ycoord * (180 / sizewy)
                         lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                         latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                         xo = lonc - deglog / 2
                         yo = latc + deglat / 2
                         lono = xo + Xcoord * (deglog / sizewx)
                         lato = yo - Ycoord * (deglat / sizewy)
                         If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                            'fudge factor for inaccuracy of linear degree approx for large size map
                            lono = lono - 0.006906
                            lato = lato + 0.003878
                            End If
                          End If
                      End If
                     Call dipcoord
                  End Select
'             Maps.Label1.Caption = "ITMx"
'             Maps.Label2.Caption = "ITMy"
'             Maps.Text1.Text = kmxc
'             Maps.Text2.Text = kmyc
              If noheights = False And world = False And jumpworld = False Then 'determine height at that point
                 kmxo = kmxc: kmyo = kmyc
                 Call heights(kmxo, kmyo, hgt)
              ElseIf noheights = False And world = True Then
                 lg = lon: lt = lat
                 Call worldheights(lg, lt, hgt)
                 If hgt = -9999 Then hgt = 0
              ElseIf noheights = True Then
                 hgt = 0#
                 End If
              'Maps.Text7.Text = hgt
              Maps.Text3.Text = hgt
              End If
           End If
        End If
        If Maps.Label5.Caption = "ITMx" And skymove = False And jumpworld = False Then
           kmxc = Maps.Text5.Text
           kmyc = Maps.Text6.Text
           kmxsky = kmxc: kmysky = kmyc
           If noheights = False Then 'determine height at that point
              kmxo = kmxc: kmyo = kmyc
              Call heights(kmxo, kmyo, hgt)
           ElseIf noheights = True Then
              hgt = 0#
              End If
           Maps.Text7.Text = hgt
           End If
        If world = False Then
            kmxoo = kmxc
            kmyoo = kmyc
            kmx50c = kmxc
            kmy50c = kmyc
            hgt50c = Maps.Text7.Text
            kmx400c = kmxc
            kmy400c = kmyc
            hgt400c = hgt50c
            End If
        'If worldmove = True Then
        '   lResult = FindWindow(vbNullString, "Extracting relevant portion of the DTM")
        '   If lResult > 0 Then Exit Sub
        '   End If
        Call blitpictures
        If mapPictureform.Visible = True And worldmove = False Then
'           ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
           BringWindowToTop (mapPictureform.hwnd)
           Exit Sub
           End If
        lResult = FindWindow(vbNullString, terranam$)
        If tblbuttons(18) = 0 And lResult > 0 And terranam$ <> "" Then 'also move terraviewer
           If Maps.Label5.Caption <> "long." And Maps.Text5.Text <> sEmpty Then
             Call skyTERRAgoto
             'first move pointer to TerraViewer window in order to activate it
             'dx1 = 0 '30 '-30
             'dy1 = 240 ' -60
             'Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item

             If gotobutton = True Then

             '***old program
'               'wait to get there, then go to minimum elevation
'                waitime = Timer
'                Do Until Timer > waitime + 1
'                   DoEvents
'                Loop
'                Call BringWindowToTop(lResult)
'                dx1 = 0 '30 '-30
'                dy1 = 240 ' -60
'                Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item
'                For i% = 1 To 30 'hold down the keys for a bit
'                   Call keybd_event(VK_SHIFT, 0, 0, 0)
'                   Call keybd_event(Asc("X"), 0, 0, 0)  'goes into Settings menu
'                Next i%
'                Call keybd_event(Asc("X"), 0, KEYEVENTF_KEYUP, 0)
'                Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)

                'move pointer back to goto button
                dx1 = -10 '-30 '30
                dy1 = -212 '60
                Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
                End If
             End If
           End If
     Exit Sub
gotoerror:
      lResult = FindWindow(vbNullString, terranam$)
      If lResult > 0 Then
          ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
      response = MsgBox("The coordinates are not in the appropriate format, ABORT!", vbCritical + vbOKOnly, "Maps & More")
'      ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      BringWindowToTop (lResult)
      Exit Sub
End Sub
Public Sub showcoord()
     If world = False Then
        kmxo = kmxc: kmyo = kmyc
        If noheights = False Then
           Call heights(kmxo, kmyo, hgt)
        ElseIf noheights = True Then
           hgt = 0#
           End If
        Maps.Text3.Text = Str$(hgt)
        End If
     Select Case coordmode%
       Case 1 'ITM
          Maps.Text1.Text = Format(kmxc, "#####0")
          Maps.Text2.Text = Format(kmyc, "######0")
       Case 2 'GEO
          If world = True Then
            lg = lon
            lt = lat
            If noheights = False Then
               Call worldheights(lg, lt, hgt)
               If hgt = -9999 Then hgt = 0
               Maps.Text3.Text = Str$(hgt)
               End If
          Else
             kmxo = kmxc: kmyo = kmyc
             Call casgeo(kmxo, kmyo, lg, lt)
             End If
          lgdeg = Fix(lg)
          lgmin = Abs(Fix((lg - Fix(lg)) * 60))
          lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
          ltdeg = Fix(lt)
          ltmin = Abs(Fix((lt - Fix(lt)) * 60))
          ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
          If ltdeg = 0 And lt < 0 Then
             Maps.Text2.Text = "-" + Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
          Else
             Maps.Text2.Text = Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
             End If
          If lgdeg = 0 And lg < 0 Then
             Maps.Text1.Text = "-" + Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
          Else
             Maps.Text1.Text = Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
             End If
       Case 3 'UTM
          kmxo = kmxc: kmyo = kmyc
          Call casgeo(kmxo, kmyo, lg, lt)
          Call GEOUTM(lt, lg, Z%, G1, G2)
          Maps.Text1.Text = Fix(G1)
          Maps.Text2.Text = Fix(G2)
       Case 4 'SKYLINE UTM
          Mode% = 1
          kmxo = kmxc: kmyo = kmyc
          Call ITMSKY(kmxo, kmyo, T1, T2, Mode%)
          Maps.Text1.Text = T1
          Maps.Text2.Text = T2
       Case 5 'distance, viewangle, azimuth
          kmxoo = kmxc: kmyoo = kmyc
          'kmxcd = Maps.Text5.Text
          'kmycd = Maps.Text6.Text
          If world = True Then
             If mag > 1 Then
               lonc = lon '+ fudx / mag
               latc = lat '+ fudy / mag
               'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               'lono = xo + Xcoord * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
               'lato = yo - Ycoord * (180# / (sizewy * mag))
               xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               lono = xo + Xcoord * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
               lato = yo - Ycoord * (deglat / (sizewy * mag))
             Else
               'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
               'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
               'xo = lonc - 90#
               'yo = latc + 90#
               'lono = xo + Xcoord * (180 / sizewx)
               'lato = yo - Ycoord * (180 / sizewy)
               lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
               latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
               xo = lonc - deglog / 2
               yo = latc + deglat / 2
               lono = xo + Xcoord * (deglog / sizewx)
               lato = yo - Ycoord * (deglat / sizewy)
               If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                  'fudge factor for inaccuracy of linear degree approx for large size map
                  lono = lono - 0.006906
                  lato = lato + 0.003878
                  End If
               End If
            End If
          Call dipcoord
       End Select
End Sub
Public Sub obstructions(Dest)
  'draw in obstructions on current map sheet
  If printing = False Then
     mult% = 4 'mag '*****allows to printscreen until debug
     If mag * 1.5 > mult% Then mult% = mag * 1.5 '***allows to printscreen until debug
     colr% = 0 '255
  Else
     mult% = 11
     colr% = 0
     End If
'  If mag > 1 Then
'    kmxcc = kmxc
'    kmycc = kmyc
'  Else

    If map50 = True Then
       kmxcc = kmxc + (km50x / mag) * (mapwi - mapwi2 + mapxdif) / 2
       kmycc = kmyc - (km50y / mag) * (maphi - maphi2 + mapydif) / 2
    ElseIf map400 Then
       kmxcc = kmxc + (km400x / mag) * (mapwi - mapwi2 + mapxdif) / 2
       kmycc = kmyc - (km400y / mag) * (maphi - maphi2 + mapydif) / 2
       End If
'    End If

  'Line Input #obsfilnum%, doclin$
  'Line Input #obsfilnum%, doclin$ 'Input #obsfilnum%, kmxob, kmyob, hgtob, Aob, Bob, Cob, Dob, Eob
  If map400 = True And world = False Then
     kmx400orig = kmxcc - (km400x / mag) * sizex / 2 '80000 + (n400x% - 1) * 80000 + dxx% * 40000
     kmy400orig = kmycc - (km400y / mag) * sizey / 2 '840000 + (n400y% - 1) * 80000 + dyy% * 40000
     X400c = ((kmxobs - kmx400orig) / (km400x / mag)) '* (Dest.Width / sizex)  '(mapwi2 - mapxdif) / 2
     Y400c = sizey - ((kmyobs - kmy400orig) / (km400y / mag)) '* (Dest.Height / sizey)  '(maphi2 - mapydif) / 2
     Dest.DrawMode = 13
     Dest.DrawWidth = 2
     Dest.Circle (X400c, Y400c), 100, colr% '100 * mag, colr%
     Dest.DrawWidth = 2 * mult%
     Dest.Circle (X400c, Y400c), 20, colr% '20 * mag, colr%
     Dest.DrawWidth = mult%
'     Do Until EOF(obsfilnum%)
'        Input #obsfilnum%, aziob, vaob, kmxob, kmyob, c, D
'        Xpnt = (((kmxob * 1000 - kmx400orig) / (km400x / mag))) '* (Dest.Width / sizex)
'        Ypnt = sizey - ((kmyob * 1000 + 1000000) - kmy400orig) / (km400y / mag) '* (Dest.Height / sizey)
     'ReDim obs(2, obsnum%)
     For i% = 1 To obsnum%
        kmxob = obs(1, i%)
        kmyob = obs(2, i%)
        Xpnt = (((kmxob - kmx400orig) / (km400x / mag)))  '* (Dest.Width / sizex)
        Ypnt = sizey - (kmyob - kmy400orig) / (km400y / mag) '* (Dest.Height / sizey)
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr%
           xpno = Xpnt: ypno = Ypnt
        Else
           flg% = 0
           If printing = True Then
              If (Xpnt < 0 And Xpnt = xpno) Or (Xpnt < 0 And xpno < 0) Or _
                 (Xpnt < 0 And xpno > mapPictureform.mapPicture.Width) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno < 0) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno > mapPictureform.mapPicture.Width) Then
                 xpno = Xpnt: ypno = Ypnt
                 GoTo o40
              ElseIf Xpnt < 0 And xpno < mapPictureform.mapPicture.Width And xpno >= 0 And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (-xpno) / (Xpnt - xpno) + ypno
                 Xpnt = 0
              ElseIf Xpnt > mapPictureform.mapPicture.Width And xpno < mapPictureform.mapPicture.Width And xpno >= 0 And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 Xpnt = Printer.Width
              ElseIf Xpnt <= mapPictureform.mapPicture.Width And Xpnt >= 0 And xpno > mapPictureform.mapPicture.Width And Xpnt <> xpno Then
                 flg% = 1
                 ypno = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 xpno = Printer.Width
                 xp1 = Xpnt
                 yp1 = Ypnt
                 End If
                 Ypnt = Ypnt - printeroffset * mag
                 Xpnt = Xpnt - printeroffset * mag
              End If
            If ((xpno > sizex And Xpnt > sizex) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizey And Xpnt > sizey) Or (ypno < 0 And Ypnt < 0)) Then GoTo o35
            Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr%
o35:        xpno = Xpnt: ypno = Ypnt
            If flg% = 1 Then
               xpno = xp1: ypno = yp1
               flg% = 0
               End If
            End If
o40: Next i%
'     Seek #obsfilnum%, 1 'rewind the file to the beginning
   ElseIf map50 = True And world = False Then
     kmx50orig = kmxcc - (km50x / mag) * sizex / 2 '80000 + (n400x% - 1) * 80000 + dxx% * 40000
     kmy50orig = kmycc - (km50y / mag) * sizey / 2 '840000 + (n400y% - 1) * 80000 + dyy% * 40000
     X50c = (kmxobs - kmx50orig) / (km50x / mag) '* (Dest.Width / sizex)  '(mapwi2 - mapxdif) / 2
     Y50c = sizey - ((kmyobs - kmy50orig) / (km50y / mag))
     Dest.DrawMode = 13
     Dest.DrawWidth = mult%
     Dest.Circle (X50c, Y50c), 100, colr% '100 * mag, colr%
     Dest.DrawWidth = 2 * mult%
     Dest.Circle (X50c, Y50c), 20, colr% '20 * mag, colr%
     Dest.DrawWidth = mult%
'     Do Until EOF(obsfilnum%)
'        Input #obsfilnum%, aziob, vaob, kmxob, kmyob, c, D
'        Xpnt = (((kmxob * 1000 - kmx50orig) / (km50x / mag))) '* (Dest.Width / sizex)
'        Ypnt = sizey - ((kmyob * 1000 + 1000000) - kmy50orig) / (km50y / mag) ' * (Dest.Height / siezey)
     'ReDim obs(2, obsnum%)
     For i% = 1 To obsnum%
        kmxob = obs(1, i%)
        kmyob = obs(2, i%)
        Xpnt = (((kmxob - kmx50orig) / (km50x / mag)))  '* (Dest.Width / sizex)
        Ypnt = sizey - (kmyob - kmy50orig) / (km50y / mag) ' * (Dest.Height / siezey)
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr%
           xpno = Xpnt: ypno = Ypnt
        Else
           If printing = True Then
              If (Xpnt < 0 And Xpnt = xpno) Or (Xpnt < 0 And xpno < 0) Or _
                 (Xpnt < 0 And xpno > mapPictureform.mapPicture.Width) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno < 0) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno > mapPictureform.mapPicture.Width) Then
                 xpno = Xpnt: ypno = Ypnt
                 GoTo o50
              ElseIf Xpnt < 0 And xpno < mapPictureform.mapPicture.Width And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (-xpno) / (Xpnt - xpno) + ypno
                 Xpnt = 0
              ElseIf Xpnt > mapPictureform.mapPicture.Width And xpno < mapPictureform.mapPicture.Width And xpno >= 0 And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 Xpnt = Printer.Width
              ElseIf Xpnt <= mapPictureform.mapPicture.Width And Xpnt >= 0 And xpno > mapPictureform.mapPicture.Width And Xpnt <> xpno Then
                 flg% = 1
                 ypno = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 xpno = Printer.Width
                 xp1 = Xpnt
                 yp1 = Ypnt
                 End If
                 Ypnt = Ypnt - printeroffset * mag
                 Xpnt = Xpnt - printeroffset * mag
              End If
            If ((xpno > sizex And Xpnt > sizex) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizey And Xpnt > sizey) Or (ypno < 0 And Ypnt < 0)) Then GoTo o45
           Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr%
o45:       xpno = Xpnt: ypno = Ypnt
           If flg% = 1 Then
              xpno = xp1: ypno = yp1
              flg% = 0
              End If
           End If
o50: Next i%
'     Seek #obsfilnum%, 1 'rewind the file to the beginning
   ElseIf world = True Then
   
       If mag > 1 Then
         lonc = lon '+ fudx / mag
         latc = lat '+ fudy / mag
         'wxorigin = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
         'wyorigin = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
         wxorigin = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
         wyorigin = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
      Else
         'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
         'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
         'wxorigin = lonc - 90#
         'wyorigin = latc + 90#
         lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
         latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
         wxorigin = lonc - deglog / 2
         wyorigin = latc + deglat / 2
         End If

     Dest.DrawWidth = mult%
     Dest.DrawMode = 13
     X50c = (lonobs - wxorigin) / (deglog / (sizewx * mag))
     Y50c = (wyorigin - latobs) / (deglat / (sizewy * mag))
     Dest.Circle (X50c, Y50c), 100, colr% '100 * mag, colr%
     Dest.DrawWidth = 2 * mult%
     Dest.Circle (X50c, Y50c), 20, colr% '20 * mag, colr%
     Dest.DrawWidth = mult%
     
     For i% = 1 To obsnum%
        T1 = obs(1, i%)
        T2 = obs(2, i%)
         If mag > 1 Then
           'Xpnt = (T1 - wxorigin) / (180# / (sizewx * mag))
           'Ypnt = (wyorigin - T2) / (180# / (sizewy * mag))
           Xpnt = (T1 - wxorigin) / (deglog / (sizewx * mag))
           Ypnt = (wyorigin - T2) / (deglat / (sizewy * mag))
        Else
           'Xpnt = (T1 - wxorigin) / (180# / sizewx)
           'Ypnt = (wyorigin - T2) / (180# / sizewy)
           Xpnt = (T1 - wxorigin) / (deglog / sizewx)
           Ypnt = (wyorigin - T2) / (deglat / sizewy)
           End If
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr
           xpno = Xpnt: ypno = Ypnt
        Else
            If ((xpno > sizewx And Xpnt > sizewx) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizewy And Xpnt > sizewy) Or (ypno < 0 And Ypnt < 0)) Then GoTo o955
           Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr
o955:      xpno = Xpnt: ypno = Ypnt
           If flg% = 1 Then
              xpno = xp1: ypno = yp1
              flg% = 0
              End If
           End If
     Next i%
     End If
End Sub
Public Sub routeform()
  Dim bRtn As Boolean, lResult As Long, lMain As Long, lOpen As Long
  If world = True Then '<<<<<<<TO DO>>>>>>>>
    insiderouteform = True
    routeload = True
    lResult = FindWindow(vbNullString, "3D Viewer")
    If lResult <> 0 Then
       nmsg = SendMessage(lResult, WM_COMMAND, 1002, 0)
       insiderouteform = False
       Exit Sub
    Else
       routeload = False
       insiderouteform = False
       Exit Sub
       End If
    End If
  'check that terraviewer is activated,
  'lMain = FindWindow(vbNullString, Maps.Caption)
  lResult = FindWindow(vbNullString, terranam$)
  If lResult > 0 Then
      insiderouteform = True
r10:  Maps.Timer2.Enabled = False
      'tblbuttons(18) = 0
      'Maps.Toolbar1.Buttons(18).Value = tbrUnpressed
      Call BringWindowToTop(lResult)
      'lMain = FindWindow(vbNullString, Maps.Caption)
      'make it the top window
      'open the Route editor dialog box
      waitime = Timer
      Do Until Timer > waitime + 3 '0.1 <--changed
         DoEvents
      Loop
      Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
      Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
      'Call keybd_event(Asc("T"), 0, 0, 0)  'goes into tools menu
      'Call keybd_event(Asc("T"), 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(Asc("R"), 0, 0, 0)    'calls for the Route editor
      Call keybd_event(Asc("R"), 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(Asc("L"), 0, 0, 0)  'goes into tools menu
      Call keybd_event(Asc("L"), 0, KEYEVENTF_KEYUP, 0)
      'Call keybd_event(VK_DOWN, 0, 0, 0)  'calls the open dialog box
      'Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
      'Call keybd_event(VK_RETURN, 0, 0, 0)
      'Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
'
      routeload = True
      'Call BringWindowToTop(lMain)
      waitime = Timer
      Do Until Timer > waitime + 0.1 '0.1 <---changed
         DoEvents
      Loop
      'Maps.Timer2.Enabled = False
      nloops% = 0
      'tblbuttons(18) = 0
      'Maps.Toolbar1.Buttons(18).Value = tbrUnpressed
r20:  lOpen = FindWindow(vbNullString, "Open")
      Do Until lOpen <> 0
         'Maps.Timer2.Enabled = False
         'tblbuttons(18) = 0
         'Maps.Toolbar1.Buttons(18).Value = tbrUnpressed
         DoEvents
         lOpen = FindWindow(vbNullString, "Open")
         nloops% = nloops% + 1
         If nloops% > 100 Then GoTo r10
      Loop

r50:  If lOpen = 0 Then GoTo r20
      Call BringWindowToTop(lOpen)
      Clipboard.Clear
      C1$ = terradir$ + "\*.trf"
      Clipboard.SetText C1$
      Call keybd_event(VK_SHIFT, 0, 0, 0)
      Call keybd_event(VK_INSERT, 0, 0, 0)
      Call keybd_event(VK_INSERT, 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(VK_RETURN, 0, 0, 0)
      Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)

      'lOpen = FindWindow(vbNullString, "Open")
      'Maps.Text7.Text = lOpen
      'Call BringWindowToTop(lOpen)
      waitime = Timer
      Do Until Timer > waitime + 0.1 '0.1 <--changed
         DoEvents
      Loop
      Call keybd_event(VK_TAB, 0, 0, 0)
      Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(VK_TAB, 0, 0, 0)
      Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(VK_TAB, 0, 0, 0)
      Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(VK_TAB, 0, 0, 0)
      Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(VK_TAB, 0, 0, 0)
      Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
      '   waitime = Timer
      '   Do Until Timer > waitime + 0.05 '0.05 <--changed
      '      DoEvents
      '   Loop
      'If Maps.Timer2.Interval < traveltimerinteral Then Maps.Timer2.Interval = traveltimerinteral
      Maps.Toolbar1.Buttons(18).value = tbrPressed
      tblbuttons(18) = 1
      'routnum% = routnum% + 1
      'activate first entry in open dialog box
      Call keybd_event(VK_DOWN, 0, 0, 0)
      Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(VK_UP, 0, 0, 0)
      Call keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)
      For i% = 1 To routnum% 'load next file <----check here!!!!!
         Call keybd_event(VK_DOWN, 0, 0, 0)
         Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
         waitime = Timer
         Do Until Timer > waitime + 0.05 '0.05 <--changed
            DoEvents
         Loop
      Next i%
      routnum% = routnum% + 1
      'Call keybd_event(VK_LMENU, 0, 0, 0)
      'Call keybd_event(Asc("O"), 0, 0, 0)
      'Call keybd_event(Asc("O"), 0, KEYEVENTF_KEYUP, 0)
      'Call keybd_event(VK_LMENU, 0, KEYEVENTF_KEYUP, 0)

      'Call keybd_event(VK_RETURN, 0, 0, 0)
      'Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)

      'check if open window is still there due to timming bug
      'tblbuttons(18) = 0
      'Maps.Toolbar1.Buttons(18).Value = tbrUnpressed
      timechecked% = 0
r100: waitime = Timer
      Do Until Timer > waitime + 0.1 ' 0.1 <-------changed
         DoEvents
      Loop
      lOpen = FindWindow(vbNullString, "Open")
      nloops% = 0
      If lOpen <> 0 Then
         'check one more time
         timechecked% = timechecked% + 1
         'Maps.Timer2.Enabled = False
         'tblbuttons(18) = 0
         'Maps.Toolbar1.Buttons(18).Value = tbrUnpressed
         If timechecked% = 1 Then
            GoTo r100
         Else
            GoTo r50
            End If
         End If
     ''ground mode
      ' ret = BringWindowToTop(lResult)
      'Call keybd_event(VK_F8, 0, 0, 0)
      'Call keybd_event(VK_F8, 0, KEYEVENTF_KEYUP, 0)

      'now begin running the playback
      ret = BringWindowToTop(lResult)
      waitime = Timer
      Do Until Timer > waitime + 0.01 ' 0.01 <---changed
         DoEvents
      Loop
      Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
      Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
      'Call keybd_event(Asc("T"), 0, 0, 0)  'goes into tools menu
      'Call keybd_event(Asc("T"), 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(Asc("R"), 0, 0, 0)    'calls for the Route editor
      Call keybd_event(Asc("R"), 0, KEYEVENTF_KEYUP, 0)
      Call keybd_event(Asc("P"), 0, 0, 0)    'begins playback of loaded route
      Call keybd_event(Asc("P"), 0, KEYEVENTF_KEYUP, 0)
      Maps.Toolbar1.Buttons(17).Enabled = False
      'Call keybd_event(VK_DOWN, 0, 0, 0)  'calls the open dialog box
      'Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
      'Call keybd_event(VK_RETURN, 0, 0, 0)
      'Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)

      'now use a callback routine to find the desired file name
      'and then input a return

      'routename$ = "trial4.trf"
      'lOpen = FindWindow(vbNullString, "Open")
      'bRtn = EnumChildWindows(lOpen, AddressOf EnumFunc, 0) 'read captions
      'routeload = False
      'Call keybd_event(VK_RETURN, 0, 0, 0)
      'Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
      'dx1 = 30 '-30
      'dy1 = 240 ' -60
      'Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item

      End If
      insiderouteform = False
  'and input keyboard functions
End Sub
Public Sub showtheroute(Dest)
  'draw in route on current map sheet
  If printing = False Or showroute = True Then
     mult% = mag '1
     If world = True Then mult% = 1
     colr = 255
     If showroute = True Then colr = QBColor(13) '8454143
  Else
     mult% = 11
     colr = 0
     End If
''  If mag > 1 Then
''    kmxcc = kmxc
''    kmycc = kmyc
''  Else
'    If map50 = True Then
'       kmxcc = kmxc + (km50x / mag) * (mapwi - mapwi2 + mapxdif) / 2
'       kmycc = kmyc - (km50y / mag) * (maphi - maphi2 + mapydif) / 2
'    ElseIf map400 Then
'       kmxcc = kmxc + (km400x / mag) * (mapwi - mapwi2 + mapxdif) / 2
'       kmycc = kmyc - (km400y / mag) * (maphi - maphi2 + mapydif) / 2
'       End If
''    End If

    If crosssection = True Then
       If world = False Then
            obsnum% = 2
            ReDim obs(2, obsnum%)
            obs(1, 1) = crosssectionpnt(0, 0)
            obs(2, 1) = crosssectionpnt(0, 1) '- 1000000
            obs(1, 2) = crosssectionpnt(1, 0)
            obs(2, 2) = crosssectionpnt(1, 1) '- 1000000
       Else
            If greatcircle = False Then
                'then show straight line on world map
                'if greatcircle=True then already loaded travel points
                'in routine MapCrossSection
                travelnum% = 2
                ReDim travel(2, travelnum%)
                travel(1, 1) = crosssectionpnt(0, 0)
                travel(2, 1) = crosssectionpnt(0, 1)
                travel(1, 2) = crosssectionpnt(1, 0)
                travel(2, 2) = crosssectionpnt(1, 1)
                End If
            End If
       End If


    If world = False Then
        If map50 = True Then
           kmxcc = kmxc + (km50x / mag) * (mapwi - mapwi2 + mapxdif) / 2
           kmycc = kmyc - (km50y / mag) * (maphi - maphi2 + mapydif) / 2
        ElseIf map400 Then
           kmxcc = kmxc + (km400x / mag) * (mapwi - mapwi2 + mapxdif) / 2
           kmycc = kmyc - (km400y / mag) * (maphi - maphi2 + mapydif) / 2
           End If
        End If

  If travelmode = True Or world = True Then GoTo o900

  'For i% = 1 To 9   'skip headers
  '   Line Input #openfilnum%, doclin$
  'Next i%
  If map400 = True Then
     kmx400orig = kmxcc - (km400x / mag) * sizex / 2 '80000 + (n400x% - 1) * 80000 + dxx% * 40000
     kmy400orig = kmycc - (km400y / mag) * sizey / 2 '840000 + (n400y% - 1) * 80000 + dyy% * 40000
     'X400c = ((kmxobs - kmx400orig) / (km400x / mag)) '* (Dest.Width / sizex)  '(mapwi2 - mapxdif) / 2
     'Y400c = sizey - ((kmyobs - kmy400orig) / (km400y / mag)) '* (Dest.Height / sizey)  '(maphi2 - mapydif) / 2
     'Dest.DrawMode = 13
     'Dest.DrawWidth = mult%
     'Dest.Circle (X400c, Y400c), 100, colr
     'Dest.DrawWidth = 2 * mult%
     'Dest.Circle (X400c, Y400c), 20, colr
     Dest.DrawWidth = mult%
'     Do Until EOF(openfilnum%)
'        Line Input #openfilnum%, doclin$
'        skyxposit% = InStr(1, doclin$, " = ") + 3
'        skyyposit% = InStr(skyxposit%, doclin$, " ")
'        positend% = InStr(skyyposit% + 1, doclin$, " ")
'        T1 = Val(Mid$(doclin$, skyxposit%, skyyposit% - skyxposit%))
'        T2 = Val(Mid$(doclin$, skyyposit% + 1, positend% - skyyposit% - 1))
'        mode% = 2 'inverse transform from SKY to ITM
'        Call ITMSKY(G11, G22, T1, T2, mode%)
'        kmxob = G11: kmyob = G22

     For i% = 1 To obsnum%
        kmxob = obs(1, i%)
        kmyob = obs(2, i%)
        Xpnt = (kmxob - kmx400orig) / (km400x / mag) '* (Dest.Width / sizex)
        Ypnt = sizey - (kmyob - kmy400orig) / (km400y / mag) '* (Dest.Height / sizey)
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr
           xpno = Xpnt: ypno = Ypnt
        Else
           flg% = 0
           If printing = True Then
              If (Xpnt < 0 And Xpnt = xpno) Or (Xpnt < 0 And xpno < 0) Or _
                 (Xpnt < 0 And xpno > mapPictureform.mapPicture.Width) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno < 0) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno > mapPictureform.mapPicture.Width) Then
                 xpno = Xpnt: ypno = Ypnt
                 GoTo o40
              ElseIf Xpnt < 0 And xpno < mapPictureform.mapPicture.Width And xpno >= 0 And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (-xpno) / (Xpnt - xpno) + ypno
                 Xpnt = 0
              ElseIf Xpnt > mapPictureform.mapPicture.Width And xpno < mapPictureform.mapPicture.Width And xpno >= 0 And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 Xpnt = Printer.Width
              ElseIf Xpnt <= mapPictureform.mapPicture.Width And Xpnt >= 0 And xpno > mapPictureform.mapPicture.Width And Xpnt <> xpno Then
                 flg% = 1
                 ypno = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 xpno = Printer.Width
                 xp1 = Xpnt
                 yp1 = Ypnt
                 End If
                 Ypnt = Ypnt - printeroffset * mag
                 Xpnt = Xpnt - printeroffset * mag
              End If
            If ((xpno > sizex And Xpnt > sizex) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizey And Xpnt > sizey) Or (ypno < 0 And Ypnt < 0)) Then GoTo o35
            Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr
o35:        xpno = Xpnt: ypno = Ypnt
            If flg% = 1 Then
               xpno = xp1: ypno = yp1
               flg% = 0
               End If
            End If
o40: Next i%
     'Seek #openfilnum%, 1 'rewind route file
   ElseIf map50 = True Then
     kmx50orig = kmxcc - (km50x / mag) * sizex / 2 '80000 + (n400x% - 1) * 80000 + dxx% * 40000
     kmy50orig = kmycc - (km50y / mag) * sizey / 2 '840000 + (n400y% - 1) * 80000 + dyy% * 40000
     'X50c = (kmxobs - kmx50orig) / (km50x / mag) '* (Dest.Width / sizex)  '(mapwi2 - mapxdif) / 2
     'Y50c = sizey - ((kmyobs - kmy50orig) / (km50y / mag))
     'Dest.DrawMode = 13
     'Dest.DrawWidth = mult%
     'Dest.Circle (X50c, Y50c), 100, colr
     'Dest.DrawWidth = 2 * mult%
     'Dest.Circle (X50c, Y50c), 20, colr
     Dest.DrawWidth = mult%
'     Do Until EOF(openfilnum%)
'        Line Input #openfilnum%, doclin$
'        skyxposit% = InStr(1, doclin$, " = ") + 3
'        skyyposit% = InStr(skyxposit%, doclin$, " ")
'        positend% = InStr(skyyposit% + 1, doclin$, " ")
'        T1 = Val(Mid$(doclin$, skyxposit%, skyyposit% - skyxposit%))
'        T2 = Val(Mid$(doclin$, skyyposit% + 1, positend% - skyyposit% - 1))
'        mode% = 2 'inverse transform from SKY to ITM
'        Call ITMSKY(G11, G22, T1, T2, mode%)
'        kmxob = G11: kmyob = G22
     For i% = 1 To obsnum%
        kmxob = obs(1, i%)
        kmyob = obs(2, i%)
        Xpnt = (kmxob - kmx50orig) / (km50x / mag) '* (Dest.Width / sizex)
        Ypnt = sizey - (kmyob - kmy50orig) / (km50y / mag) ' * (Dest.Height / siezey)
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr
           xpno = Xpnt: ypno = Ypnt
        Else
           If printing = True Then
              If (Xpnt < 0 And Xpnt = xpno) Or (Xpnt < 0 And xpno < 0) Or _
                 (Xpnt < 0 And xpno > mapPictureform.mapPicture.Width) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno < 0) Or _
                 (Xpnt > mapPictureform.mapPicture.Width And xpno > mapPictureform.mapPicture.Width) Then
                 xpno = Xpnt: ypno = Ypnt
                 GoTo o50
              ElseIf Xpnt < 0 And xpno < mapPictureform.mapPicture.Width And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (-xpno) / (Xpnt - xpno) + ypno
                 Xpnt = 0
              ElseIf Xpnt > mapPictureform.mapPicture.Width And xpno < mapPictureform.mapPicture.Width And xpno >= 0 And Xpnt <> xpno Then
                 flg% = 1
                 xp1 = Xpnt: yp1 = Ypnt
                 Ypnt = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 Xpnt = Printer.Width
              ElseIf Xpnt <= mapPictureform.mapPicture.Width And Xpnt >= 0 And xpno > mapPictureform.mapPicture.Width And Xpnt <> xpno Then
                 flg% = 1
                 ypno = (Ypnt - ypno) * (Printer.Width - xpno) / (Xpnt - xpno) + ypno
                 xpno = Printer.Width
                 xp1 = Xpnt
                 yp1 = Ypnt
                 End If
                 Ypnt = Ypnt - printeroffset * mag
                 Xpnt = Xpnt - printeroffset * mag
              End If
            If ((xpno > sizex And Xpnt > sizex) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizey And Xpnt > sizey) Or (ypno < 0 And Ypnt < 0)) Then GoTo o45
           Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr
o45:       xpno = Xpnt: ypno = Ypnt
           If flg% = 1 Then
              xpno = xp1: ypno = yp1
              flg% = 0
              End If
           End If
o50: Next i%
'     Seek #openfilnum%, 1 'rewind route file
     End If
  Exit Sub

o900:
  If map400 = True And world = False Then
     kmx400orig = kmxcc - (km400x / mag) * sizex / 2
     kmy400orig = kmycc - (km400y / mag) * sizey / 2
     Dest.DrawWidth = mult%
     For i% = 1 To travelnum%
        T1 = travel(1, i%)
        T2 = travel(2, i%)
        Mode% = 2 'inverse transform from SKY to ITM
        Call ITMSKY(G11, G22, T1, T2, Mode%)
        kmxob = G11: kmyob = G22
        Xpnt = (kmxob - kmx400orig) / (km400x / mag)
        Ypnt = sizey - (kmyob - kmy400orig) / (km400y / mag)
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr
           xpno = Xpnt: ypno = Ypnt
        Else
           flg% = 0
            If ((xpno > sizex And Xpnt > sizex) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizey And Xpnt > sizey) Or (ypno < 0 And Ypnt < 0)) Then GoTo o935
            Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr
o935:       xpno = Xpnt: ypno = Ypnt
            If flg% = 1 Then
               xpno = xp1: ypno = yp1
               flg% = 0
               End If
            End If
     Next i%
   ElseIf map50 = True And world = False Then
     kmx50orig = kmxcc - (km50x / mag) * sizex / 2
     kmy50orig = kmycc - (km50y / mag) * sizey / 2
     Dest.DrawWidth = mult%
     For i% = 1 To travelnum%
        T1 = travel(1, i%)
        T2 = travel(2, i%)
        Mode% = 2 'inverse transform from SKY to ITM
        Call ITMSKY(G11, G22, T1, T2, Mode%)
        kmxob = G11: kmyob = G22
        Xpnt = (kmxob - kmx50orig) / (km50x / mag)
        Ypnt = sizey - (kmyob - kmy50orig) / (km50y / mag)
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr
           xpno = Xpnt: ypno = Ypnt
        Else
            If ((xpno > sizex And Xpnt > sizex) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizey And Xpnt > sizey) Or (ypno < 0 And Ypnt < 0)) Then GoTo o945
           Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr
o945:      xpno = Xpnt: ypno = Ypnt
           If flg% = 1 Then
              xpno = xp1: ypno = yp1
              flg% = 0
              End If
           End If
     Next i%
  ElseIf world = True Then

     '(first check for busy signal from egg.exe)
      If Maps.Timer2.Enabled = True Then
         myfile = Dir(ramdrive + ":\wait.x")
         If myfile <> sEmpty Then
            waitime = Timer
            Do Until Timer > waitime + 0.5
               DoEvents
            Loop
            Exit Sub
            End If
         End If

      If mag > 1 Then
         lonc = lon '+ fudx / mag
         latc = lat '+ fudy / mag
         'wxorigin = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
         'wyorigin = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
         wxorigin = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
         wyorigin = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
      Else
         'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
         'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
         'wxorigin = lonc - 90#
         'wyorigin = latc + 90#
         lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
         latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
         wxorigin = lonc - deglog / 2
         wyorigin = latc + deglat / 2
         End If

     Dest.DrawWidth = mult%
     For i% = 1 To travelnum%
        T1 = travel(1, i%)
        T2 = travel(2, i%)
        If mag > 1 Then
           'Xpnt = (T1 - wxorigin) / (180# / (sizewx * mag))
           'Ypnt = (wyorigin - T2) / (180# / (sizewy * mag))
           Xpnt = (T1 - wxorigin) / (deglog / (sizewx * mag))
           Ypnt = (wyorigin - T2) / (deglat / (sizewy * mag))
        Else
           'Xpnt = (T1 - wxorigin) / (180# / sizewx)
           'Ypnt = (wyorigin - T2) / (180# / sizewy)
           Xpnt = (T1 - wxorigin) / (deglog / sizewx)
           Ypnt = (wyorigin - T2) / (deglat / sizewy)
           End If
        npnt% = npnt% + 1
        If npnt% = 1 Then
           Dest.PSet (Xpnt, Ypnt), colr
           xpno = Xpnt: ypno = Ypnt
        Else
            If ((xpno > sizewx And Xpnt > sizewx) Or (xpno < 0 And Xpnt < 0)) Or _
               ((ypno > sizewy And Xpnt > sizewy) Or (ypno < 0 And Ypnt < 0)) Then GoTo o955
           Dest.Line (xpno, ypno)-(Xpnt, Ypnt), colr
o955:      xpno = Xpnt: ypno = Ypnt
           If flg% = 1 Then
              xpno = xp1: ypno = yp1
              flg% = 0
              End If
           End If
     Next i%
     End If

End Sub
Public Sub sunrisesunset(Mode%)
    Dim XDIM As Double, YDIM As Double, skip As Boolean, MeanTemp As Double
    Dim skipkmx As Double, skipkmy As Double, lt As Double, lg As Double
    Dim beglog As Double, endlog As Double, kmx As Double, MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer
    Dim beglat As Double, endlat As Double, kmy As Double, ier As Integer
    Dim lt1 As Integer, lt2 As Integer, lg1 As Integer, lg2 As Integer
    Dim filt$, filt1$, filn$, filn2$, filn3$, l1 As Double, l2 As Double
    Dim filn33$, L2ch As Double, L1ch As Double, hgtch As Double, angch As Double, apch As Double, modch%
    Dim beglogch As Double, endlogch As Double, beglatch As Double, endlatch As Double
    Dim nstat%, maxstat%, manytiles As Boolean, tnn11 As Double, endfil%, endflag%
    Dim AddWait As Integer, treehgt As Double, treehgtch As Double
'    Dim A$, A1$, A2$
'    Dim integ&, intege%, T1#, E1#, D1#

   If world = True Then
      Screen.MousePointer = vbHourglass
      
      If DTMflag > 0 Then
         diflogss% = diflogs%
         diflatss% = diflats%
         maxangss% = maxangs%
         fullrangess% = fullranges%
         viewmodess% = viewmodes%
         modevalss = modevals
      Else
         diflogss% = diflog%
         diflatss% = diflat%
         maxangss% = maxang%
         fullrangess% = fullrange%
         viewmodess% = viewmode%
         modevalss = modeval
         End If
      
'      If WinVer <> 5 And WinVer <> 261 And Not AutoProf Then 'delete everything in the ramdrive
'        myfile = Dir(ramdrive + ":\*.*")
'        On Error GoTo out50
'        Do Until Not AutoProf And myfile = sEmpty
'           Kill ramdrive + ":\" + myfile
'           myfile = Dir(ramdrive + ":\*.*")
'        Loop
'        End If
     myfile = Dir(ramdrive + ":\*.bin")
     If myfile <> sEmpty Then Kill ramdrive + ":\" + myfile
     myfile = Dir(ramdrive + ":\*.bi1")
     If myfile <> sEmpty Then Kill ramdrive + ":\" + myfile

     myfile = Dir(drivjk_c$ + "newreaddtm.end") 'delete old end flag file if exists
     If myfile <> sEmpty Then Kill drivjk_c$ + "newreaddtm.end"
      
out50: Err.Clear

      If viewsearch = False Then
        l1 = Maps.Text6.Text
        l2 = Maps.Text5.Text
        If Maps.Text7 <> "" Then
           hgtworld = Maps.Text7
        Else
          If noheights = False Then
             Call worldheights(l1, l2, hgt)
             If hgt = -9999 Then hgt = 0
             Maps.Text7.Text = Str$(hgt)
             hgtworld = hgt
          Else
             hgtworld = 0
             Maps.Text7.Text = "0"
             End If
          End If
      Else
        l1 = lat
        l2 = lon
        hgtworld = searchhgt
        End If

      AutoPress = False
      If Mode% = 10 Then 'automatic directX button press
         Mode% = 1
         AutoPress = True
         End If
      sunmode% = Mode%
      Select Case Mode%
         Case Is >= 1 'sunrise
            'use averaged minimum temperature
            Mode% = 1
            beglat = l1 - diflatss% / 2
            endlat = beglat + diflatss%
            beglog = l2 - 0.1 - fullrangess% * (diflogss% / 2 - 0.1)
            endlog = beglog + (1 + fullrangess%) * diflogss% / 2
            'determine appropriate tile, and see if it is in the CD drive
         Case Is <= 0 'sunset
            Mode% = 0
            beglat = l1 - diflatss% / 2
            endlat = beglat + diflatss%
            endlog = l2 + 0.1 + fullrangess% * (diflogss% / 2 - 0.1)
            beglog = endlog - (1 + fullrangess%) * diflogss% / 2
            'determine appropriate tile, and see if it is in the CD drive
         Case Else
      End Select
      
     'now determine temperature for terrestrial refraction
     Call Temperatures(l1, l2, MinT, AvgT, MaxT, ier)
     MeanTemp = 0
     For ii = 1 To 12
       MeanTemp = AvgT(ii) + MeanTemp
     Next ii
     MeanTemp = MeanTemp / 12
'          Select Case Mode%
'             Case Is >= 1 'sunrise
'                'use averaged minimum temperature
'                MeanTemp = 0
'                For ii = 1 To 12
'                   MeanTemp = MT(ii) + MeanTemp
'                Next ii
'                MeanTemp = MeanTemp / 12
'             Case Is <= 0 'sunset
'                'use averaged average temperature
'                MeanTemp = 0
'                For ii = 1 To 12
'                   MeanTemp = AT(ii) + MeanTemp
'                Next ii
'                MeanTemp = MeanTemp / 12
'             Case Else
'            End Select
         'Alternatively, set MeanTemp = 0 and let readDTM read the WorldClim files
            
          If TemperatureModel% > 2 Then 'request that user input the ground temperature
W100:
             NewMeanTemp$ = InputBox("Enter ground temperature (deg C)", "Ground Temperature", Val(Format(Str$(MeanTemp), "###0.0")))
             If Val(NewMeanTemp$) <> MeanTemp Then
                Select Case MsgBox("The entered ground temperature (deg C) is equal to: " & NewMeanTemp$ _
                                   & vbCrLf & "" _
                                   & vbCrLf & "Is this correct?" _
                                   , vbYesNo Or vbQuestion Or vbDefaultButton1, "New ground temperature")
                
                    Case vbYes
                        MeanTemp = Val(NewMeanTemp$)
                        If MeanTemp < -30 Or MeanTemp > 40 Then
                           Select Case MsgBox("The suggested range of temperatures is from -30C to 40C." _
                                              & vbCrLf & "You inputed: " & Str$(MeanTemp) _
                                              & vbCrLf & "Do you want to keep your inputed value?" _
                                              , vbYesNo Or vbInformation Or vbDefaultButton2, "Ground temperature")
                           
                            Case vbYes
                           
                            Case vbNo
                              GoTo W100
                           End Select
                           End If
                           
                    Case vbNo
                        'try again
                        GoTo W100
                End Select
                End If
             End If

      Select Case viewmodess%
         Case 0
            sunmode% = Mode%
         Case 1
            If Mode% = 1 Then
               sunmode% = 2
            ElseIf Mode% = 0 Then
               sunmode% = -2
               End If
         Case 2
            If Mode% = 1 Then
               sunmode% = 3
            ElseIf Mode% = 0 Then
               sunmode% = -3
               End If
         Case 3
            If Mode% = 1 Then
               sunmode% = 4
            ElseIf Mode% = 0 Then
               sunmode% = -4
               End If
         Case Else
      End Select
      End If

      If fullrangess% = 1 Then 'view both horizons
        If Mode% = 1 Then
           sunmode% = 4
        ElseIf Mode% = 0 Then
           sunmode% = -4
           End If
        End If
        
      If autoazirange% = 1 Then
         'automatic determination of azimuth range
         maxangss% = MaxHalfAzimuthRange(l1)
         'finally, don't let the minimum be less than 45 degrees for uniformity purpose when
         'calculating most places in the world
         If maxangss% < 45 Then maxangss% = 45
      Else
         'check if recorded azimuth range is sufficient for the calculations
         maxangtmp% = MaxHalfAzimuthRange(l1)
         If maxangtmp% > maxangss% Then
            Call MsgBox("The recorded azimuth range may be too small for visible zemanim calculations at your chosen latitude!" _
                        & vbCrLf & "" _
                        & vbCrLf & "The recorded value is: " & Str$(maxangss%) _
                        & vbCrLf & "The recomended value is: " & Str$(maxangtmp%) _
                        & vbCrLf & "" _
                        & vbCrLf & "(Hint: you can turn on automatic azimuth ranges by checking the check box in the DTM limits dialog.)" _
                        , vbInformation Or vbDefaultButton1, "Maximum azimuth range")
            Exit Sub
            End If
         End If
 
     'query user for treehgt
     If Not AutoProf Or (AutoProf And AutoNum& = 0) Then
        treehgtStored = 0
        treehgtStr$ = InputBox("Enter tree height" & vbCrLf & "(leave zero if nothing to add)", "Add ''tree'' height to all DTM heights", 0)
        treehgt = Val(treehgtStr$)
     ElseIf AutoProf And AutoNum& > 0 Then
        treehgt = treehgtStored
        End If
     
50   'build data and header files of USGS EROS tile
     manytiles = False
     GoSub findtile 'determine if need to read multiple tiles, or CD's
     filn$ = worlddtm + ":\" + filt$ + "\" + filt$ 'input DEM file directory
     If lt1 < lt2 Then
        lt1 = lt2
        lg1 = lg2
        filn$ = worlddtm + ":\" + filt1$ + "\" + filt1$
        End If
     filn2$ = filn$
     filn3$ = ramdrive + ":\" + Mid$(filn$, 4, 7) 'file directory name for output
     'check if it exists
     On Error GoTo diskerrhandler
     myfile = Dir(worlddtm + ":\" + filt$, vbDirectory)
     If myfile = sEmpty Then
        If checkdtm = True Then
          'check for an old DTM file
          GoTo 80
          End If
       'else check if there is any USGS EROS CD in the CD player
       If DTMflag = 0 Then
          myfile = Dir(worlddtm + ":\E020N40\E020N40.GIF") 'Dir(worlddtm + ":\Gt30dem.gif")
       Else 'SRTM
          If Dir(srtmdtm & ":\3AS\", vbDirectory) = sEmpty And _
             Dir(srtmdtm & ":\USA\", vbDirectory) = sEmpty Then
             Screen.MousePointer = vbDefault
             ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             response = MsgBox("Can't find the SRTM tiles!", vbOKOnly + vbExclamation, "Maps & More")
'             ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             BringWindowToTop (mapPictureform.hwnd)
             Maps.Toolbar1.Buttons(26).value = tbrUnpressed
             Maps.Toolbar1.Buttons(27).value = tbrUnpressed
             Screen.MousePointer = vbDefault
             Exit Sub
          Else
             GoTo skipcheck
             End If
          End If
       If myfile = sEmpty Then
         'determine which is the right CD '<<<>>>
         If lt1 > -60 Then
            nx% = Fix((lg1 + 180) * 0.025)
            ny% = Int((90 - lt1) * 0.02)
            numCD% = worldCD%(ny% * 9 + nx% + 1)
         Else
            numCD% = 5
            End If
         Screen.MousePointer = vbDefault
         ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         response = MsgBox("Please insert USGS EROS CD#" + LTrim$(Str(numCD%)) + " in the CD drive, and try again.", vbOKOnly + vbExclamation, "Maps & More")
'         ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (mapPictureform.hwnd)
         Maps.Toolbar1.Buttons(26).value = tbrUnpressed
         Maps.Toolbar1.Buttons(27).value = tbrUnpressed
         Screen.MousePointer = vbDefault
         Exit Sub
         End If
      End If
skipcheck:
    'check that this is a new run
    myfile = Dir(drivjk_c$ + "eros.tm3")
    mapEROSDTMwarn.Visible = False
    If myfile <> sEmpty Then
       filtmp% = FreeFile
       Open drivjk_c$ + "eros.tm3" For Input As #filtmp%
       Line Input #filtmp%, filn33$
       Input #filtmp%, L1ch, L2ch, hgtch, angch, apch, modch%, modvalch%
       Input #filtmp%, beglogch, endlogch, beglatch, endlatch
       Input #filtmp%, noVoidflagch, CalculateProfilech
       Input #filtmp%, AziStepch
       Input #filtmp%, IgnoreTilesch
       tmflag% = 1
       Input #filtmp%, TemperatureModelch
       tmflag% = 2
       Input #filtmp%, MeanTempch
       tmpflag% = 3
       Input #filtmp%, treehgtch
       tmflag% = 0
       Close #filtmp%
       myfile2 = Dir(filn33$)
       'L2ch - L2 < 0.08 And L1ch - L1 < 0.08 And
       skip = False
       cc = beglog - beglogch
       If filn33$ = filn3$ + ".BIN" And modch% = sunmode% And modval% = modvalch% And myfile2 <> sEmpty And _
          ((sunmode% >= 1 And Abs(endlogch - endlog) < 0.05 And beglog - beglogch >= -0.0001 And _
          Abs(hgtch - hgtworld) < 0.1) Or _
          (sunmode% <= 0 And endlogch - endlog >= -0.0001 And Abs(beglogch - beglog) < 0.05 And _
          Abs(hgtch - hgtworld) < 0.1)) And _
          (Abs(beglatch - beglat) < 0.05 And Abs(endlatch - endlat) < 0.05) And _
          noVoidflagch = noVoidflag And CalculateProfilech = CalculateProfile And AziStepch = AziStepf% * 0.01 And _
          IgnoreTilesch = IgnoreTiles% And TemperatureModelch = TemperatureModel% And MeanTempch = MeanTemp And _
          treehgtch = treehgt Then
          mapprogressfm.Visible = True
          mapEROSDTMwarn.Visible = False
          beglog = beglogch
          endlog = endlogch
          beglat = beglatch
          endlatch = endlatch
          Open drivjk_c$ + "eros.tm3" For Output As #filtmp%
          Print #filtmp%, filn33$
          'Print #filtmp%, Format(L1, "#0.00000"), Format(L2, "##0.00000"), Format(hgtworld, "###0.0"), maxang%, apprn, sunmode%, modeval%
          'Print #filtmp%, Format(beglog, "##0.00000"), Format(endlog, "##0.00000"), Format(beglat, "#0.00000"), Format(endlat, "#0.00000")
     
          Write #filtmp%, l1, l2, hgtworld, maxangss%, apprn, sunmode%, modevalss
          Write #filtmp%, beglog, endlog, beglat, endlat
          'write flag whether to remove SRTM voids
          Write #filtmp%, noVoidflag, CalculateProfile
          Write #filtmp%, AziStepf% * 0.01
          Write #filtmp%, IgnoreTiles%
          Write #filtmp%, TemperatureModel%
          Write #filtmp%, MeanTemp
          Write #filtmp%, treehgt
          Close #filtmp%
          'write eros.tm5 file detailing if this is sunrise,sunset, or both views
          Open ramdrive + ":\eros.tm5" For Output As #filtmp%
          Write #filtmp%, sunmodess%, modevalss
          Close #filtmp%
          skip = True
          GoTo tst2
          End If
       End If
 '  Else 'check if there is a BIN, BI1 and eros.tm3 file stored in c:\dtm
      'that matches the desired extraction range
80    doclin$ = Dir(drivdtm$ & "*.BIN")
      myfile = Dir(drivdtm$ & "eros.tm3")
      If doclin$ <> sEmpty And myfile <> sEmpty And Dir(drivdtm$ & "*.BI1") <> sEmpty Then 'read this eros.tm3 and check it
         filtmp% = FreeFile
         Open drivdtm$ & "eros.tm3" For Input As #filtmp%
         Line Input #filtmp%, filn33$
         Input #filtmp%, L1ch, L2ch, hgtch, angch, apch, modch%, modvalch%
         Input #filtmp%, beglogch, endlogch, beglatch, endlatch
         Input #filtmp%, noVoidflagch, CalculateProfilech
         Input #filtmp%, AziStepTmpch
         Input #filtmp%, IgnoreTiles%
         tmpflag% = 1
         Input #filtmp%, TemperatureModel%
         tmpflag% = 2
         Input #filtmp%, MeanTemp
         tmpflag% = 3
         Input #filtmp%, treehgt
         tmpflag% = 0
         Close #filtmp%
         'L2ch - L2 < 0.08 And L1ch - L1 < 0.08 And
         skip = False
         If ramdrive + ":\" + doclin$ = filn3$ + ".BIN" And modch% = sunmode% And modval% = modvalch% And _
            ((sunmode% >= 1 And Abs(endlogch - endlog) < 0.05 And beglog - beglogch >= -0.0001 And _
            Abs(hgtch - hgtworld) < 0.1) Or _
            (sunmode% <= 0 And endlogch - endlog >= -0.0001 And Abs(beglogch - beglog) < 0.05 And _
            Abs(hgtch - hgtworld) < 0.1)) And _
            Abs(beglatch - beglat) < 0.05 And Abs(endlatch - endlat) < 0.05 And _
            noVoidflagch = noVoidflag Then
            mapprogressfm.Visible = True
            mapEROSDTMwarn.Visible = False
            beglog = beglogch
            endlog = endlogch
            beglat = beglatch
            endlatch = endlatch
            Open drivjk_c$ + "eros.tm3" For Output As #filtmp%
            Print #filtmp%, filn33$
            'Print #filtmp%, Format(L1, "#0.00000"), Format(L2, "##0.00000"), Format(hgtworld, "###0.0"), maxang%, apprn, sunmode%, modeval%
            'Print #filtmp%, Format(beglog, "##0.00000"), Format(endlog, "##0.00000"), Format(beglat, "#0.00000"), Format(endlat, "#0.00000")
            Write #filtmp%, l1, l2, hgtworld, maxangss%, apprn, sunmode%, modevalss
            Write #filtmp%, beglog, endlog, beglat, endlat
            'write flag whether to remove SRTM voids
            Write #filtmp%, noVoidflag, CalculateProfile
            Write #filtmp%, AziStep% * 0.01
            Write #filtmp%, IgnoreTiles%
            Write #filtmp%, TemperatureModel%
            Write #filtmp%, MeanTemp
            Write #filtmp%, treehgt
            Close #filtmp%
            Open ramdrive + ":\eros.tm5" For Output As #filtmp%
            Write #filtmp%, sunmode%, modeval
            Close #filtmp%
            FileCopy drivdtm$ & "" + doclin$, ramdrive + ":\" + doclin$
            FileCopy drivdtm$ & "" + Mid$(doclin$, 1, Len(doclin$) - 4) + ".BI1", ramdrive + ":\" + Mid$(doclin$, 1, Len(doclin$) - 4) + ".BI1"
            myfile = Dir(drivjk_c$ + "eros.tmp")
            If myfile <> sEmpty And AutoPress = False Then Kill drivjk_c$ + "eros.tmp"
            'also look for a DirectX file land.x
            If Dir(drivdtm$ & "land.x") <> sEmpty And Dir(drivdtm$ & "land.x") <> sEmpty Then
               FileCopy drivdtm$ & "land.x", ramdrive + ":\land.x"
               FileCopy drivdtm$ & "land.tm3", ramdrive + ":\land.tm3"
               End If
            skip = True
            GoTo tst2
            End If
         ElseIf checkdtm = True And Not NoCDWarning Then
'            ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            BringWindowToTop (mapPictureform.hwnd)
            response = MsgBox("USGS EROS CD not found!  Please enter the appropriate CD, and then press the DTM button!", vbCritical + vbOKOnly, "Maps & More")
            ret = SetWindowPos(mapPictureform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            NoCDWarning = True
            Exit Sub
         End If
'       mapprogressfm.Visible = True
'       mapEROSDTMwarn.Visible = False
'       End If
    'if got here, then write new eros.tm3 file
     filtmp% = FreeFile
     Open drivjk_c$ + "eros.tm3" For Output As #filtmp%
     Print #filtmp%, filn3$ + ".BIN"
'     Print #filtmp%, Format(L1, "#0.00000"), Format(L2, "##0.00000"), Format(hgtworld, "###0.0"), maxang%, apprn, sunmode%, modeval%
'     Print #filtmp%, Format(beglog, "##0.00000"), Format(endlog, "##0.00000"), Format(beglat, "#0.00000"), Format(endlat, "#0.00000")
     Write #filtmp%, l1, l2, hgtworld, maxangss%, apprn, sunmode%, modevalss
     Write #filtmp%, beglog, endlog, beglat, endlat
     'write flag whether to remove SRTM voids
     Write #filtmp%, noVoidflag, CalculateProfile
     Write #filtmp%, AziStepf% * 0.01
     Write #filtmp%, IgnoreTiles%
     Write #filtmp%, TemperatureModel%
     Write #filtmp%, MeanTemp
     Write #filtmp%, treehgt
     Close #filtmp%
     Open ramdrive + ":\eros.tm5" For Output As #filtmp%
     If DTMflag <= 0 Then
        Write #filtmp%, sunmode%, modeval
     Else
        Write #filtmp%, sunmode%, modevals
        End If
     
     Close #filtmp%

     myfile = Dir(ramdrive + ":\*.bin")
     If myfile <> sEmpty Then Kill ramdrive + ":\" + myfile
     myfile = Dir(ramdrive + ":\*.bi1")
     If myfile <> sEmpty Then Kill ramdrive + ":\" + myfile

     myfile = Dir(drivjk_c$ + "newreaddtm.end") 'delete old end flag file if exists
     If myfile <> sEmpty Then Kill drivjk_c$ + "newreaddtm.end"

'     '*******************************************************
'     'replace the following code with the C++ program readDTM
'     'clear the RAMDRIVE directory
'
'     GoSub dtmfiles 'read EROS info files and open DEM file
'     filnum2% = FreeFile
'     Open filn3$ + ".BIN" For Random As #filnum2% Len = 2
'
'      'Read the chosen TILE
'      skipkmx = xdim
'      skipkmy = ydim
'      numrc& = 0
'      Maps.Toolbar1.Refresh
'      mapprogressfm.Visible = True
'      mapEROSDTMwarn.Visible = False
'      mapprogressfm.StatusBar1.Panels(1) = "Extracting the relavant portion of the DTM"
'      ret = SetWindowPos(mapprogressfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'      mapprogressfm.ProgressBar1.Max = Int((Abs((beglat - endlat) / skipkmy) + 1) * (Abs((endlog - beglog) / skipkmx) + 1))
'      mapprogressfm.Command1.Visible = True
'      mapprogressfm.Refresh
'
'            ' vbreadfil% = FreeFile '<-------------------
'            ' Open drivjk$ + "vbread.out" For Output As #vbreadfil%
'
'      For kmy = endlat To beglat Step -skipkmy
'          DoEvents
'          If abortDTM = True Then 'pressed X button
'             abortDTM = False
'             GoTo mm999
'             End If
'          IKMY& = Int((lt1 - kmy) / skipkmy) + 1
'          For kmx = beglog To endlog Step skipkmx
'              If manytiles = True Then 'check for new tile at each data point
'                 GoSub findtiles
'                 If filt1$ <> filt$ Then
'                    'close opened DEM file
'                    Close #filnum1%
'                    IKMY& = Int((lt1 - kmy) / skipkmy) + 1
'                    filt$ = filt1$
'                    filn$ = "j:\" + filt$ + "\" + filt$  'input DEM file directory
'                    'check if it exists
'                    'On Error GoTo diskerrhandler
'                    myfile = Dir("j:\" + filt$, vbDirectory)
'                    If myfile = sEmpty Then
'                       'determine which CD it's on
'                       GoSub findCD
'                       GoSub dtmfiles
'                    Else 'open new files
'                       GoSub dtmfiles
'                       End If
'                    End If
'                 End If
'              numrc& = numrc& + 1
'              mapprogressfm.ProgressBar1.Value = numrc&
'              nn1$ = LTrim$(Str(CInt((mapprogressfm.ProgressBar1.Value / mapprogressfm.ProgressBar1.Max) * 100))) + "%"
'              If mapprogressfm.Label1.Caption <> nn1$ Then
'                 mapprogressfm.Label1.Caption = nn1$
'                 mapprogressfm.Label1.Refresh
'                 End If
'              IKMX& = Int((kmx - lg1) / skipkmx) + 1
'              tncols = ncols%
'              numrec& = (IKMY& - 1) * tncols + IKMX&
'
''              For itrial& = 2853599 To 71000000
''                  DoEvents
''                  Get #filnum1%, (itrial& - 1) * 2 + 1, io%
''                  If io% <> -3624 Then
''                     Write #vbreadfil%, itrial&, io%
''                     End If
''              Next itrial&
''              Close #vbreadfil%
'
'              Get #filnum1%, (numrec& - 1) * 2 + 1, io%
'
'              'Write #vbreadfil%, (numrec& - 1) * 2, io% '<------------
'
'              Put #filnum2%, numrc&, io%
'          Next kmx
'       Next kmy
'       Close #filnum1%
'       Close #filnum2%
'       Close #vbreadfil% '<----------------
'     '*******************************************************
       '>>>>>>>>>>>>>>>>>>>>>>>>>>>auto section<<<<<<<<<<<<<<<<<<<<<<<<<<

       'now save .BI1 file in the dtm directory as a temporary file
tst2:  myfile = Dir(ramdrive + ":\*.BI1")
       If myfile <> sEmpty Then
        FileCopy ramdrive + ":\" + myfile, drivdtm$ & "" + Mid$(myfile, 1, Len(myfile) - 4) + ".tBI"
        End If

       If AutoPress Then 'automatically press the DirectX button
           If DTMflag <= 0 Then
              mapprogressfm.optGTOPO30.value = True
           ElseIf DTMflag = 1 Then
              mapprogressfm.optSRTM1.value = True
           ElseIf DTMflag = 2 Then
              mapprogressfm.optSRTM2.value = True
              End If
           mapprogressfm.Command2.value = True
           GoTo at100
       Else
           mapEROSDTMwarn.Visible = False
           End If

       With mapprogressfm
        .Visible = True
        .ProgressBar1.Visible = False
        .Label1.Visible = False
        .StatusBar1.Panels(1) = "For 2D, choose how much to shave. and press graph, or pick 3D View"
        .Text2.Visible = True
        .UpDown1.Visible = True
        .Picture1.Visible = True
        .Command1.Visible = True
        .Command2.Visible = True
        .Label3.Visible = True
        .Acceptbut.Visible = True
        .frmDTM.Visible = True
        .Text2 = apprn
       End With
       If skip = True Then mapprogressfm.Text2 = apch
'       ret = SetWindowPos(mapprogressfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       BringWindowToTop (mapprogressfm.hwnd)
       '>>>>>>>>>>>>>>>>>>>>>>>>>>>auto section<<<<<<<<<<<<<<<<<<<<<<<<<<
       
       Screen.MousePointer = vbDefault
       IntOld% = 0
       waitime = Timer
       Do Until mapprogressfm.Visible = False
          DoEvents
          If AutoProf Then
             If Delay% <> 0 Then
                AddWait = Delay%
             Else
                AddWait = 10
                End If
             If Timer > waitime + AddWait Then
                'push the button automatically after a minute
                mapprogressfm.Acceptbut.value = True
             Else
                If Int(10 - Timer + waitime) <> IntOld% Then
                   mapprogressfm.StatusBar1.Panels(1).Text = "Auto Mode...Calc. starting in" & Str$(Int(AddWait - Timer + waitime)) & " sec."
                   'mapprogressfm.StatusBar1.Panels(1).Text = "Auto Mode...Calc. starting in" & Str$(Int(10 - Timer + waitime)) & " sec."
                   mapprogressfm.StatusBar1.Refresh
                   IntOld% = Int(10 - Timer + waitime)
                   End If
                End If
             End If
       Loop
       
       'now write batch file
at100: If viewer3D = True Then
          myfile = Dir(drivjk_c$ + "eros.tmp")
          If myfile <> sEmpty And AutoPress = False Then
             Kill drivjk_c$ + "eros.tmp"
             'erase this file, since can only be accessed if
             'just calculated
             End If
          'output eros.tm4 file which signals directx
          'If DTMflag > 0 Then
             filtmp4% = FreeFile
             Open ramdrive & ":\eros.tm4" For Output As #filtmp4
             Write #filtmp4%, l2
             Write #filtmp4%, l1
             Write #filtmp4%, beglog
             Write #filtmp4%, endlog
             Write #filtmp4%, endlat
             Write #filtmp4%, beglat
             'flag newreadDTM to write initial land.x file
             Write #filtmp4%, 0
             'flag whether to cosmetically remove radar shadow voids
             'Write #filtmp4%, noVoidflag
             Close #filtmp4%
             'End If
          accept = True
          
          'GoTo mm600
          
       Else 'erase any old eros.tm4 file
          myfile = Dir(ramdrive + ":\eros.tm4")
          If myfile <> sEmpty Then Kill ramdrive & ":\eros.tm4"
       
          End If
       If accept = False Then
           'X button pressed in mapprogressfm
           GoTo mm999
          End If
       If skip = True And apprn = apch And Dir(drivjk_c$ + "eros.tmp") <> sEmpty Then
          GoTo mm500
          End If
       Screen.MousePointer = vbHourglass
       filtm3num% = FreeFile
       Open drivjk_c$ + "eros.tm3" For Output As #filtm3num%
       Print #filtm3num%, filn3$ + ".BIN"
'       Print #filtm3num%, Format(L1, "#0.00000"), Format(L2, "##0.00000"), Format(hgtworld, "###0.0"), maxang%, apprn, sunmode%, modeval%
'       Print #filtm3num%, Format(beglog, "##0.00000"), Format(endlog, "##0.00000"), Format(beglat, "#0.00000"), Format(endlat, "#0.00000")
       Write #filtmp%, l1, l2, hgtworld, maxangss%, apprn, sunmode%, modevalss
       Write #filtmp%, beglog, endlog, beglat, endlat
       'write flag whether to remove SRTM voids
       Write #filtmp%, noVoidflag, CalculateProfile
       Write #filtmp%, AziStepf% * 0.01
       Write #filtmp%, IgnoreTiles%
       Write #filtmp%, TemperatureModel%
       Write #filtmp%, MeanTemp
       Write #filtmp%, treehgt
       Close filtm3num%
       'define default drives to write to
       ChDrive "c"
       ChDir drivjk_c$

'*************this is C++ version of rderos2--it turned out to be slower than PROFORT!!!!********
'     myfile = Dir(drivjk$ + "rderos3.end") 'delete old end flag file if exists
'     If myfile <> sEmpty Then Kill drivjk$ + "rderos3.end"
'
''   C++ program
'     RetVal = Shell("c:\progra~1\micros~2\MyProjects\rderos3\release\rderos3.exe", vbNormalFocus)
'    'don't go on until finished computing
'ts3: endflag% = 0
'     myfile = Dir(drivjk$ + "rderos3.end")
'     If myfile = sEmpty Then
'        waittim = Timer + 0.2
'        Do Until Timer > waittim
'           DoEvents
'        Loop
'        GoTo ts3
'     Else
'        endfil% = FreeFile
'        Open drivjk$ + "rderos3.end" For Input As #endfil%
'        Input #endfil%, endflag%
'        Close #endfil%
'        End If
'     'do second check using Windows API
'     lResult = FindWindow(vbNullString, "Calculating the profile")
'     If lResult > 0 Then GoTo ts3 'it is still working, so keep on looping
'     Kill (drivjk$ + "rderos3.end")
'     If endflag% = 1 Then GoTo mm999 'data extraction was unsuccessful
'*********************************end of C++ rderos2 program***************
       
       'erase old signal files if they exists
       myfile = Dir(drivjk_c$ + "erosend.tmp")
       If myfile <> sEmpty Then Kill drivjk_c$ + "erosend.tmp"
       myfile = Dir(drivjk_c$ + "erostat.tmp")
       If myfile <> sEmpty Then Kill drivjk_c$ + "erostat.tmp"
       
       Close 'close any open files
       'if eros.tmp exists, erase it
       If Dir(drivjk_c$ & "eros.tmp") <> sEmpty And AutoPress = False Then Kill drivjk_c$ & "eros.tmp"
       
'    run C++ program that extracts the relevant parts of the DTMs
     'Dim RetVal
     'If viewer3D And DTMflag <= 0 Then 'use old version of program until finished debugging newer version
     '   RetVal = WinExec("c:\samples\vc98\sdk\graphics\directx\readdtm_old\release\newreadDTM.exe", SW_SHOW)
     'Else
        'RetVal = Shell("c:\samples\vc98\sdk\graphics\directx\readdtm\release\newreaddtm.exe", vbNormalFocus)
        RetVal = Shell(drivjk_c$ + "newreaddtm.exe", 1) ' Run new version of newreaddtm that also includes rderos2

     '   End If
    'don't go on until finished computing
ts5: endflag% = 0
     myfile = Dir(drivjk_c$ + "newreaddtm.end")
     If myfile = sEmpty Then
        waittim = Timer + 0.2
        Do Until Timer > waittim
           DoEvents
        Loop
        GoTo ts5
     Else
        endfil% = FreeFile
        Open drivjk_c$ + "newreaddtm.end" For Input As #endfil%
        Input #endfil%, endflag%
        Close #endfil%
        End If
     'do second check using Windows API
     lResult = FindWindow(vbNullString, "Extracting relevant portion of the DTM")
     If lResult > 0 Then GoTo ts5 'it is still working, so keep on looping
     Kill (drivjk_c$ + "newreaddtm.end")
     If endflag% = 1 Then GoTo mm999 'data extraction was unsuccessful
       
     myfile = Dir(drivjk_c$ + "erosend.tmp")
     If myfile <> sEmpty Then Kill drivjk_c$ + "erosend.tmp"
     
     'now run egg if flagged
     If viewer3D = True Then
        ret = Shell(SamplesDir$ & "vc98\sdk\graphics\directx\egg\debug\egg.exe", vbNormalFocus)
        GoTo mm600
        End If
     
      If OnlyExtractFile Then 'don't run rderos2, just extract data
          mapprogressfm.Visible = False
          mapPictureform.Refresh
          Maps.Picture1.Refresh
          GoTo mm500
          End If
      
      If CalculateProfile = 1 Then
        GoTo ms700 'ms400 'newest version includes rderos2 inside the c-code of newreaddtm.exe
        
      ElseIf CalculateProfile = 0 Then 'calculate profiles using rderos2.exe
         'if not running egg, then run rderos2
         
         RetVal = Shell(drivjk_c$ + "rderos2.exe", 2) ' Run rderos2 as DOS shell
         
         waitime = Timer
         Do Until Timer > waitime + 1
            DoEvents
         Loop
         
         GoTo ms400 'new c version without stat
         
         'output filed is written to file c:\jk\eros.tmp
         waitime = Timer 'wait a bit for Chaim-PIV to activate the slave hard disk F:
         If WinVer <> 5 And WinVer <> 261 Then waiting% = 1 Else waiting% = 10
         Do Until Timer > waitime + waiting%
            DoEvents
         Loop
         lResult1 = FindWindow(vbNullString, "rderos2")
         lResult2 = FindWindow(vbNullString, drivjk_c$ & "rderos2.exe")
         If lResult1 <> 0 Or lResult2 <> 0 Then
             mapprogressfm.Visible = True
             mapEROSDTMwarn.Visible = False
'             ret = SetWindowPos(mapprogressfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             BringWindowToTop (mapprogressfm.hwnd)
             With mapprogressfm
               .ProgressBar1.Visible = True
               .StatusBar1.Panels(1) = "Calculating the profile"
               .Text2.Visible = False
               .UpDown1.Visible = False
               .Picture1.Visible = False
               .Command1.Visible = False
               .Command2.Visible = False
               .frmDTM.Visible = False
             End With
          
             waitime = Timer
             Do Until Timer > waitime + 0.5
                DoEvents
             Loop
             'now plot the eros.tmp file output from rderos2 or from newreaddtm
             myfile = Dir(drivjk_c$ + "erostat.tmp")
             If myfile <> sEmpty Then
suns500:        tmpfil% = FreeFile
                Open drivjk_c$ + "erostat.tmp" For Input As #tmpfil%
                nlines% = 0
                Do Until EOF(tmpfil%)
                   Input #tmpfil%, nstat%, maxstat%
                   nlines% = 1
                Loop
                Close #tmpfil%
                If nlines% = 0 Then GoTo suns500
                mapprogressfm.ProgressBar1.Max = maxstat%
                nn1$ = LTrim$(CStr(CInt(nstat * 100 / maxstat%))) + "%"
                If mapprogressfm.Label1.Caption <> nn1$ Then
                   mapprogressfm.Label1.Caption = nn1$
                   mapprogressfm.Label1.Refresh
                   End If
                End If
             End If
          nstat% = 1
          
ms400:
          lResult = FindWindow(vbNullString, drivjk_c$ & "rderos2.exe")
          Do Until lResult = 0
             DoEvents
             lResult = FindWindow(vbNullString, drivjk_c$ & "rderos2.exe")
          Loop
          
          GoTo ms700 'new c version without stats

mm450:    Do Until lResult = 0

             If Dir(drivjk_c$ + "erosend.tmp") <> sEmpty Then
                nstat% = 100
                mapprogressfm.ProgressBar1.value = nstat%
                mapprogressfm.Label1.Caption = "100"
                mapprogressfm.Label1.Refresh
                Exit Do
                End If
   
'             ret = SetWindowPos(mapprogressfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             BringWindowToTop (mapprogressfm.hwnd)
             myfile = Dir(drivjk_c$ + Trim$(Str$(nstat%)))
             tmpfil% = FreeFile
             If myfile <> sEmpty Then
                Open drivjk_c$ + LTrim$(Str(nstat%)) For Input As #tmpfil%
                nlines% = 0
                Do Until EOF(tmpfil%)
                   Input #tmpfil%, nstat%
                   nlines% = 1
                Loop
                Close #tmpfil%
                tnn11 = nstat% / CDbl(maxstat%)
                If tnn11 > 1 Then tnn11 = 1
                nn1$ = LTrim$(Format(tnn11 * 100, "##0")) + "%"
                If nlines% = 1 Then
                   nstat% = nstat% + 1
                Else
                   'didn't finsih writing
                   End If
                'wait a bit
                'waitime = Timer
                'Do Until Timer > waitime + 0.05
                'Loop
             Else
                nstat% = nstat% - 1 'overshot
                If nstat% = 0 Then nstat% = 1
                End If
             If mapprogressfm.Label1.Caption <> nn1$ Then
                If nstat% > maxstat% Then nstat% = maxstat%
                mapprogressfm.ProgressBar1.value = nstat%
                mapprogressfm.Label1.Caption = nn1$
                mapprogressfm.Label1.Refresh
                mapprogressfm.ProgressBar1.Refresh
                mapprogressfm.Refresh
                End If
             lResult = FindWindow(vbNullString, drivjk_c$ & "rderos2.exe")
             DoEvents
          Loop
          'check again if rderos2 disappeared
          
ms600:
          lResult = FindWindow(vbNullString, drivjk_c$ & "rderos2.exe")
          If lResult <> 0 Then GoTo mm450
          
ms700:
          mapprogressfm.Visible = False
          mapPictureform.Refresh
          Maps.Picture1.Refresh
   
          'now erase status files
          'first wait a bit
          waitime = Timer + 0.1
          Do Until Timer > waitime
             DoEvents
          Loop
          If Dir(drivjk_c$ + "erostat.tmp") <> sEmpty Then Kill drivjk_c$ + "erostat.tmp"
          If Dir(drivjk_c$ + "erosend.tmp") <> sEmpty Then Kill drivjk_c$ + "erosend.tmp"
          For i% = 1 To maxstat%
             If Dir(drivjk_c$ + Trim$(Str$(i%))) <> sEmpty Then Kill drivjk_c$ + Trim$(Str$(i%))
          Next i%
          End If

       waitime = Timer
       Do While Dir(drivjk_c$ + "eros.tmp") = sEmpty And Timer <= waitime + 1
          DoEvents
       Loop
       Screen.MousePointer = vbDefault
       myfile = Dir(drivjk_c$ & "eros.tmp")
       If myfile = sEmpty Then 'if empty then, something is wrong
          MsgBox "Can't find the eros.tmp output file (horizon profile)." & vbLf & _
                 "This usually means something is wrong with either the" & vclf & _
                 "newreaddtm or rderos2 analysis." & vbLf & _
                 "Hint: check your longitude and/or latitude ranges.", _
                  vbCritical + vbOKOnly, "Maps & More"
          GoTo mm999
          End If

mm500: mapgraphfm.Visible = True
'       ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       BringWindowToTop (mapgraphfm.hwnd)
       '>>>>>>>>>>>>>>>>>>>>>>>>>>>auto section<<<<<<<<<<<<<<<<<<<<<<<<<<
       waitime = Timer
       firstime% = 1
       IntOld% = 0
       
        With mapgraphfm
        .Picture1.Refresh
        .UpDown2.SetFocus
        .Line1.Refresh
        .Line2.Refresh
        .Line3.Refresh
        .Line4.Refresh
        .Label1.Refresh
        .Label2.Refresh
        .Label3.Refresh
        .Label4.Refresh
        .Label5.Refresh
        .Label6.Refresh
        .Label7.Refresh
        .Label8.Refresh
        .Label9.Refresh
        .MSFlexGrid1.Refresh
        .Command1.Refresh
        .Command2.Refresh
        .Command3.Refresh
        .frmObstructions.Refresh
        .TimeZonebut.Refresh
        .Text3.Refresh
        .restorelimitsbut.Refresh
        .Command3.value = True 'show the obstructions
        End With
        
            
       Do While killpicture = False
          DoEvents
          
          If Delay% >= 15 Then
            If firstime% = 1 Then waiting% = 15
          ElseIf Delay% < 15 Then
            If firstime% = 1 Then waiting% = Delay% + 5 'add five seconds to see the profile and give time to skip it
            End If
          If Not killpicture Then Delay% = Val(mapgraphfm.txtDelay)
          If firstime% = 0 Then waiting% = Delay% ' 45
          
          If AutoProf Then 'automatic pilot for calculating profiles
             If Int(waiting% - Timer + waitime) <> IntOld% Then
                mapgraphfm.StatusBar1.Panels(1).Text = "Automatic mode...calendar button will be pressed in:" & Str$(Int(waiting% - Timer + waitime)) & " sec."
                mapgraphfm.StatusBar1.Refresh
                IntOld% = Int(waiting% - Timer + waitime)
                End If
             If Timer > waitime + waiting% Then
               mapgraphfm.Calendarbut.value = True
               If firstime% = 1 Then 'pushed once
                  firstime% = 0
                  waitime = Timer
               ElseIf firstime% = 0 Then 'pushed twice
                  'profile written and bat file updated, so exit
                  waitime2 = Timer 'give it a little time to finish all processes
                  Do Until Timer > waitime2 + 1
                     DoEvents
                  Loop
                  mapgraphfm.cmdExit.value = True
                  End If
               End If
             End If
       
       Loop
       
       killpicture = False
       'If Dir(drivjk$ + "eros.tmp") <> sEmpty Then Kill drivjk$ + "eros.tmp"

'       ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       BringWindowToTop (mapPictureform.hwnd)
       GoTo mm999
mm600: init% = 0
mm650: If routeload = True Or showroute = True Then GoTo mm999
       lResult = FindWindow(vbNullString, "3D Viewer")
       If lResult <> 0 Then
          'wait a bit a try again
          waitime = Timer
          If init% = 0 Then
             waiter = 5
             init% = 1
          Else
             waiter = 1
             End If
          Do While Timer < waitime + waiter
             DoEvents
          Loop
          GoTo mm650
          End If
        viewer3D = False

mm999: Maps.Toolbar1.Buttons(26).value = tbrUnpressed
       Maps.Toolbar1.Buttons(27).value = tbrUnpressed
       tblbuttons(26) = 0
       tblbuttons(27) = 0
       If Maps.Toolbar1.Buttons(18).value = tbrPressed Then
          Maps.Toolbar1.Buttons(18).value = tbrUnpressed
          tblbuttons(18) = 0
          End If
       If showroute = False And routeload = False Then
          Maps.Timer2.Enabled = False
          showroute = False
          For i% = 23 To 25
             tblbuttons(i%) = 0
             Maps.Toolbar1.Buttons(i%).value = tbrUnpressed
             Maps.Toolbar1.Buttons(i%).Enabled = False
          Next i%
          End If

       Screen.MousePointer = vbDefault
Exit Sub

findtile:
   If beglat > -60 Then
      nx% = Fix((beglog + 180) * 0.025)
      lg1 = -180 + nx% * 40
      If Abs(lg1) >= 100 Then
         lg1ch$ = RTrim$(LTrim$(Str$(Abs(lg1))))
      Else
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
         End If
      If lg1 < 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ny% = Int((90 - beglat) * 0.02)
      lt1 = 90 - 50 * ny%
      lt1ch$ = LTrim$(RTrim$(Str$(Abs(lt1))))
      If lt1 > 0 Then
         ns$ = "N"
      Else
         ns$ = "S"
         End If
      filt1$ = EW$ + lg1ch$ + ns$ + lt1ch$
      NROWS = 6000
      NCOLS = 4800
   Else 'Antartic - Cd #5
      nx% = Fix((beglog + 180) / 60)
      lg1 = -180 + nx% * 60
      If Abs(lg1) >= 100 Then
         lg1ch$ = LTrim$(RTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) < 100 And Abs(lg1) <> 0 Then
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) = 0 Then
         lg1ch$ = "000"
         End If
      If lg1 <= 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ns$ = "S"
      lt1 = -60
      lt1ch$ = "60"
      filt1$ = EW$ + lg1ch$ + ns$ + lt1ch$
      NROWS = 3600
      NCOLS = 7200
      End If

   If endlat > -60 Then
      nx% = Fix((endlog + 180) * 0.025)
      lg2 = -180 + nx% * 40
      If Abs(lg2) >= 100 Then
         lg1ch$ = RTrim$(LTrim$(Str$(Abs(lg2))))
      Else
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg2))))
         End If
      If lg2 < 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ny% = Int((90 - endlat) * 0.02)
      lt2 = 90 - 50 * ny%
      lt1ch$ = LTrim$(RTrim$(Str$(Abs(lt2))))
      If lt2 > 0 Then
         ns$ = "N"
      Else
         ns$ = "S"
         End If
      NROWS = 6000
      NCOLS = 4800
      filt$ = EW$ + lg1ch$ + ns$ + lt1ch$
   Else 'Antartic - Cd #5
      nx% = Fix((endlog + 180) / 60)
      lg2 = -180 + nx% * 60
      If Abs(lg2) >= 100 Then
         lg1ch$ = LTrim$(RTrim$(Str$(Abs(lg2))))
      ElseIf Abs(lg2) < 100 And Abs(lg2) <> 0 Then
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg2))))
      ElseIf Abs(lg2) = 0 Then
         lg1ch$ = "000"
         End If
      If lg2 <= 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ns$ = "S"
      lt2 = -60
      lt1ch$ = "60"
      filt$ = EW$ + lg1ch$ + ns$ + lt1ch$
      NROWS = 3600
      NCOLS = 7200
      End If

    If filt$ <> filt1$ Then
'       ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'       response = MsgBox("DTM elevation data are found on different tiles, do you wish to proceed?", vbYesNo + vbInformation, "Maps & More")
'       ret = SetWindowPos(mapPictureform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'       If response = vbYes Then
          manytiles = True
          Return
          End If
'       Screen.MousePointer = vbDefault
'       GoTo mm999
'       End If
    Return

findtiles:
   If kmy > -60 Then
      nx% = Fix((kmx + 180) * 0.025)
      lg1 = -180 + nx% * 40
      If Abs(lg1) >= 100 Then
         lg1ch$ = RTrim$(LTrim$(Str$(Abs(lg1))))
      Else
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
         End If
      If lg1 < 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ny% = Int((90 - kmy) * 0.02)
      lt1 = 90 - 50 * ny%
      lt1ch$ = LTrim$(RTrim$(Str$(Abs(lt1))))
      If lt1 > 0 Then
         ns$ = "N"
      Else
         ns$ = "S"
         End If
      NROWS = 6000
      NCOLS = 4800
      filt1$ = EW$ + lg1ch$ + ns$ + lt1ch$
   Else 'Antartic - Cd #5
      nx% = Fix((kmx + 180) / 60)
      lg1 = -180 + nx% * 60
      If Abs(lg1) >= 100 Then
         lg1ch$ = LTrim$(RTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) < 100 And Abs(lg1) <> 0 Then
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) = 0 Then
         lg1ch$ = "000"
         End If
      If lg1 <= 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ns$ = "S"
      lt1 = -60
      lt1ch$ = "60"
      filt1$ = EW$ + lg1ch$ + ns$ + lt1ch$
      NROWS = 3600
      NCOLS = 7200
      End If

'    If kmy <= -60 Then
'       lt1 = -60
'    ElseIf kmy > -60 And kmy <= -10 Then
'       lt1 = -10
'    ElseIf kmy > -10 And kmy <= 40 Then
'       lt1 = 40
'    ElseIf kmy > 40 Then
'       lt1 = 90
'       End If
'
'    If kmx >= -180 And kmx < -140 Then
'       lg1 = -180
'    ElseIf kmx >= -140 And kmx < -100 Then
'       lg1 = -140
'    ElseIf kmx >= -100 And kmx < -60 Then
'       lg1 = -100
'    ElseIf kmx >= -60 And kmx < -20 Then
'       lg1 = -60
'    ElseIf kmx >= -20 And kmx < 20 Then
'       lg1 = -20
'    ElseIf kmx >= 20 And kmx < 60 Then
'       lg1 = 20
'    ElseIf kmx >= 60 And kmx < 100 Then
'       lg1 = 60
'    ElseIf kmx >= 100 And kmx < 140 Then
'       lg1 = 100
'    ElseIf kmx >= 140 And kmx <= 180 Then
'       lg1 = 140
'       End If
'
'    If lt1 > 0 Then
'       ns$ = "n"
'    Else
'       ns$ = "s"
'       End If
'    ltc$ = LTrim$(RTrim$(Str$(Abs(lt1))))
'    If lg1 > 0 Then
'       we$ = "e"
'    Else
'       we$ = "w"
'       End If
'    lgc$ = LTrim$(RTrim$(Str$(Abs(lg1))))
'    If Abs(lg1) < 100 Then we$ = we$ + "0"
'    filt1$ = we$ + lgc$ + ns$ + ltc$
    Return


dtmfiles:
       'open header files and read elevation limits and step sizes
        hdrfil% = FreeFile
        Open filn$ + ".STX" For Input As #hdrfil%
        Input #hdrfil%, T1#, elevmin%, elevmax%, D1#, e1#
        Close #hdrfil%
        Open filn$ + ".HDR" For Input As #hdrfil%
        npos% = 0
        Do Until EOF(hdrfil%)
           npos% = npos% + 1
           Line Input #hdrfil%, doclin$
           If npos% = 3 Then
              NROWS% = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
           ElseIf npos% = 4 Then
              NCOLS% = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
           ElseIf npos% = 13 Then
              XDIM = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
           ElseIf npos% = 14 Then
              YDIM = Val(Mid$(doclin$, 15, Len(doclin$) - 14))
              End If
        Loop
        Close #hdrfil%
        'open DEM file and output BIN file
        filnum1% = FreeFile
        Open filn$ + ".DEM" For Binary As #filnum1% '<<<<>>>>
Return


findCD:
   If kmy > -60 Then
      nx% = Fix((kmx + 180) * 0.025)
      lg1 = -180 + nx% * 40
      If Abs(lg1) >= 100 Then
         lg1ch$ = RTrim$(LTrim$(Str$(Abs(lg1))))
      Else
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
         End If
      If lg1 < 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ny% = Int((90 - kmy) * 0.02)
      lt1 = 90 - 50 * ny%
      lt1ch$ = LTrim$(RTrim$(Str$(Abs(lt1))))
      If lt1 > 0 Then
         ns$ = "N"
      Else
         ns$ = "S"
         End If
      DEMfile0$ = EW$ + lg1ch$ + ns$ + lt1ch$
      DEMfile1$ = worlddtm + ":\" + DEMfile0$ + "\" + DEMfile0$
      DEMfile$ = DEMfile1$ + ".dem"
      NROWS = 6000
      NCOLS = 4800
      numCD% = worldCD%(ny% * 9 + nx% + 1)
   Else 'Antartic - Cd #5
      nx% = Fix((kmx + 180) / 60)
      lg1 = -180 + nx% * 60
      If Abs(lg1) >= 100 Then
         lg1ch$ = LTrim$(RTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) < 100 And Abs(lg1) <> 0 Then
         lg1ch$ = "0" + RTrim$(LTrim$(Str$(Abs(lg1))))
      ElseIf Abs(lg1) = 0 Then
         lg1ch$ = "000"
         End If
      If lg1 <= 0 Then
         EW$ = "W"
      Else
         EW$ = "E"
         End If
      ns$ = "S"
      lt1 = -60
      lt1ch$ = "60"
      DEMfile0$ = EW$ + lg1ch$ + ns$ + lt1ch$
      DEMfile1$ = worlddtm + ":\" + DEMfile0$ + "\" + DEMfile0$
      DEMfile$ = DEMfile1$ + ".dem"
      numCD% = 5
      NROWS = 3600
      NCOLS = 7200
      End If

    ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    ret = SetWindowPos(mapprogressfm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    Screen.MousePointer = vbDefault
    response = MsgBox("Please insert USGS EROS CD#" + LTrim$(Str$(numCD%)) + " and enter OK after it has finished loading", vbInformation + vbOKCancel, "Maps & More")
    If response = vbCancel Then
       Unload mapprogressfm
       Exit Sub
    Else 'wait until find expected file
       On Error GoTo DTMerrreturn
       myfile = Dir(DEMfile$)
       If myfile <> sEmpty Then
          Do Until myfile <> sEmpty
             myfile = Dir(DEMfile$)
          Loop
          End If
       Screen.MousePointer = vbHourglass
       Maps.Picture1.Refresh
       mapPictureform.Refresh
'       ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'       ret = SetWindowPos(mapprogressfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       BringWindowToTop (mapPictureform.hwnd)
       BringWindowToTop (mapprogressfm.hwnd)
       End If
Return


diskerrhandler:
   Screen.MousePointer = vbDefault
   
   If checkdtm = True Then
      GoTo 80
   ElseIf Err.Number = 62 And tmflag% > 0 Then
      'missing parameter due to old eros.tm3 version file format
      Resume Next
   Else
      If AutoProf Then 'continue to next item without error message
         Close
         AutoNum& = AutoNum& - 1 'redo last point
         Exit Sub
         End If
      For i% = 0 To Forms.count - 1
         ret = SetWindowPos(Forms(i%).hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
      Next i%
      response = MsgBox("Encountered error #: " & Trim$(Str$(Err.Number)) & vbLf & _
                        Err.Description & vbLf & _
                        "Want to retry?", vbOKCancel + vbCritical, "Maps & More")
      Resume Next
      If response = vbOK Then
         For i% = 0 To Forms.count - 1
'            ret = SetWindowPos(Forms(i%).hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            BringWindowToTop (Forms(i%).hwnd)
         Next i%
         Resume
      Else
         Close
         Exit Sub
         End If
      End If

    For i% = 23 To 27
       tblbuttons(i%) = 0
       Maps.Toolbar1.Buttons(i%).value = tbrUnpressed
       Maps.Toolbar1.Buttons(i%).Enabled = False
    Next i%

   myfile = Dir(drivjk_c$ + "eros.tm3")
   If myfile <> sEmpty Then Kill drivjk_c$ + "eros.tm3"
   myfile = Dir(ramdrive + ":\*.bin")
   If myfile <> sEmpty Then Kill ramdrive + ":\" + myfile
   myfile = Dir(ramdrive + ":\*.bi1")
   If myfile <> sEmpty Then Kill ramdrive + ":\" + myfile
   Screen.MousePointer = vbDefault
   Exit Sub

DTMerrreturn:
   Resume

End Sub
Sub mapCrossSections()

    Dim LimHeight As Integer, filcross%
   On Error GoTo mapCrossSections_Error

    LimHeight = 0

      'determine if second point is visible from first point
      'and dump crosssection into csection.tmp

      'now determine cross section view angles, heights, as funct. of distance along the cross section
      'first find va
   If world = False Then
       Call casgeo(crosssectionpnt(0, 0), crosssectionpnt(0, 1), lg, lt)
       lg1 = lg
       lt1 = lt
       hgt1 = crosssectionhgt(0)
       Call casgeo(crosssectionpnt(1, 0), crosssectionpnt(1, 1), lg, lt)
       lg2 = lg
       lt2 = lt
       hgt2 = crosssectionhgt(1)
       lg2v = lg2
       lt2v = lt2
   Else
       lg1 = -crosssectionpnt(0, 0)
       lt1 = crosssectionpnt(0, 1)
       hgt1 = crosssectionhgt(0)
       lg2 = -crosssectionpnt(1, 0)
       lt2 = crosssectionpnt(1, 1)
       hgt2 = crosssectionhgt(1)
       lg2v = lg2
       lt2v = lt2
       End If
    
    
    'response = MsgBox("Do you wan't to add standard observer height of 1.6 meters to the first point?", vbYesNo, "Maps & More")
    'If response = vbYes Then
    If ObsHeight = True Then
       hgt1 = hgt1 + 1.6
       End If
    GoSub viewan
    viewang0 = va
    
    If SearchCrossSection Then GoTo cs50 'skip responses for calculating shed for search points
    response = MsgBox("Unobstructed Viewangle = " & LTrim$(Format(va, "##0.000")), vbOKCancel, "Maps & More")
    If response = vbCancel Then
       GoCrossSection = False
       Exit Sub
       End If
       
cs50:
    'now begin calculation of other points along trajectory
    If world = False Then
       runit1 = (crosssectionpnt(1, 0) - crosssectionpnt(0, 0))
       runit2 = (crosssectionpnt(1, 1) - crosssectionpnt(0, 1))
    Else
       X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
       X2 = Cos(lt2 * cd) * Cos(lg2 * cd)
       Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
       Y2 = Cos(lt2 * cd) * Sin(lg2 * cd)
       Z1 = Sin(lt1 * cd)
       Z2 = Sin(lt2 * cd)
       End If
      
    If Not SearchCrossSection Then
       sectnumpnt& = Val(mapCrossSection.txtNumPoints.Text)
    Else 'determine number of way points for calculating shed
       If Not world Then
          'check for obstructions every 10 meters
          sectnumpnt& = Abs(runit1) * 0.1
       Else
          sectnumpnt& = Abs(X2 - X1) * 6371315# * Cos(lt1 * cd) / 30
          End If
       
       If sectnumpn& > 500 Then sectnumpnt& = 500
       End If

    'Close
    If Not SearchCrossSection Then
    
      filcross% = FreeFile
      Open drivjk_c$ + "crossect.tmp" For Output As #filcross%
      Print #filcross%, "Cross Section Graph File"
      If world = False Then
         Print #filcross%, "Begin, End, Number of Points" & "(" & Str(crosssectionpnt(0, 0)) & "," _
                                                               & Str(crosssectionpnt(0, 1)) & "," _
                                                               & Str(crosssectionhgt(0)) & ")," _
                                                      & " (" & Str(crosssectionpnt(1, 0)) & "," _
                                                              & Str(crosssectionpnt(1, 1)) & "," _
                                                              & Str(crosssectionhgt(1)) & ")" & "," _
                                                              & Str(sectnumpnt&)
      Else
         Print #filcross%, "Begin, End, Number of Points" & "(" & Str(-lg1) & "," _
                                                               & Str(lt1) & "," _
                                                               & Str(crosssectionhgt(0)) & ")," _
                                                      & " (" & Str(-lg2) & "," _
                                                              & Str(lt2) & "," _
                                                              & Str(crosssectionhgt(1)) & ")" & "," _
                                                              & Str(sectnumpnt&)
         End If
      End If

    'record first point
    Screen.MousePointer = vbHourglass
    If Not SearchCrossSection Then
       Write #filcross%, Val(Format(azi, "###0.000")), Val(Format(va, "##0.000")), Val(Format(crosssectionpnt(0, 1), "######.000")), Val(Format(crosssectionpnt(0, 0), "#######.000")), Val(Format(0, "#####.0")), Val(Format(hgt1, "####0.0"))
       End If

    obstruct% = 0
    nprogress% = 0
    If Not SearchCrossSection Then
      mapCrossSection.CSectProgressBar.value = nprogress%
      mapCrossSection.CSectProgressBar.Visible = True
      End If
    If world = False Then
       For i& = 1 To sectnumpnt& - 1
         DoEvents
          If GoCrossSection = False Then 'Cancel button pressed
             mapCrossSection.CSectProgressBar.Visible = False
             mapCrossSection.lblObstruction.Caption = sEmpty
             Screen.MousePointer = vbDefault
             Close #filcross%
             Exit Sub
             End If
         coord1 = (i& / (sectnumpnt& - 1)) * runit1 + crosssectionpnt(0, 0)
         coord2 = (i& / (sectnumpnt& - 1)) * runit2 + crosssectionpnt(0, 1)
         'find height at this point
         corx = coord1
         cory = coord2
         Call heights(corx, cory, hgtpnt)
         corx = coord1
         cory = coord2
         'find longitude, latitude at this point
         Call casgeo(corx, cory, lg, lt)
         hgt2 = hgtpnt
         
         '/////////////added 080521/////////////////////////
         'limit negative elevations to level of Dead Sea
         If hgt2 < -430 And LimHeight = 0 Then
               Select Case MsgBox("The elevation is less than the water level of the Dead Sea." _
                                  & vbCrLf & "" _
                                  & vbCrLf & "Do you want to limit the elevation to above that elevation?" _
                                  , vbYesNo Or vbInformation Or vbDefaultButton1, "Below Sea Level")
               
                Case vbYes
                
                   LimHeight = 1 'limit elevation to the Dead Sea's water level
               
                Case vbNo
                
                   LimHeight = -1 'gnore further checks
               
               End Select
               
            ElseIf hgt2 < -430 And LimHeight = 1 Then
               
               hgt2 = -430 'water height of the Dead Sea 080521
            
               End If
         '////////////////////////////////////////////////////////
         
         lg2v = lg
         lt2v = lt
         GoSub viewan
         dista = Sqr((coord1 - crosssectionpnt(0, 0)) ^ 2 + (coord2 - crosssectionpnt(0, 1)) ^ 2)
         If va > viewang0 Then
            obstruct% = 1
            If SearchCrossSection Then GoTo mcs500 'obstruction found so abort further search
            End If
         If Not SearchCrossSection Then
            Write #filcross%, Val(Format(azi, "###0.000")), Val(Format(va, "##0.000")), Val(Format(coord1, "######.000")), Val(Format(coord2, "#######.000")), Val(Format(dista, "#####.0")), Val(Format(hgt2, "####0.0"))
            End If
         
         'increment progressbar
         If Not SearchCrossSection Then
            nprogress2% = Int(i& * 100# / (sectnumpnt& - 1))
            If nprogress2% <> nprogress% Then
               nprogress% = nprogress2%
               mapCrossSection.CSectProgressBar.value = nprogress%
               End If
            End If

       Next i&
       Close #filcross%
    Else
     lt2vo = crosssectionpnt(0, 1)
     lg2vo = -crosssectionpnt(0, 0)
     lt2voo = crosssectionpnt(0, 1)
     lg2voo = -crosssectionpnt(0, 0)
     lt2vf = crosssectionpnt(1, 1)
     lg2vf = -crosssectionpnt(1, 0)
     
     Dim cosang As Double
     If greatcircle = True Then 'earth projection of shortest distance between the two points
       'move along unit vector between point1 and point2.  This will give
       'uniform spacing only for small distances
        
       'beginning point for trajectory line to be drawn on world map
       If sectnumpnt& > 100 Then
          travelnum% = 100
          pntmod& = sectnumpnt& / 100
       Else
          travelnum% = sectnumpnt&
          pntmod& = 1
          End If
       ReDim travel(2, 1)
       travel(1, 1) = crosssectionpnt(0, 0)
       travel(2, 1) = crosssectionpnt(0, 1)
       travelnn% = 1
       ipntmod& = pntmod&
       
       dista = 0
       'total distance in radians is:
       Dist = 2 * DASIN(Sqr((Sin((lt1 - lt2) * cd / 2)) ^ 2 + _
            Cos(lt1 * cd) * Cos(lt2 * cd) * (Sin((lg1 - lg2) * cd / 2)) ^ 2))
       'this formula fails for small angles
       'D = cd * atan2((Sin(lt1 * cd) * Sin(lt2 * cd) + Cos(lt1 * cd) * Cos(lt2 * cd) * Cos((lg1 - lg2) * cd)), Sqr((Cos(lt2 * cd) * Sin((lg1 - lg2) * cd)) ^ 2 + (Cos(lt1 * cd) * Sin(lt2 * cd) - Sin(lt1 * cd) * Cos(lt2 * cd) * Cos((lg1 - lg2) * cd)) ^ 2))
       
       For i& = 1 To sectnumpnt& - 1
           DoEvents
          If GoCrossSection = False And Not SearchCrossSection Then  'Cancel button pressed
             mapCrossSection.CSectProgressBar.Visible = False
             mapCrossSection.lblObstruction.Caption = sEmpty
             Screen.MousePointer = vbDefault
             Close #filcross%
             Exit Sub
             End If
             
          'fraction traveled
           sn = i& / (sectnumpnt& - 1)
           
           'way points, see Aviation Formulary V1.4
           a = Sin((1 - sn) * Dist) / Sin(Dist)
           b = Sin(sn * Dist) / Sin(Dist)
           xx1 = a * Cos(lt1 * cd) * Cos(lg1 * cd) + b * Cos(lt2 * cd) * Cos(lg2 * cd)
           yy1 = a * Cos(lt1 * cd) * Sin(lg1 * cd) + b * Cos(lt2 * cd) * Sin(lg2 * cd)
           zz1 = a * Sin(lt1 * cd) + b * Sin(lt2 * cd)
           lt2v = atan2(zz1, Sqr(xx1 ^ 2 + yy1 ^ 2)) / cd
           lg2v = atan2(yy1, xx1) / cd
           
'          xx1 = X1 + sn * (X2 - X1)
'          yy1 = Y1 + sn * (Y2 - Y1)
'          zz1 = Z1 + sn * (Z2 - Z1)
'          norm = Sqr(xx1 ^ 2 + yy1 ^ 2 + zz1 ^ 2)
'          xx1 = xx1 / norm
'          yy1 = yy1 / norm
'          zz1 = zz1 / norm
'          If xx1 = 0 Then
'             lg2v = 0 'or 180 'determine this from quadrants
'          Else
'             lg2v = Atn(yy1 / xx1)
'             End If
'          If lg2v / cd = -90 Or lg2v / cd = 90 Then
'             lt2v = DASIN(yy1 / Sin(lg2v)) 'in radians
'          Else
'             lt2v = DACOS(xx1 / Cos(lg2v)) 'in radians
'             End If
'          If lt2v / cd > 90 Then
'             lt2v = pi - lt2v
'          ElseIf lt2v / cd < -90 Then
'             lt2v = -pi - lt2v
'             End If
'
'          'check for continuity
'          lg2v = lg2v / cd
'          lt2v = lt2v / cd
'
'          'optimized discontinuity test for longitudes
'
'          'check for crossing of poles and date line
'          Dim xarray(4) As Double
'          xarray(1) = Abs(lg2v - lg2vo)
'          xarray(2) = Abs(180 - lg2v - lg2vo)
'          xarray(3) = Abs(lg2v - 180 - lg2vo)
'          xarray(4) = Abs(180 + lg2v - lg2vo)
'          minNum% = MinArray(xarray(), 4)
'          If minNum% = 2 Then
'             lg2v = 180 - lg2v
'          ElseIf minNum% = 3 Then
'             lg2v = lg2v - 180
'          ElseIf minNum% = 4 Then
'             lg2v = 180 + lg2v
'             End If
'          'else lg2v=lg2v
'
'          'If lg2v <= -91 Then
'          '   MinNum% = 6
'          '   GoTo cmc50
'          'ElseIf lg2v >= -89 And lg2v <= 89 Then
'          '   MinNum% = 1
'          '   GoTo cmc50
'          'ElseIf lg2v >= 91 Then
'          '   MinNum% = 4
'          '   GoTo cmc50
'          '   End If
'
'          'Dim xarray(6) As Double
'          'xarray(1) = Abs(lg2v - lg2vo)
'          'xarray(2) = Abs(-lg2v - lg2vo)
'          'xarray(3) = Abs(180 - lg2v - lg2vo)
'          'xarray(4) = Abs(lg2v - 180 - lg2vo)
'          'xarray(5) = Abs(-180 - lg2v - lg2vo)
'          'xarray(6) = Abs(180 + lg2v - lg2vo)
'          'MinNum% = MinArray(xarray(), 6)
'cmc50:     If MinNum% = 2 Then
'           '  lg2v = -lg2v
'          'ElseIf MinNum% = 3 Then
'          '   lg2v = 180 - lg2v
'          'ElseIf MinNum% = 4 Then
'          '   lg2v = lg2v - 180
'          ' ElseIf MinNum% = 5 Then
'          '   lg2v = -180 - lg2v
'          'ElseIf MinNum% = 6 Then
'          '   lg2v = 180 + lg2v
'          '   End If
'          'else lg2v=lg2v
'
'          If lg2vf > lg2voo And lg2v > lg2vo And Abs(lg2vf - lg2voo) > 180 Then
'             'discontinuity at 90 degrees after crossing date line eastbound
'             lg2v = 180 - lg2v
'             End If
'          If i& <> 1 And ((lg2vf < lg2voo And lg2v > lg2vo) Or (lg2vf > lg2voo And lg2v < lg2vo)) Then
'             'lack of continuity detected, so fix it
'              If lg2v > 80 And lg2v < 100 And sgn1 < 0 Then
'                 'discontinuity at 90 degrees
'                 If Abs(lg2vf - lg2voo) > 180 Then
'                 Else
'                    lg2v = 180 - lg2v
'                 End If
'              ElseIf lg2v > -110 And lg2v < -80 And sgn1 < 0 Then
'                 'discontinuity at -90 degrees
'                 lg2v = -180 - lg2v
'              ElseIf lg2v > -1 And lg2v < 1 Then
'                 'discontinuity at 0 degrees
'                 lg2v = -lg2v
'                 End If
'              End If
'
'          'look for crossing the date line moving westward
'          sgn2 = 1
'          If i& > 2 And lg2v > 179 And sgn1 <> 0 Then
'             If (lg2v - lg2vo) / sgn1 < 0 Then 'switched directions
'                lg2v = -Abs(lg2v)
'                End If
'             End If
'          'now look for crossing the date line moving eastward
'          If lg2v < -180 Then
'             lg2v = 360 + lg2v
'             If sgn2 = -1 Then
'                sgn2 = 1 'already switched, so don't switch again
'             Else
'                sgn2 = -1 'switched signs, so signal
'                End If
'             End If
'          'now find lt2v
'
'          'optimized discontinuity test for latitudes
'          If Abs(-lt2v - lt2vo) < Abs(lt2v - lt2vo) Then
'             lt2v = -lt2v
'             End If
'          'xarray(1) = Abs(lt2v - lt2vo)
'          'xarray(2) = Abs(-lt2v - lt2vo)
'          'xarray(3) = Abs(180 - lt2v - lt2vo)
'          'xarray(4) = Abs(lt2v - 180 - lt2vo)
'          'xarray(5) = Abs(-180 - lt2v - lt2vo)
'          'xarray(6) = Abs(180 + lt2v - lt2vo)
'          'MinNum% = MinArray(xarray(), 2)
'          'MinNum2% = MinNum%
'          'If MinNum% = 2 Then
'          '   lt2v = -lt2v
'          'ElseIf MinNum% = 3 Then
'          '   lt2v = 180 - lt2v
'          'ElseIf MinNum% = 4 Then
'          '   lt2v = lt2v - 180
'          'ElseIf MinNum% = 5 Then
'          '   lt2v = -180 - lt2v
'          'ElseIf MinNum% = 6 Then
'          '   lt2v = 180 + lt2v
'          '   End If
'          'else lt2v=lt2v
'          If i& <> 1 And (lt2vf < lt2voo And lt2v > lt2vo) Or (lt2vf > lt2voo And lt2v < lt2vo) Then
'             'discontinuity at 0 degrees, so fix it
'             If lt2v < 1 And lt2v > -1 Then
'                lt2v = -lt2v
'                End If
'             End If
'
'          'Now calculate distance assuming spherical geoid.
'          'For small increments in angle, the distance on the
'          'sphere is approximately equal to the distance between
'          'the vectors of the present and previous points
'          X1X = Cos(lt2vo * cd) * Cos(lg2vo * cd)
'          X2X = Cos(lt2v * cd) * Cos(lg2v * cd)
'          Y1Y = Cos(lt2vo * cd) * Sin(lg2vo * cd)
'          Y2Y = Cos(lt2v * cd) * Sin(lg2v * cd)
'          Z1Z = Sin(lt2vo * cd)
'          Z2Z = Sin(lt2v * cd)
'          'shortest geodesic distance is Re * Angle between vectors
'          'cos(Angle between unit vectors) = Dot product of unit vectors
'          cosang = X1X * X2X + Y1Y * Y2Y + Z1Z * Z2Z
'          dista = dista + 6371315 * DACOS(cosang)
          dista = 6371315 * sn * Dist
          'now find the height at lg2,lt2
          lgtmp = -lg2v: lttmp = lt2v
           If bAirPath Then 'air path calculations so don't give heights
           Else
              Call worldheights(lgtmp, lttmp, hgt2)
              End If
           If hgt2 = -9999 Then hgt2 = 0
           GoSub viewan 'find view angle
           If va > viewang0 Then
              obstruct% = 1
              End If

           Write #filcross%, Val(Format(azi, "###0.000")), Val(Format(va, "##0.000")), Val(Format(lt2v, "######.000")), Val(Format(-lg2v, "#######.000")), Val(Format(dista, "#####.0")), Val(Format(hgt2, "####0.0"))
'           sgn1 = sgn2 * (lg2v - lg2vo)
'           lt2vo = lt2v
'           lg2vo = lg2v
           If i& Mod ipntmod& = 0 Then
              ipntmod& = ipntmod& + pntmod&
              travelnn% = travelnn% + 1
              ReDim Preserve travel(2, travelnn%)
              travel(1, travelnn%) = -lg2v
              travel(2, travelnn%) = lt2v
              End If
           
         nprogress2% = Int(i& * 100# / (sectnumpnt& - 1))
         If nprogress2% <> nprogress% Then
            nprogress% = nprogress2%
            If Not SearchCrossSection Then mapCrossSection.CSectProgressBar.value = nprogress%
            End If
       
       Next i&
       If travelnn% < 99 Then
          travelnum% = travelnn% + 1
          End If
       'end points for trajectory line
       ReDim Preserve travel(2, travelnum%)
       travel(1, travelnum%) = crosssectionpnt(1, 0)
       travel(2, travelnum%) = crosssectionpnt(1, 1)
       
     Else 'follow line drawn on mercator projection
       'travel(points) are just straight line between beginning and end points
       
       'determine which is larger slope
       If lg2vf = lg2voo Then
          slopetype% = 0 'latitudes
'          lt2v = (lt2vf - lt2voo) * i& / sectnumpnt& + lt2voo
'          lg2v = lg2voo
       ElseIf lt2vf = lt2voo Then
          slopetype% = 1 'longitudes
'          lg2v = (lg2vf - lg2voo) * i& / sectnumpnt& + lg2voo
'          lt2v = lt2voo
       Else 'determine smallest slope and use this
          aslopelat = Abs((lg2vf - lg2voo) / (lt2vf - lt2voo))
          aslopelon = Abs((lt2vf - lt2voo) / (lg2vf - lg2voo))
          If slopelat > slopelon Then
             slopetype% = 2
             Slope = (lg2vf - lg2voo) / (lt2vf - lt2voo)
'             lt2v = (lt2vf - lt2voo) * i& / sectnumpnt& + lt2voo
'             lg2v = slope * (lt2v - lt2voo) + lg2voo
          Else
             slopetype% = 3
             Slope = (lt2vf - lt2voo) / (lg2vf - lg2voo)
'             lg2v = (lg2vf - lg2voo) * i& / sectnumpnt& + lg2voo
'             lt2v = slope * (lg2v - lg2voo) + lt2voo
             End If
          End If
       dista = 0
       lt2vo = lt1
       lg2vo = lg1
       For i& = 1 To sectnumpnt&
          DoEvents
          If GoCrossSection = False And Not SearchCrossSection Then 'Cancel button pressed
             mapCrossSection.CSectProgressBar.Visible = False
             mapCrossSection.lblObstruction.Caption = sEmpty
             Screen.MousePointer = vbDefault
             Close #filcross%
             Exit Sub
             End If
          If slopetype% = 0 Then
             lt2v = (lt2vf - lt2voo) * i& / sectnumpnt& + lt2voo
             lg2v = lg2voo
          ElseIf slopetype% = 1 Then
             lg2v = (lg2vf - lg2voo) * i& / sectnumpnt& + lg2voo
             lt2v = lt2voo
          ElseIf slopetype% = 2 Then
             lt2v = (lt2vf - lt2voo) * i& / sectnumpnt& + lt2voo
             lg2v = Slope * (lt2v - lt2voo) + lg2voo
          ElseIf slopetype% = 3 Then
             lg2v = (lg2vf - lg2voo) * i& / sectnumpnt& + lg2voo
             lt2v = Slope * (lg2v - lg2voo) + lt2voo
             End If
          'now calculate distance assuming spherical geoid
          X1X = Cos(lt2vo * cd) * Cos(lg2vo * cd)
          X2X = Cos(lt2v * cd) * Cos(lg2v * cd)
          Y1Y = Cos(lt2vo * cd) * Sin(lg2vo * cd)
          Y2Y = Cos(lt2v * cd) * Sin(lg2v * cd)
          Z1Z = Sin(lt2vo * cd)
          Z2Z = Sin(lt2v * cd)
          'shortest geodesic distance is Re * Angle between vectors
          'cos(Angle between unit vectors) = Dot product of unit vectors
          cosang = X1X * X2X + Y1Y * Y2Y + Z1Z * Z2Z
          dista = dista + 6371315 * DACOS(cosang)
          'find heights
          lgtmp = -lg2v: lttmp = lt2v
          Call worldheights(lgtmp, lttmp, hgt2)
          If hgt2 = -9999 Then hgt2 = 0
          GoSub viewan 'find view angle
          If va > viewang0 Then
             obstruct% = 1
             If SearchCrossSection Then GoTo mcs500 'obstruction found so abort further search
             End If
           
         If Not SearchCrossSection Then
            Write #filcross%, Val(Format(azi, "###0.000")), Val(Format(va, "##0.000")), Val(Format(lt2v, "######.000")), Val(Format(-lg2v, "#######.000")), Val(Format(dista, "#####.0")), Val(Format(hgt2, "####0.0"))
            End If
         nprogress2% = Int(i& * 100# / sectnumpnt&)
         If nprogress2% <> nprogress% Then
            nprogress% = nprogress2%
            If Not SearchCrossSection Then mapCrossSection.CSectProgressBar.value = nprogress%
            End If
         lt2vo = lt2v
         lg2vo = lg2v
       Next i&
     End If
     If Not SearchCrossSection Then Close #filcross%
    End If
    Screen.MousePointer = vbDefault
    If Not SearchCrossSection Then mapCrossSection.CSectProgressBar.Visible = False
    
mcs500:

    If SearchCrossSection Then
       SearchCrossObstruct = False
       If obstruct% = 1 Then SearchCrossObstruct = True
       Exit Sub 'go back to calculating other search points
       End If

    If obstruct% = 1 Then
'       response = MsgBox("Second point is obstructed", vbOKOnly, "Maps & More")
        mapCrossSection.lblObstruction.Caption = "First point is obstructed"
    Else
'       response = MsgBox("Second point is ***NOT*** obstructed", vbOKOnly, "Maps & More")
        mapCrossSection.lblObstruction.Caption = "First point is not obstructed"
        End If
    
    GoCrossSection = False
    If bAirPath Then 'just show route on map
       crosssection = True
       'draw out the cross section line on the map
       Call showtheroute(mapPictureform.mapPicture)
       Exit Sub
       End If

    response = MsgBox("Display the cross section?", vbYesNoCancel, "Maps & More")
    mapCrossSection.lblObstruction.Caption = sEmpty
    If response = vbYes Then
       'Unload mapCrossSection
       'Set mapCrossSection = Nothing
       crosssection = True
       'draw out the cross section line on the map
       Call showtheroute(mapPictureform.mapPicture)
       'graph out the detailed height vs dist profile
       mapCrossSection.Visible = False
       mapgraphfm.Visible = True
       End If

   Exit Sub

viewan:
     X1v = Cos(lt1 * cd) * Cos(lg1 * cd)
     X2v = Cos(lt2v * cd) * Cos(lg2v * cd)
     Y1v = Cos(lt1 * cd) * Sin(lg1 * cd)
     Y2v = Cos(lt2v * cd) * Sin(lg2v * cd)
     Z1v = Sin(lt1 * cd)
     Z2v = Sin(lt2v * cd)
     Re = 6371315#
     re1 = (hgt1 + Re)
     re2 = (hgt2 + Re)
     X1v = re1 * X1v
     Y1v = re1 * Y1v
     Z1v = re1 * Z1v
     X2v = re2 * X2v
     Y2v = re2 * Y2v
     Z2v = re2 * Z2v
     dist1 = re1
     dist2 = re2
     Angle = DACOS((X1v * X2v + Y1v * Y2v + Z1v * Z2v) / (dist1 * dist2))
     viewang = Atn((-re1 + re2 * Cos(Angle)) / (re2 * Sin(Angle)))
     va = viewang / cd
     D = (dist1 - dist2 * Cos(Angle)) / dist1
     x1d = X1v * (1 - D) - X2v
     y1d = Y1v * (1 - D) - Y2v
     z1d = Z1v * (1 - D) - Z2v
     'x1p = -Y1
     'y1p = X1
     'z1p = 0
     'azicos = (x1p * x1d + y1p * y1d) / Sqr(X1 ^ 2 + Y1 ^ 2)
     x1p = -Sin(lg1 * cd)
     y1p = Cos(lg1 * cd)
     z1p = 0
     azicos = (x1p * x1d + y1p * y1d)
     x1s = -Cos(lg1 * cd) * Sin(lt1 * cd)
     y1s = -Sin(lg1 * cd) * Sin(lt1 * cd)
     z1s = Cos(lt1 * cd)
     azisin = (x1s * x1d + y1s * y1d + z1s * z1d)
     azi = Atn(azisin / azicos)
     azi = azi / cd
Return

   On Error GoTo 0
   Exit Sub

mapCrossSections_Error:
    
    If Err.Number = 54 Then
        'start over
        Close #filcross%
        GoTo cs50
        End If
    If filcross% > 0 Then Close #filcross%
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mapCrossSections of Module MapModule"
'    Resume
End Sub
Function MinArray(x() As Double, NumArray As Integer) As Integer
   'returns smallest member of array x() having NumArray members
   MinArray = 1
   xmin = x(1)
   For i% = 2 To NumArray
      If x(i%) < xmin Then
         xmin = x(i%)
         MinArray = i%
         End If
   Next i%
End Function
Function RdHalTrue() As Boolean
   'First look for Windows 2000. Windows 2000 doesn't support
   '16 bit Ramdrive.sys.  So have to write to disk instead.
   If WinVer >= 5 Or WinVer = 261 Then
      RdHalTrue = True
      Exit Function
      End If

   'the three ramdrives are ramdrive, then next 2 letters.
   'the largest ramdrives are always the first two letters.
   '(can also do this using API function: GetDiskFreeSpace,
   'see MSDN article Q147686)
   Screen.MousePointer = vbHourglass
   
   rd1$ = ramdrive
   rd2$ = Chr$(Asc(rd1$) + 1)
   rd1$ = ramdrive & ":"
   rd2$ = rd2$ + ":"
   'now do directory search of these ramdrives
   On Error GoTo errhand
   
   testnum% = 0
   rd$ = "dir " & rd1$ & "\*.* > c:\dirlist.dir"
   
10 filbat% = FreeFile
   Open "dirRam.bat" For Output As #filbat%
   Print #filbat%, rd$
   Close #filbat%
   ret = Shell("dirRam.bat", vbNormalFocus)
   'wait a bit for the directory listing to finish
   waitime = Timer + 0.5
   Do Until Timer > waitime
      DoEvents
   Loop
   fildir% = FreeFile
   Open "c:\dirlist.dir" For Input As #fildir%
   Do Until EOF(fildir%)
      Line Input #fildir%, doclin$
   Loop
   Close #fildir%
   'now determine number of free bytes
   pos1% = InStr(doclin$, "dir(s)") + 6
   pos2% = InStr(doclin$, "bytes")
   'Remove apostrophes in order to make number
   numch$ = sEmpty
   For i% = pos1% To pos2% - 1
       CH$ = Mid$(doclin$, i%, 1)
       If CH$ <> "," Then
          numch$ = numch$ + CH$
       End If
   Next
   numbytes& = Val(numch$)
   If numbytes& >= 32000000 Then 'there's enough RAM memory
      'usually 32MB of RAM is enough, so take a change since
      'Pentium 233 can only spare 32MB
      'If testnum% = 0 Then
      '   rd$ = "dir " & rd2$ & "\*.* > dirlist.dir"
      '   testnum% = 1
      '   GoTo 10
      'Else
         RdHalTrue = True
      'End If
    Else
       RdHalTrue = False
    End If
    
    On Error Resume Next
    Kill "c:\dirlist.dir"
    Screen.MousePointer = vbDefault
    
   Exit Function
errhand:
   RdHalTrue = False
   Screen.MousePointer = vbDefault

End Function
Sub EYsunrisesunset(Mode%)
   On Error GoTo errcancel
   
   Dim lon As Double, lat As Double, kmx As Double, kmy As Double, TempSet As Boolean
   Dim MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer, ier As Integer
   
   sunmode% = Mode%

  'add the current place to the placlist.txt file
   Screen.MousePointer = vbHourglass
   plac$ = drivjk_c$ + "placlist.txt"
   myfile = Dir(plac$)
   filplac% = FreeFile
   If myfile = sEmpty Then
      Open plac$ For Output As #filplac%
   Else
      Open plac$ For Append As #filplac%
      End If
      
   If AutoProf = False Then
      response = InputBox("Input place name", "New EY profile", "000-xxxx", Maps.Picture1.Width / 2, Maps.Picture1.Height / 2)
   ElseIf AutoProf = True Then
      response = InputBox("Input template name", "EY Template Name", "000-xxxx", Maps.Picture1.Width / 2, Maps.Picture1.Height / 2)
      End If
   ln1% = Len(Trim$(response))
   If ln1% >= 20 Then
      txt1$ = "'" + Mid$(Trim$(response), 1, 20) + "'"
   Else
      txt1$ = "'" + Trim$(response) + String(20 - ln1%, " ") + "'"
      End If
   txt3$ = "0"
   
   If AutoProf = False Then
   
      txt1$ = txt1$ + "," + Trim$(Maps.Text5.Text * 0.001) + "," + _
      Trim$((Maps.Text6.Text - 1000000) * 0.001) + "," + _
      Trim$(Maps.Text7.Text) + ",1,0," + txt3$
       
      Print #filplac%, txt1$
      Close #filplac%
      Screen.MousePointer = vbDefault
      
    Else 'using coordinates found in mapsearchfm
       If Dir(drivjk_c$ & "mappoints.sav") <> sEmpty Then
          'determine number of rows
           savfil% = FreeFile
           Open drivjk_c$ & "mappoints.sav" For Input As #savfil%
           Do Until EOF(savfil%)
              Input #savfil%, savlat, savlon, savhgt, savdis
              txt0$ = txt1$ + "," + Trim$(Str$(savlat * 0.001)) + "," + _
              Trim$(Str$((savlon - 1000000) * 0.001)) + "," + _
              Trim$(Str$(savhgt)) + ",1,0," + txt3$
       
              Print #filplac%, txt0$
           Loop
           Close #savfil%
           Close #filplac%
           AutoProf = False
           AutoVer = False
           Screen.MousePointer = vbDefault
       Else
           MsgBox "Can't read the saved map points", vbOKOnly + vbCritical, "Maps&More"
           AutoProf = False
           AutoVer = False
           Exit Sub
           End If
          
   End If
   
   'Display the placlist.txt file with check boxes.
   FileCopy drivjk_c$ + "placlist.txt", drivjk_c$ + "viewin.tmp"
   FileViewName = drivjk_c$ + "placlist.txt"
   mapFileViewfm.Visible = True
   FileView = True
   'Now wait until it is closed
   Do Until Not FileView
      DoEvents
   Loop
   If Not FileViewError And FileEdit Then
      'first back up old placlist.txt
      If Dir(drivjk_c$ + "placlist.bak") <> sEmpty Then
         filbak% = FreeFile
         Open drivjk_c$ + "placlist.bak" For Append As #filbak%
         filplac% = FreeFile
         Open drivjk_c$ + "placlist.txt" For Input As #filplac%
         Do Until EOF(filplac%)
            Line Input #filplac%, doclin$
            Print #filbak%, doclin$
         Loop
         Close #filbak%
         Close #filplac%
         End If
         
      FileCopy drivjk_c$ + "viewout.tmp", drivjk_c$ + "placlist.txt"
   Else
      If FileEdit Then GoTo errcancel
      End If
   
   'now add directory information and write to \jk\placlis2.txt
   If sunmode% = 1 Then
      netzskiy$ = "\netz\"
   Else
      netzskiy$ = "\skiy\"
      End If
      
   filin% = FreeFile
   Open drivjk_c$ & "placlist.txt" For Input As #filin%
   filout% = FreeFile
   Open drivjk_c$ & "placlis2.txt" For Output As #filout%
   nsuffix% = 1
   uniqroot$ = sEmpty
   UniqueRoots% = 0
   ReDim FileViewFileName(7, UniqueRoots%)
   ReDim FileViewFileType(UniqueRoots%)
   AbrevDir$ = sEmpty
   Do Until EOF(filin%)
      Input #filin%, doc1$, kmx, kmy, hgt, N1, n2, n3  'find way to add MeanTEmp <<<<<<<<<<<<<<<<<<<
      If sunmode% = 1 Then
         s1 = kmx
         e1 = 260
         n4 = 0
      Else
         s1 = 60
         e1 = kmx
         n4 = 1
         End If
      'find abbreviated form of city directory for use
      'with Fortran programs
      If AbrevDir$ <> sEmpty Then GoTo ab1 'already found
      If Dir(drivcities$ & "comp") = sEmpty Then
         resp = InputBox("Can't find the city file: ""comp"" that contains" & _
                "the 8 letter abbreviated (old DOS) names of the cities." & vbLf & vbLf & _
                "Input the abbreviated form, e.g., ""jerusa~2""" & _
                "instead of ""jerusalem"".", "Input the abbreviated city name", sEmpty)
         If resp = sEmpty Then Exit Sub 'user canceled
         AbrevDir$ = resp
      Else 'open it and determined the abbreviated city name
         compfil% = FreeFile
         Open drivcities$ & "comp" For Input As #compfil%
         nn% = 0
         AbrevDir$ = sEmpty
         Do Until EOF(compfil%)
            Line Input #compfil%, complin$
            nn% = nn% + 1
            If nn% > 7 Then
               If Len(complin$) > 45 Then
                  newdir$ = Mid$(complin$, 45, Len(complin$) - 44)
                  If LCase$(FileViewDir$) = drivcities$ & LCase$(Mid$(complin$, 45, Len(complin$) - 44)) Then
                     AbrevDir$ = Mid$(complin$, 1, 8)
                     Exit Do
                     End If
                  End If
               End If
         Loop
         Close #compfil%
ab1:     If AbrevDir$ = sEmpty Then 'ask user to input it
            resp = InputBox("The abbreviated city name doesn't exist!" & vbLf & vbLf & _
                   "Enter the 8 letter abbreviated (old DOS) name of the city directory." & vbLf & vbLf & _
                   "For example: ""jerusa~2"" instead of ""jerusalem"".", _
                   "Input the abbreviated city name", sEmpty)
            If resp = sEmpty Then Exit Sub 'user canceled
            AbrevDir$ = resp
            End If
         End If
      doc2$ = drivcities$ & Trim$(AbrevDir$) & netzskiy$
      'check if directory really exists
      If Dir(doc2$, vbDirectory) = sEmpty Then
         AbrevDir$ = sEmpty
         GoTo ab1
         End If
      Print #filout%, "'" & doc2 & "'"
      Print #filout%, "'" & doc2 & "',0,' '"
      'determine unique file name
f50:  If Mid$(doc1$, 2, 8) <> uniqroot$ Then
         nsuffix% = 1 'started new root name
         extnum% = 0
      Else
         nsuffix% = nsuffix% + 1
         End If
      uniqroot$ = Mid$(doc1$, 2, 8)
      If nsuffix% < 10 Then
         ext$ = ".pr" & Trim$(CStr(nsuffix%))
      ElseIf nsuffix% >= 10 And nsuffix% < 100 Then
         ext$ = ".p" & Trim$(CStr(nsuffix%))
      ElseIf nsuffix% >= 100 And nsuffix% < 1000 Then
         ext$ = "." & Trim$(CStr(nsuffix%))
      Else
         MsgBox "More than 999 profiles in directory!--abort", vbCritical + vbOKOnly, "Maps&More"
         Close #filin%
         Close #filout%
         Exit Sub
         End If
         
        'now determine temperature for terrestrial refraction
         If Not TempSet Then
         If ggpscorrection = True Then 'apply conversion from Clark geoid to WGS84
            Dim N As Long
            Dim E As Long
            N = kmy * 1000 + 1000000
            E = kmx * 1000
            Call ics2wgs84(N, E, lat, lon)
         Else
            Call casgeo(kmx * 1000, kmy * 1000 + 1000000, lon, lat)
            lon = -lon
            End If
            
          Call Temperatures(lat, lon, MinT, AvgT, MaxT, ier)
          MeanTemp = 0
          For ii = 1 To 12
            MeanTemp = AvgT(ii) + MeanTemp
          Next ii
          MeanTemp = MeanTemp / 12
          End If
             
'          Select Case Mode%
'             Case Is >= 1 'sunrise
'                'use averaged minimum temperature
'                MeanTemp = 0
'                For ii = 1 To 12
'                   MeanTemp = MT(ii) + MeanTemp
'                Next ii
'                MeanTemp = MeanTemp / 12
'             Case Is <= 0 'sunset
'                'use averaged average temperature
'                MeanTemp = 0
'                For ii = 1 To 12
'                   MeanTemp = AT(ii) + MeanTemp
'                Next ii
'                MeanTemp = MeanTemp / 12
'             Case Else
'            End Select
          
          'remove old MeanTemp.txt file
          If Dir(drivjk_c$ & "TRLapseRate.txt") <> sEmpty Then Kill drivjk_c$ & "TRLapseRate.txt"
          If Dir(drivjk_c$ & "MeanTemp.txt") <> sEmpty Then Kill drivjk_c$ & "MeanTemp.txt"
          
          If TemperatureModel% > 2 And Not TempSet Then
EY100:
             NewMeanTemp$ = InputBox("Enter ground temperature (deg C)", "Ground Temperature", Val(Format(Str$(MeanTemp), "###0.0")))
             If Val(NewMeanTemp$) <> MeanTemp Then
                Select Case MsgBox("The entered ground temperature (deg C) is equal to: " & NewMeanTemp$ _
                                   & vbCrLf & "" _
                                   & vbCrLf & "Is this correct?" _
                                   , vbYesNo Or vbQuestion Or vbDefaultButton1, "New ground temperature")
                
                    Case vbYes
                        MeanTemp = Val(NewMeanTemp$)
                        If MeanTemp < -30 Or MeanTemp > 40 Then
                           Select Case MsgBox("The suggested range of temperatures is from -30C to 40C." _
                                              & vbCrLf & "You inputed: " & Str$(MeanTemp) _
                                              & vbCrLf & "Do you want to keep your inputed value?" _
                                              , vbYesNo Or vbInformation Or vbDefaultButton2, "Ground temperature")
                           
                            Case vbYes
                           
                            Case vbNo
                              GoTo EY100
                           End Select
                           End If
                           
                        TempSet = True

                    Case vbNo
                        'try again
                        GoTo EY100
                End Select
                End If
             End If
       
      proFile$ = doc2$ & uniqroot$ & ext$
      If Trim$(proFile$) = sEmpty Then GoTo ey500
      myfile = Dir(proFile$)
      If Dir(doc2$ & uniqroot$ & ext$) <> sEmpty Then
         'determine new extension
         GoTo f50
      Else
         filn$ = uniqroot$ & ext$
         UniqueRoots% = UniqueRoots% + 1
         ReDim Preserve FileViewFileName(7, UniqueRoots%)
         ReDim Preserve FileViewFileType(UniqueRoots%)
         'record file names
         FileViewFileName(0, UniqueRoots% - 1) = filn$
         'record each extension
         FileViewFileName(1, UniqueRoots% - 1) = Mid$(ext$, 2, 3)
         'record each file type (netz or skiy)
         If sunmode% = 1 Then
           FileViewFileType(UniqueRoots% - 1) = 1 'sunrise begins at 1
         Else
           FileViewFileType(UniqueRoots% - 1) = -4 'sunset begins at 4
           End If
         
         Print #filout%, "'" & filn$ & String(8, " ") & "'" & "," & _
               CStr(kmx) & "," & CStr(kmy) & "," & CStr(hgt) & "," & _
               CStr(N1) & "," & CStr(n2) & "," & CStr(n3) & "," & _
               CStr(s1) & ","; CStr(e1) & "," & CStr(n4)
         End If
ey500:
   Loop
   
   Close #filin%
   Close #filout%
   
   If TemperatureModel% > 2 Then 'added 010922
      'record the MeanTemp to the PlacList.txt file
      'write a MeanTemp file
      filmt% = FreeFile
      Open drivjk_c$ & "MeanTemp.txt" For Output As #filmt%
      Write #filmt%, MeanTemp
      Close #filmt%
      End If
   
   If Dir(drivjk_c$ & "scanlist.txt") <> sEmpty Then Kill drivjk_c$ & "scanlist.txt"
   'Then run readlst3.exe
   ret = Shell(drivjk_c$ + "Readlst3.exe", vbNormalFocus)
   'give it 5 second to complete (in case the hard disk spinned down)
   Screen.MousePointer = vbHourglass
   waitime = Timer + 5
   Do Until Timer > waitime
      DoEvents
   Loop
   'Then display scanlist.txt for editing without check boxes
   FileCopy drivjk_c$ + "scanlist.txt", drivjk_c$ + "viewin.tmp"
   FileViewName = drivjk_c$ + "scanlist.txt"
   mapFileViewfm.Visible = True
   ret = ShowWindow(mapFileViewfm.hwnd, 1)
   FileView = True
   Screen.MousePointer = vbDefault
   FileView = True
   'Now wait until it is closed
   Do Until Not FileView
      DoEvents
   Loop
   
   If FileViewError Then Exit Sub
   
   
   'now dump the edited scanlist to scanlist.txt, add the mean temperature, and
   'start rdhal2
   If Not FileEdit Then GoTo rhal 'no editing requested
   filtmp1% = FreeFile
   Open drivjk_c$ & "viewout.tmp" For Input As #filtmp1%
   filtmp2% = FreeFile
   Open drivjk_c$ & "scanlist.txt" For Output As #filtmp2%
   'find carriage returns and rebuild scanlist.txt
   Do Until EOF(filtmp1%)
      Line Input #filtmp1%, docline$
      doclin$ = sEmpty
      For i% = 1 To Len(docline$)
          CH$ = Mid$(docline$, i%, 1)
          If (CH$ <> vbNewLine) And (CH$ <> vbLf) And (CH$ <> vbCrLf) Then
             doclin$ = doclin$ & CH$
          Else
             Print #filtmp2%, doclin$
             doclin$ = sEmpty
             End If
      Next i%
      Print #filtmp2%, doclin$ 'print last line
   Loop
   Close #filtmp1%
   Close #filtmp2%
   GoTo rhal
   
  'Restore and count the edited scanlist. Also
  'determine the file names that will be generated
  'by rdhal2, and determine if they are sunrise or
  'sunset or both.
'  If Not FileViewError Then
'     FileCopy drivjk$ + "viewout.tmp", drivjk$ + "scanlist.txt"
'  Else
'     GoTo errcancel
'     End If
'  filtmp% = FreeFile
'  Open drivjk$ + "scanlist.txt" For Input As #filtmp%
'  statusnum% = 0
'  UniqueRoots% = 0
'  ReDim FileViewFileName(7, 5)
'  Do Until EOF(filtmp%)
'     If statusnum% >= 5 Then 'redimension the array
'        ReDim Preserve FileViewFileName(7, statusnum% + 1)
'        End If
'     Line Input #filtmp%, doclin$
'
'     'Determine unique root file names and the extensions
'     '****THIS ASSUMES THAT ALL THE FILES ARE IN ORDER******
'     'i.e.,root.001,root.002,root.003, etc, then new root name, etc
'     pos% = InStr(doclin$, "\PROM") - 2 'starting position of name
'     If statusnum% = 0 Then 'first root name
'        extnum% = 1
'        FileViewFileName(0, UniqueRoots%) = Mid$(doclin$, pos%, 16)
'        'first extension name
'        FileViewFileName(extnum%, UniqueRoots%) = Mid$(doclin$, pos% + 17, 3)
'        'determine it's file extension and the number of them
'        If InStr(doclin$, ".001") <> 0 Or _
'           InStr(doclin$, ".002") <> 0 Or _
'           InStr(doclin$, ".003") <> 0 Then
'           FileViewFileType(UniqueRoots%) = 1 'sunrise begins at 1
'        Else
'           FileViewFileType(UniqueRoots%) = -4 'sunset begins at 4
'           End If
'     Else 'determine if this new filename or just new extension
'        newname% = 0
'        For j% = 0 To UniqueRoots%
'           If Mid$(doclin$, pos%, 16) = FileViewFileName(0, j%) Then
'              'just a new extension of the same name
'              'record the extension and continue
'              extnum% = extnum% + 1
'              FileViewFileName(extnum%, UniqueRoots%) = Mid$(doclin$, pos% + 17, 3)
'              FileViewFileType(UniqueRoots%) = FileViewFileType(UniqueRoots%) + 1
'              newname% = 1
'              Exit For
'              End If
'        Next j%
'        If newname% = 0 Then 'this is a new rootname name
'           UniqueRoots% = UniqueRoots% + 1
'           extnum% = 1
'           FileViewFileName(0, UniqueRoots%) = Mid$(doclin$, pos%, 16)
'           FileViewFileName(extnum%, UniqueRoots%) = Mid$(doclin$, pos% + 17, 3)
'           'determine the number of files and the file type
'           If InStr(doclin$, ".001") <> 0 Or _
'              InStr(doclin$, ".002") <> 0 Or _
'              InStr(doclin$, ".003") <> 0 Then
'              FileViewFileType(UniqueRoots%) = 1 'sunrise begins at 1
'          Else
'               FileViewFileType(UniqueRoots%) = -4 'sunset begins at -4
'            End If
'         End If
'      End If
'
'      statusnum% = statusnum% + 1 'total file count
'   Loop
'   Close #filtmp%
'   Open drivjk$ + "status.txt" For Output As #filtmp%
'   Write #filtmp%, 1, statusnum%
'   Close #filtmp%
'   '-----------------------old methods-----------------
   
rhal:
   'Then run rdhal (maybe modify rdhalbat.for so that it signals where
   'it is holding by writing files giving the x,lowery, upper y
   'coordinates.
   'check status information
   response = MsgBox("Run rdhalba3/4 program to calculate profiles?", vbYesNo + vbQuestion, "Maps&More")
   If response = vbYes Then
      If WinVer >= 5 Or WinVer = 261 Then
      
        '/////////////////////changes of 061120/////////////////////////////////
         'add Terrestrial Refraction Type Calculation flag to each line in scanlist.txt
        filtmp1% = FreeFile
        Open drivjk_c$ & "viewout.tmp" For Input As #filtmp1%
        filtmp2% = FreeFile
        Open drivjk_c$ & "scanlist.txt" For Output As #filtmp2%
        Do Until EOF(filtmp1%)
           Line Input #filtmp1%, doclin$
           Print #filtmp2%, doclin$ & "," & Str$(TemperatureModel%)
        Loop
        Close #filtmp1%
        Close #filtmp2%
        
         'ret = Shell(drivjk$ & "rdhal3.bat", vbNormalFocus)
         ret = WinExec(drivjk_c$ & "rdhalba4.exe", SW_SHOWNORMAL)
      Else
         response = MsgBox("(1) Run rdhalba3 using ramdrives?" & vbLf & _
                         "(2) Run rdhalba4 from a hard drive?" & vbLf & vbLf & _
                         "For option (1)-Answer ""Yes""" & vbLf & _
                         "For option (2)-Answer ""No""", _
                         vbYesNoCancel + vbQuestion, "Maps&More")
         Select Case response
            Case vbYes
               ret = Shell(drivjk_c$ & "rdhal2.bat", vbNormalFocus)
            Case vbNo
               ret = WinExec(drivjk_c$ & "rdhalba4.exe", SW_SHOWNORMAL)
            Case Else
               Close
               Exit Sub
          End Select
         End If
      End If
   'Now present one file root at a time and ask what
   'to do with it.  This uses the VB version of analyze.bas
   mapAnalyzefm.Visible = True
'   ret = SetWindowPos(mapAnalyzefm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (mapAnalyzefm.hwnd)
   
   'Then graph the results, and then the user can
   'push the calendar button and the profile will be
   'written to the city directory/netz or skiy.
   
   tblbuttons(26) = 0
   tblbuttons(27) = 0
   Maps.Toolbar1.Buttons(26).value = tbrUnpressed
   Maps.Toolbar1.Buttons(27).value = tbrUnpressed
   Exit Sub
   
errcancel:
   Screen.MousePointer = vbDefault
   Close
   tblbuttons(26) = 0
   tblbuttons(27) = 0
   Maps.Toolbar1.Buttons(26).value = tbrUnpressed
   Maps.Toolbar1.Buttons(27).value = tbrUnpressed
   MsgBox "Encountered error number: " & CStr(Err.Number) & vbLf & _
          Err.Description, vbExclamation + vbOKOnly, "Maps& More"

End Sub
Function AVREF(deltd, distd) As Double
    Dim M As Double, b As Double
    If deltd <= 0 Then
       M = 0.000782 - deltd * 0.000000311
       b = -0.0141 + deltd * 0.000034
    ElseIf deltd > 0 Then
       M = 0.000764 + deltd * 0.000000309
       b = -0.00915 - deltd * 0.0000269
       End If
    AVREF = M * distd + b
    If AVREF < 0 Then AVREF = 0
End Function


'convert screen physical coordinates to geo coordinates
Sub ScreenToGeo(dragCoordX, dragCoordY, kmxDrag, kmyDrag, Mode%, ier%)
  'mode=1 'convert screen coordinates to geo
  'mode=2 'convert geo to screen coordinates
  
  On Error GoTo errhand
  
  ier% = 0
  
  If Mode% = 1 Then 'convert (dragCoordX,dragCoordY) -> (kmxDrag,kmyDrag)

      x = dragCoordX
      y = dragCoordY
      If world = True Then GoTo m10
      If map400 = True Then
         If mag > 1 Then
            kmxcc = kmxc
            kmycc = kmyc
            xo = kmxcc - (km400x / mag) * (mapwi2 - mapxdif) * 0.5
            yo = kmycc + (km400y / mag) * (maphi2 - mapydif) * 0.5
            kmxDrag0 = Fix(xo + x * km400x / mag) 'mapdif accounts for size of frame around picture
            kmyDrag0 = Fix(yo - y * km400y / mag)
         Else
            kmxcc = kmxc + (km400x) * (mapwi - mapwi2 + mapxdif) / 2
            kmycc = kmyc - (km400y) * (maphi - maphi2 + mapydif) / 2
            'middle of screen corresponds to kmxc,kmyc
            'so topleft corner=origin corresponds to:
            xo = kmxcc - km400x * sizex / 2  'mapPictureform.mapPicture.Width / 2
            yo = kmycc + km400y * sizey / 2 'mapPictureform.mapPicture.Height / 2
            kmxDrag0 = Fix(xo + x * km400x)   'mapdif accounts for size of frame around picture
            kmyDrag0 = Fix(yo - y * km400y)
            End If
       ElseIf map50 = True Then
         If mag > 1 Then
            kmxcc = kmxc
            kmycc = kmyc
            xo = kmxcc - (km50x / mag) * (mapwi2 - mapxdif) * 0.5
            yo = kmycc + (km50y / mag) * (maphi2 - mapydif) * 0.5
            kmxDrag0 = Fix(xo + x * km50x / mag) 'mapdif accounts for size of frame around picture
            kmyDrag0 = Fix(yo - y * km50y / mag)
         Else
            kmxcc = kmxc + (km50x) * (mapwi - mapwi2 + mapxdif) / 2
            kmycc = kmyc - (km50y) * (maphi - maphi2 + mapydif) / 2
            xo = kmxcc - km50x * sizex / 2  'mapPictureform.mapPicture.Width / 2
            yo = kmycc + km50y * sizey / 2 'mapPictureform.mapPicture.Height / 2
            kmxDrag0 = Fix(xo + x * km50x)
            kmyDrag0 = Fix(yo - y * km50y)
            End If
         End If
m10:     Select Case coordmode%
           Case 1 'ITM
              kmxDrag = kmxDrag0
              kmyDrag = kmyDrag0
           Case 2 'GEO
              If world = True Then
                 If mag > 1 Then
                   lonc = lon '+ fudx / mag
                   latc = lat '+ fudy / mag
                   xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                   yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                   kmxDrag = xo + x * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
                   kmyDrag = yo - y * (deglat / (sizewy * mag))
                 Else
                   lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                   latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                   xo = lonc - deglog / 2
                   yo = latc + deglat / 2
                   kmxDrag = xo + x * (deglog / sizewx)
                   kmyDrag = yo - y * (deglat / sizewy)
                   If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                       'fudge factor for inaccuracy of linear degree approx for large size map
                       kmxDrag = kmxDrag - 0.006906
                       kmyDrag = kmyDrag + 0.003878
                       End If
                   End If
                Else
                   kmxDrag = 0 'option not supported yet
                   kmyDrag = 0
                End If
             Case Else
                kmxDrag = 0 'option not supported yet
                kmyDrag = 0
          End Select

  ElseIf Mode% = 2 Then 'convert (kmxDrag,kmyDrag) -> (dragCoordX,dragCoordY)
  
      kmxDrag0 = kmxDrag
      kmyDrag0 = kmyDrag
      If world = True Then GoTo m100
      If map400 = True Then
         If mag > 1 Then
            kmxcc = kmxc
            kmycc = kmyc
            xo = kmxcc - (km400x / mag) * (mapwi2 - mapxdif) * 0.5
            yo = kmycc + (km400y / mag) * (maphi2 - mapydif) * 0.5
            x = mag * (kmxDrag0 - xo) / km400x 'mapdif accounts for size of frame around picture
            y = mag * (yo - kmyDrag0) / km400y
         Else
            kmxcc = kmxc + (km400x) * (mapwi - mapwi2 + mapxdif) / 2
            kmycc = kmyc - (km400y) * (maphi - maphi2 + mapydif) / 2
            'middle of screen corresponds to kmxc,kmyc
            'so topleft corner=origin corresponds to:
            xo = kmxcc - km400x * sizex / 2  'mapPictureform.mapPicture.Width / 2
            yo = kmycc + km400y * sizey / 2 'mapPictureform.mapPicture.Height / 2
            x = (kmxDrag0 - xo) / km400x 'mapdif accounts for size of frame around picture
            y = (yo - kmyDrag0) / km400y
            End If
       ElseIf map50 = True Then
         If mag > 1 Then
            kmxcc = kmxc
            kmycc = kmyc
            xo = kmxcc - (km50x / mag) * (mapwi2 - mapxdif) * 0.5
            yo = kmycc + (km50y / mag) * (maphi2 - mapydif) * 0.5
            x = mag * (kmxDrag0 - xo) / km50x 'mapdif accounts for size of frame around picture
            y = mag * (yo - kmyDrag0) / km50y
         Else
            kmxcc = kmxc + (km50x) * (mapwi - mapwi2 + mapxdif) / 2
            kmycc = kmyc - (km50y) * (maphi - maphi2 + mapydif) / 2
            xo = kmxcc - km50x * sizex / 2  'mapPictureform.mapPicture.Width / 2
            yo = kmycc + km50y * sizey / 2 'mapPictureform.mapPicture.Height / 2
            x = (kmxDrag0 - xo) / km50x 'mapdif accounts for size of frame around picture
            y = (yo - kmyDrag0) / km50y
            End If
         End If
m100:     Select Case coordmode%
           Case 1 'ITM
              dragCoordX = x
              dragCoordY = y
           Case 2 'GEO
              If world = True Then
                 If mag > 1 Then
                   lonc = lon '+ fudx / mag
                   latc = lat '+ fudy / mag
                   xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                   yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                   x = (kmxDrag - xo) * (sizewx * mag / deglog)
                   y = (yo - kmyDrag) * (sizewy * mag / deglat)
                 Else
                   lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                   latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                   
                   xo = lonc - deglog / 2
                   yo = latc + deglat / 2
                   x = (kmxDrag - xo) * (sizewx / deglog)
                   y = (yo - kmyDrag) * (sizewy / deglat)
                   End If
                dragCoordX = x
                dragCoordY = y
                End If
             Case Else 'option not supported yet
                dragCoordX = x
                dragCoordY = y
          End Select
  
  End If
  Exit Sub
  
errhand:
   ier% = -1
  
End Sub



Sub FindSearchResult(x As Single, y As Single)
   'finds nearest search result to right clicked point
   'and moves the position of the Search Result DataGrid to that point
   
   Screen.MousePointer = vbHourglass
   
   On Error GoTo errhand
   
   'convert screen coordinates to kmx,kmy, or to lon,lat
   Call ScreenToGeo(x, y, GeoX0, GeoY0, 1, ier%)
   If ier% < 0 Then Exit Sub
   
   'now search through the DataGrid for the nearest point
   For i& = 1 To mapsearchfm.sky2.Rows - 1
        GeoX = mapsearchfm.sky2.TextArray(mapsearchfm.skyp2(i&, 1))
        GeoY = mapsearchfm.sky2.TextArray(mapsearchfm.skyp2(i&, 2))
        If world = True Then
           tmpGeoX = GeoX
           tmpGeoY = GeoY
           GeoX = tmpGeoY
           GeoY = tmpGeoX
           End If
        If i& = 1 Then
           Dist = Sqr((GeoX0 - GeoX) ^ 2 + (GeoY0 - GeoY) ^ 2)
           GeoXmin = GeoX
           GeoYmin = GeoY
           geoi& = i&
        Else
           distNew = Sqr((GeoX0 - GeoX) ^ 2 + (GeoY0 - GeoY) ^ 2)
           If distNew < Dist Then
              Dist = distNew
              GeoXmin = GeoX
              GeoYmin = GeoY
              geoi& = i&
              End If
           End If
   Next i&
   
   'now move that row to the first position
   mapsearchfm.sky2.RowPosition(geoi&) = 1
   mapsearchfm.sky2.row = 1 'highlight this row
   BringWindowToTop (mapsearchfm.hwnd)
   
   Screen.MousePointer = vbDefault
   Exit Sub
   
errhand:
   Screen.MousePointer = vbDefault
   MsgBox "Encountered error number: " & Str$(Err.Number) & vbLf & Err.Description, vbCritical + vbOKOnly, "Maps&More"
   

End Sub
'---------------------------------------------------------------------------------------
' Procedure : SysVersions
' DateTime  : 2/25/2003 09:15
' Author    : chaim keller
' Purpose   : determine windows version
'---------------------------------------------------------------------------------------
'
   Function SysVersions()
      Dim ver As Long
      'Dim DosVer As Long, WindowsVersion As String, DosVersion As String

      ver = GetVersion()

      WinVer = ver And &HFFF&
'     WindowsVersion = Format((WinVer Mod 256) + _
'        ((WinVer \ 256) / 100), "Fixed")

'      DosVer = ver \ &H10000
'      DosVersion = Format((DosVer \ 256) + _
'         ((DosVer Mod 256) / 100), "Fixed")
'
'      MsgBox "Windows Version: " & WindowsVersion _
'         & Chr(13) & "DOS Version: " & DosVersion

   End Function
Public Sub sCenterForm(tmpF As Form)
'centers a form in the middle of the program's main form

Dim x As Integer, y As Integer

On Error GoTo sCenterForm_Error

    x = Maps.Left + 0.5 * Maps.Width - 0.5 * tmpF.Width
    y = Maps.Top + 0.5 * Maps.Height - 0.5 * tmpF.Height
    
    tmpF.Move x, y
    
    On Error GoTo 0
    Exit Sub
    
sCenterForm_Error:
    
End Sub
Public Sub Temperatures(lat As Double, lon As Double, MinTemp() As Integer, AvgTemp() As Integer, MaxTemp() As Integer, ier As Integer)

'extract the WorldClim averaged minimum and average temperature for months 1-12 for this lat,lon
'constants of the bil files
Dim NROWS As Long
NROWS = 21600 'number of rows of the bil files
Dim NCOLS As Long
NCOLS = 43200 'number of columns of the bil files
Dim XDIM As Double
XDIM = 8.33333333333333E-03 'longitude steps of bil files in degrees
Dim YDIM As Double
YDIM = 8.33333333333333E-03 'latitude steps of bil files in degrees
Dim NODATA As Long
NODATA = -9999 'no temp data flag of bil files
Dim ULXMAP As Double
ULXMAP = -179.995833333333 'top left corner longitude of bil files
Dim ULYMAP As Double
ULYMAP = 89.9958333333333 'top left corner latitude of bil files

Dim FilePathBil As String
Dim FileNameBil As String
Dim DirPath$

Dim tncols As Long, IKMY&, IKMX&, numrec&, IO%, Tempmode%

FilePathBil = App.Path & "\WorldClim_bil"
If Dir(FilePathBil, vbDirectory) <> sEmpty Then
    DirPath$ = FilePathBil
    FilePathBil = DirPath$
Else
    'first try default
    FilePathBil = Mid$(USADir$, 1, 1) & "c:\devstudio\vb" & "\WorldClim_bil"
    If Dir(FilePathBil, vbDirectory) <> sEmpty Then
        DirPath$ = FilePathBil
        FilePathBil = DirPath$
    Else
        Call MsgBox("Can't find the bil directory at the following location:" _
                    & vbCrLf & FilePathBil _
                    & vbCrLf & vbCrLf & "Please select the correct direcotry location." _
                    , vbExclamation, "Missing bil file directory")
        DirPath$ = BrowseForFolder(Drukfrm.hwnd, "Choose Directory")
        If Dir(DirPath$, vbDirectory) <> "" Then
           FilePathBil = DirPath$
        Else
           ier = -1
           Exit Sub
           End If
        End If
    End If
'first extract minimum temperatures

 Tempmode% = 0
T50:
 If Tempmode% = 0 Then 'minimum temperatures to be used for sunrise calculations
    FilePathBil = DirPath$ & "\min_"
 ElseIf Tempmode% = 1 Then 'average temperatures to be used for sunset calculations
    FilePathBil = DirPath$ & "\avg_"
 ElseIf Tempmode% = 2 Then 'average temperatures to be used for sunset calculations
    FilePathBil = DirPath$ & "\max_"
    End If
    
 For i = 1 To 12
        
    FileNameBil = FilePathBil

    Select Case i
       Case 1
          FileNameBil = FileNameBil & "Jan"
       Case 2
          FileNameBil = FileNameBil & "Feb"
       Case 3
          FileNameBil = FileNameBil & "Mar"
       Case 4
          FileNameBil = FileNameBil & "Apr"
       Case 5
          FileNameBil = FileNameBil & "May"
       Case 6
          FileNameBil = FileNameBil & "Jun"
       Case 7
          FileNameBil = FileNameBil & "Jul"
       Case 8
          FileNameBil = FileNameBil & "Aug"
       Case 9
          FileNameBil = FileNameBil & "Sep"
       Case 10
          FileNameBil = FileNameBil & "Oct"
       Case 11
          FileNameBil = FileNameBil & "Nov"
       Case 12
          FileNameBil = FileNameBil & "Dec"
    End Select
    FileNameBil = FileNameBil + ".bil"
    
    If Dir(FileNameBil) <> sEmpty Then
       FileIn% = FreeFile
       Open FileNameBil For Binary As #FileIn%
   
        y = lat
        x = lon
        
        IKMY& = CLng((ULYMAP - y) / YDIM) + 1
        IKMX& = CLng((x - ULXMAP) / XDIM) + 1
        tncols = NCOLS
        numrec& = (IKMY& - 1) * tncols + IKMX&
        Get #FileIn%, (numrec& - 1) * 2 + 1, IO%
        If IO% = NODATA Then IO% = 0#
        If Tempmode% = 0 Then
            MinTemp(i) = IO%
        ElseIf Tempmode% = 1 Then
            AvgTemp(i) = IO%
        ElseIf Tempmode% = 2 Then
            MaxTemp(i) = IO%
            End If
            
        Close #FileIn%
    Else
        Call MsgBox("Can't find the bil file at the following location:" _
                & vbCrLf & FileNameBil _
                , vbExclamation, "Missing bil file")
        ier = -2
        Exit Sub
        End If
        
  Next i
  
  'now go back and calculate the AvgTemps
  If Tempmode% = 0 Then
     Tempmode% = 1
     GoTo T50
  ElseIf Tempmode% = 1 Then
     Tempmode% = 2
     GoTo T50
     End If

End Sub

Public Function MaxHalfAzimuthRange(latitude) As Integer
   Dim maxdecl As Double, l1 As Double
   
   'make sure latitude parameter is number
   l1 = CDbl(latitude)
   
   'calculate approximate half azimuth range
    maxdecl = (66.5 - Abs(l1)) + 23.5 'approximate formula for maximum declination as function of latitude (l1)
    'source: chart of declination range vs latitude from http://en.wikipedia.org/wiki/Declination
    MaxHalfAzimuthRange = CInt((Abs(DASIN(Cos(maxdecl * cd))) / cd)) 'approximate formula for half maximum angle range as function of declination
    'source: http://en.wikipedia.org/wiki/Solar_azimuth_angle
    
   'add additional range for case of high terrain that needs added azimuth range for visible calculations
    If Abs(l1) >= 60 And Abs(latitude) < 62 Then
       MaxHalfAzimuthRange = MaxHalfAzimuthRange + 3
    ElseIf Abs(l1) >= 62 And Abs(l1) < 65 Then
       MaxHalfAzimuthRange = MaxHalfAzimuthRange + 10
    ElseIf Abs(l1) >= 65 Then
       MaxHalfAzimuthRange = MaxHalfAzimuthRange + 20
       End If

End Function
