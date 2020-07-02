Attribute VB_Name = "modGDModule"
'*************************************************************
'Date: Iyar 5764, (May 2004) Jerusalem, Eretz Yisroel
'
'This is the Module containing global subroutines used by the
'MapDigitizer Program.  The routines in this program were written using
'Microsoft Visual Basic Version 5.0.  However the Active X
'components, as well as the DAO library uses VB 6.0 components.
'These components were obtained free from Microsoft from two
'sources: (1) By downloading and installing their DirectX 8 SDK
'for VB, (2) By using Microsoft's freeware ActiveX upgrade program.
'
'This software should be compatible with later versions of Visual Basic
'until the API's start using 64 bit integers.  At that time any Long
'Integer variables will probably become 64 bit integers.  It is
'impossible to anticipate any other changes at this time.
'
'This software was tested under Windows 9X, 2000, and XP.
'I expect it to work under Windows XP+ and beyond.
'
'The main purpose of this program is to provide searches over the
'two paleontologic databases at the GSI: (1) The active database, and
'(2) the inactive database produced from the scanned forms of the old records
'
'*****************Windows API functions, subroutines and constants*********
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Public Const SW_MAXIMIZE = 3 'defined already in other module
'Public Const SW_MINIMIZE = 6 'defined already in other module
Public Const SW_NORMAL = 1
'Public Const SW_SHOWMAXIMIZED = 3 'defined already in other module
'Public Const SW_SHOWMINIMIZED = 2 'defined already in other module
'Public Const SW_SHOWNORMAL = 1 'defined already in other module
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOSIZE = &H1
Public Const sEmpty = ""
Public Const MAX_PATH = 260
'Public Const NUM_FOSSIL_TYPES = 8
Public Const Installation_Type = 1 '0 for without GTCO digitizer interface
                                   '1 for with GTCO digitizer interface

'*********************program constants and parameters*************
Public Const cd = 1.74532927777778E-02 'conv deg to rad
Public Const PI As Double = 3.14159265358979
Public Const Rearthkm As Double = 6378.137 'WGS84 equitorial radium in kilometers
Public Const Rearth As Double = 6371315#

'*****************default user name used for activating msascess option****
Public Const ADMIN_USERNAME As String = "GSI_ADMIN"

'*********map parameters***************************
Public pixwi, pixhi, twipsx As Long, twipsy As Long, picnam$
Public pixwi0, pixhi0, x10, x20, y10, y20, picnam0$, picTopo$
Public GeoMap As Boolean, TopoMap As Boolean, direction As Integer
Public shiftmag As Boolean, lblX As String, LblY As String
Public Geo As Boolean, GeoX As Single, GeoY As Single, GeoHgt As Single
Public GpsCorrection As Boolean, GeoGoto As Boolean, ShowContGeo As Boolean
Public GeoDecDeg As Boolean, se&, ce&, ShowDetails As Boolean
Public CenterBlinkState As Boolean, GeoMapMode%
Public PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
Public MeLeft As Long, MeTop As Long, MeWidth As Long, MeHeight As Long
Public MaxColorHeight As Single, MinColorHeight As Single
Public HeightPrecision As Integer
Public xLL As Double, yLL As Double, nRowLL As Long, nColLL As Long
Public XStepLL As Double, YStepLL As Double, zminLL As Double, zmaxLL As Double
Public AngLL As Double, blank_LL As Double

'**********Digitizer flags and parameters**************** '<<<<<<<<<<<<digi changes
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const SRCCOPY_digi = &HCC0020

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Type POINTGEO
   XGeo As Double
   YGeo As Double
End Type

Public Type PointsAPI
   x As Long
   Y As Long
   Z As Single
End Type

Public Type pointRS
   xScreen As Long
   yScreen As Long
   XGeo As Double
   YGeo As Double
End Type

Public Type ColorEnum
   x As Single 'x pixel
   Y As Single 'y pixel
   RedColor As Byte 'RGB values 0-255
   GreenColor As Byte
   BlueColor As Byte
End Type

Public DigitizeOn As Boolean '<<<<<<<<<<<<digi changes
Public DigitizeMagvis As Boolean
Public DigitizeMagInit As Boolean
Public DigitizePoint As Boolean
Public DigitizeBlankPoint As Boolean
Public DigitizeLine As Boolean
Public DigitizeBeginLine As Boolean
Public DigitizeEndLine As Boolean
Public DigitizeContour As Boolean
Public DigitizePadVis As Boolean
Public DigitizeContinueContour As Boolean
Public DigiContourStart As Boolean
Public DigiLogFileOpened As Boolean
Public Digilogfilnum%
Public DigiRS As Boolean
Public GDRSfrmVis As Boolean
Public DigitizerEraser As Boolean
Public DigitizerSweep As Boolean
Public Digitizing As Boolean
Public DigiEraseBrushSize As Integer
Public MinDigiEraserBrushSize As Integer
Public DigitizeExtendGrid As Boolean
Public DigiExtendFirstPoint As Boolean
Public DigitizeHardy As Boolean
Public DigitizeDeleteContour As Boolean
Public HardyCoordinateOutput As Boolean 'set to true to output coordinates instead of pixels for plotting
Public DigiGDIfailed As Boolean
Public DigiPicFileOpened As Boolean
Public Picfilnum%
Public DigiZoomed As Boolean
Public ULGeoX As Double, ULGeoY As Double, LRGeoX As Double, LRGeoY As Double
Public ULPixX As Double, ULPixY As Double, LRPixX As Double, LRPixY As Double, LRGridX As Double, LRGridY As Double, ULGridX As Double, ULGridY As Double
Public XStepITM As Double, YStepITM As Double, XStepDTM As Double, YStepDTM As Double, HalfAzi As Double, StepAzi As Double, Apprn As Double
Public RotatedGrid As Boolean, numXYZpoints As Long
Public PixToCoordX As Double, PixToCoordY As Double
Public MapParms() As String
Public DigiMagnify As Integer
Public DigiReDrawContours As Boolean
Public DigiTableWorksOpen As Boolean
Public DigiRightButtonIndex As Integer
Public DigiBackground As Long
Public DigiEntered As Boolean
Public numDistContour As Integer, numDistLines As Integer, numSensitivity As Integer, numContours As Integer ' arcdir, mxddir
Public ChainCodeMethod As Integer
Public PointCenterClick As Integer
Public DigiEditPoints As Boolean
Public InitDigiGraph As Boolean
Public filnumImage%
Public ImagePointFile As Boolean
Public XpixLast As Long, YpixLast As Long, HighLightColor As Long
Public DigiEditMode As Integer
Public DigiSearchRegion As Integer
Public HeightSearch As Boolean
Public GenerateContours As Boolean
Public DTMcreating As Boolean
Public BasisDTMheights As Boolean
Public basedtm%
Public DigiConvertToMeters As Double
Public MapUnits As Double
Public Belgier_Smoothing As Boolean
Public InvElev As Double
Public HorizMode%
Public CoordListVis As Boolean
Public MarkerColor As Long
Public CoordListZoom As Boolean

'----------------GPS global constants------------------------
Public Const MAX_PORT = 15 ' maximum number of com ports to search
Dim Max_Com_Port As Integer 'largest available com port number
Public Const MAX_GPS_WAIT = 18 'maximum number of timer intervals to wait for GPS satellite signal
Public Const MAX_WAIT_NUMBER = 300 'maximum number of GPS timer intervels to test for zero velocity
Public Const MAX_REPEAT_DISTANCE_TEST = 5 'maximum number of cycles to test spurious coordinates
Public Const MAX_NO_SIGNAL = 300 'maximum number of cycles to allow before showing "no signal" message box
Public GPSconnected As Boolean
Public GPScommunication_Established As Boolean
Public ComPort% 'com port that gps is connected to
Public GPS_timer_trials As Integer
Public GPS_signal_lost As Boolean
Public Const GPS_Dist_Resolution = 5#  'km
Public GPSConnectString As String
Public GPSConnectString0 As String
Public GPSenabled As Boolean 'flag to determine if gps connection was ever established
Public GPSSetupVis As Boolean
'Public distkmTraveld As Double 'elapsed distance log
Public g_lat0 As Double, g_lon0 As Double, g_Vcruise0 As Double 'elapsed coordinates log
Public numTestGPSpnts% 'gps test points
Public GPSNow
Public GPS_no_message As Boolean
Public GPS_off As Boolean
Public GPSstarted As Boolean
Public GPSended As Boolean
Public GPS_test_loaded As Boolean
Public LastGPSDistance As Double
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
Public GPS_test_distance_number As Long
Public GPS_no_signal_number As Long
Public GPSlog_filename$
'Public GPSfillog%
Public GPSlog As Boolean
Public GPSOKButton As Boolean
Public GPSDescentVel() As Double
Public NumWayPointsPast As Long
Public ProlificGPS As Boolean
Public DeviceType_Init As Boolean
Public Control_Num As Integer

Public g_ier As Integer

Public g_PointImage() As Byte

Public WinVer As Integer

'Public DigiPage As Integer 'used for diagnostics (especially touch screen blit problem)

'positions of gddigitizerfrm buttons and elevation text box

Public Type PicZoom
   left As Long
   top As Long
   Zoom As Single
   LastZoom As Single
End Type

Public DigiZoom As PicZoom

Public Type ContourPoint
    X1 As Single 'x pixel of contour line's start point
    Y1 As Single 'y pixel of contour line's start point
    X2 As Single 'x pixel of contour line's end point
    Y2 As Single 'y pixel of contour line's end point
    color As Long 'color of contour line
End Type

Public ContourPoints() As ContourPoint
Public numContourPoints As Long

Public TraceColor As Long
Public TracingColor As couleur

Public PlotInfo() '7, numfiles%)

Public ContourHeight As Single

Public oGestionImageSrc As New CGestionImage

Public digi_last As PointsAPI
Public digi_begin As PointsAPI

Public digiextendgrid_last As POINTAPI
Public digiextendgrid_begin As POINTAPI

'Public dhwnd_digi As Long, dhdc_digi As Long
Public x_digi As Integer, y_digi As Integer
Public w_digi As Integer, h_digi As Integer
Public sw_digi As Integer, sh_digi As Integer
Public nearmouse_digi As POINTAPI
Public new_digi As PointsAPI
Public newblit As Boolean

Public StepInX As Boolean
Public StepInY As Boolean
Public DigiRSStepType As Integer
Public XGridSteps As Double
Public YGridSteps As Double

Public Wtotal As Long, Htotal As Long
Public w1 As Long, w2 As Long
Public h1 As Long, h2 As Long
Public w1new As Long, w2new As Long
Public h1new As Long, h2new As Long
Public oldFileBytes As Long

'digitizing buffers
Public DigiPoints() As PointsAPI, DigiLines() As PointsAPI
Public DigiContours() As PointsAPI
'Public DigiContourColors() As ColorEnum
Public DigiErasePoints() As POINTAPI
Public DigiExtendPoints() As POINTAPI
Public DigiHardyPoints() As PointsAPI
Public numDigiPoints As Long, numDigiLines As Long, numDigiContours As Long
Public numDigiErase As Long, numDigiExtendPoints As Long, numDigiHardyPoints As Long

'rubber sheeting flags and buffers
Public Next_Point As POINTAPI
Public RS() As pointRS
Public RSfilnum%
Public RSopenedfile As Boolean
Public numRS As Long
Public DigiRubberSheeting As Boolean
Public RSMethod1 As Boolean
Public RSMethod2 As Boolean
Public RSMethod0 As Boolean
Public RSMethodBoth As Boolean
Public NX_CALDAT As Long
Public NY_CALDAT As Long
Public SX_CALDAT() As Double, SY_CALDAT() As Double, GX_CALDAT() As Double, GY_CALDAT() As Double
Public SX_CALDAT_2() As Double, SY_CALDAT_2() As Double, GX_CALDAT_2() As Double, GY_CALDAT_2() As Double
Public SX_ROT() As Double, SY_ROT() As Double

'**********DTM variables**************
Public CHMAP(14, 26) As String * 2, filnumg%
Public CHMNE As String * 2, CHMNEO As String * 2, SF As String * 2

'*********directory information and flags************
Public dirNewDTM As String
Public dbdir2 As String
Public NEDdir As String
Public dtmdir As String
Public ASTERdir As String
Public DTMtype As Integer
Public JKHDTM As Boolean
Public ASTERbilOpen As Boolean
Public ASTERNorth As Integer
Public ASTEREast As Integer
Public ASTERfilename As String
Public ASTERNrows%
Public ASTERNcols%
Public ASTERxdim As Double
Public ASTERydim As Double
Public ASTERfilnum%
Public SRTMfileOpen As Boolean
Public SRTMNorth As Integer
Public SRTMEast As Integer
Public SRTMfil$
Public SRTMfilnum%
Public topodir As String
Public arcdir As String
Public accdir As String
Public googledir As String
Public kmldir As String
Public URL_OutCrop As String
Public URL_Well As String
Public UseNewDTM%, UsingNewDTM As Boolean
'Public linked As Boolean
'Public linkedOld As Boolean
'Public linkedpiv As Boolean
Public heights As Boolean
'Public topos As Boolean
'Public arcs As Boolean
'Public acc As Boolean
Public google As Boolean
'Public SearchDBs%
Public numMaxHighlight&, SaveClose%, Save_xyz%
Public ReportCoord(1) As POINTAPI


'********global scanned (old) database parameters**********
'Public arrN03(3) As String 'core/cutting D_KEY=2
'Public arrN05(3) As String 'well/outcropping D_KEY=3
'Public arrN11() As String 'names D_KEY=5
'Public arrN11_piv() As String 'names D_KEY=5, used for merging databases
'Public numN11 As Long 'number of elements of arrN11
'Public numN11_piv As Long 'number of elements of arrN11_piv
'Public arrN06(4) As String 'prefix ages D_KEY=6
'Public arrN07() As String 'ages D_KEY=7
'Public numN07 As Long 'number of elements in arrN07
'Public arrN10() As String 'formations D_KEY=4
'Public numN10 As Long 'number of elements in arrN10
'Public arrOldDates() As Long 'conversion from Old database age numbers to Active database age numbers
'Public numOldDates As Long 'number of elements in arrOldDates
'Public arrOldFormation() As Long 'conversion from Old database formation numbers to Active database formation numbers
'Public strSQLOld As String, modeEdit%, GotPassword As Boolean
'Public MaxAgeNum&, OKeyClick As Boolean, ONameClick As Boolean
'Public LoadingEditForm As Boolean, StepDocNo As Boolean, PwdCancel As Boolean
Public PwdCancel As Boolean

'variables used for undoing editing
'Public txtITMx00$, txtITMy00$
'Public txtPreE00$, txtEarlyAge00$
'Public txtPreL00$, txtLaterAge00$
'Public txtFormation00$, txtNames00$
'Public foscat00%, foscat%, oldSource00%, oldSource%
'Public txtGL00$, txtDepth00$, LastOKey&
'Public minDigits As String, maxDigits As String

'*************open directory structure to pick directory**************
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2

Public Declare Function SHBrowseForFolder Lib "Shell32" _
                                  (lpBI As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib "Shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Public Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'---------------librarires used for checking if running with administrator privilege----------------
' dwPlatformId values
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

' These functions are for getting the process token information, which IsUserAnAdministrator uses to
' handle detecting an administrator that’s running in a non-elevated process under UAC.

Private Const TOKEN_READ As Long = &H20008
Private Const TOKEN_ELEVATION_TYPE As Long = 18

Public Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer
Public Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

'---------API routine for detecting if the platform is vista or higher-------------

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

'***********Fancy progress bar global variables and API
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As _
 Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As _
 Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
 ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As _
 Long) As Long
Public pbScaleWidth As Long

'***********Quick Options Button**********************
'Public bNames As Boolean, bGroundLevels As Boolean, bDepths As Boolean
'Public bFossilTypes As Boolean, bCoordinates As Boolean, bFormations As Boolean
'Public bGeologicAges As Boolean, bSampleSources As Boolean
'Public QuickVis As Boolean

'*********general global variables and arrays************
Public SizeX, SizeY, direct$, direct2$, buttonstate&(70), numArc&
Public dragbegin As Boolean, dragbox As Boolean, drawbox As Boolean
Public drag1x, drag1y, drag2x, drag2y, magclose As Boolean, W, h
Public formwidth, formheight, mag, maginit As Boolean, magvis As Boolean
Public datevis As Boolean, SearchDigi As Boolean, PicSum As Boolean
Public blink_mark As POINTAPI
Public ReportPnts(), NumReportPnts&, ArcDump As Boolean, stepsearch&
Public numReport&, NearestPnt&, Searching As Boolean, sFormSearchOld As String
Public XResol As Integer, PreviewOrderNum&, OptionsVis As Boolean
Public GDform1Height As Long, MapGraphVis As Boolean, plotfile$
Public MagWidth As Long
'Public RangeOfDates As Boolean, sDateRange As String, XResol As Integer
'Public PreEarlyDate&, PreLateDate&, sFormSearch As String, OptionsVis As Boolean
'Public sAnalystSearch As String, StopSearch As Boolean, EditDBVis As Boolean
'Public EditScannedDBVis As Boolean
Public g_nrows As Long, g_ncols 'records number of rows and cols in topo_coord export file
Public numfiles%, Files() As String
Public SavedNewScannedFile As Boolean, wizardCanceled As Boolean
Public bPaleoZone As Boolean, CheckDuplicatePoints As Boolean, Old_OKEY&
Public ScreenDump As Boolean, MinimizeReport As Boolean, OrderNum&
Public EraseMaps As Boolean, StopPlotting As Boolean, PrintMag As Boolean
Public NewHighlighted&, Highlighted() As Long, numHighlighted&, iFlex As Integer
'Public SearchAll As Boolean, ReportPaths&, bFossilNames As Boolean
'Public PreviewOrderNum&, CloseSearchWizard As Boolean, Closing As Boolean
'Public Previewing As Boolean, MaxOrderNum&(2), SearchVis As Boolean
Public ReportPaths&, infonum&
Public MinOrder&, MaxOrder&, DetailRecordNum&, SplashVis As Boolean
Public IgnoreAutoRedrawError%, PreviousInstance As Boolean, Direc%
'Public tifDir$, tifViewerDir$, tifCommandLine$, nCombine&, StoreWidthAdd%, StoreHeightAdd%
'Public ReplaceWellZ As Boolean, ReplaceOtherZ As Boolean, optEDS%
Public ActivatedVersion As Boolean, GoogleDump As Boolean
'Public AveEastSearch, AveNorthSearch

'*******strings that keep track of Boolean expressions for Fossil types
'Public strSql1 As String, strSqlCategory As String

''***********string arrays that keep track of selected paleo zones w.r.t id's
'Public sArrConodonta() As String
'Public sArrDiatom() As String
'Public sArrForaminifera() As String
'Public sArrMegafauna() As String
'Public sArrNannoplankton() As String
'Public sArrPalynology() As String
'Public sArrOstracoda() As String
''***********long int arrays that keep track of selected fossil species
'Public lArrCono() As Long, numFosCono As Long
'Public lArrDiatom() As Long, numFosDiatom As Long
'Public lArrForam() As Long, numFosForam As Long
'Public lArrMega() As Long, numFosMega As Long
'Public lArrNano() As Long, numFosNano As Long
'Public lArrOstra() As Long, numFosOstra As Long
'Public lArrPaly() As Long, numFosPaly As Long
''***********string array that hold Fossil Names w.r.t IDs************
'Public sArrConoNames() As String
'Public sArrDiatomNames() As String
'Public sArrForamNames() As String
'Public sArrMegaNames() As String
'Public sArrNanoNames() As String
'Public sArrOstraNames() As String
'Public sArrPalyNames() As String

'***************contour detection***********************
'Type couleur permet de stocker les valeurs
'R g et B d'un pixel plus elegement
Public Type couleur
    R As Long
    V As Long
    b As Long
End Type

Public Type pixel
    x As Long
    Y As Long
    couleur As couleur
End Type

Public Start_Point As POINTAPI '<<<<<<<<<<<<<changes

'Ensemble des pixels de contour
Public valpix As Long

Public PointStart As Boolean
Public Point_Color As couleur 'keep track of color under mouse
Public Start_Color As couleur 'keep track of color at click point
Public RedColor As couleur
Public digi_LoweredSensitivity As Boolean

Public Const INIT_VALUE = 9999999

Public color_init As couleur

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'-------------------API for mouse control using GTCO digitizer--------------------------
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Public Const MOUSEEVENTF_LEFTDOWN = &H2      ' left button down
Public Const MOUSEEVENTF_LEFTUP = &H4        ' left button up
Public Const MOUSEEVENTF_ABSOLUTE = &H8000   ' absolute move
Public Const MOUSEEVENTF_MOVE = &H1          ' move

Public TabletControlVis As Boolean
Public TabletControlOn As Boolean

''''''''''''''''''''conrec variables''''''''''''''''''''''''
Public xmin As Double, xmax As Double
Public ymin As Double, ymax As Double
Public zmin As Double, zmax As Double
Public cpt() As Integer 'gmt rainbow palette
Public numcpt As Integer

'-----------------Mouse Wheel events------------------------------------------------------------
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

'Constants
Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Public Const blank_value As Double = 1.70141E+38 'value of "blank" elevation expected by Global Mapper when reading surfer v7 binary grd files

'KeyDown events
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

'--------------------------API for super duper Shell function-------------------------------
' Constants - ShellExecute.nShowCmd
Public Enum ShowStates
  SW_HIDE = 0            ' Hides the window and activates another window.
  SW_MAXIMIZE = 3        ' Maximizes the specified window.
  SW_MINIMIZE = 6        ' Minimizes the specified window and activates the next top-level window in the z-order.
  SW_RESTORE = 9         ' Activates and displays the window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when restoring a minimized window.
  SW_SHOW = 5            ' Activates the window and displays it in its current size and position.
  SW_SHOWDEFAULT = 10    ' Sets the show state based on the SW_ flag specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application. An application should call ShowWindow with this flag to set the initial show state of its main window.
  SW_SHOWMAXIMIZED = 3   ' Activates the window and displays it as a maximized window.
  SW_SHOWMINIMIZED = 2   ' Activates the window and displays it as a minimized window.
  SW_SHOWMINNOACTIVE = 7 ' Displays the window as a minimized window. The active window remains active.
  SW_SHOWNA = 8          ' Displays the window in its current state. The active window remains active.
  SW_SHOWNOACTIVATE = 4  ' Displays a window in its most recent size and position. The active window remains active.
  SW_SHOWNORMAL = 1      ' Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
End Enum

' Constants = ShellExecute.strAction
Public Enum ShowActions
  SA_Open = 0       ' Opens the file specified by the strWhatToShell parameter. The file can be an executable file, a document file, or a folder.
  SA_Explore = 1    ' Explores the folder specified by strWhatToShell.
  SA_Edit = 2       ' Launches an editor and opens the document for editing. If strWhatToShell is not a document file, the function will fail.
  SA_Print = 3      ' Prints the document file specified by strWhatToShell. If strWhatToShell is not a document file, the function will fail.
  SA_Properties = 4 ' Displays the file or folder's properties.
  SW_Find = 5       ' Initiates a search starting from the specified directory.
  SW_Play = 6       ' Plays audio files such as WAV, MID, and RMI files.
End Enum

' Constants - ShellExecute Return Codes (Errors)
Public Const ERROR_OOM = 0               ' The operating system is out of memory or resources.
Public Const ERROR_FILE_NOT_FOUND = 2    ' The specified file was not found.
Public Const ERROR_PATH_NOT_FOUND = 3    ' The specified path was not found.
Public Const ERROR_BAD_FORMAT = 11       ' The .exe file is invalid (non-Win32® .exe or error in .exe image).
Public Const SE_ERR_ACCESSDENIED = 5     ' The operating system denied access to the specified file.
Public Const SE_ERR_ASSOCINCOMPLETE = 27 ' The file name association is incomplete or invalid.
Public Const SE_ERR_DDEBUSY = 30         ' The DDE transaction could not be completed because other DDE transactions were being processed.
Public Const SE_ERR_DDEFAIL = 29         ' The DDE transaction failed.
Public Const SE_ERR_DDETIMEOUT = 28      ' The DDE transaction could not be completed because the request timed out.
Public Const SE_ERR_DLLNOTFOUND = 32     ' The specified dynamic-link library was not found.
Public Const SE_ERR_FNF = 2              ' The specified file was not found.
Public Const SE_ERR_NOASSOC = 31         ' There is no application associated with the given file name extension. This error will also be returned if you attempt to print a file that is not printable.
Public Const SE_ERR_OOM = 8              ' There was not enough memory to complete the operation.
Public Const SE_ERR_PNF = 3              ' The specified path was not found.
Public Const SE_ERR_SHARE = 26           ' A sharing violation occurred.

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal strAction As String, ByVal strWhatToShell As String, ByVal strParameters As String, ByVal strDefaultDir As String, ByVal nShowCmd As Long) As Long


Public Function SuperShell(ByVal File_Email_URL As String, _
                           Optional ByVal Parameters As String = vbNullString, _
                           Optional ByVal ShowAction As ShowActions = SA_Open, _
                           Optional ByVal ShowSate As ShowStates = SW_SHOWNORMAL, _
                           Optional ByVal DefaultDirectory As String = vbNullString, _
                           Optional ByVal WindowHandle As Long = -1, _
                           Optional ByVal ShowErrorMsg As Boolean = True, _
                           Optional ByRef Return_ErrNum As Long, _
                           Optional ByRef Return_ErrSrc As String, _
                           Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnValue As Long
  Dim strVerb     As String
  Dim strErrMsg   As String
  
  Return_ErrNum = 0
  Return_ErrSrc = sEmpty
  Return_ErrDesc = sEmpty
  
  ' Make sure the parameters passed are valid
  File_Email_URL = Trim(File_Email_URL)
  If File_Email_URL = sEmpty Then Err.Raise -1, "SuperShell()", "No file, Email, or URL specified to open"
  If WindowHandle = -1 Then WindowHandle = App.hInstance
  If right(File_Email_URL, 1) <> Chr(0) Then File_Email_URL = File_Email_URL & Chr(0)
  If right(DefaultDirectory, 1) <> Chr(0) And DefaultDirectory <> vbNullString Then DefaultDirectory = DefaultDirectory & Chr(0)
  If right(Parameters, 1) <> Chr(0) And Parameters <> vbNullString Then Parameters = Parameters & Chr(0)
  
  ' Get the verb that will be used to specify the action to take
  Select Case ShowAction
    Case SA_Open:       strVerb = "open"
    Case SA_Explore:    strVerb = "explore"
    Case SA_Edit:       strVerb = "edit"
    Case SA_Print:      strVerb = "print"
    Case SA_Properties: strVerb = "properties"
    Case SW_Find:       strVerb = "find"
    Case SW_Play:       strVerb = "play"
    Case Else:          strVerb = "open"
  End Select
  
  ' Start the file, document, URL, etc.
  ReturnValue = ShellExecute(WindowHandle, strVerb, File_Email_URL, Parameters, DefaultDirectory, ShowSate)
  
  ' Check if there was an error starting the program
  If ReturnValue <= 32 Then
    Select Case ReturnValue
      Case ERROR_OOM:              strErrMsg = "The operating system is out of memory or resources."
      Case ERROR_FILE_NOT_FOUND:   strErrMsg = "The specified file was not found."
      Case ERROR_PATH_NOT_FOUND:   strErrMsg = "The specified path was not found."
      Case ERROR_BAD_FORMAT:       strErrMsg = "The .exe file is invalid (non-Win32® .exe or error in .exe image)."
      Case SE_ERR_ACCESSDENIED:    strErrMsg = "The operating system denied access to the specified file."
      Case SE_ERR_ASSOCINCOMPLETE: strErrMsg = "The file name association is incomplete or invalid."
      Case SE_ERR_DDEBUSY:         strErrMsg = "The DDE transaction could not be completed because other DDE transactions were being processed."
      Case SE_ERR_DDEFAIL:         strErrMsg = "The DDE transaction failed."
      Case SE_ERR_DDETIMEOUT:      strErrMsg = "The DDE transaction could not be completed because the request timed out."
      Case SE_ERR_DLLNOTFOUND:     strErrMsg = "The specified dynamic-link library was not found."
      Case SE_ERR_FNF:             strErrMsg = "The specified file was not found."
      Case SE_ERR_NOASSOC:         strErrMsg = "There is no application associated with the given file name extension. This error will also be returned if you attempt to print a file that is not printable."
      Case SE_ERR_OOM:             strErrMsg = "There was not enough memory to complete the operation."
      Case SE_ERR_PNF:             strErrMsg = "The specified path was not found."
      Case SE_ERR_SHARE:           strErrMsg = "A sharing violation occurred."
      Case Else:                   strErrMsg = "Unknown Error"
    End Select
    Err.Raise ReturnValue, "ShellExecute()", strErrMsg
  Else
    SuperShell = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = Err.Number
  Return_ErrSrc = Err.Source
  Return_ErrDesc = Err.Description
  Err.Clear
  If ShowErrorMsg = True Then
    MsgBox "The following error occured while trying to open the specified file, folder, URL, Email, or document:" & Chr(13) & Chr(13) & _
           "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
           "Error Source = " & Return_ErrSrc & Chr(13) & _
           "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation + vbSystemModal, "  Error"
  End If
  
End Function


'********************************************************
'---------------------------------------------------------------------------------------
' Procedure : GetSystemPath
' DateTime  : 5/4/2004 10:38
' Author    : Chaim Keller
' Purpose   : Finds path to Windows/system32
'---------------------------------------------------------------------------------------
'
Public Function GetSystemPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetSystemPath = left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetSystemPath = ""
End If
End Function
Public Sub gotocoord()
 'shift this point to middle of screen
 'this will be the case when (X,Y) = (picture1.width/2, picture1.height/2)
 
 On Error GoTo errhand
 
 Dim x As Single, Y As Single
 
 Dim kmx As Long, kmy As Long, ier As Integer
 Dim lt2 As Double, lg2 As Double, hgt2 As Integer
 Dim XGeo As Double, YGeo As Double
 Dim ShiftX As Double, ShiftY As Double
 
 Dim CurrentX As Double, CurrentY As Double
 
 Dim VarD As Double
 Dim BytePosit As Long
 
 Dim Tolerance As Double
 Dim XDif As Double, YDif As Double
 Tolerance = 0.00001 'this is the tolerance in degrees in determining the goto pixel coordinate after iteratrion
 Dim SecondOrderShift As Boolean
 
 SecondOrderShift = False
 
 If Digitizing Then
 
    If LRGeoX = ULGeoX Or ULGeoY = LRGeoY Or Not DigiRubberSheeting Then
       'use rubber sheeting to determine them
       GDMDIform.Toolbar1.Buttons(8).Enabled = False
       Exit Sub
       End If
 
    Dim GeoX As Double, GeoY As Double
    Dim GeoToPixelX As Double, GeoToPixelY As Double
    
    GeoX = val(GDMDIform.Text5.Text)
    GeoY = val(GDMDIform.Text6.Text)
    
    If Not IsNumeric(GeoX) Then
       GDMDIform.Text5.Text = gsEmpty
       Exit Sub
       End If
       
    If Not IsNumeric(GeoY) Then
       GDMDIform.Text6.Text = gsEmpty
       Exit Sub
       End If
    
    
    GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
    GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
    
'    CurrentX = CLng((((GeoX - ULGeoX) * GeoToPixelX) + ULPixX) * DigiZoom.LastZoom)
'    CurrentY = CLng((((ULGeoY - GeoY) * GeoToPixelY) + ULPixY) * DigiZoom.LastZoom)
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
'            Call MsgBox("Inverse coordinate transformation unsuccessful" _
'                        & vbCrLf & "Coordinate grid rotation too large for first approx." _
'                        & vbCrLf & vbCrLf & "(Redo using a less-rotated grid as reference...)" _
'                        , vbInformation, "Goto Error")
'              Screen.MousePointer = vbDefault
'              GDMDIform.picProgBar.Visible = False
'              GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
'              GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
'              Exit Sub
'              End If
        
'        GDMDIform.StatusBar1.Panels(1).Text = ShiftX & ", " & ShiftY
        
'        Call ShiftMap(CSng((CurrentX + ShiftX) * DigiZoom.LastZoom), CSng((CurrentY + ShiftY) * DigiZoom.LastZoom))
        Call ShiftMap(CSng((CurrentX) * DigiZoom.LastZoom), CSng((CurrentY) * DigiZoom.LastZoom))
        
   Else
   
        Call ShiftMap(CSng(CurrentX * DigiZoom.LastZoom), CSng(CurrentY * DigiZoom.LastZoom))
        End If
    
    'determine the height at the new place if possible
    If heights And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
    
      If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
          'convert from ITM to WGS84
          kmx = GDMDIform.Text5
          kmy = GDMDIform.Text6
          
          If kmx < 80000 Or kmx > 260000 Or kmy < 80000 Or kmy > 11350000 Then
             Call MsgBox("Your defined the map coordinate system to be old  ITM." _
                         & vbCrLf & "However, the goto coordinates are not within the" _
                         & vbCrLf & "boundaries of such a coordinate system." _
                         , vbInformation, "ITM coordinate error")
             Exit Sub
             End If
             
          Call ics2wgs84(kmy, kmx, lt2, lg2)
      Else
          lg2 = GDMDIform.Text5
          lt2 = GDMDIform.Text6
          
          If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And _
             (lg2 < 80000 Or lg2 > 260000 Or lt2 < 80000 Or lt2 > 11350000) Then
             Call MsgBox("Your defined the map coordinate system to be old  ITM." _
                         & vbCrLf & "However, the goto coordinates are not within the" _
                         & vbCrLf & "boundaries of such a coordinate system." _
                         , vbInformation, "ITM coordinate error")
             Exit Sub
             End If

          
          If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" And _
             (lg2 < -180 Or lg2 > 180 Or lt2 < -90 Or lt2 > 90) Then
             Call MsgBox("Your defined the map coordinate system to be degrees lon/lat." _
                         & vbCrLf & "However, the goto coordinates are not within the" _
                         & vbCrLf & "boundaries of such a coordinate system." _
                         , vbInformation, "Degrees lon/lat coordinate error")
             Exit Sub
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
       
       GDMDIform.Text7.Text = Format(str$(hgt2), "######0.0#")
       
       End If
    
    Exit Sub
    
    End If
 
 'The ITM coordinates are:
 ITMx = val(GDMDIform.Text5.Text)
 ITMy = val(GDMDIform.Text6.Text)
 
 'Check if the coordinates are within the map range
 If ITMx < x10 Or ITMx > x20 Then
    response = MsgBox( _
        "Can't move to the coordinates!" & vbLf & _
        "The entered coordinates are not within the map's boundaries!", _
        vbExclamation + vbOKOnly, "MapDigitizer")
    Exit Sub
    End If
 If ITMy > y10 Or ITMy < y20 Then
    response = MsgBox( _
        "Can't move to the coordinates!" & vbLf & _
        "The entered coordinates are not within the map's boundaries!", _
        vbExclamation + vbOKOnly, "MapDigitizer")
    Exit Sub
    End If

'    If Not Digitizing And heights = True And lblX = "ITMx" And LblY = "ITMy" Then 'display heights
'       kmx = ITMx
'       kmy = ITMy
'       'Call DTMheight(kmx, kmy, hgt)
'       Dim hgt As Integer
'       Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
'       GDMDIform.Text7 = str(hgt)
'       End If
    If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
    
        If BasisDTMheights And UseNewDTM% Then
           'use background dtm as height reference
           kmx = GDMDIform.Text5
           kmy = GDMDIform.Text6
           
           If kmx >= xLL And kmy >= yLL Then
                BytePosit = 101 + 8 * Nint((kmx - xLL) / XStepLL) + 8 * nColLL * Nint((kmy - yLL) / YStepLL)
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
    
            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                'convert from ITM to WGS84
                kmx = GDMDIform.Text5
                kmy = GDMDIform.Text6
                Call ics2wgs84(kmy, kmx, lt2, lg2)
            Else
                lg2 = GDMDIform.Text5
                lt2 = GDMDIform.Text6
                End If
                
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
       
       GDMDIform.Text3.Text = hgt2
       
       End If

  
    'record these values
    Call UpdatePositionFile(ITMx * DigiZoom.LastZoom, ITMy * DigiZoom.LastZoom, hgt)
    
'    If TopoMap Then 'move topo maps to new position
'       ShowTopoMap (0)
'       GoTo gocd999
'       End If
    'else move large scale geologic map to new position
    
    If Not Digitizing Then
        'Convert (ITMx,ITMy) to pixels
        ITMx0 = ((ITMx - ULGeoX) / (LRGeoX - ULGeoX)) * pixwi
        ITMy0 = ((ULGeoY - ITMy) / (ULGeoY - LRGeoY)) * pixhi
        x = ITMx0 * twipsx
        Y = ITMy0 * twipsy

        'move to the desired place
        Call ShiftMap(x, Y)
        
        'put click coordinates into coordinate boxes
        GDMDIform.Text5 = str(Int(ITMx))
        GDMDIform.Text6 = str(Int(ITMy))
                        
        'Display height at click point
'        If heights = True And lblX = "ITMx" And LblY = "ITMy" Then 'display heights
'           kmx = ITMx
'           kmy = ITMy
'           'Call DTMheight(kmx, kmy, hgt)
'           Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
'           GDMDIform.Text7 = str(hgt)
'           End If
                        
        If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
        
            If BasisDTMheights And UseNewDTM% Then
               'use background dtm as height reference
               kmx = GDMDIform.Text5
               kmy = GDMDIform.Text6
               
               If kmx >= xLL And kmy >= yLL Then
                    BytePosit = 101 + 8 * Nint((kmx - xLL) / XStepLL) + 8 * nColLL * Nint((kmy - yLL) / YStepLL)
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
        
                If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                    'convert from ITM to WGS84
                    kmx = GDMDIform.Text5
                    kmy = GDMDIform.Text6
                    Call ics2wgs84(kmy, kmx, lt2, lg2)
                Else
                    lg2 = GDMDIform.Text5
                    lt2 = GDMDIform.Text6
                    End If
                    
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
           
           GDMDIform.Text3.Text = hgt2
           
           End If

        'Write position file to the hard disk
        'These coordinates will be used as the
        'starting position for the next time the
        'user logs into the program.  It can be also
        'used by the Access database program for
        'automatically inputing coordinates (to reduce
        'human error while inputing coordinates)
        Call UpdatePositionFile(ITMx * DigiZoom.LastZoom, ITMy * DigiZoom.LastZoom, hgt)
        
        End If

'print prompts on statusbar
If SearchDigi Then
  If NumReportPnts& = 0 Then
     GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
  Else
     GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
     End If
Else
     GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
 End If
 
'reenable blinker if haven't done so yet
GDMDIform.CenterPointTimer.Enabled = True

gocd999:
 If EditDBVis Then Ret = BringWindowToTop(GDform1.hwnd)
 If Geo = True Then Ret = BringWindowToTop(GDGeoFrm.hwnd)

Exit Sub

errhand:
   
   Screen.MousePointer = vbDefault
   
   Select Case Err.Number
      Case 52
         'problem with base dtm's file number
         'close it and reopen
         ier = OpenCloseBaseDTM(0)
         Resume
      Case 63
         'bad record number caused by being off the map sheet
         'return the blank height value
         hgt2 = blank_value
         Resume Next
     End Select

   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          "in module: gotocoord", vbCritical + vbOKOnly, "MapDigitizer"

End Sub

Sub DrawPoint(XPnt, YPnt, mode&)
'This routine plots the highlighted points in the search listing
'Mode& = 0 'outcroppings = X
'      = 1 'wells = plus symbol
    
    Screen.MousePointer = vbHourglass
    
   'Convert pixel coordinates to ITM
'    Xtwip = twipsx * pixwi * (XPnt - ULGeoX) / (LRGeoX - ULGeoX)
'    Ytwip = twipsy * pixhi * (YPnt - ULGeoY) / (LRGeoY - ULGeoY)
     
     Xtwip = XPnt
     Ytwip = YPnt

'   record old plot parameters
    'oldfil& = GDform1.Picture2.FillStyle
    oldfilcol& = GDform1.Picture2.FillColor
    olddm& = GDform1.Picture2.DrawMode
    oldfs& = GDform1.Picture2.FillStyle
    'draw new circle at click point
    GDform1.Picture2.DrawMode = 13
    If GeoMap = True Then
        GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
        GDform1.Picture2.FillStyle = 0
        If mode& = 0 Then
            'GDform1.Picture2.FillColor = PointColor&
            'GDform1.Picture2.Line (Xtwip - 15, Ytwip - 15)-(Xtwip + 15, Ytwip + 15), PointColor&, BF
'            GDform1.Picture2.Line (Xtwip - 30, Ytwip - 30)-(Xtwip + 30, Ytwip + 30), PointColor&
'            GDform1.Picture2.Line (Xtwip - 30, Ytwip + 30)-(Xtwip + 30, Ytwip - 30), PointColor&
            GDform1.Picture2.FillColor = QBColor(10)
            GDform1.Picture2.Circle (Xtwip, Ytwip), Max(2, CInt(DigiZoom.LastZoom)), QBColor(10)
        ElseIf mode& = 1 Then
'            GDform1.Picture2.Line (Xtwip - 45, Ytwip)-(Xtwip + 45, Ytwip), LineColor&
'            GDform1.Picture2.Line (Xtwip, Ytwip)-(Xtwip, Ytwip - 45), LineColor&
            GDform1.Picture2.FillColor = QBColor(11)
            GDform1.Picture2.Circle (Xtwip, Ytwip), Max(2, CInt(DigiZoom.LastZoom)), QBColor(11)
        ElseIf mode& = 2 Then
            GDform1.Picture2.FillColor = QBColor(13)
            GDform1.Picture2.Circle (Xtwip, Ytwip), Max(2, CInt(DigiZoom.LastZoom)), QBColor(13)
        ElseIf mode& = 3 Then 'unknown type
            'GDform1.Picture2.FillColor = QBColor(0)
            'GDform1.Picture2.Circle (Xtwip, Ytwip), 30, UnknownColor&
'            GDform1.Picture2.Line (Xtwip - 15, Ytwip - 15)-(Xtwip + 15, Ytwip + 15), UnknownColor&, B
            GDform1.Picture2.FillColor = QBColor(14)
            GDform1.Picture2.Circle (Xtwip, Ytwip), Max(2, CInt(DigiZoom.LastZoom)), QBColor(14)
            End If
'    ElseIf TopoMap = True Then
'        GDform1.Picture2.DrawWidth = 4
'        If mode& = 0 Then
'            GDform1.Picture2.Line (Xtwip - 200, Ytwip - 200)-(Xtwip + 200, Ytwip + 200), ContourColor&
'            GDform1.Picture2.Line (Xtwip - 200, Ytwip + 200)-(Xtwip + 200, Ytwip - 200), ContourColor&
'        ElseIf mode& = 1 Then
'            GDform1.Picture2.Line (Xtwip - 200, Ytwip)-(Xtwip + 200, Ytwip), RSColor&
'            GDform1.Picture2.Line (Xtwip, Ytwip)-(Xtwip, Ytwip - 200), RSColor&
'        ElseIf mode& = 3 Then 'unknown type
'            'GDform1.Picture2.Circle (Xtwip, Ytwip), 115, UnknownColor&
'            GDform1.Picture2.Line (Xtwip - 115, Ytwip - 115)-(Xtwip + 115, Ytwip + 115), UnknownColor&, B
'            End If
        End If
    
    'GDform1.Picture2.FillStyle = oldfil&
    GDform1.Picture2.FillColor = oldfilcol&
    GDform1.Picture2.DrawMode = olddm&
    GDform1.Picture2.FillStyle = oldfs&
    
    Screen.MousePointer = vbDefault

End Sub




Public Sub ShiftMap(x As Single, Y As Single)
'This routine shifts the map in order to put the requested
'coordinate as close to the center of the picture frame as
'possible

     On Error GoTo errhand
     
     'shut off blinkers before shifting maps
     GDMDIform.CenterPointTimer.Enabled = False
        
     'pixel coordinates of the cursor is
     ITMx0 = x / twipsx
     ITMy0 = Y / twipsy
     
     'we want it to be at middle of Picture1, i.e., at
     ITMx1 = GDform1.Picture1.Width / 2
     ITMy1 = GDform1.Picture1.Height / 2
     
     'Shift the scroll bars in order to accomplish the above
     h1 = ITMx0 - ITMx1 '<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>
     If h1 < GDform1.HScroll1.min Or h1 > GDform1.HScroll1.Max Then
        If (drag1x = drag2x And drag1y = drag2y) Then
'           'response = MsgBox("Sorry, your choice would move the map beyond it's boundaries!", vbCritical + vbOKOnly, "GDB")
'
           'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
           If GDform1.Picture2.Width > GDform1.HScroll1.Width Then

              If h1 < GDform1.HScroll1.min Then
                 GDform1.HScroll1.value = GDform1.HScroll1.min
              ElseIf h1 > GDform1.HScroll1.Max Then
                 GDform1.HScroll1.value = GDform1.HScroll1.Max
                 End If

              End If
'           Exit Sub
        Else 'check if this is end of drag operation that defines box dimensions
           Exit Sub
           End If
     ElseIf (drag1x = drag2x And drag1y = drag2y) Then
        GDform1.HScroll1.value = h1
        End If
     
     h2 = ITMy0 - ITMy1
     If h2 < 0 Or h2 > GDform1.VScroll1.Max Then
        If (drag1x = drag2x And drag1y = drag2y) Then
           'response = MsgBox("Sorry, your choice would move the map beyond it's boundaries!", vbCritical + vbOKOnly, "GDB")
           
'           'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
           If GDform1.Picture2.Height > GDform1.VScroll1.Height Then

              If h2 < GDform1.VScroll1.min Then
                 GDform1.VScroll1.value = GDform1.VScroll1.min
              ElseIf h2 > GDform1.VScroll1.Max Then
                 GDform1.VScroll1.value = GDform1.VScroll1.Max
                 End If

              End If
'
'           Exit Sub
        Else 'check if this is end of drag operation that defines box dimensions
           Exit Sub
           End If
     ElseIf (drag1x = drag2x And drag1y = drag2y) Then
        GDform1.VScroll1.value = h2
        End If
        
    If Not DigitizeContour And Not DigitizeOn And Not DigitizerEraser Then 'show blinking cursor
        
        'now draw mark at center point if requested
         '************CENTER CIRCLE************
         'erase old plot mark, and plot new one if not already blinked off
         'this is used only when left-clicking on the map on during print-previewing
         If CenterBlinkState And ce& = 1 Then
            Call DrawPlotMark(0, 0, 1)
            End If
         
         'draw new center circle if ce& = 1
         'This occurs when the maps are made visible
         'and when left-clicking on the map, and for print previewing
         If ce& = 1 Then Call DrawPlotMark(x, Y, 0)
         
         'reenable blinkers
         GDMDIform.CenterPointTimer.Enabled = True
         
    Else
         'still keep track of center of picture
         blink_mark.x = x
         blink_mark.Y = Y
         
         End If
     
     newblit = True
     
'    If DigitizePadVis And (DigitizeLine Or DigitizeContour Or DigitizePoint) Then
'       BringWindowToTop (GDDigitizerfrm.hWnd)
'       End If
     
   Exit Sub
   
errhand:
   
   Screen.MousePointer = vbDefault

   Select Case Err.Number
      Case 480
         If IgnoreAutoRedrawError% = 0 Then
            MsgBox "The pixel size of this map is too big for your memory!" & vbLf & vbLf & _
                   "If you wish to use this map and ignore such errors," & vbLf & _
                   "then check the ""Ignore AutoRedraw errors"" in the" & vbLf & _
                   """Settings"" tab of ""Path/Options"" form.", vbExclamation + vbOKOnly, "MapDigitizer"
            Exit Sub
         Else 'ignore this error
            Resume Next
            End If
      Case Else
        MsgBox "Encountered error #: " & Err.Number & vbLf & _
               Err.Description & vbLf & _
               "in module: ShiftMap", vbCritical + vbOKOnly, "MapDigitizer"
   End Select
End Sub
Sub ShowError()
'This routine displays the missing files detected in
'the initialization of the GSI program

'   msgerr$ = gsEmpty
'   If linked = False Then
'      msgerr$ = "Path to the Active database incorrect or undefined!" & vbLf & _
'                "Use the Files: Paths/Options menu to help find it."
'      End If
'   If linkedOld = False Then
'      msgerr$ = msgerr$ & vbLf & vbLf & _
'                "Path to the Scanned database incorrect or undefined!" & vbLf & _
'                "Use the Files: Paths/Options menu to help find it."
'      End If
'   If heights = False Then
'      msgerr$ = msgerr$ & vbLf & vbLf & _
'                "Path to the ASTER or 25m DTM incorrect or undefined!" & vbLf & _
'                "Heights can't be displayed." & vbLf & _
'                "Use the Files: Paths/Options menu to help find the paths."
'      End If
'   If topos = False Then 'Or XResol > 1152 Then
''      msgerr$ = msgerr$ & vbLf & vbLf & _
''                "Path to the 1:50000 topo maps incorrect or undefined!" & vbLf & _
''                "or the screen resolution is greater than 1152 x 864." & vbLf & _
''                "This option won't be enabled." & vbLf & _
''                "Use the Files: Paths/Options menu to help find them."
'      msgerr$ = msgerr$ & vbLf & vbLf & _
'                "Path to the 1:50000 topo maps incorrect or undefined!" & vbLf & vbLf & _
'                "This option won't be enabled." & vbLf & _
'                "Use the Files: Paths/Options menu to help find them."
'      End If
'   If topos And XResol > 1152 Then
''      msgerr$ = msgerr$ & vbLf & vbLf & _
''                "The screen resolution is greater than 1152 x 864." & vbLf & _
''                "The 1:50000 topo maps option will be disenabled." & vbLf & _
''                "Change your screen resolution to correct this problem."
'      msgerr$ = msgerr$ & vbLf & vbLf & _
'                "The screen resolution is greater than 1152 x 864." & vbLf & _
'                "The map tiles will not fill the map canvas." & vbLf & _
'                "Lower your screen resolution to correct this problem."
'                'GDMDIform.Toolbar1.Buttons(3).Enabled = False
'                'topos = False
'      End If
'   If arcs = False Then
'      msgerr$ = msgerr$ & vbLf & vbLf & _
'                "Path to ArcMap incorrect or undefined!" & vbLf & _
'                "You won't be able to run that program." & vbLf & _
'                "Use the Files: Paths/Options menu to help find it."
'      End If
      
   If google = False Then
      msgerr$ = msgerr$ & vbLf & vbLf & _
                "Path to Google Earth incorrect or undefined!" & vbLf & _
                "You won't be able to run that program." & vbLf & _
                "Use the Files: Paths/Options menu to help find it."
      End If
      
'   If acc = False Then
'
'      If Not ActivatedVersion Then
'
'         msgerr$ = msgerr$ & vbLf & vbLf & _
'                   "MS Access option is not activated " & vbLf & _
'                   "Use the Files: Paths/Options menu to activate it."
'      Else
'
'         msgerr$ = msgerr$ & vbLf & vbLf & _
'                   "Path to MS Access incorrect or undefined!" & vbLf & _
'                   "You won't be able to run that program." & vbLf & _
'                   "Use the Files: Paths/Options menu to change this."
'                   End If
'   Else
'      'enable the Access button
'      GDMDIform.Toolbar1.Buttons(11).Enabled = True
'      End If
      
   If msgerr$ <> gsEmpty And ReportPaths& = 0 Then 'display the error message
      'if ReportPaths& = 1 then user requested to hide error messages
      MsgBox msgerr$, vbExclamation + vbOKOnly, "MapDigitizer"
      End If

End Sub



Sub ShowGeoMap(modeshow As Integer)
'This routine is used to display the large scale Geologic Map
    
    On Error GoTo errhand
    
    Select Case modeshow
       Case 0 'show geo map
            If TopoMap = True Then
            
               Unload GDform1 'unload last map
               TopoMap = False
               blink_mark.x = 0: blink_mark.Y = 0
               
              'stop blinking search points for 1:50000 maps
               GDMDIform.CenterPointTimer.Enabled = False
               ce& = 0 'reset flag that draws blinking cursor
               
               End If
            GeoMap = True
            
            'restore Geo map parameters
            picnam$ = picnam0$
            ULGeoX = x10
            ULGeoY = y10
            LRGeoX = x20
            LRGeoY = y20
            pixwi = pixwi0
            pixhi = pixhi0
            
            
            'PictureClip1.Picture = LoadPicture(picnam$)
            GDMDIform.Picture4.Visible = True
            GDform1.Visible = True
    
            'make memory copy of geo map file to be used for the eraser tool
            Set oGestionImageSrc.PictureBox = GDform1.Picture2
    
            'shift the map to the last recorded point
            On Error Resume Next
            Call InputPositionFile(ITMx, ITMy, hgt)
   
            'display recorded coordinates in goto edit boxes
            GDMDIform.Text5.Text = Fix(ITMx)
            GDMDIform.Text6.Text = Fix(ITMy)
            GDMDIform.Text7.Text = hgt
            
            If Not Digitizing Then
                ITMx0 = ((ITMx - ULGeoX) / (LRGeoX - ULGeoX)) * pixwi
                ITMy0 = ((ULGeoY - ITMy) / (ULGeoY - LRGeoY)) * pixhi
                Dim x As Single, Y As Single
                x = ITMx0 * twipsx
                Y = ITMy0 * twipsy
            Else
               x = ITMx
               Y = ITMy
               End If
            'move the map so this point is as close as possible
            'to center
            ce& = 1 'request that center mark be drawn when map becomes visible
            Call ShiftMap(x, Y)
            
            GDMDIform.CenterPointTimer.Enabled = True
            'ce& = 1
       
       Case 1 'remove Geo map from screen
            If g_ier = -2 Then
              GDMDIform.Picture4.Visible = False
              Unload GDform1
              g_ier = 0
            Else
               GDform1.Visible = False
                GDMDIform.Picture4.Visible = False
                Unload GDform1
               End If
            GeoMap = False
            blink_mark.x = 0: blink_mark.Y = 0 'erase memory of last center point
       
            GDMDIform.CenterPointTimer.Enabled = False
            ce& = 0 'reset blinking cursor flag
       
       Case Else
    End Select
    
   Exit Sub

errhand:
   
   Screen.MousePointer = vbDefault
   
   If g_ier = -1 Or g_ier = -2 Then Exit Sub 'skip the warning in order to exit the memory error gracefully
   
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          "in module: ShowGeoMap", vbCritical + vbOKOnly, "MapDigitizer"

End Sub

Sub ConvertPixToCoord(xin, Yin, Xout, Yout)

  'Convert coordinates to pixels
  Xcoord = xin / twipsx
  Ycoord = Yin / twipsy
  
  'Convert pixel coordinates to ITM
  Xout = ((LRGeoX - ULGeoX) / pixwi) * Xcoord + ULGeoX
  Yout = ULGeoY - ((ULGeoY - LRGeoY) / pixhi) * Ycoord
  
End Sub

Public Sub casgeo(kmx, kmy, lg, lt)
        G1# = kmy - 1000000
        G2# = kmx
        R# = 57.2957795131
        B2# = 0.03246816
        f1# = 206264.806247096
        s1# = 126763.49
        S2# = 114242.75
        e4# = 0.006803480836
        C1# = 0.0325600414007
        c2# = 2.55240717534E-09
        c3# = 0.032338519783
        xc1# = 1170251.56
        yc1# = 1126867.91
        yc2# = G1#
'       GN & GE
        xc2# = G2#
        If (xc2# > 700000#) Then GoTo ca5
        xc1# = xc1# - 1000000#
ca5:    If (yc2# > 550000#) Then GoTo ca10
        yc1# = yc1# - 1000000#
ca10:   xc1# = xc2# - xc1#
        yc1# = yc2# - yc1#
        D1# = yc1# * B2# / 2#
        O1# = S2# + D1#
        O2# = O1# + D1#
        A3# = O1# / f1#
        A4# = O2# / f1#
        B3# = 1# - e4# * Sin(A3#) ^ 2#
        B4# = B3# * Sqr(B3#) * C1#
        C4# = 1# - e4# * Sin(A4#) ^ 2#
        C5# = Tan(A4#) * c2# * C4# ^ 2#
        C6# = C5# * xc1# ^ 2#
        D2# = yc1# * B4# - C6#
        C6# = C6# / 3#
'LAT
        l1# = (S2# + D2#) / f1#
        R3# = O2# - C6#
        R4# = R3# - C6#
        r2# = R4# / f1#
        A2# = 1# - e4# * Sin(l1#) ^ 2#
        lt = R# * (l1#)
        A5# = Sqr(A2#) * c3#
        d3# = xc1# * A5# / Cos(r2#)
' LON
        lg = R# * ((s1# + d3#) / f1#)
'       THIS IS THE EASTERN HEMISPHERE!
        lg = -lg
        If GpsCorrection Then
           'Use the approximate conversion factor
           'from Clark (1888) to WGS84 (GPS).
           lg = lg - 0.0013
           lt = lt + 0.0013
           End If
        

End Sub
Public Sub GEOCASC(L11, L22, G11, G22)
      Dim D1 As Double, D2 As Double, D5 As Double, E3 As Double
      Dim G1 As Double, G2 As Double, G3 As Double
      Dim l1 As Double, l2 As Double, l3 As Double, l4 As Double
      Dim s1 As Double, S2 As Double, AL As Double
      Dim m1 As Integer, GpsCorrOff As Boolean
      If GpsCorrection Then
         'Use the approximate correction factor
         'in order to agree with GPS.
         L22 = L22 - 0.0013
         L11 = L11 - 0.0013
         GpsCorrOff = True
         GpsCorrection = False
         End If
      l1 = L11
      l2 = L22
      G1 = G11
      G2 = G22
      m1 = 0
      G1 = 100000#
      G2 = 100000#
      s1 = 31.4896370431
      S2 = 34.4727144086
      E3 = 0.0001
      l3 = l1
      l4 = l2
5     AL = (s1 + l3) / (2# * 57.2957795131)
      G3 = fnm(AL)
      G4 = fnp(AL)
      D1 = (l3 - s1) * G3
      D2 = (l4 - S2) * G4
      G1 = G1 + D1
      G2 = G2 + D2
      m1 = m1 + 1
      D5 = Sqr(D1 ^ 2 + D2 ^ 2)
      If (D5 < E3) Then GoTo 10
      If (m1 > 10) Then GoTo 15
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
      If GpsCorrOff Then GpsCorrection = True
      Exit Sub
15
      If GpsCorrOff Then GpsCorrection = True
      MsgBox "GEOCASC didn't converge!", vbCritical + vbOKOnly, "MapDigitizer"
End Sub
Public Function fnp(p As Double) As Double
'  Length of a degree parallel: P must be latitude in radians
   'Dim P As Double
   fnp = 111415.13 * Cos(p) - 94.55 * Cos(3# * p) + 0.012 * Cos(5# * p)
End Function

Public Function fnm(p As Double) As Double
'  Length of a degree meridian: P must be latitude in radians
   'Dim P As Double
   fnm = 111132.09 - 566.05 * Cos(2# * p) + 1.2 * Cos(4# * p) - 0.002 * Cos(6# * p)
End Function
Sub EraseOldSearchPoints()

  'erase old search points plotted on the map
  
  EraseMaps = True
  If TopoMap Then
     'erase old search points by repainting the maps
     Call gotocoord

     If ScreenDump Then 'refresh search points
        If NumReportPnts& <> 0 Then
          For i& = 1 To NumReportPnts&
            XPnt = ReportPnts(0, i& - 1)
            YPnt = ReportPnts(1, i& - 1)
            If GeoMap = True Or TopoMap = True Then
               Call DrawPoint(XPnt, YPnt, Int(ReportPnts(2, i& - 1))) 'draw the point on the map
               End If
          Next i&
         End If
       End If
     
  ElseIf GeoMap Then
    'redraw the search points or reload map to erase

    'accomplish this simply by reloading the picture
    GDform1.Picture2.Picture = LoadPicture(picnam$)
     
 End If
 
EraseMaps = False
  
End Sub

Sub DrawPlotMark(x As Single, Y As Single, ModeMark As Integer)
   'This routine draws and erases center click mark
   
   Select Case ModeMark
      Case 1 'erase old mark
            
            If DigiZoomed Then 'don't erase since blited new picture
               DigiZoomed = False
               Exit Sub
               End If
            
            'oldfil& = GDform1.Picture2.FillStyle
            oldfilcol& = GDform1.Picture2.FillColor
            olddm& = GDform1.Picture2.DrawMode
            If blink_mark.x <> 0 Or blink_mark.Y <> 0 Then
               GDform1.Picture2.DrawMode = 7
               'GDform1.Picture2.FillStyle = 1
               GDform1.Picture2.DrawWidth = 5
               GDform1.Picture2.FillColor = QBColor(15)
               GDform1.Picture2.Circle (CLng(blink_mark.x), CLng(blink_mark.Y)), 115 * (twipsx / 15), QBColor(15)
               CenterBlinkState = Not CenterBlinkState 'blink
               'GDform1.Picture2.Circle (Xo_blink_Mark, Yo_blink_Mark), 75, QBColor(0)
               'GDform1.Picture2.Circle (Xo_blink_Mark, Yo_blink_Mark), 35, QBColor(15)
               End If
            'GDform1.Picture2.FillStyle = oldfil&
            GDform1.Picture2.FillColor = oldfilcol&
            GDform1.Picture2.DrawMode = olddm&
               
      Case 0 'plot new mark
      
            'draw circle at click point
            'oldfil& = GDform1.Picture2.FillStyle
            oldfilcol& = GDform1.Picture2.FillColor
            olddm& = GDform1.Picture2.DrawMode
            
            GDform1.Picture2.DrawMode = 7
            'GDform1.Picture2.FillStyle = 1
            GDform1.Picture2.DrawWidth = 5
            GDform1.Picture2.FillColor = QBColor(15)
            GDform1.Picture2.Circle (CLng(x), CLng(Y)), 115 * (twipsx / 15), QBColor(15)
            CenterBlinkState = True 'reset blinker to on
            'GDform1.Picture2.Circle (X, Y), 75, QBColor(0)
            'GDform1.Picture2.Circle (X, Y), 35, QBColor(14)
            blink_mark.x = x: blink_mark.Y = Y
            
            'GDform1.Picture2.FillStyle = oldfil&
            GDform1.Picture2.FillColor = oldfilcol&
            GDform1.Picture2.DrawMode = olddm&
      
      Case Else
   End Select

End Sub
'---------------------------------------------------------------------------------------
' Procedure : PlotNewSearchPoints
' DateTime  : 6/13/2004 20:21
' Author    : Chaim Keller
' Purpose   : Plot desired search points
'---------------------------------------------------------------------------------------
'
Sub PlotNewSearchPoints()
    
    Dim IgnoreDuplicates As Boolean 'flag for ignoring multiple samples at one coordinate
    Dim IgnoreOffGeoMapBoundary As Boolean 'flags for ignoring off map boundaries
    Dim IgnoreOffTopoMapBoundary As Boolean
    Dim IgnoreMaxPoints
    Dim ier As Integer
    
   On Error GoTo PlotNewSearchPoints_Error

    'initialize flag values
    IgnoreDuplicates = False
    IgnoreOffGeoMapBoundary = False
    IgnoreOffTopoMapBoundary = False
    IgnoreMaxPoints = False
    
'    If EditDBVis Then 'user may be editing records that are off the
'       'map so don't give warning messages
'       IgnoreDuplicates = True
'       IgnoreOffGeoMapBoundary = True
'       IgnoreOffTopoMapBoundary = True
'       CheckDuplicatePoints = True
'       End If

    ier = ReDrawMap(0)
    If Not InitDigiGraph Then
       InputDigiLogFile 'load up saved digitizing data for the current map sheet
    Else
       ier = RedrawDigiLog
       End If
    
    NumReportPnts& = 0
    ReDim ReportPnts(2, 0) 'clear memory of search plot points
    
    'determine if there is more then one selected point
     numSelectedPnts& = -1
     For i& = 1 To numReport& 'search over all the search results
          If GDReportfrm.lvwReport.ListItems(i&).Selected Then
             numSelectedPnts& = numSelectedPnts& + 1
             If numSelectedPnts& >= 1 Then Exit For
             End If
     Next i&
        
    'find selected points
     For i& = 1 To numReport& 'search over all the search results
     
          If GDReportfrm.lvwReport.ListItems(i&).Selected Then
              XPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(2))
              YPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(3))
              
              'Check for duplicate coordinates.
              'This is a no-no, since if the same point is plotted twice
              'then it erases itself for DrawMode = 7.
              'However, as a compromise just check the last point, since
              'this works good for a small number of points, and for a large
              'number of points, any erasures are not noticeable
              If NumReportPnts& < 2 Then
                 BeginSearchPnt& = 1
              Else
                 BeginSearchPnt& = NumReportPnts& - 1
                 End If
                 
              For j& = BeginSearchPnt& To NumReportPnts&
                 If XPnt = ReportPnts(0, j& - 1) And YPnt = ReportPnts(1, j& - 1) Then
                    If Not CheckDuplicatePoints Then
                       If Not IgnoreDuplicates Then
                            'give message one time, i.e., at first press of "Locate highlighted records on map" button (GDReportfrm)
                             response = MsgBox("More than one record is located at the coordinates:" & vbLf & _
                                "Xpix = " & str(XPnt) & "; Ypix = " & str(YPnt) & vbLf & _
                                "Only the first record with that coordinate will be plotted!" & vbLf & _
                                "Do you wish to handle any other duplicates in this manner?", vbExclamation + vbYesNoCancel, "MapDigitizer")
                        Else
                           response = vbYes 'don't announce duplicates
                           End If
                           
                        If response = vbCancel Then
                           StopPlotting = True
                           GoTo s100
                           
                        Else
                        
                           If response = vbYes And Not IgnoreDuplicates Then
                              IgnoreDuplicates = True 'don't announce other duplicates
                              End If
                              
                           'unselect duplicate
                           GDReportfrm.lvwReport.ListItems(i&).Selected = False
                           If numHighlighted& > 0 Then 'remove it from the highlight array
                              Highlighted(i& - 1) = 0
                              End If
                           End If
                        End If
                    GoTo s50
                    End If
              Next j&
              
              'If got here then it can be plotted, so
              'add it to the plot array and plot it
              NumReportPnts& = NumReportPnts& + 1
              'check that this number doesn't exceed setting for maximum
              'number of plotted points
              ReDim Preserve ReportPnts(2, NumReportPnts& - 1) 'redimension array
              ReportPnts(0, NumReportPnts& - 1) = XPnt
              ReportPnts(1, NumReportPnts& - 1) = YPnt
              'decide which plot mark to draw
              If InStr(GDReportfrm.lvwReport.ListItems(i&).Text, "point") Then
                 ReportPnts(2, NumReportPnts& - 1) = 1
              ElseIf InStr(GDReportfrm.lvwReport.ListItems(i&).Text, "line") Then
                 ReportPnts(2, NumReportPnts& - 1) = 0
              ElseIf InStr(GDReportfrm.lvwReport.ListItems(i&).Text, "Contour") Then
                 ReportPnts(2, NumReportPnts& - 1) = 3
              ElseIf InStr(GDReportfrm.lvwReport.ListItems(i&).Text, "Erased") Then
                 ReportPnts(2, NumReportPnts& - 1) = 2
                 End If
              'draw point on map
              If GeoMap = True Or TopoMap = True And Not IgnoreMaxPoints Then
                 Call DrawPoint(CLng(XPnt * DigiZoom.LastZoom), CLng(YPnt * DigiZoom.LastZoom), Int(ReportPnts(2, NumReportPnts& - 1))) 'draw the point on the map
                 End If
                 
              If NumReportPnts& + 1 > numMaxHighlight Then
                 If Not IgnoreMaxPoints Then 'give warning
                    MsgBox "Have exceeded your setting for the maximum allowed number" & vbLf & _
                        "of plotted points (" & Trim$(str$(numMaxHighlight&)) & " points)!" & vbLf & vbLf & _
                        "You can change this setting on the Paths/Options menu." & vbLf & vbLf & _
                        "As of now, no more points will be plotted.", _
                        vbExclamation + vbOKOnly, "MapDigitizer"
                        IgnoreMaxPoints = True
                    'unselect rest of search records
                    For j& = i& To numReport&
                       GDReportfrm.lvwReport.ListItems(j&).Selected = False
                    Next j&
                    GoTo s100
                    End If
                  GoTo s100
                  End If
                 
          End If
s50:
      Next i&
      
s100: CheckDuplicatePoints = True 'already checked, so don't check again

   On Error GoTo 0
   Exit Sub

PlotNewSearchPoints_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbLf & _
           "in procedure PlotNewSearchPoints of Module modGDModule." & vbLf & vbLf & _
           "Record this error and notify your system manager.", _
           vbCritical + vbOKOnly, App.Title

End Sub
Sub ClearHighlightedPoints()
    'clear all the highlighted records in the list view of GDReportfrm
    For i& = 1 To numReport&
        If GDReportfrm.lvwReport.ListItems(i&).Selected Then
           GDReportfrm.lvwReport.ListItems(i&).Selected = False
           End If
    Next i&
End Sub
''---------------------------------------------------------------------------------------
'' Procedure : ShowDetailedReport
'' DateTime  : 11/23/2002 18:45
'' Purpose   : Routine that controls the loading up of the forms
''             of the Detailed Report Form
''---------------------------------------------------------------------------------------
''
'Sub ShowDetailedReport()
'
'    On Error GoTo errhand
'
'    'populate the DetailedReport Form with the search results
'    'after right click on search report record
'    'make GDdetailreportfrm visible with all its windows
'    'also add the other records at these coordinates and name
'
'    Dim PlaceITMx As Single, PlaceITMy As Single
'
'    PlaceITMx = GDReportfrm.lvwReport.ListItems(NearestPnt&).SubItems(1)
'    PlaceITMy = GDReportfrm.lvwReport.ListItems(NearestPnt&).SubItems(2)
'
'    GDDetailReportfrm.lvwDetailReport.ListItems.Clear
'
'    'find all search record with the above coordinates
'    Dim i&, recFirst&
'
'    recFirst& = 0
'    For i& = NearestPnt& - 1 To 1 Step -1
'       If val(GDReportfrm.lvwReport.ListItems(i&).SubItems(1)) = PlaceITMx And _
'          val(GDReportfrm.lvwReport.ListItems(i&).SubItems(2)) = PlaceITMy Then
'          'found another record
'          recFirst& = i&
'       Else
'          Exit For 'no more adjacent records from the same place
'          End If
'    Next i&
'
'    If recFirst& <> 0 And recFirst& < NearestPnt& Then
'       NearestPnt& = recFirst& 'found earlier records
'       End If
'
'    'now load the records into the listview
'    For i& = NearestPnt& To numReport&
'       If val(GDReportfrm.lvwReport.ListItems(i&).SubItems(1)) = PlaceITMx And _
'          val(GDReportfrm.lvwReport.ListItems(i&).SubItems(2)) = PlaceITMy Then
'          'this record is at same place
'          Call LoadDetailedReportListView(i&) 'load this record into the list view
'       Else
'          Exit For 'no more adjacent records from the same place
'          End If
'    Next i&
'
'   'ensure that first item is visible
'    GDDetailReportfrm.lvwDetailReport.ListItems(1).EnsureVisible
'
'   'now load up treeView Control for current record in list
'   If DetailRecordNum& = 0 Then
'      Call LoadTreeView(NearestPnt&)
'      GDDetailReportfrm.lvwDetailReport.ListItems(1).Selected = True
'   Else
'      Call LoadTreeView(DetailRecordNum&)
'      GDDetailReportfrm.lvwDetailReport.ListItems(DetailRecordNum& - NearestPnt& + 1).Selected = True
'      End If
'
'  Exit Sub
'
'errhand:
'
'    Screen.MousePointer = vbDefault
'
'    MsgBox "Encountered error #: " & Err.Number & vbLf & _
'           Err.Description & vbLf & _
'           "You probably won't be able to obtain a complete detailed report!" & vbLf & _
'           sEmpty, vbExclamation + vbOKOnly, "MapDigitizer"
'
'
'End Sub
'Sub PopulateListView(FossilTag$, FossilTable$, FosIcon$)
'
'   On Error GoTo errhand
'
'   Dim qdfDV As QueryDef
'   Dim rstDV As Recordset
'   Dim strSqlDV As String
'   Dim sPreAge As String
'
'   'load up the colum headers in the Detailed Report
'   'and clear previous info
'   With GDDetailReportfrm
'
'     .lvwDetailReport.ListItems.Clear
'     .lvwDetailReport.ColumnHeaders.Clear
'     'set up headers for List View
'     .lvwDetailReport.ColumnHeaders.Add , , "Analyst", 2000
'     .lvwDetailReport.ColumnHeaders.Add , , "date", 1100
'     .lvwDetailReport.ColumnHeaders.Add , , "preageo", 800
'     .lvwDetailReport.ColumnHeaders.Add , , "Earlier Age Date", 1500
'     .lvwDetailReport.ColumnHeaders.Add , , "preagey", 800
'     .lvwDetailReport.ColumnHeaders.Add , , "Later Age Date", 1500
'     .lvwDetailReport.ColumnHeaders.Add , , "prezoneo", 900
'     .lvwDetailReport.ColumnHeaders.Add , , "Earlier Zone", 2500
'     .lvwDetailReport.ColumnHeaders.Add , , "prezoney", 900
'     .lvwDetailReport.ColumnHeaders.Add , , "Later Zone", 2500
'
'
'     .lvwDetailReport.ColumnHeaders(1).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(2).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(3).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(4).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(5).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(6).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(7).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(8).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(9).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(10).Alignment = lvwColumnLeft
'
'     .lvwDetailReport.LabelEdit = lvwManual
'
'   End With
'
'  'The user clicked a branch in the DetailedReport Tree. So
'  'query the database for the requested detailed information
'  'and then populate the ListView in that form with the results
'   If InStr(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), "*") Then
'      'this is record from scanned database and there is only one table
'      'to query
'      pos1& = InStr(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), "*")
'      pos2& = InStr(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), "/")
'      OKEY$ = Mid$(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), _
'                 pos1& + 1, pos2& - pos1& - 1)
'      GoTo Olddb
'
'   Else 'record from active database
'      pos1& = InStr(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(3), FossilTag$)
'      pos2& = InStr(pos1& + 1, GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(3), "),")
'      'the res_id (Display Number) of the relevant fossil tables is idvalue
'      idvalue$ = Mid$(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(3), _
'                    pos1& + Len(FossilTag$) + 2, pos2& - pos1& - Len(FossilTag$) - 2)
'      End If
'
'
'   Dim FossCat As String
'   Dim sqfoss As String
'   Dim rstfoss As Recordset
'   Dim Name2 As String
'   Dim Name1 As String
'
'   'create query string
'   'strSqlDV = "SELECT * FROM " & FossilTable$ & " WHERE res_id = " & idvalue$
'   strSqlDV = "SELECT " & FossilTable$ & _
'                ".*, CHKPALdic.* From (" & FossilTable$ & _
'                " INNER JOIN CHKPALdic ON " & FossilTable$ & _
'                ".[analist] = CHKPALdic.[id]) WHERE " & _
'                FossilTable$ & "![res_id] = " & idvalue$
'
'   'query the database
'   Set qdfDV = gdbs.CreateQueryDef(sEmpty, strSqlDV & ";")
'
'   'Create a temporary snapshot-type Recordset.
'   Set rstDV = qdfDV.OpenRecordset(dbOpenSnapshot)
'
'   'populate the list view with this result
'
'   'Analyst name
'   Set mitem = GDDetailReportfrm.lvwDetailReport.ListItems.Add()
'   mitem.SmallIcon = FosIcon$
'
'   rstDV.MoveFirst
'   If IsNull(rstDV![Name]) Then
'      mitem.Text = sEmpty
'   Else
'      mitem.Text = Trim$(rstDV![Name])
'      End If
'
'   'Analysis date
'   If IsNull(rstDV![Date]) Then
'      mitem.SubItems(1) = sEmpty
'   Else
'      mitem.SubItems(1) = rstDV![Date]
'      End If
'
'   'preageo field
'   If IsNull(rstDV![preageo]) Then
'      mitem.SubItems(2) = sEmpty
'   Else
'      num& = rstDV![preageo]
'      GoSub findprezone
'      mitem.SubItems(2) = sPreAge
'      End If
'
'   'Earlier age field
'   If IsNull(rstDV![ch_o_age]) Then
'      mitem.SubItems(3) = sEmpty
'   Else
'      mitem.SubItems(3) = rstDV![ch_o_age]
'      End If
'
'   'preagey field
'   If IsNull(rstDV![preagey]) Then
'      mitem.SubItems(4) = sEmpty
'   Else
'      num& = rstDV![preagey]
'      GoSub findprezone
'      mitem.SubItems(4) = sPreAge
'      End If
'
'   'Later age field
'   If IsNull(rstDV![ch_y_age]) Then
'      mitem.SubItems(5) = sEmpty
'   Else
'      mitem.SubItems(5) = rstDV![ch_y_age]
'      End If
'
'   'prezoneo field
'   If IsNull(rstDV![prezoneo]) Then
'      mitem.SubItems(6) = sEmpty
'   Else
'      num& = rstDV![prezoneo]
'      GoSub findprezone
'      mitem.SubItems(6) = sPreAge
'      End If
'
'   'Earlier zone
'   If IsNull(rstDV![zoneo]) Or val(rstDV![zoneo]) = 0 Then
'      mitem.SubItems(7) = sEmpty
'   Else
'      'load earlier genus and species name (zone)
'      num& = rstDV![zoneo]
'      GoSub findzone
'      mitem.SubItems(7) = Trim$(zone$)
'      End If
'
'   'prezoney field
'   If IsNull(rstDV![prezoney]) Then 'Later pre zone
'      mitem.SubItems(8) = sEmpty
'   Else
'      num& = rstDV![prezoney]
'      GoSub findprezone
'      mitem.SubItems(8) = sPreAge
'      End If
'
'   'Later zone field
'   If IsNull(rstDV![zoney]) Or val(rstDV![zoney]) = 0 Then
'      mitem.SubItems(9) = sEmpty
'   Else
'      'load later genus and species name (zone)
'      num& = rstDV![zoney]
'      GoSub findzone
'      mitem.SubItems(9) = Trim$(zone$)
'      End If
'
'   rstDV.Close
'
'   Exit Sub
'
''-----------begin section for Scanned (Incactive) database
'Olddb:
'
'   'Geologic Age info for record from scanned database
'
'   'SQL query string
'   strSqlDV = "SELECT OBJECTS2.* FROM OBJECTS2 WHERE O_KEY = " & OKEY$
'
'   'query the database
'   Set qdfDV = gdbsOld.CreateQueryDef(sEmpty, strSqlDV & ";")
'
'   'Create a temporary snapshot-type Recordset.
'   Set rstDV = qdfDV.OpenRecordset(dbOpenSnapshot)
'
'   'populate the list view with this result
'   rstDV.MoveFirst
'
'   Set mitem = GDDetailReportfrm.lvwDetailReport.ListItems.Add()
'   mitem.SmallIcon = FosIcon$
'
'   'no Analyst name in this database
'    mitem.Text = sEmpty
'
'   'date
'   If IsNull(rstDV![N01]) Then
'      mitem.SubItems(1) = sEmpty
'   ElseIf val(rstDV![N01]) = 0 Then
'      mitem.SubItems(1) = sEmpty
'   Else
'      mitem.SubItems(1) = Mid$(rstDV![N01], 7, 2) & "/" & _
'                          Mid$(rstDV![N01], 5, 2) & "/" & _
'                          Mid$(rstDV![N01], 1, 4)
'      End If
'
'   'preageo field
'   If IsNull(rstDV![N06]) Then
'      mitem.SubItems(2) = sEmpty
'   ElseIf val(rstDV![N06]) = 0 Then
'      mitem.SubItems(2) = sEmpty
'   Else
'      mitem.SubItems(2) = arrN06(val(rstDV![N06]) - 1)
'      End If
'
'   'Earlier age field
'   If IsNull(rstDV![N07]) Then
'      mitem.SubItems(3) = sEmpty
'   ElseIf val(rstDV![N07]) = 0 Then
'      mitem.SubItems(3) = sEmpty
'   Else
'      mitem.SubItems(3) = arrN07(val(rstDV![N07]) - 1)
'      End If
'
'   'preagey field
'   If IsNull(rstDV![N08]) Then
'      mitem.SubItems(4) = sEmpty
'   ElseIf val(rstDV![N08]) = 0 Then
'      mitem.SubItems(4) = sEmpty
'   Else
'      mitem.SubItems(4) = arrN06(val(rstDV![N08]) - 1)
'      End If
'
'   'Later age field
'   If IsNull(rstDV![N09]) Then
'      mitem.SubItems(5) = sEmpty
'   ElseIf val(rstDV![N09]) = 0 Then
'      mitem.SubItems(5) = sEmpty
'   Else
'      mitem.SubItems(5) = arrN07(val(rstDV![N09]) - 1)
'      End If
'
'   'prezoneo field
'    mitem.SubItems(6) = sEmpty
'
'   'Earlier zone
'    mitem.SubItems(7) = sEmpty
'
'   'prezoney field
'    mitem.SubItems(8) = sEmpty
'
'   'Later zone field
'    mitem.SubItems(9) = sEmpty
'
'    rstDV.Close
'
'    Exit Sub
'
''--------------------------------------------------------
'
'findprezone: 'inline gosub that determines the prezone string
'      Select Case num&
'         Case 1
'           sPreAge = "Lower"
'         Case 2, 0
'           sPreAge = sEmpty
'         Case 3
'           sPreAge = "Upper"
'         Case 4
'           sPreAge = "Middle"
'      End Select
'Return
'
'
'findzone: 'inline gosub that reads the species (zone) name
'          'from the multi species list boxes on the GDSearchfrm
'    If num& = 0 Then
'       zone$ = sEmpty
'       Return
'       End If
'
'    Select Case FossilTable$
'       Case "condores"
'          FossCat = "CONOZONEcat"
'       Case "diatores"
'          FossCat = "DIATOZONEcat"
'       Case "foramres"
'          FossCat = "FORAZONEcat"
'       Case "megares"
'          FossCat = "MEGAZONEcat"
'       Case "nanores"
'          FossCat = "NANOZONEcat"
'       Case "ostrares"
'          FossCat = "OSTRAZONcat"
'       Case "palynres"
'          FossCat = "PALIZONEcat"
'    End Select
'
'    'query fossil zone dictionary
'
'    sqfoss = "SELECT " & FossCat & ".* FROM " & FossCat & _
'             " WHERE " & FossCat & ".id = " & str$(num&)
'
'    Set rstfoss = gdbs.OpenRecordset(sqfoss, dbOpenSnapshot)
'
'    With rstfoss
'       .MoveFirst
'       'Genera
'       If Not IsNull(rstfoss![name_2]) Then
'          Name2 = rstfoss![name_2]
'       Else
'          Name2 = sEmpty
'          End If
'
'       'Species
'       If Not IsNull(rstfoss![name_1]) Then
'          Name1 = rstfoss![name_1]
'       Else
'          Name1 = sEmpty
'          End If
'
'       zone$ = LTrim$(Trim$(Name2) & " " & Trim$(Name1)) 'Genera and species
'
'    End With
'    rstfoss.Close
'
'    Return
'
'errhand:
'
'   Screen.MousePointer = vbDefault
'
'    MsgBox "Encountered error #: " & Err.Number & vbLf & _
'           Err.Description & vbLf & _
'           "You probably won't be able to obtain a complete detailed report!" & vbLf & _
'           sEmpty, vbExclamation + vbOKOnly, "MapDigitizer"
'
'
'End Sub

Sub UpdatePositionFile(Xposit, Yposit, hgtposit)
    'this routine records the last map position and
    'map animation timer interval
                   
    On Error Resume Next
                 
    filcoord& = FreeFile
    Open direct$ + "\Mapcoord.txt" For Output As #filcoord&
    hgtwrite = hgtposit
    If IsNull(hgtposit) Then hgtwrite = 0
    Write #filcoord&, "This file is used by the MapDigitizer program. Don't erase it!"
    Write #filcoord&, Xposit, Yposit, hgtwrite
'    Write #filcoord&, GDMDIform.Timer1.Interval
    Close #filcoord&

End Sub

Sub InputPositionFile(Xposit, Yposit, hgtposit)
    'this routine reads the last recorded map position
    'and map animation timer interval
    
    On Error Resume Next
    
    filcoord& = FreeFile
    Open direct$ + "\Mapcoord.txt" For Input As #filcoord&
    Input #filcoord&, doclin$
    Input #filcoord&, Xposit, Yposit, hgtposit
    If Xposit = 0 And Yposit = 0 Then
       'position in middle of picture
       Xposit = pixwi * 0.5
       Yposit = pixhi * 0.5
       End If
    If IsNull(hgtposit) Then hgtposit = 0
'    Input #filcoord&, IntervalTimer
    Close #filcoord&

End Sub



Sub SaveExcel()

   'This routine saves the search results to a file.
   'There are three possible output formats:
   
   '(1) txt file
   'this option also allows the user to load up a stored
   'search file into the report form in order to plot and
   'to obtain detailed reports that can be printed
   
   '(4) keyhole markup language file (kml)
   'this file can be displayed by Google Earth, ArcExplorer, etc.
   
   'check if report is available
'   If Not PicSum Then Exit Sub
   
'   If numReport& <= 0 Then
'      MsgBox "No records have been found yet!", vbExclamation + vbOKOnly, _
'          "MapDigitizer"
'      Exit Sub
'      End If

   Dim i&, j&, Selected As Boolean, numSelected&, numRow&
   
10 On Error GoTo errhand

   GDMDIform.CommonDialog1.FileName = sEmpty
   GDMDIform.CommonDialog1.Filter = _
       "Drawing Exchange File (*.dxf)|*.dxf|Comma separated text (*.txt)|*.txt|Google Earth kml file (*.kml)|*.kml"
   GDMDIform.CommonDialog1.FilterIndex = 1
   GDMDIform.CommonDialog1.ShowSave
   'check for existing files, and for wrong save directories
  
   If GDMDIform.CommonDialog1.FileName = sEmpty Then Exit Sub
   
      
   ext$ = RTrim$(Mid$(GDMDIform.CommonDialog1.FileName, InStr(1, _
       GDMDIform.CommonDialog1.FileName, ".") + 1, 3))
  
'  'check if some records are selected and count them
'   numSelected& = 0
'   For j& = 1 To numReport&
'      If GDReportfrm.lvwReport.ListItems(j&).Selected = True Then
'         numSelected& = numSelected& + 1
'         End If
'   Next j&
'
'   Selected = False
'   If numSelected& > 1 Then
'
'         Select Case MsgBox("Some or all of the search results are selected." _
'                            & vbCrLf & "Do you want to save just these records?" _
'                            & vbCrLf & "" _
'                            & vbCrLf & "Answer: Yes:  to save ONLY the selected records" _
'                            & vbCrLf & "               No:   to save ALL the records""" _
'                            , vbYesNo + vbQuestion + vbDefaultButton1, App.Title)
'
'            Case vbYes
'               Selected = True
'            Case vbNo
'               Selected = False
'         End Select
'
'   End If
     
   myfile = Dir(GDMDIform.CommonDialog1.FileName)
   If myfile <> sEmpty And ext$ <> "xls" Then
      response = MsgBox("Write over existing file?", vbYesNoCancel + vbQuestion, _
          "Cal Program")
      If response = vbNo Then
         GoTo 10
      ElseIf response = vbCancel Then
         Exit Sub
         End If
      End If

'-----------------------save to csv txt file-----------------
25  If ext$ = "dxf" Then

       If numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Then
          'save all the points to a dxf file if rubber sheeting done
          If DigiRubberSheeting Then
          
             Dim ier As Integer
             Dim SaveCoord(1) As POINTAPI

             SaveCoord(0).x = 0
             SaveCoord(0).Y = 0
             SaveCoord(1).x = CLng(GDform1.Picture2.Width / DigiZoom.LastZoom)
             SaveCoord(1).Y = CLng(GDform1.Picture2.Height / DigiZoom.LastZoom)
          
             ier = FindPointsHardy(SaveCoord, 0, sEmpty)
          
          Else
            Call MsgBox("You must first define a coordinate system." _
                        & vbCrLf & "" _
                        & vbCrLf & "(Hint: click on the Rubber Sheeting Button on the toolbar)" _
                        , vbInformation, "Save error message")
            
            End If
          
          End If


    ElseIf ext$ = "txt" Then

      Screen.MousePointer = vbHourglass
      
      Close
      filtm1& = FreeFile
      Open GDMDIform.CommonDialog1.FileName For Output As #filtm1&
      'write identifying information
      Print #filtm1&, "MapDigitizer Search Results, Date/Time: " & Now()
      Print #filtm1&, sEmpty
      'write number of columns and rows:
      Print #filtm1&, "[Number of Columns and Rows]"
      If Not Selected Then
         Write #filtm1&, GDReportfrm.lvwReport.ColumnHeaders.count, numReport&
      Else
         Write #filtm1&, GDReportfrm.lvwReport.ColumnHeaders.count, numSelected&
         End If
      Print #filtm1&, sEmpty
      'write column headers
      Print #filtm1&, "[Column Headers]"
      Dim sColum$
      sColum$ = sEmpty
      For i& = 1 To GDReportfrm.lvwReport.ColumnHeaders.count
         If i& = 1 Then
            sColum$ = Chr(34) & GDReportfrm.lvwReport.ColumnHeaders(i&) & Chr(34)
         Else
            sColum$ = sColum$ & "," & Chr(34) & GDReportfrm.lvwReport.ColumnHeaders(i&) & Chr(34)
            End If
      Next i&
      Print #filtm1&, sColum$
      Print #filtm1&, sEmpty
      Print #filtm1&, "[Search Results]"
      
      For j& = 1 To numReport&
        
        If Selected And GDReportfrm.lvwReport.ListItems(j&).Selected = False Then
           GoTo 100 'this record is not selected so skip it
           End If
           
        sColum$ = sEmpty
        For i& = 1 To GDReportfrm.lvwReport.ColumnHeaders.count
           If i& = 1 Then
              sColum$ = Chr(34) & GDReportfrm.lvwReport.ListItems( _
                  j&).Text & Chr(34)
           Else
              sColum$ = sColum$ & "," & Chr( _
                  34) & GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1) & Chr( _
                  34)
              End If
        Next i&
        Print #filtm1&, sColum$
100   Next j&
      Close #filtm1&
      

        
     
  ElseIf ext$ = "kml" Then
  
     Call ExportReportToGoogleEarth(GDMDIform.CommonDialog1.FileName, Selected, ier%)
   
     If ier% = -1 Then 'error detected, abort
        Exit Sub
      
     ElseIf ier% < -1 Then
        Call MsgBox("The Illegal character, ""&"", was found in search report #: " & Trim$(str$(-ier% - 1)) & "'s place name." _
                & vbCrLf & "Where ever the Illegal character was found it was replaced with a ""+""" _
                , vbInformation, "Export to Google Earth")
        End If
      
     
     End If
     
     Exit Sub
     


'--------------error handling----------------
errhand:
    
    Select Case Err.Number
       Case 53
          Screen.MousePointer = vbDefault
          Close
          Select Case MsgBox("File can't be overwritten!" & vbLf & _
                 "Pick a new file name.", vbCritical + vbOKCancel, "MapDigitizer")
             Case vbOK
                GoTo 10
             Case vbCancel
                Exit Sub
          End Select
       Case 3021
          'empty record
          LastRow0& = LastRow&
          Return
       Case 3315 'zero length string
          Resume Next
       Case Else
    End Select
    
errhand2:
    Select Case Err.Number
       Case 53
          Screen.MousePointer = vbDefault
          Close
          Select Case MsgBox("File can't be overwritten!" & vbLf & _
                 "Pick a new file name.", vbCritical + vbOKCancel, "MapDigitizer")
             Case vbOK
                GoTo 10
             Case vbCancel
                Exit Sub
          End Select
       Case 3021
          'empty record
          LastRow0& = LastRow&
          Return
       Case Else
          
    End Select

End Sub
'---------------------------------------------------------------------------------------
' Procedure : ExportReportToGoogleEarth
' DateTime  : 12/28/2008 01:49
' Author    : Chaim Keller
' Purpose   : Exports search results to Keyhole Markup Language File (Kml) file
'             that can be displayed by Google Earth, ArcExplorer, etc.
'
'              SomeSelected = false -- find out if selected results
'                       = true -- don't need to search (already searched)
'---------------------------------------------------------------------------------------
'
Sub ExportReportToGoogleEarth(KMLFileName$, SomeSelected As Boolean, ier%)

   On Error GoTo ExportReportToGoogleEarth_Error

   Dim numSelected&
   Dim Selected As Boolean
   
   ier% = 0 'initialize error flag
   
   If Not SomeSelected Then
     
     'check if some records are selected and count them
      numSelected& = 0
      For j& = 1 To numReport&
         If GDReportfrm.lvwReport.ListItems(j&).Selected = True Then
            numSelected& = numSelected& + 1
            If numSelected& > 1 Then
               GoTo 100
               End If
            End If
      Next j&
   
100
      Selected = False
      If numSelected& > 1 Then
            
            Select Case MsgBox("Some or all of the search results are selected." & vbCrLf & "Do you want to save just these records?" & vbCrLf & "" & vbCrLf & "Answer: Yes:  to save ONLY the selected records" & vbCrLf & "               No:   to save ALL the records""", vbYesNo + vbQuestion + vbDefaultButton1, App.Title)
            
               Case vbYes
                  Selected = True
               Case vbNo
                  Selected = False
            End Select
            
      End If
      
   Else
   
      Selected = SomeSelected
      
      End If
      
     
   Screen.MousePointer = vbHourglass
   
   If Trim$(KMLFileName$) = sEmpty Then 'generate default filename and write data to it
   
       KMLFileName$ = kmldir$ & "\" & "GSI_Search_" & Format(Now(), "hh-mm-ss~mm-dd-yyyy") & ".kml"
       
   'Else 'use passed name and path
   
       End If
   
   filkml% = FreeFile
   Open KMLFileName$ For Output As #filkml%
   
   'start writing kml (keyhole markup language) file to be openend by Google Earth
   
   Print #filkml%, "<?xml version=""1.0"" encoding=""UTF-8""?>"
   Print #filkml%, "<kml xmlns=""http://www.opengis.net/kml/2.2"">"
   Print #filkml%, "  <Document>"
   Print #filkml%, "    <name>MapDigitizer Search Results</name>" 'output" & Format(Now, "hh:mm:ss mm/dd/yyyy") & "</name>"
   Print #filkml%, "    <open>1</open>"
   Print #filkml%, "    <description>Created (Time-Date): " & Format(Now, "hh:mm:ss mm/dd/yyyy") & "</description>"
   Print #filkml%, "       <Style id=""OutCropIcon"">"
   Print #filkml%, "        <IconStyle>"
   Print #filkml%, "          <Icon>"
   Print #filkml%, "           <href>" & URL_OutCrop & "</href>"
   Print #filkml%, "       </Icon>"
   Print #filkml%, "        </IconStyle>"
   Print #filkml%, "     <LineStyle>"
   Print #filkml%, "          <width>2</width>"
   Print #filkml%, "     </LineStyle>"
   Print #filkml%, "        <BalloonStyle>"
   Print #filkml%, "          <text><![CDATA["
   Print #filkml%, "            <b>$[name]</b>"
   Print #filkml%, "            <br /><br />"
   Print #filkml%, "            $[description]"
   Print #filkml%, "            ]]></text>"
   Print #filkml%, "        </BalloonStyle>"
   Print #filkml%, "        </Style>"
   Print #filkml%, "       <Style id=""WellIcon"">"
   Print #filkml%, "        <IconStyle>"
   Print #filkml%, "          <Icon>"
   Print #filkml%, "           <href>" & URL_Well & "</href>"
   Print #filkml%, "          </Icon>"
   Print #filkml%, "        </IconStyle>"
   Print #filkml%, "     <LineStyle>"
   Print #filkml%, "          <width>2</width>"
   Print #filkml%, "     </LineStyle>"
   Print #filkml%, "        <BalloonStyle>"
   Print #filkml%, "          <text><![CDATA["
   Print #filkml%, "            <b>$[name]</b>"
   Print #filkml%, "            <br /><br />"
   Print #filkml%, "            $[description]"
   Print #filkml%, "            ]]></text>"
   Print #filkml%, "        </BalloonStyle>"
   Print #filkml%, "        </Style>"
   Print #filkml%, "        <Style id=""highlightPlacemark"">"
   Print #filkml%, "          <IconStyle>"
   Print #filkml%, "            <Icon>"
   Print #filkml%, "              <href>http://maps.google.com/mapfiles/kml/paddle/red-stars.png</href>"
   Print #filkml%, "            </Icon>"
   Print #filkml%, "          </IconStyle>"
   Print #filkml%, "          <BalloonStyle>"
   Print #filkml%, "            <text><![CDATA["
   Print #filkml%, "            <b>$[name]</b>"
   Print #filkml%, "            <br /><br />"
   Print #filkml%, "            $[description]"
   Print #filkml%, "            ]]></text>"
   Print #filkml%, "          </BalloonStyle>"
   Print #filkml%, "        </Style>"
   Print #filkml%, "        <Style id=""normalPlacemark"">"
   Print #filkml%, "          <IconStyle>"
   Print #filkml%, "            <Icon>"
   Print #filkml%, "              <href>http://maps.google.com/mapfiles/kml/paddle/wht-blank.png</href>"
   Print #filkml%, "            </Icon>"
   Print #filkml%, "          </IconStyle>"
   Print #filkml%, "          <BalloonStyle>"
   Print #filkml%, "            <text><![CDATA["
   Print #filkml%, "            <b>$[name]</b>"
   Print #filkml%, "            <br /><br />"
   Print #filkml%, "            $[description]"
   Print #filkml%, "            ]]></text>"
   Print #filkml%, "          </BalloonStyle>"
   Print #filkml%, "        </Style>"
   Print #filkml%, "        <StyleMap id=""exampleStyleMap"">"
   Print #filkml%, "          <Pair>"
   Print #filkml%, "            <key>normal</key>"
   Print #filkml%, "            <styleUrl>#normalPlacemark</styleUrl>"
   Print #filkml%, "          </Pair>"
   Print #filkml%, "          <Pair>"
   Print #filkml%, "            <key>highlight</key>"
   Print #filkml%, "            <styleUrl>#highlightPlacemark</styleUrl>"
   Print #filkml%, "          </Pair>"
   Print #filkml%, "        </StyleMap>"
   Print #filkml%, "    <Folder>"
   Print #filkml%, "      <name>Placemarks</name>"
   Print #filkml%, "      <description>Click on the individual placemarks for details of each search result</description>"
   
   'use WGS84 Geoid
   ChangeGeoid = False
   If Not GpsCorrection Then
      GpsCorrection = True
      ChangeGeoid = True 'set flag to change back the geoid after writing the kml file
      End If
      
   'look at average coordinates
   
   Call casgeo(AveEastSearch, AveNorthSearch, lgGoogle, ltGoogle)

   Print #filkml%, "      <LookAt>"
   Print #filkml%, "        <longitude>" & str$(-lgGoogle + 0.16246) & "</longitude>"
   Print #filkml%, "        <latitude>" & str$(ltGoogle - 1.34416) & "</latitude>"
   Print #filkml%, "        <altitude>0</altitude>"
   Print #filkml%, "        <heading>0</heading>"
   Print #filkml%, "        <tilt>38.8308932722697</tilt>"
   Print #filkml%, "        <range>299911.085945536</range>"
   Print #filkml%, "        <altitudeMode>relativeToGround</altitudeMode>"
   Print #filkml%, "      </LookAt>"

   For j& = 1 To numReport&
      
      If Selected And GDReportfrm.lvwReport.ListItems( _
          j&).Selected = False Then
         GoTo 1200 'this record is not selected so skip it
         End If
      
      docGoogle$ = sEmpty 'initialize table
      docGoogle$ = "<table cellspacing = ""0"" cellpading = ""0"" align = ""center"" >"
      
      placeName$ = sEmpty
      itmxplace = 0
      ITMyPlace = 0
      googlePlace = 0
      
      For i& = 1 To GDReportfrm.lvwReport.ColumnHeaders.count
      
          If i& = 1 Then
             doclin1 = GDReportfrm.lvwReport.ListItems(j&).Text
          Else
             doclin1 = GDReportfrm.lvwReport.ListItems(j&).SubItems(i& - 1)
             End If
             
          If IsNull(doclin1) Then
             doclin1 = sEmpty
          Else
             doclin1 = Trim$(doclin1)
             End If
             
          'enter search report information into table's fields
          Select Case GDReportfrm.lvwReport.ColumnHeaders(i&)
             Case "Type"
                placeName$ = doclin1
                
                'replace any illegal charcters
                If InStr(placeName$, "&") Then
                   placeName$ = Replace(placeName$, "&", "+")
                   ier% = -i& - 1
                   End If
                   
'                If InStr(doclin1, "(out") Then googlePlace = 0
'                If InStr(doclin1, "(well") Then googlePlace = 1
             Case "lon."
                longlace = CSng(doclin1)
                docGoogle$ = docGoogle$ & "<tr><td>longitutde</td><td>" & doclin1 & "</td></tr>"
                lgGoogle = longlace
             Case "lat"
                latplace = CSng(doclin1)
                docGoogle$ = docGoogle$ & "<tr><td>latitude</td><td>" & doclin1 & "</td></tr>"
                ltGoogle = latplace
'             Case "Order Number"
'                placeName$ = placeName$ & "; ON#: " & doclin1
'             Case Else 'other fields
'                If InStr(GDReportfrm.lvwReport.ColumnHeaders(i&), "Display Number") And Trim$(doclin1) <> sEmpty Then
'                   docGoogle$ = docGoogle$ & "<tr><td>Fossil Type</td><td>" & doclin1 & "</td></tr>"
'                ElseIf Trim$(doclin1) <> sEmpty And Trim$(doclin1) <> "." Then
'                   docGoogle$ = docGoogle$ & "<tr><td>" & GDReportfrm.lvwReport.ColumnHeaders(i&) & "</td><td>" & doclin1 & "</td></tr>"
'                   End If
          End Select
           
      Next i&
             
      docGoogle$ = docGoogle$ & "</table>" 'finish writing table of search report info
      
      'add viewing information for Google Earth
      
      If googlePlace = 0 Then 'outcropping
         Print #filkml%, "      <Placemark>"
         Print #filkml%, "      <name>" & placeName$ & "</name>"
         Print #filkml%, "      <visibility>1</visibility>"
         Print #filkml%, "      <description>" & docGoogle$ & "</description>"
         Print #filkml%, "      <LookAt>"
         Print #filkml%, "         <longitude>" & str$(-(lgGoogle - 0.0001)) & "</longitude>"
         Print #filkml%, "         <latitude>" & str$(ltGoogle - 0.0001) & "</latitude>"
         Print #filkml%, "         <altitude>0</altitude>"
         Print #filkml%, "         <heading>-148.413587348652</heading>"
         Print #filkml%, "         <tilt>40.558025245345</tilt>"
         Print #filkml%, "         <range>982.789307313033</range>"
         Print #filkml%, "         <altitudeMode>relativeToGround</altitudeMode>"
         Print #filkml%, "      </LookAt>"
         Print #filkml%, "      <styleUrl>#OutCropIcon</styleUrl>"
         Print #filkml%, "      <Point>"
         Print #filkml%, "         <extrude>1</extrude>"
         Print #filkml%, "         <altitudeMode>relativeToGround</altitudeMode>"
         Print #filkml%, "         <coordinates>" & str$(-lgGoogle) & "," & str$(ltGoogle) & "</coordinates>"
         Print #filkml%, "      </Point>"
         Print #filkml%, "      </Placemark>"
      ElseIf googlePlace = 1 Then 'well
         Print #filkml%, "      <Placemark>"
         Print #filkml%, "      <name>" & placeName$ & "</name>"
         Print #filkml%, "      <visibility>1</visibility>"
         Print #filkml%, "      <description>" & docGoogle$ & "</description>"
         Print #filkml%, "      <LookAt>"
         Print #filkml%, "         <longitude>" & str$(-(lgGoogle - 0.0027)) & "</longitude>"
         Print #filkml%, "         <latitude>" & str$(ltGoogle - 0.0001) & "</latitude>"
         Print #filkml%, "         <altitude>0</altitude>"
         Print #filkml%, "         <heading>-148.413587348652</heading>"
         Print #filkml%, "         <tilt>40.558025245345</tilt>"
         Print #filkml%, "         <range>982.789307313033</range>"
         Print #filkml%, "         <altitudeMode>relativeToGround</altitudeMode>"
         Print #filkml%, "      </LookAt>"
         Print #filkml%, "      <styleUrl>#WellIcon</styleUrl>"
         Print #filkml%, "      <Point>"
         Print #filkml%, "         <extrude>1</extrude>"
         Print #filkml%, "         <altitudeMode>relativeToGround</altitudeMode>"
         Print #filkml%, "         <coordinates>" & str$(-lgGoogle) & "," & str$(ltGoogle) & "</coordinates>"
         Print #filkml%, "      </Point>"
         Print #filkml%, "      </Placemark>"
         End If
        
1200 Next j&

   'write kml closing
   
   Print #filkml%, "    </Folder>"
   Print #filkml%, "  </Document>"
   Print #filkml%, "</kml>"
   
   Close #filkml%
   
   If ChangeGeoid Then
      GpsCorrection = False 'reset the geoid
      ChangeGeoid = False
      End If
   
   Screen.MousePointer = vbDefault
   
   'now start up Google Earth

   On Error GoTo 0
   
   Exit Sub

ExportReportToGoogleEarth_Error:

   Screen.MousePointer = vbDefault
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExportReportToGoogleEarth of Module modGDModule", vbCritical, "MapDigitizer Export to Google Earth Error"
   ier% = -1
End Sub
Sub PreviewPrintDetails()
'
'      'Display details of search report in a print ready form.
'      'Also writes temporary print file that can be used for
'      'saving the print preview to an output file.
'
'      'Dear Porgrammer, you change these values according to your
'      'document.  The following tools are available:
'      'All coordinates are in inches!
'
'      '(1) PrintFilledBox x,y,width,height,color
'      '               prints filled box whose upper left corner is
'      '               at x,y inches, with width,height in inches
'      '               color = RGB(rednumber,greennumber,bluenumber)
'      '                       or QBColor(number)
'      '(2) PrintBox 'same as above but prints unfilled box
'      '(3) PrintLine x1,y1,x2,y2,color
'      '(4) PrintFontName 'Sets the font for the following line
'      '(5) PrintFontBold 'What follows is in bold face
'      '(6) PrintFontRegular 'What follows is not in bold face
'      '(5) PrintCurrentX(Y) 'prints at this coordinate
'      '(6) PrintFontSize 'Sets the font size
'      '(7) PrintPrint 'The actual line to be printed
'      '(8) PrintCircle x,y,radius,color
'
'   On Error GoTo errhand
'
'   Screen.MousePointer = vbHourglass
'
'   Dim Xo As Single, Yo As Single
'
'   Dim numOrder(125) As Variant
'
'   Dim AnalystName As String
'   Dim AnalysisDate As String
'   Dim AnalystNames(NUM_FOSSIL_TYPES) As Variant
'   Dim AnalysisDates(NUM_FOSSIL_TYPES) As Variant
'   Dim FosIDCono As Long
'   Dim FosIDDiatom As Long
'   Dim FosIDForam As Long
'   Dim FosIDMega As Long
'   Dim FosIDNano As Long
'   Dim FosIDOstra As Long
'   Dim FosIDPaly As Long
'   Dim sAnum&
'   Dim FossilTbl$
'
'   'open temprorary print file
'   filprnt% = FreeFile
'   Open direct$ & "\print_tmp.txt" For Output As #filprnt%
'
''------------------------------------------------------
'
'      Xo = 0.25 'topleft origins
'      Yo = 0.85
'      Write #filprnt%, Xo, Yo
'      PrintFontName "Arial"
'      PrintFontSize 12
'
'      PrintCurrentX Xo + 1.1  '2.3
'      PrintCurrentY Yo - 0.09
'      PrintFontBold
'
'      'query the database
'
'      Call QueryToPrintSave(numOrder(), DocName$, AnalystName, AnalysisDate, AnalystNames(), AnalysisDates(), sAnum&)
'
'      FosIDCono = numOrder(18)
'      FosIDDiatom = numOrder(17)
'      FosIDForam = numOrder(12)
'      FosIDMega = numOrder(15)
'      FosIDNano = numOrder(16)
'      FosIDOstra = numOrder(13)
'      FosIDPaly = numOrder(14)
'
'      'header
'
'      If PicSum Then
'         PrintPrint "MapDigitizer SEARCH RESULT # " & Trim$(str$(NewHighlighted&)) & " ; DATE/TIME: " & Now
'         Write #filprnt%, 1, 1, "MapDigitizer SEARCH RESULT # " & Trim$(str$(NewHighlighted&)) & " ; DATE/TIME: " & Now
'         If OrderNum& > 0 Then DocName$ = Trim$(str$(OrderNum&))
'      Else
'         If OrderNum& < 0 Then 'record from scanned database
'             PrintPrint "MapDigitizer Scanned DBase No. " & DocName$ & " ; DATE/TIME: " & Now
'             Write #filprnt%, 1, 1, "MapDigitizer Scanned DBase No. " & DocName$ & " ; DATE/TIME: " & Now
'         Else
'             DocName$ = Trim$(str$(OrderNum&))
'             PrintPrint "MapDigitizer Order No. # " & DocName$ & " ; DATE/TIME: " & Now
'             Write #filprnt%, 1, 1, "MapDigitizer Order No. # " & DocName$ & " ; DATE/TIME: " & Now
'             End If
'
'         End If
'
'      PrintFontRegular
'
'      PrintFontName "Arial"
'      PrintFontSize 10
'
'
'      If OrderNum& > 0 And Not SearchDB And SearchVis Then
'         'gdsearchfrm only popped up due to querying it's
'         'genus and species lists, but otherwise we don't want it
'         'so make it invisible
'         GDSearchfrm.Visible = False
'         End If
'
'      '**********************END QUERYING DATABASES*****************
'
'pp700:
'         'place the map at this point if not already there
'         If GeoMap Or TopoMap Then
'            GDMDIform.Text5 = numOrder(1)
'            GDMDIform.Text6 = numOrder(2)
'            ce& = 1
'            Call gotocoord 'move the map to the record's coordinates
'            End If
'
'
'      '**********************DISPLAY RESULTS***********************
'      '________GENERAL SAMPLE INFO-FIRST LEFT COLUMN___________
'
'      Yo = Yo + 0.15
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18
'      PrintPrint "CLIENT: " & numOrder(23)
'      Write #filprnt%, 1, 2, "CLIENT: " & numOrder(23)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 2
'      PrintPrint "COMPANY/DIVISION: " & numOrder(25)
'      Write #filprnt%, 1, 3, "COMPANY/DIVISION: " & numOrder(25)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 3
'      PrintPrint "PROJECT: " & numOrder(24)
'      Write #filprnt%, 1, 4, "PROJECT: " & numOrder(24)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 4
'      PrintPrint "FORMATION: " & numOrder(3)
'      Write #filprnt%, 1, 5, "FORMATION: " & numOrder(3)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 5
'
'      If sAnum& = 1 Then 'well
'         PrintPrint "SAMPLE METHOD: WELL"
'         Write #filprnt%, 1, 6, "SAMPLE METHOD: WELL"
'         PrintCurrentX Xo + 0.5
'         PrintCurrentY Yo + 0.18 * 6
'         PrintPrint "WELL NAME: " & numOrder(0)
'         Write #filprnt%, 1, 7, "WELL NAME: " & numOrder(0)
'      ElseIf sAnum& = 0 Then 'surface
'         PrintPrint "SAMPLE METHOD: SURFACE"
'         Write #filprnt%, 1, 6, "SAMPLE METHOD: SURFACE"
'         PrintCurrentX Xo + 0.5
'         PrintCurrentY Yo + 0.18 * 6
'         PrintPrint "PLACE NAME: " & numOrder(0)
'         Write #filprnt%, 1, 7, "PLACE NAME: " & numOrder(0)
'      ElseIf sAnum& = -1 Then 'unknown type
'         PrintPrint "SAMPLE METHOD: UNKNOWN"
'         Write #filprnt%, 1, 6, "SAMPLE METHOD: UNKNOWN"
'         PrintCurrentX Xo + 0.5
'         PrintCurrentY Yo + 0.18 * 6
'         PrintPrint "PLACE NAME: " & numOrder(0)
'         Write #filprnt%, 1, 7, "PLACE NAME: " & numOrder(0)
'         End If
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 7
'      PrintPrint "ITMx: " & numOrder(1)
'      Write #filprnt%, 1, 8, "ITMx: " & numOrder(1)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 8
'      PrintPrint "ITMy: " & numOrder(2)
'      Write #filprnt%, 1, 9, "ITMy: " & numOrder(2)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 9
'      If numOrder(11) = sEmpty Then
'         PrintPrint "REMARKS: " & numOrder(11)
'         Write #filprnt%, 1, 10, "REMARKS: " & numOrder(11)
'      Else
'         numOrder(11) = Trim$(numOrder(11))
'         lenRemark& = Len(numOrder(11))
'
'        '*******************************
'         'Detect Hebrew characters and long lines.  Note,
'         'Remarks is sometimes too large to fit in one page,
'         'or sometimes it has line feed and other control characters.
'         'So parse out the control characters and break up the line.
'         'Note: Only a maximum of two lines of text is supported.  Any
'         'more text will not be in the margins.
'         '********************************
'
'          fndHebrew& = 0
'          For i& = 1 To lenRemark&
'            If Asc(Mid$(numOrder(11), i&, 1)) <= 20 Then
'               Mid$(numOrder(11), i&, 1) = " "
'            ElseIf fndHebrew& = 0 And Asc(Mid$(numOrder(11), i&, 1)) >= 128 Then
'              'Hebrew characters found so switch ObjPrint to
'              'Hebrew compatible font.
'              PrintFontName "Arial (Hebrew)"
'              fndHebrew& = 1
'              End If
'          Next i&
'
'          'new length
'          lenRemark& = Len(numOrder(11)) 'length
'
'          MaxNumCh& = 85 'maximum number of characters on one line
'
'          If lenRemark& > MaxNumCh& Then 'line is too long (more than 80 characters)
'            'find first space after MaxNumCh& characters and break it there
'            pos1& = InStr(MaxNumCh& + 1, numOrder(11), " ")
'            If pos1& <> 0 Then
'               PrintPrint "REMARKS: " & Mid$(numOrder(11), 1, pos1& - 1)
'               Write #filprnt%, 1, 10, "REMARKS: " & Mid$(numOrder(11), 1, pos1& - 1)
'               PrintCurrentX Xo + 0.5
'               PrintPrint "              " & Mid$(numOrder(11), pos1& + 1, lenRemark& - pos1& - 1)
'               Write #filprnt%, 1, 11, Mid$(numOrder(11), pos1& + 1, lenRemark& - pos1& - 1)
'            Else 'no space found so just break it up in middle of word (add "-" to indicate this break)
'               PrintPrint "REMARKS: " & Mid$(numOrder(11), 1, MaxNumCh&) & "-"
'               Write #filprnt%, 1, 10, "REMARKS: " & Mid$(numOrder(11), 1, MaxNumCh&) & "-"
'               PrintCurrentX Xo + 0.5
'               PrintPrint "              " & Mid$(numOrder(11), MaxNumCh& + 1, lenRemark& - MaxNumCh&)
'               Write #filprnt%, 1, 11, Mid$(numOrder(11), MaxNumCh& + 1, lenRemark& - MaxNumCh&)
'               End If
'          Else 'normal length
'            PrintPrint "REMARKS: " & numOrder(11)
'            Write #filprnt%, 1, 10, "REMARKS: " & numOrder(11)
'            End If
'
'         'change back to original font
'          PrintFontName "Arial"
'
'       End If
'
'      '________GENERAL SAMPLE INFO-FIRST RIGHT COLUMN___________
'
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18
'      PrintPrint "LIM UP:         " & numOrder(5)
'      Write #filprnt%, 2, 2, "LIM UP:         " & numOrder(5)
'
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18 * 2
'      PrintPrint "LIM DOWN:    " & numOrder(4)
'      Write #filprnt%, 2, 3, "LIM DOWN:    " & numOrder(4)
'
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18 * 3
'
'      If val(numOrder(7)) = 1 And sAnum& = 1 Then
'         PrintPrint "WELL SAMPLE TYPE: CUTTING"
'         Write #filprnt%, 2, 4, "WELL SAMPLE TYPE: CUTTING"
'      ElseIf val(numOrder(7)) = 2 And sAnum& = 1 Then
'         PrintPrint "WELL SAMPLE TYPE: CORE"
'         Write #filprnt%, 2, 4, "WELL SAMPLE TYPE: CORE"
'      Else
'         PrintPrint "WELL SAMPLE TYPE: "
'         Write #filprnt%, 2, 4, "WELL SAMPLE TYPE: "
'         End If
'
'ppd800:
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18 * 4
'      PrintPrint "CORE NUMBER: " & numOrder(26)
'      Write #filprnt%, 2, 5, "CORE NUMBER: " & numOrder(26)
'
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18 * 5
'      PrintPrint "BOX NUMBER: " & numOrder(27)
'      Write #filprnt%, 2, 6, "BOX NUMBER: " & numOrder(27)
'
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18 * 6
'      PrintPrint "FIELD NUMBER: " & numOrder(9)
'      Write #filprnt%, 2, 7, "FIELD NUMBER: " & numOrder(9)
'
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18 * 7
'      PrintPrint "ORDER NUMBER: " & DocName$
'      Write #filprnt%, 2, 8, "ORDER NUMBER: " & DocName$
'
'      PrintCurrentX Xo + 4.68
'      PrintCurrentY Yo + 0.18 * 8
'      PrintPrint "DATE: " & numOrder(21)
'      Write #filprnt%, 2, 9, "DATE: " & numOrder(21)
'
'      If InStr(PrintPreview.cmbPages.Text, "Page 1") Then
'         'summary page
'         GoTo ppd750
'      ElseIf InStr(PrintPreview.cmbPages.Text, "Page 1") = 0 Then
'         'fossil names
'         GoTo ppd850
'         End If
'
'ppd750: 'Previewing/printing summary page, so
'      'step through each fossil and present results if they exist.
'      Dim Ylast As Single
'      Ylast = Yo + 0.18 * 3
'
'      '-----------print fossil and age information for scanned database
'      If OrderNum& < 0 Then
'         PrintCurrentX Xo + 0.5
'         Ylast = Ylast + 0.18 * 8
'         PrintCurrentY Ylast
'         PrintFontName "Arial"
'         PrintFontSize 12
'         PrintPrint "FOSSIL CATEGORY: " & FossilTbl$
'         Write #filprnt%, 1, CInt((Ylast - Yo) / 0.18 + 1), "FOSSIL CATEGORY: " & FossilTbl$
'         PrintFontName "Arial"
'         PrintFontSize 10
'         PrintCurrentX Xo + 0.5
'         PrintPrint "DATE: " & numOrder(21)
'         Write #filprnt%, 1, CInt((Ylast - Yo) / 0.18 + 2), "DATE: " & numOrder(21)
'         PrintCurrentX Xo + 0.5
'         PrintPrint "Earlier AGE: " & numOrder(28) & " " & numOrder(29)
'         Write #filprnt%, 1, CInt((Ylast - Yo) / 0.18 + 3), "Earlier AGE: " & numOrder(28) & " " & numOrder(29)
'         PrintCurrentX Xo + 0.5
'         PrintPrint "LATER AGE: " & numOrder(30) & " " & numOrder(31)
'         Write #filprnt%, 1, CInt((Ylast - Yo) / 0.18 + 4), "LATER AGE: " & numOrder(30) & " " & numOrder(31)
'         PrintCurrentX Xo + 0.5
'         PrintPrint "Earlier ZONE: "
'         Write #filprnt%, 1, CInt((Ylast - Yo) / 0.18 + 5), "Earlier ZONE: "
'         PrintCurrentX Xo + 0.5
'         PrintPrint "LATER ZONE: "
'         Write #filprnt%, 1, CInt((Ylast - Yo) / 0.18 + 6), "LATER ZONE: "
'         PrintCurrentX Xo + 0.5
'         PrintPrint "REMARK: "
'         Write #filprnt%, 1, CInt((Ylast - Yo) / 0.18 + 7), "REMARK: "
'
'         Screen.MousePointer = vbDefault
'         Close #filprnt%
'         Exit Sub
'         End If
'      '-------------------------------------------------
'
'      'first conod
'      strnum& = 29
'      If val(numOrder(18)) <> 0 Then
'         FossilTbl$ = "CONODONTA"
'         FosNum& = 18
'         Call PrintFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, filprnt%)
'         End If
'
'      'then diato
'      strnum& = strnum& + 12
'      If val(numOrder(17)) <> 0 Then
'         FossilTbl$ = "DIATOM"
'         FosNum& = 17
'         Call PrintFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, filprnt%)
'         End If
'
'      'then foram
'      strnum& = strnum& + 12
'      If val(numOrder(12)) <> 0 Then
'         FossilTbl$ = "FORAMINIFERA"
'         FosNum& = 12
'         Call PrintFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, filprnt%)
'         End If
'
'      'then megaf
'      strnum& = strnum& + 12
'      If val(numOrder(15)) <> 0 Then
'         FossilTbl$ = "MEGAFAUNA"
'         FosNum& = 15
'         Call PrintFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, filprnt%)
'         End If
'
'      'then nanno
'      strnum& = strnum& + 12
'      If val(numOrder(16)) <> 0 Then
'         FossilTbl$ = "NANNOPLANKTON"
'         FosNum& = 16
'         Call PrintFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, filprnt%)
'         End If
'
'      'then ostra
'      strnum& = strnum& + 12
'      If val(numOrder(13)) <> 0 Then
'         FossilTbl$ = "OSTRACODA"
'         FosNum& = 13
'         Call PrintFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, filprnt%)
'         End If
'
'      'then palin
'      strnum& = strnum& + 12
'      If val(numOrder(14)) <> 0 Then
'         FossilTbl$ = "PALYNOLOGY"
'         FosNum& = 14
'         Call PrintFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, filprnt%)
'         End If
'
'ppd850:
'      If InStr(PrintPreview.cmbPages.Text, "Page 1") <> 0 Then
'         'previewing/printing summary page, so skip the following lines
'         Close #filprnt% 'close temporary print file
'         Screen.MousePointer = vbDefault
'         Exit Sub
'         End If
'
'      'previewing fossil name pages
'      PrintLine Xo + 0.5, Yo + 0.18 * 12.5, Xo + 7.2, Yo + 0.18 * 12.5, QBColor(0)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 13
'      PrintFontBold
'      PrintPrint "CHECK METHOD: " & UCase$(Mid$(PrintPreview.cmbPages.Text, 9, Len(PrintPreview.cmbPages.Text) - 8))
'      Write #filprnt%, 1, 14, "CHECK METHOD: " & UCase$(Mid$(PrintPreview.cmbPages.Text, 9, Len(PrintPreview.cmbPages.Text) - 8))
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 14
'      PrintFontRegular
'      PrintPrint "ANALYST: " & AnalystName
'      Write #filprnt%, 1, 15, "ANALYST: " & AnalystName
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 15
'      PrintPrint "ANALYSIS DATE: " & AnalysisDate
'      Write #filprnt%, 1, 16, "ANALYSIS DATE: " & AnalysisDate
'
'      PrintLine Xo + 0.5, Yo + 0.18 * 16.5, Xo + 7.2, Yo + 0.18 * 16.5, QBColor(0)
'
'      PrintCurrentX Xo + 0.5
'      PrintCurrentY Yo + 0.18 * 18
'      PrintFontBold
'      PrintPrint "FOSSIL NAMES"
'      Write #filprnt%, 1, 19, "FOSSIL NAMES"
'
'      PrintCurrentX Xo + 3.5
'      PrintCurrentY Yo + 0.18 * 18
'      PrintFontBold
'      PrintPrint "SEMI QUANT"
'      Write #filprnt%, 2, 19, "SEMI QUANT"
'
'      PrintCurrentX Xo + 5
'      PrintCurrentY Yo + 0.18 * 18
'      PrintFontBold
'      PrintPrint "FEATURES"
'      Write #filprnt%, 3, 19, "FEATURES"
'
'      PrintCurrentX Xo + 6.5
'      PrintCurrentY Yo + 0.18 * 18
'      PrintFontBold
'      PrintPrint "QUANTITY"
'      Write #filprnt%, 4, 19, "QUANTITY"
'
'      'now query database for the fossil names,
'      'semi quant, features, quantity
'
'      Dim FossilTag$, fosstbl$, FosTbl$, FosDic$
'      Dim Xfos As Single, Yfos As Single
'
'      Xfos = Xo + 0.5
'      Yfos = Yo + 0.18 * 19.5
'      PrintFontRegular
'
'      Select Case Mid$(PrintPreview.cmbPages.Text, 9, Len(PrintPreview.cmbPages.Text) - 8)
'        Case "Conodonta"
'            fosstbl$ = "condores"
'            FosTbl$ = "condofos"
'            FosDic$ = "Conodsdic"
'            'query for foram fossil names
'            Call PrintFosNames(fosstbl$, FosTbl$, FosDic$, FosIDCono, Xfos, Yfos, filprnt%)
'        Case "Diatom"
'            fosstbl$ = "diatores"
'            FosTbl$ = "diatofos"
'            FosDic$ = "Diatomsdic"
'            'query for foram fossil names
'            Call PrintFosNames(fosstbl$, FosTbl$, FosDic$, FosIDDiatom, Xfos, Yfos, filprnt%)
'        Case "Foraminifera"
'            fosstbl$ = "foramres"
'            FosTbl$ = "foramfos"
'            FosDic$ = "Foramsdic"
'            'query for foram fossil names
'            Call PrintFosNames(fosstbl$, FosTbl$, FosDic$, FosIDForam, Xfos, Yfos, filprnt%)
'        Case "Megafauna"
'            fosstbl$ = "megares"
'            FosTbl$ = "megafos"
'            FosDic$ = "Megadic"
'            'query for foram fossil names
'            Call PrintFosNames(fosstbl$, FosTbl$, FosDic$, FosIDMega, Xfos, Yfos, filprnt%)
'        Case "Nannoplankton"
'            fosstbl$ = "nanores"
'            FosTbl$ = "nanofos"
'            FosDic$ = "Nanodic"
'            'query for foram fossil names
'            Call PrintFosNames(fosstbl$, FosTbl$, FosDic$, FosIDNano, Xfos, Yfos, filprnt%)
'        Case "Ostracoda"
'            fosstbl$ = "ostrares"
'            FosTbl$ = "ostrafos"
'            FosDic$ = "Ostracoddic"
'            'query for foram fossil names
'            Call PrintFosNames(fosstbl$, FosTbl$, FosDic$, FosIDOstra, Xfos, Yfos, filprnt%)
'        Case "Palynology"
'            fosstbl$ = "palynres"
'            FosTbl$ = "palynfos"
'            FosDic$ = "Palyndic"
'            'query for foram fossil names
'            Call PrintFosNames(fosstbl$, FosTbl$, FosDic$, FosIDPaly, Xfos, Yfos, filprnt%)
'        Case Else
'      End Select
'
'      Close #filprnt% 'close temporary print file
'
'      Screen.MousePointer = vbDefault
'
'   Exit Sub
'
'errhand:
'
'    Screen.MousePointer = vbDefault
'
'    Select Case Err.Number
'       Case 94, 3021
'          'null field or empty record
'          Resume Next
'       Case Else
'          If filprnt% <> 0 Then Close #filprnt%
'          filprnt% = 0
'          Screen.MousePointer = vbDefault
'          response = MsgBox("Encountered error #: " & Err.Number & vbLf & Err.Description & vbLf & "The rest of the print preview can probably be salvaged." & vbLf & "Attempt resuming the preview?" & vbLf & sEmpty, vbExclamation + vbYesNoCancel, "MapDigitizer")
'          If response = vbYes Then
'             Resume Next
'             End If
'    End Select
'
End Sub

'Sub QueryFossil(Fossil$, ByRef numOrder(), StartNum&, FossilNum&)
'
'    'querys fossil and analyst tables for displaying and printing
'         On Error Resume Next
'
'         Dim sqFossil As String
'         Dim rstFossil As Recordset
'         Dim numTmp$, FossCat As String, zone$
'         Dim sqfoss As String
'         Dim rstfoss As Recordset
'         Dim Name2 As String
'         Dim Name1 As String
'
'         sqFossil = "SELECT * FROM " & Fossil$ & " WHERE " & Fossil$ & "![res_id] = " & numOrder(FossilNum&)
'         Set rstFossil = gdbs.OpenRecordset(sqFossil & ";", dbOpenSnapshot)
'         With rstFossil
'            .MoveFirst
'            numOrder(StartNum&) = rstFossil![analist] 'Analyst
'            numOrder(StartNum& + 1) = rstFossil![Date]  'date of entry
'            numOrder(StartNum& + 2) = rstFossil![preageo] 'earlier age prefix
'            'convert this number to words
'            numTmp$ = numOrder(StartNum& + 2)
'            GoSub Prefixes
'            numOrder(StartNum& + 2) = numTmp$
'            numOrder(StartNum& + 3) = rstFossil![preagey] 'later age prefix
'            'convert this number to words
'            numTmp$ = numOrder(StartNum& + 3)
'            GoSub Prefixes
'            numOrder(StartNum& + 3) = numTmp$
'            numOrder(StartNum& + 4) = rstFossil![ch_o_age] 'earlier age
'            numOrder(StartNum& + 5) = rstFossil![ch_y_age] 'later age
'            numOrder(StartNum& + 6) = rstFossil![prezoneo] 'earlier zone prefix
'            'convert this number to words
'            numTmp$ = numOrder(StartNum& + 6)
'            GoSub Prefixes
'            numOrder(StartNum& + 6) = numTmp$
'            numOrder(StartNum& + 7) = rstFossil![prezoney] 'later zone prefix
'            'convert this number to words
'            numTmp$ = numOrder(StartNum& + 7)
'            GoSub Prefixes
'            numOrder(StartNum& + 7) = numTmp$
'
'            numOrder(StartNum& + 8) = rstFossil![zoneo] 'earlier zone
'            If val(numOrder(StartNum& + 8)) = 0 Then
'               numOrder(StartNum& + 8) = sEmpty
'            Else 'query for the genus and species
'               num& = numOrder(StartNum& + 8)
'               GoSub QueryZone
'               numOrder(StartNum& + 8) = zone$
'               End If
'
'            numOrder(StartNum& + 9) = rstFossil![zoney] 'later zone
'            If val(numOrder(StartNum& + 9)) = 0 Then
'               numOrder(StartNum& + 9) = sEmpty
'            Else 'query for the genus and species
'               num& = numOrder(StartNum& + 9)
'               GoSub QueryZone
'               numOrder(StartNum& + 9) = zone$
'               End If
'
'            numOrder(StartNum& + 10) = rstFossil![remark] 'remark
'            If val(numOrder(StartNum& + 10)) = 0 Then
'               numOrder(StartNum& + 10) = sEmpty
'            Else
'                Dim sqRemark As String
'                Dim rstRemark As Recordset
'                sqRemark = "SELECT * FROM Remark WHERE Remark![id] = " & numOrder(StartNum& + 10)
'                Set rstRemark = gdbs.OpenRecordset(sqRemark & ";", dbOpenSnapshot)
'                With rstRemark
'                   .MoveFirst
'                   If IsNull(rstRemark![Line]) Then
'                      numOrder(StartNum& + 10) = sEmpty
'                   Else
'                      numOrder(StartNum& + 10) = rstRemark![Line] 'text of remarks
'                      End If
'                End With
'                rstRemark.Close
'                End If
'
'         End With
'         rstFossil.Close
'
'         'now query analyst
'         If val(numOrder(StartNum&)) <> 0 Then
'            Dim sqAnalyst As String
'            Dim rstAnalyst As Recordset
'            sqAnalyst = "SELECT * FROM CHKPALdic WHERE CHKPALdic![id] = " & numOrder(StartNum&)
'            Set rstAnalyst = gdbs.OpenRecordset(sqAnalyst & ";", dbOpenSnapshot)
'            With rstAnalyst
'               .MoveFirst
'               numOrder(StartNum& + 11) = rstAnalyst![Name] 'Analyst's name
'            End With
'            rstAnalyst.Close
'            End If
'
'         Exit Sub
'
'Prefixes:
'    If IsNull(numTmp$) Then
'    Else
'       Select Case val(numTmp$)
'          Case 1
'            numTmp$ = "Lower"
'          Case 2, 0
'            numTmp$ = sEmpty
'          Case 3
'            numTmp$ = "Upper"
'          Case 4
'            numTmp$ = "Middle"
'       End Select
'       End If
'Return
'
'
'QueryZone: 'inline gosub that queries the Fossil$ table for
'           'the Paleo zone name
'    If num& = 0 Then
'       zone$ = sEmpty
'       Return
'       End If
'
'    Select Case Fossil$
'       Case "condores"
'          FossCat = "CONOZONEcat"
'       Case "diatores"
'          FossCat = "DIATOZONEcat"
'       Case "foramres"
'          FossCat = "FORAZONEcat"
'       Case "megares"
'          FossCat = "MEGAZONEcat"
'       Case "nanores"
'          FossCat = "NANOZONEcat"
'       Case "ostrares"
'          FossCat = "OSTRAZONcat"
'       Case "palynres"
'          FossCat = "PALIZONEcat"
'    End Select
'
'    'query fossil (zone) dictionary
'
'    sqfoss = "SELECT " & FossCat & ".* FROM " & FossCat & _
'             " WHERE " & FossCat & ".id = " & str$(num&)
'    Set rstfoss = gdbs.OpenRecordset(sqfoss, dbOpenSnapshot)
'
'    With rstfoss
'       .MoveFirst
'       'Genera
'       If Not IsNull(rstfoss![name_2]) Then
'          Name2 = rstfoss![name_2]
'       Else
'          Name2 = sEmpty
'          End If
'
'       'Species
'       If Not IsNull(rstfoss![name_1]) Then
'          Name1 = rstfoss![name_1]
'       Else
'          Name1 = sEmpty
'          End If
'
'       zone$ = LTrim$(Trim$(Name2) & " " & Trim$(Name1)) 'Genera and species
'
'    End With
'    rstfoss.Close
'
'    Return
'
'
'End Sub
'
'Sub PrintFossilInfo(FossilTbl$, numOrder() As Variant, StartNum&, FossilNum&, Xo As Single, Ylast As Single, filprnt%)
'
'    On Error GoTo errhand
'
'    PrintCurrentX Xo + 0.5
'    Ylast = Ylast + 0.18 * 9
'    PrintCurrentY Ylast
'    PrintFontName "Arial"
'    PrintFontSize 12
'    PrintPrint FossilTbl$ & ": DISPLAY NUMBER: " & numOrder(FossilNum&) & "; ANALYST: " & numOrder(StartNum& + 11)
'    Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 3, FossilTbl$ & ": DISPLAY NUMBER: " & numOrder(FossilNum&) & "; ANALYST: " & numOrder(StartNum& + 11)
'    PrintFontName "Arial"
'    PrintFontSize 10
'    PrintCurrentX Xo + 0.5
'    PrintPrint "DATE: " & numOrder(StartNum& + 1)
'    Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 4, "DATE: " & numOrder(StartNum& + 1)
'    PrintCurrentX Xo + 0.5
'    PrintPrint "Earlier AGE: " & numOrder(StartNum& + 2) & " " & numOrder(StartNum& + 4)
'    Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 5, "Earlier AGE: " & numOrder(StartNum& + 2) & " " & numOrder(StartNum& + 4)
'    PrintCurrentX Xo + 0.5
'    PrintPrint "LATER AGE: " & numOrder(StartNum& + 3) & " " & numOrder(StartNum& + 5)
'    Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 6, "LATER AGE: " & numOrder(StartNum& + 3) & " " & numOrder(StartNum& + 5)
'    PrintCurrentX Xo + 0.5
'    PrintPrint "Earlier ZONE: " & numOrder(StartNum& + 6) & " " & numOrder(StartNum& + 8)
'    Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 7, "Earlier ZONE: " & numOrder(StartNum& + 6) & " " & numOrder(StartNum& + 8)
'    PrintCurrentX Xo + 0.5
'    PrintPrint "LATER ZONE: " & numOrder(StartNum& + 7) & " " & numOrder(StartNum& + 9)
'    Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 8, "LATER ZONE: " & numOrder(StartNum& + 7) & " " & numOrder(StartNum& + 9)
'    PrintCurrentX Xo + 0.5
'
'   '*******************************
'    'Fossil Remark sometimes is too large to fit in one page,
'    'or sometimes it has line feed and other control characters.
'    'So parse out the control characters and break up the line.
'    'Note: Only a maximum of two lines of text is supported.  Any
'    'more text will not be in the margins.
'    '********************************
'
'    numOrder(StartNum& + 10) = Trim$(numOrder(StartNum& + 10))
'    lenRemark& = Len(numOrder(StartNum& + 10)) 'length
'
'    'Remove control characters from string.
'    'Also check for Hebrew characters.
'    fndHebrew& = 0
'    For i& = 1 To lenRemark&
'       If Asc(Mid$(numOrder(StartNum& + 10), i&, 1)) <= 20 Then
'          Mid$(numOrder(StartNum& + 10), i&, 1) = " "
'       ElseIf fndHebrew& = 0 And Asc(Mid$(numOrder(StartNum& + 10), i&, 1)) >= 128 Then
'          'Hebrew characters found so switch ObjPrint to
'          'Hebrew compatible font.
'          PrintFontName "Arial (Hebrew)"
'          fndHebrew& = 1
'          End If
'    Next i&
'
'    'new length
'    lenRemark& = Len(numOrder(StartNum& + 10)) 'length
'
'    MaxNumCh& = 65 'maximum number of characters on one line
'
'    If lenRemark& > MaxNumCh& Then 'line is too long (more than 80 characters)
'       'find first space after MaxNumCh& characters and break it there
'       pos1& = InStr(MaxNumCh& + 1, numOrder(StartNum& + 10), " ")
'       If pos1& <> 0 Then
'          PrintPrint "REMARK: " & Mid$(numOrder(StartNum& + 10), 1, pos1& - 1)
'          'check for the need of a third line
'          If Len(Mid$(numOrder(StartNum& + 10), pos1& + 1, lenRemark& - pos1&)) > MaxNumCh& Then
'             'break it up into 3rd line
'              pos2& = InStr(MaxNumCh&, Mid$(numOrder(StartNum& + 10), pos1& + 1, lenRemark& - pos1&), " ")
'              If pos2& <> 0 Then
'                 PrintCurrentX Xo + 0.5
'                 PrintPrint "              " & Mid$(Mid$(numOrder(StartNum& + 10), pos1& + 1, lenRemark& - pos1&), 1, pos2&)
'                 PrintCurrentX Xo + 0.5
'                 PrintPrint "              " & Mid$(Mid$(numOrder(StartNum& + 10), pos1& + 1, lenRemark& - pos1&), pos2& + 1, Len(Mid$(numOrder(StartNum& + 10), 1, pos1& - 1)) - 1)
'                 Ylast = Ylast + 0.18 'notify program to reserve space for the third line
'              Else 'just print the remainder on the second line
'                 PrintCurrentX Xo + 0.5
'                 PrintPrint "              " & Mid$(numOrder(StartNum& + 10), pos1& + 1, lenRemark& - pos1&)
'                 End If
'          Else 'just print two lines
'             PrintCurrentX Xo + 0.5
'             PrintPrint "              " & Mid$(numOrder(StartNum& + 10), pos1& + 1, lenRemark& - pos1&)
'             End If
'          Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 9, "REMARK: " & numOrder(StartNum& + 10)
'       Else 'no space found so just break it up in middle of word (add "-" to indicate this break)
'          PrintPrint "REMARK: " & Mid$(numOrder(StartNum& + 10), 1, MaxNumCh&) & "-"
'          PrintCurrentX Xo + 0.5
'          PrintPrint "              " & Mid$(numOrder(StartNum& + 10), MaxNumCh& + 1, lenRemark& - MaxNumCh&)
'          Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 9, "REMARK: " & numOrder(StartNum& + 10)
'          End If
'    Else 'normal length
'       PrintPrint "REMARK: " & numOrder(StartNum& + 10)
'       Write #filprnt%, 1, CInt((Ylast - 0.85 - 0.18 * 3) / 0.18) + 9, "REMARK: " & numOrder(StartNum& + 10)
'       End If
'
'    Exit Sub
'
'errhand:
'   Resume Next
'
'End Sub

Sub MoveToFirstLocatedPoint()
   'user is locating search results on topo map
   'so move the topo map to the first search point
   'that is within the geologic map boundary since
   'in general this is also the boundaries of the 1:50000 topo maps
   
     Dim XPnt As Single, YPnt As Single
     For i& = 1 To numReport&
          If GDReportfrm.lvwReport.ListItems(i&).Selected Then
              XPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(2))
              YPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(3))
              'Check for duplicate coordinates or for zeros.
            If XPnt < 0 Or XPnt > pixwi * DigiZoom.LastZoom Or YPnt < 0 Or YPnt > pixhi * DigiZoom.LastZoom Then
                'coordinates are beyond the map boundaries
              Else
                'record these coordinates and move to them
                If Not DigiRubberSheeting Then
                   GDMDIform.Text5 = XPnt
                   GDMDIform.Text6 = YPnt
                Else
                   GDMDIform.Text5 = GDReportfrm.lvwReport.ListItems(i&).SubItems(4)
                   GDMDIform.Text6 = GDReportfrm.lvwReport.ListItems(i&).SubItems(5)
                   End If
                ce& = 0 'maps have been reloaded so make sure that blinker flag is reset
                Call ShiftMap(CSng(XPnt), CSng(YPnt))
'                Call gotocoord
                Exit For
                End If
             End If
      Next i&
   
End Sub
'Sub CloseDatabase()
'    If linked = True Then gdbs.Close 'close the data base
'    linked = False
'    'erase the temporary database
'    If Dir(direct$ + "\pal_dt_tmp.mdb") <> sEmpty Then Kill direct$ + "\pal_dt_tmp.mdb"
'End Sub
'Sub CloseDatabaseOld()
'    If linkedOld = True Then gdbsOld.Close 'close the data base
'    linkedOld = False
'    GotPassword = False
'    'erase the temporary database
'    If Dir(direct$ + "\pal_old_tmp.mdb") <> sEmpty Then Kill direct$ + "\pal_old_tmp.mdb"
'End Sub
'Sub CloseDatabasepiv()
'    If linkedpiv = True Then gdbspiv.Close 'close the data base
'    linkedpiv = False
'    GotPassword = False
'    'erase the temporary database
'    If Dir(direct$ + "\pal_old_piv_tmp.mdb") <> sEmpty Then Kill direct$ + "\pal_old_piv_tmp.mdb"
'End Sub

'Sub FossilNames(sFossilId As String, sFossilTable As String, _
'                sFosTable As String, combotxt As String, _
'                numFos As Long, sArr() As String, lArr() As Long, _
'                sFosDict As String, FosNames As String)
'                'numFos As Long, LstBox, lArr() As Long, _
'                'sFosDict As String, FosNames As String)
'
'   'this subroutine queries the fossil dic's for the names of
'   'the recorded fossils, if any
'   'inputs are:
'   'combotxt = 'AND or OR combo box value of fossil name search
'   'sFosId = display number, res_id
'   'sFoslTabl = name of Fossil table to query
'   'FosDict = name of Fossil Name Dictionary
'   'lArr = long array containing fossil name id's
'   'FosName = returned string containing names of found
'   '          fossils, if any.
'   '          If FosName is returned empty then search
'   '          was unsuccessful.
'
'    Dim strSQLFos1 As String, strSQLFos As String
'    Dim qdFos As QueryDef
'    Dim rstFos As Recordset
'    Dim i&, numfound&
'
'    On Error GoTo errhand
'
'    FosName = sEmpty
'
'    strSQLFos1 = "SELECT " & sFosTable & ".name FROM " & _
'             sFossilTable & ", " & sFosTable & " " & _
'             "WHERE " & sFossilTable & ".foss = " & sFosTable & ".foss_id " & _
'             "AND " & sFossilTable & ".res_id = " & sFossilId
'
'    If combotxt = "OR" Then
'       'add OR clauses unless searching over any fossil name
'       If lArr(0) = 0 Then
'          'searching for any name, so don't need OR clauses
'          strSQLFos = strSQLFos1
'       Else
'          strSQLFos = strSQLFos1 & " AND ("
'          For i& = 1 To numFos - 1
'             strSQLFos = strSQLFos & sFosTable & ".name = " & Trim$(str$(lArr(i& - 1))) & " OR "
'          Next i&
'          strSQLFos = strSQLFos & sFosTable & ".name = " & Trim$(str$(lArr(numFos - 1))) & ")"
'          End If
'    ElseIf combotxt = "AND" Then
'      'add AND clauses unless searching over any fossil name
'       If lArr(0) = 0 Then
'          'searching for any name, so don't need UNION SELECT clauses
'          strSQLFos = strSQLFos1
'       Else
'          'add UNION SELECT for each fossil name--this makes an AND search
'          strSQLFos = strSQLFos1 & " AND " & sFosTable & ".name = " & Trim$(str$(lArr(0)))
'          For i& = 2 To numFos
'             strSQLFos = strSQLFos & " UNION " & strSQLFos1
'             strSQLFos = strSQLFos & " AND " & sFosTable & ".name = " & Trim$(str$(lArr(i& - 1)))
'          Next i&
'          End If
'       End If
'
'    'query the database
'    Set qdFos = gdbs.CreateQueryDef(sEmpty, strSQLFos & ";")
'
'    'Create a temporary snapshot-type Recordset.
'    Set rstFos = qdFos.OpenRecordset(dbOpenSnapshot)
'
'    With rstFos
'       .MoveLast
'       numfound& = .RecordCount
'       If numfound& <= 0 Then 'no matches
'          rstFos.Close
'          Exit Sub
'       ElseIf combotxt = "AND" And _
'              numfound& < numFos Then 'didn't find all of them
'          rstFos.Close
'          Exit Sub
'       Else 'search successful, record fossil names
'          .MoveFirst
'          Do Until .EOF
'             FosNames = FosNames & sArr(val(rstFos!Name) - 1) & ", "
'             FosNames = FosNames & LstBox.List(Val(rstFos!Name) - 1) & ", "
'
'             'slow way of doing it
'             Dim sqfoss As String
'             Dim rstfoss As Recordset
'
'             sqfoss = "SELECT " & sFosDict & ".* FROM " & sFosDict & _
'                      " WHERE " & sFosDict & ".id = " & rstFos![Name]
'             Set rstfoss = gdbs.OpenRecordset(sqfoss, dbOpenSnapshot)
'
'             With rstfoss
'               .MoveFirst
'
'               'Genera
'               If Not IsNull(rstfoss![name_2]) Then
'                  Name2 = rstfoss![name_2]
'               Else
'                  Name2 = sEmpty
'                  End If
'
'               'Species
'               If Not IsNull(rstfoss![name_1]) Then
'                  Name1 = rstfoss![name_1]
'               Else
'                  Name1 = sEmpty
'                  End If
'
'               'Third name
'               If Not IsNull(rstfoss![name_3]) Then
'                  Name3 = rstfoss![name_3]
'               Else
'                  Name3 = sEmpty
'                  End If
'
'               FosNames = FosNames & LTrim$(Trim$(Name2) & " " & Trim$(Name3) & _
'                    " " & Trim$(Name1)) & ", " 'Genera and species
'
'             End With
'             rstfoss.Close
'
'             .MoveNext
'          Loop
'          FosNames = Trim$(FosNames)
'          End If
'    End With
'    rstFos.Close
'
'    Exit Sub
'
'errhand:
'    Select Case Err.Number
'       Case 3360
'          Screen.MousePointer = vbDefault
'         'stop the animation if activated
'          GDSearchfrm.picAnimation.Visible = False
'          GDSearchfrm.anmSearch.Visible = False
'          GDReportfrm.picAnimation.Visible = False
'          GDReportfrm.anmReport.Visible = False
'          GDReportfrm.anmReport.Stop
'          GDSearchfrm.anmSearch.Stop
'
'          'exceeded maximum number of allowed query statements
'          MsgBox "You have exceeded the maximum number of individual species" & vbLf & _
'                 "that can be searched for at one time!" & vbLf & vbLf & _
'                 "You can search for all the species at once (Any species) or" & vbLf & _
'                 "select fewer (< 250) individual species names and then retry." & vbLf & vbLf & _
'                 "This search will be aborted.", _
'                 vbExclamation + vbOKOnly, "MapDigitizer"
'          StopSearch = True 'Abort the search
'       'Case Else
'       '  let calling routine do other error checking
'    End Select
'
'End Sub


Sub FillPrintCombo()
'    'fill the pages combo box with fossil categories
'
'     PrintPreview.cmbPages.Visible = True
'     PrintPreview.lblPages.Visible = True
'
'     PrintPreview.cmbPages.Clear
'     PrintPreview.cmbPages.AddItem "Page 1: Summary"
'     Pages% = 1
'
'     'make previous and next buttons appear
'     ButtonsPrevNext
'
'     'Find Order number:
'     'If Not PicSum Then
'     If PreviewOrderNum& <> 0 Then
'        OrderNum& = PreviewOrderNum&
'        If OrderNum& < 0 Then
'           'this  is record from old scanned database
'
'           PrintPreview.cmdEditScannedDB.Visible = True
'           PrintPreview.cmdTifView.Visible = True
'           If Dir(tifViewerDir$) <> sEmpty Then
'              PrintPreview.cmdTifView.Enabled = True
'              End If
'
'           GoTo fpcend
'           If Not EditDBVis Then PrintPreview.cmdEditScannedDB.Visible = True
'        Else
'           PrintPreview.cmdEditScannedDB.Visible = False
'           PrintPreview.cmdTifView.Visible = False
'           PrintPreview.cmdTifView.Enabled = False
'           End If
'     Else
'        If InStr(GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(4), "*") Then
'           'old scanned database record
'           pos1& = InStr(GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(4), "*")
'           pos2& = InStr(GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(4), "/")
'           OrderNum& = -val(Mid$(GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(4), _
'                 pos1& + 1, pos2& - pos1& - 1))
'           If Not EditDBVis Then PrintPreview.cmdEditScannedDB.Visible = True
'
'           If OrderNum& < 0 Then
'              PrintPreview.cmdEditScannedDB.Visible = True
'              PrintPreview.cmdTifView.Visible = True
'              If Dir(tifViewerDir$) <> sEmpty Then
'                 PrintPreview.cmdTifView.Enabled = True
'                 End If
'           Else
'              PrintPreview.cmdEditScannedDB.Visible = False
'              PrintPreview.cmdTifView.Visible = False
'              PrintPreview.cmdTifView.Enabled = False
'              End If
'
'           GoTo fpcend
'           End If
'
'        OrderNum& = GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(4) 'order_id
'
'        If OrderNum& < 0 Then
'           PrintPreview.cmdEditScannedDB.Visible = True
'           PrintPreview.cmdTifView.Visible = True
'           If Dir(tifViewerDir$) <> sEmpty Then
'              PrintPreview.cmdTifView.Enabled = True
'              End If
'        Else
'           PrintPreview.cmdEditScannedDB.Visible = False
'           PrintPreview.cmdTifView.Visible = False
'           PrintPreview.cmdTifView.Enabled = False
'           End If
'
'        End If
'
'    'query database for fossil types
'    'and other Order_new fields (can't use the Gdreportfrm
'    'list view info since this is not necessary complete list
'    'depending on the type of search performed)
'     Dim sqOrder As String
'     Dim rstOrder As Recordset
'
'     sqOrder = "SELECT * FROM Order_new WHERE order_id = " & str$(OrderNum&)
'     Set rstOrder = gdbs.OpenRecordset(sqOrder, dbOpenSnapshot)
'
'     doclin$ = sEmpty
'     With rstOrder
'       .MoveFirst
'
'       If IsNull(rstOrder!conod) Then
'       Else
'          If val(rstOrder!conod) <> 0 Then
'             doclin$ = doclin$ & " cono"
'             End If
'          End If
'
'       If IsNull(rstOrder!diato) Then
'       Else
'          If val(rstOrder!diato) <> 0 Then
'             doclin$ = doclin$ & " diato"
'             End If
'          End If
'
'       If IsNull(rstOrder!foram) Then
'       Else
'          If val(rstOrder!foram) <> 0 Then
'             doclin$ = doclin$ & " foram"
'             End If
'          End If
'
'       If IsNull(rstOrder!megaf) Then
'       Else
'          If val(rstOrder!megaf) <> 0 Then
'             doclin$ = doclin$ & " megaf"
'             End If
'          End If
'
'       If IsNull(rstOrder!nano) Then
'       Else
'          If val(rstOrder!nano) <> 0 Then
'             doclin$ = doclin$ & " nanno"
'             End If
'          End If
'
'       If IsNull(rstOrder!ostra) Then
'       Else
'          If val(rstOrder!ostra) <> 0 Then
'             doclin$ = doclin$ & " ostra"
'             End If
'          End If
'
'       If IsNull(rstOrder!palin) Then
'       Else
'          If val(rstOrder!palin) <> 0 Then
'             doclin$ = doclin$ & " palyn"
'             End If
'          End If
'
'     End With
'     rstOrder.Close
'
'    'find the icon associated with this record
'
'    IconNum& = 0
'    If InStr(doclin$, "cono") Then
'       Pages% = Pages% + 1
'       PrintPreview.cmbPages.AddItem "Page " & Trim$(str$(Pages%)) & ": Conodonta"
'       End If
'
'    If InStr(doclin$, "diato") Then
'       Pages% = Pages% + 1
'       PrintPreview.cmbPages.AddItem "Page " & Trim$(str$(Pages%)) & ": Diatom"
'       End If
'
'    If InStr(doclin$, "foram") Then
'       Pages% = Pages% + 1
'       PrintPreview.cmbPages.AddItem "Page " & Trim$(str$(Pages%)) & ": Foraminifera"
'       End If
'
'    If InStr(doclin$, "megaf") Then
'       Pages% = Pages% + 1
'       PrintPreview.cmbPages.AddItem "Page " & Trim$(str$(Pages%)) & ": Megafauna"
'       End If
'
'    If InStr(doclin$, "nan") Then
'       Pages% = Pages% + 1
'       PrintPreview.cmbPages.AddItem "Page " & Trim$(str$(Pages%)) & ": Nannoplankton"
'       End If
'
'    If InStr(doclin$, "ostra") Then
'       Pages% = Pages% + 1
'       PrintPreview.cmbPages.AddItem "Page " & Trim$(str$(Pages%)) & ": Ostracoda"
'       End If
'
'    If InStr(doclin$, "palyn") Then
'       Pages% = Pages% + 1
'       PrintPreview.cmbPages.AddItem "Page " & Trim$(str$(Pages%)) & ": Palynology"
'       End If
'
'fpcend:
'    PrintPreview.cmbPages.ListIndex = 0 'set to first item
'
End Sub
'Sub PrintFosNames(fosstbl$, FosTbl$, FosDic$, _
'              FosID As Long, Xfos As Single, Yfos As Single, filprnt%)
'
'   'print fossil names on printer, print preview
'
'    On Error GoTo errhand
'
'    Dim idvalue$
'
'    idvalue$ = str$(FosID)
'
'    Dim qdfDV As QueryDef
'    Dim rstDV As Recordset
'    Dim strSqlDV As String
'
'    'create query string
'    strSqlDV = "SELECT " & FosTbl$ & ".* FROM " & _
'             fosstbl$ & ", " & FosTbl$ & " " & _
'             "WHERE " & fosstbl$ & ".foss = " & FosTbl$ & ".foss_id " & _
'             "AND " & fosstbl$ & ".res_id = " & idvalue$
'
'    'query the database
'    Set qdfDV = gdbs.CreateQueryDef(sEmpty, strSqlDV & ";")
'
'    'Create a temporary snapshot-type Recordset.
'    Set rstDV = qdfDV.OpenRecordset(dbOpenSnapshot)
'
'    Dim SemiQua As String
'    Dim Features As String
'    Dim Quantity As String
'    Dim num&, numPrint&
'    Dim FossilName As String
'    Dim SemiQuant As String
'    Dim Featr As String
'
'    'defaults
'    FossilName = "No recorded fossil names"
'    Quantity = sEmpty
'    SemiQuant = sEmpty
'    Featr = sEmpty
'
'    numPrint& = 0
'
'   'print out the results
'    With rstDV
'       .MoveLast
'       numfound& = .RecordCount
'       If numfound& <= 0 Then 'no matches
'          'print defaults
'          numPrint& = numPrint& + 1
'          GoSub PrintFossilNames
'       Else 'search successful, record fossil names
'          .MoveFirst
'          Do Until .EOF
'
'            numPrint& = numPrint& + 1
'
'            'Fossil Name
'            If IsNull(rstDV![Name]) Then
'               'print the defaults
'               numPrint& = numPrint& + 1
'               GoSub PrintFossilNames
'            Else
'
'           '--------query dic table-----------
'              'If Not PicSum Then
'                 'query dic table for fossil name
'
'                 Dim sqfoss As String
'                 Dim rstfoss As Recordset
'
'                 sqfoss = "SELECT " & FosDic$ & ".* FROM " & FosDic$ & _
'                          " WHERE " & FosDic$ & ".id = " & rstDV![Name]
'                 Set rstfoss = gdbs.OpenRecordset(sqfoss, dbOpenSnapshot)
'
'                 With rstfoss
'                   .MoveFirst
'
'                   'Genera
'                   If Not IsNull(rstfoss![name_2]) Then
'                      Name2 = rstfoss![name_2]
'                   Else
'                      Name2 = sEmpty
'                      End If
'
'                   'Species
'                   If Not IsNull(rstfoss![name_1]) Then
'                      Name1 = rstfoss![name_1]
'                   Else
'                      Name1 = sEmpty
'                      End If
'
'                   'Third name
'                   If Not IsNull(rstfoss![name_3]) Then
'                      Name3 = rstfoss![name_3]
'                   Else
'                      Name3 = sEmpty
'                      End If
'
'                   FossilName = LTrim$(Trim$(Name2) & " " & Trim$(Name3) & _
'                        " " & Trim$(Name1))  'Genera and species
'
'                 End With
'                 rstfoss.Close
'
'                 '-----------end query dic table--------
'
'              'Else
'               '   FossilName = Trim$(LstBox.List(Val(rstDV![Name]) - 1))
'               '   End If
'
'           'semi qua
'           If IsNull(rstDV![sq]) Then
'              SemiQuant = sEmpty
'           Else
'              num& = rstDV![sq]
'              GoSub findsq
'              SemiQuant = SemiQua
'              End If
'
'           'feature
'           If IsNull(rstDV![Feature]) Then
'              Featr = sEmpty
'           Else
'              num& = rstDV![Feature]
'              GoSub findFeature
'              Featr = Features
'              End If
'
'           'quantity
'           If IsNull(rstDV![Qun]) Then
'              Quantity = sEmpty
'           Else
'              If val(rstDV![Qun]) = 0 Then
'                 'obviously there were more than 0, but
'                 'the analyst didn't input how many,
'                 'and the database adds 0 automatically as the default
'                 Quantity = sEmpty
'              Else
'                 Quantity = Trim$(rstDV![Qun])
'                 End If
'              End If
'
'           'print the results
'           GoSub PrintFossilNames
'
'           End If
'
'
'            .MoveNext
'          Loop
'
'          End If
'    End With
'    rstDV.Close
'
'   Exit Sub
'
'findsq: 'inline gosub that determines the prezone string
'      Select Case num&
'         Case 0, 1
'           SemiQua = sEmpty
'         Case 2
'           SemiQua = "Abundant"
'         Case 3
'           SemiQua = "Common"
'         Case 4
'           SemiQua = "Frequent"
'         Case 5
'           SemiQua = "Rare"
'      End Select
'Return
'
'findFeature: 'inline gosub that determines the prezone string
'      Select Case num&
'         Case 0
'           Features = sEmpty
'         Case 1
'           Features = "CF"
'         Case 2
'           Features = "Caving"
'         Case 3
'           Features = "?"
'         Case 4
'           Features = "Reworked"
'      End Select
'Return
'
'PrintFossilNames:
'           PrintCurrentX Xfos
'           PrintCurrentY Yfos + 0.2 * (numPrint& - 1)
'           PrintPrint FossilName
'           Write #filprnt%, 1, CInt((Yfos - 0.85) / 0.18) + numPrint&, FossilName
'
'           PrintCurrentX Xfos + 3
'           PrintCurrentY Yfos + 0.2 * (numPrint& - 1)
'           PrintPrint SemiQuant
'           Write #filprnt%, 2, CInt((Yfos - 0.85) / 0.18) + numPrint&, SemiQuant
'
'           PrintCurrentX Xfos + 4.5
'           PrintCurrentY Yfos + 0.2 * (numPrint& - 1)
'           PrintPrint Featr
'           Write #filprnt%, 3, CInt((Yfos - 0.85) / 0.18) + numPrint&, Featr
'
'           PrintCurrentX Xfos + 6
'           PrintCurrentY Yfos + 0.2 * (numPrint& - 1)
'           PrintPrint Quantity
'           Write #filprnt%, 4, CInt((Yfos - 0.85) / 0.18) + numPrint&, Quantity
'Return
'
'errhand:
'    Select Case Err.Number
'       Case 3021
'          'empty record, resume next to print defaults
'          Resume Next
'       Case Else
'
'           Screen.MousePointer = vbDefault
'
'           MsgBox "Encountered error #: " & Err.Number & vbLf & _
'                  Err.Description & vbLf & _
'                  "You probably won't be able to obtain a complete print preview!" & vbLf & _
'                  sEmpty, vbExclamation + vbOKOnly, "MapDigitizer"
'    End Select
'
'
'End Sub
Sub ButtonsPrevNext()
   'make previous and next buttons appear
   If Not PicSum Then
      PrintPreview.cmdPrevious.Visible = True
      PrintPreview.cmdNext.Visible = True
   ElseIf (PicSum And numReport& > 1) Then
      PrintPreview.cmdPrevious.Visible = True
      PrintPreview.cmdNext.Visible = True
      End If
      
   If Not PicSum Then
   
'      If SearchDBs% = 1 Then 'searching over all active database
'         MaxOrder& = MaxOrderNum&(0)
'         MinOrder& = -MaxOrderNum&(1) '-59264
'      ElseIf SearchDBs% = 2 Then 'searching only over active database
'         MaxOrder& = MaxOrderNum&(0)
'         MinOrder& = 1
'      ElseIf SearchDBs% = 3 Then 'searching only over the scanned database
'         MaxOrder& = -1
'         MinOrder& = -MaxOrderNum&(1) '-59264
'         End If
        
      If PreviewOrderNum& < MaxOrder& Then
         PrintPreview.cmdNext.Enabled = True
      Else
         PrintPreview.cmdNext.Enabled = False
         End If
           
      If PreviewOrderNum& > MinOrder& Then
         PrintPreview.cmdPrevious.Enabled = True
      Else
         PrintPreview.cmdPrevious.Enabled = False
         End If
           
   Else
        
      If NewHighlighted& < numReport& Then
         PrintPreview.cmdNext.Enabled = True
      Else
         PrintPreview.cmdNext.Enabled = False
         End If
           
      If NewHighlighted& > 1 Then
         PrintPreview.cmdPrevious.Enabled = True
      Else
         PrintPreview.cmdPrevious.Enabled = False
         End If
           
      End If
      

End Sub
''---------------------------------------------------------------------------------------
'' Procedure : loadOldDbArrays
'' DateTime  : 10/21/2002 23:01
'' Purpose   : load arrays that determine text values for old paleont database
''---------------------------------------------------------------------------------------
''
'Sub LoadOldDbArrays()
'    On Error GoTo errhand
'
'    Dim qdFos As QueryDef
'    Dim rstFos As Recordset
'    'Dim rstForm As Recordset
'    Dim strSQLFos As String
'    'Dim strSQLForm As String
'    Dim lenForm As Integer
'
'    'open database and data tables
'
'    arrN03(0) = sEmpty
'    arrN03(1) = "core"
'    arrN03(2) = "cutting"
'    arrN05(0) = sEmpty
'    arrN05(1) = "well"
'    arrN05(2) = "surface"
'    arrN06(3) = "."
'    arrN06(0) = "Early"
'    arrN06(1) = "Middle"
'    arrN06(2) = "Late"
'
'    If numOldDates > 0 Then
'      'already filled up arrays, so skip this routine
'      'after filling up place/well name dictionary for scanned database
'
'       If numN11 > 0 Then
'            'load suggested search strings (scanned database names) into cmbDictonary
'            'and enabled the dictionary
'
'            With GDSearchfrm
'               .frmDictionary.Enabled = True
'               .cmdPasteDictionary.Enabled = True
'               .cmbDictionary.Enabled = True
'
'               .cmbDictionary.Clear
'               For i& = 1 To numN11
'                   If arrN11(i& - 1) = sEmpty Then
'                   Else
'                      .cmbDictionary.AddItem arrN11(i& - 1)
'                      End If
'               Next i&
'               .cmbDictionary.ListIndex = 0
'             End With
'
'        End If
'
'        Exit Sub
'    Else
'       'array containing conversion from the old database's age date numbers
'       'to the new (active) database's age number (which is displayed in the TreeView)
'       ReDim arrOldDates(GDSearchfrm.lstDates.ListCount)
'       numOldDates = GDSearchfrm.lstDates.ListCount
'       End If
'
'
'    Screen.MousePointer = vbHourglass
'
'
'    '------------Fill names array----------------------------------
'    If numN11 = 0 Then 'haven't yet filled it
'        strSQLFos = "SELECT DOMAINS_DATA.V_KEY, DOMAINS_DATA.V_VAL FROM DOMAINS_DATA " & _
'                    "WHERE D_KEY = 5"
'
'        'query the database
'        Set qdFos = gdbsOld.CreateQueryDef(sEmpty, strSQLFos & ";")
'
'        'Create a temporary snapshot-type Recordset.
'        Set rstFos = qdFos.OpenRecordset(dbOpenSnapshot)
'
'        With rstFos
'            .MoveLast
'            numfound& = .RecordCount
'            .MoveFirst
'            nn& = 0
'            Do Until .EOF
'               If val(rstFos!V_KEY) > nn& Then
'                  nn& = rstFos!V_KEY
'                  ReDim Preserve arrN11(nn&)
'                  numN11 = nn&
'                  End If
'               arrN11(val(rstFos!V_KEY) - 1) = Trim$(rstFos!V_VAL)
'               .MoveNext
'            Loop
'        End With
'        rstFos.Close
'        End If
'
'       'load suggested search strings (scanned database names) into cmbDictonary
'       'and enabled the dictionary
'       With GDSearchfrm
'          .frmDictionary.Enabled = True
'          .cmdPasteDictionary.Enabled = True
'          .cmbDictionary.Enabled = True
'
'          .cmbDictionary.Clear
'          For i& = 1 To numN11
'              If arrN11(i& - 1) = sEmpty Then
'              Else
'                 .cmbDictionary.AddItem arrN11(i& - 1)
'                 End If
'          Next i&
'          .cmbDictionary.ListIndex = 0
'        End With
'
'
'    '------------Fill ages array----------------------------------
'    If numN07 = 0 Then 'didn't yet fill it
'
'        strSQLFos = "SELECT DOMAINS_DATA.V_KEY, DOMAINS_DATA.V_VAL FROM DOMAINS_DATA " & _
'                    "WHERE D_KEY = 7"
'
'        'query the database
'        Set qdFos = gdbsOld.CreateQueryDef(sEmpty, strSQLFos & ";")
'
'        'Create a temporary snapshot-type Recordset.
'        Set rstFos = qdFos.OpenRecordset(dbOpenSnapshot)
'
'        With rstFos
'            .MoveLast
'            numfound& = .RecordCount
'            .MoveFirst
'            nn& = 0
'            Do Until .EOF
'               If val(rstFos!V_KEY) > nn& Then
'                  nn& = rstFos!V_KEY
'                  ReDim Preserve arrN07(nn&)
'                  numN07 = nn&
'                  End If
'               arrN07(val(rstFos!V_KEY) - 1) = Trim$(rstFos!V_VAL)
'
'                'now convert this scanned age number to the active database age number
'                'and store the conversion in the array arrOldDates
'                For i& = 0 To GDSearchfrm.lstDates.ListCount - 1
'                   If UCase$(GDSearchfrm.lstDates.List(i&)) = UCase$(Trim$(rstFos!V_VAL)) Then
'                      arrOldDates(i&) = val(rstFos!V_KEY)
'                      Exit For
'                      End If
'                Next i&
'
'              .MoveNext
'            Loop
'        End With
'        rstFos.Close
'
'   Else 'already filled age array, so just fill conversion array
'      For j& = 1 To numN07
'
'        If arrN07(j& - 1) = sEmpty Then GoTo loa250
'
'        For i& = 0 To GDSearchfrm.lstDates.ListCount - 1
'           If UCase$(GDSearchfrm.lstDates.List(i&)) = UCase$(arrN07(j& - 1)) Then
'              arrOldDates(i&) = j&
'              Exit For
'              End If
'        Next i&
'
'loa250:
'      Next j&
'
'   End If
'
'    '------------Fill formation array----------------------------------
'    'This takes so much time it might have to be skipped
'    'and the selected formation names have to be added to the
'    'SQL statement as names and not as numbers.  The best thing would be
'    'to convert the names and numbers in the database itself
'
'    GDMDIform.StatusBar1.Panels(1) = "Loading scanned (old) database's Formation Names, please wait..."
'    GDMDIform.StatusBar1.Panels(2) = "0 %"
'
'    If numN10 = 0 Then 'haven't yet loaded the formation names
'
'        strSQLFos = "SELECT DOMAINS_DATA.V_KEY, DOMAINS_DATA.V_VAL FROM DOMAINS_DATA " & _
'                    "WHERE D_KEY = 4"
'
'        'query the database
'        Set qdFos = gdbsOld.CreateQueryDef(sEmpty, strSQLFos & ";")
'
'        'Create a temporary snapshot-type Recordset.
'        Set rstFos = qdFos.OpenRecordset(dbOpenSnapshot)
'
'        'filout% = FreeFile
'        'Open "c:\jk\OldFormation.txt" For Output As #filout%
'        With rstFos
'            .MoveLast
'            numfound& = .RecordCount
'
'            GDMDIform.prbSearch.Enabled = True
'            GDMDIform.prbSearch.Visible = True
'            GDMDIform.prbSearch.Left = 10200 '10000
'            GDMDIform.prbSearch.Width = 1895
'            GDMDIform.prbSearch.Min = 0
'            GDMDIform.prbSearch.Max = numfound&
'            GDMDIform.prbSearch.Value = 0
'
'            .MoveFirst
'            nn& = 0
'            mm& = 0
'            Do Until .EOF
'               mm& = mm& + 1
'               GDMDIform.prbSearch.Value = mm&
'               GDMDIform.StatusBar1.Panels(2) = str$(CLng(100 * mm& / numfound&)) & "%"
'
'               If val(rstFos!V_KEY) > nn& Then
'                  nn& = rstFos!V_KEY
'                  ReDim Preserve arrN10(nn&)
'                  numN10 = nn&
'                  ReDim Preserve arrOldFormation(nn&)
'                  End If
'
'               arrN10(val(rstFos!V_KEY) - 1) = Trim$(rstFos!V_VAL)
'               'there is a quite a difference between the formation names
'               'some of them only match with the first word so find it
'               If InStr(1, arrN10(val(rstFos!V_KEY) - 1), " ") Then
'                  OldForm$ = UCase$(Mid$(arrN10(val(rstFos!V_KEY) - 1), 1, InStr(1, arrN10(val(rstFos!V_KEY) - 1), " ") - 1))
'               Else
'                  OldForm$ = UCase$(arrN10(val(rstFos!V_KEY) - 1))
'                  End If
'
'               'Now convert this old formation number to a new formation number
'               'and store the conversion in the array arrOldFormation.
'               ''This must be done using a query statement since it takes too long
'               ''to loop through all the formations to find the right name.
'               ''strSQLForm = "SELECT FORMdic.* FROM FORMdic WHERE UCase(Formdic![name]) = ""SEDOM FM."""
'               'strSQLForm = "SELECT FORMdic.[id] FROM FORMdic " & _
'               '         "WHERE UCase(FORMdic![Name]) = " & Chr(34) & arrN10(Val(rstFos!V_KEY) - 1) & Chr(34) & _
'               '         " OR UCase(FORMdic![Name]) = " & Chr(34) & arrN10(Val(rstFos!V_KEY) - 1) & " Fm." & Chr(34)
'               '
'               'Set rstForm = gdbs.OpenRecordset(strSQLForm, dbOpenSnapshot)
'        '
'        '       With rstForm
'        '          .MoveFirst
'        '          arrOldFormation(Val(rstForm!ID) - 1) = nn&
'        '       End With
'        '       rstForm.Close
'
'                found% = 0
'                For i& = 0 To GDSearchfrm.lstFormationUnsorted.ListCount - 1
'                   'When identifying remove all the 2nd word qualifiers like "Fm.", "Group", etc."
'                   If InStr(1, GDSearchfrm.lstFormationUnsorted.List(i&), " ") Then
'                      NewForm$ = UCase$(Mid$(GDSearchfrm.lstFormationUnsorted.List(i&), 1, InStr(1, GDSearchfrm.lstFormationUnsorted.List(i&), " ") - 1))
'                   Else
'                      NewForm$ = UCase$(GDSearchfrm.lstFormationUnsorted.List(i&))
'                      End If
'                   If NewForm$ = OldForm$ Then
'                      arrOldFormation(val(rstFos!V_KEY) - 1) = i&
'                      found% = 1
'                      Exit For
'                      End If
'                Next i&
'
'                'In the ideal case, any new formation would be added to the Formation catalogue
'                'of the new database.  However, this may never be practical 100%, so try this fix instead
'                If found% = 0 Then 'didn't find this formation, so add it to list
'                   GDSearchfrm.lstFormation.AddItem Trim$(rstFos!V_VAL)
'                   GDSearchfrm.lstFormationUnsorted.AddItem Trim$(rstFos!V_VAL)
'                   arrOldFormation(val(rstFos!V_KEY) - 1) = GDSearchfrm.lstFormationUnsorted.ListCount - 1
'                   End If
'
'              .MoveNext
'            Loop
'        End With
'        rstFos.Close
'
'    Else 'already loaded arrN10, so just do conversion from old scanned
'         'database formation names to the active database's formation names
'
'       numfound& = numN10
'       GDMDIform.prbSearch.Enabled = True
'       GDMDIform.prbSearch.Visible = True
'       GDMDIform.prbSearch.Left = 10200 '10000
'       GDMDIform.prbSearch.Width = 1895
'       GDMDIform.prbSearch.Min = 0
'       GDMDIform.prbSearch.Max = numfound&
'       GDMDIform.prbSearch.Value = 0
'
'       ReDim Preserve arrOldFormation(numN10)
'       For j& = 1 To numN10
'
'            GDMDIform.prbSearch.Value = j&
'            GDMDIform.StatusBar1.Panels(2) = str$(CLng(100 * j& / numfound&)) & "%"
'
'            If InStr(1, arrN10(j& - 1), " ") Then
'               OldForm$ = UCase$(Mid$(arrN10(j& - 1), 1, InStr(1, arrN10(j& - 1), " ") - 1))
'            Else
'               OldForm$ = UCase$(arrN10(j& - 1))
'               End If
'
'            If OldForm$ = sEmpty Then GoTo loa500
'
'            found% = 0
'            For i& = 0 To GDSearchfrm.lstFormationUnsorted.ListCount - 1
'               'When identifying remove all the 2nd word qualifiers like "Fm.", "Group", etc."
'               If InStr(1, GDSearchfrm.lstFormationUnsorted.List(i&), " ") Then
'                  NewForm$ = UCase$(Mid$(GDSearchfrm.lstFormationUnsorted.List(i&), 1, InStr(1, GDSearchfrm.lstFormationUnsorted.List(i&), " ") - 1))
'               Else
'                  NewForm$ = UCase$(GDSearchfrm.lstFormationUnsorted.List(i&))
'                  End If
'               If NewForm$ = OldForm$ Then
'                  arrOldFormation(j& - 1) = i&
'                  found% = 1
'                  Exit For
'                  End If
'            Next i&
'
'            'In the ideal case, any new formation would be added to the Formation catalogue
'            'of the new database.  However, this may never be practical 100%, so try this fix instead
'            If found% = 0 Then 'didn't find this formation, so add it to list
'               GDSearchfrm.lstFormation.AddItem UCase$(arrN10(j& - 1))
'               GDSearchfrm.lstFormationUnsorted.AddItem UCase$(arrN10(j& - 1))
'               arrOldFormation(j& - 1) = GDSearchfrm.lstFormationUnsorted.ListCount - 1
'               End If
'
'loa500:
'       Next j&
'
'
'    End If
'    'Close #filout%
'
'    GDMDIform.prbSearch.Enabled = False
'    GDMDIform.prbSearch.Visible = False
'    GDMDIform.StatusBar1.Panels(1) = sEmpty
'    GDMDIform.StatusBar1.Panels(2) = sEmpty
'
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'
'errhand:
'   Select Case Err.Number
'       Case 3021
'          Resume Next
'       Case Else
'          Screen.MousePointer = vbDefault
'          GDMDIform.prbSearch.Enabled = False
'          GDMDIform.prbSearch.Visible = False
'          GDMDIform.StatusBar1.Panels(1) = sEmpty
'          GDMDIform.StatusBar1.Panels(2) = sEmpty
'          MsgBox "Encountered error #: " & Err.Number & vbLf & _
'               Err.Description & vbLf & _
'               sEmpty, vbCritical + vbOKOnly, "MapDigitizer"
'   End Select
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : CreateOldDbSql
'' DateTime  : 10/21/2002 23:15
'' Purpose   : create SQL string for searching old paleont database
''---------------------------------------------------------------------------------------
''
'Sub CreateOldDbSql()
'
'   Dim NullCoord As Boolean
'   Dim strSqlCoord As String
'   Dim strSqlAge As String
'   Dim strSqlAgeEarly As String
'   Dim strSqlAgeLate As String
'   Dim sSortLoc As String
'   Dim sSortWell As String
'   Dim FossilSelected As Boolean
'   Dim strNot As String
'
'   On Error GoTo CreateOldDbSql_Error
'
'''---------------------------------------------------------
'    With GDSearchfrm
'
'    'check if searching over analysts, companies, clients, or
'    'dates or project no.'s of the active database
'    If sAnalystSearch <> sEmpty Then
'       strSQLOld = gsEmpty 'don't search over scanned database
'       Exit Sub
'       End If
'    If sFormSearch <> sEmpty Then
'       If .lstClient.Enabled Or _
'          .lstCompany.Enabled Then
'          strSQLOld = gsEmpty 'don't search over scanned database
'          Exit Sub
'          End If
'      If (.lstDate.Enabled Or _
'         .lstProject) And _
'         sFormSearchOld = sEmpty Then
'         strSQLOld = gsEmpty 'don't search over scanned database
'         Exit Sub
'         End If
'      End If
'
'   'now form the query string for the scanned database
'
'    NullCoord = False
'    If val(.txtEastMin) = 0 And val(.txtEastMax) = 0 And val(.txtNorthMin) = 0 And val(.txtNorthMax) = 0 Then
'       NullCoord = True 'don't search over coordinate boundaries
'       End If
'
'    ' Define the parameters clause unless all coordinates are 0.
'    strSqlCoord = sEmpty
'    If Not NullCoord Then
'
'      strSqlCoord = "AND VAL(OBJECTS2![A08]) >= " & .txtEastMin & " AND " & _
'                    "VAL(OBJECTS2![A08]) <= " & .txtEastMax & " AND " & _
'                    "VAL(OBJECTS2![A09]) >= " & .txtNorthMin & " AND " & _
'                    "VAL(OBJECTS2![A09]) <= " & .txtNorthMax & " "
'
'       End If
'
'    'check if any of the fossil categories have been selected
'    FossilSelected = False
'    If .chkConodonta.Value = vbChecked Or .chkDiatom.Value = vbChecked Or _
'      .chkForaminifera.Value = vbChecked Or .chkMegafauna.Value = vbChecked Or _
'      .chkNanoplankton.Value = vbChecked Or .chkOstracoda.Value = vbChecked Or _
'      .chkPalynology.Value = vbChecked Or .chkShekef.Value = vbChecked Then
'      FossilSelected = True
'      End If
'
'    '------------SQL clause for Fossil Types-------------
'    If .txtSQLdb2.Visible = True And .txtSQLdb2.Enabled = True Then
'      'use the user defined value of strSqlCategory
'    Else
'      '--form the BOOLEAN search string over fossil tables
'      Createdb2SqlFossil
'      End If
'
'    If FossilSelected And strSqlCategory = gsEmpty Then
'       strSQLOld = gsEmpty 'don't search over scanned database
'       Exit Sub
'       End If
'
'    '----------define sample depth SQL string (Min and Max depths w.r.t. sea level
'    'is equal to: Ground Level - limup, Ground Level - limdo)
'    If .txtLimdo = 0 And .txtLimup = 0 Then
'       strSqlDepth = " "
'    Else
'       strSqlDepth = " AND (OBJECTS2![N04] >= " & Trim$(.txtLimup) & _
'                     " AND OBJECTS2![N04] <= " & Trim$(.txtLimdo) & ")"
'       End If
'
'    '----------------can do this with a regular search over the results------
'   ' '------------make Formation SQL string
'   '  strSqlFormation = sEmpty
'   '  If .lstFormation.Enabled = False Then
'   '    'searching over all formations
'   '  Else
'   '    found% = 0
'   '    For j& = 1 To .lstFormationUnsorted.ListCount
'   '      If .lstFormationUnsorted.Selected(j& - 1) Then
'   '        If found% = 0 Then
'   '           strSqlFormation = " AND (OBJECTS2![N10] = " & Trim$(Str$(arrOldFormation(j&)))
'   '        Else
'   '           strSqlFormation = strSqlFormation & " OR OBJECTS2![N10] = " & Trim$(Str$(arrOldFormation(j&)))
'   '           End If
'   '        End If
'   '        found% = 1
'   '    Next j&
'   '    strSqlFormation = strSqlFormation & ") "
'   '    End If
'
'    '--------------make Geologic AGge SQL string
'    strSqlAge = sEmpty
'    If .TreeView1.Enabled = True Then
'       'first determine Earlier time prefix
'       Select Case .cmbEPre.ListIndex
'          Case 0
'             numPreEarly% = 1
'          Case 1
'             numPreEarly% = 4
'          Case 2
'             numPreEarly% = 3
'          Case 3
'             numPreEarly% = 2
'        End Select
'
'       'now determine Early Age SQL string
'        For i& = 0 To .lstDates.ListCount - 1
'           If .lstDates.List(i&) = Trim$(.txtEarlier) Then
'              numEarlier& = i&
'              Exit For
'              End If
'        Next i&
'
'       'now determine Later time prefix
'       Select Case .cmbLPre.ListIndex
'          Case 0
'             numPreLate% = 1
'          Case 1
'             numPreLate% = 4
'          Case 2
'             numPreLate% = 3
'           Case 3
'             numPreLate% = 2
'       End Select
'
'       'now determine Later Age SQL string
'        For i& = 0 To .lstDates.ListCount - 1
'           If .lstDates.List(i&) = Trim$(.txtLater) Then
'              numLater& = i&
'              Exit For
'              End If
'        Next i&
'
'        'now put them together
'        If RangeOfDates Then 'range of dates
''           'require both earlier and later dates be in range
''           StartAge% = 0
''           For i& = numLater& To numEarlier&
''              If StartAge% = 0 And arrOldDates(i&) <> 0 Then
''                 strSqlAgeEarly = " AND (OBJECTS2![N07] = " & Trim$(Str$(arrOldDates(i&)))
''                 strSqlAgeLate = " AND (OBJECTS2![N09] = " & Trim$(Str$(arrOldDates(i&)))
''                 StartAge% = 1
''              ElseIf StartAge% = 1 And arrOldDates(i&) <> 0 Then
''                 strSqlAgeEarly = strSqlAgeEarly & " OR OBJECTS2![N07] = " & Trim$(Str$(arrOldDates(i&)))
''                 strSqlAgeLate = strSqlAgeLate & " OR OBJECTS2![N09] = " & Trim$(Str$(arrOldDates(i&)))
''                 End If
''           Next i&
''           strSqlAge = strSqlAgeEarly & ")" & strSqlAgeLate & ") "
'
''          'don't require that both earlier and later dates be in range
'           'rather, even if one of them is in the range, then accept the record
'           StartAge% = 0
'           For i& = numLater& To numEarlier&
'              If StartAge% = 0 And arrOldDates(i&) <> 0 Then
'                 strSqlAge = " AND (OBJECTS2![N07] = " & Trim$(str$(arrOldDates(i&)) & " OR OBJECTS2![N09] = " & Trim$(str$(arrOldDates(i&))))
'                 StartAge% = 1
'              ElseIf StartAge% = 1 And arrOldDates(i&) <> 0 Then
'                 strSqlAge = strSqlAge & " OR OBJECTS2![N07] = " & Trim$(str$(arrOldDates(i&))) & " OR OBJECTS2![N09] = " & Trim$(str$(arrOldDates(i&)))
'                 End If
'           Next i&
'           strSqlAge = strSqlAge & ") "
'
'        Else 'exact dates
'
'           strSqlAge = " AND (OBJECTS2![N06] = " & Trim$(str$(numPreEarly%)) & _
'                       " AND OBJECTS2![N07] = " & Trim$(str$(arrOldDates(numEarlier&))) & _
'                       " AND OBJECTS2![N08] = " & Trim$(str$(numPreLate%)) & _
'                       " AND OBJECTS2![N09] = " & Trim$(str$(arrOldDates(numLater&))) & ") "
'           End If
'
'       End If
'
'
'    SearchAll = False 'set this as default at start of all searches
'
'    '-------------add sorting clause to SQL query since sorting
'    sSortLoc = " " '"ORDER BY Trim(Loc![place]) ASC"
'    sSortWell = " " '"ORDER BY Trim(Wellscat![Name]) ASC"
'    '(N.B., The ORDER BY phrase was removed in order to optimize searching.
'    'Sorting will be done using the sort option of the list view control)
'
'    If .chkOutcroppings.Value = vbChecked And .chkWells.Value = vbChecked Then
'       SearchAll = True
'       'search both outcroppings and wells (and records not marked as either)
'
'            strSQLOld = "SELECT OBJECTS2.* FROM OBJECTS2 " & _
'                     "WHERE OBJECTS2![N05] = 2 " & _
'                     strSqlCoord & _
'                     strSqlCategory & _
'                     strSqlAge & _
'                     "UNION ALL SELECT OBJECTS2.* FROM OBJECTS2 " & _
'                     "WHERE OBJECTS2![N05] = 1 " & _
'                     strSqlCoord & _
'                     strSqlDepth & _
'                     strSqlCategory & _
'                     strSqlAge & _
'                     "UNION ALL SELECT OBJECTS2.* FROM OBJECTS2 " & _
'                     "WHERE OBJECTS2![N05] = 0 " & _
'                     strSqlCoord & _
'                     strSqlCategory & _
'                     strSqlAge
'
'    ElseIf .chkOutcroppings.Value = vbChecked And .chkWells.Value = vbUnchecked Then
'
'            strSQLOld = "SELECT OBJECTS2.* FROM OBJECTS2 " & _
'                     "WHERE OBJECTS2![N05] = 2 " & _
'                     strSqlCoord & _
'                     strSqlCategory & _
'                     strSqlAge & _
'                     sSortLoc
'
'    ElseIf .chkOutcroppings.Value = vbUnchecked And .chkWells.Value = vbChecked Then
'         If .chkAllWells.Value = vbChecked Then
'
'               strSQLOld = "SELECT OBJECTS2.* FROM OBJECTS2 " & _
'                     "WHERE OBJECTS2![N05] = 1 " & _
'                     strSqlCoord & _
'                     strSqlDepth & _
'                     strSqlCategory & _
'                     strSqlAge & _
'                     sSortWell
'
'         ElseIf .chkJustCuttings.Value = vbChecked Then
'
'               strSQLOld = "SELECT OBJECTS2.* FROM OBJECTS2 " & _
'                     "WHERE OBJECTS2![N05] = 1 AND OBJECTS2![N03] = 2 " & _
'                     strSqlCoord & _
'                     strSqlDepth & _
'                     strSqlCategory & _
'                     strSqlAge & _
'                     sSortWell
'
'         ElseIf .chkJustCores.Value = vbChecked Then
'
'               strSQLOld = "SELECT OBJECTS2.* FROM OBJECTS2 " & _
'                     "WHERE OBJECTS2![N05] = 1 AND OBJECTS2![N03] = 1 " & _
'                     strSqlCoord & _
'                     strSqlDepth & _
'                     strSqlCategory & _
'                     strSqlAge & _
'                     sSortWell
'
'             End If
'         End If
'
'    End With
'
'    Exit Sub 'skip the test
'
'    '------debugging test of the SQL------------------
'
'    filtst% = FreeFile
'    Open "c:\jk\SQLtest.txt" For Output As #filtst%
'    Print #filtst%, strSQLOld
'    Close #filtst%
'
'    Dim rst As Recordset
'
'    Set rst = gdbsOld.OpenRecordset(strSQLOld, dbOpenSnapshot)
'
'    rst.MoveLast 'populate the recordset
'
'    RecordNum& = rst.RecordCount
'    GDMDIform.StatusBar1.Panels(2) = "Found: " & RecordNum&
'
'    rst.Close
'
'   On Error GoTo 0
'   Exit Sub
'
'CreateOldDbSql_Error:
'    rst.Close
'    Select Case Err.Number
'       Case 3021 'nor results found
'         GDMDIform.StatusBar1.Panels(2) = "Found: 0"
'       Case Else
'         Screen.MousePointer = vbDefault
'         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateOldDbSql of Module modGDModule"
'    End Select
'
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : LoadOldDb
'' DateTime  : 10/28/2002 21:26
'' Purpose   : store geologic age dates of scanned database into their arrays
''              if not already done
''              Also load up formation names of old database dictionary if
''              not already done
''---------------------------------------------------------------------------------------
''
'Sub LoadOldDb()
'
'    On Error GoTo errhand
'
'    If numN07 > 0 And numN10 > 0 And numN11 > 0 Then
'      'already filled up arrays, so skip this routine
'       Exit Sub
'       End If
'
'    Screen.MousePointer = vbHourglass
'
'loaa10:
'    arrN03(0) = sEmpty
'    arrN03(1) = "core"
'    arrN03(2) = "cutting"
'    arrN05(0) = sEmpty
'    arrN05(1) = "well"
'    arrN05(2) = "surface"
'    arrN06(3) = "."
'    arrN06(0) = "Early"
'    arrN06(1) = "Middle"
'    arrN06(2) = "Late"
'
'
'    '------------Fill ages array----------------------------------
'    Dim strSQLFos As String
'    Dim qdFos As QueryDef
'    Dim rstFos As Recordset
'
'    strSQLFos = "SELECT DOMAINS_DATA.V_KEY, DOMAINS_DATA.V_VAL FROM DOMAINS_DATA " & _
'                "WHERE D_KEY = 7"
'
'    'query the database
'    Set qdFos = gdbsOld.CreateQueryDef(sEmpty, strSQLFos & ";")
'
'    'Create a temporary snapshot-type Recordset.
'    Set rstFos = qdFos.OpenRecordset(dbOpenSnapshot)
'
'    With rstFos
'        .MoveFirst
'        nn& = 0
'        Do Until .EOF
'           If val(rstFos!V_KEY) > nn& Then
'              nn& = rstFos!V_KEY
'              ReDim Preserve arrN07(nn&)
'              numN07 = nn&
'              End If
'           arrN07(val(rstFos!V_KEY) - 1) = Trim$(rstFos!V_VAL)
'
'          .MoveNext
'        Loop
'    End With
'    rstFos.Close
'
'    '----------------now load up formation names
'    strSQLFos = "SELECT DOMAINS_DATA.V_KEY, DOMAINS_DATA.V_VAL FROM DOMAINS_DATA " & _
'                "WHERE D_KEY = 4"
'
'    'query the database
'    Set qdFos = gdbsOld.CreateQueryDef(sEmpty, strSQLFos & ";")
'
'    'Create a temporary snapshot-type Recordset.
'    Set rstFos = qdFos.OpenRecordset(dbOpenSnapshot)
'
'    With rstFos
'        .MoveFirst
'        nn& = 0
'        Do Until .EOF
'           If val(rstFos!V_KEY) > nn& Then
'              nn& = rstFos!V_KEY
'              ReDim Preserve arrN10(nn&)
'              numN10 = nn&
'              End If
'
'           arrN10(val(rstFos!V_KEY) - 1) = Trim$(rstFos!V_VAL)
'           .MoveNext
'        Loop
'    End With
'    rstFos.Close
'
'    '------------Fill names array----------------------------------
'    strSQLFos = "SELECT DOMAINS_DATA.V_KEY, DOMAINS_DATA.V_VAL FROM DOMAINS_DATA " & _
'                "WHERE D_KEY = 5"
'
'    'query the database
'    Set qdFos = gdbsOld.CreateQueryDef(sEmpty, strSQLFos & ";")
'
'    'Create a temporary snapshot-type Recordset.
'    Set rstFos = qdFos.OpenRecordset(dbOpenSnapshot)
'
'    With rstFos
'        .MoveLast
'        numfound& = .RecordCount
'        .MoveFirst
'        nn& = 0
'        Do Until .EOF
'           If val(rstFos!V_KEY) > nn& Then
'              nn& = rstFos!V_KEY
'              ReDim Preserve arrN11(nn&)
'              numN11 = nn&
'              End If
'           arrN11(val(rstFos!V_KEY) - 1) = Trim$(rstFos!V_VAL)
'           .MoveNext
'        Loop
'    End With
'    rstFos.Close
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'
'errhand:
'   Select Case Err.Number
'       Case 3021
'          Resume Next
'       Case Else
'          Screen.MousePointer = vbDefault
'          MsgBox "Encountered error #: " & Err.Number & vbLf & _
'               Err.Description & vbLf & _
'               sEmpty, vbCritical + vbOKOnly, "MapDigitizer"
'   End Select
'End Sub
'---------------------------------------------------------------------------------------
' Procedure : UpdateStatus
' Author    : Dr-John-K-Hall
' Date      : 2/18/2015
' Purpose   : Updates Status of fancy progress bar
'---------------------------------------------------------------------------------------
'
Public Sub UpdateStatus(Form1 As Form, ShowStatusProgress As Boolean, FileBytes As Long)
'--------------------------------------------------------------------
' This routine generates the picProgBar fancy progress bar on Form
'--------------------------------------------------------------------

    Dim progress As Long
    Const SRCCOPY = &HCC0020
    Dim Txt$

   On Error GoTo UpdateStatus_Error

   If pbScaleWidth = 0 Then pbScaleWidth = 100 '0 to 100% is the default

   If FileBytes = oldFileBytes And FileBytes <> 0 Then
      Exit Sub 'just old value, don't repaint
   Else
      oldFileBytes = FileBytes
      End If

    With Form1

        BringWindowToTop (Form1.hwnd)
        .picProgBar.Visible = True

        If FileBytes > pbScaleWidth Then
           progress = pbScaleWidth
           .picProgBar.Visible = False
        Else
           progress = FileBytes
           End If


        Txt$ = Format$(CLng((progress / pbScaleWidth) * 100)) + "%..."

        If ShowStatusProgress Then
           GDMDIform.StatusBar1.Panels(2).Text = Format$(CLng((progress / pbScaleWidth) * 100)) + "%..."
           End If

        .picProgBar.Cls
        .picProgBar.ScaleWidth = pbScaleWidth
        .picProgBar.CurrentX = (pbScaleWidth - .picProgBar.TextWidth(Txt$)) \ 2
        .picProgBar.CurrentY = (.picProgBar.ScaleHeight - .picProgBar.TextHeight(Txt$)) \ 2
        .picProgBar.Print Txt$
        .picProgBar.Line (0, 0)-(progress, .picProgBar.ScaleHeight), .picProgBar.ForeColor, BF
        R = BitBlt(.picProgBar.hdc, 0, 0, pbScaleWidth, .picProgBar.ScaleHeight, .picProgBar.hdc, 0, 0, SRCCOPY)
        .picProgBar.Refresh

    End With

   On Error GoTo 0
   Exit Sub

UpdateStatus_Error:

    Resume Next

End Sub
'
'
'
''---------------------------------------------------------------------------------------
'' Procedure : LoadTreeView
'' DateTime  : 11/23/2002 18:15
'' Purpose   : Loads up the tree view control of the Detail Report form
''             for search report record no. = ReportNum
''---------------------------------------------------------------------------------------
''
'Sub LoadTreeView(ReportNum&)
'
'  On Error GoTo errhand
'
'    Dim mNode As Node
'    Dim intIndex As Integer
'
'    'reload treeview1
'    GDDetailReportfrm.TreeView1.Nodes.Clear
'
'    GDDetailReportfrm.TreeView1.Sorted = True
'    Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add()
'
'    'order number
'    OrderNums$ = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(4)
'
'    'Title node
'    mNode.Text = "Specimen (# " & Trim$(OrderNums$) & ")"
'    mNode.Tag = "Specimen Name"
'    mNode.Key = "Root"
'    mNode.Image = "specimen"
'
'    'Child nodes
'    Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(1, tvwChild)
'    mNode.Text = "Categories"
'    mNode.Tag = "categories"
'    mNode.Key = "Child1"
'    mNode.Image = "categories"
'    intIndex = mNode.Index
'
'    'Second child node
'    'determine how many sub childs
'    intIndex = mNode.Index
'    If InStr(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3), "conod") <> 0 Then
'       Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex, tvwChild)
'       mNode.Text = "Conodonta"
'       mNode.Tag = "conod fossil"
'      mNode.Key = "conod"
'       mNode.Image = "cono"
'          'add fossil information
'           intIndex2 = mNode.Index
'           Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex2, tvwChild)
'           mNode.Text = "fossils"
'           mNode.Tag = "conod fossil tag"
'           mNode.Key = "conod fossil key"
'           mNode.Image = "fossil"
'       End If
'    If InStr(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3), "diato") <> 0 Then
'       Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex, tvwChild)
'       mNode.Text = "Diatom"
'       mNode.Tag = "diatom fossil"
'       mNode.Key = "diato"
'       mNode.Image = "diatom"
'          'add fossil information
'           intIndex2 = mNode.Index
'           Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex2, tvwChild)
'           mNode.Text = "fossils"
'           mNode.Tag = "diatom fossil tag"
'           mNode.Key = "diatom fossil key"
'           mNode.Image = "fossil"
'       End If
'    If InStr(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3), "foram") <> 0 Then
'       Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex, tvwChild)
'       mNode.Text = "Foraminifera"
'       mNode.Tag = "foram fossil"
'       mNode.Key = "foram"
'       mNode.Image = "foram"
'          'add fossil information
'           intIndex2 = mNode.Index
'           Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex2, tvwChild)
'           mNode.Text = "fossils"
'           mNode.Tag = "foram fossil tag"
'           mNode.Key = "foram fossil key"
'           mNode.Image = "fossil"
'       End If
'    If InStr(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3), "megaf") <> 0 Then
'       Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex, tvwChild)
'       mNode.Text = "Megafauna"
'       mNode.Tag = "megaf fossil"
'       mNode.Key = "megaf"
'       mNode.Image = "mega"
'          'add fossil information
'           intIndex2 = mNode.Index
'           Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex2, tvwChild)
'           mNode.Text = "fossils"
'           mNode.Tag = "mega fossil tag"
'           mNode.Key = "mega fossil key"
'           mNode.Image = "fossil"
'       End If
'    If InStr(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3), "nan") <> 0 Then
'       Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex, tvwChild)
'       mNode.Text = "Nannoplankton"
'       mNode.Tag = "nano fossil"
'       mNode.Key = "nanno"
'       mNode.Image = "nano"
'          'add fossil information
'           intIndex2 = mNode.Index
'           Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex2, tvwChild)
'           mNode.Text = "fossils"
'           mNode.Tag = "nano fossil tag"
'           mNode.Key = "nano fossil key"
'           mNode.Image = "fossil"
'       End If
'    If InStr(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3), "ostra") <> 0 Then
'       Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex, tvwChild)
'       mNode.Text = "Ostracoda"
'       mNode.Tag = "ostra fossil"
'       mNode.Key = "ostra"
'       mNode.Image = "ostra"
'          'add fossil information
'           intIndex2 = mNode.Index
'           Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex2, tvwChild)
'           mNode.Text = "fossils"
'           mNode.Tag = "ostra fossil tag"
'           mNode.Key = "ostra fossil key"
'           mNode.Image = "fossil"
'       End If
'    If InStr(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3), "palyn") <> 0 Then
'       Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex, tvwChild)
'       mNode.Text = "Palynology"
'       mNode.Tag = "palyn fossil"
'       mNode.Key = "palyn"
'       mNode.Image = "paly"
'          'add fossil information
'           intIndex2 = mNode.Index
'           Set mNode = GDDetailReportfrm.TreeView1.Nodes.Add(intIndex2, tvwChild)
'           mNode.Text = "fossils"
'           mNode.Tag = "paly fossil tag"
'           mNode.Key = "paly fossil key"
'           mNode.Image = "fossil"
'       End If
'
'   Exit Sub
'
'errhand:
'
'    Screen.MousePointer = vbDefault
'
'    MsgBox "Encountered error #: " & Err.Number & vbLf & _
'           Err.Description & vbLf & _
'           "You probably won't be able to obtain a complete detailed report!" & vbLf & _
'           sEmpty, vbExclamation + vbOKOnly, "MapDigitizer"
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : LoadDetailedReportListView
'' DateTime  : 11/23/2002 18:43
'' Purpose   : Load ListView form of DetailedReport Form
''---------------------------------------------------------------------------------------
''
'Sub LoadDetailedReportListView(ReportNum&)
'
'    On Error GoTo errhand
'
'    If InStr(GDDetailReportfrm.lvwDetailReport.ColumnHeaders(1), "Place Name") = 0 And _
'       InStr(GDDetailReportfrm.lvwDetailReport.ColumnHeaders(1), "Well Name") = 0 Then
'       LoadDefaultDetailReportInfo 'load up default column headers
'       End If
'
'    Set mitem = GDDetailReportfrm.lvwDetailReport.ListItems.Add()
'    mitem.Text = GDReportfrm.lvwReport.ListItems(ReportNum&).Text
'    GDMDIform.StatusBar1.Panels(2) = "Record #: " & Trim$(str(NearestPnt&))
'    GDDetailReportfrm.lvwDetailReport.LabelEdit = lvwManual
'
'    'find the icon associated with this record
'    doclin$ = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3)
'    IconNum& = 0
'    If InStr(doclin$, "cono") Then
'       IconNum& = 1
'       End If
'
'    If InStr(doclin$, "diato") Then
'       If IconNum& <> 0 Then
'          IconNum& = -1
'          GoTo 50
'       Else
'          IconNum& = 2
'          End If
'       End If
'
'    If InStr(doclin$, "foram") Then
'       If IconNum& <> 0 Then
'          IconNum& = -1
'          GoTo 50
'       Else
'          IconNum& = 3
'          End If
'       End If
'
'    If InStr(doclin$, "mega") Then
'       If IconNum& <> 0 Then
'          IconNum& = -1
'          GoTo 50
'       Else
'          IconNum& = 4
'          End If
'       End If
'
'    If InStr(doclin$, "nan") Then
'       If IconNum& <> 0 Then
'          IconNum& = -1
'          GoTo 50
'       Else
'          IconNum& = 5
'          End If
'       End If
'
'    If InStr(doclin$, "ostra") Then
'       If IconNum& <> 0 Then
'          IconNum& = -1
'          GoTo 50
'       Else
'          IconNum& = 6
'          End If
'       End If
'
'    If InStr(doclin$, "palyn") Then
'       If IconNum& <> 0 Then
'          IconNum& = -1
'          GoTo 50
'       Else
'          IconNum& = 7
'          End If
'       End If
'
'50      Select Case IconNum&
'       Case -1
'          mitem.SmallIcon = "multi"
'       Case 0
'          mitem.SmallIcon = "blank"
'       Case 1
'          mitem.SmallIcon = "cono"
'       Case 2
'          mitem.SmallIcon = "diatom"
'       Case 3
'          mitem.SmallIcon = "foram"
'       Case 4
'          mitem.SmallIcon = "mega"
'       Case 5
'          mitem.SmallIcon = "nano"
'       Case 6
'          mitem.SmallIcon = "ostra"
'       Case 7
'          mitem.SmallIcon = "paly"
'       Case Else
'    End Select
'
'    mitem.SubItems(1) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(1)
'    mitem.SubItems(2) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(2)
'    mitem.SubItems(3) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(3)
'    mitem.SubItems(4) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(4)
'    mitem.SubItems(5) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(5)
'    If Not IsNull(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(6)) Then
'       mitem.SubItems(6) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(6)
'       End If
'    If Not IsNull(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(7)) Then
'       mitem.SubItems(7) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(7)
'       End If
'    If Not IsNull(GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(8)) Then
'       mitem.SubItems(8) = GDReportfrm.lvwReport.ListItems(ReportNum&).SubItems(8)
'       End If
'
'    Exit Sub
'
'errhand:
'    Screen.MousePointer = vbDefault
'    MsgBox "Encountered error #: " & Err.Number & vbLf & _
'         Err.Description & vbLf & _
'         "in module LoadDetailedReportListView" & vbLf & _
'         "You probably won't be able to get a full Detailed Report.", _
'         vbCritical + vbOKOnly, "MapDigitizer"
'
'End Sub
'Sub LoadDefaultDetailReportInfo()
'    'load default list-view column names for detailed report form
'
'    On Error GoTo errhand:
'
'    With GDDetailReportfrm
'
'     .lvwDetailReport.ListItems.Clear
'     .lvwDetailReport.ColumnHeaders.Clear
'     'set up headers for List View
'     If GDReportfrm.lvwReport.ColumnHeaders(1).Text = "Well Name" Then
'        .lvwDetailReport.ColumnHeaders.Add , , "Well Name", 1500
'     Else
'        .lvwDetailReport.ColumnHeaders.Add , , "Place Name", 1500
'        End If
'     .lvwDetailReport.ColumnHeaders.Add , , "ITMx", 1000
'     .lvwDetailReport.ColumnHeaders.Add , , "ITMy", 1000
'     .lvwDetailReport.ColumnHeaders.Add , , "Fossils (Display Number)", 2000
'     .lvwDetailReport.ColumnHeaders.Add , , "Order Number", 1200
'     .lvwDetailReport.ColumnHeaders.Add , , "Formation", 1500
'     .lvwDetailReport.ColumnHeaders.Add , , "Mean Depth (m)", 1350
'     .lvwDetailReport.ColumnHeaders.Add , , "Gnd Level (m)", 1200
'     .lvwDetailReport.ColumnHeaders.Add , , "Z (meter)", 900
'
'     .lvwDetailReport.ColumnHeaders(1).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(2).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(3).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(4).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(5).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(6).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(7).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(8).Alignment = lvwColumnLeft
'     .lvwDetailReport.ColumnHeaders(9).Alignment = lvwColumnLeft
'
'    'don't allow the user to edit the column names
'    .lvwDetailReport.LabelEdit = lvwManual
'
'   End With
'
'   Exit Sub
'
'errhand:
'    Screen.MousePointer = vbDefault
'    MsgBox "Encountered error #: " & Err.Number & vbLf & _
'         Err.Description & vbLf & _
'         "in module LoadDefaultDetailReportInfo" & vbLf & _
'         "You probably won't get a complete detailed report.", _
'         vbCritical + vbOKOnly, "MapDigitizer"
'
'End Sub
'---------------------------------------------------------------------------------------
' Procedure : CheckDirect2
' DateTime  : 12/1/2002 21:48
' Purpose   : checks if mainpath$ exists and can be written to
'           : in case of protection error, use App.path instead
'---------------------------------------------------------------------------------------
'
Sub CheckDirect2(mainpath$)
   On Error GoTo errhand

   'attempt to write and erase on mainpath$
50
   filtmp% = FreeFile
   Open mainpath$ & "\check.tmp" For Output As #filtmp%
   Print #filtmp%, "Check i/o"
   Close #filtmp%

   Kill mainpath$ & "\check.tmp"

   Exit Sub

errhand:

   Screen.MousePointer = vbDefault

   If Err.Number = 75 Then
      mainpath$ = App.Path
      GoTo 50
      End If

   'if got here, then something is wrong, so try next letter
   Close
   mainpath$ = Chr$(Asc(mainpath$) + 1) & ":"
   GoTo 50

End Sub
'
'
'Sub Createdb1FossilSql()
'   'form the BOOLEAN search string over fossil tables
'   'of the active database
'
'    With GDSearchfrm
'
'         strSql1 = gsEmpty
'
'         andstr$ = " AND "
'         If .chkConodonta.Value = vbChecked Then
'            If strSql1 = gsEmpty Then
'               strNot = sEmpty
'               If .Combo1.Text = "AND NOT" Then strNot = "NOT "
'               strSql1 = andstr$ & "(" & strNot & "Order_new![conod] <> 0 "
'            End If
'         End If
'         If .chkDiatom.Value = vbChecked Then
'            If strSql1 = gsEmpty Then
'               strNot = sEmpty
'               If .Combo2.Text = "AND NOT" Then strNot = "NOT "
'               strSql1 = andstr$ & "(" & strNot & "Order_new![diato] <> 0 "
'            Else
'               strSql1 = strSql1 & .Combo2.Text & " Order_new![diato] <> 0 "
'               End If
'         End If
'         If .chkForaminifera.Value = vbChecked Then
'            If strSql1 = gsEmpty Then
'               strNot = sEmpty
'               If .Combo3.Text = "AND NOT" Then strNot = "NOT "
'               strSql1 = andstr$ & "(" & strNot & "Order_new![foram] <> 0 "
'            Else
'               strSql1 = strSql1 & .Combo3.Text & " Order_new![foram] <> 0 "
'               End If
'         End If
'         If .chkMegafauna.Value = vbChecked Then
'            If strSql1 = gsEmpty Then
'               strNot = sEmpty
'               If .Combo4.Text = "AND NOT" Then strNot = "NOT "
'               strSql1 = andstr$ & "(" & strNot & "Order_new![megaf] <> 0 "
'            Else
'               strSql1 = strSql1 & .Combo4.Text & " Order_new![megaf] <> 0 "
'               End If
'         End If
'         If .chkNanoplankton.Value = vbChecked Then
'            If strSql1 = gsEmpty Then
'               strNot = sEmpty
'               If .Combo5.Text = "AND NOT" Then strNot = "NOT "
'               strSql1 = andstr$ & "(" & strNot & "Order_new![nano] <> 0 "
'            Else
'               strSql1 = strSql1 & .Combo5.Text & " Order_new![nano] <> 0 "
'               End If
'         End If
'         If .chkOstracoda.Value = vbChecked Then
'            If strSql1 = gsEmpty Then
'               strNot = sEmpty
'               If .Combo6.Text = "AND NOT" Then strNot = "NOT "
'               strSql1 = andstr$ & "(" & strNot & "Order_new![ostra] <> 0 "
'            Else
'               strSql1 = strSql1 & .Combo6.Text & " Order_new![ostra] <> 0 "
'               End If
'         End If
'         If .chkPalynology.Value = vbChecked Then
'            If strSql1 = gsEmpty Then
'               strNot = sEmpty
'               If .Combo7.Text = "AND NOT" Then strNot = "NOT "
'               strSql1 = andstr$ & "(" & strNot & "Order_new![palin] <> 0 "
'            Else
'               strSql1 = strSql1 & .Combo7.Text & " Order_new![palin] <> 0 "
'               End If
'         End If
'         If .chkShekef.Value = vbChecked And _
'            .Combo15.Text = "AND" Then
'            strSql1 = gsEmpty 'AND searches over SHEKEF turns off searches of active database
'            End If
'
'         If strSql1 <> gsEmpty Then
'            strSql1 = strSql1 & ") "
'         Else
'            End If
'
'    End With
'
'End Sub
'Sub Createdb2SqlFossil()
'   'form the BOOLEAN search string over fossil tables
'   'of the scanned database
'
'   strSqlCategory = sEmpty
'
'   With GDSearchfrm
'
'      'if any of the fossil zones or fossil names are included, then
'      'don't search over the scanned database for that fossil since
'      'the scanned database doesn't contain any zone or fossil names
'
'      ' -------------------Define a SQL statement----
'        andstr$ = "AND "
'        If .chkConodonta.Value = vbChecked And _
'           .chkActCono.Value = vbUnchecked And _
'           .chkDicCono.Value = vbUnchecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo1.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 8 "
'           End If
'        End If
'        If .chkDiatom.Value = vbChecked And _
'           .chkActDiatom = vbUnchecked And _
'           .chkDicDiatom = vbUnchecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo2.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 7 "
'           Else
'              strSqlCategory = strSqlCategory & .Combo2.Text & " OBJECTS2![N02] = 7 "
'              End If
'        End If
'        If .chkForaminifera.Value = vbChecked And _
'           .chkActForam.Value = vbUnchecked And _
'           .chkDicForam.Value = vbUnchecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo3.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 1 "
'           Else
'              strSqlCategory = strSqlCategory & .Combo3.Text & " OBJECTS2![N02] = 1 "
'              End If
'        End If
'        If .chkShekef.Value = vbChecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo15.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 2 "
'           Else
'              strSqlCategory = strSqlCategory & .Combo15.Text & " OBJECTS2![N02] = 2 "
'              End If
'        End If
'        If .chkMegafauna.Value = vbChecked And _
'           .chkActMega.Value = vbUnchecked And _
'           .chkDicMega.Value = vbUnchecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo4.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 5 "
'           Else
'              strSqlCategory = strSqlCategory & .Combo4.Text & " OBJECTS2![N02] = 5 "
'              End If
'        End If
'        If .chkNanoplankton.Value = vbChecked And _
'           .chkActNano.Value = vbUnchecked And _
'           .chkDicNano.Value = vbUnchecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo5.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 6 "
'           Else
'              strSqlCategory = strSqlCategory & .Combo5.Text & " OBJECTS2![N02] = 6 "
'              End If
'        End If
'        If .chkOstracoda.Value = vbChecked And _
'           .chkActOstra.Value = vbUnchecked And _
'           .chkDicOstra.Value = vbUnchecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo6.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 3 "
'           Else
'              strSqlCategory = strSqlCategory & .Combo6.Text & " OBJECTS2![N02] = 3 "
'              End If
'        End If
'        If .chkPalynology.Value = vbChecked And _
'           .chkActPaly.Value = vbUnchecked And _
'           .chkDicPaly.Value = vbUnchecked Then
'           If strSqlCategory = gsEmpty Then
'              strNot = sEmpty
'              If .Combo7.Text = "AND NOT" Then strNot = "NOT "
'              strSqlCategory = andstr$ & "(" & strNot & "OBJECTS2![N02] = 4 "
'           Else
'              strSqlCategory = strSqlCategory & .Combo7.Text & " OBJECTS2![N02] = 4 "
'              End If
'        End If
'        If strSqlCategory <> gsEmpty Then
'           strSqlCategory = strSqlCategory & ") "
'           End If
'
'   End With
'
'End Sub
'Sub CheckFossilSQLSyntax(ier%)
'   'check the User edited Fossil Types SQL syntax
'   'for the beginning "AND" clause and
'   'for equal numbers of left and right parenthesis
'
'   ier% = 0
'
'   With GDSearchfrm
'
'     Dim numLeft&, numRight&, ch$
'
'     'check for equal numbers of left and right parenthesis
'     numLeft& = 0
'     numRight& = 0
'     If .txtSQLdb1.Enabled = True And .txtSQLdb1.Visible = True Then
'        For i% = 1 To Len(.txtSQLdb1.Text)
'           ch$ = Mid$(.txtSQLdb1.Text, i%, 1)
'           If ch$ = "(" Then numLeft& = numLeft& + 1
'           If ch$ = ")" Then numRight& = numRight& + 1
'        Next i%
'        End If
'     If numLeft& <> numRight& Then
'        ier% = -1
'     Else
'        strSql1 = Trim$(.txtSQLdb1.Text)
'        End If
'
'     numLeft& = 0
'     numRight& = 0
'     If .txtSQLdb2.Enabled = True And .txtSQLdb2.Visible = True Then
'        For i% = 1 To Len(.txtSQLdb2.Text)
'           ch$ = Mid$(.txtSQLdb2.Text, i%, 1)
'           If ch$ = "(" Then numLeft& = numLeft& + 1
'           If ch$ = ")" Then numRight& = numRight& + 1
'        Next i%
'        End If
'     If numLeft& <> numRight& Then
'        ier% = ier% - 2
'     Else
'        strSqlCategory = .txtSQLdb2.Text
'        End If
'
'     If ier% <> 0 Then Exit Sub 'skip any other error checking
'
'     'check for "AND " clause at beginning of SQL
'     If .txtSQLdb1.Enabled = True And .txtSQLdb1.Visible = True Then
'        If Trim$(.txtSQLdb1.Text) <> sEmpty Then
'           If Mid$(Trim$(.txtSQLdb1.Text), 1, 4) <> "AND " Then
'              ier% = -4
'              End If
'           End If
'        End If
'
'     If .txtSQLdb2.Enabled = True And .txtSQLdb2.Visible = True Then
'        If Trim$(.txtSQLdb2.Text) <> sEmpty Then
'           If Mid$(Trim$(.txtSQLdb2.Text), 1, 4) <> "AND " Then
'              ier% = ier% - 5
'              End If
'           End If
'        End If
'
'   End With
'
'End Sub


Function InStrR(st1 As String, st2 As String) As Integer
   'works just like InStr, but if one string is null and
   'the second is not, then returns 0
   If st1 = sEmpty And st2 <> sEmpty Then
      InStrR = 0
   ElseIf st1 <> sEmpty And st2 = sEmpty Then
      InStrR = 0
   ElseIf st1 = sEmpty And st2 = sEmpty Then
      InStrR = 1
   Else
      InStrR = InStr(st1, st2)
      End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : LocatePointOnMap
' DateTime  : 3/30/2005 15:58
' Author    : Chaim Keller
' Purpose   : locates points on map
'           (mode=0 entry from GDReportfrm locate button)
'           (mode=1 entry from right click on GDReportfrm)
'---------------------------------------------------------------------------------------
'
Sub LocatePointOnMap(modeL%)
    
   On Error GoTo LocatePointOnMap_Error

    Screen.MousePointer = vbHourglass
    
    GDMDIform.StatusBar1.Panels(1) = sEmpty
    GDMDIform.StatusBar1.Panels(2) = sEmpty
    
    'if map is not visible, then make geo map visible
    init& = 0
    If Not GeoMap And Not TopoMap Then
       'display the geo map
        myfile = Dir(picnam$)
        If myfile = sEmpty Or Trim$(picnam$) = sEmpty Then
           response = MsgBox("Can't find map!" & vbLf & _
                      "Use the Files/Geologic map options menu to help find it.", _
                      vbExclamation + vbOKOnly, "GSIDB")
           'take further response
            GeoMap = False
            Exit Sub
        Else
            With GDMDIform
                
'                Screen.MousePointer = vbHourglass
'                buttonstate&(3) = 0
'                .Toolbar1.Buttons(3).Value = tbrUnpressed
'                For i& = 4 To 7
'                  .Toolbar1.Buttons(i&).Enabled = False
'                Next i&
'                .Toolbar1.Buttons(9).Enabled = False
'                If buttonstate&(15) = 1 Then 'search still activated
'                   .Toolbar1.Buttons(15).Value = tbrPressed
'                   End If
'
'                .mnuGeo.Enabled = False 'disenable menu of geo. coordinates display
'                .Toolbar1.Buttons(2).Value = tbrPressed
'                If topos = True Then .Toolbar1.Buttons(3).Enabled = True
'                buttonstate&(2) = 1
'                .Toolbar1.Buttons(8).Enabled = True
'                .Toolbar1.Buttons(10).Enabled = True
'                .mnuPrintMap.Enabled = True
'                .Label1 = lblX
'                .Label5 = lblX
'                .Label2 = LblY
'                .Label6 = LblY
                
                'now redepress or undepress the buttons, and refresh
                For i& = 1 To .Toolbar1.Buttons.count
                   If buttonstate&(i&) = 1 Then
                      .Toolbar1.Buttons(i&).value = tbrPressed
                   Else
                      .Toolbar1.Buttons(i&).value = tbrUnpressed
                      End If
                Next i&
                .Toolbar1.Refresh 'refresh the visual state of the toolbar
                  
                'load up Geo map
                Call ShowGeoMap(0)
                init& = 1
            
            End With
            
            End If
            
       End If
       
    If StopPlotting Then GoTo L100
       
    'first turn off blinker
    GDMDIform.CenterPointTimer.Enabled = False
    
    'reset the blinker flag
    CenterBlinkState = False
    
    'refresh the maps (to erase any old searchpoints)
'    GDform1.Picture2.Picture = LoadPicture(picnam$)
    
    If init& = 0 Then
       CheckDuplicatePoints = False 'havn't checked yet for duplicate points
       End If
    
    If modeL% = 0 Then
        'plot new search points (entry from Locate Points on Map button on GDReportfrm)
        PlotNewSearchPoints
        If StopPlotting Then 'user stopped plotting
           StopPlotting = False 'reset flag
           Exit Sub
           End If
        
        If NumReportPnts& = 0 Then
           Screen.MousePointer = vbDefault
           MsgBox "You didn't select (highlight) any points in the report for plotting!", vbExclamation + vbOKOnly, "MapDigitizer"
           Exit Sub
           End If
        
        'move geo maps to the first highlighted search point
        'that's within the geologic map boundary
        If GeoMap Then MoveToFirstLocatedPoint
    
    ElseIf modeL% = 1 Then 'right clicked on GDReportfrm
    
        'restore marks on map
        PlotNewSearchPoints
       
       'move to position of record that right clicked on GDReportfrm record
        XPnt = val(GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(2))
        YPnt = val(GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(3))
        If Not DigiRubberSheeting Then
           GDMDIform.Text5 = XPnt
           GDMDIform.Text6 = YPnt
        Else
           GDMDIform.Text5 = GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(4)
           GDMDIform.Text6 = GDReportfrm.lvwReport.ListItems(NewHighlighted&).SubItems(5)
           End If
        Call gotocoord
'        Call ShiftMap(CSng(XPnt), CSng(YPnt))
        End If

    'blink center marker
'    GDMDIform.CenterPointTimer.Enabled = True
    
    'now replot center point at first plot point
    Dim x As Single, Y As Single
'    ITMx0 = ((val(GDMDIform.Text5) - ULGeoX) / (LRGeoX - ULGeoX)) * pixwi
'    ITMy0 = ((ULGeoY - val(GDMDIform.Text6)) / (ULGeoY - LRGeoY)) * pixhi
'    X = ITMx0 * twipsx
'    Y = ITMy0 * twipsy
'    xo_blank_mark = X: Yo_blank_mark = Y 'record current map position
       
    'record highlighted points
    RecordHighlighted
     
    Ret = BringWindowToTop(GDform1.hwnd)
    
L100: StopPlotting = False
    
    Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Sub

LocatePointOnMap_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LocatePointOnMap of Module modGDModule"

End Sub
Sub RecordHighlighted()
      
'this routine records highlighted records to be used
'for restoring them after left or right clicking

'---------------------------------------------------
'shut off timers during this operation
ce& = 0
If GDMDIform.CenterPointTimer.Enabled = True Then
   ce& = 1
   GDMDIform.CenterPointTimer.Enabled = False
   End If
'----------------------------------------------------

        'left clicking will erase the highlighted points, so store
        'them in order to restore them after the mouse is released
        '(numHighlighted& = 0 check makes sure that even if the double click was
        'so slow that it registered as a single click, nevertheless the highlighted
        'points are not erased)
        ReDim Preserve Highlighted(numReport& - 1)
        For i& = 1 To numReport& 'search over all the search results
            If GDReportfrm.lvwReport.ListItems(i&).Selected Then
               Highlighted(i& - 1) = 1
               numHighlighted& = numHighlighted& + 1
            Else
               Highlighted(i& - 1) = 0
               End If
            DoEvents 'yield to windows messaging
        Next i&
            
        'add delay to rectify asynchronous processing of Window's messages
        If numReport& >= 100 Then
           addtime = 50 / numReport&
        Else
           addtime = 0.5
           End If
        waitime = Timer
        Do Until Timer > waitime + addtime
          DoEvents
        Loop
      
'---------------------------------------------------------------
'restore timers shut off during this operation
If GDMDIform.CenterPointTimer.Enabled = False And ce& = 1 Then
   ce& = 0
   GDMDIform.CenterPointTimer.Enabled = True
   End If
'---------------------------------------------------------------

End Sub



''---------------------------------------------------------------------------------------
'' Procedure : FindNumberRecords
'' DateTime  : 11/15/2008 20:32
'' Author    : Chaim Keller
'' Purpose   : Query Databases to find maximum number of records
''---------------------------------------------------------------------------------------
''
'Public Function FindNumberRecords(mode%) As Long
'
'   Dim sqOrder As String
'   Dim rstOrder As Recordset
'
'   On Error GoTo FindNumberRecords_Error
'
'   If (mode% = 0) Then 'find maximum number of records in active database
'
'      Screen.MousePointer = vbHourglass
'
'      sqOrder = "SELECT Order_new.order_id FROM Order_new "
'      Set rstOrder = gdbs.OpenRecordset(sqOrder, dbOpenSnapshot)
'      With rstOrder
'         .MoveLast
'         RecordNum& = rstOrder.RecordCount - 1
'      End With
'      rstOrder.Close
'
'      Screen.MousePointer = vbDefault
'
'      FindNumberRecords = RecordNum&
'      Exit Function
'
'   ElseIf (mode% = 1) Then 'find maximum number of records in scanned database
'
'      'query old database for number of records
'
'      Screen.MousePointer = vbHourglass
'
'      sqOrder = "SELECT MAX(OBJECTS2.O_KEY) AS [MAXIMUM_OKEY] FROM OBJECTS2 "
'      Set rstOrder = gdbsOld.OpenRecordset(sqOrder, dbOpenSnapshot)
'      With rstOrder
'        .MoveLast 'populate recordset
'        RecordNum& = val(rstOrder![MAXIMUM_OKEY])
'        'can also use: RecordNum& = rstOrder.Fields(0)
'      End With
'      rstOrder.Close
'
'
'      Screen.MousePointer = vbDefault
'
'      FindNumberRecords = RecordNum&
'      Exit Function
'      End If
'
'   On Error GoTo 0
'   Exit Function
'
'FindNumberRecords_Error:
'
'   Screen.MousePointer = vbDefault
'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FindNumberRecords of Module modGDModule"
'
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : SaveScannedFile
'' DateTime  : 11/15/2008 23:49
'' Author    : Chaim Keller
'' Purpose   : Save new scanned file by its Serial Number
''query OBJECTS2 to find maximum O_KEY
''then query OBJECTS_VER2 for maximum OV_KEY
''then write FileName$ to OBJECTS_VER2 with OV_KEY = maximum OV_KEY + 1, O_KEY + 1
''then write Serial_Number$ to OBJECTS2 with the new OV_KEY
''---------------------------------------------------------------------------------------
''
'Public Sub SaveScannedFile(Serial_Number$, FileName$, New_OKEY&, ier%)
'   Dim strSQLEdit As String
'   Dim rstEdit As Recordset
'   Dim maxID&, MaxOKEY&
'   Dim DontUpdateOV_KEY As Boolean
'
'   On Error GoTo SaveScannedFile_Error
'
'   ier% = 0
'   UpdateOV_KEY = True
'
'   'first determine if file was already added to database
'   If Old_OKEY& <> 0 Then
'
'      New_OKEY& = Old_OKEY&
'
'      UpdateOV_KEY = False 'just edit the old record if the user desires
'
'      Select Case MsgBox("A file with name: " & Chr$(34) & FileName$ & Chr$(34) _
'                  & " has already been added to database" & vbCrLf _
'                  & vbCrLf & "For your reference, its O_KEY = " & str$(New_OKEY&) & vbCrLf _
'                  & vbCrLf & "Do you want to overwirte?..." _
'                  , vbYesNo Or vbExclamation Or vbDefaultButton2, App.Title)
'          Case vbYes
'
'          Case vbNo
'              ier% = -2
'              Exit Sub
'      End Select
'
'   Else 'add new file to the scanned database
'
'     'first find if O_NAME is unique, if not warn the user
'     Call CheckIfSNExists(Serial_Number$, FileName2$)
'     If Old_OKEY& <> 0 Then 'serial number is not unique
'
'         Select Case MsgBox("A file with Serial Number: " & Chr$(34) & Serial_Number$ & Chr$(34) _
'                     & " has already been added to database" & vbCrLf _
'                     & vbCrLf & "For your reference, its O_KEY = " & str$(Old_OKEY&) & vbCrLf _
'                     & vbCrLf & "Its image filename is: " & FileName2$ & vbCrLf _
'                     & vbCrLf & "Do you want to view the this image file that was" & vbCrLf _
'                     & "previously entered in the database with this identical Serial Number?..." _
'                     , vbYesNo Or vbExclamation Or vbDefaultButton2, App.Title)
'             Case vbYes
'
'                  'display the duplicate record
'                  If Dir(tifDir$ & "\" & UCase$(FileName$)) <> sEmpty Then
'                     Shell (tifCommandLine$ & " " & tifDir$ & "\" & FileName2$)
'                  Else
'                     Call MsgBox("The path: " & tifDir$ & "\" & UCase$(FileName2$) & " was not found or is not accessible!" & vbLf & vbLf & _
'                                 "Check the defined path to the tif files in the options menu, and try again", vbExclamation + vbOKOnly, App.Title)
'                     Old_OKEY& = 0
'                     ier% = -4
'                     Exit Sub
'                     End If
'
'                  'now ask the user what he want's to do with this duplication after waiting a bit
'                  waitime = Timer
'                  Do Until Timer > waitime + 1.5
'                     DoEvents
'                  Loop
'                  Select Case MsgBox("You can take three possible actions to handle this duplication" _
'                                     & vbCrLf & "" _
'                                     & vbCrLf & "       Answer ""Yes"" -        add the duplicate serial number to the database" _
'                                     & vbCrLf & "                    ""No"" -         don't add the duplicate serial number, rather replace the" _
'                                     & vbCrLf & "                                       previous image file with the newer clearer image file" _
'                                     & vbCrLf & "                    ""Cancel"" -  don't do anything (default)" _
'                                     , vbYesNoCancel Or vbQuestion Or vbDefaultButton3 Or vbSystemModal, "")
'
'                     Case vbYes
'                        'go on to record the duplicate serial number
'
'                     Case vbNo
'                        Call ReplaceImageFile(Old_OKEY&, Files(iFlex - 1))
'                        Exit Sub
'
'                     Case vbCancel
'                        Old_OKEY& = 0
'                        ier% = -4
'                        Exit Sub
'
'                  End Select
'
'             Case vbNo
'                 ier% = -3
'                 Exit Sub
'         End Select
'
'        End If
'
'     'find the current maximum O_KEY
'     MaxOKEY& = FindNumberRecords(1)
'     New_OKEY& = MaxOKEY& + 1
'
'     'find maximum OV_KEY
'     UpdateOV_KEY = True
'     strSQLEdit = "SELECT Max(OV_KEY) AS [MAX OVKEY] FROM OBJECTS_VER2 "
'     Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenSnapshot)
'     With rstEdit
'        .MoveLast 'populate recordset
'        maxID& = val(rstEdit![MAX OVKEY]) 'current maximum OV_KEY
'     End With
'     rstEdit.Close
'
'     'now update the OBJECTS_VER2 table
'     strSQLEdit = "SELECT OBJECTS_VER2.* FROM OBJECTS_VER2 "
'
'     'Create a dynaset-type Recordset for editing.
'     Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenDynaset)
'
'     With rstEdit
'        .LockEdits = True 'pessimistic locking (upon onset of editing,
'                        'locks write permission to this user only)
'
'        .MoveLast 'move to last record
'        .AddNew 'add a new record, so add "one" to the current maximum indices
'        !OV_KEY = maxID& + 1
'        !O_KEY = New_OKEY&
'        !O_FILE = FileName$
'        .Update 'update the table
'        .Close
'
'     End With
'     End If
'
'   'now update the OBJECTS2 table
'
'   If UpdateOV_KEY Then
'      strSQLEdit = "SELECT OBJECTS2.* FROM OBJECTS2 "
'
'     'Create a dynaset-type Recordset for editing.
'      Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenDynaset)
'
'      With rstEdit
'
'         .LockEdits = True 'pessimistic locking (upon onset of editing,
'                        'locks write permission to this user only)
'
'         .MoveLast 'populate the recordset
'         .AddNew 'add a new record
'         !O_KEY = New_OKEY&
'         !A08 = "0"
'         !A09 = "0"
'         !N01 = 0
'         !N02 = 0
'         !N03 = 0
'         !N04 = 0
'         !N05 = 0
'         !N06 = 0
'         !N07 = 0
'         !N08 = 0
'         !N09 = 0
'         !N10 = 0
'         !N11 = 0
'         !GL = 0
'
'      End With
'
'   Else
'
'      strSQLEdit = "SELECT OBJECTS2.* FROM OBJECTS2 WHERE " & _
'                   "OBJECTS2.[O_KEY] = " & New_OKEY&
'
'     'Create a dynaset-type Recordset for editing.
'      Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenDynaset)
'
'      rstEdit.LockEdits = True 'pessimistic locking (upon onset of editing,
'                        'locks write permission to this user only)
'      rstEdit.MoveLast 'populate the recordset
'      rstEdit.Edit
'
'      End If
'
'   rstEdit!O_NAME = Serial_Number$
'   rstEdit!O_MODIFY = Now 'date and time mmodified
'
'   rstEdit.Update 'update the table
'   rstEdit.Close
'
'   Exit Sub
'
'   On Error GoTo 0
'   Exit Sub
'
'SaveScannedFile_Error:
'   If Err.Number = 3021 Then
'      numrecord& = 0
'      Resume Next
'      End If
'   Screen.MousePointer = vbDefault
'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveScannedFile of Module modGDModule"
'   ier% = -1
'
'End Sub

Public Sub ShortPath(sPath As String, MaxDirLen As Integer, sShortPath As String, sRealPath As String)
   'This routine finds abreviated path names to fit in
   'the plot buffer list box
   
      'find last "\"
      For i% = Len(sPath) To 1 Step -1
         If Mid$(sPath, i%, 1) = "\" Then
            sRealPath = Mid$(sPath, 1, i%)
            Exit For
            End If
      Next i%
      
      'now shorten the path if necessary
      'remove final "\"
      sRealPath = Mid$(sRealPath, 1, Len(sRealPath) - 1)
   
      If Len(sRealPath) > MaxDirLen Then 'try to find short version
         pos1% = InStr(sRealPath, "\")
         'find drive letter
         sDriveLetter = Mid$(sRealPath, 1, pos1%)
         'Now find abbreviated form of the path for
         'displaying in the list box.
         'Abbreviated version just contains most inner directory
         If pos1% <> 0 Then
            pos1% = InStr(pos1% + 1, sRealPath, "\")
            If pos1% <> 0 Then
               For i% = Len(sRealPath) To 1 Step -1
                  If Mid$(sRealPath, i%, 1) = "\" Then
                     sShortPath = sDriveLetter & "...\" & Mid$(sRealPath, i% + 1, Len(sRealPath) - i%) & "\"
                     Exit For
                     End If
               Next i%
            Else
               sShortPath = sDriveLetter & "...\"
               End If
            End If
      Else
         sShortPath = sRealPath & "\" 'put back final "\"
         End If

End Sub

''---------------------------------------------------------------------------------------
'' Procedure : CheckIfFNExists
'' DateTime  : 11/19/2008 20:01
'' Author    : Chaim Keller
'' Purpose   : Determines if form has already been added to database
''---------------------------------------------------------------------------------------
''
'Public Sub CheckIfFNExists(FileName$, SN$)
'
'   Dim strSQLEdit As String
'   Dim rstEdit As Recordset
'
'   On Error GoTo CheckIfFNExists_Error
'
'   strSQLEdit = "SELECT OBJECTS_VER2.* FROM OBJECTS_VER2 WHERE " & _
'                "OBJECTS_VER2.[O_FILE] = " & Chr$(34) & Trim$(FileName$) & Chr$(34)
'
'   'Create a snapshot-type Recordset.
'   Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenSnapshot)
'
'   rstEdit.MoveLast
'   numrecord& = rstEdit.RecordCount
'
'   If numrecord& > 0 Then 'found it
'      Old_OKEY& = rstEdit!O_KEY
'   Else
'      SN$ = sEmpty
'      End If
'
'   rstEdit.Close
'
'   If numrecord& > 0 Then
'      'query the OBJECTS2 table for the corresponding serial number
'
'      strSQLEdit = "SELECT OBJECTS2.* FROM OBJECTS2 WHERE " & _
'                   "OBJECTS2.[O_KEY] = " & str$(Old_OKEY&)
'
'      'Create a snapshot-type Recordset.
'      Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenSnapshot)
'
'      rstEdit.MoveLast
'      numrecord& = rstEdit.RecordCount
'
'      If numrecord& > 0 Then 'found it, so give warning and request action
'         SN$ = rstEdit!O_NAME
'         End If
'
'      rstEdit.Close
'
'      End If
'
'   On Error GoTo 0
'   Exit Sub
'
'CheckIfFNExists_Error:
'
'   Screen.MousePointer = vbDefault
'
'   If Err.Number = 3021 Then
'      numrecord& = 0
'      Resume Next
'      End If
'
'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckIfFNExists of Module modGDModule"
'
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : CheckIfSNExists
'' DateTime  : 11/19/2008 20:01
'' Author    : Chaim Keller
'' Purpose   : Determines if form with same O_NAME = serial number has already been added to database
''             if it already exists, then returns the name of the earlier image file, FN$
''---------------------------------------------------------------------------------------
''
'Public Sub CheckIfSNExists(SN$, FN$)
'
'   Dim strSQLEdit As String
'   Dim rstEdit As Recordset
'   Dim fnp As String
'
'   On Error GoTo CheckIfSNExists_Error
'
'   strSQLEdit = "SELECT OBJECTS2.* FROM OBJECTS2 WHERE " & _
'                "OBJECTS2.[O_NAME] = " & Chr$(34) & Trim$(SN$) & Chr$(34)
'
'   'Create a snapshot-type Recordset.
'   Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenSnapshot)
'
'   rstEdit.MoveLast
'   numrecord& = rstEdit.RecordCount
'
'   If numrecord& > 0 Then 'found it
'      Old_OKEY& = rstEdit!O_KEY
'   Else
'      Old_OKEY& = 0
'      End If
'
'   rstEdit.Close
'
'   'now find the duplicate image file name
'   strSQLEdit = "SELECT OBJECTS_VER2.* FROM OBJECTS_VER2 WHERE " & _
'                "OBJECTS_VER2.[O_KEY] = " & str$(Old_OKEY&)
'
'   'Create a snapshot-type Recordset.
'   Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenSnapshot)
'
'   rstEdit.MoveLast
'   numrecord& = rstEdit.RecordCount
'
'   If numrecord& > 0 Then 'found it
'
'      FN$ = rstEdit!O_FILE
'
'      'make sure that this is not the same filename that was already replaced
'      fnp = Files(iFlex - 1)
'      If InStr(fnp, tifDir$) Then 'string must be stripped off tif directory
'         strlen% = Len(tifDir$ & "\")
'         fnp$ = Mid$(fnp, strlen% + 1, Len(fnp) - strlen%)
'         End If
'
'      If InStr(FN$, fnp) Then 'its the same file!
'         Old_OKEY& = 0
'         End If
'
'      End If
'
'   rstEdit.Close
'
'
'   On Error GoTo 0
'   Exit Sub
'
'CheckIfSNExists_Error:
'
'   Screen.MousePointer = vbDefault
'
'   If Err.Number = 3021 Then
'      numrecord& = 0
'      Resume Next
'      End If
'
'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckIfSNExists of Module modGDModule"
'
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : ReplaceImageFile
'' DateTime  : 12/3/2008 21:23
'' Author    : Chaim Keller
'' Purpose   : Replaces the old tif image with the clearer newer jpg scanned image
''---------------------------------------------------------------------------------------
''
'Public Sub ReplaceImageFile(OKEYMod&, FN As String)
'
'
'   On Error GoTo ReplaceImageFile_Error
'
'
'   Dim strSQLEdit As String
'   Dim rstEdit As Recordset
'   Dim fnp As String
'
'   fnp = FN
'   If InStr(fnp, tifDir$) Then 'string must be stripped off tif directory
'      strlen% = Len(tifDir$ & "\")
'      fnp = Mid$(fnp, strlen% + 1, Len(fnp) - strlen%)
'      End If
'
'   strSQLEdit = "SELECT OBJECTS_VER2.* FROM OBJECTS_VER2 WHERE " & _
'                "OBJECTS_VER2.[O_KEY] = " & str$(OKEYMod&)
'
'   'Create a snapshot-type Recordset.
'   Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenDynaset)
'
'   With rstEdit
'
'      .LockEdits = True 'pessimistic locking (upon onset of editing,
'                        'locks write permission to this user only)
'      .MoveLast 'populate the recordset
'      .Edit
'
'      !O_FILE = fnp
'
'      .Update 'update the table
'      .Close
'
'   End With
'
'   'now update the modification date
'
'   strSQLEdit = "SELECT OBJECTS2.* FROM OBJECTS2 WHERE " & _
'                "OBJECTS2.[O_KEY] = " & str$(OKEYMod&)
'
'   'Create a snapshot-type Recordset.
'   Set rstEdit = gdbsOld.OpenRecordset(strSQLEdit, dbOpenDynaset)
'
'   With rstEdit
'
'      .LockEdits = True 'pessimistic locking (upon onset of editing,
'                        'locks write permission to this user only)
'      .MoveLast 'populate the recordset
'      .Edit
'
'      !O_MODIFY = Now
'
'      .Update 'update the table
'      .Close
'
'   End With
'
'   On Error GoTo 0
'   Exit Sub
'
'
'ReplaceImageFile_Error:
'
'   If Err.Number = 3021 Then
'      numrecord& = 0
'      Resume Next
'      End If
'
'   Screen.MousePointer = vbDefault
'
'   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReplaceImageFile of Module modGDModule"
'End Sub

'Epsilon Algorithm, Created by Simon Johnson
'Uses storing encrypted passwd's, or producing message digests.
Function Hash(ByVal Text As String) As String
    Dim Hashch$
    AA = 1
    For i = 1 To Len(Text)
        AA = Sqr(AA * i * Asc(Mid(Text, i, 1))) 'Numeric Hash
    Next i
    Rnd (-1)
    Randomize AA 'seed PRNG
    
    For i = 1 To 12
        Hashch$ = Chr(Int(Rnd * 256))
        Hashch$ = str$(Asc(Hashch$))
        Hashch$ = Mid$(Hashch$, Len(Hashch$), 1)
        Hash = Hash & Hashch$
    Next i
End Function
Function CheckPassword(PassWord As String) As Boolean

    'checkes if inputed password is correct (returns true if correct)
    
    Dim s As String
    Dim Cnt As Long
    Dim dl As Long
    Dim CurUser As String, CurUserPass$, CurUserPass1$
    
    Cnt = 199
    s = String$(200, 0)
    dl = GetUserName(s, Cnt)
    If dl <> 0 Then CurUser = left$(s, Cnt) Else CurUser = sEmpty
    
    'use default User Name in case User Name is NULL or can't be retrieved
    If CurUser = sEmpty Then CurUser = ADMIN_USERNAME
    
    'remove the null termination
    CurUser = Mid$(CurUser, 1, Len(CurUser) - 1)
    
    'now generate a 4-4-4 digit number based on the user name
    
    CurUserPass$ = Hash(CurUser)
    CurUserPass1$ = Hash(ADMIN_USERNAME) 'GSI_USER will work as a User Name for any computer
    
    If PassWord = CurUserPass$ Or PassWord = CurUserPass1$ Then CheckPassword = True
    If PassWord <> CurUserPass$ And PassWord <> CurUserPass1$ Then CheckPassword = False

End Function



'---------------------------------------------------------------------------------------
' Procedure : HebrewCheck
' Author    : Chaim Keller
' Date      : 7/24/2011
' Purpose   : Checks for Hebrew characters
'---------------------------------------------------------------------------------------
'
Public Function HebrewCheck(str As Variant) As Boolean
    
   On Error GoTo HebrewCheck_Error

   HebrewCheck = False
   
    For i& = 1 To Len(str)
       If Asc(Mid$(str, i&, 1)) >= 128 Then
          'Hebrew characters found. don't need to check anymore
          HebrewCheck = True
          Exit For
          End If
    Next i&

   On Error GoTo 0
   Exit Function

HebrewCheck_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HebrewCheck of Module modGDModule"
End Function
'the Julian day number ("n" of the Astronomical Almanac)
Function JulianDayNumber(hrsjd As Double, tdjd As Single, mday%, mon%, _
                        yrjd As Integer, yljd As Integer, dayyrjd As Integer) As Double

        Dim Y%, m%, yd As Integer
        
'       if want JulianDayNumber at 00:00 use hrsjd = 0
'       if want JulainDayNumber at 12:00 use hrsjd = 12.0
        
'        If datenowjd$ = sEmpty Then Exit Sub

'        'determine starting day number and type of year
'        yrjd = Year(Format(datenowjd$, "mm/dd/yyyy"))
'        mday% = Day(Format(datenowjd$, "mm/dd/yyyy"))
'        mon% = Month(Format(datenowjd$, "mm/dd/yyyy"))
        
        yljd = DaysinYear(yrjd) 'number of days in civil year (either 365 or 366)
        
        'day number, Meeus p. 65
        dayyrjd = DayNumber(yljd, mon%, mday%)
        
        UT = hrsjd - tdjd - dstflag% 'universal time
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'three different ways to calculate the Julain day number depending on the
        'year of the calculation
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If yrjd >= 1900 And yrjd < 2100 Then
           'use Meeus chapter 7 formula for dyfjd to speed up things a bit

           yfjd = 367 * CDbl(yrjd) - 7 * (yrjd + (mon% + 9) \ 12) \ 4 + 275 * mon% \ 9 + mday% - 730531.5

           'day number
           'dayyrjd = yfjd - (367 * Year(datenow$) - 7 * (Year(datenow$) + 10 \ 12) \ 4 + 275 * 1 \ 9 + 1 - 730531.5) + 1

           JulianDayNumber = yfjd + UT / 24#
           
           'jd# = yfjd + J2000

        ElseIf (yrjd > -4712 And yrjd < 1900) Or (yrjd >= 2100) Then
        
           'use Meeus formula 7.1

           m% = mon%
           Y% = yrjd

           If (m <= 2) Then
               Y% = Y% - 1
               m% = m% + 12
               End If

           'Determine whether date is in Julian or Gregorian calendar based on
           'canonical date of calendar reform.

           If ((yrjd < 1582) Or ((yrjd = 1582) And ((mon% < 9) Or (mon% = 9 And mday% < 5)))) Then
               b% = 0
           Else
               a% = (Y% \ 100)
               b% = 2 - a% + (a% \ 4)
               End If

           yfjd = Fix(365.25 * (Y% + 4716)) + Fix(30.6001 * (m% + 1)) + _
                        mday% + b% - 1524.5  'Julian date, "JD", "jd", etc

           JulianDayNumber = yfjd + UT / 24# - J2000   'day "number" from J2000: "n", "d", etc...

           'jd# = yfjd + J2000

        Else 'pretend that Gregorian calendar existed at these dates

           dayyrn = dayyrjd + UT / 24#

           yfjd = 0
           If yd < 0 Then
              For i% = 1995 To yrjd Step -1
                 yfjd = yfjd - DaysinYear(i%)
              Next i%
           ElseIf yd >= 0 Then
              For i% = 1996 To yrjd - 1 Step 1
                 yfjd = yfjd + DaysinYear(i%)
              Next i%
              End If
              
           JulianDayNumber = yfjd + dayyrn - 1462.5 'number of days from J2000 at UT hours
        
           End If

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
   
   'determines daynumber for any month = mon%, day = mday%
   'yljd = 365 for regular year, 366 for leap year
   'based on Meeus' formula, p. 65
   
    kk% = 2
    If yljd = 366 Then kk% = 1
    DayNumber = (275 * mon%) \ 9 - kk * ((mon% + 9) \ 12) + mday% - 30
   
      
   End Function

Public Function CheckIfAdmin()
  
  If Not IsUserAnAdministrator Then
  
     Select Case MsgBox("You don't seem to be running the program with administrator privilege." _
            & vbCrLf & "Administrator privilege is required for running this program properly." _
            & vbCrLf & vbCrLf & "To elevate the privilege, do the following after exiting from the program: " _
            & vbCrLf & "(1) Right click on the program icon that you use to run the program." _
            & vbCrLf & "(2) Left click on ''Properties''." _
            & vbCrLf & "(3) Click on the ''Advanced'' button." _
            & vbCrLf & "(4) Check the box next to ''Run as administrator''." _
            & vbCrLf & vbCrLf & "Exit the program to change the privilige (recommended)?", vbQuestion + vbYesNoCancel, "Administrative Privilege check")
            
         Case vbYes
            'exit from the program
            GDMDIform.UnloadAllForms (sEmpty)
            Set GDMDIform = Nothing
            End
            
         Case Else
         
      End Select
         
     End If

End Function

'checks whether user is running the code as administrator for Vista and above
'source: David's blog: http://www.davidmoore.info/2011/06/20/how-to-check-if-the-current-user-is-an-administrator-even-if-uac-is-on/

Public Function IsUserAnAdministrator() As Boolean

Dim result As Long
Dim hProcessID As Long
Dim hToken As Long
Dim lReturnLength As Long
Dim tokenElevationType As Long

On Error GoTo IsUserAnAdministratorError

IsUserAnAdministrator = False

If IsUserAnAdmin() Then 'running as administrator
    IsUserAnAdministrator = True
    Exit Function
    End If

' If we’re on Vista onwards, check for UAC elevation token
' as we may be an admin but we’re not elevated yet, so the
' IsUserAnAdmin() function will return false

Dim myOS As OSVERSIONINFOEX
myOS.dwOSVersionInfoSize = Len(myOS)
GetVersionEx myOS

If myOS.dwPlatformId <> VER_PLATFORM_WIN32_NT Or myOS.dwMajorVersion < 6 Then
'   If the user is not on Vista or greater, then there’s no UAC, so don’t bother checking.
    Exit Function
    End If

' We need to get the token for the current process
hProcessID = GetCurrentProcess()

If hProcessID <> 0 Then

   If OpenProcessToken(hProcessID, TOKEN_READ, hToken) = 1 Then

      result = GetTokenInformation(hToken, TOKEN_ELEVATION_TYPE, tokenElevationType, 4, lReturnLength)

      If result = 0 Then
         ' Couldn’t get token information
         Exit Function
         End If

      If tokenElevationType <> 1 Then
        IsUserAnAdministrator = True
        End If

      CloseHandle hToken

      End If

    CloseHandle hProcessID

    End If


Exit Function


IsUserAnAdministratorError:

 ' Handle errors

End Function

Public Function isdiff2(color_0 As couleur, color_c As couleur, val As Long) As Boolean
   'find approximate color difference by taking the square root of the squared differences if r,g,b
   Dim dif As Double
   
   dif = Sqr((color_0.R - color_c.R) ^ 2 + (color_0.V - color_c.V) ^ 2 + (color_0.b - color_c.b) ^ 2)
   
   If dif <= 140 - val Then
      isdiff2 = True
   Else
      isdiff2 = False
      End If
End Function


Public Function recupcouleur(couleur As Long) As couleur

'Renvoie une valeur de type couleur contenant
'les composantes r v et b de la couleur de
'type long passée
'Returns a color value containing
'components R G and B of the color
'passed long type

Dim blue As Double
Dim green As Double
Dim red As Double

blue = Fix((couleur / 256) / 256)
green = Fix((couleur - ((blue * 256) * 256)) / 256)
red = Fix(couleur - ((blue * 256) * 256) - (green * 256))

recupcouleur.R = Abs(red)
recupcouleur.b = Abs(blue)
recupcouleur.V = Abs(green)

End Function

Public Sub tracecontours7(Pic As PictureBox, val As Long, trait As Boolean)
    'use bug walk to find the borders of the contour
    'first find boundary and move outside it.
    
    Dim R As Long
    Dim Test_Color As couleur
    Dim StartChain As POINTAPI
    Dim StepSizeX As Long
    Dim StepSizeY As Long
    Dim contour_color As couleur
    Dim PointChain As POINTAPI
    
    Dim CompletedContour As Boolean
    Dim ContourError As Boolean
    Dim CantFindBorder As Boolean
    Dim ReaachedBoundary As Boolean
    Dim EscapeKeyPressed As Boolean
    
    Dim PlusX As Boolean
    Dim PlusY As Boolean
    Dim MinusX As Boolean
    Dim MinusY As Boolean
    
    Dim CheckforContourConflicts As Boolean
    Dim ContourConflict As Boolean
    Dim numSpaceContour As Integer
    
'    Dim TraceColor As Long
'    Dim TracingColor As couleur
    
    Dim NumTurnsLeft As Integer
    Dim NumTurnsRight As Integer
    Dim numPixelSep As Integer ' connect contour points separated up to this amount of pixels
    Dim numBeginContour As Long
    
    Dim Starting As Boolean

    numPixelSep = numSpaceContour * Sqr(2)
    Pic.DrawMode = 13
    Pic.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
    
    CheckforContourConflicts = True 'set this for true to check if contour is turning around to crawl on the opposite border
    numSpaceContour = numDistContour '5 'only record contour point if it is numSpaceContour distant from the last one recorded
    
    TraceColor = ContourColor& 'QBColor(12)
    TracingColor = recupcouleur(TraceColor)
    
    contour_color = Start_Color
    
    StepSizeX = twipsx 'Screen.TwipsPerPixelX
    StepSizeY = twipsy 'Screen.TwipsPerPixelY
    
    GDMDIform.StatusBar1.Panels(1).Text = "Contour tracing activated....Press ''Esc'' to stop."
    
       
    If Not DigiLogFileOpened Then
'       pos% = InStr(picnam$, ".")
'       picext$ = Mid$(picnam$, pos% + 1, 3)
       DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
       Digilogfilnum% = FreeFile
       Open DigiLogfilnam$ For Append As #Digilogfilnum%
       DigiLogFileOpened = True
       End If
       
  
    'find the top edge and follow a Freeman chain staring from left to right clockwise
    For i = Start_Point.Y - StepSizeY To Pic.top Step -StepSizeY
        R = Pic.Point(Start_Point.x, i)
        If R <> -1 Then
            Test_Color = recupcouleur(R)
            If Not isdiff2(contour_color, Test_Color, val) Then
               'try one more just to be sure
               R = Pic.Point(Start_Point.x, i - StepSizeY)
               Test_Color = recupcouleur(R)
               If Not isdiff2(contour_color, Test_Color, val) Then
                   'this is top, so make this starting point for bug walk
                   StartChain.x = Start_Point.x
                   StartChain.Y = i 'this is one pixel above upper contour border
                   Exit For
                   End If
               End If
            End If
     Next i
     
     'now turn to the right to find the contour pixel
     'turn to the left after finding contour pixel
     
     PointChain.x = StartChain.x
     PointChain.Y = StartChain.Y
     
     Starting = True
     numBeginContour = numDigiContours
     
     Pic.DrawMode = 13
     Pic.DrawWidth = Max(1, CInt(1 * DigiZoom.LastZoom))

     'start by PlusX translation
     PointChain.x = PointChain.x + StepSizeX
     PlusX = True
     PlusY = False
     MinusX = False
     MinusY = False
     GoTo 700
     
100: 'turn right
     NumTurnsLeft = 0
     NumTurnsRight = NumTurnsRight + 1
     If NumTurnsRight > 4 Then
        'can't find the border
        CantFindBorder = True
        GoTo 900
        End If
     If PlusX Then
        PointChain.Y = PointChain.Y + StepSizeY
        PlusX = False
        PlusY = True
     ElseIf PlusY Then
        PointChain.x = PointChain.x - StepSizeX
        PlusY = False
        MinusX = True
     ElseIf MinusX Then
        PointChain.Y = PointChain.Y - StepSizeY
        MinusX = False
        MinusY = True
     ElseIf MinusY Then
        PointChain.x = PointChain.x + StepSizeX
        MinusY = False
        PlusX = True
        End If
        
     GoTo 700
        
500:  'turn Left
     NumTurnsRight = 0
     NumTurnsLeft = NumTurnsLeft + 1
     If NumTurnsLeft > 4 Then
        'can't find the border
        CantFindBorder = True
        GoTo 900
        End If
     If PlusX Then
        PointChain.Y = PointChain.Y - StepSizeY
        PlusX = False
        MinusY = True
     ElseIf PlusY Then
        PointChain.x = PointChain.x + StepSizeX
        PlusY = False
        PlusX = True
     ElseIf MinusX Then
        PointChain.Y = PointChain.Y + StepSizeY
        MinusX = False
        PlusY = True
     ElseIf MinusY Then
        PointChain.x = PointChain.x - StepSizeX
        MinusY = False
        MinusX = True
        End If
        
700: 'first check if returned to beginning
     If PointChain.x = StartChain.x And PointChain.Y = StartChain.Y Then
        'reached the beginning
        'draw the contour trace and exit
'        For i = 1 To numDigiContours - 1
'           Pic.Line (DigiContours(i).X, DigiContours(i).Y)-(DigiContours(i - 1).X, DigiContours(i - 1).Y), TraceColor
'        Next i
'        Pic.Line (DigiContours(0).X, DigiContours(0).Y)-(DigiContours(numDigiContours - 1).X, DigiContours(numDigiContours - 1).Y), TraceColor
        CompletedContour = True
        GoTo 900
     ElseIf PointChain.x <= Pic.ScaleLeft Or PointChain.x >= Pic.ScaleLeft + Pic.ScaleWidth Or _
        PointChain.Y <= Pic.ScaleTop Or PointChain.Y >= Pic.ScaleTop + Pic.ScaleHeight Then
        'reached boundary
        'draw the contour trace and exit
'        For i = 1 To numDigiContours - 1
'           Pic.Line (DigiContours(i).X, DigiContours(i).Y)-(DigiContours(i - 1).X, DigiContours(i - 1).Y), TraceColor
'        Next i
'        Pic.Line (DigiContours(0).X, DigiContours(0).Y)-(DigiContours(numDigiContours - 1).X, DigiContours(numDigiContours - 1).Y), TraceColor
        ReaachedBoundary = True
        GoTo 900
     Else
        'check if this point is inside the contour or not
        'if inside the contour turn left
        R = Pic.Point(PointChain.x, PointChain.Y)
        If R <> -1 Then
           Test_Color = recupcouleur(R)
           If isdiff4(Test_Color, contour_color, TracingColor, val) Then
           
              If Starting Then 'record starting point
                 StartChain.x = PointChain.x
                 StartChain.Y = PointChain.Y
                 Starting = False
                 End If
                 
              'space contour points at least numSpaceContour pixels apart, draw line between them
              If numDigiContours > 0 Then
              
                  'distance from last recorded point on contour
                  DisPix = Sqr((CLng(PointChain.x / DigiZoom.LastZoom) - DigiContours(numDigiContours - 1).x) ^ 2# + (CLng(PointChain.Y / DigiZoom.LastZoom) - DigiContours(numDigiContours - 1).Y) ^ 2#)
                  
                  If DisPix >= numSpaceContour Then
                  
                      If CheckforContourConflicts Then
                         ContourConflict = NearNeighbors(Pic, PointChain.x, PointChain.Y)
                         If ContourConflict Then GoTo 900
                         End If
                  
                      'record point and turn left
                      If numDigiContours = 0 Then
                         ReDim DigiContours(0)
'                         ReDim DigiContourColors(0)
                      Else
                         ReDim Preserve DigiContours(numDigiContours)
'                         ReDim Preserve DigiContourColors(numDigiContours)
                         End If
                         
                      DigiContours(numDigiContours).x = CLng(PointChain.x / DigiZoom.LastZoom)
                      DigiContours(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
                      DigiContours(numDigiContours).Z = ContourHeight * InvElev
                  
                      If DigiContours(numDigiContours).Z < MinColorHeight Then MinColorHeight = ContourHeight * InvElev
                      If DigiContours(numDigiContours).Z > MaxColorHeight Then MaxColorHeight = ContourHeight * InvElev
                      
                      If ImagePointFile Then 'record into byte array
                         ier = RecordDigiPointsImage(CLng(DigiContours(numDigiContours).x), CLng(DigiContours(numDigiContours).Y), 1)
                         End If
                      
'                      DigiContourColors(numDigiContours).X = CLng(PointChain.X / DigiZoom.LastZoom)
'                      DigiContourColors(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
'                      DigiContourColors(numDigiContours).RedColor = Test_Color.R
'                      DigiContourColors(numDigiContours).GreenColor = Test_Color.v
'                      DigiContourColors(numDigiContours).BlueColor = Test_Color.b
                      
                      numDigiContours = numDigiContours + 1
        
                      Write #Digilogfilnum%, CLng(PointChain.x / DigiZoom.LastZoom), CLng(PointChain.Y / DigiZoom.LastZoom), ContourHeight * InvElev, 1 '1 is flag of a contour height
                      Call ShiftMap(CSng(PointChain.x), CSng(PointChain.Y))
                      
                      Pic.PSet (PointChain.x, PointChain.Y), TraceColor
                      
                      'draw line from last recorded point on contour if it is close enough to be a continuation
'                      DisPix = Sqr((CLng(PointChain.x / DigiZoom.LastZoom) - DigiContours(numDigiContours - 2).x) ^ 2# + (CLng(PointChain.Y / DigiZoom.LastZoom) - DigiContours(numDigiContours - 2).Y) ^ 2#)
'                      If DisPix <= numPixelSep Then
                            If numDigiContours - 1 >= numBeginContour Then
                               Pic.Line (PointChain.x, PointChain.Y)-(DigiContours(numDigiContours - 2).x * DigiZoom.LastZoom, DigiContours(numDigiContours - 2).Y * DigiZoom.LastZoom), TraceColor
                               End If
                            Pic.Refresh
'                            End If
'                  Else
'                     'draw line between them, but don't record the point since the freeman chain don't lie on a straight line
'                     Pic.Line (DigiContours(numDigiContours - 1).X, DigiContours(numDigiContours - 1).Y)-(PointChain.X, PointChain.Y), TraceColor
                     End If
              ElseIf numDigiContours = 0 Then
                   If numDigiContours = 0 Then
                     ReDim DigiContours(0)
'                     ReDim DigiContourColors(0)
                  Else
                     ReDim Preserve DigiContours(numDigiContours)
'                     ReDim Preserve DigiContourColors(numDigiContours)
                     End If
                     
                  DigiContours(numDigiContours).x = CLng(PointChain.x / DigiZoom.LastZoom)
                  DigiContours(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
                  DigiContours(numDigiContours).Z = ContourHeight * InvElev
                  
                  If DigiContours(numDigiContours).Z < MinColorHeight Then MinColorHeight = ContourHeight * InvElev
                  If DigiContours(numDigiContours).Z > MaxColorHeight Then MaxColorHeight = ContourHeight * InvElev
                      
                  If ImagePointFile Then 'record into byte array
                     ier = RecordDigiPointsImage(CLng(DigiContours(numDigiContours).x), CLng(DigiContours(numDigiContours).Y), 1)
                     End If
                      
'                  DigiContourColors(numDigiContours).X = CLng(PointChain.X / DigiZoom.LastZoom)
'                  DigiContourColors(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
'                  DigiContourColors(numDigiContours).RedColor = Test_Color.R
'                  DigiContourColors(numDigiContours).GreenColor = Test_Color.v
'                  DigiContourColors(numDigiContours).BlueColor = Test_Color.b
                      
                  numDigiContours = numDigiContours + 1
    
                  Write #Digilogfilnum%, CLng(PointChain.x / DigiZoom.LastZoom), CLng(PointChain.Y / DigiZoom.LastZoom), ContourHeight * InvElev, 1 '1 is flag of a contour height
                  Call ShiftMap(CSng(PointChain.x), CSng(PointChain.Y))
                  Pic.PSet (PointChain.x, PointChain.Y), TraceColor
                  Pic.Refresh
                  
                  End If
              
              '---------------------break on ESC key-------------------------------
              If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
                 EscapeKeyPressed = True
                 GoTo 900
                 End If
              '---------------------------------------------------------------------
             
              DoEvents
              GoTo 500 'turn left to find border again
           Else
              GoTo 100 'turn right
              End If
        Else
           ContourError = True
           GoTo 900
           End If
        End If
     
900:
    Pic.Refresh
    DoEvents
    
    If numDigiContours > 0 And GDMDIform.mnuEraser.Enabled = False Then
       GDMDIform.Toolbar1.Buttons(40).Enabled = True
       GDMDIform.mnuEraser.Enabled = True
       GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
       buttonstate&(40) = 0
       End If
       
    If numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0 And GDMDIform.mnuDigiSweep.Enabled = False Then
       GDMDIform.Toolbar1.Buttons(41).Enabled = True
       GDMDIform.mnuDigiSweep.Enabled = True
       GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
       buttonstate&(41) = 0
       End If
    
  GDMDIform.StatusBar1.Panels(1).Text = sEmpty
  
  If CompletedContour And DigitizeContour Then
     'finished the contour
'     Select Case MsgBox("Contour has been traced back to the starting point." _
'                        & vbCrLf & "" _
'                        & vbCrLf & "Do you want to continue tracing contours with current elevation?             " _
'                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Completed contour...")
'
'        Case vbYes
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case vbNo
'
'           GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Contour has been traced back to the starting point." _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, will click on new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
    
     End Select
    
     
     
  ElseIf ReaachedBoundary Then
     'reached picturebox boundary
''     Select Case MsgBox("Contour has been traced back to the picture's boundary." _
''                        & vbCrLf & "" _
''                        & vbCrLf & "Do you want to continue tracing contours with current elevation?             " _
''                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Hit boundary...")
''
''        Case vbYes
''            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
''
''        Case vbNo
''
''           GoTo 9999
''
''     End Select
'
'     frmMsgBox.MsgCstm "Contour has been traced back to the picture's boundary." _
'                      & vbCrLf & "" _
'                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
'                      "Contour following", mbQuestion, 1, False, _
'                      "Yes, continue at a new place with the current elevation", "No, Stop here"
'
'     Select Case frmMsgBox.g_lBtnClicked
'
'        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case 0, 2
'            GoTo 9999
'
'     End Select
     
     
  ElseIf ContourError And DigitizeContour Then
     'couldn't read color
'     Select Case MsgBox("Error encountered in reading the pixel color." _
'                        & vbCrLf & "" _
'                        & vbCrLf & "Do you want to continue tracing contours with current elevation?" _
'                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Contour Error...")
'
'        Case vbYes
'
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case vbNo
'
'          GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Error encountered in reading the pixel color." _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, will click on new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
            
     End Select
     
  ElseIf CantFindBorder And DigitizeContour Then
'     Select Case MsgBox("Lost the contour!" _
'                        & vbCrLf & "" _
'                        & vbCrLf & "Do you want to continue tracing contours with current elevation?" _
'                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Lost Contour border....")
'
'        Case vbYes
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case vbNo
'
'           GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Lost the contour!" _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, will click on new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
            
     End Select
     
  ElseIf ContourConflict And DigitizeContour Then
  
      frmMsgBox.MsgCstm "Contour seems to be turning on itself!" _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, I will click on a new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
            
     End Select
    
     
  ElseIf EscapeKeyPressed And DigitizeContour Then
  
'     Select Case MsgBox("Escape key pressed, tracing halted..." _
'                       & vbCrLf & "" _
'                       & vbCrLf & "Resume?" _
'                       & vbCrLf & "" _
'                       & vbCrLf & "Answer: ""Yes"" to resume" _
'                       & vbCrLf & "               ""No"" to pick new starting point with last elevation" _
'                       & vbCrLf & "               ""Cancel"" to pick a new elevation" _
'                       , vbYesNoCancel Or vbInformation Or vbDefaultButton1, "Escape...")
'
'       Case vbYes
'         GoTo 500 'turn left
'
'       Case vbNo
'         MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'       Case vbCancel
'         GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Escape key pressed, tracing halted..." _
                      & vbCrLf & "" _
                      & vbCrLf & "Resume?", _
                      "Ignore the pause?", mbQuestion, 2, False, _
                      "Yes, ignore it", "No, pick a new starting point", "No, cancel digitizing contours"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
            GoTo 500 'turn left
    
        Case 2
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
            
        Case 0, 3
            GoTo 9999
            
     End Select
     
  
     End If
     
  PointStart = False 'let user keep on clicking to extend contour
  
  Exit Sub
  
9999
    If Not DigitizeOn Then
       Call GDMDIform.mnuDigitizeEndContour_Click 'end digitizing contours
    Else
       'keypress enter
       KeyDown (vbKeyReturn)
       End If
'    DigitizeContour = False
'    Call ShiftMap(CSng(PointChain.x), CSng(PointChain.y))
'    'reset blinking
'    GDMDIform.CenterPointTimer.Enabled = True
'    ce& = 1
        
End Sub
Public Function isdiff4(color_t As couleur, color_c As couleur, TC As couleur, val As Long) As Boolean
   'find approximate color difference by taking the square root of the squared differences if r,g,b
   Dim dif As Double
'   Dim dif_Black As Double
   
   If (color_t.R = TC.R And color_t.V = TC.V And color_t.b = TC.b) Then
      isdiff4 = True
      Exit Function
      End If
   
   'Euclidean color difference from contour color
   dif = Sqr((color_t.R - color_c.R) ^ 2 + (color_t.V - color_c.V) ^ 2 + (color_t.b - color_c.b) ^ 2)
   'Euclidean color difference from Black
'   dif_Black = Sqr((color_t.R) ^ 2 + (color_t.v) ^ 2 + (color_t.b) ^ 2)
   
   If dif <= 140 - val Then 'And dif_Black > 170 Then  'within tolerance of the contour color and not too close to Black
      isdiff4 = True
   Else
      isdiff4 = False
      End If
      
End Function
Public Sub tracecontours8(Pic As PictureBox, val As Long, trait As Boolean)
    'use Freeman 8 direction chain method to find the borders of the contour
    'first find boundary and move outside it.
    
'    Dim oGestionImageSrc As New CGestionImage
'    Dim iFor1 As Integer 'stocke les valeurs de la boucle For->Next
'    Dim iFor2 As Integer 'stocke les valeurs de la boucle For->Next
'    Dim iBleu As Byte 'stocke la composante bleue à récupèrer
'    Dim iVert As Byte 'stocke la composante verte à récupèrer
'    Dim iRouge As Byte 'stocke la composante rouge à récupèrer
'
'    'on définit les contrôles sources et destination
'    Set oGestionImageSrc.PictureBox = pic

    Dim R As Long
    Dim Test_Color As couleur
    Dim StartChain As POINTAPI
    Dim StepSizeX As Long
    Dim StepSizeY As Long
    Dim contour_color As couleur
    Dim PointChain As POINTAPI
    
    Dim CompletedContour As Boolean
    Dim ContourError As Boolean
    Dim CantFindBorder As Boolean
    Dim ReaachedBoundary As Boolean
    
    Dim PlusX As Boolean
    Dim Plus45 As Boolean
    Dim PlusY As Boolean
    Dim Plus135 As Boolean
    Dim MinusX As Boolean
    Dim Minus225 As Boolean
    Dim MinusY As Boolean
    Dim Minus315 As Boolean
    Dim CheckforContourConflicts As Boolean
    Dim ContourConflict As Boolean
    Dim numSpaceContour As Integer
    Dim numPixelSep As Integer ' connect contour points separated up to this amount of pixels
    Dim numBeginContour As Long
    
'    Dim TraceColor As Long
'    Dim TracingColor As couleur
    
    Dim NumTurnsLeft As Integer
    Dim NumTurnsRight As Integer
    
    Dim Starting As Boolean
    
'    numDigiContours = 0
    
    numPixelSep = numSpaceContour * Sqr(2)

    Pic.DrawMode = 13
    Pic.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
    CheckforContourConflicts = True 'set this for true to check if contour is turning around to crawl on the opposite border
    numSpaceContour = numDistContour '5 'only record contour point if it is numSpaceContour distant from the last one recorded
    
    TraceColor = ContourColor& 'QBColor(12)
    TracingColor = recupcouleur(TraceColor)
    
    contour_color = Start_Color
    
    StepSizeX = twipsx 'Max(1, CLng(twipsx * DigiZoom.LastZoom))  'Screen.TwipsPerPixelX
    StepSizeY = twipsy 'Max(1, CLng(twipsy * DigiZoom.LastZoom)) 'Screen.TwipsPerPixelY
    
    GDMDIform.StatusBar1.Panels(1).Text = "Contour tracing activated....Press ''Esc'' to stop."
    
    If Not DigiLogFileOpened Then
'       pos% = InStr(picnam$, ".")
'       picext$ = Mid$(picnam$, pos% + 1, 3)
       DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
       Digilogfilnum% = FreeFile
       Open DigiLogfilnam$ For Append As #Digilogfilnum%
       DigiLogFileOpened = True
       End If
       
    'fixes some sort of bug when using the easer without first digitizing contours
    If val = INIT_VALUE Then
       Exit Sub
       End If
  
    'find the top edge and follow a Freeman chain staring from left to right clockwise
    For i = Start_Point.Y - StepSizeY To Pic.top Step -StepSizeY
        R = Pic.Point(Start_Point.x, i)
        If R <> -1 Then
            Test_Color = recupcouleur(R)
            If Not isdiff2(contour_color, Test_Color, val) Then
               'try one more just to be sure
               R = Pic.Point(Start_Point.x, i - StepSizeY)
               Test_Color = recupcouleur(R)
               If Not isdiff2(contour_color, Test_Color, val) Then
                   'this is top, so make this starting point for bug walk
                   StartChain.x = Start_Point.x
                   StartChain.Y = i  'this is one pixel above upper contour border
                   Exit For
                   End If
               End If
            End If
     Next i
     
     'now turn to the right to find the contour pixel
     'turn to the left after finding contour pixel
     
     PointChain.x = StartChain.x
     PointChain.Y = StartChain.Y
     
     Starting = True
     numBeginContour = numDigiContours
     
     Pic.DrawMode = 13
     Pic.DrawWidth = Max(1, CInt(1 * DigiZoom.LastZoom))

     'start by PlusX translation
     PointChain.x = PointChain.x + StepSizeX
     PlusX = True
     PlusY = False
     MinusX = False
     MinusY = False
     GoTo 700
     
100: 'turn right
     NumTurnsLeft = 0
     NumTurnsRight = NumTurnsRight + 1
     If NumTurnsRight > 7 Then
        'can't find the border
        CantFindBorder = True
        GoTo 900
        End If
     If PlusX Then
        PointChain.Y = PointChain.Y + StepSizeY
        PointChain.x = PointChain.x + StepSizeX
        PlusX = False
        Minus315 = True
     ElseIf Minus315 = True Then
        PointChain.Y = PointChain.Y + StepSizeY
        Minus315 = False
        MinusY = True
     ElseIf MinusY Then
        PointChain.x = PointChain.x - StepSizeX
        PointChain.Y = PointChain.Y + StepSizeY
        MinusY = False
        Minus225 = True
     ElseIf Minus225 Then
        PointChain.x = PointChain.x - StepSizeX
        Minus225 = False
        MinusX = True
     ElseIf MinusX Then
        PointChain.Y = PointChain.Y - StepSizeY
        PointChain.x = PointChain.x - StepSizeX
        MinusX = False
        Plus135 = True
     ElseIf Plus135 Then
        PointChain.Y = PointChain.Y - StepSizeY
        Plus135 = False
        PlusY = True
     ElseIf PlusY Then
        PointChain.x = PointChain.x + StepSizeX
        PointChain.Y = PointChain.Y - StepSizeY
        PlusY = False
        Plus45 = True
     ElseIf Plus45 Then
        PointChain.x = PointChain.x + StepSizeX
        Plus45 = False
        PlusX = True
        End If
        
     GoTo 700
        
500:  'turn Left
     NumTurnsRight = 0
     NumTurnsLeft = NumTurnsLeft + 1
     If NumTurnsLeft > 7 Then
        'can't find the border
        CantFindBorder = True
        GoTo 900
        End If
     If PlusX Then
        PointChain.Y = PointChain.Y - StepSizeY
        PlusX = False
        Plus45 = True
     ElseIf Plus45 = True Then
        PointChain.x = PointChain.x - StepSizeX
        Plus45 = False
        PlusY = True
     ElseIf PlusY Then
        PointChain.x = PointChain.x - StepSizeX
        PlusY = False
        Plus135 = True
     ElseIf Plus135 Then
        PointChain.Y = PointChain.Y + StepSizeY
        Plus135 = False
        MinusX = True
     ElseIf MinusX Then
        PointChain.Y = PointChain.Y + StepSizeY
        MinusX = False
        Minus225 = True
     ElseIf Minus225 Then
        PointChain.x = PointChain.x + StepSizeX
        Minus225 = False
        MinusY = True
     ElseIf MinusY Then
        PointChain.x = PointChain.x + StepSizeX
        MinusY = False
        Minus315 = True
     ElseIf Minus315 Then
        PointChain.Y = PointChain.Y - StepSizeY
        Minus315 = False
        PlusX = True
        End If
        
700: 'first check if returned to beginning
     If PointChain.x = StartChain.x And PointChain.Y = StartChain.Y And Not Starting Then
        'reached the beginning
        'draw the contour trace and exit
'        For i = 1 To numDigiContours - 1
'           Pic.Line (DigiContours(i).X, DigiContours(i).Y)-(DigiContours(i - 1).X, DigiContours(i - 1).Y), TraceColor
'        Next i
'        Pic.Line (DigiContours(0).X, DigiContours(0).Y)-(DigiContours(numDigiContours - 1).X, DigiContours(numDigiContours - 1).Y), TraceColor
        CompletedContour = True
        GoTo 900
     ElseIf PointChain.x <= Pic.ScaleLeft Or PointChain.x >= Pic.ScaleLeft + Pic.ScaleWidth Or _
        PointChain.Y <= Pic.ScaleTop Or PointChain.Y >= Pic.ScaleTop + Pic.ScaleHeight Then
        'reached boundary
        'draw the contour trace and exit
'        For i = 1 To numDigiContours - 1
'           Pic.Line (DigiContours(i).X, DigiContours(i).Y)-(DigiContours(i - 1).X, DigiContours(i - 1).Y), TraceColor
'        Next i
'        Pic.Line (DigiContours(0).X, DigiContours(0).Y)-(DigiContours(numDigiContours - 1).X, DigiContours(numDigiContours - 1).Y), TraceColor
        ReaachedBoundary = True
        GoTo 900
     Else
        'check if this point is inside the contour or not
        R = Pic.Point(PointChain.x, PointChain.Y)
        If R <> -1 Then
           Test_Color = recupcouleur(R)
           If isdiff4(Test_Color, contour_color, TracingColor, val) Then
           
              If Starting Then 'record starting point
                 StartChain.x = PointChain.x
                 StartChain.Y = PointChain.Y
                 Starting = False
                 End If
                 
              'space contour points at least numSpaceContour pixels apart, draw line between them
              If numDigiContours > 0 Then
              
                  'distance from last recorded point on contour
                  DisPix = Sqr((CLng(PointChain.x / DigiZoom.LastZoom) - DigiContours(numDigiContours - 1).x) ^ 2# + (CLng(PointChain.Y / DigiZoom.LastZoom) - DigiContours(numDigiContours - 1).Y) ^ 2#)
                  
                  If DisPix >= numSpaceContour Then
                  
                      If CheckforContourConflicts Then
                         ContourConflict = NearNeighbors(Pic, PointChain.x, PointChain.Y)
                         If ContourConflict Then GoTo 900
                         End If
                  
                      'record point and turn left
                      If numDigiContours = 0 Then
                         ReDim DigiContours(0)
'                         ReDim DigiContourColors(0)
                      Else
                         ReDim Preserve DigiContours(numDigiContours)
'                         ReDim Preserve DigiContourColors(numDigiContours)
                         End If
                         
                      DigiContours(numDigiContours).x = CLng(PointChain.x / DigiZoom.LastZoom)
                      DigiContours(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
                      DigiContours(numDigiContours).Z = ContourHeight * InvElev
                  
                      If DigiContours(numDigiContours).Z < MinColorHeight Then MinColorHeight = ContourHeight * InvElev
                      If DigiContours(numDigiContours).Z > MaxColorHeight Then MaxColorHeight = ContourHeight * InvElev
                      
                      If ImagePointFile Then 'record into byte array
                         ier = RecordDigiPointsImage(CLng(DigiContours(numDigiContours).x), CLng(DigiContours(numDigiContours).Y), 1)
                         End If
                      
'                      DigiContourColors(numDigiContours).X = CLng(PointChain.X / DigiZoom.LastZoom)
'                      DigiContourColors(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
'                      DigiContourColors(numDigiContours).RedColor = Test_Color.R
'                      DigiContourColors(numDigiContours).GreenColor = Test_Color.v
'                      DigiContourColors(numDigiContours).BlueColor = Test_Color.b
        
                      Write #Digilogfilnum%, CLng(PointChain.x / DigiZoom.LastZoom), CLng(PointChain.Y / DigiZoom.LastZoom), ContourHeight * InvElev, 1 '1 is flag of a contour height
                      Call ShiftMap(CSng(PointChain.x), CSng(PointChain.Y))
                      
                      Pic.PSet (PointChain.x, PointChain.Y), TraceColor
                      'draw line from last recorded point on contour if it is close enough to be a continuation
'                      DisPix = Sqr((CLng(PointChain.X / DigiZoom.LastZoom) - DigiContours(numDigiContours - 1).X) ^ 2# + (CLng(PointChain.Y / DigiZoom.LastZoom) - DigiContours(numDigiContours - 1).Y) ^ 2#)
'                      If DisPix <= numPixelSep Then
                         If numDigiContours - 1 >= numBeginContour Then
                           Pic.Line (PointChain.x, PointChain.Y)-(DigiContours(numDigiContours - 1).x * DigiZoom.LastZoom, DigiContours(numDigiContours - 1).Y * DigiZoom.LastZoom), TraceColor
                           End If
                         Pic.Refresh
'                         End If
                      
                      numDigiContours = numDigiContours + 1

'                  Else
'                     'draw line between them, but don't record the point since the freeman chain don't lie on a straight line
'                     Pic.Line (DigiContours(numDigiContours - 1).X, DigiContours(numDigiContours - 1).Y)-(PointChain.X, PointChain.Y), TraceColor
                     End If
              ElseIf numDigiContours = 0 Then
                   If numDigiContours = 0 Then
                     ReDim DigiContours(0)
'                     ReDim DigiContourColors(0)
                  Else
                     ReDim Preserve DigiContours(numDigiContours)
'                     ReDim Preserve DigiContourColors(numDigiContours)
                     End If
                     
                  DigiContours(numDigiContours).x = CLng(PointChain.x / DigiZoom.LastZoom)
                  DigiContours(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
                  DigiContours(numDigiContours).Z = ContourHeight * InvElev
                  
                  If DigiContours(numDigiContours).Z < MinColorHeight Then MinColorHeight = ContourHeight * InvElev
                  If DigiContours(numDigiContours).Z > MaxColorHeight Then MaxColorHeight = ContourHeight * InvElev
                      
                  If ImagePointFile Then 'record into byte array
                     ier = RecordDigiPointsImage(CLng(DigiContours(numDigiContours).x), CLng(DigiContours(numDigiContours).Y), 1)
                     End If
                      
'                  DigiContourColors(numDigiContours).X = CLng(PointChain.X / DigiZoom.LastZoom)
'                  DigiContourColors(numDigiContours).Y = CLng(PointChain.Y / DigiZoom.LastZoom)
'                  DigiContourColors(numDigiContours).RedColor = Test_Color.R
'                  DigiContourColors(numDigiContours).GreenColor = Test_Color.v
'                  DigiContourColors(numDigiContours).BlueColor = Test_Color.b
                      
                  numDigiContours = numDigiContours + 1
    
                  Write #Digilogfilnum%, CLng(PointChain.x / DigiZoom.LastZoom), CLng(PointChain.Y / DigiZoom.LastZoom), ContourHeight * InvElev, 1 '1 is flag of a contour height
                  Call ShiftMap(CSng(PointChain.x), CSng(PointChain.Y))
                  Pic.PSet (PointChain.x, PointChain.Y), TraceColor
                  Pic.Refresh
                  
                  End If
                  
              '---------------------break on ESC key-------------------------------
              If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
                 EscapeKeyPressed = True
                 GoTo 900
                 End If
              '---------------------------------------------------------------------
             
              DoEvents
              GoTo 500 'turn left to find border again
           Else
              GoTo 100 'turn right to keep on looking for contour
              End If
        Else
           ContourError = True
           GoTo 900
           End If
        End If
     
900:
    Pic.Refresh
    DoEvents
    
    If numDigiContours > 0 And GDMDIform.mnuEraser.Enabled = False Then
       GDMDIform.Toolbar1.Buttons(40).Enabled = True
       GDMDIform.mnuEraser.Enabled = True
       GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
       buttonstate&(40) = 0
       End If
       
    If numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0 And GDMDIform.mnuDigiSweep.Enabled = False Then
       GDMDIform.Toolbar1.Buttons(41).Enabled = True
       GDMDIform.mnuDigiSweep.Enabled = True
       GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
       buttonstate&(41) = 0
       End If
       
    
  GDMDIform.StatusBar1.Panels(1).Text = sEmpty
  
  If CompletedContour And DigitizeContour Then
     'finished the contour
'     Select Case MsgBox("Contour has been traced back to the starting point." _
'                        & vbCrLf & "" _
'                        & vbCrLf & "Do you want to continue tracing contours with current elevation?             " _
'                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Completed contour...")
'
'        Case vbYes
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case vbNo
'
'           GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Contour has been traced back to the starting point." _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to click on new begenning of contour with current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, will click at new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
    
     End Select
    
     
     
  ElseIf ReaachedBoundary And DigitizeContour Then
     'reached picturebox boundary
''     Select Case MsgBox("Contour has been traced back to the picture's boundary." _
''                        & vbCrLf & "" _
''                        & vbCrLf & "Do you want to continue tracing contours with current elevation?             " _
''                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Hit boundary...")
''
''        Case vbYes
''            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
''
''        Case vbNo
''
''           GoTo 9999
''
''     End Select
'
'     frmMsgBox.MsgCstm "Contour has been traced back to the picture's boundary." _
'                      & vbCrLf & "" _
'                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
'                      "Contour following", mbQuestion, 1, False, _
'                      "Yes, continue at a new place with the current elevation", "No, Stop here"
'
'     Select Case frmMsgBox.g_lBtnClicked
'
'        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case 0, 2
'            GoTo 9999
'
'     End Select
     
     
  ElseIf ContourError And DigitizeContour Then
     'couldn't read color
'     Select Case MsgBox("Error encountered in reading the pixel color." _
'                        & vbCrLf & "" _
'                        & vbCrLf & "Do you want to continue tracing contours with current elevation?" _
'                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Contour Error...")
'
'        Case vbYes
'
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case vbNo
'
'          GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Error encountered in reading the pixel color." _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to click on new begenning of contour with current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, will click at new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
            'MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
            
     End Select
     
  ElseIf CantFindBorder And DigitizeContour Then
'     Select Case MsgBox("Lost the contour!" _
'                        & vbCrLf & "" _
'                        & vbCrLf & "Do you want to continue tracing contours with current elevation?" _
'                        , vbYesNo Or vbInformation Or vbDefaultButton1, "Lost Contour border....")
'
'        Case vbYes
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'        Case vbNo
'
'           GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Lost the contour!" _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, will click at new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
            
     End Select
     
  ElseIf ContourConflict And DigitizeContour Then
  
      frmMsgBox.MsgCstm "Contour seems to be turning on itself!" _
                      & vbCrLf & "" _
                      & vbCrLf & "Do you want to continue tracing contours with the current elevation?", _
                      "Contour following", mbQuestion, 1, False, _
                      "Yes, will click at new start", "No, Stop here"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
    
        Case 0, 2
            GoTo 9999
            
     End Select
    
     
  ElseIf EscapeKeyPressed And DigitizeContour Then
  
'     Select Case MsgBox("Escape key pressed, tracing halted..." _
'                       & vbCrLf & "" _
'                       & vbCrLf & "Resume?" _
'                       & vbCrLf & "" _
'                       & vbCrLf & "Answer: ""Yes"" to resume" _
'                       & vbCrLf & "               ""No"" to pick new starting point with last elevation" _
'                       & vbCrLf & "               ""Cancel"" to pick a new elevation" _
'                       , vbYesNoCancel Or vbInformation Or vbDefaultButton1, "Escape...")
'
'       Case vbYes
'         GoTo 500 'turn left
'
'       Case vbNo
'         MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
'
'       Case vbCancel
'         GoTo 9999
'
'     End Select
     
     frmMsgBox.MsgCstm "Escape key pressed, tracing halted..." _
                      & vbCrLf & "" _
                      & vbCrLf & "Resume?", _
                      "Ignore the pause?", mbQuestion, 2, False, _
                      "Yes, ignore it", "No, pick a new starting point", "No, cancel digitizing contours"
    
     Select Case frmMsgBox.g_lBtnClicked
    
        Case 1
            GoTo 500 'turn left
    
        Case 2
'            MsgBox "Click on the new beginning...", vbOKOnly + vbInformation, "Restart contour.."
            
        Case 0, 3
            GoTo 9999
            
     End Select
     
  
     End If
     
  PointStart = False 'let user keep on clicking to extend contour
  
  Exit Sub
  
9999
    If Not DigitizeOn Then
       Call GDMDIform.mnuDigitizeEndContour_Click 'end digitizing contours
    Else
       'keypress enter
       KeyDown (vbKeyReturn)
       End If

'    DigitizeContour = False
'    Call ShiftMap(CSng(PointChain.x), CSng(PointChain.y))
'    'reset blinking
'    GDMDIform.CenterPointTimer.Enabled = True
'    ce& = 1
        
End Sub

Public Sub UpdateDigiLogFile()

   'updates the log file after a deletion
   
   If DigiLogFileOpened Then
      Close #Digilogfilnum%
      Digilogfilnum% = 0
      DigiLogFileOpened = False
      End If
   
   'open file for writing
'   pos% = InStr(picnam$, ".")
'   picext$ = Mid$(picnam$, pos% + 1, 3)
   DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
   Digilogfilnum% = FreeFile
   Open DigiLogfilnam$ For Output As #Digilogfilnum%

   'record lines
   For i& = 0 To numDigiLines - 1
       Write #Digilogfilnum%, DigiLines(0, i&).x, DigiLines(0, i&).Y, DigiLines(0, i&).Z, 3 '3 is the flag for line digitizing beginning
       Write #Digilogfilnum%, DigiLines(1, i&).x, DigiLines(1, i&).Y, DigiLines(1, i&).Z, 4 '4 is the flag for line digitizing end
       If DigiLines(1, i&).Z < MinColorHeight Then MinColorHeight = DigiLines(1, i&).Z
       If DigiLines(1, i&).Z > MaxColorHeight Then MaxColorHeight = DigiLines(1, i&).Z
   Next i&
   
   'record contours
   For i& = 0 To numDigiContours - 1
       Write #Digilogfilnum%, DigiContours(i&).x, DigiContours(i&).Y, DigiContours(i&).Z, 1 '1 is the flag for contour generated points
       If DigiContours(i&).Z < MinColorHeight Then MinColorHeight = DigiContours(i&).Z
       If DigiContours(i&).Z > MaxColorHeight Then MaxColorHeight = DigiContours(i&).Z
   Next i&
   
   'record points 'record last to cause it to be plotted last
   For i& = 0 To numDigiPoints - 1
       Write #Digilogfilnum%, DigiPoints(i&).x, DigiPoints(i&).Y, DigiPoints(i&).Z, 2 '2 is the flag for point digitizing
       If DigiPoints(i&).Z < MinColorHeight Then MinColorHeight = DigiPoints(i&).Z
       If DigiPoints(i&).Z > MaxColorHeight Then MaxColorHeight = DigiPoints(i&).Z
   Next i&

   'record erasures
   For i& = 0 To numDigiErase - 1
       Write #Digilogfilnum%, DigiErasePoints(i&).x, DigiErasePoints(i&).Y, 0, 5 '5 is the flag for erasures
   Next i&
   
   Close #Digilogfilnum%
   
   DigiLogFileOpened = False

End Sub

'---------------------------------------------------------------------------------------
' Procedure : InputDigiLogFile
' Author    : Chaim Keller
' Date      : 2/6/2015
' Purpose   : read formaly recorded digitization for this map file
'---------------------------------------------------------------------------------------
'
Public Sub InputDigiLogFile()

   Dim FirstLineVertex As Boolean
   
   'on définit la couleur du pixel courant à partir des pixels alentours
   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer
   
   Dim numPixelSep As Integer ' connect contour points separated up to this amount of pixels
   Dim DisPix As Double
   
   Dim Xpix As Single, Ypix As Single, Zpix As Single
      
   On Error GoTo InputDigiLogFile_Error
   
   Screen.MousePointer = vbHourglass
   GDMDIform.StatusBar1.Panels(1).Text = "Please wait, loading and plotting stored digitized data."
   
   'stop blinking search points for 1:50000 maps
   GDMDIform.CenterPointTimer.Enabled = False
   ce& = 0 'reset flag that draws blinking cursor

   If DigiLogFileOpened Then
      Close #Digilogfilnum%
      Digilogfilnum% = 0
      DigiLogFileOpened = False
      End If
    
'    pos% = InStr(picnam$, ".")
'    picext$ = Mid$(picnam$, pos% + 1, 3)
    DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
    Digilogfilnum% = FreeFile
    myfile$ = Dir(DigiLogfilnam$)
    If myfile$ <> sEmpty Then
       Open DigiLogfilnam$ For Input As #Digilogfilnum%
    Else
       Screen.MousePointer = vbDefault
       Exit Sub 'no log file found for this map file
       End If
          
'   'clear the picture of any plot marks
'   GDform1.Picture2.Cls
'   GDform1.Picture2.Picture = LoadPicture(picnam$)
'   If DigiZoom.LastZoom <> 1 Then Call PictureBoxZoom(GDform1.Picture2, 0, 0, 0, 0) 'zoom out according to last zoom
   
   numPixelSep = numDistContour * Sqr(2)
   TraceColor = ContourColor& 'QBColor(12)
   
'   'clear buffers
'   ReDim DigiPoints(0)
'   ReDim DigiLines(1, 0)
'   ReDim DigiContours(0)
'   ReDim DigiErasePoints(0)
'
   'clear counters
   numDigiPoints = 0
   numDigiLines = 0
   numDigiContours = 0
   numDigiErase = 0
   
   gdm = GDform1.Picture2.DrawMode
   gdw = GDform1.Picture2.DrawWidth
    
   GDform1.Picture2.DrawMode = 13
   GDform1.Picture2.DrawWidth = Max(2, CInt(2 * DigiZoom.LastZoom))
   
   'read the log file
   Do Until EOF(Digilogfilnum%)
      Input #Digilogfilnum%, Xpix, Ypix, zhgt, flag%
      
      If zhgt > MaxColorHeight Then MaxColorHeight = zhgt
      If zhgt < MinColorHeight Then MinColorHeight = zhgt
      
      Select Case flag%
      
          Case 1 'contours
             FirstLineVertex = False
             
             If numDigiContours > 0 Then
                ReDim Preserve DigiContours(numDigiContours)
             Else
                ReDim DigiContours(0)
                End If
             DigiContours(numDigiContours).x = CLng(Xpix)
             DigiContours(numDigiContours).Y = CLng(Ypix)
             DigiContours(numDigiContours).Z = zhgt
             
             'now draw it
             If Not DigiEditPoints Then
                GDform1.Picture2.PSet (CLng(DigiContours(numDigiContours).x * DigiZoom.LastZoom), CLng(DigiContours(numDigiContours).Y * DigiZoom.LastZoom)), TraceColor
                
                If numDigiContours > 0 Then
                     DisPix = Sqr((Xpix - DigiContours(numDigiContours - 1).x) ^ 2# + (Ypix - DigiContours(numDigiContours - 1).Y) ^ 2#)
                     
                     If DisPix <= numPixelSep Then
                        GDform1.Picture2.Line (CLng(DigiContours(numDigiContours - 1).x * DigiZoom.LastZoom), CLng(DigiContours(numDigiContours - 1).Y * DigiZoom.LastZoom))-(CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), TraceColor
                        End If
                        
                   End If
                End If
                
             numDigiContours = numDigiContours + 1
          
          Case 2 'points
             FirstLineVertex = False
             
             If numDigiPoints > 0 Then
                ReDim Preserve DigiPoints(numDigiPoints)
             Else
                ReDim DigiPoints(0)
                End If
             DigiPoints(numDigiPoints).x = CLng(Xpix)
             DigiPoints(numDigiPoints).Y = CLng(Ypix)
             DigiPoints(numDigiPoints).Z = zhgt
             
             'now draw it
             GDform1.Picture2.Line (CLng(DigiPoints(numDigiPoints).x * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(numDigiPoints).Y * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(DigiPoints(numDigiPoints).x * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(numDigiPoints).Y * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), PointColor& 'TraceColor
             GDform1.Picture2.Line (CLng(DigiPoints(numDigiPoints).x * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(numDigiPoints).Y * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(DigiPoints(numDigiPoints).x * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(numDigiPoints).Y * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), PointColor& 'TraceColor
             
            'write the elevation value if zoomm >= 1
            If CInt(DigiZoom.LastZoom) >= 1# Then
               GDform1.Picture2.CurrentX = DigiPoints(numDigiPoints).x * DigiZoom.LastZoom + Max(4, CInt(DigiZoom.LastZoom))
               GDform1.Picture2.CurrentY = DigiPoints(numDigiPoints).Y * DigiZoom.LastZoom
               GDform1.Picture2.FontSize = CInt(8 * DigiZoom.LastZoom)
               GDform1.Picture2.Font = "Ariel"
               GDform1.Picture2.ForeColor = PointColor&
               GDform1.Picture2.Print str$(zhgt)
               End If
             
             numDigiPoints = numDigiPoints + 1
             
          Case 3 'first vertex of line
             FirstLineVertex = True
             
             If numDigiLines > 0 Then
                ReDim Preserve DigiLines(1, numDigiLines)
             Else
                ReDim DigiLines(1, 0)
                End If
             DigiLines(0, numDigiLines).x = CLng(Xpix)
             DigiLines(0, numDigiLines).Y = CLng(Ypix)
             DigiLines(0, numDigiLines).Z = zhgt
             
            
          Case 4 'second vertex of line (must occur right after case 3)
             If FirstLineVertex Then
                FirstLineVertex = False
                
                DigiLines(1, numDigiLines).x = CLng(Xpix)
                DigiLines(1, numDigiLines).Y = CLng(Ypix)
                DigiLines(1, numDigiLines).Z = zhgt
                
                'now draw the line
                If Not DigiEditPoints Then
                    GDform1.Picture2.PSet (CLng(DigiLines(0, numDigiLines).x * DigiZoom.LastZoom), CLng(DigiLines(0, numDigiLines).Y * DigiZoom.LastZoom)), LineColor& 'TraceColor
                    GDform1.Picture2.Line (CLng(DigiLines(0, numDigiLines).x * DigiZoom.LastZoom), CLng(DigiLines(0, numDigiLines).Y * DigiZoom.LastZoom))-(CLng(DigiLines(1, numDigiLines).x * DigiZoom.LastZoom), CLng(DigiLines(1, numDigiLines).Y * DigiZoom.LastZoom)), LineColor& 'TraceColor
                    End If
                    
                numDigiLines = numDigiLines + 1
                
                End If
                
          Case 5 'erase a point
             FirstLineVertex = False
             
             If numDigiErase > 0 Then
                ReDim Preserve DigiErasePoints(numDigiErase)
             Else
                ReDim DigiErasePoints(0)
                End If
             DigiErasePoints(numDigiErase).x = CLng(Xpix)
             DigiErasePoints(numDigiErase).Y = CLng(Ypix)
          
'             XX = DigiErasePoints(numDigiErase).x
'             Yy = DigiErasePoints(numDigiErase).y
        
             If Not DigiEditPoints Then
                'retrieve original RGB color
                If Not DigiGDIfailed Then
                   ier = oGestionImageSrc.GetPixelRGB(Xpix, Ypix, iRouge, iVert, iBleu)
                Else
                   ier = GetSimplePixelRGB(GDform1.Picture2, Xpix, Ypix, iRouge, iVert, iBleu)
                   End If
                
                If ier = 0 Then
                    'restore original RBG color within square area defined by brush size
                    GDform1.Picture2.PSet (CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), RGB(Int(iRouge), Int(iVert), Int(iBleu))
                    End If
                End If
                 
             numDigiErase = numDigiErase + 1
             
          Case Else
             'read error, so close file and exit loop
             Screen.MousePointer = vbDefault
             Close #Digilogfilnum%
             Exit Do

      End Select

   Loop
   
   GDform1.Picture2.DrawMode = gdm
   GDform1.Picture2.DrawWidth = gdw
   
   Close #Digilogfilnum%
   
   GDMDIform.StatusBar1.Panels(1).Text = sEmpty

   Screen.MousePointer = vbDefault
   
   'reenable blinking
   GDMDIform.CenterPointTimer.Enabled = True
   ce& = 0 'reset blinking cursor flag
   
   InitDigiGraph = True
   
   On Error GoTo 0
   Exit Sub
   

InputDigiLogFile_Error:
    Screen.MousePointer = vbDefault
    If Digilogfilnum% > 0 Then Close #Digilogfilnum%
    Digilogfilnum% = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InputDigiLogFile of Module modGDModule"
   'reenable blinking
   GDMDIform.CenterPointTimer.Enabled = True
   ce& = 0 'reset blinking cursor flag
    
End Sub

Public Function InputGuideLines() As Integer

    Dim FirstVertex As Boolean
    Dim X1, X2

'    'stop blinking search points for 1:50000 maps
'    GDMDIform.CenterPointTimer.Enabled = False
'    ce& = 0 'reset flag that draws blinking cursor
    
    'draw the lines to be recalled when the Rubber Sheeting button is pushed
'    pos% = InStr(picnam$, ".")
'    picext$ = Mid$(picnam$, pos% + 1, 3)
    GuideLineFilname$ = App.Path & "\" & RootName(picnam$) & "-RSG" & ".txt"
    If Dir(GuideLineFilname$) <> sEmpty Then
        filnum% = FreeFile
        Open GuideLineFilname$ For Input As #filnum%
        
        gddm = GDform1.Picture2.DrawMode
        gddw = GDform1.Picture2.DrawWidth
        GDform1.Picture2.DrawMode = 13
        GDform1.Picture2.DrawWidth = 1
        
        Do Until EOF(filnum%)
           Input #filnum%, X1, Y1, flag%
           If flag% = 1 Then
              Input #filnum%, X2, Y2, flag%
              If flag% = 2 Then 'draw the line
                 GDform1.Picture2.Line (X1 * DigiZoom.LastZoom, Y1 * DigiZoom.LastZoom)-(X2 * DigiZoom.LastZoom, Y2 * DigiZoom.LastZoom), QBColor(12)
                 End If
              End If
        Loop
        
        Close #filnum%
        GDform1.Picture2.DrawMode = gddm
        GDform1.Picture2.DrawWidth = gddw
        End If

End Function
'The following function will return the inverse tangent in the proper
'quadrant determined by the signs of x and y.
'http://computer-programming-forum.com/16-visual-basic/f6b1e67cca79ee85.htm
Function Atan2(x As Double, Y As Double) As Double
'-pi < Atan2 <= pi
    Const PI As Double = 3.14159265358979
    If x > 0 Then Atan2 = Atn(Y / x): Exit Function     '1st & 4th quadrants
    If x < 0 And Y > 0 Then Atan2 = Atn(Y / x) + PI: Exit Function      '2nd quadrant
    If x < 0 And Y < 0 Then Atan2 = Atn(Y / x) - PI: Exit Function      '3rd quadrant
    If x = 0 And Y > 0 Then Atan2 = PI / 2: Exit Function
    If x = 0 And Y < 0 Then Atan2 = -PI / 2
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
      DACOS = -Atn(XX / Sqr(-XX * XX + 1#)) + PI / 2
      End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : NearNeighbors
' Author    : Dr-John-K-Hall
' Date      : 2/19/2015
' Purpose   : checks if there are contour traces near neighbors in the 8 directions
'---------------------------------------------------------------------------------------
'
Public Function NearNeighbors(Pic As PictureBox, x, Y) As Boolean

   On Error GoTo NearNeighbors_Error
   
   Dim Xpix, Ypix
   Dim R As Long
   Dim Test_Color As couleur
   Dim CP As Boolean
   Dim numPixelstoCheck As Integer
   Dim i As Integer
   
   numPixelstoCheck = 2

   'check to the right
   For i = 1 To numPixelstoCheck
      Xpix = x + i * DigiZoom.LastZoom
      Ypix = Y
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i
       
   'check 45 degrees
   For i = 1 To numPixelstoCheck
      Xpix = x + i * DigiZoom.LastZoom
      Ypix = Y - i * DigiZoom.LastZoom
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i

               
   'check 90 degrees
   For i = 1 To numPixelstoCheck
      Xpix = x
      Ypix = Y - i * DigiZoom.LastZoom
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i

               
   'check 135 degrees
   For i = 1 To numPixelstoCheck
      Xpix = x - i * DigiZoom.LastZoom
      Ypix = Y - i * DigiZoom.LastZoom
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i

               
   'check 180 degrees
   For i = 1 To numPixelstoCheck
      Xpix = x - i * DigiZoom.LastZoom
      Ypix = Y * DigiZoom.LastZoom
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i

               
   'check 225 degrees
   For i = 1 To numPixelstoCheck
      Xpix = x - i * DigiZoom.LastZoom
      Ypix = Y + i * DigiZoom.LastZoom
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i

               
   'check 270 degrees
   For i = 1 To numPixelstoCheck
      Xpix = x
      Ypix = Y + i * DigiZoom.LastZoom
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i

               
   'check 315 degrees
   For i = 1 To numPixelstoCheck
      Xpix = x + i * DigiZoom.LastZoom
      Ypix = Y + i * DigiZoom.LastZoom
      GoSub CheckPoint
      If CP Then
         NearNeighbors = True
         Exit Function
         End If
   Next i

           
   NearNeighbors = False 'if got here, it is false

   On Error GoTo 0
   Exit Function
   
CheckPoint:
    R = Pic.Point(CLng(Xpix), CLng(Ypix))
    If R <> -1 Then
       Test_Color = recupcouleur(R)
       If Test_Color.R = TracingColor.R And _
          Test_Color.V = TracingColor.V And _
          Test_Color.b = TracingColor.b Then
          CP = True
       Else
          CP = False
          End If
    Else
       CP = False
       End If
Return

NearNeighbors_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NearNeighbors of Module modGDModule"
End Function
'---------------------------------------------------------------------------------------
' Procedure : GetSimplePixelRGB
' Author    : Dr-John-K-Hall
' Date      : 2/25/2015
' Purpose   : Reads pixel color from random access fie containing that info
'---------------------------------------------------------------------------------------
'
Public Function GetSimplePixelRGB(Pic As PictureBox, x, Y, iRouge As Byte, iVert As Byte, iBleu As Byte) As Integer

   Dim RecNumber As Long
   Dim NewColorEnum As ColorEnum
   Dim ier As Integer
   
   ier = 0

   On Error GoTo GetSimplePixelRGB_Error

    'figure out record number
    RecNumber = (Y - 1) * Pic.ScaleWidth + x
    Get Picfilnum%, RecNumber, NewColorEnum
    iRouge = NewColorEnum.RedColor
    iVert = NewColorEnum.GreenColor
    iBleu = NewColorEnum.BlueColor
    
    GetSimplePixelRGB = ier
       
   On Error GoTo 0
   Exit Function

GetSimplePixelRGB_Error:

    ier = -1
    GetSimplePixelRGB = ier
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSimplePixelRGB of Class Module CGestionImage"

End Function

'use this sub to center a child form in the middle of the parent form
Public Sub sCenterForm(tmpF As Form)

'centers a form in the middle of the programs main form

Dim x As Integer, Y As Integer
   
   On Error GoTo sCenterForm_Error

   x = GDMDIform.left + 0.5 * GDMDIform.Width - 0.5 * tmpF.Width
   Y = GDMDIform.top + 0.5 * GDMDIform.Height - 0.5 * tmpF.Height

tmpF.Move x, Y

   On Error GoTo 0
   Exit Sub

sCenterForm_Error:

End Sub

Public Function ConvertCoordToString(CoordIn) As String

Dim Coord As Double, CoordAbs As Double
Dim Deg As Integer, MinDeg As Integer, SecDeg As Integer
Dim DegStr As String, MinStr As String, SecStr As String

    signval = Sgn(val(CoordIn))
    
    Coord = Abs(val(CoordIn))
    If Coord > 180 Then 'this is not degrees
       ConvertCoordToString = CoordIn
       Exit Function
       End If
       
    Deg = Fix(Coord)
    MinDeg = Fix((Coord - Deg) * 60)
    SecDeg = CInt(((Coord - Deg) * 60 - MinDeg) * 60) 'don't keep fractional second
      
    If SecDeg = 60 Then
       MinDeg = MinDeg + 1
       SecDeg = 0
       End If
    
    If MinDeg = 60 Then
       Deg = Deg + 1
       MinDeg = 0
       End If
       
    DegStr = str$(Deg)
    If Deg = 0 Then DegStr = "00"
       
    If MinDeg = 0 Then
       MinStr = "00"
    Else
       MinStr = Trim$(str$(MinDeg))
       If Len(MinStr) = 1 Then MinStr = "0" & MinStr
       End If
       
    If SecDeg = 0 Then
       SecStr = "00"
    Else
       SecStr = Trim$(str$(SecDeg))
       If Len(SecStr) = 1 Then SecStr = "0" & SecStr
       End If
    
    If signval = 1 Or signval = 0 Then
       ConvertCoordToString = DegStr & "-" & MinStr & "-" & SecStr
    Else
       ConvertCoordToString = "-" & DegStr & "-" & MinStr & "-" & SecStr
       End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : ConvertDegToNumber
' Author    : Dr-John-K-Hall
' Date      : 3/11/2015
' Purpose   : Convert coordinate string into Deg-Min-Sec format or back to number if already a number
'---------------------------------------------------------------------------------------
'
Public Function ConvertDegToNumber(CoordIn As String) As String
    
    Dim Coord() As String
    Dim CoordNew As String
    Dim SignStr$
    
   On Error GoTo ConvertDegToNumber_Error
   
    CoordNew = Trim$(CoordIn)
    
    If Abs(val(CoordNew)) > 180 Then
       'this is not degrees
       ConvertDegToNumber = CoordNew
       Exit Function
       End If
    
    'detect and treat negative numbers
    SignStr$ = Mid$(Trim$(CoordNew), 1, 1)
    If SignStr$ = "-" Then
       'remove it noew, add it later
       CoordNew = Mid$(Trim$(CoordIn), 2, Len(CoordIn) - 1)
    Else
       SignStr$ = sEmpty
       End If
    
    Coord = Split(CoordNew, "-")
    If UBound(Coord) = 0 Then
      'convert back to coordinate format
      ConvertDegToNumber = ConvertCoordToString(CoordIn)
      Exit Function
    ElseIf UBound(Coord) = 1 Then
      ConvertDegToNumber = Format(str$(val(Coord(0)) + val(Coord(1)) / 60#), "####.#######0")
    ElseIf UBound(Coord) = 2 Then
      ConvertDegToNumber = Format(str$(val(Coord(0)) + val(Coord(1)) / 60# + val(Coord(2)) / 3600), "####.#######0")
      End If
      
   ConvertDegToNumber = SignStr$ & ConvertDegToNumber

   On Error GoTo 0
   Exit Function

ConvertDegToNumber_Error:

      MsgBox "Input is not in the required format:" _
             & "(Hint: Input degrees,minutes,seconds and separate them by ''-'')", _
             vbOKOnly + vbExclamation, "Input Error"
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : JustConvertDegToNumber
' Author    : Dr-John-K-Hall
' Date      : 3/11/2015
' Purpose   : Only converts in one direction, i.e., from deg-min-sec string to a number string
'---------------------------------------------------------------------------------------
'
Public Function JustConvertDegToNumber(CoordIn As String) As String
    Dim Coord() As String
    Dim CoordNew As String
    Dim SignStr$
    
   On Error GoTo JustConvertDegToNumber_Error
   
    CoordNew = Trim$(CoordIn)
    
    If Abs(val(CoordNew)) > 180 Then
       'this is not degrees
       JustConvertDegToNumber = CoordNew
       Exit Function
       End If
    
    'detect and treat negative nunbers
    SignStr$ = Mid$(Trim$(CoordNew), 1, 1)
    If SignStr$ = "-" Then
       'remove it noew, add it later
       CoordNew = Mid$(Trim$(CoordIn), 2, Len(CoordIn) - 1)
    Else
       SignStr$ = sEmpty
       End If
    
    Coord = Split(CoordNew, "-")
    If UBound(Coord) = 0 Then
      'do nothing, just return
      JustConvertDegToNumber = SignStr$ & CoordNew
      Exit Function
    ElseIf UBound(Coord) = 1 Then
      JustConvertDegToNumber = Format(str$(val(Coord(0)) + val(Coord(1)) / 60#), "####.#######0")
    ElseIf UBound(Coord) = 2 Then
      JustConvertDegToNumber = Format(str$(val(Coord(0)) + val(Coord(1)) / 60# + val(Coord(2)) / 3600), "####.#######0")
      End If
      
   JustConvertDegToNumber = SignStr$ & JustConvertDegToNumber

   On Error GoTo 0
   Exit Function

JustConvertDegToNumber_Error:

      MsgBox "Input is not in the required format:" _
             & "(Hint: Input degrees,minutes,seconds and separate them by ''-'')", _
             vbOKOnly + vbExclamation, "Input Error"

End Function

'---------------------------------------------------------------------------------------
' Procedure : ReDrawMap
' Author    : Dr-John-K-Hall
' Date      : 3/12/2015
' Purpose   : clears the picture box and repaints it at the last zoom
'---------------------------------------------------------------------------------------
'
Public Function ReDrawMap(mode As Integer) As Integer

   On Error GoTo ReDrawMap_Error

    Screen.MousePointer = vbHourglass
    
    GDMDIform.StatusBar1.Panels(1).Text = "Please wait, refreshing the map..."
    
    'stop blinking search points for 1:50000 maps
    GDMDIform.CenterPointTimer.Enabled = False
    ce& = 0: CenterBlinkState = False 'these flags keep from erasing the last position since the picture has been repainted
    
    GDform1.Picture2.Cls
    GDform1.Picture2.Picture = LoadPicture(picnam$)
    If DigiZoom.LastZoom <> 1 Then Call PictureBoxZoom(GDform1.Picture2, 0, 0, 0, 0, mode) 'zoom out according to last zoom
    
    GDMDIform.StatusBar1.Panels(1).Text = sEmpty
    
    Screen.MousePointer = vbDefault
    
    'reenable blinking
    GDMDIform.CenterPointTimer.Enabled = True
    ce& = 0 'reset blinking cursor flag
    
    ReDrawMap = 0
    
   On Error GoTo 0
   Exit Function

ReDrawMap_Error:
    
    Screen.MousePointer = vbDefault

    ReDrawMap = -1

End Function
'---------------------------------------------------------------------------------------
' Procedure : EraseSweepPoints
' Author    : Dr-John-K-Hall
' Date      : 3/19/2015
' Purpose   : sweep erasure of all digitization withing the rectangular region ReCoord
'---------------------------------------------------------------------------------------
'
Public Function EraseSweepPoints(RectCoord() As POINTAPI) As Integer

   Dim ier As Integer
   Dim X1, Y1, X2, Y2
   Dim DisPix As Double

   On Error GoTo EraseSweepPoints_Error

   Select Case MsgBox("This erasure can't be reversed!" _
                      & vbCrLf & "" _
                      & vbCrLf & "Proceed?" _
                      , vbOKCancel Or vbQuestion Or vbDefaultButton1, "Sweep erasure notice")
   
    Case vbOK
       'proceed with the sweeping
    Case vbCancel
    
        EraseSweepPoints = 0
        Exit Function
   
   End Select
   
   X1 = RectCoord(0).x
   Y1 = RectCoord(0).Y
   X2 = RectCoord(1).x
   Y2 = RectCoord(1).Y
   
   Screen.MousePointer = vbHourglass
   GDMDIform.StatusBar1.Panels(1).Text = "Please wait, sweeping away unwanted digitized points..."
   
   If DigiLogFileOpened Then
      Close #Digilogfilnum%
      Digilogfilnum% = 0
      DigiLogFileOpened = False
      End If
    
'    pos% = InStr(picnam$, ".")
'    picext$ = Mid$(picnam$, pos% + 1, 3)
    DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
    Digilogfilnum% = FreeFile
    myfile$ = Dir(DigiLogfilnam$)
    If myfile$ <> sEmpty Then
       Open DigiLogfilnam$ For Input As #Digilogfilnum%
    Else
       Screen.MousePointer = vbDefault
       Exit Function 'no log file found for this map file
       End If
   
   'clear buffers
   ReDim DigiPoints(0)
   ReDim DigiLines(1, 0)
   ReDim DigiContours(0)
   ReDim DigiErasePoints(0)
'
   'clear counters
   numDigiPoints = 0
   numDigiLines = 0
   numDigiContours = 0
   numDigiErase = 0
   
   'read the log file
   Do Until EOF(Digilogfilnum%)
      Input #Digilogfilnum%, Xpix, Ypix, zhgt, flag%
      
      Select Case flag%
      
          Case 1 'contours
             FirstLineVertex = False
                
             If Xpix >= X1 And Xpix <= X2 And Ypix >= Y1 And Ypix <= Y2 Then
                'don't include this contour point
             Else
                If numDigiContours > 0 Then
                   ReDim Preserve DigiContours(numDigiContours)
                Else
                   ReDim DigiContours(0)
                   End If
             
                DigiContours(numDigiContours).x = Xpix
                DigiContours(numDigiContours).Y = Ypix
                DigiContours(numDigiContours).Z = zhgt
             
                numDigiContours = numDigiContours + 1
                
                End If
          
          Case 2 'points
             FirstLineVertex = False
             
             If Xpix >= X1 And Xpix <= X2 And Ypix >= Y1 And Ypix <= Y2 Then
                'don't include this contour point
             Else
             
                If numDigiPoints > 0 Then
                   ReDim Preserve DigiPoints(numDigiPoints)
                Else
                   ReDim DigiPoints(0)
                   End If
                   
                DigiPoints(numDigiPoints).x = Xpix
                DigiPoints(numDigiPoints).Y = Ypix
                DigiPoints(numDigiPoints).Z = zhgt
             
                numDigiPoints = numDigiPoints + 1
                
                End If
             
          Case 3 'first vertex of line
          
             If Xpix >= X1 And Xpix <= X2 And Ypix >= Y1 And Ypix <= Y2 Then
                'don't include this contour point
                FirstLineVertex = False
             Else
                FirstLineVertex = True
             
                If numDigiLines > 0 Then
                   ReDim Preserve DigiLines(1, numDigiLines)
                Else
                   ReDim DigiLines(1, 0)
                   End If
                DigiLines(0, numDigiLines).x = Xpix
                DigiLines(0, numDigiLines).Y = Ypix
                DigiLines(0, numDigiLines).Z = zhgt
                
                End If
            
          Case 4 'second vertex of line (must occur right after case 3)
          
             If Xpix >= X1 And Xpix <= X2 And Ypix >= Y1 And Ypix <= Y2 Then
                'don't include this contour point
                If FirstLineVertex Then FirstLineVertex = False
             Else
                If FirstLineVertex Then
                   FirstLineVertex = False
                   
                   DigiLines(1, numDigiLines).x = Xpix
                   DigiLines(1, numDigiLines).Y = Ypix
                   DigiLines(1, numDigiLines).Z = zhgt
                
                   numDigiLines = numDigiLines + 1
                
                   End If
                 End If
                
          Case 5 'erase a point
             FirstLineVertex = False
             
             If Xpix >= X1 And Xpix <= X2 And Ypix >= Y1 And Ypix <= Y2 Then
                'don't include this contour point
             Else
                If numDigiErase > 0 Then
                   ReDim Preserve DigiErasePoints(numDigiErase)
                Else
                   ReDim DigiErasePoints(0)
                   End If
                DigiErasePoints(numDigiErase).x = Xpix
                DigiErasePoints(numDigiErase).Y = Ypix
        
                numDigiErase = numDigiErase + 1
                
                End If
             
          Case Else
             'read error, so close file and exit loop
             Exit Do

      End Select

   Loop
   
   Close #Digilogfilnum%
   
   'now store the changes in the digilogfile
   UpdateDigiLogFile
   
   ier = ReDrawMap(0)
   If Not InitDigiGraph Then
      InputDigiLogFile 'load up saved digitizing data for the current map sheet
   Else
      ier = RedrawDigiLog
      End If
      
    If DigitizeMagvis Then
       DoEvents
       Ret = SetWindowPos(GDDigiMagfrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       End If
      
   
   GDMDIform.StatusBar1.Panels(1).Text = sEmpty

   Screen.MousePointer = vbDefault
   
   EraseSweepPoints = 0
   
   On Error GoTo 0
   Exit Function

EraseSweepPoints_Error:

    Screen.MousePointer = vbDefault
    EraseSweepPoints = -1
    
End Function
'determines screen coordinates of client control
Public Sub ConvClientToScreen(Cont As Control, R As RECT)
    Dim pt2 As POINTAPI
    ClientToScreen Cont.hwnd, pt2
    R.X1 = pt2.x
    R.X2 = R.X1 + Cont.Width / Screen.TwipsPerPixelX
    R.Y1 = pt2.Y
    R.Y2 = R.Y1 + Cont.Height / Screen.TwipsPerPixelY
End Sub

Public Sub KeyDown(KCC As KeyCodeConstants)
    'following Key Down emulations are from: http://www.developerfusion.com/code/274/simulating-keyboard-events/
    keybd_event KCC, 0, 0, 0
End Sub

Public Sub KeyUp(KCC As KeyCodeConstants)
    keybd_event KCC, 0, KEYEVENTF_KEYUP, 0
End Sub

Public Sub KeyPress(KCC As KeyCodeConstants)
    KeyDown KCC
    KeyUp KCC
End Sub

Public Sub ShiftOnn()
    KeyDown vbKeyShift
End Sub

Public Sub ShiftOff()
    KeyUp vbKeyShift
End Sub

Public Sub CtrlOnn()
    KeyDown vbKeyControl
End Sub

Public Sub CtrlOff()
    KeyUp vbKeyControl
End Sub
Rem Check if using Vista, returns 0 or the set correction in pixels
Public Function GetVista() As Integer
    Dim myOS As OSVERSIONINFOEX
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    GetVista = myOS.dwMajorVersion
End Function
'---------------------------------------------------------------------------------------
' Procedure : RedrawDigiLog
' Author    : Dr-John-K-Hall
' Date      : 4/27/2015
' Purpose   : redraws the stored digitized points in memory
'---------------------------------------------------------------------------------------
'
Public Function RedrawDigiLog() As Integer

   On Error GoTo RedrawDigiLog_Error
   
   'on définit la couleur du pixel courant à partir des pixels alentours
   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer
   
   Dim numPixelSep As Integer ' connect contour points separated up to this amount of pixels
   Dim DisPix As Double, i&, ier
   
   Dim color_line As Long
   
   ier = 0
   
   Screen.MousePointer = vbHourglass
   GDMDIform.StatusBar1.Panels(1).Text = "Please wait, loading and plotting stored digitized data."
   
   'stop blinking search points for 1:50000 maps
   GDMDIform.CenterPointTimer.Enabled = False
   ce& = 0 'reset flag that draws blinking cursor
   
   numPixelSep = numDistContour * Sqr(2)
   TraceColor = ContourColor& 'QBColor(12)
   
   gdm = GDform1.Picture2.DrawMode
   gdw = GDform1.Picture2.DrawWidth
    
   GDform1.Picture2.DrawMode = 13
   GDform1.Picture2.DrawWidth = Max(2, CInt(2 * DigiZoom.LastZoom))
   
   'read the digitized data stored in memory
   
   'first contours
   If numDigiContours > 0 Then 'And Not DigiEditPoints Then
   
      For i& = 0 To numDigiContours - 1
             
        Xpix = DigiContours(i&).x
        Ypix = DigiContours(i&).Y
        
        If LineElevColors& = 1 And numcpt > 0 And Not DigitizeContour Then
            'determine color
            colornum% = ((DigiContours(i&).Z - MinColorHeight) / (MaxColorHeight - MinColorHeight)) * UBound(cpt, 2) + 1
            color_line = RGB(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))
        
            'now draw the contour
            GDform1.Picture2.PSet (CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), color_line
            
            If i& > 0 Then
                 DisPix = Sqr((Xpix - DigiContours(i& - 1).x) ^ 2# + (Ypix - DigiContours(i& - 1).Y) ^ 2#)
                 
                 If DisPix <= numPixelSep Then
                    GDform1.Picture2.Line (CLng(DigiContours(i& - 1).x * DigiZoom.LastZoom), CLng(DigiContours(i& - 1).Y * DigiZoom.LastZoom))-(CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), color_line
                    End If
                    
               End If
        
        Else  'don't draw with height colors
        
            GDform1.Picture2.PSet (CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), TraceColor
        
            If i& > 0 Then
                 DisPix = Sqr((Xpix - DigiContours(i& - 1).x) ^ 2# + (Ypix - DigiContours(i& - 1).Y) ^ 2#)
                 
                 If DisPix <= numDistContour Then 'numPixelSep Then
                    GDform1.Picture2.Line (CLng(DigiContours(i& - 1).x * DigiZoom.LastZoom), CLng(DigiContours(i& - 1).Y * DigiZoom.LastZoom))-(CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), TraceColor
                    End If
                    
               End If
               
            End If
          
      Next i&
      
      End If
          
    'next lines
    If numDigiLines > 0 Then 'And Not DigiEditPoints Then
    
       For i& = 0 To numDigiLines - 1
       
         If LineElevColors& = 1 And numcpt > 0 And Not DigitizeContour Then
            'determine color
            colornum% = ((DigiLines(0, i&).Z - MinColorHeight) / (MaxColorHeight - MinColorHeight)) * UBound(cpt, 2) + 1
            color_line = RGB(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))
        
            'now draw the line
            GDform1.Picture2.PSet (CLng(DigiLines(0, i&).x * DigiZoom.LastZoom), CLng(DigiLines(0, i&).Y * DigiZoom.LastZoom)), color_line
            GDform1.Picture2.Line (CLng(DigiLines(0, i&).x * DigiZoom.LastZoom), CLng(DigiLines(0, i&).Y * DigiZoom.LastZoom))-(CLng(DigiLines(1, i&).x * DigiZoom.LastZoom), CLng(DigiLines(1, i&).Y * DigiZoom.LastZoom)), color_line
        
         ElseIf Not DigitizeContour Then
            'now draw the line with the standard color
            GDform1.Picture2.PSet (CLng(DigiLines(0, i&).x * DigiZoom.LastZoom), CLng(DigiLines(0, i&).Y * DigiZoom.LastZoom)), LineColor& 'TraceColor
            GDform1.Picture2.Line (CLng(DigiLines(0, i&).x * DigiZoom.LastZoom), CLng(DigiLines(0, i&).Y * DigiZoom.LastZoom))-(CLng(DigiLines(1, i&).x * DigiZoom.LastZoom), CLng(DigiLines(1, i&).Y * DigiZoom.LastZoom)), LineColor& 'TraceColor
            
         ElseIf DigitizeContour Then
            'now draw the line with the standard color
            GDform1.Picture2.PSet (CLng(DigiLines(0, i&).x * DigiZoom.LastZoom), CLng(DigiLines(0, i&).Y * DigiZoom.LastZoom)), TraceColor
            GDform1.Picture2.Line (CLng(DigiLines(0, i&).x * DigiZoom.LastZoom), CLng(DigiLines(0, i&).Y * DigiZoom.LastZoom))-(CLng(DigiLines(1, i&).x * DigiZoom.LastZoom), CLng(DigiLines(1, i&).Y * DigiZoom.LastZoom)), TraceColor
            End If
            
       Next i&
    
       End If
       
    'next points
    If numDigiPoints > 0 Then
    
       For i& = 0 To numDigiPoints - 1
       
         Xpix = DigiPoints(i&).x
         Ypix = DigiPoints(i&).Y
         zhgt = DigiPoints(i&).Z
       
        'now draw it
        GDform1.Picture2.Line (CLng(DigiPoints(i&).x * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(i&).Y * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom)))-(CLng(DigiPoints(i&).x * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(i&).Y * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom))), PointColor& 'TraceColor
        GDform1.Picture2.Line (CLng(DigiPoints(i&).x * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(i&).Y * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)))-(CLng(DigiPoints(i&).x * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)), CLng(DigiPoints(i&).Y * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom))), PointColor& 'TraceColor
         
        'write the elevation value if zoomm >= 1
        If CInt(DigiZoom.LastZoom) >= 1# Then
           GDform1.Picture2.CurrentX = DigiPoints(i&).x * DigiZoom.LastZoom + Max(4, CInt(DigiZoom.LastZoom))
           GDform1.Picture2.CurrentY = DigiPoints(i&).Y * DigiZoom.LastZoom
           GDform1.Picture2.FontSize = CInt(8 * DigiZoom.LastZoom)
           GDform1.Picture2.Font = "Ariel"
           GDform1.Picture2.ForeColor = PointColor&
           GDform1.Picture2.Print str$(zhgt)
           End If
               
       Next i&
    
       End If
       
    'now erase points
    If numDigiErase > 0 And Not DigiEditPoints Then
    
       For i& = 0 To numDigiErase - 1
          Xpix = DigiErasePoints(i&).x
          Ypix = DigiErasePoints(i&).Y
          
          'retrieve original RGB color
          If Not DigiGDIfailed Then
             ier = oGestionImageSrc.GetPixelRGB(Xpix, Ypix, iRouge, iVert, iBleu)
          Else
             ier = GetSimplePixelRGB(GDform1.Picture2, Xpix, Ypix, iRouge, iVert, iBleu)
             End If
          
            If ier = 0 Then
                'restore original RBG color within square area defined by brush size
                GDform1.Picture2.PSet (CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), RGB(Int(iRouge), Int(iVert), Int(iBleu))
                End If
          
       Next i&
       
       End If
   
   GDform1.Picture2.DrawMode = gdm
   GDform1.Picture2.DrawWidth = gdw
   
   GDMDIform.StatusBar1.Panels(1).Text = sEmpty

   Screen.MousePointer = vbDefault
   
   'reenable blinking
   GDMDIform.CenterPointTimer.Enabled = True
   ce& = 0 'reset blinking cursor flag
   
   RedrawDigiLog = ier

   On Error GoTo 0
   Exit Function

RedrawDigiLog_Error:

    ier = -1
    GDMDIform.StatusBar1.Panels(1).Text = sEmpty
    Screen.MousePointer = vbDefault
    'reenable blinking
    GDMDIform.CenterPointTimer.Enabled = True
    ce& = 0 'reset blinking cursor flag
    
    RedrawDigiLog = ier

End Function

'---------------------------------------------------------------------------------------
' Procedure : InitDigiPointsImage
' Author    : Dr-John-K-Hall
' Date      : 4/27/2015
' Purpose   : 'writes the file used for point editing
'---------------------------------------------------------------------------------------
Public Function InitDigiPointsImage() As Integer
   
   Dim PointImage As Byte, ier As Integer, pos%, TmpImageFile$, picext$
   Dim i As Long, j As Long, NumDigiTotal As Long
   Dim Xpix As Long, Ypix As Long
   Dim RecNum&
   
   On Error GoTo InitDigiPointsImage_Error
   
   Byte0 = 0
   Byte1 = 1
   
   ier% = 0
   
   
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
   
   If Not ImagePointFile Then 'try creating buffer file now, maybe user lower the UAC settings
   
'       pos% = InStr(picnam$, ".")
'       picext$ = Mid$(picnam$, pos% + 1, 3)
       TmpImageFile$ = App.Path & "\" & RootName(picnam$) & "-IMG" & ".buf"
       
       If Dir(TmpImageFile$) = sEmpty Then
       
           filnumImage% = FreeFile
           Open TmpImageFile$ For Random Access Read Write As #filnumImage% Len = Len(PointImage)
        
           Call UpdateStatus(GDMDIform, 1, 0)
           GDMDIform.StatusBar1.Panels(1).Text = "Creating point digitizing buffer, please wait..."
           
           For j = 1 To pixhi
           
               For i = 1 To pixwi
               
               RecNum& = i + (j - 1) * pixwi
           
               Put #filnumImage%, RecNum&, Byte0
                   
        '           DoEvents
               Next i
               Call UpdateStatus(GDMDIform, 1, 100 * j / pixhi)
               DoEvents
           Next j
           
           ImagePointFile = True
           
       Else
       
           filnumImage% = FreeFile
           Open TmpImageFile$ For Random Access Read Write As #filnumImage% Len = Len(PointImage)
           
           ImagePointFile = True

           End If
       
       End If
   
   'now write the bytes for digitized points, lines
   If numDigiPoints > 0 Or numDigiContours > 0 Or numDigiLines > 0 Or numDigiErase > 0 And ImagePointFile Then
   
      If numDigiPoints > 0 Then
         NumDigiTotal = numDigiPoints + 1
         End If
         
      If numDigiContours > 0 Then
         NumDigiTotal = NumDigiTotal + numDigiContours + 1
         End If
         
      If numDigiLines > 0 Then
         NumDigiTotal = NumDigiTotal + numDigiLines + 1
         End If
         
      If numDigiErase > 0 Then
         NumDigiTotal = NumDigiTotal + numDigiErase + 1
         End If
      
      'first contours
   
      For i& = 0 To numDigiContours - 1
             
        Xpix = CLng(DigiContours(i&).x)
        Ypix = CLng(DigiContours(i&).Y)
        
        ier = RecordDigiPointsImage(Xpix, Ypix, 1)
        
        Call UpdateStatus(GDMDIform, 1, 100 * (i + 1) / (NumDigiTotal + 4))
         
      Next i&
          
      'next points
    
      For i& = 0 To numDigiPoints - 1
       
         Xpix = CLng(DigiPoints(i&).x)
         Ypix = CLng(DigiPoints(i&).Y)
        
         ier = RecordDigiPointsImage(Xpix, Ypix, 2)
        
         Call UpdateStatus(GDMDIform, 1, 100 * (i + 2 + numDigiContours) / NumDigiTotal)
               
       Next i&
    
       'next lines
       For i& = 0 To numDigiLines - 1
       
        Xpix = CLng(DigiLines(0, i&).x)
        Ypix = CLng(DigiLines(0, i&).Y)
        
        ier = RecordDigiPointsImage(Xpix, Ypix, 3)
        
        Xpix = DigiLines(1, i&).x
        Ypix = DigiLines(1, i&).Y
        
        ier = RecordDigiPointsImage(Xpix, Ypix, 4)
        
        Call UpdateStatus(GDMDIform, 1, 100 * (i + 3 + numDigiContours + numDigiPoints) / NumDigiTotal)
        
       Next i&
       
      'now erase points
    
       For i& = 0 To numDigiErase - 1
       
          Xpix = CLng(DigiErasePoints(i&).x)
          Ypix = CLng(DigiErasePoints(i&).Y)
          
          ier = RecordDigiPointsImage(Xpix, Ypix, 5)
       
          Call UpdateStatus(GDMDIform, 1, 100 * (i + 4 + numDigiLines + numDigiContours + numDigiPoints) / NumDigiTotal)
          
       Next i&
       
      GDMDIform.picProgBar.Visible = False
      GDMDIform.StatusBar1.Panels(1) = sEmpty
      GDMDIform.StatusBar1.Panels(2) = sEmpty
   
   ElseIf numDigiPoints = 0 And numDigiContours = 0 And numDigiLines = 0 And numdigieraes = 0 And ImagePointFile Then
   
      'open the DIG file and read in content
      
      Screen.MousePointer = vbHourglass
      
'      pos% = InStr(picnam$, ".")
'      picext$ = Mid$(picnam$, pos% + 1, 3)
      DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
      Digilogfilnum% = FreeFile
      myfile$ = Dir(DigiLogfilnam$)
      If myfile$ <> sEmpty Then
         Open DigiLogfilnam$ For Input As #Digilogfilnum%
      Else
         Screen.MousePointer = vbDefault
         ier = -1
         InitDigiPointsImage = ier
         Exit Function 'no log file found for this map file
         End If
      
      
       'read the log file
       Do Until EOF(Digilogfilnum%)
          Input #Digilogfilnum%, Xpix, Ypix, zhgt, flag%
          
          Select Case flag%
          
              Case 1 'contours
              
                 ier = RecordDigiPointsImage(Xpix, Ypix, 1)
              
              Case 2 'points
              
                 ier = RecordDigiPointsImage(Xpix, Ypix, 2)
                 
              Case 3 'first vertex of line
              
                 ier = RecordDigiPointsImage(Xpix, Ypix, 3)
                
              Case 4 'second vertex of line (must occur right after case 3)
              
                 ier = RecordDigiPointsImage(Xpix, Ypix, 4)
                    
              Case 5 'erase a point
              
                 ier = RecordDigiPointsImage(Xpix, Ypix, 5)
                 
              Case Else
                 'read error, so close file and exit loop
                 Screen.MousePointer = vbDefault
                 Close #Digilogfilnum%
                 Exit Do
    
          End Select
    
       Loop
       
       Close #Digilogfilnum%
       DigiLogFileOpened = False
       
       Screen.MousePointer = vbDefault
          
       GDMDIform.StatusBar1.Panels(1) = sEmpty
       GDMDIform.StatusBar1.Panels(2) = sEmpty
       
       ImagePointFile = True
       
       End If
  
   ier = 0
   InitDigiPointsImage = ier

   On Error GoTo 0
   Exit Function

InitDigiPointsImage_Error:
    
    ier = -1
    Screen.MousePointer = vbDefault
    InitDigiPointsImage = ier

End Function
'---------------------------------------------------------------------------------------
' Procedure : RedrawDigiPoints
' Author    : Dr-John-K-Hall
' Date      : 4/27/2015
' Purpose   : 'restores a rectangular region around the old point using the rgb colors
'             records the new point in the byte file and then replots the points in that region
'             Position of edited point is (Xpix, Ypix)
'             EditMode = 0 reserved
'             EditMode = 1 contours
'             EditMode = 2 points
'             EditMode = 3 'frist vertex of line
'             EditMode = 4 'second vertex of line
'             EditMode = 5 'erased points
'             mode = 0 'replace point and replot
'             mode = 1 'kill point and replot

'---------------------------------------------------------------------------------------
'
Public Function RedrawDigiPoints(Xpix As Long, Ypix As Long, EditMode As Integer, mode As Integer) As Integer

   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer
   
   Dim XpixTest As Long, YpixTest As Long
   
   Dim ier As Integer, RecNum&
   Dim i As Long, j As Long
   Dim SearchRegionSize
   Dim Byte0 As Byte
   
   Dim DigiByte As Byte
   
'   HighLightColor defined in function HighLightPoint

   On Error GoTo RedrawDigiPoints_Error
   
   ier = 0
   
   Byte0 = 0
   
   SearchRegionSize = DigiSearchRegion
   
   If GeoMap Then
   
      Screen.MousePointer = vbHourglass
   
      Select Case EditMode
         Case 0 'blank
            DigiByte = 0
         Case 1 'contour point
            DigiByte = 1
            'now find the old countour point in the digicontours array, and change the coordinates
            For i = 0 To numDigiContours - 1
               If CLng(XpixLast) = CLng(DigiContours(i).x) And CLng(YpixLast) = CLng(DigiContours(i).Y) Then
                  iPoint = i
                  Exit For
                  End If
            Next i
            
         Case 2 'digitied point
            DigiByte = 2
            'now find the old point in the digipoints array, and change the coordinates
            For i = 0 To numDigiPoints - 1
               If CLng(XpixLast) = CLng(DigiPoints(i).x) And CLng(YpixLast) = CLng(DigiPoints(i).Y) Then
                  iPoint = i
                  Exit For
                  End If
            Next i
            
         Case 3 'first vertex of line
            DigiByte = 3
            'now find the old vertex in the digilines array, and change the coordinates
            For i = 0 To numDigiLines - 1
               If CLng(XpixLast) = CLng(DigiLines(0, i).x) And CLng(YpixLast) = CLng(DigiLines(0, i).Y) Then
                  iPoint = i
                  Exit For
                  End If
            Next i
         Case 4 'second vertex of line
            DigiByte = 4
            'now find the old vertex in the diglines array, and change the coordinates
            For i = 0 To numDigiLines - 1
               If CLng(XpixLast) = CLng(DigiLines(1, i).x) And CLng(YpixLast) = CLng(DigiLines(1, i).Y) Then
                  iPoint = i
                  Exit For
                  End If
            Next i
         Case 5 'deleted point using eraser
            DigiByte = 5
            'now find the old erased point in the digieraser array, and change the coordinates
            For i = 0 To numDigiLines - 1
               If CLng(XpixLast) = CLng(DigiErasePoints(i).x) And CLng(YpixLast) = CLng(DigiErasePoints(i).Y) Then
                  iPoint = i
                  Exit For
                  End If
            Next i
            
      End Select

      'erase the old point, and redraw it
      'firt remove last highlight
      
      If EditMode = 0 Then
        ier = -1
        RedrawDigiPoints = ier
        Exit Function
        End If
      
      gdm = GDform1.Picture2.DrawMode
      gdw = GDform1.Picture2.DrawWidth
           
      GDform1.Picture2.DrawMode = 7
      GDform1.Picture2.DrawWidth = 2
      
      If EditMode = 1 Then 'contours
      
'          'undo last highlighted plot mark
'          GDform1.Picture2.Circle (CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom)), Max(2, 2 * DigiZoom.LastZoom), HighLightColor
          
'          'undo last highlighted plot mark
'          If XpixLast <> -1 And YpixLast <> -1 Then
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            End If
            
          'restore drawmode
          GDform1.Picture2.DrawMode = gdm
          GDform1.Picture2.DrawWidth = gdw
          
          'now zero old byte position and record new pix position
          If ImagePointFile Then
             
             'erase old byte
             RecNum& = XpixLast + (YpixLast - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
             Put #filnumImage%, RecNum&, Byte0
            
             'add new one
             If mode = 0 Then
                RecNum& = Xpix + (Ypix - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
                Put #filnumImage%, RecNum&, DigiByte
                End If
        
             End If
             
          If mode = 0 Then
            'record new position
            DigiContours(iPoint).x = CLng(Xpix)
            DigiContours(iPoint).Y = CLng(Ypix)
          ElseIf mode = 1 Then
            'kill contour vertex by moving indices up
            For i = iPoint To numDigiContours - 2
               DigiContours(i).x = DigiContours(i + 1).x
               DigiContours(i).Y = DigiContours(i + 1).Y
               DigiContours(i).Z = DigiContours(i + 1).Z
            Next i
            numDigiContours = numDigiContours - 1
            End If
          
          'store changes
          UpdateDigiLogFile
          
          'no need for erasing old highlighted points
          XpixLast = -1
          YpixLast = -1
          
          'replot
          ier = ReDrawMap(0)
          If Not InitDigiGraph Then
             InputDigiLogFile 'load up saved digitizing data for the current map sheet
          Else
             ier = RedrawDigiLog
             End If
      
      ElseIf EditMode = 2 Then
           
'          'undo last highlighted plot mark
'          If XpixLast <> -1 And YpixLast <> -1 Then
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            End If
            
    '      GDform1.Picture2.Circle (CLng(Xpix * DigiZoom.LastZoom), CLng(Ypix * DigiZoom.LastZoom)), 5 * DigiZoom.LastZoom, HighLightColor
                
          'restore drawmode
          GDform1.Picture2.DrawMode = gdm
          GDform1.Picture2.DrawWidth = gdw
'
'          'restore the canvas
'          For i = XpixLast * DigiZoom.LastZoom - SearchRegionSize To XpixLast + SearchRegionSize
'             For j = YpixLast - SearchRegionSize To YpixLast + SearchRegionSize
'
'                If Not DigiGDIfailed Then
'                   ier = oGestionImageSrc.GetPixelRGB(i, j, iRouge, iVert, iBleu)
'                Else
'                   ier = GetSimplePixelRGB(GDform1.Picture2, i, j, iRouge, iVert, iBleu)
'                   End If
'
'                If ier = 0 Then
'                    'restore original RBG color within square area defined by brush size
'                    GDform1.Picture2.PSet (CLng(i * DigiZoom.LastZoom), CLng(j * DigiZoom.LastZoom)), rgb(Int(iRouge), Int(iVert), Int(iBleu))
'                    End If
'
'             Next j
'          Next i
              
          'now zero old byte position and record new pix position
          If ImagePointFile Then
             
             'erase old byte
             RecNum& = XpixLast + (YpixLast - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
             Put #filnumImage%, RecNum&, Byte0
            
             'add new one
             If mode = 0 Then
                RecNum& = Xpix + (Ypix - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
                Put #filnumImage%, RecNum&, DigiByte
                End If
        
             End If
             
          If mode = 0 Then
            'record new position
            DigiPoints(iPoint).x = CLng(Xpix)
            DigiPoints(iPoint).Y = CLng(Ypix)
          Else 'kill point by moving indices up
            For i = iPoint To numDigiPoints - 2
               DigiPoints(i).x = DigiPoints(i + 1).x
               DigiPoints(i).Y = DigiPoints(i + 1).Y
               DigiPoints(i).Z = DigiPoints(i + 1).Z
            Next i
            numDigiPoints = numDigiPoints - 1
            End If
          
          'update log file
          UpdateDigiLogFile
          
          'no need for erasing old highlighted points
          XpixLast = -1
          YpixLast = -1
          
          'replot
          ier = ReDrawMap(0)
          If Not InitDigiGraph Then
             InputDigiLogFile 'load up saved digitizing data for the current map sheet
          Else
             ier = RedrawDigiLog
             End If
          
'          'replot within the erased region and a bit more
'          SearchRegionSize = SearchRegionSize + 2
'
'          For i = 0 To numDigiPoints - 1
'
'            XpixTest = DigiPoints(i).X
'            YpixTest = DigiPoints(i).Y
'
'            If XpixTest >= XpixLast - SearchRegionSize And XpixTest <= XpixLast + SearchRegionSize And _
'               YpixTest >= YpixLast - SearchRegionSize And ypixtext <= YpixLast + SearchRegionSize Then
'
'               GDform1.Picture2.Line (CLng(XpixTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), PointColor& 'TraceColor
'               GDform1.Picture2.Line (CLng(XpixTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), PointColor& 'TraceColor
'
'               'write the elevation value if zoomm >= 1
'               If CInt(DigiZoom.LastZoom) >= 1# Then
'                  GDform1.Picture2.CurrentX = XpixTest * DigiZoom.LastZoom + Max(4, CInt(DigiZoom.LastZoom))
'                  GDform1.Picture2.CurrentY = YpixTest * DigiZoom.LastZoom
'                  GDform1.Picture2.Fontsize = CInt(8 * DigiZoom.LastZoom)
'                  GDform1.Picture2.Font = "Ariel"
'                  GDform1.Picture2.ForeColor = PointColor&
'                  GDform1.Picture2.Print str$(DigiPoints(i).Z)
'                  End If
'
'               End If
'
'          Next i
          
      ElseIf EditMode = 3 Then
      
'          'undo last highlighted plot mark
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
          
          'undo last highlighted plot mark
'          If XpixLast <> -1 And YpixLast <> -1 Then
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            End If
                      
          'restore drawmode
          GDform1.Picture2.DrawMode = gdm
          GDform1.Picture2.DrawWidth = gdw
          
          'now zero old byte position and record new pix position
          If ImagePointFile Then
             
             'erase old byte
             RecNum& = XpixLast + (YpixLast - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
             Put #filnumImage%, RecNum&, Byte0
            
             'add new one
             If mode = 0 Then
                RecNum& = Xpix + (Ypix - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
                Put #filnumImage%, RecNum&, DigiByte
                End If
        
             End If
             
          If mode = 0 Then
            'record point's new position
            DigiLines(0, iPoint).x = Xpix
            DigiLines(0, iPoint).Y = Ypix
          
            If iPoint - 1 > 0 And XpixLast = DigiLines(1, iPoint - 1).x And YpixLast = DigiLines(1, iPoint - 1).Y Then
               'also redo the 2nd vertix of the last line
               
                'record new position of last line's 2nd vertex
                DigiLines(1, iPoint - 1).x = CLng(Xpix)
                DigiLines(1, iPoint - 1).Y = CLng(Ypix)
               
               End If
          ElseIf mode = 1 Then
             'kill point by moving indices up, and also kill the end vertice of this line
            For i = iPoint To numDigiLines - 2
               DigiLines(0, i).x = DigiLines(0, i + 1).x
               DigiLines(0, i).Y = DigiLines(0, i + 1).Y
               DigiLines(0, i).Z = DigiLines(0, i + 1).Z
               DigiLines(1, i).x = DigiLines(1, i + 1).x
               DigiLines(1, i).Y = DigiLines(1, i + 1).Y
               DigiLines(1, i).Z = DigiLines(1, i + 1).Z
            Next i
            numDigiLines = numDigiLines - 1
            End If
          
          'store changes
          UpdateDigiLogFile
          
          'no need for erasing old highlighted points
          XpixLast = -1
          YpixLast = -1
          
          'replot
          ier = ReDrawMap(0)
          If Not InitDigiGraph Then
             InputDigiLogFile 'load up saved digitizing data for the current map sheet
          Else
             ier = RedrawDigiLog
             End If
      
      ElseIf EditMode = 4 Then
      
'          'undo last highlighted plot mark
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
          
          'undo last highlighted plot mark
'          If XpixLast <> -1 And YpixLast <> -1 Then
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            End If
                      
          'restore drawmode
          GDform1.Picture2.DrawMode = gdm
          GDform1.Picture2.DrawWidth = gdw
          
          'now zero old byte position and record new pix position
          If ImagePointFile Then
             
             'erase old byte
             RecNum& = XpixLast + (YpixLast - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
             Put #filnumImage%, RecNum&, Byte0
            
             'add new one
             If mode = 0 Then
                RecNum& = Xpix + (Ypix - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
                Put #filnumImage%, RecNum&, DigiByte
                End If
        
             End If
             
          If mode = 0 Then
            'record new position
            DigiLines(1, iPoint).x = CLng(Xpix)
            DigiLines(1, iPoint).Y = CLng(Ypix)
            
            If iPoint + 1 <= numDigiLines - 1 And XpixLast = DigiLines(0, iPoint + 1).x And YpixLast = DigiLines(0, iPoint + 1).Y Then
               'also redo the 1st vertex of the next line
               
                'record new position of next line's first vertex
                DigiLines(0, iPoint + 1).x = CLng(Xpix)
                DigiLines(0, iPoint + 1).Y = CLng(Ypix)
               
                End If
                
          ElseIf mode = 1 Then
            'delete the vertice by moving the indices, also for first vertex of this line
            For i = iPoint To numDigiLines - 2
               DigiLines(0, i).x = DigiLines(0, i + 1).x
               DigiLines(0, i).Y = DigiLines(0, i + 1).Y
               DigiLines(0, i).Z = DigiLines(0, i + 1).Z
               DigiLines(1, i).x = DigiLines(1, i + 1).x
               DigiLines(1, i).Y = DigiLines(1, i + 1).Y
               DigiLines(1, i).Z = DigiLines(1, i + 1).Z
            Next i
            numDigiLines = numDigiLines - 1
            End If
          
          'store changes
          UpdateDigiLogFile
          
          'no need for erasing old highlighted points
          XpixLast = -1
          YpixLast = -1
          
          'replot
          ier = ReDrawMap(0)
          If Not InitDigiGraph Then
             InputDigiLogFile 'load up saved digitizing data for the current map sheet
          Else
             ier = RedrawDigiLog
             End If
      
      ElseIf EditMode = 5 Then
          
'          'restore last highlighted plot mark
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), YpixLast * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)), HighLightColor 'TraceColor
'          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor   'TraceColor
          
'          'undo last highlighted plot mark
'          If XpixLast <> -1 And YpixLast <> -1 Then
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'            End If
                     
         'restore drawmode
          GDform1.Picture2.DrawMode = gdm
          GDform1.Picture2.DrawWidth = gdw
          
          'now zero old byte position and record new pix position
          If ImagePointFile Then
             
             'erase old byte
             RecNum& = XpixLast + (YpixLast - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
             Put #filnumImage%, RecNum&, Byte0
            
             'add new one
             If mode = 0 Then
                RecNum& = Xpix + (Ypix - 1) * pixwi + 1 'we stored pixel (0,0) in record 1
                Put #filnumImage%, RecNum&, DigiByte
                End If
        
             End If
             
          If mode = 0 Then
            'record point's new position
            DigiErasePoints(iPoint).x = CLng(Xpix)
            DigiErasePoints(iPoint).Y = CLng(Ypix)
          Else
             'kill point by moving indices up, and also kill the end vertice of this line
            For i = iPoint To numDigiErase - 2
               DigiErasePoints(i).x = DigiErasePoints(i + 1).x
               DigiErasePoints(i).Y = DigiErasePoints(i + 1).Y
            Next i
            numDigiErase = numDigiErase - 1
            End If
          
          'store changes
          UpdateDigiLogFile
          
          'no need for erasing old highlighted points
          XpixLast = -1
          YpixLast = -1
          
          'replot
          ier = ReDrawMap(0)
          If Not InitDigiGraph Then
             InputDigiLogFile 'load up saved digitizing data for the current map sheet
          Else
             ier = RedrawDigiLog
             End If
      
          End If
          
       End If

   Screen.MousePointer = vbDefault
    
   RedrawDigiPoints = ier

   On Error GoTo 0
   Exit Function

RedrawDigiPoints_Error:

    ier = -1
    RedrawDigiPoints = ier
    Screen.MousePointer = vbDefault
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RedrawDigiPoints of Class Module CGestionImage"
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : RecordDigiPointsImage
' Author    : Dr-John-K-Hall
' Date      : 4/27/2015
' Purpose   : Records a digitized point as a 1 byte in the g_ImagePoint array at the pixel position of the digitized point
'             mode = 0 for recording erased point using the eraser sweep
'                  = 1 for recording digitized point
'                  = 2 for contour
'                  = 3 for first vertex of line
'                  = 4 for second vertex of line
'                  = 5 for deleted point using the eraser
'---------------------------------------------------------------------------------------
'
Public Function RecordDigiPointsImage(Xpix As Long, Ypix As Long, mode As Integer) As Integer

   Dim RecNum&, ier As Integer
   Dim DigiByte As Byte
   
   Select Case mode
       Case 0 'recording erased point using the eraser sweep
          DigiByte = 0
       Case 1 'recording digitized point
          DigiByte = 1
       Case 2 'recording digitied contour point
          DigiByte = 2
       Case 3 'recording first vertex of line
          DigiByte = 3
       Case 4 'recording second vertex of line
          DigiByte = 4
       Case 5 'recording deleted point using eraser
          DigiByte = 5
   End Select

   On Error GoTo RecordDigiPointsImage_Error
   
   ier = 0

   'check if the image file exists
    If ImagePointFile Then
     
       RecNum& = Xpix + (Ypix - 1) * pixwi + 1 'we placed pixel (0,0) in record 1
       Put #filnumImage%, RecNum&, DigiByte
       
    Else
       
       ier = -1
    
       End If
   
   RecordDigiPointsImage = ier

   On Error GoTo 0
   Exit Function

RecordDigiPointsImage_Error:

    If Err.Number = 52 Or Err.Number = 54 Then 'the file got closed somehow, so reopen it
       TmpImageFile$ = App.Path & "\" & RootName(picnam$) & "-IMG" & ".buf"
   
       filnumImage% = FreeFile
       Open TmpImageFile$ For Random Access Read Write As #filnumImage% Len = Len(DigiByte)
       
       Resume
       
       End If

    ier = -1
    RecordDigiPointsImage = ier
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RecordDigiPointsImage of Class Module CGestionImage"
End Function
'---------------------------------------------------------------------------------------
' Procedure : HighLightPoint
' Author    : Dr-John-K-Hall
' Date      : 4/27/2015
' Purpose   : highlights closest point to cursor position, while dehighlighting any other point
'           mode = 0 for searching blanks
'                = 1 for searching digitized contour points
'                = 2 for searching digitized points
'                = 3 for searching digitized lines (first vertex)
'                = 4 for searching digitized lines (second vertex)
'                = 5 for searching for erased points using the eraser
'           ByteTest - byte value found at this location
'---------------------------------------------------------------------------------------
'
Public Function HighLightPoint(Xpix As Long, Ypix As Long, XpixLast As Long, YpixLast As Long, ByteTest As Byte, mode)

   On Error GoTo HighLightPoint_Error

   Dim ier As Integer
   Dim SearchRegionSize As Long
   Dim i As Long, j As Long
   Dim Dist As Double, DistTest As Double
   Dim iTest As Long, jTest As Long
   Dim RecNum&
   Dim Byte0 As Byte
   Dim Byte1 As Byte
   Dim Byte2 As Byte
   Dim Byte3 As Byte
   Dim Byte4 As Byte
   Dim Byte5 As Byte
   Dim ByteFind As Byte
   Dim DigiByte As Byte
   
   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer
   
   HighLightColor = RGB(10, 238, 242) 'QBColor(14) 'rgb(10, 238, 242)
   
   ier = 0
   
   Byte0 = 0
   Byte1 = 1
   Byte2 = 2
   Byte3 = 3
   Byte4 = 4
   Byte5 = 5
   
   Select Case mode
       Case 0 'highlighting all points
          DigiByte = 0
       Case 1 'highlighting contour point
          DigiByte = 1
       Case 2 'highlighting digitied point
          DigiByte = 2
       Case 3, 4 'highlighting first vertex of line
          DigiByte = 3
'       Case 4 'hightlight second vertex of line
'          DigiByte = 4
'          GDMDIform.StatusBar1.Panels(4).Text = "L2"
       Case 5 'recording deleted point using eraser
          DigiByte = 5
   End Select
   
   iTest = -1
   jTest = -1
   
   SearchRegionSize = DigiSearchRegion 'search 2 * SearchRegionSize x 2 * SearchRegionSize pixels
   
   If ImagePointFile Then
   
      Dist = INIT_VALUE
      
      For i = CLng(Xpix) - SearchRegionSize To CLng(Xpix) + SearchRegionSize
      
         For j = CLng(Ypix) - SearchRegionSize To CLng(Ypix) + SearchRegionSize
         
            RecNum& = i + (j - 1) * pixwi + 1 'we placed pixel (0,0) in record 1
            Get #filnumImage%, RecNum&, ByteFind
            
            If (mode <> 0 And mode <> 3 And mode <> 4 And ByteFind = DigiByte) Or _
               (mode = 0 And ByteFind <> Byte0) Or _
               (mode = 3 Or mode = 4 And ByteFind = Byte3 Or ByteFind = Byte4) Then
               'calculate the distance to it
               DistTest = Sqr((Xpix - i) ^ 2 + (Ypix - j) ^ 2)
               If DistTest < Dist Then
                  iTest = i
                  jTest = j
                  Dist = DistTest
                  ByteTest = ByteFind
                  End If
                  
               End If
             
         Next j
         
     Next i
     
     If mode = 0 And ByteTest = Byte0 Then
       'just blanks
       If XpixLast <> -1 And YpixLast <> -1 Then
           
           gdm = GDform1.Picture2.DrawMode
           gdw = GDform1.Picture2.DrawWidth
           
           GDform1.Picture2.DrawMode = 7
           GDform1.Picture2.DrawWidth = 2
       
          'erase last highlight and return
          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
          GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
          
           'restore drawmode
           GDform1.Picture2.DrawMode = gdm
           GDform1.Picture2.DrawWidth = gdw
          
          XpixLast = -1
          YpixLast = -1
          End If
       
       ier = 0
       HighLightPoint = ier
       Exit Function
       End If
     
     Select Case ByteTest
     
        Case Byte0
          GDMDIform.StatusBar1.Panels(4).Text = "A"
        
        Case Byte1
          GDMDIform.StatusBar1.Panels(4).Text = "C"
        
        Case Byte2
          GDMDIform.StatusBar1.Panels(4).Text = "P"
        
        Case Byte3, Byte4
          GDMDIform.StatusBar1.Panels(4).Text = "L"
        
        Case Byte5
          GDMDIform.StatusBar1.Panels(4).Text = "E"
        
     End Select
     
     If iTest <> -1 And jTest <> -1 Then
        'found nearest point
        'highlight it
        'first restore last point
        If iTest <> XpixLast Or jTest <> YpixLast Then
           'dehighlight last point and highlight this point
           gdm = GDform1.Picture2.DrawMode
           gdw = GDform1.Picture2.DrawWidth
           
           GDform1.Picture2.DrawMode = 7
           GDform1.Picture2.DrawWidth = 2
           
           If (mode = 0 And ByteTest = 1) Or mode = 1 Then
                'undo last marker
'                GDform1.Picture2.Circle (CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom)), Max(2, 2 * DigiZoom.LastZoom), HighLightColor
''                GDform1.Picture2.Line (XpixLast * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom), YpixLast * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom))-(XpixLast * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom), YpixLast * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)), HighLightColor 'TraceColor
''                GDform1.Picture2.Line (XpixLast * DigiZoom.LastZoom, YpixLast * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom))-(XpixLast * DigiZoom.LastZoom, YpixLast * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom)), HighLightColor   'TraceColor
'
'               'draw new marker
'                GDform1.Picture2.Circle (CLng(iTest * DigiZoom.LastZoom), CLng(jTest * DigiZoom.LastZoom)), Max(2, 2 * DigiZoom.LastZoom), HighLightColor
''                GDform1.Picture2.Line (iTest * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom), jTest * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom))-(iTest * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom), jTest * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)), HighLightColor 'TraceColor
''                GDform1.Picture2.Line (iTest * DigiZoom.LastZoom, jTest * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom))-(iTest * DigiZoom.LastZoom, jTest * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom)), HighLightColor   'TraceColor
                If XpixLast <> -1 And YpixLast <> -1 Then
                    'undo last marker
                    GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                    GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                    End If
                    
               'draw new marker
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor    'TraceColor
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
          
           
           ElseIf (mode = 0 And ByteTest = 2) Or mode = 2 Then
                If XpixLast <> -1 And YpixLast <> -1 Then
                    'undo last marker
                     GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                     GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                     End If
                     
               'draw new marker
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor    'TraceColor
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
            
           ElseIf (mode = 0 And (ByteTest = 3 Or ByteTest = 4)) Or mode = 3 Or mode = 4 Then
           
'                'undo last marker
'                GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor       'TraceColor
'                GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor      'TraceColor
'                GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor      'TraceColor
'
'               'draw new marker at latest point
'                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor       'TraceColor
'                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor        'TraceColor
'                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor          'TraceColor
          
                If XpixLast <> -1 And YpixLast <> -1 Then
                    'undo last marker
                     GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                     GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                     End If
                     
               'draw new marker
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor    'TraceColor
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
           
           ElseIf (mode = 0 And ByteTest = 5) Or mode = 5 Then
'                'undo last marker
'                GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'                GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom), CLng(YpixLast * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom))), HighLightColor   'TraceColor
'
'               'draw new marker
'                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
'                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom), CLng(jTest * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom), CLng(jTest * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom))), HighLightColor   'TraceColor
               
                If XpixLast <> -1 And YpixLast <> -1 Then
                    'undo last marker
                     GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                     GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
                     End If
                     
               'draw new marker
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor    'TraceColor
                GDform1.Picture2.Line (CLng(iTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(iTest * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(jTest * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor  'TraceColor
                
                End If
               
           'restore drawmode
           GDform1.Picture2.DrawMode = gdm
           GDform1.Picture2.DrawWidth = gdw
           
           XpixLast = iTest
           YpixLast = jTest
           
           ier = 0
           
           End If
      
     Else
        ier = -2
        End If
   
   Else
   
      ier = -1
      
      End If
      
   HighLightPoint = ier

   On Error GoTo 0
   Exit Function

HighLightPoint_Error:

   ier = -1
   HighLightPoint = ier

End Function

Public Function RootName(FullPath As String) As String
'Input: Name/Full Path of a file
'Returns: Name of file

    Dim sPath As String
    Dim sList() As String
    Dim sAns As String
    Dim iArrayLen As Integer
    Dim pos%

    If Len(FullPath) = 0 Then Exit Function
    sList = Split(FullPath, "\")
    iArrayLen = UBound(sList)
    sAns = IIf(iArrayLen = 0, "", sList(iArrayLen))
    
    If sAns = sEmpty Then sAns = FullPath 'nothing to extract
    
    'now remove the extension
    pos% = InStr(sAns, ".")
    sAns = Mid$(sAns, 1, pos% - 1)
    
    RootName = sAns

End Function
'emulation of Fortran Nint function = round to nearest integer
Public Function Nint(ByVal Number) As Long
      Dim Trial As Long
      Trial = CLng(Number)
      If Abs(Number - Trial) > Abs(Trial + 1 - Number) Then
         Nint = Trial + 1
      Else
         Nint = Trial
         End If
End Function
'source: http://ccm.net/faq/10416-vba-vb6-rounding-function-greater-or-less-than-n-digits
Public Function Rounding(ByVal Number, ByVal DecimalPlace)
      Rounding = Int(Number * 10 ^ DecimalPlace + 1 / 2) / 10 ^ DecimalPlace
End Function
