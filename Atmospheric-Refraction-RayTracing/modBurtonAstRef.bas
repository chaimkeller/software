Attribute VB_Name = "modBurtonAstRef"
Public Const MaxViewAngles& = 501
Public Const MaxViewSteps& = 21001
Public Const TotNumSunAlt = 1001
Public Const NumSuns = 30
'Public Const PI As Double = 3.14159265358979

'Public CNST(1000) As Double
Public WXYZ(4, 82) As Double
Public pi As Double
Public CONV As Double
Public cd As Double
Public cpt() As Integer
Public PaletteLoaded As Boolean
Public MaxTemp As Double, MinTemp As Double
Public nc As Long, NumTemp As Long
Public sEmpty As String

Public CV(2001, 2001, 4) As Double
Public ELV(MaxViewSteps&) As Double, TMP(MaxViewSteps&) As Double, PRSR(MaxViewSteps&) As Double, RCV(82, MaxViewSteps&) As Double
Public IndexRefraction(MaxViewSteps&) As Double
Public ALFA(82, MaxViewSteps&) As Double, ALFT(82, MaxViewSteps&) As Double, SSR(82, MaxViewSteps&) As Double
Public AA(MaxViewSteps&) As Double, AT(MaxViewSteps&) As Double, VRefDeg As Double, CalcSondes As Boolean
Public den As Double, RefCalcType%
Public EDIS(82) As Double, IDCT(10001) As Double, IEND(10001) As Double
Public AIRM(MaxViewSteps&) As Double, ADEN(MaxViewSteps&) As Double, RC As Double
Public SINV As Double, EINV As Double, HGTSCALE As Double, DTINV As Double
Public ALT(NumSuns + 1) As Double, AZM(NumSuns + 1) As Double ', CNST(1000) As Double

Public KSTOP As Long, III As Long, IIS As Long, j As Long, ISSR As Integer, INVFLAG As Integer
Public AMZOBS As Double, BETA As Double, x As Double, DIS As Double
Public Theta As Double, H As Double, UP As Double, XD As Double, HD As Double
Public R1 As Double, R2 As Double, z As Double, ARGT As Double, UsingHSatmosphere As Boolean
Public DZ As Double, AMGAM As Double, xc As Double, YC As Double, ZC As Double
Public ISTND As Long, BETAD As Double, Top As Double, BOT As Double, DAB As Double
Public ZL As Double, ZR As Double, ISTOP As Long, CON As Double, U As Double
Public RMINTMP As Double, RMAXTMP As Double, RMINELV As Double, RMAXELV As Double
Public IJK As Long, MMM As Long, IX As Long, IY As Long, IPER As Long
Public IG As Long, IR As Long, IB As Long, MM As Long, NN As Long, n_size As Long, msize As Long
Public r As Double, g As Double, b As Double, CalcComplete As Boolean
Public DELALT As Double, DELAZM As Double, STARTAZM As Double, TransferCurve() As Variant
Public ROBJ As Double, n As Long, m As Long, KMIN As Long, KMAX As Long, KSTEP As Long
Public GAM As Double, max As Long, HDEG As Double, PPAM As Double, DELTA As Double, XMAX As Double
Public SSRMAX As Double, RMX As Double, TSUN As Double, ITRAN As Long, Image1 As Long, IPLOT As Long
Public SunAngles(NumSuns, TotNumSunAlt) As Long, NumSunAlt(TotNumSunAlt) As Long, HeightStep As Double
Public DirectOut$, FinishedTracing As Boolean, SkipRepeats As Boolean, cmdVDW_error As Integer
Public FileNameAtmOut As String, fileoutatm As Integer, LoopingAtmTracing As Boolean, DateNameAtm As String
Public HillHugging As Boolean, ReNormHeight As Boolean, ZeroRefTesting As Boolean

Public PlotMode As Integer, CalcMode As Integer, SunPlotMode As Integer
Public RayTrace(MaxViewAngles&, MaxViewSteps&) As POINTAPI, NumTraces(MaxViewAngles&) As Long, TracesLoaded As Boolean
Public FilNm As String

'van der Werf global variables
Public AD As Double, AW As Double, BD As Double, BW As Double
Public RELH As Double, AMASSW As Double, RBOLTZ As Double
Public HLIMIT As Double, AMASSD As Double
Public OPTVAP As Double
Public HMAXP1 As Double
Public GRAVC As Double, OBSLAT As Double, DEG2RAD As Double, Rearth As Double
Public PDM1 As Double
Public PDM10 As Double
Public RADCON As Double
Public HCROSS As Double
Public TGROUND As Double
Public KeyPressed As Long
Public OLAT As Double, s2 As Double, MAXIND As Double
Public HL(701) As Double, TL(701) As Double, LRL(701) As Double 'layer heights, temperatures, lapse rates 'arrays used for other atmopsheres
Public LRL0 As Double, TL0 As Double, HL0 As Double, AInv As Double, BInv As Double, CInv As Double
Public TempLoop As Boolean

Public RE As Double
Public Mult As Double, Multiplication As Double, EarthRadius As Single, EarthOrigin As Double
Public PicCenterX As Single, PicCenterY As Single, HOBS As Double, Mult0 As Double
Public Xorigin As Double, Yorigin As Double, RefZoomed As Boolean, PicsResized As Boolean

Public Zoom As Double, twipsx As Double, twipsy As Double, pixwi As Double, pixhi As Double

Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
      
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public fildiag%, DiagnoseIndex As Boolean, DiagnoseIndexHgt As Double

'--------------API constants used for generating terminator (floods region with specified color)-------------
Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Public Type POINTAPI
    x As Double
    y As Double
End Type

Public Type PicZoom
   Left As Long
   Top As Long
   Zoom As Single
   LastZoom As Single
End Type

Public RefZoom As PicZoom

Public nearmouse_digi As POINTAPI

'Public Type zz 'used in Menat ray tracing
' hj(50) As Double
' tj(50) As Double
' pj(50) As Double
' AT(50) As Double
' ct(50) As Double
'End Type

Public Type zz 'used in Menat ray tracing as working array, and as temp array in VDW raytracing
 hj As Double
 tj As Double
 pj As Double
 AT As Double
 ct As Double
End Type

'**********Digitizer flags and parameters**************** '<<<<<<<<<<<<digi changes
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public SizeX, SizeY, direct$, direct2$, buttonstate&(70), numArc&
Public dragbegin As Boolean, dragbox As Boolean, drawbox As Boolean
Public drag1x, drag1y, drag2x, drag2y, magclose As Boolean
Public worldcd%(28)

Public Const INIT_VALUE = 9999999

'-------------------API for mouse control using GTCO digitizer--------------------------
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Public Const MOUSEEVENTF_LEFTDOWN = &H2      ' left button down
Public Const MOUSEEVENTF_LEFTUP = &H4        ' left button up
Public Const MOUSEEVENTF_ABSOLUTE = &H8000   ' absolute move
Public Const MOUSEEVENTF_MOVE = &H1          ' move

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
'----------------screen resoultion variables---------------
   Type RECT
       X1 As Long
       Y1 As Long
       X2 As Long
       Y2 As Long
   End Type

   ' NOTE: The following declare statements are case sensitive.

   Declare Function GetDesktopWindow Lib "user32" () As Long
   Declare Function GetWindowRect Lib "user32" _
      (ByVal hwnd As Long, Rectangle As RECT) As Long
      
      '***********Fancy progress bar global variables and API
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As _
 Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As _
 Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
 ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As _
 Long) As Long
Public pbScaleWidth As Long

'*****************Windows API functions, subroutines and constants*********
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nsize As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nsize As Long) As Long

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

'Van der Werf raytracing dll
Public Declare Function RayTracing Lib "AtmRef.dll" (StarAng As Double, EndAng As Double, StepAng As Double, LastVA As Double, NAngles As Long, _
                                                     DistTo As Double, VAwo As Double, H21 As Double, Tolerance As Double, FileMode As Integer, _
                                                     HOBS As Double, TGROUND As Double, HMAXT As Double, ByVal File_Path As String, StepSize As Integer, _
                                                     GPress As Double, WAVELN As Double, HUMID As Double, OBSLAT As Double, NSTEPS As Long, _
                                                     RecordTLoop As Boolean, TempStart As Double, TempEnd As Double, ByVal pFunc As Long) As Long



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

Private Declare Function lstrcat Lib "kernel32" _
   Alias "lstrcatA" (ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long
   
Private Declare Function SHBrowseForFolder Lib "shell32" _
   (lpBI As BrowseInfo) As Long
   
Private Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long
   
'**********DTM variables**************
Public CHMAP(14, 26) As String * 2, filnumg%
Public CHMNE As String * 2, CHMNEO As String * 2, SF As String * 2

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


      ' FUNCTION: GetScreenResolution()
      '
      ' PURPOSE:
      '   To determine the current screen size or resolution.
      '
      ' RETURN:
      '   The current screen resolution. Typically one of the following:
      '      640 x 480
      '      800 x 600
      '      1024 x 768
      '
      '*****************************************************************
      Function GetScreenResolution() As String
          Dim r As RECT
          Dim hwnd As Long
          Dim retval As Long
          hwnd = GetDesktopWindow()
          retval = GetWindowRect(hwnd, r)
          GetScreenResolution = (r.X2 - r.X1) & "x" & (r.Y2 - r.Y1)
      End Function
      Function GetScreenAspectRatio() As Double
          Dim r As RECT
          Dim hwnd As Long
          Dim retval As Long
          hwnd = GetDesktopWindow()
          retval = GetWindowRect(hwnd, r)
          GetScreenAspectRatio = (r.X2 - r.X1) / (r.Y2 - r.Y1)
      End Function
'---------------------------------------------------------------------------------------
' Procedure : StatusMessage
' Author    : Dr-John-K-Hall
' Date      : 11/27/2018
' Purpose   : handles messages on the status bar
' PanelNumber = panel number starting with 1
' flag = 0 to display new message
'      = 1 to delete last message
'---------------------------------------------------------------------------------------
'
Public Sub StatusMessage(StatusMes As String, PanelNumber As Long, flag As Long)
    
   On Error GoTo StatusMessage_Error
   
   With MDIAtmRef

        If flag = 0 Then 'display new message
        
           .StatusBar.Panels(PanelNumber).Text = sEmpty
           .StatusBar.Panels(PanelNumber).Text = StatusMes
           .StatusBar.Refresh
           
        ElseIf flag = 1 Then 'clear the message panel
        
           .StatusBar.Panels(PanelNumber).Text = sEmpty
           .StatusBar.Refresh
        
           End If
           
    End With

   On Error GoTo 0
   Exit Sub

StatusMessage_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure StatusMessage of Form MDIAtmRef"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FileDialog
' Author    : Dr-John-K-Hall
' Date      : 11/27/2018
' Purpose   : handles the commondialog
' FNM is the output filename
' FilePath is the path to look in, should be terminated with "\"
' FileFilter is in format ".ext", is default filter of the filename that is being searched, it must have "."
' flag - not yet implemented
'---------------------------------------------------------------------------------------
'
Public Sub FileDialog(DialogTit As String, FNM As String, FilePath As String, FiledOpened As String, FileFilter As String, flag As Long, ier As Long)

   On Error GoTo FileDialog_Error
   
   With MDIAtmRef

      .comdlg.CancelError = True
      
      If FileFilter <> sEmpty Then
      
         'check format
         If Len(FileFilter) > 4 Then
            Call MsgBox("FileFilter parameter is not correct" & vbCrLf & "In routine FileDialog", vbCritical, "CommonDialog  caaling error")
            
            ier = -2
            Exit Sub
         Else
            Dim pos%
            pos% = InStr(FileFilter, ".")
            If pos% <> 1 Then
               Call MsgBox("FileFilter parameter is missing '.'" & vbCrLf & "In routine FileDialog", vbCritical, "CommonDialog  caaling error")
               
               ier = -3
               Exit Sub
               End If
            End If
      
         .comdlg.Filter = "Fixed type (" & FileFilter & ")|*" & FileFilter & "|Text Files (*.txt)|*.txt|All files (^.^)|*.*"
         .comdlg.FilterIndex = 1
         
      Else
      
         .comdlg.Filter = "Text Files (*.txt)|*.txt|All files (^.^)|*.*"
         .comdlg.FilterIndex = 1
         
         End If
      
      If FilePath <> sEmpty Then
      
         'check for terminator
         If Mid$(FilePath, Len(FilePath), 1) <> "\" Then
            FilePath = FilePath & "\"
            End If
            
         If FileFilter = sEmpty Then
            .comdlg.filename = FilePath
         Else
            .comdlg.filename = FilePath & "*" & FileFilter
            End If
            
         End If
         
      .comdlg.DialogTitle = DialogTit
      .comdlg.ShowOpen
      FiledOpened = .comdlg.filename
      
      ier = 0
   
   End With
    

   On Error GoTo 0
   Exit Sub

FileDialog_Error:

   ier = -1

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FileDialog of Module modBurtonAstRef"
End Sub

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
Public Function DMAX1(X1 As Double, X2 As Double, X3 As Double) As Double

    Dim TMPEntry As Double
    If X1 >= X2 Then
       TMPEntry = X1
    Else
       TMPEntry = X2
       End If
       
   If TMPEntry >= X3 Then
   Else
      TMPEntry = X3
      End If
      
   DMAX1 = TMPEntry
End Function
      
Public Function RandomNumber(ByVal MaxValue As Long, Optional _
ByVal MinValue As Long = 0)

  On Error Resume Next
  Randomize Timer
  RandomNumber = Int((MaxValue - MinValue + 1) * Rnd) + MinValue

End Function
Public Function GetNextObj() As String
    Dim A As Byte, Res As String
    
    'eliminate whitespace
    Do
        Get #1, , A
    Loop Until A > 32 Or EOF(1)
    
    
    If A = 35 Then     'it's a comment
        Do Until A = 13 Or A = 10
            Res = Res + Chr(A)
            Get #1, , A
        Loop
        GetNextObj = Res
        Exit Function
    End If

    Do Until A <= 32
        Res = Res + Chr(A)
        Get #1, , A
    Loop
    GetNextObj = Res
End Function
Public Sub DoPPM(FName As String, FObject As Object)
    Dim FWidth As Long, FHeight As Long
    Dim FColors As Integer, dX As Long, dy As Long
    Dim str1 As String, FType As String, theHdc As Long
    Dim r As Integer, g As Integer, b As Integer
    
    'déclaration des variables privées
    Dim oGestionImageSrc As New CGestionImage
    Dim oGestionImageDest As New CGestionImage
    Dim iFor1 As Integer 'stocke les valeurs de la boucle For->Next
    Dim iFor2 As Integer 'stocke les valeurs de la boucle For->Next
    Dim iBleu As Byte 'stocke la composante bleue à récupèrer
    Dim iVert As Byte 'stocke la composante verte à récupèrer
    Dim iRouge As Byte 'stocke la composante rouge à récupèrer
    Dim iBleuCouleur As Double 'stocke la composante bleue à appliquer
    Dim iVertCouleur As Double 'stocke la composante verte à appliquer
    Dim iRougeCouleur As Double 'stocke la composante rouge à appliquer
    
    'on définit les contrôles sources et destination
'    Set oGestionImageSrc.PictureBox = pctSource
'    Set oGestionImageDest.PictureBox = BrutonAtmReffm.picRef 'pctDest
    
    On Error GoTo ErrorTrap0
    
    Open FName For Binary As #1
    
    dX = 0: dy = 0
    
    FType = GetNextObj
    If FType <> "P3" And FType <> "P6" Then
        Close #1
        MsgBox "This is not a valid PPM File.", , "Error"
        Exit Sub
    End If
    
    Do
        str1 = GetNextObj
        If Left(str1, 1) <> "#" Then Exit Do
    Loop
    
    AtmRefPicSunfm.Visible = True
    
    FWidth = Val(str1)
    FHeight = Val(GetNextObj)
    FColors = Val(GetNextObj)
    
    theHdc = FObject.hdc
    If FType = "P3" Then
        For dy = 0 To FHeight - 1
            For dX = 0 To FWidth - 1
                r = Val(GetNextObj)
                g = Val(GetNextObj)
                b = Val(GetNextObj)
                '''''''''''''''''''editing "g" part of rgb to fix overly green image''''''EK 121918
                g = g - 68
                If g < 0 Then g = 0
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                SetPixel theHdc, dX, dy, RGB(r, g, b)
'                Call oGestionImageDest.SetPixelRGB(CLng(dx), CLng(dy), CByte(r), CByte(g), CByte(b))
            Next dX
            If dy Mod 16 = 0 Then FObject.Refresh: DoEvents
        Next dy
        Close #1
    Else
        ReDim Scan(0 To FWidth * FHeight * 3 - 1) As Byte
        Get #1, , Scan
        Close #1
        For dy = 0 To FHeight - 1
            For dX = 0 To FWidth - 1
                r = Scan(3 * (dX + dy * FWidth))
                g = Scan(3 * (dX + dy * FWidth) + 1)
                b = Scan(3 * (dX + dy * FWidth) + 2)
'                SetPixel theHdc, dx, dy, RGB(r, g, b)
                Call oGestionImageDest.GetPixelRGB(CLng(dX), CLng(dy), CByte(r), CByte(g), CByte(b))
            Next dX
            If dy Mod 32 = 0 Then FObject.Refresh: DoEvents
        Next dy
        Erase Scan
    End If
    Exit Sub

ErrorTrap0:
    Close #1
    MsgBox "There was an error reading the file:" + vbCrLf + Error$, vbCritical, "Error"
End Sub

'C
''C*********************************************************************
'C
'C
'    Public Sub RADCUR(II As Long, RC As Double, DEN As Double, ByRef ELV() As Double, ByRef TMP() As Double, WL As Double, AtmModel As Integer)
    Public Sub RADCUR(II As Long, RC As Double, den As Double, wl As Double, AtmModel As Integer)
'C
'C   This subroutine calculates the radius of curvature
'C   RC for horizontal rays in the layer II as well as
'C   average density (kg/cm^3).
'C
'       IMPLICIT DOUBLE PRECISION (A-H,O-Z)
'       Option Explicit
       
       Dim PRS As Double, PRW As Double
       Dim XCPPM As Double, SIGMA As Double
       Dim RIAS As Double, RIAXS As Double, RIWS As Double
       Dim DAXS As Double, DW As Double
       Dim DA As Double, DWS As Double, P As Double, T As Double, PW As Double
       Dim DEN1 As Double, RI1 As Double, DEN2 As Double, RI2 As Double
       Dim RKAPPA As Double, DEGC As Double
'C
       If (II > 0) Then
'C   DRY AIR AND WATER VAPOR PARTIAL PRESSURES
          If (AtmModel = 1) Then
            PRS = 101325# 'atmospheric pressure in pascals
            PRW = 0.04 * PRS 'vapor pressure of about 100% saturation at 25 degrees C
          ElseIf AtmModel = 2 Then
            PRS = PRSR(0) * 100# 'pressure in pascals on the ground
            'determine vapor pressure of 100% saturation
            PRW = PRS 'use this pressure to calculate the saturated vaopor pressure at this pressure
            End If
'C
'C   CIDDOR INDEX OF REFRACTION
'C
          XCPPM = 405#  '450# 'ppm CO2
'C   INDEX OF REFRACTION FOR DRY AIR
          SIGMA = 1000# / wl
          RIAS = 1# + 0.00000001 * (5792105# / (238.0185 - SIGMA ^ 2#) + 167917# / (57.362 - SIGMA ^ 2#))
          RIAXS = 1# + (RIAS - 1#) * (1# + 0.000000534 * (XCPPM - 450#))
          RIWS = 1# + 0.00000001 * 1.022 * (295.235 + 2.6422 * SIGMA ^ 2# - 0.03238 * SIGMA ^ 4 + 0.004028 * SIGMA ^ 6#)

'C   BOTTOM OF THE LAYER
          If AtmModel = 1 Then
            Call DENSITY(DAXS, DW, 288.15, 101325#, 0#, XCPPM, 0, AtmModel)
            Call DENSITY(DA, DWS, 293.15, 1333#, 1333#, XCPPM, 1, AtmModel)
            P = PRS * Exp(-ELV(II - 1) / 8400#)
            T = TMP(II - 1)
            PW = PRW * Exp(-ELV(II - 1) / 8400#)
            Call DENSITY(DA, DW, T, P, PW, XCPPM, 2, AtmModel)
            DEN1 = DA + DW
            RI1 = 1# + (DA / DAXS) * (RIAXS - 1#) + (DW / DWS) * (RIWS - 1#)
          ElseIf AtmModel = 2 Then
            Call DENSITY(DAXS, DW, 288.15, 101325#, 0#, XCPPM, 0, AtmModel)
            Call DENSITY(DA, DWS, 293.15, 1333#, 1333#, XCPPM, 1, AtmModel)
            P = PRSR(II - 1) * 100# 'convert to bars
            T = TMP(II - 1)
            PW = P
            Call DENSITY(DA, DW, T, P, PW, XCPPM, 2, AtmModel)
            DEN1 = DA + DW
            RI1 = 1# + (DA / DAXS) * (RIAXS - 1#) + (DW / DWS) * (RIWS - 1#)
            If prjAtmRefMainfm.optMenat.Value = True Then
               RI1 = 1 + 0.000001 * (77.46 + 0.459 * SIGMA * SIGMA) * (PRSR(II - 1) / TMP(II - 1))
               End If
            End If

'C   TOP OF THE LAYER
          If AtmModel = 1 Then
            P = PRS * Exp(-ELV(II) / 8400#)
            T = TMP(II)
            PW = PRW * Exp(-ELV(II) / 8400#)
            Call DENSITY(DA, DW, T, P, PW, XCPPM, 2, AtmModel)
            DEN2 = DA + DW
            RI2 = 1# + (DA / DAXS) * (RIAXS - 1#) + (DW / DWS) * (RIWS - 1#)
          ElseIf AtmModel = 2 Then
            P = PRSR(II) * 100# 'convert to bars
            T = TMP(II)
            PW = P
            Call DENSITY(DA, DW, T, P, PW, XCPPM, 2, AtmModel)
            DEN2 = DA + DW
            RI2 = 1# + (DA / DAXS) * (RIAXS - 1#) + (DW / DWS) * (RIWS - 1#)
            If prjAtmRefMainfm.optMenat.Value = True Then
               RI2 = 1 + 0.000001 * (77.46 + 0.459 * SIGMA * SIGMA) * (PRSR(II) / TMP(II))
               End If
            End If

'C   RADIUS OF CURVATURE
          DELV = ELV(II) - ELV(II - 1)
          RKAPPA = -2# * (RI2 - RI1) / (DELV * (RI2 + RI1))
          RC = 1# / RKAPPA

'C   AVERAGE DENSITY
          den = (DEN1 + DEN2) / 2#
          If AtmModel = 1 Then
            den = (Exp(-ELV(II) / 8400#) + Exp(-ELV(II - 1) / 8400#)) / 2#
          ElseIf AtmModel = 2 Then
            den = (1 / PRW) * 0.5 * (PRSR(II) + PRSR(II - 1))
            End If
          den = 1.225 * den

       Else
          RC = 1E+20
          den = 0#
          End If
'       Return
End Sub
'C
'C*********************************************************************
'C
'C
   Public Sub DENSITY(DA As Double, DW As Double, T As Double, P As Double, PW As Double, XCPPM As Double, IXW As Long, AtmModel As Integer)
'C
'C   This subroutine calculates the density of moist air
'C   as described by P.E. Ciddor, Applied Optics, Vol. 35,
'C   No. 9., March 20, 1996. MKS units.
'C
'C   DA = density of dry air component
'C   DW = density of water vapor component
'C   T = temperature in Kelvin
'C   P = total pressure
'C   PW = partial pressure of water vapor
'C   XCPPM = CO2 content in parts per million
'C   Fractional Humidity = PW/SVP
'C   IXW = flag for pure components
'C
'       IMPLICIT DOUBLE PRECISION (A-H,O-Z)
'      Option Explicit
      
      Dim A As Double, b As Double, C As Double, d As Double, SVP As Double
      Dim ALPHA As Double, BETA As Double, gamma As Double, F As Double
      Dim xw As Double, RMA As Double, RMW As Double, z As Double
      
'C
'C   SATURATION VAPOR PRESSURE OF WATER VAPOR IN AIR
'C
       A = 0.000012378847
       b = -0.019121316
       C = 33.93711047
       d = -6343.1645
       SVP = Exp(A * T * T + b * T + C + d / T)
'C
'C   ENHANCEMENT FACTOR OF WATER VAPOR IN AIR
'C
       ALPHA = 1.00062
       BETA = 0.0000000314
       gamma = 0.00000056
       F = ALPHA + BETA * P + gamma * (T - 273.15) ^ 2
'C
'C   MOLAR FRACTION OF WATER VAPOR IN MOIST AIR
'C
       If (IXW = 0) Then
          xw = 0#
       ElseIf (IXW = 1) Then
          xw = 1#
       Else
          xw = F * PW / P
          If AtmModel = 2 Then
'            xw = (Val(prjAtmRefMainfm.txtHumid.Text) / 100) * SVP * F * PW / P
            xw = (Val(prjAtmRefMainfm.txtHumid.Text) / 100) * SVP * F / P
            End If
          End If
'C
'
'C   MOLAR MASS OF DRY AIR AND MOLAR MASS OF WATER VAPOR
'C
       RMA = 0.001 * (28.9635 + 0.000012011 * (XCPPM - 400#))
       RMW = 0.018015
'C
'C   COMPRESSIBILITIES OF DRY AIR AND PURE WATER VAPOR
'C
       Call COMPRESS(z, T, P, xw)
'C
'C   DENSITY OF THE COMPONENTS OF MOIST AIR
'C
       r = 8.31451
       DA = P * RMA * (1# - xw) / (z * r * T)
       DW = P * RMW * xw / (z * r * T)
'       Return
End Sub
'C
'C*********************************************************************
'C
'C
   Public Sub COMPRESS(comp As Double, T As Double, P As Double, xw As Double)
'C
'C   This subroutine calculates the compressibility of moist air
'C   as described by P.E. Ciddor, Applied Optics, Vol. 35,
'C   No. 9., March 20, 1996. MKS units.
'C
'C   T = temperature in Kelvin
'C   P = total pressure
'C   XW =    molar fraction of water vapor
'C
'       IMPLICIT DOUBLE PRECISION (A-H,O-Z)

'        Option Explicit
        
        Dim A As Double, A1 As Double, A2 As Double, b As Double, B1 As Double
        Dim C As Double, C1 As Double, d As Double, e As Double, ST As Double
'C
       A = 0.00000158123
       A1 = -0.000000029331
       A2 = 0.00000000011043
       b = 0.000005707
       B1 = -0.00000002051
       C = 0.00019898
       C1 = -0.000002376
       d = 0.0000000000183
       e = -0.00000000765
       ST = T - 273.15
       comp = 1# - (P / T) * (A + A1 * ST + A2 * ST ^ 2# + (b + B1 * ST) * xw + (C + C1 * ST) * xw ^ 2#) + (P / T) ^ 2# * (d + e * xw ^ 2#)
'       Return
End Sub
'C
'C  *********************************************************************
'C
'
'C                  APPENDIX  C
'C        MULTILAYER  MODEL  INPUT  FILE
'
'C        170.01D0   1   1   0   -20.0D0 -10.0D0 1500000.0
'C        #HOBS  ITRAN       Image1       IPLOT       STARTALT DELALT XMAX
'C        200    0   0.0 1   81  40
'C        #HEIGHT    WIDTH   PPAM    KMIN    KMAX    KSTEP
''C stmod1.dat
'C        #FILENAME
'
'C        =======================================================================
'C        VALUES DEFAULT DESCRIPTION
'C        =======================================================================
'C        HOBS   5.0    Height of the Observer  (5.0)
'C        ITRAN  1   Transfer Curve :    0 - Stnd. Atmos. Tran. Curv.
'C        1  - Use layer model
'C        2  - Read Transfer Curve
'C        Image1  1   Make a ppm image?
'C        IPLOT  0   Plot ray trajectories?
'C        STARTALT -20.0    Starting Altitude
'C        DELALT 0   Altitude shift of the sun in arcminutes
'C        XMAX   150000  Maximum circumference of the earth to trace ray
'C (ducting)
'C        HEIGHT 100 Image height in pixels
'C        WIDTH  300 Image width in pixels
'C        PPAM   2   Pixels per arcminute
'C        KMIN    1     Wavelength minimum (380nm)
'C        KMAX   81    Wavelength maximum (780nm)
'C        KSTEP    1       Wavelength increment (5nm)
'C        #  Comment line
'C        =======================================================================

'C             APPENDIX  D
'C           COLOR SOURCE CODE
'C
'C   COLOR SUBROUTINES
'C   by Dan Bruton (astro@tamu.edu)
'C   March 27, 1996
'C
'C
'C**********************************************************************
'C
'C   XYZ VALUES FROM ENERGY DISTRIBUTION
'C
'C   The XYZ values are determined by
'C   "integrating" the product of the wavelength distribution of
'C   energy and the XYZ functions.
'C
   Public Sub EXYZ(x As Double, y As Double, z As Double, ByRef EDIS() As Double)
'       IMPLICIT DOUBLE PRECISION (A-H,O-Z)
'        Option Explicit
        
'       Dim WXYZ(3, 81) As Double,

       Dim i As Long, j As Long
'C
'C   CIE Color Matching Functions (x_bar,y_bar,z_bar)
'C   for wavelengths in 5 nm increments from 380 nm to 780 nm.
'C
'      DATA ((WXYZ(I,J),I=1,3),J=1,81)/
'     *       0.0014, 0.0000, 0.0065, 0.0022, 0.0001, 0.0105,
'     *       0.0042, 0.0001, 0.0201, 0.0076, 0.0002, 0.0362,
'     *       0.0143, 0.0004, 0.0679, 0.0232, 0.0006, 0.1102,
'     *       0.0435, 0.0012, 0.2074, 0.0776, 0.0022, 0.3713,
'     *       0.1344, 0.0040, 0.6456, 0.2148, 0.0073, 1.0391,
'     *       0.2839, 0.0116, 1.3856, 0.3285, 0.0168, 1.6230,
'     *       0.3483, 0.0230, 1.7471, 0.3481, 0.0298, 1.7826,
'     *       0.3362, 0.0380, 1.7721, 0.3187, 0.0480, 1.7441,
'     *       0.2908, 0.0600, 1.6692, 0.2511, 0.0739, 1.5281,
'     *       0.1954, 0.0910, 1.2876, 0.1421, 0.1126, 1.0419,
'     *       0.0956, 0.1390, 0.8130, 0.0580, 0.1693, 0.6162,
'     *       0.0320, 0.2080, 0.4652, 0.0147, 0.2586, 0.3533,
'     *       0.0049, 0.3230, 0.2720, 0.0024, 0.4073, 0.2123,
'     *       0.0093, 0.5030, 0.1582, 0.0291, 0.6082, 0.1117,
'     *       0.0633, 0.7100, 0.0782, 0.1096, 0.7932, 0.0573,
'     *       0.1655, 0.8620, 0.0422, 0.2257, 0.9149, 0.0298,
'     *       0.2904, 0.9540, 0.0203, 0.3597, 0.9803, 0.0134,
'     *       0.4334, 0.9950, 0.0087, 0.5121, 1.0000, 0.0057,
'     *       0.5945, 0.9950, 0.0039, 0.6784, 0.9786, 0.0027,
'     *       0.7621, 0.9520, 0.0021, 0.8425, 0.9154, 0.0018,
'     *       0.9163, 0.8700, 0.0017, 0.9786, 0.8163, 0.0014,
'     *       1.0263, 0.7570, 0.0011, 1.0567, 0.6949, 0.0010,
'     *       1.0622, 0.6310, 0.0008, 1.0456, 0.5668, 0.0006,
'     *       1.0026, 0.5030, 0.0003, 0.9384, 0.4412, 0.0002,
'     *       0.8544, 0.3810, 0.0002, 0.7514, 0.3210, 0.0001,
'     *       0.6424, 0.2650, 0.0000, 0.5419, 0.2170, 0.0000,
'     *       0.4479, 0.1750, 0.0000, 0.3608, 0.1382, 0.0000,
'     *       0.2835, 0.1070, 0.0000, 0.2187, 0.0816, 0.0000,
'     *       0.1649, 0.0610, 0.0000, 0.1212, 0.0446, 0.0000,
'     *       0.0874, 0.0320, 0.0000, 0.0636, 0.0232, 0.0000,
'     *       0.0468, 0.0170, 0.0000, 0.0329, 0.0119, 0.0000,
'     *       0.0227, 0.0082, 0.0000, 0.0158, 0.0057, 0.0000,
'     *       0.0114, 0.0041, 0.0000, 0.0081,  0.0029, 0.0000,
'     *       0.0058, 0.0021, 0.0000, 0.0041, 0.0015, 0.0000,
'     *       0.0029, 0.0010, 0.0000, 0.0020, 0.0007, 0.0000,
'     *       0.0014, 0.0005, 0.0000, 0.0010, 0.0004, 0.0000,
'     *       0.0007, 0.0002, 0.0000, 0.0005, 0.0002, 0.0000,
'     *       0.0003, 0.0001, 0.0000, 0.0002, 0.0001, 0.0000,
'     *       0.0002, 0.0001, 0.0000, 0.0001, 0.0000, 0.0000,
'     *       0.0001, 0.0000, 0.0000, 0.0001, 0.0000, 0.0000,
'     *       0.0000, 0.0000, 0.0000/
''     C
       XX = 0#
       YY = 0#
       zz = 0#
       For k = 1 To 81
          XX = XX + EDIS(k) * WXYZ(1, k)
          YY = YY + EDIS(k) * WXYZ(2, k)
          zz = zz + EDIS(k) * WXYZ(3, k)
       Next k
       x = XX
       y = YY
       z = zz
'       Return

'.0014,0.,.0065
'.0022,1e-4,.0105
'.0042,1e-4,.0201
'.0076,2e-4,.0362
'.0143,4e-4,.0679
'.0232,6e-4,.1102
'.0435,.0012,.2074
'.0776,.0022,.3713
'.1344,.004,.6456
'.2148,.0073,1.0391
'.2839,.0116,1.3856
'.3285,.0168,1.623
'.3483,.023,1.7471
'.3481,.0298,1.7826
'.3362,.038,1.7721
'.3187,.048,1.7441
'.2908,.06,1.6692
'.2511,.0739,1.5281
'.1954,.091,1.2876
'.1421,.1126,1.0419
'.0956,.139,.813
'.058,.1693,.6162
'.032,.208,.4652
'.0147,.2586,.3533
'.0049,.323,.272
'.0024,.4073,.2123
'.0093,.503,.1582
'.0291,.6082,.1117
'.0633,.71,.0782
'.1096,.7932,.0573
'.1655,.862,.0422
'.2257,.9149,.0298
'.2904,.954,.0203
'.3597,.9803,.0134
'.4334,.995,.0087
'.5121,1.,.0057
'.5945,.995,.0039
'.6784,.9786,.0027
'.7621,.952,.0021
'.8425,.9154,.0018
'.9163,.87,.0017
'.9786,.8163,.0014
'1.0263,.757,.0011
'1.0567,.6949,.001
'1.0622,.631,8e-4
'1.0456,.5668,6e-4
'1.0026,.503,3e-4
'.9384,.4412,2e-4
'.8544,.381,2e-4
'.7514,.321,1e-4
'.6424,.265,0.
'.5419,.217,0.
'.4479,.175,0.
'.3608,.1382,0.
'.2835,.107,0.
'.2187,.0816,0.
'.1649,.061,0.
'.1212,.0446,0.
'.0874,.032,0.
'.0636,.0232,0.
'.0468,.017,0.
'.0329,.0119,0.
'.0227,.0082,0.
'.0158,.0057,0.
'.0114,.0041,0.
'.0081,.0029,0.
'.0058,.0021,0.
'.0041,.0015,0.
'.0029,.001,0.
'.002,7e-4,0.
'.0014,5e-4,0.
'.001,4e-4,0.
'7e-4,2e-4,0.
'5e-4,2e-4,0.
'3e-4,1e-4,0.
'2e-4,1e-4,0.
'2e-4,1e-4,0.
'1e-4,0.,0.
'1e-4,0.,0.
'1e-4,0.,0.
'0.,0.,0.

End Sub
'C
'C******************************************************************
'C
   Public Sub XYZTORGB(xc As Double, YC As Double, ZC As Double, r As Double, g As Double, b As Double)
'C
'C   This subroutine convert a color from CIE XYZ space
'C   to RGB space.
'C
'       IMPLICIT REAL*8 (a-h,o-z)
'    Option Explicit
    
    Dim XR As Double, yr As Double, XG As Double, YG As Double, XB As Double, YB As Double
    Dim ZR As Double, ZG As Double, Zb As Double
'C
'C   Chromaticity Coordinates for Red, Green, and Blue
'C
       XR = 0.64
       yr = 0.33
       XG = 0.29
       YG = 0.6
       XB = 0.15
       YB = 0.06
       ZR = 1# - (XR + yr)
       ZG = 1# - (XG + YG)
       Zb = 1# - (XB + YB)
'C
'C   PERFORM MATRIX OPERATION
'C
       r = (-XG * YC * Zb + xc * YG * Zb + XG * YB * ZC - XB * YG * ZC - xc * YB * ZG + XB * YC * ZG) / (XR * YG * Zb - XG * yr * Zb - XR * YB * ZG + XB * yr * ZG + XG * YB * ZR - XB * YG * ZR)
       g = (XR * YC * Zb - xc * yr * Zb - XR * YB * ZC + XB * yr * ZC + xc * YB * ZR - XB * YC * ZR) / (XR * YG * Zb - XG * yr * Zb - XR * YB * ZG + XB * yr * ZG + XG * YB * ZR - XB * YG * ZR)
       b = (XR * YG * ZC - XG * yr * ZC - XR * YC * ZG + xc * yr * ZG + XG * YC * ZR - xc * YG * ZR) / (XR * YG * Zb - XG * yr * Zb - XR * YB * ZG + XB * yr * ZG + XG * YB * ZR - XB * YG * ZR)
       If (r <= 0#) Then r = 0#
       If (g <= 0#) Then g = 0#
       If (b <= 0#) Then b = 0#
'       Return
End Sub
'C
'C*********************************************************************
'C
Public Function RAN(ISEED As Long)
'C --------------------------------------------------------------
'C Returns a uniform random deviate between 0.0 and 1.0.
'C Based on: Park and Miller's "Minimal Standard" random number
'C generator(Comm.ACM, 31, 1192, 1988)
'C --------------------------------------------------------------
'      IMPLICIT NONE
      Dim k As Long, IM As Long, IA As Long, IQ As Long, IR As Long
      Dim AM As Double
      IM = 2147483647
      IA = 16807
      IQ = 127773
      IR = 2836
      AM = 128# / IM
      k = ISEED / IQ
      ISEED = IA * (ISEED - k * IQ) - IR * k
      If ISEED < 0 Then ISEED = ISEED + IM
      RAN = AM * (ISEED / 128)
      
End Function

'---------------------------------------------------------------------------------------
' Procedure : PlotRayTracing
' Author    : Dr-John-K-Hall
' Date      : 12/24/2018
' Purpose   : Plots the ray tracing
' FormRef is the name of the form calling the routine
' picRef is the name of the picture on that form to draw the ray tracing on
' cmbSun is the combo box that has a list of suns that are seen above the horizon
' cmbAlt are the view angles of the rays that were traced for those suns
'---------------------------------------------------------------------------------------
'
Public Sub PlotRayTracing(FormRef As Object, picRef As Object, cmbSun As Object, cmbAlt As Object)

Dim colorattributes() As String

On Error GoTo errhand

If NumTemp = 0 Then
   Call MsgBox("You first must perform the raytracing by pressing the ''Calculate'' button", vbExclamation, "Ray Tracing")
   Exit Sub
   End If
   

FormRef.Visible = True

   
'DoPPM App.Path & "\temp.ppm", BrutonAtmReffm.picRef
If Not PaletteLoaded Then
'load rainbow color palette
         myfile = Dir(App.Path & "\rainbow.cpt")
         If myfile = sEmpty Then
            GoTo op850
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
              If Trim$(cc$) <> sEmpty Then
                 If numloop% = 0 Then
                    If Val(cc$) >= numpercent Then
                        num% = Val(cc$)
                        
                        If num% - 1 > UBound(cpt, 2) Then
                           ReDim Preserve cpt(3, UBound(cpt, 2) + 1)
                           End If
                        
                        cpt(0, num% - 1) = Val(cc$)
                        numloop% = 1
                        numpercent = Val(cc$)
                        nowread = True
                    Else
                        nowread = False
                        End If
                 ElseIf numloop% = 1 Then
                    If nowread Then cpt(1, num% - 1) = Val(cc$)
                    numloop% = 2
                 ElseIf numloop% = 2 Then
                    If nowread Then cpt(2, num% - 1) = Val(cc$)
                    numloop% = 3
                 ElseIf numloop% = 3 Then
                    If nowread Then cpt(3, num% - 1) = Val(cc$)
                    numloop% = 0
                    nowread = False
                    Exit For
                    End If
                 End If
                 
                 numcpt = num%
                 
            Next i
         Loop
         Close #filenum%

        PaletteLoaded = True
           
        nc = NumTemp
        
'        nc = (MaxTemp - MinTemp) / NumTemp
        
'        colornum% = (K + 1) * (UBound(cpt, 2) + 1) / nc
'        Color = RGB(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))

      'draw earth image in center
          
     End If
     
op850:

   
picRef.Cls

'now determine temperature extremes and set color palette

Dim xpath0 As Single, xpath As Single, ypath0 As Single, ypath As Single
Dim XP0 As Single, XP As Single, YP As Single, YP0 As Single
Dim BorderColor As Long, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
Dim ANGLE As Double, A1 As Double, A2 As Double, TwipFactor As Double
Dim RecordVertices As Boolean, SunAngle As Single, VA As Single ', MultElv As Double
Dim Refr As Double, TRUANG As Double, InvLabels() As POINTAPI, NumInversions As Integer, InvText() As String
Dim DeltaX As Double, DeltaY As Double, HgtMult As Double
'Dim InversionLayer As Boolean

InversionLayer = False

HgtMult = 1#
If prjAtmRefMainfm.OptionSelby Then HgtMult = 1000#
'If RefCalcType% = 2 Then HgtMult = 1#

PicCenterX = picRef.Width * 0.5 + Xorigin
PicCenterY = picRef.height * 0.5 - Yorigin
EarthRadius = picRef.Width / 10
'Mult = 100000#
If Mult = 0 Then
   Mult = 1#
   RefZoom.LastZoom = Mult
   RefZoom.Zoom = Mult
   End If

If RefZoom.LastZoom = 0 Then
   RefZoom.LastZoom = 1#
   RefZoom.Zoom = 1#
   End If
   
'starting multiplication is that earth radius is 1/20 of picref's width
Multiplication = Mult * EarthRadius / RE 'EarthRadius * 100250 / RE

If CalcMode = 1 Then
'change dimensions of picture2 based on dimensions of earth and three times its atmosphere
'    If Multiplication * RE > picRef.Width / 2 Or Multiplication * RE > picRef.Height / 2 Then
'       Dim ShiftX As Double, ShiftY As Double
'       Dim NewAspect As Double
'       NewAspect = RefZoom.Zoom / RefZoom.LastZoom
'       ShiftX = (NewAspect - 1) * 0.5 * picRef.Width
'       ShiftY = (NewAspect - 1) * 0.5 * picRef.Height
'       picRef.Width = picRef.Width * (1# + NewAspect)
'       picRef.Height = picRef.Height * (1# + NewAspect)
'       picRef.Left = picRef.Left - ShiftX
'       picRef.Top = picRef.Top + ShiftY
'
'       'Determine if the child picture will fill up the screen
'       'If so, there is no need to use scroll bars.
'       FormRef.VScroll1.Visible = (FormRef.Picture2.Height < FormRef.picture1.Height)
'       FormRef.HScroll1.Visible = (FormRef.Picture2.Width < FormRef.picture1.Width)
'
'       FormRef.HScroll1.Max = Maximum(0, 2# * ShiftX)
'       FormRef.VScroll1.Max = Maximum(0, 2# * ShiftY)
'       FormRef.HScroll1.min = 0
'       FormRef.VScroll1.min = 0
'       FormRef.HScroll1.Value = ShiftX
'       FormRef.VScroll1.Value = ShiftY
'       End If
   End If

FormRef.txtStartMult = Mult

'set origin as the observer's height
'center of picture is a observer's x,and y
EarthOrigin = Multiplication * (RE + HOBS) 'RE and HOBS are in meters, as are the ELV of the layers
PicCenterY = PicCenterY + EarthOrigin

picRef.DrawWidth = 1
picRef.FillStyle = 0
'PicRef.FillColor = QBColor(1)

'
''draw earth
'PicRef.Circle (PicCenterX, PicCenterY), EarthRadius, QBColor(1)
'ier = ExtFloodFill(PicRef.hdc, PicCenterX, PicCenterY, QBColor(1), FLOODFILLBORDER)

'draw atmospheric layers using color scalar based on temperature

Screen.MousePointer = vbHourglass

'PicRef.AutoRedraw = False

For i = NumTemp To 1 Step -1
    IShift = 0
    'only plot it if it is visible
'    If PicCenterX + CSng(Multiplication * (RE + ELV(I - 1))) > picRef.Width Then IShift = 1
'    If PicCenterX - CSng(Multiplication * (RE + ELV(I - 1))) < 0 Then IShift = IShift + 1
'    If PicCenterY - Multiplication * HOBS - CSng(Multiplication * (RE + ELV(I - 1))) < 0 Then IShift = IShift + 1
'    If PicCenterY - Multiplication * HOBS + CSng(Multiplication * (RE + ELV(I - 1))) > picRef.Height * Screen.TwipsPerPixelY Then Ishift = Ishift + 1
   
'    If IShift < 3 Then

        colornum& = (TMP(i - 1) - MinTemp) / (MaxTemp - MinTemp) * UBound(cpt, 2) + 1
        BorderColor = RGB(cpt(1, colornum& - 1), cpt(2, colornum& - 1), cpt(3, colornum& - 1))
        picRef.FillColor = BorderColor
        picRef.FillStyle = 0
'        picRef.Circle (PicCenterX, PicCenterY - Multiplication * HOBS), CSng(Multiplication * (RE + ELV(i - 1))), BorderColor
        picRef.Circle (PicCenterX, PicCenterY), CSng(Multiplication * (RE + ELV(i - 1) * HgtMult)), BorderColor
        picRef.FillStyle = 1
        'mark when inversion layer reverses
        If (i > 1) Then 'check for change of slope
           If i = NumTemp Then
              t0 = TMP(i - 1)
              e0 = ELV(i - 1)
              slope0 = 1#
           ElseIf i < NumTemp Then
              T1 = TMP(i - 1)
              e1 = ELV(i - 1)
              If e1 <> e0 Then
                slope1 = (T1 - t0) / (e1 - e0)
              Else
                GoTo skipstep
                End If
              End If
              
'           If TMP(i - 1) > TMP(i) And ELV(i - 1) < 11000 And Not InversionLayer Then
'               InversionLayer = True
'           ElseIf InversionLayer And TMP(i - 1) > TMP(i) And ELV(i - 1) < 11000 Then
'               InversionLayer = False
              'mark the inversion layer
            If i = 19 Then
               ccc = 1
               End If
'               picRef.Circle (PicCenterX, PicCenterY - Multiplication * HOBS), CSng(Multiplication * (RE + ELV(i - 1))), QBColor(1)
            If slope0 * slope1 < 0 And ELV(i - 1) * HgtMult < HCROSS And i < NumTemp And i > 1 Then
               picRef.Circle (PicCenterX, PicCenterY), CSng(Multiplication * (RE + ELV(i - 1) * HgtMult)), QBColor(1)
               NumInversions = NumInversions + 1
               ReDim Preserve InvLabels(NumInversions)
               InvLabels(NumInversions - 1).x = CLng(picRef.Width * 0.5)
               InvLabels(NumInversions - 1).y = CLng(PicCenterY - CSng(Multiplication * (RE + ELV(i - 1) * HgtMult)))
               'if shifted by translation, calculate how much to shift in y
               DeltaX = picRef.Width * 0.5 - PicCenterX
               InvLabels(NumInversions - 1).y = CLng(PicCenterY - Sqr((Multiplication * (RE + ELV(i - 1) * HgtMult)) ^ 2 - DeltaX ^ 2))
               ReDim Preserve InvText(NumInversions)
               InvText(NumInversions - 1) = "Inversion at: " & Str(ELV(i - 1) * HgtMult) & " meters"
               End If
              
             If i < NumTemp Then
                t0 = T1
                e0 = e1
                slope0 = slope1
                End If
                
           End If
       
'       End If
    'mark layers
'    picRef.Circle (PicCenterX, PicCenterY - multiplication * HOBS), CSng(multiplication * twipfactor* (RE + ELV(I - 1))), QBColor(7)

'    picRef.Circle (PicCenterX, PicCenterY), CSng(EarthRadius + Multiplication * ELV(I)), BorderColor

'    Y1 = Multiplication * ELV(I - 1)
'    y3 = Multiplication * ELV(I)
'    y4 = 0.5 * Multiplication * (ELV(I) - ELV(I - 1)) + Y1
'    X1 = PicCenterX - CLng(Multiplication * 0.5 * (ELV(I) + ELV(I - 1)) - EarthRadius)
'    X2 = PicCenterX + CLng(Multiplication * 0.5 * (ELV(I) + ELV(I - 1)) + EarthRadius)
'    Y1 = PicCenterY + CLng(Multiplication * 0.5 * (ELV(I) + ELV(I - 1)) + EarthRadius)
'    Y2 = PicCenterY - CLng(Multiplication * 0.5 * (ELV(I) + ELV(I - 1)) - EarthRadius)
'    picRef.FillColor = CSng(ColorLayer)
'    ExtFloodFill picRef.hdc, CLng(PicCenterX), CLng(Y2), BorderColor, FLOODFILLSURFACE

    'Use the API to fill this region
    'Parameters: (picture box), (x coordinate), (y coordinate), (color being replaced), (fill method - always 1)
'    ExtFloodFill picRef.hdc, X, Y, picDemo.Point(X, Y), 1
'    ExtFloodFill picRef.hdc, CLng(PicCenterX), CLng(Y2), picRef.Point(CLng(plotcenterx), CLng(Y2)), 1
     
'     picRef.Circle (X1, PicCenterY), 10, QBColor(4)
'     picRef.Circle (X2, PicCenterY), 10, QBColor(4)
'     picRef.Circle (PicCenterX, Y1), 10, QBColor(4)
'     picRef.Circle (PicCenterX, Y2), 10, QBColor(4)
'temp = ELV(I - 1)
'If temp = 0 Then
'   picRef.Circle (PicCenterX, PicCenterY - Multiplication * HOBS), CSng(Multiplication * (RE + ELV(I - 1))), QBColor(1)
'   End If
skipstep:
Next i

'draw earth
picRef.FillStyle = 0
picRef.FillColor = QBColor(2)
picRef.Circle (PicCenterX, PicCenterY), Multiplication * RE, QBColor(2)

'now plot pedestal corresponding to observer height
picRef.Line (PicCenterX, PicCenterY - Multiplication * RE)-(PicCenterX, PicCenterY - Multiplication * (RE + HOBS)), QBColor(12)

'now add labels to inversions if any
If NumInversions > 0 Then
   For i = 1 To NumInversions
       picRef.CurrentX = InvLabels(i - 1).x
       picRef.CurrentY = InvLabels(i - 1).y
       
       picRef.FontSize = 12
       picRef.FontName = "Arial"
       picRef.fontcolor = QBColor(1)
       picRef.Print InvText(i - 1)
   Next i
   End If

If PlotMode = 0 Then
  Screen.MousePointer = vbDefault
  Exit Sub
  End If
  
'pltmv:

Screen.MousePointer = vbHourglass
picRef.DrawWidth = 1

LineColor = QBColor(15)

'now plot the ray traces
RecordVertices = True
SunAngle = 0#

If RefCalcType% = 0 Then 'Brutton formulation of raytracing used

   If Not TracesLoaded Then

    If Dir(App.Path & "\test.dat") <> sEmpty Then
       XP0 = 0
       YP0 = 0
       xpath0 = 0
       ypath0 = 0
       filnum% = FreeFile
       Open App.Path & "\test.dat" For Input As #filnum%
       numStep& = 0
       NumViewAngles& = 0
       'only plot the selected view angle
       Do Until EOF(filnum%)
          Input #filnum%, XP, YP, VA
          ANGLE = XP / RE 'angle XP subtends in radians
          A1 = (XP * Cos(ANGLE) + (RE + YP) * Sin(ANGLE)) * Multiplication
          A2 = (-XP * Sin(ANGLE) + (RE + YP) * Cos(ANGLE)) * Multiplication
          DoEvents
    '      If YP = 65.66956 Then
    '         cc = 1
    '         linecolor = QBColor(4)
    '         End If
          If numStep& = 0 Then
             xpath0 = A1
             ypath0 = A2
             RayTrace(NumViewAngles&, numStep&).x = xpath0 / Multiplication
             RayTrace(NumViewAngles&, numStep&).y = ypath0 / Multiplication
             XP0 = XP
             YP0 = YP
             NumTraces(NumViewAngles&) = numStep&
             numStep& = numStep& + 1
             
'             If Val(cmbAlt.Text) = SunAngle And RecordVertices Then 'record the ray coordinates
'                Write #fileout%, A1, A2
'                End If
                
          Else
            If XP >= XP0 Then
               xpath = A1
               ypath = A2
               RayTrace(NumViewAngles&, numStep&).x = xpath / Multiplication
               RayTrace(NumViewAngles&, numStep&).y = ypath / Multiplication
               If NumViewAngles& = cmbAlt.ListIndex Or cmbAlt.ListIndex = cmbAlt.ListCount - 1 Then
                  If ypath <> 0 Then
                    picRef.Line (PicCenterX + xpath0, PicCenterY - ypath0)-(PicCenterX + xpath, PicCenterY - ypath), LineColor
                    DoEvents
                  Else
                    Exit Do
                    End If
                  End If
               xpath0 = xpath
               ypath0 = ypath
               XP0 = XP
               YP0 = YP
               If (numStep& + 1 > MaxViewSteps&) Then
                  Call MsgBox("Exceeded max. array size for ray tracing steps", vbExclamation, "Ray Tracing Plot")
                  Exit Do
                  End If
               NumTraces(NumViewAngles&) = numStep&
               numStep& = numStep& + 1
               
'               If Val(cmbAlt.Text) = SunAngle And RecordVertices Then 'record the ray coordinates
'                  Write #fileout%, A1, A2
'                  End If
                  
            Else
               NumViewAngles& = NumViewAngles& + 1
               If (NumViewAngles& > MaxViewAngles&) Then
                  Call MsgBox("Exceeded max. array size for traced rays", vbExclamation, "Ray Tracing Plot")
                  Exit Do
                  End If
               numStep& = 0
               xpath0 = A1
               ypath0 = A2
               RayTrace(NumViewAngles&, numStep&).x = xpath0 / Multiplication
               RayTrace(NumViewAngles&, numStep&).y = ypath0 / Multiplication
               XP0 = XP
               YP0 = YP
               NumTraces(NumViewAngles&) = numStep&
               numStep& = numStep& + 1
               
'               If Val(cmbAlt.Text) = SunAngle And RecordVertices Then 'finnished recording this view angle, close file
'                  RecordVertices = False
'                  End If
                  
               End If
            End If
    
       Loop
       Close #filnum%
       
'       If RecordVertices Then
'          Close #fileout%
'          End If
       
       TracesLoaded = True
       
       If CalcMode = 1 Then
          FormRef.TabRef.Tab = 4
          End If
       
    Else
    
       Call MsgBox("Can't find trace file test.dat", vbExclamation, "Ray Tracing Plot")
       
       End If
       
   Else 'traces already stored, so use stored traces
   
      If cmbAlt.ListIndex <> cmbAlt.ListCount - 1 Then
      
         If RecordVertices And Val(cmbAlt.Text) = SunAngle Then
            'record one angle
            fileout% = FreeFile
            Open App.Path & "\Bruton_vertices.txt" For Output As #fileout%
            End If
      
         'ALFA(K, J) = (CDbl(N / 2 - J) / PPAM)
         NA = Val(cmbSun.List(cmbSun.ListIndex))
         NumViewAngles& = SunAngles(NA - 1, cmbAlt.ListIndex) - 1
         For i = 0 To NumTraces(NumViewAngles& - 1) - 1
        
            xpath0 = RayTrace(NumViewAngles&, i).x * Multiplication
            ypath0 = RayTrace(NumViewAngles&, i).y * Multiplication
            xpath = RayTrace(NumViewAngles&, i + 1).x * Multiplication
            ypath = RayTrace(NumViewAngles&, i + 1).y * Multiplication
           
            If ypath > 0 Then
                picRef.Line (PicCenterX + xpath0, PicCenterY - ypath0)-(PicCenterX + xpath, PicCenterY - ypath), LineColor
                End If
            
            If RecordVertices And Val(cmbAlt.Text) = SunAngle Then
               Write #fileout%, RayTrace(NumViewAngles&, i).x, RayTrace(NumViewAngles&, i).y - RE
               End If
            
         Next i
         
         If RecordVertices And Val(cmbAlt.Text) = SunAngle Then
            Close #fileout%
            RecordVertices = False
            End If
         
      Else 'chose "All" so display all the angles
         For j = 0 To cmbAlt.ListCount - 2
             NA = Val(cmbSun.List(cmbSun.ListIndex))
             NumViewAngles& = SunAngles(NA - 1, j) - 1
             For i = 0 To NumTraces(NumViewAngles& - 1) - 1
            
                xpath0 = RayTrace(NumViewAngles&, i).x * Multiplication
                ypath0 = RayTrace(NumViewAngles&, i).y * Multiplication
                xpath = RayTrace(NumViewAngles&, i + 1).x * Multiplication
                ypath = RayTrace(NumViewAngles&, i + 1).y * Multiplication
               
                If ypath > 0 Then
                    picRef.Line (PicCenterX + xpath0, PicCenterY - ypath0)-(PicCenterX + xpath, PicCenterY - ypath), LineColor
                    End If
             Next i
         Next j
         End If
      End If
      
   ElseIf RefCalcType% >= 1 Then
      'Hohenkerk and Sinclair formulation of refraction used
      
    If Not TracesLoaded Then
    
'    If RecordVertices Then
'       'record one angle
'       fileout% = FreeFile
'       Open App.Path & "\Bruton_vertices_HS.txt" For Output As #fileout%
'       SunAngle = 0#
'       End If

    If RefCalcType% = 1 Then
       FilNm = App.Path & "\test_HS.dat"
    ElseIf RefCalcType% = 2 Then
       FilNm = App.Path & "\test_M.dat"
    ElseIf RefCalcType% = 3 Then
        'name already stored in FilNm
'       FilNm = App.Path & "\TR_VDW_" & Trim(Str(Fix(TGROUND))) & "_" & Trim(Str(Fix(HOBS))) & "_" & Trim(Str(Fix(OBSLAT))) & ".dat"
       End If

    If Dir(FilNm) <> sEmpty Then
       XP0 = 0
       YP0 = 0
       xpath0 = 0
       ypath0 = 0
       filnum% = FreeFile
       Open FilNm For Input As #filnum%
       numStep& = 0
       NumViewAngles& = 0
       'only plot the selected view angle
       If prjAtmRefMainfm.OptionSelby.Value = True Or prjAtmRefMainfm.chkDucting.Value = vbChecked Then
          'skip doc line
          Line Input #filnum%, doclin$
          End If
       Do Until EOF(filnum%)
          Input #filnum%, XP, PATHLENGTH, YP, VA, TRUANG, Refr
'          If VA < 86.5 Then
'             ccc = 1
'             End If
'          If XP = 330.952223204299 And YP = 4.54424982890487E-02 Then
'             ccc = 1
'             End If
          ANGLE = XP / RE 'angle XP subtends in radians
          A1 = (XP * Cos(ANGLE) + (RE + YP) * Sin(ANGLE)) * Multiplication
          A2 = (-XP * Sin(ANGLE) + (RE + YP) * Cos(ANGLE)) * Multiplication
'          A1 = XP * Multiplication
'          A2 = YP * Multiplication
          DoEvents
    '      If YP = 65.66956 Then
    '         cc = 1
    '         linecolor = QBColor(4)
    '         End If
          If numStep& = 0 Then
             xpath0 = A1
             ypath0 = A2
             RayTrace(NumViewAngles&, numStep&).x = xpath0 / Multiplication
             RayTrace(NumViewAngles&, numStep&).y = ypath0 / Multiplication
             XP0 = XP
             YP0 = YP
             NumTraces(NumViewAngles&) = numStep&
             numStep& = numStep& + 1
             
'             If Val(cmbAlt.Text) = SunAngle And RecordVertices Then 'record the ray coordinates
'                Write #fileout%, A1, A2
'                End If
                
          Else
            If XP >= XP0 Then
               xpath = A1
               ypath = A2
               RayTrace(NumViewAngles&, numStep&).x = xpath / Multiplication
               RayTrace(NumViewAngles&, numStep&).y = ypath / Multiplication
               If NumViewAngles& = cmbAlt.ListIndex Or cmbAlt.ListIndex = cmbAlt.ListCount - 1 Then
                  If ypath > 0 Then
                     picRef.Line (PicCenterX + xpath0, PicCenterY - ypath0)-(PicCenterX + xpath, PicCenterY - ypath), LineColor
                     DoEvents
                  Else
                     Exit Do
                     End If
                  End If
               xpath0 = xpath
               ypath0 = ypath
               bb = cmbAlt.ListIndex
               XP0 = XP
               YP0 = YP
               If (numStep& + 1 > MaxViewSteps&) Then
                  Call MsgBox("Exceeded max. array size for ray tracing steps", vbExclamation, "Ray Tracing Plot")
                  Exit Do
                  End If
               NumTraces(NumViewAngles&) = numStep&
               numStep& = numStep& + 1
               
'               If Val(cmbAlt.Text) = SunAngle And RecordVertices Then 'record the ray's vertices
'                  Write #fileout%, A1, A2
'                  End If
                  
            Else
               NumViewAngles& = NumViewAngles& + 1
               If (NumViewAngles& > MaxViewAngles&) Then
                  Call MsgBox("Exceeded max. array size for traced rays", vbExclamation, "Ray Tracing Plot")
                  Exit Do
                  End If
               numStep& = 0
               xpath0 = A1
               ypath0 = A2
               RayTrace(NumViewAngles&, numStep&).x = xpath0 / Multiplication
               RayTrace(NumViewAngles&, numStep&).y = ypath0 / Multiplication
               XP0 = XP
               YP0 = YP
               NumTraces(NumViewAngles&) = numStep&
               numStep& = numStep& + 1
               
'               If Val(cmbAlt.Text) = SunAngle And RecordVertices Then 'finnished recording this view angle, close file
'                  RecordVertices = False
'                  End If
               
               End If
            End If
    
       Loop
       Close #filnum%
       
'       If RecordVertices Then
'          Close #fileout%
'          End If
       
       TracesLoaded = True
       
       If CalcMode = 1 Then
          FormRef.TabRef.Tab = 4
          End If
       
    Else
    
       Call MsgBox("Can't find trace file: " & vbCrLf & vbCrLf & FilNm, vbExclamation, "Ray Tracing Plot")
       
       End If
       
   Else 'traces already stored, so use stored traces
      If cmbAlt.ListIndex <> cmbAlt.ListCount - 1 Then
         'ALFA(K, J) = (CDbl(N / 2 - J) / PPAM)
         NA = Val(cmbSun.List(cmbSun.ListIndex))
         NumViewAngles& = SunAngles(NA - 1, cmbAlt.ListIndex) '- 1
      
         If RecordVertices And Val(cmbAlt.Text) = SunAngle Then
            'record one angle
            fileout% = FreeFile
            Open App.Path & "\Bruton_vertices_HS.txt" For Output As #fileout%
            End If
               
         For i = 0 To NumTraces(NumViewAngles& - 1) - 1
        
            xpath0 = RayTrace(NumViewAngles& - 1, i).x * Multiplication
            ypath0 = RayTrace(NumViewAngles& - 1, i).y * Multiplication
            xpath = RayTrace(NumViewAngles& - 1, i + 1).x * Multiplication
            ypath = RayTrace(NumViewAngles& - 1, i + 1).y * Multiplication
           
            If ypath <> 0 Then
               picRef.Line (PicCenterX + xpath0, PicCenterY - ypath0)-(PicCenterX + xpath, PicCenterY - ypath), LineColor
               End If
            
            If RecordVertices And Val(cmbAlt.Text) = SunAngle Then
               Write #fileout%, RayTrace(NumViewAngles& - 1, i).x, RayTrace(NumViewAngles&, i).y - RE
               End If
           
         Next i
         
         If RecordVertices And Val(cmdalt.Text) = SunAngle Then
            Close #fileout%
            RecordVertices = False
            End If
            
      Else 'chose "All" so display all the angles
         For j = 0 To cmbAlt.ListCount - 2
             NA = Val(cmbSun.List(cmbSun.ListIndex))
             NumViewAngles& = SunAngles(NA - 1, j) - 1
             For i = 0 To NumTraces(NumViewAngles& - 1) - 1
            
                xpath0 = RayTrace(NumViewAngles& - 1, i).x * Multiplication
                ypath0 = RayTrace(NumViewAngles& - 1, i).y * Multiplication
                xpath = RayTrace(NumViewAngles& - 1, i + 1).x * Multiplication
                ypath = RayTrace(NumViewAngles& - 1, i + 1).y * Multiplication
               
                If ypath <> 0 Then
                    picRef.Line (PicCenterX + xpath0, PicCenterY - ypath0)-(PicCenterX + xpath, PicCenterY - ypath), LineColor
                    End If
             Next i
         Next j
         End If
      End If
      
   
      End If
      
   Dim StatusMes As String
   StatusMes = "Sky depiction complete..."
   Call StatusMessage(StatusMes, 1, 0)
      
   If CalcMode = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
        End If
   
   
   With FormRef
   Screen.MousePointer = vbDefault
   Exit Sub
   
'       .Left = 0
'       .Top = 0
'       .Width = MDIAtmRef.ScaleWidth
'       .Height = MDIAtmRef.ScaleHeight
         If PicsResized Then GoTo pr500


        .Rayfrm.Width = .TabRef.Width - 720
        .Rayfrm.height = .TabRef.height - 720
        .picture1.Width = .Rayfrm.Width - .VScroll1.Width - 10
        .picture1.height = .Rayfrm.height - .HScroll1.height - .cmdLarger.Top - 20
        .VScroll1.Left = .picture1.Left + .picture1.Width + 10
        .HScroll1.Left = .picture1.Left
        .HScroll1.Width = .picture1.Width
        .HScroll1.Top = .picture1.Top + .picture1.height + 10
       
'       .paramfrm.Width = .TabRef.Width - 2 * .paramfrm.Left
'       .paramfrm.Height = .TabRef.Height - .paramfrm.Top - .paramfrm.Left
'       .picture1.Left = 0
'       .picture1.Top = 0
'       .picture1.Width = .paramfrm.Width
'       .picture1.Height = .paramfrm.Height
       
       'Set ScaleMode to pixels (since picture size is in pixels)
       .ScaleMode = vbPixels
       .picture1.ScaleMode = vbPixels
       .Picture2.ScaleMode = vbPixels
       
       pixwi = .picture1.Width
       pixhi = .picture1.height
       
       'Autosize is set to True so that the boundaries of
       'Picture2 are expanded to the size of the actual map
       .Picture2.AutoSize = True
       
       'Set the BorderStyle of each picture box to None.
       'prjAtmRefMainfm.Picture1.BorderStyle = 0
       .Picture2.BorderStyle = 0
        
       'load map to buffer
'       .picRef.Picture = LoadPicture(picnam$)
       
'       RefZoom.Zoom = 1#
'       RefZoom.LastZoom = 1#
       RefZoom.Left = 0
       RefZoom.Top = 0
       
       MDIAtmRef.StatusBar.Panels(3).Text = CInt(100 * RefZoom.LastZoom) & "%"
       
       AAA = .Picture2.Width
       bbb = pixwi
       Cdd = .Picture2.height
       Ddd = pixhi
       
       'Load the default map to the visible picturebox at 100% zoom
       
       .Picture2.Width = CLng(RefZoom.Zoom * pixwi)
       .Picture2.height = CLng(RefZoom.Zoom * pixhi)
       
       AAA = .Picture2.Width
       bbb = pixwi
       Cdd = .Picture2.height
       Ddd = pixhi
       
       If pixwi > .Picture2.Width Then pixwi = .Picture2.Width
       If pixhi > .Picture2.height Then pixhi = .Picture2.height
       
       'check for maps that are larger than the maximum pixel size
       If .Picture2.Width < pixwi Then
          Select Case MsgBox("The picture is wider than the maximum allowed width of: " & .Picture2.Width _
                             & vbCrLf & sEmpty _
                             & vbCrLf & "If you continue working with this zoom level, it will be truncated." _
                             & vbCrLf & sEmpty _
                             & vbCrLf & "Proceed anyways?" _
                             , vbYesNoCancel Or vbExclamation Or vbDefaultButton1 Or vbDefaultButton2, "Error in loading picture")
          
            Case vbYes
          
            Case vbNo, vbCancel
               g_ier = -2
               Exit Sub
          End Select
          End If
          
       If .Picture2.height < pixhi Then
          Select Case MsgBox("The picture is taller than the maximum allowed height of: " & .Picture2.height _
                             & vbCrLf & sEmpty _
                             & vbCrLf & "If you continue working with this map, it will be truncated." _
                             & vbCrLf & sEmpty _
                             & vbCrLf & "Proceed?" _
                             , vbYesNoCancel Or vbExclamation Or vbDefaultButton1 Or vbDefaultButton2, "Error in loading picture")
          
            Case vbYes
          
            Case vbNo, vbCancel
               g_ier = -2
               Exit Sub
          End Select
          End If
          
pr500:
       .picRef.Width = pixwi
       .picRef.height = pixhi
       ier = StretchBlt(.Picture2.hdc, RefZoom.Left, RefZoom.Top, CLng(RefZoom.Zoom * pixwi), CLng(RefZoom.Zoom * pixhi), .picRef.hdc, 0, 0, .picRef.Width, .picRef.height, vbSrcCopy)
       
       'this is how to divide the picture and dump
'       ier = StretchBlt(.Picture2.hdc, RefZoom.left, RefZoom.top, CLng(RefZoom.Zoom * pixwi), CLng(RefZoom.Zoom * pixhi), .PictureBlit.hdc, CLng(RefZoom.Zoom * pixwi) / 2, 0, pixwi, pixhi, vbSrcCopy)
       If ier = 0 Then 'stretchblt failed
    '      use Default
           cc = 1
           'have to draw directly onto picture2
'          .picture2.Picture = LoadPicture(picnam$)
          End If
       If PicsResized Then
          Screen.MousePointer = vbDefault
          Exit Sub
          End If
       
       RefZoom.Left = INIT_VALUE
       RefZoom.Top = INIT_VALUE
       
       'Initialize location of both pictures
       .picture1.Move 0, 0, .ScaleWidth - VScroll1.Width, .ScaleHeight - .HScroll1.height
       .Picture2.Move 0, 0
       
       'Position the horizontal scroll bar
       .HScroll1.Top = .picture1.height
       .HScroll1.Left = .picture1.Left
       .HScroll1.Width = .picture1.Width
       
       'Position the vertical scroll bar
       .VScroll1.Top = 0
       .VScroll1.Left = .picture1.Width
       .VScroll1.height = .picture1.height
       
       'Set the Max property for the scroll bars.
       .HScroll1.max = .Picture2.Width - .picture1.Width
       .VScroll1.max = .Picture2.height - .picture1.height
       
       'Determine if the child picture will fill up the screen
       'If so, there is no need to use scroll bars.
       .VScroll1.Visible = (.picture1.height < .Picture2.height)
       .HScroll1.Visible = (.picture1.Width < .Picture2.Width)
       
       'Initiate Scroll Step Sizes
       .HScroll1.LargeChange = .HScroll1.max / 20
       .HScroll1.SmallChange = .HScroll1.max / 60
          
       .VScroll1.LargeChange = .VScroll1.max / 20
       .VScroll1.SmallChange = .VScroll1.max / 60
   
   End With

   PicsResized = True
   
   Screen.MousePointer = vbDefault

'        If IsMissing(BorderColor) Then
'            ' get color at given coordinates
'            BorderColor = .Point(X, Y)
'            ' change all the pixels with that color
'            ExtFloodFill .hdc, X2, Y2, BorderColor, FLOODFILLSURFACE
'        Else
'            ExtFloodFill .hdc, X2, Y2, BorderColor, FLOODFILLBORDER
'        End If

   Exit Sub
   
errhand:

   If err.Number = 6 Then 'overflow
      MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_Resize of Form BrutonAtmReffm"
      Resume Next
   Else
      Resume Next
      Close
      Screen.MousePointer = vbDefault
      End If
   
End Sub
'---------------------------------------------------------------------------------------
' Procedure : PictureBoxZoom
' Author    : Dr-John-K-Hall
' Date      : 3/1/2015
' Purpose   : stretchblt with mouse wheel to zoom based on two different sources:
'             1. stretchblt method source: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=39776&lngWId=1
'             2. wheel mouse hook: http://www.vbforums.com/showthread.php?388222-VB6-MouseWheel-with-Any-Control-%28originally-just-MSFlexGrid-Scrolling%29
'               two files examples: WheelHook-AllControls.zip, WheelHook-NestedControls.zip
'           mode = 0 'redraw all the lines, etc.
'           mode >= 1 'don't redraw the rubber sheeting points
'           mode >= 2 'dont' redraw the guidelines
'---------------------------------------------------------------------------------------
Public Sub PictureBoxZoom(ByRef picBox As PictureBox, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long, Mode As Integer)
'  picBox.Cls
'  picBox.Print "MouseWheel " & IIf(Rotation < 0, "Down", "Up")
   Dim Bltier As Long
   Dim NewLeft As Single
   Dim NewTop As Single
   Dim NewZoom As Single
   Dim ier As Integer
   Dim AA As Long, bb As Long
'   Dim XCenter As Single, YCenter As Single 'center of zoomed out screen
      
   On Error GoTo PictureBoxZoom_Error
   
   ier = 0

   
   'if Rotation < 0 then zoom out
   'if Rotation > 0 then zoom in
   'if Rotation = 0 then stay at current zoom

   NewZoom = Maximum(0.1, RefZoom.LastZoom + 0.05 * Sgn(Rotation))
   
10:
  If ier = -1 Then 'exiting graefully from memory error
     ier = 0 'reset error flag
     NewZoom = Maximum(0.1, RefZoom.LastZoom - 0.05 * Sgn(Rotation))
     MDIAtmRef.StatusBar.Panels(1).Text = "You computer's memory limits the maximum zoom to: " & RefZoom.LastZoom * 100 & "%"
     End If
     
   picBox.Cls
   AA = CLng(NewZoom * pixwi)
   bb = CLng(NewZoom * pixhi)
   picBox.Width = AA 'CLng(NewZoom * pixwi)
   picBox.height = bb 'CLng(NewZoom * pixhi)
   
   If picBox.Width <> AA Or picBox.height <> bb Then
      'some sort of memory bug for large pictures
      
      Call MsgBox("You have reached the zoom in limit for this map." _
                  & vbCrLf & "To magnify further, use the magnify tool." _
                  , vbInformation, "Zoom in limit")
      
      NewZoom = Maximum(0.1, RefZoom.LastZoom - 0.05 * Sgn(Rotation))
      AA = CLng(NewZoom * pixwi)
      bb = CLng(NewZoom * pixhi)
      picBox.Width = AA 'CLng(NewZoom * pixwi)
      picBox.height = bb 'CLng(NewZoom * pixhi)
      End If
   
   If AA - prjAtmRefMainfm.picture1.Width > 32767 Then
      Call MsgBox("Reached maximum horizontal zoom!", vbInformation, "Horizontal Zoom error")
      Exit Sub
      End If
   
   If bb - prjAtmRefMainfm.picture1.height > 32767 Then
      Call MsgBox("Reached maximum vertical zoom!", vbInformation, "Vertical Zoom error")
      Exit Sub
      End If
      
    'reSet the Max property for the scroll bars.
    prjAtmRefMainfm.HScroll1.max = Maximum(0, AA - prjAtmRefMainfm.picture1.Width)
    prjAtmRefMainfm.VScroll1.max = Maximum(0, bb - prjAtmRefMainfm.picture1.height)
        
    'Determine if the child picture will fill up the screen
    'If so, there is no need to use scroll bars.
'    prjAtmRefMainfm.VScroll1.Visible = (prjAtmRefMainfm.Picture1.Height < AA)
'    prjAtmRefMainfm.HScroll1.Visible = (prjAtmRefMainfm.Picture1.Width < BB)
    
    If prjAtmRefMainfm.HScroll1.max > 0 Then
       prjAtmRefMainfm.HScroll1.Visible = (prjAtmRefMainfm.picture1.Width < bb)
        'Initiate Scroll Step Sizes
        If prjAtmRefMainfm.HScroll1.Visible Then
            prjAtmRefMainfm.HScroll1.LargeChange = Maximum(1, Fix(prjAtmRefMainfm.HScroll1.max / 20))
            prjAtmRefMainfm.HScroll1.SmallChange = Maximum(1, Fix(prjAtmRefMainfm.HScroll1.max / 60))
            End If
        End If
        
    If prjAtmRefMainfm.VScroll1.max > 0 Then
       prjAtmRefMainfm.VScroll1.Visible = (prjAtmRefMainfm.picture1.height < AA)
       If prjAtmRefMainfm.VScroll1.Visible Then
          prjAtmRefMainfm.VScroll1.LargeChange = Maximum(1, Fix(prjAtmRefMainfm.VScroll1.max / 20))
          prjAtmRefMainfm.VScroll1.SmallChange = Maximum(1, Fix(prjAtmRefMainfm.VScroll1.max / 60))
          End If
       End If
   
   'move the blit in order to keep the same pixels in the center
   'this is center
'   XCenter = prjAtmRefMainfm.Picture1.Width * 0.5 'CLng(nearmouse_digi.X * picBox.Width / pixwi)
'   YCenter = prjAtmRefMainfm.Picture1.Height * 0.5 'CLng(nearmouse_digi.Y * picBox.Height / pixhi)
   
   'move them to the center of picBox
   If RefZoom.Left = INIT_VALUE And RefZoom.Top = INIT_VALUE Then
      RefZoom.Left = blink_mark.x
      RefZoom.Top = blink_mark.y
      End If
   NewLeft = RefZoom.Left * NewZoom 'CLng(XCenter - nearmouse_digi.X) ' * NewZoom)
   NewTop = RefZoom.Top * NewZoom 'CLng(YCenter - nearmouse_digi.Y) ' * NewZoom)
    
'   Bltier = StretchBlt(prjAtmRefMainfm.Picture2.hdc, NewLeft, NewTop, CLng(NewZoom * pixwi), CLng(NewZoom * pixhi), prjAtmRefMainfm.PictureBlit.hdc, 0, 0, pixwi, pixhi, vbSrcCopy)
'   Bltier = StretchBlt(prjAtmRefMainfm.Picture2.hdc, NewLeft, NewTop, CLng(NewZoom * pixwi), CLng(NewZoom * pixhi), prjAtmRefMainfm.PictureBlit.hdc, 0, 0, pixwi, pixhi, vbSrcCopy)
   Bltier = StretchBlt(prjAtmRefMainfm.Picture2.hdc, 0, 0, AA, bb, prjAtmRefMainfm.picRef.hdc, 0, 0, pixwi, pixhi, vbSrcCopy)
   RefZoomed = True
   Call ShiftMap(NewLeft, NewTop)
'   prjAtmRefMainfm.Refresh
   picBox.Refresh
   
   If Bltier = 0 Then
       err.Raise vbObjectError + 50, "prjAtmRefMainfm.picture2", "Zoom failed..."
   Else
       'record changes
       RefZoom.Zoom = NewZoom
       RefZoom.LastZoom = RefZoom.Zoom
       End If

 '--------------------------------------------------------------------------

   On Error GoTo 0
   Exit Sub
   
'GeotoCoord:
'
'    CurrentX = ((GeoX - ULGeoX) * GeoToPixelX) + ULPixX
'    CurrentY = ((ULGeoY - GeoY) * GeoToPixelY) + ULPixY
'
'    If RSMethod1 Or RSMethod2 Then
'
'       If RSMethod1 Then
'          ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
'       ElseIf RSMethod2 Then
'          ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
'          End If
'
'        Dim DifX As Double, DifY As Double
'        DifX = Abs(GeoX - XGeo)
'        DifY = Abs(GeoY - YGeo)
'
'        ShiftX = CurrentX - (((XGeo - ULGeoX) * GeoToPixelX) + ULPixX)
'        ShiftY = CurrentY - (((ULGeoY - YGeo) * GeoToPixelY) + ULPixY)
'
'        CurrentX = CurrentX + ShiftX
'        CurrentY = CurrentY + ShiftY
'
'        If RSMethod1 Then
'           ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
'         ElseIf RSMethod2 Then
'           ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
'           End If
'
'        If Abs(GeoX - XGeo) > DifX Then
'           CurrentX = CurrentX - ShiftX
'           End If
'
'        If Abs(GeoY - YGeo) > DifY Then
'           CurrentY = CurrentY - ShiftY
'           End If
'
''        If Abs(GeoX - XGeo) > DifX And Abs(GeoY - YGeo) > DifY Then
'''        If Abs(GeoX - XGeo) > Tolerance Or Abs(GeoY - YGeo) > Tolerance Then
''                Call MsgBox("Inverse coordinate transformation unsuccessful" _
''                        & vbCrLf & "Coordinate grid rotation too large for first approx." _
''                        & vbCrLf & vbCrLf & "(Redo using a less-rotated grid as reference...)" _
''                        , vbInformation, "Picture Box Zoom Error")
''              Screen.MousePointer = vbDefault
''              GDMDIform.picProgBar.Visible = False
''              GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
''              GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
''              Exit Sub
''              End If
'
'   Else
'        'cuurentx, currenty are the pixel coordinates
'        End If
'Return

PictureBoxZoom_Error:

   If err.Number = 480 Then
      'out of memory, can't autoredraw, recover gracefully
      ier = -1
      GoTo 10
      End If

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure PictureBoxZoom of Module modHook"
End Sub

Public Sub ShiftMap(x As Single, y As Single)
'This routine shifts the map in order to put the requested
'coordinate as close to the center of the picture frame as
'possible

     On Error GoTo errhand
        
     'pixel coordinates of the cursor is
     ITMx0 = x / twipsx
     ITMy0 = y / twipsy
     
     'we want it to be at middle of Picture1, i.e., at
     ITMx1 = prjAtmRefMainfm.picture1.Width / 2
     ITMy1 = prjAtmRefMainfm.picture1.height / 2
     
     'Shift the scroll bars in order to accomplish the above
     H1 = ITMx0 - ITMx1 '<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>
     If H1 < prjAtmRefMainfm.HScroll1.Min Or H1 > prjAtmRefMainfm.HScroll1.max Then
        If (drag1x = drag2x And drag1y = drag2y) Then
'           'response = MsgBox("Sorry, your choice would move the map beyond it's boundaries!", vbCritical + vbOKOnly, "GDB")
'
           'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
           If prjAtmRefMainfm.Picture2.Width > prjAtmRefMainfm.HScroll1.Width Then

              If H1 < prjAtmRefMainfm.HScroll1.Min Then
                 prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.Min
              ElseIf H1 > prjAtmRefMainfm.HScroll1.max Then
                 prjAtmRefMainfm.HScroll1.Value = prjAtmRefMainfm.HScroll1.max
                 End If

              End If
'           Exit Sub
        Else 'check if this is end of drag operation that defines box dimensions
           Exit Sub
           End If
     ElseIf (drag1x = drag2x And drag1y = drag2y) Then
        prjAtmRefMainfm.HScroll1.Value = H1
        End If
     
     H2 = ITMy0 - ITMy1
     If H2 < 0 Or H2 > prjAtmRefMainfm.VScroll1.max Then
        If (drag1x = drag2x And drag1y = drag2y) Then
           'response = MsgBox("Sorry, your choice would move the map beyond it's boundaries!", vbCritical + vbOKOnly, "GDB")
           
'           'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
           If prjAtmRefMainfm.Picture2.height > prjAtmRefMainfm.VScroll1.height Then

              If H2 < prjAtmRefMainfm.VScroll1.Min Then
                 prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.Min
              ElseIf H2 > prjAtmRefMainfm.VScroll1.max Then
                 prjAtmRefMainfm.VScroll1.Value = prjAtmRefMainfm.VScroll1.max
                 End If

              End If
'
'           Exit Sub
        Else 'check if this is end of drag operation that defines box dimensions
           Exit Sub
           End If
     ElseIf (drag1x = drag2x And drag1y = drag2y) Then
        prjAtmRefMainfm.VScroll1.Value = H2
        End If
        
   
     newblit = True
     
'    If DigitizePadVis And (DigitizeLine Or DigitizeContour Or DigitizePoint) Then
'       BringWindowToTop (GDDigitizerfrm.hWnd)
'       End If
     
   Exit Sub
   
errhand:
   
   Screen.MousePointer = vbDefault
   

   Select Case err.Number
      Case 480
         If IgnoreAutoRedrawError% = 0 Then
            MsgBox "The pixel size of this map is too big for your memory!" & vbLf & vbLf & _
                   "If you wish to use this map and ignore such errors," & vbLf & _
                   "then check the ""Ignore AutoRedraw errors"" in the" & vbLf & _
                   """Settings"" tab of ""Path/Options"" form.", vbExclamation + vbOKOnly, "AtmRef"
            Exit Sub
         Else 'ignore this error
            Resume Next
            End If
      Case Else
        MsgBox "Encountered error #: " & err.Number & vbLf & _
               err.Description & vbLf & _
               "in module: ShiftMap", vbCritical + vbOKOnly, "AtmRef"
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
      MsgBox msgerr$, vbExclamation + vbOKOnly, "AtmRef"
      End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : UpdateStatus
' Author    : Dr-John-K-Hall
' Date      : 2/18/2015
' Purpose   : Updates Status of fancy progress bar
'---------------------------------------------------------------------------------------
'
Public Sub UpdateStatus(Form1 As Form, picProgBar As PictureBox, ShowStatusProgress As Boolean, FileBytes As Long)
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

    BringWindowToTop (Form1.hwnd)
    picProgBar.Visible = True

    If FileBytes > pbScaleWidth Then
       progress = pbScaleWidth
       picProgBar.Visible = False
    Else
       progress = FileBytes
       End If


    Txt$ = Format$(CLng((progress / pbScaleWidth) * 100)) + "%..."

    If ShowStatusProgress Then
       MDIAtmRef.StatusBar.Panels(3).Text = Format$(CLng((progress / pbScaleWidth) * 100)) + "%..."
       End If

    picProgBar.Cls
    picProgBar.ScaleWidth = pbScaleWidth
    picProgBar.CurrentX = (pbScaleWidth - picProgBar.TextWidth(Txt$)) \ 2
    picProgBar.CurrentY = (picProgBar.ScaleHeight - picProgBar.TextHeight(Txt$)) \ 2
    picProgBar.Print Txt$
    picProgBar.Line (0, 0)-(progress, picProgBar.ScaleHeight), picProgBar.ForeColor, BF
    r = BitBlt(picProgBar.hdc, 0, 0, pbScaleWidth, picProgBar.ScaleHeight, picProgBar.hdc, 0, 0, SRCCOPY)
    picProgBar.Refresh
    DoEvents

   On Error GoTo 0
   Exit Sub

UpdateStatus_Error:

    Resume Next

End Sub
Public Function Minimum(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Minimum = v1
Else: Minimum = v2
End If
End Function

Public Function Maximum(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Maximum = v2
Else: Maximum = v1
End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : LoadAtmospheres
' Author    : Dr-John-K-Hall
' Date      : 1/4/2019
' Purpose   : Loads up temperature, pressure profiles of model atmospheres
' parameters:
'           FileName = filename of external atomosphere file to read
'           AtmType = Type of atmosphere
'           AtmNumber = Number designation of the AtmType atmosphere type
'---------------------------------------------------------------------------------------
'
Public Function LoadAtmospheres(filename As String, AtmType As Integer, AtmNumber As Integer, _
                                LapseRate As Double, TempTop As Double, PressTop As Double, _
                                NumRecords As Long, Mode As Integer) As Long

   Dim ElevSource(1) As Double, PresSource(1) As Double, LapseSource(1) As Double
   Dim NumSource As Long, StatusMes As String
   Dim HH0 As Double, HH As Double, Lapse0 As Double, Lapse As Double, Hgts As Double
   Dim temp As Double, Temp0 As Double, press As Double, Press0 As Double
   Dim StartTemp As Double, StepHgt As Double, AConst As Double, CConst As Double
   Dim MultStep As Double, Isothermic As Boolean, StepBorder As Double, StepNudge As Double
   Dim LapseRateCalc As Double, DocSplit() As String, OffsetHgt As Double
   
   On Error GoTo LoadAtmospheres_Error

     Screen.MousePointer = vbHourglass

     Select Case AtmType
     
        Case 1
          
          filnum% = FreeFile
          Open filename For Input As #filnum%
          FNM = filename
          StatusMes = "Reading temperature profile from " & FNM
          Call StatusMessage(StatusMes, 1, 0)
    
          Input #filnum%, NNN
          NNN = NNN - 1
          lpaseratecalc = 0#
          Call UpdateStatus(prjAtmRefMainfm, prjAtmRefMainfm.picProgBar, 1, 0)
          For II = 0 To NNN
             Input #filnum%, ELV(II), TMP(II)
             PRSR(II) = 1013.25 * Exp(-ELV(II) / 8400#)
             
             If Mode = 3 And II = 0 Then
                TempTop = TMP(0)
                PressTop = PRSR(0)
                LapseRateCalc = TMP(0)
             ElseIf Mode = 3 And II > 0 And ELV(II) > 11000 Then
                LapseRateCalc = Abs((TMP(II - 1) - LapseRateCalc) / (ELV(II - 1)))
                LapseRate = LapseRateCalc
                Exit Function
                End If
             
             If Mode = 2 Then 'also calclate the index of refraction
'                Call INDEX_CIDDOR(wl, t, P, h, xc, DA, DW, Index)
                End If
             
            If II = 0 Then
               MinTemp = TMP(0)
               MaxTemp = MinTemp
            Else
               If TMP(II) > MaxTemp Then MaxTemp = TMP(II)
               If TMP(II) < MinTemp Then MinTemp = TMP(II)
               End If
                   
             Call UpdateStatus(prjAtmRefMainfm, prjAtmRefMainfm.picProgBar, 1, CLng(100# * II / (NNN - 1)))
    
          Next II
          
          Close #filnum%
          NumRecords = NNN
          prjAtmRefMainfm.progressfrm.Visible = False
          Call UpdateStatus(prjAtmRefMainfm, prjAtmRefMainfm.picProgBar, 1, 0)
        
        Case 2
        
           Select Case AtmNumber
              Case 0
                 FNM = App.Path & "\VDW-atmosphere.txt"
                 GoSub ConvertToElevTempPress3
              Case 1
                 FNM = App.Path & "\Bruton-US-standard.txt"
                 'generate height-temperature data using the gradinets for every 10 meters up to final height
                 'pressure determined from Bruton's model including humidity
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 2
                 FNM = App.Path & "\Selby-tropical.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 3
                 FNM = App.Path & "\Selby-midlatitude-winter.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 4
                 FNM = App.Path & "\Selby-midlatitude-summer.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 5
                 FNM = App.Path & "\Selby-subartic-winter.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                   ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                    End If
              Case 6
                 FNM = App.Path & "\Selby-subartic-summer.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 7
                 FNM = App.Path & "\Selby-US-standard.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 8
                 FNM = App.Path & "\Menat-EY-winter.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
'                 CalcType = 2
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 9
                 FNM = App.Path & "\Menat-EY-summer.txt"
                 'generate height-temperature,pressure data for a HeightStep meter step size
                  If Mode = 1 Then
                     GoSub ConvertToElevTempPress
                  ElseIf Mode > 2 And Mode < 5 Then
                     GoSub ConvertToElevTempPress2
                  ElseIf Mode = 5 Then
                     GoSub ConvertToElevTempPress3
                     End If
              Case 10
                 FNM = filename
                 'generate height-temperature,pressure data for a HeightStep meter step size
                 'query what type of calculation type it is
                 
                 If LoopingAtmTracing Then
                    If Mode = 4 Then
                       GoSub ConvertToElevTempPress2
                    ElseIf Mode = 5 Then
                       GoSub ConvertToElevTempPress3
                       End If
                 Else
                    If CalcSondes Then
                       'skip notifications
                       MDIAtmRef.StatusBar.Panels(2).Text = prjAtmRefMainfm.txtOther.Text & " Lowtran atms. file - convert elevations to km"
                        If Mode = 1 Then
                           GoSub ConvertToElevTempPress
                        ElseIf Mode > 2 And Mode < 5 Then
                           GoSub ConvertToElevTempPress2
                        ElseIf Mode = 5 Then
                           GoSub ConvertToElevTempPress3
                           End If
                        GoTo la100
                        End If
                              
                    Select Case MsgBox("Confirm that the file of this file is the same as the Lowtran atmospheric files," _
                                       & vbCrLf & sEmpty _
                                       & vbCrLf & "i.e., rows having three columns corresponding to:" _
                                       & vbCrLf & sEmpty _
                                       & vbCrLf & "Inreasing Height )km), Temperature (Kelvin), Pressure (mb)" _
                                       & vbCrLf & sEmpty _
                                       , vbOKCancel Or vbQuestion Or vbDefaultButton1, "Other atmospher file upload")
                    
                       Case vbOK
                       
                           If prjAtmRefMainfm.chkMeters.Value = vbChecked Then
                               Select Case MsgBox("Converting the heights from meters to km?", vbOKCancel Or vbQuestion Or vbDefaultButton1, "conversion of heights")
                               
                                   Case vbOK
                                      MDIAtmRef.StatusBar.Panels(2).Text = prjAtmRefMainfm.txtOther.Text
                                   Case vbCancel
                                       LoadAtmospheres = -1
                                       Exit Function
                               
                               End Select
                               End If
                           
                           If Mode = 1 Then
                              GoSub ConvertToElevTempPress
                           ElseIf Mode > 2 And Mode < 5 Then
                              GoSub ConvertToElevTempPress2
                           ElseIf Mode = 5 Then
                              GoSub ConvertToElevTempPress3
                              End If
                              
                       Case vbCancel
                           LoadAtmospheres = -1
                           Exit Function
                    End Select
                    End If

           End Select
        
     End Select
la100:
      Screen.MousePointer = vbDefault
      LoadAtmospheres = 0
      Exit Function
      
ConvertToElevTempPress:
   'open file and read in values
   If Dir(FNM) <> "" Then
   
     StartTemp = prjAtmRefMainfm.txtGroundTemp
     StepHgt = prjAtmRefMainfm.txtHeightStepSize
   
    StepBorder = StepHgt
    StepNudge = 0#
     
     NumSource = 0
     
     filnum% = FreeFile
     Open FNM For Input As #filnum%
     
     Select Case AtmNumber
           
       Case 1
         'list of heights and temperature gradients = lapse
         'use Bruton's exponential model for pressures as a function of height

         Do Until EOF(filnum%)
            If NumSource = 0 Then
               Input #filnum%, HH0, Lapse0
                  
               Input #filnum%, HH, Lapse
               HH0 = HH0 * 1000 'convert to meters
               HH = HH * 1000

               Lapse0 = Lapse0 * 0.001 'convert lapse to degrees K/m
               Lapse = Lapse * 0.001
               For Hgts = HH0 + StepNudge To HH - StepBorder Step StepHgt
                  ELV(NumSource) = Hgts
                  TMP(NumSource) = StartTemp + Lapse0 * (Hgts - HH0)
                  PRSR(NumSource) = 1013.25 * Exp(-ELV(NumSource) / 8400#)
                  NumSource = NumSource + 1
                  
                If NumSource = 1 Then
                   MinTemp = TMP(0)
                   MaxTemp = MinTemp
                Else
                   If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
                   If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                   End If
                
               Next Hgts
            Else
               HH0 = HH
               StartTemp = TMP(NumSource - 1)
               Lapse0 = Lapse
               Input #filnum%, HH, Lapse
               HH = HH * 1000 'convert to meters
               Lapse = Lapse * 0.001 'convert to deg K/meter
               
                StepNudge = 0#
                MultStep = 1#
                If HH >= 2000 And HH < 10000 Then
                    MultStep = 5
                 ElseIf HH >= 10000 And HH < 30000 Then
                    MultStep = 10
                 ElseIf HH >= 30000 Then
                    MultStep = 20
                    End If
               
               For Hgts = HH0 + StepNudge To HH - StepBorder Step StepHgt * MultStep
                  ELV(NumSource) = Hgts
                  TMP(NumSource) = StartTemp + Lapse0 * (Hgts - HH0)
                  PRSR(NumSource) = 1013.25 * Exp(-ELV(NumSource) / 8400#)
                     
                  NumSource = NumSource + 1
                  
                If NumSource = 1 Then
                   MinTemp = TMP(0)
                   MaxTemp = MinTemp
                Else
                   If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
                   If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                   End If
                  
               Next Hgts
               End If
               
             If NumSource = 1 Then
                MinTemp = TMP(0)
                MaxTemp = MinTemp
             Else
                If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
                If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                End If
                
           Loop
           NumRecords = NumSource - 1
       
       Case Else
          'list of heights, temperatures, and pressures
         
         If InStr(FNM, "HS-atmosphere.txt") Then
            NumSource = 0
            Do Until EOF(filnum%)
               Input #filnum%, ELV(NumSource), TMP(NumSource), PRSR(NumSource), IndexRefraction(NumSource)
                
               NumSource = NumSource + 1
               
                If NumSource = 1 Then
                   MinTemp = TMP(0)
                   MaxTemp = MinTemp
                Else
                   If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
                   If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                   End If
                   
            Loop
            Close #filnum%
            NumRecords = NumSource
            UsingHSatmosphere = True
            Exit Function
            End If
          
          LapseRateFound = False
          Do Until EOF(filnum%)
             If NumSource = 0 Then
                Input #filnum%, HH0, Temp0, Press0

                Input #filnum%, HH, temp, press
                HH0 = HH0 * 1000 'convert to meters
                HH = HH * 1000
                AConst = (temp - Temp0) / (HH - HH0)
                
                StepNudge = 0#
                   
                'check for isothermal regions
                If temp <> Temp0 Then
                   CConst = Log(press / Press0) / Log(temp / Temp0)
                   Isothermic = False
                Else
                   Isothermic = True
                   'should be exponential decay, but have to fit to linear effect
                   'should be -- > CConst = 9.81 / (8.31451 * Temp)
                   'and prsr = press0 * exp(-CConst * (Hgts - HHH))
                   'instead model it is linear
                   CConst = (press - Press0) / (HH - HH0)
                   End If
                   
                For Hgts = HH0 + StepNudge To HH - StepBorder Step StepHgt
                   ELV(NumSource) = Hgts
                   TMP(NumSource) = Temp0 + (Hgts - HH0) * AConst
                   
                   If Not Isothermic Then
                      PRSR(NumSource) = Press0 * (TMP(NumSource) / Temp0) ^ CConst
                   Else
'                      PRSR(NumSource) = Press0 * Exp(-CConst * (Hgts - HH0))
                       PRSR(NumSource) = Press0 + (Hgts - HH0) * CConst
                      End If
                      
                   NumSource = NumSource + 1
                   
                    If NumSource = 1 Then
                       MinTemp = TMP(0)
                       MaxTemp = MinTemp
                    Else
                       If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
                       If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                       End If

                Next Hgts
             Else
                HH0 = HH
                Temp0 = temp
                Press0 = press
                Input #filnum%, HH, temp, press
                HH = HH * 1000
                
                StepNudge = 0#
                MultStep = 1#
                If HH >= 2000 And HH < 10000 Then
                   MultStep = 5
                ElseIf HH >= 10000 And HH < 30000 Then
                   MultStep = 10
                ElseIf HH >= 30000 Then
                   MultStep = 20
                   End If
                   
                AConst = (temp - Temp0) / (HH - HH0)
                
                'check for isothermal regions
                If temp <> Temp0 Then
                   CConst = Log(press / Press0) / Log(temp / Temp0)
                   Isothermic = False
                Else
                   Isothermic = True
                   'should be exponential decay, but have to fit to linear effect
                   'should be -- > CConst = 9.81 / (8.31451 * Temp)
                   'and prsr = press0 * exp(-CConst * (Hgts - HHH))
                   'instead model it is linear
                   CConst = (press - Press0) / (HH - HH0)
                   End If
                   
                For Hgts = HH0 + StepNudge To HH - StepBorder Step StepHgt * MultStep
                   ELV(NumSource) = Hgts
                   TMP(NumSource) = Temp0 + (Hgts - HH0) * AConst
                   
                   If Not Isothermic Then
                      PRSR(NumSource) = Press0 * (TMP(NumSource) / Temp0) ^ CConst
                   Else
'                      PRSR(NumSource) = Press0 * Exp(-CConst * (Hgts - HH0))
                       PRSR(NumSource) = Press0 + (Hgts - HH0) * CConst
                      End If
                   
                   NumSource = NumSource + 1
                   
                    If NumSource = 1 Then
                       MinTemp = TMP(0)
                       MaxTemp = MinTemp
                    Else
                       If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
                       If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                       End If
                   
                Next Hgts
                End If
                
            Loop
            NumRecords = NumSource - 1
            
            'now load up the temperature and pressure diagrams
            
     End Select
     Close #filnum%
     
   Else
      Screen.MousePointer = vbDefault
      ier = MsgBox("Can't find file: " & FNM _
                         & vbCrLf & sEmpty _
                         , vbOKOnly Or vbExclamation Or vbDefaultButton1, "Other atmospher file upload")
      
      LoadAtmospheres = -1
      
      Exit Function
      End If

Return

ConvertToElevTempPress2:
   'open file and read in values
   If Dir(FNM) <> sEmpty Then
   
     StartTemp = prjAtmRefMainfm.txtGroundTemp
     
     NumSource = 0
     
     filnum% = FreeFile
     Open FNM For Input As #filnum%
     
     Select Case AtmNumber
           
       Case 1
         'list of heights and temperature gradients = lapse
         'use Bruton's exponential model for pressures as a function of height

         Do Until EOF(filnum%)
            If NumSource = 0 Then
               Input #filnum%, HH0, Lapse0
               HH0 = HH0 * 1000 'convert to meters
               Lapse0 = Lapse0 * 0.001 'convert to deg K/meters
               
               ELV(NumSource) = HH0 * 1000
               TMP(NumSource) = StartTemp
               PRSR(NumSource) = prjAtmRefMainfm.txtGroundPressure
               
               If Mode = 3 Then
                  LapseRate = Abs(Lapse0)
                  TempTop = prjAtmRefMainfm.txtGroundTemp
                  PressTop = prjAtmRefMainfm.txtGroundPressure
                  LapseRateCalc = TempTop
                  Exit Function
                  End If
                  
               MinTemp = TMP(0)
               MaxTemp = MinTemp
               
               NumSource = NumSource + 1
                   
            Else
               Input #filnum%, HH, Lapse
               HH = HH * 1000 'convert to meters
               Lapse = Lapse * 0.001 'convert to deg K/m
               
               ELV(NumSource) = HH
               TMP(NumSource) = StartTemp + Lapse0 * (HH - HH0)
               PRSR(NumSource) = prjAtmRefMainfm.txtGroundPressure * Exp(-ELV(NumSource) / 8400#)
                  
               Lapse0 = Lapse
               HH0 = HH
                
               If Mode = 3 And ELV(NumSource) > 11000 Then
                   LapseRateCalc = Abs((TMP(NumSource - 1) - LapseRateCalc) / ELV(NumSource - 1))
                   LapseRate = LapseRateCalc
                   Exit Function
                   End If
                   
               NumSource = NumSource + 1
    
               If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
               If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                     
               End If
               
           Loop
           NumRecords = NumSource - 1
       
       Case Else
          'list of heights, temperatures, and pressures
         
         If InStr(FNM, "HS-atmosphere.txt") Then
            NumSource = 0
            Do Until EOF(filnum%)
               Input #filnum%, ELV(NumSource), TMP(NumSource), PRSR(NumSource), IndexRefraction(NumSource)
               
               If Mode = 3 And NumSource = 0 Then
                   TempTop = TMP(NumSource)
                   PressTop = PRSR(NumSource)
                   LapseRateCalc = TMP(NumSource)
               ElseIf Mode = 3 And NumSource > 0 Then
                   If ELV(NumSource) > 11000 Then
                       LapseRateCalc = Abs((TMP(NumSource - 1) - LapseRateCalc) / ELV(NumSource - 1))
                       LapseRate = LapseRateCalc
                       Exit Function
                       End If
                   End If
               
               NumSource = NumSource + 1
               
               If NumSource = 1 Then
                   MinTemp = TMP(0)
                   MaxTemp = MinTemp
               Else
                   If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
                   If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                   End If
                   
            Loop
            
            Close #filnum%
            NumRecords = NumSource
            UsingHSatmosphere = True
            Exit Function
            End If
          
          LapseRateFound = False
          NumSource = 0
          Do Until EOF(filnum%)
          
             Input #filnum%, ELV(NumSource), TMP(NumSource), PRSR(NumSource)
             
             If prjAtmRefMainfm.chkMeters.Value = vbChecked Then
                ELV(NumSource) = ELV(NumSource) * 0.001 'convert to kilometers
                End If
               
              If TMP(NumSource) < 150 Then 'must be centigrade, so convert to Kelvin
                 TMP(NumSource) = TMP(NumSource) + 273.15
                 End If
                    
             If Mode = 3 Then ELV(NumSource) = ELV(NumSource) * 1000#  'convert to meters
             
             If NumSource = 0 Then
             
                If Mode = 3 Then
                   TempTop = TMP(NumSource)
                   PressTop = PRSR(NumSource)
                   LapseRateCalc = TempTop
                   End If
                   
                If prjAtmRefMainfm.chkReNorm.Value = vbChecked Then
                   OffsetHgt = ELV(0)
                   ELV(0) = 0#
                   End If
                   
                prjAtmRefMainfm.txtTGROUND.Text = TMP(0)
                prjAtmRefMainfm.txtPress0.Text = PRSR(0)
                If ZeroRefTesting Then prjAtmRefMainfm.txtHOBS.Text = ELV(0) * 1000#
                   
                MinTemp = TMP(0)
                MaxTemp = MinTemp
                   
             Else
             
                If prjAtmRefMainfm.chkReNorm.Value = vbChecked Then
                   ELV(NumSource) = ELV(NumSource) - OffsetHgt
                   End If
             
                If TMP(NumSource) > MaxTemp Then MaxTemp = TMP(NumSource)
                If TMP(NumSource) < MinTemp Then MinTemp = TMP(NumSource)
             
                If Mode = 3 Then
                   If ELV(NumSource) > 11000 Then
                       LapseRateCalc = Abs((TMP(NumSource - 1) - LapseRateCalc) / ELV(NumSource - 1))
                       LapseRate = LapseRateCalc
                       Exit Function
                       End If
                   End If
                
                End If
                
             NumSource = NumSource + 1
                          
          Loop
          
          NumRecords = NumSource - 1
            
          'now load up the temperature and pressure diagrams
            
     End Select
     Close #filnum%
     
   Else
      Screen.MousePointer = vbDefault
      ier = MsgBox("Can't find file: " & FNM _
                         & vbCrLf & sEmpty _
                         , vbOKOnly Or vbExclamation Or vbDefaultButton1, "Other atmospher file upload")
      
      LoadAtmospheres = -1
      
      Exit Function
      End If

Return

ConvertToElevTempPress3:
   'open file and read in values
   If Dir(FNM) <> sEmpty Then
   
     prjAtmRefMainfm.txtTGROUND = prjAtmRefMainfm.txtGroundTemp
     prjAtmRefMainfm.txtPress0 = prjAtmRefMainfm.txtGroundPressure
     
     StartTemp = prjAtmRefMainfm.txtGroundTemp
     
     NumSource = 0
     
     filnum% = FreeFile
     Open FNM For Input As #filnum%
     
     Select Case AtmNumber
     
       Case 0 'van der Werf's US standard atmosphere
          
          Do Until EOF(filnum%)
            Input #filnum%, HL(NumSource), TL(NumSource), LRL(NumSource)
            NumSource = NumSource + 1
          Loop
          Close #filnum%
          NumRecords = NumSource - 1
           
       Case 1
         'list of heights and temperature gradients = lapse
         'use Bruton's exponential model for pressures as a function of height

         Do Until EOF(filnum%)
            If NumSource = 0 Then
               Input #filnum%, HH0, Lapse0
               HH0 = HH0 'convert to meters
               Lapse0 = Lapse0  'in units of deg K/km
               
               HL(NumSource) = HH0
               TL(NumSource) = StartTemp
               LRL(NumSource) = Lapse0
                  
               MinTemp = TMP(0)
               MaxTemp = MinTemp
               
               NumSource = NumSource + 1
                   
            Else
               Input #filnum%, HH, Lapse
               
               HL(NumSource) = HH
               TL(NumSource) = StartTemp + Lapse0 * (HH - HH0)
               LRL(NumSource) = Lapse
                  
               Lapse0 = Lapse
               HH0 = HH
                
               NumSource = NumSource + 1
    
               If TMP(NumSource - 1) > MaxTemp Then MaxTemp = TMP(NumSource - 1)
               If TMP(NumSource - 1) < MinTemp Then MinTemp = TMP(NumSource - 1)
                     
               End If
               
           Loop
           NumRecords = NumSource - 1
           prjAtmRefMainfm.txtHMAXT.Text = CInt(HH0 * 10#)
           UsingHSatmosphere = True
           Exit Function
       
       Case Else
          'list of heights, temperatures, and pressures
                      
         Dim PPP As Double, IRR As Double
         
         If InStr(FNM, "HS-atmosphere.txt") Then
            NumSource = 0
               
            Do Until EOF(filnum%)
               Input #filnum%, HL(NumSource), TL(NumSource), PPP, IRR
                   
              If NumSource = 0 Then
                   HH0 = HL(0)
                   txtTGROUND.Text = TL(0)
                   txtPress0 = PPP
               ElseIf NumSource > 0 Then
                   LapseRateCalc = ((TL(NumSource - 1) - TL(NumSource)) / (HL(NumSource) - HL(NumSource - 1)))
                   LapseRate = LapseRateCalc
                   LRL(NumSource - 1) = LapseRate
                   End If
               
               NumSource = NumSource + 1
               
               If NumSource = 1 Then
                   MinTemp = TL(NumSource - 1)
                   MaxTemp = MinTemp
               Else
                   If TL(NumSource - 1) > MaxTemp Then MaxTemp = TL(NumSource - 1)
                   If TL(NumSource - 1) < MinTemp Then MinTemp = TL(NumSource - 1)
                   End If
                   
            Loop
            
            Close #filnum%
            prjAtmRefMainfm.txtHMAXT = HL(NumSource - 1) * 10#
            LRL(NumSource - 1) = 0#
            NumRecords = NumSource
            UsingHSatmosphere = True
            Exit Function
            End If
          
          LapseRateFound = False
          NumSource = 0
          OffsetHgt = 0#
          Do Until EOF(filnum%)
             If InStr(FNM, "-sondes.txt") Then
                'Beit Dagan sondes determine the atmosphere
                If NumSource = 0 Then
                   'some sort of extra undetected characters on first line, so line input first line and split it
                   Line Input #filnum%, doclin$
                   DocSplit = Split(doclin$, ",")
                   'now filter out the nonnumerical characters
                   filtch$ = ""
                   For i = Len(DocSplit(0)) To 1 Step -1
                      If InStr("0123456789", Mid$(DocSplit(0), i, 1)) Then
                         filtch$ = Mid$(DocSplit(0), i, 1) & filtch$
                         End If
                   Next i
                   HL(0) = Val(filtch$)
                   TL(0) = Val(DocSplit(1))
                   'check for erroneous sonde caused by missing ground temperature
                   If TL(0) = 0 Then
                      'this is likely due to missing ground temperature
                      'so exit with error flag
                      LoadAtmospheres = -1
                      Exit Function
                      End If
                   PPP = Val(DocSplit(2))
                   If prjAtmRefMainfm.chkReNorm.Value = vbChecked Then
                      OffsetHgt = HL(0)
                      HL(0) = 0#
                      End If
                   If LoopingAtmTracing And prjAtmRefMainfm.chkHgtProfile.Value = vbChecked Then
                      prjAtmRefMainfm.txtHOBS.Text = DistModel(0) 'use approx Rabbi Druk's observation altitude for these calculations
                      End If
                 Else
    '                Input #filnum%, HL(NumSource), TL(NumSource), PPP
                    Line Input #filnum%, doclin$
                    DocSplit = Split(doclin$, ",")
                    If UBound(DocSplit) = 2 Then
                       HL(NumSource) = Val(DocSplit(0))
                       TL(NumSource) = Val(DocSplit(1))
                       PPP = Val(DocSplit(2))
                       If prjAtmRefMainfm.chkReNorm.Value = vbChecked Then
                           HL(NumSource) = HL(NumSource) - OffsetHgt
                           End If
                    Else
                      'this is eof with some extra characters
                      Exit Do
                      End If
                    End If
                Else
                   Input #filnum%, HL(NumSource), TL(NumSource), PPP
                   If NumSource = 0 Then
                      If prjAtmRefMainfm.chkReNorm.Value = vbChecked Then
                         OffsetHgt = HL(0)
                         HL(0) = 0#
                         End If
                   Else
                       If prjAtmRefMainfm.chkReNorm.Value = vbChecked Then
                           HL(NumSource) = HL(NumSource) - OffsetHgt
                           End If
                       End If
                   End If
                   
                   
            If prjAtmRefMainfm.chkMeters.Value = vbChecked Then
               HL(NumSource) = HL(NumSource) * 0.001 'convert to kilometers
               End If
               
             If TL(NumSource) < 150 Then 'must be centigrade, so convert to Kelvin
                TL(NumSource) = TL(NumSource) + 273.15
                End If
             
             If NumSource = 0 Then
             
                prjAtmRefMainfm.txtTGROUND = TL(0)
                prjAtmRefMainfm.txtPress0 = PPP
                If prjAtmRefMainfm.chkReNorm.Value = vbChecked Then
                  If Val(prjAtmRefMainfm.txtHOBS.Text) = 0 Then
                     HL(0) = 0#
                     End If
                  End If
                  
                HH0 = HL(0)
                   
                MinTemp = TL(0)
                MaxTemp = MinTemp
                   
             Else
             
                If TL(NumSource) > MaxTemp Then MaxTemp = TL(NumSource)
                If TL(NumSource) < MinTemp Then MinTemp = TL(NumSource)
             
                LapseRateCalc = ((TL(NumSource - 1) - TL(NumSource)) / (HL(NumSource) - HL(NumSource - 1)))
                LapseRate = LapseRateCalc
                LRL(NumSource - 1) = LapseRate
               
                End If
                
             NumSource = NumSource + 1
                          
          Loop
          
          NumRecords = NumSource
          prjAtmRefMainfm.txtHMAXT = HL(NumSource - 1) * 10#
          LRL(NumSource - 1) = 0#
          UsingHSatmosphere = True
            
          'now load up the temperature and pressure diagrams
            
     End Select
     Close #filnum%
     
   Else
      Screen.MousePointer = vbDefault
      ier = MsgBox("Can't find file: " & FNM _
                         & vbCrLf & sEmpty _
                         , vbOKOnly Or vbExclamation Or vbDefaultButton1, "Other atmospher file upload")
      
      LoadAtmospheres = -1
      
      Exit Function
      End If

Return


   On Error GoTo 0
   LoadAtmospheres = 0
   Exit Function

LoadAtmospheres_Error:
    Resume
    Close
    
   prjAtmRefMainfm.cmdCalc.Enabled = True
   prjAtmRefMainfm.cmdRefWilson.Enabled = True
   prjAtmRefMainfm.cmdMenat.Enabled = True
   prjAtmRefMainfm.cmdVDW.Enabled = True

   
    LoadAtmospheres = -1
    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure LoadAtmospheres of Module modBurtonAstRef"
'    prjAtmRefMainfm.cmdCalc.Enabled = True
End Function



' -*- coding: utf-8 -*-
' Author: Mikhail Polyanskiy
' Last modified: 2017-11-23
' Original data: Ciddor 1996, https://doi.org/10.1364/AO.35.001566
' P. E. Ciddor. Refractive index of air: new equations for the visible and near infrared, Appl. Optics 35, 1566-1573 (1996)
' source: https://github.com/polyanskiy/refractiveindex.info-scripts/blob/master/scripts/Ciddor%201996%20-%20air.py
' source: https://refractiveindex.info/?shelf=other&book=air&page=Ciddor

'##############################################################################

'import numpy as np
'import matplotlib.pyplot as plt
'P = np.PI


'def Z(T, P, xw): 'compressibility
Public Function ZComp(TK As Double, P As Double, xw As Double) As Double
    Dim a0 As Double, A1 As Double, A2 As Double, b0 As Double, B1 As Double
    Dim c0 As Double, C1 As Double, d As Double, e As Double, TC As Double
    
    TC = TK - 273.15
    a0 = 0.00000158123 'K·Pa^-1
    A1 = -0.000000029331 'Pa^-1
    A2 = 0.00000000011043 'K^-1·Pa^-1
    b0 = 0.000005707  'K·Pa^-1
    B1 = -0.00000002051 'Pa^-1
    c0 = 0.00019898   'K·Pa^-1
    C1 = -0.000002376 'Pa^-1
    d = 0.0000000000183 'K^2·Pa^-2
    e = -0.00000000765 'K^2·Pa^-2
    ZComp = 1 - (P / TK) * (a0 + A1 * TC + A2 * TC ^ 2 + (b0 + B1 * TC) * xw + (c0 + C1 * TC) * xw ^ 2) + (P / TK) ^ 2 * (d + e * xw ^ 2)
    
End Function



'def n(?,t,p,h,xc):
Public Sub INDEX_CIDDOR(wl As Double, TK As Double, P As Double, H As Double, xc As Double, PA As Double, PW As Double, Index As Double)

    Dim s As Double, r As Double, k0 As Double, k1 As Double, k2 As Double, k3 As Double
    Dim w0 As Double, w1 As Double, w2 As Double, w3 As Double, TC As Double
    Dim A As Double, b As Double, C As Double, d As Double, ALPHA As Double, BETA As Double, gamma As Double
    Dim SVP As Double, xw As Double, Ma As Double, mw As Double, Za As Double, Zw As Double
    Dim paxs As Double, pws As Double, nas As Double, naxs As Double, nws As Double, F As Double

    ' WAV: wavelength, 0.3 to 1.69 µm
    ' t: temperature, -40 to +100 °C
    ' p: pressure, 80000 to 120000 Pa
    ' h: fractional humidity, 0 to 1
    ' xc: CO2 concentration, 0 to 2000 ppm
    
    'input requires Pressure in mb, e.g., 1013.25, Temperature in degrees Kelvin, e.g., 298.6
    TC = TK - 273.15 'degrees Celsius

    s = 1000# / wl        ' 'µm^-1, i.e., wl in units of nm, e.g., 574 nm
    
    r = 8.31451       'gas constant, J/(mol·K)
    
    k0 = 238.0185     'µm^-2
    k1 = 5792105      'µm^-2
    k2 = 57.362       'µm^-2
    k3 = 167917       'µm^-2
 
    w0 = 295.235      'µm^-2
    w1 = 2.6422       'µm^-2
    w2 = -0.03238     'µm^-4
    w3 = 0.004028     'µm^-6
    
    A = 0.000012378847 'K^-2
    b = -0.019121316  'K^-1
    C = 33.93711047
    d = -6343.1645    'K
    
    ALPHA = 1.00062
    ßbeta = 0.0000000314  'Pa^-1,
    gamma = 0.00000056    '°C^-2

    'saturation vapor pressure of water vapor in air at temperature T
    If (T >= 0) Then
        SVP = Exp(A * TK ^ 2 + b * TK + C + d / TK) 'Pa
    Else
        SVP = 10 ^ (-2663.5 / TK + 12.537)
        End If
    
    'enhancement factor of water vapor in air
    F = ALPHA + ßbeta * P + gamma * TC ^ 2
    
    'molar fraction of water vapor in moist air
    xw = F * H * SVP / P
    
    Ma = 0.001 * (28.9635 + 0.000012011 * (xc - 400)) 'molar mass of dry air, kg/mol
    mw = 0.018015                            'molar mass of water vapor, kg/mol
    
    Za = ZComp(288.15, 101325, 0)                'compressibility of dry air
    Zw = ZComp(293.15, 1333, 1)                  'compressibility of pure water vapor
    
    'Eq.4 with (TK,P,xw) = (288.15, 101325, 0)
    paxs = 101325 * Ma / (Za * r * 288.15) 'density of standard air
    
    'Eq 4 with (TK,P,xw) = (293.15, 1333, 1)
    pws = 1333 * mw / (Zw * r * 293.15) 'density of standard water vapor
    
    ' two parts of Eq.4: ?=?a+?w
    PA = P * Ma / (ZComp(TK, P, xw) * r * TK) * (1 - xw) 'density of the dry component of the moist air
    PW = P * mw / (ZComp(TK, P, xw) * r * TK) * xw 'density of the water vapor component
    
    'refractive index of standard air at 15 °C, 101325 Pa, 0% humidity, 450 ppm CO2
    nas = 1 + (k1 / (k0 - s ^ 2) + k3 / (k2 - s ^ 2)) * 0.00000001
    
    'refractive index of standard air at 15 °C, 101325 Pa, 0% humidity, xc ppm CO2
    naxs = 1 + (nas - 1) * (1 + 0.000000534 * (xc - 450))
    
    'refractive index of water vapor at standard conditions (20 °C, 1333 Pa)
    nws = 1 + 1.022 * (w0 + w1 * s ^ 2 + w2 * s ^ 4 + w3 * s ^ 6) * 0.00000001
    
    Index = 1 + (PA / paxs) * (naxs - 1) + (PW / pws) * (nws - 1)
    
End Sub
'C
''C*********************************************************************
'C
'C
'    Public Sub RADCUR(II As Long, RC As Double, DEN As Double, ByRef ELV() As Double, ByRef TMP() As Double, WL As Double, AtmModel As Integer)
    Public Sub RADCUR_new(II As Long, RC As Double, den As Double, wl As Double, AtmModel As Integer)
'C
'C   This subroutine calculates the radius of curvature
'C   RC for horizontal rays in the layer II as well as
'C   average density (kg/cm^3).
'C
'       IMPLICIT DOUBLE PRECISION (A-H,O-Z)
       Dim SIGNA As Double, P1 As Double, T1 As Double, RI1 As Double, DEN1 As Double
       Dim P2 As Double, T2 As Double, RI2 As Double, DELV As Double, DEN2 As Double
       Dim RKAPPA As Double, A As Double, b As Double, C As Double, d As Double
       Dim ALPHA As Double, BETA As Double, gamma As Double, DA As Double, DW As Double
       Dim T As Double, P As Double, SVP As Double, F As Double, H As Double, xc As Double
       Dim Index As Double
                
       If II > 0 Then

          H = Val(prjAtmRefMainfm.txtHumid.Text) * 0.01 'percent humidity 0-1
       
          '////////////////////MENAT's INDEX OF REFRACTION MODEL////////////////////////////////
       
          If prjAtmRefMainfm.optMenat.Value = True Then 'Or TMP(II) < 233 Or PRSR(II) < 800 Then
          
             'use simplified Menat refraction, pressure must be in mb, temperature in K
                    
             SIGMA = 1000# / wl 'wavelength in micrometers

             P1 = PRSR(II - 1)
             T1 = TMP(II - 1)
             RI1 = 1 + 0.000001 * (77.46 + 0.459 / (SIGMA * SIGMA)) * (P1 / T1) 'pressure in mb, T in Kelvin
             P2 = PRSR(II)
             T2 = TMP(II)
             RI2 = 1 + 0.000001 * (77.46 + 0.459 / (SIGMA * SIGMA)) * (P2 / T2)
            'radius of curvature
             DELV = ELV(II) - ELV(II - 1)
             RKAPPA = -2# * (RI2 - RI1) / (DELV * (RI2 + RI1))
             If RKAPPA <> 0 Then
                RC = 1# / RKAPPA
             Else 'radius of curvature is infinite since ray is not bending, so set it very large number
                RC = 1000# * RE
                End If
               
             den = (122.5 / 101325) * (P1 + P2) * 0.5 'pressures are in mb 'average density according to bruton's expression
       
          '//////////////////////////////////////Brutons's and general expression for index of refraction//////////////////////////////////////
       
          Else
          
             If AtmModel = 1 Then 'Bruton's model for atmsopheric Pressure
             
                xc = 450 'Bruton's very high value for the concentration of CO2
                
                A = 0.000012378847 'K^-2
                b = -0.019121316  'K^-1
                C = 33.93711047
                d = -6343.1645    'K
                
                ALPHA = 1.00062
                ßbeta = 0.0000000314  'Pa^-1,
                gamma = 0.00000056    '°C^-2
                
                'bottom of layer
                T = TMP(II - 1)
                P = 101325 * Exp(-ELV(II - 1) / 8400#)
            
                'saturation vapor pressure of water vapor in air at temperature T
                If (T >= 0) Then
                    SVP = Exp(A * T ^ 2 + b * T + C + d / T) 'Pa
                Else
                    SVP = 10 ^ (-2663.5 / T + 12.537)
                    End If
                
                'enhancement factor of water vapor in air
                F = ALPHA + ßbeta * P + gamma * (T - 273.15) ^ 2
                
                H = 0.04 * P / SVP 'Bruton's model of the humidity

                Call INDEX_CIDDOR(wl, T, P, H, xc, DA, DW, Index)
                DEN1 = DA + DW
                RI1 = Index
                
                'top of layer
                T = TMP(II)
                P = 101325 * Exp(-ELV(II) / 8400#)
            
                'saturation vapor pressure of water vapor in air at temperature T
                If (T >= 0) Then
                    SVP = Exp(A * T ^ 2 + b * T + C + d / T) 'Pa
                Else
                    SVP = 10 ^ (-2663.5 / T + 12.537)
                    End If
                
                'enhancement factor of water vapor in air
                F = ALPHA + ßbeta * P + gamma * (T - 273.15) ^ 2
                
                H = 0.04 * P / SVP 'Bruton's model of the humidity

                Call INDEX_CIDDOR(wl, T, P, H, xc, DA, DW, Index)
                DEN2 = DA + DW
                RI2 = Index
             
                DELV = ELV(II) - ELV(II - 1)
                RKAPPA = -2# * (RI2 - RI1) / (DELV * (RI2 + RI1))
                
                If RKAPPA <> 0 Then
                   RC = 1# / RKAPPA
                Else 'radius of curvature is infinite since ray is not bending, so set it very large number
                   RC = 1000# * RE
                   End If
                   
                den = (Exp(-ELV(II) / 8400#) + Exp(-ELV(II - 1) / 8400#)) / 2# 'Bruton's ignores Ciddor's values
                
             ElseIf AtmModel = 2 Then
             
                If UsingHSatmosphere Then
                
                   den = (PRSR(II - 1) + PRSR(II)) / (PRSR(0) * 2#)
                   
                   RI1 = IndexRefraction(II - 1)
                   RI2 = IndexRefraction(II)
                   DELV = ELV(II) - ELV(II - 1)
                   RKAPPA = -2# * (RI2 - RI1) / (DELV * (RI2 + RI1))
                   
                   If RKAPPA <> 0 Then
                      RC = 1# / RKAPPA
                   Else 'radius of curvature is infinite since ray is not bending, so set it to very large number
                      RC = 1000# * RE
                      End If
                      
                Else
             
                   'fulll models based on Ciddor's expressions for index of refraction and atmospheric densities
                   'note that range of validity of Ciddor's fit is only down to 800 mb of pressure to -40 C Celsius = 233.15 K
                   xc = 402 '2017 value of atmospheric concentration of CO2
                   
                   'bottom of layer
                   T = TMP(II - 1)
                   P = PRSR(II - 1) * 100#  'convert to pascals
                   'note that range of validity of Ciddor's fit is only down to 800 mb of pressure to -40 C Celsius = 233.15 K
                   Call INDEX_CIDDOR(wl, T, P, H, xc, DA, DW, Index)
                   DEN1 = DA + DW
                   RI1 = Index
                   
                   'top of layer
                   T = TMP(II)
                   P = PRSR(II) * 100#  ' convert to pascals
                   Call INDEX_CIDDOR(wl, T, P, H, xc, DA, DW, Index)
                   DEN2 = DA + DW
                   RI2 = Index
                
                   DELV = ELV(II) - ELV(II - 1)
                   RKAPPA = -2# * (RI2 - RI1) / (DELV * (RI2 + RI1))
                   
                   If RKAPPA <> 0 Then
                      RC = 1# / RKAPPA
                   Else 'radius of curvature is infinite since ray is not bending, so set it to very large number
                      RC = 1000# * RE
                      End If
                   
                   den = (DEN1 + DEN2) / 2#
                   
                   End If
                   
                End If
               
             End If
             
       Else
            
          RC = 1E+20
          den = 0#
          
          End If
         

End Sub
'The following function will return the inverse tangent in the proper
'quadrant determined by the signs of x and y.
'http://computer-programming-forum.com/16-visual-basic/f6b1e67cca79ee85.htm
Public Function Atan2(x As Double, y As Double) As Double
'-pi < Atan2 <= pi
'    If x > 0 Then Atan2 = Atn(y / x): Exit Function     '1st & 4th quadrants
'    If x < 0 And y > 0 Then Atan2 = Atn(y / x) + PI: Exit Function      '2nd quadrant
'    If x < 0 And y < 0 Then Atan2 = Atn(y / x) - PI: Exit Function      '3rd quadrant
'    If x = 0 And y > 0 Then Atan2 = PI / 2: Exit Function
'    If x = 0 And y < 0 Then Atan2 = -PI / 2

If x Then
        Atan2 = Atn(y / x) - (x > 0) * 3.14159265358979
    Else
        Atan2 = 1.5707963267949 + (y > 0) * 3.14159265358979
End If
End Function
'The following function will return the inverse tangent in the proper
'quadrant determined by the signs of x and y.
'http://computer-programming-forum.com/16-visual-basic/f6b1e67cca79ee85.htm
Function Atan2_2(x As Double, y As Double) As Double
'-pi < Atan2 <= pi
    Const pi As Double = 3.14159265358979
    If x > 0 Then Atan2_2 = Atn(y / x): Exit Function     '1st & 4th quadrants
    If x < 0 And y > 0 Then Atan2_2 = Atn(y / x) + pi: Exit Function      '2nd quadrant
    If x < 0 And y < 0 Then Atan2_2 = Atn(y / x) - pi: Exit Function      '3rd quadrant
    If x = 0 And y > 0 Then Atan2_2 = pi / 2: Exit Function
    If x = 0 And y < 0 Then Atan2_2 = -pi / 2
End Function

Public Function DistTrav(lat_0 As Double, lon_0 As Double, lat_1 As Double, lon_1 As Double, Mode%) As Double
   
   'calclates angular distance (central angle) on spherical earth from (lat0,lon1) to (lat1,lon1)
   'to convert to kilometers multiply by Rearthkm
   'to convert to meters, multiply by Reathkm * 1000.0
   
   'mode% = 0 'passed coordinates are in radians
   'mode% = 1 'passed coordinaes are in degrees
   'mode% = 2 'passed coordinates are in radians, use Vincenty formula source: http://en.wikipedia.org/wiki/Great-circle_distance
   'mode% = 3 'passed coordinates are in degrees, use Vincenty formula
   
   Dim AA As Double, bb As Double, cc As Double
   Dim lat0 As Double, lat1 As Double, lon0 As Double, lon1 As Double
   Dim DifDist As Double, ccc
   
    pi = 4# * Atn(1#) '3.141592654
    CONV = pi / (180# * 60#) 'conversion of minutes of arc to radians
    cd = pi / 180# 'conversion of degrees to radians
   
   If Mode% = 0 Or Mode% = 2 Then
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

'   If mode% <= 1 Then
'      aa = Sin((lat0 - lat1) * 0.5)
'      bb = Sin((lon0 - lon1) * 0.5)
'      DistTrav = 2# * DASIN(Sqr(aa * aa + Cos(lat0) * Cos(lat1) * bb * bb)) 'central angle in radians

'   ElseIf mode% > 1 Then 'use Vincenty formula (more accurate for small distances and for 32 bit calculations)
      AA = Cos(lat1) * Sin(lon1 - lon0)
      bb = Cos(lat0) * Sin(lat1) - Sin(lat0) * Cos(lat1) * Cos(lon1 - lon0)
      cc = Sin(lat0) * Sin(lat1) + Cos(lat0) * Cos(lat1) * Cos(lon1 - lon0)
      DistTrav = Atan2_2(cc, Sqr(AA * AA + bb * bb)) 'central angle in radians
      
'      End If
      
   
End Function
'Public Function fun_(h__ As Double, zz_1 As zz, num_Layers As Integer) As Double
'---------------------------------------------------------------------------------------
' Procedure : fun_
' Author    : chaim
' Date      : 8/13/2021
' Purpose   : inputs: 1. height of ray (km)
'                     2.elevation (km),temperature (deg K), pressure (mb) array zz_1
'                     2. number of layers
'                     3. distance along circumference of earth (km)
'           : outputs: fun_ = variable part of index of refraction (unitless)
'                      Temparature, T, at the height, h__ (degrees Kelvin)
'                      Pressure, P, at the height, h__ (mb)
'---------------------------------------------------------------------------------------
'
Public Function fun_(h__ As Double, zz_1() As zz, num_Layers As Integer, Dist As Double, T As Double, P As Double) As Double
'{
'    /* System generated locals */
     Dim ret_val As Double, d__1 As Double, GrndHgt As Double

'    /* Builtin functions */
'    'double pow_dd(double *, double *)
'
'    /* Local variables */
    Dim m As Integer, n As Integer
'    Dim P As Double, T As Double

   'note that if height, h__ is greater than the last layer,
   'then thsi routine uses the temperature and pressure of the last layer

'    For n = 1 To num_Layers
'        m = n - 1
'        If (h__ - zz_1.hj(n - 1) <= 0#) Then
'            Exit For
'            'GoTo L15
''        Else
''            GoTo L10
'            End If
''L10:
'
'    Next n
''L15:
'    If m < 1 Then 'hit the earth
'       ret_val = -1
'       fun_ = ret_val
'       Exit Function
'       End If
'
'    T = zz_1.tj(m - 1) + zz_1.AT(m - 1) * (h__ - zz_1.hj(m - 1))
'    d__1 = T / zz_1.tj(m - 1)
'
'    If (zz_1.tj(m) <> zz_1.tj(m - 1)) Then 'non-isothermic region
'
'        P = zz_1.pj(m - 1) * d__1 ^ zz_1.ct(m - 1)
'        ret_val = P * 0.000079 / T
'
'    Else 'interpolate pressure between the pressure of the two adjoining layers
'
'        P = zz_1.pj(m - 1) + zz_1.ct(m - 1) * (h__ - zz_1.hj(m - 1))
'        ret_val = P * 0.000079 / T
'        End If
        
'''''''''''''''''''''new format 072521''''''''''''''''''''''''''''''
    
'////////////////ground hugging added to Menat atmospheres on 081321 //////////////////////////////
    If prjAtmRefMainfm.OptionSelby.Value = True And prjAtmRefMainfm.chkHgtProfile.Value = vbChecked And prjAtmRefMainfm.chkDruk.Value = vbChecked And Dist <> -1 Then
        GrndHgt = DistModel(Dist) 'assumes that atmospheric layers are displaced vertically
        GrndHgt = GrndHgt * 0.001
        
        '///////////////fixed on 081321////////////////////
        If h__ < GrndHgt Then h__ = GrndHgt 'light doesn't travel underground!!!!
        '///////////////////////////////////////////
        
    Else
        GrndHgt = 0#
        End If
'
'    found% = 0
'    For n = 2 To num_Layers - 1
'        m = n - 1
'        If (h__ - zz_1(n - 1).hj - GrndHgt <= 0#) Then
'            found% = 1
'            Exit For
'            End If
''L10:
'
'    Next n
''L15:
'    If found% = 0 Then
'        'hit the ground
'       ret_val = -1
'       fun_ = ret_val
'       Exit Function
'       End If
       
    For n = 1 To num_Layers
        m = n - 1
        If (h__ - zz_1(n - 1).hj - GrndHgt < 0#) Then
            Exit For
        ElseIf (h__ - zz_1(n - 1).hj - GrndHgt = 0#) Then
            m = 1
            Exit For

            'GoTo L15
'        Else
'            GoTo L10
            End If
'L10:

    Next n
L15:
    If m < 1 Then 'hit the earth
       ret_val = -1
       fun_ = ret_val
       Exit Function
       End If
       
    T = zz_1(m - 1).tj + zz_1(m - 1).AT * (h__ - zz_1(m - 1).hj - GrndHgt)
    d__1 = T / zz_1(m - 1).tj

    If (zz_1(m).tj + GrndHgt <> zz_1(m - 1).tj + GrndHgt) Then 'non-isothermic region

        P = zz_1(m - 1).pj * d__1 ^ zz_1(m - 1).ct
        ret_val = P * 0.000079 / T
        
    Else 'interpolate pressure between the pressure of the two adjoining layers

        P = zz_1(m - 1).pj + zz_1(m - 1).ct * (h__ - zz_1(m - 1).hj - GrndHgt)
        ret_val = P * 0.000079 / T
        End If

'/* L25: */
'    return ret_val
    fun_ = ret_val
End Function
'Public Sub layers_int(h__ As Double, zz_1 As zz, num_Layers As Integer, P As Double, T As Double)
Public Sub layers_int(h__ As Double, zz_1() As zz, num_Layers As Integer, P As Double, T As Double)
'increase resolution of Menat atmospheres
Dim n As Integer, m As Integer, d__1 As Double

'    For n = 1 To num_Layers
'       m = n - 1
'       If (h__ - zz_1.hj(n - 1) <= 0#) Then
'          Exit For
'          End If
'    Next n
'
'    If m < 1 Then 'hit the earth, just return the ground temp and pressure
'       P = PRSR(0)
'       T = TMP(0)
'       Exit Sub
'       End If
'
'    T = zz_1.tj(m - 1) + zz_1.AT(m - 1) * (h__ - zz_1.hj(m - 1))
'    d__1 = T / zz_1.tj(m - 1)
'
'    If (zz_1.tj(m) <> zz_1.tj(m - 1)) Then 'non-isothermic region
'
'        P = zz_1.pj(m - 1) * d__1 ^ zz_1.ct(m - 1)
'
'    Else 'interpolate pressure between the pressure of the two adjoining layers
'
'        P = zz_1.pj(m - 1) + zz_1.ct(m - 1) * (h__ - zz_1.hj(m - 1))
'        End If
                
''''''''''''''''''''new format 072521''''''''''''''''''''''''
    For n = 1 To num_Layers
       m = n - 1
       If (h__ - zz_1(n - 1).hj <= 0#) Then
          Exit For
          End If
    Next n

    If m < 1 Then 'hit the earth, just return the ground temp and pressure
       P = PRSR(0)
       T = TMP(0)
       Exit Sub
       End If
   
    T = zz_1(m - 1).tj + zz_1(m - 1).AT * (h__ - zz_1(m - 1).hj)
    d__1 = T / zz_1(m - 1).tj

    If (zz_1(m).tj <> zz_1(m - 1).tj) Then 'non-isothermic region

        P = zz_1(m - 1).pj * d__1 ^ zz_1(m - 1).ct
        
    Else 'interpolate pressure between the pressure of the two adjoining layers

        P = zz_1(m - 1).pj + zz_1(m - 1).ct * (h__ - zz_1(m - 1).hj)
        End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MyCallback
' Author    : Dr-John-K-Hall
' Date      : 9/11/2015
' Purpose   : Callback function from dll used to move the progrress bar
'---------------------------------------------------------------------------------------
'
Public Sub MyCallback(ByVal parm As Long)

   On Error GoTo MyCallback_Error

   Call UpdateStatus(prjAtmRefMainfm, prjAtmRefMainfm.picProgBar, 1, parm)
   
   DoEvents

   On Error GoTo 0
   Exit Sub

MyCallback_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure MyCallback of Module modHardy"
End Sub

Public Function FunTPI(x As Double) As Double
    'removes multiples of 2*pi
    pi2 = 2 * pi
    FunTPI = (x / (pi2) - Int(x / (pi2))) * pi2
End Function
Public Sub casgeo(kmx, kmy, lg, lt)

        If kmy > 9999 Then
            g1# = kmy - 1000000
        Else
            g1# = kmy * 1000#
            End If
            
        If kmx < 9999 Then
            G2# = kmx * 1000#
        Else
            G2# = kmx
            End If
            
        r# = 57.2957795131
        B2# = 0.03246816
        f1# = 206264.806247096
        S1# = 126763.49
        s2# = 114242.75
        e4# = 0.006803480836
        C1# = 0.0325600414007
        C2# = 2.55240717534E-09
        C3# = 0.032338519783
        X1# = 1170251.56
        yy1# = 1126867.91
        yy2# = g1#
'       GN & GE
        X2# = G2#
        If (X2# > 700000!) Then GoTo ca5
        X1# = X1# - 1000000#
ca5:    If (yy2# > 550000#) Then GoTo ca10
        yy1# = yy1# - 1000000#
ca10:   X1# = X2# - X1#
        yy1# = yy2# - yy1#
        D1# = yy1# * B2# / 2#
        O1# = s2# + D1#
        O2# = O1# + D1#
        a3# = O1# / f1#
        A4# = O2# / f1#
        B3# = 1# - e4# * Sin(a3#) ^ 2
        B4# = B3# * Sqr(B3#) * C1#
        C4# = 1# - e4# * Sin(A4#) ^ 2
        C5# = Tan(A4#) * C2# * C4# ^ 2
        C6# = C5# * X1# ^ 2
        D2# = yy1# * B4# - C6#
        C6# = C6# / 3#
'LAT
        l1# = (s2# + D2#) / f1#
        R3# = O2# - C6#
        R4# = R3# - C6#
        R2# = R4# / f1#
        A2# = 1# - e4# * Sin(l1#) ^ 2
        lt = r# * (l1#)
        A5# = Sqr(A2#) * C3#
        d3# = X1# * A5# / Cos(R2#)
' LON
        lg = r# * ((S1# + d3#) / f1#)
'       THIS IS THE EASTERN HEMISPHERE!
        lg = -lg

End Sub
Public Sub Temperatures(lat As Double, lon As Double, MinTemp() As Integer, AvgTemp() As Integer, ier As Integer)

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

Dim tncols As Long, IKMY&, IKMX&, numrec&, IO%, Tempmode%

FilePathBil = App.Path & "\WorldClim_bil"
If Dir(FilePathBil, vbDirectory) <> sEmpty Then
    FileNameBil = FilePathBil
Else
    Call MsgBox("Can't find the bil directory at the following location:" _
                & vbCrLf & FileNameBil _
                , vbExclamation, "Missing bil file directory")
    ier = -1
    Exit Sub
    End If
'first extract minimum temperatures

 Tempmode% = 0
T50:
 If Tempmode% = 0 Then 'minimum temperatures to be used for sunrise calculations
    FilePathBil = App.Path & "/WorldClim_bil" & "/min_"
 ElseIf Tempmode% = 1 Then 'average temperatures to be used for sunset calculations
    FilePathBil = App.Path & "/WorldClim_bil" & "/avg_"
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
       filein% = FreeFile
       Open FileNameBil For Binary As #filein%
   
        y = lat
        x = lon
        
        IKMY& = CLng((ULYMAP - y) / YDIM) + 1
        IKMX& = CLng((x - ULXMAP) / XDIM) + 1
        tncols = NCOLS
        numrec& = (IKMY& - 1) * tncols + IKMX&
        Get #filein%, (numrec& - 1) * 2 + 1, IO%
        If IO% = NODATA Then IO% = 0#
        If Tempmode% = 0 Then
            MinTemp(i) = IO%
        ElseIf Tempmode% = 1 Then
            AvgTemp(i) = IO%
            End If
            
        Close #filein%
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
     End If
     
End Sub

Public Sub worldheights(lg, lt, hgt)
   Dim leros As Long, lmag As Long
   Dim world As Boolean, srtmdtm As String, DTMflag As Integer
   Dim NCOLS As Integer, NROWS As Integer, AA$, j%
   
   On Error GoTo worlderror
   
   If lt > 90 Or lt < -90 Or lg < -180 Or lg > 180 Then Exit Sub
   
   'make 30 m SRTM default for this program, residing in c directory
   world = True
   srtmdtm = "c"
   DTMflag = 1
   
   'check if have correct CD in the drive, if not present error message
   If (world = False And IsraelDTMsource% = 1) Or (DTMflag > 0 And (lt >= -60 And lt <= 61)) Then 'SRTM
      
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
      numCD% = worldcd%(ny% * 9 + nx% + 1)
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
         ret = SetWindowPos(mapEROSDTMwarn.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
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
   tncols = NCOLS
   C% = worldfnum%
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
    TR = T1 * 256 + T2
    integ1& = TR
mer130:
    If IO% < 0 Or integ1& > elevmax% Then 'modular division failed use HEX swap
       a0$ = LTrim$(RTrim$(Hex$(IO%)))
       AA$ = sEmpty
       'swap the two bytes using their hex representation
       'e.g., ABCD --> CDAB, etc.
       If Len(a0$) = 4 Then
          A1$ = Mid$(a0$, 1, 2)
          A2$ = Mid$(a0$, 3, 2)
          If Mid$(A2$, 3, 1) = "0" And Mid$(A2$, 3, 2) <> "0" Then
             A2$ = Mid$(a0$, 4, 1)
          ElseIf Mid$(A2$, 3, 1) = "0" And Mid$(A2$, 3, 2) = "0" Then
             A2$ = sEmpty
             End If
          AA$ = A2$ + A1$
       ElseIf Len(a0$) = 3 Then
          A1$ = "0" + Mid$(a0$, 1, 1)
          A2$ = Mid$(a0$, 2, 2)
          If Mid$(a0$, 2, 1) = "0" Then A2$ = Mid$(a0$, 3, 1)
          AA$ = A2$ + A1$
       ElseIf Len(a0$) = 2 Or Len(a0$) = 1 Then
          A1$ = "00"
          A2$ = a0$
          AA$ = A2$ + A1$
          End If
    
        'convert swaped hexadecimel to an integer value
        leng% = Len(LTrim$(RTrim$(AA$)))
        integ1& = 0
        For j% = leng% To 1 Step -1
            v$ = Mid$(LTrim$(RTrim$(AA$)), j%, 1)
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
   ret = SetWindowPos(mapPictureform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   If response = vbCancel Then Exit Sub
   Resume
End Sub

Public Sub heights(kmx, kmy, hgt2)

      Dim israeldtm As String
      Dim Jg%, IG%
      Dim IR%, IC%, NCOL%, IFN&

      israeldtm = "c" 'jkh's dtm root directory
      
      On Error GoTo g35
      
'      If IsraelDTMsource% = 1 Then 'convert to long,lat and use SRTM extraction
         'Call casgeo(kmx, kmy, lgh, lth)

'         If ggpscorrection = True Then 'apply conversion from Clark geoid to WGS84
'            Dim N As Long
'            Dim E As Long
'            Dim lat As Double
'            Dim lon As Double
'            N = kmy
'            E = kmx
'            Call ics2wgs84(N, E, lat, lon)
'            lgh = lon
'            lth = lat
'            'Call casgeo(kmx, kmy, lgh, lth)
'            ggpscorrection = False
'         Else
'            Call casgeo(kmx, kmy, lgh, lth)
'            End If
'
'         Call worldheights(lgh, lth, hgt2)
'         GoTo g99
'         End If
      
      If kmx > 1000 Then kmx = kmx * 0.001
      If kmy > 1000 Then
        kmy = (kmy - 1000000) * 0.001
        End If
      IKMX& = Int((kmx + 20!) * 40!) + 1
      IKMY& = Int((380! - kmy) * 40!) + 1
      NROW% = IKMY&: NCOL% = IKMX&

'       GETZ FINDS THE HEIGHT OF A POINT AT THE NORW AND NCOL FROM 380N
'       AND -20E WHERE 1,1 IS THAT CORNER POINT
'       FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
g15:    Jg% = 1 + Int((NROW% - 2) / 800)
        IG% = 1 + Int((NCOL% - 2) / 800)
        CHMNE = CHMAP(IG%, Jg%)
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
        IC% = NCOL% - (IG% - 1) * 800
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

Public Function DistModel(Dist As Double) As Double
   Dim DistModelCoef(12) As Double, i As Integer
   Dim dist1 As Double, dist2 As Double, hgt1 As Double, hgt2 As Double, numHgts As Long
   
   DistModel = 0
   
   If prjAtmRefMainfm.txtHgtProfile.Text = "External terrain profile file path (m,m)" Then
        '10th order polynomial fit to height vs distance profile between Rabbi Druk's obsdervation place and due East for 80 km
        
        If (Dist > 0 And Dist < 80000 And Dist <> -1) Then
            'at zero azimuth, i.e., due East
            DistModelCoef(0) = 719.595830740319
            DistModelCoef(1) = -4.81685432649849E-02
            DistModelCoef(2) = -5.00526065276487E-06
            DistModelCoef(3) = 1.73763617916305E-09
            DistModelCoef(4) = -1.88656492725314E-13
            DistModelCoef(5) = 9.71806281191486E-18
            DistModelCoef(6) = -2.741321305037E-22
            DistModelCoef(7) = 4.50300692209001E-27
            DistModelCoef(8) = -4.30862997078586E-32
            DistModelCoef(9) = 2.22937939762269E-37
            DistModelCoef(10) = -4.82303278939318E-43
            DistModelCoef(11) = 0#

            DistModel = 0
            For i = 0 To 11
               DistModel = DistModel + DistModelCoef(i) * (Dist ^ CDbl(i))
            Next i
            
            If DistModel < -430 Then DistModel = -430 'can't be lower than water level of Dead Sea
            
'             'for winter at 30 degrees azimuth
'             DistModelCoef(0) = 832.772132844035
'             DistModelCoef(1) = -0.244714550731683
'             DistModelCoef(2) = 6.84359980795917E-05
'             DistModelCoef(3) = -1.00460673653427E-08
'             DistModelCoef(4) = 8.25791986979131E-13
'             DistModelCoef(5) = -4.26494539579878E-17
'             DistModelCoef(6) = 1.44294055937789E-21
'             DistModelCoef(7) = -3.21094617296788E-26
'             DistModelCoef(8) = 4.62408002443649E-31
'             DistModelCoef(9) = -4.13125184516663E-36
'             DistModelCoef(10) = 2.07681721286578E-41
'             DistModelCoef(11) = -4.4844262868453E-47
'
'            DistModel = 0
'            For i = 0 To 11
'               DistModel = DistModel + DistModelCoef(i) * (Dist ^ CDbl(i))
'            Next i
            
         ElseIf Dist = 0 Then
            DistModel = 800.5 'last height of profile in Harei Moav
         ElseIf Dist = -1 Or Dist > 80000 Then
            DistModel = 0#
            End If
        
   Else 'use model values to interpolate height
      If Dir(prjAtmRefMainfm.txtHgtProfile.Text) <> vbNullString Then
         DistModel = 0
         filein = FreeFile
         Open prjAtmRefMainfm.txtHgtProfile.Text For Input As #filein
         numHgts = 0
         Input #filein, dist1, hgt1
100:
        If EOF(filein) Then GoTo 900
        Input #filein, dist2, hgt2
        If Dist >= dist1 And Dist < dist2 And (dist2 - dist1) <> 0 Then
           DistModel = ((hgt2 - hgt1) / (dist2 - dist1)) * (Dist - dist1) + hgt1
           GoTo 900
        Else
           hgt1 = hgt2
           dist1 = dist2
           GoTo 100
           End If
           
        End If
900:    Close #filein
      End If
    
End Function

' Return a list of files that match the patterns.
Public Function FindFiles(ByVal dirname As String, ByVal patterns As String) As Collection
Dim pattern_array() As String
Dim pattern As Variant
Dim files As Collection
Dim filename As String

    ' Separate the patterns.
    pattern_array = Split(patterns, ";")

    Set files = New Collection
    ' Loop through the files in the directory.
    filename = Dir$(dirname)
    Do While Len(filename) > 0
        ' See if the name matches any pattern.
        For Each pattern In pattern_array
            If LCase$(filename) Like LCase$(pattern) Then
                files.Add filename
                Exit For
            End If
        Next pattern

        filename = Dir$()
    Loop

    Set FindFiles = files
End Function
'---------------------------------------------------------------------------------------
' Procedure : CalcVDWRef
' Author    : Dr-John-K-Hall
' Date      : 3/24/2020
' Purpose   : Calculate VDW refraction in mrad for any lat (deg),lon (deg),daynumber (1-366), year, viewangle (deg), height (m)
'---------------------------------------------------------------------------------------
'
Public Function CalcVDWRef(lat As Double, lon As Double, height As Double, DayNumber As Integer, year As Integer, _
                           viewangle As Double) As Double
                           
   Dim Coef(4, 10) As Double, ref As Double, VA As Double
   Dim CA(10) As Double
   Dim vbweps(6) As Double, vdwref(6) As Double, ier As Integer, TK As Double
   Dim VDWSF As Double, VDWALT As Double, lnhgt As Double, pi As Double, cd As Double
   Dim sumref(7) As Double, winref(7) As Double, sumrefo As Double, winrefo As Double
   Dim weather%, ns1 As Integer, ns2 As Integer, ns3 As Integer, ns4 As Integer
   Dim MT(12) As Integer, AT(12) As Integer
   Dim RefExponent As Double, RefNorms As Double, refexp(1) As Double, refnorm(4) As Double
   Dim refFromExp As Double, RefFromHorToInfExp As Double, RefFromHorToInfFit As Double
   Dim CalculatedRefFromHgtToHoriz As Double, RefFromHgtToHorizonExp As Double
    
   On Error GoTo CalcVDWRef_Error

    weather% = 5  ' = 3 for mixed winter-summer Menat atmospheres
                  ' = 5 for van der Werf standard atmosphere
                  
   If lat = 0 And lon = 0 Then 'use Beit Dagan's coordinates
     lat = 32#
     lon = 34.81
     End If
     
  If height = 0 Then 'use Beit Dagan's height
     height = 35 'meters
     End If
    
   'constants
   pi = 4 * Atn(1)
   pi2 = 2 * pi
   ch = 360# / (pi2 * 15)  '57.29578 / 15  'conv rad to hr
   cd = pi / 180#  'conv deg to rad
   
   If weather% = 3 Then
   
'                  weather similar to Eretz Israel */
        ns1 = 85
        ns2 = 290
'               Times for ad-hoc fixes to the visible and astr. sunrise */
'               (to fit observations of the winter netz in Neve Yaakov). */
'               This should affect  sunrise and sunset equally. */
'               However, sunset hasn't been observed, and since it would */
'               make the sunset times later, it's best not to add it to */
'               the sunset times as a chumrah. */
        ns3 = 30
        ns4 = 330
                    
   'Menat atmospheres
    sumrefo = 8.899
    sumref(0) = 2.791796282
    sumref(1) = 0.5032840405
    sumref(2) = 0.001353422287
    sumref(3) = 0.0007065245866
    sumref(4) = 1.050981251
    sumref(5) = 0.4931095603
    sumref(6) = -0.02078600882
    sumref(7) = -0.00315052518

    winrefo = 9.85
    winref(0) = 2.779751597
    winref(1) = 0.5040818795
    winref(2) = 0.001809029729
    winref(3) = 0.0007994475831
    winref(4) = 1.188723157
    winref(5) = 0.4911777019
    winref(6) = -0.0221410531
    winref(7) = -0.003454047139
    
    If height > 0 Then lnhgt = Log(height * 0.001)
    If dy <= ns1 Or dy >= ns2 Then 'winter refraction
       ref = 0: EPS = 0
       If height <= 0 Then GoTo 690
       ref = Exp(winref(4) + winref(5) * lnhgt + _
           winref(6) * lnhgt * lnhgt + winref(7) * lnhgt * lnhgt * lnhgt)
'           ref = ((winref(2, n2%) - winref(2, n1%)) / 2) * (hgt - h1) + winref(2, n1%)
       EPS = Exp(winref(0) + winref(1) * lnhgt + _
            winref(2) * lnhgt * lnhgt + winref(3) * lnhgt * lnhgt * lnhgt)
'           eps = ((winref(1, n2%) - winref(1, n1%)) / 2) * (hgt - h1) + winref(1, n1%)
690    Air = 90 * cd + (EPS + ref + winrefo) / 1000
       AirMenatRefDip = EPS + ref + winrefo
'       lblMenatAir.Caption = AirMenatRefDip & " mrad"
    ElseIf dy > ns1 And dy < ns2 Then
       ref = 0: EPS = 0
       If height <= 0 Then GoTo 695
       ref = Exp(sumref(4) + sumref(5) * lnhgt + _
           sumref(6) * lnhgt * lnhgt + sumref(7) * lnhgt * lnhgt * lnhgt)
       'ref = ((sumref(2, n2%) - sumref(2, n1%)) / 2) * (hgt - h1) + sumref(2, n1%)
       EPS = Exp(sumref(0) + sumref(1) * lnhgt + _
            sumref(2) * lnhgt * lnhgt + sumref(3) * lnhgt * lnhgt * lnhgt)
'           eps = ((sumref(1, n2%) - sumref(1, n1%)) / 2) * (hgt - h1) + sumref(1, n1%)
695    Air = 90 * cd + (EPS + ref + sumrefo) / 1000
       AirMenatRefDip = EPS + ref + sumrefo
'       lblMenatAir.Caption = AirMenatRefDip & " mrad"
       End If
    End If
    
  If weather% = 5 Then
    'use Rabbi Druk's observation place latitude and longitude
    'lon = 35.237435642287 '81333572129 '-35.238456 '5 'longitude at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
    'lat = 31.748552568177 '8959288296 '31.749942 'latitude at at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
    
    
    Call Temperatures(lat, lon, MT, AT, ier)
    
       'determine the minimum and average temperature for this day for current place
       'use Meeus's forumula p. 66 to convert daynumber to month,
       'no need to interpolate between temepratures -- that is overkill
       'take year as regular
       yl = 365
       dy = DayNumber
       k% = 2
       If (yl = 366) Then k% = 1
       mMonth% = Int(9 * (k% + dy) / 275 + 0.98)
    
'       If optMin.Value = True Then
          TK = MT(mMonth%) + 273.15 'mean minimum temperature for this month in degrees Kelvin
'       Else
'          TK = AT(MMonth%) + 273.15 'mean minimum temperature for this month in degrees Kelvin
'          End If
       
       'calculate van der Werf temperature scaling factor for refraction
       VDWSF = (288.15 / TK) ^ 1.7081
       'calculate van der Werf scaling factor for view angles
       VDWALT = (288.15 / TK) ^ 0.69
    
    'vdW dip angle vs height polynomial fit coefficients
    vbweps(0) = 2.77346593151086
    vbweps(1) = 0.497348466526589
    vbweps(2) = 2.53874620975453E-03
    vbweps(3) = 6.75587054940366E-04
    vbweps(4) = 3.94973974451576E-05
  
    'vdW atmospheric refraction vs height polynomial fit coefficients
    vdwref(0) = 1.16577538442405
    vdwref(1) = 0.468149166683532
    vdwref(2) = -0.019176833246687
    vdwref(3) = -4.8345814464145E-03
    vdwref(4) = -4.90660400743218E-04
    vdwref(5) = -1.60099622077352E-05

    Coef(0, 0) = 9.56267125268496
    Coef(1, 0) = -8.6718429211079E-04
    Coef(2, 0) = 3.1664677349513E-08
    Coef(3, 0) = -2.04067678864827E-13
    Coef(4, 0) = -6.21413591282229E-17
    Coef(0, 1) = -3.54681762174248
    Coef(1, 1) = 3.05885370538294E-04
    Coef(2, 1) = -3.48413989765623E-09
    Coef(3, 1) = -3.27424677578751E-12
    Coef(4, 1) = 4.85180396156723E-16
    Coef(0, 2) = 1.00487516923555
    Coef(1, 2) = -7.12411305623716E-05
    Coef(2, 2) = -1.30264294792463E-08
    Coef(3, 2) = 6.08386198681256E-12
    Coef(4, 2) = -8.26564806865056E-16
    Coef(0, 3) = -0.234676117102
    Coef(1, 3) = 4.23105602906229E-06
    Coef(2, 3) = 1.25823603467313E-08
    Coef(3, 3) = -4.77064146898649E-12
    Coef(4, 3) = 6.38020633504241E-16
    Coef(0, 4) = 4.55474692911979E-02
    Coef(1, 4) = 4.20546127818185E-06
    Coef(2, 4) = -5.71051596715397E-09
    Coef(3, 4) = 2.05052061222564E-12
    Coef(4, 4) = -2.73486893484326E-16
    Coef(0, 5) = -7.07867490693562E-03
    Coef(1, 5) = -1.77071205323987E-06
    Coef(2, 5) = 1.51606775168871E-09
    Coef(3, 5) = -5.33770875527936E-13
    Coef(4, 5) = 7.12762362198471E-17
    Coef(0, 6) = 8.32295487796478E-04
    Coef(1, 6) = 3.56012191465623E-07
    Coef(2, 6) = -2.51083454939817E-10
    Coef(3, 6) = 8.78112651062853E-14
    Coef(4, 6) = -1.17556897803785E-17
    Coef(0, 7) = -6.96285190742393E-05
    Coef(1, 7) = -4.19488560840601E-08
    Coef(2, 7) = 2.62823518700555E-11
    Coef(3, 7) = -9.18795592984757E-15
    Coef(4, 7) = 1.23368951173372E-18
    Coef(0, 8) = 3.85246830558751E-06
    Coef(1, 8) = 2.93954176835877E-09
    Coef(2, 8) = -1.6910427905383E-12
    Coef(3, 8) = 5.92963984786967E-16
    Coef(4, 8) = -7.98575122972107E-20
    Coef(0, 9) = -1.25306160093963E-07
    Coef(1, 9) = -1.13531297031204E-10
    Coef(2, 9) = 6.10692317760114E-14
    Coef(3, 9) = -2.15222164778055E-17
    Coef(4, 9) = 2.90679288057578E-21
    Coef(0, 10) = 1.80519843190424E-09
    Coef(1, 10) = 1.8626817474919E-12
    Coef(2, 10) = -9.47783437562306E-16
    Coef(3, 10) = 3.36109376083127E-19
    Coef(4, 10) = -4.55153453427451E-23

    VA = viewangle
     
    If VA = 0 Then
     
         If height >= 0 Then
        'calculate the refraction for the observer's height looking at an apparent zero view angle
           
           CA(0) = Coef(0, 0) + Coef(1, 0) * height + Coef(2, 0) * (height ^ 2) + Coef(3, 0) * (height ^ 3) + Coef(4, 0) * (height ^ 4)
         ElseIf height < 0 Then
           CA(0) = 9.56267125268573 - 8.67184292115619E-04 * height + 3.1664677356332E-08 * (height ^ 2) _
                  - 2.0406768223753E-13 * (height ^ 3) - 6.21413585867198E-17 * (height ^ 4)
           
           End If
           
           ref = CA(0)
           
    Else
        
        'calculate for range of view angles
        ref = 0#
        For i = 0 To 10
            CA(i) = Coef(0, i) + Coef(1, i) * height + Coef(2, i) * (height ^ 2) + Coef(3, i) * (height ^ 3) + Coef(4, i) * (height ^ 4)
        Next i
        For i = 0 To 10
            ref = ref + CA(i) * VA ^ i
        Next i
         
        End If
               
   RefFromHorToInfExp = ref 'calculated from exponent
   CalcVDWRef = ref
   Exit Function
'        lblref1.Caption = Str(Ref) & " mrad"
               
                
    '    now calculate refraction from the observer's height to the horizon, as well as the geometric dip angle, then the total refraction and dip.
    '/*  All refraction terms have units of mrad */
    
     d__2 = 288.15 / TK
     VDWSF = d__2 ^ 1.7081
'/*          calculate van der Werf scaling factor for view angles */
     VDWALT = d__2 ^ 0.69
     
     If (height <= 0#) Then GoTo L690
     lnhgt = Log(height * 0.001)
    
     ref2 = Exp(vdwref(0) + vdwref(1) * lnhgt + vdwref(2) * lnhgt * _
         lnhgt + vdwref(3) * lnhgt * lnhgt * lnhgt + vdwref(4) _
         * lnhgt * lnhgt * lnhgt * lnhgt + vdwref(5) * lnhgt * _
         lnhgt * lnhgt * lnhgt * lnhgt)
     EPS = Exp(vbweps(0) + vbweps(1) * lnhgt + vbweps(2) * lnhgt * _
         lnhgt + vbweps(3) * lnhgt * lnhgt * lnhgt + vbweps(4) _
         * lnhgt * lnhgt * lnhgt * lnhgt)
'/*         now add the all the contributions together due to the observer's height */
'/*         along with the value of atm. ref. from hgt<=0, where view angle, a1 = 0 */
'/*         for this calculation, leave the refraction in units of mrad */

    
'        lbleps.Caption = eps & " mrad"
'        lblref2.Caption = VDWSF * ref2 & " mrad"
            
        RefFromHorToInfFit = VDWSF * ref2 '+ RefFromHorToInf  '& " mrad" 'two components:  VDWSF*ref2 is contrib. from height to horizon, RefFromHorToInf is contribution from horizon to infinity
        CalcVDWRef = 0.5 * (RefFromHorToInfExp + RefFromHorToInfFit) 'take weighted average of them
        Exit Function
L690:
        A1 = 0#
        Air = cd * 90# + (EPS + VDWSF * (ref2 + 9.56267125268496)) / 1000#
        
        TotalRefWithDip = EPS + VDWSF * (ref2 + 9.56267125268496) '& " mrad"
        
        A1 = (ref2 + ref) / 1000#
'/*         leave a1 in radians */
        A1 = Atn(Tan(A1) * VDWALT)
        
        'now height dependent ref determinaton
        ' TR_VDW_200 -3000 - Ref.csv
        'Polynomical coeficients of Plot program's LS fit to 1th degree polynomial. vdW ref exponent b vs height(m), where ref= a* (288.15/Tk)**b
        refexp(0) = 2.20734384287553
        refexp(1) = -2.86255933358013E-05
        '========================================
        'TR_VDW_200 -3000 - Ref.csv
        'Polynomical coeficients of Plot program's LS fit to 4th degree polynomial. vdW ref normalization a vs height, where ref=a*(288.15/Tk)**b
        refnorm(0) = 0.767089721048164
        refnorm(1) = 3.89787475596774E-03
        refnorm(2) = -2.02184590999692E-06
        refnorm(3) = 6.5747954161702E-10
        refnorm(4) = -8.27402155995415E-14
        RefExponent = refexp(0) + refexp(1) * height
        RefNorms = refnorm(0) + refnorm(1) * height + refnorm(2) * height ^ 2 + refnorm(3) * height ^ 3 + refnorm(4) * height ^ 4
        refFromExp = RefNorms * (288.15 / TK) ^ RefExponent
'        lblrefexponent.Caption = refFromExp & " mrad"
        TotalRefFromHgtExp = refFromExp
        Air = cd * 90# + (EPS + refFromExp + VDWSF * 9.56267125268496) / 1000#
        
        TotalGeoRef = EPS + refFromExp + VDWSF * 9.56267125268496 '& " mrad"
        Exit Function
        End If

   On Error GoTo 0
   Exit Function

CalcVDWRef_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure CalcVDWRef of Module modBurtonAstRef"

End Function
Public Function DayNumber(yljd As Integer, mon%, mday%) As Integer
   
   'determines daynumber for any month = mon%, day = mday%
   'yljd = 365 for regular year, 366 for leap year
   'based on Meeus' formula, p. 65
   
    KK% = 2
    If yljd = 366 Then KK% = 1
    DayNumber = (275 * mon%) \ 9 - KK * ((mon% + 9) \ 12) + mday% - 30
   
   
   End Function
   
Public Function DaysinYear(yrdy As Integer) As Integer

    'function calculates number of day in the civil year, yrdy
    
    Dim yd As Integer
    
    'determine if it is a leap year
    yd = yrdy - 1996
    DaysinYear = 365
    If yd Mod 4 = 0 Then DaysinYear = 366 'its a leap year
    'exclude century years that are not multiple of 400
    If yd Mod 4 = 0 And yrdy Mod 100 = 0 And yrdy Mod 400 <> 0 Then DaysinYear = 365
    
End Function
