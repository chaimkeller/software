Attribute VB_Name = "CalProgram"
'Option Explicit
   'version: 04/08/2003
  
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const WM_CLOSE = &H10
Public Const WM_CANCELMODE = &H1F

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public lngTimerID As Long 'Timer ID that kills internet process if necessary
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const VK_SNAPSHOT = &H2C
Public Const VK_RETURN = &HD
Public Const KEYEVENTF_KEYUP = &H2
Public Const sEmpty = ""
Public Const katzsep% = 107

'requires Bullzip print to pdf printer installed
Public Const SETTINGS_PROGID = "Bullzip.PDFSettings"
Public Const UTIL_PROGID = "Bullzip.PDFUtil"
 
Public currentdir As String, citynames$(800), cityhebnames$(800), numcities%, citnam$
Public errorfnd As Boolean, hebcityname$, s1blk, s2blk, defdriv$, SRTMflag As Integer
Public nearobstnez(3, 601), nearobstski(3, 601), nfind%, nfind1%, rhday%
Public nrnez%(601), nrski%(601), nearnez As Boolean, nearski As Boolean, calnodevis As Boolean
Public initdir As Boolean, tblmesag%, autocancel As Boolean, parshiotEY As Boolean
Public Katz As Boolean, katznum%, katzhebnam$, katztotal%, NearWarning(1) As Boolean
Public astrplaces$(100), astcoord(5, 100), autoNoCDcheck As Boolean, cityAutoEng$, cityAutoHeb$, numAstPlaces%
Public nsetflag%, yrheb%, skiya As Boolean, checklst%, maxang%, parshiotdiaspora As Boolean
Public mypath As String, automatic As Boolean, runningscan As Boolean, ZmanTitle$
Public myname As String, stage%, numautolst%, numautocity%, newpagenum%, autonum%
Public currentdrive As String, startedscan As Boolean, nearauto As Boolean, nearautoedited As Boolean, DSTcheck As Boolean
Public nearyesval As Boolean, netzskiyok As Boolean, arrStrSedra(1, 54) As String
Public ntmp%, nstat%, nstato%, tmpsetflg%, CN4netz$, CN4skiy$, arrStrParshiot(1, 62) As String
Public paperheight, paperwidth, leftmargin, rightmargin, topmargin, bottommargin
Public captmp$, suntop%, nn4%, numchecked%, rescale As Boolean, holidays(1, 11) As String
Public netzski$(1, 1000), First As Boolean, hebleapyear As Boolean, dayRoshHashono%, PaperFormatVis As Boolean
Public nchecked%(999), magnify As Boolean, nearcolor As Boolean, yeartype%, RefHebYear%, RefCivilYear%
Public newhebout As Boolean, hebcal As Boolean, Marginshow As Boolean, SponsorLine$, TitleLine$
Public rescal, portrait As Boolean, prespap%, Loadcombo%, endyr%, RemoveUnderline As Boolean
Public xc(2), y1(2), y2(2), y3(2), y4(2), y5(2), de(2), ys(2), difdyy%, autoprint As Boolean, autosave As Boolean
Public tim$(3, 366), papername$(20), papersize(2, 20) As Integer, numpaper%, margins(4, 20)
Public monthe$(12), monthh$(1, 14), mdates$(2, 13), mmdate%(2, 13), montheh$(1, 12), calnearsearchVis As Boolean
Public dx, dy, xot, yot, xo, yo, dey(2), fillcol, geo As Boolean, eros As Boolean, nweatherflag As Integer
Public stortim$(6, 12, 32), stormon$(12), storheader$(1, 5), geotz!, Option1b As Boolean, Option2b As Boolean
Public astronplace As Boolean, astkmx, astkmy, asthgt, astname$, distlim As Single, distlimnum As Integer
Public drivjk$, drivfordtm$, drivcities$, goahead As Boolean, title$, address$, drivprom$, drivprof$
Public citynodenum%, optionheb As Boolean, progvernum As Single, datavernum As Single, AddObsTime As Integer
Public cushion(5) As Integer, obsdistlim(5) As Integer, obscushion As Integer, outdistlim As Double
Public ecnam$, eroscityflag As Boolean, citnamp$, aveusa As Boolean, errorreport As Boolean
Public erosstatenum As Integer
Public erosstates(50) As String
Public erosstatesindex(50) As Integer
Public eroscitylong(3000) As Double
Public eroscitylat(3000) As Double
Public eroscityhgt(3000) As Double
Public eroscityarea(3000) As String
Public eroscountries(3000) As String
Public eroslongitude As Double, eroslatitude As Double, erosareabat As String
Public foundvantage As Boolean
Public eroscity$, eroshebcity$, IsraelNeighborhood As Boolean, userinput As Boolean, yrcal$, astronfm As Boolean
Public mNode As Node, searchradius, viseros As Boolean
Public optiond%, optiont%, optionz%, options%, zmanopen As Boolean, ProgExec$
Public optionsun1%, optionsun2%, optionround%, numsort%, reorder As Boolean
Public zmantimes(99, 385) As String * 9, zmannumber%(1, 99), zmansetflag%
Public newzmans(99) As String * 9, zmannames$(99), newnum%, neworder As Boolean
Public zmannetz As Boolean, zmanskiy As Boolean, zmantotal% ', yl1%, yl2%
Public vis As Boolean, ast As Boolean, mis As Boolean, vis0%, mis0%, ast0%, resortbutton As Boolean, savehtml As Boolean
Public avekmxnetz, avekmynetz, avehgtnetz, avekmxskiy, avekmyskiy, avehgtskiy
Public internet As Boolean, servnam$, dirint$, zmanyes%, typezman%, zmantype%
Public heb1$(70), heb2$(20), heb3$(32), heb4$(8), heb5$(40), heb6$(10), nettype$, eroscountry$
Public optiondmish%, optiontmish%, dirnet$, RoundSeconds%, myear0%, fshabos0%, htmldir$
Public yrstrt%(1), yrend%(1), visauto As Boolean, mishorauto As Boolean, astauto As Boolean
Public BeginningYear$, EndYear$, NumCivilYears%, NumCivilYearsInc%, BeginCivilRun As Boolean
Public PDFprinter As Boolean, SunriseCalc As Boolean, SunsetCalc As Boolean
'Public MaxHourZemanios As Double
Public Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal hMem As Long)

Public Declare Function lstrcat Lib "kernel32" _
   Alias "lstrcatA" (ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long
   
Public Declare Function SHBrowseForFolder Lib "shell32" _
   (lpBI As BrowseInfo) As Long
   
Public Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long
   
'**********DTM variables**************
Public CHMAP(14, 26) As String * 2, filnumg%
Public CHMNE As String * 2, CHMNEO As String * 2, SF As String * 2

'*********for Google Maps API calls***************
'based on code by: http://www.vb-helper.com/howto_google_map.html
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3

'setting printer default, ref: https://bytes.com/topic/visual-basic/insights/641541-set-default-printer
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
 
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Const HWND_BROADCAST = &HFFFF&
    Public Const WM_WININICHANGE = &H1A
     
 Public Function SetDefaultPrinter(objPrn As Printer) As Boolean
    Dim X As Long, sztemp As String
    sztemp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    X = WriteProfileString("windows", "device", sztemp)
    X = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")
End Function

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




Public Function FNarsin(X As Double) As Double
   On Error GoTo sin50
   FNarsin = Atn(X / Sqr(-X * X + 1))
   Exit Function
sin50: If internet = True Then
          lognum% = FreeFile
          Open drivjk$ + "calprog.log" For Append As #lognum%
          Print #lognum%, "Fatal error: Arcsin argument X*X > 1 . Abort program."
          Close #lognum%
          For i% = 0 To Forms.Count - 1
            Unload Forms(i%)
          Next i%
      
          myfile = Dir(drivfordtm$ + "busy.cal")
          If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
          
          'kill timer
          If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)

          'end program
          End
          End If
End Function
Public Function FNarco(X As Double) As Double
   On Error GoTo cos50
   FNarco = -Atn(X / Sqr(-X * X + 1)) + 2 * Atn(1#)
   Exit Function
cos50: If internet = True Then
          lognum% = FreeFile
          Open drivjk$ + "calprog.log" For Append As #lognum%
          Print #lognum%, "Fatal error: Arccos argument X*X > 1 . Abort program."
          Close #lognum%
          For i% = 0 To Forms.Count - 1
            Unload Forms(i%)
          Next i%
      
          myfile = Dir(drivfordtm$ + "busy.cal")
          If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
          'kill timer
          If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)

          'end program
          End
          End If
End Function




Public Sub readfont()
   If hebcal = True Then
      If hebleapyear = False Then
         ext$ = ".heb"
      ElseIf hebleapyear = True Then
         ext$ = ".hly"
         End If
   ElseIf hebcal = False Then
      ext$ = ".eng"
      End If
   If portrait = True Then
      suffix$ = "Potr"
   ElseIf portrait = False Then
      suffix$ = "Lnsc"
      End If
   Prefix$ = Trim$(Mid$(papername$(prespap%), 1, 4))
   formfilname$ = drivjk$ + sEmpty + Prefix$ + suffix$ + ext$
   filfont% = FreeFile
   myfile = Dir(formfilname$)
   If myfile = sEmpty Then
      MsgBox "CalProgram could not find Font file...will use defaults.", vbInformation, "Cal Program"
      xo = Val(newhebcalfm.Text20.Text): yo = Val(newhebcalfm.Text21.Text)
      xot = Val(newhebcalfm.Text22.Text): yot = Val(newhebcalfm.Text23.Text)
      dx = Val(newhebcalfm.Text24.Text): dy = Val(newhebcalfm.Text25.Text)
      ys(1) = 0: ys(2) = Val(newhebcalfm.Text29.Text)
      xc(1) = Val(newhebcalfm.Text16.Text): xc(2) = Val(newhebcalfm.Text33.Text)
      y1(1) = Val(newhebcalfm.Text17.Text): y1(2) = Val(newhebcalfm.Text34.Text)
      y2(1) = Val(newhebcalfm.Text18.Text): y2(2) = Val(newhebcalfm.Text35.Text)
      y3(1) = Val(newhebcalfm.Text19.Text): y3(2) = Val(newhebcalfm.Text36.Text)
      y4(1) = Val(newhebcalfm.Text26.Text): y4(2) = Val(newhebcalfm.Text37.Text)
      y5(1) = Val(newhebcalfm.Text39.Text): y5(2) = Val(newhebcalfm.Text40.Text)
      de(1) = Val(newhebcalfm.Text30.Text): de(2) = Val(newhebcalfm.Text38.Text)
      dey(1) = Val(newhebcalfm.Text27.Text): dey(2) = Val(newhebcalfm.Text28.Text)
      fillcol = 8454143
      Exit Sub
      End If
   Open formfilname$ For Input As #filfont%
   Line Input #filfont%, doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text3.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text4.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text5.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text6.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text7.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text8.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text9.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text10.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text11.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text12.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text13.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text14.Text = doclin$
   Line Input #filfont%, doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text20.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text21.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text29.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text22.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text23.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text24.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text25.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text27.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text28.Text = doclin$
   Line Input #filfont%, doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text16.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text17.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text18.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text19.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text26.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text39.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text30.Text = doclin$
   Line Input #filfont%, doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text33.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text34.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text35.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text36.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text37.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text40.Text = doclin$
   Input #filfont%, doclin$
   newhebcalfm.Text38.Text = doclin$
   Line Input #filfont%, doclin$
   Input #filfont%, fillcol
   xo = Val(newhebcalfm.Text20.Text): yo = Val(newhebcalfm.Text21.Text)
   xot = Val(newhebcalfm.Text22.Text): yot = Val(newhebcalfm.Text23.Text)
   dx = Val(newhebcalfm.Text24.Text): dy = Val(newhebcalfm.Text25.Text)
   ys(1) = 0: ys(2) = Val(newhebcalfm.Text29.Text)
   xc(1) = Val(newhebcalfm.Text16.Text): xc(2) = Val(newhebcalfm.Text33.Text)
   y1(1) = Val(newhebcalfm.Text17.Text): y1(2) = Val(newhebcalfm.Text34.Text)
   y2(1) = Val(newhebcalfm.Text18.Text): y2(2) = Val(newhebcalfm.Text35.Text)
   y3(1) = Val(newhebcalfm.Text19.Text): y3(2) = Val(newhebcalfm.Text36.Text)
   y4(1) = Val(newhebcalfm.Text26.Text): y4(2) = Val(newhebcalfm.Text37.Text)
   y5(1) = Val(newhebcalfm.Text39.Text): y5(2) = Val(newhebcalfm.Text40.Text)
   de(1) = Val(newhebcalfm.Text30.Text): de(2) = Val(newhebcalfm.Text38.Text)
   dey(1) = Val(newhebcalfm.Text27.Text): dey(2) = Val(newhebcalfm.Text28.Text)
   Close #filfont%
End Sub

Public Sub readpaper()
   If hebcal = True Then
      If hebleapyear = False Then
         ext$ = ".heb"
      ElseIf hebleapyear = True Then
         ext$ = ".hly"
         End If
   ElseIf hebcal = False Then
      ext$ = ".eng"
      End If
   filpaper% = FreeFile
   myfile = Dir(drivjk$ + "Calpaper" + ext$)
   If myfile <> sEmpty Then
      Open drivjk$ + "Calpaper" + ext$ For Input As #filpaper%
      Input #filpaper%, prespap%, paperorien%
      'If paperorien% = 1 Then
      '   portrait = True
      'ElseIf paperorien% = 2 Then
      '   portrait = False
      '   End If
      numpaper% = 0
      Do Until EOF(filpaper%)
         numpaper% = numpaper% + 1
         Input #filpaper%, papername$(numpaper%)
         Input #filpaper%, papersize(1, numpaper%), papersize(2, numpaper%)
         Input #filpaper%, margins(1, numpaper%), margins(2, numpaper%), margins(3, numpaper%), margins(4, numpaper%)
      Loop
      Close #filpaper%
      If automatic = True Then 'determine paper type
         If SunriseCalc And SunsetCalc Then
            'use Standard paper format
            For i% = 1 To numpaper%
               If papername$(i%) = "Standard" Then
                  prespap% = i%
                  Exit For
                  End If
            Next i%
         ElseIf (SunriseCalc And Not SunsetCalc) Or (SunsetCalc And Not SunriseCalc) Then
            'use A4 paper format
            For i% = 1 To numpaper%
               If papername$(i%) = "A4" Then
                  prespap% = i%
                  Exit For
                  End If
            Next i%
            End If
         End If
      If internet = True Then 'always use A4 paper format
         For i% = 1 To numpaper%
            If papername$(i%) = "A4" Then
               prespap% = i%
               Exit For
               End If
         Next i%
         End If
    Else  'use defaults
      MsgBox "CalProgram could not find file containing paper formats...will use defaults.", vbInformation, "Cal Program"
      papername$(1) = "A4"
      papersize(1, 1) = 209: papersize(2, 1) = 296
      papername$(2) = "Standard"
      papersize(1, 2) = 165: papersize(2, 2) = 239
      papername$(3) = "Kovetz"
      papersize(1, 3) = 145: papersize(2, 3) = 230
      papername$(4) = "Variable"
      papersize(1, 4) = 200: papersize(2, 4) = 250
      numpaper% = 4
      prespap% = 1
      For i% = 1 To 4
         For j% = 1 To 4
            margins(j%, i%) = 10
         Next j%
      Next i%
      End If
      paperwidth = papersize(1, prespap%)
      paperheight = papersize(2, prespap%)
      leftmargin = margins(1, prespap%)
      rightmargin = margins(2, prespap%)
      topmargin = margins(3, prespap%)
      bottommargin = margins(4, prespap%)
End Sub
Public Sub savepaper()
       filpaper% = FreeFile
       If hebcal = True Then
          If hebleapyear = False Then
             ext$ = ".heb"
          ElseIf hebleapyear = True Then
             ext$ = ".hly"
             End If
       ElseIf hebcal = False Then
          ext$ = ".eng"
          End If
       Open drivjk$ + "Calpaper" + ext$ For Output As #filpaper%
       If portrait = True Then
          paperorien% = 1
       ElseIf portrait = False Then
          paperorien% = 2
          End If
       Write #filpaper%, prespap%, paperorien%
       For i% = 1 To numpaper%
          Write #filpaper%, papername$(i%)
          Write #filpaper%, papersize(1, i%), papersize(2, i%)
          Write #filpaper%, margins(1, i%), margins(2, i%), margins(3, i%), margins(4, i%)
       Next i%
       Close #filpaper%
End Sub
Public Sub savefont()
    If hebcal = True Then
      If hebleapyear = False Then
         ext$ = ".heb"
      ElseIf hebleapyear = True Then
         ext$ = ".hly"
         End If
    ElseIf hebcal = False Then
       ext$ = ".eng"
       End If
    If portrait = True Then
       suffix$ = "Potr"
    ElseIf portrait = False Then
       suffix$ = "Lnsc"
       End If
    Prefix$ = Trim$(Mid$(papername$(prespap%), 1, 4))
    formfilname$ = drivjk$ + sEmpty + Prefix$ + suffix$ + ext$
    filfont% = FreeFile
    Open formfilname$ For Output As #filfont%
    Print #filfont%, "font parameters"
    Write #filfont%, newhebcalfm.Text3.Text
    Write #filfont%, newhebcalfm.Text4.Text
    Write #filfont%, newhebcalfm.Text5.Text
    Write #filfont%, newhebcalfm.Text6.Text
    Write #filfont%, newhebcalfm.Text7.Text
    Write #filfont%, newhebcalfm.Text8.Text
    Write #filfont%, newhebcalfm.Text9.Text
    Write #filfont%, newhebcalfm.Text10.Text
    Write #filfont%, newhebcalfm.Text11.Text
    Write #filfont%, newhebcalfm.Text12.Text
    Write #filfont%, newhebcalfm.Text13.Text
    Write #filfont%, newhebcalfm.Text14.Text
    Print #filfont%, "portrait general orientation parameters and fonts"
    Write #filfont%, newhebcalfm.Text20.Text
    Write #filfont%, newhebcalfm.Text21.Text
    Write #filfont%, newhebcalfm.Text29.Text
    Write #filfont%, newhebcalfm.Text22.Text
    Write #filfont%, newhebcalfm.Text23.Text
    Write #filfont%, newhebcalfm.Text24.Text
    Write #filfont%, newhebcalfm.Text25.Text
    Write #filfont%, newhebcalfm.Text27.Text
    Write #filfont%, newhebcalfm.Text28.Text
    Print #filfont%, "portrait top calendar parameters"
    Write #filfont%, newhebcalfm.Text16.Text
    Write #filfont%, newhebcalfm.Text17.Text
    Write #filfont%, newhebcalfm.Text18.Text
    Write #filfont%, newhebcalfm.Text19.Text
    Write #filfont%, newhebcalfm.Text26.Text
    Write #filfont%, newhebcalfm.Text39.Text
    Write #filfont%, newhebcalfm.Text30.Text
    Print #filfont%, "portrait bottom calendar parameters"
    Write #filfont%, newhebcalfm.Text33.Text
    Write #filfont%, newhebcalfm.Text34.Text
    Write #filfont%, newhebcalfm.Text35.Text
    Write #filfont%, newhebcalfm.Text36.Text
    Write #filfont%, newhebcalfm.Text37.Text
    Write #filfont%, newhebcalfm.Text40.Text
    Write #filfont%, newhebcalfm.Text38.Text
    Print #filfont%, "fill RGB color"
    Write #filfont%, fillcol
    Close #filfont%
End Sub
Public Sub hebnum(k%, cha$)
'generates hebrew letters: k%=0 returns cha$ = "א", k%=1 returns cha$="ב", etc.
   If k% <= 10 Then
      cha$ = Trim$((Chr$(k% + 223))) + " "
   ElseIf k% > 10 And k% < 20 Then
      cha$ = Trim$((Chr$(233))) + Trim$(Chr$(k% - 10 + 223))
   ElseIf k% = 20 Then
      cha$ = Trim$(Chr$(235)) + " "
   ElseIf k% > 20 And k% < 30 Then
      cha$ = Trim$(Chr$(235)) + Trim$(Chr$(k% - 20 + 223))
   ElseIf k% = 30 Then
      cha$ = Trim$(Chr$(236)) + " "
   ElseIf k% = 31 Then
      cha$ = Trim$(Chr$(236)) + Trim$(Chr$(224))
      End If
   If k% = 15 Then cha$ = Trim$(Chr$(232)) + Trim$(Chr$(229))
   If k% = 16 Then cha$ = Trim$(Chr$(232)) + Trim$(Chr$(230))
   cha$ = Trim$(cha$)
End Sub
Public Sub hebweek(dayweek%, cha$)
   Select Case dayweek%
      Case 1
         cha$ = heb4$(1)
      Case 2
         cha$ = heb4$(2)
      Case 3
         cha$ = heb4$(3)
      Case 4
         cha$ = heb4$(4)
      Case 5
         cha$ = heb4$(5)
      Case 6
         cha$ = heb4$(6)
      Case 0, 7
         cha$ = heb4$(7)
      Case Else
   End Select
End Sub

Sub TimerProc(ByVal hwnd As Long, _
               ByVal uMsg As Long, _
               ByVal idEvent As Long, _
               ByVal dwTime As Long)
               
'this timer procedure kills internet process if it's still around
'after 5 minutes

    'Abort program after updating log file.
     lognum% = FreeFile
     Open drivjk$ + "calprog.log" For Append As #lognum%
     Print #lognum%, "Step #0-0: Exceeded wait time-Abort (see Cal_timr)."
     Close #lognum%
     errlog% = FreeFile
     Open drivjk$ + "Cal_timr.log" For Output As #errlog%
     Print #errlog%, "Cal Prog exceeded alloted time limit of 5 minutes-Aborting!"
     Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
     Close #errlog%
     Close
      
     'unload forms
     For i% = 0 To Forms.Count - 1
        Unload Forms(i%)
     Next i%
      
     myfile = Dir(drivfordtm$ + "busy.cal")
     If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
     'kill the timer
     If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
     
     'end the process abruptly
     End

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadParshiotNames
' DateTime  : 8/14/2003 09:51
' Author    : Chaim Keller
' Purpose   : loads up names of Eretz Yisroel and diaspora torah reading
'---------------------------------------------------------------------------------------
'
Sub LoadParshiotNames()
    'load up names of shabbos torah reading
   On Error GoTo LoadParshiotNames_Error

   If Dir(drivjk$ & "calparshiot_data.txt") <> sEmpty And _
      Dir(drivjk$ & "calparshiot.txt") <> sEmpty Then
      Iparsha$ = sEmpty 'torah reading for Eretz Yisroel
      Dparsha$ = sEmpty 'torah reading for diaspora
      'determine cycle name
      Select Case hebleapyear
         Case False 'Hebrew calendar non-leapyear
            Iparsha$ = "[0_"
         Case True 'Hebrew calendar leapyear
            Iparsha$ = "[1_"
      End Select
      Select Case dayRoshHashono% 'day of the week of RoshHashono
         Case 2 'Monday
            Iparsha$ = Iparsha$ & "2_"
         Case 3 'Tuesday
            Iparsha$ = Iparsha$ & "3_"
         Case 5 'Thursday
            Iparsha$ = Iparsha$ & "5_"
         Case 7 'Shabbos
            Iparsha$ = Iparsha$ & "7_"
      End Select
      Select Case yeartype% 'chaser, kesidrah, shalem
         Case 1 'chaser
            Iparsha$ = Iparsha$ & "1"
         Case 2 'kesidrah
            Iparsha$ = Iparsha$ & "2"
         Case 3 'shalem
            Iparsha$ = Iparsha$ & "3"
      End Select
      Dparsha$ = Iparsha$
      
      Select Case Iparsha$
         Case "[0_2_1", "[0_5_3", "[0_7_1", "[0_7_3", _
              "[1_5_1", "[1_5_3", "[1_7_1" 'alike for Eretz Yisroel and diaspora
            Iparsha$ = Iparsha$ & "]"
            Dparsha$ = Dparsha$ & "]"
         Case Else 'different for Eretz Yisroel and the diaspora
            Iparsha$ = Iparsha$ & "_israel]"
            Dparsha$ = Dparsha$ & "_diaspora]"
      End Select
      
     'read Hebrew and English names of the parshiot
      parnum% = FreeFile
      Open drivjk$ & "calparshiot.txt" For Input As #parnum%
      PMode% = -1
      Do Until EOF(parnum%)
         Line Input #parnum%, doclin$
         If doclin$ = "[hebparshiot]" Then 'Hebrew parshiot names
            PMode% = 0
            nn% = 0
         ElseIf doclin$ = "[engparshiot]" Then 'English parshiot names
            PMode% = 1
            nn% = 0
         ElseIf Trim$(doclin$) <> sEmpty Then
            Select Case PMode%
               Case 0, 1
                  arrStrParshiot(PMode%, nn%) = doclin$
                  nn% = nn% + 1
               Case Else
                  'keep on looping
            End Select
         ElseIf Trim$(doclin$) = sEmpty And PMode% = 1 Then
            Exit Do
            End If
      Loop
      Close #parnum%
      
      'load up the yom tovim
      PMode% = -1
      parnum% = FreeFile
      Open drivjk$ & "calparshiot.txt" For Input As #parnum%
      Do Until EOF(parnum%)
         Line Input #parnum%, doclin$
         If doclin$ = "[hebholidays]" Then 'Hebrew Yom Tovim names
            PMode% = 0
            numhol% = 0
         ElseIf doclin$ = "[engholidays]" Then 'English Yom Tovim names
            PMode% = 1
            numhol% = 0
         ElseIf doclin$ <> sEmpty And PMode% <> -1 Then
            holidays(PMode%, numhol%) = doclin$
            numhol% = numhol% + 1
         ElseIf doclin$ = sEmpty And PMode% = 1 Then
            'finished
            Exit Do
            End If
      Loop
      Close #parnum%
      
      'load up the sedra for Eretz Yisroel and the diaspora
      parnum% = FreeFile
      Open drivjk$ & "calparshiot_data.txt" For Input As #parnum%
      PMode% = -1
      nfound% = 0
      Do Until EOF(parnum%)
         Line Input #parnum%, doclin$
         Select Case PMode%
            Case -1
                If Trim$(doclin$) = Iparsha$ And Trim$(doclin$) = Dparsha$ Then 'parse out both sedraot
                   PMode% = 2
                   nfound% = nfound% + 2
                ElseIf Trim$(doclin$) = Iparsha$ And Trim$(doclin$) <> Dparsha$ Then 'parse out Eretz Yisroel sedra
                   PMode% = 0
                   nfound% = nfound% + 1
                ElseIf Trim$(doclin$) = Dparsha$ And Trim$(doclin$) <> Iparsha$ Then 'parse out diaspora sedra
                   PMode% = 1
                   nfound% = nfound% + 1
                   End If
            Case 0, 1, 2 'parse out EY, diaspora sedra
               Call ParseSedra(doclin$, PMode%)
               If nfound% = 1 Then PMode% = -1 'look for sedra of Eretz Yisroel
               If nfound% = 2 Then Exit Do 'found everything
            Case Else 'keep on looping
         End Select
      Loop
      Close #parnum%
      If PMode% = -1 Then 'couldn't find parshiot
         If internet = False Then
            MsgBox "Warning--couldn't find the correct sedra!", _
                   vbExclamation + vbOKOnly, "Cal Program"
            End If
         End If
   Else
      'can't load up parshiot since data missing
      If internet = True Then
        lognum% = FreeFile
        Open drivjk$ + "calprog.log" For Append As #lognum%
        Print #lognum%, "Can't load up sedra names--file(s) missing."
        Close #lognum%
        End If
     End If

   On Error GoTo 0
   Exit Sub

LoadParshiotNames_Error:
    
    If internet = False Then
       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadParshiotNames of Module CalProgram", vbCritical + vbOKOnly, "Cal Program"
    Else
      errlog% = FreeFile
      Open drivjk$ + "Cal_OKbh.log" For Output As errlog%
      Print #errlog%, "Cal Prog exited from SunriseSunset: Error in reading parshiot"
      Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
      Close #errlog%
      Close
      
      'unload forms
      For i% = 0 To Forms.Count - 1
        Unload Forms(i%)
      Next i%
      
      myfile = Dir(drivfordtm$ + "busy.cal")
      If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      'kill the timer
      If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
      'bring program to abrupt end
      End
    
      End If
      
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParseSedra
' DateTime  : 8/14/2003 10:49
' Author    : Chaim Keller
' Purpose   : Parses the sedra numbers in the file calparshiot_data.txt
'---------------------------------------------------------------------------------------
'
Sub ParseSedra(doclin$, PMode%)

   On Error GoTo ParseSedra_Error
   
   Dim doubleparsha As Boolean
   If optionheb Then
      hc% = 0
   Else
      hc% = 1
      End If
   mParshaI% = 0
   mParshaD% = 0
   parsnum$ = sEmpty
   For i% = 1 To Len(doclin$)
      cha$ = Mid$(doclin$, i%, 1)
      Select Case cha$
         Case ",", "}"
           If PMode% = 0 Or PMode% = 1 Then
              If PMode% = 0 Then mParsha% = mParshaI%
              If PMode% = 1 Then mParsha% = mParshaD%
              
              If Not doubleparsha Then 'single parsha
                 arrStrSedra(PMode%, mParsha%) = arrStrParshiot(hc%, Val(parsnum$))
              Else 'double parsha
                 arrStrSedra(PMode%, mParsha%) = arrStrParshiot(hc%, Val(parsnum$)) & _
                                                 "-" & arrStrParshiot(hc%, Val(parsnum$) + 1)
                 doubleparsha = False 'reset double parsha flag
                 End If
              
              If PMode% = 0 Then mParshaI% = mParshaI% + 1
              If PMode% = 1 Then mParshaD% = mParshaD% + 1
              parsnum$ = sEmpty

           ElseIf PMode% = 2 Then 'both EY and diaspora have same parsha
              
              If Not doubleparsha Then 'single parsha
                 arrStrSedra(0, mParshaI%) = arrStrParshiot(hc%, Val(parsnum$))
                 arrStrSedra(1, mParshaD%) = arrStrSedra(0, mParshaI%)
                 mParshaI% = mParshaI% + 1 'increment the parsha number
                 mParshaD% = mParshaD% + 1 'increment the parsha number
              Else 'double parsha
                 arrStrSedra(0, mParshaI%) = arrStrParshiot(hc%, Val(parsnum$)) & _
                                                 "-" & arrStrParshiot(hc%, Val(parsnum$) + 1)
                 arrStrSedra(1, mParshaD%) = arrStrSedra(0, mParshaI%)
                 
                 doubleparsha = False 'reset double parsha flag
                 
                 mParshaI% = mParshaI% + 1 'increment the parsha number
                 mParshaD% = mParshaD% + 1 'increment the parsha number
                 
                 End If
              End If
           
           parsnum$ = sEmpty 'reset the sedra designation number
         
         Case "D"
           parsnum$ = sEmpty
           doubleparsha = True
         Case " "
         Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            parsnum$ = parsnum$ & cha$
         Case Else
      End Select
   Next i%

   On Error GoTo 0
   Exit Sub

ParseSedra_Error:

    If internet = False Then
       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadParshiotNames of Module CalProgram", vbCritical + vbOKOnly, "Cal Program"
    Else
      errlog% = FreeFile
      Open drivjk$ + "Cal_OKbh.log" For Output As errlog%
      Print #errlog%, "Cal Prog exited from SunriseSunset: Error in reading parshiot"
      Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
      Close #errlog%
      Close
      
      'unload forms
      For i% = 0 To Forms.Count - 1
        Unload Forms(i%)
      Next i%
      
      myfile = Dir(drivfordtm$ + "busy.cal")
      If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      'kill the timer
      If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
      'bring program to abrupt end
      End
      End If
      
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InsertHolidays
' DateTime  : 8/14/2003 21:34
' Author    : Chaim Keller
' Purpose   : Insert holiday info into the day field
'---------------------------------------------------------------------------------------
'
Sub InsertHolidays(calday$, i%, k%)

   'passed variables are:
   '(1) the day field = calday$
   '(2) the increment variables i%,k% used to determine
   '    (a) the date field = CalDate$
   '    (b) the Shabbos Parsha field = CalParsha$

   On Error GoTo InsertHolidays_Error
   
   '================================================
   'add parshiot
   
    calday$ = Trim$(stortim$(4, i% - 1, k% - 1))
    If optionheb = False Then
       Select Case calday$
          Case heb4$(1)
             calday$ = "Sunday"
          Case heb4$(2)
             calday$ = "Monday"
          Case heb4$(3)
             calday$ = "Tuesday"
          Case heb4$(4)
             calday$ = "Wednesday"
          Case heb4$(5)
             calday$ = "Thursday"
          Case heb4$(6)
             calday$ = "Friday"
          Case heb4$(7)
             calday$ = "Shabbos"
       End Select
       End If
    If parshiotEY Then
       If Trim$(stortim$(5, i% - 1, k% - 1)) <> "-----" Then
          If optionheb Then
             If InStr(stortim$(5, i% - 1, k% - 1), heb4$(6)) Then
                calday$ = Trim$(stortim$(5, i% - 1, k% - 1))
             Else
                calday$ = heb5$(37) & Trim$(stortim$(5, i% - 1, k% - 1)) '<<<<<<<<<<<
                End If
          Else 'calday$ = "Shabbos Parshos " & Trim$(stortim$(5, i% - 1, k% - 1))
             If InStr(stortim$(5, i% - 1, k% - 1), "Shabbos") Then
                calday$ = Trim$(stortim$(5, i% - 1, k% - 1))
             Else
                calday$ = "Shabbos_" & Trim$(stortim$(5, i% - 1, k% - 1))
                End If
             End If
          End If
    ElseIf parshiotdiaspora Then
       If Trim$(stortim$(6, i% - 1, k% - 1)) <> "-----" Then
          If optionheb Then
             If InStr(stortim$(6, i% - 1, k% - 1), heb4$(6)) Then
                calday$ = Trim$(stortim$(6, i% - 1, k% - 1))
             Else
                calday$ = heb5$(37) & Trim$(stortim$(6, i% - 1, k% - 1))
                End If
          Else 'calday$ = "Shabbos Parshos " & Trim$(stortim$(5, i% - 1, k% - 1))
             If InStr(stortim$(6, i% - 1, k% - 1), "Shabbos") Then
                calday$ = Trim$(stortim$(6, i% - 1, k% - 1))
             Else
                calday$ = "Shabbos_" & Trim$(stortim$(6, i% - 1, k% - 1))
                End If
             End If
          End If
       End If
 
  '=========================================================
  'now add holidays to Shabbosim
   
   caldate$ = Trim$(stortim$(3, i% - 1, k% - 1))
   If parshiotEY Then
      CalParsha$ = stortim$(5, i% - 1, k% - 1)
   ElseIf parshiotdiaspora Then
      CalParsha$ = stortim$(6, i% - 1, k% - 1)
      End If
   
   If CalParsha$ <> "-----" Then
      'add in Shabbos Chanukah and Shabbos Shushan Purim when appropriate
      If optionheb Then
         Select Case Trim$(caldate$)
            Case heb5$(13), heb5$(14), heb5$(15), heb5$(16), _
                 heb5$(17), heb5$(18), heb5$(19), heb5$(20), heb5$(21) 'Chanukah
                 If (yeartype% = 2 Or yeartype% = 3) And caldate$ = heb5$(21) Then GoTo remU
                 calday$ = calday$ & heb5$(36) & holidays(0, 6)
                 GoTo remU
            Case heb5$(24), heb5$(25)
               calday$ = calday$ & heb5$(36) & holidays(0, 8)
               GoTo remU
            Case Else
               GoTo remU
         End Select
      Else
         Select Case Trim$(caldate$)
            Case "25-Kislev", "26-Kislev", "27-Kislev", "28-Kislev", _
                 "29-Kislev", "30-Kislev", "1-Teves", "2-Teves", "3-Teves"
                  If (yeartype% = 2 Or yeartype% = 3) And caldate$ = "3-Teves" Then GoTo remU
                 calday$ = holidays(1, 6) & ",_" & calday$
                 GoTo remU
            Case "15-Adar", "15-Adar II" 'Shushan Purim
               calday$ = holidays(1, 8) & ",_" & calday$
               GoTo remU
            Case Else
               GoTo remU
          End Select
          End If
       End If
   
  '=======================================================
   'add in week day holidays
   If optionheb Then
      Select Case Trim$(caldate$)
         Case heb5$(1), heb5$(2)
            calday$ = calday$ & heb5$(36) & holidays(0, 0)
         Case heb5$(3)
            calday$ = calday$ & heb5$(36) & holidays(0, 1)
         Case heb5$(4)
            calday$ = calday$ & heb5$(36) & holidays(0, 2)
         Case heb5$(5)
            If parshiotEY Then
               calday$ = calday$ & heb5$(36) & holidays(0, 3)
            ElseIf parshiotdiaspora Then
               calday$ = calday$ & heb5$(36) & holidays(0, 2)
               End If
         Case heb5$(6), heb5$(7), heb5$(8), heb5$(9), heb5$(10)
            calday$ = calday$ & heb5$(36) & holidays(0, 3)
         Case heb5$(11)
            calday$ = calday$ & heb5$(36) & holidays(0, 4)
         Case heb5$(12)
            If parshiotdiaspora Then calday$ = calday$ & heb5$(36) & holidays(0, 5)
         Case heb5$(13), heb5$(14), heb5$(15), heb5$(16), _
              heb5$(17), heb5$(18), heb5$(19), heb5$(20), heb5$(21) 'Chanukah
              If (yeartype% = 2 Or yeartype% = 3) And caldate$ = heb5$(21) Then GoTo remU
              calday$ = calday$ & heb5$(36) & holidays(0, 6)
         Case heb5$(22), heb5$(23)
            calday$ = calday$ & heb5$(36) & holidays(0, 7)
         Case heb5$(24), heb5$(25)
            calday$ = calday$ & heb5$(36) & holidays(0, 8)
         Case heb5$(26)
            calday$ = calday$ & heb5$(36) & holidays(0, 9)
         Case heb5$(27)
            If parshiotEY Then
                calday$ = calday$ & heb5$(36) & holidays(0, 10)
            ElseIf parshiotdiaspora Then
                calday$ = calday$ & heb5$(36) & holidays(0, 9)
               End If
         Case heb5$(28), heb5$(29), heb5$(30), heb5$(31)
            calday$ = calday$ & heb5$(36) & holidays(0, 10)
         Case heb5$(32)
            calday$ = calday$ & heb5$(36) & holidays(0, 9)
         Case heb5$(33)
            If parshiotdiaspora Then calday$ = calday$ & heb5$(36) & holidays(0, 9)
         Case heb5$(34)
            calday$ = calday$ & heb5$(36) & holidays(0, 11)
         Case heb5$(35)
            If parshiotdiaspora Then calday$ = calday$ & heb5$(36) & holidays(0, 11)
         Case Else
      End Select
   Else
      Select Case Trim$(caldate$)
        Case "1-Tishrey", "2-Tishrey" 'Rosh Hoshono
           calday$ = holidays(1, 0) & ",_" & calday$
        Case "10-Tishrey" 'Yom Hakipurim
           calday$ = holidays(1, 1) & ",_" & calday$
        Case "15-Tishrey" 'Succos
           calday$ = holidays(1, 2) & ",_" & calday$
        Case "16-Tishrey" 'Second day Succos for diaspora
           If parshiotEY Then 'Chol Hamoed Succos
              calday$ = holidays(1, 3) & ",_" & calday$
           ElseIf parshiotdiaspora Then 'Succos
              calday$ = holidays(1, 2) & ",_" & calday$
              End If
        Case "17-Tishrey", "18-Tishrey", "19-Tishrey", _
              "20-Tishrey", "21-Tishrey" 'Chol Hamoed Succos
           calday$ = holidays(1, 3) & ",_" & calday$
        Case "22-Tishrey" 'Shmini Azeres
           calday$ = holidays(1, 4) & ",_" & calday$
        Case "23-Tishrey" 'Simchas Torah
           If parshiotdiaspora Then calday$ = holidays(1, 5) & ",_" & calday$
        Case "25-Kislev", "26-Kislev", "27-Kislev", "28-Kislev", _
             "29-Kislev", "30-Kislev", "1-Teves", "2-Teves", "3-Teves"
              If (yeartype% = 2 Or yeartype% = 3) And caldate$ = "3-Teves" Then GoTo remU
             calday$ = holidays(1, 6) & ",_" & calday$
        Case "14-Adar", "14-Adar II" 'Purim
           calday$ = holidays(1, 7) & ",_" & calday$
        Case "15-Adar", "15-Adar II" 'Shushan Purim
           calday$ = holidays(1, 8) & ",_" & calday$
        Case "15-Nisan" 'Pesach
           calday$ = holidays(1, 9) & ",_" & calday$
        Case "16-Nisan" 'Second day Pesach for diaspora
           If parshiotEY Then
              calday$ = holidays(1, 10) & ",_" & calday$
           ElseIf parshiotdiaspora Then
              calday$ = holidays(1, 9) & ",_" & calday$
              End If
        Case "17-Nisan", "18-Nisan", "19-Nisan", "20-Nisan"
           calday$ = holidays(1, 10) & ",_" & calday$
        Case "21-Nisan" 'last day of Pesach
           calday$ = holidays(1, 9) & ",_" & calday$
        Case "22-Nisan" 'Second day last day of Pesach for diaspora
           If parshiotdiaspora Then calday$ = holidays(1, 9) & ",_" & calday$
        Case "6-Sivan" 'Shavuos
           calday$ = holidays(1, 11) & ",_" & calday$
        Case "7-Sivan" 'Second day of Shavuos for diaspora
           If parshiotdiaspora Then calday$ = holidays(1, 11) & ",_" & calday$
        Case Else 'don't change anything
     End Select
     End If
     
   '========================================================
   'remove "_" from calday$
remU:
   If RemoveUnderline Then
      For ir% = 1 To Len(calday$)
         If Mid$(calday$, ir%, 1) = "_" Then Mid$(calday$, ir%, 1) = " "
      Next ir%
      End If
   

   On Error GoTo 0
   Exit Sub

InsertHolidays_Error:
   If internet = False Then
       MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InsertHolidays of Module CalProgram", vbCritical + vbOKOnly, "Cal Program"
   Else
      errlog% = FreeFile
      Open drivjk$ + "Cal_Zemanim.log" For Output As errlog%
      Print #errlog%, "Cal Prog exited from InsertHolidays: Error in reading parshiot"
      Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
      Close #errlog%
      Close
      
      'unload forms
      For i% = 0 To Forms.Count - 1
        Unload Forms(i%)
      Next i%
      
      myfile = Dir(drivfordtm$ + "busy.cal")
      If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      'kill the timer
      If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
      'bring program to abrupt end
      End
      End If
   
End Sub
Sub hebyear(yr%, yrcal$) 'convert hebrew year in numbers to hebrew characters
   hund% = Fix((yr% - 5000) * 0.01)
   tens% = Fix((yr% - 5000 - hund% * 100) * 0.1)
   ones% = Fix(yr% - 5000 - hund% * 100 - tens% * 10)
   If hund% = 0 Then
      yrc$ = sEmpty
   ElseIf hund% = 1 Then
      yrc$ = Chr$(247) ' "ק"
   ElseIf hund% = 2 Then
      yrc$ = Chr$(248) '"ר"
   ElseIf hund% = 3 Then
      yrc$ = Chr$(249) '"ש"
   ElseIf hund% = 4 Then
      yrc$ = Chr$(250) '"ת"
   ElseIf hund% = 5 Then
      yrc$ = Chr$(250) & Chr$(247) '"תק"
   ElseIf hund% = 6 Then
      yrc$ = Chr$(250) & Chr$(248) '"תר"
   ElseIf hund% = 7 Then
      yrc$ = Chr$(250) & Chr$(249) '"תש"
   ElseIf hund% = 8 Then
      yrc$ = Chr$(250) & Chr$(250) '"תת"
   ElseIf hund% = 9 Then
      yrc$ = Chr$(250) & Chr$(250) & Chr$(247) '"תתק"
   ElseIf hund% = 10 Then
      yrc$ = Chr$(250) & Chr$(250) & Chr$(34) & Chr$(248) ' "תת" + Chr$(34) + "ר"
      End If
   If tens% = 1 Then
      yrt$ = Chr$(233) '"י"
   ElseIf tens% = 2 Then
      yrt$ = Chr$(235) '"כ"
   ElseIf tens% = 3 Then
      yrt$ = Chr$(236) '"ל"
   ElseIf tens% = 4 Then
      yrt$ = Chr$(238) '"מ"
   ElseIf tens% = 5 Then
      yrt$ = Chr$(240) '"נ"
   ElseIf tens% = 6 Then
      yrt$ = Chr$(241) '"ס"
   ElseIf tens% = 7 Then
      yrt$ = Chr$(242) '"ע"
   ElseIf tens% = 8 Then
      yrt$ = Chr$(244) '"פ"
   ElseIf tens% = 9 Then
      yrt$ = Chr$(246) '"צ"
      End If
   If ones% <> 0 Then
      yron$ = Chr$(ones% + 223)
      yrcal$ = yrc$ + yrt$ + Chr$(34) + yron$
   Else
      yrcal$ = yrc$ + Chr$(34) + yrt$
      End If
End Sub
Sub WriteTables(filroot$, root$, ext$)
'this routine writes the csv, htm (zip), and xml files for both
'the console or internet modes
'filroot$ = output file name with directory
'root$ = output file name without directory
'ext$ = the file extension either: "htm", "zip", "csv", or "xml"

        Dim closerow As Boolean
        
        On Error GoTo generrhand
        
        '-------------------------------
        'if error report to be made, branch here
        If errorreport Then
           If ext$ <> "xml" Then 'just copy html file
              sourcehtm = filroot$ & ".html"
              destinhtm = filroot$ & "." & ext$
              FileCopy sourcehtm, destinhtm
           Else 'write error report in xml format
              destinhtm = filroot$ & "." & ext$
              tmpfilxml% = FreeFile
              Open destinhtm For Output As #tmpfilxml%
              Print #tmpfilxml%, "<?xml version='1.0' encoding = 'windows-1255'?>"
              Print #tmpfilxml%, "<html>"
              Print #tmpfilxml%, "<head>"
              Print #tmpfilxml%, "<title>Chai Tables Error Report</title>"
              Print #tmpfilxml%, "</head>"
              Print #tmpfilxml%, "<body>"
              Print #tmpfilxml%, "<h2>The Chai Tables</h2>"
              Print #tmpfilxml%, "<p></p>"
              Print #tmpfilxml%, "No calculated vantage point could be found."
              Print #tmpfilxml%, "Please check your inputs."
              Print #tmpfilxml%, "If you wish, you can attempt to increase the search radius, or calculate astronomical times"
              Print #tmpfilxml%, "<p></p>"
              Print #tmpfilxml%, "</body>"
              Print #tmpfilxml%, "</html>"
              Close #tmpfilxml%
              End If
           Exit Sub
           End If
        '--------------------------------
        
        'write html,xml version of sorted zmanim table
        tmpnum% = FreeFile

        If ext$ = "zip" Then
           Open filroot$ + "." + "htm" For Output As #tmpnum%
        Else
           Open filroot$ + "." + ext$ For Output As #tmpnum%
           End If
        If ext$ = "xml" Then 'write xsl file
           xslnum% = FreeFile 'also write xsl file
           Open filroot$ + ".xsl" For Output As #xslnum%
           End If
        
        If Not internet Then Screen.MousePointer = vbHourglass
        
        'write html header
        If ext$ = "zip" Or ext$ = "htm" Then
           Print #tmpnum%, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
           If Not optionheb Then
              Print #tmpnum%, "<HTML>"
           Else
              Print #tmpnum%, "<HTML dir = " & Chr$(34) & "rtl" & Chr$(34) & ">"
              End If
        ElseIf ext$ <> "csv" Then
           Print #tmpnum%, "<?xml version='1.0' encoding = 'windows-1255'?>"  'Window(Hebrews) uses code page 1255
           If internet Then
              Print #tmpnum%, "<?xml-stylesheet type='text/xsl' href='" & "../data/" & Mid(root$, 1, 8) & ".xsl'?>"
           Else
              Print #tmpnum%, "<?xml-stylesheet type='text/xsl' href='" & root$ & ".xsl'?>"
              End If
           Print #tmpnum%, "<chai xml:space='preserve'>"
           Print #xslnum%, "<xsl:stylesheet version = '1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>"
           Print #xslnum%, "   <xsl:template match='/'>"
           If Not optionheb Then
              Print #xslnum%, "       <HTML>"
           Else
              Print #xslnum%, "       <HTML dir=" & Chr$(34) & "rtl" & Chr$(34) & ">"
              End If
           End If
        
        If ext$ <> "xml" Then
           tmpn% = tmpnum%
        Else
           tmpn% = xslnum%
           End If
           
        If ext$ <> "csv" Then
           Print #tmpn%, "<HEAD>"
           Print #tmpn%, "    <TITLE>Your Chai Z'manim Table</TITLE>"
           Print #tmpn%, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1255""/>"
           Print #tmpn%, "    <META NAME=""AUTHOR"" CONTENT=""Chaim Keller, Chai Tables, www.chaitables.com""/>"
           Print #tmpn%, "</HEAD>"
           Print #tmpn%, "<BODY>"
           End If
        
        If optionheb Then 'convert hebrew year in numbers to hebrew characters
           Call hebyear(yrheb%, yearheb$)
           End If
           
        If ext$ = "xml" Then
             If optionheb = False Then
                If eroscity$ = sEmpty Then
                   'Print #tmpnum%, "<hd2><hd22>Chai Tables for " & citnamp$ & " for the year " & Str(yrheb%) & "</hd22></hd2>"
                   Print #tmpnum%, "<hd2><hd22>" & TitleLine$ & " Tables for " & citnamp$ & " for the year " & Str(yrheb%) & "</hd22></hd2>"
                Else
                   'Print #tmpnum%, "<hd2><hd22>Chai Tables for " & eroscity$ & " (" & citnamp$ & ")" & " for the year " & Str(yrheb%) & "</hd22></hd2>"
                   Print #tmpnum%, "<hd2><hd22>" & TitleLine$ & " Tables for " & eroscity$ & " (" & citnamp$ & ")" & " for the year " & Str(yrheb%) & "</hd22></hd2>"
                   End If
             Else
                If eroscity$ = sEmpty Then
                   Print #tmpnum%, "<hd2><hd22>" & heb3$(1) & hebcityname$ & heb3$(2) & yearheb$ & "</hd22></hd2>"
                Else
                   Print #tmpnum%, "<hd2><hd22>" & " (" & eroscity$ & ") " & heb3$(12) & " " & hebcityname$ & heb3$(2) & yearheb$ & "</hd22></hd2>"
                   End If
                End If
         Else
            If optionheb = False Then
               If eroscity$ = sEmpty Then
                  'Print #tmpnum%, "<h2><center>Chai Tables for " & citnamp$ & " for the year " & Str(yrheb%) & "</center></h2>"
                  If ext$ = "zip" Or ext$ = "htm" Then
                     Print #tmpnum%, "<h2><center>" & TitleLine$ & " Tables for " & citnamp$ & " for the year " & Str(yrheb%) & "</center></h2>"
                  ElseIf ext$ = "csv" Then
                     Write #tmpnum%, TitleLine$ & " Tables for " & citnamp$ & " for the year " & Str(yrheb%)
                     End If
               Else
                  'Print #tmpnum%, "<h2><center>Chai Tables for " & eroscity$ & " (" & citnamp$ & ")" & " for the year " & Str(yrheb%) & "</center></h2>"
                  If ext$ = "zip" Or ext$ = "htm" Then
                     Print #tmpnum%, "<h2><center>" & TitleLine$ & " Tables for " & eroscity$ & " (" & citnamp$ & ")" & " for the year " & Str(yrheb%) & "</center></h2>"
                  ElseIf ext$ = "csv" Then
                     Write #tmpnum%, TitleLine$ & " Tables for " & eroscity$ & " (" & citnamp$ & ")" & " for the year " & Str(yrheb%)
                     End If
                  End If
            Else
               If eroscity$ = sEmpty Then
                  If ext$ = "zip" Or ext$ = "htm" Then
                     Print #tmpnum%, "<h2><center>" & heb3$(1) & hebcityname$ & heb3$(2) & yearheb$ & "</h2>"
                  ElseIf ext$ = "csv" Then
                     Write #tmpnum%, heb3$(1) & hebcityname$ & heb3$(2) & yearheb$
                     End If
               Else
                  If ext$ = "zip" Or ext$ = "htm" Then
                     Print #tmpnum%, "<h2><center>" & " (" & eroscity$ & ") " & heb3$(12) & " " & hebcityname$ & heb3$(2) & yearheb$ & "</center></h2>"
                  ElseIf ext$ = "csv" Then
                     Write #tmpnum%, "(" & eroscity$ & ") " & heb3$(12) & " " & hebcityname$ & heb3$(2) & yearheb$
                     End If
                  End If
               End If
            End If
            
      address$ = heb2$(14) & Str(datavernum) & "." & Str(progvernum) & " ©"
      If optionheb = False Then
         address$ = "© Luchos Chai, Fahtal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: " & Str(progvernum) & "." & Str(datavernum)
         End If
    
        
        'now write name of table
         
         If Not internet Then 'ask for description of the table
            Screen.MousePointer = vbDefault
            response = InputBox("Input a short description of the z'manim", "Z'manim header", ZmanTitle$)
            hdr$ = response
            Screen.MousePointer = vbHourglass
            GoTo wt100
            End If

      If ZmanTitle$ <> sEmpty Then
         hdr$ = ZmanTitle$
      Else
        
        If typezman% = 0 Then
           If optionheb = True Then
              'hdr$ = "זמנים אלו מבוססים על שיטת המגן אברהם כשה העלות השחר והצה""כ הם 16.1 מעלות (72 דק')"
              hdr$ = heb3$(3)
           Else
              hdr$ = "Z'manim calculated acc. to the opinion of the Mogen Avrohom using a 16.1 degrees dawn and twilight (approx. 72 min.) "
              End If
       ElseIf typezman% = 1 Then
           If optionheb = True Then
              'hdr$ = "זמנים אלו מבוססים על שיטת המגן אברהם כשה העלות השחר והצה""כ הם 16.1 מעלות (90 דק')"
              hdr$ = heb3$(4)
           Else
              hdr$ = "Z'manim calculated acc. to the opinion of the Mogen Avrohom using a 19.75 degrees dawn and twilight (approx. 90 min.) "
              End If
       ElseIf typezman% = 2 Then
           If optionheb = True Then
              'hdr$ = "זמנים אלול מבוססים על שיטת הגר""א ועל הזריחה המישורית ועל השקיעה המישורית"
              'hdr$ = "זמנים אלול מבוססים על שיטת הגר" + Chr$(34) + "א ועל הזריחה המישורית ועל השקיעה המישורית"
              hdr$ = heb3$(5) + Chr$(34) + heb3$(6)
           Else
              hdr$ = "Z'manim calculated acc. to the opinion of the Groh using the mishor sunrise and sunset"
              End If
       ElseIf typezman% = 3 Then
           If optionheb = True Then
              'hdr$ = "זמנים אלול מבוססים על שיטת הגר""א ועל הזריחה האסטרונומית ועל השקיעה האסטרונומית"
              'hdr$ = "זמנים אלול מבוססים על שיטת הגר" + Chr$(34) + "א ועל הזריחה האסטרונומית ועל השקיעה האסטרונומית"
              hdr$ = heb3$(5) + Chr$(34) + heb3$(7)
           Else
              hdr$ = "Z'manim calculated acc. to the opinion of the Groh using the astronomical sunrise and sunset"
              End If
       ElseIf typezman% = 4 Then
           If optionheb = True Then
              'hdr$ = "זמנים אלול מבוססים על שיטת הבן איש חי"
              hdr$ = heb3$(8)
           Else
              hdr$ = "Z'manim calculated acc. to opinion of the Ben Ish Chai"
              End If
       ElseIf typezman% = 5 Then
           If optionheb = True Then
              'hdr$ = "זמנים אלול מבוססים על שיטת בעעל התניא"
              hdr$ = heb3$(14)
           Else
              hdr$ = "Z'manim calculated acc. to opinion of the Baal Hatanya"
              End If
       ElseIf typezman% = 6 Then
           If optionheb = True Then
              hdr$ = heb3$(15)
           Else
              hdr$ = "Z'manim calculated acc. to opinion of Harav Zalman Baruch Melamid"
              End If
       ElseIf typezman% = 7 Then 'Yedidia's times
           If optionheb = True Then
              'hdr$ = "זמני שקיעה האסטרונומית"
              hdr$ = heb3$(13)
           Else
              hdr$ = "Astronomical Sunset"
              End If
           End If
        End If
        
        hdr$ = Trim$(hdr$)
           
wt100: If ext$ = "xml" Then
          Print #xslnum%, "<xsl:for-each select='chai/hd2'>"
          Print #xslnum%, "   <h2><center><xsl:value-of select='hd22'/></center></h2>"
          Print #xslnum%, "</xsl:for-each>"
          Print #xslnum%, "<xsl:for-each select='chai/hd3'>"
          Print #xslnum%, "   <h3><center><xsl:value-of select='hd33'/></center></h3>"
          Print #xslnum%, "</xsl:for-each>"
          Print #xslnum%, "<xsl:for-each select='chai/address'>"
          Print #xslnum%, "   <h5><center><xsl:value-of select='name'/></center></h5>"
          Print #xslnum%, "</xsl:for-each>"
          If SponsorLine$ <> sEmpty Then
             Print #xslnum%, "<xsl:for-each select='chai/sponsor'>"
             Print #xslnum%, "   <h5><center><xsl:value-of select='sponsorlogo'/></center></h5>"
             Print #xslnum%, "</xsl:for-each>"
             End If
          Print #tmpn%, ""
          Print #tmpnum%, "<hd3><hd33>" & hdr$ & "</hd33></hd3>"
          Print #tmpnum%, "<address><name>" & address$ & "</name></address>"
          If SponsorLine$ <> sEmpty Then
             Print #tmpnum%, "<sponsor><sponsorlogo>" & SponsorLine$ & "</sponsorlogo></sponsor>"
             End If
       Else
          If ext$ = "zip" Or ext$ = "htm" Then
             Print #tmpnum%, "<h3><center>" & hdr$ & "</center></h3>"
             Print #tmpnum%, "<h3><center>" & address$ & "</center></h3>"
             If SponsorLine$ <> sEmpty Then
                Print #tmpnum%, "<h5><center>" & SponsorLine$ & "</center></h5>"
                End If
          ElseIf ext$ = "csv" Then
             Write #tmpnum%, hdr$
             Write #tmpnum%, address$
             If SponsorLine$ <> sEmpty Then
                Write #tmpnum%, SponsorLine$
                End If
             End If
          End If
           
        'generate table column legend
        If reorder = True Then
           totnum% = numsort%
        Else
           totnum% = newnum%
           End If
        
        nn% = -1
        For m% = totnum% To 0 Step -1
           
           If ext$ <> "csv" Then
              'attach letters or numbers to labels
              If optionheb Then
                 Call hebnum(m% + 4, cha$)
                 outdoc$ = cha$ & " = " & zmannames$(m%) + " | " + outdoc$
              Else
                 nn% = nn% + 1
                 cha$ = Trim$(Str$(nn% + 4))
                 outdoc$ = outdoc$ + " | " + cha$ & " = " & zmannames$(nn%)
                 End If
           
           ElseIf ext$ = "csv" Then
           
              'remove "_" from column header
              For nn% = 1 To Len(zmannames$(m%))
                  If Mid$(zmannames$(m%), nn%, 1) = "_" Then Mid$(zmannames$(m%), nn%, 1) = " "
              Next nn%
              
              outdoc$ = Chr$(34) & zmannames$(m%) & Chr$(34) + "," + outdoc$
              
              End If
        
        Next m%
        
        'now add the the dates, and days legends
        If ext$ <> "csv" Then
           If optionheb = True Then
             Call hebnum(1, ch1$)
             Call hebnum(2, ch2$)
             Call hebnum(3, ch3$)
             outdoc$ = " | " + ch1$ & " = " & heb3$(11) + _
                      " | " + ch2$ & " = " & heb3$(10) + _
                      " | " + ch3$ & " = " & heb3$(9) + " | " + outdoc$
           ElseIf optionheb = False Then
             outdoc$ = "| 1 = hebrew date | 2 = day | 3 = civil date " + outdoc$
             End If
        
           Print #tmpn%, "<hr/>" 'spacers between the headers and column headers
           Print #tmpn%, "<br/>"
        
        Else
           If optionheb = True Then
             outdoc$ = Chr$(34) & heb3$(11) & Chr$(34) & "," & Chr$(34) & heb3$(10) & Chr$(34) & "," & Chr$(34) & heb3$(9) & Chr$(34) & "," + outdoc$
           ElseIf optionheb = False Then
             outdoc$ = Chr$(34) & "hebrew date" & Chr$(34) & "," & Chr$(34) & "day" & Chr$(34) & "," & Chr$(34) & "civil date" & Chr$(34) & "," + outdoc$
             End If
             
           Print #tmpnum%, sEmpty 'spacers between the headers and column headers
           End If
        
        If ext$ = "xml" Then
           Print #tmpnum%, "<hd4><hd44>" & outdoc$ & "</hd44></hd4>" 'this caption line of zemanim"
           Print #xslnum%, "<xsl:for-each select='chai/hd4'>"
           Print #xslnum%, "   <xsl:value-of select='hd44'/>"
           Print #xslnum%, "</xsl:for-each>"
        Else
           Print #tmpnum%, outdoc$ 'print the legend for the csv, htm, and zip files
           End If
           
        If ext$ <> "csv" Then
           Print #tmpn%, "<br/>" 'spacers between columnheaders and time
           Print #tmpn%, "<br/>"
           Print #tmpn%, "<TABLE BORDER='1' CELLPADDING='1' CELLSPACING='1' ALIGN='CENTER' >"
        Else
           Print #tmpnum%, sEmpty 'spacers between column headers and times
           End If
           
        If ext$ = "xml" Then
            Print #xslnum%, "   <xsl:for-each select='chai/T1'>"
            Print #xslnum%, "      <TR ALIGN='center' FONTSIZE='1'>"
            Print #xslnum%, "         <xsl:choose>"
            Print #xslnum%, "            <xsl:when test ='position() mod 2 = 1'>"
            Print #xslnum%, "               <xsl:attribute name='STYLE'>background-color:yellow"
            Print #xslnum%, "               </xsl:attribute>"
            Print #xslnum%, "            </xsl:when>"
            Print #xslnum%, "            <xsl:otherwise>"
            Print #xslnum%, "               <xsl:attribute name='STYLE'>background-color:white"
            Print #xslnum%, "               </xsl:attribute>"
            Print #xslnum%, "            </xsl:otherwise>"
            Print #xslnum%, "         </xsl:choose>"
            
            'add commands for each column
            xslchild% = 0
            For m% = 0 To totnum% + 3 'there are a total of totnum% + 3 columns in the table
                xslchild% = xslchild% + 1
                Print #xslnum%, "             <TD><font size='3'><xsl:value-of select='t" & Trim$(Str$(xslchild%)) & "'/></font></TD>"
            Next m%
            xslchild% = 0
            
            'closing tags
            Print #xslnum%, "      </TR>"
            Print #xslnum%, "   </xsl:for-each>"
            Print #xslnum%, "</TABLE>"
            Print #xslnum%, "</BODY>"
            Print #xslnum%, "</HTML>"
            Print #xslnum%, "</xsl:template>"
            Print #xslnum%, "</xsl:stylesheet>"
            Close #xslnum%
            End If
            
        'create table's column headers
        totdoc$ = sEmpty 'totdoc$ contains one csv line of text
        linnum% = 1 'column headers are first row in the htm document
        If ext$ <> "csv" Then '(if csv, then the legend already acted as the column headers for the csv file)
           outdoc$ = sEmpty
           
          'add new line tag for this first line
          If ext$ = "htm" Or ext$ = "zip" Then
             Print #tmpnum%, "    <TR>"
          ElseIf ext$ = "xml" Then
             Print #tmpnum%, "<T1>"
             End If
           
           nn% = 0
           For m% = 1 To totnum% + 4
              If optionheb Then
                 Call hebnum(m%, cha$)
              Else
                 nn% = nn% + 1
                 cha$ = Trim$(Str$(nn%))
                 End If
              outdoc$ = cha$
              If m% = totnum% + 4 Then closerow = True
              GoSub parseit
              'outdoc$ = outdoc$ & comma$ & spacerb$ & cha$ & spacera$
           Next m%
           'GoSub parseit
           End If
        
        'now create the actual table
        numday% = -1
        For i% = 1 To endyr%
           If mmdate%(2, i%) > mmdate%(1, i%) Then
              k% = 0
              totdoc$ = sEmpty
              
              For j% = mmdate%(1, i%) To mmdate%(2, i%)
                  numday% = numday% + 1
                  k% = k% + 1
                  
                  'increment line number
                  linnum% = linnum% + 1
                  
                  'add new line tag for each new line
                  If ext$ = "htm" Or ext$ = "zip" Then
                    Print #tmpnum%, "    <TR>"
                  ElseIf ext$ = "xml" Then
                    Print #tmpnum%, "<T1>"
                    End If
              
                  outdoc$ = Trim$(stortim$(3, i% - 1, k% - 1))
                  GoSub parseit
                  Call InsertHolidays(calday$, i%, k%)
                  outdoc$ = calday$
                  GoSub parseit
                  outdoc$ = Trim$(stortim$(2, i% - 1, k% - 1))
                  GoSub parseit
                  For m% = 0 To totnum%
                     If Mid$(zmantimes(m%, numday%), 1, 2) = "00" Then
                        zmantimes(m%, numday%) = String$(6, "-")
                     ElseIf Mid$(zmantimes(m%, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m%, numday%), 1, 2) = "00"
                        End If
                     outdoc$ = Trim$(zmantimes(m%, numday%))
                     If m% = totnum% Then closerow = True
                     GoSub parseit
                  Next m%
                  
              Next j%
           ElseIf mmdate%(2, i%) < mmdate%(1, i%) Then
              k% = 0
              totdoc$ = sEmpty
              
              For j% = mmdate%(1, i%) To yrend%(0) 'yl1%
                  numday% = numday% + 1
                  k% = k% + 1
                  
                  'increment line number
                  linnum% = linnum% + 1
                  
                  'add new line tag for each new line
                  If ext$ = "htm" Or ext$ = "zip" Then
                    Print #tmpnum%, "    <TR>"
                  ElseIf ext$ = "xml" Then
                    Print #tmpnum%, "<T1>"
                    End If
              
                  outdoc$ = Trim$(stortim$(3, i% - 1, k% - 1))
                  GoSub parseit
                  Call InsertHolidays(calday$, i%, k%)
                  outdoc$ = calday$
                  GoSub parseit
                  outdoc$ = Trim$(stortim$(2, i% - 1, k% - 1))
                  GoSub parseit
                  For m% = 0 To totnum%
                     If Mid$(zmantimes(m%, numday%), 1, 2) = "00" Then
                        zmantimes(m%, numday%) = String$(6, "-")
                     ElseIf Mid$(zmantimes(m%, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m%, numday%), 1, 2) = "00"
                        End If
                     outdoc$ = Trim$(zmantimes(m%, numday%))
                     If m% = totnum% Then closerow = True
                     GoSub parseit
                  Next m%
              Next j%
              
              yrn% = yrn% + 1
              totdoc$ = sEmpty
              
              For j% = 1 To mmdate%(2, i%)
                  k% = k% + 1
                  
                  'increment line number
                  linnum% = linnum% + 1
                  
                  'add new line tag for each new line
                  If ext$ = "htm" Or ext$ = "zip" Then
                    Print #tmpnum%, "    <TR>"
                  ElseIf ext$ = "xml" Then
                    Print #tmpnum%, "<T1>"
                    End If
              
                  numday% = numday% + 1
                  outdoc$ = Trim$(stortim$(3, i% - 1, k% - 1))
                  GoSub parseit
                  Call InsertHolidays(calday$, i%, k%)
                  outdoc$ = calday$
                  GoSub parseit
                  outdoc$ = Trim$(stortim$(2, i% - 1, k% - 1))
                  GoSub parseit
                  For m% = 0 To totnum%
                     If Mid$(zmantimes(m%, numday%), 1, 2) = "00" Then
                        zmantimes(m%, numday%) = String$(6, "-")
                     ElseIf Mid$(zmantimes(m%, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m%, numday%), 1, 2) = "00"
                        End If
                     outdoc$ = Trim$(zmantimes(m%, numday%))
                     If m% = totnum% Then closerow = True
                     GoSub parseit
                  Next m%
              Next j%
              End If
        Next i%
        If ext$ = "htm" Or ext$ = "zip" Then
           Print #tmpnum%, "          </TABLE>"
           Print #tmpnum%, "<P><BR/><BR/>"
           Print #tmpnum%, "</P>"
           Print #tmpnum%, "</BODY>"
           Print #tmpnum%, "</HTML>"
        ElseIf ext$ = "xml" Then
           Print #tmpnum%, "</chai>"
           End If
        Close #tmpnum%
        eroscity$ = sEmpty
        
        If ext$ <> "zip" Then GoTo w500
'       now zip it and erase temporary htm file
        If Dir(drivjk$ & "pkzip.exe") <> sEmpty Then
            ret = Shell(drivjk$ & "pkzip " + filroot$ + ".zip " + filroot$ + ".htm", 6)
            nsloop% = 0
w50:        waitime = Timer
            Do Until Timer > waitime + 7 '<--!! 0.1 suffices for fast computer
               DoEvents
            Loop
            ret = FindWindow(vbNullString, "pkzip " + filroot$ + ".zip " + filroot$ + ".htm")
            nsloop% = nsloop% + 1
            If ret <> 0 And nsloop% < 10 Then GoTo w50
            
            If internet Then
               lognum% = FreeFile
               Open drivjk$ + "calprog.log" For Append As #lognum%
               Print #lognum%, "Step #11g: Zemanim table were zipped successfully"
               Close #lognum%
               End If
         Else
            If internet Then
               lognum% = FreeFile
               Open drivjk$ + "calprog.log" For Append As #lognum%
               Print #lognum%, "***PKZIP.EXE NOT FOUND IN THE JK DIRECTORY***"
               Close #lognum%
               End If
            End If
         
        On Error GoTo generrhand
        delfile$ = filroot$ + ".htm"
        lognum% = FreeFile
        Open drivjk$ + "calprog.log" For Append As #lognum%
        Print #lognum%, "Step #11g2: Trying to delete file: " + delfile$
        Close #lognum%
        waitime = Timer
        Do Until Timer > waitime + 0.5 '<--!! 0.1 suffices for fast computer
           DoEvents
        Loop
        If Dir(delfile$) <> sEmpty Then Kill delfile$
w500:   If internet Then
           lognum% = FreeFile
           Open drivjk$ + "calprog.log" For Append As #lognum%
           If zmantype% = 1 Then
              Print #lognum%, "Step #11h: Zemanim htm table deleted successfully"
           Else
              Print #lognum%, "Step #11g: Zemanim csv/xml table written successfully"
             End If
           Close #lognum%
           End If
        
        Screen.MousePointer = vbDefault
        
        If internet Then
           'unload the zmanim forms
           Unload Zmanimlistfm
           Unload Zmanimform
           End If
           
        Exit Sub
        
parseit: 'adds tags to html, and accumlates text for the csv file
        If ext$ = "csv" Then
           If totdoc$ = sEmpty Then
              totdoc$ = Chr$(34) & Trim$(outdoc$) & Chr$(34)
           Else
              totdoc$ = totdoc$ & "," & Chr$(34) & Trim$(outdoc$) & Chr$(34)
              End If
           If closerow Then Print #tmpnum%, totdoc$
        Else
           timlet$ = Trim$(outdoc$)
           'parse lines looking for spaces
           widthtim$ = CInt(Str$(Len(timlet$) * 4.5))
           If ext$ <> "xml" Then
              If linnum% Mod 2 = 0 Then
                 Print #tmpnum%, "        <TD WIDTH=" + Chr$(34) & widthtim$ & Chr$(34) + " BGCOLOR=""#ffff00"" >"
              Else
                 Print #tmpnum%, "        <TD WIDTH=" + Chr$(34) & widthtim$ & Chr$(34) + " >"
                 End If
                       
              Print #tmpnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""3"" >" & Trim$(timlet$) & "</FONT></P>"
              Print #tmpnum%, "        </TD>"
           Else
              xslchild% = xslchild% + 1
              Print #tmpnum%, "   <t" & Trim$(Str$(xslchild%)) & ">" & Trim$(timlet$) & "</t" & Trim$(Str$(xslchild%)) & ">"
              End If
         
           If closerow Then 'end of row signaled, so print end of row tags
              If ext$ = "xml" Then
                 Print #tmpnum%, "</T1>"
              Else
                 Print #tmpnum%, "    </TR>"
                 End If
              End If
              
           End If
            
       If closerow Then
          closerow = False
          totdoc$ = sEmpty
          xslchild% = 0
          End If
            
Return
        
generrhand:
     Screen.MousePointer = vbDefault
     If Err.Number = 53 Then
        response = MsgBox("Got it", vbCritical + vbOKOnly, "Cal Program")
        End If
     If internet = True And Err.Number >= 0 Then 'exit the program
        'abort the program with a error messages
        errlog% = FreeFile
        Open drivjk$ + "Cal_zzbgeh.log" For Output As errlog%
        Print #errlog%, "Cal Prog exited from Zmanimlistfm zmanbut_Click with runtime error message " + Str(Err.Number)
        Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
        Close #errlog%
        Close
      
       'unload forms
        For i% = 0 To Forms.Count - 1
          Unload Forms(i%)
        Next i%
      
        myfile = Dir(drivfordtm$ + "busy.cal")
        If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      
        'kill the timer
        If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
        'bring program to abrupt end
        End
     Else
        response = MsgBox("Enountered error #: " & Trim$(Str$(Err.Number)) & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & _
               "Do you want to abort the program?", vbYesNoCancel + vbCritical, "Cal Program")
        response = MsgBox("Zmanimlistfm zmanbut_click encountered error number: " + Str(Err.Number) + ".  Do you want to abort?", vbYesNoCancel + vbCritical, "Cal Program")
        If response <> vbYes Then
           Close
           Exit Sub
        Else
           Close
           For i% = 0 To Forms.Count - 1
             Unload Forms(i%)
           Next i%
           End
           End If
        End If
        
End Sub
Sub WriteTufikTables(filroot$, root$, ext$)

        'writes xml file in old format so it can be read by tufik program used for finding 4 years l'chumrah'filroot$ = output file name with directory
        'root$ = output file name without directory
        'ext$ = the file extension either: "htm", "zip", "csv", or "xml"

        'write html,xml version of sorted zmanim table
        tmpnum% = FreeFile

        If ext$ = "zip" Then
           Open filroot$ + "." + "htm" For Output As #tmpnum%
        Else
           Open filroot$ + "." + ext$ For Output As #tmpnum%
           End If
           
        If ext$ = "xml" Then 'write xsl file
           xslnum% = FreeFile 'also write xsl file
           Open filroot$ + ".xsl" For Output As #xslnum%
           End If
           
        filnam$ = filroot$ & ext$
        
        If ext$ = "html" Then
           FileCopy drivjk$ + "table.new", filnam$
        ElseIf ext$ = "htm" Or ext$ = "zip" Or ext$ = "xml" Then
           If ext$ = "zip" Then Mid$(filnam$, Len(filnam$) - 2, 3) = "htm"
           response = InputBox("Input a short description of the z'manim", "Z'manim header", sEmpty)
           hdr$ = response
           filnum% = FreeFile
           Open drivjk$ + "table.new" For Input As #filnum%
'           tmpnum% = FreeFile
'           Open filnam$ For Output As #tmpnum%
           If ext$ = "xml" Then 'also write the xml style sheet
'              xslnum% = FreeFile
'              xslfil$ = fileroot$ + ".xsl"
'              Open xslfil$ For Output As #xslnum%
              End If
           If ext$ <> "xml" Then
              Print #tmpnum%, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0//EN"">"
              Print #tmpnum%, "<HTML>"
           Else
              Print #tmpnum%, "<?xml version='1.0' encoding = 'iso-8859-8'?>"  'Window(Hebrews) uses code page 1255
              Print #tmpnum%, "<?xml-stylesheet type='text/xsl' href='" & root$ & ".xsl'?>"
              Print #tmpnum%, "<chai xml:space='preserve'>"
              Print #xslnum%, "<xsl:stylesheet xmlns:xsl='http://www.w3.org/TR/WD-xsl'>"
              Print #xslnum%, "   <xsl:template match='/'>"
              Print #xslnum%, "       <HTML>"
              End If
           If ext$ <> "xml" Then
              tmpn% = tmpnum%
           Else
              tmpn% = xslnum%
              End If
           Print #tmpn%, "<HEAD>"
           Print #tmpn%, "    <TITLE>Your Chai Z'manim Table</TITLE>"
           Print #tmpn%, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=iso-8859-8i""/>"
           Print #tmpn%, "    <META NAME=""AUTHOR"" CONTENT=""Chaim Keller, Chai Tables, www.zemanim.org""/>"
           Print #tmpn%, "</HEAD>"
           Print #tmpn%, "<BODY>"
           
           If ext$ = "xml" Then
                If optionheb = False Then
                   If eroscity$ = sEmpty Then
                      Print #tmpnum%, "<hd2><hd22>Chai Tables for " & citnamp$ & " for the year " & Str(yrheb%) & "</hd22></hd2>"
                   Else
                      Print #tmpnum%, "<hd2><hd22>Chai Tables for " & eroscity$ & " (" & citnamp$ & ")" & " for the year " & Str(yrheb%) & "</hd22></hd2>"
                      End If
                Else
                   If eroscity$ = sEmpty Then
                      Print #tmpnum%, "<hd2><hd22>" & heb3$(1) & hebcityname$ & heb3$(2) & Str(yrheb%) & "</hd22></hd2>"
                   Else
                      Print #tmpnum%, "<hd2><hd22>" & " (" & eroscity$ & ") " & heb3$(12) & " " & hebcityname$ & heb3$(2) & Str(yrheb%) & "</hd22></hd2>"
                      End If
                   End If
            Else
                If optionheb = False Then
                   If eroscity$ = sEmpty Then
                      Print #tmpnum%, "<h2><center>Chai Tables for " & citnamp$ & " for the year " & Str(yrheb%) & "</center></h2>"
                   Else
                      Print #tmpnum%, "<h2><center>Chai Tables for " & eroscity$ & " (" & citnamp$ & ")" & " for the year " & Str(yrheb%) & "</center></h2>"
                      End If
                Else
                   If eroscity$ = sEmpty Then
                      Print #tmpnum%, "<h2><center>" & heb3$(1) & hebcityname$ & heb3$(2) & Str(yrheb%) & "</h2>"
                   Else
                      Print #tmpnum%, "<h2><center>" & " (" & eroscity$ & ") " & heb3$(12) & " " & hebcityname$ & heb3$(2) & Str(yrheb%) & "</center></h2>"
                      End If
                   End If
               End If
            
         address$ = heb2$(14) & Str(datavernum) & "." & Str(progvernum) & " ©"
         If optionheb = False Then
            address$ = "© Luchos Chai, Fattal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: " & Str(progvernum) & "." & Str(datavernum)
            End If
           
           
        If ext$ = "xml" Then
          Print #xslnum%, "<xsl:for-each select='chai/hd2'>"
          Print #xslnum%, "   <h2><center><xsl:value-of select='hd22'/></center></h2>"
          Print #xslnum%, "</xsl:for-each>"
          Print #xslnum%, "<xsl:for-each select='chai/hd3'>"
          Print #xslnum%, "   <h3><center><xsl:value-of select='hd33'/></center></h3>"
          Print #xslnum%, "</xsl:for-each>"
          Print #xslnum%, "<xsl:for-each select='chai/address'>"
          Print #xslnum%, "   <h5><center><xsl:value-of select='name'/></center></h5>"
          Print #xslnum%, "</xsl:for-each>"
          Print #tmpn%, ""
          Print #tmpnum%, "<hd3><hd33>" & hdr$ & "</hd33></hd3>"
          Print #tmpnum%, "<address><name>" & address$ & "</name></address>"
        Else
          Print #tmpnum%, "<h3><center>" & hdr$ & "</center></h3>"
          Print #tmpnum%, "<h5><center>" & address$ & "</center></h5>"
          End If
              
              
           Line Input #filnum%, doclin$
           'parse zemanim caption: replace 3 spaces with delimiter line
           doclin$ = LTrim$(RTrim$(doclin$))
           doclen% = Len(doclin$)
           isr% = 1
h25:       If Mid$(doclin$, isr%, 1) = " " Then
              If Mid$(doclin$, isr%, 3) = "   " Then
                 Mid$(doclin$, isr%, 3) = " | "
                 isr% = isr% + 3
              Else
                 isr% = isr% + 1
                 End If
           Else
              isr% = isr% + 1
              End If
           If isr% <= doclen% - 2 Then GoTo h25

           Print #tmpn%, "<hr/>"
           Print #tmpn%, "<br/>"
           If ext$ = "xml" Then
              Print #tmpnum%, "<hd4><hd44>" & doclin$ & "</hd44></hd4>" 'this caption line of zemanim"
              Print #xslnum%, "<xsl:for-each select='chai/hd4'>"
              Print #xslnum%, "   <xsl:value-of select='hd44'/>"
              Print #xslnum%, "</xsl:for-each>"
           Else
              Print #tmpnum%, doclin$
              End If
           Print #tmpn%, "<br/>"
           Print #tmpn%, "<br/>"
           Print #tmpn%, "<TABLE BORDER='1' CELLPADDING='1' CELLSPACING='1' ALIGN = 'CENTER' >"
           linnum% = 0
           noxsl = False
           Do Until EOF(filnum%)
              Line Input #filnum%, doclin$
              linnum% = linnum% + 1
              'pos% = InStr(doclin$, "-")
              'If pos% <> 0 Then
              '   apart$ = Mid(doclin$, 1, pos% - 1)
              '   bpart$ = Mid(doclin$, pos%, 6)
              '   cpart$ = Mid(doclin$, pos% + 9, Len(doclin$) - 3)
              '   doclin$ = apart$ + bpart$ + cpart$
              '   'doclin$ = Mid(doclin$, 1, pos% - 1) + Mid(doclin$, pos%, pos% + 2) + Mid(doclin$, pos% + 9, Len(doclin$) - 4)
              '   End If
              If ext$ = "xml" Then
                 Print #tmpnum%, "<T1>"
                 If noxsl = False Then
                    Print #xslnum%, "   <xsl:for-each select='chai/T1'>"
                    Print #xslnum%, "      <TR ALIGN='center' FONTSIZE='1'>"
                    Print #xslnum%, "         <xsl:choose>"
                    Print #xslnum%, "            <xsl:when expr='(childNumber(this) % 2) == 1'>"
                    Print #xslnum%, "               <xsl:attribute name='STYLE'>background-color:yellow"
                    Print #xslnum%, "               </xsl:attribute>"
                    Print #xslnum%, "            </xsl:when>"
                    Print #xslnum%, "            <xsl:otherwise>"
                    Print #xslnum%, "               <xsl:attribute name='STYLE'>background-color:white"
                    Print #xslnum%, "               </xsl:attribute>"
                    Print #xslnum%, "            </xsl:otherwise>"
                    Print #xslnum%, "         </xsl:choose>"
                    End If
              Else
                 Print #tmpnum%, "    <TR>"
                 End If
              'parse lines looking for spaces
              fin1% = 0
              xslchild% = 0
              For ispa% = 1 To Len(doclin$)
                 If Mid$(doclin$, ispa%, 1) = " " And fin1% = 0 Then
                    fin1% = ispa%
                 ElseIf Mid$(doclin$, ispa%, 1) = " " Or ispa% = Len(doclin$) And fin1% <> 0 Then
                    timlet$ = Mid$(doclin$, fin1% + 1, ispa% - fin1%)
                    If LTrim$(RTrim$(timlet$)) = sEmpty Then GoTo h50
                    widthtim$ = CInt(Str$(Len(timlet$) * 4.5))
                    xslchild% = xslchild% + 1
                    If ext$ <> "xml" Then
                        If linnum% Mod 2 = 0 Then
                           Print #tmpnum%, "        <TD WIDTH="" + widthtim$ + "" BGCOLOR=""#ffff00"" >"
                        Else
                           Print #tmpnum%, "        <TD WIDTH="" + widthtim$ + "" >"
                           End If
                        Print #tmpnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""3"" >" & LTrim$(RTrim$(timlet$)) & "</FONT></P>"
                        Print #tmpnum%, "        </TD>"
                    Else
                       Print #tmpnum%, "   <t" & LTrim$(RTrim$(Str$(xslchild%))) & ">" & LTrim$(RTrim$(timlet$)) & "</t" & LTrim$(RTrim$(Str$(xslchild%))) & ">"
                       If noxsl = False Then Print #xslnum%, "             <TD><font size='3'><xsl:value-of select='t" & LTrim$(RTrim$(Str$(xslchild%))) & "'/></font></TD>"
                       End If
                    fin1% = 0
                    End If
                    
h50:
               Next ispa%
               If ext$ = "xml" Then
                  If noxsl = False Then
                     Print #tmpn%, "    </TR>"
                     Print #tmpn%, "  </xsl:for-each>"
                     End If
                  If ext$ = "xml" Then
                     Print #tmpnum%, "</T1>"
                     noxsl = True
                     End If
               Else
                  Print #tmpnum%, "    </TR>"
                  End If
              'Print #tmpnum%, "<BR/>" & doclin$
           Loop
           Print #tmpn%, "          </TABLE>"
           If ext$ = "xml" Then
              Print #xslnum%, "           </BODY>"
              Print #xslnum%, "          </HTML>"
              Print #xslnum%, "        </xsl:template>"
              Print #xslnum%, "      </xsl:stylesheet>"
              Print #tmpnum%, "</chai>"
              Close #xslnum%
           Else
              Print #tmpnum%, "<P><BR/><BR/>"
              Print #tmpnum%, "</P>"
              Print #tmpnum%, "</BODY>"
              Print #tmpnum%, "</HTML>"
              End If
           Close #filnum%
           Close #tmpnum%
           
           If ext$ = "zip" Then
             ret = Shell("pkzip " + fileroot$ + ".zip " + filnam$, 6)
             waitime = Timer
             Do Until Timer > waitime + 0.1
                DoEvents
             Loop
             If Dir(filnam$) <> sEmpty Then Kill filnam$
             End If
           End If
        Screen.MousePointer = vbDefault
        changes = False
c3error:
        Exit Sub

End Sub
'---------------------------------------------------------------------------------------
' Procedure : PrinttoDev
' DateTime  : 11/8/2005 03:06
' Author    : Chaim Keller
' Purpose   : generates print preview and prints to printer
' Dev = device to be previewed on printed on
' PrinterFlag = True if printing
'             = False if previewing
'---------------------------------------------------------------------------------------
'
Sub PrinttoDev(Dev, PrinterFlag As Boolean)

Dim lResult As Long, cirx As Single, ciry As Single

'//////////////////DST support for Israel, USA added 082921/////////////////////////////////////
Dim stryrDST%, endyrDST%, strdaynum(1) As Integer, enddaynum(1) As Integer

Dim MarchDate As Integer
Dim OctoberDate As Integer
Dim NovemberDate As Integer
Dim YearLength As Integer
Dim DSThour As Integer
   
Dim DSTadd As Boolean

Dim DSTPerpetualIsrael As Boolean
Dim DSTPerpetualUSA As Boolean

'set to true when these countries adopt universal DST
DSTPerpetualIsrael = False
DSTPerpetualUSA = False

If CalMDIform.mnuDST.Checked = True Then
   DSTadd = True
   End If
   
If Option2b Then
   If yrheb% < 1918 Then DSTadd = False
Else
   If yrheb% < 5678 Then DSTadd = False
   End If
  
If DSTadd Then

   If Not Option2b Then 'hebrew years
      stryrDST% = yrheb% + RefCivilYear% - RefHebYear% '(yrheb% - 5758) + 1997
      endyrDST% = yrheb% + RefCivilYear% - RefHebYear% + 1 '(yrheb% - 5758) + 1998
   Else
      stryrDST% = yrheb%
      endyrDST% = yrheb%
      End If
      
   'find beginning and ending day numbers for each civil year
   Select Case eroscountry$
   
      Case "Israel", "" 'EY eros or cities using 2017 DST rules
      
          MarchDate = (31 - (Fix(stryrDST% * 5 / 4) + 4) Mod 7) - 2 'starts on Friday = 2 days before EU start on Sunday
          OctoberDate = (31 - (Fix(stryrDST% * 5 / 4) + 1) Mod 7)
          YearLength% = DaysinYear(stryrDST%)
          strdaynum(0) = DayNumber(YearLength%, 3, MarchDate)
          enddaynum(0) = DayNumber(YearLength%, 10, OctoberDate)
          
          If DSTPerpetualIsrael Then
             strdaynum(0) = 1
             enddaynum(0) = YearLength%
             End If
             
          MarchDate = (31 - (Fix(endyrDST% * 5 / 4) + 4) Mod 7) - 2 'starts on Friday = 2 days before EU start on Sunday
          OctoberDate = (31 - (Fix(endyrDST% * 5 / 4) + 1) Mod 7)
          YearLength% = DaysinYear(endyrDST%)
          strdaynum(1) = DayNumber(YearLength%, 3, MarchDate)
          enddaynum(1) = DayNumber(YearLength%, 10, OctoberDate)

          If DSTPerpetualIsrael Then
             strdaynum(1) = 1
             enddaynum(1) = YearLength%
             End If
        
      
      Case "USA", "Canada" 'English {USA DST rules}
      
        'not all states in the US have DST
        If InStr(eroscity$, "Phoenix") Or InStr(eroscity$, "Honolulu") Or InStr(eroscity$, "Regina") Then
           DSTadd = False
        Else
      
          MarchDate = 14 - (Fix(1 + stryrDST% * 5 / 4) Mod 7)
          NovemberDate = 7 - (Fix(1 + stryrDST% * 5 / 4) Mod 7)
          YearLength% = DaysinYear(stryrDST%)
          strdaynum(0) = DayNumber(YearLength%, 3, MarchDate)
          enddaynum(0) = DayNumber(YearLength%, 11, NovemberDate)
          
          If DSTPerpetualUSA Then
             strdaynum(0) = 1
             enddaynum(0) = YearLength%
             End If
             
          MarchDate = 14 - (Fix(1 + endyrDST% * 5 / 4) Mod 7)
          NovemberDate = 7 - (Fix(1 + endyrDST% * 5 / 4) Mod 7)
          YearLength% = DaysinYear(endyrDST%)
          strdaynum(1) = DayNumber(YearLength%, 3, MarchDate)
          enddaynum(1) = DayNumber(YearLength%, 11, NovemberDate)
          
          If DSTPerpetualUSA Then
             strdaynum(1) = 1
             enddaynum(1) = YearLength%
             End If
             
          End If
             
      Case "England", "UK", "France", "Germany", "Netherlands", "Belgium", _
           "Northern_Ireland", "Yugoslavia", "Slovakia", "Romania", "Hungary", _
           "Denmark", "Ireland", "Switzerland", "Finland", "Ukraine", "Norway", _
           "France", "Czechoslovakia", "Sweden", "Italy", "Europe"

          MarchDate = (31 - (Fix(stryrDST% * 5 / 4) + 4) Mod 7) 'starts on Sunday, 2 days after EY
          OctoberDate = (31 - (Fix(stryrDST% * 5 / 4) + 1) Mod 7)
          YearLength% = DaysinYear(stryrDST%)
          strdaynum(0) = DayNumber(YearLength%, 3, MarchDate)
          enddaynum(0) = DayNumber(YearLength%, 10, OctoberDate)
          
          If DSTPerpetualIsrael Then
             strdaynum(0) = 1
             enddaynum(0) = YearLength%
             End If
             
          MarchDate = (31 - (Fix(endyrDST% * 5 / 4) + 4) Mod 7) 'starts on Sunday = 2 days after EY
          OctoberDate = (31 - (Fix(endyrDST% * 5 / 4) + 1) Mod 7)
          YearLength% = DaysinYear(endyrDST%)
          strdaynum(1) = DayNumber(YearLength%, 3, MarchDate)
          enddaynum(1) = DayNumber(YearLength%, 10, OctoberDate)

          If DSTPerpetualIsrael Then
             strdaynum(1) = 1
             enddaynum(1) = YearLength%
             End If

      
      Case Else 'not implemented yet for other countries
         DSTadd = False
      
   End Select
   End If
   
'///////////////////////////////////////////////////////////////////////////////////////////////////////

On Error GoTo generrhand

If Not PrinterFlag Then
   Screen.MousePointer = vbHourglass
Else
   rescal = 1
   End If

If Not PrinterFlag Then Screen.MousePointer = vbHourglass
If Option2b = True Then
   previewfm.zmanbut.Enabled = False
Else
   previewfm.zmanbut.Enabled = True
   End If
zmannetz = False: zmanskiy = False
'first check that the calculations finished
TotalWaitTime = 0
5 myfile = Dir(drivjk$ + "netzend.tmp")
  If myfile = sEmpty Then
     Waitfm.Visible = True
     'ret = SetWindowPos(Waitfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
     waittim = Timer + 0.2
     TotalWaitTime = TotalWaitTime + 0.2
     Do Until Timer > waittim
        DoEvents
     Loop
     'check if exceeded maximum wait time of 5 minutes (300 seconds)
     If TotalWaitTime > 300 Then
        If internet = True Then
           'Abort program after updating log file.
            lognum% = FreeFile
            Open drivjk$ + "calprog.log" For Append As #lognum%
            Print #lognum%, "Step #11-0: Exceeded wait time-Abort (see Cal_pbgeh)."
            Close #lognum%
            errlog% = FreeFile
            Open drivjk$ + "Cal_pbgeh.log" For Output As #errlog%
            Print #errlog%, "Cal Prog aborted during previewbut code due to exceeding time limit"
            Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
            Close #errlog%
            Close
      
           myfile = Dir(drivfordtm$ + "busy.cal")
           If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
           
           'unload forms
           For i% = 0 To Forms.Count - 1
             Unload Forms(i%)
           Next i%
      
          'kill the timer
          If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
          'bring program to abrupt end
          End
        Else
           response = MsgBox("Have waited for already 5 minutes.  Do you wan't to wait some more?", vbQuestion + vbYesNo, "Cal Program")
           If response = vbNo Then
              newhebExitbut.Value = 1
           Else
              'double the wait time
              TotalWaitTime = TotalWaitTime - 300
           End If
        End If
     End If
     GoTo 5
     End If
  'do second check using Windows API
'  If internet = False Then
'     lResult = FindWindow(vbNullString, ProgExec$)
'  Else
'     lResult = FindWindow(vbNullString, ProgExec$)
'     End If
'  If lResult <> 1443154 And lResult <> 0 Then GoTo 5 'it is still working, so keep on looping
  Waitfm.Visible = False
  
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-a: newhebcalfm: beginning filling arrays"
'Close #lognum%
'End If
 
 
 If hebcal = True Then
   ntable% = 30
   endyr% = 12
   If hebleapyear = True Then endyr% = 13
 ElseIf hebcal = False Then
   ntable% = 31
   endyr% = 12
   End If
'Dim tbl1$(31)
Dim header$(6)
'Dim monthe$(12), monthh$(iheb%,14), mdates$(3, 13), mmdate%(2, 13), montheh$(12)
'Dim tim$(2, 366)
Dim CN4 As String * 4
Dim coordxlab1(2), coordylab1(2), coordxlab2(2), coordylab2(2)
Dim coordxlab3(2), coordylab3(2), coordxlab4(2), coordylab4(2)
Dim coordxlab5(2), coordylab5(2), waiterror As Boolean
Dim coordxlab6(2), coordylab6(2)
Dim coordxmonreg(13), coordymon(2) ', coordxmonleap(13)
Dim coordxreg(2, 13), coordy(2, 31) ', coordxleap(2, 13)
Dim cap1$(2), cap2$(2), cap3$(2), cap4$(2), cap5$(2), cap6$(2)

monthe$(1) = "Jan-"
monthe$(2) = "Feb-"
monthe$(3) = "Mar-"
monthe$(4) = "Apr-"
monthe$(5) = "May-"
monthe$(6) = "Jun-"
monthe$(7) = "Jul-"
monthe$(8) = "Aug-"
monthe$(9) = "Sep-"
monthe$(10) = "Oct-"
monthe$(11) = "Nov-"
monthe$(12) = "Dec-"
monthh$(1, 1) = "Tishrey"
monthh$(1, 2) = "Chesvan"
monthh$(1, 3) = "Kislev"
monthh$(1, 4) = "Teves"
monthh$(1, 5) = "Shvat"
monthh$(1, 6) = "Adar I"
monthh$(1, 7) = "Adar II"
monthh$(1, 8) = "Nisan"
monthh$(1, 9) = "Iyar"
monthh$(1, 10) = "Sivan"
monthh$(1, 11) = "Tamuz"
monthh$(1, 12) = "Av"
monthh$(1, 13) = "Elul"
monthh$(1, 14) = "Adar"
myfile = Dir(drivjk$ + "Calhebmonths.txt")
If myfile = sEmpty Then
    monthh$(0, 1) = "תשרי"
    monthh$(0, 2) = "מרחשון"
    monthh$(0, 3) = "כסלו"
    monthh$(0, 4) = "טבת"
    monthh$(0, 5) = "שבט"
    monthh$(0, 6) = "'אדר א"
    monthh$(0, 7) = "'אדר ב"
    monthh$(0, 8) = "ניסן"
    monthh$(0, 9) = "אייר"
    monthh$(0, 10) = "סיון"
    monthh$(0, 11) = "תמוז"
    monthh$(0, 12) = "אב"
    monthh$(0, 13) = "אלול"
    monthh$(0, 14) = "אדר"
Else
   calhebnum% = FreeFile
   Open drivjk$ + myfile For Input As #calhebnum%
   For ical% = 1 To 14
      Input #calhebnum%, monthh$(0, ical%)
   Next ical%
   Close #calhebnum%
   End If
iheb% = 0: If optionheb = False Then iheb% = 1
montheh$(0, 1) = "ינואר"
montheh$(0, 2) = "פברואר"
montheh$(0, 3) = "מרץ"
montheh$(0, 4) = "אפריל"
montheh$(0, 5) = "מאי"
montheh$(0, 6) = "יוני"
montheh$(0, 7) = "יולי"
montheh$(0, 8) = "אוגוסט"
montheh$(0, 9) = "ספטמבר"
montheh$(0, 10) = "אוקטובר"
montheh$(0, 11) = "נובמבר"
montheh$(0, 12) = "דצמבר"
montheh$(1, 1) = "January"
montheh$(1, 2) = "February"
montheh$(1, 3) = "March"
montheh$(1, 4) = "April"
montheh$(1, 5) = "May"
montheh$(1, 6) = "June"
montheh$(1, 7) = "July"
montheh$(1, 8) = "August"
montheh$(1, 9) = "September"
montheh$(1, 10) = "October"
montheh$(1, 11) = "November"
montheh$(1, 12) = "December"
'montheh$(1) = "יאנ"
'montheh$(2) = "פבר"
'montheh$(3) = "מרץ"
'montheh$(4) = "אפרי"
'montheh$(5) = "מאי"
'montheh$(6) = "יוני"
'montheh$(7) = "יולי"
'montheh$(8) = "אוג"
'montheh$(9) = "ספט"
'montheh$(10) = "אוקט"
'montheh$(11) = "נוב"
'montheh$(12) = "דצמ"
'For i% = 0 To 1
'   For j% = 0 To 12
'      For k% = 0 To 30
'          stortim$(1, 12, 30) = sEmpty
'      Next k%
'   Next j%
'Next i%
waiterror = False
yr% = yrheb%
If yr% < 5000 Then
   yrcal$ = Str$(yr%)
   GoTo n100
   End If
If hebcal = True Then 'translate into hebrew characters
   Call hebyear(yr%, chyear$)
   yrcal$ = chyear$
Else
   yrcal$ = Trim$(CStr(yr%))
   End If

n100:
'********Katz changes*********
katzyo% = 0
'*****************************

'define coordinates for headers and colums (numbers in mm)
'top calendar:
   If rescale = False Then Dev.ScaleMode = 6  'chose mm as the scale
   If PrinterFlag Then
      Dev.ScaleMode = 6 'scale is in mm's
      If portrait = False Then 'set paper orientation
         Printer.Orientation = vbPRORLandscape
      ElseIf portrait = True Then
         Printer.Orientation = vbPRORPortrait
         End If
      End If

   conv = 0.01 'conversion from inputed to actual .01 mm values
   xo = Val(newhebcalfm.Text20.Text): yo = Val(newhebcalfm.Text21.Text)
   xot = Val(newhebcalfm.Text22.Text): yot = Val(newhebcalfm.Text23.Text)
   dx = Val(newhebcalfm.Text24.Text): dy = Val(newhebcalfm.Text25.Text)
   ys(1) = 0: ys(2) = Val(newhebcalfm.Text29.Text)
   xc(1) = Val(newhebcalfm.Text16.Text): xc(2) = Val(newhebcalfm.Text33.Text)
   y1(1) = Val(newhebcalfm.Text17.Text): y1(2) = Val(newhebcalfm.Text34.Text)
   y2(1) = Val(newhebcalfm.Text18.Text): y2(2) = Val(newhebcalfm.Text35.Text)
   y3(1) = Val(newhebcalfm.Text19.Text): y3(2) = Val(newhebcalfm.Text36.Text)
   y4(1) = Val(newhebcalfm.Text26.Text): y4(2) = Val(newhebcalfm.Text37.Text)
   y5(1) = Val(newhebcalfm.Text39.Text): y5(2) = Val(newhebcalfm.Text40.Text)
   de(1) = Val(newhebcalfm.Text30.Text): de(2) = Val(newhebcalfm.Text38.Text)
   dey(1) = Val(newhebcalfm.Text27.Text): dey(2) = Val(newhebcalfm.Text28.Text)
   'cap1$(1) = newhebcalfm.Combo1.Text + " לשנת " + yrcal$: cap1$(2) = newhebcalfm.Combo6.Text + " לשנת " + yrcal$
   cap1$(1) = newhebcalfm.Combo1.Text + heb2$(1) + " " + yrcal$: cap1$(2) = newhebcalfm.Combo6.Text + heb2$(1) + " " + yrcal$
   If optionheb = False Then
      cap1$(1) = newhebcalfm.Combo1.Text + " for the year " + Str(yrheb%): cap1$(2) = newhebcalfm.Combo6.Text + " for the year " + LTrim$(Str(yrheb%))
      End If
   cap2$(1) = newhebcalfm.Combo2.Text: cap2$(2) = newhebcalfm.Combo7.Text
   cap3$(1) = newhebcalfm.Combo3.Text: cap3$(2) = newhebcalfm.Combo8.Text
   cap4$(1) = newhebcalfm.Combo4.Text: cap4$(2) = newhebcalfm.Combo9.Text
   cap5$(1) = newhebcalfm.Combo5.Text: cap5$(2) = newhebcalfm.Combo10.Text
   If newhebcalfm.Check4.Value = vbChecked Then
      If optionheb = True Then
        'cap5$(1) = cap5$(1) + "  " + "השבתות מסומנות ע" + Chr$(34) + "י קו תחתון."
        'cap5$(2) = cap5$(2) + "  " + "השבתות מסומנות ע" + Chr$(34) + "י קו תחתון."
        cap5$(1) = cap5$(1) + "  " + heb2$(2) + Chr$(34) + heb2$(3)
        cap5$(2) = cap5$(2) + "  " + heb2$(2) + Chr$(34) + heb2$(3)
      Else
         cap5$(1) = cap5$(1) + " Shabbosim are underlined."
         cap5$(2) = cap5$(2) + " Shabbosim are underlined."
         End If
      End If
   
   'address$ = "© לוחות חי, טל/פקס: 5713765(02).  גרסא: " & Str(datavernum) & "." & Str(progvernum)
   address$ = heb2$(4) & Str(datavernum) & "." & Str(progvernum)
   If optionheb = False Then
      address$ = "© Luchos Chai, Fahtal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: " & Str(progvernum) & "." & Str(datavernum)
      End If
   
   If newhebcalfm.Check2.Value = vbChecked Then
      cap5$(1) = cap5$(1) + "  " + address$
      cap5$(2) = cap5$(2) + "  " + address$
      End If
   '***********Katz changes********
   If Katz = True Then
      cap1$(1) = newhebcalfm.Combo1.Text: cap1$(2) = newhebcalfm.Combo6.Text
      cap2$(1) = sEmpty: cap2$(2) = sEmpty
      cap3$(1) = sEmpty: cap3$(2) = sEmpty
      cap4$(1) = sEmpty: cap4$(2) = sEmpty
      cap5$(1) = sEmpty: cap5$(2) = sEmpty
      End If
   '*******************************
   If nearcolor = True Then
'הסתרים אלו לגמרי
      If nearnez = True Then
         'cap6$(1) = ".הזמנים הכתובים בצבע בהיר הם בימים שישנם הסתרים קרובים" + " (פחות מ-" + RTrim$(Str(distlim)) + " ק" + Chr$(34) + "מ), ויתכן שזמנים אלו אינם מדוייקים מספיק, או שיש לנכות הסתרים אלו לגמרי"
         cap6$(1) = heb2$(5) + heb2$(6) + RTrim$(Str(distlim)) + heb2$(7) + Chr$(34) + heb2$(8)
         If optionheb = False Then
            cap6$(1) = "The lighter colors denote times that may be inaccurate due to near obstructions (closer than " & RTrim$(Str(distlim)) & "km).  It is possible that such near obstructions should be ignored."
            End If
         If automatic = True Then
            'cap6$(1) = "(!הזמנים הכתובים בצבע בהיר הם בימים שישנם הסתרים קרובים" + " (פחות מ-" + RTrim$(Str(distlim)) + " ק" + Chr$(34) + "מ), ויתכן שזמנים אלו אינם מדוייקים מספיק, או שיש לנכות הסתרים אלו לגמרי" + " (" + "ראה מבוא"
            cap6$(1) = heb2$(9) + heb2$(6) + RTrim$(Str(distlim)) + heb2$(7) + Chr$(34) + heb2$(8) + " (" + heb2$(10)
            End If
'         cap6$(1) = "הזמנים הכתובים בצבע בהיר מבוססים על אופק קרוב מידי, ויתכן שאינם מדוייקים"
         End If
      If nearski = True Then
         'cap6$(2) = ".הזמנים הכתובים בצבע בהיר הם בימים שישנם הסתרים קרובים" + " (פחות מ-" + RTrim$(Str(distlim)) + " ק" + Chr$(34) + "מ), ויתכן שזמנים אלו אינם מדוייקים מספיק, או שיש לנכות הסתרים אלו לגמרי"
         cap6$(2) = heb2$(5) + heb2$(6) + RTrim$(Str(distlim)) + heb2$(7) + Chr$(34) + heb2$(8)
         If optionheb = False Then
            cap6$(2) = "The lighter colors denote times that may be inaccurate due to near obstructions (closer than " & RTrim$(Str(distlim)) & "km).  It is possible that such near obstructions should be ignored."
            End If
         If automatic = True Then
            'cap6$(2) = "(!הזמנים הכתובים בצבע בהיר הם בימים שישנם הסתרים קרובים" + " (פחות מ-" + RTrim$(Str(distlim)) + " ק" + Chr$(34) + "מ), ויתכן שזמנים אלו אינם מדוייקים מספיק, או שיש לנכות הסתרים אלו לגמרי" + " (" + "ראה מבוא"
            cap6$(2) = heb2$(9) + heb2$(6) + RTrim$(Str(distlim)) + heb2$(7) + Chr$(34) + heb2$(8) + " (" + heb2$(10)
            End If
'         cap6$(2) = "הזמנים הכתובים בצבע בהיר מבוססים על אופק קרוב מידי, ויתכן שאינם מדוייקים"
         End If
      End If
      If AddObsTime = 1 Then
         'add captions for adding additional time for near obstructions
         If optionheb Then
            cap6$(1) = heb2$(15)
            cap6$(2) = heb2$(15)
         Else
            cap6$(1) = "A larger cushion has been used for days where the horizon is obstructed by near obstructions."
            cap6$(2) = "A larger cushion has been used for days where the horizon is obstructed by near obstructions."
            End If
         End If
   For i% = 1 To 2
      coordxlab1(i%) = (xo + xc(i%)) * conv ' - Dev.TextWidth(cap1$(i%)) ' * Val(newhebcalfm.Text13.Text) * 0.12 / 2
      coordylab1(i%) = (yo + ys(i%) + y1(i%)) * conv
      coordxlab2(i%) = (xo + xc(i%)) * conv ' - Dev.TextWidth(cap2$(i%)) ' * Val(newhebcalfm.Text14.Text) * 0.13 / 2
      coordylab2(i%) = (yo + ys(i%) + y2(i%)) * conv
      coordxlab3(i%) = (xo + xc(i%)) * conv ' - Dev.TextWidth(cap3$(i%)) ' * Val(newhebcalfm.Text14.Text) * 0.13 / 2 '- Len(cap3$(i%)) * siz2 / 2
      coordylab3(i%) = (yo + ys(i%) + y3(i%)) * conv
      coordxlab4(i%) = (xo + xc(i%)) * conv ' - Dev.TextWidth(cap4$(i%)) * Val(newhebcalfm.Text14.Text) * 0.13 / 2 '- Len(cap4$(i%)) * siz2 / 2
      coordylab4(i%) = (yo + ys(i%) + y4(i%)) * conv
      coordxlab5(i%) = (xo + xc(i%)) * conv
      coordylab5(i%) = (yo + ys(i%) + y5(i%)) * conv
      coordxlab6(i%) = coordxlab5(i%)
      coordylab6(i%) = coordylab5(i%) + 350 * conv '250 * conv
      coordymon(i%) = (yo + ys(i%) + yot) * conv '(yo + ys(i%) + ygrid - de(i%)) * conv '17,235
   Next i%
   
   For i% = 1 To endyr%
     coordxreg(1, i%) = (xo + xot - (i% - 1) * dx) * conv '140 - (i% - 1) * 10.2
     coordxreg(2, i%) = coordxreg(1, i%)
'     If i% < 6 Then
'        k% = i%
'     ElseIf k% = 6 Then
'        k% = 14
'     ElseIf k% > 6 Then
'        k% = i% + 1
'        End If
'     If hebcal = True Then
        coordxmonreg(i%) = (xo + xot - (i% - 1) * dx) * conv '+ (8 - Len(monthh$(iheb%,k%))) / 2 ' 140 + (8 - Len(monthh$(iheb%,k%))) / 2 - (i% - 1) * 10.2
'     Else
'        coordxmonreg(i%) = (xo + xot - (i% - 1) * dx) * conv '+ (8 - Len(montheh$(i%))) / 2 ' 140 + (8 - Len(montheh$(i%))) / 2 - (i% - 1) * 10.2
'        End If
   Next i%
   If optionheb = False Then
       j% = 0
       For i% = endyr% To 1 Step -1
         j% = j% + 1
         coordxreg(1, j%) = (xo + xot - (i% - 1) * dx) * conv '140 - (i% - 1) * 10.2
         coordxreg(2, j%) = coordxreg(1, j%)
    '     If i% < 6 Then
    '        k% = i%
    '     ElseIf k% = 6 Then
    '        k% = 14
    '     ElseIf k% > 6 Then
    '        k% = i% + 1
    '        End If
    '     If hebcal = True Then
            coordxmonreg(j%) = (xo + xot - (i% - 1) * dx) * conv '+ (8 - Len(monthh$(iheb%,k%))) / 2 ' 140 + (8 - Len(monthh$(iheb%,k%))) / 2 - (i% - 1) * 10.2
    '     Else
    '        coordxmonreg(i%) = (xo + xot - (i% - 1) * dx) * conv '+ (8 - Len(montheh$(i%))) / 2 ' 140 + (8 - Len(montheh$(i%))) / 2 - (i% - 1) * 10.2
    '        End If
       Next i%
      End If
   For i% = 1 To 31
     coordy(1, i%) = (yo + yot + (i% - 1) * dy) * conv '+ (i% - 1) * 2.12 '20 + (i% - 1) * 2.12
     coordy(2, i%) = (yo + ys(2) + yot + (i% - 1) * dy) * conv
   Next i%
'   For i% = 1 To 13
'     coordxleap(1, i%) = 150 - (i% - 1) * 12
'     coordxleap(2, i%) = coordxleap(1, i%)
'     coordxmonleap(i%) = 150 - (i% - 1) * 12
'   Next i%

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-b: newhebcalfm: before .visible=true"
'Close #lognum%
'End If

previewfm.Visible = True
'ret = SetWindowPos(previewfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'If Pageformatfm.Visible = True Then
'   ret = SetWindowPos(Pageformatfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'   End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-c: newhebcalfm: before reading netzskiy.tm3"
'Close #lognum%
'End If


On Error GoTo newheberror     '<<<<<<<<????????
nheb5: filtm3% = FreeFile
Open drivjk$ + "netzskiy.tm3" For Input As #filtm3%
Input #filtm3%, yr%           'hebrew or civil year for calculation
Input #filtm3%, nsetflag%, geon% '1=netz only,2=skiy only,3=both; geon%=geotz!
zmansetflag% = nsetflag%
If nsetflag% <= -4 And eros = False Then
   nsetflag% = nsetflag% + 3
ElseIf nsetflag% <= -4 And eros = True Then
   nsetflag = Abs(nsetflag% + 3)
   End If
Input #filtm3%, numfilo%
Input #filtm3%, CN4           'place file prefix
Close #filtm3%

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-d: newhebcalfm: after reading netzskiy.tm3"
'Close #lognum%
'End If

If PrinterFlag Then GoTo pf100

First = True
If First = True And portrait = True Then
   Dev.Cls
   previewfm.Visible = True
   'ret = SetWindowPos(previewfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   Dev.DrawMode = 13
   If paperwidth = 0 Then paperwidth = 200
   If paperheight = 0 Then paperheight = 250
   If leftmargin = 0 Then leftmargin = 10
   If rightmargin = 0 Then rightmargin = 10
   If topmargin = 0 Then topmargin = 10
   If bottommargin = 0 Then bottommargin = 10
   Dev.Line (0, 0)-(CSng(paperwidth), paperheight), QBColor(15), BF
   End If
If First = True And portrait = False Then
   Dev.Cls
   Dev.Visible = True
   Dev.DrawMode = 13
   paperwi = paperheight
   paperhi = paperwidth
   If paperwidth = 0 Then paperwidth = 250
   If paperheight = 0 Then paperheight = 200
   If leftmargin = 0 Then leftmargin = 10
   If rightmargin = 0 Then rightmargin = 10
   If topmargin = 0 Then topmargin = 10
   If bottommargin = 0 Then bottommargin = 10
   Dev.Line (0, 0)-(paperwi, paperhi), QBColor(15), BF
   End If
   
pf100:
   skiya = False
10 filtm2% = FreeFile
 If Abs(nsetflag%) = 1 Or (Abs(nsetflag%) = 3 And skiya = False) Then
     tmpsetflg% = 1
     zmannetz = True
     zmanskiy = False
     If tblmesag% = 1 Then GoTo 20
     Open drivfordtm$ + "netz\netzskiy.tm2" For Input As #filtm2%
     Input #filtm2%, numplac%
     Input #filtm2%, placnam$
     Close #filtm2%
  ElseIf Abs(nsetflag%) = 2 Or (Abs(nsetflag%) = 3 And skiya = True) Then
     tmpsetflg% = 2
     zmanskiy = True
     If nsetflag% <> 3 Then zmannetz = False
     If tblmesag% = 2 Then GoTo 20
     Open drivfordtm$ + "skiy\netzskiy.tm2" For Input As #filtm2%
     Input #filtm2%, numplac%
     Input #filtm2%, placnam$
     Close #filtm2%
     End If
     
If hebcal Then fshabos% = fshabos0%

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-e: newhebcalfm: before if portrait=true"
'Close #lognum%
'End If


'************check for message*********************
20 If portrait = True Then

   'first print header name if abs(nsetflag%)=2 to tblmesag%=3
    'If abs(nsetflag%) = 2 Or tblmesag% = 1 Or tblmesag% = 3 Then
      If PrinterFlag Then Dev.DrawMode = 9
       
      Dev.Font = "David" 'newhebcalfm.Text5.Text
      Dev.FontSize = 20 * rescal 'Val(newhebcalfm.Text13.Text) * rescal * 1.5
      Dev.FontBold = True
      Dev.FontItalic = True

      headertop$ = hebcityname$
     
      Dev.CurrentX = coordxlab1(tmpsetflg%) - Dev.TextWidth(headertop$) / 2
      Dev.CurrentY = 10  'coordylab1(tmpsetflg%) - Dev.TextHeight(headertop$) / 2
      '**********Katz changes*************
      If Katz = True And katznum% = 0 Then
         headertop$ = katzhebnam$
         Dev.FontSize = 16 * rescal 'Val(newhebcalfm.Text13.Text) * rescal * 1.5
         Dev.CurrentX = coordxlab1(tmpsetflg%) - Dev.TextWidth(headertop$) / 2
      ElseIf Katz = True And katznum% >= 1 Then
         katzyo% = katzsep% '***newest
         Dev.FontSize = 16 * rescal 'Val(newhebcalfm.Text13.Text) * rescal * 1.5
         Dev.CurrentY = 10 + katzyo%
         Dev.CurrentX = coordxlab1(tmpsetflg%) - Dev.TextWidth(headertop$) / 2
         End If
      '***********************************
      Dev.Print headertop$
    '  End If

   GoSub messages
   If (tblmesag% = 1 Or tblmesag% = 3) And (Abs(nsetflag%) = 1 Or (Abs(nsetflag%) = 3 And skiya = False)) Then
      If hebcal = True Then
         GoTo 375
      Else
         GoTo 475
         End If
   ElseIf (tblmesag% = 2 Or tblmesag% = 3) And (Abs(nsetflag%) = 2 Or (Abs(nsetflag%) = 3 And skiya = True)) Then
      If hebcal = True Then
         GoTo 375
      Else
         GoTo 475
         End If
      End If
   End If
'***************************************************

   pos1% = InStr(placnam$, "netz")
   pos2% = InStr(placnam$, "skiy")
   If pos1% <> 0 And pos2% = 0 Then
     CN4 = Mid$(placnam$, pos1% + 5, 4)
   ElseIf pos2% <> 0 And pos1% = 0 Then
     CN4 = Mid$(placnam$, pos2% + 5, 4)
     End If
   'CN4 = Mid$(placnam$, 16, 4) 'OUTPUT FILE NAME ROOT
   
   If automatic = True Then
      If Abs(nsetflag%) = 1 Or Abs(nsetflag%) = 3 And skiya = False Then
         If CN4netz$ = sEmpty Then
            CN4netz$ = currentdir
         Else
            If currentdir = CN4netz$ And Not PDFprinter And Val(Caldirectories.Text2.Text) <> numautolst% + 1 And waiterror = False Then 'error
               GoSub founderror
            Else
               CN4netz$ = currentdir
               waiterror = False
               End If
            End If
      ElseIf Abs(nsetflag%) = 2 Or Abs(nsetflag%) = 3 And skiya = True Then
         If CN4skiy$ = sEmpty Then
            CN4skiy$ = currentdir
         Else
            If currentdir = CN4skiy$ And Not PDFprinter And Val(Caldirectories.Text2.Text) <> numautolst% + 1 And waiterror = False Then  'error
               GoSub founderror
            Else
               CN4skiy$ = currendir
               waiterror = False
               End If
            End If
         End If
      End If
   
If Abs(nsetflag%) = 1 Or (Abs(nsetflag%) = 3 And skiya = False) Then
   direct$ = drivfordtm$ + "netz\"
   place$ = CN4
   If numplac% = 1 And Not internet Then
      ext$ = ".pl1"
      '********Katz changes**********
      If Katz = True And katznum% = 0 Then
         ext$ = ".pl0"
         End If
      '*****************************
   Else
      ext$ = ".com"
      End If
   'plachdr1$ = "לוח לזריחת החמה ל"
   plachdr1$ = heb2$(11)
   setflag% = 0 ' = 0 for sunrise, 1 for sunset, -1 for sunsets in 12 hr clock
   steps = Val(newhebcalfm.Text1.Text)
   accur = Val(newhebcalfm.Text2.Text)
ElseIf Abs(nsetflag%) = 2 Or (Abs(nsetflag%) = 3 And skiya = True) Then
   direct$ = drivfordtm$ + "skiy\"
   place$ = CN4
   If numplac% = 1 And Not internet Then
      ext$ = ".pl1"
      '********Katz changes**********
      If Katz = True And katznum% = 0 Then
         ext$ = ".pl0"
         End If
      '*****************************
   Else
      ext$ = ".com"
      End If
   'plachdr1$ = "לוח לשקיעת החמה ב"
   plachdr1$ = heb2$(12)
   setflag% = -1 ' = 0 for sunrise, 1 for sunset, -1 for sunsets in 12 hr clock
   steps = Val(newhebcalfm.Text31.Text)
   accur = Val(newhebcalfm.Text32.Text)
   End If
   
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-f: newhebcalfm: before calculating moladim"
'Close #lognum%
'End If

   
'plachdr2$ = "לשנת "
plachdr2$ = heb2$(13)
plachdr2$ = yrcal$ + plachdr2$
   rdflag% = 1  '= 0, for no rounding
             '= 1, rounds sunsets to earlier steps sec/ sunrises to later steps sec
   'LOCATE 12, 5: INPUT "Round the times to the which nearest second (e.g., 5, 15, 20, etc)?", crsteps$

'   steps = 5   'for rdflag% = 1, then round sunrise times to nearest steps sec/sunsets to later steps sec
'             '= 2, rounds sunsets to earlier .1 sec/ sunrises to later .1 sec
'             '= 3, like option 2 but outputs numbers in Menat's format, e.g., 4:52:06 = 452.2
'             '(< 0, outputs row of numbers rounded to steps secs as above)
   'LOCATE 13, 5: INPUT "What time cushion to add (+) or subtract (-)", crcus$
   If hebcal = False Then GoTo 400  'goto program tables
'  accur = 15  '= 0 uses the tabulated numbers in the .PLC file
             '<> 0 adds accur seconds to the tabulated numbers in the .PLC file
'  yr% = 5758  'Hebrew year for calculation
   
'hebyr$ = "‡" + Chr$(34) + "™"

'<<<<<<<<<<<<<<<<<<<changes
'If Not hebcal Then GoTo pd1000
'
'difdec% = 0
'stdyy% = 299    'English date paramters of Rosh Hashanoh year 1
''stdyy% = 275   'English date paramters of Rosh Hashanoh 5758
'styr% = -3760
''styr% = 1997
'yl% = 366
''yl% = 365
'rh2% = 2       'Rosh Hashanoh year 1 is on day 2 (Monday)
''rh2% = 5      'Rosh Hashanoh 5758 is on day 5 (Thursday)
'yrr% = 2     '5758 is a regular kesidrah year of 354 days
'leapyear% = 0
'ncal1% = 2: ncal2% = 5: ncal3% = 204  'molad of Tishri 1 year 1-- day;hour;chelakim
''ncal1% = 5: ncal2% = 4: ncal3% = 129  'molad of Tishri 1 5758 day;hour;chelakim
'nt1% = ncal1%: nt2% = ncal2%: nt3% = ncal3%: n1rhoo% = nt1%
'leapyr2% = leapyear%
'n1yreg% = 4: n2yreg% = 8: n3yreg% = 876 'change in molad after 12 lunations of reg. year
'n1ylp% = 5: n2ylp% = 21: n3ylp% = 589 'change in molad after 13 lunations of leap year
'n1mon% = 1: n2mon% = 12: n3mon% = 793 'monthly change in molad after 1 lunation
'n11% = ncal1%: n22% = ncal2%: n33% = ncal3%  'initialize molad
'
''Cls
''chosen year to calculate monthly moladim
'yrstep% = yr% - 1
''yrstep% = yr% - 5758
'
'nyear% = 0: flag% = 0
'For kyr% = 1 To yrstep%
'   nyear% = nyear% + 1
'   nnew% = 1
'   GoSub newdate
'   n1rhooo% = n1rho%
'Next kyr%
''now calculate molad of Tishri 1 of next year in order to
''determine if the desired year is choser,kesidrah, or sholem
'leapyr2% = leapyear%
'nyear% = nyear% + 1
'flag% = 1
'nnew% = 0
'GoSub newdate
''now calculate english date and molad of each rosh chodesh of desired year, yr%
'n1rh% = n1rhoo%: n2rh% = nt2%: n3rh% = nt3%
'constdif% = n1rhooo% - rh2%
'rhday% = rh2%: If rhday% = 0 Then rhday% = 7
'
'GoSub dmh
'
'dyy% = stdyy% '- difdec%
'hdryr$ = "-" + Trim$(Str$(styr%))
'newschulyr% = 0
'GoSub engdate
'mdates$(1, 1) = monthh$(iheb%, 1): mdates$(2, 1) = dates$: mmdate%(1, 1) = dyy%
'If newhebcalfm.Check4.Value = vbChecked Then
'   'calculate dyy% for first shabbos
'   fshabos% = 7 - rhday% + dyy%
'   End If
'
''now calculate other molados and their english date
'endyr% = 12: If leapyear% = 1 Then endyr% = 13
''If magnify = True Then GoTo 250
'For k% = 2 To endyr%
'   n33% = n3rh% + n3mon%
'   cal3 = n33% / 1080
'   ncal3% = CInt((cal3 - Fix(cal3)) * 1080)
'   n22% = n2rh% + n2mon%
'   cal2 = (n22% + Fix(cal3)) / 24
'   ncal2% = CInt((cal2 - Fix(cal2)) * 24)
'   n11% = n1rh% + n1mon%
'   cal1 = (n11% + Fix(cal2)) / 7
'   ncal1% = CInt((cal1 - Fix(cal1)) * 7)
'   n1rh% = ncal1%
'   n2rh% = ncal2%
'   n3rh% = ncal3%
'   GoSub dmh
'   n1day% = n1rh%: If n1day% = 0 Then n1day% = 7
'    If k% = 2 Then
'      dyy% = dyy% + 30
'   ElseIf k% = 3 Then
'      If yrr% <> 3 Then dyy% = dyy% + 29
'      If yrr% = 3 Then dyy% = dyy% + 30
'   ElseIf k% = 4 Then
'      If yrr% = 1 Then dyy% = dyy% + 29
'      If yrr% <> 1 Then dyy% = dyy% + 30
'   ElseIf k% = 5 Then
'      dyy% = dyy% + 29
'   ElseIf k% = 6 Then
'      dyy% = dyy% + 30
'   ElseIf k% >= 7 And leapyear% = 0 Then
'      If k% = 7 Then dyy% = dyy% + 29
'      If k% = 8 Then dyy% = dyy% + 30
'      If k% = 9 Then dyy% = dyy% + 29
'      If k% = 10 Then dyy% = dyy% + 30
'      If k% = 11 Then dyy% = dyy% + 29
'      If k% = 12 Then dyy% = dyy% + 30
'   ElseIf k% >= 7 And leapyear% = 1 Then
'      If k% = 7 Then dyy% = dyy% + 30
'      If k% = 8 Then dyy% = dyy% + 29
'      If k% = 9 Then dyy% = dyy% + 30
'      If k% = 10 Then dyy% = dyy% + 29
'      If k% = 11 Then dyy% = dyy% + 30
'      If k% = 12 Then dyy% = dyy% + 29
'      If k% = 13 Then dyy% = dyy% + 30
'      End If
'   hdryr$ = "-" + Trim$(Str$(styr%))
'   dyy% = dyy% - 1
'   GoSub engdate
'   mdates$(3, k% - 1) = dates$: mmdate%(2, k% - 1) = dyy%
'   dyy% = dyy% + 1
'   GoSub engdate
'   mdates$(1, k%) = monthhh$: mdates$(2, k%) = dates$: mmdate%(1, k%) = dyy%
' Next k%
'250 If styr% = 1997 Then styr% = 1998
'dyy% = dyy% + 28: hdryr$ = "-" + Trim$(Str$(styr%))
'GoSub engdate ': LOCATE 13, 5: Print "end of year: "; dates$
'mdates$(3, endyr%) = dates$: mmdate%(2, endyr%) = dyy%
'
''!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''If endyr% = 13 And internet = True Then
''lognum% = FreeFile
''Open drivjk$ + "calprog.log" For Append As #lognum%
''Print #lognum%, "Step #11-g: newhebcalfm: before reading time files"
''Close #lognum%
''End If
'
'pd1000:
myear% = myear0%
'<<<<<<<<<<<<<end of changes

nadd% = 0: If Abs(setflag%) = 1 Then nadd% = 1
nshabos% = 0
ntotshabos% = 0
NearWarning(0) = False
NearWarning(1) = False
For k% = 1 To 2
   myear% = myear% + (k% - 1)
   myear1% = Abs(myear%)
   If myear1% < 1000 Then
      myear1% = myear1% + 1000
      End If
   name100$ = direct$ + Mid$(place$, 1, 4) + Trim$(Str$(myear1%)) + ext$

   yd% = myear% - 1988
   yl% = 365
   If yd% Mod 4 = 0 Then yl% = 366
   If yd% Mod 4 = 0 And myear% Mod 100 = 0 And myear% Mod 400 <> 0 Then yl% = 365
   
'   If k% = 1 Then yl1% = yl%
'   If k% = 2 Then yl2% = yl%
   filplc% = FreeFile
   Open name100$ For Input As #filplc%
'   leapyr% = 0
'   If yl% = 366 Then leapyr% = 1 'leap years
'   If k% = 1 And lenyr1% = 0 Then
'      lenyr1% = 365 + leapyr%  'length of first of the two years
'      End If
   lenyr1% = yrend%(0)
   'If magnify = True Then GoTo 380
'   For i% = 1 To 365 + leapyr%
   For i% = yrstrt%(k% - 1) To yrend%(k% - 1)
      Input #filplc%, tims$
      caldate$ = Mid$(tims$, 1, 11)
      dobs = Val(Mid$(tims$, Len(tims$) - 5, 6))
      
      
      '////////////////////////EK 062022 Arctic Circle handling///////////////////////////////////////
      'look for error flags in the line
      If InStr(tims$, ">max azi") Or InStr(tims$, "55:00:00") Then
         tims$ = "?" 'since there might be a visible sunrise or sunset, just don't have all the info to calculate it
         dobs = 999 'dobs unknown, so print it in black
         
         azim% = 0
         If tmpsetflg% = 1 And nearnez = True Then  'test for near obstructions
           If dobs <= distlim Then
              azim% = 1
              NearWarning(0) = True
              End If
         ElseIf tmpsetflg% = 2 And nearski = True Then
           If dobs <= distlim Then
              azim% = 1
              NearWarning(1) = True
              End If
            End If
        
         GoTo 365
         End If
      If InStr(tims$, "NoSunris") Or _
         InStr(tims$, "NoSunset") Or _
         InStr(tims$, "**:00:00") Or _
         InStr(tims$, "99:00:00") Then
         'perpetual day/night
         If optionheb Then
            tims$ = heb2$(17)
         Else
            tims$ = "none"
            End If
         dobs = 999 'dobs unknown, so print it in black
         azim% = 0
         If tmpsetflg% = 1 And nearnez = True Then  'test for near obstructions
           If dobs <= distlim Then
              azim% = 1
              NearWarning(0) = True
              End If
         ElseIf tmpsetflg% = 2 And nearski = True Then
           If dobs <= distlim Then
              azim% = 1
              NearWarning(1) = True
              End If
            End If
         
         GoTo 365
         End If
      '//////////////////////////////////////////////////////////////
         
      nadd1% = nadd%
      posit% = InStr(12, tims$, ":")
      If skiya = False Then
         If posit% > 15 Then nadd1% = 1
         End If
         
      azim% = 0
      If tmpsetflg% = 1 And nearnez = True Then  'test for near obstructions
         If dobs <= distlim Then
            azim% = 1
            NearWarning(0) = True
            End If
      ElseIf tmpsetflg% = 2 And nearski = True Then
         If dobs <= distlim Then
            azim% = 1
            NearWarning(1) = True
            End If
         End If
      
      tims$ = Mid$(tims$, posit% - 1 - nadd1%, 7 + nadd1%)
      If accur <> 0 Then  'convert to fractional hours/ add accur/ convert back
         hrs = Val(Mid$(tims$, 1, 1 + nadd1%))
         Min = Val(Mid$(tims$, 3 + nadd1%, 2))
         Sec = Val(Mid$(tims$, 6 + nadd1%, 2))
         t3sub = hrs + (Min + (Sec + accur) / 60#) / 60#
         t3hr = Fix(t3sub): t3min = Fix((t3sub - t3hr) * 60#)
         t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'         If setflag% = 0 Then
'            t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'         ElseIf Abs(setflag%) = 1 Then
'            t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'            End If
         If t3sec = 60 Then t3min = t3min + 1: t3sec = 0
         If t3min = 60 Then t3hr = t3hr + 1: t3min = 0
         tshr$ = Trim$(Str$(t3hr))
         tsmin$ = Trim$(Str$(t3min))
         If Len(tsmin$) = 1 Then tsmin$ = "0" + tsmin$
         tssec$ = Trim$(Str$(t3sec))
         If Len(tssec$) = 1 Then tssec$ = "0" + tssec$
         tims$ = tshr$ + ":" + tsmin$ + ":" + tssec$
      Else
         tims$ = Trim$(tims$)
         End If
      If Abs(rdflag%) = 2 Or Abs(rdflag%) = 3 Then
         t3hr = Val(Mid$(tims$, 1, 1 + nadd1%))
         t3min = Val(Mid$(tims$, 3 + nadd1%, 2))
         t3sec = Val(Mid$(tims$, 6 + nadd1%, 2))
         t3sec = t3sec / 60#
         If setflag% = 0 Then
            t3sec = CInt((t3sec + 0.04) * 10) / 10#
         ElseIf Abs(setflag%) = 1 Then
            t3sec = CInt((t3sec - 0.04) * 10) / 10#
            If Abs(rdflag%) = 3 Then t3hr = t3hr - 12
            End If
         If t3sec = 1 Then
            t3min = t3min + 1
            t3sec = 0
            End If
         If t3min = 60 Then
            t3hr = t3hr + 1
            t3min = 0
            t3sec = 0
            End If
        
         t3hrr$ = Trim$(Str$(t3hr))
         t3minn$ = Trim$(Str$(t3min))
         lent3min% = Len(t3minn$)
         If lent3min% = 1 Then t3minn$ = "0" + LTrim$(t3minn$)
         If Abs(rdflag%) = 2 Then
            t3sec = 60 * t3sec
            t3secc$ = Trim$(Str$(t3sec))
            If t3sec < 10 Then t3secc$ = "0" + t3secc$
            tims$ = t3hrr$ + ":" + t3minn$ + ":" + t3secc$
         ElseIf Abs(rdflag%) = 3 Then
            t3secc$ = Trim$(Str$(t3sec))
            If t3sec = 0 Then t3secc$ = ".0"
            tims$ = " " + t3hrr$ + t3minn$ + t3secc$ + "  "
            End If
         End If
      If Abs(rdflag%) = 1 Then GoSub round
      If setflag% = -1 Then '***** new changes (check round and nadd1%)
         If Mid$(tims$, 1, 2) < 22 Then
            Mid$(tims$, 1, 2) = " " + Trim$(Str$(Val(Mid$(tims$, 1, 2)) - 12))
         Else
            Mid$(tims$, 1, 2) = Trim$(Str$(Val(Mid$(tims$, 1, 2)) - 12))
            End If
         End If
      'If setflag% = -1 Then Mid$(tims$, 1, 2) = " " + LTrim$(RTrim$(Str$(Val(Mid$(tims$, 1, 2)) - 12)))
365:
      tim$(k% - 1, i%) = tims$
      tim$(k% + 1, i%) = caldate$
      If azim% = 1 Then
         If nearcolor = True Then
            tim$(k% - 1, i%) = "@" + tims$
         Else
            tim$(k% - 1, i%) = "#:##:##"
            End If
         End If
   
   Next i%
380   Close #filplc%
Next k%


'now generate table
'For i% = 1 To ntable%
'   tbl1$(i%) = String$(160, " ")
'   k% = i%
'   cha$ = "  "
'Next i%


'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If endyr% = 13 And internet = True Then
lognum% = FreeFile
Open drivjk$ + "calprog.log" For Append As #lognum%
Print #lognum%, "********DEBUG VERSION 2.0************"
Print #lognum%, "Step #11-h: newhebcalfm: before mainfont sub"
Close #lognum%
End If

nx% = 0: If setflag% = -1 Then nx% = -1
yrn% = 1

If PrinterFlag Then
   Dev.DrawMode = 13
   If paperwidth = 0 Then paperwidth = 200
   If paperheight = 0 Then paperheight = 250
   If leftmargin = 0 Then leftmargin = 10
   If rightmargin = 0 Then rightmargin = 10
   If topmargin = 0 Then topmargin = 10
   If bottommargin = 0 Then bottommargin = 10
   End If

'*************set fonts for main text**********************
GoSub mainfont
'************************************************************

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-i: newhebcalfm: after mainfont sub, before writing tables"
'Close #lognum%

If PrinterFlag Then Dev.DrawWidth = 1
nshabos% = 0: newschulyr% = 0: dayweek% = rhday% - 1: addmon% = 0
ntotshabos% = 0
For i% = 1 To endyr%

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If i% = endyr% Then
'  cc = 1
'  End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-i1: newhebcalfm: inside writing loop, i%= " & Str(i%)
'Close #lognum%
'End If

   If mmdate%(2, i%) > mmdate%(1, i%) Then
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-i2: newhebcalfm: inside writing loop, mmdate%(2, i%) > mmdate%(1, i%)"
'Close #lognum%
'End If
   
      If i% > 1 And mmdate%(1, i%) = 1 Then
         'Rosh Chodesh is on January 1, so increment the year
         '(this happens in the year 2006, 2036)
         yrn% = yrn% + 1
         End If
         
      stortim$(tmpsetflg% - 1, i% - 1, 29) = sEmpty
      k% = 0
      For j% = mmdate%(1, i%) To mmdate%(2, i%)
          k% = k% + 1
          tims$ = Trim$(tim$(yrn% - 1, j%))
            
          'If endyr% = 12 Then 'regular year
             Dev.CurrentX = coordxreg(tmpsetflg%, i%)
             Dev.CurrentY = coordy(tmpsetflg%, k%)
          'ElseIf endyr% = 13 Then 'leap year
          '   Dev.CurrentX = coordxleap(tmpsetflg%, i%)
          '   Dev.CurrentY = coordy(tmpsetflg%, k%)
          '   End If
          changit% = 0
          If newhebcalfm.Check4.Value = vbChecked And fshabos% + nshabos% * 7 = j% Then 'this is shabbos
             nshabos% = nshabos% + 1 '<<<2--->>>
             ntotshabos% = ntotshabos% + 1
             stortim$(5, i% - 1, k% - 1) = arrStrSedra(0, ntotshabos% - 1)
             stortim$(6, i% - 1, k% - 1) = arrStrSedra(1, ntotshabos% - 1)
             'add sedra information
             changit% = 1: Dev.FontUnderline = True
             'check for end of year
             If fshabos% + nshabos% * 7 > lenyr1% Then
                newschulyr% = 1
                fshabos% = 7 - (lenyr1% - (fshabos% + (nshabos% - 1) * 7))
                nshabos% = 0
                End If
          Else
             stortim$(5, i% - 1, k% - 1) = "-----"
             stortim$(6, i% - 1, k% - 1) = "-----"
             End If
          intflag% = 0
          If Mid$(tims$, 1, 1) = "@" Then
             intflag% = 1
             tims$ = LTrim$(Mid$(tims$, 2, Len(tims$) - 1))
             frcolor = Dev.ForeColor
             Dev.ForeColor = QBColor(6) '16711935 '4227327
          ElseIf Not PrinterFlag Then
             frcolor = QBColor(0)
             End If
          
          '//////////////added 082921--DST support//////////////
          If DSTadd And tims$ <> "NA" Then
             DSThour = Mid$(tims$, 1, InStr(1, tims$, ":") - 1)
             'add hour for DST
             If j% >= strdaynum(yrn% - 1) And j% < enddaynum(yrn% - 1) Then
                DSThour = DSThour + 1
                Mid$(tims$, 1, InStr(1, tims$, ":") - 1) = Trim$(Str$(DSThour))
                End If
             End If
          '//////////////////////////////////////////////////
             
          Dev.Print tims$
          Dev.ForeColor = frcolor
          stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = tims$
          If intflag% = 1 And Not PrinterFlag Then 'and internet = True Then
             stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = stortim$(tmpsetflg% - 1, i% - 1, k% - 1) & "*"
             End If
          If changit% = 1 And Not PrinterFlag Then 'Dev.FontUnderline = True Then 'and internet = True
             stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = stortim$(tmpsetflg% - 1, i% - 1, k% - 1) & "_"
             End If
          If changit% = 1 Then Dev.FontUnderline = False
          If Not PrinterFlag Then
             stortim$(2, i% - 1, k% - 1) = tim$(yrn% + 1, j%)
             If optionheb = True Then
                Call hebnum(k%, cha$)
             Else
                cha$ = Trim$(Str(k%))
                End If
             stortim$(3, i% - 1, k% - 1) = cha$ + "-" + monthh$(iheb%, i% + addmon%)
             If endyr% = 12 And i% = 6 Then
                 stortim$(3, i% - 1, k% - 1) = cha$ + "-" + monthh$(iheb%, 14)
                addmon% = 1
                End If
             dayweek% = dayweek% + 1
             If dayweek% = 8 Then dayweek% = 1
             Call hebweek(dayweek%, cha$)
             stortim$(4, i% - 1, k% - 1) = cha$
             End If
      Next j%
   ElseIf mmdate%(2, i%) < mmdate%(1, i%) Then
      
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-i3: newhebcalfm: inside writing loop, mmdate%(2, i%) < mmdate%(1, i%)"
'Close #lognum%
'End If

      
      If Not PrinterFlag Then stortim$(tmpsetflg% - 1, i% - 1, 29) = sEmpty
      k% = 0
      For j% = mmdate%(1, i%) To yrend%(0)
          k% = k% + 1
          
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop: newhebcalfm: inside writing loop, j%,k%=" & Str(j%) + " ," & Str(k%)
'Close #lognum%
'End If

          tims$ = Trim$(tim$(yrn% - 1, j%))
          'If endyr% = 12 Then 'regular year
          Dev.CurrentX = coordxreg(tmpsetflg%, i%)
          Dev.CurrentY = coordy(tmpsetflg%, k%)
          'ElseIf endyr% = 13 Then 'leap year
          '   Dev.CurrentX = coordxleap(tmpsetflg%, i%)
          '   Dev.CurrentY = coordy(tmpsetflg%, k%)
          '   End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop1: newhebcalfm: inside writing loop"
'Close #lognum%
'End If
          changit% = 0
          If newhebcalfm.Check4.Value = vbChecked And fshabos% + nshabos% * 7 = j% Then 'this is shabbos
             nshabos% = nshabos% + 1
             ntotshabos% = ntotshabos% + 1
             stortim$(5, i% - 1, k% - 1) = arrStrSedra(0, ntotshabos% - 1)
             stortim$(6, i% - 1, k% - 1) = arrStrSedra(1, ntotshabos% - 1)
             changit% = 1: Dev.FontUnderline = True
             'check for end of year
             If fshabos% + nshabos% * 7 > lenyr1% Then
                newschulyr% = 1
                fshabos% = 7 - (lenyr1% - (fshabos% + (nshabos% - 1) * 7))
                nshabos% = 0
                End If
          ElseIf Not PrinterFlag Then
             stortim$(5, i% - 1, k% - 1) = "-----"
             stortim$(6, i% - 1, k% - 1) = "-----"
             End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop2: newhebcalfm: inside writing loop"
'Close #lognum%
'End If
          
          intflag% = 0
          If Mid$(tims$, 1, 1) = "@" Then
             intflag% = 1
             tims$ = LTrim$(Mid$(tims$, 2, Len(tims$) - 1))
             frcolor = Dev.ForeColor
             Dev.ForeColor = QBColor(6) '16711935 '4227327
          ElseIf Not PrinterFlag Then
             frcolor = QBColor(0)
             End If
          
          '//////////////added 082921--DST support//////////////
          If DSTadd And tims$ <> "NA" Then
             DSThour = Mid$(tims$, 1, InStr(1, tims$, ":") - 1)
             'add hour for DST
             If j% >= strdaynum(yrn% - 1) And j% < enddaynum(yrn% - 1) Then
                DSThour = DSThour + 1
                Mid$(tims$, 1, InStr(1, tims$, ":") - 1) = Trim$(Str$(DSThour))
                End If
             End If
          '//////////////////////////////////////////////////
             
          Dev.Print tims$
          Dev.ForeColor = frcolor
          
          If Not PrinterFlag Then
            stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = tims$
          
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop3: newhebcalfm: inside writing loop"
'Close #lognum%
'End If
          
            If intflag% = 1 Then 'and internet = True Then
               stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = stortim$(tmpsetflg% - 1, i% - 1, k% - 1) & "*"
               End If
          'If internet = True And Dev.FontUnderline = True Then

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop4: newhebcalfm: inside writing loop"
'Close #lognum%
'End If
          
            If Dev.FontUnderline = True Then 'and internet = True
               stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = stortim$(tmpsetflg% - 1, i% - 1, k% - 1) & "_"
               End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop6: newhebcalfm: inside writing loop"
'Close #lognum%
'End If
          End If
          
          If changit% = 1 Then Dev.FontUnderline = False
          
          If Not PrinterFlag Then
            stortim$(2, i% - 1, k% - 1) = tim$(yrn% + 1, j%)
            If optionheb = True Then
               Call hebnum(k%, cha$)
            Else
               cha$ = Trim$(Str(k%))
               End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop7: newhebcalfm: inside writing loop"
'Close #lognum%
'End If
          
            stortim$(3, i% - 1, k% - 1) = cha$ + "-" + monthh$(iheb%, i% + addmon%)

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop8: newhebcalfm: inside writing loop"
'Close #lognum%
'End If
          
            If endyr% = 12 And i% = 6 Then
               stortim$(3, i% - 1, k% - 1) = Trim$(cha$) + "-" + monthh$(iheb%, 14)
               addmon% = 1
               End If
            dayweek% = dayweek% + 1
            If dayweek% = 8 Then dayweek% = 1
            Call hebweek(dayweek%, cha$)
            stortim$(4, i% - 1, k% - 1) = cha$

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9: newhebcalfm: inside writing loop"
'Close #lognum%
'End If

       End If

'!!!!!!!!!!!!!!!!!!!!!!!!
'If j% = 365 And k% = 26 And i% = 4 Then
'   cc = 1
'   End If
          
      Next j%
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-j0: newhebcalfm: left loop"
'Close #lognum%
'End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!
If endyr% = 13 And internet = True Then
On Error Resume Next
End If
      
      textwi = Dev.TextWidth(tims$)
      texthi = Dev.TextHeight(tims$)

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9-1, newhebcalfm"
'Close #lognum%
'End If

      yrn% = yrn% + 1
      For j% = 1 To mmdate%(2, i%)
          k% = k% + 1

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9-2, newhebcalfm, j%,k%=" + Str(j%) + " ," + Str(k%)
'Close #lognum%
'End If

          tims$ = Trim$(tim$(yrn% - 1, j%))
          'If endyr% = 12 Then 'regular year
          Dev.CurrentX = coordxreg(tmpsetflg%, i%)
          Dev.CurrentY = coordy(tmpsetflg%, k%)
          'ElseIf endyr% = 13 Then 'leap year
          '   Dev.CurrentX = coordxleap(tmpsetflg%, i%)
          '   Dev.CurrentY = coordy(tmpsetflg%, k%)
          '   End If
          If Not PrinterFlag Then changit% = 0

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9-3, newhebcalfm"
'Close #lognum%
'End If

          If newhebcalfm.Check4.Value = vbChecked And fshabos% + nshabos% * 7 = j% Then 'this is shabbos
             nshabos% = nshabos% + 1  '<<<--->>>
             ntotshabos% = ntotshabos% + 1
             If Not PrinterFlag Then
                stortim$(5, i% - 1, k% - 1) = arrStrSedra(0, ntotshabos% - 1)
                stortim$(6, i% - 1, k% - 1) = arrStrSedra(1, ntotshabos% - 1)
                End If
             changit% = 1: Dev.FontUnderline = True
             'check for end of year
             If fshabos% + nshabos% * 7 > lenyr1% Then
                newschulyr% = 1
                fshabos% = 7 - (lenyr1% - (fshabos% + (nshabos% - 1) * 7))
                nshabos% = 0
                End If
          ElseIf Not PrinterFlag Then
             stortim$(5, i% - 1, k% - 1) = "-----"
             stortim$(6, i% - 1, k% - 1) = "-----"
             End If
          intflag% = 0

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9-4, newhebcalfm"
'Close #lognum%
'End If

          If Mid$(tims$, 1, 1) = "@" Then
             If Not PrinterFlag Then intflag% = 1
             tims$ = LTrim$(Mid$(tims$, 2, Len(tims$) - 1))
             frcolor = Dev.ForeColor
             Dev.ForeColor = QBColor(6) '16711935 '4227327
          ElseIf Not PrinterFlag Then
             frcolor = QBColor(0)
             End If
          
          '//////////////added 082921--DST support//////////////
          If DSTadd And tims$ <> "NA" Then
             DSThour = Mid$(tims$, 1, InStr(1, tims$, ":") - 1)
             'add hour for DST
             If j% >= strdaynum(yrn% - 1) And j% < enddaynum(yrn% - 1) Then
                DSThour = DSThour + 1
                Mid$(tims$, 1, InStr(1, tims$, ":") - 1) = Trim$(Str$(DSThour))
                End If
             End If
          '//////////////////////////////////////////////////
             
          Dev.Print tims$
          Dev.ForeColor = frcolor
          stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = tims$

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9-5, newhebcalfm"
'Close #lognum%
'End If

       If Not PrinterFlag Then
          If intflag% = 1 Then 'and internet = True Then
             stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = stortim$(tmpsetflg% - 1, i% - 1, k% - 1) & "*"
             End If
          If Dev.FontUnderline = True Then 'and internet = True
          'If internet = True And Dev.FontUnderline = True Then
             stortim$(tmpsetflg% - 1, i% - 1, k% - 1) = stortim$(tmpsetflg% - 1, i% - 1, k% - 1) & "_"
             End If
          If changit% = 1 Then Dev.FontUnderline = False
          stortim$(2, i% - 1, k% - 1) = tim$(yrn% + 1, j%)
          'Call hebnum(k%, cha$)
          If optionheb = True Then
             Call hebnum(k%, cha$)
          Else
             cha$ = Trim$(Str(k%))
             End If
          

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9-6, newhebcalfm"
'Close #lognum%
'End If

          stortim$(3, i% - 1, k% - 1) = cha$ + "-" + monthh$(iheb%, i% + addmon%)
          If endyr% = 12 And i% = 6 Then
             stortim$(3, i% - 1, k% - 1) = cha$ + "-" + monthh$(iheb%, 14)
             addmon% = 1
             End If
          dayweek% = dayweek% + 1
          If dayweek% = 8 Then dayweek% = 1
          Call hebweek(dayweek%, cha$)
          stortim$(4, i% - 1, k% - 1) = cha$
          
          End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-loop9-7, newhebcalfm"
'Close #lognum%
'End If
      Dev.FontUnderline = False '<<<--->>> 'reset underline flag
      changit% = 0

      Next j%
      End If
Next i%
textwi = Dev.TextWidth(tims$)
texthi = Dev.TextHeight(tims$)

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-j: newhebcalfm: after writing tables, before printrest"
'Close #lognum%
'End If


'***************print out months/days/grid lines/etc**********
GoSub printrest
'*************************************************************
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-k: newhebcalfm: after printrest"
'Close #lognum%
'End If


375:
If Abs(nsetflag%) = 3 And skiya = False Then
   'nsetflag% = 2
   skiya = True
   First = False
   GoTo 10
   End If
skiya = False
GoTo 999
'***********************END HEBREW CALENDAR**************************
'____________________________________________________________________
'--------------------------------------------------------------------
400
'version 2:00*******************TABLES.BAS****************************
'Cls
'stortim$(0, 1, 28) = sEmpty 'zero values of Feb. 29,30,31
'stortim$(0, 1, 29) = sEmpty
'stortim$(0, 1, 30) = sEmpty
'stortim$(1, 1, 28) = sEmpty
'stortim$(1, 1, 29) = sEmpty
'stortim$(1, 1, 30) = sEmpty

If Abs(nsetflag%) = 1 Or (Abs(nsetflag%) = 3 And skiya = False) Then
   setflag% = 0
   tmpsetflg% = 1
   stortim$(0, 1, 28) = sEmpty 'zero values of Feb. 29,30,31
   stortim$(0, 1, 29) = sEmpty
   stortim$(0, 1, 30) = sEmpty
ElseIf Abs(nsetflag%) = 2 Or (Abs(nsetflag%) = 3 And skiya = True) Then
   tmpsetflg% = 2
   setflag% = -1 'use 12 hour clock for simplicity
   nadd% = 1
   stortim$(1, 1, 28) = sEmpty 'zero values of Feb. 29,30,31
   stortim$(1, 1, 29) = sEmpty
   stortim$(1, 1, 30) = sEmpty
End If
name100$ = direct$ + Mid$(place$, 1, 4) + Trim$(Str$(Abs(yr%))) + ext$

'setflag% = 0 ' = 0 for sunrise, 1 for sunset, -1 for sunsets in 12 hr clock
'hdryr$ = "1998"
'name100$ = "a:\neighb~1\netz\599v1998.pl1" 'input name.plc file
'name101$ = drivjk$+":\jk\neighbor.tbl" 'output table name
rdflag% = 1 '= 0, for no rounding
'             '= 1, rounds sunsets to earlier steps sec/ sunrises to later steps sec
'             '=-1, ", but gives rows of numbers
'             '= 2, rounds sunsets to earlier .1 sec/ sunrises to later .1 sec
'             '= 3, like option 2 but outputs numbers in Menat's format, e.g., 4:52:06 = 452.2
'             '(< 0, outputs row of numbers rounded to steps secs as above)
'steps = 5   'for rdflag% = 1, then round sunrise times to nearest steps sec/sunsets to later steps sec
'accur = 0  '= 0 uses the tabulated numbers in the .PLC file
'             '<> 0 adds accur seconds to the tabulated numbers in the .PLC file
'
yd = yr% - 1988
yl = 365
If yr% Mod 4 = 0 Then yl = 366
If ((yr% Mod 4 = 0) And (yr% Mod 100 = 0) And (yr% Mod 400 <> 0)) Then yl = 365
ntable% = 31
'Dim tbl1$(ntable%)
'Dim header$(2, 5)
'Dim months$(12)
   If Abs(setflag%) = 1 Then nadd% = 1
'   hdrneigh$ = LTrim$(RTrim$(place$))
'   If setflag% = 0 Then
'      header$(1, 2) = String$(52, " ") + hdryr$ + " ™ €… ‰˜„ ’ „€˜„ „‡„ ‡‰˜† ‰† ™ ‡…"
'   ElseIf Abs(setflag%) = 1 Then
'      If placeflg% = 1 Then
'         header$(1, 2) = "  " + hdryr$ + " ™ " + hdrneigh$ + "  …‹™ „ƒ…„‰ ‰˜„ ’ „€˜„ „‡„ ’‰—™ ‰† ™ ‡…"
'      ElseIf placeflg% = 0 Then
'         header$(1, 2) = String$(52, " ") + hdryr$ + " ™ " + "„ƒ…„‰ ‰˜„ ’ „€˜„ „‡„ ’‰—™ ‰† ™ ‡…"
'         End If
'      End If
'   header$(1, 4) = " …‰      –ƒ         …         ˆ—…€        ˆ"‘         ‚…€         ‰…‰        ‰…‰        ‰€         ‰˜"€        ‘˜         ˜"         €‰      …‰ "
'   header$(2, 2) = header$(1, 2)
'   months$(1) = hdryr$ + "-˜€…‰"
'   months$(2) = hdryr$ + "-˜€…˜"""
'   months$(3) = hdryr$ + "-•˜"
'   months$(4) = hdryr$ + "-‰˜"€"
'   months$(5) = hdryr$ + "-‰€"
'   months$(6) = hdryr$ + "-‰…‰"
'   months$(7) = hdryr$ + "-‰…‰"
'   months$(8) = hdryr$ + "-ˆ‘…‚…€"
'   months$(9) = hdryr$ + "-˜ˆ"‘"
'   months$(10) = hdryr$ + "-˜…ˆ—…€"
'   months$(11) = hdryr$ + "-˜…"
'   months$(12) = hdryr$ + "-˜–ƒ"
   leapyr% = 0
   If yl = 366 Then leapyr% = 1 'leap years


'    For i% = 1 To ntable%
'       tbl1$(i%) = String$(160, " ")
'       Mid$(tbl1$(i%), 152, 3) = Str$(i%)
'       'tbl2$(i%) = STRING$(80, " ")
'       Mid$(tbl1$(i%), 2, 3) = Str$(i%)
'    Next i%
    On Error GoTo tableerror  '<<<<<<<<<<??????????
tab5: fil100% = FreeFile
    Open name100$ For Input As #fil100%
'    For i% = 1 To 365 + leapyr%
    For i% = yrstrt%(0) To yrend%(0)
       Input #fil100%, doclin$

       dat$ = RTrim$(Mid$(doclin$, 1, 11))
       nadd1% = nadd%
       posit% = InStr(1, doclin$, ":")
       If skiya = False Then
          If posit% > 15 Then nadd1% = 1
          End If
          
       dobs = Val(Mid$(doclin$, Len(doclin$) - 5, 6))
       
       '///////////////////EK 062022 Arctic Circle handling///////////////////////////////////
       
       If InStr(doclin$, ">max azi") Or InStr(doclin$, "55:00:00") Then
          timss$ = "?"
          dobs = 999 'dobs unknown, so print it in black
          azim% = 0
          If tmpsetflg% = 1 And nearnez = True Then  'test for near obstructions
             If dobs <= distlim Then azim% = 1
          ElseIf tmpsetflg% = 2 And nearski = True Then
             If dobs <= distlim Then azim% = 1
             End If
          GoTo 440
       ElseIf InStr(doclin$, "NoSunris") Or _
          InStr(doclin$, "NoSunset") Or _
          InStr(doclin$, "**:00:00") Or _
          InStr(doclin$, "99:00:00") Then
          If optionheb Then
             timss$ = heb2$(17)
          Else
             timss$ = "none"
             End If
          dobs = 999 'dobs unknown so print it in black
          azim% = 0
          If tmpsetflg% = 1 And nearnez = True Then  'test for near obstructions
             If dobs <= distlim Then azim% = 1
          ElseIf tmpsetflg% = 2 And nearski = True Then
             If dobs <= distlim Then azim% = 1
             End If
          GoTo 440 'skip all the time manipulation
       Else
          timss$ = Mid$(doclin$, posit% - 1 - nadd1%, 7 + nadd1%)
          End If
     '///////////////////////////////////////////////////////////
       
      azim% = 0
      If tmpsetflg% = 1 And nearnez = True Then  'test for near obstructions
         If dobs <= distlim Then azim% = 1
      ElseIf tmpsetflg% = 2 And nearski = True Then
         If dobs <= distlim Then azim% = 1
         End If
       
       'LOCATE 17, 5: PRINT "                                 "
       'LOCATE 17, 5: PRINT tim$
       If accur <> 0 Then  'convert to fractional hours/ add accur/ convert back
          hrs = Val(Mid$(timss$, 1, 1 + nadd1%))
          Min = Val(Mid$(timss$, 3 + nadd1%, 2))
          Sec = Val(Mid$(timss$, 6 + nadd1%, 2))
          '**********Katz change**********
          If Katz = True Then 'always add/subtract 5 minutes
             rdflag% = 1
             steps = 60 'round to latter/earlier minute
             If skiya = False Then
                accur = 300
             Else
                accur = -300
                End If
             End If
          '***************************
          t3sub = hrs + (Min + (Sec + accur) / 60#) / 60#
          t3hr = Fix(t3sub): t3min = Fix((t3sub - t3hr) * 60)
          'If skiya = True And t3hr >= 22 Then
          '   cc = 1
          '   End If
          t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'          If setflag% = 0 Then
'             t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'          ElseIf Abs(setflag%) = 1 Then
'             t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'             End If
          If t3sec = 60 Then t3min = t3min + 1: t3sec = 0
          If t3min = 60 Then t3hr = t3hr + 1: t3min = 0
          tshr$ = Trim$(Str$(t3hr))
          tsmin$ = Trim$(Str$(t3min))
          If Len(tsmin$) = 1 Then tsmin$ = "0" + tsmin$
          tssec$ = Trim$(Str$(t3sec))
          If Len(tssec$) = 1 Then tssec$ = "0" + tssec$
          timss$ = tshr$ + ":" + tsmin$ + ":" + tssec$
       Else
          timss$ = Trim$(timss$)
          End If
       If Abs(rdflag%) = 2 Or Abs(rdflag%) = 3 Then
          t3hr = Val(Mid$(timss$, 1, 1 + nadd1%))
          t3min = Val(Mid$(timss$, 3 + nadd1%, 2))
          t3sec = Val(Mid$(timss$, 6 + nadd1%, 2))
          t3sec = t3sec / 60
          If setflag% = 0 Then
             t3sec = CInt((t3sec + 0.04) * 10) / 10#
          ElseIf Abs(setflag%) = 1 Then
             t3sec = CInt((t3sec - 0.04) * 10) / 10#
             If Abs(rdflag%) = 3 Then t3hr = t3hr - 12
             End If
          If t3sec = 1 Then
             t3min = t3min + 1
             t3sec = 0
             End If
          If t3min = 60 Then
             t3hr = t3hr + 1
             t3min = 0
             t3sec = 0
             End If

          t3hrr$ = Trim$(Str$(t3hr))
          t3minn$ = Trim$(Str$(t3min))
          lent3min% = Len(t3minn$)
          If lent3min% = 1 Then t3minn$ = "0" + LTrim$(t3minn$)
          If Abs(rdflag%) = 2 Then
             t3sec = 60 * t3sec
             t3secc$ = Trim$(Str$(t3sec))
             If t3sec < 10 Then t3secc$ = "0" + t3secc$
             timss$ = t3hrr$ + ":" + t3minn$ + ":" + t3secc$
          ElseIf Abs(rdflag%) = 3 Then
             t3secc$ = Trim$(Str$(t3sec))
             If t3sec = 0 Then t3secc$ = ".0"
             timss$ = " " + t3hrr$ + t3minn$ + t3secc$ + "  "
             End If
          End If
       If Abs(rdflag%) = 1 Then GoSub round
440:
       If setflag% = -1 Then '***** new changes (check round and nadd1%)
          If timss$ = sEmpty Or timss$ = "none" Then GoTo 442
          If Mid$(timss$, 1, 2) < 22 Then
             Mid$(timss$, 1, 2) = " " + Trim$(Str$(Val(Mid$(timss$, 1, 2)) - 12))
          Else
             Mid$(timss$, 1, 2) = Trim$(Str$(Val(Mid$(timss$, 1, 2)) - 12))
             End If
          End If
       If azim% = 1 Then
         If nearcolor = True Then
            timss$ = "@" + timss$
         Else
            timss$ = "#:##:##"
            End If
         End If
442:
         If Not PrinterFlag Then previewfm.Visible = True
'   Next i%
'*************set fonts for main text**********************
GoSub mainfont
'************************************************************


450    nx% = 0: If setflag% = -1 Then nx% = -1
       timss$ = Trim$(timss$)
       If Abs(rdflag%) = 3 Then timss$ = " " + timss$
       lnt% = Len(timss$)
       If Mid$(dat$, 1, 3) = "Jan" Then
          j% = 1: k% = i%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Feb" Then
          j% = 2: k% = i% - 31
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Mar" Then
          j% = 3: k% = i% - 59 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Apr" Then
          j% = 4: k% = i% - 90 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "May" Then
          j% = 5: k% = i% - 120 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Jun" Then
          j% = 6: k% = i% - 151 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Jul" Then
          j% = 7: k% = i% - 181 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Aug" Then
          j% = 8: k% = i% - 212 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Sep" Then
          j% = 9: k% = i% - 243 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Oct" Then
          j% = 10: k% = i% - 273 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Nov" Then
          j% = 11: k% = i% - 304 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
       ElseIf Mid$(dat$, 1, 3) = "Dec" Then
          j% = 12: k% = i% - 334 - leapyr%
          If k% = 1 Then KK% = 0 '*****Katz change*****
          End If
       Dev.CurrentX = coordxreg(tmpsetflg%, j%)
       Dev.CurrentY = coordy(tmpsetflg%, k%)
       If Mid$(timss$, 1, 1) = "@" Then
          timss$ = LTrim$(Mid$(timss$, 2, Len(timss$) - 1))
          frcolor = Dev.ForeColor
          Dev.ForeColor = QBColor(6) '16711935 '4227327
       Else
          frcolor = QBColor(0)
          End If
       '**********Katz change***********
       If Katz = True Then
          If katznum% >= 1 Then katzyo% = katzsep%
          If k% Mod 2 <> 0 Then
             KK% = KK% + 1
             Dev.CurrentY = coordy(tmpsetflg%, KK%) + katzyo%
             Dev.Print timss$
             Dev.ForeColor = frcolor
             End If
       Else
          Dev.Print timss$
          Dev.ForeColor = frcolor
          End If
       stortim$(tmpsetflg% - 1, j% - 1, k% - 1) = timss$
470 Next i%
textwi = Dev.TextWidth(timss$)
If textwi < 12 Then 'if last entry is empty, e.g., for Arctic regions for the sunset and using a civil calendar
   textwi = 12 'need to redefine the text width for the non-empty entries
   End If
texthi = Dev.TextHeight(timss$)
Close #fil100%

'******************print headers,grids,days,etc.*******************
GoSub printrest
'******************************************************************

475:
    If Abs(nsetflag%) = 3 And skiya = False Then
      '*******Katz change*******
      If Katz = True And katznum% >= 1 Then GoTo 485
      '***********************
      'nsetflag% = 2
      skiya = True
      setflag% = -1
      nadd% = 1
      GoTo 10 '400
      End If
    '******Katz changes**********
485: katzyo% = katzsep% * 100
    If Katz = True And skiya = True And katznum% = 0 Then
       katznum% = 1
       skiya = False
       setflag% = 0
       nadd% = 0
       GoTo 10
    ElseIf Katz = True And katznum% = 1 And skiya = False Then
       skiya = True
       katznum% = 2
       setflag% = -1
       nadd% = 1
       GoTo 10
       End If
    '**************************
    skiya = False
    katznum% = 0
    nadd% = 0
'    If rdflag% < 0 Then
'       Close #2
'       GoTo 999
'       End If
'    Open name101$ For Output As #1
'    Print #1, " "
'    For i% = 1 To 5
'       Print #1, header$(1, i%)
'    Next i%
'    For i% = 1 To ntable%
'       Print #1, tbl1$(i%)
'    Next i%
'    Close #1
'    GoTo 999
'
'round:
'           sp% = 0: If Abs(setflag%) = 1 Then sp% = 1
'           t3sub$ = tim$
'           secmin = Val(Mid$(t3sub$, 6 + sp%, 2))
'           minad = 0
'           If secmin Mod steps = 0 Then GoTo rnd50
'           If setflag% = 0 Then 'round up
'              ssec = secmin / steps
'              secmins = CInt(Fix((secmin / steps) * 10) / 10 + 0.499999)
'              If ssec - Fix(ssec) + 0.000001 < 0.1 Then secmins = secmins + 1
'              secmin = steps * secmins
'              If secmin = 60 Then
'                 secmin = 0
'                 minad = 1 - sp%
'                 End If
'              'IF secmin > 0 AND secmin <= 15 THEN
'              '   secmin = 15 * ABS(sp% - 1)
'              'ELSEIF secmin > 15 AND secmin <= 30 THEN
'              '   secmin = 15 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin > 30 AND secmin <= 45 THEN
'              '   secmin = 30 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin > 45 THEN
'              '   secmin = 45 * sp%
'              '   minad = 1 - sp%
'              '   END IF
'           ElseIf Abs(setflag%) = 1 Then 'round down
'              secmin = steps * (Int(Fix((secmin / steps) * 10) / 10))
'              'IF secmin >= 0 AND secmin < 15 THEN
'              '   secmin = 15 * ABS(sp% - 1)
'              'ELSEIF secmin >= 15 AND secmin < 30 THEN
'              '   secmin = 15 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin >= 30 AND secmin < 45 THEN
'              '   secmin = 30 + 15 * ABS(sp% - 1)
'              'ELSEIF secmin >= 45 THEN
'              '   secmin = 45 * sp%
'              '   minad = 1 - sp%
'              '   END IF
'              End If
''rnd50:     If secmin <> 0 Then
'              If secmin < 10 Then
'                 Mid$(t3sub$, 6 + sp%, 1) = "0"
'                 Mid$(t3sub$, 7 + sp%, 1) = LTrim$(RTrim$(Str$(secmin)))
'              Else
'                 Mid$(t3sub$, 6 + sp%, 2) = LTrim$(RTrim$(Str$(secmin)))
'                 End If
'           Else
'              Mid$(t3sub$, 6 + sp%, 2) = "00"
'              End If
'           minmin = Val(Mid$(t3sub$, 3 + sp%, 2)) + minad
'           If minmin = 60 Then
'              Mid$(t3sub$, 1 + sp%, 1) = LTrim$(RTrim$(Str$(Val(Mid$(t3sub$, 1, 1)) + 1)))
'              Mid$(t3sub$, 3 + sp%, 2) = "00"
'           Else
'              If minmin < 10 Then
'                 Mid$(t3sub$, 3 + sp%, 1) = "0"
'                 Mid$(t3sub$, 4 + sp%, 1) = LTrim$(RTrim$(Str$(minmin)))
'              Else
'                 Mid$(t3sub$, 3 + sp%, 2) = LTrim$(RTrim$(Str$(minmin)))
'                 End If
'              End If
'           'IF setflag% = -1 THEN MID$(t3sub$, 1, 2) = " " + LTRIM$(RTRIM$(STR$(VAL(MID$(t3sub$, 1, 2)) - 12)))
'           tim$ = t3sub$
'Return
'
'
'999:  End
'

999 If notprinterflag Then magnify = False
    GoTo 9999
newheberror:
   If Err.Number >= 52 And Err.Number <= 63 Then
      If automatic = True Then 'wait another 5 seconds
         newwait = Timer + 5#
         Do While newwait > Timer
            DoEvents
         Loop
         waiterror = True
         GoTo nheb5
         End If
      If PrinterFlag And internet Then
         'exit program with error message
         Close
         myfile = Dir(drivfordtm$ + "busy.cal")
         If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
           
         lognum% = FreeFile
         Open drivjk$ + "calprog.log" For Append As #lognum%
         Print #lognum%, "Fatal error in previewfm: couldn't find the *.com or *.pl1 file."
         Close #lognum%
         errlog% = FreeFile
         Open drivjk$ + "Cal_PrOK.log" For Output As errlog%
         Print #errlog%, "Cal Prog exited from Previewfm OK button with runtime error message " + Str(Err.Number)
         Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
         Close #errlog%
         Close
      
       'unload forms
        For i% = 0 To Forms.Count - 1
          Unload Forms(i%)
        Next i%
      
        'kill the timer
        If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
        'bring program to abrupt end
        End
        End If
      
      
      MsgBox "Newhebcalfm requests that you PLEASE wait until the NETZSKI3 button on the Menu bar is cleared.", vbInformation, "Cal Program"
      waiterror = True
      GoTo nheb5
   Else
      On Error GoTo nhe50
      If Not PDFprinter Then
        response = MsgBox("Newhebcalfm encountered unexpected error number" + Str$(Err.Number) & vbCrLf & _
                           Err.Description & vbCrLf & vbCrLf & _
                           "Do you want to abort the program?", vbYesNoCancel + vbCritical, "Cal Program")
        If response <> vbYes Then
           Close
           Exit Sub
           End If
      Else
         Close
         Exit Sub
         End If
         
      Close
      If Dir(drivfordtm$ + "netz\*.*") <> sEmpty Then Kill drivfordtm$ + "netz\*.*"
      If Dir(drivfordtm$ + "skiy\*.*") <> sEmpty Then Kill drivfordtm$ + "skiy\*.*"
      If Dir(drivcities$ + "ast\netz\*.*") <> sEmpty Then Kill drivcities$ + "ast\netz\*.*"
      If Dir(drivcities$ + "ast\skiy\*.*") <> sEmpty Then Kill drivcities$ + "ast\skiy\*.*"
      If Dir(drivjk$ + "netzskiy.*") <> sEmpty Then Kill drivjk$ + "netzskiy.*"
     For i% = 0 To Forms.Count - 1
         Unload Forms(i%)
      Next i%
      End If
nhe50:
    End
tableerror:
   Screen.MousePointer = vbDefault
'   Resume
   If Err.Number >= 52 And Err.Number <= 63 Then
      MsgBox "Newhebcalfm requests that you PLEASE wait until the NETZSKI3 button on the Menu bar is cleared.", vbInformation, "Cal Program"
      waiterror = True
      Close
      GoTo tab5
   Else
      On Error GoTo tble50
      response = MsgBox("Newhebcalfm encountered unexpected error number" + Str$(Err.Number) & vbCrLf & _
                         Err.Description & vbCrLf & vbCrLf & _
                         "Abort the program?", vbCritical + vbYesNoCancel, "Cal Program")
      If response <> vbYes Then
         Close
         Exit Sub
         End If
      Close
      If Dir(drivfordtm$ + "netz\*.*") <> sEmpty Then Kill drivfordtm$ + "netz\*.*"
      If Dir(drivfordtm$ + "skiy\*.*") <> sEmpty Then Kill drivfordtm$ + "skiy\*.*"
      If Dir(drivcities$ + "ast\netz\*.*") <> sEmpty Then Kill drivcities$ + "ast\netz\*.*"
      If Dir(drivcities$ + "ast\skiy\*.*") <> sEmpty Then Kill drivcities$ + "ast\skiy\*.*"
      If Dir(drivjk$ + "netzskiy.*") <> sEmpty Then Kill drivjk$ + "netzskiy.*"
      For i% = 0 To Forms.Count - 1
         Unload Forms(i%)
      Next i%
      End If
tble50:
      End
      
founderror:
      response = MsgBox("Newhebcalfm may have detected that Cal Progarm is not advancing." & vbLf & _
                        "It is suspected that there is an error in one of the directory's hebrew names." & vbLf & _
                        "Do you want to abort automatic operation on the next iteration?", vbYesNo + vbQuestion, "Cal Program")
      If response = vbYes Then
         autocancel = True
         End If
Return
      
'newdate:
'    If nyear% = 20 Then nyear% = 1
'    Select Case nyear%
'       Case 3, 6, 8, 11, 14, 17, 19
'          leapyear% = 1
'          n111% = n1ylp%: n222% = n2ylp%: n333% = n3ylp%
'       Case Else
'          leapyear% = 0
'          n111% = n1yreg%: n222% = n2yreg%: n333% = n3yreg%
'    End Select
'    n33% = n33% + n333%
'    cal3 = n33% / 1080
'    ncal3% = CInt((cal3 - Fix(cal3)) * 1080)
'    n22% = n22% + n222%
'    cal2 = (n22% + Fix(cal3)) / 24
'    ncal2% = CInt((cal2 - Fix(cal2)) * 24)
'    n11% = n11% + n111%
'    cal1 = (n11% + Fix(cal2)) / 7
'    ncal1% = CInt((cal1 - Fix(cal1)) * 7)
'    n11% = ncal1%
'    n22% = ncal2%
'    n33% = ncal3%   'molad of Tishri 1 of this iteration
'    'now use dechiyos to determine which day of week Rosh Hashanoh falls on
'    n1rh% = n11%: n2rh% = n22%: n3rh% = n33%
'    'difd% = 0
'    n1rho% = n1rh%
'    Select Case nyear%
'       Case 3, 6, 8, 11, 14, 17, 19
'          If n11% = 2 And n22% + n33% / 1080 > 15.545 Then
'             n1rh% = n1rh% + 1
'             'difd% = 1
'             GoTo 500
'             End If
'    End Select
'    If n2rh% >= 18 Then
'       n1rh% = n1rh% + 1
'       'difd% = 1
'       If n1rh% = 8 Then n1rh% = 1
'       If n1rh% = 1 Or n1rh% = 4 Or n1rh% = 6 Then
'          n1rh% = n1rh% + 1
'          End If
'       GoTo 500
'       End If
'    If n1rh% = 1 Or n1rh% = 4 Or n1rh% = 6 Then
'       n1rh% = n1rh% + 1
'       'difd% = 1
'       End If
'    'GOTO 500
'
''    IF (leapyear% <> 1 AND flag% = 0) AND n1rh% = 3 AND n2rh% + n3rh% / 1080 > 9.188 THEN
''       n1rh% = 5
''       END IF
'    If (flag% = 0 Or flag% = 1 Or kyr% = yrstep%) And n1rh% = 3 And n2rh% + n3rh% / 1080 > 9.188 Then
'       Select Case nyear% + 1
'          Case 20, 2, 4, 5, 7, 9, 12, 13, 15, 16, 18
'             n1rh% = 5
'       End Select
'       End If
'    If n1rh% = 0 Then n1rh% = 7
'
'500    If rh2% >= n1rh% Then difrh% = 7 - rh2% + n1rh%
'       If rh2% < n1rh% Then difrh% = n1rh% - rh2%
'       If nnew% = 1 Then n1rhoo% = n1rh%
'       If (leapyear% = 0 And difrh% = 3) Or (leapyear% = 1 And difrh% = 5) Then
'          yrr% = 1
'       ElseIf (leapyear% = 0 And difrh% = 4) Or (leapyear% = 1 And difrh% = 6) Then
'          yrr% = 2
'       ElseIf (leapyear% = 0 And difrh% = 5) Or (leapyear% = 1 And difrh% = 7) Then
'          yrr% = 3
'          End If
'       If leapyear% = 0 Then
'          If yrr% = 1 Then difdyy% = 353
'          If yrr% = 2 Then difdyy% = 354
'          If yrr% = 3 Then difdyy% = 355
'       ElseIf leapyear% = 1 Then
'          If yrr% = 1 Then difdyy% = 383
'          If yrr% = 2 Then difdyy% = 384
'          If yrr% = 3 Then difdyy% = 385
'          'INPUT "cr", crr$
'          End If
'       If flag% <> 1 Then
'          dyy% = stdyy% + difdyy% - yl% '- difdec%
'          stdyy% = dyy%
'          styr% = styr% + 1
'          yd% = styr% - 1988
'          yl% = 365
'          'LOCATE 18, 1: PRINT "styr%,dy;difdyy%="; styr%; dyy%; difdyy%
'          'LOCATE 19, 1: PRINT "rh2%,n1rh%;leapyear;leapyr2%="; rh2%; n1rh%; leapyear%; leapyr2%
'          'LOCATE 20, 1: PRINT "leapyear%;leapyr2%;year%;kyr%"; leapyear%; leapyr2%; kyr% - 1 + 5758
'          'INPUT "cr", crr$
'          If yd% Mod 4 = 0 Then yl% = 366
'          If yd% Mod 4 = 0 And styr% Mod 100 = 0 And styr% Mod 400 <> 0 Then yl% = 365
'          rh2% = n1rh% 'rh2% is the day of the week (1-7) of the Rosh Hashonoh
'          leapyr2% = leapyear%
'          nt1% = n11%
'          nt2% = n22%
'          nt3% = n33%
'          End If
'Return
'
'engdate:
'   newyear% = 0
'   ydeng% = styr% - 1988
'   yreng% = styr%
'   yleng% = 365
'   If ydeng% Mod 4 = 0 Then yleng% = 366
'   If ydeng% Mod 4 = 0 And yreng% Mod 100 = 0 And yreng% Mod 400 <> 0 Then yleng% = 365
'   If dyy% > yleng% Or newschulyr% = 1 Then
'      newschulyr% = 0
'      myear% = yreng%
'      yreng% = yreng% + 1
'      hdryr$ = "-" + Trim$(Str$(yreng%))
'      dyy% = dyy% - yleng%
'      ydeng% = Val(hdryr$) - 1988
'      yleng% = 365
'      If ydeng% Mod 4 = 0 Then yleng% = 366
'      If ydeng% Mod 4 = 0 And yreng% Mod 100 = 0 And yreng% Mod 400 <> 0 Then yleng% = 365
'      styr% = yreng%: yl% = yleng%
'      newyear% = 1
'      End If
'   leapyr% = 0
'   If yl% = 366 Then leapyr% = 1 'leap years
'   If dyy% >= 1 And dyy% < 32 Then dates$ = monthe$(1) + Trim$(Str$(dyy%)) + hdryr$
'   If dyy% >= 32 And dyy% < 60 + leapyr% Then dates$ = monthe$(2) + Trim$(Str$(dyy% - 31)) + hdryr$
'   If dyy% >= 60 + leapyr% And dyy% < 91 + leapyr% Then dates$ = monthe$(3) + Trim$(Str$(dyy% - 59 - leapyr%)) + hdryr$
'   If dyy% >= 91 + leapyr% And dyy% < 121 + leapyr% Then dates$ = monthe$(4) + Trim$(Str$(dyy% - 90 - leapyr%)) + hdryr$
'   If dyy% >= 121 + leapyr% And dyy% < 152 + leapyr% Then dates$ = monthe$(5) + Trim$(Str$(dyy% - 120 - leapyr%)) + hdryr$
'   If dyy% >= 152 + leapyr% And dyy% < 182 + leapyr% Then dates$ = monthe$(6) + Trim$(Str$(dyy% - 151 - leapyr%)) + hdryr$
'   If dyy% >= 182 + leapyr% And dyy% < 213 + leapyr% Then dates$ = monthe$(7) + Trim$(Str$(dyy% - 181 - leapyr%)) + hdryr$
'   If dyy% >= 213 + leapyr% And dyy% < 244 + leapyr% Then dates$ = monthe$(8) + Trim$(Str$(dyy% - 212 - leapyr%)) + hdryr$
'   If dyy% >= 244 + leapyr% And dyy% < 274 + leapyr% Then dates$ = monthe$(9) + Trim$(Str$(dyy% - 243 - leapyr%)) + hdryr$
'   If dyy% >= 274 + leapyr% And dyy% < 305 + leapyr% Then dates$ = monthe$(10) + Trim$(Str$(dyy% - 273 - leapyr%)) + hdryr$
'   If dyy% >= 305 + leapyr% And dyy% < 335 + leapyr% Then dates$ = monthe$(11) + Trim$(Str$(dyy% - 304 - leapyr%)) + hdryr$
'   If dyy% >= 335 + leapyr% And dyy% < 365 + leapyr% Then dates$ = monthe$(12) + Trim$(Str$(dyy% - 334 - leapyr%)) + hdryr$
'   'IF newyear% = 1 AND yl% = 366 THEN dyy% = dyy% - 1
'Return
'
'dmh:
'   Hourr = n2rh% + 6
'   If Hourr < 12 Then
'      tm$ = " PM night"
'   ElseIf Hourr >= 12 Then
'      Hourr = Hourr - 12
'      If Hourr < 12 Then
'         If Hourr = 0 Then Hourr = 12
'         tm$ = " AM"
'      ElseIf Hourr >= 12 Then
'         Hourr = Hourr - 12
'         If Hourr = 0 Then Hourr = 12
'         tm$ = " PM afternoon"
'         End If
'      End If
'   minc = n3rh% * (60 / 1080)
'   Min = Fix(minc)
'   hel = CInt((minc - Min) * 18)
'Return

round:
           sp% = 0: If Abs(setflag%) = 1 Then sp% = 1
           'If nadd1% = 1 Then sp% = 1
           If hebcal = True Then
             t3subb$ = tims$
           Else
              t3subb$ = timss$
              End If
           If Len(RTrim$(LTrim$(t3subb$))) > 7 Then sp% = 1
           secmin = Val(Mid$(t3subb$, 6 + sp%, 2))
           minad = 0
           If secmin Mod steps = 0 Then GoTo rnd50
           If setflag% = 0 Then 'round up
              ssec = secmin / steps
              secmins = CInt(Fix((secmin / steps) * 10) / 10 + 0.499999)
              If ssec - Fix(ssec) + 0.000001 < 0.1 Then secmins = secmins + 1
              secmin = steps * secmins
              If secmin = 60 Then
                 secmin = 0
                 If skiya = True Then '***changes
                    minad = 1 - sp%
                 Else
                    minad = 1
                    End If
                 End If
           ElseIf Abs(setflag%) = 1 Then 'round down
              secmin = steps * (Int(Fix((secmin / steps) * 10) / 10))
              End If
rnd50:     If secmin <> 0 Then
              If secmin < 10 Then
                 Mid$(t3subb$, 6 + sp%, 1) = "0"
                 Mid$(t3subb$, 7 + sp%, 1) = Trim$(Str$(secmin))
              Else
                 Mid$(t3subb$, 6 + sp%, 2) = Trim$(Str$(secmin))
                 End If
           Else
              Mid$(t3subb$, 6 + sp%, 2) = "00"
              End If
           minmin = Val(Mid$(t3subb$, 3 + sp%, 2)) + minad
           If minmin = 60 Then
              If Len(Trim$(t3subb$)) = 8 Then sp% = 1  '***changes
              If skiya = False And sp% = 0 Then
                 If Trim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1)) >= 10 Then
                    chtmp$ = Trim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1))
                    sp% = 1
                    Mid$(t3subb$, 1, 2) = chtmp$
                    Mid$(t3subb$, 3, 6) = ":00:0"
                    t3subb$ = t3subb$ + "0"
                 Else
                    Mid$(t3subb$, 1 + sp%, 1) = Trim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1))
                    End If
              Else
                 'Mid$(t3subb$, 1 + sp%, 1) = LTrim$(RTrim$(Str$(Val(Mid$(t3subb$, 1 + sp%, 1)) + 1)))
                 newhour% = Val(Mid$(t3subb$, 1, 1 + sp%)) + 1
                 t3subb$ = Trim$(Str$(Val(newhour%))) & Mid$(t3subb$, 2 + sp%, 6)
                 End If
              Mid$(t3subb$, 3 + sp%, 2) = "00"
           Else
              If minmin < 10 Then
                 Mid$(t3subb$, 3 + sp%, 1) = "0"
                 Mid$(t3subb$, 4 + sp%, 1) = Trim$(Str$(minmin))
              Else
                 Mid$(t3subb$, 3 + sp%, 2) = Trim$(Str$(minmin))
                 End If
              End If
           If hebcal = True Then
              tims$ = t3subb$
           Else
              timss$ = t3subb$
              '********Katz changes*********
              If Katz = True Then
                 timss$ = sEmpty
                 timss$ = Mid$(t3subb$, 1, Len(t3subb$) - 3)
                 End If
              '**********end changes************
              End If
Return
  
'hebnum:
'   If k% <= 10 Then
'      cha$ = LTrim$(RTrim$((Chr$(k% + 223)))) + " "
'   ElseIf k% > 10 And k% < 20 Then
'      cha$ = LTrim$(RTrim$((Chr$(233)))) + LTrim$(RTrim$(Chr$(k% - 10 + 223)))
'   ElseIf k% = 20 Then
'      cha$ = LTrim$(RTrim$(Chr$(235))) + " "
'   ElseIf k% > 20 And k% < 30 Then
'      cha$ = LTrim$(RTrim$(Chr$(235))) + LTrim$(RTrim$(Chr$(k% - 20 + 223)))
'   ElseIf k% = 30 Then
'      cha$ = LTrim$(RTrim$(Chr$(236))) + " "
'      End If
'   If k% = 15 Then cha$ = LTrim$(RTrim$(Chr$(232))) + LTrim$(RTrim$(Chr$(229)))
'   If k% = 16 Then cha$ = LTrim$(RTrim$(Chr$(232))) + LTrim$(RTrim$(Chr$(230)))
'Return

schultim:
         hrs = Val(Mid$(timnez$, 1, 1 + nadd1%))
         Min = Val(Mid$(timnez$, 3 + nadd1%, 2))
         Sec = Val(Mid$(timnez$, 6 + nadd1%, 2))
         accur = -13.5 * 60  'schochen ad
         GoSub caltim
         timchonad$ = tims$
         accur = -6.5 * 60!   'schma
         GoSub caltim
         timschma$ = tims$
         accur = -2! * 60!    'emes
         GoSub caltim
         timemes$ = tims$
         accur = -0.25 * 60!  'goal israel
         GoSub caltim
         timgoal$ = tims$
Return

caltim:
            t3sub = hrs + (Min + (Sec + accur) / 60#) / 60#
            t3hr = Fix(t3sub): t3min = Fix((t3sub - t3hr) * 60)
            t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'            If setflag% = 0 Then
'               t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'            ElseIf Abs(setflag%) = 1 Then
'               t3sec = Int((t3sub - t3hr - t3min / 60#) * 3600 + 0.5)
'               End If
            If t3sec = 60 Then t3min = t3min + 1: t3sec = 0
            If t3min = 60 Then t3hr = t3hr + 1: t3min = 0
            tshr$ = Trim$(Str$(t3hr))
            tsmin$ = Trim$(Str$(t3min))
            If Len(tsmin$) = 1 Then tsmin$ = "0" + tsmin$
            tssec$ = Trim$(Str$(t3sec))
            If Len(tssec$) = 1 Then tssec$ = "0" + tssec$
            tims$ = tshr$ + ":" + tsmin$ + ":" + tssec$
Return

mainfont:
    Dev.Font = newhebcalfm.Text3.Text
    Dev.FontSize = Val(newhebcalfm.Text11.Text) * rescal
    If newhebcalfm.Text7.Text = "Bold" Or newhebcalfm.Text7.Text = "Bold Italic" Then
       Dev.FontBold = True
    Else
       Dev.FontBold = False
       End If
    If newhebcalfm.Text7.Text = "Italic" Or newhebcalfm.Text7.Text = "Bold Italic" Then
       Dev.FontItalic = True
    Else
       Dev.FontItalic = False
       End If
Return

printrest:
'***********************set fonts for months and days***************
Dev.Font = newhebcalfm.Text4.Text
Dev.FontSize = Val(newhebcalfm.Text12.Text) * rescal
If newhebcalfm.Text8.Text = "Bold" Or newhebcalfm.Text8.Text = "Bold Italic" Then
   Dev.FontBold = True
Else
   Dev.FontBold = False
   End If
If newhebcalfm.Text8.Text = "Italic" Or newhebcalfm.Text8.Text = "Bold Italic" Then
   Dev.FontItalic = True
Else
   Dev.FontItalic = False
   End If
'*******************************************************************
If hebcal = True Then
    If endyr% = 12 Then
       For i% = 1 To endyr%
          If i% < 6 Then
            k% = i%
          ElseIf i% = 6 Then
            k% = 14
          ElseIf i% > 6 Then
            k% = i% + 1
            End If
          Dev.CurrentX = coordxmonreg(i%) + textwi * 0.5 - Dev.TextWidth(monthh$(iheb%, k%)) / 2
          Dev.CurrentY = coordymon(tmpsetflg%) - dey(tmpsetflg%) * conv '* (newhebcalfm.Texthi / textwi) ' * conv '- de(1)
          Dev.Print monthh$(iheb%, k%)
          stormon$(i% - 1) = monthh$(iheb%, k%)
        Next i%
    ElseIf endyr% = 13 Then
        For i% = 1 To endyr%
          k% = i%
          Dev.CurrentX = coordxmonreg(i%) + textwi * 0.5 - Dev.TextWidth(monthh$(iheb%, k%)) / 2
          Dev.CurrentY = coordymon(tmpsetflg%) - dey(tmpsetflg%) * conv '* (newhebcalfm.Texthi / textwi) ' * conv '- de(1)
          Dev.Print monthh$(iheb%, k%)
          stormon$(i% - 1) = monthh$(iheb%, k%)
        Next i%
        End If
ElseIf hebcal = False Then
    For i% = 1 To 12
       Dev.CurrentX = coordxmonreg(i%) + textwi * 0.5 - Dev.TextWidth(montheh$(iheb%, i%)) / 2
       Dev.CurrentY = coordymon(tmpsetflg%) - dey(tmpsetflg%) * conv
       '*******Katz change********
       If Katz = True And katznum% >= 1 Then
          Dev.CurrentY = coordymon(tmpsetflg%) - dey(tmpsetflg%) * conv + katzsep%
          End If
       '***************************
       Dev.Print montheh$(iheb%, i%)
       If Not PrinterFlag Then stormon$(i% - 1) = montheh$(iheb%, i%)
    Next i%
 End If
 '*********changes for Katz*********
 If Katz = True Then
    stepday = 2
 Else
    stepday = 1
    End If
 KK% = 0
 '******************
 For k% = 1 To ntable% Step stepday 'print out days *****stepday = Katz change
    If hebcal = True Then
       'Call hebnum(k%, cha$)
       If optionheb = True Then
          Call hebnum(k%, cha$)
       Else
          cha$ = Trim$(Str(k%))
          End If
       'GoSub hebnum
       If k% = 21 Then
          textwi21 = Dev.TextWidth(cha$)
          End If
       If k% = 10 Then
          textwi10 = Dev.TextWidth(cha$)
          End If
     ElseIf hebcal = False Then
        cha$ = CStr(k%)
        If k% = 1 Then
           textwi10 = Dev.TextWidth(cha$)
           End If
        If k% = 25 Then
           textwi21 = Dev.TextWidth(cha$)
           End If
        End If
        
    '*********Katz changes****************
    If Katz = True And k% Mod 2 <> 0 Then
       KK% = KK% + 1
    Else
       KK% = k%
       End If
    Dev.CurrentY = (yo + ys(tmpsetflg%) + yot + (KK% - 1) * dy) * conv '20 + (k% - 1) * 2.12
    
    If Katz = True And katznum% >= 1 Then
       katzyo% = katzsep% * 100
       Dev.CurrentY = (yo + katzyo% + ys(tmpsetflg%) + yot + (KK% - 1) * dy) * conv '20 + (k% - 1) * 2.12
       End If
    '*****************************
    Dev.CurrentX = coordxreg(1, endyr%) - de(tmpsetflg%) * conv - Dev.TextWidth(cha$) / 2 '+ 3.1 '24 - Dev.TextWidth(cha$) / 2 '- numshift%
    If optionheb = False Then
       Dev.CurrentX = coordxreg(1, 1) - de(tmpsetflg%) * conv - Dev.TextWidth(cha$) / 2 '+ 3.1 '24 - Dev.TextWidth(cha$) / 2 '- numshift%
       End If
    Dev.Print cha$
    Dev.CurrentY = (yo + yot + ys(tmpsetflg%) + (KK% - 1) * dy) * conv '20 + (kk% - 1) * 2.12
    '*********Katz changes********
    If Katz = True And katznum% >= 1 Then
       katzyo% = katzsep% * 100
       Dev.CurrentY = (yo + katzyo% + yot + ys(tmpsetflg%) + (KK% - 1) * dy) * conv '20 + (kk% - 1) * 2.12
       End If
    '*****************************
    Dev.CurrentX = coordxreg(1, 1) + textwi + de(tmpsetflg%) * conv - Dev.TextWidth(cha$) / 2 '- 3.1 - numshift%'152 - Dev.TextWidth(cha$) / 2 '- numshift%
    If optionheb = False Then
       Dev.CurrentX = coordxreg(1, endyr%) + textwi + de(tmpsetflg%) * conv - Dev.TextWidth(cha$) / 2 '- 3.1 - numshift%'152 - Dev.TextWidth(cha$) / 2 '- numshift%
       End If
    Dev.Print cha$ '" " + cha$
 Next k%

''***********************Print Header*******************************
'If portrait = True And skiya = False Then
'   Dev.Font = newhebcalfm.Text5.Text
'   Dev.FontSize = Val(newhebcalfm.Text13.Text) * rescal * 1.5
'   Dev.FontBold = True
'   Dev.FontItalic = True
'
'   headertop$ = hebcityname$
'
'   Dev.CurrentX = coordxlab1(tmpsetflg%) - Dev.TextWidth(headertop$) / 2
'   Dev.CurrentY = 10  'coordylab1(tmpsetflg%) - Dev.TextHeight(headertop$) / 2
'   Dev.Print headertop$
'   End If

'***********************set fonts for Title Header****************
Dev.Font = newhebcalfm.Text5.Text
Dev.FontSize = Val(newhebcalfm.Text13.Text) * rescal
If newhebcalfm.Text9.Text = "Bold" Or newhebcalfm.Text9.Text = "Bold Italic" Then
   Dev.FontBold = True
Else
   Dev.FontBold = False
   End If
If newhebcalfm.Text9.Text = "Italic" Or newhebcalfm.Text9.Text = "Bold Italic" Then
   Dev.FontItalic = True
Else
   Dev.FontItalic = False
   End If
'*******************************************************************


header$(1) = cap1$(tmpsetflg%)
     
Dev.CurrentX = coordxlab1(tmpsetflg%) - Dev.TextWidth(header$(1)) / 2
Dev.CurrentY = coordylab1(tmpsetflg%) - Dev.TextHeight(header$(1)) / 2
'**********Katz change***********
If Katz = True And katznum% >= 1 Then
   Dev.CurrentY = coordylab1(tmpsetflg%) - Dev.TextHeight(header$(1)) / 2 + katzsep%
   End If
'********************************
Dev.Print header$(1)

'***********************set fonts for other captions****************
Dev.Font = newhebcalfm.Text6.Text
Dev.FontSize = Val(newhebcalfm.Text14.Text) * rescal
If newhebcalfm.Text10.Text = "Bold" Or newhebcalfm.Text10.Text = "Bold Italic" Then
   Dev.FontBold = True
Else
   Dev.FontBold = False
   End If
If newhebcalfm.Text10.Text = "Italic" Or newhebcalfm.Text10.Text = "Bold Italic" Then
   Dev.FontItalic = True
Else
   Dev.FontItalic = False
   End If
cirx = coordxlab5(tmpsetflg%) - Dev.TextWidth(cap5$(tmpsetflg%)) / 2
ciry = coordylab5(tmpsetflg%) '+ Dev.TextHeight(cap5$(tmpsetflg%)) / 2
'*******************************************************************


header$(2) = cap2$(tmpsetflg%)
header$(3) = cap3$(tmpsetflg%)
header$(4) = cap4$(tmpsetflg%)
header$(5) = cap5$(tmpsetflg%)
header$(6) = cap6$(tmpsetflg%)

If (tmpsetflg% = 1 And NearWarning(0) = False And AddObsTime = 0) Or _
   (tmpsetflg% = 2 And NearWarning(1) = False And AddObsTime = 0) Then
   header$(6) = sEmpty 'no near obstructions detected
   End If
   
   
For stori% = 0 To 5
   storheader$(tmpsetflg% - 1, stori%) = header$(stori% + 1)
Next stori%
     
Dev.CurrentX = coordxlab2(tmpsetflg%) - Dev.TextWidth(header$(2)) / 2
Dev.CurrentY = coordylab2(tmpsetflg%) - Dev.TextHeight(header$(2)) / 2
Dev.Print header$(2)
Dev.CurrentX = coordxlab3(tmpsetflg%) - Dev.TextWidth(header$(3)) / 2
Dev.CurrentY = coordylab3(tmpsetflg%) - Dev.TextHeight(header$(3)) / 2
Dev.Print header$(3)
Dev.CurrentX = coordxlab4(tmpsetflg%) - Dev.TextWidth(header$(4)) / 2
Dev.CurrentY = coordylab4(tmpsetflg%) - Dev.TextHeight(header$(4)) / 2
Dev.Print header$(4)
Dev.CurrentX = coordxlab5(tmpsetflg%) - Dev.TextWidth(header$(5)) / 2
Dev.CurrentY = coordylab5(tmpsetflg%) - Dev.TextHeight(header$(5)) / 2
Dev.Print header$(5)
If nearcolor = True Or AddObsTime = 1 Then
   If nearnez = True Or nearski = True Or AddObsTime = 1 Then
      Dev.CurrentX = coordxlab6(tmpsetflg%) - Dev.TextWidth(header$(6)) / 2
      Dev.CurrentY = coordylab6(tmpsetflg%) - Dev.TextHeight(header$(6)) / 2
      Dev.Print header$(6)
      End If
   End If

'Now print בס"ד and copyright circle if checked and update table of contents for automatic mode
helpfromshemyim$ = "בס" + Chr$(34) + "ד"
If newhebcalfm.Check1.Value = vbChecked And automatic = False Then
   If portrait = True Then
      Dev.CurrentX = paperwidth - rightmargin - 2 * Dev.TextWidth(helpfromshemyim$)
      Dev.CurrentY = topmargin + 2 * Dev.TextHeight(helpfromshemyim$)
   ElseIf portrait = False Then
      Dev.CurrentX = paperheight - rightmargin - 2 * Dev.TextWidth(helpfromshemyim$)
      Dev.CurrentY = topmargin + 2 * Dev.TextHeight(helpfromshemyim$)
      End If
   Dev.Print helpfromshemyim$
   End If
If automatic = False Then GoTo 9000
If automatic = True And Caldirectories.Check1.Value = vbChecked Then 'paginate
   If Abs(nsetflag%) = 3 And skiya = True And tblmesag% = 0 Then GoTo 9000
   Dev.Font = "David"
   Dev.FontSize = 12 * rescal
   'numautolst% = numautolst% + 1
   autonum% = autonum% + 1
   pagnum% = newpagenum% + autonum% + 1 'Val(Caldirectories.Text2.Text) + numautolst% - 1 - newpagenum%
   If PDFprinter Then pagnum% = numautolst%
   'record in temporary file
   
   numd% = FreeFile
   Open drivjk$ & "numdirec.txt" For Output As #numd%
   Write #numd%, pagnum%
   Close #numd%
   
   'pagnum% = numautolst% - 1 - newpagenum%
   pagnums$ = Trim$(CStr(pagnum%))
   'print page number on automatic mode page

   If SunriseCalc And SunsetCalc Then

        If pagnum% Mod 2 = 0 Then
           Dev.CurrentX = paperwidth - rightmargin - Dev.TextWidth(pagnums$)
        ElseIf pagnum% Mod 2 <> 0 Then
           Dev.CurrentX = leftmargin
           End If
        Dev.CurrentY = topmargin - Dev.TextHeight(pagnums$)
        
   ElseIf (SunriseCalc And Not SunsetCalc) Or (SunsetCalc And Not SunriseCalc) Then
   
        If pagnum% Mod 2 = 0 Then
'           Dev.CurrentX = paperwidth - rightmargin - Dev.TextWidth(pagnums$)
           Dev.CurrentX = paperwidth * 1.3 - Dev.TextWidth(pagnums$) '* 2 - rightmargin - Dev.TextWidth(pagnums$)
        ElseIf pagnum% Mod 2 <> 0 Then
           Dev.CurrentX = leftmargin
           End If
        Dev.CurrentY = topmargin - Dev.TextHeight(pagnums$)
        If Dev.CurrentY < 0 Then Dev.CurrentY = 5
        End If
        
   Dev.Print pagnums$
   'now append to the table of contents file
   filcont% = FreeFile
   myfile = Dir(drivcities$ + "tablcont.txt")
   If myfile = sEmpty Then
      Open drivcities$ + "tablcont.txt" For Output As #filcont%
   Else
      If numautolst% = 1 And Val(Caldirectories.Text2.Text) = 1 Then
         Open drivcities$ + "tablcont.txt" For Output As #filcont%
      Else
         Open drivcities$ + "tablcont.txt" For Append As #filcont%
         End If
      End If
'    Print #filcont%, Tab(1); pagnums$; Tab(20); hebcityname$
'   textlen1 = Abs(Dev.TextWidth(hebcityname$))
'   textlen2 = Abs(Dev.TextWidth("."))
   'textlen3 = Dev.TextWidth(pagnums$)
'   totlen% = 150 '15 centimeters
'   totleft% = totlen% - textlen1
'   mult% = totleft% / textlen2
   'mult% = CInt((textlen1 + textlen3) / textlen2)
'   Print #filcont%, hebcityname$ + String$(mult% - 50, ".") + "," + pagnums$
    Print #filcont%, hebcityname$ + "," + pagnums$
    Close #filcont%
    End If
    
9000:
If newhebcalfm.Check2.Value = vbChecked Then
   'Dev.Font = "Ariel"
   If PrinterFlag Then Dev.DrawMode = 9
   Dev.Font = newhebcalfm.Text6.Text
   Dev.FontSize = Val(newhebcalfm.Text14.Text) * rescal
   'cirx = cirx - 3 * Dev.TextWidth("c")
   ''ciry = ciry + 0.5 * Dev.TextHeight("c") / 2
   'cirx = cirx + 3.2 * Dev.TextWidth("c") / 2
   ''ciry = ciry + 2.9 * Dev.TextWidth("c") / 2
   'Dev.Circle (cirx, ciry), 0.2 * Dev.TextWidth("c"), , 3.14159 / 4, 7.4 * 3.14159 / 4
   'Dev.Circle (cirx, ciry), 0.6 * Dev.TextWidth("c")
   End If

'now draw grid lines
If Not PrinterFlag Then
   Dev.DrawWidth = 1
Else
   Dev.DrawWidth = 2
   End If
If newhebcalfm.Check3.Value = vbChecked Then
   If Not PrinterFlag Then
      Dev.DrawMode = 13
   Else
      Dev.DrawMode = 9
      End If
   For i% = 1 To ntable% + 1
      'horizontal grid for main text
      linx1 = (xo + xot) * conv + textwi + (dx * conv - textwi) / 2
      linx2 = (xo + xot - (endyr% - 1) * dx) * conv - (dx * conv - textwi) / 2 '1 / 14 * textwi
      liny = (yo + yot + ys(tmpsetflg%) + (i% - 1) * dy) * conv
      If PrinterFlag Then Dev.DrawMode = 13
      '************Katz change**********
      If Katz = True And i% <= ntable% / 2 + 1 Then
        If Not PrinterFlag Then
         If katznum% >= 1 Then katzyo% = katzsep% * 100
         If newhebcalfm.Check5.Value = vbChecked Then
           If i% Mod 3 = 0 Then
              Dev.DrawWidth = 2
           Else
              Dev.DrawWidth = 1
              End If
        ElseIf PrinterFlag Then
            myfile = Dir(drivjk$ + "printinfo.sav")
            If myfile <> sEmpty Then
               filtmp% = FreeFile
               Open drivjk$ + "printinfo.sav" For Input As #filtmp%
               Input #filtmp%, drawwidths%
               Close #filtmp%
            Else
               drawwidths% = 9
               End If
            If i% Mod 3 = 0 Then
               Dev.DrawWidth = drawwidths%
            Else
               Dev.DrawWidth = 3
               End If
        
           End If
         End If
         If PrinterFlag Then
            If katznum% >= 1 Then katzyo% = katzsep% * 100
            End If
         liny = (yo + katzyo% + yot + ys(tmpsetflg%) + (i% - 1) * dy) * conv
         If PrinterFlag Then Dev.DrawMode = 13
         Dev.Line (linx1, liny)-(linx2, liny)
         'now horizontal-grid for hebrew numbers
         linxo = coordxreg(1, endyr%) - de(tmpsetflg%) * conv
         If PrinterFlag Then Dev.DrawMode = 13
         Dev.Line (linxo - textwi21 / 2 - textwi10, liny)-(linxo + textwi21 / 2 + textwi10, liny)
         linxo = coordxreg(1, 1) + textwi + de(tmpsetflg%) * conv
         If PrinterFlag Then Dev.DrawMode = 13
         Dev.Line (linxo - textwi21 / 2 - textwi10, liny)-(linxo + textwi21 / 2 + textwi10, liny)
      ElseIf Katz = False Then
        If Not PrinterFlag Then
           Dev.Line (linx1, liny)-(linx2, liny)
        Else
           Dev.Line (linx1, liny)-(linx2, liny), QBColor(0), B
           End If
        
        'now horizontal-grid for hebrew numbers
        linxo = coordxreg(1, endyr%) - de(tmpsetflg%) * conv
        If optionheb = False Then
           linxo = coordxreg(1, 1) - de(tmpsetflg%) * conv
           End If
        Dev.Line (linxo - textwi21 / 2 - textwi10, liny)-(linxo + textwi21 / 2 + textwi10, liny)
        If Not PrinterFlag Then
           Dev.Line (linxo - textwi21 / 2 - textwi10, liny)-(linxo + textwi21 / 2 + textwi10, liny)
        Else
           Dev.DrawMode = 13
           Dev.Line (linxo - textwi21 / 2 - textwi10, liny)-(linxo + textwi21 / 2 + textwi10, liny), QBColor(0), B
           End If
        
        linxo = coordxreg(1, 1) + textwi + de(tmpsetflg%) * conv
        If optionheb = False Then
           linxo = coordxreg(1, endyr%) + textwi + de(tmpsetflg%) * conv
           End If
        If Not PrinterFlag Then
           Dev.Line (linxo - textwi21 / 2 - textwi10, liny)-(linxo + textwi21 / 2 + textwi10, liny)
        Else
           Dev.DrawMode = 13
           Dev.Line (linxo - textwi21 / 2 - textwi10, liny)-(linxo + textwi21 / 2 + textwi10, liny), QBColor(0), B
           End If
        End If
      '*********************************
   Next i%
   For i% = 1 To endyr% - 1
      'vertical-grid for main text
      liny1 = (yo + yot + ys(tmpsetflg%)) * conv
      '***********Katz change***************
      If Katz = False Then
         liny2 = (yo + yot + ys(tmpsetflg%) + ntable% * dy) * conv
      Else
         If katznum% >= 1 Then katzyo% = katzsep% * 100
         liny1 = (yo + katzyo% + yot + ys(tmpsetflg%)) * conv
         liny2 = (yo + katzyo% + yot + ys(tmpsetflg%) + (ntable% / 2 + 1 / 2) * dy) * conv
         End If
      '*****************************
      linx = (xo + xot - (i% - 1) * dx) * conv - (dx * conv - textwi) / 2
      If PrinterFlag Then Dev.DrawMode = 13
      If Not PrinterFlag Then
         Dev.Line (linx, liny1)-(linx, liny2)
      Else
         Dev.Line (linx, liny1)-(linx, liny2), QBColor(0), B
         End If

      'vertical grid for hebrew months
      liny1 = coordymon(tmpsetflg%) - dey(tmpsetflg%) * conv
      liny2 = liny1 + texthi
      '*********Katz change********
      If Katz = True And katznum% >= 1 Then
         liny1 = liny1 + katzsep%
         liny2 = liny2 + katzsep%
         End If
      If PrinterFlag Then Dev.DrawMode = 13
      If Not PrinterFlag Then
         Dev.Line (linx, liny1)-(linx, liny2)
      Else
         Dev.Line (linx, liny1)-(linx, liny2), QBColor(0), B
         End If
   Next i%
   'thick box for main text
   If PrinterFlag Then Dev.DrawMode = 9
   linx1 = (xo + xot) * conv + textwi + (dx * conv - textwi) / 2
   linx2 = (xo + xot - (endyr% - 1) * dx) * conv - (dx * conv - textwi) / 2
   liny1 = (yo + yot + ys(tmpsetflg%)) * conv
   '***********Katz change************************
   If Katz = False Then
      liny2 = (yo + yot + ys(tmpsetflg%) + ntable% * dy) * conv
   Else
      If katznum% >= 1 Then katzyo% = katzsep% * 100
      liny1 = (yo + katzyo% + yot + ys(tmpsetflg%)) * conv
      liny2 = (yo + katzyo% + yot + ys(tmpsetflg%) + (ntable% / 2 + 1 / 2) * dy) * conv
      End If
'gives error in Win 10, so just comment out
'   '************************CHANGE FOR DIFFERENT DRIVERS********
'   'DriverName for HP LaserJet 1100 is HPPTA, HP LaserJet 6L is "HPW", for IIIP and 4P it is "HPPCL5MS"
'   If Printer.DriverName = "HPW" Or Printer.DriverName = "HPPTA" Or Printer.DriverName = "winspool" Then
'      Printer.DrawWidth = 5
'   Else
'      Printer.DrawWidth = 3
'      End If
'   '*******************************************************
   If Not PrinterFlag Then
      If portrait = True Then Dev.DrawWidth = 2
      If portrait = False Then Dev.DrawWidth = 3
      Dev.Line (linx1, liny1)-(linx2, liny2), , B
      End If
   If PrinterFlag Then
      Dev.Line (linx1, liny1)-(linx2, liny2), QBColor(0), B
      Printer.FillStyle = vbFSTransparent
      Printer.DrawStyle = vbSolid
      End If
   'thick box for hebrew numbers
   linxo = coordxreg(1, endyr%) - de(tmpsetflg%) * conv
   If optionheb = False Then
      linxo = coordxreg(1, 1) - de(tmpsetflg%) * conv
      End If
   linx1 = linxo - textwi21 / 2 - textwi10
   linx2 = linxo + textwi21 / 2 + textwi10
   If Not PrinterFlag Then
      Dev.Line (linx1, liny1)-(linx2, liny2), , B
   Else
      Dev.Line (linx1, liny1)-(linx2, liny2), QBColor(0), B
      End If
   linxo = coordxreg(1, 1) + textwi + de(tmpsetflg%) * conv
   If optionheb = False Then
      linxo = coordxreg(1, endyr%) + textwi + de(tmpsetflg%) * conv
      End If
   linx1 = linxo - textwi21 / 2 - textwi10
   linx2 = linxo + textwi21 / 2 + textwi10
   If Not PrinterFlag Then
      Dev.Line (linx1, liny1)-(linx2, liny2), , B
   Else
      Dev.Line (linx1, liny1)-(linx2, liny2), QBColor(0), B
      End If
   'thick box for hebrew months
   liny1 = coordymon(tmpsetflg%) - dey(tmpsetflg%) * conv
   liny2 = liny1 + texthi
   '********Katz change*********
   If Katz = True And katznum% >= 1 Then
      liny1 = liny1 + katzsep%
      liny2 = liny2 + katzsep%
      End If
   '***************************
   linx1 = (xo + xot) * conv + textwi + (dx * conv - textwi) / 2
   linx2 = (xo + xot - (endyr% - 1) * dx) * conv - (dx * conv - textwi) / 2
   If Not PrinterFlag Then
      Dev.Line (linx1, liny1)-(linx2, liny2), , B
   Else
      Dev.Line (linx1, liny1)-(linx2, liny2), QBColor(0), B
      End If
   
   Dev.DrawWidth = 1
   End If

If PrinterFlag Then Dev.DrawMode = 9
'now put in fill
If newhebcalfm.Check5.Value = vbChecked And Katz = False Then
   Dev.DrawMode = 9
   For i% = 1 To ntable%
      'horizontal grid for main text
      linx1 = (xo + xot) * conv + textwi + (dx * conv - textwi) / 2
      linx2 = (xo + xot - (endyr% - 1) * dx) * conv - (dx * conv - textwi) / 2 '1 / 14 * textwi
      liny1 = (yo + yot + ys(tmpsetflg%) + (i% - 1) * dy) * conv
      liny2 = (yo + yot + ys(tmpsetflg%) + i% * dy) * conv
      If newhebcalfm.Option1.Value = True And (i% = 1 Or i% = 2 Or i% = 3 Or i% = 7 Or i% = 8 Or i% = 9 _
         Or i% = 13 Or i% = 14 Or i% = 15 Or i% = 19 Or i% = 20 Or i% = 21 Or i% = 25 Or i% = 26 Or i% = 27) Then
         Dev.Line (linx1, liny1)-(linx2, liny2), fillcol, BF
      ElseIf newhebcalfm.Option2.Value = True And (i% = 1 Or i% = 2 Or i% = 3 Or i% = 4 Or i% = 9 Or i% = 10 Or i% = 11 Or i% = 12 _
         Or i% = 17 Or i% = 18 Or i% = 19 Or i% = 20 Or i% = 25 Or i% = 26 Or i% = 27 Or i% = 28) Then
         Dev.Line (linx1, liny1)-(linx2, liny2), fillcol, BF
         End If
      'now horizontal-grid for hebrew/english numbers
      linxo = coordxreg(1, endyr%) - de(tmpsetflg%) * conv
      If optionheb = False Then
         linxo = coordxreg(1, 1) - de(tmpsetflg%) * conv
         End If
      linx1 = linxo - textwi21 / 2 - textwi10
      linx2 = linxo + textwi21 / 2 + textwi10
      If newhebcalfm.Option1.Value = True And (i% = 1 Or i% = 2 Or i% = 3 Or i% = 7 Or i% = 8 Or i% = 9 _
         Or i% = 13 Or i% = 14 Or i% = 15 Or i% = 19 Or i% = 20 Or i% = 21 Or i% = 25 Or i% = 26 Or i% = 27) Then
         Dev.Line (linx1, liny1)-(linx2, liny2), fillcol, BF
      ElseIf newhebcalfm.Option2.Value = True And (i% = 1 Or i% = 2 Or i% = 3 Or i% = 4 Or i% = 9 Or i% = 10 Or i% = 11 Or i% = 12 _
         Or i% = 17 Or i% = 18 Or i% = 19 Or i% = 20 Or i% = 25 Or i% = 26 Or i% = 27 Or i% = 28) Then
         Dev.Line (linx1, liny1)-(linx2, liny2), fillcol, BF
         End If
      linxo = coordxreg(1, 1) + textwi + de(tmpsetflg%) * conv
      If optionheb = False Then
         linxo = coordxreg(1, endyr%) + textwi + de(tmpsetflg%) * conv
         End If
      linx1 = linxo - textwi21 / 2 - textwi10
      linx2 = linxo + textwi21 / 2 + textwi10
      If newhebcalfm.Option1.Value = True And (i% = 1 Or i% = 2 Or i% = 3 Or i% = 7 Or i% = 8 Or i% = 9 _
         Or i% = 13 Or i% = 14 Or i% = 15 Or i% = 19 Or i% = 20 Or i% = 21 Or i% = 25 Or i% = 26 Or i% = 27) Then
         Dev.Line (linx1, liny1)-(linx2, liny2), fillcol, BF
      ElseIf newhebcalfm.Option2.Value = True And (i% = 1 Or i% = 2 Or i% = 3 Or i% = 4 Or i% = 9 Or i% = 10 Or i% = 11 Or i% = 12 _
         Or i% = 17 Or i% = 18 Or i% = 19 Or i% = 20 Or i% = 25 Or i% = 26 Or i% = 27 Or i% = 28) Then
         Dev.Line (linx1, liny1)-(linx2, liny2), fillcol, BF
         End If
   Next i%
   End If
Return

messages:
   If Not PrinterFlag Then
      Dev.Font = "David"
      Dev.FontSize = 10 * rescal
      Dev.FontBold = True
      Dev.FontItalic = False
       If (Abs(nsetflag%) = 1 Or (Abs(nsetflag%) = 3 And skiya = False)) And (tblmesag% = 1 Or tblmesag% = 3) Then
          messg$ = ".האופק המזרחי של יישוב זה מוסתר ע" + Chr$(34) + "י הסתרים קרובים ב- % " + _
                   Format(s1blk, "###") + " של השנה"
          ym01 = (yo + ys(1) + yot + 11 * dy) * conv
          xm01 = (xo + xc(1)) * conv - Dev.TextWidth(messg$) / 2
          xm02 = (xo + xc(1)) * conv + Dev.TextWidth(messg$) / 2
          Dev.CurrentX = xm01
          Dev.CurrentY = ym01
          Dev.Print messg$
          messg$ = ".בגלל הסתרים אלו אי " + "אפשר לחשב לוח מדויק לזריחה הנראית מהמודל הטופוגרפי הממוחשב"
    '      messg$ = ".בגלל ההסתרים האלו א" + Chr$(34) + "א לחשב לוח מדויק של הזריחה הנראית מהמודל הטופוגרפי הממוחשב"
          ym1 = (yo + ys(1) + yot + 12 * dy) * conv + 0.5
          xm11 = (xo + xc(1)) * conv - Dev.TextWidth(messg$) / 2
          xm12 = (xo + xc(1)) * conv + Dev.TextWidth(messg$) / 2
          Dev.CurrentY = ym1
          Dev.CurrentX = xm11
          Dev.Print messg$
          messg$ = ".נא לעיין בהקדמת ספר זה"
          xm2 = (xo + xc(1)) * conv - Dev.TextWidth(messg$) / 2
          ym21 = (yo + ys(1) + yot + 13 * dy) * conv + 1 '+ Dev.TextHeight(messg$)
          ym22 = (yo + ys(1) + yot + 14 * dy) * conv + 1.5 '+ 2 * Dev.TextHeight(messg$)
          Dev.CurrentX = xm2
          Dev.CurrentY = ym21
          Dev.Print messg$
          Dev.Line (xm11 - 1.5, ym01 - 1.5)-(xm12 + 1.5, ym22 + 1.5), , B
       ElseIf (Abs(nsetflag%) = 2 Or (Abs(nsetflag%) = 3 And skiya = True)) And (tblmesag% = 2 Or tblmesag% = 3) Then
          messg$ = ".האופק המערבי של יישוב זה מוסתר ע" + Chr$(34) + "י הסתרים קרובים ב- % " + _
                   Format(s2blk, "###") + " של השנה"
          ym01 = (yo + ys(2) + yot + 11 * dy) * conv
          xm01 = (xo + xc(2)) * conv - Dev.TextWidth(messg$) / 2
          xm02 = (xo + xc(2)) * conv + Dev.TextWidth(messg$) / 2
          Dev.CurrentX = xm01
          Dev.CurrentY = ym01
          Dev.Print messg$
          messg$ = ".בגלל הסתרים אלו אי " + "אפשר לחשב לוח מדויק לשקיעה הנראית מהמודל הטופוגרפי הממוחשב"
    '      messg$ = ".בגלל ההסתרים האלו א" + Chr$(34) + "א לחשב לוח מדויק של השקיעה הנראית מהמודל הטופוגרפי הממוחשב"
          ym1 = (yo + ys(2) + yot + 12 * dy) * conv + 0.5
          xm11 = (xo + xc(2)) * conv - Dev.TextWidth(messg$) / 2
          xm12 = (xo + xc(2)) * conv + Dev.TextWidth(messg$) / 2
          Dev.CurrentY = ym1
          Dev.CurrentX = xm11
          Dev.Print messg$
          messg$ = ".נא לעיין בהקדמת ספר זה"
          Dev.CurrentX = (xo + xc(2)) * conv - Dev.TextWidth(messg$) / 2
          ym21 = (yo + ys(2) + yot + 13 * dy) * conv + 1 '+ Dev.TextHeight(messg$)
          ym22 = (yo + ys(2) + yot + 14 * dy) * conv + 1.5 '+ 2 * Dev.TextHeight(messg$)
          Dev.CurrentY = ym21
          Dev.Print messg$
          Dev.Line (xm11 - 1.5, ym01 - 1.5)-(xm12 + 1.5, ym22 + 1.5), , B
          End If
      
   Else
   
      Dev.DrawMode = 9
      Dev.Font = "David"
      Dev.FontSize = 10
      Dev.FontBold = False
      Dev.FontItalic = False
      Dev.DrawStyle = vbSolid
       If (Abs(nsetflag%) = 1 Or (Abs(nsetflag%) = 3 And skiya = False)) And (tblmesag% = 1 Or tblmesag% = 3) Then
          messg$ = ".האופק המזרחי של יישוב זה מוסתר ע" + Chr$(34) + "י הסתרים קרובים ב- % " + _
                   Format(s1blk, "###") + " של השנה"
          ym01 = (yo + ys(1) + yot + 11 * dy) * conv
          xm01 = (xo + xc(1)) * conv - Printer.TextWidth(messg$) / 2
          xm02 = (xo + xc(1)) * conv + Printer.TextWidth(messg$) / 2
          ym1 = (yo + ys(1) + yot + 12 * dy) * conv + 0.5
          messg$ = ".בגלל הסתרים אלו אי " + "אפשר לחשב לוח מדויק לזריחה הנראית מהמודל הטופוגרפי הממוחשב"
          xm11 = (xo + xc(1)) * conv - Printer.TextWidth(messg$) / 2
          xm12 = (xo + xc(1)) * conv + Printer.TextWidth(messg$) / 2
          messg$ = ".נא לעיין בהקדמת ספר זה"
          xm2 = (xo + xc(1)) * conv - Printer.TextWidth(messg$) / 2
          ym21 = (yo + ys(1) + yot + 13 * dy) * conv + 1 '+ printer.TextHeight(messg$)
          ym22 = (yo + ys(1) + yot + 14 * dy) * conv + 1.5 '+ 2 * printer.TextHeight(messg$)
          Printer.DrawMode = 9
          Printer.FillStyle = vbFSTransparent
          Printer.DrawStyle = vbSolid
          Printer.DrawWidth = 2
          Printer.Line (xm11 - 1.5, ym01 - 1.5)-(xm12 + 1.5, ym22 + 1.5), QBColor(0), B
          messg$ = ".האופק המזרחי של יישוב זה מוסתר ע" + Chr$(34) + "י הסתרים קרובים ב- % " + _
                   Format(s1blk, "###") + " של השנה"
          Printer.CurrentX = xm01
          Printer.CurrentY = ym01
          Printer.Print messg$
          messg$ = ".בגלל הסתרים אלו אי " + "אפשר לחשב לוח מדויק לזריחה הנראית מהמודל הטופוגרפי הממוחשב"
    '      messg$ = ".בגלל הסתרים אלו א" + Chr$(34) + "א לחשב לוח מדויק לזריחה הנראית מהמודל הטופוגרפי הממוחשב"
          Printer.CurrentY = ym1
          Printer.CurrentX = xm11
          Printer.Print messg$
          messg$ = ".נא לעיין בהקדמת ספר זה"
          Printer.CurrentX = xm2
          Printer.CurrentY = ym21
          Printer.Print messg$
       ElseIf (Abs(nsetflag%) = 2 Or (Abs(nsetflag%) = 3 And skiya = True)) And (tblmesag% = 2 Or tblmesag% = 3) Then
          messg$ = ".האופק המערבי של יישוב זה מוסתר ע" + Chr$(34) + "י הסתרים קרובים ב- % " + _
                   Format(s2blk, "###") + " של השנה"
          ym01 = (yo + ys(2) + yot + 11 * dy) * conv
          xm01 = (xo + xc(2)) * conv - Printer.TextWidth(messg$) / 2
          xm02 = (xo + xc(2)) * conv + Printer.TextWidth(messg$) / 2
          ym1 = (yo + ys(2) + yot + 12 * dy) * conv + 0.5
          messg$ = ".בגלל הסתרים אלו אי " + "אפשר לחשב לוח מדויק לשקיעה הנראית מהמודל הטופוגרפי הממוחשב"
          xm11 = (xo + xc(2)) * conv - Printer.TextWidth(messg$) / 2
          xm12 = (xo + xc(2)) * conv + Printer.TextWidth(messg$) / 2
          messg$ = ".נא לעיין בהקדמת ספר זה"
          ym21 = (yo + ys(2) + yot + 13 * dy) * conv + 1 '+ printer.TextHeight(messg$)
          ym22 = (yo + ys(2) + yot + 14 * dy) * conv + 1.5 '+ 2 * printer.TextHeight(messg$)
          Printer.DrawMode = 9 '7
          Printer.DrawStyle = vbSolid
          Printer.FillStyle = vbFSTransparent
          Printer.DrawWidth = 2
          'If tblmesag% = 3 Then Printer.DrawMode = 9
          Printer.Line (xm11 - 1.5, ym01 - 1.5)-(xm12 + 1.5, ym22 + 1.5), QBColor(0), B
          messg$ = ".האופק המערבי של יישוב זה מוסתר ע" + Chr$(34) + "י הסתרים קרובים ב- % " + _
                   Format(s2blk, "###") + " של השנה"
          Printer.CurrentX = xm01
          Printer.CurrentY = ym01
          Printer.Print messg$
          messg$ = ".בגלל הסתרים אלו אי " + "אפשר לחשב לוח מדויק לשקיעה הנראית מהמודל הטופוגרפי הממוחשב"
          Printer.CurrentY = ym1
          Printer.CurrentX = xm11
          Printer.Print messg$
          messg$ = ".נא לעיין בהקדמת ספר זה"
          Printer.CurrentX = (xo + xc(2)) * conv - Printer.TextWidth(messg$) / 2
          Printer.CurrentY = ym21
          Printer.Print messg$
          End If
      
      End If
Return

9999 If PrinterFlag Then 'printing
        Dev.EndDoc
        
        If PDFprinter Then
            Rem -- Wait for runonce settings file to disappear
            Dim runonce As String
            runonce = settings.GetSettingsFilePath(True)
            While Dir(runonce, vbNormal) <> ""
                Sleep 100
            Wend
            
            MsgBox "myfile.pdf was saved on your desktop", vbInformation, "PDF Created"
            End If
           
        katznum% = 0
        Exit Sub
        End If

     If automatic = True Then '<<<<<<<<<<<<<autosave>>>>>>>>>>>>>>>>>
        If autoprint Then 'print the tables to the default printer
            previewfm.PreviewOKbut.Value = True
            waittime = Timer + 5#
            Do While waittime > Timer
               DoEvents
            Loop
            previewfm.PreviewExitbut.Value = True
        ElseIf autosave Then 'save the tables as html to save directory and add to html TOC
           previewfm.prevASCIIfilbut.Value = True
            waittime = Timer + 1#
            Do While waittime > Timer
               DoEvents
            Loop
            previewfm.PreviewExitbut.Value = True
           End If
        End If
     If internet = True Then
     
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-l: newhebcalfm: before prevASCIIfilbut.value=true"
'Close #lognum%
'End If
     
     
        previewfm.prevASCIIfilbut.Value = True
        'previewfm.PreviewExitbut.Value = True


'!!!!!!!!!!!!!!!!!!
'If endyr% = 13 And internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-m: newhebcalfm: after prevASCIIfilbut.value=true"
'Close #lognum%
'End If

        'newhebcalfm.newhebExitbut.Value = True
        'Call MDIform_queryunload(0, 0)
        End If
     Screen.MousePointer = vbDefault
     Exit Sub
     
     
generrhand:
     Screen.MousePointer = vbDefault
     If internet = True And Err.Number >= 0 Then 'exit the program
        'abort the program with a error messages
        errlog% = FreeFile
        Open drivjk$ + "Cal_pbgeh.log" For Output As #errlog%
        Print #errlog%, "Cal Prog exited from previewbut with runtime error message " + Str(Err.Number)
        Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
        Close #errlog%
        Close
      
       'unload forms
        For i% = 0 To Forms.Count - 1
          Unload Forms(i%)
        Next i%
      
        myfile = Dir(drivfordtm$ + "busy.cal")
        If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
      
      
        'kill the timer
        If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
        'bring program to abrupt end
        End
     Else
        response = MsgBox("Newhebcalfm previewbut encountered error number: " + Str(Err.Number) + ".  Do you want to abort?", vbYesNoCancel + vbCritical, "Cal Program")
        If response <> vbYes Then
           Close
           Exit Sub
        Else
           Close
           For i% = 0 To Forms.Count - 1
             Unload Forms(i%)
           Next i%
           End
           End If
        End If
     
End Sub
Public Function fileToString(ByVal uri As String) As String
Dim FF As Integer
Dim fileByte() As Byte
FF = FreeFile()
Open (uri) For Binary As #FF
ReDim fileByte(LOF(FF))
Do While Not (EOF(FF))
Get FF, , fileByte
Loop
Close FF
End Function
Function UnicodeToAscii(ByRef pstrUnicode)
     Dim llngLength
     Dim llngIndex
     Dim llngAscii
     Dim lstrAscii
         
     llngLength = Len(pstrUnicode)
         
     For llngIndex = 1 To llngLength
          llngAscii = Asc(Mid(pstrUnicode, llngIndex, 1))
          lstrAscii = lstrUnicode & ChrB(llngAscii)
     Next
         
     UnicodeToAscii = lstrAscii
End Function


Function AsciiToUnicode(ByRef pstrAscii)
         
     Dim llngLength
     Dim llngIndex
     Dim llngAscii
     Dim lstrUnicode
         
     llngLength = LenB(pstrAscii)
         
     For llngIndex = 1 To llngLength
          llngAscii = AscB(MidB(pstrAscii, llngIndex, 1))
          lstrUnicode = lstrUnicode & Chr(llngAscii)
     Next
         
     AsciiToUnicode = lstrUnicode
         
End Function

Public Function PrinterIndex(ByVal printerName As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(printerName) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function


Public Function FNms(X As Double) As Double
    FNms = mp + mc * X
End Function

Public Function FNaas(X As Double) As Double
    FNaas = ap + ac * X
End Function
Public Function FNes(aas As Double) As Double
    FNes = ms + ec * Sin(aas) + e2c * Sin(2 * aas)
End Function
Public Function FNha(X As Double) As Double
    FNha = FNarco((-Tan(lr) * Tan(D)) + (Cos(X) / Cos(lr) / Cos(D))) * ch
End Function
Public Function FNfrsum(X As Double) As Double
    FNfrsum = (P / (t + 273)) * (0.1419 - 0.0073 * X + 0.00005 * X * X) / (1 + 0.3083 * X + 0.01011 * X * X)
End Function
Public Function FNfrwin(X As Double) As Double
    FNfrwin = (P / (t + 273)) * (0.1561 - 0.0082 * X + 0.00006 * X * X) / (1 + 0.3254 * X + 0.01086 * X * X)
End Function
Public Function FNref(X As Double) As Double
    FNref = (P / (t + 273)) * (0.1594 + 0.0196 * X + 0.00002 * X * X) / (1 + 0.505 * X + 0.0845 * X * X)
End Function
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

Dim tncols As Long, IKMY&, IKMX&, numrec&, IO%, Tempmode%

FilePathBil = App.Path & "\WorldClim_bil"
If Dir(FilePathBil, vbDirectory) <> sEmpty Then
    FileNameBil = FilePathBil
Else
    Call MsgBox("Can't find the bil directory at the following location:" _
                & vbCrLf & FilePathBil _
                & vbCrLf & vbCrLf & "Please select the correct direcotry location." _
                , vbExclamation, "Missing bil file directory")
    FileNameBil = BrowseForFolder(Drukfrm.hwnd, "Choose Directory")
    If Dir(FileNameBil, vbDirectory) = sEmpty Then
       ier = -1
       Exit Sub
       End If
    End If
'first extract minimum temperatures

 Tempmode% = 0
T50:
   
 For i = 1 To 12
        
    If Tempmode% = 0 Then 'minimum temperatures to be used for sunrise calculations
       FilePathBil = FileNameBil & "\min_"
    ElseIf Tempmode% = 1 Then 'average temperatures to be used for sunset calculations
       FilePathBil = FileNameBil & "\avg_"
    ElseIf Tempmode% = 2 Then 'average temperatures to be used for sunset calculations
       FilePathBil = FileNameBil & "\max_"
       End If
    
    Select Case i
       Case 1
          FilePathBil = FilePathBil & "Jan"
       Case 2
          FilePathBil = FilePathBil & "Feb"
       Case 3
          FilePathBil = FilePathBil & "Mar"
       Case 4
          FilePathBil = FilePathBil & "Apr"
       Case 5
          FilePathBil = FilePathBil & "May"
       Case 6
          FilePathBil = FilePathBil & "Jun"
       Case 7
          FilePathBil = FilePathBil & "Jul"
       Case 8
          FilePathBil = FilePathBil & "Aug"
       Case 9
          FilePathBil = FilePathBil & "Sep"
       Case 10
          FilePathBil = FilePathBil & "Oct"
       Case 11
          FilePathBil = FilePathBil & "Nov"
       Case 12
          FilePathBil = FilePathBil & "Dec"
    End Select
    FilePathBil = FilePathBil + ".bil"
    
    If Dir(FilePathBil) <> sEmpty Then
       filein% = FreeFile
       Open FilePathBil For Binary As #filein%
   
        Y = lat
        X = lon
        
        IKMY& = CLng((ULYMAP - Y) / YDIM) + 1
        IKMX& = CLng((X - ULXMAP) / XDIM) + 1
        tncols = NCOLS
        numrec& = (IKMY& - 1) * tncols + IKMX&
        Get #filein%, (numrec& - 1) * 2 + 1, IO%
        If IO% = NODATA Then IO% = 0#
        If Tempmode% = 0 Then
            MinTemp(i) = IO%
        ElseIf Tempmode% = 1 Then
            AvgTemp(i) = IO%
        ElseIf Tempmode% = 2 Then
            MaxTemp(i) = IO%
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
  ElseIf Tempmode% = 1 Then
     Tempmode% = 2
     GoTo T50
     End If

End Sub


