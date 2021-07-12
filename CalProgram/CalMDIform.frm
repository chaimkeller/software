VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm CalMDIform 
   BackColor       =   &H8000000C&
   Caption         =   "Calendar Programs"
   ClientHeight    =   3195
   ClientLeft      =   3210
   ClientTop       =   3240
   ClientWidth     =   4680
   Icon            =   "CalMDIform.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "calfm"
            Object.ToolTipText     =   "Cal Program"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Katzform"
            Object.ToolTipText     =   "Program for Akiba Katz"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "worldbut"
            Object.ToolTipText     =   "wrold DTM program"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "astrobut"
            Object.ToolTipText     =   "Calculate astronomical and mishor times for inputed coordinates"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ChaiAirTables"
            Object.ToolTipText     =   "Chai Air Travel Tables"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DSTbut"
            Object.ToolTipText     =   "Set onset and end of DST for EY"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   3240
         TabIndex        =   1
         Top             =   60
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1980
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":169C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":1DD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":1F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalMDIform.frx":228A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu filefm 
      Caption         =   "&File"
      Begin VB.Menu mnuStndRef 
         Caption         =   "Use Standard Refraction"
      End
      Begin VB.Menu mnuVDW 
         Caption         =   "Use VDW raytracing Ref."
         Checked         =   -1  'True
      End
      Begin VB.Menu exitfm 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu viewfm 
      Caption         =   "&Calendars"
      Begin VB.Menu eretzfm 
         Caption         =   "&Eretz Israel"
      End
      Begin VB.Menu katzfm 
         Caption         =   "&Katz tables"
      End
      Begin VB.Menu worldsunfm 
         Caption         =   "&World sunrise tables"
      End
      Begin VB.Menu caltablefm 
         Caption         =   "&Astronomic times "
      End
      Begin VB.Menu mnuChaiAir 
         Caption         =   "&Chai Air Times"
      End
   End
   Begin VB.Menu versionfm 
      Caption         =   "&Version"
   End
   Begin VB.Menu helpfm 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "CalMDIform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub caltablefm_Click()
       eroscityflag = False
       ret = SetWindowPos(CalMDIform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       astronplace = True
       AstronForm.Visible = True
       BringWindowToTop (AstronForm.hwnd)
End Sub

Private Sub eretzfm_Click()
    eroscityflag = False
    Caldirectories.Visible = True
End Sub

Private Sub exitfm_Click()
   Call MDIform_queryunload(i%, j%)
End Sub

Private Sub katzfm_Click()
        eroscityflag = False
        Katz = True
        AstronForm.Visible = True
        katztotal% = 0
        AstronForm.Combo1.ListIndex = 0
        AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
        If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
           AstronForm.Option2.Value = True
        ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
           AstronForm.Option1.Value = True
           End If
End Sub

Private Sub MDIForm_Load()
   'version: 08/22/2003
   
   On Error GoTo generrhand
   
   'constants
   pi = 4 * Atn(1)
   pi2 = 2 * pi
   ch = 360# / (pi2 * 15)  '57.29578 / 15  'conv rad to hr
   cd = pi / 180#  'conv deg to rad
   
'look for PathsZones_dtm.txt file and if exists read in cushions, else set up cushion defaults
myfile = Dir(App.Path & "\PathsZones_dtm.txt")
If myfile <> sEmpty Then
   filin% = FreeFile
   Open App.Path & "\PathsZones_dtm.txt" For Input As #filin%
   Do Until EOF(filin%)
      Line Input #filin%, doclin$
      Select Case doclin$
         Case "[time cushions]"
            Input #filin%, cushion(0), cushion(1), cushion(2), cushion(3), cushion(4)
         Case "[obstruction limit]"
            Input #filin%, obsdistlim(0), obsdistlim(1), obsdistlim(2), obsdistlim(3)
            Exit Do
      End Select
   Loop
   Close #filin%
Else
   cushion(0) = 15 'EY JKH DTM
   cushion(1) = 45 'GTOPO30
   cushion(2) = 35 'SRTM 90 m
   cushion(3) = 20 'SRTM 30 m, NED 1 arcsec DTM
   cushion(4) = 15 'astronomical calculations
   obsdistlim(0) = 5 'EY JKH DTM
   obsdistlim(1) = 30 'GTOPO30
   obsdistlim(2) = 10 'SRTM 90 m
   obsdistlim(3) = 6 'SRTM 30 m, NED 1 arcsec DTM (30 m)
   End If
   
   If Not internet Then 'allow user to calculate tables for entire Hebrew calendar
      RefHebYear% = 1
      RefCivilYear% = -3760
   ElseIf internet Then 'speed up calculation by using more recent reference Hebrew year
      RefHebYear% = 5758
      RefCivilYear% = 1997
      End If
     
   'find location of default drives
   Screen.MousePointer = vbHourglass
   numdriv% = Drive1.ListCount
   driveletters$ = "cdefghijklmnop"
   s1% = 0: S2% = 0: S3% = 0
   For i% = 1 To numdriv%
'   For i% = 4 To numdriv%
      drivlet$ = Mid$(driveletters$, i%, 1)
      ChDrive drivlet$
      mypath = drivlet$ + ":\" ' Set the path.

      myname = LCase(Dir(mypath, vbDirectory))   ' Retrieve the first entry.
      Do While myname <> sEmpty   ' Start the loop.
         'Ignore the current directory and the encompassing directory.
         If myname <> "." And myname <> ".." Then
            'Use bitwise comparison to make sure MyName is a directory.
            If (GetAttr(mypath & myname) And vbDirectory) = vbDirectory Then
               myname = LCase(myname)
               If s1% = 0 And myname = "jk" Then
                  s1% = 1: drivjk$ = drivlet$ + ":\jk\"

               ElseIf S2% = 0 And myname = "fordtm" Then
                  S2% = 1: drivfordtm$ = drivlet$ + ":\fordtm\"

                  'write busy signal
               ElseIf S3% = 0 And myname = "cities" Then
                  S3% = 1: drivcities$ = drivlet$ + ":\cities\"
                  defdriv$ = drivlet$

                  End If
               If s1% = 1 And S2% = 1 And S3% = 1 Then GoTo cdc1
               End If 'it represents a directory
            End If
         myname = Dir 'Get next entry
      Loop
   Next i%

cdc1: If s1% = 0 Then
         drivjk$ = InputBox("Can't find the ""jk"" directory, please give the full path name below " + _
                  "(e.g., if jk is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
         If drivjk$ <> sEmpty Then 'check the directory
            myname = Dir(drivjk$, vbDirectory)
            If myname = sEmpty Then
               response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
               If response = vbYes Then
                  GoTo cdc1
               Else
                  GoTo ce10
                  End If
               End If
         ElseIf drivjk$ = sEmpty Then 'user canceled the operation
            GoTo ce10
            End If
         End If
      
cdc2: If S2% = 0 Then
         drivfordtm$ = InputBox("Can't find the ""fordtm"" directory, please give the full path name below " + _
                  "(e.g., if fordtm is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
         If drivfordtm$ <> sEmpty Then 'check the directory
            myname = Dir(drivfordtm$, vbDirectory)
            If myname = sEmpty Then
               response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
               If response = vbYes Then
                  GoTo cdc2
               Else
                  GoTo ce10
                  End If
               End If
         ElseIf drivfordtm$ = sEmpty Then 'user canceled the operation
            GoTo ce10
            End If
         End If
      
cdc3: If S3% = 0 Then
         drivcities$ = InputBox("Can't find the ""cities"" directory, please give the full path name below " + _
                  "(e.g., if cities is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
         If drivcities$ <> sEmpty Then 'check the directory
            myname = Dir(drivcities$, vbDirectory)
            If myname = sEmpty Then
               response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
               If response = vbYes Then
                  GoTo cdc3
               Else
                  GoTo ce10
                  End If
               End If
         ElseIf drivcities$ = sEmpty Then 'user canceled the operation
            GoTo ce10
            End If
         End If
      

'      mydir = Dir(drivlet$ + ":\cities\*.*")
'      If mydir <> sEmpty Then
'         defdriv$ = drivlet$
'         GoTo 5
'         End If
      If s1% = 1 And S2% = 1 And S3% = 1 Then GoTo 5
   'if got here, means that couldn't find the cities directory
ce10: If internet = False Then
         MsgBox "Can't locate necessary directories! ABORTING program...Sorry", vbCritical + vbOKOnly, "Cal Programs"
         End If
      Call MDIform_queryunload(i%, j%)
      Exit Sub
      
   '******skip splash screen for Internet Server Version**********
5  internet = False
   optionheb = True
   If Dir(drivcities$ + "internet.yes") <> sEmpty Then
      intnum% = FreeFile
      Open drivcities$ & "internet.yes" For Input As #intnum%
      Input #intnum%, intflag%
      Input #intnum%, dirint$
      Close #intnum%
      If intflag% = 1 Then
         internet = True
         
        If Dir(drivfordtm$ & "busy.cal") = sEmpty Then
           busynum% = FreeFile
           Open drivfordtm$ + "busy.cal" For Output As #busynum%
           Print #busynum%, "Busy!"
           Close #busynum%
           End If
         
         lognum% = FreeFile
         'read last calprog.log--check if ended sucessfully
         'if ended sucessfully then start new log file,
         'if not ended sucessfully then append to it.
         appendlog% = 0
         If Dir(drivjk$ + "calprog.log") <> sEmpty Then
            Open drivjk$ + "calprog.log" For Input As #lognum%
            appendlog% = 1
            Do Until EOF(lognum%)
               Line Input #lognum%, doclin$
               If doclin$ = "Success! Cal program terminated normally." Then
                  appendlog% = 0
                  Exit Do
                  End If
            Loop
            Close lognum%
            End If
         lognum% = FreeFile
         If appendlog% = 0 Then
            Open drivjk$ + "calprog.log" For Output As #lognum%
         Else
            Open drivjk$ + "calprog.log" For Append As #lognum%
            Print #lognum%, "-----last scan didn't terminate successfully, appending log file for this new run-----------"
            End If
         Print #lognum%, "log file opened at system time/date: " & Trim$(Str$(Time)) & ", " & Trim$(Str$(Date))
         Close #lognum%
         End If
      End If
   If Dir(drivcities$ + "version.num") <> sEmpty Then
      vernum% = FreeFile
      Open drivcities$ + "version.num" For Input As #vernum%
      Input #vernum%, progvernum
      Close #vernum%
   Else
      If internet = True Then
         GoTo cd10
      Else
         Call versionfm_Click
         End If
      End If
   
   If internet = True Then GoTo cd10 'skip splash screen
   
   frmSplash.Show 0, Me
   frmSplash.NewLabel.Visible = True
   frmSplash.Label1.Visible = False
   ret = SetWindowPos(frmSplash.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)

   'erase any old *.tmp files on startup
    mypath = drivjk$ & "*.tmp" ' Set the path.
    myname = LCase(Dir(mypath, vbNormal))   ' Retrieve the first entry.
    Do While myname <> sEmpty   ' Start the loop.
       DoEvents
       If InStr(LCase(myname), "stat") <> 0 Then
          Kill drivjk$ & myname
          End If
       myname = Dir
    Loop
   
'   For i% = 1 To 1500
'      DoEvents
'      filstat$ = drivjk$ + "stat"
'      If i% <= 9 Then
'        filstat$ = filstat$ + "000" + Trim$(CStr(i%)) + ".tmp"
'      ElseIf i >= 10 And i% < 100 Then
'        filstat$ = filstat$ + "00" + Trim$(CStr(i%)) + ".tmp"
'      ElseIf i% >= 100 And i% < 1000 Then
'        filstat$ = filstat$ + "0" + Trim$(CStr(i%)) + ".tmp"
'      ElseIf i% >= 1000 And i% < 10000 Then
'        filstat$ = filstat$ + Trim$(CStr(i%)) + ".tmp"
'        End If
'      myfile = Dir(filstat$)
'      If myfile <> sEmpty Then Kill filstat$
'   Next i%
   
cd10:
   myfile = Dir(drivjk$ + "netzend.tmp")
   If myfile <> sEmpty Then Kill drivjk$ + "netzend.tmp"
   
   myfile = Dir(drivjk$ + "refflag.tmp")
   If myfile <> sEmpty Then Kill drivjk$ + "refflag.tmp"

If internet = False Then
   waitime = Timer
   Do Until Timer > waitime + 2
      DoEvents
   Loop
   CalMDIform.Visible = True
Else
   CalMDIform.WindowState = vbNormal
   End If
If internet = False Then
   waitime = Timer
   Do Until Timer > waitime + 1
      DoEvents
   Loop
   frmSplash.Visible = False
   End If
Screen.MousePointer = vbDefault

If internet = True Then
    
    'Set up timer to monitor progress of program.  If process remains
    'active after 5 minutes = , then the timer kills the process
    lngTimerID = SetTimer(0, 0, 300000, AddressOf TimerProc)
    'read lattest *.sev file
   'On Error GoTo errhand 'if file is being read or written, retry
   'ChDrive "c"
   'ChDir "c:/inetpub/webpub/data"
   Open drivjk$ + "calprog.log" For Append As #lognum%
   Print #lognum%, "Step #1: finding server batch file '*.ser'"
   Close #lognum%
   
   mypath = drivfordtm$
   myname = Dir(mypath, vbNormal)
   found% = 0
   Do While myname <> sEmpty
      If InStr(1, myname, ".ser") <> 0 Then
         found% = 1
         servnam$ = myname
         Exit Do
         End If
      myname = Dir
   Loop
   If found% = 1 Then
     lognum% = FreeFile
     Open drivjk$ + "calprog.log" For Append As #lognum%
     Print #lognum%, "Step #1.1: search for server '*.ser' file was successfull"
     Close #lognum%
   Else 'file not found
     Err.Raise vbObjectError + 1999, "CalMDIform", "server.dir is empty!"
     End If
   
   'open file to keep track of users' choices
   filuser% = FreeFile
   Open drivjk$ + "userlog.log" For Append As #filuser%
   lognum% = FreeFile
   Open drivjk$ + "calprog.log" For Append As #lognum%
   Print #lognum%, "Step #2: server.bat executed and following server file read: " & servnam$
   Print #lognum%, "The following is the contents of the server file files:"
   Print #lognum%, String$(50, "-")
   servnum% = FreeFile
   Open drivfordtm$ + servnam$ For Input As #servnum%
   Do Until EOF(servnum%)
      Line Input #servnum%, docserv$
      Print #lognum%, docserv$
   Loop
   Print #lognum%, String$(50, "-")
   Close servnum%
   Close #lognum%
   
   'now open server file and read city name, etc.
   servnum% = FreeFile
   Open drivfordtm$ + servnam$ For Input As #servnum%
   Line Input #servnum%, doclin$
   nettype$ = doclin$
   
   If doclin$ = "BY" Then 'Eretz Israel city
      parshiotEY = True
      Line Input #servnum%, doclin$
      If Asc(Mid$(doclin$, 1, 1)) < 128 + 96 Then
         currentdir = drivcities$ + doclin$
         userloglin$ = doclin$ & ", Eretz Yisroel"
         'Print #filuser%, doclin$ & ", Eretz Yisroel"
      ElseIf Asc(Mid$(doclin$, 1, 1)) >= 128 + 96 And Asc(Mid$(doclin$, 1, 1)) <= 154 + 96 Then 'hebrew city name detected
         'convert to english directory name
         hebcit% = FreeFile
         Open drivcities$ + "citynams_heb_w1255.txt" For Input As hebcit%
         foundhebcity% = 0
         Do Until EOF(hebcit%)
            Line Input #hebcit%, engcitynam$
            Line Input #hebcit%, hebrewcitynam$
            If doclin$ = hebrewcitynam$ Then
               currentdir = drivcities$ + engcitynam$
               userloglin$ = engcitynam$ & ", Eretz Yisroel"
               'Print #filuser%, engcitynam$ & ", Eretz Yisroel"
               foundhebcity% = 1
               Exit Do
               End If
         Loop
         Close #hebcit%
         If foundhebcity% = 0 Then
            'exit program with error message
             Close
             myfile = Dir(drivfordtm$ + "busy.cal")
             If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
           
             lognum% = FreeFile
             Open drivjk$ + "calprog.log" For Append As #lognum%
             Print #lognum%, "Fatal error: couldn't find corresponding english directory name. Abort program."
             Close #lognum%
           
             For i% = 0 To Forms.Count - 1
               Unload Forms(i%)
             Next i%
             
              'kill timer
              If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
    
              'end program abruptly
              End

             End If
         End If
      Line Input #servnum%, doclin$
      yrheb% = Val(doclin$)
      Print #filuser%, userloglin$ + " ," + Str(yrheb%) 'write this info to userlog.log
      Option1b = True 'Hebrew years only
      SunriseSunset.Combo1.Text = doclin$
      Line Input #servnum%, doclin$
      nsetflag% = Val(doclin$)
      Line Input #servnum%, doclin$
      typeflag% = Val(doclin$)
      Line Input #servnum%, doclin$
      zmanyes% = Val(doclin$)
      Line Input #servnum%, doclin$
      typezman% = Val(doclin$)
      Line Input #servnum%, doclin$
      optionheb = Val(doclin$)
      Line Input #servnum%, doclin$
      zmantype% = Val(doclin$)
      RoundSeconds% = 0
      If Not EOF(servnum%) Then
         Line Input #servnum%, doclin$
         RoundSeconds% = Val(doclin$)
         End If
      Close #servnum%
      'now use these parameters to set forms in the Cal Prog
      If typeflag% = 0 Then 'visible sunrise/sunset
        SunriseSunset.Check3.Value = 1 'check for near mountains
        If nsetflag% = 1 Then
           SunriseSunset.Check1.Value = 1
           SunriseSunset.Check2.Value = 0
        ElseIf nsetflag% = 2 Then
           SunriseSunset.Check1.Value = 0
           SunriseSunset.Check2.Value = 1
           End If
      ElseIf typeflag% = 2 Then 'astronomical sunrise/sunset
         SunriseSunset.Check3.Value = 0
         If nsetflag% = 1 Then
            SunriseSunset.Check4.Value = 1
            SunriseSunset.Check5.Value = 0
         ElseIf nsetflag% = 2 Then
            SunriseSunset.Check4.Value = 0
            SunriseSunset.Check5.Value = 1
            End If
      ElseIf typeflag% = 1 Then  'mishor sunrise/sunset
         SunriseSunset.Check3.Value = 0
         If nsetflag% = 1 Then
            SunriseSunset.Check6.Value = 1
            SunriseSunset.Check7.Value = 0
         ElseIf nsetflag% = 2 Then
            SunriseSunset.Check6.Value = 0
            SunriseSunset.Check7.Value = 1
            End If
         End If
      'find citnam
       If optionheb = True Then
          SunriseSunset.Option3.Value = True
       Else
          SunriseSunset.Option4.Value = True
          End If
      Caldirectories.Visible = True
      'SunriseSunset.OKbut(0).Value = True
   ElseIf Mid$(doclin$, 1, 4) = "Chai" Then 'eros cities
      parshiotdiaspora = True
      eroscountry$ = Mid$(doclin$, 5, Len(doclin$) - 4)
      eros = True
      geo = True
      eroscityflag = True
      'now read in parameters
      Input #servnum%, numUSAcities1%
      'this defines name of the metro area = erosareabat
         filnum% = FreeFile
         Open drivcities$ & "eros\" + eroscountry$ + "cities1.dir" For Input As #filnum%
         cnum% = 0
         Do Until cnum% = numUSAcities1%
            cnum% = cnum% + 1
            Line Input #filnum%, doclin2$
         Loop
         Close #filnum%
         erosareabat = doclin2$
      Input #servnum%, numUSAcities2%
         'this will define the eroscity name unless flagged by
         'yesmetro = 1 to use the inputed placename and coordinates
      Input #servnum%, searchradius
      Input #servnum%, yesmetro%
      If yesmetro% = 0 Then 'found the city in the city list, use it
        filnum% = FreeFile
        Open drivcities$ & "eros\" + eroscountry$ + "cities2.dir" For Input As #filnum%
        cnum% = 0
        Do Until cnum% = numUSAcities2%
           Line Input #filnum%, doclin2$
           cnum% = cnum% + 1
        Loop
        Close #filnum%
        pos% = InStr(doclin2$, "/")
        eroscity$ = Mid(doclin2$, pos% + 1, Len(doclin2) - pos%)
        userloglin$ = eroscountry$ & " ," & doclin2$
        'Print #filuser%, eroscountry$, doclin2$
        'now find the eroslongitude,eroslatitude
        foundcity% = 0
        filnum% = FreeFile
        Open drivcities$ & "eros\" + eroscountry$ + "cities3.dir" For Input As #filnum%
        Do Until EOF(filnum%)
            Input #filnum%, doclin3$, eroslatitude, eroslongitude, eroshgt
            'strip out city name
            poscity1% = InStr(doclin3$, "/")
            poscity2% = InStr(doclin3$, "\")
            checkcity$ = Mid$(doclin3$, poscity1% + 1, poscity2% - poscity1% - 1)
            If InStr(doclin3$, erosareabat$) <> 0 Then
               If InStr(eroscity$, checkcity$) <> 0 Then
                  foundcity% = 1
                  Exit Do
                  End If
               End If
        Loop
        Close #filnum%
        If foundcity% = 0 And Err.Number >= 0 Then 'inconsistent city and metro areas, abort
           Close
           myfile = Dir(drivfordtm$ + "busy.cal")
           If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
           
           lognum% = FreeFile
           Open drivjk$ + "calprog.log" For Append As #lognum%
           Print #lognum%, "Fatal error: inconsistent city and metro areas. Abort program."
           Close #lognum%
           
           For i% = 0 To Forms.Count - 1
             Unload Forms(i%)
           Next i%
          
           'kill timer
           If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)

           'end program abruptly
           End
           End If
        Line Input #servnum%, doclin$
        Line Input #servnum%, doclin$
        Line Input #servnum%, doclin$
        Line Input #servnum%, doclin$
     Else 'use in the recorded parameters
        Line Input #servnum%, eroscity$
        Input #servnum%, eroslatitude
        Input #servnum%, eroslongitude
        Input #servnum%, eroshgt
        End If
     Input #servnum%, yrheb%
     Print #filuser%, userloglin$ & " ," & Str(yrheb%) 'Update; userlog.Log
     Input #servnum%, types%
        'SunriseSunset.Check3.Value = 0 'for the time being, don't
                          'check for near obstructions
        Select Case types%
           Case 0
              viseros = True
              nsetflag% = 1 'visible sunrise
              SunriseSunset.Check1.Value = 1
              SunriseSunset.Check2.Value = 0
           Case 1
              viseros = False
              nsetflag% = 1 'mishor sunrise
              SunriseSunset.Check6.Value = 1
              SunriseSunset.Check7.Value = 0
           Case 2
              viseros = False
              nsetflag% = 1 'astron. sunrise
              SunriseSunset.Check4.Value = 1
              SunriseSunset.Check5.Value = 0
           Case 3
              viseros = False
              nsetflag% = 2 'mishor sunset
              SunriseSunset.Check6.Value = 0
              SunriseSunset.Check7.Value = 1
           Case 4
              viseros = False
              nsetflag% = 2 'astron. sunset
              SunriseSunset.Check4.Value = 0
              SunriseSunset.Check5.Value = 1
           Case 5
              viseros = True
              nsetflag% = 0 'visible sunset
              SunriseSunset.Check1.Value = 0
              SunriseSunset.Check2.Value = 1
        End Select
      Line Input #servnum%, doclin$
      zmanyes% = Val(doclin$)
      Line Input #servnum%, doclin$
      typezman% = Val(doclin$)
      Line Input #servnum%, doclin$
      optionheb = Val(doclin$)
      Line Input #servnum%, doclin$
      zmantype% = Val(doclin$)
      Close #servnum%
      erosareabat = erosareabat & "_" & eroscountry$
       If optionheb = True Then
          SunriseSunset.Option3.Value = True
       Else
          SunriseSunset.Option4.Value = True
          End If
      Caldirectories.Visible = True
   ElseIf doclin$ = "Astr" Then 'astronomical times
      geo = True
      astronplace = True
      Line Input #servnum%, doclin$
      astname$ = doclin$
      userloglin$ = "Astronomical name: " & doclin$
      'Print #filuser%, "Astronomical name: " & doclin$
      citnamp$ = astname$
      Line Input #servnum%, doclin$
      lat = Val(doclin$)
      Line Input #servnum%, doclin$
      lon = Val(doclin$)
      Line Input #servnum%, doclin$
      hgt = Val(doclin$)
      Line Input #servnum%, doclin$
      geotz! = Val(doclin$)
      'determine which sedra to use
      If lon < -34.21 And lon > -35.8333 And _
         lat > 29.55 And lat < 33.3417 And geotz! = 2 Then
         parshiotEY = True 'inside the borders of Eretz Yisroel
      Else
         parshiotdiaspora = True 'somewhere in the diaspora
         End If
      Line Input #servnum%, doclin$
      yrheb% = Val(doclin$)
      Print #filuser%, userloglin$ & " ," & Str(yrheb%) 'update userlog.log
      Option1b = True
      Line Input #servnum%, doclin$
      nsetflag% = Val(doclin$)
      Line Input #servnum%, doclin$
      typeflag% = Val(doclin$)
      Line Input #servnum%, doclin$
      zmanyes% = Val(doclin$)
      Line Input #servnum%, doclin$
      typezman% = Val(doclin$)
      Line Input #servnum%, doclin$
      optionheb = Val(doclin$)
      Line Input #servnum%, doclin$
      zmantype% = Val(doclin$)
      Close #servnum%
      filnum% = FreeFile
      If nsetflag% = 1 Then
         Open drivcities$ + "ast\netz\astr.bat" For Output As #filnum%
         Write #filnum%, drivfordtm$ + "netz\astronom.pr1", lat, lon, hgt
         Print #filnum%, "version"; ","; "1"; ","; "0"; ","; "0"
         Close #filnum%
         'now make dummy profile files to place in c:\cities\ast netz and skiy subdirectories
         filnum% = FreeFile
         Open drivcities$ + "ast\netz\astronom.pr1" For Output As #filnum%
         Write #filnum%, "FILENAME, KMX, KMY, HGT: ", drivprof$ + "astronom.fnz", lat, lon, hgt
      ElseIf nsetflag% = 2 Then
          Open drivcities$ + "ast\skiy\astr.bat" For Output As #filnum%
          Write #filnum%, drivfordtm$ + "skiy\astronom.pr1", lat, lon, hgt
          Print #filnum%, "version"; ","; "1"; ","; "0"; ","; "0"
          Close #filnum%
         'now make dummy profile files to place in c:\cities\ast netz and skiy subdirectories
          filnum% = FreeFile
          Open drivcities$ + "ast\skiy\astronom.pr1" For Output As #filnum%
          Write #filnum%, "FILENAME, KMX, KMY, HGT: ", drivprof$ + "astronom.fsk", lat, lon, hgt
          End If
       'now make dummy profile files to place in c:\cities\ast netz and skiy subdirectories
       Print #filnum%, "  AZI  VIEWANG+REFRACT   FLGSUM   FLGWIN"
       For i% = 1 To 601
          xentry = -30 + (i% - 1) * 0.1
          If xentry <= -10 Then
             Print #filnum%, Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
          ElseIf xentry > -10 And xentry < 0 Then
             Print #filnum%, " "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
          ElseIf xentry >= 0 And xentry < 10 Then
             Print #filnum%, "  "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
          ElseIf xentry >= 10 Then
             Print #filnum%, " "; Format(xentry, "#0.0#"); "     "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000"); "    "; Format(0#, "#0.0000")
             End If
       Next i%
       Close #filnum%
       startedscan = False
     
      If typeflag% = 2 Then 'astronomical sunrise/sunset
         SunriseSunset.Check3.Value = 0
         If nsetflag% = 1 Then
            SunriseSunset.Check4.Value = 1
            SunriseSunset.Check5.Value = 0
         ElseIf nsetflag% = 2 Then
            SunriseSunset.Check4.Value = 0
            SunriseSunset.Check5.Value = 1
            End If
      ElseIf typeflag% = 1 Then  'mishor sunrise/sunset
         SunriseSunset.Check3.Value = 0
         If nsetflag% = 1 Then
            SunriseSunset.Check6.Value = 1
            SunriseSunset.Check7.Value = 0
         ElseIf nsetflag% = 2 Then
            SunriseSunset.Check6.Value = 0
            SunriseSunset.Check7.Value = 1
            End If
         End If
       avekmxnetz = lon 'parameters used for z'manim tables
       avekmynetz = lat
       avehgtnetz = hgt
       avekmxskiy = lon
       avekmyskiy = lat
       avehgtskiy = hgt
       aveusa = True 'invert coordinates in SunriseSunset due to
                     'silly nonstandard coordinate format
       
       If optionheb = True Then
          SunriseSunset.Option3.Value = True
       Else
          SunriseSunset.Option4.Value = True
          End If
       
       Caldirectories.Visible = True
       Caldirectories.Text1.Text = drivcities$ + "ast\" + LTrim$(astname$)
       currentdir = Trim$(Caldirectories.Text1.Text)
      End If
   End If
   Close #filuser%
Exit Sub

errhand:
   Resume 'try again
   
generrhand:
     If internet = True Then
         'abort the program with a error messages
        errlog% = FreeFile
        Open drivjk$ + "Cal_ssgeh.log" For Output As #errlog%
        Print #errlog%, "Cal Prog exited from MDIForm with runtime error number" + Str(Err.Number) & vbLf & Err.Description
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
        If Err.Number = 68 Then 'empty CD drive, so skip it
           i% = i% + 1 'increment to next drive letter
           drivlet$ = Mid$(driveletters$, i%, 1)
           Resume Next 'try again
        ElseIf Err.Number = 76 Then
           Resume Next
           End If
        response = MsgBox("CalMDIform encountered error number: " + Str(Err.Number) + vbLf & Err.Description & vbLf & "Do you want to abort?", vbYesNoCancel + vbCritical, "Cal Program")
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

Private Sub mnuChaiAir_Click()
   calAirfm.Visible = True
End Sub

Private Sub mnuStndRef_Click()
    If mnuStndRef.Checked Then
    Else
       nweatherflag = 1
       mnuStndRef.Checked = True
       mnuVDW.Checked = False
       
       'write refflag file
       filflag% = FreeFile
       Open drivjk$ & "refflag.tmp" For Output As #filflag%
       Print #filflag%, nweatherflag%
       Close #filflag%
       
       End If
End Sub

Private Sub mnuVDW_Click()
  If mnuVDW.Checked Then
  Else
    nweatherflag = 0
    mnuVDW.Checked = True
    mnuStndRef.Checked = False
       
    'write refflag file
    filflag% = FreeFile
    Open drivjk$ & "refflag.tmp" For Output As #filflag%
    Print #filflag%, nweatherflag%
    Close #filflag%
    
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
   On Error GoTo Toolbar1_ButtonClick_Error

   Select Case Button.Key
     Case "calfm" 'Cal Program
        If eroscityflag = True Then
           Unload calnode
           Set calnode = Nothing
           eroscityflag = False
           End If
        'Screen.MousePointer = vbHourglass
        Caldirectories.Visible = True
     Case "Katzform" 'Katz Program
        If eroscityflag = True Then
           Unload calnode
           Set calnode = Nothing
           eroscityflag = False
           End If
        Katz = True
        If Dir(drivjk$ + openfil$) = sEmpty Then
            Call MsgBox("The file: " & drivjk$ & "Katzplaces.sav" & " doesn't exit!" _
                        & vbCrLf & "" _
                        & vbCrLf & "Restore the file to the directory: " + drivjk$ _
                        , vbExclamation, "astronomical places")
           Exit Sub
           End If
           
        If numAstPlaces% = 0 Then
            Call MsgBox("Warning: The file: " & drivjk$ & "Katzplaces.sav" & " is empty!" _
                        & vbCrLf & "" _
                        & vbCrLf & "Restore the file to the directory: " + drivjk$ _
                        , vbExclamation, "astronomical places")
           End If
           
        AstronForm.Visible = True
        If numAstPlaces% = 0 Then Exit Sub
        katztotal% = 0
        AstronForm.Combo1.ListIndex = 0
        AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
        If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
           AstronForm.Option2.Value = True
        ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
           AstronForm.Option1.Value = True
           End If
     Case "worldbut"
       ecdir$ = drivcities$ + "eros\"
       myfile = Dir(ecdir$ & "eroscity.sav")
       If myfile = sEmpty Then
          response = MsgBox("Can't find the eroscity.sav file", vbCritical + vbOKOnly, "Maps & More")
          Exit Sub
          End If
       eroscityflag = True
       calnode.Visible = True
       ret = SetWindowPos(CalMDIform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       BringWindowToTop (calnode.hwnd)
     Case "astrobut"
        If eroscityflag = True Then
           Unload calnode
           Set calnode = Nothing
           eroscityflag = False
           End If
       ret = SetWindowPos(CalMDIform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       astronplace = True
       AstronForm.Visible = True
       BringWindowToTop (AstronForm.hwnd)
     Case "ChaiAirTables"
        calAirfm.Visible = True
     Case Else
  End Select

   On Error GoTo 0
   Exit Sub

Toolbar1_ButtonClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Toolbar1_ButtonClick of Form CalMDIform"
End Sub
Public Sub MDIform_queryunload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo mqerr50
   myfile = Dir(drivfordtm$ + "busy.cal")
   If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
   
   For i% = 0 To Forms.Count - 1
      Unload Forms(i%)
   Next i%
   'Unload CalMDIform
   
mqerr50:
   Close
   End
End Sub

Private Sub versionfm_Click()
   response = InputBox("Please input the version number for the calendars. The version number must be in decimal form (e.g., 5, 5.0, 6.1, but NOT 5.1.1 !)  Warning: do not change the default value unless instructed to do so by Chaim Keller", "Version Number", progvernum)
   If response <> Trim$(Str(progvernum)) Then
      newres = MsgBox("Do you really want to change the version number?", vbExclamation + vbYesNoCancel, "Cal Program Version Number")
      If newres <> vbYes Then
         Exit Sub
      Else
         progvernum = response
         vernum% = FreeFile
         Open drivcities$ + "version.num" For Output As #vernum%
         Write #vernum%, progvernum
         Close #vernum%
         End If
      End If
End Sub

Private Sub worldsunfm_Click()
       ecdir$ = drivcities$ + "eros\"
       myfile = Dir(ecdir$ & "eroscity.sav")
       If myfile = sEmpty Then
          response = MsgBox("Can't find the eroscity.sav file", vbCritical + vbOKOnly, "Maps & More")
          Exit Sub
          End If
       eroscityflag = True
       calnode.Visible = True
       ret = SetWindowPos(CalMDIform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
       BringWindowToTop (calnode.hwnd)
End Sub
