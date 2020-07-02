VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form calAirfm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chai Air Travel Tables Entry Form"
   ClientHeight    =   5715
   ClientLeft      =   2520
   ClientTop       =   1710
   ClientWidth     =   7410
   Icon            =   "calAirfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7410
   Begin VB.Frame frameArrival 
      Caption         =   "Arrivals"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2715
      Left            =   120
      TabIndex        =   9
      Top             =   2340
      Width           =   7155
      Begin VB.CommandButton cmdEquate 
         Caption         =   "&Equate Departure and Arrival Dates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   16
         Top             =   2280
         Width           =   4815
      End
      Begin VB.ComboBox cmbArrival 
         Height          =   315
         Left            =   420
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   360
         Width           =   6375
      End
      Begin VB.CheckBox chkArrival 
         Caption         =   "Departure Point is currently on Daily Savings Times"
         Height          =   195
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   4035
      End
      Begin MSComCtl2.DTPicker dtTimeArrival 
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   1380
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm"
         Format          =   119275522
         CurrentDate     =   37292
      End
      Begin MSComCtl2.DTPicker dtDateArrival 
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   900
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   192
         CalendarTitleForeColor=   16777215
         Format          =   119275521
         CurrentDate     =   37292
      End
      Begin VB.Label lblArrivalDate 
         Caption         =   "Arrival Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2100
         TabIndex        =   15
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblArrivalTime 
         Caption         =   "Estimated Arrival Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         Top             =   1440
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Times"
      Height          =   375
      Left            =   3660
      TabIndex        =   2
      Top             =   5220
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate Times"
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   5220
      Width           =   1575
   End
   Begin VB.Frame frameDeparture 
      Caption         =   "Departures"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      Begin MSComCtl2.DTPicker dtTimeDeparture 
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   1380
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm"
         Format          =   119275522
         CurrentDate     =   37292
      End
      Begin MSComCtl2.DTPicker dtDateDeparture 
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   900
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   16777215
         Format          =   119275521
         CurrentDate     =   37292
      End
      Begin VB.CheckBox chkDeparture 
         Caption         =   "Departure Point is currently on Daily Savings Times"
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   1920
         Width           =   4035
      End
      Begin VB.ComboBox cmbDeparture 
         Height          =   315
         Left            =   420
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label lblDepartureTime 
         Caption         =   "Expected Departure Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblDepartureDate 
         Caption         =   "Departure Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   1395
      End
   End
End
Attribute VB_Name = "calAirfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lt1 As Double, lg1 As Double, hgt1 As Double
Dim lt2 As Double, lg2 As Double, hgt2 As Double
Dim TZ1 As Double, TZ2 As Double, TotalDist As Double
Private Sub cmdCalculate_Click()
   On Error GoTo errhand
   
   Dim DayDeparture As Integer, MonthDeparture As Integer
   Dim YearDeparture As Integer, DSTDeparture As Integer
   Dim HourDeparture As Integer, MinuteDeparture As Integer
   Dim DayArrival As Integer, MonthArrival As Integer
   Dim YearArrival As Integer, DSTArrival As Integer
   Dim HourArrival As Integer, MinuteArrival As Integer
   Dim TimeZoneDeparture As Double, TimeZoneArrival As Double
   

   DayDeparture = dtDateDeparture.day
   MonthDeparture = dtDateDeparture.month
   YearDeparture = dtDateDeparture.Year
   If YearDeparture < 1996 Then
      MsgBox "Years earlier than 1996 not supported", vbExclamation + vbOKOnly, "Cal Program"
      Exit Sub
      End If
   HourDeparture = dtTimeDeparture.Hour
   MinuteDeparture = dtTimeDeparture.Minute
   If chkDeparture.Value = vbChecked Then
      DSTDeparture = -1
   Else
      DSTDeparture = 0
      End If
   
   DayArrival = dtDateArrival.day
   MonthArrival = dtDateArrival.month
   YearArrival = dtDateArrival.Year
   If YearArrival < 1996 Then
      MsgBox "Years earlier than 1996 not supported", vbExclamation + vbOKOnly, "Cal Program"
      Exit Sub
      End If
   HourArrival = dtTimeArrival.Hour
   MinuteArrival = dtTimeArrival.Minute
   If chkArrival.Value = vbChecked Then
      DSTArrival = -1
   Else
      DSTArrival = 0
      End If
      
   'now check if times are possible
   CheckCoord
   TimeZoneDeparture = TZ1
   TimeZoneArrival = TZ2
   yfDepart = jd(DayDeparture, MonthDeparture, YearDeparture, _
       HourDeparture, MinuteDeparture, TimeZoneDeparture, _
       DSTDeparture)
   
   yfArrival = jd(DayArrival, MonthArrival, YearArrival, _
       HourArrival, MinuteArrival, TimeZoneArrival, _
       DSTArrival)
       
   If yfArrival <= yfDepart Then
      MsgBox "The arrival must be later then the departure time!", vbExclamation + vbOKOnly, "Cal Program"
      Exit Sub
      End If
     
   'Now check the speed
   Speed = TotalDist / ((yfArrival - yfDepart) * 24)
   If Speed < 400 And Speed > 1600 Then
      response = MsgBox("Something is wrong with your inputs that makes the plane travel too fast!" + vbLf + _
             "Do you wan't to try again?", vbExclamation + vbYesNoCancel, "Cal Program")
      If response = vbYes Then Exit Sub
      End If
             
  'Finished checks, now calculate the table
   ChDrive Mid$(drivjk$, 1, 2)
   ChDir drivjk$
   cdir$ = CurDir
   
   shel$ = "ChaiAirTimes.exe" & Str(cmbDeparture.ListIndex + 1) + Str(DayDeparture) + Str(MonthDeparture) + Str(YearDeparture) + Str(HourDeparture) + Str(MinuteDeparture) + " " + Str(DSTDeparture) + _
          Str(cmbArrival.ListIndex + 1) & Str(DayArrival) + Str(MonthArrival) + Str(YearArrival) + Str(HourArrival) + Str(MinuteArrival) + " " + Str(DSTArrival)
   Shell (shel$) 'execute ChaiAirTimes program
   Exit Sub
   
errhand:
   MsgBox "Encountered Error Number: " & Str(Err.Number) + vbLf + _
          "Error description follows." & vbLf & _
          Err.Description, vbOKOnly, "Cal Program"
   
End Sub

Private Sub cmdEquate_Click()
  dtDateArrival.day = dtDateDeparture.day
  dtDateArrival.month = dtDateDeparture.month
  dtDateArrival.Year = dtDateDeparture.Year
End Sub

Private Sub cmdView_Click()
   ret = SetWindowPos(calAirfm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   ret = SetWindowPos(CalMDIform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   currentdir$ = drivjk$
   If Dir("e:\progra~1\intern~1\iexplore.exe") <> sEmpty Then
      'Windows XP, Pentium IV
      ret = Shell("e:\progra~1\intern~1\iexplore " & currentdir$ & "AirTimes.html", vbMaximizedFocus)
   ElseIf Dir("f:\progra~1\intern~1\iexplore.exe") <> sEmpty Then
      'Windows XP, Pentium III
      ret = Shell("f:\progra~1\intern~1\iexplore " & currentdir$ & "AirTimes.html", vbMaximizedFocus)
   ElseIf Dir("c:\progra~1\plus!\micros~1\iexplore.exe") <> sEmpty Then
      'Pentium I at Kollel
      ret = Shell("c:\progra~1\plus!\micros~1\iexplore " & currentdir$ & "AirTimes.html", vbMaximizedFocus)
   ElseIf Dir("c:\progra~1\intern~1\iexplore.exe") <> sEmpty Then
      'Windows 98, Pentium III,IV
      ret = Shell("c:\progra~1\intern~1\iexplore " & currentdir$ & "AirTimes.html", vbMaximizedFocus)
      End If
End Sub

Private Sub Form_Load()
   'version: 04/08/2003
   On Error GoTo errhand
   
   'open Airplaces.txt and populate combo boxes
   myfile = Dir(drivjk$ + "Airplaces.txt")
   If myfile <> sEmpty Then
      filair% = FreeFile
      cmb% = 1
      Open drivjk$ + "Airplaces.txt" For Input As #filair%
      Line Input #filair%, doclin$
      Do Until EOF(filair%)
         Line Input #filair%, doclin1$
         If doclin1$ <> sEmpty And cmb% = 1 Then
            cmbDeparture.AddItem doclin1$
            Line Input #filair%, doclin2$
         ElseIf doclin1$ = sEmpty Then
            cmb% = 2
            Line Input #filair%, doclin2$
         ElseIf doclin1$ <> sEmpty And cmb% = 2 Then
            cmbArrival.AddItem doclin1$
            Line Input #filair%, doclin2$
         End If
      Loop
   End If
   cmbDeparture.ListIndex = 0 'place on first item
   cmbArrival.ListIndex = 0
   ret = SetWindowPos(calAirfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   
   'default date is today's date
   dtDateDeparture.Year = Year(Now)
   dtDateDeparture.month = month(Now)
   dtDateDeparture.day = day(Now)
   
   dtDateArrival.Year = Year(Now)
   dtDateArrival.month = month(Now)
   dtDateArrival.day = day(Now)
    
   Exit Sub
  
errhand:
   MsgBox "Encountered Error Number: " & Str(Err.Number) + vbLf + _
          "Error description follows." & vbLf & _
          Err.Description, vbOKOnly, "Cal Program"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set calAirfm = Nothing
   ret = SetWindowPos(CalMDIform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

Function jd(day As Integer, month As Integer, Year As Integer, _
       Hour As Integer, Minute As Integer, TimeZone As Double, _
       DST As Integer) As Double
    
    Dim yr As Integer, yd As Integer, yl As Integer, yrtst As Integer
    Dim yltst As Integer, i As Integer
    Dim yf As Double

    'Calculate the Julian Date of Departure */
    yr = Year
    yf = 0
    yd = yr - 1996 '(Note: years earlier than 1996 not supported)
    '* Astronomical constants used were calculated for
    'the year 1996.  So calculate them for latter years.
    '(N.B., astronomical constants are accurate to 6 seconds
    'for periods of 50 years.  For these calculations, this
    'is all the accuracy needed.) */
    For i = 1996 To yr - 1
        yrtst = i
        yltst = 365 ' Normal civil year */
        If (yrtst - 1996 Mod 4 = 0) Then
            yltst = 366 ' Civil Leap Year */
            End If
        If (yrtst Mod 100 = 0 And yrtst Mod 400 <> 0) Then
            yltst = 365 ' Gregorian cycle non leap year */
            End If
        yf = yf + yltst
    Next i
    ' yf number due to years from J2000 = Jan.0. 12:00 P.M. = 0 UTM */
    yf = yf - 1462.5 ' Number of days from J2000.0 at Jan. 0. 12:00 P.M. = 0 UTM */

    '* now add contribution from Month, Day and Time */
    Dim nleap As Integer, dayyr As Integer

    yl = 365 '; /* Normal civil year */
    If (yr - 1996 Mod 4 = 0) Then
        yl = 366 '; /*Civil leap year */
        End If
    If (yr Mod 100 = 0 And yr Mod 400 <> 0) Then
        yl = 365 '; /* Gregorian cycle non leap year */
        End If
    If (yl = 366) Then nleap = 1

    If (month = 1) Then
       dayyr = 0
    ElseIf (month = 2) Then
       dayyr = 31
    ElseIf (month = 3) Then
       dayyr = 59 + nleap
    ElseIf (month = 4) Then
       dayyr = 90 + nleap
    ElseIf (month = 5) Then
       dayyr = 120 + nleap
    ElseIf (month = 6) Then
       dayyr = 151 + nleap
    ElseIf (month = 7) Then
       dayyr = 181 + nleap
    ElseIf (month = 8) Then
       dayyr = 212 + nleap
    ElseIf (month = 9) Then
       dayyr = 243 + nleap
    ElseIf (month = 10) Then
       dayyr = 273 + nleap
    ElseIf (month = 11) Then
       dayyr = 304 + nleap
    ElseIf (month = 12) Then
       dayyr = 334 + nleap
       End If

    Dim UT As Double ' /* Universal Time */
    UT = Hour + Minute / 60# - TimeZone + DST
    Dim dayyrn As Double
    dayyrn = dayyr + day + UT / 24#
    jd = yf + dayyrn ' /* total Julian Number including contrib.
                     '    from Months, Days, Hours, TimeZone, and
                     '    DST */

End Function

Sub CheckCoord()
   filair% = FreeFile
   myfile = Dir(drivjk$ + "Airplaces.txt")
   If myfile <> sEmpty Then
      filair% = FreeFile
      cmb% = 1
      num% = 0
      Open drivjk$ + "Airplaces.txt" For Input As #filair%
      Line Input #filair%, doclin$
      Do Until EOF(filair%)
         Line Input #filair%, doclin1$
         If doclin1$ <> sEmpty And cmb% = 1 Then
            If num% = cmbDeparture.ListIndex Then
               Input #filair%, lg1, lt1, hgt1, TZ1
            Else
               Line Input #filair%, doclin2$
               End If
         ElseIf doclin1$ = sEmpty Then
            cmb% = 2
            num% = 0
            Line Input #filair%, doclin2$
         ElseIf doclin1$ <> sEmpty And cmb% = 2 Then
            If num% = cmbArrival.ListIndex + 1 Then
               Input #filair%, lg2, lt2, hgt2, TZ2
               Exit Do
            Else
               Line Input #filair%, doclin2$
               End If
         End If
      num% = num% + 1
      Loop
   End If
   
   'calculate Total Distance
'   pi = 4 * Atn(1)
'   cd = pi / 180#  'conv deg to rad
   X11 = Cos(lt1 * cd) * Cos(lg1 * cd)
   X22 = Cos(lt2 * cd) * Cos(lg2 * cd)
   y11 = Cos(lt1 * cd) * Sin(lg1 * cd)
   y22 = Cos(lt2 * cd) * Sin(lg2 * cd)
   Z11 = Sin(lt1 * cd)
   Z22 = Sin(lt2 * cd)
   'distance is Re * Angle between vectors
   'cos(Angle between unit vectors) = Dot product of unit vectors
   Dim cosang As Double
   cosang = X11 * X22 + y11 * y22 + Z11 * Z22
   TotalDist = 6371.315 * FNarco(cosang)
   
End Sub
