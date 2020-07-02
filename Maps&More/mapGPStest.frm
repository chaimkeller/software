VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form GPStest 
   Caption         =   "GPS Reader"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   4920
      ScaleHeight     =   1635
      ScaleWidth      =   2475
      TabIndex        =   76
      Top             =   2760
      Width           =   2535
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "Ed Keller/Alex Etchells 2011/2006"
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "Com Port, BaudRate set:"
         Height          =   255
         Left            =   360
         TabIndex        =   80
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Baud Rate is set to 38400"
         Height          =   255
         Left            =   360
         TabIndex        =   79
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "CommPort is sert to 5 "
         Height          =   255
         Left            =   480
         TabIndex        =   78
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "The Basics of Using VB6NMEAinterpreter"
         Height          =   495
         Left            =   240
         TabIndex        =   77
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.TextBox Textdgpsage 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Text            =   "-"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TextAlt 
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox TextSpeed 
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox TextBearing 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TextLon 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox TextLat 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "quit"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2520
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   38400
   End
   Begin VB.Label Label12 
      Caption         =   "Mode"
      Height          =   255
      Left            =   5880
      TabIndex        =   75
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Satellites in view = "
      Height          =   255
      Left            =   4920
      TabIndex        =   74
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "Quality"
      Height          =   255
      Left            =   5880
      TabIndex        =   73
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Satellites in use = "
      Height          =   255
      Left            =   4920
      TabIndex        =   72
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   71
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   70
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   12
      Left            =   1320
      TabIndex        =   69
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   68
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   11
      Left            =   3600
      TabIndex        =   67
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   11
      Left            =   2520
      TabIndex        =   66
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   11
      Left            =   1320
      TabIndex        =   65
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   64
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   63
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   62
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   61
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   60
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   59
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   58
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   57
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   56
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   55
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   54
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   8
      Left            =   1320
      TabIndex        =   53
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   52
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   51
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   50
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   49
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   48
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   47
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   46
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   45
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   44
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   43
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   42
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   41
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   40
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   39
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   38
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   37
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   36
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   35
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   34
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   33
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   32
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   31
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   30
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   29
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   28
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   27
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   26
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   25
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   24
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label LabelsatSNR 
      Caption         =   "Sat SNR"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   23
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LabelSatAzi 
      Caption         =   "Sat Azimuth"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   22
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LabelSatEle 
      Caption         =   "Sat Elevation"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   21
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LabelSatID 
      Caption         =   "Sat ID"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label LabelStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label LabelDGPSstat 
      Caption         =   "DGPS stat ID = "
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "DGPS age:"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Altitude:"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label LabelAltunit 
      Caption         =   "?"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label LabelTime 
      Caption         =   "time"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label LabelDate 
      Caption         =   "date"
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "knots"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Speed:"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Bearing:"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "HDOP"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Lon:"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Lat:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "GPStest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' USING VB6NMEAinterpreter - the basics
' There's nothing elegant here =o)
' Alex Etchells 2006    a.etchells@gmail.com

Option Explicit

Dim WithEvents gps As VB6NmeaInterpreter
Attribute gps.VB_VarHelpID = -1
Dim currentHDOP As Double
Dim satsInView As Integer
Dim lastSatsInView As Integer
Dim inString As String
Dim Iport%, Iport0%

Private Sub Command1_Click()
MSComm1.PortOpen = False
End
End Sub

Private Sub form_load()

GPS_test_loaded = True

Set gps = New VB6NmeaInterpreter

'scan over ports backward, find first one that fits baud rate and parity
'this should be the Prolific USB to Serial COM port used for the GPS

On Error GoTo errhand

   GPSconnected = False

   Iport0% = MAX_PORT
   For Iport% = MAX_PORT To 1 Step -1
      MSComm1.CommPort = Iport%
      MSComm1.Settings = GPSConnectString '"38400,N,8,1"
      MSComm1.PortOpen = True
   Next Iport%

Exit Sub
errhand:

   If Iport% = Iport0% Then
      Iport0% = Iport% - 1
      If Iport0% = 0 Then 'couldn't find any GPS from COM ports 1 to MAX_PORT
      
         If GPSSetupVis Then Unload GPSsetup
         
         If GPSenabled Then Exit Sub
         
         Select Case MsgBox("GSP device not found or no signal detected after searching for " _
                    & LTrim$(Str$(MAX_GPS_WAIT)) & " minutes." _
                    & vbCrLf & "Do you want to find the connection to your GPS yourself?" _
                    & vbCrLf & vbCrLf _
                    & "(If you forgot to insert your GPS device into a USB connection, " _
                    & vbCrLf & "answer ''No'', and insert your device into a USB port.)" _
                    , vbYesNoCancel + vbQuestion, "GPS Serial Port")
                    
             Case vbYes
             
                GPSsetup.Visible = True
             
             Case vbNo, vbCancel
             
'                Maps.GPS_timer.Enabled = True
                GPS_timer_trials = 0
                GPS_no_message = True
                Unload GPStest
                Maps.GPS_connect
                
                Exit Sub
                
         End Select
         
        End If
   Else
      'must have skipped the error handler at Iport0%, meaning that GPS is on Com port Iport0%
      ComPort% = Iport0%
      Label14.Caption = "CommPort is set to " & LTrim$(Str$(ComPort%))
'      Maps.GPS_timer.Enabled = True 'check for connection
      If GPSSetupVis Then
         waitimeGPSvis = Timer
         Do Until Timer > waitimeGPSvis + 2
            DoEvents
         Loop
         Unload GPSsetup
         End If
      Exit Sub
      End If
   
   Resume Next
   
End Sub
Private Sub MSComm_OnComm()

   If Not GPS_test_loaded Then Exit Sub
   
   Select Case MSComm1.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of
                        ' chars.
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   GPS_test_loaded = False
End Sub

Private Sub gps_AltitudeReceived(ByVal altitude As Double)
GPS_altitude = altitude
TextAlt.Text = CStr(altitude)
End Sub

Private Sub gps_AltitudeUnitsReceived(ByVal altitudeUnits As String)
GPS_altitudeunits = altitudeUnits
LabelAltunit.Caption = altitudeUnits
End Sub

Private Sub GPS_AutoManModeReceived(ByVal autoManMode As String)
GPS_ModeReceived = "Mode = " & autoManMode
Label12.Caption = "Mode = " & autoManMode
End Sub

Private Sub gps_BearingReceived(ByVal bearing As Double)
GPS_bearing = bearing
TextBearing.Text = CStr(bearing)
End Sub

Private Sub gps_Ch1SatReceived(ByVal ch1Sat As Integer)
If ch1Sat = 0 Then
    LabelSatID(1).ForeColor = &H80000011
Else
    LabelSatID(1).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch2SatReceived(ByVal ch2Sat As Integer)
If ch2Sat = 0 Then
    LabelSatID(2).ForeColor = &H80000011
Else
    LabelSatID(2).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch3SatReceived(ByVal ch3Sat As Integer)

If ch3Sat = 0 Then
    LabelSatID(3).ForeColor = &H80000011
Else
    LabelSatID(3).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch4SatReceived(ByVal ch4Sat As Integer)
If ch4Sat = 0 Then
    LabelSatID(4).ForeColor = &H80000011
Else
    LabelSatID(4).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch5SatReceived(ByVal ch5Sat As Integer)
If ch5Sat = 0 Then
    LabelSatID(5).ForeColor = &H80000011
Else
    LabelSatID(5).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch6SatReceived(ByVal ch6Sat As Integer)
If ch6Sat = 0 Then
    LabelSatID(6).ForeColor = &H80000011
Else
    LabelSatID(6).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch7SatReceived(ByVal ch7Sat As Integer)
If ch7Sat = 0 Then
    LabelSatID(7).ForeColor = &H80000011
Else
    LabelSatID(7).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch8SatReceived(ByVal ch8Sat As Integer)
If ch8Sat = 0 Then
    LabelSatID(8).ForeColor = &H80000011
Else
    LabelSatID(8).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch9SatReceived(ByVal ch9Sat As Integer)
If ch9Sat = 0 Then
    LabelSatID(9).ForeColor = &H80000011
Else
    LabelSatID(9).ForeColor = &H80000012
End If
End Sub


Private Sub gps_Ch10SatReceived(ByVal ch10Sat As Integer)
If ch10Sat = 0 Then
    LabelSatID(10).ForeColor = &H80000011
Else
    LabelSatID(10).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch11SatReceived(ByVal ch11Sat As Integer)
If ch11Sat = 0 Then
    LabelSatID(11).ForeColor = &H80000011
Else
    LabelSatID(11).ForeColor = &H80000012
End If
End Sub

Private Sub gps_Ch12SatReceived(ByVal ch12Sat As Integer)
If ch12Sat = 0 Then
    LabelSatID(12).ForeColor = &H80000011
Else
    LabelSatID(12).ForeColor = &H80000012
End If
End Sub

Private Sub gps_DateChanged(ByVal satDate As String)
GPS_date = satDate
LabelDate.Caption = satDate
End Sub

Private Sub gps_DGPSageReceived(ByVal dGPSage As Integer)
Textdgpsage.Text = CStr(dGPSage)
End Sub

Private Sub gps_DGPSstationIDReceived(ByVal dGPSstationID As String)
LabelDGPSstat.Caption = "DGPS stat ID = " & dGPSstationID
End Sub

Private Sub gps_FixLost()

GPSconnected = False 'gps fix lost
GPS_no_message = True
'Maps.GPS_timer.Enabled = True 'look for a reconnect every 10 seconds
Maps.mnuGPS_goto.Enabled = False
Maps.mnuGPS.Checked = False

LabelStatus.Caption = "Fix lost"
Maps.StatusBar1.Panels(2).Text = "GPS fix lost"

End Sub

Private Sub gps_FixObtained()

    On Error GoTo errhand

    Dim GPSday As Integer, GPSmonth As Integer, GPSyear As Integer
    Dim GPShour As Integer, GPSminute As Integer, GPSsecond As Integer
    Dim lat1 As Double, lon1 As Double

    LabelStatus.Caption = "Fix Obtained"
    Maps.StatusBar1.Panels(2).Text = "GPS fix..."
    
    If Not GPS_test_loaded Then Exit Sub
    
    If GPSSetupVis Then Unload GPSsetup
    
    GPSconnected = True 'gps position fixed
    GPSenabled = True 'this connection works, so don't go searching for it again
    Maps.mnuGPS_goto.Enabled = True
    Maps.mnuGPS.Checked = True
    
    'if baud rate is different then initial values, record the new connect string
    If GPSConnectString <> GPSConnectString0 Then
       GPSConnectString0 = GPSConnectString
       SaveSetting App.Title, "Settings", "GPS serial-USB connection string", GPSConnectString
       End If

   
   'use the departure coordinates as the first gps point, and this point as the next one
   '(assume that it was missed since airlines don't allow computers to operate
   'during takeoff until the airplane reaches a minimum cruising altitude)
   lat1 = GPS_latitude 'Val(GPStest.TextLat.text)
   lon1 = GPS_longitude 'Val(GPStest.TextLon.text)
    
    'use GPS time for calculating the time difference
    GPSday = Val(Mid$(GPS_date, 1, 2))
    GPSmonth = Val(Mid$(GPS_date, 4, 2))
    GPSyear = 2000 + Val(Mid$(GPS_date, 7, 2))
    GPShour = Val(Mid$(GPS_time, 1, 2))
    GPSminute = Val(Mid$(GPS_time, 4, 2))
    GPSsecond = Val(Mid$(GPS_time, 7, 2))
    
    GPS_warning_number = 0
       
   
errhand:
   
End Sub

Private Sub gps_HDOPReceived(ByVal value As Double)
currentHDOP = value
Label3.Caption = "HDOP = " & CStr(value)
End Sub

Private Sub gps_Mode3DReceived(ByVal mode3D As String)
Label10.Caption = "Mode = " & mode3D
End Sub

Private Sub gps_PositionReceived(ByVal latitude As String, ByVal longitude As String)
' If currentHDOP <= 6 Then
      ' Yes.  Display the current position
   
       GPS_latitude = latitude
       GPS_longitude = longitude
      
      
      TextLat.Text = latitude
      TextLon.Text = longitude
'    Else
'      TextLat.Text = "Poor"
'      TextLon.Text = "Signal!"
 '   End If
End Sub

Private Sub gps_SatelliteReceived(ByVal satelliteNumber As Integer, ByVal pseudoRandomCode As Integer, ByVal azimuth As Integer, ByVal elevation As Integer, ByVal signalToNoiseRatio As Integer)

On Error GoTo errhand

LabelSatID(satelliteNumber).Caption = CStr(pseudoRandomCode)
LabelSatEle(satelliteNumber).Caption = CStr(elevation)
LabelSatAzi(satelliteNumber).Caption = CStr(azimuth)
LabelsatSNR(satelliteNumber).Caption = CStr(signalToNoiseRatio)

'grey out those not in use
LabelSatEle(satelliteNumber).ForeColor = LabelSatID(satelliteNumber).ForeColor
LabelSatAzi(satelliteNumber).ForeColor = LabelSatID(satelliteNumber).ForeColor
LabelsatSNR(satelliteNumber).ForeColor = LabelSatID(satelliteNumber).ForeColor

Exit Sub

errhand:
   Handle_GPS_error

End Sub

Private Sub gps_SatellitesInViewReceived(ByVal satellitesInView As Integer)

On Error GoTo errhand

Dim satCount As Integer

Label11.Caption = "Satellites in view = " + CStr(satellitesInView)
satsInView = satellitesInView

'only show sats in view
If satsInView <> lastSatsInView Then
    For satCount = 1 To satsInView
        LabelSatID(satCount).Visible = True
        LabelSatEle(satCount).Visible = True
        LabelSatAzi(satCount).Visible = True
        LabelsatSNR(satCount).Visible = True
    Next satCount
    If satsInView < 12 Then
        For satCount = satsInView + 1 To 12
            LabelSatID(satCount).Visible = False
            LabelSatEle(satCount).Visible = False
            LabelSatAzi(satCount).Visible = False
            LabelsatSNR(satCount).Visible = False
        Next satCount
    End If
    lastSatsInView = satsInView

End If

errhand:
'   Handle_GPS_error
   
End Sub

Private Sub gps_SatellitesUsedReceived(ByVal satellitesUsed As Integer)
Label7.Caption = "Satellites in use = " & CStr(satellitesUsed)
End Sub

Private Sub gps_SpeedReceived(ByVal speed As Double)
GPS_speed = speed * 1.852
TextSpeed.Text = CStr(speed * 1.852) 'speed converted from knots to km/hr
End Sub

Private Sub gps_TimeChanged(ByVal Time As String)
GPS_time = Time
LabelTime.Caption = Time
End Sub

Private Sub MSComm1_OnComm()
Dim InBuff As String

If Not GPS_test_loaded Then Exit Sub

         Select Case MSComm1.CommEvent
         ' Handle each event or error by placing
         ' code below each case statement.

         ' This template is found in the Example
         ' section of the OnComm event Help topic
         ' in VB Help.

         ' Errors
            Case comEventBreak   ' A Break was received.
            Case comEventCDTO    ' CD (RLSD) Timeout.
            Case comEventCTSTO   ' CTS Timeout.
            Case comEventDSRTO   ' DSR Timeout.
            Case comEventFrame   ' Framing Error.
            Case comEventOverrun ' Data Lost.
            Case comEventRxOver  ' Receive buffer overflow.
            Case comEventRxParity   ' Parity Error.
            Case comEventTxFull  ' Transmit buffer full.
            Case comEventDCB     ' Unexpected error retrieving DCB]

         ' Events
            Case comEvCD   ' Change in the CD line.
            Case comEvCTS  ' Change in the CTS line.
            Case comEvDSR  ' Change in the DSR line.
            Case comEvRing ' Change in the Ring Indicator.
            Case comEvReceive ' Received RThreshold # of chars.
               
               InBuff = MSComm1.Input
               Call HandleInput(InBuff)
             
            Case comEvSend ' There are SThreshold number of
                           ' characters in the transmit buffer.
            Case comEvEOF  ' An EOF character was found in the
                           ' input stream.
         End Select
         
         GPSNow = Now

End Sub

Public Sub HandleInput(Sinput As String)
Dim cluster() As String
Dim counter As Integer

On Error GoTo errhand

If Not GPS_test_loaded Then Exit Sub

If Left$(Sinput, 1) = "$" Then 'start of string
    inString = Sinput
Else
    inString = inString + Sinput
End If
cluster = Split(inString, vbCrLf)
For counter = 0 To UBound(cluster) - 1
        cluster(counter) = Trim(cluster(counter))
        If cluster(counter) <> vbNullString Then gps.Parse (cluster(counter))
Next counter
Exit Sub
    
errhand:
'   MsgBox "Error in GPStest::HandleInput, error #: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
   Handle_GPS_error
   
End Sub

Public Sub Handle_GPS_error()

   Dim waitime As Long

   'error handler for GPStest -- resets GPS communication
   Unload GPStest
   
   waitime = Timer
   Do Until Timer > waitime + 1
      DoEvents
   Loop
   
   Maps.GPS_connect
End Sub
