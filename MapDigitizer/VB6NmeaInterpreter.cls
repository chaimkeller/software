VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VB6NmeaInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************************
'                           Class VB6NmeaInterpreter
'                          A VB6   NMEA interpreter
'   Initially derived from Jon Person's VB.NET  NMEA interpreter info@gpsdotnet.com
'                    Alex Etchells  2006    a.etchells@gmail.com
'
' differences from Jon Person's VB.NET  NMEA interpreter:-
' 1 lacks international support
' 2 assumes the user will want GMT, British date format and speed in knots. (multiply by 1.85200 to convert to km/hr)
' 3 does not check for speeding, i.e., car speeds in excess of 55mph. '<<<<<<<<this was for driving! -- don't need it for airplanes!
' 4 DateTimeChanged event handler is replaced by separate events for date and time
' 5 Parses    $GPGGA (Fix data) and $GPGLL(Position)
'             in addition to $GPGSA (Active satellites)
'                            $GPGSV (Satellites in View)
'                            $GPRMC (Position and Time)   NMEA sentences
' 6 SatelliteReceived event includes 'satelliteNumber' - the channel (1 to 12) being used for this satellite
' 7 The Following new event handlers
'   PositionFixReceived
'   SatellitesUsedReceived
'   AltitudeReceived
'   AltitudeUnitsReceived
'   GeoidSeparationReceived
'   GeoidSeparationUnitsReceived
'   DGPSageReceived
'   DGPSstationIDReceived
'   AutoManModeReceived
'   Mode3DReceived
'   Ch1SatReceived
'   Ch2SatReceived
'   Ch3SatReceived
'   Ch4SatReceived
'   Ch5SatReceived
'   Ch6SatReceived
'   Ch7SatReceived
'   Ch8SatReceived
'   Ch9SatReceived
'   Ch10SatReceived
'   Ch11SatReceived
'   Ch12SatReceived
'   SatellitesInViewReceived
'********************************************************************************************

Option Explicit

'DECLARATIONS

Dim words() As String
Dim latitude As String
Dim longitude As String
Dim satelliteTime As String
Dim satelliteDate As String
Dim satelliteNumber As Long
Dim pseudoRandomCode As Long
Dim azimuth As Long
Dim elevation As Long
Dim signalToNoiseRatio As Long
Dim positionFix As String
Dim autoManMode As String
Dim mode3D As String
Dim noOfMessages As Long
Dim gPGSVsequence As Long
Dim satellitesInView As Long

'EVENTS

Public Event PositionReceived(ByVal latitude As String, ByVal longitude As String)
Public Event TimeChanged(ByVal Time As String)
Public Event DateChanged(ByVal satDate As String)
Public Event BearingReceived(ByVal Bearing As Double)
Public Event SpeedReceived(ByVal Speed As Double)
Public Event FixObtained()
Public Event FixLost()
Public Event SatelliteReceived(ByVal satelliteNumber As Long, _
                               ByVal pseudoRandomCode As Long, _
                               ByVal azimuth As Long, _
                               ByVal elevation As Long, _
                               ByVal signalToNoiseRatio As Long)
Public Event HDOPReceived(ByVal value As Double)
Public Event VDOPReceived(ByVal value As Double)
Public Event PDOPReceived(ByVal value As Double)
Public Event PositionFixReceived(ByVal positionFix As String)
Public Event SatellitesUsedReceived(ByVal satellitesUsed As Long)
Public Event AltitudeReceived(ByVal altitude As Double)
Public Event AltitudeUnitsReceived(ByVal altitudeUnits As String)
Public Event GeoidSeparationReceived(ByVal geoidSeparation As Double)
Public Event GeoidSeparationUnitsReceived(ByVal geoidSeparationUnits As String)
Public Event DGPSageReceived(ByVal dGPSage As Long)
Public Event DGPSstationIDReceived(ByVal dGPSstationID As String)
Public Event AutoManModeReceived(ByVal autoManMode As String)
Public Event Mode3DReceived(ByVal mode3D As String)
Public Event Ch1SatReceived(ByVal ch1Sat As Long)
Public Event Ch2SatReceived(ByVal ch2Sat As Long)
Public Event Ch3SatReceived(ByVal ch3Sat As Long)
Public Event Ch4SatReceived(ByVal ch4Sat As Long)
Public Event Ch5SatReceived(ByVal ch5Sat As Long)
Public Event Ch6SatReceived(ByVal ch6Sat As Long)
Public Event Ch7SatReceived(ByVal ch7Sat As Long)
Public Event Ch8SatReceived(ByVal ch8Sat As Long)
Public Event Ch9SatReceived(ByVal ch9Sat As Long)
Public Event Ch10SatReceived(ByVal ch10Sat As Long)
Public Event Ch11SatReceived(ByVal ch11Sat As Long)
Public Event Ch12SatReceived(ByVal ch12Sat As Long)
Public Event SatellitesInViewReceived(ByVal satellitesInView As Long)

'FUNCTIONS

' Processes information from the GPS receiver
 Public Function Parse(ByVal sentence As String) As Boolean

On Error Resume Next

'On Error GoTo errhand

'If Timer1_Wait Then Exit Function 'And Not GPS_trial_connect Then Exit Function 'don't accept input when waiting for calculation to finish
'
'If Timer - waitimeTimer1 > 2# * RegressionForm.Timer1.Interval * 1000# Then Exit Function 'inside a loop, get out

  ' Discard the sentence if its checksum does not match calculated checksum
  If Not IsValid(sentence) Then
    Parse = False
    Exit Function
  End If
  ' Divide the sentence into words and Look at the first word to decide where to go next
  Select Case GetWords(sentence)(0)
   Case "$GPGGA"    ' Global positioning system fixed data
    ' Indicate that the sentence was recognised
    Parse = ParseGPGGA(sentence)
    Exit Function
   Case "$GPGLL"    ' Geographic position - latitude / longitude
    ' Indicate that the sentence was recognised
    Parse = ParseGPGLL(sentence)
    Exit Function
   Case "$GPGSA"    '    GNSS DOP and active satellites
    ' Indicate that the sentence was recognised
    Parse = ParseGPGSA(sentence)
    Exit Function
   Case "$GPGSV"        ' A "Satellites in View" sentence was recieved
    ' Indicate that the sentence was recognised
    Parse = ParseGPGSV(sentence)
    Exit Function
   Case "$GPRMC"    ' A "Recommended Minimum" sentence was found!
    ' Indicate that the sentence was recognised
    Parse = ParseGPRMC(sentence)
    Exit Function
   Case Else
    ' Indicate that the sentence was not recognised
    Parse = False
    Exit Function
  End Select
  
  Exit Function
  
'errhand:
''  MsgBox "Error detected in NmealInterpreter::Parse, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'   GPS_Error

 End Function
 
 ' Divides a sentence into individual words
Private Function GetWords(ByVal sentence As String) As String()

On Error Resume Next

'On Error GoTo errhand

  'strip off final * + checksum
  If Len(sentence) > 3 Then sentence = left$(sentence, Len(sentence) - 3)
  ' Divide sentence into individual words
  GetWords = Split(sentence, ",")
  
  Exit Function
  
'errhand:
''  MsgBox "Error detected in NmealInterpreter::GetWords, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'   GPS_Error
   
End Function

Private Function ParseGPGGA(ByVal sentence As String) As Boolean  ' Interprets a $GPGGA message

On Error Resume Next

'On Error GoTo errhand
'
'  If Timer1_Wait Then Exit Function 'don't except messages when waiting for calculation to finish
'
'  If Timer - waitimeTimer1 > 2# * RegressionForm.Timer1.Interval * 1000# Then Exit Function 'inside a loop, get out
  
  ' Divide the sentence into words
  words = GetWords(sentence)
  ' satellite-derived time?
  If words(1) <> vbNullString Then Call GetTime(words(1))
  ' location?
  If words(2) <> vbNullString And words(3) <> vbNullString And words(4) <> vbNullString _
                    And words(5) <> vbNullString Then Call _
                    GetLatLon(words(2), words(3), words(4), words(5))
  'Position Fix
  If words(6) <> vbNullString Then Call GetPositionFix(words(6))
  'Satellites Used
  If words(7) <> vbNullString Then RaiseEvent SatellitesUsedReceived(CInt(words(7)))
  'HDOP
  If words(8) <> vbNullString Then RaiseEvent HDOPReceived(CDbl(words(8)))
  'Altitude
'  If words(9) <> vbNullString Then RaiseEvent AltitudeReceived(CDbl(words(9)))
  'Altitude Units
  If words(10) <> vbNullString Then RaiseEvent AltitudeUnitsReceived(words(10))
  'Geoid Separation
  If words(11) <> vbNullString Then RaiseEvent GeoidSeparationReceived(CDbl(words(11)))
  'Geoid Separation Units
  If words(12) <> vbNullString Then RaiseEvent GeoidSeparationUnitsReceived(words(12))
  'DGPS age
  If words(13) <> vbNullString Then RaiseEvent DGPSageReceived(CInt(words(13)))
  'DGPS Station ID
  If words(14) <> vbNullString Then RaiseEvent DGPSstationIDReceived(words(14))
  
  ParseGPGGA = True
  
  Exit Function
  
'errhand:
''  MsgBox "Error detected in NmealInterpreter::ParseGPGGA, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'  'sever parse error that hangs up communication
'  'therefore unload the event handler and reload it
'  GPS_Error

End Function

Private Function ParseGPGLL(ByVal sentence As String) As Boolean  ' Interprets a $GPGLL message

On Error Resume Next

'On Error GoTo errhand

'If Timer1_Wait Then Exit Function 'don't accept input when in the midst of calculations
'
'If Timer - waitimeTimer1 > 2# * RegressionForm.Timer1.Interval * 1000# Then Exit Function 'inside a loop, get out
  
  ' Divide the sentence into words
  words = GetWords(sentence)
  
  'location?
  If words(1) <> vbNullString And words(2) <> vbNullString And words(3) <> vbNullString _
                    And words(4) <> vbNullString Then Call _
                    GetLatLon(words(1), words(2), words(3), words(4))
  'satellite-derived time?
  If words(5) <> vbNullString Then Call GetTime(words(5))
  'status
  If words(6) <> vbNullString Then Call GetStatus(words(6))
  
  ParseGPGLL = True
  
  Exit Function
  
'errhand:
''  MsgBox "Error detected in NmealInterpreter::ParseGPGLL, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'  GPS_Error
  
End Function

Private Function ParseGPGSA(ByVal sentence As String) As Boolean

On Error Resume Next

'On Error GoTo errhand

'If Timer1_Wait Then Exit Function 'don't accept input when in the midst of calculations
'
'If Timer - waitimeTimer1 > 2# * RegressionForm.Timer1.Interval * 1000# Then Exit Function 'inside a loop, get out

' Divide the sentence into words
words = GetWords(sentence)
'Auto or manual mode
If words(1) <> vbNullString Then GetAutoManMode (words(1))
'3D mode
If words(2) <> vbNullString Then Get3DMode (words(2))
'ch1
If words(3) <> vbNullString Then
    RaiseEvent Ch1SatReceived(CInt(words(3)))
Else
    RaiseEvent Ch1SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch2
If words(4) <> vbNullString Then
    RaiseEvent Ch2SatReceived(CInt(words(4)))
Else
    RaiseEvent Ch2SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch3
If words(5) <> vbNullString Then
    RaiseEvent Ch3SatReceived(CInt(words(5)))
Else
    RaiseEvent Ch3SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch4
If words(6) <> vbNullString Then
    RaiseEvent Ch4SatReceived(CInt(words(6)))
Else
    RaiseEvent Ch4SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch5
If words(7) <> vbNullString Then
    RaiseEvent Ch5SatReceived(CInt(words(7)))
Else
    RaiseEvent Ch5SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch6
If words(8) <> vbNullString Then
    RaiseEvent Ch6SatReceived(CInt(words(8)))
Else
    RaiseEvent Ch6SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch7
If words(9) <> vbNullString Then
    RaiseEvent Ch7SatReceived(CInt(words(9)))
Else
    RaiseEvent Ch7SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch8
If words(10) <> vbNullString Then
    RaiseEvent Ch8SatReceived(CInt(words(10)))
Else
    RaiseEvent Ch8SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch9
If words(11) <> vbNullString Then
    RaiseEvent Ch9SatReceived(CInt(words(11)))
Else
    RaiseEvent Ch9SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch10
If words(12) <> vbNullString Then
    RaiseEvent Ch10SatReceived(CInt(words(12)))
Else
    RaiseEvent Ch10SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch11
If words(13) <> vbNullString Then
    RaiseEvent Ch11SatReceived(CInt(words(13)))
Else
    RaiseEvent Ch11SatReceived(0) ' using 0 to indicate chan not in use
End If
'ch12
If words(14) <> vbNullString Then
    RaiseEvent Ch12SatReceived(CInt(words(14)))
Else
    RaiseEvent Ch12SatReceived(0) ' using 0 to indicate chan not in use
End If

' Update the DOP values
If words(15) <> vbNullString Then RaiseEvent PDOPReceived(CDbl(words(15)))
If words(16) <> vbNullString Then RaiseEvent HDOPReceived(CDbl(words(16)))
If words(17) <> vbNullString Then RaiseEvent VDOPReceived(CDbl(words(17)))
    
ParseGPGSA = True

Exit Function

'errhand:
''  MsgBox "Error detected in NmealInterpreter::ParseGPGSA, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'   GPS_Error

End Function

Private Function ParseGPGSV(ByVal sentence As String) As Boolean

On Error Resume Next

'On Error GoTo errhand

'If Timer1_Wait Then Exit Function 'don't accept input when in the midst of calculations
'
'If Timer - waitimeTimer1 > 2# * RegressionForm.Timer1.Interval * 1000# Then Exit Function 'inside a loop, get out

' Divide the sentence into words
words = GetWords(sentence)
' number of messgaes in complete message
If words(1) <> vbNullString Then
     noOfMessages = CInt(words(1))
'    RaiseEvent NumberOfMessagesReceived(noOfMessages)
End If
' GPGSV sequence
If words(2) <> vbNullString Then
    gPGSVsequence = CInt(words(2))
'    RaiseEvent GPGSVsequenceReceived(gPGSVsequence)
End If
' Satellites in view
If words(3) <> vbNullString Then
    satellitesInView = CInt(words(3))
    RaiseEvent SatellitesInViewReceived(satellitesInView)
End If
  
'Each sentence contains four blocks of satellite information.
'Read each block and report each satellite's information
Dim count As Long
For count = 1 To 4
    'identify satellite
    If gPGSVsequence <> 0 Then satelliteNumber = (gPGSVsequence - 1) * 4 + count
    ' Does the sentence have enough words to analyse?
    If UBound(words) >= (count * 4 + 3) Then
    ' Yes.  Proceed with analysing the block.  Does it contain any information?
    If words(count * 4) <> vbNullString And words(count * 4 + 1) <> vbNullString _
                              And words(count * 4 + 2) <> vbNullString _
                              And words(count * 4 + 3) <> vbNullString Then
        
        ' Yes. Extract satellite information and report it
        pseudoRandomCode = CInt(words(count * 4))
        elevation = CInt(words(count * 4 + 1))
        azimuth = CInt(words(count * 4 + 2))
        signalToNoiseRatio = CInt(words(count * 4 + 3))
        ' Notify of this satellite's information
        RaiseEvent SatelliteReceived(satelliteNumber, pseudoRandomCode, azimuth, _
                                  elevation, signalToNoiseRatio)
    End If
   End If
Next
' Indicate that the sentence was recognised
ParseGPGSV = True

Exit Function

'errhand:
''  MsgBox "Error detected in NmealInterpreter::ParseGPGSV, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'  GPS_Error

End Function

 Private Function ParseGPRMC(ByVal sentence As String) As Boolean  ' Interprets a $GPRMC message
 
On Error Resume Next

'On Error GoTo errhand
 
' If Timer1_Wait Then Exit Function 'don't accept input when in the midst of calculations
'
' If Timer - waitimeTimer1 > 2# * RegressionForm.Timer1.Interval * 1000# Then Exit Function 'inside a loop, get out
  
  ' Divide the sentence into words
  words = GetWords(sentence)
   ' Do we have enough values to parse satellite-derived time?
  If words(1) <> vbNullString Then Call GetTime(words(1))
    ' Does the device currently have a satellite fix?
  If words(2) <> vbNullString Then Call GetStatus(words(2))
    ' Do we have enough values to describe our location?
  If words(3) <> vbNullString And words(4) <> vbNullString And words(5) <> vbNullString _
                    And words(6) <> vbNullString Then Call _
                    GetLatLon(words(3), words(4), words(5), words(6))
   ' Do we have enough information to extract the current speed?
  If words(7) <> vbNullString Then RaiseEvent SpeedReceived(CDbl(words(7))) '<<<<<<<<convert to km/hr or mi/hr
    '  that's in knots     * 1.150779 would give mph
  ' Do we have enough information to extract bearing?
  If words(8) <> vbNullString Then RaiseEvent BearingReceived(CDbl(words(8)))
  'Date
  If words(9) <> vbNullString Then GetDate (words(9))
   ' Indicate that the sentence was recognised

  ParseGPRMC = True
  Exit Function
  
'errhand:
''
''  MsgBox "Error detected in NmealInterpreter::ParseGPRMC, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'   GPS_Error

 End Function
 
  
'------------------original version---------------------------------
' Returns True if a sentence's checksum matches the calculated checksum
' Private Function IsValid(ByVal sentence As String) As Boolean
'  ' Compare the characters after the asterisk to the calculation
'  If Right$(sentence, 2) = GetChecksum(sentence) Then
'    IsValid = True
'  Else
'    IsValid = False
'  End If
'End Function
'-----------------------------------------------------------
Private Function IsValid(ByVal sentence As String) As Boolean
' Compare the characters after the asterisk to the calculation
' Returns True if a sentence's checksum matches the calculated checksum

On Error Resume Next

'On Error GoTo errhand

If Len(sentence) < 3 Then
    IsValid = False
    Exit Function
End If


If Mid$(sentence, Len(sentence) - 2, 1) <> "*" Then
    IsValid = False
    Exit Function
End If



If right$(sentence, 2) = GetChecksum(sentence) Then
    IsValid = True
Else

    IsValid = False
End If

Exit Function

'errhand:
''  MsgBox "Error detected in NmealInterpreter::IsValid, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'   GPS_Error

End Function
 
Private Function GetChecksum(ByVal sentence As String) As String ' Calculates the checksum for a sentence
  
  Dim Character As String
  Dim charCount As Long
  Dim Checksum As Long
  
  On Error Resume Next

'  On Error GoTo errhand
  
'  If Timer1_Wait Then Exit Function 'don't process input when waiting for calculation to finish
'
'  If Timer - waitimeTimer1 > 2# * RegressionForm.Timer1.Interval * 1000# Then Exit Function 'inside a loop, get out
  
  For charCount = 1 To Len(sentence)     ' Loop through all chars to get a checksum
   Character = right$(left$(sentence, charCount), 1)
   Select Case Character
    Case "$"
     ' Ignore the dollar sign
    Case "*"
     ' Stop processing before the asterisk
     Exit For
    Case Else
     ' Is this the first value for the checksum?
     If Checksum = 0 Then
      ' Yes. Set the checksum to the value
      Checksum = val(Asc(Character))
     Else
      ' No. XOR the checksum with this character's value
      Checksum = Checksum Xor val(Asc(Character))
     End If
   End Select
  Next charCount
  ' Return the checksum formatted as a two-character hexadecimal
  GetChecksum = right$("00" & Hex(Checksum), 2)
  
  Exit Function
  
'errhand:
''
''  MsgBox "Error detected in NmealInterpreter::GetChecksum, error number: " & Str$(Err.Number) & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "GPS unit error"
'   GPS_Error
   
  
End Function


'SUBROUTINES

Private Sub GetTime(ByVal word As String)

   On Error Resume Next

'   On Error GoTo errhand:
   
   satelliteTime = left$(word, 2) + ":" + Mid$(word, 3, 2) + ":" + Mid$(word, 5, 2)
   ' Extract milliseconds if it is available
   If Len(word) > 7 Then satelliteTime = satelliteTime + Mid$(word, 7, Len(word) - 6)
   satelliteTime = satelliteTime + " GMT" '<<<<<<<<<<<<adjust to chosen time zone
   RaiseEvent TimeChanged(satelliteTime)
   
Exit Sub
  
'errhand:
'   GPS_Error
   
End Sub

Private Sub GetDate(ByVal word As String)

   On Error Resume Next

'   On Error GoTo errhand:
   
   satelliteDate = left$(word, 2) + "/" + Mid$(word, 3, 2) + "/" + Mid$(word, 5, 2)
   RaiseEvent DateChanged(satelliteDate)
   
Exit Sub
  
'errhand:
'   GPS_Error
   
End Sub

Private Sub GetLatLon(ByVal latWord As String, ByVal latHemi As String, ByVal lonWord As String, ByVal lonHemi As String)

   On Error Resume Next

'   On Error GoTo errhand:

''latitude
'   latitude = Left$(latWord, 2) & "° " 'hours/degrees
'   latitude = latitude & Right$(latWord, Len(latWord) - 2) & "'"      ' Append minutes
'   latitude = latHemi & " " & latitude    ' start with the hemisphere
''longitude
'   longitude = Left$(lonWord, 3) & "° " 'hours/degrees
'   longitude = longitude & Right$(lonWord, Len(lonWord) - 3) & "'"     ' Append minutes
'   longitude = lonHemi & " " & longitude    ' start with the hemisphere
   
   If InStr(latHemi, "N") <> 0 Then
      latitude = Format(left$(latWord, 2) & Format(Trim$(str$(val(right$(latWord, Len(latWord) - 2)) / 60#)), ".0####"), "#0.0####")
   Else
      'negative latitude for Southern Hemisphere
      latitude = Format("-" & left$(latWord, 2) & Format(Trim$(str$(val(right$(latWord, Len(latWord) - 2)) / 60#)), ".0####"), "##0.0####")
      End If
      
   If InStr(lonHemi, "W") <> 0 Then
      'negative longitude for Western Hemisphere
      longitude = Format("-" & left$(lonWord, 3) & Format(Trim$(str$(val(right$(lonWord, Len(lonWord) - 3)) / 60#)), ".0####"), "##0.0####")
   Else
      longitude = Format(left$(lonWord, 3) & Format(Trim$(str$(val(right$(lonWord, Len(lonWord) - 3)) / 60#)), ".0####"), "###0.0####")
      End If

   ' Notify the calling application of the change
   
   RaiseEvent PositionReceived(latitude, longitude)
   
Exit Sub
  
'errhand:
'   GPS_Error
   
End Sub

Private Sub GetPositionFix(ByVal word As String)

On Error Resume Next

'On Error GoTo errhand

Select Case word
    Case "3"
        positionFix = "Valid PPS"
    Case "2"
        positionFix = "Valid DGPS"
    Case "1"
        positionFix = "Valid SPS"
    Case Else
        positionFix = "Invalid"
End Select
RaiseEvent PositionFixReceived(positionFix)

Exit Sub
  
'errhand:
'   GPS_Error
   
End Sub

Private Sub GetStatus(ByVal word As String)

On Error Resume Next

'On Error GoTo errhand

   ' Does the device currently have a satellite fix?
  If word <> vbNullString Then
   Select Case word
    Case "A"
     RaiseEvent FixObtained
    Case "V"
     RaiseEvent FixLost
   End Select
  End If
  
Exit Sub
  
'errhand:
'   GPS_Error
   
End Sub

Private Sub GetAutoManMode(ByVal word As String)

On Error Resume Next

'On Error GoTo errhand

 If word <> vbNullString Then
   Select Case word
    Case "A"
     autoManMode = "Auto 2D/3D"
    Case "M"
     autoManMode = "Forced 2D/3D"
    Case Else
     autoManMode = "Unknown"
   End Select
  End If
  RaiseEvent AutoManModeReceived(autoManMode)
  
Exit Sub
  
'errhand:
'   GPS_Error
   
End Sub

Private Sub Get3DMode(ByVal word As String)

On Error Resume Next

'On Error GoTo errhand

 If word <> vbNullString Then
   Select Case word
    Case "2"
     mode3D = "2D"
    Case "3"
     mode3D = "3D"
    Case Else
     mode3D = "No Fix"
   End Select
  End If
  RaiseEvent Mode3DReceived(mode3D)
  
Exit Sub
  
'errhand:
'   GPS_Error
   
End Sub

Private Sub GPS_Error()

  'error detected in one of the modules
  'handle the error by reseting GPS communication
  Dim waitime As Single
  
  GPStest.MSComm1.PortOpen = False 'Unload GPStest
  
  waitime = Timer
  Do Until Timer > waitime + 1
     DoEvents
  Loop
  
  'reload it
  GPStest.MSComm1.PortOpen = True
  
End Sub
