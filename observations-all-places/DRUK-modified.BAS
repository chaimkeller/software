Attribute VB_Name = "Module1"
Public cd As Double, ob As Double, mp As Double, mc As Double, ap As Double, ac As Double
Public ms As Double, ec As Double, e2c As Double, lr As Double, d As Double
Public ch As Double, P As Double, T As Double, aas As Double, pi As Double, pi2 As Double
Public Initialize As Boolean, HeightLoop As Boolean
Public MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer, lat As Double, lon As Double


Public weather%

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

Public Function FNarsin(x As Double) As Double
    FNarsin = Atn(x / Sqr(-x * x + 1))
End Function

Public Function FNarco(x As Double) As Double
    FNarco = -Atn(x / Sqr(-x * x + 1)) + pi / 2
End Function

Public Function FNms(x As Double) As Double
    FNms = mp + mc * x
End Function

Public Function FNaas(x As Double) As Double
    FNaas = ap + ac * x
End Function
Public Function FNes(aas As Double) As Double
    FNes = ms + ec * Sin(aas) + e2c * Sin(2 * aas)
End Function
Public Function FNha(x As Double) As Double
    FNha = FNarco((-Tan(lr) * Tan(d)) + (Cos(x) / Cos(lr) / Cos(d))) * ch
End Function
Public Function FNfrsum(x As Double) As Double
    FNfrsum = (P / (T + 273)) * (0.1419 - 0.0073 * x + 0.00005 * x * x) / (1 + 0.3083 * x + 0.01011 * x * x)
End Function
Public Function FNfrwin(x As Double) As Double
    FNfrwin = (P / (T + 273)) * (0.1561 - 0.0082 * x + 0.00006 * x * x) / (1 + 0.3254 * x + 0.01086 * x * x)
End Function
Public Function FNref(x As Double) As Double
    FNref = (P / (T + 273)) * (0.1594 + 0.0196 * x + 0.00002 * x * x) / (1 + 0.505 * x + 0.0845 * x * x)
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
Dim DirPath$

Dim tncols As Long, IKMY&, IKMX&, numrec&, IO%, Tempmode%

FilePathBil = App.Path & "\WorldClim_bil"
If Dir(FilePathBil, vbDirectory) <> sEmpty Then
    DirPath$ = FilePathBil
    FilePathBil = DirPath$
Else
    'first try default
    FilePathBil = "c:\devstudio\vb" & "\WorldClim_bil"
    If Dir(FilePathBil, vbDirectory) <> sEmpty Then
        DirPath$ = FilePathBil
        FilePathBil = DirPath$
    Else
        Call MsgBox("Can't find the bil directory at the following location:" _
                    & vbCrLf & FilePathBil _
                    & vbCrLf & vbCrLf & "Please select the correct direcotry location." _
                    , vbExclamation, "Missing bil file directory")
        DirPath$ = BrowseForFolder(Drukfrm.hWnd, "Choose Directory")
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
       filein% = FreeFile
       Open FileNameBil For Binary As #filein%
   
        Y = lat
        x = lon
        
        IKMY& = CLng((ULYMAP - Y) / YDIM) + 1
        IKMX& = CLng((x - ULXMAP) / XDIM) + 1
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
Public Function SortArray(ByRef TheArray As Variant)
Sorted = False
Do While Not Sorted
    Sorted = True
For x = 0 To UBound(TheArray) - 1
    If TheArray(x) > TheArray(x + 1) Then
        Temp = TheArray(x + 1)
        TheArray(x + 1) = TheArray(x)
        TheArray(x) = Temp
        Sorted = False
    End If
Next x
Loop
End Function
' min and max are the minimum and maximum indexes
' of the items that might still be out of order.
'copied from http://www.vb-helper.com/tut1.htm and modified to handle two dimensional double value array
Public Sub BubbleSort(list() As Double, ByVal min As Integer, _
    ByVal max As Integer)
Dim last_swap As Integer
Dim i As Integer
Dim j As Integer
Dim tmp1 As Double, tmp2 As Double

    ' Repeat until we are done.
    Do While min < max
        ' Bubble up.
        last_swap = min - 1
        ' For i = min + 1 To max
        i = min + 1
        Do While i <= max
            ' Find a bubble.
            If list(0, i - 1) > list(0, i) Then
                ' See where to drop the bubble.
                tmp1 = list(0, i - 1)
                tmp2 = list(1, i - 1)
                j = i
                Do
                    list(0, j - 1) = list(0, j)
                    list(1, j - 1) = list(1, j)
                    j = j + 1
                    If j > max Then Exit Do
                Loop While list(0, j) < tmp1
                list(0, j - 1) = tmp1
                list(1, j - 1) = tmp2
                last_swap = j - 1
                i = j + 1
            Else
                i = i + 1
            End If
        Loop
        ' Update max.
        max = last_swap - 1

        ' Bubble down.
        last_swap = max + 1
        ' For i = max - 1 To min Step -1
        i = max - 1
        Do While i >= min
            ' Find a bubble.
            If list(0, i + 1) < list(0, i) Then
                ' See where to drop the bubble.
                tmp1 = list(0, i + 1)
                tmp2 = list(1, i + 1)
                j = i
                Do
                    list(0, j + 1) = list(0, j)
                    list(1, j + 1) = list(1, j)
                    j = j - 1
                    If j < min Then Exit Do
                Loop While list(0, j) > tmp1
                list(0, j + 1) = tmp1
                list(1, j + 1) = tmp2
                last_swap = j + 1
                i = j - 1
            Else
                i = i - 1
            End If
        Loop
        ' Update min.
        min = last_swap + 1
    Loop
End Sub

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

'---------------------------------------------------------------------------------------
' Procedure : DayNumToDate
' Author    : Dr-John-K-Hall
' Date      : 4/23/2020
' Purpose   : converts daynumber to Gregorian date
' Input     : length of year (yl), daynumber (dydy), year (yrdy)
'           : mode = 0 for US date order = mm/dd/yyyy
'           :      = 1 for EU date order = dd-mmm-yyyy
'---------------------------------------------------------------------------------------
'
Public Function DayNumToDate(yl As Integer, dydy As Integer, yrdy As Integer, mode) As String
    nleap = 0
    If yl = 366 Then nleap = 1
    If dydy <= 31 Then
       If mode = 0 Then
          DayNumToDate = "Jan" & "/" & Trim$(Str$(dydy)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy)) & "-Jan-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 31 And dydy <= 59 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Feb" & "/" & Trim$(Str$(dydy - 31)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 31)) & "-Feb-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 59 + nleap And dydy <= 90 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Mar" & "/" & Trim$(Str$(dydy - 59 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 59 - nleap)) & "-Mar-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 90 + nleap And dydy <= 120 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Apr" & "/" & Trim$(Str$(dydy - 90 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 90 - nleap)) & "-Apr-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 120 + nleap And dydy <= 151 + nleap Then
       If mode = 0 Then
          DayNumToDate = "May" & "/" & Trim$(Str$(dydy - 120 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 120 - nleap)) & "-May-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 151 + nleap And dydy <= 181 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Jun" & "/" & Trim$(Str$(dydy - 151 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 151 - nleap)) & "-Jun-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 181 + nleap And dydy <= 212 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Jul" & "/" & Trim$(Str$(dydy - 181 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 181 - nleap)) & "-Jul-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 212 + nleap And dydy <= 243 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Aug" & "/" & Trim$(Str$(dydy - 212 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 212 - nleap)) & "-Aug-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 243 + nleap And dydy <= 273 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Sep" & "/" & Trim$(Str$(dydy - 243 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 243 - nleap)) & "-Sep-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 273 + nleap And dydy <= 304 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Oct" & "/" & Trim$(Str$(dydy - 273 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 273 - nleap)) & "-Oct-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 304 + nleap And dydy <= 334 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Nov" & "/" & Trim$(Str$(dydy - 304 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 304 - nleap)) & "-Nov-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
    ElseIf dydy > 334 + nleap Then
       If mode = 0 Then
          DayNumToDate = "Dec" & "/" & Trim$(Str$(dydy - 334 - nleap)) & "/" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "mm/dd/yyyy")
       ElseIf mode = 1 Then
          DayNumToDate = Trim$(Str$(dydy - 334 - nleap)) & "-Dec-" & Trim$(Str$(yrdy))
          DayNumToDate = Format(DayNumToDate, "dd-mmm-yyyy")
          End If
       End If
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
                           
   Dim Coef(4, 10) As Double, Ref As Double, VA As Double
   Dim CA(10) As Double
   Dim vbweps(6) As Double, vdwref(6) As Double, ier As Integer, TK As Double
   Dim VDWSF As Double, VDWALT As Double, lnhgt As Double, pi As Double, cd As Double
   Dim VDWEPSR As Double, VDWREFR As Double
   Dim sumref(7) As Double, winref(7) As Double, sumrefo As Double, winrefo As Double
   Dim weather%, ns1 As Integer, ns2 As Integer, ns3 As Integer, ns4 As Integer
   Dim MinT(12) As Integer, AvgT(12) As Integer, MaxT(12) As Integer
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
       Ref = 0: eps = 0
       If height <= 0 Then GoTo 690
       Ref = Exp(winref(4) + winref(5) * lnhgt + _
           winref(6) * lnhgt * lnhgt + winref(7) * lnhgt * lnhgt * lnhgt)
'           ref = ((winref(2, n2%) - winref(2, n1%)) / 2) * (hgt - h1) + winref(2, n1%)
       eps = Exp(winref(0) + winref(1) * lnhgt + _
            winref(2) * lnhgt * lnhgt + winref(3) * lnhgt * lnhgt * lnhgt)
'           eps = ((winref(1, n2%) - winref(1, n1%)) / 2) * (hgt - h1) + winref(1, n1%)
690    air = 90 * cd + (eps + Ref + winrefo) / 1000
       AirMenatRefDip = eps + Ref + winrefo
'       lblMenatAir.Caption = AirMenatRefDip & " mrad"
    ElseIf dy > ns1 And dy < ns2 Then
       Ref = 0: eps = 0
       If height <= 0 Then GoTo 695
       Ref = Exp(sumref(4) + sumref(5) * lnhgt + _
           sumref(6) * lnhgt * lnhgt + sumref(7) * lnhgt * lnhgt * lnhgt)
       'ref = ((sumref(2, n2%) - sumref(2, n1%)) / 2) * (hgt - h1) + sumref(2, n1%)
       eps = Exp(sumref(0) + sumref(1) * lnhgt + _
            sumref(2) * lnhgt * lnhgt + sumref(3) * lnhgt * lnhgt * lnhgt)
'           eps = ((sumref(1, n2%) - sumref(1, n1%)) / 2) * (hgt - h1) + sumref(1, n1%)
695    air = 90 * cd + (eps + Ref + sumrefo) / 1000
       AirMenatRefDip = eps + Ref + sumrefo
'       lblMenatAir.Caption = AirMenatRefDip & " mrad"
       End If
    End If
    
  If weather% = 5 Then
    'use Rabbi Druk's observation place latitude and longitude
    'lon = 35.237435642287 '81333572129 '-35.238456 '5 'longitude at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
    'lat = 31.748552568177 '8959288296 '31.749942 'latitude at at Rabbi Druk's shul at Armon HaNaziv N. (place of observations)
    
    
    Call Temperatures(lat, lon, MinT, AvgT, MaxT, ier)
    
       'determine the minimum and average temperature for this day for current place
       'use Meeus's forumula p. 66 to convert daynumber to month,
       'no need to interpolate between temepratures -- that is overkill
       'take year as regular
       yl = 365
       dy = DayNumber
       k% = 2
       If (yl = 366) Then k% = 1
       MMonth% = Int(9 * (k% + dy) / 275 + 0.98)
    
''       If optMin.Value = True Then
'          TK = MaxT(MMonth%) + 273.15 'mean minimum temperature for this month in degrees Kelvin
''       Else
''          TK = AT(MMonth%) + 273.15 'mean minimum temperature for this month in degrees Kelvin
''          End If
'
'       'calculate van der Werf temperature scaling factor for refraction
'       VDWSF = (288.15 / TK) ^ 1.7081
'       'calculate van der Werf scaling factor for view angles
'       VDWALT = (288.15 / TK) ^ 0.69
       
 '/////////////////////////////////////////////
         'TK = MT(m%) + 273.15 'mean minimum temperature for this month in degrees Kelvin
        TK = MaxT(MMonth%) + 273.15 'use maximum temperature for zemanim, etc.
        'calculate van der Werf temperature scaling factor for refraction
        VDWSF = (288.15 / TK) ^ 1.687 ' 1.7081
        'calculate van der Werf scaling factor for view angles
        VDWALT = (288.15 / TK) ^ 0.69
        'calculate VDW scaling factor for eps
        VDWEPSR = (288.15 / TK) ^ -0.2
        'calculate VDW scaling for ref
        VDWREFR = (288.15 / TK) ^ 2.18
               
'static doublereal c_b175 = 1.687; //1.7081; //1.686701538132792; //1.7081;
'static doublereal c_b176 = .69;
'static doublereal c_b177 = 73.;
'static doublereal c_b178 = 9.56267125268496f; //9.572702286470884f; //9.56267125268496f;
'static doublereal c_b179 = 2.18; //exponent for ref  (need expln why it turned out different than c_b075 ????)
'                                 //but makes dif for Jerusalem astronomical is no more than 2 seconds.
'static doublereal c_b180 = 0.5; //reduced vdw exponent for astronomical altitudes
'static doublereal c_b181 = -0.2; //exponent for eps
'
'        d__2 = 288.15f / tk;
'        vdwsf = pow_dd(&d__2, &c_b175);
'/*          calculate van der Werf scaling factor for view angles */
'        d__2 = 288.15f / tk;
'        vdwalt = pow_dd(&d__2, &c_b176);
'        //scaling law for ref
'        d__2 = 288.15f / tk;
'        vdwrefr = pow_dd(&d__2, &c_b179);
'        //scaling for astronomical times
'        d__2 = 288.15f / tk;
'        vbwast = pow_dd(&d__2, &c_b180);
'        //scaling for local ray altitude (compliment of zenith angle)
'        d__2 = 288.15f / tk;
'        vbweps = pow_dd(&d__2, &c_b181);
'
'        air = cd * 90. + (vbweps * eps + vdwrefr * ref + vdwsf * refrac1) / 1e3;

 
 '/////////////////////////////////////////
    
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
           
           Ref = CA(0)
           
    Else
        
        'calculate for range of view angles
        Ref = 0#
        For i = 0 To 10
            CA(i) = Coef(0, i) + Coef(1, i) * height + Coef(2, i) * (height ^ 2) + Coef(3, i) * (height ^ 3) + Coef(4, i) * (height ^ 4)
        Next i
        For i = 0 To 10
            Ref = Ref + CA(i) * VA ^ i
        Next i
         
        End If
               
   RefFromHorToInfExp = Ref 'calculated from exponent
   CalcVDWRef = Ref
   Exit Function
'        lblref1.Caption = Str(Ref) & " mrad"
               
                
    '    now calculate refraction from the observer's height to the horizon, as well as the geometric dip angle, then the total refraction and dip.
    '/*  All refraction terms have units of mrad */
    
'     d__2 = 288.15 / TK
'     VDWSF = d__2 ^ 1.7081
''/*          calculate van der Werf scaling factor for view angles */
'     VDWALT = d__2 ^ 0.69
     
     If (height <= 0#) Then GoTo L690
     lnhgt = Log(height * 0.001)
    
     ref2 = Exp(vdwref(0) + vdwref(1) * lnhgt + vdwref(2) * lnhgt * _
         lnhgt + vdwref(3) * lnhgt * lnhgt * lnhgt + vdwref(4) _
         * lnhgt * lnhgt * lnhgt * lnhgt + vdwref(5) * lnhgt * _
         lnhgt * lnhgt * lnhgt * lnhgt)
     eps = Exp(vbweps(0) + vbweps(1) * lnhgt + vbweps(2) * lnhgt * _
         lnhgt + vbweps(3) * lnhgt * lnhgt * lnhgt + vbweps(4) _
         * lnhgt * lnhgt * lnhgt * lnhgt)

'        lbleps.Caption = eps & " mrad"
'        lblref2.Caption = VDWSF * ref2 & " mrad"
            
        RefFromHorToInfFit = VDWREFR * ref2 '+ RefFromHorToInf  '& " mrad" 'two components:  VDWSF*ref2 is contrib. from height to horizon, RefFromHorToInf is contribution from horizon to infinity
        CalcVDWRef = 0.5 * (RefFromHorToInfExp + RefFromHorToInfFit) 'take weighted average of them
        Exit Function
L690:
        a1 = 0#
        air = cd * 90# + (VDWREFR * eps + VDWREFR * ref2 + VDWSF * 9.56267125268496) / 1000#
        
        TotalRefWithDip = VDWREFR * eps + VDWREFR * ref2 + VDWSF * 9.56267125268496 '& " mrad"
        
        a1 = (ref2 + Ref) / 1000#
'/*         leave a1 in radians */
        a1 = Atn(Tan(a1) * VDWALT)
        
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
        air = cd * 90# + (eps + refFromExp + VDWSF * 9.56267125268496) / 1000#
        
        TotalGeoRef = eps + refFromExp + VDWSF * 9.56267125268496 '& " mrad"
        Exit Function
        End If

   On Error GoTo 0
   Exit Function

CalcVDWRef_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CalcVDWRef of Module modBurtonAstRef"

End Function


