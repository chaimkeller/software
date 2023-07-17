VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   Caption         =   "Map"
   ClientHeight    =   7875
   ClientLeft      =   5505
   ClientTop       =   3525
   ClientWidth     =   8505
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   8505
   Begin MSComDlg.CommonDialog comdlgJSON 
      Left            =   840
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "JASON addess"
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      ToolTipText     =   "Read Jason coordinates and move map to address' coordinates"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame frmConv 
      Caption         =   "ITM to Geo Conversion"
      Height          =   495
      Left            =   4200
      TabIndex        =   19
      Top             =   1000
      Visible         =   0   'False
      Width           =   4215
      Begin VB.OptionButton optGeo 
         Caption         =   "Molendensky (web)"
         Height          =   195
         Left            =   2040
         TabIndex        =   21
         ToolTipText     =   "Like used in web program"
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optcasgeo 
         Caption         =   "standard JHK"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "use casgeo and geocasc"
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdGPS 
      Height          =   375
      Left            =   2880
      Picture         =   "frmMap.frx":0F4A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Your GPS coordinates"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdHelp 
      Height          =   375
      Left            =   2160
      Picture         =   "frmMap.frx":1C44
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Help"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveMaps 
      Height          =   375
      Left            =   3360
      Picture         =   "frmMap.frx":1D46
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Center Maps and More to above coordinates"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Mapsbut 
      Height          =   375
      Left            =   3720
      Picture         =   "frmMap.frx":2188
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "display map at Mapys & More's center coordinate"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Map Lat/Long"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      ToolTipText     =   "move map to the  inputed coordinates"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "JASON file"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      ToolTipText     =   "Download JASON file from BING REST service"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtLong 
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      ToolTipText     =   "Longitude, Positive for Western Hemisphere"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtLat 
      Height          =   285
      Left            =   6000
      TabIndex        =   11
      ToolTipText     =   "Latitude, positive for North Hemisphere"
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtZipCode 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   960
      TabIndex        =   8
      ToolTipText     =   "Enter City Name"
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtStreet 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Enter street address"
      Top             =   120
      Width           =   4215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label6 
      Caption         =   "Long"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Lat"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Zip Code"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "State (Country)"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Enter state (or country if no state applies)"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "City"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Street"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TypeConv As Integer

Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control
' Save the controls' positions and sizes.
   On Error GoTo SaveSizes_Error

ReDim m_ControlPositions(1 To Controls.count)
i = 1
For Each ctl In Controls
    With m_ControlPositions(i)
        If TypeOf ctl Is Line Then
            .Left = ctl.X1
            .Top = ctl.Y1
            .Width = ctl.X2 - ctl.X1
            .Height = ctl.Y2 - ctl.Y1
        Else
            .Left = ctl.Left
            .Top = ctl.Top
            .Width = ctl.Width
            .Height = ctl.Height
            On Error Resume Next
            .FontSize = ctl.Font.size
            On Error GoTo 0
        End If
    End With
    i = i + 1
Next ctl
' Save the form's size.
m_FormWid = ScaleWidth
m_FormHgt = ScaleHeight

   On Error GoTo 0
   Exit Sub

SaveSizes_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveSizes of Form frmMap"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdGPS_Click
' Author    : chaim
' Date      : 7/11/2022
' Purpose   : Acquire GPS coordinates and move Google map to that position
'---------------------------------------------------------------------------------------
'
Private Sub cmdGPS_Click()

   On Error GoTo cmdGPS_Click_Error

   If GPSconnected Then
      txtLat = Val(GPStest.TextLat)
      txtlon = Val(GPStest.TextLon)
      Maps.Text5.Text = Format(lono, "###0.0#####")
      Maps.Text6.Text = Format(lato, "##0.0#####")
      Mapsbut.value = True
   
   Else
      Call MsgBox("GPS is not avaialbe!" _
                  & vbCrLf & "" _
                  & vbCrLf & "To confiugre it, click on the ""GPS"" menu item above the tool bar" _
                  , vbInformation, "GPS not configured")
      Exit Sub
      End If

   On Error GoTo 0
   Exit Sub

cmdGPS_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGPS_Click of Form frmMap"
End Sub

Private Sub cmdHelp_Click()
  MsgBox "To retrieve coordinates on the BING map," & vbCrLf & _
  "right click on the point of interest," & vbCrLf & _
  "then click on the COPY coordinates link." & vbCrLf & _
  "The BING map will confirm that the coordinates have been saved to the clipboard." & vbCrLf & vbCrLf & _
  "Now you can send those coordinates to the main program via the bullseye button." & vbCrLf & vbCrLf & _
  "(N.b., the position that will be shon the imported maps is only approximate)", vbOKOnly + vbInformation, "Google Map Help"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdMoveMaps_Click
' Author    : chaim
' Date      : 6/3/2022
' Purpose   : Read clipboard where coordintaes should be stored and moves the program map
'---------------------------------------------------------------------------------------
'
Private Sub cmdMoveMaps_Click()
    Dim StrTxt() As String
    Dim N As Long
    Dim E As Long
    Dim lat As Double
    Dim lon As Double
    
   On Error GoTo cmdMoveMaps_Click_Error

    textcoord$ = Clipboard.GetText
    If textcoord = sEmpty Then
       MsgBox "No coordinates found!  Try again" & vbCrLf & vbCrLf & _
       "Click the ""Help"" button for further information.", vbInformation + vbOKOnly, _
       "No coordinates stored to the clipboard!"
       Exit Sub
       End If
       
    StrTxt = Split(textcoord$, ",")
    
    If UBound(StrTxt) < 1 Then
       MsgBox "No coordinates found in the clipboard!  Try again" & vbCrLf & vbCrLf & _
       "Click the ""Help"" button for further information.", vbInformation + vbOKOnly, _
       "No coordinates stored to the clipboard!"
       Exit Sub
       End If
       
    If IsNumeric(Val(StrTxt(0))) And IsNumeric(Val(StrTxt(1))) Then
    Else
       MsgBox "No recognizable coordinates found in the clipboard!  Try again" & vbCrLf & vbCrLf & _
       "Click the ""Help"" button for further information.", vbInformation + vbOKOnly, _
       "No coordinates stored to the clipboard!"
       Exit Sub
       End If
       
    If world Then
        txtLat.Text = StrTxt(0)
        txtLong.Text = StrTxt(1)
        Maps.Text6.Text = txtLat.Text
        Maps.Text5.Text = txtLong.Text
        Call goto_click
    Else
        If TypeConv = 1 Then
            'convert geo corrdinates to ITM
            txtLat.Text = StrTxt(0)
            txtLong.Text = StrTxt(1)
            lg = Val(txtLong.Text)
            lt = Val(txtLat.Text)
    '        ggpscorrection = False
            Call GEOCASC(lt, lg, ITMx, ITMy)
            Maps.Text6.Text = ITMx * 0.001
            Maps.Text5.Text = ITMy * 0.001
            Call goto_click
        ElseIf TypeConv = 2 Then
            lat = Val(txtLat.Text)
            lon = Val(txtLong.Text)
            'convert lat lon to ITM
            Call wgs842ics(lat, lon, N, E)
            kmy = N
            kmx = E
            Maps.Text5.Text = kmx
            Maps.Text6.Text = kmy
           End If
        End If
        


   On Error GoTo 0
   Exit Sub

cmdMoveMaps_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMoveMaps_Click of Form frmMap"

End Sub

Private Sub Command1_Click()
Dim street As String
Dim city As String
Dim state As String
Dim zip As String
Dim queryAddress As String
Dim BingKeyCoad As String

If Trim$(txtStreet.Text) = "" And Trim$(txtCity.Text) = "" And Trim$(txtState.Text) = "" Then
   MsgBox "You first need to add an address to the adress boxes", vbOKOnly + vbInformation, "Maps&More"
   Exit Sub
   End If

BingKeyCoad = "&key=AnImm57iA90PhX2Ou3jl7l5o-PxhM0bazl9l5yCePLauUP55_MesJClgmRte5ch0"

'queryAddress = "http://google.com/maps?q="
'queryAddress = "https://www.bing.com/maps?cp="
queryAddress = "http://dev.virtualearth.net/REST/v1/Locations?countryRegion=&locality=&addressLine="
' build street part of query string
If txtStreet.Text <> "" Then
    street = txtStreet.Text
    If Trim$(street) = sEmpty Then
       'Call MsgBox("Enter a valid street address in the box provided!", vbInformation, "Google Map Interface")
       Call MsgBox("Enter a valid street address in the box provided!", vbInformation, "Bing Map Interface")
       Exit Sub
       End If
    'queryAddress = queryAddress & street + "," & "+"
    queryAddress = queryAddress & street + ","

End If
' build city part of query string
If txtCity.Text <> "" Then
    city = txtCity.Text
    If Trim$(city) = sEmpty Then
       'Call MsgBox("Enter a valid city in the box provided!", vbInformation, "Google Map Interface")
       Call MsgBox("Enter a valid city in the box provided!", vbInformation, "Bing Map Interface")
       Exit Sub
       End If
'    queryAddress = queryAddress & city + "," & "+"
    queryAddress = queryAddress & city + ","
End If
' build state part of query string
If txtState.Text <> "" Then
    state = txtState.Text
'    queryAddress = queryAddress & state + "," & "+"
    queryAddress = queryAddress & state + ","
End If
' build zip code part of query string
If txtZipCode.Text <> "" Then
    zip = txtZipCode.Text
    queryAddress = queryAddress & zip
End If
' pass the url with the query string to web browser control
queryAddress = queryAddress & BingKeyCoad

WebBrowser1.Navigate queryAddress
'source: https://www.tek-tips.com/viewthread.cfm?qid=1807492
'shows route on static map
'WebBrowser1.Navigate "http://dev.virtualearth.net/REST/v1/Imagery/Map/Road/Routes?wp.0=Seattle,WA;64;1&wp.1=Redmond,WA;66;2&key=AnImm57iA90PhX2Ou3jl7l5o-PxhM0bazl9l5yCePLauUP55_MesJClgmRte5ch0"
'WebBrowser1.Navigate "http://dev.virtualearth.net/REST/v1/Locations?countryRegion={countryRegion}&adminDistrict={adminDistrict}&locality={locality}&postalCode={postalCode}&addressLine={addressLine}&userLocation={userLocation}&userIp={userIp}&usermapView={usermapView}&includeNeighborhood={includeNeighborhood}&maxResults={maxResults}&key={BingMapsKey}"

'downloads json!
WebBrowser1.Navigate queryAddress
'WebBrowser1.Navigate "http://dev.virtualearth.net/REST/v1/Locations?countryRegion=&locality=&addressLine=12531 Collins St., Valley Village, CA&key=AnImm57iA90PhX2Ou3jl7l5o-PxhM0bazl9l5yCePLauUP55_MesJClgmRte5ch0"

'following don't work
'WebBrowser1.Navigate "http://dev.virtualearth.net/REST/v1/Imagery/Map/Road/&addressLine=12531 Collins St., Valley Village, CA&key=AnImm57iA90PhX2Ou3jl7l5o-PxhM0bazl9l5yCePLauUP55_MesJClgmRte5ch0"
'WebBrowser1.Navigate "https://dev.virtualearth.net/REST/v1/Imagery/Map/Streetside/12531 Collins St., Valley Village, CA;&key=AnImm57iA90PhX2Ou3jl7l5o-PxhM0bazl9l5yCePLauUP55_MesJClgmRte5ch0"

End Sub

Private Sub Command2_Click()
If txtLat.Text = "" Or txtLong.Text = "" Then
    MsgBox "Supply a latitude and longitude value.", vbOKOnly, "Missing Data"
End If
Dim lat As String
Dim lon As String
Dim queryAddress As String
'https://www.bing.com/maps/?cp=32.093046%7E34.784775&lvl=15.0
'https://www.bing.com/maps/?cp=23.523462%7E-50.976065&lvl=3.0
'querAddress = "http:/google.com/maps?q="
queryAddress = "https://www.bing.com/maps?cp="
If txtLat.Text <> "" Then
    lat = txtLat.Text
    'queryAddress = queryAddress & lat + "%2C"
    queryAddress = queryAddress & lat + "%7E"
End If
' build longitude part of query string
If txtLong.Text <> "" Then
    lon = txtLong.Text
    queryAddress = queryAddress & lon
End If
WebBrowser1.Navigate queryAddress & "&lvl=15.0"
End Sub

Private Sub Command3_Click()
  'open JASON file and read the coordinates
   On Error GoTo Command3_Click_Error
   
   Dim StrngSplit() As String

  With comdlgJSON
    .CancelError = True
    .Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*|"
    .FilterIndex = 1
    .FileName = App.Path + "\*.json"
    .ShowOpen
    JASONfile$ = .FileName
  End With
  
  FileJSN% = FreeFile
  Open JASONfile$ For Input As #FileJSN%
  'search for string: "coordinates":"
  Line Input #FileJSN%, doclin$
  pos% = InStr(doclin$, """coordinates""")
  If pos% > 0 Then
     pos2% = InStr(pos% + 15, doclin$, "]},")
     If pos2% > 0 Then
        CoordStr$ = Mid$(doclin$, pos% + 15, pos2% - pos% - 15)
        StrngSplit = Split(CoordStr$, ",")
        If UBound(StrngSplit) > 0 Then
           txtLat = StrngSplit(0)
           txtLong = StrngSplit(1)
           Command2.value = True 'move the BING map to those coordinates
           End If
        End If
    End If
  
   On Error GoTo 0
   Exit Sub

Command3_Click_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command3_Click of Form frmMap"
End Sub

Private Sub Form_Load()
SaveSizes
GoogleMapVis = True
WebBrowser1.Silent = True

If world Then
   frmConv.Visible = False
Else
   frmConv.Visible = True
   End If
   
If TypeConv = 0 Then TypeConv = 1
   
'ggpscorrection = False
'ret = SetWindowPos(frmMap.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Sub

Private Sub Form_Resize()
ResizeControls
End Sub

Private Sub ResizeControls()
Dim i As Integer
Dim ctl As Control
Dim x_scale As Single
Dim y_scale As Single
' Don't bother if we are minimized.
   On Error GoTo ResizeControls_Error

If WindowState = vbMinimized Then Exit Sub
' Get the form's current scale factors.
x_scale = ScaleWidth / m_FormWid
y_scale = ScaleHeight / m_FormHgt
' Position the controls.
i = 1
For Each ctl In Controls
    With m_ControlPositions(i)
        If TypeOf ctl Is Line Then
            ctl.X1 = x_scale * .Left
            ctl.Y1 = y_scale * .Top
            ctl.X2 = ctl.X1 + x_scale * .Width
            ctl.Y2 = ctl.Y1 + y_scale * .Height
        Else
            ctl.Left = x_scale * .Left
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            If Not (TypeOf ctl Is ComboBox) Then
                ' Cannot change height of ComboBoxes.
                ctl.Height = y_scale * .Height
            End If
            On Error Resume Next
            ctl.Font.size = y_scale * .FontSize
            On Error GoTo 0
        End If
    End With
    i = i + 1
Next ctl

   On Error GoTo 0
   Exit Sub

ResizeControls_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ResizeControls of Form frmMap"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  GoogleMapVis = False
  tblbuttons(30) = 0
  Maps.Toolbar1.Buttons(30).value = tbrUnpressed
  Set frmMap = Nothing
End Sub

Private Sub Mapsbut_Click()

Dim lt2 As Double, lg2 As Double

   If world Then
      txtLat = Maps.Text6
      txtLong = Maps.Text5
      Command2.value = True
   Else
      'convert ITM to geo
      If InStr(Maps.Text5, ".") And Val(Maps.Text5) < 1000 Then
         lon1 = Val(Maps.Text5) * 1000
      Else
         lon1 = Val(Maps.Text5)
         End If
      If InStr(Maps.Text6, ".") And Val(Maps.Text6) < 1000 Then
         lat1 = 1000000 + Val(Maps.Text6) * 1000
      Else
         lat1 = Val(Maps.Text6)
         End If
         
      'convert itm coorinates to wgs34 geo
      If TypeConv = 2 Then
         Call ics2wgs84(CLng(lat1), CLng(lon1), lt2, lg2)
         txtLong = lg2
         txtLat = lt2
         Command2.value = True
      ElseIf TypeConv = 1 Then
         Call casgeo(lon1, lat1, lg, lt)
         txtLong = -lg
         txtLat = lt
         Command2.value = True
         End If
         
       End If
      
End Sub

Private Sub optcasgeo_Click()
   TypeConv = 1
End Sub

Private Sub optGeo_Click()
   TypeConv = 2
End Sub
