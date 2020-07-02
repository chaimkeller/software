VERSION 5.00
Begin VB.Form GPSsetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Com-Port Setup for NMEA 0183 GPS"
   ClientHeight    =   2835
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4320
   Icon            =   "GPSsetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer_bubble 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3240
      Top             =   1560
   End
   Begin VB.Frame frmND100 
      Caption         =   "(default GPS device)"
      Height          =   560
      Left            =   120
      TabIndex        =   9
      Top             =   40
      Width           =   4095
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Height          =   195
         Left            =   3120
         TabIndex        =   12
         ToolTipText     =   "Other device, choose its baud raute below"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optND100S 
         Caption         =   "ND-100S/105C"
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         ToolTipText     =   "ND-100S  SiRF-II or ND-105C Micro GPS USB Receiver Dongle (baud: 4800 bps)"
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optND100 
         Caption         =   "ND-100"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "ND-100 GPS USB Dongle (baud: 38400 bps)"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmExplain 
      Caption         =   "Preset Settings"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1620
      Width           =   4095
      Begin VB.Label lblExplain 
         Alignment       =   2  'Center
         Caption         =   "Parity: N (none), Data Bits:: 8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   280
         Width           =   3855
      End
   End
   Begin VB.Frame frmSettings 
      Caption         =   "Settings"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   620
      Width           =   4095
      Begin VB.ComboBox cboCom 
         Height          =   315
         Left            =   3000
         TabIndex        =   8
         ToolTipText     =   "Available Com ports (choose one)"
         Top             =   360
         Width           =   900
      End
      Begin VB.ComboBox cmbBaud 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Choose Baud Rate"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Baud Rate:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   200
         TabIndex        =   4
         Top             =   380
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan Com Port"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Pick the baud rate, and click to scan for a com port"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Accept displayed values"
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "GPSsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim TT1 As New CBalloonToolTip                             '//On Demand tooltip
Dim TT2 As New CBalloonToolTip                             '//mouse over tooltip


Private Sub CancelButton_Click()

   GPSSetupVis = False
   Unload Me

End Sub

Private Sub cboCom_click()

   Dim DeviceTypeNum As Integer
   
   DeviceTypeNum = val(GetSetting(App.Title, "Settings", "GPS_device_name"))
   
   If ComPort% = 0 And DeviceTypeNum = 3 Then
    
        ComPort% = val(Mid$(GPSsetup.cboCom.Text, 4, Len(GPSsetup.cboCom.Text) - 3)) 'Val(Mid$(GPSsetup.cboCom.List(GPSsetup.cboCom.ListIndex), 4, Len(GPSsetup.cboCom.List(GPSsetup.cboCom.ListIndex)) - 3))
        If IsNumeric(ComPort%) Then SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%
        
        End If

End Sub

Private Sub cmbBaud_click()

  Dim DeviceTypeNum As Integer
     
  DeviceTypeNum = val(GetSetting(App.Title, "Settings", "GPS_device_name"))

   If DeviceTypeNum = 3 And DeviceType_Init Then 'ComPort% = 0 And DeviceTypeNum = 3 Then
       'set COM port
       
        Control_Num = 8
        TT1.Style = TTBalloon
        TT1.Icon = TTIconInfo
        TT1.Title = "GPS COM port"
        TT1.TipText = "Chosse the COM port of your GPS device, and press ''OK''"
        TT1.PopupOnDemand = True
        TT1.VisibleTime = 6000                                 'After 6 Seconds tooltip will go away
        TT1.CreateToolTip cboCom.hwnd
        TT1.Show GPSsetup.cboCom.left / Screen.TwipsPerPixelX - 15, GPSsetup.cboCom.Height / Screen.TwipsPerPixelX + 5 '//In Pixel only
    '    Timer_bubble.Enabled = True 'timer will kill bubble after 6 seconds
        
        End If

End Sub

Private Sub cmdScan_Click()

  Dim i As Long, Max_Com_Port As Integer, waitime As Single
  Dim temp As String
  Dim Lister As New DevLister, PortGUID() As GUID_Storage
  Dim pos1%, pos2%, Last_Port
   
  'try again to find available com ports
  cboCom.Clear
  temp = Lister.GetGUIDByName("Ports", PortGUID)
  If temp = "OK" Then
      temp = Lister.AddToList(PortGUID(0), 1)
      If temp = "OK" Then
          temp = ""
          For i = 1 To Lister.ListCount
              If ProlificGPS Then
                 cboCom.Clear
                 ComPort% = 0
                 If InStr(1, Lister.Item(i).GetName, "Prolific") Then
                    pos1% = InStr(Lister.Item(i).GetName, "COM")
                    pos2% = InStr(pos1% + 3, Lister.Item(i).GetName, ")")
                    ComPort% = val(Mid$(Lister.Item(i).GetName, pos1% + 3, pos2% - pos1% - 3))
                    SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%
                    cboCom.AddItem "COM" & ComPort%
                    cboCom.ListIndex = 0
                    Exit For
                    End If
              Else
                 'list all the COM's
                 pos1% = InStr(Lister.Item(i).GetName, "COM")
                 pos2% = InStr(pos1% + 3, Lister.Item(i).GetName, ")")
                 cboCom.AddItem "COM" & Mid$(Lister.Item(i).GetName, pos1% + 3, pos2% - pos1% - 3)
                 End If
            Next i
        Else
            MsgBox temp, vbCritical, "Error Getting Device List"
        End If
    Else
        MsgBox temp, vbCritical, "Error Finding 'Ports' GUID"
    End If
    
    cboCom.ListIndex = 0
 
  
'  If GPSSimulation Then
'     cboCom.AddItem "Com" & Trim$(Str$(GPSSimulPort))
'     Max_Com_Port = GPSSimulPort
'  Else
'    For i = 1 To MAX_PORT
'       If COMAvailable(i) Then
'          cboCom.AddItem "Com" & Trim$(Str$(i))
'          Max_Com_Port = i
'          End If
'    Next i
'    End If
    
  'set to last available com port
'  cboCom.ListIndex = cboCom.ListCount - 1 'show last available com port -- this is the right one
   
   GPSConnectString = cmbBaud.List(cmbBaud.ListIndex) & ",N,8,1" 'scan for GPS COM port using the chosen Baud rate
   Load GPStest
   GDMDIform.GPS_timer.Enabled = True
   lblExplain.ForeColor = &H400000
   lblExplain.Caption = "Testing Baud Rate: " & cmbBaud.List(cmbBaud.ListIndex) & ", " & cboCom.Text
   
   lblExplain.Refresh
   waitime = Timer
   Do Until Timer > waitime + 2
   Loop
      

End Sub

Private Sub Form_Load()

  Dim Ret As Long, i As Long, Max_Com_Port As Integer, DeviceTypeNum As Integer ', BaudRate As Long
  Dim temp As String
  Dim Lister As New DevLister, PortGUID() As GUID_Storage
  Dim pos1%, pos2%
   
   Dim numPort%, FoundPort%
   
   GPSSetupVis = True

   Ret = SetWindowPos(GPSsetup.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   
   cmbBaud.AddItem 1200
   cmbBaud.AddItem 2400
   cmbBaud.AddItem 4800
   cmbBaud.AddItem 9600
   cmbBaud.AddItem 14400
   cmbBaud.AddItem 19200
   cmbBaud.AddItem 38400
   cmbBaud.AddItem 56000
   cmbBaud.AddItem 128000
   cmbBaud.AddItem 256000
   
   cmbBaud.ListIndex = 0 'set to first item
   
   DeviceTypeNum = val(GetSetting(App.Title, "Settings", "GPS_device_name"))
   If DeviceTypeNum = 0 Then 'not defined yet
'      'ask the user
   ElseIf DeviceTypeNum = 1 Then 'ND-100
      optND100.value = True
   ElseIf DeviceTypeNum = 2 Then 'ND-100S
      optND100S.value = True
   ElseIf DeviceTypeNum = 3 Then 'other, but defined
      optOther.value = True
      End If
   
  Max_Com_Port = MAX_PORT
   
  'find available com ports
  cboCom.Clear
  
'  If GPSSimulation Then
'     cboCom.AddItem "Com" & Trim$(Str$(GPSSimulPort))
'     Max_Com_Port = GPSSimulPort
'  Else
  numPort% = 0
  temp = Lister.GetGUIDByName("Ports", PortGUID)
  If temp = "OK" Then
      temp = Lister.AddToList(PortGUID(0), 1)
      If temp = "OK" Then
          temp = ""
          For i = 1 To Lister.ListCount
              If ProlificGPS Then
                 cboCom.Clear
                 ComPort% = 0
                 If InStr(1, Lister.Item(i).GetName, "Prolific") Then
                    pos1% = InStr(Lister.Item(i).GetName, "COM")
                    pos2% = InStr(pos1% + 3, Lister.Item(i).GetName, ")")
                    ComPort% = val(Mid$(Lister.Item(i).GetName, pos1% + 3, pos2% - pos1% - 3))
                    SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%
                    cboCom.AddItem "COM" & ComPort%
                    cboCom.ListIndex = 0
                    Exit For
                    End If
              Else
                 'list all the COM's
                 pos1% = InStr(Lister.Item(i).GetName, "COM")
                 pos2% = InStr(pos1% + 3, Lister.Item(i).GetName, ")")
                 cboCom.AddItem "COM" & Mid$(Lister.Item(i).GetName, pos1% + 3, pos2% - pos1% - 3)
                 numPort% = numPort% + 1
                 If ComPort% <> 0 And i = ComPort% Then
                    FoundPort% = numPort%
                    End If
                 End If
            Next i
            
        If ComPort% <> 0 And FoundPort% <> 0 Then 'show it
           cboCom.ListIndex = FoundPort% - 1
        ElseIf ComPort% <> 0 Then 'show first item
           cboCom.ListIndex = 0
           End If
            
        Else
            MsgBox temp, vbCritical, "Error Getting Device List"
        End If
    Else
        MsgBox temp, vbCritical, "Error Finding 'Ports' GUID"
    End If

'    numPort% = 0
'    For i = 1 To MAX_PORT
'       If COMAvailable(i) Then 'Or i = ComPort% Then
'          cboCom.AddItem "Com" & Trim$(Str$(i))
'          numPort% = numPort% + 1
'          If ComPort% <> 0 And i = ComPort% Then
'             FoundPort% = numPort%
'             End If
'          Max_Com_Port = i
'          End If
'    Next i
'
'    If ComPort% <> 0 Then 'show it
'       cboCom.ListIndex = FoundPort% - 1
'    Else
'       'set to last available com port
'       cboCom.ListIndex = cboCom.ListCount - 1 'show last available com port -- this is the right one
'       End If
       
'    End If
    
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DeviceType_Init = False
   GPSSetupVis = False
   Set GPSsetup = Nothing
End Sub

Private Sub OKButton_Click()

  Dim waitime As Single, Last_Port
  
   On Error GoTo OKButton_Click_Error

  If DeviceType_Init Then 'record baud rate assuming user knows what he is doing
  
     TT1.Destroy 'close the bubble notification
     
     DeviceType_Init = False
     GPSConnectString = cmbBaud.List(cmbBaud.ListIndex) & ",N,8,1" 'communication string
     GPSConnectString0 = GetSetting(App.Title, "Settings", "GPS serial-USB connection string")

    If GPSConnectString <> GPSConnectString0 Then
       GPSConnectString0 = GPSConnectString
       SaveSetting App.Title, "Settings", "GPS serial-USB connection string", GPSConnectString
       End If
       
     GPSsetup.Visible = False
     Exit Sub
     End If

'  Last_Port = Val(Mid$(cboCom.List(cboCom.ListIndex), 4, Len(cboCom.List(cboCom.ListIndex)) - 3))
  GPSOKButton = True
  GPSConnectString = cmbBaud.List(cmbBaud.ListIndex) & ",N,8,1" 'scan for GPS COM port using the chosen Baud rate
  GPSsetup.lblExplain.ForeColor = &HFF&
  GPSsetup.lblExplain = "Establishing GPS communication..."
  
  If Trim$(GPSsetup.cboCom.Text) = sEmpty Then
     GPSsetup.lblExplain.Caption = """Prolific"" USB Device not found on any COM port!"
     Exit Sub
     End If
     
  ComPort% = val(Mid$(GPSsetup.cboCom.Text, 4, Len(GPSsetup.cboCom.Text) - 3)) 'Val(Mid$(GPSsetup.cboCom.List(GPSsetup.cboCom.ListIndex), 4, Len(GPSsetup.cboCom.List(GPSsetup.cboCom.ListIndex)) - 3))
  If IsNumeric(ComPort%) Then SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%
  
  lblExplain.Refresh
  waitime = Timer
  Do Until Timer > waitime + 2
  Loop

  Load GPStest
  

   On Error GoTo 0
   Exit Sub

OKButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OKButton_Click of Form GPSsetup"

End Sub

Private Sub optND100_Click()

Dim DeviceTypeNum As Integer

Dim temp As String, i As Long
Dim Lister As New DevLister, PortGUID() As GUID_Storage
Dim pos1%, pos2%

   On Error GoTo optND100_Click_Error

GDMDIform.Timer_bubble.Interval = 5
GDMDIform.Timer_bubble.Enabled = True 'destroy last bubble notification

GPSConnectString = "38400,N,8,1" 'default baud rate, parity, data bit, stop bit

GPSConnectString0 = GetSetting(App.Title, "Settings", "GPS serial-USB connection string")

If GPSConnectString <> GPSConnectString0 Then
   GPSConnectString0 = GPSConnectString
   SaveSetting App.Title, "Settings", "GPS serial-USB connection string", GPSConnectString
   End If
   
DeviceTypeNum = 1
SaveSetting App.Title, "Settings", "GPS_device_name", DeviceTypeNum

'determine the default baud rate
If InStr(GPSConnectString0, "1200") Then
   cmbBaud.ListIndex = 0
ElseIf InStr(GPSConnectString0, "2400") Then
   cmbBaud.ListIndex = 1
ElseIf InStr(GPSConnectString0, "4800") Then
   cmbBaud.ListIndex = 2
ElseIf InStr(GPSConnectString0, "9600") Then
   cmbBaud.ListIndex = 3
ElseIf InStr(GPSConnectString0, "14400") Then
   cmbBaud.ListIndex = 4
ElseIf InStr(GPSConnectString0, "19200") Then
   cmbBaud.ListIndex = 5
ElseIf InStr(GPSConnectString0, "38400") Then
   cmbBaud.ListIndex = 6
ElseIf InStr(GPSConnectString0, "56000") Then
   cmbBaud.ListIndex = 7
ElseIf InStr(GPSConnectString0, "128000") Then
   cmbBaud.ListIndex = 8
ElseIf InStr(GPSConnectString0, "256000") Then
   cmbBaud.ListIndex = 9
Else 'default
   cmbBaud.ListIndex = 6 '38400 is default
   End If

GPSConnectString = cmbBaud.List(cmbBaud.ListIndex) & ",N,8,1" 'scan for GPS COM port using the chosen Baud rate
'GPSsetup.lblExplain.ForeColor = &HFF&
'GPSsetup.lblExplain = "Establishing GPS communication..."

ProlificGPS = True

'determine COM port connected to the Prolific serial device

  temp = Lister.GetGUIDByName("Ports", PortGUID)
  If temp = "OK" Then
      temp = Lister.AddToList(PortGUID(0), 1)
      If temp = "OK" Then
          temp = ""
          For i = 1 To Lister.ListCount
              If ProlificGPS Then
                 cboCom.Clear
                 ComPort% = 0
                 If InStr(1, Lister.Item(i).GetName, "Prolific") Then
                    pos1% = InStr(Lister.Item(i).GetName, "COM")
                    pos2% = InStr(pos1% + 3, Lister.Item(i).GetName, ")")
                    ComPort% = val(Mid$(Lister.Item(i).GetName, pos1% + 3, pos2% - pos1% - 3))
                    SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%
                    cboCom.AddItem "COM" & ComPort%
                    cboCom.ListIndex = 0
                    Exit For
                    End If
                 End If
            Next i
            
        Else
            MsgBox temp, vbCritical, "Error Getting Device List"
        End If
    Else
        MsgBox temp, vbCritical, "Error Finding 'Ports' GUID"
    End If

    SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%


   If DeviceType_Init Then OKButton.value = True 'record the connect string

   On Error GoTo 0
   Exit Sub

optND100_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optND100_Click of Form GPSsetup"

End Sub

Private Sub optND100S_Click()

Dim DeviceTypeNum As Integer

Dim temp As String, i As Long
Dim Lister As New DevLister, PortGUID() As GUID_Storage
Dim pos1%, pos2%

   On Error GoTo optND100S_Click_Error

GDMDIform.Timer_bubble.Interval = 5
GDMDIform.Timer_bubble.Enabled = True 'destroy last bubble notification

GPSConnectString = "4800,N,8,1" 'default baud rate, parity, data bit, stop bit

GPSConnectString0 = GetSetting(App.Title, "Settings", "GPS serial-USB connection string")

If GPSConnectString <> GPSConnectString0 Then
   GPSConnectString0 = GPSConnectString
   SaveSetting App.Title, "Settings", "GPS serial-USB connection string", GPSConnectString
   End If

DeviceTypeNum = 2
SaveSetting App.Title, "Settings", "GPS_device_name", DeviceTypeNum
   
'determine the default baud rate
If InStr(GPSConnectString0, "1200") Then
   cmbBaud.ListIndex = 0
ElseIf InStr(GPSConnectString0, "2400") Then
   cmbBaud.ListIndex = 1
ElseIf InStr(GPSConnectString0, "4800") Then
   cmbBaud.ListIndex = 2
ElseIf InStr(GPSConnectString0, "9600") Then
   cmbBaud.ListIndex = 3
ElseIf InStr(GPSConnectString0, "14400") Then
   cmbBaud.ListIndex = 4
ElseIf InStr(GPSConnectString0, "19200") Then
   cmbBaud.ListIndex = 5
ElseIf InStr(GPSConnectString0, "38400") Then
   cmbBaud.ListIndex = 6
ElseIf InStr(GPSConnectString0, "56000") Then
   cmbBaud.ListIndex = 7
ElseIf InStr(GPSConnectString0, "128000") Then
   cmbBaud.ListIndex = 8
ElseIf InStr(GPSConnectString0, "256000") Then
   cmbBaud.ListIndex = 9
Else 'default
   cmbBaud.ListIndex = 6 '38400 is default
   End If
   
GPSConnectString = cmbBaud.List(cmbBaud.ListIndex) & ",N,8,1" 'scan for GPS COM port using the chosen Baud rate
'GPSsetup.lblExplain.ForeColor = &HFF&
'GPSsetup.lblExplain = "Establishing GPS communication..."

ProlificGPS = True

'determine COM port connected to the Prolific serial device

  temp = Lister.GetGUIDByName("Ports", PortGUID)
  If temp = "OK" Then
      temp = Lister.AddToList(PortGUID(0), 1)
      If temp = "OK" Then
          temp = ""
          For i = 1 To Lister.ListCount
              If ProlificGPS Then
                 cboCom.Clear
                 ComPort% = 0
                 If InStr(1, Lister.Item(i).GetName, "Prolific") Then
                    pos1% = InStr(Lister.Item(i).GetName, "COM")
                    pos2% = InStr(pos1% + 3, Lister.Item(i).GetName, ")")
                    ComPort% = val(Mid$(Lister.Item(i).GetName, pos1% + 3, pos2% - pos1% - 3))
                    SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%
                    cboCom.AddItem "COM" & ComPort%
                    cboCom.ListIndex = 0
                    Exit For
                    End If
                 End If
            Next i
            
        Else
            MsgBox temp, vbCritical, "Error Getting Device List"
        End If
    Else
        MsgBox temp, vbCritical, "Error Finding 'Ports' GUID"
    End If

    SaveSetting App.Title, "Settings", "GPS serial-USB COM port", ComPort%
    

   If DeviceType_Init Then OKButton.value = True 'record the connect string
    
   On Error GoTo 0
   Exit Sub

optND100S_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optND100S_Click of Form GPSsetup"

End Sub

Private Sub optOther_Click()

Dim temp As String, i As Long
Dim Lister As New DevLister, PortGUID() As GUID_Storage
Dim pos1%, pos2%

Dim DeviceTypeNum As Integer

   On Error GoTo optOther_Click_Error

GDMDIform.Timer_bubble.Interval = 5
GDMDIform.Timer_bubble.Enabled = True 'destroy last bubble notification

DeviceTypeNum = val(GetSetting(App.Title, "Settings", "GPS_device_name"))

GPSConnectString0 = GetSetting(App.Title, "Settings", "GPS serial-USB connection string")

If DeviceTypeNum = 3 And GPSConnectString0 <> "" Then 'baud rate already defined

    'determine the default baud rate
    If InStr(GPSConnectString0, "1200") Then
       cmbBaud.ListIndex = 0
    ElseIf InStr(GPSConnectString0, "2400") Then
       cmbBaud.ListIndex = 1
    ElseIf InStr(GPSConnectString0, "4800") Then
       cmbBaud.ListIndex = 2
    ElseIf InStr(GPSConnectString0, "9600") Then
       cmbBaud.ListIndex = 3
    ElseIf InStr(GPSConnectString0, "14400") Then
       cmbBaud.ListIndex = 4
    ElseIf InStr(GPSConnectString0, "19200") Then
       cmbBaud.ListIndex = 5
    ElseIf InStr(GPSConnectString0, "38400") Then
       cmbBaud.ListIndex = 6
    ElseIf InStr(GPSConnectString0, "56000") Then
       cmbBaud.ListIndex = 7
    ElseIf InStr(GPSConnectString0, "128000") Then
       cmbBaud.ListIndex = 8
    ElseIf InStr(GPSConnectString0, "256000") Then
       cmbBaud.ListIndex = 9
    Else 'default
       cmbBaud.ListIndex = 6 '38400 is default
       End If
       
    If DeviceType_Init Then OKButton.value = True 'record the connect string

Else

   'set baud rate
   
    Control_Num = 8
    TT1.Style = TTBalloon
    TT1.Icon = TTIconInfo
    TT1.Title = "GPS baud rate"
    TT1.TipText = "Please choose your device's baud rate."
    TT1.PopupOnDemand = True
    TT1.VisibleTime = 6000                                 'After 6 Seconds tooltip will go away
    TT1.CreateToolTip cmbBaud.hwnd
    TT1.Show GPSsetup.cmbBaud.left / Screen.TwipsPerPixelX + 5, GPSsetup.cmbBaud.Height / Screen.TwipsPerPixelX + 5  '//In Pixel only
'    Timer_bubble.Enabled = True 'timer will kill bubble after 6 seconds
  
    DeviceTypeNum = 3
    SaveSetting App.Title, "Settings", "GPS_device_name", DeviceTypeNum

   End If
   
   ProlificGPS = False

  temp = Lister.GetGUIDByName("Ports", PortGUID)
  If temp = "OK" Then
      temp = Lister.AddToList(PortGUID(0), 1)
      If temp = "OK" Then
          temp = ""
          For i = 1 To Lister.ListCount
            'list all the COM's
            pos1% = InStr(Lister.Item(i).GetName, "COM")
            pos2% = InStr(pos1% + 3, Lister.Item(i).GetName, ")")
            cboCom.AddItem "COM" & Mid$(Lister.Item(i).GetName, pos1% + 3, pos2% - pos1% - 3)
          Next i
           
          If cboCom.ListCount > 0 Then cboCom.ListIndex = 0
            
        Else
            MsgBox temp, vbCritical, "Error Getting Device List"
        End If
    Else
        MsgBox temp, vbCritical, "Error Finding 'Ports' GUID"
    End If

   On Error GoTo 0
   Exit Sub

optOther_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optOther_Click of Form GPSsetup"

End Sub

Private Sub Timer_bubble_Timer()

   TT1.Destroy
   Timer_bubble.Enabled = False

End Sub
