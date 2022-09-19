VERSION 5.00
Begin VB.Form GPSsetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Com-Port Setup for NMEA 0183 GPS"
   ClientHeight    =   2340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4425
   Icon            =   "mapGPSsetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmExplain 
      Caption         =   "Preset Settings"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   980
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
      Top             =   0
      Width           =   4095
      Begin VB.ComboBox cmbBaud 
         Height          =   315
         Left            =   1800
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
         Left            =   600
         TabIndex        =   4
         Top             =   380
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan Com Port"
      Height          =   375
      Left            =   200
      TabIndex        =   2
      ToolTipText     =   "Pick the baud rate, and click to scan for a com port"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1720
      TabIndex        =   0
      ToolTipText     =   "Accept displayed values"
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "GPSsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

   GPSSetupVis = False
   Unload Me

End Sub

Private Sub cmdScan_Click()
   
   GPSConnectString = cmbBaud.List(cmbBaud.ListIndex) & ",N,8,1" 'scan for GPS COM port using the chosen Baud rate
   Load GPStest
'   Maps.GPS_timer.Enabled = True
   lblExplain.ForeColor = &H400000
   lblExplain.Caption = "Testing Baud Rate: " & cmbBaud.List(cmbBaud.ListIndex)

End Sub

Private Sub Form_Load()

   Dim ret As Long
   
   GPSSetupVis = True

'   ret = SetWindowPos(GPSsetup.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   BringWindowToTop (GPSsetup.hWnd)
   
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
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   GPSSetupVis = False
   Set GPSsetup = Nothing
End Sub

Private Sub OKButton_Click()

  GPSConnectString = cmbBaud.List(cmbBaud.ListIndex) & ",N,8,1" 'scan for GPS COM port using the chose Baud rate
  Load GPStest

End Sub
