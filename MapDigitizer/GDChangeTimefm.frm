VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GDChangeTimefm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "               Set the Map Timer Interval"
   ClientHeight    =   1095
   ClientLeft      =   3900
   ClientTop       =   285
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GDChangeTimefm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRestartCursor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   240
      Picture         =   "GDChangeTimefm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Click to restart the map's blinking cursort"
      Top             =   40
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   465
      Left            =   1620
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Animate Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Map Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   550
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
      _Version        =   393216
      MousePointer    =   99
      MouseIcon       =   "GDChangeTimefm.frx":05CC
      LargeChange     =   100
      SmallChange     =   10
      Max             =   1000
      TickFrequency   =   20
   End
   Begin VB.Label Label1 
      Caption         =   "msec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2940
      TabIndex        =   2
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "GDChangeTimefm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRestartCursor_Click()
   'sometimes the cursor stops blinking due to
   'Window's asynchronous messaging timing problems
   'so restart it manually (can also be done by reloading the maps)
   If (GeoMap Or TopoMap) And Not GDMDIform.CenterPointTimer Then
      GDMDIform.CenterPointTimer.Enabled = True
      cmdRestartCursor.Visible = False
      End If
End Sub

Private Sub Form_Load()
   Slider1.value = GDMDIform.Timer1.Interval
   Text1.Text = Format(GDMDIform.Timer1.Interval, "####0")
   
   'if maps are visible and the cursor needs nudging, make nudging button visible
   If (GeoMap Or TopoMap) And Not GDMDIform.CenterPointTimer.Enabled Then
      cmdRestartCursor.Visible = True
      End If
      
End Sub

Private Sub Option1_Click()
   'Slider1.Value = Maps.Timer1.Interval
   'Text1.Text = Format(Maps.Timer1.Interval, "####0")
   'Maps.Timer1.Enabled = False
   'If tblbuttons(18) = 0 And routeload = False Then Maps.Timer2.Enabled = False
End Sub

Private Sub Option2_Click()
   'Slider1.Value = Maps.Timer2.Interval
   'Text1.Text = Format(Maps.Timer2.Interval, "####0")
   'Maps.Timer1.Enabled = False
   'If tblbuttons(18) = 0 And routeload = False Then Maps.Timer2.Enabled = False
End Sub

Private Sub slider1_scroll()
   If Option1.value = False Then
      'Maps.Timer2.Interval = Slider1.Value
      'Text1.Text = Maps.Timer2.Interval
   Else
      GDMDIform.Timer1.Interval = Slider1.value
      Text1.Text = GDMDIform.Timer1.Interval
      End If
   'Maps.Timer1.Enabled = False
   'If tblbuttons(18) = 0 And routeload = False Then Maps.Timer2.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Option1.value = False Then
      'Maps.Timer2.Interval = Fix(Abs(Val(Text1.Text)))
   Else
      GDMDIform.Timer1.Interval = Fix(Abs(val(Text1.Text)))
      End If
   'Maps.Timer1.Enabled = False
   'If tblbuttons(18) = 0 And routeload = False Then Maps.Timer2.Enabled = False
   Unload Me
   Set GDChangeTimefm = Nothing
   'Maps.Toolbar1.Buttons(21).Value = tbrUnpressed
   'lResult = FindWindow(vbNullString, terranam$)
   'If lResult > 0 Then
   '   ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   '   End If
End Sub

Private Sub Text1_Change()
   Slider1.value = Fix(Abs(val(Text1.Text)))
End Sub
