VERSION 5.00
Begin VB.Form GDGeoFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geographic Coordinates"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "GDGeoFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4380
   Begin VB.TextBox txtLonSec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2820
      TabIndex        =   14
      Text            =   "00"
      Top             =   420
      Width           =   555
   End
   Begin VB.TextBox txtLonMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2340
      TabIndex        =   13
      Text            =   "00"
      Top             =   420
      Width           =   375
   End
   Begin VB.TextBox txtLonDeg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1860
      TabIndex        =   12
      Text            =   "00"
      Top             =   420
      Width           =   375
   End
   Begin VB.TextBox txtLatSec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2820
      TabIndex        =   11
      Text            =   "00"
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox txtLatMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2340
      TabIndex        =   10
      Text            =   "00"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtLatDeg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1860
      TabIndex        =   9
      Text            =   "00"
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   3480
      TabIndex        =   6
      Top             =   -60
      Width           =   855
      Begin VB.CommandButton cmdGoto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         Picture         =   "GDGeoFrm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Goto inputed coordinates"
         Top             =   420
         Width           =   495
      End
      Begin VB.CheckBox chkGoto 
         Caption         =   "Goto?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Click here to input goto coordinates"
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   60
      TabIndex        =   4
      Top             =   -60
      Width           =   975
      Begin VB.OptionButton optDMS 
         Caption         =   "DMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   15
         ToolTipText     =   "Degrees Minutes Seconds"
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optDecimal 
         Caption         =   "&Decimal degrees"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Decimal degrees"
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtLon 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   1
      Text            =   "0"
      Top             =   420
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtLat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Longitude:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Latitude:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   130
      Width           =   615
   End
End
Attribute VB_Name = "GDGeoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkGoto_Click()
   If chkGoto.value = vbChecked Then
      cmdGoto.Enabled = True
      If GeoDecDeg = True Then
        txtLatDeg.Visible = False
        txtLatMin.Visible = False
        txtLatSec.Visible = False
        txtLonDeg.Visible = False
        txtLonMin.Visible = False
        txtLonSec.Visible = False
        txtLat.Visible = True
        txtLon.Visible = True
        txtLat.Enabled = True
        txtLon.Enabled = True
      Else
        txtLatDeg.Visible = True
        txtLatMin.Visible = True
        txtLatSec.Visible = True
        txtLonDeg.Visible = True
        txtLonMin.Visible = True
        txtLonSec.Visible = True
        txtLon.Visible = False
        txtLat.Visible = False
        txtLatDeg.Enabled = True
        txtLatMin.Enabled = True
        txtLatSec.Enabled = True
        txtLonDeg.Enabled = True
        txtLonMin.Enabled = True
        txtLonSec.Enabled = True
        End If
      ShowContGeo = False
   Else
      cmdGoto.Enabled = False
      txtLat.Enabled = False
      txtLon.Enabled = False
      txtLatDeg.Enabled = False
      txtLatMin.Enabled = False
      txtLatSec.Enabled = False
      txtLonDeg.Enabled = False
      txtLonMin.Enabled = False
      txtLonSec.Enabled = False
      If GeoDecDeg = True Then
        txtLatDeg.Visible = False
        txtLatMin.Visible = False
        txtLatSec.Visible = False
        txtLonDeg.Visible = False
        txtLonMin.Visible = False
        txtLonSec.Visible = False
        txtLat.Visible = True
        txtLon.Visible = True
      Else
        txtLatDeg.Visible = True
        txtLatMin.Visible = True
        txtLatSec.Visible = True
        txtLonDeg.Visible = True
        txtLonMin.Visible = True
        txtLonSec.Visible = True
        txtLon.Visible = False
        txtLat.Visible = False
      End If
      ShowContGeo = True
      End If
End Sub

Private Sub cmdGoto_Click()
   'convert geo into ITM and display in goto boxes
    If GeoDecDeg Then
       lt = val(GDGeoFrm.txtLat)
       lg = val(GDGeoFrm.txtLon)
       If lg <= 0 Then lg = -lg
    Else
       lt = val(GDGeoFrm.txtLatDeg) + val(GDGeoFrm.txtLatMin) / 60# + val(GDGeoFrm.txtLatSec) / 3600#
       'convert to a positive longitude
       lg = Abs(val(GDGeoFrm.txtLonDeg)) + Abs(val(GDGeoFrm.txtLonMin) / 60#) + Abs(val(GDGeoFrm.txtLonSec) / 3600#)
       End If
       
    If GpsCorrection Then 'wgs84
        Dim N As Long
        Dim E As Long
        Dim lat_g As Double
        Dim lon_g As Double
        lat_g = lt
        lon_g = lg
        Call wgs842ics(lat_g, lon_g, N, E)
        kmyg = N
        kmxg = E
    Else
        Call GEOCASC(lt, lg, kmyg, kmxg)
        End If
        
    kmxc = Fix(0.5 + kmxg)
    If kmyg < 870000 Then
       kmyc = Fix(0.5 + kmyg) + 1000000
    Else
       kmyc = Fix(0.5 + kmyg)
       End If
    
    ITMx = kmxc: ITMy = kmyc
    GDMDIform.Text5.Text = Int(ITMx)
    GDMDIform.Text6.Text = Int(ITMy)
    'record these coordinates
    Call UpdatePositionFile(ITMx * DigiZoom.LastZoom, ITMy * DigiZoom.LastZoom, hgt)
    
    ShowTopoMap (0)
End Sub

Private Sub Form_Load()
      'this form displays geographic coordinates and allows for
      'for geographic coordinate goto's.
      
      GDGeoFrm.Top = 0    'starting positions
      GDGeoFrm.Left = 0
      
      chkGoto.value = vbUnchecked
      txtLat.Enabled = False
      txtLon.Enabled = False
      ShowContGeo = True
      GeoDecDeg = False
      
      If GpsCorrection Then 'display this info
         GDGeoFrm.Caption = "Geographic Coordinates" & " - WGS84 geoid"
      Else
         GDGeoFrm.Caption = "Geographic Coordinates" & " - Clark geoid"
         End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set GDGeoFrm = Nothing
   Geo = False
   ShowContGeo = False
   GDMDIform.Toolbar1.Buttons(9).value = tbrUnpressed
   buttonstate&(9) = 0
End Sub

Private Sub optDecimal_Click()
      GeoDecDeg = True
      txtLatDeg.Visible = False
      txtLatMin.Visible = False
      txtLatSec.Visible = False
      txtLonDeg.Visible = False
      txtLonMin.Visible = False
      txtLonSec.Visible = False
      txtLon.Visible = True
      txtLat.Visible = True
      If chkGoto.value = vbChecked Then
        txtLat.Enabled = True
        txtLon.Enabled = True
        End If
End Sub

Private Sub optDMS_Click()
      GeoDecDeg = False
      txtLatDeg.Visible = True
      txtLatMin.Visible = True
      txtLatSec.Visible = True
      txtLonDeg.Visible = True
      txtLonMin.Visible = True
      txtLonSec.Visible = True
      txtLon.Visible = False
      txtLat.Visible = False
      If chkGoto.value = vbChecked Then
        txtLatDeg.Enabled = True
        txtLatMin.Enabled = True
        txtLatSec.Enabled = True
        txtLonDeg.Enabled = True
        txtLonMin.Enabled = True
        txtLonSec.Enabled = True
        End If
End Sub
