VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewer 
   Caption         =   "Map Interface"
   ClientHeight    =   10200
   ClientLeft      =   2280
   ClientTop       =   3270
   ClientWidth     =   14595
   Icon            =   "mapViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   680
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   973
   Begin VB.PictureBox picView 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      ClipControls    =   0   'False
      Height          =   10200
      Left            =   0
      ScaleHeight     =   676
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   969
      TabIndex        =   0
      Top             =   0
      Width           =   14595
      Begin VB.CheckBox chkLstCP 
         Height          =   495
         Left            =   14160
         Picture         =   "mapViewer.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Toggle visbility of list of CP points"
         Top             =   9480
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   9600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   1
               Object.Width           =   1852
               MinWidth        =   1852
               Object.ToolTipText     =   "Select edit mode, either rotate, resize or translate"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   1
               Object.Width           =   1852
               MinWidth        =   1852
               Object.ToolTipText     =   "translation mode for selection region (pixels or km)"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.ToolTipText     =   "Averaged ""elevation"" over the selection region"
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog comdlg 
         Left            =   12720
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   13440
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   54
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":064C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":075E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":0870
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":0982
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":0A94
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":0DE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":0EF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":100A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":111C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":1216
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":1328
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":143A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":154C
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":165E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":1770
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":1AC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":1E14
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":2166
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":24B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":280A
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":2B5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":2EAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":3200
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":3552
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":38A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":3BF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":3F48
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":4002
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":41D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":439E
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":44F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":4652
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":49A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":4CF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":5048
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":539A
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":56EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":5A3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":5D90
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":60E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":6434
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":6786
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":6AD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":6E2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":717C
               Key             =   ""
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":75CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":7A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":7D72
               Key             =   ""
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":80C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":8416
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":8768
               Key             =   ""
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":8ABA
               Key             =   ""
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":8E0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mapViewer.frx":981E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstCP 
         Height          =   2985
         Left            =   11280
         TabIndex        =   5
         Top             =   6960
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   24
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "openmapfilekey"
               Object.ToolTipText     =   "Open map file image"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "loadcpkey"
               Object.ToolTipText     =   "load a CP list to convert pixels to cgeo oordinates"
               ImageIndex      =   49
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editcpkey"
               Object.ToolTipText     =   "Edit CP list"
               ImageIndex      =   36
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "savecpkey"
               Object.ToolTipText     =   "Save CP points"
               ImageIndex      =   42
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "uploadkey"
               Object.ToolTipText     =   "upload backup points to recreate CP points"
               ImageIndex      =   53
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deletecpkey"
               Object.ToolTipText     =   "delete tast CP point"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "panmodekey"
               Object.ToolTipText     =   "Pan mode"
               ImageIndex      =   25
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "distancekey"
               Object.ToolTipText     =   "Distances on the map"
               ImageIndex      =   48
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "selectmodekey"
               Object.ToolTipText     =   "Select mode"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "changemodekey"
               Object.ToolTipText     =   "Change between rot_resize to translation"
               ImageIndex      =   47
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "switchPIXKMkey"
               Object.ToolTipText     =   "Switch from rotation/resize rectangle to move rectangle"
               ImageIndex      =   50
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "savecornerskey"
               Object.ToolTipText     =   "Save geo coordinates of selected region"
               ImageIndex      =   45
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   2
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "openxyzfilekey"
               Object.ToolTipText     =   "Open xyz file for analysis"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "averagekey"
               Object.ToolTipText     =   "Search for ""elevations"" within select rectangle and calculate average"
               ImageIndex      =   44
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "saveresults"
               Object.ToolTipText     =   "Save analysis results"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   2
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CPkey"
               Object.ToolTipText     =   "make list of CP points visible"
               ImageIndex      =   54
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   8000
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ListBox lstAverage 
            Appearance      =   0  'Flat
            Height          =   225
            Left            =   12800
            TabIndex        =   7
            Top             =   30
            Visible         =   0   'False
            Width           =   1500
         End
         Begin MSComctlLib.ProgressBar progsearch 
            Height          =   255
            Left            =   6720
            TabIndex        =   6
            Top             =   20
            Visible         =   0   'False
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin VB.Frame frmCoords 
         BackColor       =   &H80000016&
         Height          =   615
         Left            =   3750
         TabIndex        =   2
         Top             =   9480
         Width           =   7455
         Begin VB.Label lblCoords 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   7095
         End
      End
   End
   Begin VB.PictureBox picSource 
      AutoSize        =   -1  'True
      Height          =   2175
      Left            =   12480
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================
'   DATA STRUCTURES
'===============================
Private Type ControlPoint
    imgX As Double
    imgY As Double
    lat As Double
    lon As Double
End Type

Private CP() As ControlPoint
Private CPCount As Long
Private CPListUpGraded As Boolean

Dim R(3) As POINTF      ' Unrotated corners
Dim Rot(3) As POINTF    ' Rotated corners
Dim RecGeo(3) As GEOF        'Geo coordinates of rectangle's corners
Dim imgCP(3) As POINTF 'pixel coordinates of the rectangle's corners

' Forward transform coefficients (pixel -> geo)
Private a1 As Double, a2 As Double, a3 As Double
Private b1 As Double, b2 As Double, b3 As Double
Private c1 As Double, c2 As Double

' Inverse transform coefficients (geo -> pixel)
Private ia1 As Double, ia2 As Double, ia3 As Double
Private ib1 As Double, ib2 As Double, ib3 As Double
Private ic1 As Double, ic2 As Double

Private curDestX As Double
Private curDestY As Double
Private curDestW As Double
Private curDestH As Double

' --- Selection rectangle state ---
Dim SelActive As Boolean
Dim SelStartX As Double, SelStartY As Double
Dim SelEndX As Double, SelEndY As Double
Dim SelAngle As Double
Dim SelMapStartX As Double, SelMapStartY As Double
Dim SelMapEndX As Double, SelMapEndY As Double
Dim ShiftDown As Boolean, ConvertedCoordinates As Boolean
Dim CPLoaded As Boolean, numXYZpoints As Double
Dim XYZFileName As String

'used for centering the map, for markers, and moving the markers during drag operations
Dim MarkerScreenX As Double
Dim MarkerScreenY As Double
Dim MarkerVisible As Boolean
Dim MarkerLat As Double
Dim MarkerLon As Double
Dim DidDrag As Boolean

'toolbar helper state
Dim PressState As Boolean


'----flag to log all results-----
Const RecordLog As Boolean = True 'set to true to log results

' --- Selection mode ---
Dim Mode As Integer
Dim ResizeMode As Integer
Dim TranslateMode As Integer
Const MODE_PAN = 0
Const MODE_DIST = 2
Const MODE_SELECT = 1
Const MODE_ROT_RESIZE = 0
Const MODE_TRANSLATE = 1
Const MOVE_PIXELS = 0
Const MOVE_KM = 1

'---Step Sizes-----------
Private Const RESIZE_STEP As Double = 0.2   'pixels in image space
Const MoveStepImg As Double = 1   ' move 1 image pixels per keypress
Const MoveStepKM As Double = 0.01   ' 10 meters

Private Sub chkLstCP_Click()
    If chkLstCP.value = vbChecked Then
        lstCP.Visible = True
    ElseIf chkLstCP.value = vbUnchecked Then
        lstCP.Visible = False
        End If
End Sub

'' Viewer state
'Private ZoomFactor As Double
'Private PanX As Double
'Private PanY As Double
'Private Dragging As Boolean
'Private lastX As Long
'Private lastY As Long
'Private CurMouseX As Long
'Private CurMouseY As Long
'
'' Gesture state
'Private LastGestureZoom As Double
'Private LastGestureTime As Long


'===============================
'   FORM INITIALIZATION
'===============================
 Private Sub form_load()
    Set MainForm = Me
        
    BringWindowToTop (frmViewer.hWnd)
    
    ScaleMode = vbPixels
    picView.ScaleMode = vbPixels
    picView.AutoRedraw = True
   
    RestoreWindowState Me
    InitResizer Me

    ' Make the PictureBox fill the client area initially
    picView.Move 0, 0, ScaleWidth, ScaleHeight
    
'    PI = Atn(-1) * 4#
'    cd = PI / 180#
    
    ReDim CP(1 To 200)
    
    OldWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WndProc)
    
    ' Enable WM_POINTER messages for this window
    RegisterPointerInputTarget Me.hWnd, PT_TOUCH

    picSource.ScaleMode = vbPixels
    picView.ScaleMode = vbPixels
    picView.AutoRedraw = True
    picView.BackColor = &H404040

    If Dir(PictureFileName) <> "" Then 'always load up last picture used (stored in registry)

        If InStr(PictureFileName, ".png") Then
            Set picSource.Picture = LoadPNG(PictureFileName)
        Else
            picSource.Picture = LoadPicture(PictureFileName) ' <-- your image path
            End If
    
        ZoomFactor = 1#
        PanX = 0
        PanY = 0
    
        LastGestureZoom = 0
        LastGestureTime = 0
    
        If PictureFileName <> sEmpty Then RedrawView
        End If
        
    frmViewer.Toolbar1.Refresh
    Dim i
    For i = 1 To frmViewer.Toolbar1.Buttons.count
        frmViewer.Toolbar1.Buttons(i).Visible = True
        frmViewer.Toolbar1.Buttons(i).value = tbrUnpressed
        If i > 8 Then
            frmViewer.Toolbar1.Buttons(i).Enabled = False
            End If
    Next i
    
    frmViewer.Toolbar1.Refresh
    
    EditedCPPoints = False
    frmCoords.Visible = False
    StatusBar1.Visible = False
    
    Call cmdPanMode_Click    'starat in panning/zooming mode
    
    ViewerVis = True
    

End Sub

'===============================
'   VIEWER RENDERING
'===============================
'Private Sub RedrawView()
'    Dim viewW As Long, viewH As Long
'    Dim imgW As Long, imgH As Long
'    Dim destW As Double, destH As Double
'    Dim destX As Double, destY As Double
'    Dim i As Integer
'
'    viewW = picView.ScaleWidth
'    viewH = picView.ScaleHeight
'    imgW = picSource.ScaleWidth
'    imgH = picSource.ScaleHeight
'
'    destW = imgW * ZoomFactor
'    destH = imgH * ZoomFactor
'
'    destX = (viewW - destW) / 2 + PanX
'    destY = (viewH - destH) / 2 + PanY
'
'    picView.Cls
'    picView.PaintPicture picSource.Picture, destX, destY, destW, destH, 0, 0, imgW, imgH
'
'  If Mode = MODE_SELECT Then
'
'    If SelStartX <> SelEndX And SelStartY <> SelEndY Then
'
'        If Dragging And SelAngle = 0 And Not SelActive Then 'draw rectangle during drag
'            ' Compute unrotated corners
'            R(0).X = SelStartX: R(0).Y = SelStartY   ' TL
'            R(1).X = SelEndX:   R(1).Y = SelStartY   ' TR
'            R(2).X = SelEndX:   R(2).Y = SelEndY     ' BR
'            R(3).X = SelStartX: R(3).Y = SelEndY     ' BL
'            ' Convert map ? screen
'            R(0).X = SelMapStartX * ZoomFactor + destX 'PanX
'            R(0).Y = SelMapStartY * ZoomFactor + destY 'PanY
'
'            R(1).X = SelMapEndX * ZoomFactor + destX 'PanX
'            R(1).Y = SelMapStartY * ZoomFactor + destY 'PanY
'
'            R(2).X = SelMapEndX * ZoomFactor + destX 'PanX
'            R(2).Y = SelMapEndY * ZoomFactor + destY 'PanY
'
'            R(3).X = SelMapStartX * ZoomFactor + destX 'PanX
'            R(3).Y = SelMapEndY * ZoomFactor + destY 'PanY
'
'            ' --- Dashed rectangle while dragging ---
'            picView.DrawStyle = vbDash
'            picView.DrawWidth = 1
'            picView.ForeColor = vbRed
'
'            picView.Line (R(0).X, R(0).Y)-(R(1).X, R(1).Y)
'            picView.Line (R(1).X, R(1).Y)-(R(2).X, R(2).Y)
'            picView.Line (R(2).X, R(2).Y)-(R(3).X, R(3).Y)
'            picView.Line (R(3).X, R(3).Y)-(R(0).X, R(0).Y)
'
'            ConvertedCoordinates = False
'
'         ElseIf Dragging And SelAngle <> 0 And Not SelActive Then 'draw rotated rectangle during drag
'
'            ' --- Dashed rectangle while dragging ---
'            picView.DrawStyle = vbDash
'            picView.DrawWidth = 1
'            picView.ForeColor = vbRed
'
'            ' Draw rotated rectangle
'            DrawRotatedRectangle
'
'        ElseIf Not Dragging And SelActive Then 'draw solid rectangle after drag
'            ' --- Solid rectangle after drag completes ---
'            SelActive = True 'flag that selected region has been defined
'            picView.DrawStyle = vbSolid
'            picView.DrawWidth = 2
'            picView.ForeColor = vbBlue
'
'            If SelAngle = 0 Then
'                picView.Line (R(0).X, R(0).Y)-(R(1).X, R(1).Y)
'                picView.Line (R(1).X, R(1).Y)-(R(2).X, R(2).Y)
'                picView.Line (R(2).X, R(2).Y)-(R(3).X, R(3).Y)
'                picView.Line (R(3).X, R(3).Y)-(R(0).X, R(0).Y)
'            Else
'                DrawRotatedRectangle
'            End If
'
'        End If
'
'    End If
'
'    'convert map coordinates to geo coordinates necessary in order to zoom the map
'
'End If
'
'Toolbar1.Refresh
'
'picView.DrawStyle = vbSolid
'picView.DrawWidth = 1
'
'End Sub
Public Sub RedrawView()
    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim DistAlongLine As Double
    Dim BearingAlongLine As Double
    
    'diagnostics flags
    Dim Diagnostics1 As Boolean, Diagnostics2 As Boolean
    
    Diagnostics1 = False
    Diagnostics2 = False
    
    If picSource.Picture = 0 Then Exit Sub

    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    curDestW = imgW * ZoomFactor
    curDestH = imgH * ZoomFactor

    curDestX = (viewW - curDestW) / 2 + PanX
    curDestY = (viewH - curDestH) / 2 + PanY

    picView.Cls
    picView.PaintPicture picSource.Picture, curDestX, curDestY, curDestW, curDestH

    If Mode = MODE_SELECT Then
        'selecting a rectangular region
        ' only if there is some size
        If SelStartX <> SelEndX Or SelStartY <> SelEndY Then

            ' always rebuild UNROTATED screen rectangle from map coords
            R(0).x = SelMapStartX * ZoomFactor + curDestX
            R(0).y = SelMapStartY * ZoomFactor + curDestY

            R(1).x = SelMapEndX * ZoomFactor + curDestX
            R(1).y = SelMapStartY * ZoomFactor + curDestY

            R(2).x = SelMapEndX * ZoomFactor + curDestX
            R(2).y = SelMapEndY * ZoomFactor + curDestY

            R(3).x = SelMapStartX * ZoomFactor + curDestX
            R(3).y = SelMapEndY * ZoomFactor + curDestY

            '-----------------------------
            ' 1) Dragging (dashed red)
            '-----------------------------
            If Dragging And Not SelActive Then

                picView.DrawStyle = vbDash
                picView.DrawWidth = 1
                picView.ForeColor = vbRed

                If SelAngle = 0 Then
                    picView.Line (R(0).x, R(0).y)-(R(1).x, R(1).y)
                    picView.Line (R(1).x, R(1).y)-(R(2).x, R(2).y)
                    picView.Line (R(2).x, R(2).y)-(R(3).x, R(3).y)
                    picView.Line (R(3).x, R(3).y)-(R(0).x, R(0).y)
                Else
                    DrawRotatedRectangle
                End If

                ConvertedCoordinates = False

            '-----------------------------
            ' 2) After drag (solid blue)
            '-----------------------------
            ElseIf SelActive Then

                picView.DrawStyle = vbSolid
                picView.DrawWidth = 2
                picView.ForeColor = vbBlue

                If SelAngle = 0 Then
                    picView.Line (R(0).x, R(0).y)-(R(1).x, R(1).y)
                    picView.Line (R(1).x, R(1).y)-(R(2).x, R(2).y)
                    picView.Line (R(2).x, R(2).y)-(R(3).x, R(3).y)
                    picView.Line (R(3).x, R(3).y)-(R(0).x, R(0).y)
                Else
                    DrawRotatedRectangle
                End If

            End If

        End If
    ElseIf Mode = MODE_DIST And Dragging Then
        'drawing a line and determining the distance along that line
        picView.DrawStyle = vbSolid
        picView.DrawWidth = 3
        picView.ForeColor = vbYellow
        picView.Line (firstX, firstY)-(lastX, lastY)
        'calculate distance along line
        DistAlongLine = IntervalToKilometers(CDbl(firstX), CDbl(firstY), CDbl(lastX), CDbl(lastY), "screen")
        'calculate bearing along line
        BearingAlongLine = IntervalToBearing(CDbl(firstX), CDbl(firstY), CDbl(lastX), CDbl(lastY), "screen")
        'frmCoords.Visible = True
        'StatusBar1.Visible = False
        lblCoords.Caption = Format(Str$(DistAlongLine), "###0.0####") & " km, bearing: " & Format(Str$(BearingAlongLine), "##0.0##")
    End If
    
' --- Recompute marker screen position on every redraw ---
    If MarkerVisible Then
        Dim imgX As Double, imgY As Double
        GeoToPixel MarkerLat, MarkerLon, imgX, imgY
        ImageToScreen imgX, imgY, MarkerScreenX, MarkerScreenY
    End If

    'plot the bull's eye
    If MarkerVisible Then
        picView.DrawStyle = vbSolid
        picView.DrawWidth = 2
        picView.ForeColor = vbRed
    
        ' Outer circle
        picView.Circle (MarkerScreenX, MarkerScreenY), 10, vbRed
    
        ' Inner dot
        picView.Circle (MarkerScreenX, MarkerScreenY), 3, vbRed
    
        ' Crosshair lines
        picView.Line (MarkerScreenX - 12, MarkerScreenY)-(MarkerScreenX + 12, MarkerScreenY), vbRed
        picView.Line (MarkerScreenX, MarkerScreenY - 12)-(MarkerScreenX, MarkerScreenY + 12), vbRed
    End If
    
    If Diagnostics1 Then Call PlotControlPointsOnImage ' plots image-space CPs (green)
    If Diagnostics2 Then Call PlotGeoControlPoints     ' plots geo?image?screen CPs (yellow)

   

    Toolbar1.Refresh


End Sub




Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub

    ' Resize all normal controls
    ResizeControls Me

    ' Manually resize the map viewport
    picView.Move 0, 0, ScaleWidth, ScaleHeight

    RedrawView
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If OldWndProc <> 0 Then
        SetWindowLong Me.hWnd, GWL_WNDPROC, OldWndProc
        OldWndProc = 0
    End If
    
    SaveWindowState Me
    
    If CPListUpGraded And EditedCPPoints Then
    
        Dim reply As Long
        reply = MsgBox("Before closing, save your editing of CP points?", vbQuestion + vbYesNo, "Save CP points")
        Select Case reply
            Case vbYes
                cmdSaveCP_Click
            Case vbNo
        End Select
        
        CPListUpGraded = False
        End If
        
    ViewerVis = False
        
    If Not MapPictureVis Then
        Maps.Picture4.Visible = False 'hide Maps&More's coordinate frame
        Maps.searchfm.Enabled = False 'disenable the Search menu on Maps&More's menu bar
        Maps.Toolbar1.Buttons(10).Enabled = False 'disenable the goto button
        End If
        
    'save the center coordinates
    lon = Maps.Text5.Text
    lat = Maps.Text6.Text
    If Not noheights Then hgtworld = Maps.Text7.Text
    
    Maps.mnuPanZoom.Checked = False
    Set frmViewer = Nothing
    
End Sub


'===============================
'   ZOOM AT CURSOR
'===============================
'Public Sub ZoomAtCursor(ByVal factor As Double)
'    Dim mx As Long, my As Long, i As Integer
'   Dim imgR(3) As POINTF
'
'    On Error GoTo errhand
'
'    mx = CurMouseX
'    my = CurMouseY
'
''    Dim imgW As Long, imgH As Long
'    Dim oldZoom As Double
'    Dim imgX As Double, imgY As Double
'
''    imgW = picSource.ScaleWidth
''    imgH = picSource.ScaleHeight
'
'    oldZoom = ZoomFactor
'    If oldZoom <= 0 Then oldZoom = 0.0001
'
'    ' Convert cursor to image-space coordinates
'    imgX = (mx - PanX) / oldZoom
'    imgY = (my - PanY) / oldZoom
'
'    If ConvertedCoordinates And CPLoaded Then
'        For i = 0 To 3
'            imgR(i).X = (R(i).X - PanX) / oldZoom
'            imgR(i).Y = (R(i).Y - PanY) / oldZoom
'        Next i
'        End If
'
'    ' Apply zoom
'    ZoomFactor = ZoomFactor * factor
'    If ZoomFactor < 0.05 Then ZoomFactor = 0.05
'    If ZoomFactor > 50 Then ZoomFactor = 50
'
'    ' Recompute PanX/PanY so the cursor stays fixed
'    PanX = mx - imgX * ZoomFactor
'    PanY = my - imgY * ZoomFactor
'
'    If ConvertedCoordinates And CPLoaded Then
'         'adjust screen coordinates of select rectangle corners
'        For i = 0 To 3
'            R(i).X = imgR(i).X * ZoomFactor + PanX
'            R(i).Y = imgR(i).Y * ZoomFactor + PanY
'        Next i
'        End If
'
'    RedrawView
'    Exit Sub
'
'errhand:
'    MsgBox "Error #: " & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Error"
'End Sub
Public Sub ZoomAtCursor(ByVal factor As Double)
    Dim mx As Long, my As Long, i As Integer
    Dim imgR(3) As POINTF

    On Error GoTo errhand

    mx = CurMouseX
    my = CurMouseY

    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim destW As Double, destH As Double
    Dim destX As Double, destY As Double
    Dim imgX As Double, imgY As Double
    Dim oldZoom As Double

    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    oldZoom = ZoomFactor
    If oldZoom <= 0 Then oldZoom = 0.0001

    ' current drawn size and top-left (OLD zoom)
    destW = imgW * oldZoom
    destH = imgH * oldZoom
    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    ' cursor: screen -> image (OLD zoom)
    imgX = (mx - destX) / oldZoom
    imgY = (my - destY) / oldZoom

    ' rectangle corners: screen -> image (OLD zoom)
    If ConvertedCoordinates And CPLoaded Then
        For i = 0 To 3
            imgR(i).x = (R(i).x - destX) / oldZoom
            imgR(i).y = (R(i).y - destY) / oldZoom
        Next i
    End If

    ' apply zoom
    ZoomFactor = ZoomFactor * factor
    If ZoomFactor < 0.05 Then ZoomFactor = 0.05
    If ZoomFactor > 50 Then ZoomFactor = 50

    ' new drawn size
    destW = imgW * ZoomFactor
    destH = imgH * ZoomFactor

    ' new top-left so cursor stays fixed
    destX = mx - imgX * ZoomFactor
    destY = my - imgY * ZoomFactor

    ' convert destX/destY back into PanX/PanY
    PanX = destX - (viewW - destW) / 2
    PanY = destY - (viewH - destH) / 2

    ' recompute destX/destY with NEW PanX/PanY (for rectangle)
    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    ' rectangle corners: image -> screen (NEW zoom)
    If ConvertedCoordinates And CPLoaded Then
        For i = 0 To 3
            R(i).x = imgR(i).x * ZoomFactor + destX
            R(i).y = imgR(i).y * ZoomFactor + destY
        Next i
    End If
    
    If MarkerVisible Then
        GeoToPixel MarkerLat, MarkerLon, imgX, imgY
        ImageToScreen imgX, imgY, MarkerScreenX, MarkerScreenY
    End If


    RedrawView
    Exit Sub

errhand:
    MsgBox "Error #: " & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub






'===============================
'   CONTROL POINT MANAGEMENT
'===============================
Private Sub AddControlPoint(viewX As Double, viewY As Double)
    Dim imgX As Double, imgY As Double
    'determine the pixel value of the map image from the mouse position on the (zoomed) map
    GetImagePixelFromMouse imgX, imgY

    CPCount = CPCount + 1
    If CPCount > UBound(CP) Then
        ReDim Preserve CP(1 To CPCount + 50)
    End If

    CP(CPCount).imgX = imgX
    CP(CPCount).imgY = imgY

    AskForGeoCoordinates CPCount
    UpdateCPList
End Sub


Private Sub AskForGeoCoordinates(idx As Long)
    Dim sLat As String, sLon As String
    Dim sCoord() As String
    Dim TmpFileName As String
    Dim filnum%

    sLat = InputBox("Enter (Latitude,Longitude) or just latitude for point #" & idx & vbCrLf & "(Blank, enter to exit)", "Latitude,Longitude")
    
    If sLat = "" And sLon = "" Then
        CPCount = CPCount - 1
        Exit Sub
    End If
    
    sCoord = Split(sLat, ",")
    If UBound(sCoord) > 0 Then
       sLat = sCoord(0)
       sLon = sCoord(1)
    Else
       sLat = sCoord(0)
       sLon = InputBox("Enter longitude for point #" & idx, "Longitude")
       End If

    CP(idx).lat = CDbl(sLat)
    CP(idx).lon = CDbl(sLon)
    
    'store latet CP point in backup
    filnum% = FreeFile
    Open App.Path & "\MapCPPoints_backup.txt" For Append As #filnum%
    Write #filnum%, CP(idx).imgX, CP(idx).imgY, CP(idx).lat, CP(idx).lon
    Close #filnum%

    ComputeTransform
    
    EditedCPPoints = True
    
End Sub

Private Sub UpdateCPList()
    Dim i As Long
    lstCP.Clear
    For i = 1 To CPCount
        lstCP.AddItem i & ": (" & _
            Format(CP(i).imgX, "0.0") & "," & _
            Format(CP(i).imgY, "0.0") & ")  ?  " & _
            Format(CP(i).lat, "0.00000") & ", " & _
            Format(CP(i).lon, "0.00000")
    Next i
    CPListUpGraded = True
End Sub

Private Sub cmdEditCP_Click()
    If lstCP.ListIndex < 0 Then Exit Sub
    Dim idx As Long
    idx = lstCP.ListIndex + 1
    AskForGeoCoordinates idx
    UpdateCPList
End Sub

Private Sub cmdDeleteCP_Click()
    If lstCP.ListIndex < 0 Then Exit Sub
    
    Dim idx As Long
    idx = lstCP.ListIndex + 1

    Dim i As Long
    For i = idx To CPCount - 1
        CP(i) = CP(i + 1)
    Next i

    CPCount = CPCount - 1
    UpdateCPList
    ComputeTransform
    
    If CPCount < 4 Then CPLoaded = False
End Sub

'===============================
'   SAVE / LOAD CONTROL POINTS
'===============================
Private Sub cmdSaveCP_Click()
    Dim f As Integer
    Dim FileRoot As String
    Dim CPFile As String
    
    If CPCount = 0 Then
        MsgBox "You haven't saved any CP points yet!", vbOKOnly + vbExclamation, "Save CPs"
        Exit Sub
        End If

    f = FreeFile
    If Dir(PictureFileName) <> "" Then
        FileRoot = Mid$(PictureFileName, 1, Len(PictureFileName) - 4)
    Else
       MsgBox "First load a map image", vbInformation + vbOKOnly, "Missing map image"
       Exit Sub
       End If
       
    CPFile = FileRoot + ".map"
    If Dir(CPFile) <> "" Then
        Dim reply As Long
        reply = MsgBox(CPFile & " already exits; overwrite it?", vbYesNo + vbQuestion, "load control points file")
        Select Case reply
            Case vbYes
            Case vbNo
                Exit Sub
            Case Else
        End Select
        End If
    Open CPFile For Output As #f
    
    'Open "controlpoints.txt" For Output As #f

    Dim i As Long
    Print #f, CPCount
    For i = 1 To CPCount
        Print #f, CP(i).imgX & "," & CP(i).imgY & "," & CP(i).lat & "," & CP(i).lon
    Next i

    Close #f
    MsgBox "Saved."
End Sub
Private Sub BackUpCP_Click()
'===============================
'   SAVE / LOAD CONTROL POINTS
'===============================

    Dim f As Integer
    Dim FileRoot As String
    Dim CPFile As String
    
    If CPCount = 0 Then
        Exit Sub
        End If

    f = FreeFile
    If Dir(PictureFileName) <> "" Then
        FileRoot = Mid$(PictureFileName, 1, Len(PictureFileName) - 4)
    Else
       Exit Sub
       End If
       
    CPFile = FileRoot + ".bak"
    Open CPFile For Output As #f
    
    'Open "controlpoints.txt" For Output As #f

    Dim i As Long
    Print #f, CPCount
    For i = 1 To CPCount
        Print #f, CP(i).imgX & "," & CP(i).imgY & "," & CP(i).lat & "," & CP(i).lon
    Next i

    Close #f

End Sub
Private Sub cmdLoadCP_Click()
    Dim f As Integer
    Dim i As Long, line As String
    Dim FileRoot As String
    Dim CPFile As String
    Dim parts() As String
    
   On Error GoTo cmdLoadCP_Click_Error

    f = FreeFile

    On Error Resume Next
    
    f = FreeFile
    If Dir(PictureFileName) <> "" Then
        FileRoot = Mid$(PictureFileName, 1, Len(PictureFileName) - 4)
    Else
       MsgBox "First load a map image", vbInformation + vbOKOnly, "Missing map image"
       Exit Sub
       End If
       
    CPFile = FileRoot + ".map"
    If Dir(CPFile) = "" Then
        MsgBox "CP *.map File not found." & vbCrLf & vbCrLf & "Add geocoding control points (CP) by rightclicking on the map.", _
                    vbInformation + vbOKOnly, "Add CP points"
        Exit Sub
        End If
    Open CPFile For Input As #f
    
    'Open "controlpoints.txt" For Input As #f
    If Err.Number <> 0 Then
        
        Exit Sub
    End If
    On Error GoTo 0

    Input #f, CPCount
    ReDim Preserve CP(1 To CPCount)

    Dim minlat As Double
    Dim maxlat As Double
    Dim minlon As Double
    Dim maxlon As Double
    minlat = 9999
    maxlat = -9999
    minlon = 9999
    maxlon = -9999
    
    For i = 1 To CPCount
        Line Input #f, line
        parts = Split(line, ",")
        CP(i).imgX = CDbl(parts(0))
        CP(i).imgY = CDbl(parts(1))
        CP(i).lat = CDbl(parts(2))
        CP(i).lon = CDbl(parts(3))
        
        'determine ranges
        If CP(i).lat > maxlat Then
           maxlat = CP(i).lat
           End If
        If CP(i).lat < minlat Then
           minlat = CP(i).lat
           End If
        If CP(i).lon > maxlon Then
           maxlon = CP(i).lon
           End If
        If CP(i).lon < minlon Then
           minlon = CP(i).lat
           End If
           
    Next i

    Close #f

    UpdateCPList
    ComputeTransform
    CPLoaded = True 'flag that CP points are loaded
    
    MsgBox "Control points loaded.", vbOKOnly + vbInformation, "Control Points"

    Maps.Picture4.Visible = True
    Maps.Text1.Text = "long."
    Maps.Text2.Text = "lati."
    Maps.Label1.Caption = "long."
    If Not noheights Then Maps.Label2.Caption = "lati."
    
   
   Maps.searchfm.Enabled = True
   
    If CPLoaded Then
        
        For i = 1 To Toolbar1.Buttons.count
            Toolbar1.Buttons(i).value = tbrUnpressed
            Toolbar1.Buttons(i).Enabled = True
        Next i
        
        Toolbar1.Buttons(2).value = tbrPressed
        Toolbar1.Refresh
        
        frmCoords.Visible = False
        StatusBar1.Visible = False
        
        Maps.Toolbar1.Buttons(10).Enabled = True 'enable the goto button on Maps&More toolbar
        
        Maps.Picture4.Visible = True
        
        Maps.Label1.Caption = "Long."
        Maps.Label2.Caption = "lati."
        Maps.Label5.Caption = "long."
        Maps.Label6.Caption = "Lati."
        
        world = True
        
        'if stored center coordinates in range, use them
        If lat > minlat And lat < maxlat And lon > minlon And lon < maxlon Then
            Maps.Text5.Text = lon
            Maps.Text6.Text = lat
            Maps.Text7.Text = hgtworld
            Call goto_click
            End If
            
        Toolbar1.Buttons(2).value = tbrPressed
        
        End If
   
   On Error GoTo 0
   Exit Sub

cmdLoadCP_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdLoadCP_Click of Form frmViewer"
End Sub

'===============================
'   PIXEL / GEO CONVERSION
'===============================
Private Sub GetImagePixelFromMouse(ByRef imgX As Double, ByRef imgY As Double)
    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim destW As Double, destH As Double
    Dim destX As Double, destY As Double

    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    destW = imgW * ZoomFactor
    destH = imgH * ZoomFactor

    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    imgX = (CurMouseX - destX) / ZoomFactor
    imgY = (CurMouseY - destY) / ZoomFactor
End Sub

Private Sub GetImagePixelFromCP(x As Double, y As Double, ByRef imgX As Double, ByRef imgY As Double)
    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim destW As Double, destH As Double
    Dim destX As Double, destY As Double

    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    destW = imgW * ZoomFactor
    destH = imgH * ZoomFactor

    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    imgX = (x - destX) / ZoomFactor
    imgY = (y - destY) / ZoomFactor
End Sub



Public Sub ImageToScreen(imgX As Double, imgY As Double, ByRef scrX As Double, ByRef scrY As Double)
    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim destW As Double, destH As Double
    Dim destX As Double, destY As Double

    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    destW = imgW * ZoomFactor
    destH = imgH * ZoomFactor

    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    scrX = imgX * ZoomFactor + destX
    scrY = imgY * ZoomFactor + destY
End Sub


Private Sub ShowGeoCoordinates()
    Dim hgt
    
    If CPCount < 4 Then
        lblCoords.Caption = "Not georeferenced"
        Exit Sub
    End If

    Dim imgX As Double, imgY As Double
    Dim lat As Double, lon As Double

    ' get IMAGE coordinates under the mouse
    GetImagePixelFromMouse imgX, imgY
    PixelToGeo imgX, imgY, lat, lon

    lblCoords.Caption = Format(lat, "##0.000000") & ", " & Format(lon, "###0.000000")
    Maps.Text1.Text = Format(lon, "##0.00000")
    Maps.Text2.Text = Format(lat, "###0.00000")
    If Not noheights Then
       Call WorldHeights(lon, lat, hgt)
       Maps.Text3.Text = hgt
       End If
    
End Sub


'===============================
'   TRANSFORM MANAGEMENT
'===============================
Private Sub ComputeTransform()
    If CPCount < 4 Then Exit Sub

    If CPCount = 4 Then
        SolveForwardHomographyExact
        SolveInverseHomographyExact
    Else
        SolveForwardHomographyLeastSquares
        SolveInverseHomographyLeastSquares
    End If

End Sub

'===============================
'   EXACT 4-POINT SOLVER
'===============================
Private Sub SolveForwardHomographyExact()
    Dim M(1 To 8, 1 To 8) As Double
    Dim V(1 To 8) As Double
    Dim x(1 To 8) As Double
    Dim i As Long, R As Long
    Dim px As Double, py As Double

    For i = 1 To 4
        R = (i - 1) * 2 + 1
        px = CP(i).imgX
        py = CP(i).imgY

        ' Latitude equation
        M(R, 1) = px: M(R, 2) = py: M(R, 3) = 1
        M(R, 4) = 0: M(R, 5) = 0: M(R, 6) = 0
        M(R, 7) = -px * CP(i).lat
        M(R, 8) = -py * CP(i).lat
        V(R) = CP(i).lat

        ' Longitude equation
        M(R + 1, 1) = 0: M(R + 1, 2) = 0: M(R + 1, 3) = 0
        M(R + 1, 4) = px: M(R + 1, 5) = py: M(R + 1, 6) = 1
        M(R + 1, 7) = -px * CP(i).lon
        M(R + 1, 8) = -py * CP(i).lon
        V(R + 1) = CP(i).lon
    Next i

    SolveLinearSystem M, V, x

    a1 = x(1): a2 = x(2): a3 = x(3)
    b1 = x(4): b2 = x(5): b3 = x(6)
    c1 = x(7): c2 = x(8)
End Sub

'convert from lat, lon to screen coordinates
Private Sub SolveInverseHomographyExact()
    Dim M(1 To 8, 1 To 8) As Double
    Dim V(1 To 8) As Double
    Dim x(1 To 8) As Double
    Dim i As Long, R As Long

    For i = 1 To 4
        R = (i - 1) * 2 + 1
        Dim L As Double, G As Double
        L = CP(i).lat
        G = CP(i).lon

        M(R, 1) = L: M(R, 2) = G: M(R, 3) = 1
        M(R, 4) = 0: M(R, 5) = 0: M(R, 6) = 0
        M(R, 7) = -L * CP(i).imgX
        M(R, 8) = -G * CP(i).imgX
        V(R) = CP(i).imgX

        M(R + 1, 1) = 0: M(R + 1, 2) = 0: M(R + 1, 3) = 0
        M(R + 1, 4) = L: M(R + 1, 5) = G: M(R + 1, 6) = 1
        M(R + 1, 7) = -L * CP(i).imgY
        M(R + 1, 8) = -G * CP(i).imgY
        V(R + 1) = CP(i).imgY
    Next i

    SolveLinearSystem M, V, x

    ia1 = x(1): ia2 = x(2): ia3 = x(3)
    ib1 = x(4): ib2 = x(5): ib3 = x(6)
    ic1 = x(7): ic2 = x(8)
End Sub

'===============================
'   LEAST-SQUARES SOLVER
'===============================
Private Sub SolveForwardHomographyLeastSquares()
    Dim rows As Long
    rows = CPCount * 2

    Dim M() As Double, V() As Double, x(1 To 8) As Double
    ReDim M(1 To rows, 1 To 8)
    ReDim V(1 To rows)

    Dim i As Long, R As Long
    For i = 1 To CPCount
        R = (i - 1) * 2 + 1
        Dim x1 As Double, y1 As Double
        x1 = CP(i).imgX
        y1 = CP(i).imgY

        M(R, 1) = x1: M(R, 2) = y1: M(R, 3) = 1
        M(R, 4) = 0: M(R, 5) = 0: M(R, 6) = 0
        M(R, 7) = -x1 * CP(i).lat
        M(R, 8) = -y1 * CP(i).lat
        V(R) = CP(i).lat

        M(R + 1, 1) = 0: M(R + 1, 2) = 0: M(R + 1, 3) = 0
        M(R + 1, 4) = x1: M(R + 1, 5) = y1: M(R + 1, 6) = 1
        M(R + 1, 7) = -x1 * CP(i).lon
        M(R + 1, 8) = -y1 * CP(i).lon
        V(R + 1) = CP(i).lon
    Next i

    SolveLeastSquares M, V, x

    a1 = x(1): a2 = x(2): a3 = x(3)
    b1 = x(4): b2 = x(5): b3 = x(6)
    c1 = x(7): c2 = x(8)
End Sub

Private Sub SolveInverseHomographyLeastSquares()
    Dim rows As Long
    rows = CPCount * 2

    Dim M() As Double, V() As Double, x(1 To 8) As Double
    ReDim M(1 To rows, 1 To 8)
    ReDim V(1 To rows)

    Dim i As Long, R As Long
    For i = 1 To CPCount
        R = (i - 1) * 2 + 1
        Dim L As Double, G As Double
        L = CP(i).lat
        G = CP(i).lon

        M(R, 1) = L: M(R, 2) = G: M(R, 3) = 1
        M(R, 4) = 0: M(R, 5) = 0: M(R, 6) = 0
        M(R, 7) = -L * CP(i).imgX
        M(R, 8) = -G * CP(i).imgX
        V(R) = CP(i).imgX

        M(R + 1, 1) = 0: M(R + 1, 2) = 0: M(R + 1, 3) = 0
        M(R + 1, 4) = L: M(R + 1, 5) = G: M(R + 1, 6) = 1
        M(R + 1, 7) = -L * CP(i).imgY
        M(R + 1, 8) = -G * CP(i).imgY
        V(R + 1) = CP(i).imgY
    Next i

    SolveLeastSquares M, V, x

    ia1 = x(1): ia2 = x(2): ia3 = x(3)
    ib1 = x(4): ib2 = x(5): ib3 = x(6)
    ic1 = x(7): ic2 = x(8)
End Sub

'===============================
'   LINEAR / LEAST-SQUARES CORE
'===============================
Private Sub SolveLinearSystem(ByRef M() As Double, ByRef V() As Double, ByRef x() As Double)
    Dim n As Integer
    n = UBound(M, 1) ' should be 8 here

    Dim i As Integer, j As Integer, k As Integer
    Dim maxRow As Integer
    Dim tmp As Double
    Dim a() As Double

    ReDim a(1 To n, 1 To n + 1)

    For i = 1 To n
        For j = 1 To n
            a(i, j) = M(i, j)
        Next j
        a(i, n + 1) = V(i)
    Next i

    For i = 1 To n
        maxRow = i
        For k = i + 1 To n
            If Abs(a(k, i)) > Abs(a(maxRow, i)) Then
                maxRow = k
            End If
        Next k

        If maxRow <> i Then
            For j = i To n + 1
                tmp = a(i, j)
                a(i, j) = a(maxRow, j)
                a(maxRow, j) = tmp
            Next j
        End If

        If Abs(a(i, i)) < 1E-20 Then Exit Sub

        For k = i + 1 To n
            tmp = a(k, i) / a(i, i)
            For j = i To n + 1
                a(k, j) = a(k, j) - tmp * a(i, j)
            Next j
        Next k
    Next i

    For i = n To 1 Step -1
        tmp = a(i, n + 1)
        For j = i + 1 To n
            tmp = tmp - a(i, j) * x(j)
        Next j
        x(i) = tmp / a(i, i)
    Next i
End Sub

Private Sub SolveLeastSquares(ByRef M() As Double, ByRef V() As Double, ByRef x() As Double)
    Dim rows As Long, cols As Long
    rows = UBound(M, 1)
    cols = UBound(M, 2)

    Dim MT() As Double
    ReDim MT(1 To cols, 1 To rows)

    Dim i As Long, j As Long, k As Long

    For i = 1 To rows
        For j = 1 To cols
            MT(j, i) = M(i, j)
        Next j
    Next i

    Dim MTM() As Double
    ReDim MTM(1 To cols, 1 To cols)

    For i = 1 To cols
        For j = 1 To cols
            Dim sum As Double
            sum = 0
            For k = 1 To rows
                sum = sum + MT(i, k) * M(k, j)
            Next k
            MTM(i, j) = sum
        Next j
    Next i

    Dim MTV() As Double
    ReDim MTV(1 To cols)

    For i = 1 To cols
        Dim sum1 As Double
        sum1 = 0
        For k = 1 To rows
            sum1 = sum1 + MT(i, k) * V(k)
        Next k
        MTV(i) = sum1
    Next i

    SolveLinearSystem MTM, MTV, x
End Sub
Private Function IsCPInitialized() As Boolean
    On Error GoTo NotInit
    Dim n As Long
    n = UBound(CP)
    IsCPInitialized = True
    Exit Function
NotInit:
    IsCPInitialized = False
End Function

Private Sub cmdSelectMode_Click()
    
    If Toolbar1.Buttons(10).value = tbrUnpressed Then
        Mode = MODE_SELECT
        SelAngle = 0
        RedrawView
        
        Dim i
        For i = 1 To Toolbar1.Buttons.count
            Toolbar1.Buttons(i).value = tbrUnpressed
        Next i
        
        If CPLoaded Then Toolbar1.Buttons(2).value = tbrPressed
        
        Toolbar1.Buttons(10).value = tbrPressed
        Toolbar1.Refresh
        
        frmCoords.Visible = False
        StatusBar1.Visible = False
        
        End If
End Sub

Private Sub cmdPanMode_Click()

    If Toolbar1.Buttons(8).value = tbrUnpressed Then
        Mode = MODE_PAN
        SelAngle = 0
        RedrawView
        Dim i
        For i = 1 To Toolbar1.Buttons.count
            Toolbar1.Buttons(i).value = tbrUnpressed
        Next i
        Toolbar1.Buttons(8).value = tbrPressed
        Toolbar1.Refresh
        
        frmCoords.Visible = False
        StatusBar1.Visible = False
        
        If CPLoaded Then Toolbar1.Buttons(2).value = tbrPressed
        
        End If
End Sub
Private Sub cmdDistMode_Click()

    If Toolbar1.Buttons(9).value = tbrUnpressed Then
        Mode = MODE_DIST
        RedrawView
        Dim i
        For i = 1 To Toolbar1.Buttons.count
            Toolbar1.Buttons(i).value = tbrUnpressed
        Next i
        Toolbar1.Buttons(9).value = tbrPressed
        Toolbar1.Refresh
        
        frmCoords.Visible = True
        StatusBar1.Visible = False
        
        If CPLoaded Then Toolbar1.Buttons(2).value = tbrPressed

        End If

End Sub

Private Sub DrawRotatedRectangle()
    Dim cx As Double, cy As Double
    Dim i As Long

    ' compute center of unrotated rectangle
    cx = (R(0).x + R(2).x) / 2
    cy = (R(0).y + R(2).y) / 2

    ' rotate into Rot(), but DO NOT overwrite R()
    For i = 0 To 3
        RotatePoint R(i).x, R(i).y, cx, cy, SelAngle, Rot(i).x, Rot(i).y
    Next i

    ' draw using Rot(), but DO NOT store Rot() back into R()
    picView.Line (Rot(0).x, Rot(0).y)-(Rot(1).x, Rot(1).y), vbBlue
    picView.Line (Rot(1).x, Rot(1).y)-(Rot(2).x, Rot(2).y), vbBlue
    picView.Line (Rot(2).x, Rot(2).y)-(Rot(3).x, Rot(3).y), vbBlue
    picView.Line (Rot(3).x, Rot(3).y)-(Rot(0).x, Rot(0).y), vbBlue
End Sub

Private Sub RotatePoint(px As Double, py As Double, cx As Double, cy As Double, angle As Double, ByRef rx As Double, ByRef ry As Double)
    Dim s As Double, c As Double
    s = Sin(angle * cd)
    c = Cos(angle * cd)

    px = px - cx
    py = py - cy

    rx = px * c - py * s + cx
    ry = px * s + py * c + cy
End Sub
Public Sub SaveSelectionToFile(fileName As String)
    Open fileName For Output As #1
    Dim i As Long
    For i = 0 To 3
        Print #1, Rot(i).x & "," & Rot(i).y
    Next i
    Close #1
End Sub

Private Sub SaveSelectCorners()
    Dim filnum%, fileName As String, i As Integer
    fileName = App.Path & "\selectcorners.txt"
    filnum% = FreeFile
    Open fileName For Output As #filnum%
    For i = 0 To 3
        Write #filnum%, RecGeo(i).lat, RecGeo(i).lon
    Next i
    Close #filnum%
End Sub
'Private Sub Form_MouseWheel(ByVal delta As Long)
'    If Mode = MODE_SELECT Then
'        If ShiftDown Then
'            SelAngle = SelAngle + (delta / 120) * 0.05   ' ~3 degrees per wheel click
'            RedrawView
'        End If
'    End If
'End Sub

Private Sub picView_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim cx As Double, cy As Double
     Dim cxS As Double, cyS As Double
     Dim halfW As Double, halfH As Double
     Dim viewW As Long, viewH As Long
     Dim imgW As Long, imgH As Long
     Dim destW As Double, destH As Double
     Dim destX As Double, destY As Double
     Dim i As Long

    If ResizeMode = MODE_ROT_RESIZE Then  'use the arrow keys for rotation and resizing
    
        If Mode = MODE_SELECT And SelActive Then
            
            ' compute center in IMAGE space
            cx = (SelMapStartX + SelMapEndX) / 2
            cy = (SelMapStartY + SelMapEndY) / 2
    
            ' current half-width and half-height
            halfW = Abs(SelMapEndX - SelMapStartX) / 2
            halfH = Abs(SelMapEndY - SelMapStartY) / 2
    
            Select Case KeyCode
    
                '===========================
                ' ROTATION
                '===========================
                Case vbKeyAdd
                    SelAngle = SelAngle + AngleStep
                    GoTo RecomputeAndRedraw
    
                Case vbKeySubtract
                    SelAngle = SelAngle - AngleStep
                    GoTo RecomputeAndRedraw
    
                '===========================
                ' WIDTH CONTROL (up arrow / down arrow)
                '===========================
                Case vbKeyUp
                    halfW = halfW + RESIZE_STEP
    
                Case vbKeyDown
                    halfW = halfW - RESIZE_STEP
                    If halfW < 1 Then halfW = 1
    
                '===========================
                ' HEIGHT CONTROL (right arrow / left arrow)
                '===========================
                Case vbKeyRight
                    halfH = halfH + RESIZE_STEP
    
                Case vbKeyLeft
                    halfH = halfH - RESIZE_STEP
                    If halfH < 1 Then halfH = 1
    
                Case Else
                    Exit Sub
    
            End Select
            
            '===========================
            ' Update rectangle in IMAGE space
            '===========================
            SelMapStartX = cx - halfW
            SelMapEndX = cx + halfW
            SelMapStartY = cy - halfH
            SelMapEndY = cy + halfH
 
        End If
        
    ElseIf ResizeMode = MODE_TRANSLATE Then 'use arrow keys for moving the select rectangle
    
        If Mode = MODE_SELECT And SelActive Then
    
            If TranslateMode = MOVE_PIXELS Then
    
                Select Case KeyCode
    
                    Case vbKeyLeft
                        MoveSelection -MoveStepImg, 0
    
                    Case vbKeyRight
                        MoveSelection MoveStepImg, 0
    
                    Case vbKeyUp
                        MoveSelection 0, -MoveStepImg
    
                    Case vbKeyDown
                        MoveSelection 0, MoveStepImg
    
                End Select
    
            ElseIf TranslateMode = MOVE_KM Then
    
                 Dim stepImg As Double
                stepImg = KilometersToImagePixels(MoveStepKM)
    
                Select Case KeyCode
    
                    Case vbKeyLeft
                        MoveSelection -stepImg, 0
    
                    Case vbKeyRight
                        MoveSelection stepImg, 0
    
                    Case vbKeyUp
                        MoveSelection 0, -stepImg
    
                    Case vbKeyDown
                        MoveSelection 0, stepImg
    
                End Select
                
            End If
                
        End If
            
    End If
    
RecomputeAndRedraw:
    
        ' Always rebuild unrotated rectangle first
        BuildUnrotatedScreenRect
    
        ' Then apply rotation
        ApplyRotationToR
    
        RedrawView

End Sub

'===============================
'   MOUSE INTERACTION
'===============================
Private Sub picView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then

        DidDrag = False   ' ? NEW: reset drag detection

        If Mode = MODE_PAN Then
            Dragging = True
            lastX = x
            lastY = y

        ElseIf Mode = MODE_DIST Then
            Dragging = True
            firstX = x
            firstY = y

        ElseIf Mode = MODE_SELECT Then
            Dragging = True
            SelActive = False
            SelStartX = x
            SelStartY = y
            SelEndX = x
            SelEndY = y

            ScreenToImage x, y, SelMapStartX, SelMapStartY
            SelMapEndX = SelMapStartX
            SelMapEndY = SelMapStartY
        End If

    End If

End Sub

Private Sub picView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    CurMouseX = x
    CurMouseY = y

    ' Detect real dragging (movement > 3 pixels)
    If Dragging Then
        If Abs(x - lastX) > 3 Or Abs(y - lastY) > 3 Then
            DidDrag = True
        End If
    End If
    
    If Mode = MODE_PAN And Button = vbLeftButton Then
        If Dragging Then
            PanX = PanX + (x - lastX)
            PanY = PanY + (y - lastY)
            lastX = x
            lastY = y
            RedrawView
        End If
        
    ElseIf Mode = MODE_DIST And Button = vbLeftButton Then
        'don't pan, rather draw line for distance
        If Dragging Then
            lastX = x
            lastY = y
            RedrawView
            End If

    ElseIf Mode = MODE_SELECT And Button = vbLeftButton Then
        If Dragging Then
            SelEndX = x
            SelEndY = y
            ScreenToImage x, y, SelMapEndX, SelMapEndY
            RedrawView
        End If
    End If
    
    If Mode <> MODE_DIST Then ShowGeoCoordinates
    
End Sub
Private Sub picView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim imgX As Double, imgY As Double
    Dim lat As Double, lon As Double
    Dim hgt As Double
    Dim i As Integer

    If Button = vbLeftButton Then

        If Mode = MODE_PAN Then

            Dragging = False

            '-----------------------------------------
            ' CASE 1: User dragged ? PAN, no marker
            '-----------------------------------------
            If DidDrag Then
                If MarkerVisible Then
                    GeoToPixel MarkerLat, MarkerLon, imgX, imgY
                    ImageToScreen imgX, imgY, MarkerScreenX, MarkerScreenY
                End If
                RedrawView
                Exit Sub
            End If

            '-----------------------------------------
            ' CASE 2: User clicked ? DROP marker
            '-----------------------------------------
            If CPLoaded Then

                ' Screen ? Image
                ScreenToImage x, y, imgX, imgY

                ' Image ? Geo
                PixelToGeo imgX, imgY, lat, lon

                ' Update coordinate display
                Maps.Text5.Text = Format$(lon, "###0.0#####")
                Maps.Text6.Text = Format$(lat, "##0.0#####")

                If Not noheights Then
                    Call WorldHeights(lon, lat, hgt)
                    Maps.Text7.Text = CStr(hgt)
                End If

                ' Store marker
                MarkerLat = lat
                MarkerLon = lon
                MarkerVisible = True

                ' Geo ? Image ? Screen
                GeoToPixel lat, lon, imgX, imgY
                ImageToScreen imgX, imgY, MarkerScreenX, MarkerScreenY

                RedrawView
            End If

        '===========================
        '   MODE: DISTANCE
        '===========================
        ElseIf Mode = MODE_DIST Then
            Dragging = False
            lastX = x
            lastY = y
            RedrawView

        '===========================
        '   MODE: SELECT
        '===========================
        ElseIf Mode = MODE_SELECT Then
            SelActive = True
            Dragging = False
            RedrawView

            ' Convert rectangle corners to geo
            If CPLoaded Then
                For i = 0 To 3
                    GetImagePixelFromCP R(i).x, R(i).y, imgCP(i).x, imgCP(i).y
                    PixelToGeo imgCP(i).x, imgCP(i).y, RecGeo(i).lat, RecGeo(i).lon
                Next i
                ConvertedCoordinates = True
                StatusBar1.Panels(1).Text = "Rot/Resize"
                StatusBar1.Panels(2).Text = "Move Pixels"
            End If
        End If

    ElseIf Button = vbRightButton Then
        ' Add CP point
        Call AddControlPoint(CDbl(x), CDbl(y))
    End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim CornerFileName As String
    Dim avg As Double, doclin$
    Dim parts() As String, reply$
    Dim i As Integer
    
    'dim XYZstring() As String, doclin$, avg As Double
    
    Dim filnum%, oldnum As Double
     Select Case Button.Key
            Case "openmapfilekey"
                mnuOpenFile 'open map image
            Case "loadcpkey"
                Call cmdLoadCP_Click 'load map's forward homography file (pixels->coordinates)
            Case "editcpkey"
                Call cmdEditCP_Click 'edit map's forward homography file
            Case "savecpkey"
                Call cmdSaveCP_Click 'save map's forward homography file
            Case "uploadkey"
                Call cmdUpLoadCP_Click 'upload temporary storage of pixel,coordiante data and convert to CP
            Case "deletecpkey"
                Call cmdDeleteCP_Click 'delete map's forward homography file
            Case "panmodekey"
                Call cmdPanMode_Click 'panning mode
'                StatusBar1.Panels(1) = ""
'                StatusBar1.Panels(2) = ""
            Case "distancekey" 'measure distances on the map
                Call cmdDistMode_Click 'distance measurement mode
            Case "selectmodekey"
                Call cmdSelectMode_Click 'mouse drag used to defined selected region (mousewheel defines rotation)
            Case "changemodekey" 'change between rotation-resize of rectangle to translation of rectangle up/down
                If SelActive And CPLoaded Then
                    If ResizeMode = MODE_ROT_RESIZE Then
                        ResizeMode = MODE_TRANSLATE
                        StatusBar1.Panels(1).Text = "Translate"
                    ElseIf ResizeMode = MODE_TRANSLATE Then
                        ResizeMode = MODE_ROT_RESIZE
                        StatusBar1.Panels(1).Text = "Rot/Resize"
                        End If
                    End If
            Case "switchPIXKMkey" 'change from translating in pixels to kilometers
                If SelActive And CPLoaded Then
                    If TranslateMode = MOVE_PIXELS Then
                        TranslateMode = MOVE_KM
                        StatusBar1.Panels(2).Text = "Move Km"
                    ElseIf TranslateMode = MOVE_KM Then
                        TranslateMode = MOVE_PIXELS
                        StatusBar1.Panels(2).Text = "Move Pixels"
                        End If
                    End If
            Case "savecornerskey"
                SaveSelectCorners 'save geo coordinates of selection rectangle
                'Copilot's version
                CornerFileName = App.Path & "\" & Mid$(PictureFileName, Len(PictureFileName) - 4) & "-corners.txt"
                Call SaveRectangleGeoCoordinates(CornerFileName)
            Case "openxyzfilekey" 'open a xyz file with format (lat,lon,elevation)
                If Not ConvertedCoordinates Or Not CPLoaded Then
                    MsgBox "Please determine or load geosynchonization of map image", vbOKOnly + vbInformation, "Error"
                    Exit Sub
                    End If
                frmViewer.Caption = ""
                On Error GoTo errhand
                comdlg.CancelError = True
                comdlg.fileName = "d:\Curvatures\*.xyz"
                comdlg.Filter = "xyz files (.xyz)|*.xyz|all files (*.*)|*.*"
                comdlg.ShowOpen
                XYZFileName = comdlg.fileName
                frmViewer.Caption = "Opened: " & XYZFileName
                'check the filename for the correct format, and determine the number of rows.
                Toolbar1.Enabled = False 'don't allow for other operations until validation is finished
                filnum% = FreeFile
                Open XYZFileName For Input As #filnum%
                numXYZpoints = 0
                oldnum = 0
                lstAverage.Visible = True
                frmCoords.Caption = "Validating data in file....please wait"
                 Do Until EOF(filnum%)
                    Line Input #filnum%, doclin$
                    numXYZpoints = numXYZpoints + 1
                    parts = Split(doclin$, ",")
                    ' Validate: must have EXACTLY 3 parts
                    If UBound(parts) <> 2 Then
                        MsgBox "Error in XYZ file at line " & numXYZpoints & ":" & vbCrLf & _
                               doclin$ & vbCrLf & vbCrLf & _
                               "Expected format: lat,lon,elevation", vbCritical, "XYZ Format Error"
                        Close #filnum%
                        Exit Sub
                    End If

                    ' Validate numeric values
                    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
                        MsgBox "Non-numeric value in XYZ file at line " & numXYZpoints & ":" & vbCrLf & _
                               doclin$ & vbCrLf & vbCrLf & _
                               "Expected numeric: lat,lon,elevation", vbCritical, "XYZ Format Error"
                        Close #filnum%
                        Exit Sub
                    End If
                    
                     If numXYZpoints <> oldnum And numXYZpoints Mod 1000 = 0 Then
                        lblCoords.Caption = Str$(numXYZpoints) & " points"
                        oldnum = numXYZpoints
                        lstAverage.Refresh
                        DoEvents
                        End If
 
'                    XYZstring = Split(doclin$, ",")
'                    If UBound(XYZstring) <> 2 Then
'                        MsgBox "Format of xyz file is not compatible at line #: " & Str$(numXYZpoints).vbOKOnly + vbInformation, "Error in xyz file"
'                        Close #filnum%
'                        Exit Sub
'                        End If
                Loop
                Close #filnum%
                filnum% = 0
                frmCoords.Caption = ""
                 Toolbar1.Enabled = True
                               
            Case "averagekey" 'search for data points in the xyz file within the boundary of the selection rectangle and find the average elevation
                If Not SelActive Or Not ConvertedCoordinates Or Not CPLoaded Or frmViewer.Caption = "Map Interface" Then
                    MsgBox "(1) Please determine or load geosynchonization of map image" & vbCrLf & vbCrLf & "Then select search region" & vbCrLf & vbCrLf & "(3)Then choose file to analyze", vbOKOnly + vbInformation, "Error"
                    Exit Sub
                    End If
                frmViewer.Caption = "Analyzing file: " & XYZFileName
                avg = AverageElevationInRectangle(XYZFileName)

                If (avg = -1) Then  'error incalculation
                    lstAverage.AddItem "Error detected"
                    StatusBar1.Panels(3).Text = "Error detected"
                    Exit Sub
                Else
                    'record results in log file
                    If (RecordLog) Then
                        filnum% = FreeFile
                        Open App.Path & "\SearchLog.txt" For Append As #filnum%
                        Write #filnum%, XYZFileName, "average = ", avg
                        Write #filnum%, "gx , gy (0 to 3) = "
                        For i = 0 To 3
                            Write #filnum%, g_gx(i), g_gy(i)
                        Next i%
                        Close #filnum%
                    End If
                End If
                    
                 If (avg <> -9999) Then
                    lstAverage.AddItem Format(Str$(avg), "##0.0########")
                    lblCoords = Format(Str$(avg), "##0.0########")
                    StatusBar1.Panels(3).Text = Format(Str$(avg), "##0.0########")
                Else
                    lstAverage.Clear
                    lblCoords.Caption = "No results"
                    StatusBar1.Panels(3).Text = "No results"
                    End If
                    
            Case "saveresults"
                reply$ = InputBox("Enter short description", "")
                filnum% = FreeFile
                Open App.Path & "\averages-" & Mid$(XYZFileName, Len(XYZFileName) - 4) & ".txt" For Append As #filnum%
                Write #filnum%, reply$ & ", avg = " & lstAverage.List(0), g_gx(0), g_gy(0), g_gx(1), g_gy(1), g_gx(2), g_gy(2), g_gx(3), g_gy(3)
                Close #filnum%
            Case "CPkey"
                If Toolbar1.Buttons(21).value = tbrUnpressed And Not PressState Then
                   Toolbar1.Buttons(21).value = tbrPressed
                   PressState = True
                   lstCP.Visible = True
                ElseIf Toolbar1.Buttons(21).value = tbrUnpressed And PressState Then
                   Toolbar1.Buttons(21).value = tbrUnpressed
                   lstCP.Visible = False
                   PressState = False
                   End If
       
            Case Else:
     End Select
     Exit Sub
errhand:
End Sub
Private Sub mnuOpenFile()
Dim reply As Long
On Error GoTo errhand
    comdlg.CancelError = True
    comdlg.Filter = "jpg images (jpg)|(*.jpg)|bmp images (bmp)|*.bmp|gif images (gif)|*.gif|png images (png)|*.png}all files|*.*"
    If ErosCitiesDir$ <> sEmpty Then
       comdlg.fileName = ErosCitiesDir$ & "*.jpg"
    Else
        comdlg.fileName = App.Path & "\*.jpg"
        End If
    comdlg.ShowOpen
     
    If comdlg.fileName <> sEmpty Then
       PictureFileName = comdlg.fileName
    Else
       Exit Sub
       End If
    
    If InStr(PictureFileName, ".png") Then
        Set picSource.Picture = LoadPNG(PictureFileName)
    ElseIf PictureFileName <> sEmpty Then
        picSource.Picture = LoadPicture(PictureFileName) ' <-- your image path
        End If

    ZoomFactor = 1#
    PanX = 0
    PanY = 0

    LastGestureZoom = 0
    LastGestureTime = 0

    If PictureFileName <> sEmpty Then RedrawView

'     If PictureFileName = sEmpty And comdlg.fileName <> sEmpty Then
'        PictureFileName = comdlg.fileName
'        picView.Picture = LoadPicture(PictureFileName)
'        RedrawView
'     ElseIf comdlg.fileName <> sEmpty Then
'        reply = MsgBox("Change the map image?", vbQuestion + vbYesNoCancel, "Map Picture")
'        Select Case reply
'            Case vbYes
'                 picView.Cls
'                 PictureFileName = comdlg.fileName
'                picView.Picture = LoadPicture(PictureFileName)
'               RedrawView
'            Case vbNo, vbCancel
'        End Select
'     End If
     Exit Sub
errhand:
'    Unload Me
End Sub
'Private Sub ImagePixelToGeo(ByVal imgX As Double, ByVal imgY As Double, _
'                            ByRef lon As Double, ByRef lat As Double)
'
'    lon = LonMin + (imgX / picSource.ScaleWidth) * (LonMax - LonMin)
'    lat = LatMax - (imgY / picSource.ScaleHeight) * (LatMax - LatMin)
'End Sub
Private Sub ScreenToImage(ByVal sx As Double, ByVal sy As Double, _
                          ByRef imgX As Double, ByRef imgY As Double)

    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim destW As Double, destH As Double
    Dim destX As Double, destY As Double

    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    destW = imgW * ZoomFactor
    destH = imgH * ZoomFactor

    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    imgX = (sx - destX) / ZoomFactor
    imgY = (sy - destY) / ZoomFactor
End Sub
Private Sub BuildUnrotatedScreenRect()
    R(0).x = SelMapStartX * ZoomFactor + curDestX
    R(0).y = SelMapStartY * ZoomFactor + curDestY

    R(1).x = SelMapEndX * ZoomFactor + curDestX
    R(1).y = SelMapStartY * ZoomFactor + curDestY

    R(2).x = SelMapEndX * ZoomFactor + curDestX
    R(2).y = SelMapEndY * ZoomFactor + curDestY

    R(3).x = SelMapStartX * ZoomFactor + curDestX
    R(3).y = SelMapEndY * ZoomFactor + curDestY
End Sub


Private Sub ApplyRotationToR()
    If SelAngle = 0 Then Exit Sub

    Dim cxS As Double, cyS As Double
    Dim i As Long

    ' TRUE center in SCREEN space (unrotated rectangle)
    cxS = (R(0).x + R(2).x) / 2
    cyS = (R(0).y + R(2).y) / 2

    ' rotate each corner around true center
    For i = 0 To 3
        RotatePoint R(i).x, R(i).y, cxS, cyS, SelAngle, R(i).x, R(i).y
    Next i
End Sub
Public Sub SaveRectangleGeoCoordinates(ByVal fileName As String)
    Dim i As Integer
    Dim imgX As Double, imgY As Double
    Dim lat As Double, lon As Double
    Dim f As Integer

    ' Open file for output
    f = FreeFile
    Open fileName For Output As #f

    Print #f, "Rectangle Geographic Coordinates"
    Print #f, "--------------------------------"
    Print #f, ""

    ' Process each of the 4 corners
    For i = 0 To 3

        ' Convert SCREEN ? IMAGE coordinates
        ScreenToImage R(i).x, R(i).y, imgX, imgY

        ' Convert IMAGE ? GEO coordinates
        PixelToGeo imgX, imgY, lat, lon


        ' Write to file
        Print #f, "Corner " & (i + 1) & ":"
        Print #f, "    Latitude:  " & Format(lat, "0.000000")
        Print #f, "    Longitude: " & Format(lon, "0.000000")
        Print #f, ""
    Next i

    Close #f
End Sub
'///////////////////////////////////////////////////////////////xyz analysis routines///////////////////////////////////////////
Private Function PointInPolygon(ByVal x As Double, ByVal y As Double, _
                                ByRef polyX() As Double, ByRef polyY() As Double) As Boolean
    Dim i As Long, j As Long
    Dim inside As Boolean

    j = UBound(polyX)
    inside = False

    For i = 0 To UBound(polyX)
        If ((polyY(i) > y) <> (polyY(j) > y)) Then
            If (x < (polyX(j) - polyX(i)) * (y - polyY(i)) / (polyY(j) - polyY(i)) + polyX(i)) Then
                inside = Not inside
            End If
        End If
        j = i
    Next i

    PointInPolygon = inside
End Function

Private Sub ScreenToGeo(ByVal sx As Double, ByVal sy As Double, _
                        ByRef lat As Double, ByRef lon As Double)

    Dim imgX As Double, imgY As Double

    ' screen ? image
    ScreenToImage sx, sy, imgX, imgY

    ' image ? geographic
    PixelToGeo imgX, imgY, lat, lon
End Sub
'Private Sub LoadXYZFile(ByVal fileName As String, _
'                        ByRef latArr() As Double, _
'                        ByRef lonArr() As Double, _
'                        ByRef elevArr() As Double, _
'                        ByRef count As Long)
'
'    Dim f As Integer
'    Dim line As String
'    Dim parts() As String
'
'    f = FreeFile
'    Open fileName For Input As #f
'
'    count = 0
'    Do While Not EOF(f)
'        Line Input #f, line
'        If Trim(line) <> "" Then
'            parts = Split(line, ",")
'            If UBound(parts) >= 2 Then
'                count = count + 1
'                ReDim Preserve latArr(1 To count)
'                ReDim Preserve lonArr(1 To count)
'                ReDim Preserve elevArr(1 To count)
'
'                latArr(count) = CDbl(parts(0))
'                lonArr(count) = CDbl(parts(1))
'                elevArr(count) = CDbl(parts(2))
'            End If
'        End If
'    Loop
'
'    Close #f
'End Sub
Private Sub GetRectangleGeoCorners(ByRef gx() As Double, ByRef gy() As Double)
    Dim i As Long
    Dim lat As Double, lon As Double

    ReDim gx(0 To 3)
    ReDim gy(0 To 3)

    For i = 0 To 3
        Dim imgX As Double, imgY As Double

        ' screen ? image
        ScreenToImage R(i).x, R(i).y, imgX, imgY

        ' image ? geo
        PixelToGeo imgX, imgY, lat, lon

        gx(i) = lon   ' polygon X = longitude
        gy(i) = lat   ' polygon Y = latitude
        g_gx(i) = gx(i) 'store in global array
        g_gy(i) = gy(i)
    Next i
End Sub

'Public Function AverageElevationInRectangle(ByVal xyzFile As String) As Double
'    Dim latArr() As Double, lonArr() As Double, elevArr() As Double
'    Dim count As Long
'    Dim gx() As Double, gy() As Double
'    Dim i As Long
'    Dim sumElev As Double, n As Long
'
'    ' Load XYZ file
'    LoadXYZFile xyzFile, latArr, lonArr, elevArr, count
'
'    ' Get rectangle polygon in geographic coords
'    GetRectangleGeoCorners gx, gy
'
'    sumElev = 0
'    n = 0
'
'    ' Test each XYZ point
'    For i = 1 To count
'        If PointInPolygon(lonArr(i), latArr(i), gx, gy) Then
'            sumElev = sumElev + elevArr(i)
'            n = n + 1
'        End If
'    Next i
'
'    If n > 0 Then
'        AverageElevationInRectangle = sumElev / n
'    Else
'        AverageElevationInRectangle = -9999   ' or any "no data" value
'    End If
'End Function
''Dim avg As Double
''avg = AverageElevationInRectangle("C:\data\terrain.xyz")
''MsgBox "Average elevation inside rectangle = " & avg
Public Function AverageElevationInRectangle(ByVal xyzFile As String) As Double
    Dim gx() As Double, gy() As Double
    Dim f As Integer
    Dim line As String
    Dim parts() As String
    Dim lat As Double, lon As Double, elev As Double
    Dim sumElev As Double
    Dim count As Long
    Dim lineNum As Long
    Dim OldPer As Integer
    Dim Percentage As Integer
    
    On Error GoTo errhand

    ' Get rectangle polygon in geographic coordinates
    GetRectangleGeoCorners gx, gy

    sumElev = 0
    count = 0
    lineNum = 0

    f = FreeFile
    Open xyzFile For Input As #f
    
    With progsearch
        .Visible = True
        lstAverage.Clear
        lstAverage.Visible = True
        .Max = 100
        .Min = 0
        .value = 0

    OldPer = 0
    Do While Not EOF(f)
        Line Input #f, line
        lineNum = lineNum + 1
        Percentage = Min(CLng(100 * lineNum / numXYZpoints), .Max)
        If Percentage <> OldPer Then
            .value = Percentage
'            lstAverage.Text = Str$(Percentage) & "%"
            lblCoords.Caption = Str$(Percentage) & "%"
            OldPer = Percentage
            DoEvents
            .Refresh
            End If

        line = Trim(line)
        If Len(line) = 0 Then GoTo NextLine   ' skip blank lines

        ' Split by comma
        parts = Split(line, ",")

        'validation already performed when opening file to find number of lines
        
'        ' Validate: must have EXACTLY 3 parts
'        If UBound(parts) <> 2 Then
'            MsgBox "Error in XYZ file at line " & lineNum & ":" & vbCrLf & _
'                   line & vbCrLf & vbCrLf & _
'                   "Expected format: lat,lon,elevation", vbCritical, "XYZ Format Error"
'            Close #f
'            AverageElevationInRectangle = -9999
'            Exit Function
'        End If
'
'        ' Validate numeric values
'        If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
'            MsgBox "Non-numeric value in XYZ file at line " & lineNum & ":" & vbCrLf & _
'                   line & vbCrLf & vbCrLf & _
'                   "Expected numeric: lat,lon,elevation", vbCritical, "XYZ Format Error"
'            Close #f
'            AverageElevationInRectangle = -9999
'            Exit Function
'        End If

        ' Convert values
        lat = CDbl(parts(0))
        lon = CDbl(parts(1)) 'east longitude on the map is usually defined as negative
        elev = CDbl(parts(2))

        ' Test if point is inside rotated rectangle
        If PointInPolygon(lon, lat, gx, gy) Then
            sumElev = sumElev + elev
            count = count + 1
        End If

NextLine:
    Loop

    Close #f
    f = 0
    
    .Visible = False

    If count > 0 Then
        AverageElevationInRectangle = sumElev / count
    Else
        AverageElevationInRectangle = -9999   ' no points found
    End If
    
    End With
    Exit Function
    
errhand:
    If f <> 0 Then Close #f
    MsgBox "Error #: " & Str$(Err.Number) & " detected, error message: " & Err.Description, vbOKOnly + vbInformation, "Error in routine AverageElevationRectangle"
    AverageElevationInRectangle = -1
End Function
Function Min(x As Double, y As Double) As Double
    If x <= y Then
        Min = x
    Else
        Min = y
        End If
    
End Function

Private Function GeoDistanceKM(ByVal lat1 As Double, ByVal lon1 As Double, _
                               ByVal lat2 As Double, ByVal lon2 As Double) As Double
    Const R As Double = 6371#  ' Earth radius in km
    Dim dLat As Double, dLon As Double
    Dim a As Double, c As Double

    dLat = (lat2 - lat1) * (3.14159265358979 / 180#)
    dLon = (lon2 - lon1) * (3.14159265358979 / 180#)

    lat1 = lat1 * (3.14159265358979 / 180#)
    lat2 = lat2 * (3.14159265358979 / 180#)

    a = Sin(dLat / 2) ^ 2 + Cos(lat1) * Cos(lat2) * Sin(dLon / 2) ^ 2
    c = 2 * Atn2(Sqr(a), Sqr(1 - a))

    GeoDistanceKM = R * c
End Function
'Private Sub ScreenToGeo(ByVal sx As Double, ByVal sy As Double, _
'                        ByRef lat As Double, ByRef lon As Double)
'    Dim imgX As Double, imgY As Double
'
'    ScreenToImage sx, sy, imgX, imgY
'    PixelToGeo imgX, imgY, lat, lon
'End Sub
Private Sub ImageToGeo(ByVal imgX As Double, ByVal imgY As Double, _
                       ByRef lat As Double, ByRef lon As Double)
    PixelToGeo imgX, imgY, lat, lon
End Sub
Public Function IntervalToKilometers(ByVal x1 As Double, ByVal y1 As Double, _
                                     ByVal x2 As Double, ByVal y2 As Double, _
                                     ByVal coordType As String) As Double
    Dim lat1 As Double, lon1 As Double
    Dim lat2 As Double, lon2 As Double

    coordType = LCase$(Trim$(coordType))

    Select Case coordType

        Case "screen"
            ' Convert screen -> geo
            ScreenToGeo x1, y1, lat1, lon1
            ScreenToGeo x2, y2, lat2, lon2

        Case "image"
            ' Convert image -> geo
            ImageToGeo x1, y1, lat1, lon1
            ImageToGeo x2, y2, lat2, lon2

        Case Else
            MsgBox "Invalid coordType. Use 'screen' or 'image'.", vbCritical
            IntervalToKilometers = -1
            Exit Function
    End Select

    ' Compute great-circle distance
    IntervalToKilometers = GeoDistanceKM(lat1, lon1, lat2, lon2)
End Function
''//////////////////////sample usage///////////////////////////
''didstance between two screen points
'Dim d As Double
'd = IntervalToKilometers 100, 200, 350, 600, "screen"
'MsgBox "Distance = " & d & " km"
''distance between two image pixel points
'Dim d As Double
'd = IntervalToKilometers 1200, 800, 1400, 900, "image"
'MsgBox "Distance = " & d & " km"
'///////////////////////////convert image/screen pixels to bearing
Public Function IntervalToBearing(ByVal x1 As Double, ByVal y1 As Double, _
                                     ByVal x2 As Double, ByVal y2 As Double, _
                                     ByVal coordType As String) As Double
    Dim lat1 As Double, lon1 As Double
    Dim lat2 As Double, lon2 As Double

    coordType = LCase$(Trim$(coordType))

    Select Case coordType

        Case "screen"
            ' Convert screen -> geo
            ScreenToGeo x1, y1, lat1, lon1
            ScreenToGeo x2, y2, lat2, lon2

        Case "image"
            ' Convert image -> geo
            ImageToGeo x1, y1, lat1, lon1
            ImageToGeo x2, y2, lat2, lon2

        Case Else
            MsgBox "Invalid coordType. Use 'screen' or 'image'.", vbCritical
            IntervalToBearing = -1
            Exit Function
    End Select

    ' Compute great-circle distance
    IntervalToBearing = GeoBearing(lat1, lon1, lat2, lon2)
End Function

'//////////////////////////opposite, convert any interval in kilometers to screen or image coordinates
Public Sub KilometersToIntervals(ByVal km As Double, _
                                 ByRef imgPixels As Double, _
                                 ByRef screenPixels As Double)

    Dim lat1 As Double, lon1 As Double
    Dim lat2 As Double, lon2 As Double
    Dim dkm As Double

    ' Pick a reference pixel (0,0)
    PixelToGeo 0, 0, lat1, lon1


    ' Move 1 pixel to the right
    PixelToGeo 1, 0, lat2, lon2

    ' Distance of 1 pixel in km
    dkm = GeoDistanceKM(lat1, lon1, lat2, lon2)

    If dkm <= 0 Then
        imgPixels = 0
        screenPixels = 0
        Exit Sub
    End If

    ' Convert km ? image pixels
    imgPixels = km / dkm

    ' Convert image pixels ? screen pixels
    screenPixels = imgPixels * ZoomFactor
End Sub
'///////////////////////////provided earlier//////////////////
'Private Function GeoDistanceKM(ByVal lat1 As Double, ByVal lon1 As Double, _
'                               ByVal lat2 As Double, ByVal lon2 As Double) As Double
'    Const R As Double = 6371# 'Earth radius in kms (use more accurate value)
'    Dim dLat As Double, dLon As Double
'    Dim a As Double, c As Double
'
'    dLat = (lat2 - lat1) * (3.14159265358979 / 180#)
'    dLon = (lon2 - lon1) * (3.14159265358979 / 180#)
'
'    lat1 = lat1 * (3.14159265358979 / 180#)
'    lat2 = lat2 * (3.14159265358979 / 180#)
'
'    a = Sin(dLat / 2) ^ 2 + Cos(lat1) * Cos(lat2) * Sin(dLon / 2) ^ 2
'    c = 2 * Atn2(Sqr(a), Sqr(1 - a))
'
'    GeoDistanceKM = R * c
'End Function
'////////////////////////example usage/////////////////////////////////////////////////////////////////////////////
'Dim km As Double
'Dim img As Double
'Dim scr As Double
'
'km = 5   ' 5 kilometers
'
'KilometersToIntervals km, img, scr
'
'MsgBox "5 km = " & img & " image pixels, " & scr & " screen pixels"
'////////////////////////////////
'kilometerse to geographic deltas
Public Sub KilometersToGeoDelta(ByVal km As Double, _
                                ByRef dLat As Double, _
                                ByRef dLon As Double)

    Const R As Double = 6371#   ' Earth radius in km
    Dim rad As Double

    ' Convert km to radians
    rad = km / R

    ' ?lat in degrees
    dLat = rad * (180# / 3.14159265358979)

    ' ?lon depends on latitude; assume mid-latitude of your map
    ' You can replace this with the rectangle center latitude if needed
    Dim midLat As Double
    midLat = 0   ' default; caller can override

    dLon = rad * (180# / 3.14159265358979) / Cos(midLat * (3.14159265358979 / 180#))
End Sub
'kilometers to image coordinates delta (pixels)
Public Function KilometersToImagePixels(ByVal km As Double) As Double
    Dim lat1 As Double, lon1 As Double
    Dim lat2 As Double, lon2 As Double
    Dim dkm As Double

    ' Pick reference pixel (0,0)
    PixelToGeo 0, 0, lat1, lon1


    ' Move 1 pixel to the right
    PixelToGeo 1, 0, lat2, lon2

    ' Distance of 1 image pixel in km
    dkm = GeoDistanceKM(lat1, lon1, lat2, lon2)

    If dkm <= 0 Then
        KilometersToImagePixels = 0
    Else
        KilometersToImagePixels = km / dkm
    End If
End Function
'kilometers to screen coordinates delta (pixels)
Public Function KilometersToScreenPixels(ByVal km As Double) As Double
    Dim imgPixels As Double

    imgPixels = KilometersToImagePixels(km)
    KilometersToScreenPixels = imgPixels * ZoomFactor
End Function
'///////////////////////////////////example usage////////////////////////////////////////////////////
'Dim km As Double
'Dim img As Double
'Dim scr As Double
'Dim dLat As Double, dLon As Double
'
'km = 10   ' 10 kilometers
'
'' Geographic delta
'KilometersToGeoDelta km, dLat, dLon
'Debug.Print "?Lat:", dLat, "?Lon:", dLon
'
'' Image pixels
'img = KilometersToImagePixels(km)
'Debug.Print "Image pixels:", img
'
'' Screen pixels
'scr = KilometersToScreenPixels(km)
'Debug.Print "Screen pixels:", scr
'compute area of rotated rectangle
'We already know the rectangle’s corners in screen coordinates (R(0)…R(3)).
'
'We convert each to geographic coordinates, then compute the polygon area on the Earth’s surface.
'
'For small rectangles (your case), we can safely approximate using the planar shoelace formula after converting each point to a local tangent plane in kilometers.
'
Public Function RectangleAreaKM2() As Double
    Dim gx(0 To 3) As Double, gy(0 To 3) As Double
    Dim lat As Double, lon As Double
    Dim imgX As Double, imgY As Double
    Dim x(0 To 3) As Double, y(0 To 3) As Double
    Dim lat0 As Double, lon0 As Double
    Dim dx As Double, dy As Double
    Dim sum As Double
    Dim i As Long, j As Long

    ' Convert rectangle corners to geographic coordinates
    For i = 0 To 3
        ScreenToImage R(i).x, R(i).y, imgX, imgY
        PixelToGeo imgX, imgY, lat, lon
        gx(i) = lon
        gy(i) = lat
    Next i

    ' Use corner 0 as local origin
    lat0 = gy(0)
    lon0 = gx(0)

    ' Convert each point to local XY in kilometers
    For i = 0 To 3
        dx = GeoDistanceKM(lat0, lon0, lat0, gx(i))
        If gx(i) < lon0 Then dx = -dx

        dy = GeoDistanceKM(lat0, lon0, gy(i), lon0)
        If gy(i) < lat0 Then dy = -dy

        x(i) = dx
        y(i) = dy
    Next i

    ' Shoelace formula
    sum = 0
    For i = 0 To 3
        j = (i + 1) Mod 4
        sum = sum + (x(i) * y(j) - x(j) * y(i))
    Next i

    RectangleAreaKM2 = Abs(sum) / 2
End Function

'compute side lengths and bearings
Public Sub RectangleSideLengthsAndBearings( _
        ByRef distKM() As Double, _
        ByRef bearingDeg() As Double)

    Dim i As Long, j As Long
    Dim imgX1 As Double, imgY1 As Double
    Dim imgX2 As Double, imgY2 As Double
    Dim lat1 As Double, lon1 As Double
    Dim lat2 As Double, lon2 As Double

    ReDim distKM(0 To 3)
    ReDim bearingDeg(0 To 3)

    For i = 0 To 3
        j = (i + 1) Mod 4

        ScreenToImage R(i).x, R(i).y, imgX1, imgY1
        ScreenToImage R(j).x, R(j).y, imgX2, imgY2

        PixelToGeo imgX1, imgY1, lat1, lon1
        PixelToGeo imgX2, imgY2, lat2, lon2

        distKM(i) = GeoDistanceKM(lat1, lon1, lat2, lon2)
        bearingDeg(i) = GeoBearing(lat1, lon1, lat2, lon2)
    Next i
End Sub

'bearing function
Private Function GeoBearing(ByVal lat1 As Double, ByVal lon1 As Double, _
                            ByVal lat2 As Double, ByVal lon2 As Double) As Double

    Dim rad As Double
    Dim phi1 As Double, phi2 As Double
    Dim lam1 As Double, lam2 As Double
    Dim y As Double, x As Double
    Dim brng As Double

    rad = 3.14159265358979

    phi1 = lat1 * (rad / 180#)
    phi2 = lat2 * (rad / 180#)
    lam1 = lon1 * (rad / 180#)
    lam2 = lon2 * (rad / 180#)

    y = Sin(lam2 - lam1) * Cos(phi2)
    x = Cos(phi1) * Sin(phi2) - Sin(phi1) * Cos(phi2) * Cos(lam2 - lam1)

    brng = Atn2(y, x) * (180# / rad)
    If brng < 0 Then brng = brng + 360#

    GeoBearing = brng
End Function

'draw a circle
Public Sub DrawCircleKM(ByVal centerScreenX As Double, _
                        ByVal centerScreenY As Double, _
                        ByVal radiusKM As Double)

    Dim imgX As Double, imgY As Double
    Dim lat0 As Double, lon0 As Double
    Dim lat As Double, lon As Double
    Dim sx As Double, sy As Double
    Dim angleDeg As Long

    ' Convert center to geographic coordinates
    ScreenToImage centerScreenX, centerScreenY, imgX, imgY
    PixelToGeo imgX, imgY, lat0, lon0

    picView.DrawStyle = vbSolid
    picView.DrawWidth = 1
    picView.ForeColor = vbGreen

    For angleDeg = 0 To 359
        GeoPointAtDistance lat0, lon0, radiusKM, angleDeg, lat, lon
        GeoToScreen lat, lon, sx, sy

        If angleDeg = 0 Then
            picView.CurrentX = sx
            picView.CurrentY = sy
        Else
            picView.Line -(sx, sy)
        End If
    Next angleDeg
End Sub

'helper functions
Private Sub GeoPointAtDistance(ByVal lat0 As Double, ByVal lon0 As Double, _
                               ByVal km As Double, ByVal bearingDeg As Double, _
                               ByRef lat As Double, ByRef lon As Double)

    Const R As Double = 6371#
    Dim rad As Double
    Dim phi1 As Double, lam1 As Double
    Dim phi2 As Double, lam2 As Double
    Dim delta As Double
    Dim br As Double

    rad = 3.14159265358979

    delta = km / R
    br = bearingDeg * (rad / 180#)

    phi1 = lat0 * (rad / 180#)
    lam1 = lon0 * (rad / 180#)

    phi2 = Asin(Sin(phi1) * Cos(delta) + Cos(phi1) * Sin(delta) * Cos(br))
    lam2 = lam1 + Atn2(Sin(br) * Sin(delta) * Cos(phi1), _
                       Cos(delta) - Sin(phi1) * Sin(phi2))

    lat = phi2 * (180# / rad)
    lon = lam2 * (180# / rad)
End Sub

'helper geo to screen
Private Sub GeoToScreen(ByVal lat As Double, ByVal lon As Double, _
                        ByRef sx As Double, ByRef sy As Double)

    Dim imgX As Double, imgY As Double

    GeoToPixel lat, lon, imgX, imgY
    ImageToScreen imgX, imgY, sx, sy

'    sx = imgX * ZoomFactor + curDestX
'    sy = imgY * ZoomFactor + curDestY
End Sub
'This shifts the rectangle by a given delta in image pixels, then redraws.
Public Sub MoveSelection(ByVal dxImg As Double, ByVal dyImg As Double)

    ' Move map-space coordinates
    SelMapStartX = SelMapStartX + dxImg
    SelMapEndX = SelMapEndX + dxImg

    SelMapStartY = SelMapStartY + dyImg
    SelMapEndY = SelMapEndY + dyImg

    ' Redraw everything
    RedrawView
End Sub

''key handler to move the rectangle using the arrow keys
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If Mode = MODE_SELECT And SelActive Then
'
'        If MoveMode = MOVE_PIXELS Then
'
'            Select Case KeyCode
'
'                Case vbKeyLeft
'                    MoveSelection -MoveStepImg, 0
'
'                Case vbKeyRight
'                    MoveSelection MoveStepImg, 0
'
'                Case vbKeyUp
'                    MoveSelection 0, -MoveStepImg
'
'                Case vbKeyDown
'                    MoveSelection 0, MoveStepImg
'
'            End Select
'
'        ElseIf MoveMode = MOVE_KM Then
'
'             Dim stepImg As Double
'            stepImg = KilometersToImagePixels(MoveStepKM)
'
'            Select Case KeyCode
'
'                Case vbKeyLeft
'                    MoveSelection -stepImg, 0
'
'                Case vbKeyRight
'                    MoveSelection stepImg, 0
'
'                Case vbKeyUp
'                    MoveSelection 0, -stepImg
'
'                Case vbKeyDown
'                    MoveSelection 0, stepImg
'
'            End Select
'
'        End If
'
'    End If
'
'End Sub

'Optional: Move by kilometers instead of pixels
'If you want the arrow keys to move the rectangle by a real-world distance (e.g., 0.1 km per press), use this helper:
'
'Convert km -> image pixels
'usage:

'Public Function KilometersToImagePixels(ByVal km As Double) As Double
'    Dim lat1 As Double, lon1 As Double
'    Dim lat2 As Double, lon2 As Double
'    Dim dkm As Double
'
'    ' Reference pixel
'    PixelToGeo 0, 0, lat1, lon1
'
'    ' One pixel to the right
'    PixelToGeo 1, 0, lat2, lon2
'
'    ' km per pixel
'    dkm = GeoDistanceKM(lat1, lon1, lat2, lon2)
'
'    If dkm <= 0 Then
'        KilometersToImagePixels = 0
'    Else
'        KilometersToImagePixels = km / dkm
'    End If
'End Function
' Returns angle in radians, like Atan2(y, x)
Public Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    Const PI As Double = 3.14159265358979

    If x > 0# Then
        Atn2 = Atn(y / x)
    ElseIf x < 0# And y >= 0# Then
        Atn2 = Atn(y / x) + PI
    ElseIf x < 0# And y < 0# Then
        Atn2 = Atn(y / x) - PI
    ElseIf x = 0# And y > 0# Then
        Atn2 = PI / 2#
    ElseIf x = 0# And y < 0# Then
        Atn2 = -PI / 2#
    Else
        ' x = 0 and y = 0: undefined; return 0 or handle as needed
        Atn2 = 0#
    End If
End Function
' Returns arcsine in radians
Public Function Asin(ByVal x As Double) As Double
    Const PI As Double = 3.14159265358979

    If x > 1# Or x < -1# Then
        ' Out of domain; handle however you prefer
        Asin = 0#
        Exit Function
    End If

    If x = 1# Then
        Asin = PI / 2#
    ElseIf x = -1# Then
        Asin = -PI / 2#
    Else
        Asin = Atn(x / Sqr(1# - x * x))
    End If
End Function
' arccosine
Public Function Acos(ByVal x As Double) As Double
    If x > 1# Or x < -1# Then
        Acos = 0#
        Exit Function
    End If

    Acos = PI / 2# - Asin(x)
End Function
'===============================
'  Distance Between Two Points
'===============================

' Returns distance in meters between two lat/lon points
Public Function DistanceMeters( _
    ByVal lat1 As Double, ByVal lon1 As Double, _
    ByVal lat2 As Double, ByVal lon2 As Double) As Double

    Dim dLat As Double, dLon As Double
    Dim a As Double, c As Double

    ' Convert to radians
    lat1 = lat1 * PI / 180#
    lon1 = lon1 * PI / 180#
    lat2 = lat2 * PI / 180#
    lon2 = lon2 * PI / 180#

    dLat = lat2 - lat1
    dLon = lon2 - lon1

    ' Haversine formula
    a = Sin(dLat / 2#) ^ 2# + Cos(lat1) * Cos(lat2) * Sin(dLon / 2#) ^ 2#
    c = 2# * Atn2(Sqr(a), Sqr(1# - a))

    DistanceMeters = EARTH_RADIUS * c
End Function

Public Function DistanceKm(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    DistanceKm = DistanceMeters(lat1, lon1, lat2, lon2) / 1000#
End Function

'===============================
'  Bearing Between Two Points
'===============================

' Returns initial bearing in degrees (0–360)
Public Function BearingBetweenPoints( _
    ByVal lat1 As Double, ByVal lon1 As Double, _
    ByVal lat2 As Double, ByVal lon2 As Double) As Double

    Dim dLon As Double
    Dim y As Double, x As Double
    Dim brng As Double

    lat1 = lat1 * PI / 180#
    lon1 = lon1 * PI / 180#
    lat2 = lat2 * PI / 180#
    lon2 = lon2 * PI / 180#

    dLon = lon2 - lon1

    y = Sin(dLon) * Cos(lat2)
    x = Cos(lat1) * Sin(lat2) - Sin(lat1) * Cos(lat2) * Cos(dLon)

    brng = Atn2(y, x) * 180# / PI

    If brng < 0# Then brng = brng + 360#

    BearingBetweenPoints = brng
End Function

'===============================
'  Destination Point
'===============================

' Given start point, bearing, and distance, returns new lat/lon
Public Sub DestinationPoint( _
    ByVal lat As Double, ByVal lon As Double, _
    ByVal bearing As Double, ByVal distance As Double, _
    ByRef outLat As Double, ByRef outLon As Double)

    Dim angDist As Double
    Dim br As Double
    Dim lat1 As Double, lon1 As Double
    Dim lat2 As Double, lon2 As Double

    angDist = distance / EARTH_RADIUS
    br = bearing * PI / 180#

    lat1 = lat * PI / 180#
    lon1 = lon * PI / 180#

    lat2 = Asin(Sin(lat1) * Cos(angDist) + Cos(lat1) * Sin(angDist) * Cos(br))
    lon2 = lon1 + Atn2(Sin(br) * Sin(angDist) * Cos(lat1), Cos(angDist) - Sin(lat1) * Sin(lat2))

    outLat = lat2 * 180# / PI
    outLon = lon2 * 180# / PI
End Sub

Sub cmdUpLoadCP_Click()
    'nothing yet, meant to upload the backup file for restoring the CP points
End Sub
'the goto routine for the frmViewer
'it centers the picViewer on the requested coordinates
Public Sub CenterMapOnGeo(ByVal lat As Double, ByVal lon As Double)
    Dim imgX As Double, imgY As Double
    Dim scrX As Double, scrY As Double
    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim destW As Double, destH As Double
    Dim destX As Double, destY As Double

    ' 1. GEO ? IMAGE
    GeoToPixel lat, lon, imgX, imgY

    ' 2. Viewer dimensions
    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    destW = imgW * ZoomFactor
    destH = imgH * ZoomFactor

    ' 3. Current top-left of drawn image
    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    ' 4. IMAGE ? SCREEN (current position)
    scrX = imgX * ZoomFactor + destX
    scrY = imgY * ZoomFactor + destY

    ' 5. Adjust PanX/PanY so this point becomes the center
    PanX = PanX + (viewW / 2 - scrX)
    PanY = PanY + (viewH / 2 - scrY)

    ' 6. Recompute marker position EXACTLY from geo ? image ? screen
    GeoToPixel lat, lon, imgX, imgY
    ImageToScreen imgX, imgY, MarkerScreenX, MarkerScreenY
    
    MarkerLat = lat
    MarkerLon = lon
    GeoToPixel lat, lon, imgX, imgY
    ImageToScreen imgX, imgY, MarkerScreenX, MarkerScreenY

    MarkerVisible = True

    RedrawView
    
    BringWindowToTop (frmViewer.hWnd)
    
End Sub


'diagnostics routine that plots the control points from the stored image coordinates
Public Sub PlotControlPointsOnImage()
    Dim i As Integer
    Dim scrX As Double, scrY As Double
    Dim imgX As Double, imgY As Double
    Dim viewW As Long, viewH As Long
    Dim imgW As Long, imgH As Long
    Dim destW As Double, destH As Double
    Dim destX As Double, destY As Double

    If CPCount < 1 Then Exit Sub

    viewW = picView.ScaleWidth
    viewH = picView.ScaleHeight
    imgW = picSource.ScaleWidth
    imgH = picSource.ScaleHeight

    destW = imgW * ZoomFactor
    destH = imgH * ZoomFactor

    destX = (viewW - destW) / 2 + PanX
    destY = (viewH - destH) / 2 + PanY

    ' Draw CP markers on top of the already-rendered map
    picView.DrawStyle = vbSolid
    picView.DrawWidth = 2
    picView.ForeColor = vbGreen

    For i = 1 To CPCount

        ' 1. Get image coordinates of CP
        imgX = CP(i).imgX
        imgY = CP(i).imgY

        ' 2. Convert IMAGE ? SCREEN using current zoom + pan + centering
        scrX = imgX * ZoomFactor + destX
        scrY = imgY * ZoomFactor + destY

        ' 3. Draw a bulls-eye marker
        picView.Circle (scrX, scrY), 8, vbGreen
        picView.Circle (scrX, scrY), 3, vbGreen
        picView.Line (scrX - 10, scrY)-(scrX + 10, scrY), vbGreen
        picView.Line (scrX, scrY - 10)-(scrX, scrY + 10), vbGreen

    Next i

End Sub
'further diagnostics
Public Sub PlotGeoControlPoints()
    Dim i As Long
    Dim lat As Double, lon As Double
    Dim imgX As Double, imgY As Double
    Dim scrX As Double, scrY As Double

    If CPCount < 1 Then Exit Sub

    picView.DrawStyle = vbSolid
    picView.DrawWidth = 2
    picView.ForeColor = vbYellow

    For i = 1 To CPCount
        lat = CP(i).lat
        lon = CP(i).lon

        ' Geo ? Image
        GeoToPixel lat, lon, imgX, imgY

        ' Image ? Screen
        ImageToScreen imgX, imgY, scrX, scrY

        picView.Circle (scrX, scrY), 10, vbYellow
        picView.Circle (scrX, scrY), 3, vbYellow
        picView.Line (scrX - 12, scrY)-(scrX + 12, scrY), vbYellow
        picView.Line (scrX, scrY - 12)-(scrX, scrY + 12), vbYellow
    Next i
End Sub
Private Sub PixelToGeo(x As Double, y As Double, ByRef lat As Double, ByRef lon As Double)
    If CPCount < 4 Then
        lat = 0
        lon = 0
        Exit Sub
    End If

    ' X,Y are IMAGE coordinates here
    Dim W As Double
    W = c1 * x + c2 * y + 1#
    lat = (a1 * x + a2 * y + a3) / W
    lon = (b1 * x + b2 * y + b3) / W
End Sub

Private Sub GeoToPixel(lat As Double, lon As Double, ByRef imgX As Double, ByRef imgY As Double)
    If CPCount < 4 Then
        imgX = 0
        imgY = 0
        Exit Sub
    End If

    Dim W As Double
    W = ic1 * lat + ic2 * lon + 1#
    imgX = (ia1 * lat + ia2 * lon + ia3) / W
    imgY = (ib1 * lat + ib2 * lon + ib3) / W
End Sub
'general routine to plot red filled circles on map
Public Sub PlotGeoPoints(lat As Double, lon As Double)

    Dim imgX As Double, imgY As Double
    Dim scrX As Double, scrY As Double
    Dim tmpfillstyle As Integer
    Dim CircleRadius As Integer
    
    CircleRadius = 3
    
    ' Geo ? Image
    GeoToPixel lat, lon, imgX, imgY
            
    ' Image ? Screen
    ImageToScreen imgX, imgY, scrX, scrY
            
    tmpfillstyle = picView.FillStyle
    
    picView.FillStyle = 0
    
    picView.Circle (scrX, scrY), CircleRadius, vbRed
    
    picView.FillStyle = tmpfillstyle
            
End Sub
