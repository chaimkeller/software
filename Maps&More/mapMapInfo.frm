VERSION 5.00
Begin VB.Form mapMapInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Information"
   ClientHeight    =   6750
   ClientLeft      =   7245
   ClientTop       =   3465
   ClientWidth     =   4560
   Icon            =   "mapMapInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   4560
   Begin VB.CommandButton comHelp 
      Height          =   375
      Left            =   3960
      Picture         =   "mapMapInfo.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Help"
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame infofrm 
      Caption         =   "Map Information"
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtlatcenter 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Text            =   "txtlatcenter"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtloncenter 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Text            =   "txtloncenter"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtpixlat 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   20
         Text            =   "txtpixlat"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtpixlon 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Text            =   "txtpixlon"
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox txtpixkm 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Text            =   "txtpixkm"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtYCenter 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "txtYCenter"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtXcenter 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Text            =   "txtXcenter"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtVertical 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Text            =   "txtVertical"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtHorizontal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Text            =   "txtHorizontal"
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cmbFormat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "mapMapInfo.frx":0F44
         Left            =   2640
         List            =   "mapMapInfo.frx":0F46
         TabIndex        =   3
         Text            =   "cmbFormat"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbllatcenter 
         Caption         =   "Lat (deg):"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label latloncenter 
         Caption         =   "Lon. (deg):"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lbllat 
         Caption         =   "Pixels for one degree in latitude"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label lbllon 
         Caption         =   "Pixels for one degree in longitude"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   4680
         Width           =   2535
      End
      Begin VB.Label lbldegscale 
         Caption         =   "Map Scale in pixels/degrees"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   4320
         Width           =   3375
      End
      Begin VB.Label lblXScale 
         Caption         =   "Map Scale in pixels/km"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label lblY 
         Caption         =   "Y (pixels):"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblX 
         Caption         =   "X (pixels):"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblCenter 
         Caption         =   "Map """"Center"""""
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
         Left            =   960
         TabIndex        =   9
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblvertical 
         Caption         =   "Vertical (pixels)"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblHorizontal 
         Caption         =   "Horizontal (pixels)"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblSize 
         Caption         =   "Map Pixel Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label formatLbl 
         Caption         =   "Map Picture format"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label LblMapName 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "mapMapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InfoChanged As Boolean

Private Sub cmbFormat_Click()
  InfoChanged = True
End Sub

Private Sub cmdSave_Click()
   Dim FileSave$, filsav%
   On Error GoTo cmdSave_Click_Error

      mydir$ = Dir(ErosCitiesDir$ & "*.*")
   If mydir$ <> sEmpty Then
      mapdir$ = ErosCitiesDir$
   Else
      mapdir$ = Mid$(MainDir$, 1, 1) & ":\eroscities\"
      End If
   FileSave$ = mapdir$ & MapInfo.name & ".map"
   mydir$ = Dir(FileSave$)
   If mydir$ = sEmpty Then
   Else
      Select Case MsgBox("File already exists!  Do you want to overwrite?", vbYesNoCancel Or vbInformation Or vbDefaultButton2, "Map Information")
      
        Case vbYes
          
        Case vbNo
           Exit Sub
        Case vbCancel
           Exit Sub
      End Select
      End If
      
   Select Case cmbFormat.ListIndex
      Case 1
         MapInfo.type = 0 'bmp
      Case 2
         MapInfo.type = 1 'gif
      Case 3
         MapInfo.type = 2 'jpg
   End Select
      
   filsav% = FreeFile
   Open FileSave$ For Output As #filsav%
   Print #filsav%, "[format]"
   Select Case MapInfo.type
      Case 0
         Print #filsav%, "bmp"
      Case 1
         Print #filsav%, "gif"
      Case 2
         Print #filsav%, "jpg"
   End Select
   With mapMapInfo
    Write #filsav%, CInt(.txtHorizontal), CInt(.txtVertical)
    Print #filsav%, sEmpty
    Print #filsav%, "[capital]"
    Write #filsav%, CInt(.txtXcenter), CInt(.txtYCenter)
    Print #filsav%, sEmpty
    Print #filsav%, "[pixel/km]"
    Write #filsav%, CDbl(.txtpixkm)
    Print #filsav%, sEmpty
    Print #filsav%, "[pixels for deg. lon. and lat.]"
    Write #filsav%, CInt(.txtpixlon), CInt(.txtpixlat)
    Print #filsav%, sEmpty
    Print #filsav%, "[MapCenter, deg lon, deg lat]"
    Write #filsav%, CDbl(.txtloncenter), CDbl(.txtlatcenter)
   Close #filsav%
   End With
   InfoSaved = True
   Unload Me

   On Error GoTo 0
   Exit Sub

cmdSave_Click_Error:
    Close #filsav%
End Sub

Private Sub comHelp_Click()
Call MsgBox("The Map information file contains information that allows Maps & More to use the inputed map file." _
            & vbCrLf & vbCrLf & _
            "(Note: Imported map files usually use unknown non-Mercadian map projections. As a consequence, map positions can only be approximated.  Use the Google Map interface when accuracy is needed.)" _
            & vbCrLf & "" _
            & vbCrLf & "The following are required information.  Use a graphics program like ""Paint"" to determine them:" _
            & vbCrLf & "Map Format: Choose either: bmp,gif,jpg (other formats are not supported)" _
            & vbCrLf & "" _
            & vbCrLf & "Map Pixel size:  Enter the horizontal and vertical sizes in pixels." _
            & vbCrLf & "" _
            & vbCrLf & "Center Point Pixel Coordinates:  This can be any point in the map that you have a way to read off its geographic coordinates." _
            & vbCrLf & "" _
            & vbCrLf & "Number of pixels for each km:  (Calculate this from the legend using a graphics program like ""Paint"")." _
            & vbCrLf & "" _
            & vbCrLf & "Number of pixels corresponding to one degree of longitude and one degree of latitude." _
            , vbInformation, "Map Information File Help")

End Sub

Private Sub Form_Load()
   MapFormatVis = True
   
   With mapMapInfo
      .cmbFormat.AddItem "None"
      .cmbFormat.AddItem "bmp"
      .cmbFormat.AddItem "gif"
      .cmbFormat.AddItem "jpg"
      .cmbFormat.ListIndex = MapInfo.type + 1
      .LblMapName = MapInfo.name
      .txtHorizontal = MapInfo.xsize
      .txtVertical = MapInfo.ysize
      .txtXcenter = MapInfo.pixcx
      .txtloncenter = MapInfo.loncenter
      .txtYCenter = MapInfo.pixcy
      .txtlatcenter = MapInfo.latcenter
      .txtpixkm = MapInfo.pixkm
      .txtpixlon = MapInfo.pixlon
      .txtpixlat = MapInfo.pixlat
   End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Not InfoChanged Then
      Select Case MsgBox("You didn't save any editing..." _
                         & vbCrLf & "" _
                         & vbCrLf & "Do you really want to exit without saving?" _
                         , vbYesNoCancel Or vbQuestion Or vbDefaultButton1, "Map Information")
      
        Case vbYes
          
        Case vbNo
           InfoChanged = False
           Cancel = 1
        Case vbCancel
           InfoChanged = False
           Cancel = 1
      End Select
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

   MapFormatVis = False
   Maps.mnushowmapinfo.Checked = False

   Set mapMapInfo = Nothing

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:
   Resume Next
End Sub

Private Sub txtHorizontal_Change()
   InfoChanged = True
End Sub

Private Sub txtHorizontal_Click()
   InfoChanged = True
End Sub

Private Sub txtlatcenter_Change()
   InfoChanged = True
End Sub

Private Sub txtlatcenter_Click()
   InfoChanged = True
End Sub

Private Sub txtloncenter_Change()
   InfoChanged = True
End Sub

Private Sub txtloncenter_Click()
   InfoChanged = True
End Sub

Private Sub txtpixkm_Change()
   InfoChanged = True
End Sub

Private Sub txtpixkm_Click()
   InfoChanged = True
End Sub

Private Sub txtpixlat_Change()
   InfoChanged = True
End Sub

Private Sub txtpixlat_Click()
   InfoChanged = True
End Sub

Private Sub txtpixlon_Change()
   InfoChanged = True
End Sub

Private Sub txtpixlon_Click()
   InfoChanged = True
End Sub

Private Sub txtXcenter_Change()
   InfoChanged = True
End Sub

Private Sub txtXcenter_Click()
   InfoChanged = True
End Sub

Private Sub txtYCenter_Change()
   InfoChanged = True
End Sub

Private Sub txtYCenter_Click()
   InfoChanged = True
End Sub
