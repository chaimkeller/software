VERSION 5.00
Begin VB.Form GDLocationfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locations"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "GDLocationfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4110
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   1620
      Width           =   3855
      Begin VB.PictureBox picProgBar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   3585
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   -60
      Width           =   3855
      Begin VB.TextBox txtGLCat 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Text            =   "txtGLCat"
         ToolTipText     =   "Ground Level (meters)"
         Top             =   1140
         Width           =   795
      End
      Begin VB.TextBox txtITMyCat 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Text            =   "txtITMyCat"
         ToolTipText     =   "ITMy"
         Top             =   1140
         Width           =   1035
      End
      Begin VB.TextBox txtITMxCat 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   420
         TabIndex        =   6
         Text            =   "txtITMxCat"
         ToolTipText     =   "ITMx"
         Top             =   1140
         Width           =   1095
      End
      Begin VB.CommandButton cmdMap 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1980
         Picture         =   "GDLocationfrm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Locate on map"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdLocation 
         Height          =   375
         Left            =   1560
         Picture         =   "GDLocationfrm.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Load locations"
         Top             =   240
         Width           =   435
      End
      Begin VB.ComboBox cmbPlaceNames 
         Enabled         =   0   'False
         Height          =   315
         Left            =   420
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cmbPlaceNames"
         ToolTipText     =   "Place Name"
         Top             =   720
         Width           =   3015
      End
   End
End
Attribute VB_Name = "GDLocationfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPlaceNames_Click()
   If cmbPlaceNames.ListCount > 1 Then
      'determine the corresponding ITMx,ITMy,GL
      Call LoadCatCoord(GDLocationfrm, 1)
      End If
End Sub

Private Sub cmdLocation_Click()
   'load the catalogue
   LoadPlaceCat2
End Sub

Private Sub cmdMap_Click()
   'attempt to locate the coordinates on the maps
   
    'if map is not visible, then make geo map visible
    If Not GeoMap And Not TopoMap Then
       'display the geo map
        myfile = Dir(picnam$)
        If myfile = sEmpty Or Trim$(picnam$) = sEmpty Then
           response = MsgBox("Can't find map!" & vbLf & _
                      "Use the Files/Geologic map options menu to help find it.", _
                      vbExclamation + vbOKOnly, "GSIDB")
           'take further response
            GeoMap = False
            Exit Sub
        Else
            Screen.MousePointer = vbHourglass
            buttonstate&(3) = 0
            GDMDIform.Toolbar1.Buttons(3).value = tbrUnpressed
            For i& = 4 To 7
              GDMDIform.Toolbar1.Buttons(i&).Enabled = False
            Next i&
            GDMDIform.Toolbar1.Buttons(9).Enabled = False
            If buttonstate&(15) = 1 Then 'search still activated
               GDMDIform.Toolbar1.Buttons(15).value = tbrPressed
               End If
                  
            GDMDIform.mnuGeo.Enabled = False 'disenable menu of geo. coordinates display
            GDMDIform.Toolbar1.Buttons(2).value = tbrPressed
            If topos = True Then GDMDIform.Toolbar1.Buttons(3).Enabled = True
            buttonstate&(2) = 1
            GDMDIform.Label1 = lblX
            GDMDIform.Label5 = lblX
            GDMDIform.Label2 = LblY
            GDMDIform.Label6 = LblY
              
            'load up Geo map
            Call ShowGeoMap(0)
            
            End If
            
       End If
       
       'now attempt to place the map at the recorded coordinates
       GDMDIform.Text5 = txtITMxCat
       GDMDIform.Text6 = txtITMyCat
       If PicSum Then ret = ShowWindow(GDReportfrm.hWnd, SW_MINIMIZE)
       Call gotocoord 'move the map to the record's coordinates
       'remove the clutter of windows from the screen and bring
       'the map to the top of the Z order
       BringWindowToTop (GDLocationfrm.hWnd)
      

End Sub

Private Sub Form_Load()
   With GDLocationfrm
      .Top = 0
      .Left = GDMDIform.Width / 2 - .Width / 2
      .Height = 2080
      .cmbPlaceNames.Text = sEmpty
      .txtITMxCat = sEmpty
      .txtITMyCat = sEmpty
      .txtGLCat = sEmpty
      
      '------progress bar settings---------
      .picProgBar.AutoRedraw = True
      .picProgBar.BackColor = &H8000000B 'light grey
      .picProgBar.DrawMode = 10
    
      .picProgBar.FillStyle = 0
      .picProgBar.ForeColor = &H400000 'dark blue
      
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If GDLocationfrm.Height = 2760 Then
      'in middle of loading place names, can't quit in the middle
      Cancel = True
      Exit Sub
      End If

   Unload Me
   Set GDLocationfrm = Nothing
End Sub
