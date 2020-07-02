VERSION 5.00
Begin VB.Form GDRSfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grid Coordinates"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   4560
   Begin VB.Frame frmWizard 
      Caption         =   "Wizard"
      Height          =   855
      Left            =   120
      TabIndex        =   32
      Top             =   4080
      Width           =   4335
      Begin VB.CommandButton cmdNext 
         Caption         =   "Record grid intersection and go to next step"
         Height          =   375
         Index           =   1
         Left            =   500
         TabIndex        =   33
         Top             =   270
         Width           =   3375
      End
   End
   Begin VB.Frame frmType 
      Caption         =   "Screen to coordinate conversion method"
      Height          =   520
      Left            =   120
      TabIndex        =   29
      Top             =   40
      Width           =   4335
      Begin VB.CheckBox chkRS 
         Caption         =   "Rubber Sheeting"
         Height          =   255
         Left            =   2500
         TabIndex        =   31
         ToolTipText     =   "HIghly accurate screen to geo coordinate conversion"
         Top             =   220
         Width           =   1575
      End
      Begin VB.CheckBox chkSimple 
         Caption         =   "Use corner coordinates"
         Height          =   195
         Left            =   300
         TabIndex        =   30
         ToolTipText     =   "Simple conversion based on the corner coordinates"
         Top             =   220
         Width           =   2055
      End
   End
   Begin VB.Frame frmCalcType 
      Caption         =   "Calculation Method"
      Height          =   520
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   4335
      Begin VB.OptionButton optRS 
         Caption         =   "Rubber sheeting transf."
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         ToolTipText     =   "Uses Rubber Sheeting rotuines"
         Top             =   220
         Width           =   2055
      End
      Begin VB.OptionButton optDefault 
         Caption         =   "Linear extrapolation"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "Uses Rotation Matrices and linear extrapolation"
         Top             =   220
         Width           =   1695
      End
   End
   Begin VB.Frame frmUndo 
      Caption         =   "Undo"
      Height          =   735
      Left            =   3120
      TabIndex        =   20
      Top             =   5640
      Width           =   1335
      Begin VB.CommandButton cmdUndo 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   360
         Picture         =   "GDRSfrm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Undo last grid point"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame frmSteps 
      Caption         =   "Geo Coord Step Size"
      Height          =   1150
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   4335
      Begin VB.CommandButton cmdYStepConvert 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2880
         Picture         =   "GDRSfrm.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Convert coordinate display"
         Top             =   680
         Width           =   375
      End
      Begin VB.CommandButton cmdXStepConvert 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2880
         Picture         =   "GDRSfrm.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Convert coordinate display"
         Top             =   230
         Width           =   375
      End
      Begin VB.CheckBox chkStepY 
         Caption         =   "Each click steps:"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   640
         Width           =   1575
      End
      Begin VB.CheckBox chkStepX 
         Caption         =   "Each click steps:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   320
         Width           =   1575
      End
      Begin VB.TextBox txtStepY 
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
         ForeColor       =   &H00000080&
         Height          =   340
         Left            =   1680
         TabIndex        =   16
         ToolTipText     =   "Step in Geo Y (use ""-"" to separate deg-min-sec)"
         Top             =   680
         Width           =   1200
      End
      Begin VB.TextBox txtStepX 
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
         ForeColor       =   &H00800000&
         Height          =   340
         Left            =   1680
         TabIndex        =   14
         ToolTipText     =   "Step in Geo X (use ""-"" to separate degrees-minutes-sec)"
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblStepY 
         Caption         =   "Geo Y units"
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblStepX 
         Caption         =   "Geo X units"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.PictureBox picProgBar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   4305
      TabIndex        =   4
      Top             =   6480
      Width           =   4335
   End
   Begin VB.Frame frmCalclate 
      Caption         =   "Rubber Sheeting Conversion"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2895
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Activate Calculation Method"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Run affine transformation to convert grid screen coordinates to map coordinates"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frmMapCoordinates 
      Caption         =   "Map Coordinates"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1750
      Width           =   4335
      Begin VB.CommandButton cmdYConvert 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3600
         Picture         =   "GDRSfrm.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Convert coordinate display"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdXConvert 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3600
         Picture         =   "GDRSfrm.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Convert coordinate display"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtGeoY 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   340
         Left            =   1920
         TabIndex        =   8
         ToolTipText     =   "Grid's map Y coordinate (use ""-"" to separate degrees-minutes-sec)"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtGeoX 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   340
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Grid's X map coordinate (use ""-"" to separate degrees-minutes-seconds)"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblGeoY 
         Caption         =   "Grid Y Coordinate"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   640
         Width           =   1335
      End
      Begin VB.Label lblGeoX 
         Caption         =   "Grid X Coordinate"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         ToolTipText     =   "Grid X map coordinate"
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.Frame frmScreenCoordinates 
      Caption         =   "Screen Coordinates"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      Begin VB.TextBox txtY 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   340
         Left            =   2160
         TabIndex        =   6
         ToolTipText     =   "Picture's Y coordinate"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtX 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   340
         Left            =   2160
         TabIndex        =   5
         Tag             =   "Picture's X coordinate"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbltxtY 
         Caption         =   "Screen Y Coordinate"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   620
         Width           =   1575
      End
      Begin VB.Label lbltxtX 
         Caption         =   "Screen X Coordinate:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "GDRSfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkRS_Click()
   If RSMethod0 Then
      DigiRubberSheeting = False
      RSMethod0 = False
      Unload Me 'have to reload to show the entire dialog box
      End If
   GDRSfrm.chkRS.value = vbChecked
   DigiRS = True
End Sub

Private Sub chkSimple_Click()
   'calculate simple conversion
   DigiRubberSheeting = False
   
   If (ULGeoX = LRGeoX And ULGeoY = LRGeoY) Or (ULPixX = LRPixX And ULPixY = LRPixY) Then
      'corner coordinates not defined
      
      If chkSimple.value = vbChecked Then
         Call MsgBox("You need to enter the corner screen and geo coordinates" _
                  & vbCrLf & "to use this method." _
                  & vbCrLf & "" _
                  & vbCrLf & "(Hint: enter those coordinates in the Options menu/dialog)" _
                  , vbInformation, "Coordinate conversion error")
      
         Unload Me
         Exit Sub
         End If
      End If
      
   If Not RSMethod0 Then
      RSMethod1 = False
      RSMethod2 = False
      Unload Me
      End If
      
   DigiRS = False
'   GDRSfrm.chkSimple.value = vbChecked
'   GDRSfrm.Height = 1400 '7365
'   BringWindowToTop (GDRSfrm.hWnd)
   RSMethodBoth = False
   
   'determine coordinate conversion constants
   If LRPixX <> ULPixX Then PixToCoordX = (LRGeoX - ULGeoX) / (LRPixX - ULPixX)
   If ULPixY <> LRPixY Then PixToCoordY = (ULGeoY - LRGeoY) / (ULPixY - LRPixY)
   
   'set flags if rsmethod0 = false
   If Not RSMethod0 Then 'record new RSMethod0 flag value
   
        RSMethod0 = True
        RSMethod1 = False
        RSMethod2 = False
        
'        'determine coordinate conversion constants
'        If LRPixX <> ULPixX Then PixToCoordX = (LRGeoX - ULGeoX) / (LRPixX - ULPixX)
'        If ULPixY <> LRPixY Then PixToCoordY = (ULGeoY - LRGeoY) / (ULPixY - LRPixY)
        
        'store the method
        If infonum& > 0 Then
           Close #infonum&
           infonum& = 0
           End If

        infonum& = FreeFile
        Open direct$ + "\gdbinfo.sav" For Output As #infonum&
        Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
        Write #infonum&, dirNewDTM
        Write #infonum&, MinDigiEraserBrushSize
        Write #infonum&, NEDdir
        Write #infonum&, dtmdir
        Write #infonum&, ChainCodeMethod
        Write #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
        Write #infonum&, PointCenterClick
        Write #infonum&, picnam$
        Write #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
        Write #infonum&, ReportPaths&, DigiSearchRegion, numMaxHighlight&, Save_xyz%
        Write #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
        Write #infonum&, IgnoreAutoRedrawError%
        Write #infonum&, UseNewDTM%, nOtherCheck%
        Write #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
        Write #infonum&, NX_CALDAT, NY_CALDAT
        Write #infonum&, RSMethod0, RSMethod1, RSMethod2
        Write #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
        Write #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
        Close #infonum&
        End If
        
     'enable Hardy analysis
     GDMDIform.Toolbar1.Buttons(43).Enabled = True
     
     ANG = 0 'this coordinate transformation assumes that the geo axes are not rotated w.r.t. the pixel axes
     DigiRubberSheeting = True
    
     If ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
         (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) Then
         'enable GPS button
         GDMDIform.Toolbar1.Buttons(34).Enabled = True
         End If
        
    If (heights Or BasisDTMheights) And (RSMethod1 Or RSMethod2 Or RSMethod0) And DigiRubberSheeting Then
    
        'enable search height button
        GDMDIform.Toolbar1.Buttons(50).Enabled = True
        'enable contour generation
        GDMDIform.Toolbar1.Buttons(51).Enabled = True
        
        GDMDIform.Label1 = lblX
        GDMDIform.Label5 = lblX
        GDMDIform.Label2 = LblY
        GDMDIform.Label6 = LblY
        
        GDMDIform.Text3.Visible = True
        GDMDIform.Label3.Visible = True
        GDMDIform.Text7.Visible = True
        GDMDIform.Label7.Visible = True
          
        GDMDIform.Text4.Visible = True
        GDMDIform.Label4.Visible = True
        
        End If
        
     'press rubber sheeting button
     buttonstate&(39) = 1
     GDMDIform.Toolbar1.Buttons(39).value = tbrPressed
        
   
End Sub

Private Sub chkStepX_Click()
   If IsNumeric(val(GDRSfrm.txtStepX)) And Not RSMethod0 Then
      If chkStepX.value = vbChecked Then
         StepInX = True
         StepInY = False
          
        If StepInX Then
           GDRSfrm.txtGeoX = ULGridX
           GDRSfrm.txtGeoY = ULGridY
           DigiRSStepType = 0
           Call MsgBox("The rubber sheeting wizard will guide you the coordinates:," _
                     & vbCrLf _
                     & vbCrLf & "Start at the NW grid intersection of the top row" _
                     & vbCrLf & "Digitize grid intersections in that row, W to E," _
                     & vbCrLf & "descend one row to the South," _
                     & vbCrLf & "Digitize grid intersections in the second row, E to W," _
                     & vbCrLf & "descend one row to the South," _
                     & vbCrLf & "Digitize grid intersections in the third row, W to E" _
                     & vbCrLf & "etc." _
                     , vbInformation Or vbDefaultButton1, "Rubber Sheeting")
        ElseIf StepInY Then 'starting digitizing grid intersections, display hint
           GDRSfrm.txtGeoX = LRGridX
           GDRSfrm.txtGeoY = LRGridY
           DigiRSStepType = 1
           Call MsgBox("The rubber sheeting wizard will guide you with the coordinates:," _
                     & vbCrLf _
                     & vbCrLf & "Start at the SE grid intersection of the last column" _
                     & vbCrLf & "Digitize grid intersections in that column from S to N," _
                     & vbCrLf & "move left one column (one column to the West)," _
                     & vbCrLf & "Digitize grid intersections in that column, N to S," _
                     & vbCrLf & "move left one column (one column to the West)," _
                     & vbCrLf & "Digitize grid intersections in that column, S to N" _
                     & vbCrLf & "etc." _
                     , vbInformation Or vbDefaultButton1, "Rubber Sheeting")
           End If
          
      Else
         StepInX = False
         End If
   ElseIf Not RSMethod0 Then
     MsgBox "Step Size in X is not a number!", vbExclamation + vbOKOnly, "Step Size in X Error"
     End If
End Sub

Private Sub chkStepY_Click()
   If IsNumeric(val(GDRSfrm.txtStepY)) And Not RSMethod0 Then
      If chkStepY.value = vbChecked Then
         StepInY = True
         StepInX = False
         
        If StepInX Then
           GDRSfrm.txtGeoX = ULGridX
           GDRSfrm.txtGeoY = ULGridY
           DigiRSStepType = 0
           Call MsgBox("The rubber sheeting wizard will guide you the coordinates:," _
                     & vbCrLf _
                     & vbCrLf & "Start at the NW grid intersection of the top row" _
                     & vbCrLf & "Digitize grid intersections in that row, W to E," _
                     & vbCrLf & "descend one row to the South," _
                     & vbCrLf & "Digitize grid intersections in the second row, E to W," _
                     & vbCrLf & "descend one row to the South," _
                     & vbCrLf & "Digitize grid intersections in the third row, W to E" _
                     & vbCrLf & "etc." _
                     , vbInformation Or vbDefaultButton1, "Rubber Sheeting")
        ElseIf StepInY Then 'starting digitizing grid intersections, display hint
           GDRSfrm.txtGeoX = LRGridX
           GDRSfrm.txtGeoY = LRGridY
           DigiRSStepType = 1
           Call MsgBox("The rubber sheeting wizard will guide you with the coordinates:," _
                     & vbCrLf _
                     & vbCrLf & "Start at the SE grid intersection of the last column" _
                     & vbCrLf & "Digitize grid intersections in that column from S to N," _
                     & vbCrLf & "move left one column (one column to the West)," _
                     & vbCrLf & "Digitize grid intersections in that column, N to S," _
                     & vbCrLf & "move left one column (one column to the West)," _
                     & vbCrLf & "Digitize grid intersections in that column, S to N" _
                     & vbCrLf & "etc." _
                     , vbInformation Or vbDefaultButton1, "Rubber Sheeting")
           End If

         
      Else
         StepInY = False
         End If
   ElseIf Not RSMethod0 Then
      MsgBox "Step Size in Y is not a number!", vbExclamation + vbOKOnly, "Step Size in Y Error"
      End If
End Sub


Public Sub cmdConvert_Click()
   
   Dim ier As Integer
   
   'reload map without the x's
   ier = ReDrawMap(0)
   
   'first record RSmethod
   'set flags if rsmethod0 = false
   If Not RSMethod0 Then 'record new RSMethod0 flag value
   
        If ULGeoX = LRGeoX Or ULGeoY = LRGeoY Then
'           And ULPixX <> LRPixX And ULPixY <> LRPixY
            'define default lrgeox and ulgeoy using rubber sheeting
            
            Dim XGeo As Double, Xcoord As Double
            Dim YGeo As Double, Ycoord As Double
       
            If ULPixX <> 0 And ULPixY <> 0 Then
                Xcoord = ULPixX
                Ycoord = ULPixY
            Else
               ULPixX = 0
               ULPixY = 0
               Xcoord = 0
               Ycoord = 0
               End If
               
           If RSMethod1 Then
              ier = RS_pixel_to_coord2(Xcoord, Ycoord, XGeo, YGeo)
           ElseIf RSMethod2 Then
              ier = RS_pixel_to_coord(Xcoord, Ycoord, XGeo, YGeo)
           ElseIf RSMethod0 Then
              ier = Simple_pixel_to_coord(Xcoord, Ycoord, XGeo, YGeo)
              End If
              
           ULGeoX = XGeo
           ULGeoY = YGeo
               
               
          If LRPixX <> 0 And LRPixY <> 0 Then
             Xcoord = LRPixX
             Ycoord = LRPixY
          Else
             LRPixX = pixwi - 1
             LRPixY = pixhi - 1
             Xcoord = LRPixX
             Ycoord = LRPixY
             End If
               
           If RSMethod1 Then
              ier = RS_pixel_to_coord2(Xcoord, Ycoord, XGeo, YGeo)
           ElseIf RSMethod2 Then
              ier = RS_pixel_to_coord(Xcoord, Ycoord, XGeo, YGeo)
           ElseIf RSMethod0 Then
              ier = Simple_pixel_to_coord(Xcoord, Ycoord, XGeo, YGeo)
              End If
              
           LRGeoX = XGeo
           LRGeoY = YGeo
        
           End If
           
        If ULGeoX <> LRGeoX And ULGeoY <> LRGeoY _
           And ULPixX <> LRPixX And ULPixY <> LRPixY Then
           GDMDIform.Toolbar1.Buttons(8).Enabled = True  'allow goto actions
           End If
              
        'store the method
        If infonum& > 0 Then
           Close #infonum&
           infonum& = 0
           End If
        infonum& = FreeFile
        Open direct$ + "\gdbinfo.sav" For Output As #infonum&
        Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
        Write #infonum&, dirNewDTM
        Write #infonum&, MinDigiEraserBrushSize
        Write #infonum&, NEDdir
        Write #infonum&, dtmdir
        Write #infonum&, ChainCodeMethod
        Write #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
        Write #infonum&, PointCenterClick
        Write #infonum&, picnam$
        Write #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
        Write #infonum&, ReportPaths&, DigiSearchRegion, numMaxHighlight&, Save_xyz%
        Write #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
        Write #infonum&, IgnoreAutoRedrawError%
        Write #infonum&, UseNewDTM%, nOtherCheck%
        Write #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
        Write #infonum&, NX_CALDAT, NY_CALDAT
        Write #infonum&, RSMethod0, RSMethod1, RSMethod2
        Write #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
        Write #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
        Close #infonum&
        
        End If
   
   With GDRSfrm
       .picProgBar.Visible = True
       .picProgBar.Enabled = True
       pbScaleWidth = 100
       .picProgBar.ScaleWidth = 100
   End With
   
'    'now record the results onto the RS file
'
'    pos% = InStr(picnam$, ".")
'    picext$ = Mid$(picnam$, pos% + 1, 3)
'    RSfilnam$ = Mid$(picnam$, 1, pos% - 1) & "-RS" & ".txt"
'
'    If RSopenedfile Then
'       Close #RSfilnum%
'       End If
'
'    RSfilnum% = FreeFile
'    Open RSfilnam$ For Output As #RSfilnum%
'    RSopenedfile = True
'
'   For i& = 0 To numRS - 1
'       Write #RSfilnum%, RS(i&).xScreen, RS(i&).yScreen, RS(i&).XGeo, RS(i&).YGeo
'   Next i&
'
'   Close #RSfilnum%
   
   
   If RSMethod1 = False And RSMethod2 = False Then
      RSMethod1 = True
      RSMethod2 = False
      End If
   
   If RSMethod1 Then
      ier = RS_convert_init()
   ElseIf RSMethod2 Then
      If ULGeoX <> ULGridX Or LRGeoX <> LRGridX Or ULGeoY <> ULGridY Or LRGeoY <> LRGridY Then
         'corners not defined on not on the grid, so will need to use interpolation for areas not within the defined grid intersections
         ier = RS_convert_init() 'interpolation method initialization
         If ier = 0 Then RSMethodBoth = True
         If ier < 0 Then Exit Sub
         End If
      ier = Step1toStep2 'rubber sheeting intialization
      End If
   
   If ier <> 0 Then
      Call MsgBox("Conversion wasn't successful!", vbExclamation, "Rubber Sheeting")
   Else
      DigiRubberSheeting = True 'convert mouse coordinates to geo coordinates and display them
      
      ier = ReDrawMap(0) 'redraw the map
      
      If DigitizeOn Then 'redraw the digitized points
         'load previously recorded digitizing results
'         ier = ReDrawMap(0)
          If Not InitDigiGraph Then
             InputDigiLogFile 'load up saved digitizing data for the current map sheet
          Else
             ier = RedrawDigiLog
             End If
         End If
'      If DigitizeOn Then 'redraw any digitized points
'         ReadRSfile
'         End If

    
        If buttonstate&(36) = 1 Then
           If DigitizeMagvis Then 'remagnify screen
              GDDigiMagfrm.Visible = True
           Else
              DigitizeMagInit = True
              GDDigiMagfrm.Visible = True
              End If
           End If
           
      'enable Hardy analysis
      GDMDIform.Toolbar1.Buttons(43).Enabled = True
           
      If ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
          (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) Then
          'enable GPS button
          GDMDIform.Toolbar1.Buttons(34).Enabled = True
          End If
          
      If (heights Or BasisDTMheights) And (RSMethod1 Or RSMethod2 Or RSMethod0) And DigiRubberSheeting Then

          'enable search height button
          GDMDIform.Toolbar1.Buttons(50).Enabled = True
          'enable contour generation
          GDMDIform.Toolbar1.Buttons(51).Enabled = True
          
          GDMDIform.Label1 = lblX
          GDMDIform.Label5 = lblX
          GDMDIform.Label2 = LblY
          GDMDIform.Label6 = LblY
          
          GDMDIform.Text3.Visible = True
          GDMDIform.Label3.Visible = True
          GDMDIform.Text7.Visible = True
          GDMDIform.Label7.Visible = True
            
          GDMDIform.Text4.Visible = True
          GDMDIform.Label4.Visible = True
          
          End If
            
      buttonstate&(39) = 1
      GDMDIform.Toolbar1.Buttons(39).value = tbrPressed
      GDform1.Picture2.SetFocus
    
      Unload Me
      
      End If
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdNext_Click
' Author    : Dr-John-K-Hall
' Date      : 5/27/2015
' Purpose   : Records RS positions
'---------------------------------------------------------------------------------------
'
Private Sub cmdNext_Click(Index As Integer)
    'record last value and display new coordinate
    Dim RoundOffX As Double, RoundOffY As Double, i As Long
    
   On Error GoTo cmdNext_Click_Error

    If GDRSfrm.txtGeoX <> sEmpty And GDRSfrm.txtGeoY <> sEmpty And _
       GDRSfrm.txtX <> sEmpty And GDRSfrm.txtY <> sEmpty And _
       IsNumeric(val(GDRSfrm.txtGeoX)) And IsNumeric(val(GDRSfrm.txtGeoY)) Then
       
       If numRS = 0 Then
          ReDim RS(0)
       Else
          'check for repeat mistake
          ReDim Preserve RS(numRS)
          End If
          
       RS(numRS).xScreen = CLng(val(GDRSfrm.txtX.Text))
       RS(numRS).yScreen = CLng(val(GDRSfrm.txtY.Text))
       RS(numRS).XGeo = val(JustConvertDegToNumber(GDRSfrm.txtGeoX.Text))
       RS(numRS).YGeo = val(JustConvertDegToNumber(GDRSfrm.txtGeoY.Text))
       
       'now record this onto the temporary file
'       pos% = InStr(picnam$, ".")
'       picext$ = Mid$(picnam$, pos% + 1, 3)
       RSfilnam$ = App.Path & "\" & RootName(picnam$) & "-RS" & ".txt"

       If RSopenedfile Or numRS > 0 Then
          If RSfilnum% > 0 Then
             Close #RSfilnum%
             End If
          RSfilnum% = FreeFile
          Open RSfilnam$ For Append As #RSfilnum%
       ElseIf numRS = 0 Then
          RSfilnum% = FreeFile
          Open RSfilnam$ For Output As #RSfilnum%
          Write #RSfilnum%, DigiRSStepType
          RSopenedfile = True
          End If
          
       Write #RSfilnum%, RS(numRS).xScreen, RS(numRS).yScreen, RS(numRS).XGeo, RS(numRS).YGeo
       
       'mark the points done on the map
       gddm = GDform1.Picture2.DrawMode
       gddw = GDform1.Picture2.DrawWidth
       GDform1.Picture2.DrawMode = 13
       GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
       GDform1.Picture2.Line (RS(numRS).xScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)), RS(numRS).yScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)))-(RS(numRS).xScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)), RS(numRS).yScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom))), RSColor& 'QBColor(14)
       GDform1.Picture2.Line (RS(numRS).xScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)), RS(numRS).yScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)))-(RS(numRS).xScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)), RS(numRS).yScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom))), RSColor& 'QBColor(14)
       GDform1.Picture2.DrawMode = gddm
       GDform1.Picture2.DrawWidth = gddw
       
       numRS = numRS + 1 'increment counter
       
       If numRS = NX_CALDAT * NY_CALDAT Then 'enable rubber sheeting calculation
          Close #RSfilnum%
          RSopenedfile = False
          
          'make backup
          Dim BackupRSfile$
          BackupRSfile$ = RootName(RSfilnam$) & "-" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & ".bak"
          FileCopy RSfilnam$, BackupRSfile$
          
          'now reopen and write all the data
          RSfilnum% = FreeFile
          Open RSfilnam$ For Output As #RSfilnum%
          Write #RSfilnum%, DigiRSStepType
          RSopenedfile = True
          
          For i = 0 To numRS - 1
              Write #RSfilnum%, RS(i).xScreen, RS(i).yScreen, RS(i).XGeo, RS(i).YGeo
          Next i
          
          Close #RSfilnum%
          RSopenedfile = False
          
          GDRSfrm.cmdConvert.Enabled = True
          MsgBox "You have digitized all the grid intersections." _
                 & vbCrLf & vbCrLf & "Hint: Press the ''Activate Calculation Method'' button to finish.", _
                 vbInformation + vbOKOnly, "Completion of grid digitizing"
          Exit Sub
          End If

       End If
            
    RoundOffX = 0
    RoundOffY = 0
                     
    If NX_CALDAT <> 0 And NY_CALDAT <> 0 And numRS < NX_CALDAT * NY_CALDAT Then
       If (StepInX Or GDRSfrm.chkStepX.value = vbChecked) Then
          XGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepX))
          YGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepY))
          If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60)
          If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60)
          AA = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) + XGridSteps
          BB = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) - YGridSteps
          If AA >= ULGridX - RoundOffX And AA <= LRGridX + RoundOffX Then
             GDRSfrm.txtGeoX = JustConvertDegToNumber(GDRSfrm.txtGeoX) + XGridSteps
          ElseIf AA >= LRGridX + RoundOffX And BB >= LRGridY - RoundOffY Then
             GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) - YGridSteps
             GDRSfrm.txtStepX = -val(JustConvertDegToNumber(GDRSfrm.txtStepX))
          ElseIf AA <= ULGridX - RoundOffX And BB >= LRGridY - RoundOffY Then
             GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) - YGridSteps
             GDRSfrm.txtStepX = -val(JustConvertDegToNumber(GDRSfrm.txtStepX))
             End If
          End If
       If (StepInY Or GDRSfrm.chkStepY.value = vbChecked) Then
          XGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepX))
          YGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepY))
          If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60#)
          If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60#)
          AA = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
          BB = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
          If AA >= LRGridY - RoundOffY And AA <= ULGridY + RoundOffY Then
             GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
          ElseIf AA >= ULGridY + RoundOffY And BB >= ULGridX - RoundOffX Then
             GDRSfrm.txtGeoX = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
             GDRSfrm.txtStepY = -val(JustConvertDegToNumber(GDRSfrm.txtStepY))
          ElseIf AA <= LRGridY - RoundOffY And BB >= ULGridX - RoundOffX Then
             GDRSfrm.txtGeoX = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
             GDRSfrm.txtStepY = -val(JustConvertDegToNumber(GDRSfrm.txtStepY))
             End If
          End If
       End If


   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdNext_Click of Form GDRSfrm"

End Sub

Private Sub cmdUndo_Click()

    Dim RoundOffX As Double, RoundOffY As Double

   'undo last point on list
   Dim ier As Integer
   Dim NCol As Long
   Dim NRow As Long
   Dim SignStepX As Integer
   Dim SignStepY As Integer
   
   If numRS > 0 Then
   
'        Select Case MsgBox("Undo last point?" _
'                           & vbCrLf & "" _
'                           & vbCrLf & "(This operation is not reversible)" _
'                           , vbOKCancel Or vbQuestion Or vbDefaultButton1, "Undo last point")
'
'         Case vbOK

            TmpX = RS(numRS - 1).XGeo
            tmpY = RS(numRS - 1).YGeo
            TmpStepX = txtStepX
            TmpStepY = txtStepY
            
            ier = ReDrawMap(2)
            ier = InputGuideLines
            
            numRS = numRS - 1
            ReDim Preserve RS(numRS)
            
            If RSopenedfile Then
               RSopenedfile = False
               Close #RSfilnum%
               End If
               
            RSfilnum% = FreeFile
'            pos% = InStr(picnam$, ".")
'            picext$ = Mid$(picnam$, pos% + 1, 3)
            RSfilnam$ = App.Path & "\" & RootName(picnam$) & "-RS" & ".txt"
            RSfilnum% = FreeFile
            
            RSfilnum% = FreeFile
            Open RSfilnam$ For Output As #RSfilnum%
            Write #RSfilnum%, DigiRSStepType
            
            gddm = GDform1.Picture2.DrawMode
            gddw = GDform1.Picture2.DrawWidth
            GDform1.Picture2.DrawMode = 13
            GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
            
            For i& = 1 To numRS
                'mark the points done on the map
                GDform1.Picture2.Line (RS(i& - 1).xScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)), RS(i& - 1).yScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)))-(RS(i& - 1).xScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)), RS(i& - 1).yScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom))), RSColor& 'QBColor(14)
                GDform1.Picture2.Line (RS(i& - 1).xScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)), RS(i& - 1).yScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)))-(RS(i& - 1).xScreen * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)), RS(i& - 1).yScreen * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom))), RSColor& 'QBColor(14)
                Write #RSfilnum%, RS(i& - 1).xScreen, RS(i& - 1).yScreen, RS(i& - 1).XGeo, RS(i& - 1).YGeo
            Next i&
            
            Close #RSfilnum%
            
            GDform1.Picture2.DrawMode = gddm
            GDform1.Picture2.DrawWidth = gddw

            If numRS > 0 Then
              'shift map to place of erased mark
               ce& = 1 'this flag forces the pointer to move to the next shiftmap coordinate
               Call ShiftMap(CSng(RS(numRS - 1).xScreen * DigiZoom.LastZoom), CSng(RS(numRS - 1).yScreen * DigiZoom.LastZoom))
               End If
               
            'refresh coordinates
            txtGeoX = TmpX
            txtGeoY = tmpY
            txtStepX = TmpStepX
            txtStepY = TmpStepY
            
           If NX_CALDAT > 1 And NY_CALDAT > 1 Then
             XGridSteps = (LRGridX - ULGridX) / (NX_CALDAT - 1)
             YGridSteps = (ULGridY - LRGridY) / (NY_CALDAT - 1)
             End If
            
            'now determine the right sign of the steps
            If StepInX Then
               NRow = Fix((numRS - 1) / NX_CALDAT) + 1
                
                If NRow Mod 2 = 0 Then
                   NCol = NRow * (NX_CALDAT) - numRS + 1
                Else
                   NCol = numRS - (NX_CALDAT) * (NRow - 1)
                   End If
                   
                SignStepX = (-1) ^ (NRow - 1)
'                SignStepY = (-1) ^ (NCol - 1)
                
                If (NRow = 1 Or NRow Mod 2 <> 0) And NCol = NX_CALDAT Then 'reverse sign again
                   SignStepX = -1 '-1 * SignStepX
                   End If
                   
                If NCol = 1 And NRow Mod 2 = 0 Then 'reverse sign again
                   SignStepX = 1 '-1 * SignStepX
                   End If
                   
            ElseIf StepInY Then
                NCol = Fix((numRS - 1) / NY_CALDAT) + 1
                
                If NCol Mod 2 = 0 Then
                   NRow = NCol * (NY_CALDAT) - numRS + 1
                Else
                   NRow = numRS - (NY_CALDAT) * (NCol - 1)
                   End If
                   
'                SignStepX = (-1) ^ (NRow - 1)
'                SignStepY = (-1) ^ (NCol - 1)
                SignStepY = (-1) ^ NCol
                   
                If (NCol = 1 Or NCol Mod 2 <> 0) And NRow = NY_CALDAT Then 'reverse sign again
                   SignStepY = 1 '* SignStepY
                   End If
                   
'                If NCol = 1 Or NCol Mod 2 <> 0 And NRow = NY_CALDAT Then
'                   SignStepY = -1
'                   End If
                   
                If NRow = 1 And NCol Mod 2 = 0 Then 'reverse sign again
                   SignStepY = -1
                   End If
                   
                End If
               
            
            If StepInX Then
               If SignStepX = -1 Then
                  GDRSfrm.txtStepX = Trim$(str$(SignStepX * XGridSteps))
               Else
                  GDRSfrm.txtStepX = Trim$(str$(XGridSteps))
                  End If
               End If
            If StepInY Then
               If SignStepY = -1 Then
                  GDRSfrm.txtStepY = Trim$(str$(SignStepY * YGridSteps))
               Else
                  GDRSfrm.txtStepY = Trim$(str$(YGridSteps))
                  End If
               End If
                     
            If lblX = "lon." And LblY = "lat." Then
               GDRSfrm.txtStepX = JustConvertDegToNumber(GDRSfrm.txtStepX)
               GDRSfrm.txtStepY = JustConvertDegToNumber(GDRSfrm.txtStepY)
               End If
            

'                RoundOffX = 0
'                RoundOffY = 0
'
'                'adjust values of wizard
'                If GDRSfrmVis Then
'                   If StepInX Or chkStepX.value = vbChecked Then
'                      XGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepX))
'                      YGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepY))
'                      If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60#)
'                      If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60#)
'                      AA = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
'                      BB = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
'                      If AA >= ULGridX - RoundOffX And AA <= LRGridX + RoundOffX Then
'                         GDRSfrm.txtGeoX = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
'                      ElseIf AA < ULGridX + RoundOffX And BB <= ULGridY + RoundOffY Then
'                         GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
'                         GDRSfrm.txtStepX = -val(JustConvertDegToNumber(GDRSfrm.txtStepX))
'                      ElseIf AA > LRGridX - RoundOffX And BB <= ULGridY + RoundOffY Then
'                         GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
'                         GDRSfrm.txtStepX = -val(JustConvertDegToNumber(GDRSfrm.txtStepX))
'                         End If
'                   ElseIf StepInY Or chkStepY.value = vbChecked Then
'                      XGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepX))
'                      YGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepY))
'                      If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60#)
'                      If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60#)
'                      AA = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
'                      BB = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
'                      If AA >= LRGridY - RoundOffY And AA <= ULGridY + RoundOffY Then
'                         GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
'                      ElseIf AA > ULGridY - RoundOffY And BB >= ULGridX - RoundOffX Then
'                         GDRSfrm.txtGeoX = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
'                         GDRSfrm.txtStepY = -val(JustConvertDegToNumber(GDRSfrm.txtStepY))
'                      ElseIf AA < LRGridY + RoundOffY And BB >= ULGridX - RoundOffX Then
'                         GDRSfrm.txtGeoX = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
'                         GDRSfrm.txtStepY = -val(JustConvertDegToNumber(GDRSfrm.txtStepY))
'                         End If
'                      End If
'                   End If
                End If
        
      
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdXConvert_Click
' Author    : Chaim Keller
' Date      : 2/13/2015
' Purpose   : converts degrees-minutes-seconds into decimal degrees
'---------------------------------------------------------------------------------------
'
Private Sub cmdXConvert_Click()
  If lblX = "lon." And LblY = "lat." Then
     txtGeoX = ConvertDegToNumber(txtGeoX)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                & vbCrLf & "" _
                & vbCrLf & "(Hint: set the coordinates labels in the ""Options menu"" to "".lon."", "".lat"")" _
                , vbInformation, "Coordinate Conversion Error")

     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdXStepConvert_Click
' Author    : Chaim Keller
' Date      : 2/13/2015
' Purpose   : converts degrees-minutes-seconds into decimal degree
'---------------------------------------------------------------------------------------
'
Private Sub cmdXStepConvert_Click()
  If lblX = "lon." And LblY = "lat." Then
     txtStepX = ConvertDegToNumber(txtStepX)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                & vbCrLf & "" _
                & vbCrLf & "(Hint: set the coordinates labels in the ""Options menu"" to "".lon."", "".lat"")" _
                , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdYConvert_Click
' Author    : Chaim Keller
' Date      : 2/13/2015
' Purpose   : convert from degrees-minutes-seconds to decimal degrees
'---------------------------------------------------------------------------------------
'
Private Sub cmdYConvert_Click()
  If lblX = "lon." And LblY = "lat." Then
     txtGeoY = ConvertDegToNumber(txtGeoY)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                & vbCrLf & "" _
                & vbCrLf & "(Hint: set the coordinates labels in the ""Options menu"" to "".lon."", "".lat"")" _
                , vbInformation, "Coordinate Conversion Error")
     End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdYStepConvert_Click
' Author    : Chaim Keller
' Date      : 2/13/2015
' Purpose   : convert from degrees-minutes-seconds to decimal degrees
'---------------------------------------------------------------------------------------
'
Private Sub cmdYStepConvert_Click()
  If lblX = "lon." And LblY = "lat." Then
     txtStepY = ConvertDegToNumber(txtStepY)
  Else
     Call MsgBox("Only degrees latitude and longitude can be converted!" _
                & vbCrLf & "" _
                & vbCrLf & "(Hint: set the coordinates labels in the ""Options menu"" to "".lon."", "".lat"")" _
                , vbInformation, "Coordinate Conversion Error")
     End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

   If KeyAscii = Asc("Z") Or KeyAscii = Asc("z") Then
      'zoom out
      Call PictureBoxZoom(GDform1.Picture2, 0, -120, 0, 0, 0)
      End If
      
   If KeyAscii = Asc("X") Or KeyAscii = Asc("x") Then
      'zoom in
      Call PictureBoxZoom(GDform1.Picture2, 0, 120, 0, 0, 0)
      End If

End Sub

Private Sub Form_Load()

   Dim ier As Integer
   Dim SignXStep As Integer
   Dim SignYStep As Integer

   GDRSfrmVis = True
   
   Call sCenterForm(Me)
   
   Dim Ret As Long
   Ret = BringWindowToTop(GDRSfrm.hwnd)
   
'   ReadRSfile
   
   If NX_CALDAT <> 0 And NY_CALDAT <> 0 Then
      
        If Not DigiRubberSheeting Or RSMethod0 Then
           If Not RSMethod0 Then
              chkRS.value = vbChecked
           Else
              chkSimple.value = vbChecked
              End If
           End If
      
        With GDRSfrm
           '------fancy progress bar settings---------
           .picProgBar.AutoRedraw = True
           .picProgBar.BackColor = &H8000000B 'light grey
           .picProgBar.DrawMode = 10
         
           .picProgBar.FillStyle = 0
           .picProgBar.ForeColor = &H400000 'dark blue
           
           If NX_CALDAT > 1 And NY_CALDAT > 1 Then
             XGridSteps = (LRGridX - ULGridX) / (NX_CALDAT - 1)
             YGridSteps = (ULGridY - LRGridY) / (NY_CALDAT - 1)
             End If
           
           If NX_CALDAT = 1 Then
             XGridSteps = 0
             End If
             
           If NY_CALDAT = 1 Then
              YGridSteps = 0
              End If
           
           If DigiRSStepType = 0 And XGridSteps <> 0 Then
              StepInX = True
              StepInY = False
           ElseIf DigiRSStepType = 0 And XGridSteps = 0 Then
              If YGridSteps <> 0 Then
                 StepInX = False
                 StepInY = True
                 DigiRSStepType = 1
                 End If
              End If
           
           If DigiRSStepType = 1 And YGridSteps <> 0 Then
              StepInY = True
              StepInX = False
           ElseIf DigiRSStepType = 1 And YGridSteps = 0 Then
              If XGridSteps <> 0 Then
                 StepInY = False
                 StepInX = True
                 DigiRSStepType = 0
                 End If
              End If
                 
            .txtStepX.Text = ConvertCoordToString(XGridSteps)
            .txtStepY.Text = ConvertCoordToString(YGridSteps)
                 
'           If XGridSteps = 0 And YGridSteps <> 0 Then StepInY = True
'           If XGridSteps <> 0 And YGridSteps = 0 Then StepInX = True
           
           Select Case DigiRSStepType
              Case 0
                 StepInX = True
              Case 1
                 StepInY = True
           End Select
           
           If StepInX Then
              chkStepX.value = vbChecked
              chkStepY.value = vbUnchecked
           ElseIf StepInY Then
              chkStepY.value = vbChecked
              chkStepX.value = vbUnchecked
              End If
              
           If RSMethod2 Then
              optRS.value = True
           ElseIf RSMethod1 Then
              optDefault.value = True
              End If
              
           If lblX = "lon." And LblY = "lat." Then
              .cmdXConvert.Enabled = True
              .cmdXStepConvert.Enabled = True
              .cmdYConvert.Enabled = True
              .cmdYStepConvert.Enabled = True
              End If
              
'            pos% = InStr(picnam$, ".")
'            picext$ = Mid$(picnam$, pos% + 1, 3)
            RSfilnam$ = App.Path & "\" & RootName(picnam$) & "-RS" & ".txt"
              
            If Trim$(Dir(RSfilnam$)) = sEmpty Then 'starting digitizing grid intersections, display hint
              
               .chkStepX.Enabled = True
               .chkStepY.Enabled = True
               
               End If
               
            If numRS = 0 Then
            
              'set defaults for Geo coordinates to start from (x1,y1)
              If StepInX Then
                .txtGeoX = ULGridX
                .txtGeoY = ULGridY
              ElseIf StepInY Then
                .txtGeoX = LRGridX
                .txtGeoY = LRGridY
                End If
                
            ElseIf numRS > 0 Then
            
'                'move map to last RS point done
                 .chkStepX.Enabled = False
                 .chkStepY.Enabled = False
                
                 If GDRSfrmVis And (GeoMap Or TopoMap) And numRS > 0 Then
                    GDRSfrm.Visible = True
                    Ret = BringWindowToTop(GDRSfrm.hwnd)
                    End If
                  
                 If numRS <> NX_CALDAT * NY_CALDAT Then
                    'move map to last RS point in the file
                    Call ShiftMap(CSng(RS(numRS - 1).xScreen * DigiZoom.LastZoom), CSng(RS(numRS - 1).yScreen * DigiZoom.LastZoom)) 'move map to last point
                 Else
                    cmdConvert.Enabled = True
                    End If
                     
                  GDRSfrm.txtX = RS(numRS - 1).xScreen
                  GDRSfrm.txtY = RS(numRS - 1).yScreen
                  
                  GDRSfrm.txtGeoX = ConvertCoordToString(RS(numRS - 1).XGeo)
                  GDRSfrm.txtGeoY = ConvertCoordToString(RS(numRS - 1).YGeo)
                  
                  'now determine the right sign of the steps
                  NRow = Fix((numRS - 1) / NX_CALDAT) + 1
                    
                  If NRow Mod 2 = 0 Then
                     NCol = NRow * (NX_CALDAT) - numRS + 1
                  Else
                     NCol = numRS - (NX_CALDAT) * (NRow - 1)
                     End If
                  
                  SignStepX = (-1) ^ (NRow - 1)
                  SignStepY = (-1) ^ (NCol - 1)
                  If StepInX Then
                     If SignStepX = -1 Then
                        GDRSfrm.txtStepX = Trim$(str$(SignStepX * val(JustConvertDegToNumber(GDRSfrm.txtStepX))))
                     Else
                        GDRSfrm.txtStepX = Trim$(str$(val(Abs(JustConvertDegToNumber(GDRSfrm.txtStepX)))))
                        End If
                     End If
                  If StepInY Then
                     If SignStepY = -1 Then
                        GDRSfrm.txtStepY = Trim$(str$(SignStepY * val(JustConvertDegToNumber(GDRSfrm.txtStepY))))
                     Else
                       GDRSfrm.txtStepY = Trim$(str$(val(Abs(JustConvertDegToNumber(GDRSfrm.txtStepY)))))
                       End If
                     End If

                 GDRSfrm.txtStepX = ConvertDegToNumber(GDRSfrm.txtStepX)
                 GDRSfrm.txtStepY = ConvertDegToNumber(GDRSfrm.txtStepY)
                  
                RoundOffX = 0
                RoundOffY = 0
                                 
                If NX_CALDAT <> 0 And NY_CALDAT <> 0 And numRS < NX_CALDAT * NY_CALDAT Then
                   If (StepInX Or GDRSfrm.chkStepX.value = vbChecked) Then
                      XGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepX))
                      YGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepY))
                      If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60)
                      If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60)
                      AA = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) + XGridSteps
                      BB = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) - YGridSteps
                      If AA >= ULGridX - RoundOffX And AA <= LRGridX + RoundOffX Then
                         GDRSfrm.txtGeoX = JustConvertDegToNumber(GDRSfrm.txtGeoX) + XGridSteps
                      ElseIf AA >= LRGridX + RoundOffX And BB >= LRGridY - RoundOffY Then
                         GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) - YGridSteps
                         GDRSfrm.txtStepX = -val(JustConvertDegToNumber(GDRSfrm.txtStepX))
                      ElseIf AA <= ULGridX - RoundOffX And BB >= LRGridY - RoundOffY Then
                         GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) - YGridSteps
                         GDRSfrm.txtStepX = -val(JustConvertDegToNumber(GDRSfrm.txtStepX))
                         End If
                      End If
                   If (StepInY Or GDRSfrm.chkStepY.value = vbChecked) Then
                      XGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepX))
                      YGridSteps = val(JustConvertDegToNumber(GDRSfrm.txtStepY))
                      If lblX = "lon." And LblY = "lat." Then RoundOffX = Abs(XGridSteps / 60#)
                      If lblX = "lon." And LblY = "lat." Then RoundOffY = Abs(YGridSteps / 60#)
                      AA = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
                      BB = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
                      If AA >= LRGridY - RoundOffY And AA <= ULGridY + RoundOffY Then
                         GDRSfrm.txtGeoY = val(JustConvertDegToNumber(GDRSfrm.txtGeoY)) + YGridSteps
                      ElseIf AA >= ULGridY + RoundOffY And BB >= ULGridX - RoundOffX Then
                         GDRSfrm.txtGeoX = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
                         GDRSfrm.txtStepY = -val(JustConvertDegToNumber(GDRSfrm.txtStepY))
                      ElseIf AA <= LRGridY - RoundOffY And BB >= ULGridX - RoundOffX Then
                         GDRSfrm.txtGeoX = val(JustConvertDegToNumber(GDRSfrm.txtGeoX)) - XGridSteps
                         GDRSfrm.txtStepY = -val(JustConvertDegToNumber(GDRSfrm.txtStepY))
                         End If
                      End If
                   End If
                  
                 End If
                 
              
'        If lblX = "lon." And LblY = "lat." Then
'           'convert coordinates back into degrees, minutes, and seconds
'           GDRSfrm.txtGeoX = ConvertCoordToString(val(GDRSfrm.txtGeoX))
'           GDRSfrm.txtGeoY = ConvertCoordToString(val(GDRSfrm.txtGeoY))
'           GDRSfrm.txtStepX = ConvertCoordToString(val(GDRSfrm.txtStepX))
'           GDRSfrm.txtStepY = ConvertCoordToString(val(GDRSfrm.txtStepY))
'           End If
              
        End With
        
     Else
     
        RSMethod1 = False
        RSMethod2 = False
        If (ULGeoX = LRGeoX And ULGeoY = LRGeoY) Or (ULPixX = LRPixX And ULPixY = LRPixY) Then
           RSMethod0 = False
           End If
'        Else
'           RSMethod0 = True
'           End If
        
        End If
        
'    'draw the extra guide lines
'    ier = InputGuideLines
'
'    ReadRSfile
'
'    'move to last point
'    If numRS > 0 Then
'       Call ShiftMap(CSng(RS(numRS - 1).xScreen * DigiZoom.LastZoom), CSng(RS(numRS - 1).yScreen * DigiZoom.LastZoom)) 'move map to last point
'       End If
'
        
'    Call WheelHook(Me.hWnd)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
   GDRSfrmVis = False
   Set GDRSfrm = Nothing
   
   If RSopenedfile Then
      Close #RSfilnum%
      numRS = 0
      ReDim RS(numRS) 'reclaim memory
      End If

   DigiRS = False
   
   If Not DigiRubberSheeting And Not RSMethod0 And buttonstate&(39) = 1 Then
'      'unpress button
'      buttonstate&(39) = 0
'      GDMDIform.Toolbar1.Buttons(39).value = tbrUnpressed
   
      GDMDIform.mnuDigitizeRubberSheeting_Click
      End If
      
   If DigitizeMagvis Then
      MagFrmSize = GDDigiMagfrm.Width
      PicFrmSize = GDform1.Width
      End If
      
   ier = ReDrawMap(0) 'remove all the marks used for the rubber sheeting digitizing
   
   If DigitizeOn Then
       If Not InitDigiGraph Then
          InputDigiLogFile 'load up saved digitizing data for the current map sheet
       Else
          ier = RedrawDigiLog
          End If
      End If
      
    waitime = Timer
    Do Until Timer > waitime + 0.1
       DoEvents
    Loop
    
    If DigitizeMagvis Then
       GDDigiMagfrm.Width = MagFrmSize
       GDform1.Width = PicFrmSize
       End If
'   Call WheelUnHook(Me.hWnd)
   
End Sub


Private Sub optDefault_Click()
   'use default type of coordinate calculation
   If optDefault.value Then
      optRS.value = False
      RSMethod1 = True
      RSMethod2 = False
      cmdConvert.Enabled = True
      RSMethodBoth = False
      End If
End Sub

Private Sub optRS_Click()
   'use JKHalls rubber sheeting
   If optRS.value Then
      optDefault.value = False
      RSMethod1 = False
      RSMethod2 = True
'      If Not DigiRubberSheeting Then ReadRSfile 'plot grid extension lines
      End If
End Sub

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' source : wheel mouse hook: http://www.vbforums.com/showthread.php?388222-VB6-MouseWheel-with-Any-Control-%28originally-just-MSFlexGrid-Scrolling%29
'          two files examples: WheelHook-AllControls.zip, WheelHook-NestedControls.zip

' ===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)

  PictureBoxZoom GDform1.Picture2, MouseKeys, Rotation, Xpos, Ypos, 0

End Sub


