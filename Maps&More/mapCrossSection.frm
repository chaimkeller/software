VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form mapCrossSection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cross Section Entries Wizard"
   ClientHeight    =   6615
   ClientLeft      =   6975
   ClientTop       =   1875
   ClientWidth     =   4680
   Icon            =   "mapCrossSection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   4680
   Begin VB.Frame Frame6 
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   5520
      Width           =   4395
      Begin MSComctlLib.ProgressBar CSectProgressBar 
         Height          =   255
         Left            =   100
         TabIndex        =   29
         Top             =   160
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblObstruction 
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
         Height          =   195
         Left            =   840
         TabIndex        =   28
         Top             =   180
         Width           =   2835
      End
   End
   Begin VB.Frame Frame5 
      Enabled         =   0   'False
      Height          =   555
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   4395
      Begin VB.TextBox txtNumPoints 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2820
         TabIndex        =   25
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label lblNumPoints 
         Caption         =   "&Number of Points in Cross Section"
         Enabled         =   0   'False
         Height          =   255
         Left            =   300
         TabIndex        =   26
         Top             =   195
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   315
      Left            =   780
      TabIndex        =   23
      Top             =   6180
      Width           =   1035
   End
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   4395
      Begin VB.CheckBox chkObsHeight 
         Caption         =   "&Add 1.6 meters to 1st Point for observer's height"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   300
         Width           =   3795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Second Point"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   9
      Top             =   1740
      Width           =   4395
      Begin VB.TextBox txthgt2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   19
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox txtlat2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   12
         Top             =   660
         Width           =   1755
      End
      Begin VB.TextBox txtlon2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   11
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton cmdSecondPoint 
         Caption         =   "&2nd Point"
         Enabled         =   0   'False
         Height          =   915
         Left            =   3120
         Picture         =   "mapCrossSection.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Accept current map coordinates"
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label lblhgt2 
         Caption         =   "height (m)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label lbllat2 
         Caption         =   "latitude"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbllon2 
         Caption         =   "longitude"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   6180
      Width           =   915
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Top             =   6180
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cross Section Trajectory"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      TabIndex        =   1
      Top             =   3420
      Width           =   4395
      Begin VB.OptionButton OptionProject 
         Caption         =   "&Minimum distance projection on spherical earth"
         Enabled         =   0   'False
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   480
         Width           =   3615
      End
      Begin VB.OptionButton OptionMercator 
         Caption         =   "&Straight line on Mercator Projection"
         Enabled         =   0   'False
         Height          =   195
         Left            =   420
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "First Point"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4395
      Begin VB.TextBox txthgt1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1020
         TabIndex        =   20
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdFirstPoint 
         Caption         =   "&1st Point"
         Height          =   915
         Left            =   3120
         Picture         =   "mapCrossSection.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Accept current map coordinates"
         Top             =   420
         Width           =   1035
      End
      Begin VB.TextBox txtlat1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         TabIndex        =   5
         Top             =   660
         Width           =   1755
      End
      Begin VB.TextBox txtlon1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         TabIndex        =   4
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lblhgt1 
         Caption         =   "height (m)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lbllat1 
         Caption         =   "latitude"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbllon1 
         Caption         =   "longitude"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "mapCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nextstep%

Private Sub chkObsHeight_Click()
   If chkObsHeight.value = vbChecked Then
      ObsHeight = True
      End If
End Sub

Private Sub cmdBack_Click()
    Select Case nextstep%
       Case 1
            lbllon1.Enabled = True
            lbllat1.Enabled = True
            lblhgt1.Enabled = True
            txtlon1.Enabled = True
            txtlat1.Enabled = True
            txthgt1.Enabled = True
            cmdFirstPoint.Enabled = True
            Frame1.Enabled = True
            
            Frame2.Enabled = False
            lbllon2.Enabled = False
            lbllat2.Enabled = False
            lblhgt2.Enabled = False
            txtlon2.Enabled = False
            txtlat2.Enabled = False
            txthgt2.Enabled = False
            cmdSecondPoint.Enabled = False
            lblObstruction.Caption = "Click on the 1st point"
        
            cmdFirstPoint.SetFocus
            cmdBack.Enabled = False
          
       Case 2
             
            lbllon2.Enabled = True
            lbllat2.Enabled = True
            lblhgt2.Enabled = True
            txtlon2.Enabled = True
            txtlat2.Enabled = True
            txthgt2.Enabled = True
            cmdSecondPoint.Enabled = True
            Frame2.Enabled = True
            
            If world = False Then
               nextstep% = nextstep% - 1
               Frame4.Enabled = False
               chkObsHeight.Enabled = False
               cmdSecondPoint.SetFocus
               Exit Sub
               End If
            Frame3.Enabled = False
            OptionProject.Enabled = False
            OptionMercator.Enabled = False
            greatcircle = False
            lblObstruction.Caption = "Click on the 2nd point"
            cmdSecondPoint.SetFocus
       Case 3
            Frame3.Enabled = True
            OptionProject.Enabled = True
            OptionMercator.Enabled = True
            
            Frame4.Enabled = False
            chkObsHeight.Enabled = False
            OptionMercator.SetFocus
       Case 4
           Frame4.Enabled = True
           chkObsHeight.Enabled = True
           
           Frame5.Enabled = False
           lblNumPoints.Enabled = False
           txtNumPoints.Enabled = False
           
           txtNumPoints.Text = sEmpty
           chkObsHeight.SetFocus
       Case 5
           Frame5.Enabled = True
           lblNumPoints.Enabled = True
           txtNumPoints.Enabled = True
           
           Frame6.Enabled = False
           lblObstruction.Enabled = False
           cmdNext.Enabled = True
           'txtNumPoints.SetFocus
       Case 6
       Case Else
    End Select
    nextstep% = nextstep% - 1

End Sub

Private Sub cmdCancel_Click()
   If GoCrossSection = True Then
      GoCrossSection = False
      Exit Sub
      End If
   Call form_queryunload(0, 0)
End Sub

Private Sub cmdFirstPoint_Click()
   txtlon1.Text = Maps.Text5.Text
   txtlat1.Text = Maps.Text6.Text
   txthgt1.Text = Maps.Text7.Text
End Sub

Private Sub cmdNext_Click()
    nextstep% = nextstep% + 1
    Select Case nextstep%
       Case 1
          If txtlon1.Text <> sEmpty And txtlat1.Text <> sEmpty Then
             lbllon1.Enabled = False
             lbllat1.Enabled = False
             lblhgt1.Enabled = False
             txtlon1.Enabled = False
             txtlat1.Enabled = False
             txthgt1.Enabled = False
             cmdFirstPoint.Enabled = False
             Frame1.Enabled = False
             
             Frame2.Enabled = True
             lbllon2.Enabled = True
             lbllat2.Enabled = True
             lblhgt2.Enabled = True
             txtlon2.Enabled = True
             txtlat2.Enabled = True
             txthgt2.Enabled = True
             cmdSecondPoint.Enabled = True
             cmdBack.Enabled = True
             lblObstruction.Caption = "Click on the 2nd point"
         
             crosssectionpnt(0, 0) = Val(txtlon1.Text)
             crosssectionpnt(0, 1) = Val(txtlat1.Text)
             crosssectionhgt(0) = Val(txthgt1.Text)
             cmdSecondPoint.SetFocus
          
          Else
             response = MsgBox("You haven't entered the 1st point's coordinates!", vbCritical + vbOKOnly, "Cross Section Entries")
             nextstep% = nextstep% - 1
             cmdFirstPoint.SetFocus
             Exit Sub
             End If
       Case 2
          If txtlon2.Text <> sEmpty And txtlat2.Text <> sEmpty Then
             
             crosssectionpnt(1, 0) = Val(txtlon2.Text)
             crosssectionpnt(1, 1) = Val(txtlat2.Text)
             crosssectionhgt(1) = Val(txthgt2.Text)
         
            If crosssectionpnt(1, 0) = crosssectionpnt(0, 0) And _
               crosssectionpnt(1, 1) = crosssectionpnt(0, 1) Then
               response = MsgBox("The second point must be different from the first point!", vbOKOnly + vbCritical, "Cross Section Wizard")
               nextstep% = nextstep% - 1
               Exit Sub
               End If
           Else
             response = MsgBox("You haven't entered the 2st point's coordinates!", vbCritical + vbOKOnly, "Cross Section Entries")
             nextstep% = nextstep% - 1
             Exit Sub
             End If
         
             
             lbllon2.Enabled = False
             lbllat2.Enabled = False
             lblhgt2.Enabled = False
             txtlon2.Enabled = False
             txtlat2.Enabled = False
             txthgt2.Enabled = False
             cmdSecondPoint.Enabled = False
             Frame2.Enabled = False
             lblObstruction.Caption = ""
             
             If world = False Then
                nextstep% = nextstep% + 1
                Frame4.Enabled = True
                chkObsHeight.Enabled = True
                chkObsHeight.SetFocus
                Exit Sub
                End If
             Frame3.Enabled = True
             OptionProject.Enabled = True
             OptionMercator.Enabled = True
             greatcircle = False
             OptionMercator.SetFocus
       Case 3
            Frame3.Enabled = False
            OptionProject.Enabled = False
            OptionMercator.Enabled = False
            
            Frame4.Enabled = True
            chkObsHeight.Enabled = True
            chkObsHeight.SetFocus
       Case 4
           Frame4.Enabled = False
           chkObsHeight.Enabled = False
           
           Frame5.Enabled = True
           lblNumPoints.Enabled = True
           txtNumPoints.Enabled = True
           
           If world = False Then
               Call casgeo(crosssectionpnt(0, 0), crosssectionpnt(0, 1), lg, lt)
               lg1 = lg
               lt1 = lt
               hgt1 = crosssectionhgt(0)
               Call casgeo(crosssectionpnt(1, 0), crosssectionpnt(1, 1), lg, lt)
               lg2 = lg
               lt2 = lt
               hgt2 = crosssectionhgt(1)
               lg2v = lg2
               lt2v = lt2
           Else
               lg1 = -crosssectionpnt(0, 0)
               lt1 = crosssectionpnt(0, 1)
               hgt1 = crosssectionhgt(0)
               lg2 = -crosssectionpnt(1, 0)
               lt2 = crosssectionpnt(1, 1)
               hgt2 = crosssectionhgt(1)
               lg2v = lg2
               lt2v = lt2
               End If
            
            If world = False Then
               totdist = Sqr((crosssectionpnt(1, 0) - crosssectionpnt(0, 0)) ^ 2 + _
                          (crosssectionpnt(1, 1) - crosssectionpnt(0, 1)) ^ 2)
            Else
               X1 = Cos(lt1 * cd) * Cos(lg1 * cd)
               X2 = Cos(lt2 * cd) * Cos(lg2 * cd)
               Y1 = Cos(lt1 * cd) * Sin(lg1 * cd)
               Y2 = Cos(lt2 * cd) * Sin(lg2 * cd)
               Z1 = Sin(lt1 * cd)
               Z2 = Sin(lt2 * cd)
               'this is a calculation of
               'the shortest geodesic distance and is given by
               'Re * Angle between vectors
               'cos(Angle between unit vectors) = Dot product of unit vectors
               'this is considerably smaller than the straight line distance
               'for large distances.  To calculate that distance you
               'need to use the CrossSection option
               Dim cosang As Double
               cosang = X1 * X2 + Y1 * Y2 + Z1 * Z2
               totdist = 6371315 * DACOS(cosang)
               End If
       
              'ask for the number of points
              If world = False Then
                 defnum& = totdist / 10
              Else
                 If DTMflag = 0 Then
                    defnum& = totdist / 500
                 ElseIf DTMflag = 1 Then
                    defnum& = totdist / 10
                 ElseIf DTMflag = 2 Then
                    defnum& = totdist / 30
                    End If
                 End If
              If defnum& > 32767 Then defnum& = 32767 '16 bit Integer limit
              txtNumPoints.Text = Str(defnum&)
              txtNumPoints.SetFocus
       Case 5
           If Val(txtNumPoints.Text) = 0 Or txtNumPoints.Text = sEmpty Then
              response = MsgBox("Input a non zero number of points!", vbCritical + vbOKOnly, "Cross Section Wizard")
              nextpoint% = nextpoint% - 1
              Exit Sub
              End If
           If Val(txtNumPoints.Text) > 32767 Then
              response = MsgBox("Input a non zero number smaller than 32767", vbCritical + vbOKOnly, "Cross Section Wizard")
              nextpoint% = nextpoint% - 1
              Exit Sub
              End If
           Frame5.Enabled = False
           lblNumPoints.Enabled = False
           txtNumPoints.Enabled = False
           cmdBack.Enabled = False
           cmdNext.Enabled = False
           cmdCancel.Enabled = True
           Frame6.Enabled = True
           lblObstruction.Enabled = True
           GoCrossSection = True 'flag to allow calculations to proceed
           Call mapCrossSections
           
           lblObstruction.Enabled = False
           Frame6.Enabled = False
           cmdNext.Enabled = False
           cmdBack.Enabled = True
           If GoCrossSection = False Then
              'calculations were aborted
              cmdBack_Click
              End If
              
       Case 6
       Case Else
    End Select
End Sub

Private Sub cmdSecondPoint_Click()
   txtlon2.Text = Maps.Text5.Text
   txtlat2.Text = Maps.Text6.Text
   txthgt2.Text = Maps.Text7.Text
End Sub

Private Sub form_load()
  nextstep% = 0
  cmdBack.Enabled = False
  lblObstruction.Caption = "Click on the 1st point"

  If world = True Then
     lbllon1.Caption = "longitude"
     lbllat1.Caption = "latitude"
     lbllon2.Caption = "longitude"
     lbllat2.Caption = "latitude"
  Else
     lbllon1.Caption = "ITMx"
     lbllat1.Caption = "ITMy"
     lbllon2.Caption = "ITMx"
     lbllat2.Caption = "ITMy"
     End If
End Sub

Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
   If crosssection = True Then
      crosssection = False
      Call blitpictures 'erase the path
      End If
   Unload Me
   Set mapCrossSection = Nothing
End Sub

Private Sub OptionMercator_Click()
   greatcircle = False
End Sub

Private Sub OptionProject_Click()
   greatcircle = True
End Sub
