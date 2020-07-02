VERSION 5.00
Begin VB.Form GDDigitizerfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digitizer"
   ClientHeight    =   4935
   ClientLeft      =   1350
   ClientTop       =   3285
   ClientWidth     =   3045
   Icon            =   "GDDigitizerrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmBathymetry 
      Caption         =   "Bathy."
      Height          =   975
      Left            =   2280
      TabIndex        =   25
      Top             =   -10
      Width           =   615
      Begin VB.CheckBox chkOcean 
         Height          =   255
         Left            =   180
         TabIndex        =   26
         ToolTipText     =   "Check to invert sign of the stored elevation (bathymetry)"
         Top             =   440
         Width           =   255
      End
   End
   Begin VB.Frame frmAtritube 
      Height          =   520
      Left            =   120
      TabIndex        =   20
      Top             =   4380
      Width           =   2800
      Begin VB.Label lblFunction 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hardy Quadratic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   350
         Left            =   50
         TabIndex        =   21
         Tag             =   "300"
         Top             =   140
         Width           =   2700
      End
   End
   Begin VB.Frame frmKeyPad 
      Caption         =   "KeyPad"
      Height          =   2610
      Left            =   120
      TabIndex        =   3
      Top             =   1790
      Width           =   2800
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Enter"
         Height          =   390
         Left            =   1920
         TabIndex        =   24
         Top             =   2140
         Width           =   735
      End
      Begin VB.CommandButton cmdSpace 
         Caption         =   "Space"
         Height          =   390
         Left            =   120
         TabIndex        =   23
         Top             =   2140
         Width           =   1695
      End
      Begin VB.CommandButton cmdD 
         Caption         =   "&D"
         Height          =   615
         Left            =   2040
         TabIndex        =   19
         ToolTipText     =   "Add decimal point"
         Top             =   1500
         Width           =   615
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "&C"
         Height          =   615
         Left            =   1440
         TabIndex        =   18
         ToolTipText     =   "Change sign"
         Top             =   1500
         Width           =   615
      End
      Begin VB.CommandButton cmd9 
         Caption         =   "&9"
         Height          =   615
         Left            =   760
         TabIndex        =   17
         Top             =   1500
         Width           =   615
      End
      Begin VB.CommandButton cmd8 
         Caption         =   "&8"
         Height          =   615
         Left            =   170
         TabIndex        =   16
         Top             =   1500
         Width           =   615
      End
      Begin VB.CommandButton cmd7 
         Caption         =   "&7"
         Height          =   615
         Left            =   2040
         TabIndex        =   15
         Top             =   860
         Width           =   615
      End
      Begin VB.CommandButton cmd6 
         Caption         =   "&6"
         Height          =   615
         Left            =   1440
         TabIndex        =   14
         Top             =   860
         Width           =   615
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "&5"
         Height          =   615
         Left            =   760
         TabIndex        =   13
         Top             =   860
         Width           =   615
      End
      Begin VB.CommandButton cmd4 
         Caption         =   "&4"
         Height          =   615
         Left            =   170
         TabIndex        =   12
         Top             =   860
         Width           =   615
      End
      Begin VB.CommandButton cmd3 
         Caption         =   "&3"
         Height          =   615
         Left            =   2040
         TabIndex        =   11
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "&2"
         Height          =   615
         Left            =   1440
         TabIndex        =   10
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&1"
         Height          =   615
         Left            =   760
         TabIndex        =   9
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmd0 
         Caption         =   "&0"
         Height          =   615
         Left            =   170
         TabIndex        =   8
         Top             =   200
         Width           =   615
      End
   End
   Begin VB.Frame frmHeight 
      Caption         =   "Elevation/Depth (map units)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   950
      Width           =   2775
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   2040
         TabIndex        =   22
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox txtelev 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   140
         TabIndex        =   0
         ToolTipText     =   "enter elevation or depth in the map's units"
         Top             =   280
         Width           =   1800
      End
   End
   Begin VB.Frame frmCoordinates 
      Caption         =   "Screen Coordinates"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   -10
      Width           =   2055
      Begin VB.TextBox txtY 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   7
         ToolTipText     =   "Map Y pixel coordinate at clicked point"
         Top             =   550
         Width           =   1335
      End
      Begin VB.TextBox txtX 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   6
         ToolTipText     =   "Map X coordinate at clicked point"
         Top             =   250
         Width           =   1335
      End
      Begin VB.Label lbYdig 
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblXdig 
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   255
      End
   End
End
Attribute VB_Name = "GDDigitizerfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub chkOcean_Click()
   If chkOcean.value = vbChecked Then
      InvElev = -1#
   Else
      InvElev = 1#
      End If
   
End Sub

Public Sub cmd8_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "8"
    Else
        txtelev.Text = txtelev.Text & "8"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd0_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "0"
    Else
        txtelev.Text = txtelev.Text & "0"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd1_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "1"
    Else
        txtelev.Text = txtelev.Text & "1"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd2_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "2"
    Else
        txtelev.Text = txtelev.Text & "2"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd3_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "3"
    Else
        txtelev.Text = txtelev.Text & "3"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd4_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "4"
    Else
        txtelev.Text = txtelev.Text & "4"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd5_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "5"
    Else
        txtelev.Text = txtelev.Text & "5"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd6_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "6"
    Else
        txtelev.Text = txtelev.Text & "6"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd7_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "7"
    Else
        txtelev.Text = txtelev.Text & "7"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmd9_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "9"
    Else
        txtelev.Text = txtelev.Text & "9"
        End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmdC_Click()
    If Trim$(txtelev.Text) = "0" Then
       txtelev.Text = "0"
       txtelev.SetFocus
    Else
       txtelev.Text = -txtelev.Text
       txtelev.SetFocus
       End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmdDigiClear_Click()

   txtelev.Text = sEmpty
   txtelev.SetFocus
   
   Call txtelev_GotFocus
   
End Sub


Private Sub cmdClear_Click()

   txtelev.Text = "0"
   txtelev.SetFocus
   
   Call txtelev_GotFocus
   
End Sub

Public Sub cmdD_Click()

    pos% = InStr(txtelev.Text, ".")
    If pos% = 0 Then 'can only be one decimal place
       txtelev.Text = txtelev.Text & "."
       End If
        
   Call txtelev_GotFocus
   
End Sub

Public Sub cmdEnter_Click()

   Call Form_KeyDown(vbKeyReturn, 0)
       
   txtelev.SetFocus
   
   Call txtelev_GotFocus
   
End Sub

Public Sub cmdSpace_Click()

   KeyDown (vbKeySpace)
   
   txtelev.SetFocus
   
   Call txtelev_GotFocus
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Dim Xcoord As Long
   Dim Ycoord As Long
             
   'if a mode is already activated, then can't use keypress event until that mode is deactivated
   If KeyCode = vbKeyReturn Then
      'pass it through
   ElseIf DigiEditPoints And (KeyCode = vbKeyR Or KeyCode = vbKeyK) Then
      'pass it through
   ElseIf _
      DigitizerSweep Or DigitizerEraser Or DigiEditPoints Or DigitizeHardy Or _
      DigitizePoint Or DigitizeBlankPoint Or DigiEditPoints Or DigitizeLine Or DigitizeContour Then
      Exit Sub 'need to deactivate one of these modes before activating new one.
      End If
   
   If KeyCode = vbKeyLeft Then 'left arrow
      h1 = -10
      GoSub ScrollHoriz
   ElseIf KeyCode = vbKeyUp Then 'up arrow
      h2 = -10
      GoSub ScrollVert
   ElseIf KeyCode = vbKeyRight Then 'right arrow
      h1 = 10
      GoSub ScrollHoriz
   ElseIf KeyCode = vbKeyDown Then 'down arrow
      h2 = 10
      GoSub ScrollVert
   ElseIf KeyCode = vbKeyPageUp Then  'PgUp
'      DigiPage = DigiPage + 1
      If GDform1.VScroll1.Visible Then GDform1.VScroll1.value = GDform1.VScroll1.Max
   ElseIf KeyCode = vbKeyPageDown Then  'PgDown
'      DigiPage = DigiPage - 1
      If GDform1.VScroll1.Visible Then GDform1.VScroll1.value = GDform1.VScroll1.min
   ElseIf KeyCode = vbKeyHome Then 'Home
      If GDform1.HScroll1.Visible Then GDform1.HScroll1.value = GDform1.HScroll1.min
   ElseIf KeyCode = vbKeyEnd Then 'End
      If GDform1.HScroll1.Visible Then GDform1.HScroll1.value = GDform1.HScroll1.Max
   
   ElseIf KeyCode = vbKeySpace Or _
          KeyCode = vbKeyB Or KeyCode = vbKeyC Or KeyCode = vbKeyD Or KeyCode = vbKeyE Or _
          KeyCode = vbKeyH Or KeyCode = vbKeyL Or KeyCode = vbKeyP Or KeyCode = vbKeyS Or _
          KeyCode = vbKeyR Or KeyCode = vbKeyK Then
          'space key or short-cut keys
          
      If KeyCode = vbKeyB Then
         DigiRightButtonIndex = 13
         
      ElseIf KeyCode = vbKeyC Then
         DigiRightButtonIndex = 5
          
      ElseIf KeyCode = vbKeyD Then 'some sort of deleting
         If DigiRightButtonIndex <> 3 And DigiRightButtonIndex <> 5 And DigiRightButtonIndex <> 7 Then 'delete last point
            DigiRightButtonIndex = 2
         ElseIf DigiRightButtonIndex <> 5 And DigiRightButtonIndex <> 7 Then 'delete last line
            DigiRightButtonIndex = 4
         ElseIf DigiRightButtonIndex <> 3 And DigiRightButtonIndex <> 7 Then 'eraser
            DigiRightButtonIndex = 6
         ElseIf DigiRightButtonIndex <> 3 And DigiRightButtonIndex <> 5 Then 'sweep erase
            DigiRightButtonIndex = 7
            End If
            
      ElseIf KeyCode = vbKeyE Then 'Edit mode'
         DigiRightButtonIndex = 14
      
      ElseIf KeyCode = vbKeyH Then 'Hardy
          DigiRightButtonIndex = 12
     
      ElseIf KeyCode = vbKeyL Then
          DigiRightButtonIndex = 3
     
      ElseIf KeyCode = vbKeyP Then 'Point digitizing
         If DigiRightButtonIndex <> 1 Then
            DigiRightButtonIndex = 0
         ElseIf DigiRightButtonIndex <> 2 Then
            DigiRightButtonIndex = 1
            End If
            
      ElseIf KeyCode = vbKeyR Then
          If DigiEditPoints Then
            'Edit Replace Mode
            'shift the point and replot
            If XpixLast <> -1 And YpixLast <> -1 Then
                  
               Xcoord = CLng(nearmouse_digi.x / (twipsx * DigiZoom.LastZoom))
               Ycoord = CLng(nearmouse_digi.Y / (twipsy * DigiZoom.LastZoom))
               
               ier = RedrawDigiPoints(Xcoord, Ycoord, DigiEditMode, 0)
               End If
             Exit Sub
             End If
      ElseIf KeyCode = vbKeyK Then
          If DigiEditPoints Then
            'Edit Kill Mode
            'delete the point and replot
            If XpixLast <> -1 And YpixLast <> -1 Then
                  
               Xcoord = CLng(nearmouse_digi.x / (twipsx * DigiZoom.LastZoom))
               Ycoord = CLng(nearmouse_digi.Y / (twipsy * DigiZoom.LastZoom))
               
               ier = RedrawDigiPoints(Xcoord, Ycoord, DigiEditMode, 1)
               End If
             Exit Sub
             End If
      ElseIf KeyCode = vbKeyS Then 'search
         DigiRightButtonIndex = 11
      
         End If
         
   
      'scroll options
        DigiRightButtonIndex = DigiRightButtonIndex + 1
        
        If DigiRightButtonIndex < 9 Or DigiRightButtonIndex > 11 Then
            DigiBackground = &HC0FFFF    'neutral
            GDDigitizerfrm.lblFunction.Enabled = True
            GDDigitizerfrm.lblFunction.BackColor = DigiBackground
        Else
            DigiBackground = &HE0E0E0       'disenabled
            GDDigitizerfrm.lblFunction.Enabled = False
            GDDigitizerfrm.lblFunction.BackColor = DigiBackground
            End If
        
        Select Case DigiRightButtonIndex
           Case 1
              GDDigitizerfrm.lblFunction.Caption = "Point (Blank)"
           Case 2
              GDDigitizerfrm.lblFunction.Caption = "Point (Repeat)"
           Case 3
              GDDigitizerfrm.lblFunction.Caption = "Delete Last Point"
           Case 4
              GDDigitizerfrm.lblFunction.Caption = "Lines"
           Case 5
              GDDigitizerfrm.lblFunction.Caption = "Delete Last Line"
           Case 6
              GDDigitizerfrm.lblFunction.Caption = "Contours"
           Case 7
              GDDigitizerfrm.lblFunction.Caption = "Erasing"
           Case 8
              GDDigitizerfrm.lblFunction.Caption = "Sweeping"
           Case 9
              GDDigitizerfrm.lblFunction.Caption = "Move in X"
           Case 10
              GDDigitizerfrm.lblFunction.Caption = "Move in Y"
           Case 11
              GDDigitizerfrm.lblFunction.Caption = "Zoom"
           Case 12
              GDDigitizerfrm.lblFunction.Caption = "Search"
           Case 13
              GDDigitizerfrm.lblFunction.Caption = "Hardy quadratic"
           Case 14
              GDDigitizerfrm.lblFunction.Caption = "Backup to dxf"
           Case 15
              GDDigitizerfrm.lblFunction.Caption = "Edit"
           Case 16
              DigiRightButtonIndex = 1
              'repeat case 1
              GDDigitizerfrm.lblFunction.Caption = "Point (Blank)"
           Case Else
        End Select
      
        DigiEntered = False
   
   ElseIf KeyCode = vbKeyReturn Then 'enter key
   
      'use to activate or deactivate function
      If Not DigiEntered Then
         'activate
         DigiEntered = True
         
           'activate an action
           If DigiRightButtonIndex < 9 Or DigiRightButtonIndex > 11 Then
              DigiBackground = &HC0FFC0
              GDDigitizerfrm.lblFunction.BackColor = DigiBackground
              End If
              
           Select Case DigiRightButtonIndex
              Case 1
                 Call GDMDIform.mnuDigitizePoint_Click
              Case 2
                 Call GDMDIform.mnuDigitizePointSameHeights_Click
              Case 3
                 Call GDMDIform.mnuDigiDeleteLastPoint_Click
              Case 4
                 Call GDMDIform.mnuDigitizeLine_Click
              Case 5
                 Call GDMDIform.mnuDigiDeleteLastLine_Click
              Case 6
                 Call GDMDIform.mnuDigitizeContour_Click
              Case 7
'                 If Not DigitizerEraser Then
                 Call GDMDIform.mnuEraser_Click
              Case 8
'                 If Not DigitizerSweep Then
                 Call GDMDIform.mnuDigiSweep_Click
              Case 9
'                 h1 = 10
'                 GoSub ScrollHoriz
              Case 10
'                 h2 = -10
'                 GoSub ScrollVert
              Case 11
'                 Call PictureBoxZoom(GDform1.Picture2, 0, 120, 0, 0, 0)
              Case 12
'                 If Not SearchDigi Then
                 Call GDMDIform.mnuSearchActivated_Click
              Case 13
'                 If Not DigiHardyPoints Then
                 Call GDMDIform.mnuDigitizeHardy_Click
              Case 14
                 Call GDMDIform.mnuSave_Click
              Case 15
'                 If Not DigiEditPoints Then
                 Call GDMDIform.EditDigitizedPoints
                 
           End Select
         
      
     Else
         'deactivate
         DigiEntered = False
         
           If DigiRightButtonIndex < 9 Or DigiRightButtonIndex > 11 Then
              DigiBackground = &HC0C0FF
              GDDigitizerfrm.lblFunction.BackColor = DigiBackground
              End If
           
           Select Case DigiRightButtonIndex
              Case 1
                 Call GDMDIform.mnuDigitizeEndPoint_Click
              Case 2
                 Call GDMDIform.mnuDigitizeEndPoint_Click
              Case 3
                 'continue deleting, and keep green background color to denote activation
                 DigiBackground = &HC0FFC0
                 GDDigitizerfrm.lblFunction.BackColor = DigiBackground
                 Call GDMDIform.mnuDigiDeleteLastPoint_Click
              Case 4
                 Call GDMDIform.mnuDigitizeEndLine_Click
              Case 5
                 'continue deleting, and keep green background color to denote activation
                 DigiBackground = &HC0FFC0
                 GDDigitizerfrm.lblFunction.BackColor = DigiBackground
                 Call GDMDIform.mnuDigiDeleteLastLine_Click
              Case 6
                 Call GDMDIform.mnuDigitizeEndContour_Click
              Case 7
                 Call GDMDIform.mnuEraser_Click
              Case 8
                 Call GDMDIform.mnuDigiSweep_Click
              Case 9
'                 h1 = -10
'                 GoSub ScrollHoriz
              Case 10
'                 h2 = 10
'                 GoSub ScrollVert
              Case 11
'                 Call PictureBoxZoom(GDform1.Picture2, 0, -120, 0, 0, 0)
              Case 12
                 Call GDMDIform.mnuSearchActivated_Click
              Case 13
                 Call GDMDIform.mnuDigitizeHardy_Click
              Case 14
'                 'do nothing
              Case 15
                  Call GDMDIform.EditDigitizedPoints
                  
           End Select
         
         End If
   
      End If
      
Exit Sub
      
ScrollHoriz:
    If GDform1.HScroll1.value + h1 < GDform1.HScroll1.min Or GDform1.HScroll1.value + h1 > GDform1.HScroll1.Max Then
          'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
          If GDform1.Picture2.Width > GDform1.HScroll1.Width Then

             If GDform1.HScroll1.value + h1 < GDform1.HScroll1.min Then
                GDform1.HScroll1.value = GDform1.HScroll1.min
             ElseIf GDform1.HScroll1.value + h1 > GDform1.HScroll1.Max Then
                GDform1.HScroll1.value = GDform1.HScroll1.Max
                End If
             End If
    Else
       GDform1.HScroll1.value = GDform1.HScroll1.value + h1
       End If
Return

ScrollVert:
    If GDform1.VScroll1.value + h2 < 0 Or GDform1.VScroll1.value + h2 > GDform1.VScroll1.Max Then
         'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
          If GDform1.Picture2.Height > GDform1.VScroll1.Height Then

             If GDform1.VScroll1.value + h2 < GDform1.VScroll1.min Then
                GDform1.VScroll1.value = GDform1.VScroll1.min
             ElseIf GDform1.VScroll1.value + h2 > GDform1.VScroll1.Max Then
                GDform1.VScroll1.value = GDform1.VScroll1.Max
                End If

             End If
    Else
       GDform1.VScroll1.value = GDform1.VScroll1.value + h2
       End If
                   
Return
 
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
     
    DigitizePadVis = True
    DigitizeOn = True
    
    InvElev = 1#
    
    txtX.Text = nearmouse_digi.x
    txtY.Text = nearmouse_digi.Y
   If Trim$(GDDigitizerfrm.txtelev) = gsEmpty Then
      GDDigitizerfrm.txtelev = Format(str$(ContourHeight / MapUnits), "#####0.0#")
      End If
    
    If DigiBackground = 0 Then
       DigiBackground = &HC0FFFF
       End If
       
    lblFunction.BackColor = DigiBackground
    
    If DigiRightButtonIndex = 0 Then DigiRightButtonIndex = 1
    
    DigiBackground = &HC0FFFF    'neutral
    GDDigitizerfrm.lblFunction.BackColor = DigiBackground
    
    Select Case DigiRightButtonIndex
       Case 1
          GDDigitizerfrm.lblFunction.Caption = "Point (Blank)"
       Case 2
          GDDigitizerfrm.lblFunction.Caption = "Point (Repeat)"
       Case 3
          GDDigitizerfrm.lblFunction.Caption = "Delete Last Point"
       Case 4
          GDDigitizerfrm.lblFunction.Caption = "Lines"
       Case 5
          GDDigitizerfrm.lblFunction.Caption = "Delete Last Line"
       Case 6
          GDDigitizerfrm.lblFunction.Caption = "Contours"
       Case 7
          GDDigitizerfrm.lblFunction.Caption = "Erasing"
       Case 8
          GDDigitizerfrm.lblFunction.Caption = "Sweeping"
       Case 9
          GDDigitizerfrm.lblFunction.Caption = "Move in X"
       Case 10
          GDDigitizerfrm.lblFunction.Caption = "Move in Y"
       Case 11
          GDDigitizerfrm.lblFunction.Caption = "Zoom"
       Case 12
          GDDigitizerfrm.lblFunction.Caption = "Search"
       Case 13
          GDDigitizerfrm.lblFunction.Caption = "Hardy quadratic"
       Case 14
          GDDigitizerfrm.lblFunction.Caption = "Backup to dxf"
       Case 15
          GDDigitizerfrm.lblFunction.Caption = "Edit"
       Case 16
          DigiRightButtonIndex = 1
          'repeat case 1
          GDDigitizerfrm.lblFunction.Caption = "Point (Blank)"
       Case Else
    End Select
    
    Call WheelHook(Me.hwnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim ier As Integer
   
   DigitizePoint = False
   DigitizeLine = False
   DigitizeBeginLine = False
   DigitizeEndLine = False
   DigitizeContour = False
   DigitizePadVis = False
   DigitizeBlankPoint = False
   DigitizePadVis = False
   ContourHeight = val(txtelev.Text) * MapUnits
   GDMDIform.SliderContour.Visible = False
   
   Set GDDigitizerfrm = Nothing
   Call WheelUnHook(Me.hwnd)
   
    buttonstate&(37) = 0
    GDMDIform.Toolbar1.Buttons(37).value = tbrUnpressed
    
    If Installation_Type = 1 Then
        If buttonstate&(47) = 1 Then
           buttonstate&(47) = 0
           GDMDIform.Toolbar1.Buttons(47).value = tbrUnpressed
           Unload TabConSample_VB_Form
           End If
        End If
       
     'disenable search drags
     If buttonstate&(15) = 1 Then
        buttonstate&(15) = 0
        GDMDIform.Toolbar1.Buttons(15).value = tbrUnpressed
        SearchDigi = False
        End If
        
     'disenable other types of drag window operations
     If DigitizeHardy Then
        DigitizeHardy = False
        buttonstate&(43) = 0
        GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
        
        XminC = 0
        YminC = 0
        XmaxC = 0
        YmaxC = 0
        
        End If
     
     If buttonstate&(40) = 1 Then
        buttonstate&(40) = 0
        GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
        DigitizerEraser = False
        GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
        End If
        
     If buttonstate&(41) = 1 Then
        buttonstate&(41) = 0
        DigitizerSweep = False
        GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
        End If
   
    GDMDIform.mnuEraser.Enabled = False
    GDMDIform.mnuDigiSweep.Enabled = False
    
'    If ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
'       (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) And _

    If (heights Or BasisDTMheights) And (RSMethod1 Or RSMethod2 Or RSMethod0) And DigiRubberSheeting Then
       GDMDIform.Toolbar1.Buttons(50).Enabled = True
       GDMDIform.Toolbar1.Buttons(51).Enabled = True
       End If
    
    'refresh map
    DigitizeOn = False
    ier = ReDrawMap(0)
    
    'renable blinking
    GDMDIform.CenterPointTimer.Enabled = True
    ce& = 1
    
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

'---------------------------------------------------------------------------------------
' Procedure : txtelev_Change
' Author    : Dr-John-K-Hall
' Date      : 4/12/2015
' Purpose   : make sure that only decimal input is entered in the txtelev textbox
'---------------------------------------------------------------------------------------
'
Private Sub txtelev_Change()
   On Error GoTo txtelev_Change_Error

   If Not IsNumeric(txtelev.Text) Then
      'remove non numberical characters
     If Mid$(Trim$(txtelev.Text), 1, 1) = "-" Then
        For i% = 2 To Len(txtelev.Text)
           ich$ = Mid$(txtelev.Text, i%, 1)
           If Not IsNumeric(ich$) And ich$ <> "." Then
              txtelev.Text = Mid$(txtelev.Text, 1, i% - 1) & Mid$(txtelev.Text, i% + 1, Len(txtelev.Text) - i%)
              End If
        Next i%
     Else
        For i% = 1 To Len(txtelev.Text)
           ich$ = Mid$(txtelev.Text, i%, 1)
           If Not IsNumeric(ich$) And ich$ <> "." Then
              txtelev.Text = Mid$(txtelev.Text, 1, i% - 1) & Mid$(txtelev.Text, i% + 1, Len(txtelev.Text) - i%)
              End If
        Next i%
        End If
      End If

   On Error GoTo 0
   Exit Sub

txtelev_Change_Error:

    
End Sub

Private Sub txtelev_Click()
   Call txtelev_GotFocus
End Sub

Public Sub txtelev_GotFocus()
   If Len(txtelev.Text) > 0 Then
      txtelev.SelStart = 0
      txtelev.SelLength = Len(txtelev.Text)
      End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowModes
' Author    : Dr-John-K-Hall
' Date      : 5/27/2015
' Purpose   : Shows the correct mode in the GDDigitizerfrm when pushing and reseting the toolbar buttons
'           mode% = 0 for green background
'                 = 1 for red background
'---------------------------------------------------------------------------------------
'
Public Sub ShowModes(activatemode%, mode%)

   Dim DigiBackground As Long
   
   DigiRightButtonIndex = activatemode%

   On Error GoTo ShowModes_Error

        Select Case DigiRightButtonIndex
           Case 1
              GDDigitizerfrm.lblFunction.Caption = "Point (Blank)"
           Case 2
              GDDigitizerfrm.lblFunction.Caption = "Point (Repeat)"
           Case 3
              GDDigitizerfrm.lblFunction.Caption = "Delete Last Point"
           Case 4
              GDDigitizerfrm.lblFunction.Caption = "Lines"
           Case 5
              GDDigitizerfrm.lblFunction.Caption = "Delete Last Line"
           Case 6
              GDDigitizerfrm.lblFunction.Caption = "Contours"
           Case 7
              GDDigitizerfrm.lblFunction.Caption = "Erasing"
           Case 8
              GDDigitizerfrm.lblFunction.Caption = "Sweeping"
           Case 9
              GDDigitizerfrm.lblFunction.Caption = "Move in X"
           Case 10
              GDDigitizerfrm.lblFunction.Caption = "Move in Y"
           Case 11
              GDDigitizerfrm.lblFunction.Caption = "Zoom"
           Case 12
              GDDigitizerfrm.lblFunction.Caption = "Search"
           Case 13
              GDDigitizerfrm.lblFunction.Caption = "Hardy quadratic"
           Case 14
              GDDigitizerfrm.lblFunction.Caption = "Backup to dxf"
           Case 15
              GDDigitizerfrm.lblFunction.Caption = "Edit"
           Case 16
              DigiRightButtonIndex = 1
              'repeat case 1
              GDDigitizerfrm.lblFunction.Caption = "Point (Blank)"
           Case Else
        End Select
        
   If mode% = 0 Then
      DigiBackground = &HC0C0FF 'green
      GDDigitizerfrm.lblFunction.BackColor = DigiBackground
   ElseIf mode% = 1 Then
      DigiBackground = &HC0FFC0 'red
      GDDigitizerfrm.lblFunction.BackColor = DigiBackground
      End If

   On Error GoTo 0
   Exit Sub

ShowModes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowModes of Form GDDigitizerfrm"
End Sub
