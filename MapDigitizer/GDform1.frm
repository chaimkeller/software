VERSION 5.00
Begin VB.Form GDform1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   -75
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   2595
      Left            =   4560
      TabIndex        =   2
      Top             =   180
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   2880
      Width           =   4035
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   2595
      Left            =   540
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   0
      Top             =   240
      Width           =   4035
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   1935
         Left            =   540
         MousePointer    =   2  'Cross
         ScaleHeight     =   1875
         ScaleWidth      =   2835
         TabIndex        =   3
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "GDform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   On Error GoTo errhand
   
   color_init.R = INIT_VALUE
   color_init.v = INIT_VALUE
   color_init.b = INIT_VALUE
   
   If Dir(picnam$) = sEmpty Then Exit Sub
   
   twipsx = Screen.TwipsPerPixelX
   twipsy = Screen.TwipsPerPixelY

   sizex = twipsx * pixwi
   sizey = twipsy * pixhi
   
   GDform1.Left = 0
   GDform1.Top = 0
   GDform1.Width = GDMDIform.Width - 210
   GDform1.Height = GDMDIform.Height - 1480 - GDMDIform.StatusBar1.Height
   
   'Set ScaleMode to pixels (since picture size is in pixels)
   GDform1.ScaleMode = vbPixels
   GDform1.Picture1.ScaleMode = vbPixels
   
   'Autosize is set to True so that the boundaries of
   'Picture2 are expanded to the size of the actual map
   GDform1.Picture2.AutoSize = True
   
   'Set the BorderStyle of each picture box to None.
   'GDform1.Picture1.BorderStyle = 0
   GDform1.Picture2.BorderStyle = 0
   
   'Load the default map
   GDform1.Picture2.Picture = LoadPicture(picnam$)
   'GDform1.Picture2.PaintPicture GDMDIform.PictureClip1.Picture, 0, 0, pixwi , pixhi , 0, 0, pixwi, pixhi
   
   
   'Initialize location of both pictures
   GDform1.Picture1.Move 0, 0, ScaleWidth - VScroll1.Width, ScaleHeight - HScroll1.Height
   GDform1.Picture2.Move 0, 0
   
   'Position the horizontal scroll bar
   HScroll1.Top = GDform1.Picture1.Height
   HScroll1.Left = GDform1.Picture1.Left
   HScroll1.Width = GDform1.Picture1.Width
   
   'Position the vertical scroll bar
   VScroll1.Top = 0
   VScroll1.Left = GDform1.Picture1.Width
   VScroll1.Height = GDform1.Picture1.Height
   
   'Set the Max property for the scroll bars.
   HScroll1.Max = GDform1.Picture2.Width - GDform1.Picture1.Width
   VScroll1.Max = GDform1.Picture2.Height - GDform1.Picture1.Height
   
   'Determine if the child picture will fill up the screen
   'If so, there is no need to use scroll bars.
   GDform1.VScroll1.Visible = (GDform1.Picture1.Height < GDform1.Picture2.Height)
   GDform1.HScroll1.Visible = (GDform1.Picture1.Width < GDform1.Picture2.Width)
   
   'Initiate Scroll Step Sizes
   HScroll1.LargeChange = HScroll1.Max / 20
   HScroll1.SmallChange = HScroll1.Max / 60
      
   VScroll1.LargeChange = VScroll1.Max / 20
   VScroll1.SmallChange = VScroll1.Max / 60
   
   Screen.MousePointer = vbDefault
   Exit Sub

errhand:
   Select Case Err.Number
      Case 481 'invalid picture
         Screen.MousePointer = vbDefault
         MsgBox "Invalid picture!", vbCritical + vbOKOnly, "GSI_PDB"
      Case 380 'scroll bar errors, just resume
         Resume Next
      Case Else 'unexpected error
         Screen.MousePointer = vbDefault
         MsgBox "Encountered error #: " & Err.Number & vbLf & _
           "in program module GDform1:Form_Load." & vbLf & _
           Err.Description, vbCritical + vbOKOnly, "GSI_PDB"
  End Select
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'buttonstate&(1) = 0
   buttonstate&(23) = 0
   GDMDIform.Toolbar1.Buttons(1).value = tbrUnpressed
   GDMDIform.Toolbar1.Buttons(23).value = tbrUnpressed
   Unload Me
End Sub



Private Sub Form_Resize()
   On Error GoTo errhand
   GDform1.Picture1.Move 0, 0, ScaleWidth - VScroll1.Width, ScaleHeight - HScroll1.Height
   GDform1.Picture2.Move 0, 0
   If magvis = True Then
      GDMagform.Left = GDform1.Left + GDform1.Width
      GDMagform.Width = GDMDIform.Width - GDform1.Width - 200
      End If
      
   If DigitizeMagvis Then
      GDDigiMagfrm.Left = GDform1.Left + GDform1.Width
      GDDigiMagfrm.Width = GDMDIform.Width - GDform1.Width - 200
      End If
      
   'Position the horizontal scroll bar
   HScroll1.Top = GDform1.Picture1.Height
   HScroll1.Left = GDform1.Picture1.Left
   HScroll1.Width = GDform1.Picture1.Width
   
   'Position the vertical scroll bar
   VScroll1.Top = 0
   VScroll1.Left = GDform1.Picture1.Width
   VScroll1.Height = GDform1.Picture1.Height
   
   'Set the Max property for the scroll bars.
   HScroll1.Max = GDform1.Picture2.Width - GDform1.Picture1.Width
   VScroll1.Max = GDform1.Picture2.Height - GDform1.Picture1.Height
   
   'Determine if the child picture will fill up the screen
   'If so, there is no need to use scroll bars.
   GDform1.VScroll1.Visible = (GDform1.Picture1.Height < GDform1.Picture2.Height)
   GDform1.HScroll1.Visible = (GDform1.Picture1.Width < GDform1.Picture2.Width)
   
   GDform1.Picture2.Left = -HScroll1.value
   GDform1.Picture2.Top = -VScroll1.value
      
   Exit Sub
errhand:
End Sub

Private Sub HScroll1_Change()
'   picx = 0
'   picy = 0
'   width1 = Picture1.Width
'   height1 = Picture1.Height
'   picx2 = (HScroll1.Value / HScroll1.Width) * pixwi 'sizex
'   picy2 = (VScroll1.Value / VScroll1.Height) * pixhi 'sizey
'   Picture1.PaintPicture GDMDIform.PictureClip1.Picture, picx, picy, width1, height1, picx2, picy2, width1, height1
   GDform1.Picture2.Left = -HScroll1.value
End Sub

Private Sub HScroll1_GotFocus()
    'print prompts on statusbar
    If SearchDB Then
      If NumReportPnts& = 0 Then
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
      Else
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
         End If
    Else
      GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define magnification boundaries."
     End If
End Sub

Private Sub Picture2_GotFocus()
    'print prompts on statusbar
    If Not DigitizeOn Then
        If SearchDB Then
          If NumReportPnts& = 0 Then
             GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
          Else
             GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
             End If
        Else
          GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define magnification boundaries."
         End If
    ElseIf DigitizeOn Then
       'enable magnification of curson position
'       GDMDIform.DigiTimer.Enabled = True
       GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location to magnify digitization region."
       End If
End Sub

Private Sub Picture2_LostFocus()

'    If DigitizeOn Then
'       GDMDIform.DigiTimer.Enabled = False
'       End If
       
    GDMDIform.StatusBar1.Panels(1) = sEmpty
   
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo errhand
   
   If Button = 1 And drawbox = False And Not DigitizeOn Then 'may be beginning of drag operation
      drag1x = x
      drag1y = y
      dragbegin = True
      drag2x = drag1x
      drag2y = drag1y
      End If
      
   If Not DigitizeOn Then
      'shut off timers during drag
      ce& = 0 'reset blinker flag
      If GDMDIform.CenterPointTimer.Enabled = True Then
         ce& = 1 'flag that timer has been shut down during drag
         GDMDIform.CenterPointTimer.Enabled = False
         End If
      End If
      
   Exit Sub
   
errhand:
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          "in module: Gdform1.Picture2_MouseDown", _
          vbCritical + vbOKOnly, "GSI_PDB"
      
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
'  On Error GoTo errhand
  
  Dim ContourHeight As Single

  nearmouse_digi.x = x
  nearmouse_digi.y = y
  
'  Dim TestColor As Long
''  Dim ColorTest As String
'  Dim Color_Test As couleur
'
'  TestColor = GDform1.Picture2.Point(nearmouse_digi.X, nearmouse_digi.y)
'''  GDMDIform.StatusBar1.Panels(2).Text = TestColor
'  Color_Test = recupcouleur(TestColor)
'
'' find difference of this from black
'  If color_init.R = INIT_VALUE Then
'     color_init.R = Color_Test.R
'     color_init.v = Color_Test.v
'     color_init.b = Color_Test.b
'  Else
'     dif = Sqr((Color_Test.R - color_init.R) ^ 2 + (Color_Test.v - color_init.v) ^ 2 + (Color_Test.b - color_init.b) ^ 2)
'     GDMDIform.StatusBar1.Panels(2).Text = Format(dif, "######.#0")
'     End If
     
'  ColorTest = getPixelColor(Color_Test.r, Color_Test.v, Color_Test.b)
'  GDMDIform.StatusBar1.Panels(2).Text = ColorTest

'
'  Dim news_mouse As PointAPI
'  Dim hDnews As Long
'  Dim Rnews As Long
'  Dim News_Color As couleur
'
'  GetCursorPos news_mouse
'  hDnews = GetDC(0)
'  Rnews = GetPixel(hDnews, news_mouse.X, news_mouse.Y)
'  If Rnews <> -1 Then
'     News_Color = recupcouleur(Rnews)
'     GDMDIform.StatusBar1.Panels(2).Text = News_Color.R & "," & News_Color.v & "," & News_Color.b
'     End If
'
'  For ii = 1 To 500 Step 50
'
'      Rnews = GDform1.Picture2.Point(news_mouse.X + ii, news_mouse.Y)
'      If Rnews <> -1 Then
'         News_Color = recupcouleur(Rnews)
'         GDMDIform.StatusBar1.Panels(2).Text = News_Color.R & "," & News_Color.v & "," & News_Color.b
'         End If
'
'  Next ii
     
  
'  Rnews = GDform1.Picture2.Point(news_mouse.X, news_mouse.Y)
'  If Rnews <> -1 Then
'     News_Color = recupcouleur(Rnews)
'     GDMDIform.StatusBar1.Panels(2).Text = News_Color.R & "," & News_Color.v & "," & News_Color.b
'     End If

     
  If GeoMap = True Or TopoMap = True Then 'if the map is visible
      xcoord = x
      ycoord = y
      
      If Geo = True Then
         'keep geo coordinate window visible
         ret = BringWindowToTop(GDGeoFrm.hWnd)
         End If
      
      Select Case Button
         Case 1  'left button
            'shift this point to middle of screen
            'this will be the case when (X,Y) = (picture1.width/2, picture1.height/2)
            
gd50:       If (drag1x = drag2x And drag1y = drag2y) And Not DigitizeOn Then
                dragbegin = False
                dragbox = False
                
                'reset center timer if flagged
                If ce& = 1 Then 'blinker was shut down during drag, so reenable it
                   ce& = 0 'reset blinker flag
                   GDMDIform.CenterPointTimer.Enabled = True
                   End If
            Else 'signales end of drag
               End If
               
               
            'erase last box and redraw it at new position
            If Not DigitizeOn And dragbox = True And dragbegin = True And ((drag2x = drag1x And drag2y <> drag1y) Or (drag2x <> drag1x And drag2y = drag1y)) And Button = 1 Then
               'defines box with no internal area, erase it and start again
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (x, y)-(drag1x, drag1y), QBColor(15), B
               Picture2.DrawMode = 13
               drawbox = False
               dragbegin = False
               'reset center timer if flagged
               If ce& = 1 Then 'blinker was shut down during drag operation
                  ce& = 0 'reset blinker flag
                  GDMDIform.CenterPointTimer.Enabled = True
                  End If
            ElseIf Not DigitizeOn And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               Picture2.DrawWidth = 2
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               Picture2.DrawMode = 13
               dragbegin = False
               dragbox = False
               drawbox = True
               Picture2.DrawWidth = 1
               magclose = False
               If drag2x < drag1x Then
                  dragtmp = drag1x
                  drag1x = drag2x
                  drag2x = dragtmp
                  End If
               If drag2y < drag1y Then
                  dragtmp = drag1y
                  drag1y = drag2y
                  drag2y = dragtmp
                  End If
                  
                'refresh toolbar1
                For i& = 1 To GDMDIform.Toolbar1.Buttons.count
                    If buttonstate&(i&) = 1 Then
                       GDMDIform.Toolbar1.Buttons(i&).value = tbrPressed
                       End If
                Next i&
                  
gd150:          If SearchDB = False Then 'show magnif if search is off
                   maginit = True
                   GDMagform.Visible = True
                   End If
                
                'enter box coordinates into search form if
                'searches are flaged
                If SearchDB Then
                   
                   'convert drag coordinates to ITM
                   Call ConvertPixToCoord(drag1x, drag1y, Xout1, Yout1)
                   Call ConvertPixToCoord(drag2x, drag2y, Xout2, Yout2)
                   
                   'erase drag box
                   Picture2.DrawMode = 7
                   Picture2.DrawWidth = 2
                   Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                   Picture2.DrawWidth = 1
                   Picture2.DrawMode = 13
                   drawbox = False
                   dragbegin = False
                   'reset drag coordinates
                   drag1x = 0
                   drag2x = 0
                   drag1y = 0
                   drag2y = 0
                    
                   If PicSum Then
                      'erase old plotted search points if any
                      If NumReportPnts& <> 0 Then
                         EraseOldSearchPoints
                         NumReportPnts& = 0 'clear buffer
                         ReDim ReportPnts(1, 0) 'clear memory
                         End If
                      End If
                      
                   'fill coordinate boundaries values in Search Wizard
                   stepsearch& = 0
                   GDSearchfrm.cmdBack.Enabled = False
                   GDSearchfrm.Visible = True
                   ret = ShowWindow(GDSearchfrm.hWnd, SW_NORMAL)
                   GDSearchfrm.tbSearch.Tab = 0
                   GDSearchfrm.txtEastMin = Fix(Xout1)
                   GDSearchfrm.txtNorthMax = Fix(Yout1)
                   GDSearchfrm.txtEastMax = Fix(Xout2)
                   GDSearchfrm.txtNorthMin = Fix(Yout2)
                   GDSearchfrm.tbSearch.Tab = 0
                   ret = BringWindowToTop(GDSearchfrm.hWnd)
                   End If
                   
                If SearchDB = False Then
                   'sit here until user closes mag box
                   'shut off timers
                   Do Until magclose = True
                     DoEvents
                   Loop
                   End If
                
               'mag box closed, so erase box
               If drawbox Then
                  Picture2.DrawMode = 7
                  Picture2.DrawWidth = 2
                  Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                  Picture2.DrawWidth = 1
                  Picture2.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  'reset drag coordinates
                  drag1x = 0
                  drag2x = 0
                  drag1y = 0
                  drag2y = 0
                  End If
               
               'reset center timer if flagged
               If ce& = 1 Then 'timer was shut down, so reset it
                   GDMDIform.CenterPointTimer.Enabled = True
                   ce& = 0 'reset blinker flag
                   End If
               If shiftmag = True Then
                  'right click recorded inside mag window
                  'so close mag window and goto to desired
                  'coordinate on unmagnified map window
                  ce& = 1 'flag to draw blinker at new location and erase old location
                  Call gotocoord
                  shiftmag = False
                  End If
                  
            Else 'shift map and draw circle at click point
            
                ce& = 1 'flag to draw blinker at new position
                Call ShiftMap(x, y)
                
                'put click coordinates into coordinate boxes
                'Convert coordinates to pixels
                xcoord = x / twipsx
                ycoord = y / twipsy
                
                'Convert pixel coordinates to ITM (or to any user's coordinate system)
                ITMx = ((X2 - x1) / pixwi) * xcoord + x1
                ITMy = y1 - ((y1 - Y2) / pixhi) * ycoord
                
                'Display the ITM coordinates of the click point
                GDMDIform.Text5 = str(Int(ITMx))
                GDMDIform.Text6 = str(Int(ITMy))
                
                'Display height at click point
                If heights = True And lblX = "ITMx" And LblY = "ITMy" Then 'display heights
                   kmx = ITMx
                   kmy = ITMy
                   'Call DTMheight(kmx, kmy, hgt)
                   Dim hgt As Integer
                   Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
                   GDMDIform.Text7 = str(hgt)
                   End If
                
                'Write position file to the hard disk
                'These coordinates will be used as the
                'starting position for the next time the
                'user logs into the program.  It can be also
                'used by the Access database program for
                'automatically inputing coordinates (to reduce
                'human error while inputing coordinates)
                Call UpdatePositionFile(ITMx, ITMy, hgt)
             
                'print prompts on statusbar
                If SearchDB Then
                  If NumReportPnts& = 0 Then
                     GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
                  Else
                     GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
                     End If
                Else
                  GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define magnification boundaries."
                 End If
                 
               With GDDigitizerfrm
               
                   If DigitizePoint Then
                   
                        If Not DigitizePadVis Then
                           GDDigitizerfrm.Visible = True
                           End If

                        BringWindowToTop (GDDigitizerfrm.hWnd)
                        
                        If IsNumeric(val(.txtelev.Text)) And Trim(.txtelev.Text) <> vbNullString Then
                          'draw x at the digitized point with the elevation
                          gddm = GDform1.Picture2.DrawMode
                          gddw = GDform1.Picture2.DrawWidth
                          GDform1.Picture2.DrawMode = 13
                          GDform1.Picture2.DrawWidth = 2
                          GDform1.Picture2.Line (digi_last.x - 30, digi_last.y - 30)-(digi_last.x + 30, digi_last.y + 30), GeoOutColor&
                          GDform1.Picture2.Line (digi_last.x - 30, digi_last.y + 30)-(digi_last.x + 30, digi_last.y - 30), GeoOutColor&
                          GDform1.Picture2.DrawMode = gddm
                          GDform1.Picture2.DrawWidth = gddw
                          digi_last.x = x
                          digi_last.y = y
                          digi_last.Z = val(.txtelev.Text)
                          
                          'record point
                          If numDigiPoints = 0 Then
                             ReDim DigiPoints(2, 0)
                             DigiPoints(0, 0) = digi_last.x
                             DigiPoints(1, 0) = digi_last.y
                             DigiPoints(2, 0) = digi_last.Z
                             numDigiPoints = numDigiPoints + 1
                          Else
                             ReDim Preserve DigiPoints(2, numDigiPoints)
                             DigiPoints(0, numDigiPoints) = digi_last.x
                             DigiPoints(1, numDigiPoints) = digi_last.y
                             DigiPoints(2, numDigiPoints) = digi_last.Z
                             numDigiPoints = numDigiPoints + 1
                             End If
                             
                        Else 'store coordinate value
                          digi_last.x = x
                          digi_last.y = y
                          digi_last.Z = val(.txtelev.Text)
                          End If
                          
                        .txtX = x
                        .txtY = y
                         
                        If DigitizeBlankPoint Then .txtelev = vbNullString
                        
                   ElseIf DigitizeLine Then
                      
                        If Not DigitizePadVis Then
                           GDDigitizerfrm.Visible = True
                           End If
                        
                        BringWindowToTop (GDDigitizerfrm.hWnd)
                        
                        If IsNumeric(val(.txtelev.Text)) And Trim(.txtelev.Text) <> vbNullString Then
                            digi_last.x = x
                            digi_last.y = y
                            digi_last.Z = val(.txtelev.Text)
                           End If
    
                        .txtX = x
                        .txtY = y
                            
                        If DigitizeBeginLine And IsNumeric(val(.txtelev.Text)) And Trim(.txtelev.Text) <> vbNullString Then
                        
                            DigitizeBeginLine = False
                            'initialize beginning coordinates of line
                            digi_begin.x = x
                            digi_begin.y = y
                            digi_begin.Z = val(.txtelev.Text)
                            digi_last.x = -999999
                            digi_last.y = -999999
                            
                            
                         ElseIf Not DigitizeBeginLine And Not DigitizeEndLine And IsNumeric(val(.txtelev.Text)) And Trim(.txtelev.Text) <> vbNullString Then
                            'draw last line
                            
                            'here is where to record subsequent vertices

                            gddm = GDform1.Picture2.DrawMode
                            gddw = GDform1.Picture2.DrawWidth

                            If digi_last.x <> -999999 And digi_last.y <> -999999 And digi_begin.x <> -999999 And digi_begin.y <> -999999 Then
                                GDform1.Picture2.DrawMode = 13
                                GDform1.Picture2.DrawWidth = 2
                                GDform1.Picture2.Line (digi_last.x, digi_last.y)-(digi_begin.x, digi_begin.y), QBColor(12)
                                End If

                            GDform1.Picture2.DrawMode = gddm
                            GDform1.Picture2.DrawWidth = gddw
                            
                            If numDigiLines = 0 Then
                               ReDim DigiLines(5, 0)
                               DigiLines(0, 0) = digi_begin.x
                               DigiLines(1, 0) = digi_begin.y
                               DigiLines(2, 0) = digi_begin.Z
                               DigiLines(3, 0) = digi_last.x
                               DigiLines(4, 0) = digi_last.y
                               DigiLines(5, 0) = digi_last.Z
                               numDigiLines = numDigiLines + 1
                            Else
                               ReDim Preserve DigiLines(5, numDigiLines)
                               DigiLines(0, numDigiLines) = digi_begin.x
                               DigiLines(1, numDigiLines) = digi_begin.y
                               DigiLines(2, numDigiLines) = digi_begin.Z
                               DigiLines(3, numDigiLines) = digi_last.x
                               DigiLines(4, numDigiLines) = digi_last.y
                               DigiLines(5, numDigiLines) = digi_last.Z
                               numDigiLines = numDigiLines + 1
                               End If
                            
                            End If
                            
                            
                      ElseIf DigitizeContour And DigitizeMagvis And Not PointStart Then
                      
                         If Not DigitizePadVis Then
                            GDDigitizerfrm.Visible = True
                            End If
                                       
                         BringWindowToTop (GDDigitizerfrm.hWnd)
                         
                         GDDigitizerfrm.txtX = blink_mark.x
                         GDDigitizerfrm.txtY = blink_mark.y
'                         GDDigitizerfrm.txtelev = ContourHeight
'                         If DigitizeContinueContour Then
'                            DigitizeContinueContour = False
'                            .txtelev = ContourHeight
'                            End If
                      
                         If IsNumeric(val(.txtelev.Text)) And Trim(.txtelev.Text) <> vbNullString Then
                         
                             ContourHeight = val(.txtelev.Text)
                             Unload GDDigitizerfrm
                             
                             Start_Point.x = nearmouse_digi.x
                             Start_Point.y = nearmouse_digi.y
                             
                             If DigitizeContinueContour Then
                                DigitizeContinueContour = False
                                
                                'record coordinates
                                Next_Point.x = Start_Point.x
                                Next_Point.y = Start_Point.y
                                
                                numTrace = numTrace + 1
                                ReDim Preserve trace(numTrace)
                                trace(numTrace).x = Next_Point.x
                                trace(numTrace).y = Next_Point.y
                                trace(numTrace).Z = ContourHeight
                                
                                TraceColor = QBColor(12)
                                
                                'draw connecting line
                                gdw = GDform1.Picture2.DrawWidth
                                GDform1.Picture2.DrawWidth = 1
                                GDform1.Picture2.PSet (trace(numTrace - 1).x, trace(numTrace - 1).y), TraceColor
                                GDform1.Picture2.Line -(trace(numTrace).x, trace(numTrace).y), TraceColor
                                GDform1.Picture2.DrawWidth = gdw
                                
                                End If
                             
                             'convert to screen coordinates
                             Dim R As Long
    '                         Dim hD As Long
    '                         Dim p As PointAPI
    '                         Dim rctCurrentView As RECT
    '                         Dim rctCurrentMain As RECT
    '                         Dim rctCurrentProg As RECT
    
    '                         'find screen coordinates of this X,Y
    '                         p.X = Start_Point.X
    '                         p.Y = Start_Point.Y
    '                         ClientToScreen GDDigiMagfrm.PictureBox1.hWnd, p
                             
    '                         Call GetWindowRect(GDform1.Picture1.hWnd, rctCurrentView)
    '                         Call GetWindowRect(GDform1.hWnd, rctCurrentMain)
    '                         Call GetWindowRect(GDMDIform.hWnd, rctCurrentProg)
    '                         difx0 = rctCurrentMain.x1 - rctCurrentProg.x1
    '                         dify0 = rctCurrentMain.y1 - rctCurrentProg.y1
    '                         difx = rctCurrentView.x1 - rctCurrentMain.x1
    '                         dify = rctCurrentView.y1 - rctCurrentMain.y1
                             
                             
'                             GDMDIform.StatusBar1.Panels(1).Text = "X: " & p.X & " Y: " & p.Y & " shifts x,y: " & test_point.X & ", " & test_point.Y & " middle: " & CInt(GDform1.Picture1.Width * 0.5 + difx + difx0 - 2) & ", " & CInt(GDform1.Picture1.Height * 0.5 + Abs(dify) + Abs(dify0) - 2)
    '                         Start_Point.X = CInt(GDform1.Picture1.Width * 0.5 + difx + difx0 - 2)
    '                         Start_Point.Y = CInt(GDform1.Picture1.Height * 0.5 + Abs(dify) + Abs(dify0) - 2)
    '
'                             hD = GetDC(0) 'get color under cursor "www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_20935677.html
'                             R = GetPixel(hD, mouse_digi.X, mouse_digi.Y)
'                             R = GetPixel(hD, 10, 10)
'                             R = GetPixel(hD, nearmouse_digi.X, nearmouse_digi.Y)
    
                             
'                             R = GetPixel(GDform1.Picture2.hDc, mouse_digi.X, mouse_digi.Y)
'                             hD = GetDC(0)
'                             GetCursorPos Start_Point
'                             R = GetPixel(hD, Start_Point.X, Start_Point.Y)
'
'                             If R <> -1 Then
'                               Start_Color = recupcouleur(R)
'                               GDMDIform.StatusBar1.Panels(1).Text = "RGB of starting point = " & Start_Color.R & "," & Start_Color.v & "," & Start_Color.b
'                               End If
                               
'                               R = GetPixel(hD, mouse_dig1.X + 1, mouse_digi.Y + 1)
'                               Start_Color = recupcouleur(R)
    '
    ''                         Call ReleaseDC(0, hD) 'release the dc
    
                              Start_Point.x = nearmouse_digi.x
                              Start_Point.y = nearmouse_digi.y
                              
                              R = GDform1.Picture2.Point(Start_Point.x, Start_Point.y)
                              If R <> -1 Then
                                 Start_Color = recupcouleur(R)
                                 GDMDIform.StatusBar1.Panels(2).Text = Start_Color.R & "," & Start_Color.v & "," & Start_Color.b
                              Else
                                 MsgBox "Contour color couldn't be determined, try again"
                                 Exit Sub
                                 End If
                              
    
'                              GDDigiMagfrm.PictureBlt.PaintPicture GDform1.Picture2.Picture, 0, 0, Screen.TwipsPerPixelX, Screen.TwipsPerPixelY, Start_Point.X, Start_Point.Y, Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
'                              R = GetPixel(GDDigiMagfrm.PictureBlt.hDc, 0, 0)
'
'                              If R <> -1 Then
'                                 Start_Color = recupcouleur(R)
'                              Else
'                                 MsgBox "Contour color couldn't be determined, try again"
'                                 Exit Sub
'                                 End If

                              'shut down blinkers
                              GDMDIform.CenterPointTimer.Enabled = False
                        
                              If CenterBlinkState And ce& = 1 Then
                                 Call DrawPlotMark(0, 0, 1)
                                 End If
                             
                             PointStart = True
                          
                             Dim Scroll As Long
                             Dim trait As Boolean
                             
                             Scroll = GDMDIform.SliderContour.value 'replace with slider
                             trait = False
'                             Call tracecontours3(GDform1.Picture2, GDform1.Picture2, ContourHeight, Scroll, trait)
'                             Call tracecontours4(GDform1.Picture2, GDform1.Picture2, ContourHeight, Scroll, trait)
'                             Call tracecontours5(GDform1.Picture2, GDform1.Picture2, ContourHeight, Scroll, trait)
'                             Call tracecontours6(GDform1.Picture2, GDform1.Picture2, ContourHeight, Scroll, trait)
                             Call tracecontours7(GDform1.Picture2, GDform1.Picture2, ContourHeight, Scroll, trait)
                             
                             End If
                      
                      End If
                      
               End With
                 
               
            End If
            
         Case 2 'right button--show record information of
                'search result nearest to right click point
                
             If PicSum = True And NumReportPnts& <> 0 Then
                
               'print prompts on statusbar
                GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
                
                Screen.MousePointer = vbHourglass
                'freeze the GeoMap blinker
                If GeoMap And ce& = 1 Then
                   GDMDIform.CenterPointTimer.Enabled = False
                   End If
                
                'shift map to this point
                Call ShiftMap(x, y)
                               
                'make list reappear at the record that has the
                'closest coordinate to the clicked point
                xcoord = x / twipsx
                ycoord = y / twipsy
                
                'Convert pixel coordinates to ITM (or to any coord system)
                ITMx = ((X2 - x1) / pixwi) * xcoord + x1
                ITMy = y1 - ((y1 - Y2) / pixhi) * ycoord
                
                'Display the ITM coordinates of the click point
                GDMDIform.Text5 = str(Int(ITMx))
                GDMDIform.Text6 = str(Int(ITMy))
                
                'Display height at click point
                If heights = True And lblX = "ITMx" And LblY = "ITMy" Then 'display heights
                   kmx = ITMx
                   kmy = ITMy
                   'Call DTMheight(kmx, kmy, hgt)
                   Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
                   GDMDIform.Text7 = str(hgt)
                   End If
                
                'write position file to hard disk
                Call UpdatePositionFile(ITMx, ITMy, hgt)
                
                'now search for closest highlighted record to this point
                NearestPnt& = -1
                SelectedPnt& = 0
                DetailRecordNum& = 0
                For i& = 1 To numReport&
                    If GDReportfrm.lvwReport.ListItems(i&).Selected Then
                       SelectedPnt& = SelectedPnt& + 1
                       XPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(1))
                       YPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(2))
                       dist = Sqr((ITMx - XPnt) ^ 2 + (ITMy - YPnt) ^ 2)
                       If NearestPnt& = -1 Then
                          NearestPnt& = i&
                          NearestDist = dist
                       Else
                          If dist < NearestDist Then
                             NearestPnt& = i&
                             NearestDist = dist
                          End If
                       End If
                    End If
                Next i&
                
                If NearestPnt& > 0 Then 'found the nearest search result to this point
                    'show detailed report of this search result record
                    ShowDetailedReport
                    
                    GDDetailReportfrm.Visible = True
                    ret = BringWindowToTop(GDDetailReportfrm.hWnd)
                    End If
                
                Screen.MousePointer = vbDefault
                
                'refresh the timers
                If GeoMap And ce& = 1 Then
                   GDMDIform.CenterPointTimer.Enabled = True
                   End If
                   
                'refresh toolbar1
                For i& = 1 To GDMDIform.Toolbar1.Buttons.count
                    If buttonstate&(i&) = 1 Then
                       GDMDIform.Toolbar1.Buttons(i&).value = tbrPressed
                       End If
                Next i&


             ElseIf DigitizeOn Then
             
                PopupMenu GDMDIform.mnuDigitize
             
             Else
                Beep
                'refresh center timer
                If ce& = 1 Then
                   GDMDIform.CenterPointTimer.Enabled = True
                   End If
             End If

         Case Else
       End Select
    End If
                
    Exit Sub
    
errhand:
   Select Case Err.Number
      Case 480 'autoredraw error--not enough memory--ignore
         If IgnoreAutoRedrawError% = 0 Then
            MsgBox "The pixel size of this map is too big for your memory!" & vbLf & vbLf & _
                   "If you wish to use this map and ignore such errors," & vbLf & _
                   "then check the ""Ignore AutoRedraw errors"" in the" & vbLf & _
                   """Settings"" tab of ""Path/Options"" form.", vbExclamation + vbOKOnly, "GSI_PDB"
            Exit Sub
         Else 'ignore this error
            Resume Next
            End If
      Case Else
         MsgBox "Encountered error #: " & Err.Number & vbLf & _
             Err.Description & vbLf & _
             "in module: Gdform1.Picture2_MouseUp", _
             vbCritical + vbOKOnly, "GSI_PDB"
   End Select
    
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'As cursor moves over map, display readout of coordinates.
  'Also detect drag.
  
  On Error GoTo errhand
  
  nearmouse_digi.x = x
  nearmouse_digi.y = y
  
  Dim next_mouse As PointAPI
  Dim hDnext As Long
  Dim R As Long
  Dim Next_Color As couleur
  
'  GetCursorPos next_mouse
'  hDnext = GetDC(0)
'  R = GetPixel(hDnext, next_mouse.X, next_mouse.Y)
'  If R <> -1 Then
'     Next_Color = recupcouleur(R)
'     GDMDIform.StatusBar1.Panels(1).Text = "RGB: " & Next_Color.R & "," & Next_Color.v & "," & Next_Color.b
'     End If
'
'  Call ReleaseDC(0, hDnext)
  
  
    '<<<<<<<<<<<<<<<<<new for digi>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> '<<<<<<<<<<<<digi changes
    If DigitizeOn Then
        'check if cursor left the picture frame, if so move scroll bars
        'to allow for dragging over entire map
        '(move the picture by the smallest increment = 1)
        If x / twipsx < Picture1.Left + HScroll1.value + 10 Then
           'scroll map to right
           
           'slow things down a bit to give time for repainting
           waitime = Timer
           Do Until Timer > waitime + 0.0001
              DoEvents
           Loop
              
           If HScroll1.value - 1 >= HScroll1.Min Then
              HScroll1.value = HScroll1.value - 1
              End If
        ElseIf x / twipsx > Picture1.Width + Picture1.Left + HScroll1.value - 10 Then
           'scroll map to left
           
           'slow things down a bit to give time for repainting
           waitime = Timer
           Do Until Timer > waitime + 0.0001
              DoEvents
           Loop
              
           If HScroll1.value + 1 <= HScroll1.Max Then
              HScroll1.value = HScroll1.value + 1
              End If
           End If
        If y / twipsy < Picture1.Top + VScroll1.value + 10 Then
           'scroll map down
           
           'slow things down a bit to give time for repainting
           waitime = Timer
           Do Until Timer > waitime + 0.0001
              DoEvents
           Loop
              
           If VScroll1.value - 1 >= VScroll1.Min Then
              VScroll1.value = VScroll1.value - 1
              End If
        ElseIf y / twipsy > Picture1.Top + Picture1.Height + VScroll1.value - 10 Then
           'scroll map up
           
           'slow things down a bit to give time for repainting
           waitime = Timer
           Do Until Timer > waitime + 0.0001
              DoEvents
           Loop
              
           If VScroll1.value + 1 <= VScroll1.Max Then
              VScroll1.value = VScroll1.value + 1
              End If
           End If
           
       'now open magnifier window if not already opened
           
'       GDMDIform.DigiTimer.Enabled = True
       
       If DigitizeLine And DigitizePadVis Then
          If Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString And Not DigitizeBeginLine Then
            'erase last line
            
'            If GDMDIform.CenterPointTimer.Enabled = True Then
'               ce& = 1 'flag that timer has been shut down during drag
'               GDMDIform.CenterPointTimer.Enabled = False
'               End If
            
            gddm = GDform1.Picture2.DrawMode
            gddw = GDform1.Picture2.DrawWidth
            
            If digi_last.x <> -999999 And digi_last.y <> -999999 And digi_begin.x <> -999999 And digi_begin.y <> -999999 Then
               'erase last line
                GDform1.Picture2.DrawMode = 7 'erase mode
                GDform1.Picture2.DrawWidth = 2
                
                digi_begin.x = blink_mark.x
                digi_begin.y = blink_mark.y
                digi_begin.Z = val(GDDigitizerfrm.txtelev.Text)
                
                If newblit = False Then
                    GDform1.Picture2.Line (new_digi.x, new_digi.y)-(digi_begin.x, digi_begin.y), QBColor(12)
                    End If
'                DigitizeDrawLine = False
                newblit = False
                GDform1.Picture2.Line (x, y)-(digi_begin.x, digi_begin.y), QBColor(12)
                new_digi.x = x
                new_digi.y = y
                new_digi.Z = val(GDDigitizerfrm.txtelev.Text)
                End If
                
            'record new position of end of line
            digi_last.x = x
            digi_last.y = y
            digi_last.Z = val(GDDigitizerfrm.txtelev.Text)
            
'            If digi_last.X <> -999999 And digi_last.Y <> -999999 And digi_begin.X <> -999999 And digi_begin.Y <> -999999 Then
                'now draw new line
                'GDform1.Picture2.DrawMode = 13 '7 '13 'drawing mode
                'GDform1.Picture2.Line (digi_last.X, digi_last.Y)-(digi_begin.X, digi_begin.Y), QBColor(12)
'                DigitizeDrawLine = True
'                End If
                
            GDform1.Picture2.DrawMode = gddm
            GDform1.Picture2.DrawWidth = gddw
            End If
          End If
          
'       If DigitizeContour And DigitizePadVis Then
'          If Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString Then
'            'draw line
'            gddm = GDform1.Picture2.DrawMode
'            gddw = GDform1.Picture2.DrawWidth
'
'            If digi_last.X <> -999999 And digi_last.Y <> -999999 Then
'                GDform1.Picture2.DrawMode = 7 '13
'                GDform1.Picture2.DrawWidth = 2
'                GDform1.Picture2.Line (X, Y)-(digi_last.X, digi_last.Y), QBColor(12)
'                End If
'
'            digi_last.X = X
'            digi_last.Y = Y
'            digi_last.Z = val(GDDigitizerfrm.txtelev.Text)
'            GDform1.Picture2.DrawMode = gddm
'            GDform1.Picture2.DrawWidth = gddw
'
''            If GDMDIform.CenterPointTimer.Enabled = False Then
''               ce& = 0 'flag that timer has been shut down during drag
''               GDMDIform.CenterPointTimer.Enabled = True
''               End If
'
'
'            End If
'          End If
    
       End If
   '<<<<<<<<<<<<<<<<<<<<<<<end of new>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
      
   If dragbegin = True And Button = 1 And dragbox = True Then 'dragging continues, draw box
      'continue dragging
      Picture2.DrawMode = 7
      Picture2.DrawStyle = vbDot
      Picture2.DrawWidth = 1
      
      'erase last drag box
      Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
      
      'check if cursor left the picture frame, if so move scroll bars
      'to allow for dragging over entire map
      '(move the picture by the smallest increment = 1)
      If x / twipsx < Picture1.Left + HScroll1.value Then
         'scroll map to right
         If HScroll1.value - 1 >= HScroll1.Min Then
            HScroll1.value = HScroll1.value - 1
            End If
      ElseIf x / twipsx > Picture1.Width + Picture1.Left + HScroll1.value Then
         'scroll map to left
         If HScroll1.value + 1 <= HScroll1.Max Then
            HScroll1.value = HScroll1.value + 1
            End If
         End If
      If y / twipsy < Picture1.Top + VScroll1.value Then
         'scroll map down
         If VScroll1.value - 1 >= VScroll1.Min Then
            VScroll1.value = VScroll1.value - 1
            End If
      ElseIf y / twipsy > Picture1.Top + Picture1.Height + VScroll1.value Then
         'scroll map up
         If VScroll1.value + 1 <= VScroll1.Max Then
            VScroll1.value = VScroll1.value + 1
            End If
         End If
      
      'draw new drag box
      Picture2.Line (x, y)-(drag1x, drag1y), QBColor(15), B
      Picture2.Refresh
      
      'record new drag end coordinates
      drag2x = x: drag2y = y
      
   ElseIf dragbegin = True And Button = 1 And dragbox = False And drawbox = False Then
      'begin dragging
      Picture2.DrawMode = 7
      Picture2.DrawStyle = vbDot
      Picture2.DrawWidth = 1
      Picture2.Line (x, y)-(drag1x, drag1y), QBColor(15), B
      drag2x = x: drag2y = y
      dragbox = True
      End If
  
  'Convert coordinates to pixels
  xcoord = x / twipsx
  ycoord = y / twipsy
  
  'Convert pixel coordinates to ITM
  ITMx = ((X2 - x1) / pixwi) * xcoord + x1
  ITMy = y1 - ((y1 - Y2) / pixhi) * ycoord
  
  'Display the ITM coordinates
  GDMDIform.Text1 = str(Int(ITMx))
  GDMDIform.Text2 = str(Int(ITMy))
  
  If heights = True And lblX = "ITMx" And LblY = "ITMy" Then 'display heights
     kmx = ITMx
     kmy = ITMy
     'Call DTMheight(kmx, kmy, hgt)
     Dim hgt As Integer
     Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
     GDMDIform.Text3 = str(hgt)
     End If
     
  If Geo And ShowContGeo Then 'also display geo coordinates
        kmxoo = ITMx: kmyoo = ITMy
        
        If GpsCorrection Then
            Dim lat_g As Double
            Dim lon_g As Double
            Dim N As Long
            Dim E As Long
            N = CLng(kmyoo)
            E = CLng(kmxoo)
            Call ics2wgs84(N, E, lat_g, lon_g)
            lt = lat_g
            lg = -lon_g
        Else
            Call casgeo(kmxoo, kmyoo, lg, lt)
            End If
            
        If GeoDecDeg = True Then
            GDGeoFrm.txtLat = Mid$(str$(lt), 1, 9)
            GDGeoFrm.txtLon = Mid$(str$(lg), 1, 9)
        Else
            lgdeg = Fix(lg)
            lgmin = Abs(Fix((lg - Fix(lg)) * 60))
            lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
            ltdeg = Fix(lt)
            ltmin = Abs(Fix((lt - Fix(lt)) * 60))
            ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
            If ltdeg = 0 And lt < 0 Then
              GDGeoFrm.txtLatDeg = "-" + str$(ltdeg) + "°"
              GDGeoFrm.txtLatMin = str$(ltmin) + "'"
              GDGeoFrm.txtLatSec = Mid$(str$(ltsec), 1, 6) + """"
            Else
              GDGeoFrm.txtLatDeg = str$(ltdeg) + "°"
              GDGeoFrm.txtLatMin = str$(ltmin) + "'"
              GDGeoFrm.txtLatSec = Mid$(str$(ltsec), 1, 6) + """"
            End If
            If lgdeg = 0 And lg < 0 Then
              GDGeoFrm.txtLonDeg = "-" + str$(lgdeg) + "°"
              GDGeoFrm.txtLonMin = str$(lgmin) + "'"
              GDGeoFrm.txtLonSec = Mid$(str$(lgsec), 1, 6) + """"
            Else
              GDGeoFrm.txtLonDeg = str$(lgdeg) + "°"
              GDGeoFrm.txtLonMin = str$(lgmin) + "'"
              GDGeoFrm.txtLonSec = Mid$(str$(lgsec), 1, 6) + """"
            End If
         End If
  End If
  Exit Sub

errhand:
   Select Case Err.Number
      Case 480
         If IgnoreAutoRedrawError% = 0 Then
            MsgBox "The pixel size of this map is too big for your memory!" & vbLf & vbLf & _
                   "If you wish to use this map and ignore such errors," & vbLf & _
                   "then check the ""Ignore AutoRedraw errors"" in the" & vbLf & _
                   """Settings"" tab of ""Path/Options"" form.", vbExclamation + vbOKOnly, "GSI_PDB"
            Exit Sub
         Else 'ignore this error
            Resume Next
            End If
      Case Else
         MsgBox "Encountered error #: " & Err.Number & vbLf & _
               Err.Description & vbLf & _
               "", vbCritical + vbOKOnly, "GSI_PDB"
   End Select
End Sub



Private Sub VScroll1_Change()
'   picx = 0
'   picy = 0
'   width1 = Picture1.Width
'   height1 = Picture1.Height
'   picx2 = (HScroll1.Value / HScroll1.Width) * pixwi 'sizex
'   picy2 = (VScroll1.Value / VScroll1.Height) * pixhi 'sizey
'   Picture1.PaintPicture GDMDIform.PictureClip1.Picture, picx, picy, width1, height1, picx2, picy2, width1, height1
   GDform1.Picture2.Top = -VScroll1.value
End Sub

Private Sub VScroll1_GotFocus()
    'print prompts on statusbar
    If SearchDB Then
      If NumReportPnts& = 0 Then
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
      Else
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
         End If
    Else
      GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define magnification boundaries."
     End If
End Sub
