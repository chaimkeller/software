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
         AutoSize        =   -1  'True
         Height          =   1935
         Left            =   540
         MousePointer    =   2  'Cross
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   189
         TabIndex        =   3
         Top             =   300
         Width           =   2895
         Begin VB.PictureBox PictureBlit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   360
            ScaleHeight     =   825
            ScaleWidth      =   945
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
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
   
   Screen.MousePointer = vbHourglass
   
   color_init.R = INIT_VALUE
   color_init.V = INIT_VALUE
   color_init.b = INIT_VALUE
   
   If Dir(picnam$) = sEmpty Then
      'try adding app path
      picnam$ = App.Path & "\" & picnam$
      If Dir(picnam$) = sEmpty Then Exit Sub
      End If
   
   twipsx = 1 'Screen.TwipsPerPixelX
   twipsy = 1 'Screen.TwipsPerPixelY

   SizeX = twipsx * pixwi
   SizeY = twipsy * pixhi
   
   With GDform1
   
       .left = 0
       .top = 0
       .Width = GDMDIform.ScaleWidth
       .Height = GDMDIform.ScaleHeight
       
       'Set ScaleMode to pixels (since picture size is in pixels)
       .ScaleMode = vbPixels
       .Picture1.ScaleMode = vbPixels
       
       'Autosize is set to True so that the boundaries of
       'Picture2 are expanded to the size of the actual map
       .Picture2.AutoSize = True
       
       'Set the BorderStyle of each picture box to None.
       'GDform1.Picture1.BorderStyle = 0
       .Picture2.BorderStyle = 0
        
       'load map to buffer
       .PictureBlit.Picture = LoadPicture(picnam$)
       
       DigiZoom.Zoom = 1#
       DigiZoom.LastZoom = 1#
       DigiZoom.left = 0
       DigiZoom.top = 0
       
       GDMDIform.StatusBar1.Panels(3).Text = CInt(100 * DigiZoom.LastZoom) & "%"
       
       'Load the default map to the visible picturebox at 100% zoom
       
       .Picture2.Width = CLng(DigiZoom.Zoom * pixwi)
       .Picture2.Height = CLng(DigiZoom.Zoom * pixhi)
       
       'check for maps that are larger than the maximum pixel size
       If .Picture2.Width < pixwi Then
          Select Case MsgBox("The map is wider than the maximum allowed width of: " & .Picture2.Width _
                             & vbCrLf & "" _
                             & vbCrLf & "If you continue working with this map, it will be truncated." _
                             & vbCrLf & "" _
                             & vbCrLf & "(Recommended to divide the map width wise into two maps" _
                             & vbCrLf & " E.g., split it with the following widths: " & .Picture2.Width & ", " & pixwi - .Picture2.Width _
                             & vbCrLf & "Then reload then separately as two separae maps.)" _
                             & vbCrLf & "" _
                             & vbCrLf & "Proceed anyways?" _
                             , vbYesNoCancel Or vbExclamation Or vbDefaultButton1 Or vbDefaultButton2, "Error in loading picture")
          
            Case vbYes
          
            Case vbNo, vbCancel
               Unload Me
               g_ier = -2
               Exit Sub
          End Select
          End If
          
       If .Picture2.Height < pixhi Then
          Select Case MsgBox("The map is taller than the maximum allowed height of: " & .Picture2.Height _
                             & vbCrLf & "" _
                             & vbCrLf & "If you continue working with this map, it will be truncated." _
                             & vbCrLf & "" _
                             & vbCrLf & "(Suggestion: divide the map height wise into two maps." _
                             & vbCrLf & " E.g., split it with the following widths: " & .Picture2.Height & ", " & pixhi - .Picture2.Height _
                             & vbCrLf & "Then reload then separately as two separae maps.)" _
                             & vbCrLf & "" _
                             & vbCrLf & "Proceed?" _
                             , vbYesNoCancel Or vbExclamation Or vbDefaultButton1 Or vbDefaultButton2, "Error in loading picture")
          
            Case vbYes
          
            Case vbNo, vbCancel
               Unload Me
               g_ier = -2
               Exit Sub
          End Select
          End If
          
       ier = StretchBlt(.Picture2.hdc, DigiZoom.left, DigiZoom.top, CLng(DigiZoom.Zoom * pixwi), CLng(DigiZoom.Zoom * pixhi), .PictureBlit.hdc, 0, 0, pixwi, pixhi, vbSrcCopy)
       
       'this is how to divide the picture and dump
'       ier = StretchBlt(.Picture2.hdc, DigiZoom.left, DigiZoom.top, CLng(DigiZoom.Zoom * pixwi), CLng(DigiZoom.Zoom * pixhi), .PictureBlit.hdc, CLng(DigiZoom.Zoom * pixwi) / 2, 0, pixwi, pixhi, vbSrcCopy)
       If ier = 0 Then 'stretchblt failed
    '      use Default
          .Picture2.Picture = LoadPicture(picnam$)
          End If
       
       DigiZoom.left = INIT_VALUE
       DigiZoom.top = INIT_VALUE
       
       'Initialize location of both pictures
       .Picture1.Move 0, 0, .ScaleWidth - VScroll1.Width, .ScaleHeight - .HScroll1.Height
       .Picture2.Move 0, 0
       
       'Position the horizontal scroll bar
       .HScroll1.top = .Picture1.Height
       .HScroll1.left = .Picture1.left
       .HScroll1.Width = .Picture1.Width
       
       'Position the vertical scroll bar
       .VScroll1.top = 0
       .VScroll1.left = .Picture1.Width
       .VScroll1.Height = .Picture1.Height
       
       'Set the Max property for the scroll bars.
       .HScroll1.Max = .Picture2.Width - .Picture1.Width
       .VScroll1.Max = .Picture2.Height - .Picture1.Height
       
       'Determine if the child picture will fill up the screen
       'If so, there is no need to use scroll bars.
       .VScroll1.Visible = (.Picture1.Height < .Picture2.Height)
       .HScroll1.Visible = (.Picture1.Width < .Picture2.Width)
       
       'Initiate Scroll Step Sizes
       .HScroll1.LargeChange = .HScroll1.Max / 20
       .HScroll1.SmallChange = .HScroll1.Max / 60
          
       .VScroll1.LargeChange = .VScroll1.Max / 20
       .VScroll1.SmallChange = .VScroll1.Max / 60
   
   End With
   
'    If DigitizePadVis And (DigitizeLine Or DigitizeContour Or DigitizePoint) Then
'       BringWindowToTop (GDDigitizerfrm.hWnd)
'       End If
   
   'iniital blit position to center of picture2
   nearmouse_digi.x = CLng(Picture1.Width * 0.5)
   nearmouse_digi.Y = CLng(Picture1.Height * 0.5)
   
   'begin listening to mouse wheel turns signifiying zooming out
   Call WheelHook(Me.hwnd)
   
   Screen.MousePointer = vbDefault
   Exit Sub

errhand:
   Select Case Err.Number
      Case 481 'invalid picture
         Screen.MousePointer = vbDefault
         MsgBox "Invalid picture!", vbCritical + vbOKOnly, "MapDigitizer"
      Case 380 'scroll bar errors, just resume
         Resume Next
      Case 480 'out of memory, picture is too big
        Call MsgBox("Insufficient video memory for drawing this map..." _
                    & vbCrLf & "" _
                    & vbCrLf & "(Hint: reduce the resolution or divide up the map.)" _
                    , vbExclamation, "Out of memory")
         g_ier = -1 'flag graceful recovery from memory error
      Case Else 'unexpected error
         Screen.MousePointer = vbDefault
         MsgBox "Encountered error #: " & Err.Number & vbLf & _
           "in program module GDform1:Form_Load." & vbLf & _
           Err.Description, vbCritical + vbOKOnly, "MapDigitizer"
  End Select
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'buttonstate&(1) = 0
   buttonstate&(23) = 0
   GDMDIform.Toolbar1.Buttons(1).value = tbrUnpressed
   GDMDIform.Toolbar1.Buttons(23).value = tbrUnpressed
   Call WheelUnHook(Me.hwnd)
   Unload Me
End Sub
Private Sub Form_Resize()

   On Error GoTo errhand
   
   Dim ResziedToZero As Boolean
   
   ResziedToZero = False
   
   If GDform1.ScaleHeight = 0 Then 'Exit Sub 'some sort of bug causes this, not sure of the source, but deal with it...
      ResziedToZero = True
      GoTo errhand
      End If

   GDform1.Picture1.Move 0, 0, GDform1.ScaleWidth - VScroll1.Width, GDform1.ScaleHeight - HScroll1.Height
   gdScHgt = GDform1.ScaleHeight
   GDform1.Picture2.Move 0, 0
   If magvis = True Then
      GDMagform.left = GDform1.left + GDform1.Width
      GDMagform.Width = GDMDIform.ScaleWidth - GDform1.Width
      End If
      
   If DigitizeMagvis Then
      GDDigiMagfrm.left = GDform1.left + GDform1.Width
      If MagWidth <> 0 And GDMDIform.ScaleWidth - GDform1.Width = 0 Then
         'restore mag window after the program was minimized
         GDDigiMagfrm.Width = MagWidth
         GDform1.Width = GDMDIform.ScaleWidth - MagWidth
         MagWidth = 0
      Else
         GDDigiMagfrm.Width = GDMDIform.ScaleWidth - GDform1.Width
         End If
      End If
      
   'Position the horizontal scroll bar
   HScroll1.top = GDform1.Picture1.Height
   HScroll1.left = GDform1.Picture1.left
   HScroll1.Width = GDform1.Picture1.Width
   
   'Position the vertical scroll bar
   VScroll1.top = 0
   VScroll1.left = GDform1.Picture1.Width
   VScroll1.Height = GDform1.Picture1.Height
   
   'Set the Max property for the scroll bars.
   HScroll1.Max = GDform1.Picture2.Width - GDform1.Picture1.Width
   VScroll1.Max = GDform1.Picture2.Height - GDform1.Picture1.Height
   
   'Determine if the child picture will fill up the screen
   'If so, there is no need to use scroll bars.
   GDform1.VScroll1.Visible = (GDform1.Picture1.Height < GDform1.Picture2.Height)
   GDform1.HScroll1.Visible = (GDform1.Picture1.Width < GDform1.Picture2.Width)
   
   GDform1.Picture2.left = -HScroll1.value
   GDform1.Picture2.top = -VScroll1.value
   
'    If DigitizePadVis And (DigitizeLine Or DigitizeContour Or DigitizePoint) Then
'       BringWindowToTop (GDDigitizerfrm.hWnd)
'       End If
     
   Exit Sub
errhand:
   If ResziedToZero Then
   
      With GDform1
        DigiZoom.left = INIT_VALUE
        DigiZoom.top = INIT_VALUE
       
       'Initialize location of both pictures
       .Picture1.Move 0, 0, .ScaleWidth - VScroll1.Width, .ScaleHeight - .HScroll1.Height
       .Picture2.Move 0, 0
       
       'Position the horizontal scroll bar
       .HScroll1.top = .Picture1.Height
       .HScroll1.left = .Picture1.left
       .HScroll1.Width = .Picture1.Width
       
       'Position the vertical scroll bar
       .VScroll1.top = 0
       .VScroll1.left = .Picture1.Width
       .VScroll1.Height = .Picture1.Height
       
       'Set the Max property for the scroll bars.
       .HScroll1.Max = .Picture2.Width - .Picture1.Width
       .VScroll1.Max = .Picture2.Height - .Picture1.Height
       
       'Determine if the child picture will fill up the screen
       'If so, there is no need to use scroll bars.
       .VScroll1.Visible = (.Picture1.Height < .Picture2.Height)
       .HScroll1.Visible = (.Picture1.Width < .Picture2.Width)
       
       'Initiate Scroll Step Sizes
       .HScroll1.LargeChange = .HScroll1.Max / 20
       .HScroll1.SmallChange = .HScroll1.Max / 60
          
       .VScroll1.LargeChange = .VScroll1.Max / 20
       .VScroll1.SmallChange = .VScroll1.Max / 60
       
     End With
   
     ier = ReDrawMap(0)
     
     ResziedToZero = False
   
   Else
      MsgBox "Error " & Err.Number & " " & Err.Description & " encountered in module GDform1.Resize", vbExclamation + vbOKOnly, "Error"
      End If
End Sub

Private Sub HScroll1_Change()
'   picx = 0
'   picy = 0
'   width1 = Picture1.Width
'   height1 = Picture1.Height
'   picx2 = (HScroll1.Value / HScroll1.Width) * pixwi 'sizex
'   picy2 = (VScroll1.Value / VScroll1.Height) * pixhi 'sizey
'   Picture1.PaintPicture GDMDIform.PictureClip1.Picture, picx, picy, width1, height1, picx2, picy2, width1, height1
   GDform1.Picture2.left = -HScroll1.value
End Sub

Private Sub HScroll1_GotFocus()
    'print prompts on statusbar
    If SearchDigi Then
      If NumReportPnts& = 0 Then
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
      Else
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
         End If
    ElseIf Not DigitizeOn And Not DigiRS Then
         GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
    ElseIf DigitizeOn Or DigiRS Then
       GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
       If DigitizePadVis Then 'And (DigitizeLine Or DigitizeContour Or DigitizePoint) Then
          GDDigitizerfrm.Visible = True
          BringWindowToTop (GDDigitizerfrm.hwnd)
          End If
       If DigiRS Then
          GDRSfrm.Visible = True
          BringWindowToTop (GDRSfrm.hwnd)
          End If
       End If
End Sub

Private Sub Picture2_GotFocus()
    'print prompts on statusbar
    If Not DigitizeOn And Not DigiRS Then
        If SearchDigi Then
          If NumReportPnts& = 0 Then
             GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
          ElseIf Not digitizon Then
             GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
          ElseIf DigitizeOn Then
             GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location and click to center the screen and choose that point"
             End If
        Else
         GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
         End If
    ElseIf DigitizeOn Or DigiRS Then
       'enable magnification of curson position
'       GDMDIform.DigiTimer.Enabled = True
       GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
       End If
    
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
'   a = KeyCode
   'up arrow = 38
   'down arrow = 40
   'left arrow = 37
   'right arrow = 39
   'PgUp = 33
   'PgDown = 34
   If KeyCode = Asc("Z") Or KeyCode = Asc("z") Then
      'zoom out
      Call PictureBoxZoom(GDform1.Picture2, 0, -120, 0, 0, 0)
      End If
      
   If KeyCode = Asc("X") Or KeyCode = Asc("x") Then
      'zoom in
      Call PictureBoxZoom(GDform1.Picture2, 0, 120, 0, 0, 0)
      End If
      
   If KeyCode = vbKeyR And DigiEditPoints Then
      'Edit replace mode
        
       'shift the point and replot
       If XpixLast <> -1 And YpixLast <> -1 Then
          ier = RedrawDigiPoints(Nint(XCoord), Nint(Ycoord), DigiEditMode, 0)
          End If
      
   ElseIf KeyCode = vbKeyK And DigiEditPoints Then
      'Edit kill digitized point mode
      
       'shift the point and replot
       If XpixLast <> -1 And YpixLast <> -1 Then
          ier = RedrawDigiPoints(Nint(XCoord), Nint(Ycoord), DigiEditMode, 1)
          End If
      
   ElseIf KeyCode = vbKeyLeft Then 'left arrow
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


Private Sub Picture2_LostFocus()

'    If DigitizeOn Then
'       GDMDIform.DigiTimer.Enabled = False
'       End If
       
    GDMDIform.StatusBar1.Panels(1) = sEmpty
   
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim ier As Integer
   ier = MouseDown(Button, Shift, x, Y)
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim ier As Integer
   ier = MouseMove(Button, Shift, x, Y)
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim ier As Integer
   ier = MouseUp(Button, Shift, x, Y)
End Sub

Private Sub VScroll1_Change()
   GDform1.Picture2.top = -VScroll1.value
End Sub

Private Sub VScroll1_GotFocus()
    'print prompts on statusbar
    'print prompts on statusbar
    If SearchDigi Then
      If NumReportPnts& = 0 Then
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
      Else
         GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
         End If
    ElseIf Not DigitizeOn And Not DigiRS Then
       GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
    ElseIf DigitizeOn Or DigiRS Then
       GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
       If DigitizePadVis Then 'And (DigitizeLine Or DigitizeContour Or DigitizePoint) Then
          GDDigitizerfrm.Visible = True
          BringWindowToTop (GDDigitizerfrm.hwnd)
       ElseIf DigiRS Then
          GDRSfrm.Visible = True
          BringWindowToTop (GDRSfrm.hwnd)
          End If
       End If
End Sub

Public Function min(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  min = v1
Else: min = v2
End If
End Function

Public Function Max(ByVal v1 As Double, ByVal v2 As Double) As Double
If (v1 < v2) Then
  Max = v2
Else: Max = v1
End If
End Function

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' source : wheel mouse hook: http://www.vbforums.com/showthread.php?388222-VB6-MouseWheel-with-Any-Control-%28originally-just-MSFlexGrid-Scrolling%29
'          two files examples: WheelHook-AllControls.zip, WheelHook-NestedControls.zip

' ===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    PictureBoxZoom GDform1.Picture2, MouseKeys, Rotation, Xpos, Ypos, 0
 
'  'original WheelWheel code for interacting with very controls on the form is below
'  Dim ctl As Control, cContainerCtl As Control
'  Dim bHandled As Boolean
'  Dim bOver As Boolean
'
'  For Each ctl In Controls
'    ' Is the mouse over the control
'    On Error Resume Next
'    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
'    On Error GoTo 0
'
'    If bOver Then
'      ' If so, respond accordingly
'      bHandled = True
'      Select Case True
'
'        Case TypeOf ctl Is MSFlexGrid
'          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
'
'        Case TypeOf ctl Is PictureBox, TypeOf ctl Is Frame
'          Set cContainerCtl = ctl
'          bHandled = False
'
'        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
'          ' These controls already handle the mousewheel themselves, so allow them to:
'          If ctl.Enabled Then ctl.SetFocus
'
'        Case Else
'          bHandled = False
'
'      End Select
'      If bHandled Then Exit Sub
'    End If
'    bOver = False
'    Debug.Print ctl.Name
'  Next ctl
'
'  If Not cContainerCtl Is Nothing Then
'    If TypeOf cContainerCtl Is PictureBox Then PictureBoxZoom GDform1.Picture2, MouseKeys, Rotation, Xpos, Ypos, 0
'  Else
'    ' Scroll was not handled by any controls, so treat as a general message send to the form
'    GDMDIform.StatusBar1.Panels(1) = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
'  End If
End Sub
'MouseDown Event for Picture2
Public Function MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) As Integer
   On Error GoTo errhand
        
    Dim twipscX As Long, twipscY As Long
    twipscX = Screen.TwipsPerPixelX
    twipscY = Screen.TwipsPerPixelY

   
    If DigitizePadVis Then
    
       Dim pt As POINTAPI, Cont As POINTAPI
       Dim currX As Long, currY As Long
       Dim pt2 As POINTAPI
       Dim df As RECT
    
      'if clicking inside domain of GDDigitizerfrm then exit
      'convert X,Y to Screen coordinate
      
      'determine screen coordinates of form
       pt.x = 0
       pt.Y = 0
    
       ClientToScreen GDDigitizerfrm.hwnd, pt
       currX = pt.x
       currY = pt.Y
       
       'determine screen coordinates of mouse
       pt.x = x
       pt.Y = Y
       ClientToScreen GDform1.Picture2.hwnd, pt
       
       'detect if mouse was clicked inside the form (the form has scalemode = vbtwips)
       With GDDigitizerfrm
        
        If pt.x >= currX And pt.Y >= currY _
           And pt.x <= currX + .ScaleWidth / twipscX _
           And pt.Y <= currY + .Height / twipscY Then
           
          'chkOcean
          Call ConvClientToScreen(.chkOcean, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             'chkOcean invert entered heights for bathymetry check box
             Call .chkOcean_Click
             MouseDown = -1
             Exit Function
             End If
          
          'txtelev
          Call ConvClientToScreen(.txtelev, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             'txtelev text box
             Call .txtelev_GotFocus
             MouseDown = -1
             Exit Function
             End If
             
          'cmdClear
          Call ConvClientToScreen(.cmdClear, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmdDigiClear_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmdspace
          Call ConvClientToScreen(.cmdSpace, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmdSpace_Click
             MouseDown = -1
             Exit Function
             End If
          
          'cmdEnter
          Call ConvClientToScreen(.cmdEnter, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmdEnter_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd0
          Call ConvClientToScreen(.cmd0, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd0_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd1
          Call ConvClientToScreen(.cmd1, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd1_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd2
          Call ConvClientToScreen(.cmd2, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd2_Click
             MouseDown = -1
             Exit Function
             End If
   
          'cmd3
          Call ConvClientToScreen(.cmd3, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd3_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd4
          Call ConvClientToScreen(.cmd4, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd4_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd5
          Call ConvClientToScreen(.cmd5, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd5_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd6
          Call ConvClientToScreen(.cmd6, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd6_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd7
          Call ConvClientToScreen(.cmd7, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd7_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd8
          Call ConvClientToScreen(.cmd8, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd8_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmd9
          Call ConvClientToScreen(.cmd9, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmd9_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmdC
          Call ConvClientToScreen(.cmdC, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmdC_Click
             MouseDown = -1
             Exit Function
             End If
             
          'cmdD
          Call ConvClientToScreen(.cmdD, df)
          If pt.x >= df.X1 And pt.x <= df.X2 And _
             pt.Y >= df.Y1 And pt.Y <= df.Y2 Then
             Call .cmdD_Click
             MouseDown = -1
             Exit Function
             End If
             
          MouseDown = -1
          Exit Function
          End If
          
       End With
    End If
   
   If Button = 1 And _
      Not DigitizerEraser Then
      drag1x = x
      drag1y = Y
      dragbegin = True
      drag2x = drag1x
      drag2y = drag1y
      End If

     
   If Not DigitizeOn And Not DigitizerEraser And Not DigitizeHardy Then
      'shut off timers during drag
      ce& = 0 'reset blinker flag
      If GDMDIform.CenterPointTimer.Enabled = True Then
         ce& = 1 'flag that timer has been shut down during drag
         GDMDIform.CenterPointTimer.Enabled = False
         End If
      End If
      
   MouseDown = 0
   Exit Function
   
errhand:
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          "in module: Gdform1.Picture2_MouseDown", _
          vbCritical + vbOKOnly, "MapDigitizer"
   MouseDown = -1

End Function
'mouseup event for gdform1.picture2
Public Function MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) As Integer
  Dim VarD As Double
  Dim BytePosit As Long
  
  Dim RecNum&
  Dim Byte0 As Byte
  Dim Byte1 As Byte
  Dim Byte2 As Byte
  Dim Byte3 As Byte
  Dim Byte4 As Byte
  Dim Byte5 As Byte
  
  Byte0 = 0
  Byte1 = 1
  Byte2 = 2
  Byte3 = 3
  Byte4 = 4
  Byte5 = 5
  
  Dim color_line As Long, colornum%
  
  'heights
  Dim kmx As Long, kmy As Long
  Dim lt2 As Double, lg2 As Double, hgt2 As Integer
  
  Dim SearchCoord(1) As POINTAPI
  Dim ContourCoord(1) As POINTAPI
  Dim SmoothCoord(1) As POINTAPI
  Dim DTMCoord(1) As POINTAPI
  
  On Error GoTo errhand
  
  nearmouse_digi.x = x
  nearmouse_digi.Y = Y
  
  GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
  
  If GeoMap = True Or TopoMap = True Then 'if the map is visible
      XCoord = x
      Ycoord = Y
      
      If Geo = True Then
         'keep geo coordinate window visible
         Ret = BringWindowToTop(GDGeoFrm.hwnd)
         End If
      
      Select Case Button
         Case 1  'left button
            'shift this point to middle of screen
            'this will be the case when (X,Y) = (picture1.width/2, picture1.height/2)
            
gd50:       If (drag1x = drag2x And drag1y = drag2y) Then 'And Not DigitizeOn Then
                dragbegin = False
                dragbox = False
                
                'reset center timer if flagged
                If ce& = 1 Then 'blinker was shut down during drag, so reenable it
                   ce& = 0 'reset blinker flag
                   GDMDIform.CenterPointTimer.Enabled = True
                   End If
            Else 'signales end of drag
               End If
               
            If dragbox = True And dragbegin = True And _
              ((drag2x = drag1x And drag2y <> drag1y) Or (drag2x <> drag1x And drag2y = drag1y)) And _
              Button = 1 And _
              (SearchDigi _
              Or HeightSearch _
              Or DTMcreating _
              Or GenerateContours _
              Or DigitizeExtendGrid _
              Or DigitizerSweep _
              Or Belgier_Smoothing) And _
              Not DigitizerEraser Then
                'defines box with no internal area, erase it and start again
                Picture2.DrawMode = 7
                Picture2.DrawStyle = vbDot
                Picture2.DrawWidth = 1
                Picture2.Line (x, Y)-(drag1x, drag1y), QBColor(15), B
                drawbox = False
                dragbegin = False
                'reset center timer if flagged
                If ce& = 1 Then 'blinker was shut down during drag operation
                   ce& = 0 'reset blinker flag
                   GDMDIform.CenterPointTimer.Enabled = True
                   End If
            
            ElseIf dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 And _
                 Not HeightSearch And _
                 Not DTMcreating And _
                 Not GenerateContours And _
                 Not SearchDigi And _
                 Not DigitizeHardy And _
                 Not DigitizeExtendGrid And _
                 Not DigitizerEraser And _
                 Not DigitizerSweep And _
                 Not Belgier_Smoothing Then
                'USER IS DRAGGING THE MAP
                'move map by difference of drag2 and drag1
                'treat this as though the scroll bars have been clicked
                'Shift the scroll bars in order to accomplish the above
                h1 = drag1x - drag2x '<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>
                If GDform1.HScroll1.value + h1 < GDform1.HScroll1.min Or GDform1.HScroll1.value + h1 > GDform1.HScroll1.Max Then
                      'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
                      If GDform1.Picture2.Width > GDform1.HScroll1.Width Then

                         If GDform1.HScroll1.value + h1 < GDform1.HScroll1.min Then
                            GDform1.HScroll1.value = GDform1.HScroll1.min
                         ElseIf GDform1.HScroll1.value + h1 > GDform1.HScroll1.Max Then
                            GDform1.HScroll1.value = GDform1.HScroll1.Max
                            End If
'
                         End If
                Else
                   GDform1.HScroll1.value = GDform1.HScroll1.value + h1
                   End If
                
                h2 = drag1y - drag2y
                If GDform1.VScroll1.value + h2 < 0 Or GDform1.VScroll1.value + h2 > GDform1.VScroll1.Max Then
'                      'scroll maximum or minimum amount possible and only if the map tile is larger than the canvas
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
            
                'reset the drag flags
               
                Picture2.DrawMode = 13
                drawbox = False
                dragbegin = False
                'reset drag coordinates
                drag1x = 0
                drag2x = 0
                drag1y = 0
                drag2y = 0

                'refresh toolbar1
                For i& = 1 To GDMDIform.Toolbar1.Buttons.count
                    If buttonstate&(i&) = 1 Then
                       GDMDIform.Toolbar1.Buttons(i&).value = tbrPressed
                       End If
                Next i&
                
                If DigitizePadVis And (DigitizeLine Or DigitizeContour Or DigitizePoint Or DigitizeBlankPoint) Then
                   GDDigitizerfrm.Visible = True
                   BringWindowToTop (GDDigitizerfrm.hwnd)
                ElseIf DigiRS Then
                   GDRSfrm.Visible = True
                   BringWindowToTop (GDRSfrm.hwnd)
                   End If
                   
'                If SearchDigi = False Then
                   'sit here until user closes mag box
                   'shut off timers
                   Do Until magclose = True Or DigitizeMagvis = False
                     DoEvents
                   Loop
'                   End If

               
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
                  
'/////////////////////////////////////////////////////////////////
            ElseIf SearchDigi And Not DTMcreating And Not HeightSearch And Not GenerateContours And Not DigitizerSweep And Not Belgier_Smoothing And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               'end of selecting points for searching for highest place
               
               'erase last dotted line
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'draw new line
               Picture2.DrawWidth = 2
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'restore drawing mode to normal
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
                
               'mag box closed, so erase box
               If drawbox Then
                  Picture2.DrawMode = 7
                  Picture2.DrawWidth = 2
                  Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                  Picture2.DrawWidth = 1
                  Picture2.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  
                  'define rectangular window for report

                  ReportCoord(0).x = min(drag1x, drag2x)
                  ReportCoord(0).Y = min(drag1y, drag2y)
                  ReportCoord(1).x = Max(drag1x, drag2x)
                  ReportCoord(1).Y = Max(drag1y, drag2y)
                  
                  'now renormalize for zooming, if any
                  ReportCoord(0).x = CLng(ReportCoord(0).x / DigiZoom.LastZoom)
                  ReportCoord(0).Y = CLng(ReportCoord(0).Y / DigiZoom.LastZoom)
                  ReportCoord(1).x = CLng(ReportCoord(1).x / DigiZoom.LastZoom)
                  ReportCoord(1).Y = CLng(ReportCoord(1).Y / DigiZoom.LastZoom)
                  
                  'reset drag coordinates
                  drag1x = 0
                  drag2x = 0
                  drag1y = 0
                  drag2y = 0
                  
                  Call GDMDIform.mnuReport_Click
                  
                  End If
                  
'//////////////////////////////////////////////////////////////////
            ElseIf HeightSearch And Not DTMcreating And Not SearchDigi And Not GenerateContours And Not DigitizerSweep And Not Belgier_Smoothing And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               'end of selecting points for sweeping erasure
               
               'erase last dotted line
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'draw new line
               Picture2.DrawWidth = 2
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'restore drawing mode to normal
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
                
               'mag box closed, so erase box
               If drawbox Then
                  Picture2.DrawMode = 7
                  Picture2.DrawWidth = 2
                  Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                  Picture2.DrawWidth = 1
                  Picture2.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  
                  'define rectangular window for search of heights

                  SearchCoord(0).x = min(drag1x, drag2x)
                  SearchCoord(0).Y = min(drag1y, drag2y)
                  SearchCoord(1).x = Max(drag1x, drag2x)
                  SearchCoord(1).Y = Max(drag1y, drag2y)
                  
                  'now renormalize for zooming, if any
                  SearchCoord(0).x = CLng(SearchCoord(0).x / DigiZoom.LastZoom)
                  SearchCoord(0).Y = CLng(SearchCoord(0).Y / DigiZoom.LastZoom)
                  SearchCoord(1).x = CLng(SearchCoord(1).x / DigiZoom.LastZoom)
                  SearchCoord(1).Y = CLng(SearchCoord(1).Y / DigiZoom.LastZoom)
                  
                  'reset drag coordinates
                  drag1x = 0
                  drag2x = 0
                  drag1y = 0
                  drag2y = 0
                  
                  ier = SearchMaxHeights(GDform1.Picture2, SearchCoord)
                  
                  End If


'/////////////////////////////////////////////////////////////////

            ElseIf DTMcreating And Not HeightSearch And Not SearchDigi And Not GenerateContours And Not DigitizerSweep And Not Belgier_Smoothing And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               'end of selecting points for sweeping erasure
               
               'erase last dotted line
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'draw new line
               Picture2.DrawWidth = 2
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'restore drawing mode to normal
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
                
               'mag box closed, so erase box
               If drawbox Then
                  
                  'define rectangular window for DTM creation

                  DTMCoord(0).x = min(drag1x, drag2x)
                  DTMCoord(0).Y = min(drag1y, drag2y)
                  DTMCoord(1).x = Max(drag1x, drag2x)
                  DTMCoord(1).Y = Max(drag1y, drag2y)
                  
                  'now renormalize for zooming, if any
                  SearchCoord(0).x = CLng(DTMCoord(0).x / DigiZoom.LastZoom)
                  SearchCoord(0).Y = CLng(DTMCoord(0).Y / DigiZoom.LastZoom)
                  SearchCoord(1).x = CLng(DTMCoord(1).x / DigiZoom.LastZoom)
                  SearchCoord(1).Y = CLng(DTMCoord(1).Y / DigiZoom.LastZoom)
                  
                  ier = CreateDTM(GDform1.Picture2, SearchCoord)
                  
                  If ier < 0 Then 'remove guide line, otherwise the merged area is highlighted and shouldn't be overwriten by drawmode = 7
                    Picture2.DrawMode = 7
                    Picture2.DrawWidth = 2
                    Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                    End If
                    
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


'/////////////////////////////////////////////////////////////////

            ElseIf GenerateContours And Not DTMcreating And Not HeightSearch And Not SearchDigi And Not DigitizerSweep And Not Belgier_Smoothing And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               'end of selecting points for sweeping erasure
               
               'erase last dotted line
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'draw new line
               Picture2.DrawWidth = 2
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'restore drawing mode to normal
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
                
               If drawbox Then
                  Picture2.DrawWidth = 1
                  Picture2.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  
                  'define rectangular window for displaying contours

                  ContourCoord(0).x = min(drag1x, drag2x)
                  ContourCoord(0).Y = min(drag1y, drag2y)
                  ContourCoord(1).x = Max(drag1x, drag2x)
                  ContourCoord(1).Y = Max(drag1y, drag2y)
                  
                  'now renormalize for zooming, if any
                  ContourCoord(0).x = CLng(ContourCoord(0).x / DigiZoom.LastZoom)
                  ContourCoord(0).Y = CLng(ContourCoord(0).Y / DigiZoom.LastZoom)
                  ContourCoord(1).x = CLng(ContourCoord(1).x / DigiZoom.LastZoom)
                  ContourCoord(1).Y = CLng(ContourCoord(1).Y / DigiZoom.LastZoom)
                  
                  'reset drag coordinates
                  drag1x = 0
                  drag2x = 0
                  drag1y = 0
                  drag2y = 0
                  
                  ier = Contours(GDform1.Picture2, ContourCoord)
                  
                  If ier < 0 Then 'redraw screen and depress button
                     Call GDMDIform.mnuContour_Click
                     MouseUp = -1
                     Exit Function
                     End If
                  
                  If Save_xyz% = 1 Then GDMDIform.Toolbar1.Buttons(52).Enabled = True 'allow profiling
                  
                  End If

'/////////////////////////////////////////////////////////////////

            ElseIf Belgier_Smoothing And Not GenerateContours And Not HeightSearch And Not SearchDigi And Not DigitizerSweep And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               'end of selecting points for sweeping erasure
               
               'erase last dotted line
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
'               'draw new line
'               Picture2.DrawWidth = 2
'               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'restore drawing mode to normal
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
                
               If drawbox Then
'                  gbfc = Picture2.ForeColor
'                  Picture2.ForeColor = QBColor(4)
'                  Picture2.DrawWidth = 1
'                  Picture2.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  
                  'define rectangular window for displaying contours

                  SmoothCoord(0).x = min(drag1x, drag2x)
                  SmoothCoord(0).Y = min(drag1y, drag2y)
                  SmoothCoord(1).x = Max(drag1x, drag2x)
                  SmoothCoord(1).Y = Max(drag1y, drag2y)
                  
                  'now renormalize for zooming, if any
                  SmoothCoord(0).x = CLng(SmoothCoord(0).x / DigiZoom.LastZoom)
                  SmoothCoord(0).Y = CLng(SmoothCoord(0).Y / DigiZoom.LastZoom)
                  SmoothCoord(1).x = CLng(SmoothCoord(1).x / DigiZoom.LastZoom)
                  SmoothCoord(1).Y = CLng(SmoothCoord(1).Y / DigiZoom.LastZoom)
                  
                  'reset drag coordinates
                  drag1x = 0
                  drag2x = 0
                  drag1y = 0
                  drag2y = 0
                  
                  Picture2.ForeColor = gbfc
                  
                  ier = Smoothing(GDform1.Picture2, SmoothCoord)
                  
                  If ier < 0 Then 'redraw screen and depress button
                     Call GDMDIform.mnuSmooth_Click
                     MouseUp = -1
                     Exit Function
                     End If
                  
                  End If

'/////////////////////////////////////////////////////////////////
                  
            ElseIf DigitizeHardy And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               'end of digihardy drag
               
               'erase last dotted line
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'draw new line
               Picture2.DrawWidth = 2
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'restore drawing mode to normal
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
                
               If drawbox Then
                  Picture2.DrawWidth = 1
                  Picture2.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  
                  'define rectangular window for Hardy Quadratic Surface analysis
                  Dim HardyDragCoord(1) As POINTAPI

                  HardyDragCoord(0).x = min(drag1x, drag2x)
                  HardyDragCoord(0).Y = min(drag1y, drag2y)
                  HardyDragCoord(1).x = Max(drag1x, drag2x)
                  HardyDragCoord(1).Y = Max(drag1y, drag2y)
                  
                  'now renormalize for zooming, if any
                  HardyDragCoord(0).x = CLng(HardyDragCoord(0).x / DigiZoom.LastZoom)
                  HardyDragCoord(0).Y = CLng(HardyDragCoord(0).Y / DigiZoom.LastZoom)
                  HardyDragCoord(1).x = CLng(HardyDragCoord(1).x / DigiZoom.LastZoom)
                  HardyDragCoord(1).Y = CLng(HardyDragCoord(1).Y / DigiZoom.LastZoom)
                  
                  'reset drag coordinates
                  drag1x = 0
                  drag2x = 0
                  drag1y = 0
                  drag2y = 0
                  End If
               
                  
                If numDigiPoints <> 0 Or numDigiContours <> 0 Or numDigiLines <> 0 And DigitizeHardy Then
                
                   'show progress bar while converting all points, lines, contours into Hardy scan, '<<<<<<<<<<<<<<<<<<<<<<<<<<<
                   'then show contours in the GDMagform picturebox.
                   Dim modeHardy As Integer
                   modeHardy = 0 'mode = 0 finds new points, mode = 1 adds to the points already found
                   ier = FindPointsHardy(HardyDragCoord, modeHardy, sEmpty)
                   If ier = -1 Then
                      Call MsgBox("Error encountered in the search for relevant points", vbExclamation, "Hardy quadratic surface error")
                      MouseUp = -1
                      
                      ier = ReDrawMap(0)
                      If Not InitDigiGraph Then
                         InputDigiLogFile 'load up saved digitizing data for the current map sheet
                      Else
                         ier = RedrawDigiLog
                         End If
                      Exit Function
                      End If
                      
                   'now run Hardy Quadratic surface analysis and contours
                   ier = HardyQuadraticSurfaces(GDform1.Picture2, HardyDragCoord)
                   
                   Inside_Hardy_Calculation = False
                   
                   If ier < 0 Then
                        If ier = -1 Then
                           Call MsgBox("Error encountered in Hardy Quadratic Surface Routine", vbExclamation, "Hardy quadratic surface error")
                        ElseIf ier = -2 Then 'out of memory
                        
                           Call MsgBox("Not enough memory available for calculating the selected region!" _
                                       & vbCrLf & "" _
                                       & vbCrLf & "Hint:" _
                                       & vbCrLf & "1. Select a smaller region." _
                                       & vbCrLf & "2. Redo the digitizing using larger spacing between" _
                                       & vbCrLf & "   line and contour vertices (Options menu)." _
                                       , vbInformation, "Hardy Quadratic Surfaces Error")
                           
                           End If
                           
                         MouseUp = -1
                         
                         ier = ReDrawMap(0)
                         If Not InitDigiGraph Then
                            InputDigiLogFile 'load up saved digitizing data for the current map sheet
                         Else
                            ier = RedrawDigiLog
                            End If
                            
                         Exit Function
                         End If
                   
                    End If
                   
            ElseIf DigitizerSweep And dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
               'end of selecting points for sweeping erasure
               
               'erase last dotted line
               Picture2.DrawMode = 7
               Picture2.DrawStyle = vbDot
               Picture2.DrawWidth = 1
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'draw new line
               Picture2.DrawWidth = 2
               Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
               'restore drawing mode to normal
               Picture2.DrawMode = 13
               dragbegin = False
               dragbox = False
               drawbox = True
               Picture2.DrawWidth = 1
1
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
                
               'mag box closed, so erase box
               If drawbox Then
                  Picture2.DrawMode = 7
                  Picture2.DrawWidth = 2
                  Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                  Picture2.DrawWidth = 1
                  Picture2.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  
                  'define rectangular window for Hardy Quadratic Surface analysis
                  Dim SweepDragCoord(1) As POINTAPI

                  SweepDragCoord(0).x = min(drag1x, drag2x)
                  SweepDragCoord(0).Y = min(drag1y, drag2y)
                  SweepDragCoord(1).x = Max(drag1x, drag2x)
                  SweepDragCoord(1).Y = Max(drag1y, drag2y)
                  
                  'now renormalize for zooming, if any
                  SweepDragCoord(0).x = CLng(SweepDragCoord(0).x / DigiZoom.LastZoom)
                  SweepDragCoord(0).Y = CLng(SweepDragCoord(0).Y / DigiZoom.LastZoom)
                  SweepDragCoord(1).x = CLng(SweepDragCoord(1).x / DigiZoom.LastZoom)
                  SweepDragCoord(1).Y = CLng(SweepDragCoord(1).Y / DigiZoom.LastZoom)
                  
                  'reset drag coordinates
                  drag1x = 0
                  drag2x = 0
                  drag1y = 0
                  drag2y = 0
                  End If
               
                  
                If numDigiPoints <> 0 Or numDigiContours <> 0 Or numDigiLines <> 0 Or numDigiErase <> 0 And DigitizerSweep Then
                
                   'show progress bar while converting all points, lines, contours into Hardy scan, '<<<<<<<<<<<<<<<<<<<<<<<<<<<
                   'then show contours in the GDMagform picturebox.
                   ier = EraseSweepPoints(SweepDragCoord)
                   If ier = -1 Then
                      Call MsgBox("Error encountered in the search for relevant points", vbExclamation, "Sweep erasure error")
                      MouseUp = -1
                      Exit Function
                      End If
                      
                   'now replot the digitized points, etc.
                   ier = ReDrawMap(0)
                   If Not InitDigiGraph Then
                      InputDigiLogFile 'load up saved digitizing data for the current map sheet
                   Else
                      ier = RedrawDigiLog
                      End If
                      
                    If DigitizeMagvis Then
                        DoEvents
                        Ret = SetWindowPos(GDDigiMagfrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                        End If
                
                   
                   End If
                   
            Else 'shift map and draw circle at click point
            
                ce& = 1 'flag to draw blinker at new position
            
                DigiZoom.left = CLng(x / DigiZoom.LastZoom)
                DigiZoom.top = CLng(Y / DigiZoom.LastZoom)
                
                If (DigitizePoint Or DigitizeBlankPoint) And PointCenterClick = 1 Then
                   'don't shift map while digitizing points
                Else
                   Call ShiftMap(x, Y)
                   End If
                
                'put click coordinates into coordinate boxes
                'Convert coordinates to pixels
                XCoord = CLng(x / (twipsx * DigiZoom.LastZoom))
                Ycoord = CLng(Y / (twipsy * DigiZoom.LastZoom))
                
                'Convert pixel coordinates to ITM (or to any user's coordinate system)
                ITMx = ((LRGeoX - ULGeoX) / pixwi) * XCoord + ULGeoX
                ITMy = ULGeoY - ((ULGeoY - LRGeoY) / pixhi) * Ycoord
                
                    
                'Display height at click point
                
'                If Not Digitizing And heights = True And lblX = "ITMx" And LblY = "ITMy" Then 'display heights
'                   kmx = ITMx
'                   kmy = ITMy
'                   'Call DTMheight(kmx, kmy, hgt)
'                   Dim hgt As Integer
'                   Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
'                   GDMDIform.Text7 = str(hgt)
'                   End If

                'Write position file to the hard disk
                'These coordinates will be used as the
                'starting position for the next time the
                'user logs into the program.  It can be also
                'used by the Access database program for
                'automatically inputing coordinates (to reduce
                'human error while inputing coordinates)
                If Not Digitizing Then
                   Call UpdatePositionFile(ITMx * DigiZoom.LastZoom, ITMy * DigiZoom.LastZoom, hgt)
                Else
                   Call UpdatePositionFile(XCoord, Ycoord, hgt)
                   End If
                
                
                If Digitizing And Not DigiRubberSheeting Then 'And Not RSMethod0 Then 'display pixel coordinates
                
                  'show pixel coordinates
                  GDMDIform.Text5 = Nint(XCoord)
                  GDMDIform.Text6 = Nint(Ycoord)
                
                ElseIf Digitizing And DigiRubberSheeting Then 'Or RSMethod0) Then
                
                  'convert screen coordinates to map coordinates and display
                  Dim XGeo As Double
                  Dim YGeo As Double
                  If RSMethod1 Then
                     ier = RS_pixel_to_coord2(CDbl(XCoord), CDbl(Ycoord), XGeo, YGeo)
                  ElseIf RSMethod2 Then
                     ier = RS_pixel_to_coord(CDbl(XCoord), CDbl(Ycoord), XGeo, YGeo)
                  ElseIf RSMethod0 Then
                     ier = Simple_pixel_to_coord(CDbl(XCoord), CDbl(Ycoord), XGeo, YGeo)
                     End If
                  If ier = 0 Then
                     If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                        GDMDIform.Text5 = Format(str$(XGeo), "#######.####0")
                        GDMDIform.Text6 = Format(str$(YGeo), "#######.####0")
                     ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                        GDMDIform.Text5 = Format(str$(XGeo), "#####0.0##") 'str$(CLng(XGeo))
                        GDMDIform.Text6 = Format(str$(YGeo), "######0.0##") 'str$(CLng(YGeo))
                     Else
                        GDMDIform.Text5 = Format(str$(XGeo), "#######.####0")
                        GDMDIform.Text6 = Format(str$(YGeo), "#######.####0")
                        End If
                     End If
                     
                If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
                
                    If BasisDTMheights And UseNewDTM% Then
                       'use background dtm as height reference
                       kmx = XGeo
                       kmy = YGeo
                       
                       If XGeo >= xLL And YGeo >= yLL Then
                            BytePosit = 101 + 8 * Nint((XGeo - xLL) / XStepLL) + 8 * nColLL * Nint((YGeo - yLL) / YStepLL)
                            If BytePosit < 0 Then
                               VarD = 0
                            Else
                               Get #basedtm%, BytePosit, VarD
                               End If
                            
                            If VarD = blank_value Then
                               VarD = -9999
                            ElseIf VarD < -100000 Or VarD > 100000 Then
                               VarD = -9999 'flag unreadible height
                               End If
                            
                            hgt2 = VarD / (DigiConvertToMeters * MapUnits)
                       Else
                            hgt2 = -9999
                            End If
                       
                       GDMDIform.Text7.Text = Format(str$(hgt2), "######0.0#")
                       
                    Else 'use stored dtm's
                
                        If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                            'convert from ITM to WGS84
                            kmx = XGeo
                            kmy = YGeo
                            Call ics2wgs84(kmy, kmx, lt2, lg2)
                        Else
                            lg2 = XGeo
                            lt2 = YGeo
                            End If
                            
                         If DTMtype = 1 Then
                            'use ASTER
                            Call ASTERheight(lg2, lt2, hgt2)
                         ElseIf DTMtype = 2 Then
                            'use JKH's DTM if ITM coordinates, else use NED, SRTM
                            If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                               Call DTMheight2(lg2, lt2, hgt2)
                            Else
                               Call worldheights(lg2, lt2, hgt2)
                               End If
                            End If
                            
                         GDMDIform.Text7.Text = Format(str$(hgt2 / MapUnits), "######0.0#")
                         
                         End If
                   
                   End If
                
                ElseIf Not Digitizing Then
                    
                    'Display the ITM coordinates of the click point
                    GDMDIform.Text5 = str(Int(ITMx))
                    GDMDIform.Text6 = str(Int(ITMy))
                       
                    End If
             
                'print prompts on statusbar
                If SearchDigi Or HeightSearch Then
                  If NumReportPnts& = 0 Then
                     GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries."
                  ElseIf GenerateContours Then
                     GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define map boundaries for generating contours."
                  Else
                     GDMDIform.StatusBar1.Panels(1) = "Move the cursor to the desired location (click to center on that point). Drag to define search boundaries. Right click on marks for details."
                     End If
                Else
                 GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
                 End If
                 
               If DigitizeOn Then
                    If Not DigitizePadVis Then
                        GDDigitizerfrm.Visible = True
                        BringWindowToTop (GDDigitizerfrm.hwnd)
                    Else
                        BringWindowToTop (GDDigitizerfrm.hwnd)
                        End If
                    End If
                 
               
               If (DigitizePoint Or DigitizeBlankPoint) And Not DigiRS And Not DigitizeExtendGrid And Not DigitizerEraser And Not DigiEditPoints Then
               
                    If Not DigitizePadVis Then
                       GDDigitizerfrm.Visible = True
                       End If

                    BringWindowToTop (GDDigitizerfrm.hwnd)
                    
                    If IsNumeric(val(GDDigitizerfrm.txtelev.Text)) And Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString And _
                       Trim$(GDDigitizerfrm.txtX) <> sEmpty And Trim$(GDDigitizerfrm.txtY) <> sEmpty Then
                      'draw x at the digitized point with the elevation
                      
                      digi_last.x = CLng(x / DigiZoom.LastZoom)
                      digi_last.Y = CLng(Y / DigiZoom.LastZoom)
                      digi_last.Z = val(GDDigitizerfrm.txtelev.Text) * MapUnits * InvElev
                      
                      gddm = GDform1.Picture2.DrawMode
                      gddw = GDform1.Picture2.DrawWidth
                      GDform1.Picture2.DrawMode = 13
                      GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
                      GDform1.Picture2.Line (digi_last.x * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)), digi_last.Y * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)))-(digi_last.x * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)), digi_last.Y * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom))), PointColor&
                      GDform1.Picture2.Line (digi_last.x * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom)), digi_last.Y * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)))-(digi_last.x * DigiZoom.LastZoom + Max(2, CInt(DigiZoom.LastZoom)), digi_last.Y * DigiZoom.LastZoom - Max(2, CInt(DigiZoom.LastZoom))), PointColor&
                      
                      'write the elevation value if zoomm >= 1
                      If CInt(DigiZoom.LastZoom) >= 1# Then
                         GDform1.Picture2.CurrentX = digi_last.x * DigiZoom.LastZoom + Max(4, CInt(DigiZoom.LastZoom))
                         GDform1.Picture2.CurrentY = digi_last.Y * DigiZoom.LastZoom
                         GDform1.Picture2.FontSize = CInt(8 * DigiZoom.LastZoom)
                         GDform1.Picture2.Font = "Ariel"
                         GDform1.Picture2.ForeColor = PointColor&
                         GDform1.Picture2.Print str$(digi_last.Z)
                         End If
                         
                      GDform1.Picture2.DrawMode = gddm
                      GDform1.Picture2.DrawWidth = gddw
                      
                      'record point
                      If numDigiPoints = 0 Then
                         ReDim DigiPoints(0)
                         DigiPoints(numDigiPoints).x = digi_last.x
                         DigiPoints(numDigiPoints).Y = digi_last.Y
                         DigiPoints(numDigiPoints).Z = digi_last.Z
                      Else
                         ReDim Preserve DigiPoints(numDigiPoints)
                         DigiPoints(numDigiPoints).x = digi_last.x
                         DigiPoints(numDigiPoints).Y = digi_last.Y
                         DigiPoints(numDigiPoints).Z = digi_last.Z
                         End If
                         
                    If digi_last.Z > MaxColorHeight Then MaxColorHeight = digi_last.Z
                    If digi_last.Z < MinColorHeight Then MinColorHeight = digi_last.Z
                         
                    If DigiEditPoints Then 'add digitized point to edit image buffer
                       ier = RecordDigiPointsImage(DigiPoints(numDigiPoints).x, DigiPoints(numDigiPoints).Y, 2)
                       End If
                         
                     numDigiPoints = numDigiPoints + 1
                    
                    If Not DigiLogFileOpened Then
'                       pos% = InStr(picnam$, ".")
'                       picext$ = Mid$(picnam$, pos% + 1, 3)
                       DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
                       Digilogfilnum% = FreeFile
                       Open DigiLogfilnam$ For Append As #Digilogfilnum%
                       DigiLogFileOpened = True
                       End If
                       
                     Write #Digilogfilnum%, digi_last.x, digi_last.Y, digi_last.Z, 2  '2 is the flag for point digitizing
                         
                  ElseIf IsNumeric(val(GDDigitizerfrm.txtelev.Text)) And Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString Then 'store coordinate value
                    digi_last.x = x / DigiZoom.LastZoom
                    digi_last.Y = Y / DigiZoom.LastZoom
                    digi_last.Z = val(GDDigitizerfrm.txtelev.Text) * MapUnits * InvElev
                    End If
                      
                  GDDigitizerfrm.txtX = CLng(x / DigiZoom.LastZoom)
                  GDDigitizerfrm.txtY = CLng(Y / DigiZoom.LastZoom)
                     
                  If DigitizeBlankPoint Then GDDigitizerfrm.txtelev = vbNullString
                    
               ElseIf DigitizeLine And Not DigitizeExtendGrid And Not DigiRS And Not DigitizerEraser Then
                  
                    If Not DigitizePadVis Then GDDigitizerfrm.Visible = True
                    
                    BringWindowToTop (GDDigitizerfrm.hwnd)
                    
                    If IsNumeric(val(GDDigitizerfrm.txtelev.Text)) And Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString Then
                        digi_last.x = CLng(x) '/ DigiZoom.LastZoom
                        digi_last.Y = CLng(Y) '/ DigiZoom.LastZoom
                        digi_last.Z = val(GDDigitizerfrm.txtelev.Text) * MapUnits * InvElev
                       End If
                       
                    digi_last.x = CLng(x) '/ DigiZoom.LastZoom
                    digi_last.Y = CLng(Y) '/ DigiZoom.LastZoom
                    
                    GDDigitizerfrm.txtX = x '/ DigiZoom.LastZoom
                    GDDigitizerfrm.txtY = Y '/ DigiZoom.LastZoom
                        
                    If DigitizeBeginLine And IsNumeric(val(GDDigitizerfrm.txtelev.Text)) And Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString Then
                    
                        DigitizeBeginLine = False
                        'initialize beginning coordinates of line
                        digi_begin.x = CLng(x) '/ DigiZoom.LastZoom
                        digi_begin.Y = CLng(Y) '/ DigiZoom.LastZoom
                        digi_begin.Z = val(GDDigitizerfrm.txtelev.Text) * MapUnits * InvElev
                        digi_last.x = INIT_VALUE
                        digi_last.Y = INIT_VALUE
                        
                        
                     ElseIf Not DigitizeBeginLine And IsNumeric(val(GDDigitizerfrm.txtelev.Text)) And Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString Then
                        'draw last line
                        
                        'here is where to record subsequent vertices

                        gddm = GDform1.Picture2.DrawMode
                        gddw = GDform1.Picture2.DrawWidth

                        If digi_last.x <> INIT_VALUE And digi_last.Y <> INIT_VALUE And digi_begin.x <> INIT_VALUE And digi_begin.Y <> INIT_VALUE Then
                            GDform1.Picture2.DrawMode = 13
                            GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
'                            GDform1.Picture2.Line (digi_last.X * DigiZoom.LastZoom, digi_last.Y * DigiZoom.LastZoom)-(digi_begin.X * DigiZoom.LastZoom, digi_begin.Y * DigiZoom.LastZoom), QBColor(12)
                            
                            If LineElevColors& = 1 And numcpt > 0 And MaxColorHeight >= 0 And MinColorHeight <> MaxColorHeight Then
                               'determine color
                               colornum% = ((DigiLines(0, i&).Z - MinColorHeight) / (MaxColorHeight - MinColorHeight)) * UBound(cpt, 2) + 1
                               color_line = RGB(cpt(1, colornum% - 1), cpt(2, colornum% - 1), cpt(3, colornum% - 1))
                               GDform1.Picture2.Line (digi_last.x, digi_last.Y)-(digi_begin.x, digi_begin.Y), color_line
                            Else
                               GDform1.Picture2.Line (digi_last.x, digi_last.Y)-(digi_begin.x, digi_begin.Y), LineColor& 'QBColor(12)
                               End If
                               
                            End If

                        GDform1.Picture2.DrawMode = gddm
                        GDform1.Picture2.DrawWidth = gddw
                        
                        If numDigiLines = 0 Then
                           ReDim DigiLines(1, 0)
                           DigiLines(0, 0).x = CLng(digi_begin.x / DigiZoom.LastZoom)
                           DigiLines(0, 0).Y = CLng(digi_begin.Y / DigiZoom.LastZoom)
                           DigiLines(0, 0).Z = digi_begin.Z
                           DigiLines(1, 0).x = CLng(digi_last.x / DigiZoom.LastZoom)
                           DigiLines(1, 0).Y = CLng(digi_last.Y / DigiZoom.LastZoom)
                           DigiLines(1, 0).Z = digi_last.Z
                        Else
                           ReDim Preserve DigiLines(1, numDigiLines)
                           DigiLines(0, numDigiLines).x = CLng(digi_begin.x / DigiZoom.LastZoom)
                           DigiLines(0, numDigiLines).Y = CLng(digi_begin.Y / DigiZoom.LastZoom)
                           DigiLines(0, numDigiLines).Z = digi_begin.Z
                           DigiLines(1, numDigiLines).x = CLng(digi_last.x / DigiZoom.LastZoom)
                           DigiLines(1, numDigiLines).Y = CLng(digi_last.Y / DigiZoom.LastZoom)
                           DigiLines(1, numDigiLines).Z = digi_last.Z
                           End If
                           
                        If digi_last.Z > MaxColorHeight Then MaxColorHeight = digi_last.Z
                        If digi_last.Z < MinColorHeight Then MinColorHeight = digi_last.Z
                         
                        If DigiEditPoints Then 'add digitized line to edit image buffer
                           ier = RecordDigiPointsImage(DigiLines(0, numDigiLines).x, DigiLines(0, numDigiLines).Y, 3)
                           ier = RecordDigiPointsImage(DigiLines(1, numDigiLines).x, DigiLines(1, numDigiLines).Y, 4)
                           End If
                           
                        numDigiLines = numDigiLines + 1
                        
                        If Not DigiLogFileOpened Then
'                           pos% = InStr(picnam$, ".")
'                           picext$ = Mid$(picnam$, pos% + 1, 3)
                           DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
                           Digilogfilnum% = FreeFile
                           Open DigiLogfilnam$ For Append As #Digilogfilnum%
                           DigiLogFileOpened = True
                           End If
                           
                         Write #Digilogfilnum%, DigiLines(0, numDigiLines - 1).x, DigiLines(0, numDigiLines - 1).Y, DigiLines(0, numDigiLines - 1).Z, 3 '3 is the flag for begninning of a line segment of line digitizing
                         Write #Digilogfilnum%, DigiLines(1, numDigiLines - 1).x, DigiLines(1, numDigiLines - 1).Y, DigiLines(1, numDigiLines - 1).Z, 4 '4 is the flag for end of a line segment of line digitizing
                        
                        End If
                        
               ElseIf DigitizeExtendGrid And Not DigitizePoint And Not DigitizeBlankPoint And Not DigitizeLine And Not DigitizeContour Then 'And Not DigiRS Then
                      'define beginning of line
                    digiextendgrid_last.x = x 'CLng(X / DigiZoom.LastZoom)
                    digiextendgrid_last.Y = Y 'CLng(Y / DigiZoom.LastZoom)
                        
                    If DigiExtendFirstPoint Then
                    
                        DigiExtendFirstPoint = False
                        'initialize beginning coordinates of line
                        digiextendgrid_begin.x = x 'CLng(X / DigiZoom.LastZoom)
                        digiextendgrid_begin.Y = Y 'CLng(Y / DigiZoom.LastZoom)
                        digiextendgrid_last.x = INIT_VALUE
                        digiextendgrid_last.Y = INIT_VALUE
                        
                        
                     ElseIf Not DigiExtendFirstPoint Then
                        'draw last line
                        
                        'here is where to record subsequent vertices

                        gddm = GDform1.Picture2.DrawMode
                        gddw = GDform1.Picture2.DrawWidth

                        If digiextendgrid_last.x <> INIT_VALUE And digiextendgrid_last.Y <> INIT_VALUE And digiextendgrid_begin.x <> INIT_VALUE And digiextendgrid_begin.Y <> INIT_VALUE Then
                            'remove last guide line
                            GDform1.Picture2.DrawMode = 7 'erase mode
                            GDform1.Picture2.DrawWidth = 1
                            GDform1.Picture2.Line (digiextendgrid_last.x, digiextendgrid_last.Y)-(digiextendgrid_begin.x, digiextendgrid_begin.Y), QBColor(12)
                            
                            'redraw the line, extending to the end of the map in the direction of the last extended point
                            GDform1.Picture2.DrawMode = 13
                            GDform1.Picture2.DrawWidth = 1
                            
                            'now finish the line to the end of the map
                            If (Abs(digiextendgrid_last.x - digiextendgrid_begin.x)) > (Abs(digiextendgrid_last.Y - digiextendgrid_begin.Y)) And digiextendgrid_last.x <> digiextendgrid_begin.x Then
                                Slope = (digiextendgrid_last.Y - digiextendgrid_begin.Y) / (digiextendgrid_last.x - digiextendgrid_begin.x)
                                If (digiextendgrid_last.x > digiextendgrid_begin.x) Then
                                   digiextendgrid_last.x = pixwi * DigiZoom.LastZoom
                                   digiextendgrid_last.Y = (digiextendgrid_last.x - digiextendgrid_begin.x) * Slope + digiextendgrid_begin.Y
                                   GDform1.Picture2.Line (digiextendgrid_last.x, digiextendgrid_last.Y)-(digiextendgrid_begin.x, digiextendgrid_begin.Y), QBColor(12)
                                ElseIf (digiextendgrid_last.x < digiextendgrid_begin.x) Then
                                   digiextendgrid_last.x = 0
                                   digiextendgrid_last.Y = (digiextendgrid_last.x - digiextendgrid_begin.x) * Slope + digiextendgrid_begin.Y
                                   GDform1.Picture2.Line (digiextendgrid_last.x, digiextendgrid_last.Y)-(digiextendgrid_begin.x, digiextendgrid_begin.Y), QBColor(12)
                                   End If
                             ElseIf (Abs(digiextendgrid_last.Y - digiextendgrid_begin.Y)) > (Abs(digiextendgrid_last.x - digiextendgrid_begin.x)) And digiextendgrid_last.Y <> digiextendgrid_begin.Y Then 'vertical lines
                                Slope = (digiextendgrid_last.x - digiextendgrid_begin.x) / (digiextendgrid_last.Y - digiextendgrid_begin.Y)
                                If (digiextendgrid_last.Y > digiextendgrid_begin.Y) Then
                                   digiextendgrid_last.Y = pixhi * DigiZoom.LastZoom
                                   digiextendgrid_last.x = (digiextendgrid_last.Y - digiextendgrid_begin.Y) * Slope + digiextendgrid_begin.x
                                   GDform1.Picture2.Line (digiextendgrid_last.x, digiextendgrid_last.Y)-(digiextendgrid_begin.x, digiextendgrid_begin.Y), QBColor(12)
                                ElseIf (digiextendgrid_last.Y < digiextendgrid_begin.Y) Then
                                   digiextendgrid_last.Y = 0
                                   digiextendgrid_last.x = (digiextendgrid_last.Y - digiextendgrid_begin.Y) * Slope + digiextendgrid_begin.x
                                   GDform1.Picture2.Line (digiextendgrid_last.x, digiextendgrid_last.Y)-(digiextendgrid_begin.x, digiextendgrid_begin.Y), QBColor(12)
                                   End If
                             End If
                            
                            'now save the extended lines to be recalled when the Rubber Sheeting button is pushed
'                            pos% = InStr(picnam$, ".")
'                            picext$ = Mid$(picnam$, pos% + 1, 3)
                            GuideLineFilname$ = App.Path & "\" & RootName(picnam$) & "-RSG" & ".txt"
                            filnum% = FreeFile
                            Open GuideLineFilname$ For Append As #filnum%
                            Write #filnum%, CLng(digiextendgrid_begin.x / DigiZoom.LastZoom), CLng(digiextendgrid_begin.Y / DigiZoom.LastZoom), 1 '1 is the flag for begninning of a line segment
                            Write #filnum%, CLng(digiextendgrid_last.x / DigiZoom.LastZoom), CLng(digiextendgrid_last.Y / DigiZoom.LastZoom), 2 '2 is the flag for end of a line segment
                            Close #filnum%
                            
                            DigiExtendFirstPoint = True
                            digiextendgrid_last.x = INIT_VALUE
                            digiextendgrid_last.Y = INIT_VALUE
                            digiextendgrid_begin.x = INIT_VALUE
                            digiextendgrid_begin.Y = INIT_VALUE
                            End If

                        GDform1.Picture2.DrawMode = gddm
                        GDform1.Picture2.DrawWidth = gddw
                        
                        End If
                        
                  ElseIf DigitizeContour And DigitizeMagvis And Not DigitizeExtendGrid And Not PointStart And Not DigiRS And Not DigitizeLine Then
                  
                     If Not DigitizePadVis Then
                        GDDigitizerfrm.Visible = True
                        End If
                                   
                     BringWindowToTop (GDDigitizerfrm.hwnd)
                     
                     GDDigitizerfrm.txtX = CLng(blink_mark.x / DigiZoom.LastZoom)
                     GDDigitizerfrm.txtY = CLng(blink_mark.Y / DigiZoom.LastZoom)
                  
                     If IsNumeric(val(GDDigitizerfrm.txtelev.Text)) And Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString Then
                     
                          ContourHeight = val(GDDigitizerfrm.txtelev.Text) * MapUnits
                          If Not DigitizeOn Then Unload GDDigitizerfrm

                          'convert to screen coordinates
                          Dim R As Long

                          Start_Point.x = nearmouse_digi.x
                          Start_Point.Y = nearmouse_digi.Y
                          
                          R = GDform1.Picture2.Point(Start_Point.x, Start_Point.Y)
                          If R <> -1 Then
                             Start_Color = recupcouleur(R)
                             GDMDIform.StatusBar1.Panels(2).Text = Start_Color.R & "," & Start_Color.V & "," & Start_Color.b
                          Else
                             MsgBox "Contour color couldn't be determined, try again"
                             MouseUp = -1
                             Exit Function
                             End If
                          
                         PointStart = True
                      
                         Dim Scroll As Long
                         Dim trait As Boolean
                         
                         Scroll = GDMDIform.SliderContour.value 'replace with slider
                         trait = False
                         If ChainCodeMethod = 0 Then '8 directional Freeman chain method
                            Call tracecontours8(GDform1.Picture2, Scroll, trait)
                         ElseIf ChainCodeMethod = 1 Then '4 directional Freeman chain method
                            Call tracecontours7(GDform1.Picture2, Scroll, trait)
                            End If
                         
                         End If
                         
'                  ElseIf DigiEditPoints And Not DigiRS And Not DigitizeExtendGrid And Not DigitizePoint And Not DigitizeLine And Not DigitizeContour And Not DigitizerEraser Then
'
'                     'shift the point and replot
'                     If XpixLast <> -1 And YpixLast <> -1 Then
'                        ier = RedrawDigiPoints(CLng(Xcoord), CLng(Ycoord), DigiEditMode, 0)
'                        End If
                  
                  End If
                      
               If DigiRS And Not DigitizeExtendGrid And Not DigitizePoint And Not DigitizeBlankPoint And Not DigitizeLine And Not DigitizeContour And Not DigitizerEraser Then
                     GDRSfrm.Visible = True
                     BringWindowToTop (GDRSfrm.hwnd)
                     GDRSfrm.txtX = CLng(nearmouse_digi.x / DigiZoom.LastZoom)
                     GDRSfrm.txtY = CLng(nearmouse_digi.Y / DigiZoom.LastZoom)
                     
                  End If
               
               
            End If
            
         Case 2 'right button--show record information of
                'search result nearest to right click point
                
             If PicSum = True And NumReportPnts& <> 0 Then
                
               'print prompts on statusbar
                GDMDIform.StatusBar1.Panels(1) = "Pan the map to the desired location (by draging, clicking, or using the scroll bars).  (Use mouse wheel or press ''z'' to zoom out, and ''x'' to zoom back.)"
                
                Screen.MousePointer = vbHourglass
                'freeze the GeoMap blinker
                If GeoMap And ce& = 1 Then
                   GDMDIform.CenterPointTimer.Enabled = False
                   End If
                
                'shift map to this point
                Call ShiftMap(x, Y)
                               
                'make list reappear at the record that has the
                'closest coordinate to the clicked point
                XCoord = x / (twipsx * DigiZoom.LastZoom)
                Ycoord = Y / (twipsy * DigiZoom.LastZoom)
                
                If Not Digitizing Then
                    'Convert pixel coordinates to ITM (or to any coord system)
                    ITMx = ((LRGeoX - ULGeoX) / pixwi) * XCoord + ULGeoX
                    ITMy = ULGeoY - ((ULGeoY - LRGeoY) / pixhi) * Ycoord
                    
                    'Display the ITM coordinates of the click point
                    GDMDIform.Text5 = str(Int(ITMx))
                    GDMDIform.Text6 = str(Int(ITMy))
                    
                    'Display height at click point
                    If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
                    
                        If BasisDTMheights And UseNewDTM% Then
                           'use background dtm as height reference
                           kmx = XGeo
                           kmy = YGeo
                           
                           BytePosit = 101 + 8 * Nint((XGeo - xLL) / XStepLL) + 8 * nColLL * Nint((kmy - yLL) / YStepLL)
                           Get #basedtm%, BytePosit, VarD
                           
                           If VarD = blank_value Then
                              VarD = -9999
                           ElseIf VarD < -100000 Or VarD > 100000 Then
                              VarD = -9999 'flag unreadible height
                           Else
                              hgt2 = VarD / (DigiConvertToMeters * MapUnits)
                              End If
                        
                           GDMDIform.Text3.Text = Format(str$(hgt2), "######0.0#")
                           
                        Else 'use stored dtm's
                             If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
                                'convert from ITM to WGS84
                                kmx = GDMDIform.Text5
                                kmy = GDMDIform.Text6
                                Call ics2wgs84(kmy, kmx, lt2, lg2)
                            Else
                                lg2 = GDMDIform.Text5
                                lt2 = GDMDIform.Text6
                                End If
                                
                             If DTMtype = 1 Then
                                'use ASTER
                                Call ASTERheight(lg2, lt2, hgt2)
                             ElseIf DTMtype = 2 Then
                                'use JKH's DTM if ITM coordinates, else use NED, SRTM
                                If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                                   Call DTMheight2(lg2, lt2, hgt2)
                                Else
                                   Call worldheights(lg2, lt2, hgt2)
                                   End If
                                End If
                             
                             GDMDIform.Text3.Text = Format(str$(hgt2 / MapUnits), "######0.0#")
                             
                             End If
                       
                       End If
                       
                    'write position file to hard disk
                    Call UpdatePositionFile(ITMx * DigiZoom.LastZoom, ITMy * DigiZoom.LastZoom, hgt)
                       
                Else
                   GDMDIform.Text5 = CLng(x / DigiZoom.LastZoom)
                   GDMDIform.Text6 = CLng(Y / DigiZoom.LastZoom)
                   
                   'write position file to hard disk
                   Call UpdatePositionFile(XCoord, Ycoord, 0)
                   
                   End If
                
                'now search for closest highlighted record to this point
                NearestPnt& = -1
                SelectedPnt& = 0
                DetailRecordNum& = 0
                For i& = 1 To numReport&
                    If GDReportfrm.lvwReport.ListItems(i&).Selected Then
                       SelectedPnt& = SelectedPnt& + 1
                       XPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(1))
                       YPnt = val(GDReportfrm.lvwReport.ListItems(i&).SubItems(2))
                       Dist = Sqr((ITMx - XPnt) ^ 2 + (ITMy - YPnt) ^ 2)
                       If NearestPnt& = -1 Then
                          NearestPnt& = i&
                          NearestDist = Dist
                       Else
                          If Dist < NearestDist Then
                             NearestPnt& = i&
                             NearestDist = Dist
                          End If
                       End If
                    End If
                Next i&
                
'                If NearestPnt& > 0 Then 'found the nearest search result to this point
'                    'show detailed report of this search result record
'                    ShowDetailedReport
'
'                    GDDetailReportfrm.Visible = True
'                    ret = BringWindowToTop(GDDetailReportfrm.hWnd)
'                    End If
                
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


             ElseIf Digitizing Or DigitizeOn Then
             
                PopupMenu GDMDIform.mnuDigitize
             
             Else
                beep
                'refresh center timer
                If ce& = 1 Then
                   GDMDIform.CenterPointTimer.Enabled = True
                   End If
             End If

         Case Else
       End Select
    End If
       
    MouseUp = 0
                
    Exit Function
    
errhand:
   Select Case Err.Number
      Case 52
         'problem with base dtm's file number
         'close it and reopen
         ier = OpenCloseBaseDTM(0)
         Resume
      Case 63
         'bad record number caused by being off the map sheet
         'return the blank height value
         VarD = blank_value
         Resume Next
      Case 11
         'division by zero - missing parameter, exit gracefully
         MouseUp = -1
      
      Case 480 'autoredraw error--not enough memory--ignore
         If IgnoreAutoRedrawError% = 0 Then
            MsgBox "The pixel size of this map is too big for your memory!" & vbLf & vbLf & _
                   "If you wish to use this map and ignore such errors," & vbLf & _
                   "then check the ""Ignore AutoRedraw errors"" in the" & vbLf & _
                   """Settings"" tab of ""Path/Options"" form.", vbExclamation + vbOKOnly, "MapDigitizer"
         Else 'ignore this error
            Resume Next
            End If
            
      Case 54 'something wrong with the digilog file
        If Not DigiLogFileOpened Then
'           pos% = InStr(picnam$, ".")
'           picext$ = Mid$(picnam$, pos% + 1, 3)
           DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
           Digilogfilnum% = FreeFile
           Open DigiLogfilnam$ For Append As #Digilogfilnum%
           DigiLogFileOpened = True
        Else
           If Digilogfilnum% > 0 Then Close #Digilogfilnum%
           DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
           Digilogfilnum% = FreeFile
           Open DigiLogfilnam$ For Append As #Digilogfilnum%
           DigiLogFileOpened = True
           End If
        Resume
        
      Case Else
         MsgBox "Encountered error #: " & Err.Number & vbLf & _
             Err.Description & vbLf & _
             "in module: Gdform1.Picture2_MouseUp", _
             vbCritical + vbOKOnly, "MapDigitizer"
   End Select
   
   MouseUp = -1

End Function
'mousemouse of Gdform1.picture2
Public Function MouseMove(Button As Integer, Step As Integer, x As Single, Y As Single) As Integer
  'As cursor moves over map, display readout of coordinates.
  'Also detect drag.
  Dim VarD As Double
  Dim BytePosit As Long
  
  Dim Byte0 As Byte
  Dim Byte1 As Byte
  Dim Byte2 As Byte
  Dim Byte3 As Byte
  Dim Byte4 As Byte
  Dim Byte5 As Byte
  
  Byte0 = 0
  Byte1 = 1
  Byte2 = 2
  Byte3 = 3
  Byte4 = 4
  Byte5 = 5
  
  On Error GoTo errhand
  
  nearmouse_digi.x = x
  nearmouse_digi.Y = Y
  
  GDMDIform.StatusBar1.Panels(2).Text = "Xpix: " & CLng(nearmouse_digi.x / DigiZoom.LastZoom) & "  Ypix: " & CLng(nearmouse_digi.Y / DigiZoom.LastZoom)
  GDMDIform.StatusBar1.Panels(3).Text = CInt(100 * DigiZoom.LastZoom) & "%"
  
  Dim next_mouse As POINTAPI
  Dim hDnext As Long
  Dim R As Long
  Dim Next_Color As couleur
  Dim ByteTest As Byte
  
  'variables used for heights
  Dim lt2 As Double, lg2 As Double, hgt2 As Integer
  Dim kmx As Long, kmy As Long
  
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
    If SearchDigi Or HeightSearch Or DTMcreating Or GenerateContours Or DigitizeOn Or DigitizeExtendGrid Or DigitizerEraser Or Digitizing Or DigitizerSweep Or DigiEditPoints Or Belgier_Smoothing Then
        'check if cursor left the picture frame, if so move scroll bars
        'to allow for dragging over entire map
        '(move the picture by the smallest increment = 1)
        If x / twipsx < Picture1.left + HScroll1.value + 10 Then
           'scroll map to right
           
           'slow things down a bit to give time for repainting
           waitime = Timer
           Do Until Timer > waitime + 0.0001
              DoEvents
           Loop
              
           If HScroll1.value - 1 >= HScroll1.min Then
              HScroll1.value = HScroll1.value - 1
              End If
        ElseIf x / twipsx > Picture1.Width + Picture1.left + HScroll1.value - 10 Then
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
        If Y / twipsy < Picture1.top + VScroll1.value + 10 Then
           'scroll map down
           
           'slow things down a bit to give time for repainting
           waitime = Timer
           Do Until Timer > waitime + 0.0001
              DoEvents
           Loop
              
           If VScroll1.value - 1 >= VScroll1.min Then
              VScroll1.value = VScroll1.value - 1
              End If
        ElseIf Y / twipsy > Picture1.top + Picture1.Height + VScroll1.value - 10 Then
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
       If DigitizeLine And DigitizePadVis And Not DTMcreating And Not DigitizerEraser And Not DigitizeExtendGrid And Not DigitizerSweep And Not SearchDigi And Not HeightSearch And Not GenerateContours And Not Belgier_Smoothing Then
          If Trim(GDDigitizerfrm.txtelev.Text) <> vbNullString And Not DigitizeBeginLine Then
            'erase last line
            
'            If GDMDIform.CenterPointTimer.Enabled = True Then
'               ce& = 1 'flag that timer has been shut down during drag
'               GDMDIform.CenterPointTimer.Enabled = False
'               End If
            
            gddm = GDform1.Picture2.DrawMode
            gddw = GDform1.Picture2.DrawWidth
            
            If digi_last.x <> INIT_VALUE And digi_last.Y <> INIT_VALUE And digi_begin.x <> INIT_VALUE And digi_begin.Y <> INIT_VALUE Then
               'erase last line
                GDform1.Picture2.DrawMode = 7 'erase mode
                GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
                
                digi_begin.x = blink_mark.x 'CLng(blink_mark.X / DigiZoom.LastZoom)
                digi_begin.Y = blink_mark.Y 'CLng(blink_mark.Y / DigiZoom.LastZoom)
                digi_begin.Z = val(GDDigitizerfrm.txtelev.Text) * MapUnits * InvElev
                
                If newblit = False Then
'                    GDform1.Picture2.Line (new_digi.X * DigiZoom.LastZoom, new_digi.Y * DigiZoom.LastZoom)-(digi_begin.X * DigiZoom.LastZoom, digi_begin.Y * DigiZoom.LastZoom), QBColor(12)
                    GDform1.Picture2.Line (new_digi.x, new_digi.Y)-(digi_begin.x, digi_begin.Y), QBColor(12)
                    End If
'                DigitizeDrawLine = False
                newblit = False
'                GDform1.Picture2.Line (X, Y)-(digi_begin.X * DigiZoom.LastZoom, digi_begin.Y * DigiZoom.LastZoom), QBColor(12)
                GDform1.Picture2.Line (x, Y)-(digi_begin.x, digi_begin.Y), QBColor(12)
                new_digi.x = x 'CLng(X / DigiZoom.LastZoom)
                new_digi.Y = Y 'CLng(Y / DigiZoom.LastZoom)
                new_digi.Z = val(GDDigitizerfrm.txtelev.Text) * MapUnits * InvElev
                End If
                
            'record new position of end of line
            digi_last.x = x 'CLng(X / DigiZoom.LastZoom)
            digi_last.Y = Y 'CLng(Y / DigiZoom.LastZoom)
            digi_last.Z = val(GDDigitizerfrm.txtelev.Text) * MapUnits * InvElev
            
'            If digi_last.X <> init_value And digi_last.Y <> init_value And digi_begin.X <> init_value And digi_begin.Y <> init_value Then
                'now draw new line
                'GDform1.Picture2.DrawMode = 13 '7 '13 'drawing mode
                'GDform1.Picture2.Line (digi_last.X, digi_last.Y)-(digi_begin.X, digi_begin.Y), QBColor(12)
'                DigitizeDrawLine = True
'                End If
                
            GDform1.Picture2.DrawMode = gddm
            GDform1.Picture2.DrawWidth = gddw
            End If
            
      ElseIf DigitizeExtendGrid And Not DTMcreating And Not SearchDigi And Not DigiExtendFirstPoint And Not DigitizeLine And Not DigitizePoint And Not DigitizeBlankPoint And Not DigitizeContour And Not DigitizerEraser And Not HeightSearch And Not GenerateContours And Not Belgier_Smoothing Then  'And Not DigiRS And Not DigitizerEraser Then
            gddm = GDform1.Picture2.DrawMode
            gddw = GDform1.Picture2.DrawWidth
            
            If digiextendgrid_last.x <> INIT_VALUE And digiextendgrid_last.Y <> INIT_VALUE And digiextendgrid_begin.x <> INIT_VALUE And digiextendgrid_begin.Y <> INIT_VALUE Then
               'erase last line
                GDform1.Picture2.DrawMode = 7 'erase mode
                GDform1.Picture2.DrawWidth = 1
                
                digiextendgrid_begin.x = blink_mark.x 'CLng(blink_mark.X / DigiZoom.LastZoom)
                digiextendgrid_begin.Y = blink_mark.Y 'CLng(blink_mark.Y / DigiZoom.LastZoom)
                
                If newblit = False Then
'                    GDform1.Picture2.Line (new_digi.X * DigiZoom.LastZoom, new_digi.Y * DigiZoom.LastZoom)-(digiextendgrid_begin.X * DigiZoom.LastZoom, digiextendgrid_begin.Y * DigiZoom.LastZoom), QBColor(12)
                    GDform1.Picture2.Line (new_digi.x, new_digi.Y)-(digiextendgrid_begin.x, digiextendgrid_begin.Y), QBColor(12)
                    End If
'                DigitizeDrawLine = False
                newblit = False
'                GDform1.Picture2.Line (X, Y)-(digiextendgrid_begin.X * DigiZoom.LastZoom, digiextendgrid_begin.Y * DigiZoom.LastZoom), QBColor(12)
                GDform1.Picture2.Line (x, Y)-(digiextendgrid_begin.x, digiextendgrid_begin.Y), QBColor(12)
                new_digi.x = x 'CLng(X / DigiZoom.LastZoom)
                new_digi.Y = Y 'CLng(Y / DigiZoom.LastZoom)
                End If
                
            'record new position of end of line
            digiextendgrid_last.x = x 'CLng(X / DigiZoom.LastZoom)
            digiextendgrid_last.Y = Y 'CLng(Y / DigiZoom.LastZoom)
            
            GDform1.Picture2.DrawMode = gddm
            GDform1.Picture2.DrawWidth = gddw
            
      ElseIf DigitizerEraser And Not DigitizeExtendGrid And Not DigitizeLine And Not DigitizerSweep And Not Belgier_Smoothing And Button = 1 Then
         'restore rgb color to places the eraser is swiping over
        
          'on définit la couleur du pixel courant à partir des pixels alentours
          Dim iBleu As Byte 'stocke la composante bleue à récupèrer
          Dim iVert As Byte 'stocke la composante verte à récupèrer
          Dim iRouge As Byte 'stocke la composante rouge à récupèrer
          
          DigiEraseBrushSize = Max(MinDigiEraserBrushSize, CInt(DigiZoom.LastZoom * 0.5))
        
        If Not DigiLogFileOpened Then
'           pos% = InStr(picnam$, ".")
'           picext$ = Mid$(picnam$, pos% + 1, 3)
           DigiLogfilnam$ = App.Path & "\" & RootName(picnam$) & "-DIG" & ".txt"
           Digilogfilnum% = FreeFile
           Open DigiLogfilnam$ For Append As #Digilogfilnum%
           DigiLogFileOpened = True
           End If

             
          GDform1.DrawWidth = 1
          For i% = -DigiEraseBrushSize To DigiEraseBrushSize Step 1
             For j% = -DigiEraseBrushSize To DigiEraseBrushSize Step 1
             
                 'retrieve original RGB color
                 If Not DigiGDIfailed Then
                    ier = oGestionImageSrc.GetPixelRGB(CLng(x / DigiZoom.LastZoom) + i%, CLng(Y / DigiZoom.LastZoom) + j%, iRouge, iVert, iBleu)
                 Else
                    ier = GetSimplePixelRGB(GDform1.Picture2, CLng(x / DigiZoom.LastZoom) + i%, CLng(Y / DigiZoom.LastZoom) + j%, iRouge, iVert, iBleu)
                    End If
                 
                 If ier = 0 Then
                    'restore original RBG color within square area defined by brush size
                    GDform1.Picture2.PSet (x + i% * DigiZoom.LastZoom, Y + j% * DigiZoom.LastZoom), RGB(Int(iRouge), Int(iVert), Int(iBleu))
                    
                    'record these values in a buffer and in a log file
                    If numDigiErase = 0 Then
                       ReDim DigiErasePoints(0)
                    Else
                       ReDim Preserve DigiErasePoints(numDigiErase)
                       End If
                       
                    DigiErasePoints(numDigiErase).x = CLng(x / DigiZoom.LastZoom) + i%
                    DigiErasePoints(numDigiErase).Y = CLng(Y / DigiZoom.LastZoom) + j%
                    
                    If DigiEditPoints Then 'record the erase in the edit point buffer
                       ier = RecordDigiPointsImage(CLng(DigiErasePoints(numDigiErase).x), CLng(DigiErasePoints(numDigiErase).Y), 5)
                       End If
                    
                    numDigiErase = numDigiErase + 1
                    
                    'record in written buffer also
                    Write #Digilogfilnum%, CLng(x / DigiZoom.LastZoom) + i%, CLng(Y / DigiZoom.LastZoom) + j%, 0, 5 '5 is the flag for erasing
                    End If
                 
             Next j%
          Next i%
          End If
          
      
       End If
   '<<<<<<<<<<<<<<<<<<<<<<<end of new>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
      
   If dragbegin = True And Button = 1 And dragbox = True And _
     (GenerateContours Or _
     DTMcreating Or _
     SearchDigi Or _
     HeightSearch Or _
     DigitizeHardy Or _
     DigitizerSweep) And _
     Not DigitizerEraser Then 'dragging continues, draw box
      'continue dragging
      Picture2.DrawMode = 7
      Picture2.DrawStyle = vbDot
      Picture2.DrawWidth = 1
      
      'erase last drag box
      Picture2.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
      
      'check if cursor left the picture frame, if so move scroll bars
      'to allow for dragging over entire map
      '(move the picture by the smallest increment = 1)
      If x / twipsx < Picture1.left + HScroll1.value Then
         'scroll map to right
         If HScroll1.value - 1 >= HScroll1.min Then
            HScroll1.value = HScroll1.value - 1
            End If
      ElseIf x / twipsx > Picture1.Width + Picture1.left + HScroll1.value Then
         'scroll map to left
         If HScroll1.value + 1 <= HScroll1.Max Then
            HScroll1.value = HScroll1.value + 1
            End If
         End If
      If Y / twipsy < Picture1.top + VScroll1.value Then
         'scroll map down
         If VScroll1.value - 1 >= VScroll1.min Then
            VScroll1.value = VScroll1.value - 1
            End If
      ElseIf Y / twipsy > Picture1.top + Picture1.Height + VScroll1.value Then
         'scroll map up
         If VScroll1.value + 1 <= VScroll1.Max Then
            VScroll1.value = VScroll1.value + 1
            End If
         End If
      
      'draw new drag box
      Picture2.Line (x, Y)-(drag1x, drag1y), QBColor(15), B
'      GDMDIform.StatusBar1.Panels(1).Text = GDMDIform.StatusBar1.Panels(1).Text & "X,Y,drag1x,drag1y= " & str(x) + ", " & str(Y) & ", " & str(drag1x) & ", " & str(drag1y)

      Picture2.Refresh
      
      'record new drag end coordinates
      drag2x = x: drag2y = Y
      
   ElseIf dragbegin = True And Button = 1 And dragbox = False And drawbox = False And _
      (SearchDigi Or DTMcreating Or HeightSearch Or DigitizeHardy Or DigitizerSweep Or GenerateContours) And _
      Not DigitizerEraser Then
      'begin dragging for Hardy selection
      Picture2.DrawMode = 7
      Picture2.DrawStyle = vbDot
      Picture2.DrawWidth = 1
      Picture2.Line (x, Y)-(drag1x, drag1y), QBColor(15), B
      drag2x = x: drag2y = Y
      dragbox = True
      
   ElseIf dragbegin = True And Button = 1 And dragbox = False And drawbox = False And _
      Not DTMcreating And _
      Not HeightSearch And _
      Not SearchDigi And _
      Not DigitizeHardy And _
      Not DigitizerSweep And _
      Not DigitizerEraser Then
      'just drag for moving the picture
      GDform1.Picture2.MouseIcon = LoadResPicture(103, vbResCursor) 'load special drag cursor
      GDform1.Picture2.MousePointer = vbCustom
      drag2x = x: drag2y = Y
      End If
      
  'Convert coordinates to pixels
  XCoord = x / (twipsx * DigiZoom.LastZoom)
  Ycoord = Y / (twipsy * DigiZoom.LastZoom)
  
  'Convert pixel coordinates to ITM
  ITMx = ((LRGeoX - ULGeoX) / pixwi) * XCoord + ULGeoX
  ITMy = ULGeoY - ((ULGeoY - LRGeoY) / pixhi) * Ycoord
    
  'Display the ITM coordinates
  GDMDIform.Text1 = str(CLng(ITMx))
  GDMDIform.Text2 = str(CLng(ITMy))
    
  If Digitizing And Not DigiRubberSheeting Then 'And Not RSMethod0 Then 'display pixel coordinates
  
    'show pixel coordinates
    GDMDIform.Text1 = Nint(XCoord)
    GDMDIform.Text2 = Nint(Ycoord)
  
  ElseIf Digitizing And DigiRubberSheeting Then '(DigiRubberSheeting Or RSMethod0) Then
  
    'convert screen coordinates to map coordinates and display
    Dim XGeo As Double
    Dim YGeo As Double
    If RSMethod1 Then
        ier = RS_pixel_to_coord2(CDbl(XCoord), CDbl(Ycoord), XGeo, YGeo)
    ElseIf RSMethod2 Then
        ier = RS_pixel_to_coord(CDbl(XCoord), CDbl(Ycoord), XGeo, YGeo)
    ElseIf RSMethod0 Then
       ier = Simple_pixel_to_coord(CDbl(XCoord), CDbl(Ycoord), XGeo, YGeo)
       End If
    If ier = 0 Then
       If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
          GDMDIform.Text1 = Format(str$(XGeo), "#######.####0")
          GDMDIform.Text2 = Format(str$(YGeo), "#######.####0")
       ElseIf Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
          GDMDIform.Text1 = Format(str$(XGeo), "######0.0##") 'str$(CLng(XGeo))
          GDMDIform.Text2 = Format(str$(YGeo), "######0.0##") 'str$(CLng(YGeo))
       Else
          GDMDIform.Text1 = Format(str$(XGeo), "#######.####0")
          GDMDIform.Text2 = Format(str$(YGeo), "#######.####0")
          End If
       GDMDIform.StatusBar1.Panels(2) = "Xpix: " & Nint(XCoord) & ",  Ypix: " & Nint(Ycoord)
       GDMDIform.StatusBar1.Panels(3).Text = CInt(100 * DigiZoom.LastZoom) & "%"
       End If
       
     If (heights Or BasisDTMheights) And (RSMethod0 Or RSMethod1 Or RSMethod2) Then
    
       If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And Not JKHDTM Then
          'convert from ITM to WGS84
          kmx = XGeo
          kmy = YGeo
          Call ics2wgs84(kmy, kmx, lt2, lg2)
       Else
          lg2 = XGeo
          lt2 = YGeo
          End If
          
       If BasisDTMheights And UseNewDTM% Then
          'use background dtm as height reference
          kmx = XGeo
          kmy = YGeo
          
          If XGeo >= xLL And YGeo >= yLL Then
            BytePosit = 101 + 8 * Nint((XGeo - xLL) / XStepLL) + 8 * nColLL * Nint((YGeo - yLL) / YStepLL)
            Get #basedtm%, BytePosit, VarD
          
            If VarD = blank_value Then
               VarD = -9999 * (DigiConvertToMeters * MapUnits)
            ElseIf VarD < -100000 Or VarD > 100000 Then
               VarD = -9999 * (DigiConvertToMeters * MapUnits) 'flag unreadible height
            Else
               hgt2 = VarD / (DigiConvertToMeters * MapUnits)
               End If
          Else
             hgt2 = -9999
             End If
          
          GDMDIform.Text3.Text = Format(str$(hgt2), "######0.0#")
       
       Else 'use stored dtm's
          
            If DTMtype = 1 Then
               'use ASTER
               Call ASTERheight(lg2, lt2, hgt2)
            ElseIf DTMtype = 2 Then
               'use JKH's DTM if ITM coordinates, else use NED, SRTM
               If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                  kmx = lg2
                  kmy = lt2
                  Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt2)
               Else
                  Call worldheights(lg2, lt2, hgt2)
                  End If
               End If
               
               GDMDIform.Text3.Text = Format(str$(hgt2 / MapUnits), "######0.0#")
               
               End If
       
       
       If Not BasisDTMheights And JKHDTM And Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
          If val(GDMDIform.Text5.Text) >= 80000 And val(GDMDIform.Text5.Text) <= 260000 And val(GDMDIform.Text6.Text) >= 80000 And val(GDMDIform.Text6.Text) <= 11350000 Then
             lg1 = val(GDMDIform.Text5.Text)
             lt1 = val(GDMDIform.Text6.Text)
             distkm = Sqr((lt2 - lt1) ^ 2# + (lg2 - lg1) ^ 2#) * 0.001
             GDMDIform.Text4.Text = Format(str$(distkm), "###0.0###")
             End If
       ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
          If Abs(val(GDMDIform.Text6.Text)) <= 90 And Abs(val(GDMDIform.Text6.Text)) <= 180 Then
             distkm = Rearthkm * DistTrav(lt2, lg2, CDbl(GDMDIform.Text6.Text), CDbl(GDMDIform.Text5.Text), 3)
             GDMDIform.Text4.Text = Format(str$(distkm), "###0.0###")
             End If
       ElseIf BasisDTMheights And ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
          (Mid$(LCase(lblX), 1, 3) = "utm" And Mid$(LCase(LblY), 1, 3) = "utm")) Then
          lg1 = val(GDMDIform.Text5.Text)
          lt1 = val(GDMDIform.Text6.Text)
          distkm = Sqr(((lt2 - lt1) / (XStepLL * nColLL)) ^ 2# + ((lg2 - lg1) / (YStepLL * nRowLL)) ^ 2#)
          GDMDIform.Text4.Text = Format(str$(distkm), "###0.0###")
          End If
       
       End If
  
  ElseIf Not Digitizing Then 'regular MapDigitizer program coordinates
    
    'Display the ITM coordinates
    GDMDIform.Text1 = str(Int(ITMx))
    GDMDIform.Text2 = str(Int(ITMy))
    
    End If
    
  If DigiEditPoints And ImagePointFile Then
     'highlight points closest to cursor
     ier = HighLightPoint(Nint(XCoord), Nint(Ycoord), XpixLast, YpixLast, ByteTest, 0)
     
     If ier = 0 Then
        'store the type of digitized point the mouse was hovering over
        Select Case ByteTest
           Case Byte0
              DigiEditMode = 0
           Case Byte1
              DigiEditMode = 1
           Case Byte2
              DigiEditMode = 2
           Case Byte3
              DigiEditMode = 3
           Case Byte4
              DigiEditMode = 4
           Case Byte5
              DigiEditMode = 5
        End Select
        End If
     
     End If
     
'  If Geo And ShowContGeo Then 'also display geo coordinates
'        kmxoo = ITMx: kmyoo = ITMy
'
'        If GpsCorrection Then
'            Dim lat_g As Double
'            Dim lon_g As Double
'            Dim N As Long
'            Dim E As Long
'            N = CLng(kmyoo)
'            E = CLng(kmxoo)
'            Call ics2wgs84(N, E, lat_g, lon_g)
'            lt = lat_g
'            lg = -lon_g
'        Else
'            Call casgeo(kmxoo, kmyoo, lg, lt)
'            End If
'
'        If GeoDecDeg = True Then
'            GDGeoFrm.txtLat = Mid$(str$(lt), 1, 9)
'            GDGeoFrm.txtLon = Mid$(str$(lg), 1, 9)
'        Else
'            lgdeg = Fix(lg)
'            lgmin = Abs(Fix((lg - Fix(lg)) * 60))
'            lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
'            ltdeg = Fix(lt)
'            ltmin = Abs(Fix((lt - Fix(lt)) * 60))
'            ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
'            If ltdeg = 0 And lt < 0 Then
'              GDGeoFrm.txtLatDeg = "-" + str$(ltdeg) + "°"
'              GDGeoFrm.txtLatMin = str$(ltmin) + "'"
'              GDGeoFrm.txtLatSec = Mid$(str$(ltsec), 1, 6) + """"
'            Else
'              GDGeoFrm.txtLatDeg = str$(ltdeg) + "°"
'              GDGeoFrm.txtLatMin = str$(ltmin) + "'"
'              GDGeoFrm.txtLatSec = Mid$(str$(ltsec), 1, 6) + """"
'              End If
'            If lgdeg = 0 And lg < 0 Then
'              GDGeoFrm.txtLonDeg = "-" + str$(lgdeg) + "°"
'              GDGeoFrm.txtLonMin = str$(lgmin) + "'"
'              GDGeoFrm.txtLonSec = Mid$(str$(lgsec), 1, 6) + """"
'            Else
'              GDGeoFrm.txtLonDeg = str$(lgdeg) + "°"
'              GDGeoFrm.txtLonMin = str$(lgmin) + "'"
'              GDGeoFrm.txtLonSec = Mid$(str$(lgsec), 1, 6) + """"
'              End If
'            End If
'
'  End If

  MouseMove = 0
  
  If DigitizePadVis Then
     BringWindowToTop (GDDigitizerfrm.hwnd)
     End If
  
  Exit Function

errhand:
   Select Case Err.Number
      Case 52
         'problem with base dtm's file number
         'close it and reopen
         ier = OpenCloseBaseDTM(0)
         Resume
      Case 63
         'this is caused by a record number beyond the limits of the map edges
         'return a blank height
         VarD = blank_value
         Resume Next
      Case 11
         'division by zero - missing parameter, exit gracefully
         MouseMove = -1
      Case 480
         If IgnoreAutoRedrawError% = 0 Then
            MsgBox "The pixel size of this map is too big for your memory!" & vbLf & vbLf & _
                   "If you wish to use this map and ignore such errors," & vbLf & _
                   "then check the ""Ignore AutoRedraw errors"" in the" & vbLf & _
                   """Settings"" tab of ""Path/Options"" form.", vbExclamation + vbOKOnly, "MapDigitizer"
         Else 'ignore this error
            Resume Next
            End If
      Case 13
         MsgBox "Coordinate has wrong format!" & vbLf & vbLf & "Try inputing it again.", vbExclamation + vbOKOnly, "MapDigitizer"
         GDMDIform.Text5 = 0
         GDMDIform.Text6 = 0
      Case Else
         MsgBox "Encountered error #: " & Err.Number & vbLf & _
               Err.Description & vbLf & _
               "", vbCritical + vbOKOnly, "MapDigitizer"
               
               
   End Select
            
   MouseMove = -1

End Function
