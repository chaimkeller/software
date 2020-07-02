Attribute VB_Name = "modHook"
Option Explicit

' Purpose   : stretchblt with mouse wheel to zoom based on two different sources:
'             1. stretchblt method source: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=39776&lngWId=1
'             2. wheel mouse hook: http://www.vbforums.com/showthread.php?388222-VB6-MouseWheel-with-Any-Control-%28originally-just-MSFlexGrid-Scrolling%29
'               two files examples: WheelHook-AllControls.zip, WheelHook-NestedControls.zip


' Store WndProcs
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String) As Long

' Hooking
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hwnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

' Position Checking
Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hwnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetAncestor Lib "user32.dll" ( _
                ByVal hwnd As Long, _
                ByVal gaFlags As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157
Private Const GA_ROOT = 2

Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Dim RecordNum&

' Check Messages
' ================================================
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long
  Dim fFrm As Form
  Dim Ret As Long

  Select Case Lmsg
  
    Case WM_MOUSEWHEEL
    
      MouseKeys = wParam And 65535
      Rotation = wParam / 65536
      Xpos = lParam And 65535
      Ypos = lParam / 65536
      
      Set fFrm = GetForm(Lwnd)
      If fFrm Is Nothing Then
        ' it's not a form
        
        If Not IsOver(Lwnd, Xpos, Ypos) And IsOver(GetAncestor(Lwnd, GA_ROOT), Xpos, Ypos) Then
          ' it's not over the control and is over the form,
          ' so fire mousewheel on form (if it's not a dropped down combo)
          If SendMessage(Lwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
            GetForm(GetAncestor(Lwnd, GA_ROOT)).MouseWheel MouseKeys, Rotation, Xpos, Ypos
            Exit Function ' Discard scroll message to control
          End If
        End If
      Else
        ' it's a form so fire mousewheel
        If IsOver(fFrm.hwnd, Xpos, Ypos) Then fFrm.MouseWheel MouseKeys, Rotation, Xpos, Ypos
      End If
  End Select
  
  WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
  
       
'    If DigitizeMagvis Then
'       DoEvents
'       Ret = SetWindowPos(GDDigiMagfrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'       End If
  
End Function

' Hook / UnHook
' ================================================
Public Sub WheelHook(ByVal hwnd As Long)
  On Error Resume Next
  SetProp hwnd, "PrevWndProc", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook(ByVal hwnd As Long)
  On Error Resume Next
  SetWindowLong hwnd, GWL_WNDPROC, GetProp(hwnd, "PrevWndProc")
  RemoveProp hwnd, "PrevWndProc"
End Sub

' Window Checks
' ================================================
Public Function IsOver(ByVal hwnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
  Dim rectCtl As RECT
  GetWindowRect hwnd, rectCtl
  With rectCtl
    IsOver = (lX >= .left And lX <= .right And lY >= .top And lY <= .bottom)
    '--------------CNK - changes for the digitizer program--------------------------------
    If GDRSfrmVis Or DigitizePadVis Then
       'include the entire gform1.picture2 area for mouse wheel capture even though the GDRSfrm or GDDigitizerfrm forms are in focus
       IsOver = (lX >= GDform1.Picture2.ScaleLeft And lX <= GDform1.Picture2.ScaleWidth And _
                 lY >= GDform1.Picture2.ScaleTop And lY <= GDform1.Picture2.ScaleHeight)
       End If
    '-------------------------------------------------------------------------------------
  End With
End Function

Private Function GetForm(ByVal hwnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.hwnd = hwnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function

' Control Specific Behaviour
' ================================================
Public Sub FlexGridScroll(ByRef FG As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim NewValue As Long
  Dim Lstep As Single

  On Error Resume Next
  With FG
    Lstep = .Height / .RowHeight(0)
    Lstep = Int(Lstep)
    If .Rows < Lstep Then Exit Sub
    Do While Not (.RowIsVisible(.TopRow + Lstep))
      Lstep = Lstep - 1
    Loop
    If Rotation > 0 Then
        NewValue = .TopRow - Lstep
        If NewValue < 1 Then
            NewValue = 1
        End If
    Else
        NewValue = .TopRow + Lstep
        If NewValue > .Rows - 1 Then
            NewValue = .Rows - 1
        End If
    End If
    .TopRow = NewValue
  End With
End Sub

' Control Specific Behaviour
' ================================================
Public Sub ListViewScroll(ByRef FG As ListView, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)

  On Error Resume Next
  With FG
  
    If RecordNum& < 1 Then RecordNum& = 1
    If RecordNum& > .ListItems.count Then RecordNum& = .ListItems.count - 1
    
    If Rotation < 0 Then
       RecordNum& = RecordNum& + 1
       If RecordNum& < .ListItems.count Then
          .ListItems(RecordNum& - 1).EnsureVisible
       Else
          .ListItems(.ListItems.count - 1).EnsureVisible
          End If
    ElseIf Rotation > 0 Then
       RecordNum& = RecordNum& - 1
       If RecordNum& >= 1 Then
          .ListItems(RecordNum& - 1).EnsureVisible
       Else
          .ListItems(0).EnsureVisible
          End If
       End If
  End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PictureBoxZoom
' Author    : Dr-John-K-Hall
' Date      : 3/1/2015
' Purpose   : stretchblt with mouse wheel to zoom based on two different sources:
'             1. stretchblt method source: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=39776&lngWId=1
'             2. wheel mouse hook: http://www.vbforums.com/showthread.php?388222-VB6-MouseWheel-with-Any-Control-%28originally-just-MSFlexGrid-Scrolling%29
'               two files examples: WheelHook-AllControls.zip, WheelHook-NestedControls.zip
'           mode = 0 'redraw all the lines, etc.
'           mode >= 1 'don't redraw the rubber sheeting points
'           mode >= 2 'dont' redraw the guidelines
'---------------------------------------------------------------------------------------
Public Sub PictureBoxZoom(ByRef picBox As PictureBox, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long, mode As Integer)
'  picBox.Cls
'  picBox.Print "MouseWheel " & IIf(Rotation < 0, "Down", "Up")
   Dim Bltier As Long
   Dim NewLeft As Single
   Dim NewTop As Single
   Dim NewZoom As Single
   Dim ier As Integer
   Dim AA As Long, BB As Long
'   Dim XCenter As Single, YCenter As Single 'center of zoomed out screen
      
   On Error GoTo PictureBoxZoom_Error
   
   ier = 0
   
   If (DigitizeHardy Or GenerateContours) And Not DigiReDrawContours And numContourPoints > 0 Then
      DigiReDrawContours = True
   Else
      DigiReDrawContours = False
      End If
   
   'if Rotation < 0 then zoom out
   'if Rotation > 0 then zoom in
   'if Rotation = 0 then stay at current zoom

   NewZoom = Max(0.1, DigiZoom.LastZoom + 0.05 * Sgn(Rotation))
   
10:
  If ier = -1 Then 'exiting graefully from memory error
     ier = 0 'reset error flag
     NewZoom = Max(0.1, DigiZoom.LastZoom - 0.05 * Sgn(Rotation))
     GDMDIform.StatusBar1.Panels(1).Text = "You computer's memory limits the maximum zoom to: " & DigiZoom.LastZoom * 100 & "%"
     End If
     
   picBox.Cls
   AA = CLng(NewZoom * pixwi)
   BB = CLng(NewZoom * pixhi)
   picBox.Width = AA 'CLng(NewZoom * pixwi)
   picBox.Height = BB 'CLng(NewZoom * pixhi)
   
   If picBox.Width <> AA Or picBox.Height <> BB Then
      'some sort of memory bug for large pictures
      
      Call MsgBox("You have reached the zoom in limit for this map." _
                  & vbCrLf & "To magnify further, use the magnify tool." _
                  , vbInformation, "Zoom in limit")
      
      NewZoom = Max(0.1, DigiZoom.LastZoom - 0.05 * Sgn(Rotation))
      AA = CLng(NewZoom * pixwi)
      BB = CLng(NewZoom * pixhi)
      picBox.Width = AA 'CLng(NewZoom * pixwi)
      picBox.Height = BB 'CLng(NewZoom * pixhi)
      End If
   
   If AA - GDform1.Picture1.Width > 32767 Then
      Call MsgBox("Reached maximum horizontal zoom!", vbInformation, "Horizontal Zoom error")
      Exit Sub
      End If
   
   If BB - GDform1.Picture1.Height > 32767 Then
      Call MsgBox("Reached maximum vertical zoom!", vbInformation, "Vertical Zoom error")
      Exit Sub
      End If
      
    'reSet the Max property for the scroll bars.
    GDform1.HScroll1.Max = Max(0, AA - GDform1.Picture1.Width)
    GDform1.VScroll1.Max = Max(0, BB - GDform1.Picture1.Height)
        
    'Determine if the child picture will fill up the screen
    'If so, there is no need to use scroll bars.
'    GDform1.VScroll1.Visible = (GDform1.Picture1.Height < AA)
'    GDform1.HScroll1.Visible = (GDform1.Picture1.Width < BB)
    
    If GDform1.HScroll1.Max > 0 Then
       GDform1.HScroll1.Visible = (GDform1.Picture1.Width < BB)
        'Initiate Scroll Step Sizes
        If GDform1.HScroll1.Visible Then
            GDform1.HScroll1.LargeChange = Max(1, Fix(GDform1.HScroll1.Max / 20))
            GDform1.HScroll1.SmallChange = Max(1, Fix(GDform1.HScroll1.Max / 60))
            End If
        End If
        
    If GDform1.VScroll1.Max > 0 Then
       GDform1.VScroll1.Visible = (GDform1.Picture1.Height < AA)
       If GDform1.VScroll1.Visible Then
          GDform1.VScroll1.LargeChange = Max(1, Fix(GDform1.VScroll1.Max / 20))
          GDform1.VScroll1.SmallChange = Max(1, Fix(GDform1.VScroll1.Max / 60))
          End If
       End If
   
   'move the blit in order to keep the same pixels in the center
   'this is center
'   XCenter = GDform1.Picture1.Width * 0.5 'CLng(nearmouse_digi.X * picBox.Width / pixwi)
'   YCenter = GDform1.Picture1.Height * 0.5 'CLng(nearmouse_digi.Y * picBox.Height / pixhi)
   
   'move them to the center of picBox
   If DigiZoom.left = INIT_VALUE And DigiZoom.top = INIT_VALUE Then
      DigiZoom.left = blink_mark.x
      DigiZoom.top = blink_mark.Y
      End If
   NewLeft = DigiZoom.left * NewZoom 'CLng(XCenter - nearmouse_digi.X) ' * NewZoom)
   NewTop = DigiZoom.top * NewZoom 'CLng(YCenter - nearmouse_digi.Y) ' * NewZoom)
    
   ce& = 0 'reset blinker flag
   If GDMDIform.CenterPointTimer.Enabled = True Then
      ce& = 1 'flag that timer has been shut down during drag
      GDMDIform.CenterPointTimer.Enabled = False
      End If
    
'   Bltier = StretchBlt(GDform1.Picture2.hdc, NewLeft, NewTop, CLng(NewZoom * pixwi), CLng(NewZoom * pixhi), GDform1.PictureBlit.hdc, 0, 0, pixwi, pixhi, vbSrcCopy)
'   Bltier = StretchBlt(GDform1.Picture2.hdc, NewLeft, NewTop, CLng(NewZoom * pixwi), CLng(NewZoom * pixhi), GDform1.PictureBlit.hdc, 0, 0, pixwi, pixhi, vbSrcCopy)
   Bltier = StretchBlt(GDform1.Picture2.hdc, 0, 0, AA, BB, GDform1.PictureBlit.hdc, 0, 0, pixwi, pixhi, vbSrcCopy)
   DigiZoomed = True
   Call ShiftMap(NewLeft, NewTop)
'   GDform1.Refresh
   picBox.Refresh
   
   If Bltier = 0 Then
       Err.Raise vbObjectError + 50, "Gdform1.picture2", "Zoom failed..."
   Else
       'record changes
       DigiZoom.Zoom = NewZoom
       DigiZoom.LastZoom = DigiZoom.Zoom
       End If
       
    If DigitizePadVis And Not DigiRS Then 'And (DigitizeLine Or DigitizeContour Or DigitizePoint) And Not DigiRS Then
       If Not InitDigiGraph Then
          InputDigiLogFile 'load up saved digitizing data for the current map sheet
       Else
          ier = RedrawDigiLog
          End If
       GDDigitizerfrm.Visible = True
       BringWindowToTop (GDDigitizerfrm.hwnd)
       
    ElseIf Not DigitizePadVis And (DigitizeOn Or DigitizeHardy) Then
       If Not InitDigiGraph Then
          InputDigiLogFile 'load up saved digitizing data for the current map sheet
       Else
          ier = RedrawDigiLog
          End If
       
    ElseIf DigiRS Then
       If mode = 0 Then
          ier = ReadRSfile
          If ier <> 0 Then Exit Sub
          End If
       ier = InputGuideLines
       GDRSfrm.Visible = True
       BringWindowToTop (GDRSfrm.hwnd)
       
    ElseIf DigitizeExtendGrid And mode = 0 Then
       ier = InputGuideLines
       End If
       
    If DigiReDrawContours Then
       ier = ReDrawContours(GDform1.Picture2)
       End If
       
    DigiReDrawContours = False
    
    If CoordListVis Then
       'redraw the markers
       CoordListZoom = True
       Call GDCoordinateList.cmdDraw_Click
       CoordListZoom = False
       BringWindowToTop (GDCoordinateList.hwnd)
       End If
    
    If DTMcreating Then
       'redraw the merged regions

        If LRGeoX <> ULGeoX And ULGeoY <> LRGeoY Then
            Dim GeoToPixelX As Double, GeoToPixelY As Double
            Dim CurrentX As Double, CurrentY As Double
            Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
            Dim X3 As Double, Y3 As Double, X4 As Double, Y4 As Double
            Dim GeoX As Double, GeoY As Double
            Dim XGeo As Double, YGeo As Double
            Dim XDif As Double, YDif As Double
            Dim filhdr%, DTMhdrfile$, ninput%
            Dim ShiftX As Double, ShiftY As Double
            Dim Tolerance As Double
            Dim SecondOrderShift As Boolean
            Dim zminLLpz As Double, zmaxLLpz As Double
            
            SecondOrderShift = False
            Tolerance = 0.00001
            
            GeoToPixelX = (LRPixX - ULPixX) / (LRGeoX - ULGeoX)
            GeoToPixelY = (LRPixY - ULPixY) / (ULGeoY - LRGeoY)
             
             filhdr% = FreeFile
             DTMhdrfile$ = dirNewDTM & "\" & RootName(picnam$) & ".hdr"
             Open DTMhdrfile$ For Input As #filhdr%
             Input #filhdr%, nRowLL
             Input #filhdr%, nColLL
             Input #filhdr%, xLL
             Input #filhdr%, yLL
             Input #filhdr%, XStepLL
             Input #filhdr%, YStepLL
             Input #filhdr%, zminLLpz
             Input #filhdr%, zmaxLLpz
             Input #filhdr%, AngLL
             Input #filhdr%, blank_LL
             
             'now read the coordinates of the mergers and overplot them on the map as a colored box
             Dim xin(0 To 10)
             Dim gdfs, gdco, gdwi, gdds
             gdfs = GDform1.Picture2.FillStyle
             gdco = GDform1.Picture2.FillColor
             gdwi = GDform1.Picture2.DrawWidth
             gdds = GDform1.Picture2.DrawStyle
             
gdtm50:
             ninput% = 0
             Do Until EOF(filhdr%)
                Input #filhdr%, xin(ninput%)
                ninput% = ninput% + 1
                If ninput% = 10 Then Exit Do
             Loop
             
             If ninput% = 10 Then 'succcessfully read a merge region's boundaries
                'plot the merged regions
                                                 
                GeoX = xin(2)
                GeoY = xin(3)
                GoSub GeotoCoord
                X1 = CurrentX * DigiZoom.LastZoom
                Y1 = CurrentY * DigiZoom.LastZoom
                GeoX = xin(2) + xin(1) * xin(4)
                GoSub GeotoCoord
                X2 = CurrentX * DigiZoom.LastZoom
                Y2 = CurrentY * DigiZoom.LastZoom
                GeoY = xin(3) + xin(0) * xin(5)
                GoSub GeotoCoord
                X3 = CurrentX * DigiZoom.LastZoom
                Y3 = CurrentY * DigiZoom.LastZoom
                GeoX = xin(2)
                GoSub GeotoCoord
                X4 = CurrentX * DigiZoom.LastZoom
                Y4 = CurrentY * DigiZoom.LastZoom
                
                If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" And JKHDTM Then
                   'draw box
                   GDform1.Picture2.FillStyle = 4
                   GDform1.Picture2.DrawMode = 9 '8 '5 '2 '3 '5
                   GDform1.Picture2.Line (X1, Y1)-(X3, Y3), , BF
                Else
                    GDform1.Picture2.DrawMode = 13
                    GDform1.Picture2.DrawWidth = Max(5, CInt(DigiZoom.LastZoom))
                    GDform1.Picture2.DrawStyle = vbDot
                    GDform1.Picture2.Line (X1, Y1)-(X2, Y2)
                    GDform1.Picture2.Line (X2, Y2)-(X3, Y3)
                    GDform1.Picture2.Line (X3, Y3)-(X4, Y4)
                    GDform1.Picture2.Line (X4, Y4)-(X1, Y1)
                    End If
                GoTo gdtm50
             Else
                Close #filhdr%
                GDform1.Picture2.FillStyle = gdfs
                GDform1.Picture2.FillColor = gdco
                GDform1.Picture2.DrawWidth = gdwi
                GDform1.Picture2.DrawStyle = gdds
                End If
                
            Else
                Call MsgBox("Previously merged regions cannot be plotted until you define the" _
                            & vbCrLf & "pixel coordinates of the map's corners.  " _
                            & vbCrLf & "" _
                            & vbCrLf & "Use the Option menu to define them" _
                            , vbInformation Or vbDefaultButton1, "DTM merging")
            
               End If
       
       End If
        
 '-----------------------------digilines and digiextgend---------------------
    newblit = True 'flag not to erase last guide line since just repainted the map
    
    'rescale the temp line segment ends after zooming
    
    If DigitizeLine And DigitizePadVis And Not DigitizerEraser And Not DigitizeExtendGrid And Not DigitizerSweep Then
    
       digi_begin.x = digi_begin.x * DigiZoom.LastZoom
       digi_begin.Y = digi_begin.Y * DigiZoom.LastZoom
       digi_last.x = digi_last.x * DigiZoom.LastZoom
       digi_last.Y = digi_last.Y * DigiZoom.LastZoom
       new_digi.x = new_digi.x * DigiZoom.LastZoom
       new_digi.Y = new_digi.Y * DigiZoom.LastZoom
       
           
    ElseIf DigitizeExtendGrid And Not DigiExtendFirstPoint And Not DigitizeLine And Not DigitizePoint And Not DigitizeContour And Not DigiRS And Not DigitizerEraser Then
            
       digiextendgrid_begin.x = digiextendgrid_begin.x * DigiZoom.LastZoom
       digiextendgrid_begin.Y = digiextendgrid_begin.Y * DigiZoom.LastZoom
       digiextendgrid_last.x = digiextendgrid_last.x * DigiZoom.LastZoom
       digiextendgrid_last.Y = digiextendgrid_last.Y * DigiZoom.LastZoom
       new_digi.x = new_digi.x * DigiZoom.LastZoom
       new_digi.Y = new_digi.Y * DigiZoom.LastZoom
       
       End If
 '--------------------------------------------------------------------------

   On Error GoTo 0
   Exit Sub
   
GeotoCoord:
    
    CurrentX = ((GeoX - ULGeoX) * GeoToPixelX) + ULPixX
    CurrentY = ((ULGeoY - GeoY) * GeoToPixelY) + ULPixY
    
    If RSMethod1 Or RSMethod2 Then
       
       If RSMethod1 Then
          ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
       ElseIf RSMethod2 Then
          ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
          End If
          
        Dim DifX As Double, DifY As Double
        DifX = Abs(GeoX - XGeo)
        DifY = Abs(GeoY - YGeo)
       
        ShiftX = CurrentX - (((XGeo - ULGeoX) * GeoToPixelX) + ULPixX)
        ShiftY = CurrentY - (((ULGeoY - YGeo) * GeoToPixelY) + ULPixY)
        
        CurrentX = CurrentX + ShiftX
        CurrentY = CurrentY + ShiftY
        
        If RSMethod1 Then
           ier = RS_pixel_to_coord2(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
         ElseIf RSMethod2 Then
           ier = RS_pixel_to_coord(CDbl(CurrentX), CDbl(CurrentY), XGeo, YGeo)
           End If

        If Abs(GeoX - XGeo) > DifX Then
           CurrentX = CurrentX - ShiftX
           End If
           
        If Abs(GeoY - YGeo) > DifY Then
           CurrentY = CurrentY - ShiftY
           End If

'        If Abs(GeoX - XGeo) > DifX And Abs(GeoY - YGeo) > DifY Then
''        If Abs(GeoX - XGeo) > Tolerance Or Abs(GeoY - YGeo) > Tolerance Then
'                Call MsgBox("Inverse coordinate transformation unsuccessful" _
'                        & vbCrLf & "Coordinate grid rotation too large for first approx." _
'                        & vbCrLf & vbCrLf & "(Redo using a less-rotated grid as reference...)" _
'                        , vbInformation, "Picture Box Zoom Error")
'              Screen.MousePointer = vbDefault
'              GDMDIform.picProgBar.Visible = False
'              GDMDIform.StatusBar1.Panels(1).Text = gsEmpty
'              GDMDIform.StatusBar1.Panels(2).Text = gsEmpty
'              Exit Sub
'              End If
        
   Else
        'cuurentx, currenty are the pixel coordinates
        End If
Return

PictureBoxZoom_Error:

   If Err.Number = 480 Then
      'out of memory, can't autoredraw, recover gracefully
      ier = -1
      GoTo 10
      End If

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PictureBoxZoom of Module modHook"
End Sub

