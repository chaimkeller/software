VERSION 5.00
Begin VB.Form mapPictureform 
   ClientHeight    =   13995
   ClientLeft      =   60
   ClientTop       =   -315
   ClientWidth     =   14160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "mapPictureform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   ScaleHeight     =   13995
   ScaleMode       =   0  'User
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox mapPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   13890
      Left            =   0
      ScaleHeight     =   13924.4
      ScaleMode       =   0  'User
      ScaleWidth      =   14890.96
      TabIndex        =   0
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "mapPictureform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo keyerror
   If map50 = True Then
      Step = 50 / mag
   ElseIf map400 = True Then
      Step = 400 / mag
      End If
   Select Case KeyCode
      Case vbKeyReturn 'enter key
         If world = True Then
            If Maps.Text5.Text <> sEmpty Then
               If coordmode% = 2 Then
                  coordmode% = 5
                  Maps.Label4.Visible = True
                  Maps.Text4.Visible = True
                  Maps.Label1.Caption = "dist."
                  Maps.Label2.Caption = "Azim."
                  Maps.Text1.ToolTipText = "Distance (km) from goto coordinates"
                  Maps.Text2.ToolTipText = "Azimuth (degrees) w.r.t. goto coordinates"
                  If mag > 1 Then
                     lonc = lon '+ fudx / mag
                     latc = lat '+ fudy / mag
                     'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                     'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                     'lono = xo + Xcoord * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     'lato = yo - Ycoord * (180# / (sizewy * mag))
                     xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                     yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                     lono = xo + Xcoord * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     lato = yo - Ycoord * (deglat / (sizewy * mag))
                   Else
                     'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
                     'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
                     'xo = lonc - 90#
                     'yo = latc + 90#
                     'lono = xo + Xcoord * (180 / sizewx)
                     'lato = yo - Ycoord * (180 / sizewy)
                     lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                     latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                     xo = lonc - deglog / 2
                     yo = latc + deglat / 2
                     lono = xo + Xcoord * (deglog / sizewx)
                     lato = yo - Ycoord * (deglat / sizewy)
                     If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                       'fudge factor for inaccuracy of linear degree approx for large size map
                        lono = lono - 0.006906
                        lato = lato + 0.003878
                        End If
                     End If
                  Call dipcoord
               ElseIf coordmode% = 5 Then
                  coordmode% = 2
                  Maps.Label4.Visible = False
                  Maps.Text4.Visible = False
                  Maps.Label1.Caption = "long."
                  Maps.Label2.Caption = "latit."
                  Maps.Text1.ToolTipText = "Map's X coordinate"
                  Maps.Text2.ToolTipText = "Map's Y coordinate"
                  End If
               End If
               Exit Sub
            End If
         'switch coordinates
         coordmode% = coordmode% + 1
         addcoord% = 0
         If Maps.Text5.Text <> sEmpty Then addcoord% = 1
         If coordmode% = 5 And addcoord% = 1 Then
            Maps.Label4.Visible = True
            Maps.Text4.Visible = True
            Maps.Text1.ToolTipText = "Distance (km) from goto coordinates"
            Maps.Text2.ToolTipText = "Azimuth (degrees) w.r.t. goto coordinates"
         ElseIf coordmode% = 5 + addcoord% Then
            Maps.Label4.Visible = False
            Maps.Text4.Visible = False
            Maps.Text1.ToolTipText = "Map's X coordinate"
            Maps.Text2.ToolTipText = "Map's Y coordinate"
            coordmode% = 1
            End If
         'redesplay coordinates in new system
         Select Case coordmode%
            Case 1 'ITM
               Maps.Text1.Text = kmxoo
               Maps.Text2.Text = kmyoo
               Maps.Label1.Caption = "ITMx"
               Maps.Label2.Caption = "ITMy"
            Case 2 'GEO
               Call casgeo(kmxoo, kmyoo, lg, lt)
               lgdeg = Fix(lg)
               lgmin = Abs(Fix((lg - Fix(lg)) * 60))
               lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
               ltdeg = Fix(lt)
               ltmin = Abs(Fix((lt - Fix(lt)) * 60))
               ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
               Maps.Label1.Caption = "long."
               Maps.Label2.Caption = "latit."
               If ltdeg = 0 And lt < 0 Then
                  Maps.Text2.Text = "-" + Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
               Else
                  Maps.Text2.Text = Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
                  End If
               If lgdeg = 0 And lg < 0 Then
                  Maps.Text1.Text = "-" + Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
               Else
                  Maps.Text1.Text = Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
                  End If
            Case 3 'UTM
               Call casgeo(kmxoo, kmyoo, lg, lt)
               Call GEOUTM(lt, lg, Z%, G1, G2)
               Maps.Text1.Text = Fix(G1)
               Maps.Text2.Text = Fix(G2)
               Maps.Label1.Caption = "UTMx"
               Maps.Label2.Caption = "UTMy"
            Case 4 'SKYLINE UTM
               Mode% = 1
               Call ITMSKY(kmxoo, kmyoo, T1, T2, Mode%)
               Maps.Text1.Text = T1
               Maps.Text2.Text = T2
               Maps.Label1.Caption = "SKYx"
               Maps.Label2.Caption = "SKYy"
            Case 5 'distance, viewangle, azimuth
               Maps.Label1.Caption = "dist."
               Maps.Label2.Caption = "Azim."
               'kmxcd = kmxc
               'kmycd = kmyc
                If world = True Then
                   If mag > 1 Then
                     lonc = lon '+ fudx / mag
                     latc = lat '+ fudy / mag
                     'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                     'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                     'lono = xo + Xcoord * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     'lato = yo - Ycoord * (180# / (sizewy * mag))
                     xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                     yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                     lono = xo + Xcoord * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     lato = yo - Ycoord * (deglat / (sizewy * mag))
                   Else
                     'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
                     'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
                     'xo = lonc - 90#
                     'yo = latc + 90#
                     'lono = xo + Xcoord * (180 / sizewx)
                     'lato = yo - Ycoord * (180 / sizewy)
                     lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                     latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                     xo = lonc - deglog / 2
                     yo = latc + deglat / 2
                     lono = xo + Xcoord * (deglog / sizewx)
                     lato = yo - Ycoord * (deglat / sizewy)
                     If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                       'fudge factor for inaccuracy of linear degree approx for large size map
                        lono = lono - 0.006906
                        lato = lato + 0.003878
                        End If
                     End If
                  End If
               Call dipcoord
         End Select
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 3
            DoEvents
         Loop
         'lResult = FindWindow(vbNullString, "1 (Help Author On)")
         'If lResult > 0 Then
         '   Maps.Text7.Text = lResult
         '   ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         '   End If
      Case vbKeyPageUp 'page up key - switch SKY box coordinates
         If world = True Then Exit Sub
         'switch coordinates
         coordmode2% = coordmode2% + 1
         If coordmode2% = 5 Then coordmode2% = 1
         'redesplay coordinates in new system
         kmxcc = kmxsky
         kmycc = kmysky
         Select Case coordmode2%
            Case 1 'ITM
               Maps.Text5.Text = kmxcc
               Maps.Text6.Text = kmycc
               Maps.Label5.Caption = "ITMx"
               Maps.Label6.Caption = "ITMy"
            Case 2 'GEO
               Call casgeo(kmxcc, kmycc, lg, lt)
               lgdeg = Fix(lg)
               lgmin = Abs(Fix((lg - Fix(lg)) * 60))
               lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
               ltdeg = Fix(lt)
               ltmin = Abs(Fix((lt - Fix(lt)) * 60))
               ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
               Maps.Label5.Caption = "long."
               Maps.Label6.Caption = "latit."
               If ltdeg = 0 And lt < 0 Then
                  Maps.Text6.Text = "-" + Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
               Else
                  Maps.Text6.Text = Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
                  End If
               If lgdeg = 0 And lg < 0 Then
                  Maps.Text5.Text = "-" + Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
               Else
                  Maps.Text5.Text = Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
                  End If
            Case 3 'UTM
               Call casgeo(kmxcc, kmycc, lg, lt)
               Call GEOUTM(lt, lg, Z%, G1, G2)
               Maps.Text5.Text = Fix(G1)
               Maps.Text6.Text = Fix(G2)
               Maps.Label5.Caption = "UTMx"
               Maps.Label6.Caption = "UTMy"
            Case 4 'SKYLINE UTM
               Mode% = 1
               Call ITMSKY(kmxcc, kmycc, T1, T2, Mode%)
               Maps.Text5.Text = T1
               Maps.Text6.Text = T2
               Maps.Label5.Caption = "SKYx"
               Maps.Label6.Caption = "SKYy"
         End Select
      Case 38 'up arrow
         If world = True Then
            'lat = lat + 3.6 / mag
            lat = lat + deglog / (mag * 50)
         ElseIf map50 = True Then
            kmyc = kmyc + 50 / mag
         ElseIf map400 = True Then
            kmyc = kmyc + 400 / mag
            End If
         Call showcoord
         Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      Case 40 'down arrow
         If world = True Then
            'lat = lat - 3.6 / mag
            lat = lat - deglog / (mag * 50)
         ElseIf map50 = True Then
            kmyc = kmyc - 50 / mag
         ElseIf map400 = True Then
            kmyc = kmyc - 400 / mag
            End If
         Call showcoord
         Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      Case 37 'left arrow
         If world = True Then
            'lon = lon - 3.6 / mag
            lon = lon - deglog / (mag * 50)
         ElseIf map50 = True Then
            kmxc = kmxc - 50 / mag
         ElseIf map400 = True Then
            kmxc = kmxc - 400 / mag
            End If
         Call showcoord
         Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      Case 39 'right arrow
         If world = True Then
            'lon = lon + 3.6 / mag
            lon = lon + deglog / (mag * 50)
         ElseIf map50 = True Then
            kmxc = kmxc + 50 / mag
         ElseIf map400 = True Then
            kmxc = kmxc + 400 / mag
            End If
         Call showcoord
         Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      Case Else
   End Select
   Exit Sub
keyerror:
   Exit Sub
End Sub

Private Sub Form_Load()
   MapOn = True
End Sub

Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

  On Error GoTo f999
  If resizes = True Then GoTo f999
  resizes = True
  If world = False Then
    If mapPictureform.Width > sizex + 60 Then
       mapPictureform.Width = sizex + 60
       If mapPictureform.Width > Screen.Width Then
          mapPictureform.Width = Screen.Width - 60
          End If
       Exit Sub
       End If
    If mapPictureform.Width > Screen.Width Then
       mapPictureform.Width = Screen.Width - 60
       End If
    If mapPictureform.Height > sizey + 60 Then
       mapPictureform.Height = sizey + 60
       If mapPictureform.Height > Screen.Height Then
          mapPictureform.Height = Screen.Height - 1900
          End If
       Exit Sub
       End If
    If mapPictureform.Height > Screen.Height Then
       mapPictureform.Height = Screen.Height - 1900
       End If
  ElseIf world = True Then
    If mapimport Then
        pixwwi = xpix '+ 10
        pixwhi = ypix '+ 10
        sizewx = Screen.TwipsPerPixelX * pixwwi '# twips in half of picture=8850/2
        sizewy = Screen.TwipsPerPixelY * pixwhi '=8850/2
        mapPictureform.mapPicture.Width = mapPictureform.Width
        mapPictureform.mapPicture.Height = mapPictureform.Height
        If mapwi2 > sizewx + 60 Then
           mapPictureform.Width = sizewx '+ 60 '60 is the size (pixels) of the borders
           mapPictureform.mapPicture.Width = sizewx
           mapwi = mapPictureform.Width
           mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
           End If
        If maphi2 > sizewy + 60 Then
           mapPictureform.mapPicture.Height = sizewy
           maphi = mapPictureform.Height
           mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
           End If
        If world = True Then
           mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width 'mapxdif + 35
           mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height 'mapydif + 35
           End If
        kmwx = 2 * deglog / sizewx
        kmwy = deglat / sizewy

        Call loadpictures  'load appropriate map tiles into off-screen buffers
        Call blitpictures   'blit desired portions of the off-screen buffers to the screen
       End If
       
    If mapPictureform.Width > sizewx + 60 Then
       mapPictureform.Width = sizewx + 60
       Call blitpictures
       Exit Sub
       End If
    If mapPictureform.Height > sizewy + 60 Then
       mapPictureform.Height = sizewy + 60
       Call blitpictures
       Exit Sub
       End If
     End If
   
  'If topotype% = 1 Then '***********
  '   mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
  '   mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
  '   End If
     
  If mapPictureform.Visible = True And magbox = False Then
     Call blitpictures
     End If
f999:
   resizes = False
   Exit Sub

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form mapPictureform"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MapOn = False
  Unload Me
  Set mapPictureform = Nothing
End Sub

Private Sub mappicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Maps.StatusBar1.Panels(2) = "Move the cursor to the desired location (click to center on this point)"
  If impcenter = True Then
     Maps.StatusBar1.Panels(2) = "Move the cursor to the map's true center and then click."
     'determine new fudx, fudy
     End If
  
  Xcoord = x
  Ycoord = y
  If world = True Then GoTo m10
  If map400 = True Then
     If mag > 1 Then
        kmxcc = kmxc
        kmycc = kmyc
        xo = kmxcc - (km400x / mag) * (mapwi2 - mapxdif) * 0.5
        yo = kmycc + (km400y / mag) * (maphi2 - mapydif) * 0.5
        kmxo = Fix(xo + x * km400x / mag) 'mapdif accounts for size of frame around picture
        kmyo = Fix(yo - y * km400y / mag)
     Else
        kmxcc = kmxc + (km400x) * (mapwi - mapwi2 + mapxdif) / 2
        kmycc = kmyc - (km400y) * (maphi - maphi2 + mapydif) / 2
        'middle of screen corresponds to kmxc,kmyc
        'so topleft corner=origin corresponds to:
        xo = kmxcc - km400x * sizex / 2  'mapPictureform.mapPicture.Width / 2
        yo = kmycc + km400y * sizey / 2 'mapPictureform.mapPicture.Height / 2
        kmxo = Fix(xo + x * km400x)   'mapdif accounts for size of frame around picture
        kmyo = Fix(yo - y * km400y)
        End If
     If (kmxo <> kmxoo Or kmyo <> kmyoo) Then
        kmxoo = kmxo: kmyoo = kmyo
        If noheights = False Then
           Call heights(kmxo, kmyo, hgt)
        ElseIf noheights = True Then
           hgt = 0#
           End If
        End If
     Maps.Text3.Text = Str$(hgt)
   ElseIf map50 = True Then
     If mag > 1 Then
        kmxcc = kmxc
        kmycc = kmyc
        xo = kmxcc - (km50x / mag) * (mapwi2 - mapxdif) * 0.5
        yo = kmycc + (km50y / mag) * (maphi2 - mapydif) * 0.5
        kmxo = Fix(xo + x * km50x / mag) 'mapdif accounts for size of frame around picture
        kmyo = Fix(yo - y * km50y / mag)
     Else
        kmxcc = kmxc + (km50x) * (mapwi - mapwi2 + mapxdif) / 2
        kmycc = kmyc - (km50y) * (maphi - maphi2 + mapydif) / 2
        xo = kmxcc - km50x * sizex / 2  'mapPictureform.mapPicture.Width / 2
        yo = kmycc + km50y * sizey / 2 'mapPictureform.mapPicture.Height / 2
        kmxo = Fix(xo + x * km50x)
        kmyo = Fix(yo - y * km50y)
        End If
     If (kmxo <> kmxoo Or kmyo <> kmyoo) Then
        kmxoo = kmxo: kmyoo = kmyo
        If noheights = False Then
           Call heights(kmxo, kmyo, hgt)
        ElseIf noheights = True Then
           hgt = 0#
           End If
        End If
     Maps.Text3.Text = Str$(hgt)
     End If
m10: Select Case coordmode%
       Case 1 'ITM
          Maps.Text1.Text = kmxoo
          Maps.Text2.Text = kmyoo
          'Label5.Caption = "ITMx"
          'Label6.Caption = "ITMy"
       Case 2 'GEO
          If world = True Then
             If mag > 1 Then
               lonc = lon '+ fudx / mag
               latc = lat '+ fudy / mag
               'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               'lono = xo + x * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
               'lato = yo - y * (180# / (sizewy * mag))
               'X = 0
               'Y = 0
               xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               lono = xo + x * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
               lato = yo - y * (deglat / (sizewy * mag))
               lg = lono
               lt = lato
             Else
               'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
               'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
               'xo = lonc - 90#
               'yo = latc + 90#
               'lono = xo + x * (180 / sizewx)
               'lato = yo - y * (180 / sizewy)
               lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
               latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
               xo = lonc - deglog / 2
               yo = latc + deglat / 2
               lono = xo + x * (deglog / sizewx)
               lato = yo - y * (deglat / sizewy)
               If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                  'fudge factor for inaccuracy of linear degree approx for large size map
                  lono = lono - 0.006906
                  lato = lato + 0.003878
                  End If
               lg = lono
               lt = lato
               End If
            If noheights = False Then
               Call worldheights(lg, lt, hgt)
               If hgt = -9999 Then hgt = 0
               Maps.Text3.Text = Str$(hgt)
               End If
          Else
             Call casgeo(kmxoo, kmyoo, lg, lt)
             End If
          lgdeg = Fix(lg)
          lgmin = Abs(Fix((lg - Fix(lg)) * 60))
          lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
          ltdeg = Fix(lt)
          ltmin = Abs(Fix((lt - Fix(lt)) * 60))
          ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
          'Label5.Caption = "long."
          'Label6.Caption = "latit."
          If ltdeg = 0 And lt < 0 Then
             Maps.Text2.Text = "-" + Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
          Else
             Maps.Text2.Text = Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
             End If
          If lgdeg = 0 And lg < 0 Then
             Maps.Text1.Text = "-" + Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
          Else
             Maps.Text1.Text = Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
             End If
       Case 3 'UTM
          Call casgeo(kmxoo, kmyoo, lg, lt)
          Call GEOUTM(lt, lg, Z%, G1, G2)
          Maps.Text1.Text = Fix(G1)
          Maps.Text2.Text = Fix(G2)
          'Label5.Caption = "UTMx"
          'Label6.Caption = "UTMy"
       Case 4 'SKYLINE UTM
          Mode% = 1
          Call ITMSKY(kmxoo, kmyoo, T1, T2, Mode%)
          Maps.Text1.Text = T1
          Maps.Text2.Text = T2
          'Label5.Caption = "SKYx"
          'Label6.Caption = "SKYy"
       Case 5 'distance, viewangle, azimuth
          'maps.text1.text = LTrim$(Mid$(Str$(Sqr((kmxoo - kmxc) ^ 2 + (kmyoo - kmyc) ^ 2) * 0.001), 1, 10))
          'kmxcd = kmxc
          'kmycd = kmyc
          If world = True Then
             If mag > 1 Then
               lonc = lon '+ fudx / mag
               latc = lat '+ fudy / mag
               'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               'lono = xo + x * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
               'lato = yo - y * (180# / (sizewy * mag))
               xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               lono = xo + x * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
               lato = yo - y * (deglat / (sizewy * mag))
             Else
               'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
               'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
               'xo = lonc - 90#
               'yo = latc + 90#
               'lono = xo + x * (180 / sizewx)
               'lato = yo - y * (180 / sizewy)
               lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
               latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
               xo = lonc - deglog / 2
               yo = latc + deglat / 2
               lono = xo + x * (deglog / sizewx)
               lato = yo - y * (deglat / sizewy)
               If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                  'fudge factor for inaccuracy of linear degree approx for large size map
                  lono = lono - 0.006906
                  lato = lato + 0.003878
                  End If
               End If
            End If
          Call dipcoord
       End Select
   If dragbegin = True And Button = 1 And dragbox = True Then 'dragging continues, draw box
      mapPictureform.mapPicture.DrawMode = 7
      mapPictureform.mapPicture.DrawStyle = vbDot
      mapPictureform.DrawWidth = 1
      mapPictureform.mapPicture.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
      mapPictureform.mapPicture.Line (x, y)-(drag1x, drag1y), QBColor(15), B
      drag2x = x: drag2y = y
   ElseIf dragbegin = True And Button = 1 And dragbox = False And drawbox = False Then
      mapPictureform.mapPicture.DrawMode = 7
      mapPictureform.mapPicture.DrawStyle = vbDot
      mapPictureform.DrawWidth = 1
      mapPictureform.mapPicture.Line (x, y)-(drag1x, drag1y), QBColor(15), B
      drag2x = x: drag2y = y
      dragbox = True
      End If
   If travelmode = True And travelnum% >= 1 Then
      mapPictureform.mapPicture.DrawMode = 7
      mapPictureform.mapPicture.DrawStyle = vbDot
      travelX = mapPictureform.Width / 2 - mapxdif
      travelY = mapPictureform.Height / 2 - mapydif
      If newblit = False Then mapPictureform.mapPicture.Line (drag2x, drag2y)-(travelX, travelY), QBColor(15)
      newblit = False
      mapPictureform.mapPicture.Line (x, y)-(travelX, travelY), QBColor(15)
      drag2x = x: drag2y = y
      End If
   mapPictureform.mapPicture.DrawStyle = vbSolid
   mapPictureform.mapPicture.DrawMode = 13
 End Sub
Private Sub mappicture_MouseUp(Button As Integer, _
   Shift As Integer, x As Single, y As Single)
   Dim xrl As Long, yrl As Long
      Xcoord = x
      Ycoord = y
      Select Case Button
      Case 1  'left button
         
         'inactivate left clicking if program is following the TerraExplorer
         If tblbuttons(18) = 1 Or routeload = True Then Exit Sub
         
         If (drag1x = drag2x And drag1y = drag2y) Or travelmode = True Or skyleftjump = True Then
            dragbegin = False
            dragbox = False
            Maps.Text7.Text = hgt
            If Maps.Text7.Text = sEmpty Then Maps.Text7.Text = "0"
            hgtpos = hgt
            If travelmode = True Then
               travelnum% = travelnum% + 1
               ReDim Preserve travel(2, travelnum%)
               'If travelnum% > travelmax% Then 'reached array limit
               '   response = MsgBox("You have exhausted the travel array, please unpress the travel button to leave the input mode!", vbExclamation + vbOKOnly, "Maps & more")
               '   Exit Sub
               '   End If
               'convert coordinates to SKY and store in travel array
               If world = False Then
                  Mode% = 1
                  Call ITMSKY(kmxoo, kmyoo, T1, T2, Mode%)
                  travel(1, travelnum%) = T1
                  travel(2, travelnum%) = T2
               Else
                  travel(1, travelnum%) = lono
                  travel(2, travelnum%) = lato
                  End If
               End If
            If world = False Then
               'mode% = 1
               'Call ITMSKY(kmxoo, kmyoo, T1, T2, mode%)
               'Text4.Text = T1
               'Text1.Text = T2
               'Label4.Caption = "SKYx"
               'Label7.Caption = "SKYy"
               kmxc = kmxoo: kmyc = kmyoo
               kmxsky = kmxc: kmysky = kmyc
               Maps.Text5.Text = kmxc
               Maps.Text6.Text = kmyc
               Maps.Label5.Caption = "ITMx"
               Maps.Label6.Caption = "ITMy"
               coordmode2% = 1
            ElseIf world = True Then
               If impcenter = True Then GoTo mup50
               If coordmode% = 5 Then
                  txt1$ = Maps.Text1.Text
                  txt2$ = Maps.Text2.Text
                  txt3$ = Maps.Label1.Caption
                  txt4$ = Maps.Label2.Caption
                  End If

               Maps.Text5.Text = Format(lono, "###0.0#####") '-180# + X * 360# / mappictureform.mappicture.Width
               Maps.Text6.Text = Format(lato, "##0.0#####") '90# - Y * 180# / mappictureform.mappicture.Height
               Maps.Label5.Caption = "long."
               Maps.Label6.Caption = "latit."
               Xworld = x
               Yworld = y
               cirworld = True
               hgtworld = hgt
               lon = lono
               lat = lato
mup50:         If impcenter = True Then
                  If fudx = 0 And fudy = 0 Then
                     fudx = lono - lon '- 0.01275
                     fudy = lato - lat '+ 0.003
'                     fudx = lonc / lon 'fudx = lono - lon '- 0.01275
'                     fudy = latc / lat 'fudy = lato - lat '+ 0.003
'                     deglat = fudy * deglat
'                     deglog = fudx * deglog
                  Else
                      fudx = fudx + (lono - lon)
                      fudy = fudy + (lato - lat)
'                     fudy = latc / lat 'lato - lat '+ 0.003
'                     fudx = lonc / lon
'                     deglat = fudy * deglat
'                     deglog = fudx * deglog
                      End If
                  'mapPictureform.mapPicture.Circle (Xcoord, Ycoord), 20, 255
                  'impcenter = False
                  End If
               Screen.MousePointer = vbHourglass
               Call blitpictures
               Screen.MousePointer = vbDefault
               If skyleftjump = True Then
                  'go there there on 3D Viewer
                  'check if there is a USGUS EROS CD in the CD-drive
                   On Error GoTo sunrerr
                   myfile = Dir(worlddtm + ":\E020N40\E020N40.GIF") 'Dir(worlddtm + ":\Gt30dem.gif")
                   If myfile = sEmpty Then
                      'check if there are stored DTM files in c:\dtm
                       doclin$ = Dir(drivdtm$ & "*.BIN")
                       myfile = Dir(drivdtm$ & "eros.tm3")
                       If doclin$ <> sEmpty And myfile <> sEmpty And Dir(drivdtm$ & "*.BI1") <> sEmpty Then
                         'leave rest of checking for sunrisesunset routine
                          checkdtm = True
                          Call sunrisesunset(1)
                       ElseIf Not NoCDWarning Then
                          Maps.Toolbar1.Buttons(26).value = tbrUnpressed
                          ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                          response = MsgBox("USGS EROS CD not found!  Please enter the appropriate CD, and then press the DTM button!", vbCritical + vbOKOnly, "Maps & More")
'                          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                          BringWindowToTop (mapPictureform.hWnd)
                          NoCDWarning = True
                          Exit Sub
                          End If
                    Else
                       Call sunrisesunset(1)
                       End If
                    End If
               End If
            If world = True Then
               If coordmode% = 5 Then 'fix some type of timing bug that erases some of the entries
                  waitime = Timer
                  Do Until Timer > waitime + 0.001
                     DoEvents
                  Loop
                  Maps.Text1.Text = txt1$
                  Maps.Text2.Text = txt2$
                  Maps.Label1.Caption = txt3$
                  Maps.Label2.Caption = txt4$
                  End If
               Exit Sub
            ElseIf map400 = True Then
               'cir400 = True
               'cir50 = False
                X400c = x: Y400c = y
                kmx400c = kmxoo: kmy400c = kmyoo
                kmxc = kmx400c: kmyc = kmy400c
                hgt400c = hgt
               Screen.MousePointer = vbHourglass
               Call blitpictures
               Screen.MousePointer = vbDefault
               If tblbuttons(19) = 1 Then
                  skyleftjump = True
                  Call skyTERRAgoto
                  End If
            ElseIf map50 = True Then
               'cir50 = True
                X50c = x: Y50c = y
                kmx50c = kmxoo: kmy50c = kmyoo
                kmxc = kmx50c: kmyc = kmy50c
                If topotype% = 1 Then '****************
                   kmxc = kmxc + 5785 * km50x
                   kmyc = kmyc - 10615 * km50y
                   End If
                
                hgt50c = hgt
               'now calculate postion on 1:400 map
               Screen.MousePointer = vbHourglass
               Call blitpictures
               Screen.MousePointer = vbDefault
               'cir400 = True
               If tblbuttons(19) = 1 Then
                  skyleftjump = True
                  Call skyTERRAgoto
                  End If
               End If
         Else 'signales end of drag '<
               If magbox = True And tblbuttons(19) = 1 Then 'notify user that can't terra jump until close mag window
                  lResult = FindWindow(vbNullString, terranam$)
                  If lResult > 0 Then
                      ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                      End If
                  response = MsgBox("You can't jump around with the Terra Viewer until" + _
                           " you close the Magnifaction box.  Close it now?", vbInformation + vbYesNo, "Skylight")
                  If lResult > 0 Then
'                     ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                     BringWindowToTop (lResult)
                     End If
                  If response = vbYes Then
                     Unload mapMAGfm
                     Set mapMAGfm = Nothing
                     Maps.Combo1.Enabled = True
                     End If
                  Exit Sub
                  End If
                  End If
               'erase last box and redraw it
               If dragbox = True And dragbegin = True And (drag2x <> drag1x And drag2y <> drag1y) And Button = 1 Then
                  mapPictureform.mapPicture.DrawMode = 7
                  mapPictureform.mapPicture.DrawStyle = vbDot
                  mapPictureform.mapPicture.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                  mapPictureform.mapPicture.DrawWidth = 2
                  mapPictureform.mapPicture.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                  mapPictureform.mapPicture.DrawMode = 13
                  dragbegin = False
                  dragbox = False
                  drawbox = True
                  mapPictureform.mapPicture.DrawWidth = 1
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
                 If world = False Then
                    lResult = FindWindow(vbNullString, terranam$)
                    If lResult > 0 Then
                       ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                       End If
                    End If
               If Maps.mnuMagDragEnable.Checked Then
'                  ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                  BringWindowToTop (mapPictureform.hWnd)
                  Call keybd_event(VK_SNAPSHOT, 1, 0, 0)
                  Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_KEYUP, 0)
                  waitime = Timer
                  Do Until Timer > waitime + 1
                     DoEvents
                  Loop
                  Maps.Combo1.Enabled = False
                  For j% = 2 To 15
                     Maps.Toolbar1.Buttons(j%).Enabled = False
                  Next j%
                  Maps.PictureClip2.Picture = Clipboard.GetData()
                  magx = (drag2x - drag1x) / mapMAGfm.MAGpicture.Width
                  magy = (drag2y - drag1y) / mapMAGfm.MAGpicture.Height
                  mx = Fix(10 / magx) / 10
                  my = Fix(10 / magy) / 10
                  mapMAGfm.Visible = True
'                  ret = SetWindowPos(mapMAGfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                  BringWindowToTop (mapMAGfm.hWnd)
                  mapMAGfm.Caption = "Magnifcation " + Str(mx) + " x" + Str(my) + " of marked region on main map"
                  mapMAGfm.MAGpicture.DrawMode = 13
                  'now load the bitmap from the clipboard to pictureclip2
                  mapMAGfm.MAGpicture.PaintPicture Maps.PictureClip2.Picture, 0, 0, mapMAGfm.MAGpicture.Width, _
                  mapMAGfm.MAGpicture.Height, mapPictureform.Left + 88 + drag1x, mapPictureform.Top + 88 + drag1y, Abs(drag2x - drag1x), Abs(drag2y - drag1y)
'                 mapMAGfm.MAGpicture.PaintPicture Maps.PictureClip2.Picture, 0, 0, mapMAGfm.MAGpicture.Width, _
  '               mapMAGfm.MAGpicture.Height, mapPictureform.Left + Val(Maps.Text8.Text) + drag1x, mapPictureform.Top + Val(Maps.Text9.Text) + drag1y, Abs(drag2x - drag1x), Abs(drag2y - drag1y)
                  magbox = True
                  ht1 = mapMAGfm.Height
                  wt1 = mapMAGfm.Width
                  'If world = True And cirworld = True Then
                  '  mapMAGfm.DrawMode = 13
                  '  mapMAGfm.MAGpicture.DrawWidth = 1 / magx
                  '  mapMAGfm.MAGpicture.Circle ((Xworld - drag1x) / magx, (Yworld - drag1y) / magy), 100 * magx, 255
                  '  mapMAGfm.MAGpicture.DrawWidth = 2 / magx
                  '  mapMAGfm.MAGpicture.Circle ((Xworld - drag1x) / magx, (Yworld - drag1y) / magy), 20 * magx, 255
                  '  mapMAGfm.MAGpicture.DrawWidth = 1
                  '  GoTo mu50
                  '  End If
                  'If map50 = True And cir50 = True Then
                  '   mapMAGfm.DrawMode = 13
                  '   mapMAGfm.MAGpicture.DrawWidth = 1 / magx
                  '   mapMAGfm.MAGpicture.Circle ((X50c - drag1x) / magx, (Y50c - drag1y) / magy), 100 * magx, 255
                  '   mapMAGfm.MAGpicture.DrawWidth = 2 / magx
                  '   mapMAGfm.MAGpicture.Circle ((X50c - drag1x) / magx, (Y50c - drag1y) / magy), 20 * magx, 255
                  '   mapMAGfm.MAGpicture.DrawWidth = 1 / magx
                  'ElseIf map400 = True And cir400 = True Then
                  '   mapMAGfm.DrawMode = 13
                  '   mapMAGfm.MAGpicture.DrawWidth = 1 / magx
                  '   mapMAGfm.MAGpicture.Circle ((X400c - drag1x) / magx, (Y400c - drag1y) / magy), 100 * magx, 255
                  '   mapMAGfm.MAGpicture.DrawWidth = 2 / magx
                  '   mapMAGfm.MAGpicture.Circle ((X400c - drag1x) / magx, (Y400c - drag1y) / magy), 20 * magx, 255
                  '   mapMAGfm.MAGpicture.DrawWidth = 1 / magx
                  '   End If
mu50:             mapMAGfm.MAGpicture.SetFocus
                  Do Until magclose = True
                    DoEvents
                  Loop
                  
               ElseIf Maps.mnuExcelDrag.Checked Then
                  'dump 3D topo data to Excel for plotting
                  response = MsgBox("Export to Excel or to xyz file?", vbQuestion + vbYesNoCancel, "Maps&More")
                  If response = vbYes Then
                     Call ExportToExcel(drag1x, drag1y, drag2x, drag2y)
                     End If
               ElseIf Maps.mnuTrigDrag.Checked Then
                  response = MsgBox("Drag defines trig point influence?", vbQuestion + vbYesNoCancel, "Maps&More")
                  If response = vbYes And hgtTrig <> -9999 Then
                     Call TrigPointAdjust(drag1x, drag1y, drag2x, drag2y)
                  ElseIf response = vbYes And hgtTrig = -9999 Then
                     MsgBox "You haven't defined a Trig Point!" & vbLf & vbLf & _
                            "Right click on the trig point to define its coordinates." & vbLf & _
                            "Then enter its height in the ""hgt"" text box." _
                            , vbExclamation + vbOKOnly, "Maps&More"
                     End If
               ElseIf Maps.mnuColumnFix.Checked Then
                  response = MsgBox("Drag defines column fix area?", vbQuestion + vbYesNoCancel, "Maps&More")
                  If response = vbYes Then
                     XfixCoord = InputBox("Column Fix X Coord (ITMx) =", "Column Glitch Fix", "100000")
                     If XfixCoord = sEmpty Then
                        'just canceled
                     Else
                        Call GlitchFix(drag1x, drag1y, drag2x, drag2y, CLng(XfixCoord), 0)
                        End If
                     End If
               ElseIf Maps.mnuRowfix.Checked Then
                  response = MsgBox("Drag defines row fix area?", vbQuestion + vbYesNoCancel, "Maps&More")
                  If response = vbYes Then
                     YfixCoord = InputBox("Column Fix Y Coord (ITMy) =", "Row Glitch Fix", "1100000")
                     If YfixCoord = sEmpty Then
                        'just canceled
                     Else
                        Call GlitchFix(drag1x, drag1y, drag2x, drag2y, CLng(YfixCoord), 1)
                        End If
                     End If
               ElseIf Maps.mnuXYFix.Checked Then
                  ret = BringWindowToTop(mapPictureform.hWnd)
                  If drag1x <> drag2x And drag1y <> drag2y Then
                    response = MsgBox("Drag defines XY fix area?", vbQuestion + vbYesNoCancel, "Maps&More")
                    If response = vbYes Then
                       mapXYGlitchfrm.Show
                       End If
                    End If
                  End If
                  
                  BringWindowToTop (mapPictureform.hWnd)
                  mapPictureform.mapPicture.DrawMode = 7
                  mapPictureform.mapPicture.DrawWidth = 2
                  mapPictureform.mapPicture.Line (drag2x, drag2y)-(drag1x, drag1y), QBColor(15), B
                  mapPictureform.mapPicture.DrawWidth = 1
                  mapPictureform.mapPicture.DrawMode = 13
                  drawbox = False
                  dragbegin = False
                  If world = False And lResult > 0 Then
'                     ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                     BringWindowToTop (lResult)
                     End If
                  Maps.Combo1.Enabled = True
                  For j% = 2 To 15
                     Maps.Toolbar1.Buttons(j%).Enabled = True
                  Next j%
               End If
      Case 2  'right button
         If mapSearchVis Then 'determine nearest search result
            Call FindSearchResult(x, y)
            Exit Sub
            End If
      
         'If world = False Then
         '   Maps.Label5.Caption = "ITMx"
         '   Maps.Label6.Caption = "ITMy"
         '   Maps.Text7.Text = hgt
         '   Maps.Text5.Text = kmxoo
         '   Maps.Text6.Text = kmyoo
         '   End If
         'kmxc = kmxoo: kmyc = kmyoo
         Call blitpictures
         'mapPictureform.mapPicture.DrawWidth = 1
         'mapPictureform.mapPicture.Circle (X, Y), 100, 255
         mapPictureform.mapPicture.DrawWidth = 2 '2 * mag
         mapPictureform.mapPicture.Circle (x, y), 20, 255 '20 * mag, 255
         mapPictureform.mapPicture.DrawWidth = 1 '1 * mag
         If world = True Then
            Xworld = x
            Yworld = y
            cirworld = True
            hgtworld = hgt
            maprightform.Text2 = LTrim$(RTrim$(Mid$(Str$(lono), 1, 10)))
            maprightform.Text3 = LTrim$(RTrim$(Mid$(Str$(lato), 1, 10)))
            maprightform.Text4 = hgt
            'Maps.Label5.Caption = "long."
            'Maps.Label6.Caption = "latit."
            'Maps.Text7.Text = hgt
            'Maps.Text5.Text = maprightform.Text2
            'Maps.Text6.Text = maprightform.Text4
            maprightform.Text1 = "Name"
            maprightform.Visible = True
'            ret = SetWindowPos(maprightform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            BringWindowToTop (maprightform.hWnd)
            maprightform.SetFocus
            Exit Sub
            End If
'         kmxc = kmxoo: kmyc = kmyoo
'         If map400 = True Then
'            cir400 = True
'            cir50 = False
'            X400c = X: Y400c = Y
'            kmx400c = kmxoo: kmy400c = kmyoo: hgt400c = hgt
'         ElseIf map50 = True Then
'            cir50 = True
'            X50c = X: Y50c = Y
'            kmx50c = kmxoo: kmy50c = kmyoo: hgt50c = hgt
'            End If

         maprightform.Text2 = kmxoo
         maprightform.Text3 = kmyoo
         maprightform.Text4 = hgt
         maprightform.Text1 = picf$
         If world = False Then
            lResult = FindWindow(vbNullString, terranam$)
            If lResult > 0 Then
               ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               End If
            End If
         maprightform.Visible = True
'         ret = SetWindowPos(maprightform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (maprightform.hWnd)
         maprightform.SetFocus
      End Select
      
   Exit Sub

sunrerr:
   myfile = sEmpty
   Resume Next
   
   End Sub
Private Sub mappicture_mousedown(Button As Integer, _
   Shift As Integer, x As Single, y As Single)
   If Button = 1 And drawbox = False And travelmode = False And skyleftjump = False Then 'maybe beginning of drag operation
      drag1x = x
      drag1y = y
      dragbegin = True
      drag2x = drag1x
      drag2y = drag1y
      End If
   End Sub


