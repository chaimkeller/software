VERSION 5.00
Begin VB.Form mapMAGfm 
   AutoRedraw      =   -1  'True
   Caption         =   "Magnification of marked area on main map"
   ClientHeight    =   6690
   ClientLeft      =   4260
   ClientTop       =   1830
   ClientWidth     =   7665
   Icon            =   "mapMAGfm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MAGpicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6710
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   6645
      ScaleWidth      =   7590
      TabIndex        =   0
      Top             =   0
      Width           =   7650
   End
End
Attribute VB_Name = "mapMAGfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub MagClosebut_Click()
'   Call form_queryunload(i%, j%)
'End Sub
Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
    Unload mapMAGfm
    Set skyMAGm = Nothing
    ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    magclose = True
    magbox = False
End Sub
Private Sub MAGpicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Maps.StatusBar1.Panels(2) = "Move the cursor to the desired location (click to center on this point)"
  If world = True Then GoTo ma50
  If map400 = True Then
     If mag > 1 Then
        kmxcc = kmxc
        kmycc = kmyc
        xo = kmxcc - (km400x / mag) * (mapwi2 - mapxdif) * 0.5
        yo = kmycc + (km400y / mag) * (maphi2 - mapydif) * 0.5
        kmxo = Fix(xo + (drag1x + X * magx) * km400x / mag) 'mapdif accounts for size of frame around picture
        kmyo = Fix(yo - (drag1y + Y * magy) * km400y / mag)
     Else
        kmxcc = kmxc + (km400x) * (mapwi - mapwi2 + mapxdif) / 2
        kmycc = kmyc - (km400y) * (maphi - maphi2 + mapydif) / 2
        'middle of screen corresponds to kmxc,kmyc
        'so topleft corner=origin corresponds to:
        xo = kmxcc - km400x * sizex / 2  'mapPictureform.mapPicture.Width / 2
        yo = kmycc + km400y * sizey / 2 'mapPictureform.mapPicture.Height / 2
        kmxo = Fix(xo + (drag1x + X * magx) * km400x)   'mapdif accounts for size of frame around picture
        kmyo = Fix(yo - (drag1y + Y * magy) * km400y)
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
        kmxo = Fix(xo + (drag1x + X * magx) * km50x / mag) 'mapdif accounts for size of frame around picture
        kmyo = Fix(yo - (drag1y + Y * magy) * km50y / mag)
     Else
        kmxcc = kmxc + (km50x) * (mapwi - mapwi2 + mapxdif) / 2
        kmycc = kmyc - (km50y) * (maphi - maphi2 + mapydif) / 2
        xo = kmxcc - km50x * sizex / 2  'mapPictureform.mapPicture.Width / 2
        yo = kmycc + km50y * sizey / 2 'mapPictureform.mapPicture.Height / 2
        kmxo = Fix(xo + (drag1x + X * magx) * km50x)
        kmyo = Fix(yo - (drag1y + Y * magy) * km50y)
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
ma50: Select Case coordmode%
       Case 1 'ITM
          Maps.Text1.Text = kmxoo
          Maps.Text2.Text = kmyoo
          'maps.label1.caption = "ITMx"
          'maps.label2.caption = "ITMy"
       Case 2 'GEO
          If world = True Then
             If mag > 1 Then
               lonc = lon '+ fudx / mag
               latc = lat '+ fudy / mag
               'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               'lono = xo + (drag1x + x * magx) * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
               'lato = yo - (drag1y + y * magy) * (180# / (sizewy * mag))
               xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               lono = xo + (drag1x + X * magx) * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
               lato = yo - (drag1y + Y * magy) * (deglat / (sizewy * mag))
               lg = lono
               lt = lato
             Else
               'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
               'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
               'xo = lonc - 90#
               'yo = latc + 90#
               'lono = xo + (drag1x + x * magx) * (180 / sizewx)
               'lato = yo - (drag1y + y * magy) * (180 / sizewy)
               lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
               latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
               xo = lonc - deglog / 2
               yo = latc + deglat / 2
               lono = xo + (drag1x + X * magx) * (deglog / sizewx)
               lato = yo - (drag1y + Y * magy) * (deglat / sizewy)
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
          'maps.label1.caption = "long."
          'maps.label2.caption = "latit."
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
          'maps.label1.caption = "UTMx"
          'maps.label2.caption = "UTMy"
       Case 4 'SKYLINE "UTM"
          Mode% = 1
          Call ITMSKY(kmxoo, kmyoo, T1, T2, Mode%)
          Maps.Text1.Text = T1
          Maps.Text2.Text = T2
          'maps.label1.caption = "SKYx"
          'maps.label2.caption = "SKYy"
       Case 5 'distance, view angle, azimuth
          'maps.label1.caption = "d(km)"
          'maps.text1.text = LTrim(Mid$(Str$(Sqr((kmxoo - kmxc) ^ 2 + (kmyoo - kmyc) ^ 2) * 0.001), 1, 7))
          'kmxcd = kmxc
          'kmycd = kmyc
          If world = True Then
             If mag > 1 Then
               lonc = lon '+ fudx / mag
               latc = lat '+ fudy / mag
               'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               'lono = xo + (drag1x + x * magx) * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
               'lato = yo - (drag1y + y * magy) * (180# / (sizewy * mag))
               xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
               yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
               lono = xo + (drag1x + X * magx) * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
               lato = yo - (drag1y + Y * magy) * (deglat / (sizewy * mag))
             Else
               'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
               'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
               'xo = lonc - 90#
               'yo = latc + 90#
               'lono = xo + (drag1x + x * magx) * (180 / sizewx)
               'lato = yo - (drag1y + y * magy) * (180 / sizewy)
               lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
               latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
               xo = lonc - deglog / 2
               yo = latc + deglat / 2
               lono = xo + (drag1x + X * magx) * (deglog / sizewx)
               lato = yo - (drag1y + Y * magy) * (deglat / sizewy)
               If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                  'fudge factor for inaccuracy of linear degree approx for large size map
                   lono = lono - 0.006906
                   lato = lato + 0.003878
                   End If
               End If
            End If
          Call dipcoord
        End Select
End Sub
Private Sub MAGpicture_MouseUp(Button As Integer, _
   Shift As Integer, X As Single, Y As Single)
      Select Case Button
      Case 1  'left button
           If Maps.Timer2.Enabled = True Then Exit Sub
      Case 2  'right button
         'MAGpicture.DrawWidth = 1
         'MAGpicture.Circle (X, Y), 100, 255
         MAGpicture.DrawWidth = 2 / magx
         MAGpicture.Circle (X, Y), 20, 255 '20 * magx, 255
         MAGpicture.DrawWidth = 1 / magx
         If world = False Then
            'kmxc = kmxoo: kmyc = kmyoo
            'Maps.Label5.Caption = "ITMx"
            'Maps.Label6.Caption = "ITMy"
            'Maps.Text7.Text = hgt
            'Maps.Text5.Text = kmxoo
            'Maps.Text6.Text = kmyoo
         Else
            cirworld = True
            Xworld = X * magx + drag1x
            Yworld = Y * magy + drag1y
            'mapPictureform.mapPicture.DrawWidth = 1
            'mapPictureform.mapPicture.Circle (Xworld, Yworld), 100, 255
            mapPictureform.mapPicture.DrawWidth = 2 * mag
            mapPictureform.mapPicture.Circle (Xworld, Yworld), 20, 255 '20 * mag, 255
            maprightform.Label1.Caption = "long."
            maprightform.Label2.Caption = "latit."
            maprightform.Text2 = lono 'LTrim$(RTrim$(Mid$(Str$(-180# + Xworld * 360# / mapPictureform.mapPicture.Width), 1, 10)))
            maprightform.Text3 = lato 'LTrim$(RTrim$(Mid$(Str$(90# - Yworld * 180# / mapPictureform.mapPicture.Height), 1, 10)))
            maprightform.Text4 = hgt
            maprightform.Text1 = "Name"
            'Maps.Label5.Caption = "long."
            'Maps.Label6.Caption = "latit."
            'Maps.Text7.Text = hgt
            'Maps.Text5.Text = maprightform.Text2
            'Maps.Text6.Text = maprightform.Text3
            If world = False Then
               lResult = FindWindow(vbNullString, terranam$)
               If lResult > 0 Then
                  ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                  End If
               End If
            maprightform.Visible = True
            ret = SetWindowPos(maprightform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            maprightform.SetFocus
            Exit Sub
            End If
         If map50 = True Then
            X50c = X * magx + drag1x
            Y50c = Y * magy + drag1y
            'mapPictureform.mapPicture.DrawWidth = 1
            'mapPictureform.mapPicture.Circle (X50c, Y50c), 100, 255
            mapPictureform.mapPicture.DrawWidth = 2 '2 * mag
            mapPictureform.mapPicture.Circle (X50c, Y50c), 20, 255 '20 * mag, 255
            'cir50 = True
            kmx50c = kmxoo: kmy50c = kmyoo: hgt50c = hgt
         Else
            X400c = X * magx + drag1x
            Y400c = Y * magy + drag1y
            'mapPictureform.mapPicture.DrawWidth = 1
            'mapPictureform.mapPicture.Circle (X400c, Y400c), 100, 255
            mapPictureform.mapPicture.DrawWidth = 2 '2 * mag
            mapPictureform.mapPicture.Circle (X400c, Y400c), 20, 255 '20 * mag, 255
            'cir400 = True
            'cir50 = False
            kmx400c = kmxoo: kmy400c = kmyoo: hgt400c = hgt
            End If
         maprightform.Text2 = kmxoo
         maprightform.Text3 = kmyoo
         maprightform.Text4 = hgt
         maprightform.Text1 = picf$
         ret = SetWindowPos(mapMAGfm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         maprightform.Visible = True
         maprightform.SetFocus
      End Select
   End Sub
Private Sub Form_Resize()
  If mapMAGfm.Visible = True And magbox = True Then
      'Height = MAGpicture.Height + (mapMAGfm.Height - ht1)
      'Width = MAGpicture.Width + (mapMAGfm.Width - wt1)
      MAGpicture.Width = MAGpicture.Width + (mapMAGfm.Width - wt1)
      MAGpicture.Height = MAGpicture.Height + (mapMAGfm.Height - ht1)
      wt1 = MAGpicture.Width
      ht1 = MAGpicture.Height
      magx = (drag2x - drag1x) / mapMAGfm.MAGpicture.Width
      magy = (drag2y - drag1y) / mapMAGfm.MAGpicture.Height
      mx = Fix(10 / magx) / 10
      my = Fix(10 / magy) / 10
      mapMAGfm.Caption = "Magnifcation " + Str(mx) + " x" + Str(my) + " of marked region on main map"
      mapMAGfm.MAGpicture.PaintPicture Maps.PictureClip2.Picture, 0, 0, mapMAGfm.MAGpicture.Width, _
      mapMAGfm.MAGpicture.Height, mapPictureform.Left + 88 + drag1x, mapPictureform.Top + 88 + drag1y, Abs(drag2x - drag1x), Abs(drag2y - drag1y)
'      mapMAGfm.MAGpicture.PaintPicture Maps.PictureClip2.Picture, 0, 0, mapMAGfm.MAGpicture.Width, _
'      mapMAGfm.MAGpicture.Height , mapPictureform.Left + Val(Maps.Text8.Text) + drag1x, mapPictureform.Top + Val(Maps.Text9.Text) + drag1y, Abs(drag2x - drag1x), Abs(drag2y - drag1y)
      If world = True And cirworld = True Then
         mapMAGfm.DrawMode = 13
         mapMAGfm.MAGpicture.DrawWidth = 1 / magx
         mapMAGfm.MAGpicture.Circle ((Xworld - drag1x) / magx, (Yworld - drag1y) / magy), 100, 255 '100 * magx, 255
         mapMAGfm.MAGpicture.DrawWidth = 2 / magx
         mapMAGfm.MAGpicture.Circle ((Xworld - drag1x) / magx, (Yworld - drag1y) / magy), 20, 255 '20 * magx, 255
         mapMAGfm.MAGpicture.DrawWidth = 1 / magx
         Exit Sub
         End If
      If map50 = True Then 'And cir50 = True Then
         mapMAGfm.DrawMode = 13
         mapMAGfm.MAGpicture.DrawWidth = 1 / magx
         mapMAGfm.MAGpicture.Circle ((X50c - drag1x) / magx, (Y50c - drag1y) / magy), 100, 255 '100 * magx, 255
         mapMAGfm.MAGpicture.DrawWidth = 2 / magx
         mapMAGfm.MAGpicture.Circle ((X50c - drag1x) / magx, (Y50c - drag1y) / magy), 20, 255 '20 * magx, 255
         mapMAGfm.MAGpicture.DrawWidth = 1 / magx
      ElseIf map400 = True Then 'And cir400 = True Then
         mapMAGfm.DrawMode = 13
         mapMAGfm.MAGpicture.DrawWidth = 1 / magx
         mapMAGfm.MAGpicture.Circle ((X400c - drag1x) / magx, (Y400c - drag1y) / magy), 100, 255 '100 * magx, 255
         mapMAGfm.MAGpicture.DrawWidth = 2 / magx
         mapMAGfm.MAGpicture.Circle ((X400c - drag1x) / magx, (Y400c - drag1y) / magy), 20, 255 '20 * magx, 255
         mapMAGfm.MAGpicture.DrawWidth = 1 / magx
         End If
      End If
End Sub
Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo magkeyerror
   Select Case KeyCode
      Case vbKeyReturn 'enter key
         If world = True Then
            If Maps.Text5.Text <> sEmpty Then
               If coordmode% = 2 Then
                  coordmode% = 5
                  Maps.Label1.Caption = "dist."
                  Maps.Label2.Caption = "Azim."
                  Maps.Label4.Visible = True
                  Maps.Text4.Visible = True
                  'kmxcd = kmxc
                  'kmycd = kmyc
                If world = True Then
                   If mag > 1 Then
                     lonc = lon '+ fudx / mag
                     latc = lat '+ fudy / mag
                     'xo = lonc - (180 / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                     'yo = latc + (180 / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                     'lono = xo + (drag1x + Xcoord * magx) * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     'lato = yo - (drag1y + Ycoord * magy) * (180# / (sizewy * mag))
                     xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                     yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                     lono = xo + (drag1x + Xcoord * magx) * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     lato = yo - (drag1y + Ycoord * magy) * (deglat / (sizewy * mag))
                  Else
                     'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
                     'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
                     'xo = lonc - 90#
                     'yo = latc + 90#
                     'lono = xo + (drag1x + Xcoord * magx) * (180 / sizewx)
                     'lato = yo - (drag1y + Ycoord * magy) * (180 / sizewy)
                     lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                     latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                     xo = lonc - deglog / 2
                     yo = latc + deglat / 2
                     lono = xo + (drag1x + Xcoord * magx) * (deglog / sizewx)
                     lato = yo - (drag1y + Ycoord * magy) * (deglat / sizewy)
                     If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                       'fudge factor for inaccuracy of linear degree approx for large size map
                        lono = lono - 0.006906
                        lato = lato + 0.003878
                        End If
                     End If
                  End If
                  Call dipcoord
               ElseIf coordmode% = 5 Then
                  coordmode% = 2
                  Maps.Label1.Caption = "long."
                  Maps.Label2.Caption = "latit."
                  Maps.Label4.Visible = False
                  Maps.Text4.Visible = False
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
         ElseIf coordmode% = 5 + addcoord% Then
            coordmode% = 1
            Maps.Label4.Visible = False
            Maps.Text4.Visible = False
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
                     'lono = xo + (drag1x + Xcoord * magx) * (180# / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     'lato = yo - (drag1y + Ycoord * magy) * (180# / (sizewy * mag))
                     xo = lonc - (deglog / (sizewx * mag)) * (mapwi2 - mapxdif) * 0.5
                     yo = latc + (deglat / (sizewy * mag)) * (maphi2 - mapydif) * 0.5
                     lono = xo + (drag1x + Xcoord * magx) * (deglog / (sizewx * mag))  'mapdif accounts for size of frame around picture
                     lato = yo - (drag1y + Ycoord * magy) * (deglat / (sizewy * mag))
                   Else
                     'lonc = lon + (180# / sizewx) * (mapwi - mapwi2 + mapxdif) / 2
                     'latc = lat - (180# / sizewy) * (maphi - maphi2 + mapydif) / 2
                     'xo = lonc - 90#
                     'yo = latc + 90#
                     'lono = xo + (drag1x + Xcoord * magx) * (180 / sizewx)
                     'lato = yo - (drag1y + Ycoord * magy) * (180 / sizewy)
                     lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                     latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                     xo = lonc - deglog / 2
                     yo = latc + deglat / 2
                     lono = xo + (drag1x + Xcoord * magx) * (deglog / sizewx)
                     lato = yo - (drag1y + Ycoord * magy) * (deglat / sizewy)
                     If sizewx = Screen.TwipsPerPixelX * 10201 And sizewy = Screen.TwipsPerPixelY * 5489 Then
                       'fudge factor for inaccuracy of linear degree approx for large size map
                        lono = lono - 0.006906
                        lato = lato + 0.003878
                        End If
                     End If
                  End If
               Call dipcoord
         End Select
      End Select
      Exit Sub
magkeyerror:
End Sub
