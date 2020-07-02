VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GDMagform 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   6465
   ClientLeft      =   6060
   ClientTop       =   1065
   ClientWidth     =   5925
   ControlBox      =   0   'False
   Icon            =   "GDMagform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2340
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":04E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":06B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":0B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":0C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":0E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDMagform.frx":1536
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   4935
      Left            =   60
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   9
      Top             =   360
      Width           =   4935
      Begin VB.Image Image1 
         Height          =   3735
         Left            =   360
         MousePointer    =   2  'Cross
         Stretch         =   -1  'True
         Top             =   420
         Width           =   4155
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1560
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   8
      Top             =   1260
      Visible         =   0   'False
      Width           =   15
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "magnifykey"
            Object.ToolTipText     =   "Magnify by 10%"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "demagnifykey"
            Object.ToolTipText     =   "Reduce by 10%"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printkey"
            Object.ToolTipText     =   "Print this window"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "closekey"
            Object.ToolTipText     =   "Close the window"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   9350
         EndProperty
      EndProperty
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         ToolTipText     =   "Elevation (m) at cursor postion"
         Top             =   60
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   260
         Left            =   4600
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "hgt."
         ToolTipText     =   "meters"
         Top             =   60
         Width           =   315
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3400
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "Y coordinate of cursor"
         Top             =   60
         Width           =   1100
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2900
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "ITMy"
         Top             =   60
         Width           =   555
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1720
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "X coordinate of cursor"
         Top             =   60
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "ITMx"
         Top             =   60
         Width           =   495
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   5400
      Width           =   4875
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      Left            =   5160
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close Window"
      Height          =   435
      Left            =   1680
      TabIndex        =   0
      Top             =   5880
      Width           =   2295
   End
End
Attribute VB_Name = "GDMagform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
   Call Form_Unload(0)
End Sub

Public Sub Form_Load()
   On Error GoTo errhand
   
'   If DigitizeOn And Not maginit Then '<<<<<<<<<<<digi changes
'      zoom_digi = 1
'      End If
'
   magvis = True
   GDform1.Width = GDMDIform.Width / 2 - 120
   GDMagform.ScaleMode = vbTwips
   GDMagform.Text1.Text = GDMDIform.Label1
   GDMagform.Text3.Text = GDMDIform.Label2
   GDMagform.Left = GDform1.Left + GDform1.Width
   GDMagform.Top = 0
   GDMagform.Height = GDform1.Height
   GDMagform.Width = GDMDIform.Width - GDform1.Width - 200
   Picture1.Left = 60
   Picture1.Width = GDMagform.Width - VScroll1.Width - 240
   Picture1.Height = GDMagform.Height - HScroll1.Height - cmdClose.Height - 800
   formwidth = GDMagform.Width
   formheight = GDMagform.Height
   Image1.BorderStyle = 0
   cmdClose.Left = GDMagform.Width / 2 - cmdClose.Width / 2
   cmdClose.Top = HScroll1.Top + HScroll1.Height + 200
   
   Screen.MousePointer = vbHourglass
   
'   If DigitizeOn Then '<<<<<<<<<<<<digi changes
'
'        If dhdc_digi = 0 Then
'           dhwnd_digi = GDform1.Picture2.hWnd  'GetDesktopWindow ' get desktop window
'           dhdc_digi = GDform1.Picture2.hDC 'dhdc = GetDC(dhwnd)      ' get display device
''           dhwnd_digi = GetDesktopWindow ' get desktop window
''           dhdc_digi = GetDC(dhwnd)      ' get display device
'           End If
'        'blit as much as the magform will handle
'        w = GDMagform.Picture1.Width
'        h = GDMagform.Picture1.Height
'   Else
        'blit the chosen portion of the map to the image control
        'the width, and height of the selected region is W,H:
        W = drag2x - drag1x
        h = drag2y - drag1y
        'upper left corner of selected region is:
        xpic = drag1x
        ypic = drag1y
'        End If
   
   'rescale the Picture2 picturebox to the size of the blit
   GDMagform.Picture2.Width = W * (15 / twipsx)
   GDMagform.Picture2.Height = h * (15 / twipsy)
   
'   If Not DigitizeOn Then '<<<<<<<<<<<<digi changes
   
    'Zooming will be accomplished by using the Stretch property of
    'the Image control.  It is necessary first to transfer the desired
    'portion of the map to the Image control.  This can't be accomplished
    'directly using PaintPicture since Image controls don't have
    'PaintPicture methods.  So need to do it in a roundabout way, viz.,
    'by first PaintPicture'ing the desired portion to an invisible
    'Picture Box (GDMagform.Picture2) on the form.  This picture is
    'then transfered to the ClipBoard.  The contents of the
    'ClipBoard can then be loaded into the Image control
    '(GDMagform.Image1).  (There is probably a
    'more direct way to do this using API, but this is the most
    'straightforward way I can think of doing it without API.)
    'So lets...
    
    'blit this selected region to picture buffer on this form
    GDMagform.Picture2.PaintPicture GDform1.Picture2.Picture, 0, 0, W, h, xpic, ypic, W, h
    
    'save it to the clipboard
    Clipboard.Clear
    Clipboard.SetData GDMagform.Picture2.Image
    
    'transfer the clipboard to the image control
    GDMagform.Image1.Picture = Clipboard.GetData()
    
    'clear the clipboard
    Clipboard.Clear
    
    'clear the picture buffer
    Picture2.Picture = LoadPicture(sEmpty)
    
    'Set the size of the image
    Image1.Height = h
    Image1.Width = W
    Image1.Left = 0
    Image1.Top = 0
        
'        End If
   
   'set the ranges of the scroll bars
   HScroll1.Left = Picture1.Left
   HScroll1.Top = Picture1.Top + Picture1.Height
   HScroll1.Width = Picture1.Width
   VScroll1.Top = Picture1.Top
   VScroll1.Height = Picture1.Height
   VScroll1.Left = Picture1.Left + Picture1.Width
   If Image1.Width <= Picture1.Width Then
      HScroll1.Visible = False
   Else
      HScroll1.Visible = True
      HScroll1.Max = Image1.Width - Picture1.Width
      HScroll1.LargeChange = HScroll1.Max / 30
      HScroll1.SmallChange = HScroll1.Max / 60
      End If
   If Image1.Height <= Picture1.Height Then
      VScroll1.Visible = False
   Else
      VScroll1.Visible = True
      VScroll1.Max = Image1.Height - Picture1.Height
      VScroll1.LargeChange = VScroll1.Max / 30
      VScroll1.SmallChange = VScroll1.Max / 60
      End If
   maginit = False
   Screen.MousePointer = vbDefault
   Exit Sub
   
errhand:
   Screen.MousePointer = vbDefault
   If Err.Number = 6 Then
      MsgBox "Drag region is too big...image will be clipped." _
             & vbCrLf & "Reduce the magnification to recover the entire image.", _
             vbExclamation + vbOKOnly, App.Title
      Resume Next
      Exit Sub
      End If
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          "in module: GDMagform", vbCritical + vbOKOnly, "MapDigitizer"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   Set GDMagform = Nothing
   'reset parameters for fully expanded map
   mag = 0
   magclose = True
   maginit = True
   magvis = False
   
  'reexpand the regular map to fill the boundaries of the program
   GDform1.Width = GDMDIform.ScaleWidth
   
End Sub

Private Sub Form_Resize()
   On Error GoTo errhand

   If maginit = True Then Exit Sub
   'this needs reworking <<<<<<<<<<
   GDform1.Width = GDMagform.ScaleLeft - GDform1.Left
   
   formheight = GDMagform.Height
   formwidth = GDMagform.Width
   
   Picture1.Height = Picture1.Height * (GDMagform.Height) / formheight
   'Picture1.Width = Picture1.Width * (GDMagform.Width) / formwidth
   Picture1.Left = 60
   Picture1.Width = GDMagform.ScaleWidth - VScroll1.Width - 240
   Picture1.Height = GDMagform.ScaleHeight - HScroll1.Height - cmdClose.Height - 800
   
   HScroll1.Left = Picture1.Left
   HScroll1.Top = Picture1.Top + Picture1.Height
   HScroll1.Width = Picture1.Width
   VScroll1.Top = Picture1.Top
   VScroll1.Height = Picture1.Height
   VScroll1.Left = Picture1.Left + Picture1.Width
   If Image1.Width <= Picture1.Width Then
      HScroll1.Visible = False
   Else
      HScroll1.Visible = True
      HScroll1.Max = Image1.Width - Picture1.Width
      HScroll1.LargeChange = HScroll1.Max / 30
      HScroll1.SmallChange = HScroll1.Max / 60
      End If
   If Image1.Height <= Picture1.Height Then
      VScroll1.Visible = False
   Else
      VScroll1.Visible = True
      VScroll1.Max = Image1.Height - Picture1.Height
      VScroll1.LargeChange = VScroll1.Max / 30
      VScroll1.SmallChange = VScroll1.Max / 60
      End If
   cmdClose.Left = GDMagform.Width / 2 - cmdClose.Width / 2
   cmdClose.Top = HScroll1.Top + HScroll1.Height + 200
   
   Exit Sub
   
errhand:
   If Err.Number = 6 Or Err.Number = 0 Then
      'overflow error due to drag region being too big,
      'so ignore it (vscroll.max automatically set to max integer value)
      Resume Next
   ElseIf Err.Number = 380 Then
      'squishing some control, just resume
      Resume Next
   Else
      MsgBox "Error: " & str(Err.Number) & " " & Err.Description _
          & vbCrLf & "in module GDMagform", vbCritical + vbOKOnly
      End If
      
End Sub

Private Sub HScroll1_Change()
   Image1.Left = -HScroll1.Value
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
  'Convert coordinates to pixels
  If mag = 0 Then mag = 1
  Xcoord = X / (twipsx * mag)
  Ycoord = Y / (twipsy * mag)
  
  'determine new corner coordinates
  'upper-left corner coordinates of drag box is:
  x11 = ((LRGeoX - ULGeoX) / pixwi) * drag1x / twipsx + ULGeoX
  y11 = ULGeoY - ((ULGeoY - LRGeoY) / pixhi) * drag1y / twipsy
  
  'lower-right corner coordinates of drag box is:
  x22 = ((LRGeoX - ULGeoX) / pixwi) * drag2x / twipsx + ULGeoX
  y22 = ULGeoY - ((ULGeoY - LRGeoY) / pixhi) * drag2y / twipsy
   
  'Convert pixel coordinates to ITM
  ITMx = ((x22 - x11) / W) * Xcoord * (twipsx / 15) + x11
  ITMy = y11 - ((y11 - y22) / h) * Ycoord * (twipsy / 15)
  
  'Display the ITM coordinates
  GDMagform.Text2 = str(Int(ITMx))
  GDMagform.Text4 = str(Int(ITMy))
   If heights And (RSMethod0 Or RSMethod1 Or RSMethod2) And Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then 'display heights
     kmx = ITMx
     kmy = ITMy
     'Call DTMheight(kmx, kmy, hgt)
     Dim hgt As Integer
     Call DTMheight2(CDbl(kmx), CDbl(kmy), hgt)
     GDMagform.Text6 = str(hgt)
     End If
     
  If Geo And ShowContGeo Then 'also display geo coordinates
        ret = BringWindowToTop(GDGeoFrm.hWnd)
        kmxoo = ITMx: kmyoo = ITMy
        
        If GpsCorrection Then 'wgs84
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
     
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xcoord = X
   Ycoord = Y
   Select Case Button
     Case 1  'left button
     Case 2 'right button
        'right button, load these coordinates into goto coordinates
        GDMDIform.Text5 = GDMagform.Text2
        GDMDIform.Text6 = GDMagform.Text4
        GDMDIform.Text7 = GDMagform.Text6
        'close the magform, record these coordinates and move there
        cmdClose_Click
        shiftmag = True
     Case Else
   End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
      Case "magnifykey"
         mag = mag * 1.1
'         If Not DigitizeOn Then '<<<<<<<<<<<digi changes
            Call magnify
'         Else
'            zoom_digi = zoom_digi * 1.1
'            If zoom_digi > 1 Then Toolbar1.Buttons(2).Enabled = True
'            End If
      Case "demagnifykey"
         mag = mag / 1.1
'         If Not DigitizeOn Then '<<<<<<<<<<<digi changes
            Call demagnify
'         Else
'            zoom_digi = zoom_digi / 1.1
'            If zoom_digi < 1 Then zoom_digi = 1
'            If zoom_digi <= 1 Then Toolbar1.Buttons(2).Enabled = False
'            End If
      Case "printkey"
         PrintMag = True
         PreviewPrint
      Case "closekey"
         Call Form_Unload(0)
      Case Else
   End Select
End Sub

Private Sub VScroll1_Change()
   Image1.Top = -VScroll1.Value
End Sub

Private Sub magnify()
   On Error GoTo errhand

   Image1.Height = Image1.Height * 1.1
   Image1.Width = Image1.Width * 1.1
   If Image1.Width <= Picture1.Width Then
      HScroll1.Visible = False
   Else
      HScroll1.Visible = True
      HScroll1.Max = Image1.Width - Picture1.Width
      HScroll1.LargeChange = HScroll1.Max / 30
      HScroll1.SmallChange = HScroll1.Max / 60
      End If
   If Image1.Height <= Picture1.Height Then
      VScroll1.Visible = False
   Else
      VScroll1.Visible = True
      VScroll1.Max = Image1.Height - Picture1.Height
      VScroll1.LargeChange = VScroll1.Max / 30
      VScroll1.SmallChange = VScroll1.Max / 60
      End If
      
errhand:
End Sub

Private Sub demagnify()
   On Error GoTo errhand

   Image1.Height = Image1.Height / 1.1
   Image1.Width = Image1.Width / 1.1
   If Image1.Width <= Picture1.Width Then
      HScroll1.Visible = False
   Else
      HScroll1.Visible = True
      HScroll1.Max = Image1.Width - Picture1.Width
      lc = HScroll1.Max / 30
      If lc < 1 Then lc = 1
      HScroll1.LargeChange = lc
      sc = HScroll1.Max / 60
      If sc < 1 Then sc = 1
      HScroll1.SmallChange = sc
      End If
   If Image1.Height <= Picture1.Height Then
      VScroll1.Visible = False
   Else
      VScroll1.Visible = True
      VScroll1.Max = Image1.Height - Picture1.Height
      lc = VScroll1.Max / 30
      If lc < 1 Then lc = 1
      VScroll1.LargeChange = lc
      sc = VScroll1.Max / 60
      If sc < 1 Then sc = 1
      VScroll1.SmallChange = sc
      End If

errhand:
End Sub
