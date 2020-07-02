VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GDDigiMagfrm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Magnifyer"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictureBox1 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   0
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   5160
         Top             =   2640
      End
      Begin MSComctlLib.ImageList ImageListMag 
         Left            =   3000
         Top             =   1560
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
               Picture         =   "GDDigiMag,.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GDDigiMag,.frx":01D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GDDigiMag,.frx":03AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GDDigiMag,.frx":0800
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GDDigiMag,.frx":08FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GDDigiMag,.frx":0B90
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GDDigiMag,.frx":10D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GDDigiMag,.frx":122C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar ToolbarMag 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageListMag"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ZoomIn"
            Object.ToolTipText     =   "Zoom In"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ZoomOut"
            Object.ToolTipText     =   "Zoom Out"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.TextBox txtlblMag 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5800
         TabIndex        =   5
         Text            =   "magnified"
         Top             =   70
         Width           =   700
      End
      Begin VB.TextBox txtMagnify 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5160
         TabIndex        =   4
         Text            =   "txtMagnify"
         ToolTipText     =   "Magnification"
         Top             =   70
         Width           =   615
      End
      Begin VB.TextBox txtRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   840
         TabIndex        =   3
         Text            =   "Color under cursor"
         Top             =   40
         Width           =   1695
      End
      Begin VB.PictureBox PictureRGB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   320
         Left            =   2640
         ScaleHeight     =   285
         ScaleWidth      =   2265
         TabIndex        =   2
         Top             =   22
         Width           =   2295
      End
   End
End
Attribute VB_Name = "GDDigiMagfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************
' realtime magnifyer
' the code was written by neowax
' you may use it for your own application
' i wrote this, cause there were no really good magnifyer
' the advantage: it used only 2-6% of the system-resources
' while moving and zooming
' for further information feel free to
' contact me under neowax@uni.de
' please have patience while receiving answer
' *************************************************

Private Sub Form_Load()
'dhwnd_digi = GetDesktopWindow ' get desktop window GDform1.Picture1.hWnd '
'dhdc_digi = GetDC(dhwnd_digi)      ' get display device
DigitizeMagvis = True
DigiMagnify = 300
txtMagnify = DigiMagnify & "%"

   GDform1.Width = GDMDIform.ScaleWidth * 0.5
   GDDigiMagfrm.ScaleMode = vbTwips
   GDDigiMagfrm.left = GDform1.ScaleLeft + GDform1.Width
   GDDigiMagfrm.top = 0
   GDDigiMagfrm.Height = GDform1.Height
   GDDigiMagfrm.Width = GDMDIform.ScaleWidth - GDform1.Width
   PictureBox1.left = 0
   PictureBox1.Width = GDDigiMagfrm.Width '- VScroll1.Width
   PictureBox1.Height = GDDigiMagfrm.Height '- HScroll1.Height
   formwidth = GDDigiMagfrm.Width
   formheight = GDDigiMagfrm.Height
   DigitizeMagInit = False

    AutoRedraw = True
    
    If DigitizePadVis Then
       GDDigitizerfrm.Visible = True
       BringWindowToTop (GDDigitizerfrm.hwnd)
       End If

'  Ret = SetWindowPos(GDDigiMagfrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    
  GDform1.Picture2.SetFocus
'    Show
    On Error Resume Next
    
End Sub

Private Sub Form_Resize()

   On Error GoTo errhand

   If DigitizeMagInit = True Then Exit Sub
   GDform1.Width = GDDigiMagfrm.left - GDform1.ScaleLeft
   GDDigiMagfrm.Width = GDMDIform.ScaleWidth - GDform1.Width
   
   formheight = GDDigiMagfrm.Height
   formwidth = GDDigiMagfrm.Width
   
   PictureBox1.Height = PictureBox1.Height * (GDDigiMagfrm.Height) / formheight
   PictureBox1.left = 0
   PictureBox1.Width = GDDigiMagfrm.Width
   PictureBox1.Height = GDDigiMagfrm.Height
  
   PictureBox1.Width = GDDigiMagfrm.Width - PictureBox1.left
   PictureBox1.Height = GDDigiMagfrm.Height - PictureBox1.top

Exit Sub

errhand:
   If Err.Number = 6 Or Err.Number = 0 Then
      'overflow error due to drag region being too big,
      'so ignore it (vscroll.max automatically set to max integer value)
      Resume Next
   ElseIf Err.Number = 380 Then
      'squishing some control, just resume
      Resume Next
'   Else
'      MsgBox "Error: " & str(Err.Number) & " " & Err.Description _
'          & vbCrLf & "in module GDMagform", vbCritical + vbOKOnly
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call ReleaseDC(dhwnd_digi, dhdc_digi) ' it's important you free tha dc, cause windows may crash
DigitizeMagvis = False
DigitizeMagInit = True
'DigitizeLine = False
'DigitizePoint = False
'DigitizeBeginLine = False
'DigitizeEndLine = False
'DigitizeContour = False
'DigiContourStart = False
'PointStart = False
'DigitizeOn = False
   
GDMDIform.Toolbar1.Buttons(36).value = False
buttonstate&(36) = 0
'If DigitizePadVis Then Unload GDDigitizerfrm

'If DigiLogFileOpened Then
'   DigiLogFileOpened = False
'   Close #Digilogfilnum%
'   Digilogfilnum% = 0
'   End If

'DigitizeContour = False
If hD <> 0 Then
'   DigitizeContour = False
   Call ReleaseDC(0, hD) 'release the dc
   End If
   
'GDMDIform.SliderContour.Visible = False

  'reexpand the regular map to fill the boundaries of the program
   GDform1.Width = GDMDIform.ScaleWidth

End Sub

Private Sub ToolbarMag_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
      Case "ZoomIn"
         DigiMagnify = Max(100, DigiMagnify * 1.1)
      Case "ZoomOut"
         DigiMagnify = Max(100, DigiMagnify / 1.1)
      Case Else
   End Select
   txtMagnify = DigiMagnify & "%"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Timer1_Timer
' Author    : Chaim Keller
' Date      : 1/24/2015
' Purpose   : based on code "magnifier" written by neowax
'---------------------------------------------------------------------------------------
'
Private Sub Timer1_Timer()

Dim R As Long
Dim New_Color As couleur
Dim mag As Double
Dim ier As Long

   On Error GoTo Timer1_Timer_Error

'the following code only works if the picture are in pixels (NO MORE TWIPS!!!!)
R = GDform1.Picture2.Point(nearmouse_digi.x, nearmouse_digi.Y)
New_Color = recupcouleur(R)
GDDigiMagfrm.PictureRGB.BackColor = RGB(New_Color.R, New_Color.V, New_Color.b)
Me.Caption = "Magnifyer: Xpix: " & CLng(nearmouse_digi.x / DigiZoom.LastZoom) & ",  Ypix: " & CLng(nearmouse_digi.Y / DigiZoom.LastZoom) & "               RGB: " & str(New_Color.R) & "," & str(New_Color.V) + "," + str(New_Color.b) '& " X: " & nearmouse_digi.X & ", Y: " & nearmouse_digi.Y '& " ratios: " & nearmouse_digi.X / mouse_digi.X & ", " & nearmouse_digi.Y / mouse_digi.Y

w_digi = PictureBox1.ScaleWidth  ' destination width
h_digi = PictureBox1.ScaleHeight ' destination height

mag = DigiMagnify * 0.01
PictureBox1.Cls  ' clean picturebox

sw_digi = w_digi / mag ' source width
sh_digi = h_digi / mag 'source height
XX = nearmouse_digi.x - sw_digi \ 2 ' x source position (center to destination)
YY = nearmouse_digi.Y - sh_digi \ 2 ' y source position (center to destination)
ier = StretchBlt(PictureBox1.hdc, 0, 0, w_digi, h_digi, GDform1.Picture2.hdc, XX, YY, sw_digi, sh_digi, vbSrcCopy)  'copy picture and strech to picturebox (destination)
If ier = 0 Then 'outside picturebox
   GDMDIform.StatusBar1.Panels(2).Text = "Outside Picture"
   End If

'draw center mark on magnified picture
centerx = w_digi / 2
centery = h_digi / 2
sizecrosshair = w_digi / 50
AA& = PictureBox1.DrawMode
BB& = PictureBox1.DrawStyle
PictureBox1.DrawMode = 7
PictureBox1.DrawStyle = 1
PictureBox1.DrawWidth = 1
PictureBox1.Line (0, centery)-(w_digi, centery), QBColor(15)
PictureBox1.Line (centerx, 0)-(centerx, h_digi), QBColor(15)
PictureBox1.Circle (centerx, centery), sizecrosshair / 2, QBColor(0)
PictureBox1.DrawMode = AA&

   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Form GDDigiMagfrm"

End Sub

Private Sub txtMagnify_Change()
   txtMagnify = DigiMagnify & "%"
   Exit Sub
End Sub

Private Sub txtRGB_Change()
   txtRGB.Text = "Color under cursor"
   Exit Sub
End Sub
