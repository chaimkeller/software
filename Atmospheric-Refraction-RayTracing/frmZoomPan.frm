VERSION 5.00
Object = "{61683A27-FCBC-4C86-BC9D-195B8ADDC7FB}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Begin VB.Form frmZoomPan 
   Caption         =   "Zoom/Pan with PrePaint Events"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMask 
      Caption         =   "Apply Alpha Mask"
      Height          =   720
      Left            =   13440
      TabIndex        =   6
      Top             =   1800
      Width           =   900
   End
   Begin VB.CommandButton cmdSkew 
      Caption         =   "Animate on Y-Axis"
      Height          =   720
      Left            =   13440
      TabIndex        =   5
      Top             =   960
      Width           =   900
   End
   Begin VB.Timer Timer1 
      Left            =   13560
      Top             =   240
   End
   Begin VB.HScrollBar HScrollZoom 
      Height          =   285
      LargeChange     =   20
      Left            =   360
      Max             =   800
      Min             =   10
      TabIndex        =   2
      Top             =   9600
      Value           =   10
      Width           =   13725
   End
   Begin VB.HScrollBar HScrollPan 
      Height          =   285
      LargeChange     =   20
      Left            =   360
      TabIndex        =   1
      Top             =   9120
      Width           =   12180
   End
   Begin VB.VScrollBar VScrollPan 
      Height          =   8475
      LargeChange     =   20
      Left            =   12840
      TabIndex        =   0
      Top             =   360
      Width           =   300
   End
   Begin LaVolpeAlphaImg.AlphaImgCtl AlphaImgCtl1 
      Height          =   8655
      Left            =   240
      Top             =   240
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   15266
      Effects         =   "frmZoomPan.frx":0000
   End
   Begin VB.Label lblZoom 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   495
      Left            =   13680
      TabIndex        =   4
      Top             =   8040
      Width           =   840
   End
   Begin VB.Label lblPan 
      Caption         =   "Label1"
      Height          =   675
      Left            =   13680
      TabIndex        =   3
      Top             =   7080
      Width           =   675
   End
End
Attribute VB_Name = "frmZoomPan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' see cmdMask_Click. Path/Brush-related APIs
Private Declare Function GdipAddPathEllipse Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mFillMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColors As Long, ByRef mCount As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Const pi180 As Double = 3.14159265358979 / 180#
Private m_PtList(2) As POINTAPI
Private m_Mask() As Byte


Private Sub AlphaImgCtl1_DblClick()
    HScrollPan.Value = 0
    HScrollZoom.Value = 100
    VScrollPan.Value = 0
End Sub

Private Sub AlphaImgCtl1_PrePaint(hdc As Long, Left As Long, Top As Long, Width As Long, height As Long, HitTestRgn As Long, Cancel As Boolean)
    If HScrollPan.Enabled Then  ' else animation sample being run
        Dim zoomOffset As Single
        zoomOffset = HScrollZoom.Value / 100
        lblPan.Caption = "h:" & HScrollPan.Value & vbCrLf & "v:" & VScrollPan.Value
        lblZoom.Caption = "Zoom: " & Format(zoomOffset, "Percent")
        Cancel = True ' prevent rendering image, we will be rendering it below
        AlphaImgCtl1.Picture.Render hdc, (AlphaImgCtl1.Width - (AlphaImgCtl1.Width * zoomOffset)) \ 2 + HScrollPan.Value * zoomOffset, _
            (AlphaImgCtl1.height - (AlphaImgCtl1.height * zoomOffset)) \ 2 + VScrollPan.Value * zoomOffset, _
            AlphaImgCtl1.Width * zoomOffset, AlphaImgCtl1.height * zoomOffset, , , , , , _
            AlphaImgCtl1.Effects.AttributesHandle, , AlphaImgCtl1.Effects.EffectsHandle(AlphaImgCtl1.Effect)
    Else
        AlphaImgCtl1.Picture.RenderSkewed hdc, m_PtList(0).x, m_PtList(0).y, m_PtList(1).x, m_PtList(1).y, m_PtList(2).x, m_PtList(2).y
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False
    Me.ScaleMode = vbPixels
    AlphaImgCtl1.SetRedraw = False
    With HScrollPan
        .Min = -AlphaImgCtl1.Width
        .max = -.Min
        .Value = 0
    End With
    With VScrollPan
        .Min = -AlphaImgCtl1.height
        .max = -.Min
        .Value = 0
    End With
    With HScrollZoom
        .Min = 10
        .max = 800
        .Value = 100
    End With
    AlphaImgCtl1.WantPrePostEvents = True
    AlphaImgCtl1.SetRedraw = True
End Sub

Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

    Timer1.Enabled = False
    Me.ScaleMode = vbPixels
    AlphaImgCtl1.SetRedraw = False
    With prjAtmRefMainfm
        If Val(.txtXSize) > 1500 Then
           .Width = .Width + 20 * (Val(.txtXSize) - 1500)
           End If
    End With
    With HScrollPan
        .Min = -AlphaImgCtl1.Width
        .max = -.Min
        .Value = 0
    End With
    With VScrollPan
        .Min = -AlphaImgCtl1.height
        .max = -.Min
        .Value = 0
    End With
    With HScrollZoom
        .Min = 10
        .max = 800
        .Value = 100
    End With
    AlphaImgCtl1.WantPrePostEvents = True
    AlphaImgCtl1.SetRedraw = True

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

End Sub

Private Sub HScrollPan_Change()
    AlphaImgCtl1.Refresh
End Sub

Private Sub HScrollPan_Scroll()
    AlphaImgCtl1.Refresh
End Sub

Private Sub HScrollZoom_Change()
    AlphaImgCtl1.Refresh
End Sub

Private Sub HScrollZoom_Scroll()
    AlphaImgCtl1.Refresh
End Sub

Private Sub VScrollPan_Change()
    AlphaImgCtl1.Refresh
End Sub

Private Sub VScrollPan_Scroll()
    AlphaImgCtl1.Refresh
End Sub

Private Sub cmdSkew_Click()
    If Timer1.Enabled = False Then ' abort if timer already running
        HScrollPan.Value = 0
        HScrollZoom.Value = 100
        VScrollPan.Value = 0
        HScrollPan.Enabled = False
        HScrollZoom.Enabled = False
        VScrollPan.Enabled = False
        cmdMask.Enabled = False
        Timer1.Tag = "0"
        Timer1.Interval = 50
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()

 Timer1.Enabled = False ' turn off for now

 Dim x As Integer
 Dim NewX As Long, NewY As Long
 Dim SinAng1 As Double, CosAng1 As Double
 Dim SinAng2 As Double, SinAng3 As Double
 Dim Zoom As Double

 With AlphaImgCtl1.Picture
    m_PtList(0).x = -(.Width / 2)
    m_PtList(0).y = -(.height / 2)
    m_PtList(1).x = .Width / 2
    m_PtList(1).y = -(.height / 2)
    m_PtList(2).x = -(.Width / 2)
    m_PtList(2).y = (.height / 2)
 End With
 
 Timer1.Tag = Val(Timer1.Tag) + 20
 
 Zoom = Tan(45 * pi180)
 SinAng1 = Sin(90 * pi180)
 CosAng1 = Cos(90 * pi180)
 x = CInt(Timer1.Tag)
 
 SinAng2 = Sin((x + 90) * pi180) * Zoom
 SinAng3 = SinAng1 * Zoom
 
 For x = 0 To 2
    NewX = (m_PtList(x).x * SinAng1 + m_PtList(x).y * CosAng1) * SinAng2
    NewY = (m_PtList(x).y * SinAng1 - m_PtList(x).x * CosAng1) * SinAng3
    m_PtList(x).x = NewX + (AlphaImgCtl1.Picture.Width / 2)
    m_PtList(x).y = NewY + (AlphaImgCtl1.Picture.height / 2)
 Next
 
 
 If Timer1.Tag = "360" Then ' done rotating/flipping
    HScrollPan.Enabled = True
    HScrollZoom.Enabled = True
    VScrollPan.Enabled = True
    cmdMask.Enabled = True
    AlphaImgCtl1.Refresh ' update screen
 Else
    AlphaImgCtl1.Refresh ' update screen
    Timer1.Enabled = True ' continue timer
 End If

End Sub

Private Sub cmdMask_Click()

    ' routine creates a fade effect on the image
    ' the fade occurs center out. The center of image will be nearly fully opaque
    ' and each pixel from the center out will get progressively more transparent in a circular pattern
    ' Here we will use a GDI+ path to help us out vs. calculating it all ourselves

    Dim tMask() As Byte, x As Long
    Dim tImg As GDIpImage, SS As SAVESTRUCT
    Dim hGraphics As Long, hPath As Long, hBrush As Long
    
    If cmdMask.Caption = "Apply Alpha Mask" Then            ' creaet & apply mask
        cmdMask.Caption = "Restore Alpha Mask"              ' change button caption
                                                            ' get current mask & cache
        With AlphaImgCtl1.Picture                           ' create new imgae same size as source
            .AlphaMask m_Mask(), False
            SS.Width = .Width
            SS.height = .height
            SS.ColorDepth = lvicConvert_TrueColor32bpp_ARGB ' we want the alpha channel
        End With
        Set tImg = New GDIpImage
        SavePictureGDIplus Nothing, tImg, , SS              ' create the temp image
        
        GdipCreatePath 0&, hPath                            ' create a temp path
        GdipAddPathEllipse hPath, 0, 0, SS.Width + 5&, SS.height + 5& ' make it elliptical
        GdipCreatePathGradientFromPath hPath, hBrush        ' create a brush from that path
        x = 1&                                              ' set center & gradient colors of path
        GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertRGBtoARGB(vbWhite, 0), x
        GdipSetPathGradientCenterColor hBrush, ConvertRGBtoARGB(vbWhite, 225)
        ' note that colors not important, the alpha channel is what we want
        
        GdipGetImageGraphicsContext tImg.Handle, hGraphics  ' get graphics object from temp image
        GdipFillPath hGraphics, hBrush, hPath               ' fill that img with the path/brush
        GdipDeleteBrush hBrush                              ' clean up
        GdipDeletePath hPath
        GdipDeleteGraphics hGraphics
        
        tImg.AlphaMask tMask(), False                       ' get the alpha mask
        Set tImg = Nothing                                  ' no longer needed
        For x = 0 To UBound(m_Mask)                         ' multiply it against existing mask
            tMask(x) = (1& * m_Mask(x) * tMask(x)) \ 255
        Next                                                ' apply new mask
        AlphaImgCtl1.Picture.AlphaMask tMask(), True
        
    Else    ' restore original mask
        cmdMask.Caption = "Apply Alpha Mask"                ' restore cached mask
        AlphaImgCtl1.Picture.AlphaMask m_Mask(), True
        Erase m_Mask()
    End If
    AlphaImgCtl1.Refresh
    
End Sub

