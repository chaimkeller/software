Attribute VB_Name = "PlotGraphics"
Option Explicit

'The basis of the code below is taken from the www.FreeVBcode.com

'API declarations font handling
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public LineWidth As Integer, Fitting As Boolean, SaveFormat As Boolean, LastSelected%, FitWizard As Boolean

Public Const VK_DOWN = &H28
Public Const VK_UP = &H26
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_TAB = &H9
Public Const LF_FACESIZE = 32

Public Const MaxNumOverplotFiles = 1000 'maximum number of files allowed for overplotting

Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
  'lfFaceName(1 To LF_FACESIZE) As Byte 'original type declaration from API Viewer
End Type

Public Enum FontWeights
  FW_DONTCARE = 0
  FW_THIN = 100
  FW_EXTRALIGHT = 200
  FW_ULTRALIGHT = 200
  FW_LIGHT = 300
  FW_NORMAL = 400
  FW_REGULAR = 400
  FW_MEDIUM = 500
  FW_SEMIBOLD = 600
  FW_DEMIBOLD = 600
  FW_BOLD = 700
  FW_EXTRABOLD = 800
  FW_ULTRABOLD = 800
  FW_HEAVY = 900
  FW_BLACK = 900
End Enum

Public Enum FontCharSets
  ANSI_CHARSET = 0
  ARABIC_CHARSET = 178
  BALTIC_CHARSET = 186
  CHINESEBIG5_CHARSET = 136
  DEFAULT_CHARSET = 1
  EASTEUROPE_CHARSET = 238
  GB2312_CHARSET = 134
  GREEK_CHARSET = 161
  HANGEUL_CHARSET = 129
  HEBREW_CHARSET = 177
  JOHAB_CHARSET = 130
  MAC_CHARSET = 77
  OEM_CHARSET = 255
  RUSSIAN_CHARSET = 204
  SHIFTJIS_CHARSET = 128
  SYMBOL_CHARSET = 2
  THAI_CHARSET = 222
  TURKISH_CHARSET = 162
End Enum

Public Enum FontOutPrecisions
  OUT_DEFAULT_PRECIS = 0
  OUT_STRING_PRECIS = 1
  OUT_CHARACTER_PRECIS = 2
  OUT_STROKE_PRECIS = 3
  OUT_TT_PRECIS = 4
  OUT_DEVICE_PRECIS = 5
  OUT_RASTER_PRECIS = 6
  OUT_TT_ONLY_PRECIS = 7
  OUT_OUTLINE_PRECIS = 8
End Enum

Public Enum FontClipPrecisions
  CLIP_DEFAULT_PRECIS = 0
  CLIP_CHARACTER_PRECIS = 1
  CLIP_STROKE_PRECIS = 2
  CLIP_LH_ANGLES = 16
  CLIP_TT_ALWAYS = 32
  CLIP_EMBEDDED = 128
  CLIP_TO_PATH = 4097
End Enum

Public Enum FontQuality
  ANTIALIASED_QUALITY = 4
  DEFAULT_QUALITY = 0
  DRAFT_QUALITY = 1
  NONANTIALIASED_QUALITY = 3
  PROOF_QUALITY = 2
End Enum

Public Enum FontPitch
  DEFAULT_PITCH = 0
  FIXED_PITCH = 1
  VARIABLE_PITCH = 2
End Enum

Public Enum FontFamily
  FF_DONTCARE = 0
  FF_ROMAN = 16
  FF_SWISS = 32
  FF_MODERN = 48
  FF_SCRIPT = 64
  FF_DECORATIVE = 80
End Enum

'/////////////////////here's more info for LOGFONT///////////////////////////
'http://www.jasinskionline.com/windowsapi/ref/l/logfont.html
'lfHeight
'    The height of the font's character cell, in logical units (also known as the em height). If positive, the font mapper converts this value directly into device units and matches it with the cell height of the possible fonts. If 0, the font mapper uses a default character height. If negative, the font mapper converts the absolute value into device units and matches it with the character height of the possible fonts.
'lfWidth
'    The average width of the font's characters. If 0, the font mapper tries to determine the best value.
'lfEscapement
'    The angle between the font's baseline and escapement vectors, in units of 1/10 degrees. Windows 95, 98: This must be equal to lfOrientation.
'lfOrientation
'    The angle between the font's baseline and the device's x-axis, in units of 1/10 degrees. Windows 95, 98: This must be equal to lfEscapement.
'lfWeight
'    One of the following flags specifying the boldness (weight) of the font:
'
'    FW_DONTCARE
'        Default weight.
'    FW_THIN
'        Thin weight.
'    FW_EXTRALIGHT
'        Extra-light weight.
'    FW_ULTRALIGHT
'        Same as FW_EXTRALIGHT.
'    FW_LIGHT
'        Light weight.
'    FW_NORMAL
'        Normal weight.
'    FW_REGULAR
'        Same as FW_NORMAL.
'    FW_MEDIUM
'        Medium weight.
'    FW_SEMIBOLD
'        Semi-bold weight.
'    FW_DEMIBOLD
'        Same As FW_SEMIBOLD.
'    FW_BOLD
'        Bold weight.
'    FW_EXTRABOLD
'        Extra-bold weight.
'    FW_ULTRABOLD
'        Same as FW_EXTRABOLD.
'    FW_HEAVY
'        Heavy weight.
'    FW_BLACK
'        Same as FW_HEAVY.
'
'lfItalic
'    A non-zero value if the font is italicized, 0 if not.
'lfUnderline
'    A non-zero value if the font is underlined, 0 if not.
'lfStrikeOut
'    A non-zero value if the font is striked out, 0 if not.
'lfCharSet
'    Exactly one of the following flags specifying the character set of the font:
'
'    ANSI_CHARSET
'        ANSI character set.
'    ARABIC_CHARSET
'        Windows NT, 2000: Arabic character set.
'    BALTIC_CHARSET
'        Windows 95, 98: Baltic character set.
'    CHINESEBIG5_CHARSET
'        Chinese Big 5 character set.
'    DEFAULT_CHARSET
'        Default character set.
'    EASTEUROPE_CHARSET
'        Windows 95, 98: Eastern European character set.
'    GB2312_CHARSET
'        GB2312 character set.
'    GREEK_CHARSET
'        Windows 95, 98: Greek character set.
'    HANGEUL_CHARSET
'        HANDEUL character set.
'    HEBREW_CHARSET
'        Windows NT, 2000: Hebrew character set.
'    JOHAB_CHARSET
'        Windows 95, 98: Johab character set.
'    MAC_CHARSET
'        Windows 95, 98: Mac character set.
'    OEM_CHARSET
'        Original equipment manufacturer (OEM) character set.
'    RUSSIAN_CHARSET
'        Windows 95, 98: Russian character set.
'    SHIFTJIS_CHARSET
'        ShiftJis character set.
'    SYMBOL_CHARSET
'        Symbol character set.
'    THAI_CHARSET
'        Windows NT, 2000: Thai character set.
'    TURKISH_CHARSET
'        Windows 95, 98: Turkish character set.
'
'lfOutPrecision
'    Exactly one of the following flags specifying the desired precision (closeness of the match) between the logical font ideally described by the structure and the actual logical font. This value is used by the font mapper to produce the logical font.
'
'    OUT_DEFAULT_PRECIS
'        The default font mapping behavior.
'    OUT_DEVICE_PRECIS
'        Choose a device font if there are multiple fonts in the system with the same name.
'    OUT_OUTLINE_PRECIS
'        Windows NT, 2000: Choose a TrueType or other outline-based font.
'    OUT_RASTER_PRECIS
'        Choose a raster font if there are multiple fonts in the system with the same name.
'    OUT_STRING_PRECIS
'        Raster font (used for enumeration only).
'    OUT_STROKE_PRECIS
'        Windows 95, 98: Vector font (used for enumeration only). Windows NT, 2000: TrueType, outline-based, or vector font (used for enumeration only).
'    OUT_TT_ONLY_PRECIS
'        Choose only a TrueType font.
'    OUT_TT_PRECIS
'        Choose a TrueType font if there are multiple fonts in the system with the same name.
'
'lfClipPrecision
'    Exactly one of the following flags specifying the clipping precision to use when the font's characters must be clipped:
'
'    CLIP_DEFAULT_PRECIS
'        The default clipping behavior.
'    CLIP_EMBEDDED
'        This flag must be set for an embedded read-only font.
'    CLIP_LH_ANGLES
'        The direction of any rotations is determined by the coordinate system (or else all rotations are counterclockwise).
'    CLIP_STROKE_PRECIS
'        Raster, vector, or TrueType font (used for enumeration only).
'
'lfQuality
'    Exactly one of the following flags specifying the output quality of the logical font as compared to the ideal font:
'
'    ANTIALIASED_QUALITY
'        Windows 95, 98, NT 4.0 or later, 2000: The font is always antialiased if possible.
'    DEFAULT_QUALITY
'        The default quality: the appearance of the font does not matter.
'    DRAFT_QUALITY
'        The appearance of the font is less important then in PROOF_QUALITY.
'    NONANTIALIASED_QUALITY
'        Windows 95, 98, NT 4.0 or later, 2000: The font is never antialiased.
'    PROOF_QUALITY
'        The quality of the appearance of the font is more important than exactly matching the specified font attributes.
'
'lfPitchAndFamily
'    A bitwise OR combination of exactly one *_PITCH flag specifying the pitch of the font and exactly one FF_* flag specifying the font face family of the font:
'
'    DEFAULT_PITCH
'        The default pitch.
'    FIXED_PITCH
'        Fixed pitch.
'    VARIABLE_PITCH
'        Variable pitch.
'    FF_DECORATIVE
'        Showy, decorative font face.
'    FF_DONTCARE
'        Do not care about the font face.
'    FF_MODERN
'        Modern font face (monospaced, sans serif font).
'    FF_ROMAN
'        Roman font face (proportional-width, serif font).
'    FF_SCRIPT
'        Script font face which imitates script handwriting.
'    FF_SWISS
'        Swiss font face (proportional-width, sans serif font).
'
'lfFaceName
'    The name of the font face to use. This string must be terminated with a null character.

'Constant Definitions
'
'Const FW_DONTCARE = 0
'Const FW_THIN = 100
'Const FW_EXTRALIGHT = 200
'Const FW_ULTRALIGHT = 200
'Const FW_LIGHT = 300
'Const FW_NORMAL = 400
'Const FW_REGULAR = 400
'Const FW_MEDIUM = 500
'Const FW_SEMIBOLD = 600
'Const FW_DEMIBOLD = 600
'Const FW_BOLD = 700
'Const FW_EXTRABOLD = 800
'Const FW_ULTRABOLD = 800
'Const FW_HEAVY = 900
'Const FW_BLACK = 900
'Const ANSI_CHARSET = 0
'Const ARABIC_CHARSET = 178
'Const BALTIC_CHARSET = 186
'Const CHINESEBIG5_CHARSET = 136
'Const DEFAULT_CHARSET = 1
'Const EASTEUROPE_CHARSET = 238
'Const GB2312_CHARSET = 134
'Const GREEK_CHARSET = 161
'Const HANGEUL_CHARSET = 129
'Const HEBREW_CHARSET = 177
'Const JOHAB_CHARSET = 130
'Const MAC_CHARSET = 77
'Const OEM_CHARSET = 255
'Const RUSSIAN_CHARSET = 204
'Const SHIFTJIS_CHARSET = 128
'Const SYMBOL_CHARSET = 2
'Const THAI_CHARSET = 222
'Const TURKISH_CHARSET = 162
'Const OUT_DEFAULT_PRECIS = 0
'Const OUT_DEVICE_PRECIS = 5
'Const OUT_OUTLINE_PRECIS = 8
'Const OUT_RASTER_PRECIS = 6
'Const OUT_STRING_PRECIS = 1
'Const OUT_STROKE_PRECIS = 3
'Const OUT_TT_ONLY_PRECIS = 7
'Const OUT_TT_PRECIS = 4
'Const CLIP_DEFAULT_PRECIS = 0
'Const CLIP_EMBEDDED = 128
'Const CLIP_LH_ANGLES = 16
'Const CLIP_STROKE_PRECIS = 2
'Const ANTIALIASED_QUALITY = 4
'Const DEFAULT_QUALITY = 0
'Const DRAFT_QUALITY = 1
'Const NONANTIALIASED_QUALITY = 3
'Const PROOF_QUALITY = 2
'Const DEFAULT_PITCH = 0
'Const FIXED_PITCH = 1
'Const VARIABLE_PITCH = 2
'Const FF_DECORATIVE = 80
'Const FF_DONTCARE = 0
'Const FF_MODERN = 48
'Const FF_ROMAN = 16
'Const FF_SCRIPT = 64
'Const FF_SWISS = 32


'public constants used in forms
Public dPlot() As Double
Public udtMyGraphLayout As GRAPHIC_LAYOUT

'public constants used in dragging and plotting
Public drag1x As Single, drag1y As Single, EndPlot As Boolean
Public drag2x As Single, drag2y As Single, dragbegin As Boolean
Public Xo As Single, Yo As Single, PlotInfoCancel As Boolean
Public XMin As Single, YMin As Single, XMax As Single, YMax As Single
Public drm%, drs%, drw%, PlotForm() As Integer, PlotAll As Boolean
Public YMin0 As Double, YRange0 As Double, XMin0 As Double, XRange0 As Double
Public Files() As String, FilForm(4, 11) As Integer, ReSized As Boolean
Public numfiles%, PlotInfo() As String, RecordSize() As Long, direct$, numSelected% ', numPlotInfo%
Public numFilesToPlot%, directPlot$, JKHplotVis As Boolean, ScreenDump As Boolean
Public maxFilesToPlot%, numRowsToNow%, PlotInfofrmVis As Boolean
Public DefaultFileType%, dirWordpad As String, PolyDeg As Integer
Public Const sEmpty As String = ""
Public XTitle$, YTitle$, Title$

'declaration of UDT's (User Defined Types)
Public Type GRAPHIC_LAYOUT
  XTitle As String 'title X-axis
  YTitle As String 'title Y-axis
  Title As String 'chart title
  blnOrigin As Boolean 'origin is included for only pos/neg values when true
  blnGridLine As Boolean 'Gridlines are shown when true
  lStart As Long 'index of start x-Range
  lEnd As Long 'index of end x-Range
  asX As Double 'trace in array to function as "X-value"
  asY() As Variant 'Y-traces to plot
  DrawTrace() As DRAWN_AS
  X0 As Double 'minimum value of domain X-values to draw
  X1 As Double 'maximum value of domain X-values to draw
  Y0 As Double 'minimum value of domain Y-values to draw
  Y1 As Double 'maximum value of domain Y-values to draw
End Type

Public Enum DRAWN_AS
  AS_POINT
  AS_CONLINE
  AS_BAR
  AS_DASH
  AS_DOT
  AS_DASHDOT
  AS_DASHDOTDOT
  AS_CIRCLE
  AS_FILLEDCIRCLE
End Enum

Public Type COORDINATE
  X As Single
  Y As Single
End Type


  

'public declaration of screen variables - in twips
Public twp_XLeftMargin As Single 'left margin
Public twp_XRightMargin As Single 'right margin
Public twp_YTopMargin As Single 'top margin
Public twp_YBottomMargin As Single 'bottom margin
Public twp_YRange As Single 'full Y-Range
Public twp_XRange As Single 'full X-Range

'public declaration of value variables - in their own units
Public val_XMin As Double 'minimum value X
Public val_XRange As Double 'full X-Range X-values
Public val_YMin As Double 'minimum value Y
Public val_YRange As Double 'full Y-Range Y-values
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib _
    "User32" Alias "SendMessageA" (ByVal hWnd As _
    Long, ByVal wMsg As Long, ByVal wParam As _
    Long, lParam As Any) As Long

Private Const WM_USER = &H400
Private Const SB_GETRECT = (WM_USER + 10)

'---------------------------------------------------------------------------------------
' Procedure : PanelText
' Author    : chaim
' Date      : 6/3/2020
' Purpose   : source: https://stackoverflow.com/questions/58882461/visual-basic-6-add-backcolor-to-statusbar-panel
'             change color of background and font for statusbars
'---------------------------------------------------------------------------------------
'
Public Sub StatusBarPanelText(sb As StatusBar, Pic As PictureBox, index As Long, aText As String, bkColor As Long, fgColor As Long, lAlign As Integer)
    Dim R As RECT

    SendMessage sb.hWnd, SB_GETRECT, index - 1, R

    With Pic
        Set .Font = sb.Font
        .Move 0, 0, (R.Right - R.Left + 2) * Screen.TwipsPerPixelX, (R.Bottom - R.Top) * Screen.TwipsPerPixelY
        .BackColor = bkColor
        .Cls
        .ForeColor = fgColor
        .CurrentY = (.Height - .TextHeight(aText)) \ 2

        Select Case lAlign
            Case 0      ' Left Justified
                .CurrentX = 0
            Case 1      ' Right Justified
                .CurrentX = .Width - .TextWidth(aText) - Screen.TwipsPerPixelX * 2
            Case 2      ' Centered
                .CurrentX = (.Width - .TextWidth(aText)) \ 2
        End Select

        Pic.Print aText
        sb.Panels(index).Text = aText
        sb.Panels(index).Picture = .Image
    End With
End Sub



Public Sub Plot(frmSpec As Form, dArSpec() As Double, udtLayoutSpec As GRAPHIC_LAYOUT)

'/////////////////Plot Layout Definitions///////////////////////////////////

'declaration of screen variables - in twips
Dim twp_YTick As Single 'size of Y tick
Dim twp_XTick As Single 'size of X tick
Dim twp_Y0 As Single 'Y0: origin (as in 1st quadrant)
Dim twp_X0 As Single 'X0: origin (as in 1st quadrant)
Dim twp_Y0Tr As Single 'transferrred Y0: origin (as in 4 quadrants)
Dim twp_X0Tr As Single 'transferrred X0: origin (as in 4 quadrants)


'declaration of value variables - in their own units
Dim val_XMax As Double 'maximum value X
Dim val_YMax As Double 'maximum value Y
Dim val_X As Double 'value of X-value
Dim val_Y As Double 'value of Y-value

'declaration of dimensionless variables (ratios)
Dim XRatio As Double 'quotient of val_X and val_XRange
Dim YRatio As Double 'quotient of val_Y and val_YRange
Dim NumYTicks As Integer 'number of ticks Y-axis
Dim NumXTicks As Integer 'number of ticks X-axis

'declaration of general variables
Dim nI As Integer 'counter
Dim nTrace As Variant 'the traces to be plotted
Dim clr_Plot(23) As Long 'array with colors
Dim udtFont As LOGFONT 'to create a logical font type
Dim lHandleFont As Long 'handle for new (logical) font
Dim lOldFont As Long 'handle of old font
Dim lRetVal As Long 'acts for storing return value

'font sizes for chart labels
Dim XAxis_font_size As Integer
Dim YAxis_font_size As Integer
Dim Title_font_size As Integer
XAxis_font_size = Val(frmSetCond.txtTitleXfont.Text)
YAxis_font_size = Val(frmSetCond.txtTitleYfont.Text)
Title_font_size = Val(frmSetCond.txtTitlefont.Text)

On Error GoTo errhand

'*************  initialise screen  ******************
  'screen: clear and define drawwidth
  frmSpec.Cls
  frmSpec.DrawWidth = 1
  
  'screenRange, screenorigin and X/Y-ticks
  twp_XTick = 80
  twp_YTick = 80
  twp_XLeftMargin = 1000
  twp_XRightMargin = 1000
  twp_YTopMargin = 1000
  twp_YBottomMargin = 1400 '500 //////////////changes 11/23/19
  twp_Y0 = frmSpec.ScaleHeight - twp_YBottomMargin
  twp_X0 = twp_XLeftMargin
  twp_YRange = frmSpec.ScaleHeight - twp_YBottomMargin - twp_YTopMargin
  twp_XRange = frmSpec.ScaleWidth - twp_XLeftMargin - twp_XRightMargin
  
  'font (and colors) defaults
  frmSpec.Font.Name = "Ariel"
  frmSpec.Font.Size = 8
  If Val(frmSetCond.txtAxisLabelSize.Text) <> 0 Then
     frmSpec.Font.Size = Val(frmSetCond.txtAxisLabelSize.Text)
     End If
  frmSpec.Font.Bold = True
  frmSpec.Font.Italic = False
  clr_Plot(0) = RGB(100, 20, 0) 'brown
  clr_Plot(1) = RGB(0, 0, 255) 'blue
  clr_Plot(2) = RGB(255, 0, 0) 'red
  clr_Plot(3) = RGB(255, 255, 0) 'green
  clr_Plot(4) = RGB(93, 255, 201)
  clr_Plot(5) = RGB(0, 255, 255)
  clr_Plot(6) = RGB(210, 25, 210)
  clr_Plot(7) = RGB(255, 255, 255)
  clr_Plot(8) = RGB(255, 0, 115)
  clr_Plot(9) = RGB(0, 0, 115)
  clr_Plot(10) = QBColor(0) 'black
  clr_Plot(11) = QBColor(1) 'blue
  clr_Plot(12) = QBColor(2) 'green
  clr_Plot(13) = QBColor(3) 'cyan
  clr_Plot(14) = QBColor(4) 'red
  clr_Plot(15) = QBColor(5) 'magenta
  clr_Plot(16) = QBColor(6) 'yellow
  clr_Plot(17) = QBColor(8) 'gray
  clr_Plot(18) = QBColor(9) 'light blue
  clr_Plot(19) = QBColor(10) 'light green
  clr_Plot(20) = QBColor(11) 'light cyan
  clr_Plot(21) = QBColor(12) 'light red
  clr_Plot(22) = QBColor(13) 'light magenta
  clr_Plot(23) = QBColor(14) 'light yellow
  Dim PlotColor As Long
  
  'logical font
  With udtFont
    .lfEscapement = 200
    .lfFaceName = "Arial" & Chr$(0)
    .lfHeight = (9 * -20) / Screen.TwipsPerPixelY
  End With
    

'*************  determine Xmin, Xmax, Ymin and Ymax  *************
'Xmin
val_XMin = udtLayoutSpec.X0
If val_XMin > 0 And udtLayoutSpec.blnOrigin = True Then
  val_XMin = 0
  If XMin0 > 0 Then
     XMin0 = 0
     frmSetCond.txtValueX0 = 0
     End If
End If
'if val_Xmin<0 lower twp_XleftMargin to show more in window
If val_XMin < 0 Then
'  twp_XLeftMargin = 600
  twp_X0 = twp_XLeftMargin
  twp_XRange = frmSpec.ScaleWidth - twp_XLeftMargin - twp_XRightMargin
End If

'Xmax
val_XMax = udtLayoutSpec.X1
If val_XMax < 0 And udtLayoutSpec.blnOrigin = True Then
  val_XMax = 0
  If XRange0 < 0 Then
     XRange0 = 0
     frmSetCond.txtValueX1 = 0
     End If
End If
If val_XMax = val_XMin Then
  val_XMin = val_XMin - 1
  val_XMax = val_XMax + 1
End If
Dim temp As Double
If val_XMax < val_XMin Then
   temp = val_XMin
   val_XMin = val_XMax
   val_XMax = temp
   End If
val_XRange = val_XMax - val_XMin

'Ymin
val_YMin = udtLayoutSpec.Y0
If val_YMin > 0 And udtLayoutSpec.blnOrigin = True Then
  val_YMin = 0
  If YMin0 > 0 Then
     YMin0 = 0
     frmSetCond.txtValueY0 = 0
     End If
End If

'Ymax
val_YMax = udtLayoutSpec.Y1
If val_YMax < 0 And udtLayoutSpec.blnOrigin = True Then
  val_YMax = 0
  If YRange0 < 0 Then
     YRange0 = 0
     frmSetCond.txtValueY1 = 0
     End If
End If
If val_YMax = val_YMin Then
  val_YMin = val_YMin - 1
  val_YMax = val_YMax + 1
End If
If val_YMax < val_YMin Then
   temp = val_YMin
   val_YMin = val_YMax
   val_YMax = temp
   End If
val_YRange = val_YMax - val_YMin
  

'*************  prepare SpanX and twp_X0Tr  *****************
'determine Pl_SpanX
Dim Pl_SpanX As Double 'span between two ticks in own units
Dim nExp As Integer 'help to determine Pl_SpanX and Pl_SpanY

nExp = 0
If (val_XMax - val_XMin) < 1 And (val_XMax - val_XMin) > 0 Then
  Do While (val_XMax - val_XMin) < 1
    nExp = nExp + 1
    val_XMax = val_XMax * 10
    val_XMin = val_XMin * 10
  Loop
  Pl_SpanX = 10 ^ (-nExp)
  val_XMax = val_XMax * 10 ^ (-nExp) 'correct val_Xmax to original value
  val_XMin = val_XMin * 10 ^ (-nExp) 'correct val_Xmin to original value
Else
  Pl_SpanX = 1
  Do While val_XRange / Pl_SpanX > 20
    If val_XRange / Pl_SpanX > 20 Then
      Pl_SpanX = Pl_SpanX * 2
    End If
    If val_XRange / Pl_SpanX > 20 Then
      Pl_SpanX = Pl_SpanX * 2.5
    End If
    If val_XRange / Pl_SpanX > 20 Then
      Pl_SpanX = Pl_SpanX * 2
    End If
  Loop
End If

'determine twp_X0Tr (Translated twp_X0; twp_X0 is position origin in twips)
If val_XMin < 0 And val_XMax > 0 Then 'positive and negative X-values
  twp_X0Tr = twp_X0 - (val_XMin / val_XRange) * twp_XRange ' "-" because of negative value val_Xmin!
ElseIf val_XMin < 0 And val_XMax <= 0 Then 'only negative values X-values
  twp_X0Tr = twp_X0 + twp_XRange 'axis at end of X-Range
Else
  End If
  
If twp_X0Tr < twp_X0 Then twp_X0Tr = twp_X0


'*************  prepare SpanY and twp_Y0Tr  *****************
'determine Pl_SpanY
Dim Pl_SpanY As Double 'span between two ticks in own units

nExp = 0
If (val_YMax - val_YMin) < 1 And (val_YMax - val_YMin) > 0 Then
  Do While (val_YMax - val_YMin) < 1
    nExp = nExp + 1
    val_YMax = val_YMax * 10
    val_YMin = val_YMin * 10
  Loop
  Pl_SpanY = 10 ^ (-nExp)
  val_YMax = val_YMax * 10 ^ (-nExp)
  val_YMin = val_YMin * 10 ^ (-nExp)
Else
  Pl_SpanY = 1
  Do While val_YRange / Pl_SpanY > 20
    If val_YRange / Pl_SpanY > 20 Then
      Pl_SpanY = Pl_SpanY * 2
    End If
    If val_YRange / Pl_SpanY > 20 Then
      Pl_SpanY = Pl_SpanY * 2.5
    End If
    If val_YRange / Pl_SpanY > 20 Then
      Pl_SpanY = Pl_SpanY * 2
    End If
  Loop
End If

'determine twp_Y0Tr (Translated twp_Y0; twp_Y0 is position origin in twips)
If val_YMin < 0 And val_YMax > 0 Then 'positive and negative Y-values
  twp_Y0Tr = twp_Y0 + (val_YMin / val_YRange) * twp_YRange
ElseIf val_YMin < 0 And val_YMax <= 0 Then 'only negative values Y-values
  twp_Y0Tr = twp_Y0 - twp_YRange
Else '1st quadrant is shown
End If


'************  plot Y-gridlines  (vertical)  **************

Dim Dummy1 As Double
Dim Dummy2 As Double
Dim OffSetX As Double
Dim twp_OffSetX As Double
Dim twp_XTickRange As Double

Dummy1 = Int(val_XMin / Pl_SpanX)
Dummy2 = Int(val_XMax / Pl_SpanX)
If Dummy1 = Dummy2 Then Dummy2 = Dummy1 + 1
NumXTicks = Dummy2 - Dummy1
OffSetX = (val_XMin - Pl_SpanX * Int(val_XMin / Pl_SpanX)) 'offsetX in own units
twp_OffSetX = twp_XRange * OffSetX / val_XRange 'offsetX in twips
Dummy2 = (val_XMax - Pl_SpanX * Int(val_XMax / Pl_SpanX)) 'difference between val_Xmax and highest Ytick lable
Dummy2 = twp_XRange * Dummy2 / val_XRange 'and now in twips
twp_XTickRange = twp_XRange * ((twp_XRange + twp_OffSetX - Dummy2) / twp_XRange) / NumXTicks

If udtLayoutSpec.blnGridLine = True Then
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = gridline
    frmSpec.Line (twp_X0, twp_Y0)-(twp_X0, twp_Y0 - twp_YRange), &H80000016
  End If
  For nI = 1 To NumXTicks
    frmSpec.Line (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0)- _
    (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0 - twp_YRange), &H80000016
  Next nI
End If


'************  plot X-gridlines  (horizontal)  **************

Dim OffSetY As Double
Dim twp_OffSetY As Double
Dim twp_YTickRange As Double

Dummy1 = Int(val_YMin / Pl_SpanY)
Dummy2 = Int(val_YMax / Pl_SpanY)
NumYTicks = Dummy2 - Dummy1 - 1
If NumYTicks = 0 Then
  NumYTicks = 1
End If
OffSetY = Pl_SpanY - (val_YMin - Pl_SpanY * Int(val_YMin / Pl_SpanY))
twp_OffSetY = twp_YRange * OffSetY / val_YRange
Dummy2 = (val_YMax - Pl_SpanY * Int(val_YMax / Pl_SpanY))
Dummy2 = twp_YRange * Dummy2 / val_YRange
twp_YTickRange = (twp_YRange - twp_OffSetY - Dummy2) / NumYTicks
If (Int(val_YMax / Pl_SpanY) - Int(val_YMin / Pl_SpanY)) = 1 Then
  NumYTicks = 0
End If 'otherwise labeling incorrect

If udtLayoutSpec.blnGridLine = True Then 'plot gridlines
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline
    frmSpec.Line (twp_X0, twp_Y0)-(twp_X0 + twp_XRange, twp_Y0), &H80000016
  End If
  For nI = 0 To NumYTicks 'rest of gridlines
    frmSpec.Line (twp_X0, twp_Y0 - twp_OffSetY - nI * twp_YTickRange)- _
    (twp_X0 + twp_XRange, twp_Y0 - twp_OffSetY - nI * twp_YTickRange), &H80000016
  Next nI
End If
  
  
'*************  plot datapoints for every trace  *******************
Dim nNumTraces As Integer
Dim X0 As Double, Y0 As Double, store_style%

On Error GoTo errhand

nNumTraces = 0
For Each nTrace In udtLayoutSpec.asY()
  nNumTraces = nNumTraces + 1
  Select Case udtLayoutSpec.DrawTrace(nNumTraces - 1)
  
  Case AS_CONLINE
    'find value starting point
    val_X = dArSpec(nTrace - 1, 0, udtLayoutSpec.lStart)
    val_Y = dArSpec(nTrace - 1, 1, udtLayoutSpec.lStart)
    XRatio = (val_X - val_XMin) / val_XRange
    YRatio = (val_Y - val_YMin) / val_YRange
    frmSpec.CurrentX = twp_X0 + XRatio * twp_XRange
    frmSpec.CurrentY = twp_Y0 - YRatio * twp_YRange
    'find rest
'     For nI = udtLayoutSpec.lStart + 1 To udtLayoutSpec.lEnd Step 1
     For nI = 0 To RecordSize(nNumTraces - 1) - 1 Step 1
      PlotColor = Val(PlotInfo(2, nTrace - 1))
      Select Case PlotColor
         Case 0 'automatic
            PlotColor = clr_Plot(nNumTraces Mod 23)
         Case 1 'black
            PlotColor = QBColor(0)
         Case 2 'blue
            PlotColor = QBColor(1)
         Case 3 'green
            PlotColor = QBColor(2)
         Case 4 'cyan
            PlotColor = QBColor(3)
         Case 5 'red
            PlotColor = QBColor(4)
         Case 6 'magneta
            PlotColor = QBColor(5)
         Case 7 'yellow
            PlotColor = QBColor(6)
         Case 8 'gray
            PlotColor = QBColor(8)
         Case 9 'light blue
            PlotColor = QBColor(9)
      End Select
      val_X = dArSpec(nTrace - 1, 0, nI)
      val_Y = dArSpec(nTrace - 1, 1, nI)
      
      'look for abrupt ends
      If nI = 0 Then 'if nI = udtLayoutSpec.lStart + 1 Then
         X0 = val_X
         Y0 = val_Y
      Else
         If (X0 > 0 And val_X < X0) Or _
            (X0 < 0 And val_X > X0) And _
            val_X = 0 And val_Y = 0 Then
            Exit For 'abrupt end reaached
         Else
            X0 = val_X
            Y0 = val_Y
            End If
         End If
            
      frmSpec.DrawWidth = Val(PlotInfo(9, nTrace - 1))
      
      XRatio = (val_X - val_XMin) / val_XRange
      YRatio = (val_Y - val_YMin) / val_YRange
      frmSpec.Line -(twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), PlotColor
      
    Next nI
    'clear lines outside drawing pane
    frmSpec.Line (twp_X0, twp_Y0 - twp_YRange - 10)-(twp_X0 + twp_XRange, 0), frmSpec.BackColor, BF 'above drawing pane
    frmSpec.Line (twp_X0, twp_Y0 + 20)-(frmSpec.Width, frmSpec.Height), frmSpec.BackColor, BF 'below drawing pane
    frmSpec.Line (twp_X0 + twp_XRange + 20, frmSpec.Height)-(frmSpec.Width, 0), frmSpec.BackColor, BF 'below drawing pane
    frmSpec.Line (0, 0)-(twp_XLeftMargin - 20, frmSpec.Height), frmSpec.BackColor, BF 'below drawing pane

  Case AS_BAR
'    For nI = udtLayoutSpec.lStart To udtLayoutSpec.lEnd Step 1
    For nI = 0 To RecordSize(nNumTraces - 1) - 1 Step 1
      PlotColor = Val(PlotInfo(2, nTrace - 1))
      Select Case PlotColor
         Case 0 'automatic
            PlotColor = clr_Plot(nNumTraces Mod 23)
         Case 1 'black
            PlotColor = QBColor(0)
         Case 2 'blue
            PlotColor = QBColor(1)
         Case 3 'green
            PlotColor = QBColor(2)
         Case 4 'cyan
            PlotColor = QBColor(3)
         Case 5 'red
            PlotColor = QBColor(4)
         Case 6
            PlotColor = QBColor(5)
         Case 7
            PlotColor = QBColor(6)
         Case 8
            PlotColor = QBColor(8)
         Case 9
            PlotColor = QBColor(9)
      End Select
      val_X = dArSpec(nTrace - 1, 0, nI)
      val_Y = dArSpec(nTrace - 1, 1, nI)
      
      'look for abrupt ends
      If nI = 0 Then 'If nI = udtLayoutSpec.lStart + 1 Then
         X0 = val_X
         Y0 = val_Y
      Else
         If (X0 > 0 And val_X < X0) Or _
            (X0 < 0 And val_X > X0) And _
            val_X = 0 And val_Y = 0 Then
            Exit For 'abrupt end reaached
         Else
            X0 = val_X
            Y0 = val_Y
            End If
         End If
         
      frmSpec.DrawWidth = Val(PlotInfo(9, nTrace - 1))
                  
      XRatio = (val_X - val_XMin) / val_XRange
      If XRatio >= 0 And XRatio <= 1 Then
        YRatio = (val_Y - val_YMin) / val_YRange
        If YRatio > 1 Then YRatio = 1
        If YRatio < 0 Then YRatio = 0
        If val_YMin >= 0 Then
          frmSpec.Line (twp_X0 + XRatio * twp_XRange, twp_Y0)- _
          (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), PlotColor
        Else
          frmSpec.Line (twp_X0 + XRatio * twp_XRange, twp_Y0Tr)- _
          (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), PlotColor
        End If
      End If
    Next nI
  
  Case AS_POINT
 
    frmSpec.DrawWidth = Val(PlotInfo(9, nTrace - 1))
    
'    For nI = udtLayoutSpec.lStart To udtLayoutSpec.lEnd Step 1
    For nI = 0 To RecordSize(nNumTraces - 1) - 1 Step 1
      PlotColor = Val(PlotInfo(2, nTrace - 1))
      Select Case PlotColor
         Case 0 'automatic
            PlotColor = clr_Plot(nNumTraces Mod 23)
         Case 1 'black
            PlotColor = QBColor(0)
         Case 2 'blue
            PlotColor = QBColor(1)
         Case 3 'green
            PlotColor = QBColor(2)
         Case 4 'cyan
            PlotColor = QBColor(3)
         Case 5 'red
            PlotColor = QBColor(4)
         Case 6
            PlotColor = QBColor(5)
         Case 7
            PlotColor = QBColor(6)
         Case 8
            PlotColor = QBColor(8)
         Case 9
            PlotColor = QBColor(9)
      End Select
      val_X = dArSpec(nTrace - 1, 0, nI)
      val_Y = dArSpec(nTrace - 1, 1, nI)
      XRatio = (val_X - val_XMin) / val_XRange
      YRatio = (val_Y - val_YMin) / val_YRange
      frmSpec.ForeColor = PlotColor
      If XRatio >= 0 And XRatio <= 1 And YRatio >= 0 And YRatio <= 1 Then
        frmSpec.PSet (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange)
      End If
      frmSpec.ForeColor = vbBlack
    Next nI
    frmSpec.DrawWidth = 1
  
  Case AS_DASH, AS_DOT, AS_DASHDOT, AS_DASHDOTDOT
    'find value starting point
    val_X = dArSpec(nTrace - 1, 0, udtLayoutSpec.lStart)
    val_Y = dArSpec(nTrace - 1, 1, udtLayoutSpec.lStart)
    XRatio = (val_X - val_XMin) / val_XRange
    YRatio = (val_Y - val_YMin) / val_YRange
    frmSpec.CurrentX = twp_X0 + XRatio * twp_XRange
    frmSpec.CurrentY = twp_Y0 - YRatio * twp_YRange
    'find rest
    Dim Xdash0, Ydash0, oldds%, olddw%
    olddw% = frmSpec.DrawWidth
    oldds% = frmSpec.DrawStyle
    frmSpec.DrawWidth = 1
    Select Case Val(PlotInfo(1, nTrace - 1))
       Case AS_DASH
         frmSpec.DrawStyle = vbDash
       Case AS_DOT
         frmSpec.DrawStyle = vbDot
       Case AS_DASHDOT
         frmSpec.DrawStyle = vbDashDot
       Case AS_DASHDOTDOT
         frmSpec.DrawStyle = vbDashDotDot
    End Select
'    For nI = udtLayoutSpec.lStart + 1 To udtLayoutSpec.lEnd Step 1
    For nI = 0 To RecordSize(nNumTraces - 1) - 1 Step 1
      PlotColor = Val(PlotInfo(2, nTrace - 1))
      Select Case PlotColor
         Case 0 'automatic
            PlotColor = clr_Plot(nNumTraces Mod 23)
         Case 1 'black
            PlotColor = QBColor(0)
         Case 2 'blue
            PlotColor = QBColor(1)
         Case 3 'green
            PlotColor = QBColor(2)
         Case 4 'cyan
            PlotColor = QBColor(3)
         Case 5 'red
            PlotColor = QBColor(4)
         Case 6
            PlotColor = QBColor(5)
         Case 7
            PlotColor = QBColor(6)
         Case 8
            PlotColor = QBColor(8)
         Case 9
            PlotColor = QBColor(9)
      End Select
      val_X = dArSpec(nTrace - 1, 0, nI)
      val_Y = dArSpec(nTrace - 1, 1, nI)
      
      frmSpec.DrawWidth = Val(PlotInfo(9, nTrace - 1))
      
      'look for abrupt ends
      If nI = 0 Then 'If nI = udtLayoutSpec.lStart + 1 Then
         X0 = val_X
         Y0 = val_Y
      Else
         If (X0 > 0 And val_X < X0) Or _
            (X0 < 0 And val_X > X0) And _
            val_X = 0 And val_Y = 0 Then
            Exit For 'abrupt end reaached
         Else
            X0 = val_X
            Y0 = val_Y
            End If
         End If
                  
      XRatio = (val_X - val_XMin) / val_XRange
      YRatio = (val_Y - val_YMin) / val_YRange
      frmSpec.Line -(twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), PlotColor
    Next nI
    'restore original DrawWidth, DrawStyle
    frmSpec.DrawWidth = olddw%
    frmSpec.DrawStyle = oldds%
    'clear lines outside drawing pane
    frmSpec.Line (twp_X0, twp_Y0 - twp_YRange - 10)-(twp_X0 + twp_XRange, 0), frmSpec.BackColor, BF 'above drawing pane
    frmSpec.Line (twp_X0, twp_Y0 + 20)-(frmSpec.Width, frmSpec.Height), frmSpec.BackColor, BF 'below drawing pane
    frmSpec.Line (twp_X0 + twp_XRange + 20, frmSpec.Height)-(frmSpec.Width, 0), frmSpec.BackColor, BF 'below drawing pane
    frmSpec.Line (0, 0)-(twp_XLeftMargin - 20, frmSpec.Height), frmSpec.BackColor, BF 'below drawing pane
    frmSpec.DrawWidth = 1
    
  Case AS_CIRCLE
  
    frmSpec.DrawWidth = Val(PlotInfo(9, nTrace - 1))
    store_style% = frmSpec.FillStyle
    frmSpec.FillStyle = 1 'transparent fill
    
'    For nI = udtLayoutSpec.lStart To udtLayoutSpec.lEnd Step 1
    For nI = 0 To RecordSize(nNumTraces - 1) - 1 Step 1
      PlotColor = Val(PlotInfo(2, nTrace - 1))
      Select Case PlotColor
         Case 0 'automatic
            PlotColor = clr_Plot(nNumTraces Mod 23)
         Case 1 'black
            PlotColor = QBColor(0)
         Case 2 'blue
            PlotColor = QBColor(1)
         Case 3 'green
            PlotColor = QBColor(2)
         Case 4 'cyan
            PlotColor = QBColor(3)
         Case 5 'red
            PlotColor = QBColor(4)
         Case 6
            PlotColor = QBColor(5)
         Case 7
            PlotColor = QBColor(6)
         Case 8
            PlotColor = QBColor(8)
         Case 9
            PlotColor = QBColor(9)
      End Select
      val_X = dArSpec(nTrace - 1, 0, nI)
      val_Y = dArSpec(nTrace - 1, 1, nI)
      XRatio = (val_X - val_XMin) / val_XRange
      YRatio = (val_Y - val_YMin) / val_YRange
      frmSpec.ForeColor = PlotColor
      If XRatio >= 0 And XRatio <= 1 And YRatio >= 0 And YRatio <= 1 Then
        frmSpec.Circle (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), frmSpec.ScaleHeight * 0.01
      End If
      frmSpec.ForeColor = vbBlack
    Next nI
    frmSpec.DrawWidth = 1
    frmSpec.FillStyle = store_style%
  
  Case AS_FILLEDCIRCLE
  
    frmSpec.DrawWidth = Val(PlotInfo(9, nTrace - 1))
    
    frmSpec.DrawWidth = Val(PlotInfo(9, nTrace - 1))
    store_style% = frmSpec.FillStyle
    frmSpec.FillStyle = 0 'solid fill
    
'    For nI = udtLayoutSpec.lStart To udtLayoutSpec.lEnd Step 1
    For nI = 0 To RecordSize(nNumTraces - 1) - 1 Step 1
      PlotColor = Val(PlotInfo(2, nTrace - 1))
      Select Case PlotColor
         Case 0 'automatic
            PlotColor = clr_Plot(nNumTraces Mod 23)
         Case 1 'black
            PlotColor = QBColor(0)
         Case 2 'blue
            PlotColor = QBColor(1)
         Case 3 'green
            PlotColor = QBColor(2)
         Case 4 'cyan
            PlotColor = QBColor(3)
         Case 5 'red
            PlotColor = QBColor(4)
         Case 6
            PlotColor = QBColor(5)
         Case 7
            PlotColor = QBColor(6)
         Case 8
            PlotColor = QBColor(8)
         Case 9
            PlotColor = QBColor(9)
      End Select
      val_X = dArSpec(nTrace - 1, 0, nI)
      val_Y = dArSpec(nTrace - 1, 1, nI)
      XRatio = (val_X - val_XMin) / val_XRange
      YRatio = (val_Y - val_YMin) / val_YRange
      frmSpec.ForeColor = PlotColor
      frmSpec.FillColor = QBColor(0)
      If XRatio >= 0 And XRatio <= 1 And YRatio >= 0 And YRatio <= 1 Then
        frmSpec.Circle (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), frmSpec.ScaleHeight * 0.01
      End If
      frmSpec.ForeColor = vbBlack
    Next nI
    frmSpec.DrawWidth = 1
    frmSpec.FillStyle = store_style%
  
  End Select
  
Next nTrace

  
'*************  plot ticks Y-axis  ************
If val_XMin < 0 Then
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + tick
    frmSpec.Line (twp_X0Tr - twp_XTick, twp_Y0)-(twp_X0Tr, twp_Y0), vbBlack
  End If
  For nI = 0 To NumYTicks
    frmSpec.Line (twp_X0Tr - twp_XTick, twp_Y0 - twp_OffSetY - nI * twp_YTickRange)- _
    (twp_X0Tr, twp_Y0 - twp_OffSetY - nI * twp_YTickRange), vbBlack
  Next nI
Else
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + tick
    frmSpec.Line (twp_X0 - twp_XTick, twp_Y0)-(twp_X0, twp_Y0), vbBlack
  End If
  For nI = 0 To NumYTicks
    frmSpec.Line (twp_X0 - twp_XTick, twp_Y0 - twp_OffSetY - nI * twp_YTickRange)- _
    (twp_X0, twp_Y0 - twp_OffSetY - nI * twp_YTickRange), vbBlack
  Next nI
End If


'************  plot labels to ticks from Y-axis  *************
Dim nLenYLable As Integer 'length of lable Y-axis ticks

frmSpec.ForeColor = clr_Plot(0)
If val_XMin < 0 Then
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + lable
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0Tr - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick - nLenYLable * frmSpec.FontSize * 3#
      frmSpec.CurrentY = twp_Y0 - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print Format(val_YMin, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin))))
      frmSpec.CurrentX = twp_X0Tr - twp_XLeftMargin + (5 - nLenYLable) * twp_XTick - nLenYLable * frmSpec.FontSize * 10#
      frmSpec.CurrentY = twp_Y0 - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print val_YMin
    End If
  End If 'plot lable for val_Ymin = gridline
  For nI = 0 To NumYTicks
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0Tr - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick - nLenYLable * frmSpec.FontSize * 2.5
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print Format(val_YMin + OffSetY + Pl_SpanY * nI, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin + OffSetY + Pl_SpanY * nI))))
      frmSpec.CurrentX = twp_X0Tr - twp_XLeftMargin + (5 - nLenYLable) * twp_XTick - nLenYLable * frmSpec.FontSize * 10#
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print val_YMin + OffSetY + Pl_SpanY * nI
    End If
  Next nI
Else
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + lable
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0 - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick - nLenYLable * frmSpec.FontSize * 2.5
      frmSpec.CurrentY = twp_Y0 - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print Format(val_YMin, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin))))
      frmSpec.CurrentX = twp_X0 - twp_XLeftMargin + (7 - nLenYLable) * twp_XTick - nLenYLable * frmSpec.FontSize * 10#
      frmSpec.CurrentY = twp_Y0 - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print val_YMin
    End If
  End If 'plot lable for val_Ymin = gridline
  For nI = 0 To NumYTicks
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0 - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick - nLenYLable * frmSpec.FontSize * 2.5
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print Format(val_YMin + OffSetY + Pl_SpanY * nI, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin + OffSetY + Pl_SpanY * nI))))
      frmSpec.CurrentX = twp_X0 - twp_XLeftMargin + (7 - nLenYLable) * twp_XTick - nLenYLable * frmSpec.FontSize * 10#
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick - frmSpec.Font.Size * 2.7
      frmSpec.Print val_YMin + OffSetY + Pl_SpanY * nI
    End If
  Next nI
End If
frmSpec.ForeColor = vbBlack

'**********  plot Y-axis and title  ***********

'/////////////////changes 11/23/19 -- to print vertically and rotrated//////////////////////
    Dim myFont As LOGFONT
    Dim hFont As Long
    Dim hOldFont As Long
    Dim OldGraphicsMode As Long

    ' font information
    myFont.lfFaceName = "Arial"
    myFont.lfHeight = 14
    If YAxis_font_size <> 0 Then myFont.lfHeight = YAxis_font_size
    myFont.lfWeight = FW_DEMIBOLD
    myFont.lfEscapement = 8100
    
    hFont = CreateFontIndirect(myFont)
    hOldFont = SelectObject(frmSpec.hDC, hFont)
    ' position of the text
'    frmSpec.CurrentY = frmSpec.Height / 2
'    frmSpec.CurrentX = frmSpec.Width / 2
'    frmSpec.Print "Hello"
'
'    SelectObject frmSpec.hdc, hOldFont
'    DeleteObject hFont
'//////////////////////////////////////////////////////////////
    
If val_XMin < 0 And val_XMax > 0 Then 'all four quadrants are shown
  frmSpec.Line (twp_X0Tr, twp_Y0 + twp_YTick)-(twp_X0Tr, twp_Y0 - twp_YRange), vbBlack 'Y-axis
  'prepare position title Y-axis and plot
  frmSpec.CurrentX = 10 + twp_X0Tr - (Len(udtLayoutSpec.YTitle) + 0.5) * twp_XTick / 2
'  frmSpec.CurrentY = 10
'  frmSpec.Print udtLayoutSpec.Ytitle
ElseIf val_XMin < 0 And val_XMax <= 0 Then '3rd quadrant is shown
  frmSpec.Line (twp_X0Tr, twp_Y0 + twp_YTick)-(twp_X0Tr, twp_Y0 - twp_YRange), vbBlack 'Y-axis
  'prepare position title Y-axis and plot
  frmSpec.CurrentX = 10 + twp_X0Tr - (Len(udtLayoutSpec.YTitle) + 0.5) * twp_XTick
'  frmSpec.CurrentY = 10
'  frmSpec.Print udtLayoutSpec.Ytitle
Else '1st quadrant is shown
  frmSpec.Line (twp_X0, twp_Y0 + twp_YTick)-(twp_X0, twp_Y0 - twp_YRange), vbBlack 'Y-axis
  'prepare position title Y-axis and plot
  frmSpec.CurrentX = twp_XLeftMargin
'  frmSpec.CurrentY = 10
'  frmSpec.Print udtLayoutSpec.Ytitle
End If

'//////////////changes to Ytitle printing (restore from printing from rotated by -90 degrees)//////////////////////
    frmSpec.CurrentY = frmSpec.ScaleHeight / 2 + Len(udtLayoutSpec.YTitle) * myFont.lfHeight * 2.5 ^ (10 / myFont.lfHeight) ^ 1.8 '.Font.Size * 2.5 'changes to center in Y based on title size
    frmSpec.CurrentX = 160 'frmSpec.Font.Size * 20
    frmSpec.Print udtLayoutSpec.YTitle

    SelectObject frmSpec.hDC, hOldFont
    DeleteObject hFont
'/////////////////////////////////////////////////////////////
'**********  plot X-axis and title and chart title (chart title added on 041620)**********

'//////////////changes to Xtitle printing//////////////////////

    ' font information (plot title horizontally)
    myFont.lfFaceName = "Arial"
    myFont.lfWeight = FW_DEMIBOLD
    myFont.lfHeight = 17
    If XAxis_font_size <> 0 Then myFont.lfHeight = XAxis_font_size
    myFont.lfEscapement = 0
    
    hFont = CreateFontIndirect(myFont)
    hOldFont = SelectObject(frmSpec.hDC, hFont)
    
If val_YMin < 0 And val_YMax > 0 Then 'all four quadrants are shown
  frmSpec.Line (twp_X0 - twp_XTick, twp_Y0Tr)-(twp_X0 + twp_XRange, twp_Y0Tr), vbBlack 'X-axis
'  frmSpec.CurrentX = 10
  frmSpec.CurrentY = twp_Y0Tr - 3 * twp_YTick
'  frmSpec.Print udtLayoutSpec.XTitle

    frmSpec.CurrentX = frmSpec.ScaleWidth / 2 - Len(udtLayoutSpec.XTitle) * myFont.lfHeight * 2.5
    'print Xtitle above the axis
'    frmSpec.CurrentY = frmSpec.CurrentY + frmSpec.Font.Size * 100 'under center line
    frmSpec.CurrentY = frmSpec.ScaleHeight - 800 'frmSpec.Font.Size * 100 'on bottom of page
    frmSpec.Print udtLayoutSpec.XTitle
    
ElseIf val_YMin < 0 And val_YMax <= 0 Then '3rd quadrant is shown
  frmSpec.Line (twp_X0 - twp_XTick, twp_Y0Tr)-(twp_X0 + twp_XRange, twp_Y0Tr), vbBlack 'X-axis
'  frmSpec.CurrentX = 10
  frmSpec.CurrentY = twp_Y0Tr - 3 * twp_YTick
'  frmSpec.Print udtLayoutSpec.XTitle

    frmSpec.CurrentX = frmSpec.ScaleWidth / 2 - Len(udtLayoutSpec.XTitle) * myFont.lfHeight * 2.5
    'print Xtitle above the axis
    frmSpec.CurrentY = frmSpec.CurrentY - myFont.lfHeight * 2.6 - 200 'frmSpec.Font.Size
    frmSpec.Print udtLayoutSpec.XTitle
    
Else '1st quadrant is shown
  frmSpec.Line (twp_X0 - twp_XTick, twp_Y0)-(twp_X0 + twp_XRange, twp_Y0), vbBlack 'X-axis
'  frmSpec.CurrentX = 10
  frmSpec.CurrentY = frmSpec.ScaleHeight - twp_YBottomMargin
'  frmSpec.Print udtLayoutSpec.XTitle

    frmSpec.CurrentX = frmSpec.ScaleWidth / 2 - Len(udtLayoutSpec.XTitle) * myFont.lfHeight * 2.5 '.Font.Size * 2.5 'changes to center in X based on title size
    'print Xtitle below the axis
    frmSpec.CurrentY = frmSpec.ScaleHeight - myFont.lfHeight * 2.6 - 600 ' frmspec.CurrentY + 400 'frmSpec.Font.Size * 50
    frmSpec.Print udtLayoutSpec.XTitle
    
End If

'///////////add chart Title (added on 041620)///////////////////////////
If Len(udtLayoutSpec.Title) > 0 Then
    frmSpec.Font.Size = 14
    If Title_font_size <> 0 Then frmSpec.Font.Size = Title_font_size
    frmSpec.Font.Bold = True
    frmSpec.Font.Italic = True
    frmSpec.CurrentX = frmSpec.ScaleWidth / 2 - Len(udtLayoutSpec.Title) * frmSpec.Font.Size * 5
    frmSpec.CurrentY = 0 'frmSpec.CurrentY + frmSpec.Font.Size * 20
    frmSpec.Print udtLayoutSpec.Title
'    frmSpec.Font.Size = 8
    If Val(frmSetCond.txtAxisLabelSize.Text) <> 0 Then
       frmSpec.Font.Size = Val(frmSetCond.txtAxisLabelSize.Text)
       End If
    frmSpec.Font.Bold = False
    frmSpec.Font.Italic = False
    End If
'////////////////////////////////
    
    SelectObject frmSpec.hDC, hOldFont
    DeleteObject hFont
'///////////////////////////end changes//////////////////////////

'**********  plot ticks X-axis  **********
If val_YMin < 0 Then
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xtick
    frmSpec.Line (twp_X0, twp_Y0Tr + twp_XTick)-(twp_X0, twp_Y0Tr), vbBlack
  End If
  For nI = 1 To NumXTicks
    frmSpec.Line (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0Tr + twp_XTick)- _
    (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0Tr), vbBlack
  Next nI
Else
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xtick
    frmSpec.Line (twp_X0, twp_Y0 + twp_XTick)-(twp_X0, twp_Y0), vbBlack
  End If
  For nI = 1 To NumXTicks
    frmSpec.Line (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0 + twp_XTick)- _
    (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0), vbBlack
  Next nI
End If


'**********  plot labels to ticks from X-axis  **********
Dim nLenXLable As Integer 'length of lable X-axis ticks
Dim sLenXLable As Integer 'length of x label

frmSpec.ForeColor = clr_Plot(0)
If val_YMin < 0 Then 'val_Ymin < 0 implies that twp_Y0Tr has to used
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xlable
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hDC, lHandleFont) 'save old font
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hDC, lOldFont) 'reload old font
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 0 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX)))) 'to avoid to get long lables (this will misplace the label under the tick)
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2 - nLenXLable * frmSpec.FontSize * 2.5
        frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
'        frmSpec.Print val_XMin - OffSetX + Pl_SpanX * nI
        If Abs(val_XMin - OffSetX + Pl_SpanX * nI) < 0.0000001 Then
           frmSpec.Print "0"
        Else
            sLenXLable = Len(Trim$(Str$(val_XMin - OffSetX + Pl_SpanX * nI)))
            frmSpec.CurrentX = frmSpec.CurrentX - sLenXLable * frmSpec.FontSize * 2.6
            frmSpec.Print Str(val_XMin - OffSetX + Pl_SpanX * nI)
            End If
      Next nI
    End If
  Else 'val_Xmin doesn't get a lable
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hDC, lHandleFont)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hDC, lOldFont)
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 1 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX))))
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2 - nLenXLable * frmSpec.FontSize * 2.5
        frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
'        frmSpec.Print val_XMin - OffSetX + Pl_SpanX * nI
        If Abs(val_XMin - OffSetX + Pl_SpanX * nI) < 0.0000001 Then
           frmSpec.Print "0"
        Else
            sLenXLable = Len(Trim$(Str$(val_XMin - OffSetX + Pl_SpanX * nI)))
            frmSpec.CurrentX = frmSpec.CurrentX - sLenXLable * frmSpec.FontSize * 2.6
            frmSpec.Print Str(val_XMin - OffSetX + Pl_SpanX * nI)
            End If
      Next nI
    End If
  End If
Else 'val_Ymin >= 0 ; this implies that twp_Y0 is used
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xlable
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hDC, lHandleFont)
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hDC, lOldFont)
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 0 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX))))
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2
        frmSpec.CurrentY = twp_Y0 + twp_XTick * 2 - nLenXLable * frmSpec.FontSize * 2.6
'        frmSpec.Print Str(val_XMin - OffSetX + Pl_SpanX * nI)
        If Abs(val_XMin - OffSetX + Pl_SpanX * nI) < 0.0000001 Then
           frmSpec.Print "0"
        Else
            sLenXLable = Len(Trim$(Str(val_XMin - OffSetX + Pl_SpanX * nI)))
            frmSpec.CurrentX = frmSpec.CurrentX - sLenXLable * frmSpec.FontSize * 2.5
            frmSpec.Print Str(val_XMin - OffSetX + Pl_SpanX * nI)
            End If
      Next nI
    End If
  Else
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hDC, lHandleFont)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hDC, lOldFont)
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 1 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX))))
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2 - nLenXLable * frmSpec.FontSize * 2.6
        frmSpec.CurrentY = twp_Y0 + twp_XTick * 2
        If Abs(val_XMin - OffSetX + Pl_SpanX * nI) < 0.0000001 Then
           frmSpec.Print "0"
        Else
            sLenXLable = Len(Trim$(Str(val_XMin - OffSetX + Pl_SpanX * nI)))
            frmSpec.CurrentX = frmSpec.CurrentX - sLenXLable * frmSpec.FontSize * 2.6
            frmSpec.Print Str(val_XMin - OffSetX + Pl_SpanX * nI)
            End If
      Next nI
    End If
  End If
End If
frmSpec.ForeColor = vbBlack

EndPlot = True 'flag the end of the plot

Exit Sub

errhand:
MsgBox "Error: " & Err.Number & vbLf & _
       Err.Description, vbExclamation + vbOKOnly, "Plot"



End Sub

Public Sub SetZoomValues(twp_StartXMD As Single, twp_StartYMD As Single, _
 twp_WidthMD As Single, twp_HeightMD As Single)
 
Dim val_ZoomXStart As Double
Dim val_ZoomXEnd As Double
Dim val_ZoomYStart As Double
Dim val_ZoomYEnd As Double
 
'check if boundaries of drawing pane are crossed
If twp_StartXMD < twp_XLeftMargin Then twp_StartXMD = twp_XLeftMargin
If twp_StartXMD > (twp_XLeftMargin + twp_XRange) Then twp_StartXMD = (twp_XLeftMargin + twp_XRange)
If twp_StartYMD < twp_YTopMargin Then twp_StartYMD = twp_YTopMargin
If twp_StartYMD > (twp_YTopMargin + twp_YRange) Then twp_StartYMD = (twp_YTopMargin + twp_YRange)
If twp_WidthMD > twp_XRange Then twp_WidthMD = twp_XRange
If twp_HeightMD > twp_YRange Then twp_HeightMD = twp_YRange

'calculate X-value
val_ZoomXStart = val_XMin + val_XRange * (twp_StartXMD - twp_XLeftMargin) / twp_XRange
val_ZoomXEnd = val_XMin + val_XRange * (twp_StartXMD + twp_WidthMD - twp_XLeftMargin) / twp_XRange

'calculate Y-value
val_ZoomYEnd = val_YMin + val_YRange * (twp_YRange - (twp_StartYMD - twp_YTopMargin)) / twp_YRange
val_ZoomYStart = val_YMin + val_YRange * (twp_YRange - (twp_StartYMD + twp_HeightMD - twp_YTopMargin)) / twp_YRange

With udtMyGraphLayout
  .X0 = val_ZoomXStart
  .X1 = val_ZoomXEnd
  .Y0 = val_ZoomYStart
  .Y1 = val_ZoomYEnd
End With
  

End Sub

Public Function GetValues(twp_XPosition As Single, twp_Yposition As Single) As COORDINATE

Dim val_CoordX As Double 'value X-coordinate
Dim val_CoordY As Double 'value Y-coordinate
Dim flg_OutsidePane As Boolean 'true when clicked outside drawing pane

'check if boundaries of drawing pane are crossed
If twp_XPosition < twp_XLeftMargin Then flg_OutsidePane = True
If twp_XPosition > (twp_XLeftMargin + twp_XRange) Then flg_OutsidePane = True
If twp_Yposition < twp_YTopMargin Then flg_OutsidePane = True
If twp_Yposition > (twp_YTopMargin + twp_YRange) Then flg_OutsidePane = True

'get X-coordinate
'If flg_OutsidePane = False Then
  val_CoordX = val_XMin + val_XRange * (twp_XPosition - twp_XLeftMargin) / twp_XRange
'Else
'  val_CoordX = 0
'End If

'get Y-coordinate
'If flg_OutsidePane = False Then
  val_CoordY = val_YMin + val_YRange * (twp_YRange - (twp_Yposition - twp_YTopMargin)) / twp_YRange
'Else
'  val_CoordY = 0
'End If

With GetValues
  .X = val_CoordX
  .Y = val_CoordY
End With


End Function

'**************************************
' Name: Auto resize flexgrid column widths
' Description:Automatically resize the columns in any flex grid to give a nice, professional appearance.
'Public sub automatically resizes MS Flex Grid columns to match the width of the text, no matter the size of the grid or the number of columns.
'Reads first n number of rows of data, and adjusts column size to match the widest cell of text. Will even expand columns proportionately if they aren't wide enough to fill out the entire width of the grid. Configurable constraints allow you to designate
'1) Any flex grid to resize
'2) Maximum column width
'3) the maximum number of rows in depth to look for the widest cell of text.
' By: Jonathan W. Lartigue (from psc cd)
'
' Inputs:msFG (MSFlexGrid) = The name of the flex grid to resize .... MaxRowsToParse (integer) = The maximum number of rows (depth) of the table to scan for cell width (e.g. 50) .... MaxColWidth (Integer) = The maximum width of any given cell in twips (e.g. 5000)
'
' Assumes:Simply drop this public sub into your form or module and access it from anywhere in your program to automatically resize any flex grid.
'**************************************

Public Sub AutosizeGridColumns(ByRef msFG As MSFlexGrid, ByVal MaxRowsToParse As Integer, ByVal MaxColWidth As Integer)
Dim I, J As Integer
Dim txtString As String
Dim intTempWidth, BiggestWidth As Integer
Dim intRows As Integer, intBiggestWidth As Integer
Const intPadding = 150
With msFG
 For I = 0 To .Cols - 1
' Loops through every column
.Col = I
' Set the active colunm
intRows = .Rows
' Set the number of rows
If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
' If there are more rows of data, reset
' intRows to the MaxRowsToParse constant
 
intBiggestWidth = 0
' Reset some values to 0
For J = 0 To intRows - 1
 ' check up to MaxRowsToParse # of rows and obtain
 ' the greatest width of the cell contents
 
 .Row = J
 
 txtString = .Text
 intTempWidth = Len(txtString) + intPadding
 ' The intPadding constant compensates for text insets
 ' You can adjust this value above as desired.
 
 If intTempWidth > intBiggestWidth Then intBiggestWidth = intTempWidth
 ' Reset intBiggestWidth to the intMaxColWidth value if necessary
Next J
.ColWidth(I) = intBiggestWidth
 Next I
 ' Now check to see if the columns aren't as wide as the grid itself.
 ' If not, determine the difference and expand each column proportionately
 ' to fill the grid
 intTempWidth = 0
 
 For I = 0 To .Cols - 1
intTempWidth = intTempWidth + .ColWidth(I)
' Add up the width of all the columns
 Next I
 
 If intTempWidth < msFG.Width Then
' Compate the width of the columns to the width of the grid control
' and if necessary expand the columns.
intTempWidth = Fix((msFG.Width - intTempWidth) / .Cols)
' Determine the amount od width expansion needed by each column
For I = 0 To .Cols - 1
 .ColWidth(I) = .ColWidth(I) + intTempWidth
 ' add the necessary width to each column
 
Next I
 End If
End With
End Sub
' min and max are the minimum and maximum indexes
' of the items that might still be out of order.
'copied from http://www.vb-helper.com/tut1.htm and modified to handle two dimensional double value array
Public Sub BubbleSort(list() As Double, ByVal Min As Integer, _
    ByVal max As Integer)
Dim last_swap As Integer
Dim I As Integer
Dim J As Integer
Dim tmp1 As Double, tmp2 As Double

    ' Repeat until we are done.
    Do While Min < max
        ' Bubble up.
        last_swap = Min - 1
        ' For i = min + 1 To max
        I = Min + 1
        Do While I <= max
            ' Find a bubble.
            If list(0, I - 1) > list(0, I) Then
                ' See where to drop the bubble.
                tmp1 = list(0, I - 1)
                tmp2 = list(1, I - 1)
                J = I
                Do
                    list(0, J - 1) = list(0, J)
                    list(1, J - 1) = list(1, J)
                    J = J + 1
                    If J > max Then Exit Do
                Loop While list(0, J) < tmp1
                list(0, J - 1) = tmp1
                list(1, J - 1) = tmp2
                last_swap = J - 1
                I = J + 1
            Else
                I = I + 1
            End If
        Loop
        ' Update max.
        max = last_swap - 1

        ' Bubble down.
        last_swap = max + 1
        ' For i = max - 1 To min Step -1
        I = max - 1
        Do While I >= Min
            ' Find a bubble.
            If list(0, I + 1) < list(0, I) Then
                ' See where to drop the bubble.
                tmp1 = list(0, I + 1)
                tmp2 = list(1, I + 1)
                J = I
                Do
                    list(0, J + 1) = list(0, J)
                    list(1, J + 1) = list(1, J)
                    J = J - 1
                    If J < Min Then Exit Do
                Loop While list(0, J) > tmp1
                list(0, J + 1) = tmp1
                list(1, J + 1) = tmp2
                last_swap = J + 1
                I = J - 1
            Else
                I = I - 1
            End If
        Loop
        ' Update min.
        Min = last_swap + 1
    Loop
End Sub
