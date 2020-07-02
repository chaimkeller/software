Attribute VB_Name = "modPrintPreview"
Option Explicit

Public iermjwPDF As Integer
Public ierPDF As Integer

   ' The following Types, Declares, and Constants are only necessary
   ' for the PrintPicture routine
   '=======================================================================
   Type BITMAPINFOHEADER_TYPE
      biSize As Long
      biWidth As Long
      biHeight As Long
      biPlanes As Long
      biBitCount As Long
      biCompression As Long
      biSizeImage As Long
      biXPelsPerMeter As Long
      biYPelsPerMeter As Long
      biClrUsed As Long
      biClrImportant As Long
      bmiColors As String * 1024
   End Type

   Type BITMAPINFO_TYPE
      BITMAPINFOHEADER As BITMAPINFOHEADER_TYPE
      bmiColors As String * 1024
   End Type

   ' Enter each of the following Declare statements as one, single line:
   Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, _
      ByVal hBitmap As Long, ByVal nStartScan As Long, _
      ByVal nNumScans As Long, ByVal lpBits As Long, _
      BITMAPINFO As BITMAPINFO_TYPE, ByVal wUsage As Long) As Long
   Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
      ByVal DestX As Long, ByVal DestY As Long, _
      ByVal wDestWidth As Long, ByVal wDestHeight As Long, _
      ByVal SrcX As Long, ByVal SrcY As Long, _
      ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
      ByVal lpBits As Long, BitsInfo As BITMAPINFO_TYPE, _
      ByVal wUsage As Long, ByVal dwRop As Long) As Long
   Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
      ByVal lMem As Long) As Long
   Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) _
      As Long
      
      

   Global Const SRCCOPY = &HCC0020
   Global Const BI_RGB = 0
   Global Const DIB_RGB_COLORS = 0
   Global Const GMEM_MOVEABLE = 2
   Global Const gsEmpty = ""

   ' Module level variables set in PrintStartDoc flag indicating Printing
   ' or Previewing:
   Public PrinterFlag As Boolean
   ' Object used for Print Preview:
   Dim ObjPrint As Control
   ' Storage for output objects original scale mode:
   Dim sm
   ' The size ratio between the actual page and the print preview object:
   Dim Ratio
   ' Size of the non-printable area on printer:
   Dim LRGap
   Dim TBGap
   ' The actual paper size (8.5 x 11 normally):
   Public PgWidth
   Public PgHeight
   
   Public PaprSize 'Paper Size set for your printer
   Public PaprType As String 'Name of paper
   Public PaperOrientation 'Paper Orientation set for your printer

   
   Public magPrint& 'zoom percentage
   Public finishedloading As Boolean 'flag to tell program when finished loading zoomcombo
   Public PicWidth, PicHeight 'stored initial values of picture1.width/height
   Public PicLeft, PicTop 'stored initial values of picture1.left/top
   Public zoomfactor& 'zoom is defined by (1 + zoomfactor&*0.1)
                      'not used in current project
   Public LoadInit As Boolean
   Public numPrinter& 'number of installed printers
'----------------screen resoultion variables---------------
   Type RECT
       X1 As Long
       Y1 As Long
       X2 As Long
       Y2 As Long
   End Type

   ' NOTE: The following declare statements are case sensitive.

   Declare Function GetDesktopWindow Lib "user32" () As Long
   Declare Function GetWindowRect Lib "user32" _
      (ByVal hwnd As Long, Rectangle As RECT) As Long

      '*****************************************************************
      ' FUNCTION: GetScreenResolution()
      '
      ' PURPOSE:
      '   To determine the current screen size or resolution.
      '
      ' RETURN:
      '   The current screen resolution. Typically one of the following:
      '      640 x 480
      '      800 x 600
      '      1024 x 768
      '
      '*****************************************************************
      Function GetScreenResolution() As String
          Dim R As RECT
          Dim hwnd As Long
          Dim retval As Long
          hwnd = GetDesktopWindow()
          retval = GetWindowRect(hwnd, R)
          GetScreenResolution = (R.X2 - R.X1) & "x" & (R.Y2 - R.Y1)
      End Function
   

   Sub PrintStartDoc(objToPrintOn As Control, PF) ', PaperWidth, PaperHeight)
               
      'Most of this Sub is taken from MSDN's Visual Studio 6.0's example
      'of how to make a Print Preview using Visual Basic.  If you need
      'more information about how it works, ask them!
      
      Dim psm
      Dim fsm
      Dim HeightRatio
      Dim WidthRatio
      
      On Error GoTo errhand

      ' Set the flag that determines whether printing or previewing:
      PrinterFlag = PF

      'Set the physical page size: 'already set by FindPaperSize
      'PgWidth = PaperWidth '<--original code
      'PgHeight = PaperHeight

      ' Find the size of the non-printable area on the printer to
      ' use to offset coordinates. These formulas assume the
      ' non-printable area is centered on the page:
      
      If numPrinter& > 0 Then
         psm = Printer.ScaleMode
         Printer.ScaleMode = 5 'Inches
         LRGap = (PgWidth - Printer.ScaleWidth) / 2
         TBGap = (PgHeight - Printer.ScaleHeight) / 2
         Printer.ScaleMode = psm
      Else 'no printer installed, so set some defaults
         psm = 1
         LRGap = 0.85
         TBGap = 0.56
         End If

      ' Initialize printer or preview object:
      If PrinterFlag Then 'initialize printer
         sm = Printer.ScaleMode
         Printer.ScaleMode = 5 'Inches
         Printer.Print sEmpty;
      Else 'preview object
         ' Set the object used for preview:
         Set ObjPrint = objToPrintOn
         ' Scale Object to Printer's printable area in Inches:
         sm = ObjPrint.ScaleMode
         ObjPrint.ScaleMode = 5 'Inches
         ' Compare the height and width ratios to determine the
         ' ratio to use and how to size the picture box:
         HeightRatio = ObjPrint.ScaleHeight / PgHeight
         WidthRatio = ObjPrint.ScaleWidth / PgWidth
         If HeightRatio < WidthRatio Then
            Ratio = HeightRatio
            ' Re-size picture box - this does not work on a form:
            fsm = ObjPrint.Parent.ScaleMode
            ObjPrint.Parent.ScaleMode = 5 'Inches
            ObjPrint.Width = PgWidth * Ratio
            ObjPrint.Parent.ScaleMode = fsm
         Else
            Ratio = WidthRatio
            ' Re-size picture box - this does not work on a form:
            fsm = ObjPrint.Parent.ScaleMode
            ObjPrint.Parent.ScaleMode = 5 'Inches
            ObjPrint.Height = PgHeight * Ratio
            ObjPrint.Parent.ScaleMode = fsm
         End If
         ObjPrint.Scale (0, 0)-(PgWidth, PgHeight)
         ObjPrint.BackColor = QBColor(15) 'CHANGES!!!!!!
         ObjPrint.Cls
         
         ' Set default properties of picture box to match printer
         ' There are many that you could add here:
         If numPrinter& > 0 Then
            ObjPrint.FontName = Printer.FontName
            ObjPrint.FontSize = Printer.FontSize * Ratio
            ObjPrint.ForeColor = Printer.ForeColor
         Else 'no installed printer, so use defaults
            ObjPrint.FontName = "Arial"
            ObjPrint.FontSize = 8.28 * Ratio
            ObjPrint.ForeColor = 0
            End If
      
      End If
      Exit Sub
      
errhand:
   ShowPreviewError
   
   End Sub

   Sub PrintCurrentX(XVal)
      If PrinterFlag Then
         Printer.CurrentX = XVal - LRGap
      Else
         ObjPrint.CurrentX = XVal
      End If
   End Sub

   Sub PrintCurrentY(YVal)
      If PrinterFlag Then
         Printer.CurrentY = YVal - TBGap
      Else
         ObjPrint.CurrentY = YVal
      End If
   End Sub

   Sub PrintFontName(pFontName)
      If PrinterFlag Then
         Printer.FontName = pFontName
      Else
         ObjPrint.FontName = pFontName
      End If
   End Sub

   Sub PrintFontSize(pSize)
      If PrinterFlag Then
         Printer.FontSize = pSize
      Else
         ' Sized by ratio since Scale method does not effect FontSize:
         ObjPrint.FontSize = pSize * Ratio
      End If
   End Sub
   Sub PrintFontBold()
      If PrinterFlag Then
         Printer.FontBold = True
      Else
         ' Sized by ratio since Scale method does not effect FontSize:
         ObjPrint.FontBold = True
      End If
   End Sub
   Sub PrintFontRegular()
      If PrinterFlag Then
         Printer.FontBold = False
      Else
         ' Sized by ratio since Scale method does not effect FontSize:
         ObjPrint.FontBold = False
      End If
   End Sub
   

   Sub PrintPrint(PrintVar)
      If PrinterFlag Then
         Printer.Print PrintVar
      Else
         ObjPrint.Print PrintVar
      End If
   End Sub

   Sub PrintLine(bLeft0, bTop0, bLeft1, bTop1, color)
      If PrinterFlag Then
         Printer.Line (bLeft0 - LRGap, bTop0 - TBGap)- _
            (bLeft1 - LRGap, bTop1 - TBGap), color
      Else
         ObjPrint.Line (bLeft0, bTop0)-(bLeft1, bTop1), color
      End If
   End Sub

   Sub PrintBox(bLeft, bTop, bWidth, bHeight, color)
      If PrinterFlag Then
         ' Enter the following two lines as one, single line:
         Printer.Line (bLeft - LRGap, bTop - TBGap)- _
            (bLeft + bWidth - LRGap, bTop + bHeight - TBGap), color, B
      Else
         ObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), color, B
      End If
   End Sub

   Sub PrintFilledBox(bLeft, bTop, bWidth, bHeight, color)
      If PrinterFlag Then
         ' Enter the following two lines as one, single line:
         Printer.Line (bLeft - LRGap, bTop - TBGap)- _
            (bLeft + bWidth - LRGap, bTop + bHeight - TBGap), color, BF
      Else
         ' Enter the following two lines as one, single line:
         ObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), _
            color, BF
      End If
   End Sub

   Sub PrintCircle(bLeft, bTop, bRadius, color)
      If PrinterFlag Then
         Printer.Circle (bLeft - LRGap, bTop - TBGap), bRadius, color
      Else
         ObjPrint.Circle (bLeft, bTop), bRadius, color
      End If
   End Sub

   Sub PrintNewPage()
      If PrinterFlag Then
         Printer.NewPage
      Else
         ObjPrint.Cls
      End If
   End Sub
   Sub PrintPicture(picSource As Control, ByVal pLeft, ByVal pTop, _
      ByVal pWidth, ByVal pHeight)

      'This sub is taken from MSDN's Visual Studio 6.0's example
      'of how to make a Print Preview using Visual Basic.  If you need
      'more information about how it works, ask them!

      ' Picture Box should have autoredraw = False, ScaleMode = Pixel
      ' Also can have visible=false, Autosize = true

      Dim BITMAPINFO As BITMAPINFO_TYPE
      Dim DesthDC As Long
      Dim hMem As Long
      Dim lpBits As Long
      Dim R As Long
      
      On Error GoTo errhand

      ' Precaution:
      If pLeft < LRGap Or pTop < TBGap Then Exit Sub
      If pWidth < 0 Or pHeight < 0 Then Exit Sub
      If pWidth + pLeft > PgWidth - LRGap Then Exit Sub
      If pHeight + pTop > PgHeight - TBGap Then Exit Sub
      picSource.ScaleMode = 3 'Pixels
      picSource.AutoRedraw = False
      picSource.Visible = False
      picSource.AutoSize = True

      If PrinterFlag Then
         Printer.ScaleMode = 3 'Pixels
         ' Calculate size in pixels:
         pLeft = ((pLeft - LRGap) * 1440) / Printer.TwipsPerPixelX
         pTop = ((pTop - TBGap) * 1440) / Printer.TwipsPerPixelY
         pWidth = (pWidth * 1440) / Printer.TwipsPerPixelX
         pHeight = (pHeight * 1440) / Printer.TwipsPerPixelY
         Printer.Print sEmpty;
         DesthDC = Printer.hdc
      Else
         ObjPrint.Scale
         ObjPrint.ScaleMode = 3 'Pixels
         ' Calculate size in pixels:
         pLeft = ((pLeft * 1440) / Screen.TwipsPerPixelX) * Ratio
         pTop = ((pTop * 1440) / Screen.TwipsPerPixelY) * Ratio
         pWidth = ((pWidth * 1440) / Screen.TwipsPerPixelX) * Ratio
         pHeight = ((pHeight * 1440) / Screen.TwipsPerPixelY) * Ratio
         DesthDC = ObjPrint.hdc
      End If

      BITMAPINFO.BITMAPINFOHEADER.biSize = 40
      BITMAPINFO.BITMAPINFOHEADER.biWidth = picSource.ScaleWidth
      BITMAPINFO.BITMAPINFOHEADER.biHeight = picSource.ScaleHeight
      BITMAPINFO.BITMAPINFOHEADER.biPlanes = 1
      BITMAPINFO.BITMAPINFOHEADER.biBitCount = 8
      BITMAPINFO.BITMAPINFOHEADER.biCompression = BI_RGB

     
      hMem = GlobalAlloc(GMEM_MOVEABLE, (CLng(picSource.ScaleWidth + 3) _
         \ 4) * 4 * picSource.ScaleHeight) 'DWORD ALIGNED
      lpBits = GlobalLock(hMem)

      
      R = GetDIBits(picSource.hdc, picSource.Image, 0, _
         picSource.ScaleHeight, lpBits, BITMAPINFO, DIB_RGB_COLORS)
      If R <> 0 Then
         
         R = StretchDIBits(DesthDC, pLeft, pTop, pWidth, pHeight, 0, 0, _
            picSource.ScaleWidth, picSource.ScaleHeight, lpBits, _
            BITMAPINFO, DIB_RGB_COLORS, SRCCOPY)
      End If

      R = GlobalUnlock(hMem)
      R = GlobalFree(hMem)

      If PrinterFlag Then
         Printer.ScaleMode = 5 'Inches
      Else
         ObjPrint.ScaleMode = 5 'Inches
         ObjPrint.Scale (0, 0)-(PgWidth, PgHeight)
      End If
      
   Exit Sub
errhand:
   ShowPreviewError
   
   End Sub

   Sub PrintEndDoc()
      If PrinterFlag Then
         If Not ScreenDump And Not PrintMag Then 'send search results to printer
            Printer.EndDoc
            Printer.ScaleMode = sm
         ElseIf ScreenDump Or PrintMag Then 'send screen dump of map (portion) to printer
            PrintPictureToFitPage Printer, PrintPreview.Picture1
            Printer.EndDoc
            End If
      Else
         ObjPrint.ScaleMode = sm
      End If
   End Sub

Sub FindPaperSize()

On Error GoTo errhand

   'Find paper size in inches.
   If numPrinter& > 0 Then
      PaprSize = Printer.PaperSize
   Else 'no printer installed, so set Letter as default
      If PrintPreview.cmbtxtPaper.ListIndex > -1 Then
         PaprSize = PrintPreview.cmbtxtPaper.ListIndex + 1
      Else 'use default of A4
         PaprSize = 9
         End If
      End If
   
50 Select Case PaprSize
      Case 1 'Letter, 8.5 x 11 in.
         PaprType = "Letter"
         PgWidth = 8.5
         PgHeight = 11
      Case 2 'Letter, Small 8.5 x 11 in.
         PaprType = "Letter, Small"
         PgWidth = 8.5
         PgHeight = 11
      Case 3 'Tabloid, 11 x 17 in.
         PaprType = "Tabloid"
         PgWidth = 11
         PgHeight = 17
      Case 4 'Ledger, 17 x 11 in.
         PaprType = "Ledger"
         PgWidth = 17
         PgHeight = 11
      Case 5 'Legal, 8.5 x 14 in.
         PaprType = "Legal"
         PgWidth = 8.5
         PgHeight = 14
      Case 6 'Statement, 5.5 x 8.5 in.
         PaprType = "Statement"
         PgWidth = 5.5
         PgHeight = 8.5
      Case 7 'Executive, 7.5 x 10.5 in.
         PaprType = "Executive"
         PgWidth = 7.5
         PgHeight = 10.5
      Case 8 'A3, 297 x 420 mm
         PaprType = "A3"
         PgWidth = 11.69
         PgHeight = 16.54
      Case 9 'A4, 210 x 297 mm
         PaprType = "A4"
         PgWidth = 8.27
         PgHeight = 11.93
      Case 10 'A4 Small, 210 x 297 mm
         PaprType = "A4 Small"
         PgWidth = 8.25
         PgHeight = 11.93
      Case 11 'A5, 148 x 210 mm
         PaprType = "A5"
         PgWidth = 5.83
         PgHeight = 8.27
      Case 12 'B4, 250 x 354 mm
         PaprType = "B4"
         PgWidth = 9.84
         PgHeight = 13.94
      Case 13 'B5, 182 x 257 mm
         PaprType = "B5"
         PgWidth = 7.17
         PgHeight = 10.12
      Case 14 'Folio, 8.5 x 13 in.
         PaprType = "Folio"
         PgWidth = 8.5
         PgHeight = 13
      Case 15 'Quarto, 215 x 275 mm
         PaprType = "Quarto"
         PgWidth = 8.47
         PgHeight = 10.83
      Case 16 ' 10 x 14 in.
         PaprType = "10x14 in."
         PgWidth = 10
         PgHeight = 14
      Case 17 ' 11 x 17 in.
         PaprType = "11x17 in."
         PgWidth = 11
         PgHeight = 17
      Case 18 ' Note, 8.5 x 11 in
         PaprType = "Note"
         PgWidth = 8.5
         PgHeight = 11
      Case Else
         MsgBox "Warning, unrecognized paper size detected. A4 set as default. Check your printer's settings!", vbExclamation + vbOKOnly, "Print Preview"
         'set A4 as European default
         PaprSize = 9
         GoTo 50
   End Select
   
   PrintPreview.cmbtxtPaper.ListIndex = PaprSize - 1
   
   If ScreenDump Or PrintMag Then
      
      PrintPreview.cmbtxtPaper.Text = "Special"
      
      'make special page size to fit entire screen dump
      'this will depend on screen resolution
      'For 800 x 600 then 8.9" x 13" is enough
      
      'Xresol is the X resolution of the screen
      
      PgWidth = 8.9 * XResol / 800
      PgHeight = 13 * XResol / 800
      End If
   
   Exit Sub
   
errhand:
    ShowPreviewError
    'set A4 as default
    PaprSize = 9
    GoTo 50

End Sub
Sub FindPaperOrientation()
   On Error GoTo errhand
   
   Dim change As Boolean
   
   If numPrinter& = 0 Then 'no printer installed, use defaults
      If PaperOrientation = 0 Then
         PaperOrientation = 1 'set portrait as default
      ElseIf PaperOrientation = 1 And PgWidth > PgHeight Then
         change = True
      ElseIf PaperOrientation = 2 And PgWidth < PgHeight Then
         change = True
         End If
   Else
      change = False
      If Printer.Orientation <> PaperOrientation Then
         change = True 'change paper orientation
         End If
      End If

   'PaperOrientation = Printer.Orientation
50 Select Case PaperOrientation
      Case 1 'portrait
         If numPrinter& > 0 Then Printer.Orientation = vbPRORPortrait
         PrintPreview.ImgLandscape.Visible = False
         PrintPreview.imgPortrait.Visible = True
         PrintPreview.optPortrait.value = True
         PrintPreview.ImgLandscape.ToolTipText = gsEmpty
         PrintPreview.imgPortrait.ToolTipText = "Portrait orientation"
      Case 2 'landscape
         If numPrinter& > 0 Then Printer.Orientation = vbPRORLandscape
         PrintPreview.imgPortrait.Visible = False
         PrintPreview.ImgLandscape.Visible = True
         PrintPreview.optLandscape.value = True
         PrintPreview.imgPortrait.ToolTipText = gsEmpty
         PrintPreview.ImgLandscape.ToolTipText = "Landscape orientation"
   End Select
   If PaperOrientation = 1 And LoadInit Then
      'portrait orientation
   ElseIf PaperOrientation = 2 And LoadInit Then
      'landscape orientation
       Dim PgWidthTmp
       PgWidthTmp = PgWidth
       PgWidth = PgHeight
       PgHeight = PgWidthTmp
       End If
       
   If change And Not LoadInit Then 'switch between Portrait and Landscape
      PgWidthTmp = PgWidth
      PgWidth = PgHeight
      PgHeight = PgWidthTmp
      PreviewSetup 'initialize picture boxes
      PreviewPrint 'Execute Printing/Previewing
      'redisplay preview at current zoom setting
      PrintPreview.ZoomCombo_click
      End If
            
   Exit Sub
   
errhand:
    ShowPreviewError
    'set portrait as default
    PaperOrientation = 1
    GoTo 50
    
End Sub
'Sub FindPrinterName()
'   On Error GoTo errhand
'
'   PrintPreview.lbltxtPrinter = Printer.DeviceName
'   Exit Sub
'
'errhand:
'   ShowPreviewError
'   PrintPreview.lbltxtPrinter = gsEmpty
'End Sub

'------------------------------------------------------------
'this sub displays the error message with it's Err code
'------------------------------------------------------------
Sub ShowPreviewError()
  Dim sTmp As String
  
  Screen.MousePointer = vbDefault

  sTmp = "The following Error occurred:" & vbCrLf & vbCrLf
  'add the error string
  sTmp = sTmp & Err.Description & vbCrLf
  'add the error number
  sTmp = sTmp & "VB Error Number: " & Err & vbCrLf & vbCrLf
  'add a suggestion
  sTmp = sTmp & "Check your printer's settings!"
  
  beep

  MsgBox sTmp, vbCritical + vbOKOnly, "Print Preview"
  Err.Clear

End Sub

Sub ScrollBars()
   'settings for scroll bars
   
   Screen.MousePointer = vbHourglass
   
   'horizontal scroll bar
   With PrintPreview.HScroll1
      If PrintPreview.Picture1.Width + PrintPreview.Picture1.left > PrintPreview.RightBorderPictureBox.left Then
        .Visible = True
        .left = 0
        .Width = PrintPreview.BottomBorderPictureBox.Width
        .top = 0
        'settings for it
        .Max = PrintPreview.Picture2.Width + PrintPreview.Picture1.Width
        .LargeChange = .Max / 30
        .SmallChange = .Max / 60
      Else
         .Visible = False
      End If
   End With
      
   
   'now vertical scroll bar
   With PrintPreview.VScroll1
      If PrintPreview.Picture1.top + PrintPreview.Picture1.Height > PrintPreview.BottomBorderPictureBox.top Then
        .Visible = True
        .left = 0
        .top = 0
        .Height = PrintPreview.RightBorderPictureBox.Height - PrintPreview.BottomBorderPictureBox.Height - PrintPreview.TopBorderPictureBox.top
        'settings for it
'        Dim cc
        .Max = PrintPreview.Picture2.Height + PrintPreview.Picture1.Height
        .LargeChange = .Max / 30
        .SmallChange = .Max / 60
      Else
        .Visible = False
      End If
   End With
   
   Screen.MousePointer = vbDefault
   
End Sub

Sub PositionBorders()

      On Error Resume Next
      
'     positions of picture boxes that act as the form's borders

      With PrintPreview.TopBorderPictureBox
           .left = -10
           .Width = PrintPreview.Width '9700
           .Height = 735
           .top = -60
      End With
      
      With PrintPreview.LeftBorderPictureBox
        .left = 0
        .Width = 495
        .top = PrintPreview.TopBorderPictureBox.top + PrintPreview.TopBorderPictureBox.Height - 10 '667
        .Height = PrintPreview.Height - .top - 400 '7523
      End With
      
      With PrintPreview.RightBorderPictureBox
         .Width = 495 '675
         .left = PrintPreview.Width - .Width - 120 '9120
         .top = PrintPreview.TopBorderPictureBox.top + PrintPreview.TopBorderPictureBox.Height - 10 '667
         .Height = PrintPreview.LeftBorderPictureBox.Height
      End With
      
      With PrintPreview.BottomBorderPictureBox
         .left = PrintPreview.LeftBorderPictureBox.Width  'assumes that leftpicturebox is at .left=0 '480
         .Width = PrintPreview.RightBorderPictureBox.left - PrintPreview.LeftBorderPictureBox.Width '8685
         .top = PrintPreview.LeftBorderPictureBox.top + PrintPreview.LeftBorderPictureBox.Height - 450 '7740
         .Height = 495
      End With
End Sub

Sub PreviewSetup()
      'setup picture boxes that display preview
      PrintPreview.Picture1.AutoRedraw = True
      PrintPreview.Picture1.Picture = LoadPicture(sEmpty)
      PrintPreview.Picture2.AutoRedraw = False
      PrintPreview.Picture2.ScaleMode = 3 'Pixels
      PrintPreview.Picture2.Visible = False
      PrintPreview.Picture2.AutoSize = True
      PrintPreview.Picture2.Picture = LoadPicture(sEmpty)
End Sub

Sub LoadPaperSize()
'Add paper sizes in inches to combo box PrintPreview.cmbtxtPaper.

On Error GoTo errhand

      'Letter, 8.5 x 11 in.
         PaprType = "Letter"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Letter, Small 8.5 x 11 in.
         PaprType = "Letter, Small"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Tabloid, 11 x 17 in.
         PaprType = "Tabloid"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Ledger, 17 x 11 in.
         PaprType = "Ledger"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Legal, 8.5 x 14 in.
         PaprType = "Legal"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Statement, 5.5 x 8.5 in.
         PaprType = "Statement"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Executive, 7.5 x 10.5 in.
         PaprType = "Executive"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'A3, 297 x 420 mm
         PaprType = "A3"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'A4, 210 x 297 mm
         PaprType = "A4"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'A4 Small, 210 x 297 mm
         PaprType = "A4 Small"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'A5, 148 x 210 mm
         PaprType = "A5"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'B4, 250 x 354 mm
         PaprType = "B4"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'B5, 182 x 257 mm
         PaprType = "B5"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Folio, 8.5 x 13 in.
         PaprType = "Folio"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      'Quarto, 215 x 275 mm
         PaprType = "Quarto"
         PrintPreview.cmbtxtPaper.AddItem PaprType
      ' 10 x 14 in.
         PaprType = "10x14 in."
         PrintPreview.cmbtxtPaper.AddItem PaprType
      ' 11 x 17 in.
         PaprType = "11x17 in."
         PrintPreview.cmbtxtPaper.AddItem PaprType
      ' Note, 8.5 x 11 in
         PaprType = "Note"
         PrintPreview.cmbtxtPaper.AddItem PaprType
   
   Exit Sub
   
errhand:
    ShowPreviewError

End Sub

Sub LoadPrinterName()
    'load names of available printers into PrintPreview.cmbtxtPrinter combo box

    On Error GoTo errhand
       
    Dim x As Printer, foundPrinter&
    numPrinter& = 0
    For Each x In Printers
        PrintPreview.cmbtxtPrinter.AddItem x.DeviceName
        numPrinter& = numPrinter& + 1
        If x.DeviceName = Printer.DeviceName Then
           'this is default printer
           Set Printer = x
           foundPrinter& = numPrinter&
           End If
    Next
    
    'show the default printer
    If numPrinter& > 0 Then
       PrintPreview.cmbtxtPrinter.ListIndex = foundPrinter& - 1
    Else
       PrintPreview.cmbtxtPrinter.Text = "None Installed"
       End If
    Exit Sub
   
errhand:
   ShowPreviewError

End Sub

Sub LoadPaperOrientation()
   'set default paper orientation
   
   On Error GoTo errhand
      
   If LoadInit And (ScreenDump Or PrintMag) Then
      'set default as landscape
      PrintPreview.optLandscape.value = True
      Exit Sub
      End If
   
   If numPrinter& > 0 Then
     If Printer.Orientation = vbPRORPortrait Then
        PrintPreview.optPortrait.value = True
     ElseIf Printer.Orientation = vbPRORLandscape Then
        PrintPreview.optLandscape.value = True
        End If
   Else 'no printer installed, so set portrait as default
      PrintPreview.optPortrait.value = True
      End If
   
   Exit Sub
   
errhand:
   ShowPreviewError

End Sub


Sub PreviewPrint()

On Error GoTo errhand
      
      If PrinterFlag = True Then
        'Sending document to printer so
        'don't refresh preview screen.
      Else
        'Repaint the preview screen.
        PreviewSetup
        End If
        
      'This print job can go to the printer or the picture box as a preview.
      'This is determined by the PrinterFlag (= True sends to printer)
      PrintStartDoc PrintPreview.Picture1, PrinterFlag

      'Paint empty paper sheet on preview screen
      'Note that all the subs use inches
      PrintPicture PrintPreview.Picture2, 1.1, 1.1, 0.8, 0.8
      
'-----------------------------------------------------------------------------
      'Document details--This is either search results or Screen Dump of Map
      If Not ScreenDump And Not PrintMag Then
         PreviewPrintDetails 'preview detailed report on specific record
      ElseIf ScreenDump And Not PrintMag Then
         PreviewPrintScreen 'preview current map tile
      ElseIf PrintMag And Not ScreenDump Then
         PreviewPrintMag 'preview magnified portion of map
         End If
'------------------------------------------------------------------------------
      
      PrintEndDoc
      If PrinterFlag = True Then
         PrinterFlag = False
         End If
         
      Exit Sub
      
errhand:
   ShowPreviewError
      
End Sub
Sub PreviewPrintScreen() 'preview current map tile
     'freeze the blinkers at the on state until end of screen dump
    If ce& = 1 Then 'freeze center blinker
       Do Until CenterBlinkState
          DoEvents
          If CenterBlinkState Then Exit Do
       Loop
       GDMDIform.CenterPointTimer.Enabled = False
       End If
       
    'clear out the printpreview picture box
    Set PrintPreview.Picture1.Picture = Nothing
    Set PrintPreview.Picture2.Picture = Nothing
    
    'now capture the map
    Set PrintPreview.Picture1.Picture = CaptureClient(GDform1)
    
    'print center coordinates to blank bottom left portion of picture1
    PrintCurrentX 1
    PrintCurrentY 8 * XResol / 800
    PrintFontName "Arial"
    PrintFontBold
    PrintFontSize 20
    PrintPrint "Marker at " & lblX & " = " & GDMDIform.Text5 & "; " & LblY & " = " & GDMDIform.Text6 & "; hgt = " & GDMDIform.Text7 & " meters"
    PrintCurrentX 1
    PrintPrint "MapDigitizer Map, Date/Time: " & Now()
    
    'reenable blinkers
    If ce& = 1 Then
       GDMDIform.CenterPointTimer.Enabled = True
       End If

End Sub
Sub PreviewPrintMag() 'preview magnified portion of map tile
    
    'clear out the printpreview picture box
    Set PrintPreview.Picture1.Picture = Nothing
    Set PrintPreview.Picture2.Picture = Nothing
    
    'now capture the map
    
    Dim winx As Long, winy As Long
    If GDMagform.Image1.Width >= GDMagform.Picture1.Width Then
       winx = GDMagform.Picture1.Width / twipsx
    Else
       winx = GDMagform.Image1.Width / twipsx
       End If
       
    If GDMagform.Image1.Height >= GDMagform.Picture1.Height Then
       winy = GDMagform.Picture1.Height / twipsy
    Else
       winy = GDMagform.Image1.Height / twipsy
       End If
    
    'capture the map portion for previewing
    Set PrintPreview.Picture1.Picture = CaptureWindow(GDMagform.Picture1.hwnd, 0, 0, 0, winx, winy)
    
    'print center coordinates to blank bottom left portion of picture1
    PrintCurrentX 1
    PrintCurrentY 8 * XResol / 800
    PrintFontName "Arial"
    PrintFontBold
    PrintFontSize 20
    'print out magnifcation and corner coordinates
    If mag >= 1 Then
       PrintPrint "Magnification: " & str$(Int(mag * 100)) & "%"
    Else
       PrintPrint "Demagnification: " & str$(Int(mag * 100)) & "%"
       End If
    
    Dim Xout As Single, Yout As Single
    Dim x1mag As String, y1mag As String
    Dim x2mag As String, y2mag As String
    Call ConvertPixToCoord(drag1x, drag1y, Xout, Yout)
    x1mag = Fix(Xout)
    y1mag = Fix(Yout)
    Call ConvertPixToCoord(drag2x, drag2y, Xout, Yout)
    x2mag = Fix(Xout)
    y2mag = Fix(Yout)
    
    PrintCurrentX 1
    PrintPrint "Map Boundaries: (" & x1mag & "," & y1mag & ")-(" & x2mag & "," & y2mag & ")"
    
End Sub

Sub CheckPrinter(ier%)
    'check if printer is installed
    'if it is installed, then ier%=0
    'else ier%=-1
    
    Dim x As Printer, found%
    
    found% = 0
    For Each x In Printers
        found% = 1
        ier% = 0
        Exit For
    Next
    
    If found% = 0 Then 'no printer has yet been installed
       ier% = -1
       End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : SavetoFile
' DateTime  : 9/26/2004 11:03
' Author    : Chaim Keller
' Purpose   : Save Print Preview to file
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : SavetoFile
' Author    : Chaim Keller
' Date      : 7/21/2011
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub SavetoFile()

   'This routine saves the search results to a file.
   'There are three possible output formats:
   
   '(1) txt file
   'this option also allows the user to load up a stored
   'search file into the report form in order to plot and
   'to obtain detailed reports that can be printed
   
   On Error GoTo SavetoFile_Error

   On Error GoTo errhand
   
   Dim FileOutName$, pos%, filtmp&, filtm1&, doclin$, ext$
   Dim myfile, response, Xcoord%, Ycoord%, ycoord0%, TmpStr$
   Dim i%, skipRow&, numRow&
   
   If Dir(direct$ & "\print_tmp.txt") = sEmpty Then
      Call MsgBox("Can't file temporary file ""print_tmp.txt""!", vbCritical, App.Title)
      Exit Sub
      End If
   
10 FileOutName$ = PrintPreview.cmbPages.List(PrintPreview.cmbPages.ListIndex)
   pos% = InStr(FileOutName$, ":")
   Mid$(FileOutName$, pos%, 2) = "__"
   FileOutName$ = direct$ + "\" & FileOutName$
     
   FileOutName$ = FileOutName$ & "_scannedDB_" & Trim$(str$(Abs(PreviewOrderNum&))) '& ".xls"
   
   PrintPreview.CommonDialog1.CancelError = True
   PrintPreview.CommonDialog1.FileName = FileOutName$
   PrintPreview.CommonDialog1.Filter = _
       "pdf file (*.pdf)|*.pdf|Unformated text file (*.txt)|*.txt"
   PrintPreview.CommonDialog1.FilterIndex = 1
   PrintPreview.CommonDialog1.ShowSave
   'check for existing files, and for wrong save directories
  
  On Error GoTo SavetoFile_Error

   If PrintPreview.CommonDialog1.FileName = sEmpty Then Exit Sub
       
   ext$ = RTrim$(Mid$(PrintPreview.CommonDialog1.FileName, InStr(1, _
       PrintPreview.CommonDialog1.FileName, ".") + 1, 3))
     
   myfile = Dir(PrintPreview.CommonDialog1.FileName)
   If myfile <> sEmpty And ext$ <> "xls" Then
      response = MsgBox("Write over existing file?", vbYesNoCancel + vbQuestion, _
          "Cal Program")
      If response = vbNo Then
         GoTo 10
      ElseIf response = vbCancel Then
         Exit Sub
         End If
      End If
      
      'wait a little bit of time for message box to close
      Dim waitime As Long
      waitime = Timer
      Do Until Timer > waitime + 0.5
         DoEvents
      Loop

'-----------------------save to txt file-----------------
25 If ext$ = "txt" Then

      Screen.MousePointer = vbHourglass
      
      Close
      
      'open temporary buffer file and print it's contents to the
      'output file.
      filtmp& = FreeFile 'the temporary buffer file
      Open direct$ & "\print_tmp.txt" For Input As #filtmp&
      filtm1& = FreeFile 'the output file
      Open PrintPreview.CommonDialog1.FileName For Output As #filtm1&
      
      'print source name
      Print #filtm1&, PrintPreview.cmbPages.List(PrintPreview.cmbPages.ListIndex)
      Print #filtm1&, sEmpty
      
      Line Input #filtmp&, doclin$
      
      ycoord0% = 1
      Do Until EOF(filtmp&)
         Input #filtmp&, Xcoord%, Ycoord%, TmpStr$
         TmpStr$ = Trim$(TmpStr$)
         If Ycoord% <> ycoord0% Then
            For i% = 1 To Ycoord% - ycoord0%
               Print #filtm1&, sEmpty
            Next i%
            Print #filtm1&, TmpStr$
         Else
            Print #filtm1&, TmpStr$
            If ycoord0% = 1 Then Print #filtm1&, sEmpty
            End If
         ycoord0% = Ycoord% + 1
      Loop
      
      Close #filtmp&
      Close #filtm1&
     
      Screen.MousePointer = vbDefault
      
     
'  ElseIf ext$ = "pdf" Then
'
'    Screen.MousePointer = vbHourglass
'    GDMDIform.StatusBar1.Panels(1).Text = "Please wait....."
'
'    'now create it using vbPDF (hopefully faster)
'    CreatePDF2
'
'    Screen.MousePointer = vbDefault
'    GDMDIform.StatusBar1.Panels(1).Text = sEmpty
    
    End If
       
   On Error GoTo 0
   Exit Sub
   
errhand:
   Exit Sub

SavetoFile_Error:
    
    GDMDIform.StatusBar1.Panels(1).Text = sEmpty
    Screen.MousePointer = vbDefault
    If filtmp& <> 0 Then Close #filtmp&
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SavetoFile of Module modPrintPreview", vbCritical


End Sub

'---------------------------------------------------------------------------------------
' Procedure : WordWrap01
' Author    : Chaim Keller
' Date      : 7/19/2011
' Purpose   : Wraps Text

'Source: http://www.xbeat.net/vbspeed/c_WordWrap.htm
'
'Wraps a string to a given number of characters using a line break character.
'Note, that our function is a wrapper based on character counting; we are not measuring the rendered text width within a given device context. By the way, "LineWrap" is another, slightly better, term for what the function does: wrapping lines at word boundaries; nevertheless, "WordWrap" is more common.

'Note the following points:
'Text can be wrapped at space (ASCII 32) and hyphen (ASCII 45) characters.
'Words that are longer than Width are cut.
'We wrap lossless: no string characters of the original string are replaced, so that the output is fully revertible to the input by simply removing the line break characters. To achieve this, lines wrapped at a space character will end in that space character (see examples below). In other words, the space behaves exactly as a wrapping hyphen.
'The line break character is fixed to 2-byte vbCrLf (ASCII 13 & ASCII 10).
'No line breaks are added to the very end of the output (see last example below).
'Any line break characters in the original string are ignored. It's the responsability of the caller to take care of removing them.
'If Width is 0 the function returns an empty string.
'CountLines is an optional return argument that's handy when you need to know the height of the wrapped string returned or whether anything has been wrapped (CountLines > 1).
'Examples (the comma stands for vbCrLf):
'
'  WordWrap("ab cdef ghi", 2)       --> "ab ,cd,ef ,gh,i"
'  WordWrap("ab cdef ghi", 3)       --> "ab ,cde,f ,ghi"
'  WordWrap("ab cdef ghi", 4)       --> "ab ,cdef ,ghi"
'  WordWrap("ab-cdef-ghi", 4)       --> "ab-,cdef-,ghi"
'  WordWrap("abcedef", 0)           --> ""
'  WordWrap("abcedef", 1)           --> "a,b,c,d,e,f,g"
'  WordWrap("abcedef", 2)           --> "ab,cd,ef,g"
'  WordWrap("abcedef", 3)           --> "abc,def,g"
'  WordWrap("a ", 1)                --> "a " (not "a ,")
'---------------------------------------------------------------------------------------
'
Public Function WordWrap01( _
    ByRef Text As String, _
    ByVal Width As Long, _
    Optional ByRef CountLines As Long) As String
' by Donald, donald@xbeat.net, 20040913
  Dim i As Long
  Dim LenLine As Long
  Dim posBreak As Long
  Dim cntBreakChars As Long
  Dim abText() As Byte
  Dim abTextOut() As Byte
  Dim ubText As Long

  ' no fooling around
   On Error GoTo WordWrap01_Error

  If Width <= 0 Then
    CountLines = 0
    Exit Function
  End If
  If Len(Text) <= Width Then  ' no need to wrap
    CountLines = 1
    WordWrap01 = Text
    Exit Function
  End If
  
  abText = StrConv(Text, vbFromUnicode)
  ubText = UBound(abText)
  ReDim abTextOut(ubText * 3) 'dim to potential max
  
  For i = 0 To ubText
    Select Case abText(i)
    Case 32, 45 'space, hyphen
      posBreak = i
    Case Else
    End Select
    
    abTextOut(i + cntBreakChars) = abText(i)
    LenLine = LenLine + 1
    
    If LenLine > Width Then
      If posBreak > 0 Then
        ' don't break at the very end
        If posBreak = ubText Then Exit For
        ' wrap after space, hyphen
        abTextOut(posBreak + cntBreakChars + 1) = 13  'CR
        abTextOut(posBreak + cntBreakChars + 2) = 10  'LF
        i = posBreak
        posBreak = 0
      Else
        ' cut word
        abTextOut(i + cntBreakChars) = 13     'CR
        abTextOut(i + cntBreakChars + 1) = 10 'LF
        i = i - 1
      End If
      cntBreakChars = cntBreakChars + 2
      LenLine = 0
    End If
  Next
  
  CountLines = cntBreakChars \ 2 + 1
  
  ReDim Preserve abTextOut(ubText + cntBreakChars)
  WordWrap01 = StrConv(abTextOut, vbUnicode)

   On Error GoTo 0
   Exit Function

WordWrap01_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WordWrap01 of Module modPrintPreview"
  
End Function

'---------------------------------------------------------------------------------------
' Procedure : WordWrap
' Author    : Chaim Keller
' Date      : 7/20/2011
' Purpose   : Fits long text into sentances of length <= SentanceLength and returns
'             CountLines sentances as elements of an array, TheTextArray
'           'TheTextArray must be defined in the calling routine containing the Maximum number
'            of lines to look for, Max_Num_Lines
'---------------------------------------------------------------------------------------
'
Public Sub WordWrap(ByRef TheText, ByRef TheTextArray(), ByRef Max_Num_Lines, _
                    ByRef SentanceLength, ByRef CountLines)

   On Error GoTo WordWrap_Error
   
   Dim i As Long, TheChar$, posBreak As Long
   Dim NewLine As String, LenLine As Long
   Dim Begin As Long, TotalLen As Long, found%
   Dim NumBreakCharacters As Long, j As Long
   Dim posBreak2 As Long
   
   If Len(TheText) <= SentanceLength Then 'nothing to do
      CountLines = 1
      TheTextArray(CountLines - 1) = TheText
      Exit Sub
      End If
      
   LenLine = 0
   CountLines = 0
   posBreak = 0
   Begin = 1
   
50:
   
   For i = Begin To Len(TheText) + NumBreakCharacters
   
      TheChar$ = Mid$(TheText, i, 1)
      
      Select Case TheChar$
      
         Case Chr$(13) 'vbcr
            posBreak = LenLine
            posBreak2 = i
            TheChar$ = Chr$(13)
            NumBreakCharacters = NumBreakCharacters + 1
            
         Case Chr$(10) 'vblf
            posBreak = LenLine
            posBreak2 = 1
            TheChar$ = Chr$(10)
            NumBreakCharacters = NumBreakCharacters + 1
            
         Case Chr$(32) 'space
            posBreak = LenLine
            posBreak2 = i
            TheChar$ = Chr$(32)
            
         Case Else
         
            'do nothing else
            
            
      End Select
      
      LenLine = LenLine + 1
      TotalLen = i + NumBreakCharacters
      
      NewLine = NewLine & TheChar$
      
      If LenLine > SentanceLength And TotalLen < Len(TheText) + NumBreakCharacters Then  'stuff the line into array
      
         If posBreak > 0 Then
         
            CountLines = CountLines + 1
            TheTextArray(CountLines - 1) = Mid$(NewLine, 1, posBreak)
            Begin = posBreak2 + 1
            posBreak = 0
            NewLine = sEmpty
            LenLine = 0
            
            If CountLines + 1 > Max_Num_Lines Then
               Exit Sub
            Else 'keep on parsing
               GoTo 50
               End If
            
         Else
         
            'add hyphen and break the line at this point
            CountLines = CountLines + 1
            TheTextArray(CountLines - 1) = NewLine & "-"
            Begin = i + 1
            NewLine = sEmpty
            LenLine = 0
           
            If CountLines + 1 > Max_Num_Lines Then
               found% = 1
               Exit For
            Else 'keep on parsing
               GoTo 50
               End If
            
            End If
            
      ElseIf TotalLen >= Len(TheText) + NumBreakCharacters Then   'next line is just what is left
      
         'check for a few more characters
         
          For j = i + 1 To Len(TheText)
             
            TheChar$ = Mid$(TheText, j, 1)
            NewLine = NewLine + TheChar$
            
         Next j
      
         CountLines = CountLines + 1
         TheTextArray(CountLines - 1) = NewLine
         found% = 1
         Exit For
            
         End If
            
   Next i
   
   If found% = 0 Then 'never got the end
      CountLines = CountLines + 1
      TheTextArray(CountLines - 1) = NewLine
      End If

   On Error GoTo 0
   Exit Sub

WordWrap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WordWrap of Module modPrintPreview"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CreatePDF2
' Author    : Chaim Keller
' Date      : 7/24/2011
' Purpose   : Handles saving print preview to Pdf file
'---------------------------------------------------------------------------------------
'
'Public Sub CreatePDF2()
'
'  Dim clPDF As New clsPDFCreator
'
'    Dim Xo As Single, Yo As Single, YoLastLeft As Single
'    Dim Xspacing As Single, Yspacing As Single
'
'    Dim numOrder(125) As Variant, waitime As Long
'
'    Dim AnalystName As String
'    Dim AnalysisDate As String
'    Dim AnalystNames(NUM_FOSSIL_TYPES - 1) As Variant
'    Dim AnalysisDates(NUM_FOSSIL_TYPES - 1) As Variant
'    Dim FosIDCono As Long
'    Dim FosIDDiatom As Long
'    Dim FosIDForam As Long
'    Dim FosIDMega As Long
'    Dim FosIDNano As Long
'    Dim FosIDOstra As Long
'    Dim FosIDPaly As Long
'    Dim DocTitle$, DocName$, DocLine$, sAnum&
'    Dim TheTextArray(20) As Variant, TheText As String
'    Dim Max_Num_Lines As Long
'    Dim CountLines As Long
'    Dim SentanceLength As Long
'    Dim FossilTbl$, strnum&, FosNum&, Ylast As Single
'    Dim numOFile$
'
'   'query the database
'
''   On Error GoTo CreatePDF1_Error
'
'    ierPDF = 0
'
'    Call QueryToPrintSave(numOrder(), DocName$, AnalystName, AnalysisDate, AnalystNames(), AnalysisDates(), sAnum&)
'
'    Screen.MousePointer = vbHourglass
'
'    If PicSum Then
'       DocTitle$ = "GSI Paleontology Dbase SEARCH RESULT, # " & Trim$(str$(NewHighlighted&))
'       If OrderNum& > 0 Then DocName$ = Trim$(str$(OrderNum&))
'    Else
'       If OrderNum& < 0 Then 'record from scanned database
'           DocTitle$ = "GSI Paleontology Database, Dbase Order No. " & DocName$ '& " ; DATE/TIME: " & Now
'       Else
'           DocName$ = Trim$(str$(OrderNum&))
'           DocTitle$ = "GSI Paleontology Database, Dbase Order No. " & DocName$ '& " ; DATE/TIME: " & Now
'           End If
'
'       End If
'
'
''    ' Set the PDF title and filename
''    objPDF.PDFTitle = DocTitle$
''    objPDF.PDFFileName = PrintPreview.CommonDialog1.FileName
'
'  Dim strFile As String
'  Dim i As Single
'
'  ' Imposta il file di output
'  strFile = PrintPreview.CommonDialog1.FileName
'
'  With clPDF
'    .Title = DocTitle$       ' Titolo
'    .ScaleMode = pdfMillimeter         ' Unità di misura
'    .PaperSize = pdfA4                  ' Formato pagina
'    .Margin = 0                         ' Margine
'    .Orientation = pdfPortrait          ' Orientamento
'
'    .EncodeASCII85 = True '(chkASCII85.Value = Checked)
'
'    .InitPDFFile strFile                ' inizializza il file
'
'    If ierPDF = -1 Then 'file already open
'       Screen.MousePointer = vbDefault
'       MsgBox "Can't save the pdf file while it is being used by another application." & _
'              vbCrLf & "Please close that application and try again.", _
'              vbExclamation + vbOKOnly, App.Title
'       Exit Sub
'       End If
'
'
'    ' Definisce le risorse relative ai font
'    .LoadFont "Fnt1", "Times New Roman"                       ' Tipo TrueType
'    .LoadFont "Fnt2", "Arial", pdfBold 'pdfItalic                      ' Tipo TrueType
'    .LoadFont "Fnt3", "Courier New"                           ' Tipo TrueType
'    .LoadFontStandard "Fnt4", "Courier New", pdfBoldItalic    ' Tipo Type1
'    .LoadFont "Fnt5", "Arial", pdfNormal
'
'    ' Definisce le risorse relative alle immagini
'     .LoadImgFromBMPFile "Logo", App.Path & "\Gsi_03.bmp", pdfRGB
''    .LoadImgFromBMPFile "Img1", App.Path & "\img\20x20x24.bmp" ', pdfGrayScale
''    .LoadImgFromBMPFile "Img2", App.Path & "\img\200x200x24.bmp" ', pdfGrayScale
'
'    ' Definisce una risorsa comune da stampare solo sulle pagine pari
'    .StartObject "Item1", pdfAllPages ' , pdfEvenPages
'      .SetColorFill -240
'      .SetTextHorizontalScaling 100
''      .DrawText 6, 4, "Bozza", "Fnt2", 200, , 60
'      .SetColorFill 0
'    .EndObject
'
''     Inizializza la prima pagina
'    .BeginPage
'
'    'header with logo
'    .DrawImg "Logo", 12, 288, 35, 35
'    If OrderNum& < 0 Then
'       .DrawText 52, 269, DocTitle$, "Fnt2", 14, pdfAlignLeft
'    Else
'       .DrawText 60, 269, DocTitle$, "Fnt2", 14, pdfAlignLeft
'       End If
'
'    'change text color to blue
'    .SetColorFill rgb(0, 0, 255)
'    .SetColorStroke rgb(0, 0, 255)
'    .DrawText 100, 252, "Summary", "Fnt2", 14, pdfAlignLeft
'
'    'restore text color to black
'    .SetColorFill rgb(0, 0, 0)
'    .SetColorStroke rgb(0, 0, 0)
'
'    'demarcation line
'    .SetLineCap 0
'    .SetLineWidth 0.4
'    .SetColorStroke rgb(0, 0, 0)
'    .MoveTo 15, 248
'    .LineTo 195, 248
'
'  '---------------------------------------------
'      'begin summary section (page 1) of report
'  '-----------------------------------------------------------------------------
'
'     Xo = 15: Yo = 240: Yspacing = 5
'     DocLine$ = "CLIENT: " & numOrder(23)
'     .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'     Yo = Yo - Yspacing
'     DocLine$ = "COMPANY/DIVISION: " & numOrder(25)
'     .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'     Yo = Yo - Yspacing
'     DocLine$ = "PROJECT: " & numOrder(24)
'     .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'     Yo = Yo - Yspacing
'     DocLine$ = "FORMATION: " & numOrder(3)
'     .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'     If sAnum& = 1 Then 'well
'        Yo = Yo - Yspacing
'        DocLine$ = "SAMPLE METHOD: WELL"
'        .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'        Yo = Yo - Yspacing
'        DocLine$ = "WELL NAME: " & numOrder(0)
'        .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'     ElseIf sAnum& = 0 Then 'surface
'        Yo = Yo - Yspacing
'        DocLine$ = "SAMPLE METHOD: SURFACE"
'        .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'        Yo = Yo - Yspacing
'        DocLine$ = "PLACE NAME: " & numOrder(0)
'        .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'     ElseIf sAnum& = -1 Then 'unknown type
'        Yo = Yo - Yspacing
'        DocLine$ = "SAMPLE METHOD: UNKNOWN"
'        .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'        Yo = Yo - Yspacing
'        DocLine$ = "PLACE NAME: " & numOrder(0)
'        .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'        End If
'
'    Yo = Yo - Yspacing
'    DocLine$ = "ITMx: " & numOrder(1)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    Yo = Yo - Yspacing
'    DocLine$ = "ITMy: " & numOrder(2)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    TheText = "REMARKS: " & numOrder(11)
'
'    If Len(numOrder(11)) > 0 Then
'        'check for Hebrew
'        If HebrewCheck(numOrder(11)) Then
'
'           Screen.MousePointer = vbDefault
'
'           MsgBox "Hebrew characters where detected in the ''Remark''." _
'                  & vbCrLf & vbCrLf & "However, Hebrew fonts are not supported for pdf files." _
'                  & vbCrLf & vbCrLf & "(Consider reentering the remark to the database in English.)", _
'                  vbInformation + vbOKOnly, App.Title
'
'           'give a bit of time to close the message box and repaint
'           waitime = Timer
'           Do Until Timer > waitime + 0.5
'              DoEvents
'           Loop
'
'           Screen.MousePointer = vbHourglass
'
'           End If
'        End If
'
'    Max_Num_Lines = 5
'    SentanceLength = 100
'    Call WordWrap(TheText, TheTextArray(), Max_Num_Lines, SentanceLength, CountLines)
'    For i = 1 To CountLines
'       Yo = Yo - Yspacing
'       YoLastLeft = Yo
'       DocLine$ = TheTextArray(i - 1)
'      .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'    Next i
'
'     '________GENERAL SAMPLE INFO-FIRST RIGHT COLUMN___________
'
'    Xo = 135: Yo = 240: Yspacing = 5
'    DocLine$ = "LIM UP:         " & numOrder(5)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    Yo = Yo - Yspacing
'    DocLine$ = "LIM DOWN:    " & numOrder(4)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    If val(numOrder(7)) = 1 And sAnum& = 1 Then
'       Yo = Yo - Yspacing
'       DocLine$ = "WELL SAMPLE TYPE: CUTTING"
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'    ElseIf val(numOrder(7)) = 2 And sAnum& = 1 Then
'       Yo = Yo - Yspacing
'       DocLine$ = "WELL SAMPLE TYPE: CORE"
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'    Else
'       Yo = Yo - Yspacing
'       DocLine$ = "WELL SAMPLE TYPE: "
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'       End If
'
'ppd800:
'    Yo = Yo - Yspacing
'    DocLine$ = "CORE NUMBER: " & numOrder(26)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    Yo = Yo - Yspacing
'    DocLine$ = "BOX NUMBER: " & numOrder(27)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    Yo = Yo - Yspacing
'    DocLine$ = "FIELD NUMBER: " & numOrder(9)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    Yo = Yo - Yspacing
'    DocLine$ = "ORDER NUMBER: " & DocName$
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'    Yo = Yo - Yspacing
'    DocLine$ = "DATE: " & numOrder(21)
'   .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
''
''  '-----------print fossil and age information for scanned database
'    If OrderNum& < 0 Then
'
'       'draw line to delineate the fossil info from the summary
'       .SetLineWidth 0.4
'       .SetColorStroke rgb(0, 0, 0)
'       Yo = YoLastLeft - Yspacing
'       .MoveTo 15, Yo
'       .LineTo 195, Yo
'       Yo = Yo - Yspacing
'
'       Xo = 15: Yo = Yo - Yspacing
'       DocLine$ = "FOSSIL CATEGORY: " & FossilTbl$
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'       Yo = Yo - Yspacing
'       DocLine$ = "DATE: " & numOrder(21)
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'       Yo = Yo - Yspacing
'       DocLine$ = "DATE: " & numOrder(21)
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'       Yo = Yo - Yspacing
'       DocLine$ = "Earlier AGE: " & numOrder(28) & " " & numOrder(29)
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'       Yo = Yo - Yspacing
'       DocLine$ = "LATER AGE: " & numOrder(30) & " " & numOrder(31)
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'       Yo = Yo - Yspacing
'       DocLine$ = "Earlier ZONE: "
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'       Yo = Yo - Yspacing
'       DocLine$ = "LATER ZONE: "
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'       Yo = Yo - Yspacing
'       DocLine$ = "REMARK: "
'       .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
''        'now add tiff form image, if it exists, as a second page
''
''         If Dir(tifViewerDir$) <> sEmpty Then
''            Call FindTifPath(Abs(OrderNum&), numOFile$)
''
''            Select Case numOFile$
''              Case "-1" 'error flag
''
''                  'just close the document
''
''              Case Else 'add to second page of pdf
''
''                If Dir(tifDir$ & "\" & UCase$(numOFile$)) = sEmpty Then
''
''                   'just close the document
''
''                Else 'add to tiff file to second page
''
''                    objPDF.PDFNewPage
''
''                    objPDF.PDFImage tifDir$ & "\" & UCase$(numOFile$), _
''                          0, 0, 500, 800, "http://www.gsi.gov.il"
''
''                    objPDF.PDFEndPage
''
''                   End If
''
''            End Select
''            End If
'
'       'add page number
'       .DrawText 200, 10, "Page 1", "Fnt5", 12, pdfAlignRight
'
'       .EndObject
'
'       .ClosePDFFile
'
'       GDMDIform.StatusBar1.Panels(1).Text = sEmpty
'       Screen.MousePointer = vbDefault
'
'       GoTo cp29998
'
'       Exit Sub
'
'       End If
'
'    '--------summary of fossil results-----------------------------------------
'    'note: no checking is undertaken to see that the text doesn't go off the page
'    'this is highly unlikely, so hasn't been implemented
'    '--------------------------------------------------------------------------
'    'draw line to delineate the fossil info from the summary
'    .SetLineWidth 0.4
'    .SetColorStroke rgb(0, 0, 0)
'    Yo = YoLastLeft - Yspacing
'    .MoveTo 15, Yo
'    .LineTo 195, Yo
'    Yo = Yo
'
'    Xo = 15
'    Ylast = Yo
'
'    'first conod
'    strnum& = 29
'    If val(numOrder(18)) <> 0 Then
'       FossilTbl$ = "CONODONTA"
'       FosNum& = 18
'       Call PhpFossilInfo2(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, clPDF)
'       End If
'
'
'    'then diato
'    strnum& = strnum& + 12
'    If val(numOrder(17)) <> 0 Then
'       FossilTbl$ = "DIATOM"
'       FosNum& = 17
'       Call PhpFossilInfo2(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, clPDF)
'       End If
'
'    'then foram
'    strnum& = strnum& + 12
'    If val(numOrder(12)) <> 0 Then
'       FossilTbl$ = "FORAMINIFERA"
'       FosNum& = 12
'       Call PhpFossilInfo2(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, clPDF)
'       End If
'
'    'then megaf
'    strnum& = strnum& + 12
'    If val(numOrder(15)) <> 0 Then
'       FossilTbl$ = "MEGAFAUNA"
'       FosNum& = 15
'       Call PhpFossilInfo2(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, clPDF)
'       End If
'
'    'then nanno
'    strnum& = strnum& + 12
'    If val(numOrder(16)) <> 0 Then
'       FossilTbl$ = "NANNOPLANKTON"
'       FosNum& = 16
'       Call PhpFossilInfo2(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, clPDF)
'       End If
'
'    'then ostra
'    strnum& = strnum& + 12
'    If val(numOrder(13)) <> 0 Then
'       FossilTbl$ = "OSTRACODA"
'       FosNum& = 13
'       Call PhpFossilInfo2(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, clPDF)
'       End If
'
'    'then palin
'    strnum& = strnum& + 12
'    If val(numOrder(14)) <> 0 Then
'       FossilTbl$ = "PALYNOLOGY"
'       FosNum& = 14
'       Call PhpFossilInfo2(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, clPDF)
'       End If
'
'    'add page number
'    .DrawText 200, 10, "Page 1", "Fnt5", 12, pdfAlignRight
'
'    .EndPage
'
'    Dim j As Integer
'    Dim Xfos As Single, Yfos As Single
'
'    '-------------add a page for each fossil info----------------------
'
'    For i = 2 To PrintPreview.cmbPages.ListCount
'
'         FossilTbl$ = UCase$(Mid$(PrintPreview.cmbPages.List(i - 1), 9, Len(PrintPreview.cmbPages.List(i - 1)) - 8))
'
'         .BeginPage
'
'          Select Case LCase$(FossilTbl$)
'            Case "conodonta"
'                j = 0
'            Case "diatom"
'                j = 1
'            Case "foraminifera"
'                j = 2
'            Case "megafauna"
'                j = 3
'            Case "nannoplankton"
'                j = 4
'            Case "ostracoda"
'                j = 5
'            Case "palynology"
'                j = 6
'            Case Else
'          End Select
'
'          'header with logo
'          .DrawImg "Logo", 12, 288, 35, 35
'          If OrderNum& < 0 Then
'             .DrawText 52, 269, DocTitle$, "Fnt2", 14, pdfAlignLeft
'          Else
'             .DrawText 60, 269, DocTitle$, "Fnt2", 14, pdfAlignLeft
'             End If
'
'          'change text color to blue
'          .SetColorFill rgb(0, 0, 255)
'          .SetColorStroke rgb(0, 0, 255)
'          .DrawText 100 - 2.2 * (Len(FossilTbl$) - 6), 252, FossilTbl$, "Fnt2", 14, pdfAlignLeft
'
'          'restore text color to black
'          .SetColorFill rgb(0, 0, 0)
'          .SetColorStroke rgb(0, 0, 0)
'
'          'demarcation line
'          .SetLineCap 0
'          .SetLineWidth 0.4
'          .SetColorStroke rgb(0, 0, 0)
'          .MoveTo 15, 248
'          .LineTo 195, 248
'
'          '-----------------fossil data------------------------------------
'
'          Xo = 15: Yo = 240: Yspacing = 5
'          DocLine$ = "CHECK METHOD:  " & FossilTbl$
'         .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'          Yo = Yo - Yspacing
'          If IsNull(AnalystNames(j)) Then
'             DocLine$ = "ANALYST: "
'          Else
'             DocLine$ = "ANALYST: " & Trim$(AnalystNames(j))
'             End If
'         .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'          Yo = Yo - Yspacing
'          If IsNull(AnalysisDates(j)) Then
'             DocLine$ = "ANALYSIS DATE: "
'          Else
'             DocLine$ = "ANALYSIS DATE: " & Trim$(AnalysisDates(j))
'             End If
'
'         .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'         'draw another line
'         'demarcation line
'          Yo = Yo - Yspacing
'         .SetLineCap 0
'         .SetLineWidth 0.4
'         .SetColorStroke rgb(0, 0, 0)
'         .MoveTo 15, Yo
'         .LineTo 195, Yo
'
'          Yo = Yo - Yspacing
'          DocLine$ = "FOSSIL NAMES"
'         .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'          Xo = Xo + 85
'          DocLine$ = "SEMI QUANT"
'         .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'          Xo = Xo + 38
'          DocLine$ = "FEATURES"
'         .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
'          Xo = Xo + 38
'          DocLine$ = "QUANTITY"
'         .DrawText Xo, Yo, DocLine$, "Fnt5", 10, pdfAlignLeft
'
''           now query database for the fossil names,
''           semi quant, features, quantity
'
'          Dim FossilTag$, fosstbl$, FosTbl$, FosDic$
'
'          Yfos = Yo - Yspacing
'          Xfos = 15
'
'          FosIDCono = numOrder(18)
'          FosIDDiatom = numOrder(17)
'          FosIDForam = numOrder(12)
'          FosIDMega = numOrder(15)
'          FosIDNano = numOrder(16)
'          FosIDOstra = numOrder(13)
'          FosIDPaly = numOrder(14)
'
'
'          Select Case LCase$(FossilTbl$)
'             Case "conodonta"
'                  fosstbl$ = "condores"
'                  FosTbl$ = "condofos"
'                  FosDic$ = "Conodsdic"
'                  'query for foram fossil names
'                  Call PdfFosNames2(fosstbl$, FosTbl$, FosDic$, FosIDCono, Xfos, Yfos, Yspacing, clPDF)
'            Case "diatom"
'                  fosstbl$ = "diatores"
'                  FosTbl$ = "diatofos"
'                  FosDic$ = "Diatomsdic"
'                  'query for foram fossil names
'                  Call PdfFosNames2(fosstbl$, FosTbl$, FosDic$, FosIDDiatom, Xfos, Yfos, Yspacing, clPDF)
'            Case "foraminifera"
'                  fosstbl$ = "foramres"
'                  FosTbl$ = "foramfos"
'                  FosDic$ = "Foramsdic"
'                  'query for foram fossil names
'                  Call PdfFosNames2(fosstbl$, FosTbl$, FosDic$, FosIDForam, Xfos, Yfos, Yspacing, clPDF)
'            Case "megafauna"
'                  fosstbl$ = "megares"
'                  FosTbl$ = "megafos"
'                  FosDic$ = "Megadic"
'                  'query for foram fossil names
'                  Call PdfFosNames2(fosstbl$, FosTbl$, FosDic$, FosIDMega, Xfos, Yfos, Yspacing, clPDF)
'            Case "nannoplankton"
'                  fosstbl$ = "nanores"
'                  FosTbl$ = "nanofos"
'                  FosDic$ = "Nanodic"
'                  'query for foram fossil names
'                  Call PdfFosNames2(fosstbl$, FosTbl$, FosDic$, FosIDNano, Xfos, Yfos, Yspacing, clPDF)
'            Case "ostracoda"
'                  fosstbl$ = "ostrares"
'                  FosTbl$ = "ostrafos"
'                  FosDic$ = "Ostracoddic"
'                  'query for foram fossil names
'                  Call PdfFosNames2(fosstbl$, FosTbl$, FosDic$, FosIDOstra, Xfos, Yfos, Yspacing, clPDF)
'            Case "palynology"
'                  fosstbl$ = "palynres"
'                  FosTbl$ = "palynfos"
'                  FosDic$ = "Palyndic"
'                  'query for foram fossil names
'                  Call PdfFosNames2(fosstbl$, FosTbl$, FosDic$, FosIDPaly, Xfos, Yfos, Yspacing, clPDF)
'            Case Else
'         End Select
'
'       'add page number
'       .DrawText 200, 10, "Page " & Trim$(str$(i)), "Fnt5", 12, pdfAlignRight
'
'       .EndPage
'
'    Next i
'
'   .EndObject
'
'   .ClosePDFFile
'
'  End With
'
''  dblElapsed = Timer - dblElapsed
''  lblEnd.Caption = Format(dblElapsed, "0.00")
'
''  Command1.Enabled = True
'cp29998:
'    Dim iRet As Long
'
'    Select Case MsgBox("Do you want to view the pdf file?", vbQuestion + vbYesNo, App.Title)
'
'       Case vbYes
'         iRet = Shell("rundll32.exe url.dll,FileProtocolHandler " & (strFile), vbMaximizedFocus)
'
'       Case Else 'don't show it
'
'    End Select
'
''///////////////////////////////////////////////////////////////////////////////////////
''  '---------------------------------------------
''      'begin summary section (page 1) of report
'''    '-----------------------------------------------------------------------------
''    objPDF.PDFSetFont FONT_ARIAL, 10, FONT_normal
''
''     Xo = 40: Yo = 150: Yspacing = 15
''     DocLine$ = "CLIENT: " & numOrder(23)
''     objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''     Yo = Yo + Yspacing
''     DocLine$ = "COMPANY/DIVISION: " & numOrder(25)
''     objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''     Yo = Yo + Yspacing
''     DocLine$ = "PROJECT: " & numOrder(24)
''     objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''     Yo = Yo + Yspacing
''     DocLine$ = "FORMATION: " & numOrder(3)
''     objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''     If sAnum& = 1 Then 'well
''        Yo = Yo + Yspacing
''        DocLine$ = "SAMPLE METHOD: WELL"
''        objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''        Yo = Yo + Yspacing
''        DocLine$ = "WELL NAME: " & numOrder(0)
''        objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''     ElseIf sAnum& = 0 Then 'surface
''        Yo = Yo + Yspacing
''        DocLine$ = "SAMPLE METHOD: SURFACE"
''        objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''        Yo = Yo + Yspacing
''        DocLine$ = "PLACE NAME: " & numOrder(0)
''        objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''     ElseIf sAnum& = -1 Then 'unknown type
''        Yo = Yo + Yspacing
''        DocLine$ = "SAMPLE METHOD: UNKNOWN"
''        objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''        Yo = Yo + Yspacing
''        DocLine$ = "PLACE NAME: " & numOrder(0)
''        objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''        End If
''
''    Yo = Yo + Yspacing
''    DocLine$ = "ITMx: " & numOrder(1)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    Yo = Yo + Yspacing
''    DocLine$ = "ITMy: " & numOrder(2)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    Yo = Yo + Yspacing
''    TheText = "REMARKS: " & numOrder(11)
''
''    If Len(numOrder(11)) > 0 Then
''        'check for Hebrew
''        If HebrewCheck(numOrder(11)) Then
''
''           Screen.MousePointer = vbDefault
''
''           MsgBox "Hebrew characters where detected in the ''Remark''." _
''                  & vbCrLf & vbCrLf & "However, Hebrew fonts are not supported for pdf files." _
''                  & vbCrLf & vbCrLf & "(Consider reentering the remark to the database in English.)", _
''                  vbInformation + vbOKOnly, App.Title
''
''           'give a bit of time to close the message box and repaint
''           waitime = Timer
''           Do Until Timer > waitime + 0.5
''              DoEvents
''           Loop
''
''           Screen.MousePointer = vbHourglass
''
''           End If
''        End If
''
''    Max_Num_Lines = 5
''    SentanceLength = 100
''    CountLines = Int(Len(TheText) / 100#) + 1
'''        Call WordWrap(TheText, TheTextArray(), Max_Num_Lines, SentanceLength, CountLines)
'''        For i = 1 To CountLines
'''           Yo = Yo + Yspacing
'''           YoLastLeft = Yo
'''           DocLine$ = TheTextArray(i - 1)
'''           objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
'''        Next i
''     objPDF.PDFCell DocLine$, Xo, Yo, objPDF.PDFGetPageWidth - 2 * Xo, Yspacing
''     YoLastLeft = Yo + Yspacing * (CountLines - 1)
''
''    '________GENERAL SAMPLE INFO-FIRST RIGHT COLUMN___________
''
''
''    Xo = 370: Yo = 150: Yspacing = 15
''    DocLine$ = "LIM UP:         " & numOrder(5)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    Yo = Yo + Yspacing
''    DocLine$ = "LIM DOWN:    " & numOrder(4)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    If Val(numOrder(7)) = 1 And sAnum& = 1 Then
''       Yo = Yo + Yspacing
''       DocLine$ = "WELL SAMPLE TYPE: CUTTING"
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''    ElseIf Val(numOrder(7)) = 2 And sAnum& = 1 Then
''       Yo = Yo + Yspacing
''       DocLine$ = "WELL SAMPLE TYPE: CORE"
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''    Else
''       Yo = Yo + Yspacing
''       DocLine$ = "WELL SAMPLE TYPE: "
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''       End If
''
''ppd800:
''    Yo = Yo + Yspacing
''    DocLine$ = "CORE NUMBER: " & numOrder(26)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    Yo = Yo + Yspacing
''    DocLine$ = "BOX NUMBER: " & numOrder(27)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    Yo = Yo + Yspacing
''    DocLine$ = "FIELD NUMBER: " & numOrder(9)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    Yo = Yo + Yspacing
''    DocLine$ = "ORDER NUMBER: " & DocName$
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''    Yo = Yo + Yspacing
''    DocLine$ = "DATE: " & numOrder(21)
''    objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
'''
'''  '-----------print fossil and age information for scanned database
''    If OrderNum& < 0 Then
''
''       'draw line to delineate the fossil info from the summary
''       objPDF.PDFDrawLine 40, YoLastLeft + 25, objPDF.PDFGetPageWidth - 40, YoLastLeft + 25
''
''       Xo = 40: Yo = YoLastLeft + 40
''       DocLine$ = "FOSSIL CATEGORY: " & FossilTbl$
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       Yo = Yo + Yspacing
''       DocLine$ = "DATE: " & numOrder(21)
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       Yo = Yo + Yspacing
''       DocLine$ = "DATE: " & numOrder(21)
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       Yo = Yo + Yspacing
''       DocLine$ = "Earlier AGE: " & numOrder(28) & " " & numOrder(29)
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       Yo = Yo + Yspacing
''       DocLine$ = "LATER AGE: " & numOrder(30) & " " & numOrder(31)
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       Yo = Yo + Yspacing
''       DocLine$ = "Earlier ZONE: "
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       Yo = Yo + Yspacing
''       DocLine$ = "LATER ZONE: "
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       Yo = Yo + Yspacing
''       DocLine$ = "REMARK: "
''       objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''       AddPageNumber objPDF, 1
''
''       objPDF.PDFEndPage
''
'''        'now add tiff form image, if it exists, as a second page
'''
'''         If Dir(tifViewerDir$) <> sEmpty Then
'''            Call FindTifPath(Abs(OrderNum&), numOFile$)
'''
'''            Select Case numOFile$
'''              Case "-1" 'error flag
'''
'''                  'just close the document
'''
'''              Case Else 'add to second page of pdf
'''
'''                If Dir(tifDir$ & "\" & UCase$(numOFile$)) = sEmpty Then
'''
'''                   'just close the document
'''
'''                Else 'add to tiff file to second page
'''
'''                    objPDF.PDFNewPage
'''
'''                    objPDF.PDFImage tifDir$ & "\" & UCase$(numOFile$), _
'''                          0, 0, 500, 800, "http://www.gsi.gov.il"
'''
'''                    objPDF.PDFEndPage
'''
'''                   End If
'''
'''            End Select
'''            End If
''
''       objPDF.PDFEndDoc
''
''       GDMDIform.StatusBar1.Panels(1).Text = sEmpty
''       Screen.MousePointer = vbDefault
''
''       Exit Sub
''
''       End If
''
''
''    '--------summary of fossil results-----------------------------------------
''    'note: no checking is undertaken to see that the text doesn't go off the page
''    'this is highly unlikely, so hasn't been implemented
''    '--------------------------------------------------------------------------
''    'draw line to delineate the fossil info from the summary
''     objPDF.PDFDrawLine 40, YoLastLeft + 25, objPDF.PDFGetPageWidth - 40, YoLastLeft + 25
''
''     Xo = 40: Ylast = YoLastLeft + 15
''
''    'first conod
''    strnum& = 29
''    If Val(numOrder(18)) <> 0 Then
''       FossilTbl$ = "CONODONTA"
''       FosNum& = 18
''       Call PhpFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, objPDF)
''       End If
''
''
''      'then diato
''      strnum& = strnum& + 12
''      If Val(numOrder(17)) <> 0 Then
''         FossilTbl$ = "DIATOM"
''         FosNum& = 17
''         Call PhpFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, objPDF)
''         End If
''
''      'then foram
''      strnum& = strnum& + 12
''      If Val(numOrder(12)) <> 0 Then
''         FossilTbl$ = "FORAMINIFERA"
''         FosNum& = 12
''         Call PhpFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, objPDF)
''         End If
''
''      'then megaf
''      strnum& = strnum& + 12
''      If Val(numOrder(15)) <> 0 Then
''         FossilTbl$ = "MEGAFAUNA"
''         FosNum& = 15
''         Call PhpFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, objPDF)
''         End If
''
''      'then nanno
''      strnum& = strnum& + 12
''      If Val(numOrder(16)) <> 0 Then
''         FossilTbl$ = "NANNOPLANKTON"
''         FosNum& = 16
''         Call PhpFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, objPDF)
''         End If
''
''      'then ostra
''      strnum& = strnum& + 12
''      If Val(numOrder(13)) <> 0 Then
''         FossilTbl$ = "OSTRACODA"
''         FosNum& = 13
''         Call PhpFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, objPDF)
''         End If
''
''      'then palin
''      strnum& = strnum& + 12
''      If Val(numOrder(14)) <> 0 Then
''         FossilTbl$ = "PALYNOLOGY"
''         FosNum& = 14
''         Call PhpFossilInfo(FossilTbl$, numOrder, strnum&, FosNum&, Xo, Ylast, Yspacing, objPDF)
''         End If
''
''
''      AddPageNumber objPDF, 1
''
''      objPDF.PDFEndPage
''
''      Dim j As Integer, i As Integer
''
''      '-------------add a page for each fossil info----------------------
''      For i = 2 To PrintPreview.cmbPages.ListCount
''
''           FossilTbl$ = UCase$(Mid$(PrintPreview.cmbPages.List(i - 1), 9, Len(PrintPreview.cmbPages.List(i - 1)) - 8))
''
''           objPDF.PDFNewPage
''
''            Select Case LCase$(FossilTbl$)
''              Case "conodonta"
''                  j = 0
''              Case "diatom"
''                  j = 1
''              Case "foraminifera"
''                  j = 2
''              Case "megafauna"
''                  j = 3
''              Case "nannoplankton"
''                  j = 4
''              Case "ostracoda"
''                  j = 5
''              Case "palynology"
''                  j = 6
''              Case Else
''            End Select
''
''
'''            'Lets add a bookmark to the start of page 1
'''            objPDF.PDFSetBookmark "Page " & Trim$(Str$(i)), 0, 0
''
''            objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
''            objPDF.PDFSetDrawColor = vbWhite
''            objPDF.PDFSetAlignement = ALIGN_center
''            objPDF.PDFSetBorder = BORDER_none
''            objPDF.PDFSetFill = False
''
''            'add GSI logo
''            objPDF.PDFImage App.Path & "\gsi_03.jpg", _
''                  35, 25, 100, 100, "http://www.gsi.gov.il"
''
''            'add header
''            objPDF.PDFSetTextColor = vbBlack
''            objPDF.PDFCell DocTitle$, 100, 55, Len(DocTitle$) * 9, 40
''
''            'second header
''            objPDF.PDFSetTextColor = vbBlack
''            objPDF.PDFCell FossilTbl$, 100, 90, Len(DocTitle$) * 9, 40
''
''            'draw line to delineate the header
''            objPDF.PDFDrawLine 40, 140, objPDF.PDFGetPageWidth - 40, 140
''
''            '-----------------fossil data------------------------------------
''
''            objPDF.PDFSetFont FONT_ARIAL, 10, FONT_normal
''            objPDF.PDFSetAlignement = ALIGN_left
''
''            Xo = 40: Yo = 150: Yspacing = 15
''            DocLine$ = "CHECK METHOD:  " & FossilTbl$
''            objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 8, Yspacing
''
''            Xo = 40: Yo = Yo + Yspacing
''            If IsNull(AnalystNames(j)) Then
''               DocLine$ = "ANALYST: "
''            Else
''               DocLine$ = "ANALYST: " & Trim$(AnalystNames(j))
''               End If
''            objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 8, Yspacing
''
''            Xo = 40: Yo = Yo + Yspacing
''            If IsNull(AnalysisDates(j)) Then
''               DocLine$ = "ANALYSIS DATE: "
''            Else
''               DocLine$ = "ANALYSIS DATE: " & Trim$(AnalysisDates(j))
''               End If
''
''            objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 8, Yspacing
''
''            'draw another line
''            Yo = Yo + 25
''            objPDF.PDFDrawLine 40, Yo, objPDF.PDFGetPageWidth - 40, Yo
''
''            Yo = Yo + Yspacing
''            Xo = 40
''            DocLine$ = "FOSSIL NAMES"
''            objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''            Xo = Xo + 260
''            DocLine$ = "SEMI QUANT"
''            objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''            Xo = Xo + 100
''            DocLine$ = "FEATURES"
''            objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
''            Xo = Xo + 100
''            DocLine$ = "QUANTITY"
''            objPDF.PDFCell DocLine$, Xo, Yo, Len(DocLine$) * 10, Yspacing
''
'''           now query database for the fossil names,
'''           semi quant, features, quantity
''
''            Dim FossilTag$, fosstbl$, FosTbl$, FosDic$
''            Dim Xfos As Double, Yfos As Double
''
''            Yfos = Yo + Yspacing
''            Xfos = 40
''
''            FosIDCono = numOrder(18)
''            FosIDDiatom = numOrder(17)
''            FosIDForam = numOrder(12)
''            FosIDMega = numOrder(15)
''            FosIDNano = numOrder(16)
''            FosIDOstra = numOrder(13)
''            FosIDPaly = numOrder(14)
''
''
''            Select Case LCase$(FossilTbl$)
''               Case "conodonta"
''                    fosstbl$ = "condores"
''                    FosTbl$ = "condofos"
''                    FosDic$ = "Conodsdic"
''                    'query for foram fossil names
''                    Call PdfFosNames(fosstbl$, FosTbl$, FosDic$, FosIDCono, Xfos, Yfos, Yspacing, objPDF)
''              Case "diatom"
''                    fosstbl$ = "diatores"
''                    FosTbl$ = "diatofos"
''                    FosDic$ = "Diatomsdic"
''                    'query for foram fossil names
''                    Call PdfFosNames(fosstbl$, FosTbl$, FosDic$, FosIDDiatom, Xfos, Yfos, Yspacing, objPDF)
''              Case "foraminifera"
''                    fosstbl$ = "foramres"
''                    FosTbl$ = "foramfos"
''                    FosDic$ = "Foramsdic"
''                    'query for foram fossil names
''                    Call PdfFosNames(fosstbl$, FosTbl$, FosDic$, FosIDForam, Xfos, Yfos, Yspacing, objPDF)
''              Case "megafauna"
''                    fosstbl$ = "megares"
''                    FosTbl$ = "megafos"
''                    FosDic$ = "Megadic"
''                    'query for foram fossil names
''                    Call PdfFosNames(fosstbl$, FosTbl$, FosDic$, FosIDMega, Xfos, Yfos, Yspacing, objPDF)
''              Case "nannoplankton"
''                    fosstbl$ = "nanores"
''                    FosTbl$ = "nanofos"
''                    FosDic$ = "Nanodic"
''                    'query for foram fossil names
''                    Call PdfFosNames(fosstbl$, FosTbl$, FosDic$, FosIDNano, Xfos, Yfos, Yspacing, objPDF)
''              Case "ostracoda"
''                    fosstbl$ = "ostrares"
''                    FosTbl$ = "ostrafos"
''                    FosDic$ = "Ostracoddic"
''                    'query for foram fossil names
''                    Call PdfFosNames(fosstbl$, FosTbl$, FosDic$, FosIDOstra, Xfos, Yfos, Yspacing, objPDF)
''              Case "palynology"
''                    fosstbl$ = "palynres"
''                    FosTbl$ = "palynfos"
''                    FosDic$ = "Palyndic"
''                    'query for foram fossil names
''                    Call PdfFosNames(fosstbl$, FosTbl$, FosDic$, FosIDPaly, Xfos, Yfos, Yspacing, objPDF)
''              Case Else
''           End Select
''
''           AddPageNumber objPDF, i          'page number
''
''           objPDF.PDFEndPage
''
''      Next i
'''
'
'
'End Sub
