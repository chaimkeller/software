Attribute VB_Name = "modPrintPreview"
Option Explicit


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
      BitmapInfoHeader As BITMAPINFOHEADER_TYPE
      bmiColors As String * 1024
   End Type

   ' Enter each of the following Declare statements as one, single line:
   Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, _
      ByVal hBitmap As Long, ByVal nStartScan As Long, _
      ByVal nNumScans As Long, ByVal lpBits As Long, _
      BitmapInfo As BITMAPINFO_TYPE, ByVal wUsage As Long) As Long
   Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, _
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

   
   Public magPrint% 'zoom percentage
   Public finishedloading As Boolean 'flag to tell program when finished loading zoomcombo
   Public PicWidth, PicHeight 'stored initial values of picture1.width/height
   Public PicLeft, PicTop 'stored initial values of picture1.left/top
   Public zoomfactor% 'zoom is defined by (1 + zoomfactor%*0.1)
                      'not used in current project
   Public LoadInit As Boolean
   
'----------------screen resoultion variables---------------
   Type RECT
       x1 As Long
       y1 As Long
       x2 As Long
       y2 As Long
   End Type

   ' NOTE: The following declare statements are case sensitive.

   Declare Function GetDesktopWindow Lib "User32" () As Long
   Declare Function GetWindowRect Lib "User32" _
      (ByVal hWnd As Long, rectangle As RECT) As Long

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
          Dim hWnd As Long
          Dim RetVal As Long
          hWnd = GetDesktopWindow()
          RetVal = GetWindowRect(hWnd, R)
          GetScreenResolution = (R.x2 - R.x1) & "x" & (R.y2 - R.y1)
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
      psm = Printer.ScaleMode
      Printer.ScaleMode = 5 'Inches
      LRGap = (PgWidth - Printer.ScaleWidth) / 2
      TBGap = (PgHeight - Printer.ScaleHeight) / 2
      Printer.ScaleMode = psm

      ' Initialize printer or preview object:
      If PrinterFlag Then
         sm = Printer.ScaleMode
         Printer.ScaleMode = 5 'Inches
         Printer.Print sEmpty;
      Else
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
         ' Set default properties of picture box to match printer
         ' There are many that you could add here:
         ObjPrint.Scale (0, 0)-(PgWidth, PgHeight)
         ObjPrint.FontName = Printer.FontName
         ObjPrint.FontSize = Printer.FontSize * Ratio
         ObjPrint.ForeColor = Printer.ForeColor
         ObjPrint.BackColor = QBColor(15) 'CHANGES!!!!!!
         ObjPrint.Cls
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

      Dim BitmapInfo As BITMAPINFO_TYPE
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
         DesthDC = Printer.hDC
      Else
         ObjPrint.Scale
         ObjPrint.ScaleMode = 3 'Pixels
         ' Calculate size in pixels:
         pLeft = ((pLeft * 1440) / Screen.TwipsPerPixelX) * Ratio
         pTop = ((pTop * 1440) / Screen.TwipsPerPixelY) * Ratio
         pWidth = ((pWidth * 1440) / Screen.TwipsPerPixelX) * Ratio
         pHeight = ((pHeight * 1440) / Screen.TwipsPerPixelY) * Ratio
         DesthDC = ObjPrint.hDC
      End If

      BitmapInfo.BitmapInfoHeader.biSize = 40
      BitmapInfo.BitmapInfoHeader.biWidth = picSource.ScaleWidth
      BitmapInfo.BitmapInfoHeader.biHeight = picSource.ScaleHeight
      BitmapInfo.BitmapInfoHeader.biPlanes = 1
      BitmapInfo.BitmapInfoHeader.biBitCount = 8
      BitmapInfo.BitmapInfoHeader.biCompression = BI_RGB

     
      hMem = GlobalAlloc(GMEM_MOVEABLE, (CLng(picSource.ScaleWidth + 3) _
         \ 4) * 4 * picSource.ScaleHeight) 'DWORD ALIGNED
      lpBits = GlobalLock(hMem)

      
      R = GetDIBits(picSource.hDC, picSource.Image, 0, _
         picSource.ScaleHeight, lpBits, BitmapInfo, DIB_RGB_COLORS)
      If R <> 0 Then
         
         R = StretchDIBits(DesthDC, pLeft, pTop, pWidth, pHeight, 0, 0, _
            picSource.ScaleWidth, picSource.ScaleHeight, lpBits, _
            BitmapInfo, DIB_RGB_COLORS, SRCCOPY)
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
         If Not ScreenDump Then 'send search results to printer
            Printer.EndDoc
            Printer.ScaleMode = sm
         Else 'send screen dump of map to printer
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
   PaprSize = Printer.PaperSize
   
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
   
   If ScreenDump Then
      'make special page size to fit entire screen dump
      'this will depend on screen resolution
      'For 800 x 600 then 8.8" x 13" is enough
      
      'call GetScreenResolution function to return resolution string
      'E.g., if resolution is 800 x 600, then it returns "800x600"
      Dim pos%, ResMul As Single
      pos% = InStr(GetScreenResolution, "x") 'find X resolution
      ResMul = Val(Mid$(GetScreenResolution, 1, pos% - 1)) / 800 'relative resolution to 800
      PrintPreview.cmbtxtPaper.Text = "Special"
      
      PgWidth = 8.9 * ResMul
      PgHeight = 13 * ResMul
      
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
   
   change = False
   If Printer.Orientation <> PaperOrientation Then
      change = True
      End If

   'PaperOrientation = Printer.Orientation
50 Select Case PaperOrientation
      Case 1 'portrait
         Printer.Orientation = vbPRORPortrait
         PrintPreview.ImgLandscape.Visible = False
         PrintPreview.imgPortrait.Visible = True
         PrintPreview.optPortrait.Value = True
         PrintPreview.ImgLandscape.ToolTipText = gsEmpty
         PrintPreview.imgPortrait.ToolTipText = "Portrait orientation"
      Case 2 'landscape
         Printer.Orientation = vbPRORLandscape
         PrintPreview.imgPortrait.Visible = False
         PrintPreview.ImgLandscape.Visible = True
         PrintPreview.optLandscape.Value = True
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
       
   If change And Not LoadInit Then
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

  sTmp = "The following Error occurred:" & vbCrLf & vbCrLf
  'add the error string
  sTmp = sTmp & Err.Description & vbCrLf
  'add the error number
  sTmp = sTmp & "VB Error Number: " & Err & vbCrLf & vbCrLf
  'add a suggestion
  sTmp = sTmp & "Check your printer's settings!"
  
  Beep

  MsgBox sTmp, vbCritical + vbOKOnly, "Print Preview"
  Err.Clear

End Sub

Sub ScrollBars()
   'settings for scroll bars
   
   'horizontal scroll bar
   With PrintPreview.HScroll1
      If PrintPreview.Picture1.Width + PrintPreview.Picture1.Left > PrintPreview.RightBorderPictureBox.Left Then
        .Visible = True
        .Left = 0
        .Width = PrintPreview.BottomBorderPictureBox.Width
        .Top = 0
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
      If PrintPreview.Picture1.Top + PrintPreview.Picture1.Height > PrintPreview.BottomBorderPictureBox.Top Then
        .Visible = True
        .Left = 0
        .Top = 0
        .Height = PrintPreview.RightBorderPictureBox.Height - PrintPreview.BottomBorderPictureBox.Height - PrintPreview.TopBorderPictureBox.Top
        'settings for it
        .Max = PrintPreview.Picture2.Height + PrintPreview.Picture1.Height
        .LargeChange = .Max / 30
        .SmallChange = .Max / 60
      Else
        .Visible = False
      End If
   End With
End Sub

Sub PositionBorders()

      On Error Resume Next
      
'     positions of picture boxes that act as the form's borders

      With PrintPreview.TopBorderPictureBox
           .Left = -10
           .Width = PrintPreview.Width '9700
           .Height = 735
           .Top = -60
      End With
      
      With PrintPreview.LeftBorderPictureBox
        .Left = 0
        .Width = 495
        .Top = PrintPreview.TopBorderPictureBox.Top + PrintPreview.TopBorderPictureBox.Height - 10 '667
        .Height = PrintPreview.Height - .Top - 400 '7523
      End With
      
      With PrintPreview.RightBorderPictureBox
         .Width = 495 '675
         .Left = PrintPreview.Width - .Width - 120 '9120
         .Top = PrintPreview.TopBorderPictureBox.Top + PrintPreview.TopBorderPictureBox.Height - 10 '667
         .Height = PrintPreview.LeftBorderPictureBox.Height
      End With
      
      With PrintPreview.BottomBorderPictureBox
         .Left = PrintPreview.LeftBorderPictureBox.Width  'assumes that leftpicturebox is at .left=0 '480
         .Width = PrintPreview.RightBorderPictureBox.Left - PrintPreview.LeftBorderPictureBox.Width '8685
         .Top = PrintPreview.LeftBorderPictureBox.Top + PrintPreview.LeftBorderPictureBox.Height - 450 '7740
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
   
    Dim x As Printer, numPrinter%, foundPrinter%
    numPrinter% = 0
    For Each x In Printers
        PrintPreview.cmbtxtPrinter.AddItem x.DeviceName
        numPrinter% = numPrinter% + 1
        If x.DeviceName = Printer.DeviceName Then
           'this is default printer
           Set Printer = x
           foundPrinter% = numPrinter%
           End If
    Next
    
    'show the default printer
    PrintPreview.cmbtxtPrinter.ListIndex = foundPrinter% - 1
    Exit Sub
   
errhand:
   ShowPreviewError

End Sub

Sub LoadPaperOrientation()
   'set default paper orientation
   
   On Error GoTo errhand
   
   If Printer.Orientation = vbPRORPortrait Then
      PrintPreview.optPortrait.Value = True
   ElseIf Printer.Orientation = vbPRORLandscape Then
      PrintPreview.optLandscape.Value = True
      End If
      
   If LoadInit And ScreenDump Then
      'set default as landscape
      PrintPreview.optLandscape.Value = True
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
       PreviewPrintScreen
'------------------------------------------------------------------------------
      
      PrintEndDoc
      If PrinterFlag = True Then
         PrinterFlag = False
         End If
         
      Exit Sub
      
errhand:
   ShowPreviewError
      
End Sub
Sub PreviewPrintScreen()
    
    'clear out the printpreview picture box
    Set PrintPreview.Picture1.Picture = Nothing
    Set PrintPreview.Picture2.Picture = Nothing
    
    'now capture the plot
    Set PrintPreview.Picture1.Picture = CaptureClient(frmDraw)
    
End Sub
