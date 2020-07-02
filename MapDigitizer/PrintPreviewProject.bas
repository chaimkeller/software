Attribute VB_Name = "Module1"
Option Explicit


   ' The following Types, Declares, and Constants are only necessary
   ' for the PrintPicture routine
   '=======================================================================
   Type BITMAPINFOHEADER_TYPE
      biSize As Long
      biWidth As Long
      biHeight As Long
      biPlanes As Integer
      biBitCount As Integer
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
   Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Integer, _
      ByVal hBitmap As Integer, ByVal nStartScan As Integer, _
      ByVal nNumScans As Integer, ByVal lpBits As Long, _
      BitmapInfo As BITMAPINFO_TYPE, ByVal wUsage As Integer) As Integer
   Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Integer, _
      ByVal DestX As Integer, ByVal DestY As Integer, _
      ByVal wDestWidth As Integer, ByVal wDestHeight As Integer, _
      ByVal SrcX As Integer, ByVal SrcY As Integer, _
      ByVal wSrcWidth As Integer, ByVal wSrcHeight As Integer, _
      ByVal lpBits As Long, BitsInfo As BITMAPINFO_TYPE, _
      ByVal wUsage As Integer, ByVal dwRop As Long) As Integer
   Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, _
      ByVal lMem As Long) As Integer
   Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Long
   Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
   Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) _
      As Integer
      
      

   Global Const SRCCOPY = &HCC0020
   Global Const BI_RGB = 0
   Global Const DIB_RGB_COLORS = 0
   Global Const GMEM_MOVEABLE = 2

   ' Module level variables set in PrintStartDoc flag indicating Printing
   ' or Previewing:
   Dim PrinterFlag
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
   Dim PgWidth
   Dim PgHeight

   Sub PrintStartDoc(objToPrintOn As Control, PF, PaperWidth, PaperHeight)
      Dim psm
      Dim fsm
      Dim HeightRatio
      Dim WidthRatio

      ' Set the flag that determines whether printing or previewing:
      PrinterFlag = PF

      ' Set the physical page size:
      PgWidth = PaperWidth
      PgHeight = PaperHeight

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
         Printer.Print "";
      Else
         ' Set the object used for preview:
         Set ObjPrint = objToPrintOn
         ' Scale Object to Printer's printable area in Inches:
         sm = ObjPrint.ScaleMode
         ObjPrint.ScaleMode = 5 'Inches
         ' Compare the height and with ratios to determine the
         ' Ratio to use and how to size the picture box:
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
         ObjPrint.Cls
      End If
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

   Sub PrintPrint(PrintVar)
      If PrinterFlag Then
         Printer.Print PrintVar
      Else
         ObjPrint.Print PrintVar
      End If
   End Sub

   Sub PrintLine(bLeft0, bTop0, bLeft1, bTop1)
      If PrinterFlag Then
         ' Enter the following two lines as one, single line:
         Printer.Line (bLeft0 - LRGap, bTop0 - TBGap)- _
            (bLeft1 - LRGap, bTop1 - TBGap)
      Else
         ObjPrint.Line (bLeft0, bTop0)-(bLeft1, bTop1)
      End If
   End Sub

   Sub PrintBox(bLeft, bTop, bWidth, bHeight)
      If PrinterFlag Then
         ' Enter the following two lines as one, single line:
         Printer.Line (bLeft - LRGap, bTop - TBGap)- _
            (bLeft + bWidth - LRGap, bTop + bHeight - TBGap), , B
      Else
         ObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), , B
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

   Sub PrintCircle(bLeft, bTop, bRadius)
      If PrinterFlag Then
         Printer.Circle (bLeft - LRGap, bTop - TBGap), bRadius
      Else
         ObjPrint.Circle (bLeft, bTop), bRadius
      End If
   End Sub

   Sub PrintNewPage()
      If PrinterFlag Then
         Printer.NewPage
      Else
         ObjPrint.Cls
      End If
   End Sub

   ' Enter the following two lines as one, single line:
   Sub PrintPicture(picSource As Control, ByVal pLeft, ByVal pTop, _
      ByVal pWidth, ByVal pHeight)

      ' Picture Box should have autoredraw = False, ScaleMode = Pixel
      ' Also can have visible=false, Autosize = true

      Dim BitmapInfo As BITMAPINFO_TYPE
      Dim DesthDC As Integer
      Dim hMem As Integer
      Dim lpBits As Long
      Dim r As Integer

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
         Printer.Print "";
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

      ' Enter the following two lines as one, single line:
      hMem = GlobalAlloc(GMEM_MOVEABLE, (CLng(picSource.ScaleWidth + 3) _
         \ 4) * 4 * picSource.ScaleHeight) 'DWORD ALIGNED
      lpBits = GlobalLock(hMem)

      ' Enter the following two lines as one, single line:
      r = GetDIBits(picSource.hDC, picSource.Image, 0, _
         picSource.ScaleHeight, lpBits, BitmapInfo, DIB_RGB_COLORS)
      If r <> 0 Then
         ' Enter the following two lines as one, single line:
         r = StretchDIBits(DesthDC, pLeft, pTop, pWidth, pHeight, 0, 0, _
            picSource.ScaleWidth, picSource.ScaleHeight, lpBits, _
            BitmapInfo, DIB_RGB_COLORS, SRCCOPY)
      End If

      r = GlobalUnlock(hMem)
      r = GlobalFree(hMem)

      If PrinterFlag Then
         Printer.ScaleMode = 5 'Inches
      Else
         ObjPrint.ScaleMode = 5 'Inches
         ObjPrint.Scale (0, 0)-(PgWidth, PgHeight)
      End If
   End Sub

   Sub PrintEndDoc()
      If PrinterFlag Then
         Printer.EndDoc
         Printer.ScaleMode = sm
      Else
         ObjPrint.ScaleMode = sm
      End If
   End Sub


