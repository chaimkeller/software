VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form PrintPreview 
   BackColor       =   &H00808080&
   Caption         =   "Print Preview"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "PrintPreviewForm1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   9600
   WindowState     =   2  'Maximized
   Begin VB.PictureBox RightBorderPictureBox 
      BorderStyle     =   0  'None
      Height          =   7523
      Left            =   9120
      ScaleHeight     =   7530
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   667
      Width           =   495
      Begin VB.VScrollBar VScroll1 
         Height          =   7095
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox BottomBorderPictureBox 
      BorderStyle     =   0  'None
      Height          =   444
      Left            =   480
      ScaleHeight     =   450
      ScaleWidth      =   8685
      TabIndex        =   4
      Top             =   7740
      Width           =   8685
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   8475
      End
   End
   Begin VB.PictureBox LeftBorderPictureBox 
      BorderStyle     =   0  'None
      Height          =   7523
      Left            =   0
      ScaleHeight     =   7530
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   667
      Width           =   495
   End
   Begin VB.PictureBox TopBorderPictureBox 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   9675
      TabIndex        =   2
      Top             =   -60
      Width           =   9700
      Begin VB.OptionButton optLandscape 
         Caption         =   "Landscape"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5460
         TabIndex        =   19
         Top             =   420
         Width           =   975
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Portrait"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5460
         TabIndex        =   18
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cmbtxtPaper 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3960
         TabIndex        =   17
         Text            =   "cmbtxtPaper"
         Top             =   190
         Width           =   1335
      End
      Begin VB.ComboBox cmbtxtPrinter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7770
         TabIndex        =   16
         Text            =   "cmbtxtPrinter"
         Top             =   220
         Width           =   1755
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   960
         Picture         =   "PrintPreviewForm1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Send to Printer"
         Top             =   180
         Width           =   435
      End
      Begin VB.CommandButton cmdPrintSetup 
         Height          =   375
         Left            =   480
         Picture         =   "PrintPreviewForm1.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Printer Setup"
         Top             =   180
         Width           =   435
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   900
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdMagnify 
         Height          =   375
         Left            =   1320
         Picture         =   "PrintPreviewForm1.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   -120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdDemagnify 
         Height          =   375
         Left            =   1920
         Picture         =   "PrintPreviewForm1.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   -120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox ZoomCombo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         TabIndex        =   6
         Text            =   "200%"
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label lblPrinter 
         Caption         =   "Printer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Image ImgLandscape 
         Enabled         =   0   'False
         Height          =   480
         Left            =   6480
         Picture         =   "PrintPreviewForm1.frx":0DF2
         ToolTipText     =   "Landscape Orientation"
         Top             =   180
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgPortrait 
         Enabled         =   0   'False
         Height          =   480
         Left            =   6480
         Picture         =   "PrintPreviewForm1.frx":1534
         ToolTipText     =   "Portrait Orientation"
         Top             =   165
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblPaper 
         Caption         =   "Paper:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblZoom 
         Caption         =   "Zoom:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1570
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   10585
      Left            =   1080
      ScaleHeight     =   11
      ScaleMode       =   0  'User
      ScaleWidth      =   8.5
      TabIndex        =   0
      Top             =   1020
      Width           =   8012
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   3000
         Left            =   2160
         ScaleHeight     =   2940
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   780
         Width           =   2355
      End
   End
End
Attribute VB_Name = "PrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbtxtPaper_Click()
'if new paper size was chosen, then change
'printer setting to reflect this

   On Error GoTo errhand
   
   If PaprSize <> PrintPreview.cmbtxtPaper.ListIndex + 1 Then
      'reset printer settings
      Printer.PaperSize = PrintPreview.cmbtxtPaper.ListIndex + 1
      FindPaperSize

      FindPaperOrientation
      
      'resize the preview
      PreviewSetup 'initialize picture boxes
      
      'set up landscape preview if requested
      If Printer.Orientation = 2 Then 'landscape
         Dim PgWidthTmp
         PgWidthTmp = PgWidth
         PgWidth = PgHeight
         PgHeight = PgWidthTmp
         End If
      
      'now scale picture box with respect to Letter size
      Picture1.Width = PicWidth * PgWidth / 8.5
      Picture1.Height = PicHeight * PgHeight / 11
      
      PreviewPrint 'Execute Printing/Previewing
      
      'now display in requested magnification
      ZoomCombo_click
      
      End If

     
   Exit Sub
   
errhand:
   Select Case Err.Number
      Case 380
         MsgBox "Your printer doesn't support this paper size!", vbExclamation + vbOKOnly, "Print Preview"
      Case Else
         ShowPreviewError
   End Select
   'reset combo box
   PrintPreview.cmbtxtPaper.ListIndex = Printer.PaperSize - 1

End Sub

Private Sub cmbtxtPrinter_Click()
  
  On Error GoTo errhand
  
  Dim NewDeviceName As String
  
  'check if printer was changed
  If PrintPreview.cmbtxtPrinter.List(PrintPreview.cmbtxtPrinter.ListIndex) <> Printer.DeviceName Then
    'reinitialize the printer
    NewDeviceName = PrintPreview.cmbtxtPrinter.List(PrintPreview.cmbtxtPrinter.ListIndex)
    Dim X As Printer
    For Each X In Printers
        If X.DeviceName = NewDeviceName Then
           'this is default printer
           Set Printer = X
           Exit For
           End If
    Next
    
    'set the papersize and paperorientation
    Printer.Orientation = PaperOrientation
    Printer.PaperSize = PaprSize
    
    End If
     
  Exit Sub
   
errhand:
   ShowPreviewError
     
End Sub

Private Sub cmdDemagnify_Click()
   zoomfactor% = zoomfactor% - 1
   If 1 + 0.1 * zoomfactor% <= 0.1 Then
     cmdDemagnify.Enabled = False
     Exit Sub
     End If
   cmdMagnify.Enabled = True
   Picture1.Width = PicWidth * (1 + zoomfactor% * 0.1)
   Picture1.Height = PicHeight * (1 + zoomfactor% * 0.1)
   
   'refresh preview
   PreviewPrint
End Sub

Private Sub cmdMagnify_Click()
   zoomfactor% = zoomfactor% + 1
   If 1 + 0.1 * zoomfactor >= 3 Then
      cmdMagnify.Enabled = False
      Exit Sub
      End If
   cmdDemagnify.Enabled = True
   Picture1.Width = PicWidth * (1 + zoomfactor% * 0.1)
   Picture1.Height = PicHeight * (1 + zoomfactor% * 0.1)
   
   'Execute new preview
   PreviewPrint
End Sub

Private Sub cmdPrintSetup_Click()
    Dim BeginPage, EndPage, NumCopies, I
    ' Set Cancel to True.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    
    ' Display the Print Setup dialog box.
    CommonDialog1.flags = cdlPDPrintSetup
    CommonDialog1.ShowPrinter
    
    ' Get user-selected values from the dialog box.
    
    Dim X As Printer
    For Each X In Printers
        If X.DeviceName = Printer.DeviceName Then
           'this is default printer
           Set Printer = X
           Exit For
           End If
    Next
    
    If ScreenDump Then Exit Sub 'leave everything else alone
    
    FindPaperSize 'determine paper size set for the printer
    FindPaperOrientation 'determine paper orientation
    'FindPrinterName 'find name of printer
    
    'refresh printpreview
    PrinterFlag = False
    PreviewPrint
    
    'redisplay new preview at requested magnification
    ZoomCombo_click
    
    Exit Sub

ErrHandler:
    ' User pressed Cancel button.
    'check if Print Setting was changed
    
    For Each X In Printers
        If X.DeviceName = Printer.DeviceName Then
           'this is default printer
           Set Printer = X
           Exit For
           End If
    Next
    
    FindPaperSize 'determine paper size set for the printer
    FindPaperOrientation 'determine paper orientation
    'FindPrinterName 'find name of printer
    
    'refresh printpreview
    PrinterFlag = False
    PreviewPrint
    
    'redisplay new preview at current zoom
    ZoomCombo_click

End Sub

Private Sub cmdPrint_Click()
    Dim BeginPage, EndPage, NumCopies, I
    ' Set Cancel to True.
    
    If ScreenDump Then  'just print and exit
       PrinterFlag = True
       Call PrintEndDoc
       PrinterFlag = False
       Exit Sub
       End If
    
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    
    ' Display the Print dialog box.
    CommonDialog1.flags = 0
    CommonDialog1.ShowPrinter
    ' Get user-selected values from the dialog box.
    
    Dim X As Printer
    For Each X In Printers
        If X.DeviceName = Printer.DeviceName Then
           'this is default printer
           Set Printer = X
           Exit For
           End If
    Next
    
50  FindPaperSize 'determine paper size set for the printer
    FindPaperOrientation 'determine paper orientation
    'FindPrinterName 'find name of printer
    
    'refresh printpreview if printer settings were changed
    PrinterFlag = False
    PreviewPrint
    
    'present new preview at current zoom if printer settings were changed
    ZoomCombo_click
    
    'Print!
    NumCopies = CommonDialog1.Copies
    For I = 1 To NumCopies
       'send document to your printer.
       PrinterFlag = True
       PreviewPrint
    Next
    
    Exit Sub
    
ErrHandler:
    ' User pressed Cancel button.
    'check if Print Setting was changed
    For Each X In Printers
        If X.DeviceName = Printer.DeviceName Then
           'this is default printer
           Set Printer = X
           Exit For
           End If
    Next
    
    FindPaperSize 'determine paper size set for the printer
    FindPaperOrientation 'determine paper orientation
    'FindPrinterName 'find name of printer
    
    'refresh printpreview
    PrinterFlag = False
    PreviewPrint
    
    'present new preview at current zoom if printer settings were changed
    ZoomCombo_click
 

End Sub


Private Sub Form_Load()
On Error GoTo errhand

      'put borders in right place
      PositionBorders
      
      'set up scroll bars if necessary
      ScrollBars
      
      'Set up zoom control
      finishedloading = False 'Just Started!
      ZoomCombo.Clear
      For I% = 1 To 27
         ZoomCombo.AddItem LTrim$(Str(40 + 10 * (I% - 1)) & "%")
      Next I%
      ZoomCombo.ListIndex = 6
      magPrint% = Mid(ZoomCombo.Text, 1, Len(ZoomCombo.Text) - 1)
      finishedloading = True 'Finished
      
      'Store initial picture1 dimensions
      PicWidth = Picture1.Width
      PicHeight = Picture1.Height
      PicLeft = Picture1.Left
      PicTop = Picture1.Top
      
      LoadInit = True 'load defaults
      LoadPrinterName 'load combo box with available printers
      LoadPaperOrientation 'set default paper orientation
      LoadPaperSize 'load combo box with paper sizes
    
      FindPaperSize 'determine paper size set for the printer
      FindPaperOrientation 'determine paper orientation
         
      LoadInit = False 'finished loading defaults
      
      PreviewSetup 'initialize picture boxes
      
      'now scale picture box with respect to Letter size
      Picture1.Width = PicWidth * PgWidth / 8.5
      Picture1.Height = PicHeight * PgHeight / 11

      PreviewPrint 'Execute Printing/Previewing
      
      'freeze controls if previewing maps
      '(since format is not variable)
      If ScreenDump Then
         optPortrait.Enabled = False
         optLandscape.Enabled = False
         cmbtxtPaper.Enabled = False
         cmbtxtPrinter.Enabled = False
         ZoomCombo.Enabled = False
         cmdPrintSetup.Enabled = False
         ImgLandscape.Enabled = False
         End If
      
      Exit Sub
      
errhand:
   ShowPreviewError

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ScreenDump = False
  Unload Me
  Set PrintPreview = Nothing
End Sub

Private Sub Form_Resize()
      'put borders in right place
      PositionBorders
      
      'rezero scrollbars and pictures
      PrintPreview.HScroll1.Value = 0
      PrintPreview.VScroll1.Value = 0
      PrintPreview.Picture1.Top = PicTop
      PrintPreview.Picture1.Left = PicLeft
      
      'set up scroll bars if necessary
      ScrollBars

End Sub

Private Sub HScroll1_Change()
   Picture1.Left = PicLeft - HScroll1.Value
End Sub


Private Sub optLandscape_Click()
   PaperOrientation = 2
   FindPaperOrientation
End Sub

Private Sub optPortrait_Click()
   PaperOrientation = 1
   FindPaperOrientation
End Sub

Private Sub VScroll1_Change()
   Picture1.Top = PicTop - VScroll1.Value
End Sub

Public Sub ZoomCombo_click()
   On Error GoTo errhand

   'Change the dimensions of the previewed page according to
   'the zoom.
   If finishedloading = False Then Exit Sub
   magPrint% = Mid(ZoomCombo.Text, 1, Len(ZoomCombo.Text) - 1)
   If magPrint% < 10 Or magPrint% > 400 Then
     Beep
     Exit Sub
     End If
     
   'make sure that correct scaling exist between current
   'paper size and Letter (the default) size
   Picture1.Width = PicWidth * magPrint% * (PgWidth / 8.5) / 100
   Picture1.Height = PicHeight * magPrint% * (PgHeight / 11) / 100
   
   'refresh preview
   PreviewPrint
   
   'Rezero scrollbars and pictures to make sure that
   'part of picture doesn't becomes hidden as scale changes.
   PrintPreview.HScroll1.Value = 0
   PrintPreview.VScroll1.Value = 0
   PrintPreview.Picture1.Top = PicTop
   PrintPreview.Picture1.Left = PicLeft
   
   'Reset scroll bars as necessary.
   ScrollBars

   Exit Sub
   
errhand:
   ShowPreviewError
   
End Sub

Private Sub ZoomCombo_KeyPress(KeyAscii As Integer)
   
   'This sub allows user to enter variable zoom within
   'the permissible range (which is set conservatively to
   'fit with most computers' graphic's memory)
   
   On Error GoTo errhand
   
   Select Case KeyAscii
      Case 13 'carriage return
         'check for % sign
         'Check if it is integer with/without percentage sign
          ZoomCombo.Text = LTrim(RTrim(ZoomCombo.Text))
          For I% = 1 To Len((ZoomCombo.Text)) - 1
               If InStr("0123456789", Mid(Trim$(ZoomCombo.Text), I%, 1)) = 0 Then
                  'non numerical values, exit
                  response = MsgBox("Enter a nonnegative integer", vbCritical + vbOKOnly, "Gdb")
                  Exit Sub
                  End If
          Next I%
         'now check that percentage is at end if it is there
         If InStr(Mid(ZoomCombo.Text, Len(ZoomCombo.Text), 1), "%") <> 0 Then
            'leave it alone
         Else 'add % sign
            ZoomCombo.Text = ZoomCombo.Text & "%"
            End If
         ZoomCombo_click 'magnify it
      Case Else
   End Select
   
   'Rezero scrollbars and pictures to make sure that
   'part of picture doesn't becomes hidden as scale changes.
   PrintPreview.HScroll1.Value = 0
   PrintPreview.VScroll1.Value = 0
   PrintPreview.Picture1.Top = PicTop
   PrintPreview.Picture1.Left = PicLeft
   
   'Reset scroll bars as necessary.
   ScrollBars
   
   Exit Sub

errhand:
   ShowPreviewError
  
   
End Sub

