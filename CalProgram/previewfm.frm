VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form previewfm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "previewfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11865
   Begin VB.CommandButton schultimbut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shul  &times"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1740
      Width           =   855
   End
   Begin VB.CommandButton zmanbut 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Z'manim"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11040
      Picture         =   "previewfm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Define zemanim entries"
      Top             =   1020
      Width           =   855
   End
   Begin VB.CommandButton prevCLIPbut 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Clipboard"
      Height          =   555
      Left            =   11040
      Picture         =   "previewfm.frx":05CC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Snapshot to clipboard"
      Top             =   2280
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   11220
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton prevASCIIfilbut 
      Caption         =   "S&AVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11040
      Picture         =   "previewfm.frx":0716
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton PreviewMini 
      BackColor       =   &H00C0E0FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      Picture         =   "previewfm.frx":1FA8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5940
      Width           =   375
   End
   Begin VB.CommandButton PreviewMagnify 
      BackColor       =   &H00FFC0FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      Picture         =   "previewfm.frx":20FA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5940
      Width           =   375
   End
   Begin VB.CommandButton PreviewMarginsbut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Margins"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      Picture         =   "previewfm.frx":224C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4500
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   7620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton PreviewPrinterbut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Setup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11040
      Picture         =   "previewfm.frx":3D2E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7815
      LargeChange     =   10
      Left            =   10680
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      TabIndex        =   4
      Top             =   7800
      Width           =   10695
   End
   Begin VB.PictureBox previewpicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   136.79
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   191.823
      TabIndex        =   3
      Top             =   0
      Width           =   10935
      Begin VB.PictureBox previewpicture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawMode        =   7  'Invert
         Height          =   7215
         Left            =   120
         ScaleHeight     =   126.206
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   147.373
         TabIndex        =   6
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.CommandButton PreviewFormatbut 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Paper &Format"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      Picture         =   "previewfm.frx":4398
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton PreviewExitbut 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   11040
      Picture         =   "previewfm.frx":4A02
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6900
      Width           =   855
   End
   Begin VB.CommandButton PreviewOKbut 
      BackColor       =   &H00FFFFFF&
      Caption         =   " &Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      Picture         =   "previewfm.frx":4F34
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "previewfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub zmanbut_Click()
   Zmanimform.Visible = True
   BringWindowToTop (Zmanimform.hwnd)
End Sub

Private Sub schultimbut_Click()
   zmanschul.Visible = True
End Sub

Private Sub prevCLIPbut_Click()
  'send a bitmap image of the current window to the CLIPBOARD
   Screen.MousePointer = vbHourglass
   Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
   waittime = Timer
   Do Until Timer > waittime + 1
      DoEvents
   Loop
   Screen.MousePointer = vbDefault
   response = MsgBox("The BitMap of this calendar has been saved to the Clipboard. " + _
                      "It is available for other window programs by pasting from the Clipboard. " + _
                      "You can can also use MSPaint or an equivalent program to edit or print it.", vbInformation + vbOKOnly, "Cal Program")
End Sub

Private Sub Form_Load()
   'version: 04/08/2003
   rescal = 1
   rescale = False
   magnify = False
   Marginshow = False
   previewfm.ScaleMode = 6
   previewpicture.ScaleMode = 6
   previewpicture2.ScaleMode = 6
   previewpicture2.AutoSize = True
   previewpicture.BorderStyle = 0
   previewpicture2.BorderStyle = 0
   previewpicture2.AutoRedraw = True
   If automatic = True Then
      previewpicture2.Width = 320
      previewpicture2.Height = 320
   Else
     previewpicture2.Width = 600 '320
     previewpicture2.Height = 500 '400 '600 '320
     End If
   previewpicture2.Left = previewpicture.Left
   previewpicture2.Top = previewpicture.Top
   HScroll1.Max = (previewpicture2.Width - previewpicture.Width) * rescal
   VScroll1.Max = (previewpicture2.Height - previewpicture.Height) * rescal
   'VScroll1.Visible = (previewpicture.Height < previewpicture2.Height)
   'HScroll1.Visible = (previewpicture.Width < previewpicture2.Width)

End Sub

Private Sub HScroll1_Change()
   previewpicture2.Left = -HScroll1.Value
End Sub

Private Sub prevASCIIfilbut_Click()
   'Dim ExcelSheet As Object
   Dim ExcelApp As Excel.Application
   Dim ExcelBook As Excel.Workbook
   Dim ExcelSheet As Excel.Worksheet
   Dim Times() As String, Coords() As String
   Dim ITMx As Double, ITMy As Double
   Dim lat As Double, lon As Double, hgt As Double
   
'!!!!!!!!!!!!!!!!!!!!!
'If internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-l2: prevASCIIfilbut: beginning wrting html file"
'Close #lognum%
'End If
 


10 On Error GoTo errhandler
   If internet Or automatic Then
      CommonDialog2.FileName = dirint$ + "\" + Mid(servnam$, 1, 8) + ".html"
      ext$ = "html"
      GoTo 25
      End If
   CommonDialog2.FileName = sEmpty
   CommonDialog2.Filter = "List of times as outputed from netzski3 (*.ls1)|*.ls1|" + _
                          "EXCEL file (matrix) (*.xls)|*.xls|" + _
                          "WIN-ASCII List of Times vs. Month (*.ls2)|*.ls2|" + _
                          "List of dates and times (*.ls2)|*.ls2|" + _
                          "DOS-Win95 List of Times vs. Month (*.dos)|*.dos|" + _
                          "List of dates and times (*.ls3)|*.ls3|" + _
                          "Schul Davening Times and Moladim (*.sch)|*.sch|" + _
                          "Menat WIN-ASCII format (*.men)|*.men|" + _
                          "HTML format (*.html)|*.html|" + _
                          "All files (*.*)|*.*"
   CommonDialog2.FilterIndex = 1
   CommonDialog2.CancelError = True
   CommonDialog2.ShowSave
   'check for existing files, and for wrong save directories
   ext$ = RTrim$(Mid$(CommonDialog2.FileName, InStr(1, CommonDialog2.FileName, ".") + 1, 3))
   If ext$ = "htm" Then ext$ = "html"
   If ext$ <> "ls1" And ext$ <> "xls" And ext$ <> "ls2" And ext$ <> "dos" And ext$ <> "men" And ext$ <> "ls3" And ext$ <> "sch" And ext$ <> "html" And ext$ <> "*" Then
      MsgBox "Sorry, the selected save file format is not yet available. Please choose (*.ls1) format.", vbInformation, "Cal Program"
      GoTo 10
      End If
   If ext$ <> "ls1" And ext$ <> "xls" And ext$ <> "ls2" And ext$ <> "dos" And ext$ <> "ls3" And ext$ <> "sch" And ext$ <> "men" And ext$ <> "html" And ext$ <> "*" Then Exit Sub
   myfile = Dir(CommonDialog2.FileName)
   If myfile <> sEmpty And ext$ <> "xls" Then
      response = MsgBox("Write over existing file?", vbYesNoCancel + vbQuestion, "Cal Program")
      If response = vbNo Then
         GoTo 10
      ElseIf response = vbCancel Then
         Exit Sub
         End If
      End If
   If Len(CommonDialog2.FileName) > 19 Then
      If Mid$(CommonDialog2.FileName, 1, 15) = drivfordtm$ + "netz\" Or Mid$(CommonDialog2.FileName, 1, 15) = drivfordtm$ + "skiy\" Then
        MsgBox "The selected directory is not avaiable for saving files.  Please choose another directory.", vbInformation, "Cal Program"
        GoTo 10
        End If
      End If
25 If ext$ = "ls1" Then
      Screen.MousePointer = vbHourglass
      filtm1% = FreeFile
      Open CommonDialog2.FileName For Output As #filtm1%
50    filtm2% = FreeFile
      If Abs(nsetflag%) = 1 Or Abs(nsetflag%) = 3 Then
         Open drivfordtm$ + "netz\netzskiy.tm2" For Input As #filtm2%
         Input #filtm2%, numplac%
         Print #filtm1%, "List of sunrise times for city file: " + currentdir
         Do Until EOF(filtm2%)
            Input #filtm2%, placnam$
            filtm3% = FreeFile
            Open placnam$ For Input As #filtm3%
            Do Until EOF(filtm3%)
               Line Input #filtm3%, doclin$
               Print #filtm1%, doclin$
            Loop
            Close #filtm3%
         Loop
         Close #filtm2%
         If Abs(nsetflag%) = 1 Then
            Close #filtm1
            Screen.MousePointer = vbDefault
            Exit Sub
         ElseIf Abs(nsetflag%) = 3 Then
            Print #filtm1%, sEmpty
            Print #filtm1%, sEmpty
            nsetflag% = 2
            GoTo 50
            End If
      ElseIf Abs(nsetflag%) = 2 Then
         Open drivfordtm$ + "skiy\netzskiy.tm2" For Input As #filtm2%
         Input #filtm2%, numplac%
         Print #filtm1%, "List of sunset times for city file: " + currentdir
         Do Until EOF(filtm2%)
            Input #filtm2%, placnam$
            filtm3% = FreeFile
            Open placnam$ For Input As #filtm3%
            Do Until EOF(filtm3%)
               Line Input #filtm3%, doclin$
               Print #filtm1%, doclin$
            Loop
            Close #filtm3%
         Loop
         Close #filtm2%
         Close #filtm1%
         Screen.MousePointer = vbDefault
         Exit Sub
         End If
   ElseIf ext$ = "xls" Then
     Screen.MousePointer = vbHourglass
     
     Set ExcelApp = New Excel.Application
     Set ExcelBook = ExcelApp.Workbooks.Add
     Set ExcelSheet = ExcelBook.Worksheets.Add
     
     'Set ExcelSheet = CreateObject("Excel.Sheet")
     Screen.MousePointer = vbDefault
     ExcelBook.Application.Visible = True
     ExcelBook.Windows(1).Visible = True
     'ExcelSheet.Application.Visible = True
     If Abs(nsetflag%) = 2 Then ii% = 1 Else ii% = 0
     nadd% = 0: If hebcal = False Then nadd% = 1
     ExcelSheet.Cells(1, 9).Value = storheader$(ii%, 0)
     ExcelSheet.Cells(2, 9).Value = storheader$(ii%, 1)
     ExcelSheet.Cells(3, 9).Value = storheader$(ii%, 2)
     ExcelSheet.Cells(4, 9).Value = storheader$(ii%, 3)
     ExcelSheet.Cells(38 + nadd%, 9).Value = storheader$(ii%, 4)
     If Abs(nsetflag%) = 3 Then
        ExcelSheet.Cells(42 + nadd%, 9).Value = storheader$(1, 0)
        ExcelSheet.Cells(43 + nadd%, 9).Value = storheader$(1, 1)
        ExcelSheet.Cells(44 + nadd%, 9).Value = storheader$(1, 2)
        ExcelSheet.Cells(45 + nadd%, 9).Value = storheader$(1, 3)
        ExcelSheet.Cells(79 + 2 * nadd%, 9).Value = storheader$(1, 4)
        If nearcolor = True Then
          If nearnez = True Or nearski = True Then 'warning line
             ExcelSheet.Cells(80 + 2 * nadd%, 9).Value = storheader$(1, 5)
             End If
          End If
     Else
        If nearcolor = True Then
          If nearnez = True Or nearski = True Then 'warning line
             ExcelSheet.Cells(39 + nadd%, 9).Value = storheader$(ii%, 5)
             End If
          End If
        End If
     
     If hebcal = True Then
        For i% = 0 To endyr% - 1
           ExcelSheet.Cells(6, 3 + endyr% - i%).Value = stormon$(i%)
           If Abs(nsetflag%) = 3 Then
              ExcelSheet.Cells(47, 3 + endyr% - i%).Value = stormon$(i%)
              End If
           For j% = 0 To 29
             'Call hebnum(j% + 1, cha$)
             If optionheb = True Then
                Call hebnum(j% + 1, cha$)
             Else
                cha$ = LTrim$(RTrim$(j% + 1))
                End If
             If Abs(nsetflag%) = 1 Or Abs(nsetflag%) = 3 Then
                ExcelSheet.Cells(j% + 8, 3 + endyr% - i%).Value = stortim$(0, i%, j%)
                ExcelSheet.Cells(j% + 8, 2).Value = cha$
                ExcelSheet.Cells(j% + 8, 5 + endyr%).Value = cha$
             ElseIf Abs(nsetflag%) = 2 Then
                ExcelSheet.Cells(j% + 8, 3 + endyr% - i%).Value = stortim$(1, i%, j%)
                ExcelSheet.Cells(j% + 8, 2).Value = cha$
                ExcelSheet.Cells(j% + 8, 5 + endyr%).Value = cha$
                End If
             If Abs(nsetflag%) = 3 Then
                ExcelSheet.Cells(j% + 49, 3 + endyr% - i%).Value = stortim$(1, i%, j%)
                ExcelSheet.Cells(j% + 49, 2).Value = cha$
                ExcelSheet.Cells(j% + 49, 5 + endyr%).Value = cha$
                End If
           Next j%
        Next i%
     ElseIf hebcal = False Then
        For i% = 0 To endyr% - 1
           ExcelSheet.Cells(6, 3 + endyr% - i%).Value = stormon$(i%)
           If Abs(nsetflag%) = 3 Then
              ExcelSheet.Cells(47 + nadd%, 3 + endyr% - i%).Value = stormon$(i%)
              End If
           For j% = 0 To 30
             cha$ = Trim$(Str$(j% + 1))
             If Abs(nsetflag%) = 1 Or Abs(nsetflag%) = 3 Then
                ExcelSheet.Cells(j% + 8, 3 + endyr% - i%).Value = stortim$(0, i%, j%)
                ExcelSheet.Cells(j% + 8, 2).Value = cha$
                ExcelSheet.Cells(j% + 8, 5 + endyr%).Value = cha$
             ElseIf Abs(nsetflag%) = 2 Then
                ExcelSheet.Cells(j% + 8, 3 + endyr% - i%).Value = stortim$(1, i%, j%)
                ExcelSheet.Cells(j% + 8, 2).Value = cha$
                ExcelSheet.Cells(j% + 8, 5 + endyr%).Value = cha$
                End If
             If Abs(nsetflag%) = 3 Then
                ExcelSheet.Cells(j% + 49 + nadd%, 3 + endyr% - i%).Value = stortim$(1, i%, j%)
                ExcelSheet.Cells(j% + 49 + nadd%, 2).Value = cha$
                ExcelSheet.Cells(j% + 49 + nadd%, 5 + endyr%).Value = cha$
                End If
           Next j%
        Next i%
        End If
     ExcelSheet.SaveAs CommonDialog2.FileName
     previewfm.SetFocus
     'Screen.MousePointer = vbDefault
     response = MsgBox("Do you wan't to close the EXCEL window? " + _
     "(If you answer No, then EXCEL will continue running, even after " + _
     "closing Cal Program.)", vbQuestion + vbYesNo, "Cal Program")
     If response = vbYes Then
        'ExcelSheet.Application.Quit
        'Set ExcelSheet = Nothing
        ExcelApp.Quit
        
        Set ExcelApp = Nothing
        Set ExcelBook = Nothing
        Set ExcelSheet = Nothing
        End If
     Exit Sub
 ElseIf ext$ = "ls2" Then
     Screen.MousePointer = vbHourglass
     alsosunset = False
     filtm1% = FreeFile
     Open CommonDialog2.FileName For Output As #filtm1%
100  If Abs(nsetflag%) = 2 Or alsosunset = True Then ii% = 1 Else ii% = 0
     nadd% = 0: If hebcal = False Then nadd% = 1
     If (Abs(nsetflag%) = 3 And alsosunset = False) Or Abs(nsetflag%) = 1 Then 'sunrise tables
        Print #filtm1%, storheader$(ii%, 0)
        Print #filtm1%, storheader$(ii%, 1)
        Print #filtm1%, storheader$(ii%, 2)
        Print #filtm1%, storheader$(ii%, 3)
        If hebcal = True Then
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              Print #filtm1%, stormon$(i%)
              Print #filtm1%, String(20, "-")
              For j% = 0 To 29
'                Call hebnum(j% + 1, cha$)
'                If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, cha$ + ":"; Tab(10); stortim$(0, i%, j%)
                If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(0, i%, j%)
              Next j%
           Next i%
        ElseIf hebcal = False Then
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              Print #filtm1%, stormon$(i%)
              Print #filtm1%, String(20, "-")
              For j% = 0 To 30
     '           cha$ = ltrim(rtrim(str$(j% + 1)
     '           If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, cha$ + ":"; Tab(10); stortim$(0, i%, j%)
                If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(0, i%, j%)
              Next j%
           Next i%
           End If
        Print #filtm1%, storheader$(ii%, 4) 'bottom line
        
        If nearcolor = True Then
          If nearnez = True Or nearski = True Then 'warning line
             Print #filtm1%, storheader$(ii%, 5)
             End If
          End If
        End If
        
     If Abs(nsetflag%) = 2 Or alsosunset = True Then 'sunset tables
        Print #filtm1%, storheader$(1, 0)
        Print #filtm1%, storheader$(1, 1)
        Print #filtm1%, storheader$(1, 2)
        Print #filtm1%, storheader$(1, 3)
        If hebcal = True Then
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              Print #filtm1%, stormon$(i%)
              Print #filtm1%, String(20, "-")
              For j% = 0 To 29
'                Call hebnum(j% + 1, cha$)
'                If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, cha$ + ":"; Tab(10); stortim$(1, i%, j%)
                If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(1, i%, j%)
              Next j%
           Next i%
        ElseIf hebcal = False Then
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              Print #filtm1%, stormon$(i%)
              Print #filtm1%, String(20, "-")
              For j% = 0 To 30
'                cha$ = ltrim(rtrim(str$(j% + 1)
'                If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, cha$ + ":"; Tab(10); stortim$(1, i%, j%)
                If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(1, i%, j%)
              Next j%
           Next i%
           End If
        Print #filtm1%, storheader$(1, 4) 'bottom line
        
        If nearcolor = True Then
          If nearnez = True Or nearski = True Then 'warning line
             Print #filtm1%, storheader$(ii%, 5)
             End If
          End If
        End If
        
     If Abs(nsetflag%) = 3 And alsosunset = False Then 'go back and print sunset tables
        alsosunset = True
        Print #filtm1%, sEmpty
        Print #filtm1%, sEmpty
        Print #filtm1%, sEmpty
        Print #filtm1%, String(100, "-")
        GoTo 100
        End If
     alsosunset = False
     Close #filtm1%
     Screen.MousePointer = vbDefault
 ElseIf ext$ = "ls3" Then
      Screen.MousePointer = vbHourglass
      filtm1% = FreeFile
      Open CommonDialog2.FileName For Output As #filtm1%
125   filtm2% = FreeFile
      If Abs(nsetflag%) = 1 Or Abs(nsetflag%) = 3 Then
         Open drivfordtm$ + "netz\netzskiy.tm2" For Input As #filtm2%
         Input #filtm2%, numplac%
         
         If numplac% = 1 Then
         
            'find which file was selected
            For i% = 6 To nn4%
               If nchecked%(i% - 5) = 1 Then
                  Exit For
                  End If
            Next i%
         
            Print #filtm1%, "Sunrise times for: " + netzski$(0, i%) 'currentdir + "; " + CommonDialog2.FileName
            Coords = Split(netzski$(1, i%), ",")
            ITMx = Coords(0)
            ITMy = Coords(1)
            hgt = Coords(2)
            'convert to geographic latitude and longitude if necessary
            If geo And eroscountry <> "Israel" Then 'geo coordinates
               lat = ITMy
               lon = ITMx
               Print #filtm1%, "longitude, latitude, height (m)"
               Write #filtm1%, lon, lat, hgt
            Else 'EY old ITM coordinates
               Call casgeo(ITMx, ITMy, lon, lat)
               Print #filtm1%, "ITMx, ITMy, longitude, latitude, height (m)"
               Write #filtm1%, ITMx, ITMy, lon, lat, hgt
               End If
'            Print #filtm1%, netzski$(1, 6) 'coordinates,hgt,year, etc
         Else
            Print #filtm1%, "Sunrise times for: " + currentdir + "; " + CommonDialog2.FileName
            End If
            
         Do Until EOF(filtm2%)
            Input #filtm2%, placnam$
            filtm3% = FreeFile
            Open placnam$ For Input As #filtm3%
            Do Until EOF(filtm3%)
               Line Input #filtm3%, doclin$
               doclin$ = Replace(doclin$, "   ", ",")
               Times = Split(doclin$, ",")
               Print #filtm1%, Times(0) & "," & Times(1)
            Loop
            Close #filtm3%
         Loop
         Close #filtm2%
         If Abs(nsetflag%) = 1 Then
            Close #filtm1
            Screen.MousePointer = vbDefault
            Exit Sub
         ElseIf Abs(nsetflag%) = 3 Then
            Print #filtm1%, sEmpty
            Print #filtm1%, sEmpty
            nsetflag% = 2
            GoTo 125
            End If
      ElseIf Abs(nsetflag%) = 2 Then
         Open drivfordtm$ + "skiy\netzskiy.tm2" For Input As #filtm2%
         Input #filtm2%, numplac%
         
         If numplac% = 1 Then
         
            'find which file was selected
            For i% = 6 To nn4%
               If nchecked%(i% - 5) = 1 Then
                  Exit For
                  End If
            Next i%
         
            Print #filtm1%, "Sunset times for: " + netzski$(0, i%) 'currentdir + "; " + CommonDialog2.FileName
            Coords = Split(netzski$(1, i%), ",")
            ITMx = Coords(0)
            ITMy = Coords(1)
            hgt = Coords(2)
            'convert to geographic latitude and longitude if necessary
            If geo And eroscountry <> "Israel" Then 'geo coordinates
               lat = ITMy
               lon = ITMx
               Print #filtm1%, "longitude, latitude, height (m)"
               Write #filtm1%, lon, lat, hgt
            Else 'EY old ITM coordinates
               Call casgeo(ITMx, ITMy, lon, lat)
               Print #filtm1%, "ITMx, ITMy, longitude, latitude, height (m)"
               Write #filtm1%, ITMx, ITMy, lon, lat, hgt
               End If
'            Print #filtm1%, netzski$(1, 6) 'coordinates,hgt,year, etc
         Else
            Print #filtm1%, "Sunset times for: " + currentdir + "; " + CommonDialog2.FileName
            End If

         Do Until EOF(filtm2%)
            Input #filtm2%, placnam$
            filtm3% = FreeFile
            Open placnam$ For Input As #filtm3%
            Do Until EOF(filtm3%)
               Line Input #filtm3%, doclin$
               doclin$ = Replace(doclin$, "   ", ",")
               Times = Split(doclin$, ",")
               Print #filtm1%, Times(0) & "," & Times(1)
            Loop
            Close #filtm3%
         Loop
         Close #filtm2%
         Close #filtm1%
         Screen.MousePointer = vbDefault
         Exit Sub
         End If
 ElseIf ext$ = "sch" Then
     zmanschul.Visible = True
 ElseIf ext$ = "dos" Or ext$ = "men" Then
     If ext$ = "men" Then
        response = MsgBox("Have the times been rounded to the nearest 6 seconds?", vbQuestion + vbYesNoCancel, "Cal Program")
        If response <> vbYes Then
           Exit Sub
           End If
        End If
     Screen.MousePointer = vbHourglass
     alsosunset = False
     filtm1% = FreeFile
     Open CommonDialog2.FileName For Output As #filtm1%
150  If Abs(nsetflag%) = 2 Or alsosunset = True Then ii% = 1 Else ii% = 0
     nadd% = 0: If hebcal = False Then nadd% = 1
     If (Abs(nsetflag%) = 3 And alsosunset = False) Or Abs(nsetflag%) = 1 Then 'sunrise tables
        'convert headers into ascii
        hh$ = storheader$(ii%, 0)
        GoSub asciiwin
        Print #filtm1%, hhout$
        hh$ = storheader$(ii%, 1)
        GoSub asciiwin
        Print #filtm1%, hhout$
        hh$ = storheader$(ii%, 2)
        GoSub asciiwin
        Print #filtm1%, hhout$
        hh$ = storheader$(ii%, 3)
        GoSub asciiwin
        Print #filtm1%, hhout$
        If Option1b = True Then 'hebcal = True
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              hh$ = stormon$(i%)
              GoSub asciiwin
              Print #filtm1%, hhout$
              Print #filtm1%, String(20, "-")
              If ext$ <> "men" Then
                 For j% = 0 To 29
'                   Call hebnum(j% + 1, cha$)
'                   hh$ = cha$
'                   GoSub asciiwin
'                   If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, hhout$ + ":"; Tab(10); stortim$(0, i%, j%)
                   If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(0, i%, j%)
                 Next j%
              ElseIf ext$ = "men" Then
                 For j% = 0 To 29
                   If stortim$(0, i%, j%) <> sEmpty Then
                      stortim$(0, i%, j%) = Mid$(stortim$(0, i%, j%), 1, 4) + LTrim$(Format(Val(Mid$(stortim$(0, i%, j%), 6, 7)) / 60, ".0"))
                      Print #filtm1%, stortim$(0, i%, j%)
                      End If
                 Next j%
                 End If
           Next i%
        ElseIf Option2b = True Then 'hebcal = False Then
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              hh$ = stormon$(i%)
              GoSub asciiwin
              Print #filtm1%, hhout$
              Print #filtm1%, String(20, "-")
              If ext$ <> "men" Then
                 For j% = 0 To 30
'                   cha$ = ltrim(rtrim(str$(j% + 1)
'                   If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, cha$ + ":"; Tab(10); stortim$(0, i%, j%)
                   If stortim$(0, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(0, i%, j%)
                 Next j%
              ElseIf ext$ = "men" Then
                 For j% = 0 To 30
                   If stortim$(0, i%, j%) <> sEmpty Then
                      stortim$(0, i%, j%) = Mid$(stortim$(0, i%, j%), 1, 4) + LTrim$(Format(Val(Mid$(stortim$(0, i%, j%), 6, 7)) / 60, ".0"))
                      Print #filtm1%, stortim$(0, i%, j%)
                      End If
                 Next j%
                 End If
           Next i%
           End If
        hh$ = storheader$(ii%, 4)
        GoSub asciiwin
        Print #filtm1%, hhout$ 'bottom line
        
        If nearcolor = True Then
          If nearnez = True Or nearski = True Then 'warning line
             hh$ = storheader$(ii%, 5)
             GoSub asciiwin
             Print #filtm1%, hhout$ 'warning line
             End If
          End If
        
        End If
        
     If Abs(nsetflag%) = 2 Or alsosunset = True Then 'sunset tables
        hh$ = storheader$(1, 0)
        GoSub asciiwin
        Print #filtm1%, hhout$
        hh$ = storheader$(1, 1)
        GoSub asciiwin
        Print #filtm1%, hhout$
        hh$ = storheader$(1, 2)
        GoSub asciiwin
        Print #filtm1%, hhout$
        hh$ = storheader$(1, 3)
        GoSub asciiwin
        Print #filtm1%, hhout$
        If Option1b = True Then  'hebcal=true
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              hh$ = stormon$(i%)
              GoSub asciiwin
              Print #filtm1%, hhout$
              Print #filtm1%, String(20, "-")
              If ext$ <> "men" Then
                 For j% = 0 To 29
'                   Call hebnum(j% + 1, cha$)
'                   hh$ = cha$
'                   GoSub asciiwin
'                   If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, hhout$ + ":"; Tab(10); stortim$(1, i%, j%)
                   If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(1, i%, j%)
                 Next j%
              ElseIf ext$ = "men" Then
                 For j% = 0 To 29
                   If stortim$(1, i%, j%) <> sEmpty Then
                      stortim$(1, i%, j%) = Mid$(stortim$(1, i%, j%), 1, 4) + LTrim$(Format(Val(Mid$(stortim$(1, i%, j%), 6, 7)) / 60, ".0"))
                      Print #filtm1%, stortim$(1, i%, j%)
                      End If
                 Next j%
                 End If
           Next i%
        ElseIf Option2b = True Then 'hebcal=false
           For i% = 0 To endyr% - 1
              Print #filtm1%, sEmpty
              hh$ = stormon$(i%)
              GoSub asciiwin
              Print #filtm1%, hhout$
              Print #filtm1%, String(20, "-")
              If ext$ <> "men" Then
                 For j% = 0 To 30
'                   cha$ = ltrim(rtrim(str$(j% + 1)
'                   If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, cha$ + ":"; Tab(10); stortim$(1, i%, j%)
                   If stortim$(1, i%, j%) <> sEmpty Then Print #filtm1%, stortim$(1, i%, j%)
                 Next j%
              ElseIf ext$ = "men" Then
                 For j% = 0 To 30
                   If stortim$(1, i%, j%) <> sEmpty Then
                      stortim$(1, i%, j%) = Mid$(stortim$(1, i%, j%), 1, 4) + LTrim$(Format(Val(Mid$(stortim$(1, i%, j%), 6, 7)) / 60, ".0"))
                      Print #filtm1%, stortim$(1, i%, j%)
                      End If
                 Next j%
                 End If
           Next i%
           End If
        hh$ = storheader$(1, 4)
        GoSub asciiwin
        Print #filtm1%, hhout$ 'bottom line

     End If
     If Abs(nsetflag%) = 3 And alsosunset = False Then 'go back and print sunset tables
        alsosunset = True
        Print #filtm1%, sEmpty
        Print #filtm1%, sEmpty
        Print #filtm1%, sEmpty
        Print #filtm1%, String(100, "-")
        GoTo 150
        End If
     alsosunset = False
     Close #filtm1%
     Screen.MousePointer = vbDefault
     Exit Sub
     
asciiwin:
   If ext$ = "men" Then
      '---------------------------------------------
      'old Menat DOS ASCII (Win 95-98) no longer supported
      '---------------------------------------------
      hhout$ = hh$
      Return
      End If
   hhout$ = sEmpty
   iiitmp% = 0
   For iii% = 1 To Len(hh$)
      cc$ = Mid$(hh$, iii%, 1)
      If cc$ = " " Then
         If iiitmp% <> 0 Then 'first add this non-inverted non Hebrew phrase
            hhout$ = Mid$(hh$, iii% - iiitmp%, iiitmp%) + hhout$
            iiitmp% = 0
            End If
         hhout$ = " " + hhout$
      ElseIf Asc(cc$) >= 128 + 96 And Asc(cc$) <= 154 + 96 Then 'Window Hebrew characters, so convert and invert
         If iiitmp% <> 0 Then 'first add this non-inverted non Hebrew phrase
            hhout$ = Mid$(hh$, iii% - iiitmp%, iiitmp%) + hhout$
            iiitmp% = 0
            End If
         cc2$ = Chr$(Asc(cc$) - 96)
         hhout$ = cc2$ + hhout$
      Else 'don't invert
         If iiitmp% <> 0 Then
            iiitmp% = iiitmp% + 1
         Else
            iiitmp% = 1
            End If
         'hhout$ = cc$ + hhout$
         End If
   Next iii%
  If iiitmp% <> 0 Then 'add this phrase
     hhout$ = Mid$(hh$, iii% - iiitmp%, iiitmp%) + hhout$
     iiitmp% = 0
     End If
Return

 ElseIf ext$ = "sch" Then
 ElseIf ext$ = "html" Then
     Close
     If endyr% = 12 Then
        If automatic And autosave Then
           If hebcal Then
              nsetflag% = 1
              Call heb12monthHTML(htmldir$ & "\" & cityAutoEng$ & ".html")
              nsetflag% = 2
              Call heb12monthHTML(htmldir$ & "\" & cityAutoEng$ & ".html")
           Else
              If Caldirectories.chkHtmlAuto.Value = vbChecked Then
                 nsetflag% = 1
                 Call civilHTM(htmldir$ & "\" & cityAutoEng$ & ".html")
                 nsetflag% = 2
                 Call civilHTM(htmldir$ & "\" & cityAutoEng$ & ".html")
              ElseIf Caldirectories.chkListAuto.Value = vbChecked Then
                 nsetflag% = 1
                 Call civilListHTM(htmldir$ & "\" & cityAutoEng$ & "_netz_" & ".html")
                 nsetflag% = 2
                 Call civilListHTM(htmldir$ & "\" & cityAutoEng$ & "_skiy_" & ".html")
                 End If
              End If
        Else
           If (Abs(nsetflag% < 3)) Then
               If hebcal Then
                  Call heb12monthHTML(CommonDialog2.FileName)
               Else
                  Call civilHTM(CommonDialog2.FileName)
                  End If
           Else
               nsetflag0% = nsetflag%
               nsetflag% = 1
               If hebcal Then
                  Call heb12monthHTML(CommonDialog2.FileName)
               Else
                  Call civilHTM(CommonDialog2.FileName)
                  End If
               savehtml = True
               nsetflag% = 2
               If hebcal Then
                  Call heb12monthHTML(CommonDialog2.FileName)
               Else
                  Call civilHTM(CommonDialog2.FileName)
                  End If
               nsetflag% = nsetflag0%
               savehtml = False
               End If
           End If
     ElseIf endyr% = 13 Then

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-l3: prevASCIIfilbut: before heb13monthHTML"
'Close #lognum%
'End If
     
        If automatic And autosave Then
           nsetflag% = 1
           Call heb13monthHTML(htmldir$ & "\" & cityAutoEng$ & ".html")
           nsetflag% = 2
           Call heb13monthHTML(htmldir$ & "\" & cityAutoEng$ & ".html")
        Else
           If (Abs(nsetflag% < 3)) Then
               Call heb13monthHTML(CommonDialog2.FileName)
           Else
               nsetflag0% = nsetflag%
               nsetflag% = 1
               Call heb13monthHTML(CommonDialog2.FileName)
               savehtml = True
               nsetflag% = 2
               Call heb13monthHTML(CommonDialog2.FileName)
               nsetflag% = nsetflag0%
               savehtml = False
               End If
           
           End If
        
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If internet = True Then
'lognum% = FreeFile
'Open drivjk$ + "calprog.log" For Append As #lognum%
'Print #lognum%, "Step #11-l4: prevASCIIfilbut: after heb13monthHTML"
'Close #lognum%
'End If
        
        
        End If
        
        If automatic And autosave Then 'write to TOC file
           fnum% = FreeFile
           
           myfile = Dir(drivjk$ & "html_city_tables\" & "index.html")
           If myfile = sEmpty Then
              
              Open drivjk$ & "html_city_tables\" & "index.html" For Output As #fnum%
              
              Print #fnum%, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
              Print #fnum%, "<HTML dir = " & Chr$(34) & "rtl" & Chr$(34) & ">"
              Print #fnum%, "<HEAD>"
              Print #fnum%, "    <TITLE>Chai Tables of Eretz Yisroel Table of Contents</TITLE>"
              Print #fnum%, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1255""/>"
              Print #fnum%, "    <META NAME=""AUTHOR"" CONTENT=""Chaim Keller, Chai Tables, www.chaitables.com""/>"
              Print #fnum%, "</HEAD>"
              Print #fnum%, "<BODY>"
               
              Call hebyear(yrheb%, yearheb$)

              Print #fnum%, "<h3>" & "לוחות" & Chr$(34) & "חי" & Chr$(34) & " " & "לארץ ישראל הידוע בשם" & " " & "לוחות" & " " & Chr$(34) & TitleLine$ & Chr$(34) & " " & "לשנת" & " " & yearheb$ & "</h3>"
              Print #fnum%, "<a href=" & Chr$(34) & cityAutoEng$ & ".html"; Chr$(34) & ">" & cityAutoHeb$ & "</a><br/>"
               
           Else
              
              Open drivjk$ & "html_city_tables\" & "index.html" For Append As #fnum%
              Print #fnum%, "<a href=" & Chr$(34) & cityAutoEng$ & ".html" & Chr$(34) & ">" & cityAutoHeb$ & "</a><br/>"
              
              End If
              
              Close #fnum%
       
           End If
           
     End If '**************************
     
 If internet = True Then 'check if zmanim tables need to be written
                         'if not, then just close up the windows
    waitime = Timer
    Do Until Timer > waitime + 0.5 '<--!! 0.1 suffices for fast computer
       DoEvents
    Loop
        
    lognum% = FreeFile
    Open drivjk$ + "calprog.log" For Append As #lognum%
    Print #lognum%, "Step #11b: Previewfm closed, sunrise/sunset table written successfully"
    Close #lognum%
    
    'now calculate z'manim if flagged
    If zmanyes% = 1 Then
        
        lognum% = FreeFile
        Open drivjk$ + "calprog.log" For Append As #lognum%
        Print #lognum%, "Step #11c: Calculate zemanim--activate zmanimform"
        Close #lognum%
    
       previewfm.zmanbut.Value = 1
       waitime = Timer 'wait a bit for the window to load up
       Do Until Timer > waitime + 0.5
         DoEvents
       Loop
       'now load in the template
       
       lognum% = FreeFile
       Open drivjk$ + "calprog.log" For Append As #lognum%
       Print #lognum%, "Step #11d: Loading; Zemanim template"
       Close #lognum%
       
       Zmanimform.loadbut.Value = 1
       waitime = Timer 'wait a bit for the window to load up
       Do Until Timer > waitime + 0.5
         DoEvents
       Loop
       'now calculate the z'manim
       
       lognum% = FreeFile
       Open drivjk$ + "calprog.log" For Append As #lognum%
       Print #lognum%, "Step #11e: Template loaded, begin zemanim calculations"
       Close #lognum%
       
       Zmanimform.calendarbut.Value = 1
       'now wait until z'manim are generated
       lwin = FindWindow(vbNullString, Zmanimlistfm.hwnd)
       Do Until lwin = 0
          DoEvents
          lwin = FindWindow(vbNullString, Zmanimlistfm.hwnd)
       Loop
       waitime = Timer 'wait a bit more
       Do Until Timer > waitime + 0.5
         DoEvents
       Loop
       
       lognum% = FreeFile
       Open drivjk$ + "calprog.log" For Append As #lognum%
       Print #lognum%, "Step #11f: Zemanim were calculated successfully, now generate sorted zemanim table"
       Close #lognum%
       
       'now generate sorted table, and save it as a html document
       'with the correct name (server.tim)
       Zmanimlistfm.zmanbut.Value = 1
       End If
       
    
       lognum% = FreeFile
       Open drivjk$ + "calprog.log" For Append As #lognum%
       Print #lognum%, "Step #11g: Sorted zemanim table calculated successfully"
       Close #lognum%
    
    previewfm.PreviewExitbut.Value = True
    newhebcalfm.newhebExitbut.Value = True
    waitime = Timer
    Do Until Timer > waitime + 0.5 '<--!! 0.1 suffices for fast computer
       DoEvents
    Loop
    Caldirectories.ExitButton.Value = True
    End If
    
    lognum% = FreeFile
    Open drivjk$ + "calprog.log" For Append As #lognum%
    Print #lognum%, "Step #11g: Exited from previewfm and newhebcalfm successfully"
    Close #lognum%
 
 Exit Sub
 
errhandler:
   Screen.MousePointer = vbDefault
   If internet = False Then
      Close
      If Err.Number = 32755 Then Exit Sub 'canceled save
      response = MsgBox("Sorry, error " & Err.Number & " has occured!", vbExclamation + vbOKOnly, "Cal Program")
      Exit Sub
   Else
      If internet = True Then
         'exit program with error message
         Close
         myfile = Dir(drivfordtm$ + "busy.cal")
         If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
           
         lognum% = FreeFile
         Open drivjk$ + "calprog.log" For Append As #lognum%
         Print #lognum%, "Fatal error in previewfm, error number: " & Err.Number
         Print #lognum%, Err.Description
         Close #lognum%
         errlog% = FreeFile
         Open drivjk$ + "Cal_prevASCIIfilbut.log" For Output As errlog%
         Print #errlog%, "Cal Prog exited from Previewfm Save button with runtime error message " + Str(Err.Number)
         Print #errlog%, "System Date and Time: " & Str$(Date) & " " & Str$(Time)
         Close #errlog%
         Close
           
      
       'unload forms
        For i% = 0 To Forms.Count - 1
          Unload Forms(i%)
        Next i%
      
        'kill the timer
        If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)
      
        'bring program to abrupt end
        End
        
        End If
   
      End If
End Sub


Private Sub PreviewExitbut_Click()
   previewfm.Visible = False
   newhebcalfm.Visible = True
   'ret = SetWindowPos(newhebcalfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   If automatic = True Then
      newhebcalfm.newhebExitbut.Value = True
      End If
   'Unload previewfm
   'Set previewfm = Nothing
End Sub

Private Sub PreviewMagnify_Click()
   Pageformatfm.Visible = False
   X22 = previewpicture2.ScaleWidth
   y22 = previewpicture2.ScaleHeight
   X11 = previewpicture2.ScaleLeft
   y11 = previewpicture2.ScaleTop
   previewpicture2.Scale (X11, y11)-(X22 * 0.85, y22 * 0.85)
'   previewpicture2.Scale (X11, y11)-(X22 * 0.5, y22 * 0.5)
   previewpicture2.Refresh
   rescale = True
   rescal = (1 / 0.85) * rescal
   'rescal = 2 * rescal
   newhebcalfm.newhebPreviewbut.Value = True
   magnify = True
   'HScroll1.Max = (previewpicture2.Width - previewpicture.Width) * rescal
   'VScroll1.Max = (previewpicture2.Height - previewpicture.Height) * rescal
   'previewpicture2.Width = 320 * rescal
   'previewpicture2.Height = 320 * rescal
End Sub

Private Sub PreviewMarginsbut_Click()
   col = QBColor(0)
   If Marginshow = False Then
      'draw page sizes and margins
      previewfm.previewpicture2.DrawMode = 2
      previewfm.previewpicture2.DrawStyle = vbDot
      If portrait = True Then
         previewfm.previewpicture2.Line (leftmargin, 0)-(leftmargin, paperheight), col
         previewfm.previewpicture2.Line (paperwidth - rightmargin, 0)-(paperwidth - rightmargin, paperheight), col
         previewfm.previewpicture2.Line (0, topmargin)-(paperwidth, topmargin), col
         previewfm.previewpicture2.Line (0, paperheight - bottommargin)-(paperwidth, paperheight - bottommargin), col
      Else
         previewfm.previewpicture2.Line (leftmargin, 0)-(leftmargin, paperwidth), col
         previewfm.previewpicture2.Line (paperheight - rightmargin, 0)-(paperheight - rightmargin, paperwidth), col
         previewfm.previewpicture2.Line (0, topmargin)-(paperheight, topmargin), col
         previewfm.previewpicture2.Line (0, paperwidth - bottommargin)-(paperheight, paperwidth - bottommargin), col
         End If
      Marginshow = True
      previewfm.previewpicture2.DrawMode = 7
      previewfm.previewpicture2.DrawStyle = vbSolid
      previewfm.previewpicture2.Line (0, paperheight)-(paperwidth, paperheight), QBColor(15)
   ElseIf Marginshow = True Then
      previewfm.previewpicture2.DrawMode = 2
      previewfm.previewpicture2.DrawStyle = vbDot 'vbSolid
      If portrait = True Then
         previewfm.previewpicture2.Line (leftmargin, 0)-(leftmargin, paperheight), col
         previewfm.previewpicture2.Line (paperwidth - rightmargin, 0)-(paperwidth - rightmargin, paperheight), col
         previewfm.previewpicture2.Line (0, topmargin)-(paperwidth, topmargin), col
         previewfm.previewpicture2.Line (0, paperheight - bottommargin)-(paperwidth, paperheight - bottommargin), col
      Else
         previewfm.previewpicture2.Line (leftmargin, 0)-(leftmargin, paperwidth), col
         previewfm.previewpicture2.Line (paperheight - rightmargin, 0)-(paperheight - rightmargin, paperwidth), col
         previewfm.previewpicture2.Line (0, topmargin)-(paperheight, topmargin), col
         previewfm.previewpicture2.Line (0, paperwidth - bottommargin)-(paperheight, paperwidth - bottommargin), col
         End If
      Marginshow = False
      previewfm.previewpicture2.DrawMode = 7
      previewfm.previewpicture2.DrawStyle = vbSolid
      previewfm.previewpicture2.Line (0, paperheight)-(paperwidth, paperheight), QBColor(15)
      End If
End Sub

Private Sub PreviewMini_Click()
   Pageformatfm.Visible = False
   X22 = previewpicture2.ScaleWidth
   y22 = previewpicture2.ScaleHeight
   X11 = previewpicture2.ScaleLeft
   y11 = previewpicture2.ScaleTop
   previewpicture2.Scale (X11, y11)-(X22 * 1 / 0.85, y22 * 1 / 0.85)
'   previewpicture2.Scale (X11, y11)-(X22 * 2, y22 * 2)
   previewpicture2.Refresh
   rescale = True
   rescal = 0.85 * rescal
   'rescal = 0.5 * rescal
   newhebcalfm.newhebPreviewbut.Value = True
   magnify = True
   'HScroll1.Max = (previewpicture2.Width - previewpicture.Width) * rescal
   'VScroll1.Max = (previewpicture2.Height - previewpicture.Height) * rescal
   'previewpicture2.Width = 320 * rescal
   'previewpicture2.Height = 320 * rescal
End Sub
Private Sub PreviewOKbut_Click()
'sends table to printer

If PDFprinter Then
   'create pdf of table using bullzip module
   'source: https://community.spiceworks.com/topic/533230-programming-bullzip-pdf-writer-with-vb6
   'source: http://www.biopdf.com/guide/examples/vb6/
    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.SetValue "Output", drivjk$ & "\pdf_city_tables\" & Trim$(Caldirectories.Combo1.Text) & "\" & numautolst% & ".pdf"
    settings.SetValue "ConfirmOverwrite", "no"
    settings.SetValue "ShowSaveAS", "never"
    settings.SetValue "ShowSettings", "never"
    settings.SetValue "ShowPDF", "yes"
    settings.SetValue "RememberLastFileName", "no"
    settings.SetValue "RememberLastFolderName", "no"
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
    End If

Call PrinttoDev(Printer, True)

End Sub


Private Sub PreviewFormatbut_Click()
   If Marginshow = True Then
      previewfm.previewpicture2.DrawMode = 2
      previewfm.previewpicture2.DrawStyle = vbDot
      If portrait = True Then
         previewfm.previewpicture2.Line (leftmargin, 0)-(leftmargin, paperheight), col
         previewfm.previewpicture2.Line (paperwidth - rightmargin, 0)-(paperwidth - rightmargin, paperheight), col
         previewfm.previewpicture2.Line (0, topmargin)-(paperwidth, topmargin), col
         previewfm.previewpicture2.Line (0, paperheight - bottommargin)-(paperwidth, paperheight - bottommargin), col
      Else
         previewfm.previewpicture2.Line (leftmargin, 0)-(leftmargin, paperwidth), col
         previewfm.previewpicture2.Line (paperheight - rightmargin, 0)-(paperheight - rightmargin, paperwidth), col
         previewfm.previewpicture2.Line (0, topmargin)-(paperwidth, topmargin), col
         previewfm.previewpicture2.Line (0, paperwidth - bottommargin)-(paperheight, paperwidth - bottommargin), col
         End If
      Marginshow = False
      End If
      
   Pageformatfm.Combo1.ListIndex = prespap% - 1
   Pageformatfm.Visible = True
   Pageformatfm.Text1 = paperwidth
   Pageformatfm.Text2 = paperheight
   'ret = SetWindowPos(Pageformatfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub

Private Sub PreviewPrinterbut_Click()
   
   CommonDialog1.CancelError = True
   On Error GoTo errhandler
    
    ' Display the Print Setup dialog box.
    CommonDialog1.Flags = cdlPDPrintSetup
    CommonDialog1.ShowPrinter
    
    Dim X As Printer
    For Each X In Printers
        If X.DeviceName = Printer.DeviceName Then
           'this is default printer
           Set Printer = X
           Exit For
           End If
    Next
   
errhandler:
End Sub

Private Sub VScroll1_Change()
   previewpicture2.Top = -VScroll1.Value
End Sub
Private Sub heb12monthHTML(fname)
     Screen.MousePointer = vbHourglass
     alsosunset = False
     If Abs(nsetflag%) = 1 Then
        ii% = 0
     Else
        ii% = 1
        End If
    iheb% = 0: If optionheb = False Then iheb% = 1
    Close
    
    fnum% = FreeFile
    If Not automatic And Not autosave And Not savehtml Then
       Open fname For Output As #fnum%
    ElseIf automatic And autosave And ii% = 0 Then
       Open fname For Output As #fnum%
    ElseIf ((automatic And autosave) Or savehtml) And ii% = 1 Then
       Open fname For Append As #fnum%
       GoTo h100
       End If
    
    Print #fnum%, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
    Print #fnum%, "<HTML>"
    Print #fnum%, "<HEAD>"
    Print #fnum%, "    <TITLE>Your Chai Table</TITLE>"
    Print #fnum%, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1255""/>"
    Print #fnum%, "    <META NAME=""AUTHOR"" CONTENT=""Chaim Keller, Chai Tables, www.chaitables.com""/>"
    Print #fnum%, "</HEAD>"
    Print #fnum%, "<BODY>"
    
h100:
    Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""  ALIGN=""CENTER"" >"
    Print #fnum%, "    <COL WIDTH=""670"">"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""4"" STYLE=""font-size: 16pt"">" & storheader$(ii%, 0) & "</FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "    </TR>"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 1) & "</FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "    </TR>"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 3) & "</FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "    </TR>"
    Print #fnum%, "</TABLE>"
    Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""1"" CELLPADDING=""4"" CELLSPACING=""0""  ALIGN=""CENTER"" >"
    Print #fnum%, "    <COL WIDTH=""16"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""46"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""40"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""40"">"
    Print #fnum%, "    <COL WIDTH=""40"">"
    Print #fnum%, "    <COL WIDTH=""41"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""15"">"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""16"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""LEFT""><BR/>"
    Print #fnum%, "            </P>"
    Print #fnum%, "        </TD>"
    If optionheb = True Then
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 13) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""46"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 12) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 11) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 10) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 9) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 8) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 14) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 5) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 4) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 3) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""41"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 2) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 1) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""15"" VALIGN=""BOTTOM"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><BR/>"
        Print #fnum%, "            </P>"
        Print #fnum%, "        </TD>"
    Else
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 1) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""46"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 2) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 3) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 4) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 5) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 14) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 8) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 9) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 10) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 11) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""41"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 12) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 13) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""15"" VALIGN=""BOTTOM"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><BR/>"
        Print #fnum%, "            </P>"
        Print #fnum%, "        </TD>"
       End If
    Print #fnum%, "    </TR>"

    
    iii% = -1
    For i% = 1 To 5 '5 rows of groups of 3 coloured, and 3 uncoloured
        'check for shabboses (_) and for near mountains (*)
        For j% = 1 To 3
           iii% = iii% + 1
           'Call hebnum(iii% + 1, cha$)
           If optionheb = True Then
              Call hebnum(iii% + 1, cha$)
           Else
              cha$ = Trim$(Str(iii% + 1))
              End If
           
            Print #fnum%, "    <TR VALIGN=""BOTTOM"">"
            Print #fnum%, "        <TD WIDTH=""16"" BGCOLOR=""#ffff00"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            'now determine if this time is underlined or shabbos and add appropriate HTML ******
            k% = 0
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""46"" BGCOLOR=""#ffff00"">"
            k% = 1
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 2
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 3
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 4
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 5
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 6
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 7
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 8
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 9
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""41"" BGCOLOR=""#ffff00"">"
            k% = 10
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 11
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""16"" BGCOLOR=""#ffff00"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
        Next j%
        For j% = 1 To 3
           iii% = iii% + 1
           'Call hebnum(iii% + 1, cha$)
           If optionheb = True Then
              Call hebnum(iii% + 1, cha$)
           Else
              cha$ = Trim$(Str(iii% + 1))
              End If
            Print #fnum%, "    <TR VALIGN=""BOTTOM"">"
            Print #fnum%, "        <TD WIDTH=""16"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "        <TD WIDTH=""39"" >"
            'now determine if this time is underlined or shabbos and add appropriate HTML ******
            k% = 0
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""46"">"
            k% = 1
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 2
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 3
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 4
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 5
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 6
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 7
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 8
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 9
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""41"">"
            k% = 10
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 11
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""16"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
                Next j%
            Next i%
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""  ALIGN=""CENTER"" >"
            Print #fnum%, "    <COL WIDTH=""670"">"
            Print #fnum%, "    <TR>"
            Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 4) & "</FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" ALIGN=""CENTER"" >"
            Print #fnum%, "    <COL WIDTH=""670"">"
            Print #fnum%, "    <TR>"
            Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 5) & "</FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<P><BR/><BR/>"
            Print #fnum%, "</P>"
            
            If automatic And autosave And ii% = 0 Then GoTo h500
            
            Print #fnum%, "</BODY>"
            Print #fnum%, "</HTML>"
h500:
        Close #fnum%
        Screen.MousePointer = vbDefault
        Exit Sub
        
timestring:
        If optionheb = True Then
           ntim$ = stortim$(ii%, endyr% - k% - 1, iii%)
        Else
           ntim$ = stortim$(ii%, k%, iii%)
           End If
        positunder% = InStr(ntim$, "_")
        If positunder% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positunder% - 1)
           End If
        positnear% = InStr(ntim$, "*")
        If positnear% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positnear% - 1)
           End If
        If positunder% <> 0 Then
           ntim$ = "<u>" & ntim$ & "</u>"
           End If
        If positnear% <> 0 Then
           ntim$ = "<font color=""#33cc66"">" & ntim$ & "</font>"
           End If
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & ntim$ & "</FONT></P>"
        Print #fnum%, "        </TD>"
Return
    
End Sub

Sub heb13monthHTML(fname)
     Screen.MousePointer = vbHourglass
     alsosunset = False
     If Abs(nsetflag%) = 1 Then
        ii% = 0
     Else
        ii% = 1
        End If

    iheb% = 0: If optionheb = False Then iheb% = 1
    
    fnum% = FreeFile
    If Not automatic And Not autosave And Not savehtml Then
       Open fname For Output As #fnum%
    ElseIf automatic And autosave And ii% = 0 Then
       Open fname For Output As #fnum%
    ElseIf ((automatic And autosave) Or savehtml) And ii% = 1 Then
       Open fname For Append As #fnum%
       GoTo h100
       End If

Print #fnum%, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
Print #fnum%, "<HTML>"
Print #fnum%, "<HEAD>"
Print #fnum%, "    <TITLE>Your Chai Table</TITLE>"
Print #fnum%, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1255""/>"
Print #fnum%, "    <META NAME=""AUTHOR"" CONTENT=""Chaim Keller, Chai Tables, www.chaitables.com""/>"
Print #fnum%, "</HEAD>"
Print #fnum%, "<BODY>"
h100:
Print #fnum%, "<TABLE WIDTH=""696"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" ALIGN=""CENTER"" >"
Print #fnum%, "    <COL WIDTH=""696"">"
Print #fnum%, "    <TR>"
Print #fnum%, "        <TD WIDTH=""696"" VALIGN=""TOP"" BGCOLOR=""#ffffff"">"
Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""4"" STYLE=""font-size: 16pt"">" & storheader$(ii%, 0) & "</FONT></P>"
Print #fnum%, "        </TD>"
Print #fnum%, "    </TR>"
Print #fnum%, "    <TR>"
Print #fnum%, "        <TD WIDTH=""696"" VALIGN=""TOP"">"
Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 1) & "</FONT></P>"
Print #fnum%, "        </TD>"
Print #fnum%, "    </TR>"
Print #fnum%, "    <TR>"
Print #fnum%, "        <TD WIDTH=""696"" VALIGN=""TOP"">"
Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 3) & "</FONT></P>"
Print #fnum%, "        </TD>"
Print #fnum%, "    </TR>"
Print #fnum%, "</TABLE>"
Print #fnum%, "<TABLE WIDTH=""696"" BORDER=""1"" CELLPADDING=""4"" CELLSPACING=""0"" ALIGN=""CENTER"" >"
Print #fnum%, "    <COL WIDTH=""14"">"
Print #fnum%, "    <COL WIDTH=""40"">"
Print #fnum%, "    <COL WIDTH=""51"">"
Print #fnum%, "    <COL WIDTH=""39"">"
Print #fnum%, "    <COL WIDTH=""39"">"
Print #fnum%, "    <COL WIDTH=""40"">"
Print #fnum%, "    <COL WIDTH=""40"">"
Print #fnum%, "    <COL WIDTH=""43"">"
Print #fnum%, "    <COL WIDTH=""41"">"
Print #fnum%, "    <COL WIDTH=""44"">"
Print #fnum%, "    <COL WIDTH=""44"">"
Print #fnum%, "    <COL WIDTH=""45"">"
Print #fnum%, "    <COL WIDTH=""42"">"
Print #fnum%, "    <COL WIDTH=""39"">"
Print #fnum%, "    <COL WIDTH=""15"">"
Print #fnum%, "    <TR VALIGN=""TOP"">"
Print #fnum%, "        <TD WIDTH=""14"">"
Print #fnum%, "            <P ALIGN=""CENTER""><BR/>"
Print #fnum%, "            </P>"
Print #fnum%, "        </TD>"
If optionheb = True Then
    Print #fnum%, "        <TD WIDTH=""40"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 13) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""51"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 12) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""39"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 11) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""39"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 10) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""40"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 9) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""40"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 8) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""43"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 7) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""41"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 6) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""44"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 5) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""44"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 4) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""45"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 3) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""42"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 2) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""39"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 1) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""15"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><BR/>"
    Print #fnum%, "            </P>"
    Print #fnum%, "        </TD>"
 Else
    Print #fnum%, "        <TD WIDTH=""40"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 1) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""51"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 2) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""39"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 3) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""39"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 4) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""40"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 5) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""40"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 6) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""43"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 7) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""41"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 8) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""44"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 9) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""44"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 10) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""45"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 11) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""42"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 12) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""39"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & monthh$(iheb%, 13) & "</B></FONT></FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "        <TD WIDTH=""15"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><BR/>"
    Print #fnum%, "            </P>"
    Print #fnum%, "        </TD>"
    End If
Print #fnum%, "    </TR>"


'**********finish here*******
    iii% = -1
    For i% = 1 To 5 '5 rows of groups of 3 coloured, and 3 uncoloured
        'check for shabboses (_) and for near mountains (*)
        For j% = 1 To 3
           iii% = iii% + 1
           'Call hebnum(iii% + 1, cha$)
           If optionheb = True Then
              Call hebnum(iii% + 1, cha$)
           Else
              cha$ = Trim$(Str(iii% + 1))
              End If
              
            Print #fnum%, "    <TR VALIGN=""BOTTOM"">"
            Print #fnum%, "        <TD WIDTH=""14"" BGCOLOR=""#ffff00"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            'now determine if this time is underlined or shabbos and add appropriate HTML ******
            k% = 0
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""51"" BGCOLOR=""#ffff00"">"
            k% = 1
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 2
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 3
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 4
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 5
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""43"" BGCOLOR=""#ffff00"">"
            k% = 6
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""41"" BGCOLOR=""#ffff00"">"
            k% = 7
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""44"" BGCOLOR=""#ffff00"">"
            k% = 8
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""44"" BGCOLOR=""#ffff00"">"
            k% = 9
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""45"" BGCOLOR=""#ffff00"">"
            k% = 10
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""42"" BGCOLOR=""#ffff00"">"
            k% = 11
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 12
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""14"" BGCOLOR=""#ffff00"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
        Next j%
        For j% = 1 To 3
           iii% = iii% + 1
           'Call hebnum(iii% + 1, cha$)
           If optionheb = True Then
              Call hebnum(iii% + 1, cha$)
           Else
              cha$ = Trim$(Str$(iii% + 1))
              End If
           
            Print #fnum%, "    <TR VALIGN=""BOTTOM"">"
            Print #fnum%, "        <TD WIDTH=""14"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "        <TD WIDTH=""40"">"
            'now determine if this time is underlined or shabbos and add appropriate HTML ******
            k% = 0
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""51"">"
            k% = 1
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 2
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 3
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 4
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 5
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""43"">"
            k% = 6
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""41"">"
            k% = 7
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""44"">"
            k% = 8
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""44"">"
            k% = 9
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""45"">"
            k% = 10
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""42"">"
            k% = 11
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 12
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""14"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
                Next j%
            Next i%
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<TABLE WIDTH=""696"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" ALIGN=""CENTER"" >"
            Print #fnum%, "    <COL WIDTH=""696"">"
            Print #fnum%, "    <TR>"
            Print #fnum%, "        <TD WIDTH=""696"" VALIGN=""TOP"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 4) & "</FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<TABLE WIDTH=""696"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" ALIGN=""CENTER"" >"
            Print #fnum%, "    <COL WIDTH=""696"">"
            Print #fnum%, "    <TR>"
            Print #fnum%, "        <TD WIDTH=""696"" VALIGN=""TOP"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 5) & "</FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<P><BR/><BR/>"
            Print #fnum%, "</P>"
            
            If automatic And autosave And ii% = 0 Then GoTo h500
            
            Print #fnum%, "</BODY>"
            Print #fnum%, "</HTML>"
h500:
        Close #fnum%
        Screen.MousePointer = vbDefault
        Exit Sub
        
timestring:
        If optionheb = True Then
           ntim$ = stortim$(ii%, endyr% - k% - 1, iii%)
        Else
           ntim$ = stortim$(ii%, k%, iii%)
           End If
        positunder% = InStr(ntim$, "_")
        If positunder% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positunder% - 1)
           End If
        positnear% = InStr(ntim$, "*")
        If positnear% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positnear% - 1)
           End If
        If positunder% <> 0 Then
           ntim$ = "<u>" & ntim$ & "</u>"
           End If
        If positnear% <> 0 Then
           ntim$ = "<font color=""#33cc66"">" & ntim$ & "</font>"
           End If
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & ntim$ & "</FONT></P>"
        Print #fnum%, "        </TD>"
Return

End Sub

Private Sub civilHTM(fname)
     Screen.MousePointer = vbHourglass
     alsosunset = False
     If Abs(nsetflag%) = 1 Then
        ii% = 0
     Else
        ii% = 1
        End If
    iheb% = 0: If optionheb = False Then iheb% = 1
    Close
    
    fnum% = FreeFile
    If Not automatic And Not autosave And Not savehtml Then
       Open fname For Output As #fnum%
    ElseIf automatic And autosave And ii% = 0 Then
       Open fname For Output As #fnum%
    ElseIf ((automatic And autosave) Or savehtml) And ii% = 1 Then
       Open fname For Append As #fnum%
       GoTo h100
       End If
    
    Print #fnum%, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
    Print #fnum%, "<HTML>"
    Print #fnum%, "<HEAD>"
    Print #fnum%, "    <TITLE>Your Chai Table</TITLE>"
    Print #fnum%, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1255""/>"
    Print #fnum%, "    <META NAME=""AUTHOR"" CONTENT=""Chaim Keller, Chai Tables, www.chaitables.com""/>"
    Print #fnum%, "</HEAD>"
    Print #fnum%, "<BODY>"
    
h100:
    Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""  ALIGN=""CENTER"" >"
    Print #fnum%, "    <COL WIDTH=""670"">"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""4"" STYLE=""font-size: 16pt"">" & storheader$(ii%, 0) & "</FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "    </TR>"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 1) & "</FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "    </TR>"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 3) & "</FONT></P>"
    Print #fnum%, "        </TD>"
    Print #fnum%, "    </TR>"
    Print #fnum%, "</TABLE>"
    Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""1"" CELLPADDING=""4"" CELLSPACING=""0""  ALIGN=""CENTER"" >"
    Print #fnum%, "    <COL WIDTH=""16"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""46"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""40"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""40"">"
    Print #fnum%, "    <COL WIDTH=""40"">"
    Print #fnum%, "    <COL WIDTH=""41"">"
    Print #fnum%, "    <COL WIDTH=""39"">"
    Print #fnum%, "    <COL WIDTH=""15"">"
    Print #fnum%, "    <TR>"
    Print #fnum%, "        <TD WIDTH=""16"" VALIGN=""TOP"">"
    Print #fnum%, "            <P ALIGN=""LEFT""><BR/>"
    Print #fnum%, "            </P>"
    Print #fnum%, "        </TD>"
    If optionheb = True Then
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 12) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""46"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 11) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 10) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 9) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 8) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 7) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 6) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 5) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 4) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 3) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""41"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 2) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 1) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""15"" VALIGN=""BOTTOM"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><BR/>"
        Print #fnum%, "            </P>"
        Print #fnum%, "        </TD>"
    Else
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 1) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""46"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 2) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 3) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 4) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 5) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 6) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 7) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 8) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 9) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""40"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 10) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""41"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 11) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""39"" VALIGN=""TOP"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & montheh$(iheb%, 12) & "</B></FONT></P>"
        Print #fnum%, "        </TD>"
        Print #fnum%, "        <TD WIDTH=""15"" VALIGN=""BOTTOM"">"
        Print #fnum%, "            <P ALIGN=""CENTER""><BR/>"
        Print #fnum%, "            </P>"
        Print #fnum%, "        </TD>"
       End If
    Print #fnum%, "    </TR>"

    
    iii% = -1
    For i% = 1 To 5 '5 rows of groups of 3 coloured, and 3 uncoloured
        'check for shabboses (_) and for near mountains (*)
        For j% = 1 To 3
           iii% = iii% + 1
           'Call hebnum(iii% + 1, cha$)
           If optionheb = True Then
              Call hebnum(iii% + 1, cha$)
           Else
              cha$ = Trim$(Str(iii% + 1))
              End If
           
            Print #fnum%, "    <TR VALIGN=""BOTTOM"">"
            Print #fnum%, "        <TD WIDTH=""16"" BGCOLOR=""#ffff00"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            'now determine if this time is underlined or shabbos and add appropriate HTML ******
            k% = 0
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""46"" BGCOLOR=""#ffff00"">"
            k% = 1
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 2
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 3
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 4
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 5
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 6
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 7
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 8
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"" BGCOLOR=""#ffff00"">"
            k% = 9
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""41"" BGCOLOR=""#ffff00"">"
            k% = 10
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"" BGCOLOR=""#ffff00"">"
            k% = 11
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""16"" BGCOLOR=""#ffff00"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
        Next j%
        For j% = 1 To 3
           iii% = iii% + 1
           'Call hebnum(iii% + 1, cha$)
           If optionheb = True Then
              Call hebnum(iii% + 1, cha$)
           Else
              cha$ = Trim$(Str(iii% + 1))
              End If
            Print #fnum%, "    <TR VALIGN=""BOTTOM"">"
            Print #fnum%, "        <TD WIDTH=""16"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "        <TD WIDTH=""39"" >"
            'now determine if this time is underlined or shabbos and add appropriate HTML ******
            k% = 0
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""46"">"
            k% = 1
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 2
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 3
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 4
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 5
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 6
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 7
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 8
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 9
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""41"">"
            k% = 10
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 11
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""16"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
                Next j%
            Next i%
            
            'day 31
           iii% = iii% + 1
           'Call hebnum(iii% + 1, cha$)
           If optionheb = True Then
              Call hebnum(iii% + 1, cha$)
           Else
              cha$ = Trim$(Str(iii% + 1))
              End If
            Print #fnum%, "    <TR VALIGN=""BOTTOM"">"
            Print #fnum%, "        <TD WIDTH=""16"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "        <TD WIDTH=""39"" >"
            'now determine if this time is underlined or shabbos and add appropriate HTML ******
            k% = 0
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""46"">"
            k% = 1
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 2
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 3
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 4
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 5
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 6
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 7
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 8
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""40"">"
            k% = 9
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""41"">"
            k% = 10
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""39"">"
            k% = 11
            GoSub timestring
            Print #fnum%, "        <TD WIDTH=""16"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 10pt""><B>" & cha$ & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
            
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""  ALIGN=""CENTER"" >"
            Print #fnum%, "    <COL WIDTH=""670"">"
            Print #fnum%, "    <TR>"
            Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 4) & "</FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<TABLE WIDTH=""670"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" ALIGN=""CENTER"" >"
            Print #fnum%, "    <COL WIDTH=""670"">"
            Print #fnum%, "    <TR>"
            Print #fnum%, "        <TD WIDTH=""670"" VALIGN=""TOP"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & storheader$(ii%, 5) & "</FONT></P>"
            Print #fnum%, "        </TD>"
            Print #fnum%, "    </TR>"
            Print #fnum%, "</TABLE>"
            Print #fnum%, "<P><BR/><BR/>"
            Print #fnum%, "</P>"
            
            If automatic And autosave And ii% = 0 Then GoTo h500
            
            Print #fnum%, "</BODY>"
            Print #fnum%, "</HTML>"
h500:
        Close #fnum%
        Screen.MousePointer = vbDefault
        Exit Sub
        
timestring:
        If optionheb = True Then
           ntim$ = stortim$(ii%, endyr% - k% - 1, iii%)
        Else
           ntim$ = stortim$(ii%, k%, iii%)
           End If
        positunder% = InStr(ntim$, "_")
        If positunder% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positunder% - 1)
           End If
        positnear% = InStr(ntim$, "*")
        If positnear% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positnear% - 1)
           End If
        If positunder% <> 0 Then
           ntim$ = "<u>" & ntim$ & "</u>"
           End If
        If positnear% <> 0 Then
           ntim$ = "<font color=""#33cc66"">" & ntim$ & "</font>"
           End If
        Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & ntim$ & "</FONT></P>"
        Print #fnum%, "        </TD>"
Return
    
End Sub

Private Sub civilListHTM(fname)
     Screen.MousePointer = vbHourglass
     alsosunset = False
     If Abs(nsetflag%) = 1 Then
        ii% = 0
     Else
        ii% = 1
        End If
    iheb% = 0: If optionheb = False Then iheb% = 1
    Close
    
    fnum% = FreeFile
    If Not automatic And Not autosave And Not savehtml Then
       Open fname For Output As #fnum%
    ElseIf automatic And autosave Then
       myfile$ = Dir(fname)
       If Trim$(myfile$) = sEmpty Then
          Open fname For Output As #fnum%
       Else
          Open fname For Append As #fnum%
          GoTo h100
          End If
       End If
    
    Print #fnum%, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
    Print #fnum%, "<HTML>"
    Print #fnum%, "<HEAD>"
    Print #fnum%, "    <TITLE>Your Chai Table</TITLE>"
    Print #fnum%, "    <META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1255""/>"
    Print #fnum%, "    <META NAME=""AUTHOR"" CONTENT=""Chaim Keller, Chai Tables, www.chaitables.com""/>"
    Print #fnum%, "</HEAD>"
    Print #fnum%, "<BODY>"
    
h100:
    Print #fnum%, "<TABLE WIDTH=""100"" BORDER=""1"" CELLPADDING=""4"" CELLSPACING=""0""  ALIGN=""CENTER"" >"
    Print #fnum%, "    <COL WIDTH=""40"">"
    Print #fnum%, "    <COL WIDTH=""60"">"

    j% = 0
    
    For k% = 11 To 0 Step -1 'months
    
       j% = j% + 1
    
       For iii% = 0 To 30 'days
            
          ntim$ = sEmpty
          GoSub timestring
          If Trim$(ntim$) = sEmpty Then
             'don't add row
          Else
             'create row consisting of time | date
             
            'time
            Print #fnum%, "    <TR>"
            Print #fnum%, "        <TD WIDTH=""40"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT SIZE=""1"" STYLE=""font-size: 8pt"">" & ntim$ & "</FONT></P>"
            Print #fnum%, "        </TD>"
            
            'date
            Print #fnum%, "        <TD WIDTH=""60"">"
            Print #fnum%, "            <P ALIGN=""CENTER""><FONT FACE=""Arial, sans-serif""><FONT SIZE=""1"" STYLE=""font-size: 8pt""><B>" & Trim$(Str$(iii% + 1)) & "/" & Trim$(Str$(j%)) & "/" & Trim$(Caldirectories.Combo1.Text) & "</B></FONT></P>"
            Print #fnum%, "        </TD>"
        
            Print #fnum%, "    </TR>"
            End If
       
       Next iii%
    
    Next k%
    

    Print #fnum%, "</TABLE>"
    
    If automatic And autosave And Caldirectories.chkListAuto.Value = vbChecked Then GoTo h500
    
    Print #fnum%, "</BODY>"
    Print #fnum%, "</HTML>"
h500:
    If Caldirectories.chkListAuto.Value = vbChecked And _
       Mid$(Trim$(Caldirectories.txtCivil.Text), Len(Trim$(Caldirectories.txtCivil.Text)) - 3, 4) = Trim$(Caldirectories.Combo1.Text) Then
       Print #fnum%, "</BODY>"
       Print #fnum%, "</HTML>"
       End If
    
    Close #fnum%
    Screen.MousePointer = vbDefault
    Exit Sub
        
timestring:
        If optionheb = True Then
           ntim$ = stortim$(ii%, endyr% - k% - 1, iii%)
        Else
           ntim$ = stortim$(ii%, k%, iii%)
           End If
        If Trim$(ntim$) = sEmpty Then Return
        positunder% = InStr(ntim$, "_")
        If positunder% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positunder% - 1)
           End If
        positnear% = InStr(ntim$, "*")
        If positnear% <> 0 Then
           ntim$ = Mid$(ntim$, 1, positnear% - 1)
           End If
        If positunder% <> 0 Then
           ntim$ = "<u>" & ntim$ & "</u>"
           End If
        If positnear% <> 0 Then
           ntim$ = "<font color=""#33cc66"">" & ntim$ & "</font>"
           End If
Return
    
End Sub

