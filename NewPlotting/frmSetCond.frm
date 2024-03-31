VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{825967DA-1756-11D3-B695-ED78B587442C}#30.0#0"; "FlexListBox.ocx"
Begin VB.Form frmSetCond 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entries"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   1335
   ClientWidth     =   2940
   Icon            =   "frmSetCond.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   2940
   Begin ComctlLib.StatusBar StatusBarflx 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   8205
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5133
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmFiles 
      Caption         =   "Plot File Buffer"
      Height          =   3975
      Left            =   60
      TabIndex        =   24
      Top             =   0
      Width           =   2775
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1980
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FlexList.FlexListBox flxlstFiles 
         Height          =   2955
         Left            =   60
         TabIndex        =   28
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   5212
         BackColorUnselected=   14941439
         BackColorSelected=   8388608
         BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiSelect     =   -1  'True
         MousePreSelector=   0   'False
      End
      Begin VB.CommandButton cmdShowEdit 
         Caption         =   "&Show/Edit"
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Top             =   3600
         Width           =   1275
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Plot &All"
         Height          =   315
         Left            =   60
         TabIndex        =   26
         Top             =   3600
         Width           =   1395
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   3300
         Width           =   1275
      End
      Begin VB.CommandButton cmdWizard 
         Caption         =   "&Plot Wizard"
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   3300
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Plot"
      Height          =   435
      Left            =   900
      TabIndex        =   13
      Top             =   8940
      Width           =   975
   End
   Begin VB.Frame fraLayout 
      Caption         =   "Layout"
      Height          =   4155
      Left            =   60
      TabIndex        =   12
      Top             =   4020
      Width           =   2775
      Begin VB.CommandButton cmdFull 
         Caption         =   "Full"
         Height          =   495
         Left            =   2280
         TabIndex        =   42
         ToolTipText     =   "Restore Full Range"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox PicStatus 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   41
         Top             =   4080
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDownAxisLabelSize 
         Height          =   285
         Left            =   2400
         TabIndex        =   39
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   10
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtAxisLabelSize"
         BuddyDispid     =   196618
         OrigLeft        =   2400
         OrigTop         =   480
         OrigRight       =   2655
         OrigBottom      =   735
         Max             =   40
         Min             =   8
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtAxisLabelSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   38
         Text            =   "12"
         ToolTipText     =   "Axis label font size"
         Top             =   480
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDownTitle 
         Height          =   285
         Left            =   2400
         TabIndex        =   37
         Top             =   3720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   14
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtTitlefont"
         BuddyDispid     =   196619
         OrigLeft        =   2400
         OrigTop         =   3720
         OrigRight       =   2655
         OrigBottom      =   3975
         Max             =   40
         Min             =   8
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTitlefont 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   36
         Text            =   "14"
         ToolTipText     =   "Chart Title font size"
         Top             =   3720
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDownTitleY 
         Height          =   285
         Left            =   2400
         TabIndex        =   35
         Top             =   3360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   17
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtTitleYfont"
         BuddyDispid     =   196620
         OrigLeft        =   2520
         OrigTop         =   3360
         OrigRight       =   2775
         OrigBottom      =   3615
         Max             =   40
         Min             =   8
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTitleYfont 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   34
         Text            =   "17"
         ToolTipText     =   "Y axis label font size"
         Top             =   3360
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDownTitleX 
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   3000
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   17
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtTitleXfont"
         BuddyDispid     =   196621
         OrigLeft        =   2400
         OrigTop         =   3000
         OrigRight       =   2655
         OrigBottom      =   3375
         Max             =   40
         Min             =   8
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTitleXfont 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   32
         Text            =   "17"
         ToolTipText     =   "X axis label font size"
         Top             =   3000
         Width           =   375
      End
      Begin VB.ListBox ListSort 
         Height          =   255
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTitle 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Chart Title"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtValueX1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   555
      End
      Begin VB.TextBox txtValueX0 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   555
      End
      Begin VB.TextBox txtYTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Y axis label"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtXTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtValueY1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   555
      End
      Begin VB.TextBox txtValueY0 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox txtX1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Text            =   "0"
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox txtX0 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   840
         Width           =   555
      End
      Begin VB.CheckBox chkGridLine 
         Caption         =   "Check1"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkOrigin 
         Caption         =   "Check1"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lblFontSize 
         Caption         =   "Font Size"
         Height          =   255
         Left            =   1300
         TabIndex        =   40
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   3765
         Width           =   375
      End
      Begin VB.Label lblEndX1 
         AutoSize        =   -1  'True
         Caption         =   "maximum value X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   21
         Top             =   2700
         Width           =   1680
      End
      Begin VB.Label lblStartX0 
         AutoSize        =   -1  'True
         Caption         =   "minimum value X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label lblYTitle 
         AutoSize        =   -1  'True
         Caption         =   "Title Y-axis"
         Height          =   195
         Left            =   1080
         TabIndex        =   23
         Top             =   3420
         Width           =   765
      End
      Begin VB.Label lblXTitle 
         AutoSize        =   -1  'True
         Caption         =   "Title X-axis"
         Height          =   195
         Left            =   1140
         TabIndex        =   22
         ToolTipText     =   "X axis label"
         Top             =   3060
         Width           =   765
      End
      Begin VB.Label lblEndY 
         AutoSize        =   -1  'True
         Caption         =   "maximum value Y-range"
         Height          =   195
         Left            =   840
         TabIndex        =   19
         Top             =   1980
         Width           =   1680
      End
      Begin VB.Label lblStartY 
         AutoSize        =   -1  'True
         Caption         =   "minimum value Y-range"
         Height          =   195
         Left            =   840
         TabIndex        =   18
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Label lblIndexEnd 
         AutoSize        =   -1  'True
         Caption         =   "index End X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   17
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label lblIndexStart 
         AutoSize        =   -1  'True
         Caption         =   "index Start X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   900
         Width           =   1350
      End
      Begin VB.Label lblGridLine 
         AutoSize        =   -1  'True
         Caption         =   "Gridlines?"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   540
         Width           =   690
      End
      Begin VB.Label lblOrigin 
         AutoSize        =   -1  'True
         Caption         =   "Show origin?"
         Height          =   195
         Left            =   420
         TabIndex        =   14
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "&Open"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuOpen1 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuspacer 
         Caption         =   "-"
      End
      Begin VB.Menu cmdPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuRestoreOther 
         Caption         =   "Restore Other"
         Enabled         =   0   'False
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaths 
         Caption         =   "P&aths"
         Shortcut        =   ^A
      End
      Begin VB.Menu spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
   End
   Begin VB.Menu mnuSpline 
      Caption         =   "F&itting"
   End
   Begin VB.Menu mnuStatistics 
      Caption         =   "&Statistics"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadme 
         Caption         =   "&Readme"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Plot"
      End
   End
End
Attribute VB_Name = "frmSetCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmSetCond
' Date      : 02/06/2003
' Author    : Chaim Keller

'This is a general plot program programmed by Chaim Keller.
'It was originally developed to plot the bathymetric sounding
'data of John K. Hall of the GSI.  This program is based on a
'modification of freeware VB code from several sources:
'(1) Graphics program code (peeWee Technologies) and (2) Flex list box Active X
'control (Ted Schopenhouer) from www.FreeVBcode.com.
'(3) Multi-select dialog box code from www.mvps.org/vbnet.
'(I express my gratitude to theose organizations and software engineers.)
'---------------------------------------------------------------------------------------

Option Explicit

Public N As Integer 'counter
Public M As Integer 'counter
Public dYmax As Double
Public dYmin As Double
Public dXmin As Double
Public dXmax As Double



Private Sub cmdAll_Click()
  Dim response As String
  
'10
'  response = InputBox("Note: files must all have the same format." & vbLf & _
'                    "Input the format number of these files (1-11).", _
'                    "Files' format number", Str(DefaultFileType%))
'  Select Case Trim$(response)
'     Case sEmpty
'        'cancel or escape so exit sub
'        Exit Sub
'     Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"
'        DefaultFileType% = Val(response)
'     Case Else
'        response = MsgBox("The range of valid inputs is 1-11.", vbOKCancel + vbExclamation, "Plot error")
'        If response = vbOK Then
'           GoTo 10
'        Else
'           Exit Sub
'           End If
'  End Select
     
  PlotAll = True
  Dim I%
  For I% = 1 To flxlstFiles.list.Count    'lstFiles.ListCount '<<<!!!
      flxlstFiles.list.item(I%).Selected = True 'lstFiles.Selected(i% - 1) = True
  Next I%
  cmdWizard_Click
  
'      .AddItem "Hello"
'      .List.item(1).Selected = True
'
'      .AddItem "You"
'      .List.item(2).Enabeled = False
'      .List.item(2).ItemForeColor = RGB(0, 255, 0)
'      .List.item(2).ItemBackColor = RGB(255, 255, 0)
'
'      .AddItem ""
'      .List.item(3).Enabeled = False
'
'      .AddItem ""
'      .List.item(4).Enabeled = False
'
'      .AddItem "A Big"
'      .List.item(5).Selected = True
'
End Sub

Private Sub cmdClear_Click()
  flxlstFiles.Clear  'lstFiles.Clear
  'clear the plotinfo memory
  ReDim PlotInfo(9, 0) 'clear last plot info and memory
  'clear the plot data memory
  ReDim RecordSize(0)
  ReDim dPlot(maxFilesToPlot%, 1, 0) As Double
  numRowsToNow% = 0
  If JKHplotVis Then
     Unload JKHplot
     Set JKHplot = Nothing
     End If
  frmSetCond.txtX1 = 0
  numfiles% = 0 'initialize plot buffer
  frmDraw.Cls
  Unload frmDraw
  
  mnuSave.Enabled = False
  mnuSpline.Enabled = False
  cmdShowEdit.Enabled = False
  cmdWizard.Enabled = False
  cmdAll.Enabled = False
  cmdClear.Enabled = False
   
   chkOrigin.Enabled = False
   chkGridLine.Enabled = False
   txtX0.Enabled = False
   txtX1.Enabled = False
   txtValueY0.Enabled = False
   txtValueY1.Enabled = False
   txtXTitle.Enabled = False
   txtYTitle.Enabled = False
   txtTitle.Enabled = False
   txtXTitle.Text = sEmpty
   txtYTitle.Text = sEmpty
   txtTitle.Text = sEmpty
   txtValueX0.Enabled = False
   txtValueX1.Enabled = False
   fraLayout.Enabled = False
   frmSetCond.lblEndX1.Enabled = False
   frmSetCond.lblEndY.Enabled = False
   frmSetCond.lblGridLine.Enabled = False
   frmSetCond.lblIndexEnd.Enabled = False
   frmSetCond.lblOrigin.Enabled = False
   frmSetCond.lblIndexStart.Enabled = False
   frmSetCond.lblStartX0.Enabled = False
   frmSetCond.lblStartY.Enabled = False
   frmSetCond.lblXTitle.Enabled = False
   frmSetCond.lblYTitle.Enabled = False
   frmSetCond.lblTitle.Enabled = False
   txtX0 = sEmpty
   txtX1 = sEmpty
   txtValueY0 = sEmpty
   txtValueY1 = sEmpty
   txtValueX0 = sEmpty
   txtValueX1 = sEmpty
  
End Sub

Private Sub cmdFull_Click()
   frmSetCond.txtValueY0 = YMin0
   frmSetCond.txtValueY1 = YRange0
   frmSetCond.txtValueX0 = XMin0
   frmSetCond.txtValueX1 = XRange0
   DblClickForm
   dragbegin = False
End Sub

Private Sub cmdOK_Click()

'input plot data into the plot arrays dPlot
ReadValues

'Define window Layout
DefineLayout
  
'plot the data
Plot frmDraw, dPlot, udtMyGraphLayout

'redo window Layout if necessary
Dim redo%
redo% = 0
If Val(txtX0.Text) > LBound(dPlot, 3) Then
   txtX0.Text = LBound(dPlot, 3)
   redo% = 1
   End If
If Val(txtX1.Text) < UBound(dPlot, 3) Then
   txtX1.Text = UBound(dPlot, 3)
   redo% = 1
   End If

If redo% = 1 Then DblClickForm

'record default screen parameters
drm% = frmDraw.DrawMode
drw% = frmDraw.DrawWidth
drs% = frmDraw.DrawStyle
frmSetCond.txtValueX0 = XMin0
frmSetCond.txtValueX1 = XRange0
XMin0 = XMin0
XRange0 = XRange0

If JKHplotVis Then 'let it appear again
   Dim ret As Long
   ret = BringWindowToTop(JKHplot.hWnd)
   End If
End Sub


Private Sub cmdPrint_Click()
   EndPlot = False 'flag beginning of plot
   
   frmDraw.WindowState = vbMaximized

   Do
      DoEvents
      If EndPlot Then
        ScreenDump = True
        PrintPreview.Visible = True
        Exit Do
        End If
   Loop
   
   frmDraw.WindowState = vbNormal
   Dim ret As Long
   ret = BringWindowToTop(PrintPreview.hWnd)

End Sub

Private Sub cmdShowEdit_Click()
   
   'show and edit a listed file
   If Dir(dirWordpad & "Wordpad.exe") = sEmpty Then
      MsgBox "Path to Wordpad.exe incorrect or undefined!" & vbLf & _
             "Use the File/Paths option to define its path", _
             vbOKCancel + vbExclamation
      Exit Sub
      End If
       
   Dim I%, ret As Long
     
   Dim found%
   found% = 0
   For I% = 1 To flxlstFiles.list.Count  'lstFiles.ListCount
      If flxlstFiles.list.item(I%).Selected Then 'lstFiles.Selected(i% - 1) Then
         found% = 1
         ret = Shell(dirWordpad & "Wordpad.exe " & Chr$(34) & Files(I% - 1) & Chr$(34), vbNormalFocus)
         End If
   Next I%
   
   If found% = 0 Then
      MsgBox "You haven't selected any files!", vbExclamation + vbOKOnly, "Plot"
      End If
      
End Sub

Private Sub cmdWizard_Click()
  'cycle through selected files and set plot info
   On Error GoTo cmdWizard_Click_Error

  If flxlstFiles.list.Count = 0 Then
     MsgBox "Plot buffer is empty!", vbExclamation + vbOKOnly, "Plot"
     Exit Sub
     End If
  
  maxFilesToPlot% = MaxNumOverplotFiles
  numFilesToPlot% = 0
  
  cmdWizard.Enabled = False
  cmdClear.Enabled = False
  mnuFormat.Enabled = False
  mnuOpen.Enabled = False
  
  Dim I As Integer, tmpForeColor&
  Dim tmpBackColor&, waitime As Single
  
  If flxlstFiles.list.Count > 15 Then
    'scroll the flex list box to the beginning
    For I = 1 To flxlstFiles.list.Count
       flxlstFiles.SetFocus
       waitime = Timer
       Do Until Timer > waitime + 0.001
          DoEvents
       Loop
       Call keybd_event(VK_UP, 0, 0, 0)
       Call keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)
    Next I
    End If
           
  numSelected% = -1
  For I = 0 To flxlstFiles.list.Count - 1
        'scroll down one
        flxlstFiles.SetFocus
        waitime = Timer
        Do Until Timer > waitime + 0.001
           DoEvents
        Loop
        flxlstFiles.SetFocus
        Call keybd_event(VK_DOWN, 0, 0, 0)
        Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
        
        If flxlstFiles.list.item(I + 1).Selected Then
            'scroll down one item
            flxlstFiles.SetFocus
            waitime = Timer
            Do Until Timer > waitime + 0.001
               DoEvents
            Loop
            flxlstFiles.SetFocus
            Call keybd_event(VK_DOWN, 0, 0, 0)
            Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
            
           numSelected% = I
           numFilesToPlot% = numFilesToPlot% + 1
           If I% > maxFilesToPlot% Then
              MsgBox "You are allowed to overplot a maximum of " & _
              Str(maxFilesToPlot%) & "files!", vbExclamation + vbOKOnly, "Plot"
              Exit For
              End If
           tmpBackColor& = flxlstFiles.list.item(I + 1).ItemBackColor
           tmpForeColor& = flxlstFiles.list.item(I + 1).ItemForeColor
           Call keybd_event(VK_DOWN, 0, 0, 0)
           Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
           flxlstFiles.list.item(I + 1).ItemBackColor = &HFF&
           flxlstFiles.list.item(I + 1).ItemForeColor = QBColor(0)
           flxlstFiles.Refresh
           PlotInfofrm.lblFileName = Files(I)
           PlotInfofrm.Visible = True
           'wait here until form becomes invisible
           Do Until PlotInfofrmVis = False
              DoEvents
              If PlotAll Then
                 If Not PlotInfofrmVis Then Exit Do
                 If DefaultFileType% <> 0 Then 'click on file type of all the files
                    Select Case DefaultFileType%
                       Case 1
                          PlotInfofrm.optFF1.Value = True
                       Case 2
                          PlotInfofrm.optFF2.Value = True
                       Case 3
                          PlotInfofrm.optFF3.Value = True
                       Case 4
                          PlotInfofrm.optFF4.Value = True
                       Case 5
                          PlotInfofrm.optFF5.Value = True
                       Case 6
                          PlotInfofrm.optFF6.Value = True
                       Case 7
                          PlotInfofrm.optFF7.Value = True
                       Case 8
                          PlotInfofrm.optFF8.Value = True
                       Case 9
                          PlotInfofrm.optFF9.Value = True
                       Case 10
                          PlotInfofrm.optFF10.Value = True
                       Case 11
                          PlotInfofrm.optFF11.Value = True
                    End Select
                    End If
                 PlotInfofrm.chkSave.Value = vbChecked
                 'give some time to fill in other options
                 If I > 0 Then 'wait for accept on first file
                    PlotInfofrm.cmdAccept.Value = True
                    End If
                 End If
              flxlstFiles.list.item(I + 1).ItemBackColor = tmpBackColor&
              flxlstFiles.list.item(I + 1).ItemForeColor = tmpForeColor&
              If Not PlotInfofrmVis Then Exit Do
           Loop
           If Not PlotInfoCancel Then
              'user canceled, so leave the wizard
              flxlstFiles.list.item(I + 1).ItemBackColor = tmpBackColor&
              flxlstFiles.list.item(I + 1).ItemForeColor = tmpForeColor&
              flxlstFiles.Refresh
              numFilesToPlot% = 0
              cmdWizard.Enabled = True
              cmdClear.Enabled = True
              mnuFormat.Enabled = True
              mnuOpen.Enabled = True
              mnuSpline.Enabled = True
              Exit Sub
              End If
        End If
  Next I
  flxlstFiles.Refresh
  cmdWizard.Enabled = True
  cmdClear.Enabled = True
  mnuFormat.Enabled = True
  mnuOpen.Enabled = True
  mnuSpline.Enabled = True
  
  If numSelected% = -1 Then
     MsgBox "No file was selected for plotting!", vbExclamation + vbOKOnly, "Plot"
     Exit Sub
     End If

Load frmDraw
frmDraw.Show vbModeless
frmDraw.Width = Screen.Width - frmSetCond.Width
frmDraw.ScaleHeight = frmSetCond.ScaleHeight
frmDraw.Left = frmSetCond.Width
If txtXTitle.Text = "" Then txtXTitle.Text = "X-values"
If txtYTitle.Text = "" Then txtYTitle.Text = "Y-values"
'If txtTitle.Text = "" Then txtTitle.Text = "Chart Title (blank)"
cmdOK_Click
PlotAll = False 'reset plt all flag

   On Error GoTo 0
   Exit Sub

cmdWizard_Click_Error:

    If Err.Number = 384 Then 'drawing form is minizied, to maximized it
        frmDraw.WindowState = vbNormal
        Resume
        End If
        
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdWizard_Click of Form frmSetCond"

End Sub



Private Sub flxlstFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, Listitem As String)
Call DispListItems(CLng(Listitem))
End Sub
Private Sub DispListItems(Listitem As Long)
Static lItem As Long
If lItem <> Listitem And Listitem > 0 Then
   lItem = Listitem
   With flxlstFiles.list.item(lItem)
                
    StatusBarPanelText frmSetCond.StatusBarflx, frmSetCond.PicStatus, 1, FileRoot(.Text), _
       &HFFFFC0, QBColor(1), 0
'      frmSetCond.StatusBarflx.Panels(1).Text = FileRoot(.Text) 'display root filename in statusbar
'      lblItems(0) = .Text
'      lblItems(1) = .UnderlayingValue
'      lblItems(2) = .Enabeled
'      lblItems(3) = .Selected
'      lblItems(4) = .ToolTipText
   End With
Else
End If

End Sub
''---------------------------------------------------------------------------------------
'' Procedure : flxlstFiles_KeyDown
'' Author    : Dr-John-K-Hall
'' Date      : 4/23/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub flxlstFiles_KeyDown(KeyCode As Integer, Shift As Integer)
''   'utility to stuff a long file name into the tool tip for viewing
'   'to use: use the up and down keys to stuff the flxlist filenames into the tooltip for display
'   Dim J As Integer
'   On Error GoTo flxlstFiles_KeyDown_Error
'
'   Select Case KeyCode
'      Case vbKeyDown
'         If frmSetCond.flxlstFiles.ListIndex < frmSetCond.flxlstFiles.List.Count - 1 Then
'            frmSetCond.flxlstFiles.ListIndex = frmSetCond.flxlstFiles.ListIndex + 1
'            frmSetCond.flxlstFiles.ToolTipText = frmSetCond.flxlstFiles.List.item(frmSetCond.flxlstFiles.ListIndex + 1).Text
'            End If
'
'      Case vbKeyUp
'         If frmSetCond.flxlstFiles.ListIndex > 0 Then
'            frmSetCond.flxlstFiles.ListIndex = frmSetCond.flxlstFiles.ListIndex - 1
'            frmSetCond.flxlstFiles.ToolTipText = frmSetCond.flxlstFiles.List.item(frmSetCond.flxlstFiles.ListIndex + 1).Text
'            End If
'
'      Case Else
'   End Select
'
'   On Error GoTo 0
'   Exit Sub
'
'flxlstFiles_KeyDown_Error:
'
'End Sub

Private Sub Form_Load()
   Dim filplt%, oldformat%
   
   On Error GoTo errhand

   'look for information on default directories
   'first loop for old info files and transfer them
   'to the App.Path directory
     
   If Dir("c:\PlotFiles.txt") <> sEmpty And _
      Dir(App.Path & "\PlotFiles.txt") = sEmpty Then
         FileCopy "c:\PlotFiles.txt", App.Path & "\PlotFiles.txt"
      End If
      
   direct$ = App.Path
      
   Dim doclin$, DirectFilePath As String, FileIn As Integer
   If Dir(direct$ & "\PlotDirec.txt") <> sEmpty Then
      filplt% = FreeFile
      oldformat% = 0
      Open direct$ & "\PlotDirec.txt" For Input As #filplt%
      Input #filplt%, doclin$
      Input #filplt%, direct$, directPlot$, dirWordpad
      oldformat% = 1
      Input #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
      Close #filplt%
   Else
      'ask user to find plotfiles
      Call MsgBox("Can't find the ""PlotDirect.txt"" file" _
                  & vbCrLf & "This file contains the directory structure for this program." _
                  & vbCrLf & "Please browse for it." _
                  , vbInformation, "Can't find the PlotDirec.txt file")
Load500:
      CommonDialog1.CancelError = True
      CommonDialog1.Filter = "PlotDirect.txt file|*.txt|All files (*.*)|*.*"
      CommonDialog1.ShowOpen
      DirectFilePath = CommonDialog1.FileName
      FileIn = FreeFile
      Open DirectFilePath For Input As #FileIn
      Line Input #FileIn, doclin$
      Close #FileIn
      If Not InStr(doclin$, "This file is used by Plot. Don't erase it!") Then
      
        Select Case MsgBox("This is not the right file!" _
                           & vbCrLf & "" _
                           & vbCrLf & "Do you want to try again?" _
                           , vbOKCancel Or vbInformation Or vbDefaultButton1, "Not the right file")
        
            Case vbOK
              GoTo Load500
            Case vbCancel
        End Select
        
      Else
        FileIn = FreeFile
        oldformat% = 0
        Open DirectFilePath For Input As #filplt%
        Input #filplt%, doclin$
        Input #filplt%, direct$, directPlot$, dirWordpad
        oldformat% = 1
        Input #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
        Close #filplt%
        End If
    End If
    
   'load up saved formats if they exist
   'if they don't exist, warn user
   
   Dim fil%, I%, num%
   If Dir(direct$ + "\FilFormat.txt") = "" Then
      Dim response
      response = MsgBox("Can't find stored file formats." & vbLf & _
             "Do you wan't to browse for it?", vbYesNoCancel + vbExclamation, "Plot")
      If response = vbYes Then
      
           Dim sFilters As String
          
           Dim pos As Long
           Dim buff As String
           Dim sLongname As String
           Dim sShortname As String
        
           'string of filters for the dialog box
           sFilters = "Text documents (.txt)" & vbNullChar & "*.txt" & vbNullChar & _
                     "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
              
           With OFN
                'size of the OFN structure
                .nStructSize = Len(OFN)
                
                 'window owning the dialog
                .hWndOwner = frmSetCond.hWnd
                
                'filters (patterns) for the dropdown combo
                .sFilter = sFilters
                
                'index to the default filter
                .nFilterIndex = 1
                
                'default filename, plus additional padding
                'for the user's final selection(s).  Must be
                'double-null terminated
                .sFile = "FilFormat.txt" & Space$(2048) & vbNullChar & vbNullChar
                
                'the size of the buffer
                .nMaxFile = Len(.sFile)
                
                'default extension applied to
                'file if it has no extension
                .sDefFileExt = "bas" & vbNullChar & vbNullChar
                
                'space fot he file title if a single selection
                'made, double-null terminated, and its size
                .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
                
                'starting folder, double-null terminated
                .nMaxTitle = Len(OFN.sFileTitle)
                
                'the dialog title
                If direct$ = "" Then direct$ = CurDir
                .sInitialDir = direct$ & vbNullChar & vbNullChar
                
                'default open flags and multiselect
                .sDialogTitle = "Open files' format file"
                .flags = OFS_FILE_OPEN_FLAGS
          End With
          
          FileOffset = OFN.nFileOffset 'store the file offset '<<<<changes 11/20/19
          
         If GetOpenFileName(OFN) Then
            buff = TrimNull(Trim$(Left$(OFN.sFile, Len(OFN.sFile) - 2)))
            direct$ = Left$(OFN.sFile, OFN.nFileOffset)
            
            'if has trailing "\" remove it
            If Mid$(direct$, Len(direct$), 1) = "\" Then
               direct$ = Mid$(direct$, 1, Len(direct$) - 1)
               End If
            
            'now put it where it belongs
            FileCopy direct$ & "\FilFormat.txt", App.Path & "\FilFormat.txt"
            
            direct$ = App.Path
            
            'record this information
            filplt% = FreeFile
            Open App.Path & "\PlotDirec.txt" For Output As #filplt%
            Write #filplt%, "This file is used by Plot. Don't erase it!"
            Write #filplt%, direct$, directPlot$, dirWordpad
            Write #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
            Close #filplt%
            
            fil% = FreeFile
            Open buff For Input As #fil%
            Line Input #fil%, doclin$
            For I% = 0 To 10
               Input #fil%, num%, FilForm(0, I%), FilForm(1, I%), FilForm(2, I%), FilForm(3, I%), FilForm(4, I%)
            Next I%
            Close #fil%
            
         Else
            MsgBox "You'll have to reenter the file formats." & vbLf & _
                   "Use the Format menu.", vbExclamation + vbOKOnly, "Plot"
            End If
       Else
         MsgBox "You'll have to reenter the file formats." & vbLf & _
                "Use the Format menu.", vbExclamation + vbOKOnly, "Plot"
         End If
   Else
      fil% = FreeFile
      Open direct$ + "\FilFormat.txt" For Input As #fil%
      Line Input #fil%, doclin$
      For I% = 0 To 10
         Input #fil%, num%, FilForm(0, I%), FilForm(1, I%), FilForm(2, I%), FilForm(3, I%), FilForm(4, I%)
      Next I%
      Close #fil%
      
      'now record the formats to the app.path
      If UCase$(direct$) <> UCase(App.Path) Then
         FileCopy direct$ & "\FilFormat.txt", App.Path & "\FilFormat.txt"
         
         direct$ = App.Path
           
        'record this information
        filplt% = FreeFile
        Open App.Path & "\PlotDirec.txt" For Output As #filplt%
        Write #filplt%, "This file is used by Plot. Don't erase it!"
        Write #filplt%, direct$, directPlot$, dirWordpad
        Write #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
        Close #filplt%
        
        End If
   
   End If

   numfiles% = 0 'initialize plot buffer
   
   'check if file list exists
   'if it exists then enable restore menu
   If Dir(App.Path & "\PlotFiles.txt") <> sEmpty Then
      mnuRestore.Enabled = True
      mnuRestoreOther.Enabled = True
      End If
      
   'Nothing inputed yet to the file buffer
   'so disenable plotting buttons.
   mnuSpline.Enabled = False
   cmdShowEdit.Enabled = False
   cmdWizard.Enabled = False
   cmdAll.Enabled = False
   cmdClear.Enabled = False
   chkOrigin.Enabled = False
   chkGridLine.Enabled = False
   txtX0.Enabled = True
   txtX1.Enabled = True
   txtValueY0.Enabled = False
   txtValueY1.Enabled = False
   txtXTitle.Enabled = False
   txtYTitle.Enabled = False
   txtTitle.Enabled = False
   txtValueX0.Enabled = False
   txtValueX1.Enabled = False
   fraLayout.Enabled = False
   frmSetCond.lblEndX1.Enabled = False
   frmSetCond.lblEndY.Enabled = False
   frmSetCond.lblGridLine.Enabled = False
   frmSetCond.lblIndexEnd.Enabled = False
   frmSetCond.lblOrigin.Enabled = False
   frmSetCond.lblIndexStart.Enabled = False
   frmSetCond.lblStartX0.Enabled = False
   frmSetCond.lblStartY.Enabled = False
   frmSetCond.lblXTitle.Enabled = False
   frmSetCond.lblYTitle.Enabled = False
   txtX0 = sEmpty
   txtX1 = sEmpty
   txtValueY0 = sEmpty
   txtValueY1 = sEmpty
   txtValueX0 = sEmpty
   txtValueX1 = sEmpty
Exit Sub

errhand:
    If Err.Number = 70 Then Resume Next
    If Err.Number = 62 And oldformat% = 1 Then Resume Next
    MsgBox "Error Number: " & Err.Number & " encountered!" & vbLf & _
           Err.Description & vbLf & _
           "You'll have to enter the file formats again." & vbLf & _
           "Use the format menu.", vbExclamation + vbOKOnly, "Plot"
End Sub
Public Sub DefineLayout()

Dim nSum As Integer 'sum of number of traces to plot

   On Error GoTo DefineLayout_Error

With udtMyGraphLayout
  .XTitle = txtXTitle.Text
  .YTitle = txtYTitle.Text
  .Title = txtTitle.Text
  If chkOrigin.Value = 0 Then
    .blnOrigin = False
    Else
    .blnOrigin = True
  End If
  If chkGridLine.Value = 0 Then
    .blnGridLine = False
    Else
    .blnGridLine = True
  End If
  'X-range
  If Abs((Val(txtValueX1.Text) - Val(txtValueX0.Text))) >= 0 And txtValueX1.Text <> "" _
  And txtValueX0.Text <> "" Then
    .X0 = Val(txtValueX0.Text)
    .x1 = Val(txtValueX1.Text)
    Else
    .X0 = dXmin
    .x1 = dXmax
  End If
  'Y-range
  If Abs((Val(txtValueY1.Text) - Val(txtValueY0.Text))) >= 0 And txtValueY1.Text <> "" _
  And txtValueY0.Text <> "" Then
    .Y0 = Val(txtValueY0.Text)
    .y1 = Val(txtValueY1.Text)
    Else
    .Y0 = dYmin 'dPlot(LBound(dPlot, 1), 2)
    .y1 = dYmax  'dPlot(UBound(dPlot, 1), 2)
  End If
  'index start-X1
  If Val(txtX0.Text) >= LBound(dPlot, 3) And Val(txtX0.Text) <= UBound(dPlot, 3) _
  And (Val(txtX1.Text) - Val(txtX0.Text)) > 0 Then
    .lStart = Val(txtX0.Text)
    Else
    .lStart = LBound(dPlot, 3)
    txtX0.Text = Str(LBound(dPlot, 3))
    txtX1.Text = Str(UBound(dPlot, 3))
  End If
  'index end-X 'set to array bound
  If Val(txtX1.Text) >= LBound(dPlot, 3) And Val(txtX1.Text) <= UBound(dPlot, 3) _
  And (Val(txtX1.Text) - Val(txtX0.Text)) > 0 Then
    .lEnd = Val(txtX1.Text)
    Else
    .lEnd = UBound(dPlot, 3)
    txtX0.Text = Str(LBound(dPlot, 3))
    txtX1.Text = Str(UBound(dPlot, 3))
  End If
  .asX = 0
  
  'check number of files to plot
  nSum = numFilesToPlot% - 1 'nSum files to plot
  ReDim .asY(nSum)
  ReDim .DrawTrace(flxlstFiles.list.Count - 1)
  
   nSum = 0
   For N = 0 To flxlstFiles.list.Count - 1
      If flxlstFiles.list.item(N + 1).Selected Then 'lstFiles.Selected(n) Then
            Select Case Val(PlotInfo(1, N))
               Case 1
                 .DrawTrace(nSum) = AS_POINT
               Case 0
                 .DrawTrace(nSum) = AS_CONLINE
               Case 2
                 .DrawTrace(nSum) = AS_BAR
               Case 3
                 .DrawTrace(nSum) = AS_DASH
               Case 4
                 .DrawTrace(nSum) = AS_DOT
               Case 5
                 .DrawTrace(nSum) = AS_DASHDOT
               Case 6
                 .DrawTrace(nSum) = AS_DASHDOTDOT
               Case 7
                 .DrawTrace(nSum) = AS_CIRCLE
               Case 8
                 .DrawTrace(nSum) = AS_FILLEDCIRCLE
             End Select
            If N + 1 <= flxlstFiles.list.Count Then
               .asY(nSum) = N + 1 'plot trace checked (=dplot(1,n))
                nSum = nSum + 1
                End If
            End If
  Next N
End With

   On Error GoTo 0
   Exit Sub

DefineLayout_Error:
    If Err.Number = 9 Then Resume Next
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DefineLayout of Form frmSetCond"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   Unload Me
   Set frmSetCond = Nothing
   Unload frmDraw
   Set frmDraw = Nothing
   Unload frmShowValues
   Set frmShowValues = Nothing
   Unload JKHplot
   Set JKHplot = Nothing
   Unload frmSpline
   Set frmSpline = Nothing
   Unload frmStat
   Set frmStat = Nothing
   End
   
   'unload all the opened forms
'   Dim I%
'   For I% = 0 To Forms.Count - 1
'      Unload Forms(I%)
'   Next I%
'   End
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim filplt%
    'record plot info
    'record this information
    filplt% = FreeFile
    Open App.Path & "\PlotDirec.txt" For Output As #filplt%
    Write #filplt%, "This file is used by Plot. Don't erase it!"
    Write #filplt%, direct$, directPlot$, dirWordpad
    Write #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
    Close #filplt%
            
   Unload Me
   Set frmSetCond = Nothing
End Sub

Private Sub mnuAbout_Click()
   frmAbout.Visible = True
End Sub

Private Sub mnuExit_Click()

Unload frmDraw
Unload frmShowValues
Unload Me
End Sub



Private Sub mnuFormat_Click()
   FileFormatfm.Visible = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuOpen_Click
' Author    : Dr-John-K-Hall
' Date      : 5/4/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOpen_Click()

'This routine replaces the VB CommonDialog control with a
'MultiSelect GetOpenFileName Common Dialog API
'(much of this code is from "www.mvps.org/vbnet/code/")

   Dim sFilters As String
   
   Dim pos As Long
   Dim buff As String
   Dim sLongname As String
   Dim sShortname As String
   
   'string of filters for the dialog box
   On Error GoTo mnuOpen_Click_Error

   sFilters = "Csv documents (*.csv) " & vbNullChar & "*.csv" & vbNullChar & _
              "Text documents (*.txt)" & vbNullChar & "*.txt" & vbNullChar & _
              "Rel files (*.rel)" & vbNullChar & "*.rel" & vbNullChar & _
              "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
              
   With OFN
      'size of the OFN structure
      .nStructSize = Len(OFN)
      
       'window owning the dialog
      .hWndOwner = frmSetCond.hWnd
      
      'filters (patterns) for the dropdown combo
      .sFilter = sFilters
      
      'index to the default filter
      .nFilterIndex = 4
      
      'default filename, plus additional padding
      'for the user's final selection(s).  Must be
      'double-null terminated
'      .sFile = "*.*" & Space$(2048) & vbNullChar & vbNullChar
      .sFile = "*.*" & Space$(4096) & vbNullChar & vbNullChar
      
      'the size of the buffer
      .nMaxFile = Len(.sFile)
      
      'default extension applied to
      'file if it has no extension
      .sDefFileExt = "bas" & vbNullChar & vbNullChar
      
      'space fot he file title if a single selection
      'made, double-null terminated, and its size
      .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      
      'starting folder, double-null terminated
      .nMaxTitle = Len(OFN.sFileTitle)
      
      'the dialog title
      If directPlot$ = "" Then directPlot$ = CurDir
      .sInitialDir = directPlot$ & vbNullChar & vbNullChar
      
      'default open flags and multiselect
      .sDialogTitle = "(Multi)select file(s) for plotting"
      .flags = OFS_FILE_OPEN_FLAGS Or _
             OFN_ALLOWMULTISELECT
   End With
   
   If GetOpenFileName(OFN) Then
      'remove trailing pair of termnating nulls
      'and trim returned file string
      buff = Trim$(Left$(OFN.sFile, Len(OFN.sFile) - 2))
      
      'Show the members of the returned sFile string
      'It path is larger than 3 characters, only show
      'most inner path (but record the rest in the file buffer)
      Dim sTemp As String, sPath As String, sShortPath As String
      Dim I%, numlist%, pos1%, MultiSelectPath As Boolean, sPath0 As String
      Dim sDriveLetter As String, MaxDirLen As Integer
      numlist% = 0
      MaxDirLen = Int(flxlstFiles.Width / 70) - 30
      
      sPath0 = Left$(OFN.sFile, OFN.nFileOffset)
      
      If InStr(sPath0, vbNullChar) <> 0 Then
         'this is multiselect path without final "\"
         'first trim off vbnullchar
         sPath = TrimNull(sPath0)
         If Mid$(sPath, Len(sPath), 1) <> "\" Then sPath = sPath & "\"
         MultiSelectPath = True
      Else
         'this is single path with final "\"
         sPath = sPath0
         End If
         
      'record this plot information
      'if has final "\" then remove it
      directPlot$ = sPath
      If Mid$(directPlot$, Len(directPlot$), 1) = "\" Then
         directPlot$ = Mid$(directPlot$, 1, Len(directPlot$) - 1)
         End If
      Dim filplt%
      filplt% = FreeFile
      Open App.Path & "\PlotDirec.txt" For Output As #filplt%
      Write #filplt%, "This file is used by Plot. Don't erase it!"
      Write #filplt%, direct$, directPlot$, dirWordpad
      Write #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
      Close #filplt%
      
      
      'determine short form of this path consisting of
      'the innermost directory
      Call ShortPath(sPath, MaxDirLen, sShortPath)
      
      Do While Len(buff) > 3
         sTemp = StripDelimitedItem(buff, vbNullChar)
         If MultiSelectPath Then
            If Mid$(sTemp, 1, Len(sPath0)) <> sPath0 Then
                numfiles% = numfiles% + 1
                
                'redimension plotinfo array
                ReDim Preserve PlotInfo(9, numfiles%)
                
                'list everything but the directory path
                'lstFiles.AddItem sShortPath & sTemp
                flxlstFiles.AddItem sShortPath & sTemp
                flxlstFiles.Refresh
                
                ReDim Preserve Files(numfiles%)
                Files(numfiles% - 1) = sPath & "\" & TrimNull(sTemp)
                End If
         Else
            'don't repeat the directory
             numfiles% = numfiles% + 1
             
             'redimension plotinfo array
             ReDim Preserve PlotInfo(9, numfiles%)

             'lstFiles.AddItem sShortPath & Mid$(sTemp, Len(sPath) + 2, Len(sTemp) - Len(sPath) - 1)
             flxlstFiles.AddItem sShortPath & Mid$(sTemp, Len(sPath) + 2, Len(sTemp) - Len(sPath) - 1)
             flxlstFiles.list.item(numfiles%).ToolTipText = Mid$(sTemp, Len(sPath) + 2, Len(sTemp) - Len(sPath) - 1)
             flxlstFiles.Refresh
             ReDim Preserve Files(numfiles%)
             Files(numfiles% - 1) = TrimNull(sTemp)
             End If
      Loop
   End If
     
   If numfiles% > 0 Then
      mnuSave.Enabled = True
      mnuSpline.Enabled = True
      cmdShowEdit.Enabled = True
      cmdWizard.Enabled = True
      cmdAll.Enabled = True
      cmdClear.Enabled = True
      
        chkOrigin.Enabled = True
        chkGridLine.Enabled = True
        txtX0.Enabled = True
        txtX1.Enabled = True
        txtValueY0.Enabled = True
        txtValueY1.Enabled = True
        txtXTitle.Enabled = True
        txtYTitle.Enabled = True
        txtTitle.Enabled = True
        txtValueX0.Enabled = True
        txtValueX1.Enabled = True
        fraLayout.Enabled = True
        frmSetCond.lblEndX1.Enabled = True
        frmSetCond.lblEndY.Enabled = True
        frmSetCond.lblGridLine.Enabled = True
        frmSetCond.lblIndexEnd.Enabled = True
        frmSetCond.lblOrigin.Enabled = True
        frmSetCond.lblIndexStart.Enabled = True
        frmSetCond.lblStartX0.Enabled = True
        frmSetCond.lblStartY.Enabled = True
        frmSetCond.lblXTitle.Enabled = True
        frmSetCond.lblYTitle.Enabled = True
        frmSetCond.lblTitle.Enabled = True
        
   Else
      Dim ier As Integer
'      ier = MsgBox("Simultaneous file selection failed!" & vbCrLf & vbCrLf & _
'            "(Hint: Try again by selecting fewer files.)", vbOKOnly + vbCritical, "File open error")
      
      End If

   On Error GoTo 0
   Exit Sub

mnuOpen_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuOpen_Click of Form frmSetCond"

End Sub

Sub ReadValues()
'this sub does the actual data input
'it finds also the plot limits

Dim nOffSetIndex As Integer
Dim nOffSetValue As Integer
Dim numFilesToPlot As Integer
Dim nSign As Integer

'open the highlighted files and input their data
'according to the stored plot formats

Dim I%, freefil%
   On Error GoTo ReadValues_Error

numFilesToPlot = 0
For I% = 0 To flxlstFiles.list.Count - 1
   If flxlstFiles.list.item(I% + 1).Selected Then
      freefil% = FreeFile
      If Dir(PlotInfo(7, I%)) = "" Then
         MsgBox "Can't find plot file: " & PlotInfo(7, I%), vbCritical + vbOKOnly
      Else
         numFilesToPlot = numFilesToPlot + 1
         Call OpenRead(I%, numFilesToPlot)
         End If
      End If
Next I%

'determine Ymin and Ymax
Dim numTried%
numTried% = 0
For I% = 0 To flxlstFiles.list.Count - 1
   If flxlstFiles.list.item(I% + 1).Selected Then

        If numTried% = 0 Then
           dYmax = dPlot(I%, 1, 0)
           dYmin = dPlot(I%, 1, 0)
           dXmax = dPlot(I%, 0, 0)
           dXmin = dPlot(I%, 0, 0)
           numTried% = 1
           End If
        For N = 0 To UBound(dPlot, 3)
           If dYmax < dPlot(I%, 1, N) Then
             If dPlot(I%, 1, N) = 0 And dPlot(I%, 0, N) = 0 Then
             Else
                dYmax = dPlot(I%, 1, N)
                End If
           End If
           If dYmin > dPlot(I%, 1, N) Then
             If dPlot(I%, 1, N) = 0 And dPlot(I%, 0, N) = 0 Then
             Else
                dYmin = dPlot(I%, 1, N)
                End If
           End If
        Next N
          
        'determine Xmin and Xmax
        For N = 0 To UBound(dPlot, 3)
           If dXmax < dPlot(I%, 0, N) Then
              If dPlot(I%, 1, N) = 0 And dPlot(I%, 0, N) = 0 Then
              Else
                 dXmax = dPlot(I%, 0, N)
                 End If
           End If
           If dXmin > dPlot(I%, 0, N) Then
              If dPlot(I%, 1, N) = 0 And dPlot(I%, 0, N) = 0 Then
              Else
                 dXmin = dPlot(I%, 0, N)
                 End If
           End If
        Next N
        
     End If
Next I%
  
YMin0 = dYmin
YRange0 = dYmax
YMin0 = YMin0 - Abs(dYmax - dYmin) / 20
YRange0 = YRange0 + Abs(dYmax - dYmin) / 20

XMin0 = dXmin
XRange0 = dXmax
XMin0 = XMin0 - Abs(dXmax - dXmin) / 20
XRange0 = XRange0 + Abs(dXmax - dXmin) / 20

frmSetCond.txtValueY0 = YMin0
frmSetCond.txtValueY1 = YRange0
frmSetCond.txtValueX0 = XMin0
frmSetCond.txtValueX1 = XRange0

   On Error GoTo 0
   Exit Sub

ReadValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadValues of Form frmSetCond"

End Sub

Private Sub mnuOpen1_Click()
   mnuOpen_Click
End Sub

Private Sub mnuPaths_Click()
   'find path to Wordpad.exe
   
   On Error GoTo errhand
   
10 CommonDialog1.CancelError = True
   If dirWordpad = sEmpty Then dirWordpad = "c:\Progra~1\Accessories\"
   CommonDialog1.FileName = dirWordpad & "Wordpad.exe"
   CommonDialog1.Filter = "Wordpad.exe (.exe)|*.exe|All files (*.*)|*.*"
   CommonDialog1.ShowOpen
   dirWordpad = Mid$(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - 11)
   If Dir(dirWordpad & "Wordpad.exe") = sEmpty Then
      MsgBox "Wordpad not found at the chosen path!", _
              vbOKCancel + vbExclamation, "Plot Error", "Plot Error"
      GoTo 10
      End If
      
'record path
      Dim filplt%
      filplt% = FreeFile
      Open App.Path & "\PlotDirec.txt" For Output As #filplt%
      Write #filplt%, "This file is used by Plot. Don't erase it!"
      Write #filplt%, direct$, directPlot$, dirWordpad
      Write #filplt%, FitMethod%, FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg, PolyDeg, SplineType%, SplineDeg%
      Close #filplt%

errhand:
End Sub

Private Sub mnuReadme_Click()
   'present help as a wordpad readme file
   Dim helpnam$, ret As Long
   helpnam$ = App.Path & "\PlotReadme.txt"
   If Dir(helpnam$) <> sEmpty And Dir(dirWordpad & "Wordpad.exe") <> sEmpty Then
      ret = Shell(dirWordpad & "Wordpad.exe " & Chr$(34) & helpnam$ & Chr$(34), vbNormalFocus)
      End If
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuRestore_Click
' Author    : Dr-John-K-Hall
' Date      : 11/26/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuRestore_Click()
    'restore plot info from previously saved file list
    Dim NewDoclin$
       
   On Error GoTo mnuRestore_Click_Error

    If Dir(App.Path & "\PlotFiles.txt") = sEmpty Then
       MsgBox "Can't find the PlotFiles.txt file!", vbCritical + vbOKOnly, "Plot Error"
       mnuRestore.Enabled = False
       mnuRestoreOther.Enabled = False
       Exit Sub
       End If
    
    Dim MaxDirLen As Integer, sShortPath As String
    Dim filplt%, doclin$, FileOut As String, found%
    Dim sPath As String, J%, RootName$, SplitDoc() As String
    
    If numfiles% <> 0 Then 'clear present plot and reclaim memory
       cmdClear_Click
       End If
       
    
    filplt% = FreeFile
    numfiles% = 0
    Open App.Path & "\PlotFiles.txt" For Input As #filplt%
    Line Input #filplt%, doclin$
    SplitDoc = Split(doclin$, ",")
    If UBound(SplitDoc) = 1 Then
       'new format that records font size of axis labels
       txtAxisLabelSize.Text = SplitDoc(1)
       End If
    Line Input #filplt%, NewDoclin$
    If InStr(NewDoclin$, ",") = 0 Then
       'old format without format size recorded
       XTitle$ = NewDoclin$
       Input #filplt%, YTitle$
       Input #filplt%, Title$
    Else
       Dim TitlesInput() As String
       TitlesInput = Split(NewDoclin$, ",")
       XTitle$ = Mid$(TitlesInput(0), 2, Len(TitlesInput(0)) - 2) 'remove enclosing quotation marks
       txtTitleXfont.Text = TitlesInput(1)
       Line Input #filplt%, NewDoclin$
       TitlesInput = Split(NewDoclin$, ",")
       YTitle$ = Mid$(TitlesInput(0), 2, Len(TitlesInput(0)) - 2)
       txtTitleYfont.Text = TitlesInput(1)
       Line Input #filplt%, NewDoclin$
       TitlesInput = Split(NewDoclin$, ",")
       Title$ = Mid$(TitlesInput(0), 2, Len(TitlesInput(0)) - 2)
       txtTitlefont.Text = TitlesInput(1)
       txtTitlefont.Text = TitlesInput(1)
       
       Line Input #filplt%, NewDoclin$
       txtX0.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtX1.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueY0.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueY1.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueX0.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueX1.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       End If

        
    Do Until EOF(filplt%)
       numfiles% = numfiles% + 1
       ReDim Preserve PlotInfo(9, numfiles%)
       Input #filplt%, PlotInfo(0, numfiles% - 1), PlotInfo(1, numfiles% - 1), _
                     PlotInfo(2, numfiles% - 1), PlotInfo(3, numfiles% - 1), _
                     PlotInfo(4, numfiles% - 1), PlotInfo(5, numfiles% - 1), _
                     PlotInfo(6, numfiles% - 1), PlotInfo(7, numfiles% - 1), _
                     PlotInfo(8, numfiles% - 1), PlotInfo(9, numfiles% - 1)
       ReDim Preserve Files(numfiles%)
       Files(numfiles% - 1) = PlotInfo(7, numfiles% - 1)
       If PlotInfo(0, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(0, numfiles% - 1) = 0
          End If
       If PlotInfo(1, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(1, numfiles% - 1) = 0
          End If
       If PlotInfo(2, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(2, numfiles% - 1) = 0
          End If
       If PlotInfo(3, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(3, numfiles% - 1) = 0
          End If
       If PlotInfo(4, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(4, numfiles% - 1) = 0
          End If
       If PlotInfo(5, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(5, numfiles% - 1) = 0
          End If
       If PlotInfo(6, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(6, numfiles% - 1) = 0
          End If
       If PlotInfo(9, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(9, numfiles% - 1) = "1"
          End If
       'display file name in flxlstFiles list box
       'find short path name

       MaxDirLen = Int(frmSetCond.flxlstFiles.Width / 70) - 30
       'find path
       FileOut = Files(numfiles% - 1)
       found% = 0
       For J% = Len(FileOut) To 1 Step -1
          If Mid$(FileOut, J%, 1) = "\" Then
             RootName$ = Mid$(FileOut, J% + 1, Len(FileOut) - J%)
             sPath = Mid$(FileOut, 1, Len(FileOut) - Len(RootName$))
             found% = 1
             Exit For
             End If
       Next J%
       If found% = 1 Then
          'shorten this name to fit into plot buffer List Box
          Call ShortPath(sPath, MaxDirLen, sShortPath)
       Else
          sPath = ""
          End If
        
       'add the file to plot buffer
        'frmSetCond.lstFiles.AddItem sShortPath & RootName$
        frmSetCond.flxlstFiles.AddItem sShortPath & RootName$
        flxlstFiles.Refresh
    
    Loop
    Close #filplt%
    
    mnuSave.Enabled = True
    mnuSpline.Enabled = True
    cmdShowEdit.Enabled = True
    cmdWizard.Enabled = True
    cmdAll.Enabled = True
    cmdClear.Enabled = True
    
    chkOrigin.Enabled = True
    chkGridLine.Enabled = True
    txtX0.Enabled = True
    txtX1.Enabled = True
    txtValueY0.Enabled = True
    txtValueY1.Enabled = True
    txtXTitle.Enabled = True
    txtYTitle.Enabled = True
    txtTitle.Enabled = True
    txtValueX0.Enabled = True
    txtValueX1.Enabled = True
    fraLayout.Enabled = True
    frmSetCond.lblEndX1.Enabled = True
    frmSetCond.lblEndY.Enabled = True
    frmSetCond.lblGridLine.Enabled = True
    frmSetCond.lblIndexEnd.Enabled = True
    frmSetCond.lblOrigin.Enabled = True
    frmSetCond.lblIndexStart.Enabled = True
    frmSetCond.lblStartX0.Enabled = True
    frmSetCond.lblStartY.Enabled = True
    frmSetCond.lblXTitle.Enabled = True
    frmSetCond.lblYTitle.Enabled = True
    
    frmSetCond.txtXTitle = XTitle$
    frmSetCond.txtYTitle = YTitle$
    frmSetCond.txtTitle = Title$

   On Error GoTo 0
   Exit Sub

mnuRestore_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuRestore_Click of Form frmSetCond"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuRestoreOther_Click
' Author    : Dr-John-K-Hall
' Date      : 5/4/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuRestoreOther_Click()

   On Error GoTo mnuRestoreOther_Click_Error
   
   Dim FileInPlotFiles$, NewDoclin$, SplitDoc() As String
   
   'pick other restore file
   CommonDialog1.CancelError = True
   CommonDialog1.Filter = "text files (*.txt)|*.txt|All fiels (*.*)"
   CommonDialog1.FileName = "*.txt"
   CommonDialog1.ShowOpen
   FileInPlotFiles$ = CommonDialog1.FileName
   

    If Dir(FileInPlotFiles$) = sEmpty Then
       MsgBox "Can't find the " & FileInPlotFiles$ & " file!", vbCritical + vbOKOnly, "Plot Error"
       mnuRestore.Enabled = False
       mnuRestoreOther.Enabled = False
       Exit Sub
       End If
    
    Dim MaxDirLen As Integer, sShortPath As String
    Dim filplt%, doclin$, FileOut As String, found%
    Dim sPath As String, J%, RootName$
    
    If numfiles% <> 0 Then 'clear present plot and reclaim memory
       cmdClear_Click
       End If
       
    
    filplt% = FreeFile
    numfiles% = 0
    Open FileInPlotFiles$ For Input As #filplt%
    Line Input #filplt%, doclin$
    SplitDoc = Split(doclin$, ",")
    If UBound(SplitDoc) = 1 Then
       'new format that records font size of axis labels
       txtAxisLabelSize.Text = SplitDoc(1)
       End If
    'see if new format with value of font size
    Line Input #filplt%, NewDoclin$
    If InStr(NewDoclin$, ",") = 0 Then
       'old format without format size recorded
       XTitle$ = NewDoclin$
       Input #filplt%, YTitle$
       Input #filplt%, Title$
    Else
       Dim TitlesInput() As String
       TitlesInput = Split(NewDoclin$, ",")
       XTitle$ = Mid$(TitlesInput(0), 2, Len(TitlesInput(0)) - 2) 'remove enclosing quotation marks
       txtTitleXfont.Text = TitlesInput(1)
       Line Input #filplt%, NewDoclin$
       TitlesInput = Split(NewDoclin$, ",")
       YTitle$ = Mid$(TitlesInput(0), 2, Len(TitlesInput(0)) - 2)
       txtTitleYfont.Text = TitlesInput(1)
       Line Input #filplt%, NewDoclin$
       TitlesInput = Split(NewDoclin$, ",")
       Title$ = Mid$(TitlesInput(0), 2, Len(TitlesInput(0)) - 2)
       txtTitlefont.Text = TitlesInput(1)
       
       Line Input #filplt%, NewDoclin$
       txtX0.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtX1.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueY0.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueY1.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueX0.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       Line Input #filplt%, NewDoclin$
       txtValueX1.Text = Mid$(Trim$(NewDoclin$), 2, Len(Trim$(NewDoclin$)) - 2)
       End If

    Do Until EOF(filplt%)
       numfiles% = numfiles% + 1
       ReDim Preserve PlotInfo(9, numfiles%)
       Input #filplt%, PlotInfo(0, numfiles% - 1), PlotInfo(1, numfiles% - 1), _
                     PlotInfo(2, numfiles% - 1), PlotInfo(3, numfiles% - 1), _
                     PlotInfo(4, numfiles% - 1), PlotInfo(5, numfiles% - 1), _
                     PlotInfo(6, numfiles% - 1), PlotInfo(7, numfiles% - 1), _
                     PlotInfo(8, numfiles% - 1), PlotInfo(9, numfiles% - 1)
       ReDim Preserve Files(numfiles%)
       Files(numfiles% - 1) = PlotInfo(7, numfiles% - 1)
       If PlotInfo(0, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(0, numfiles% - 1) = 0
          End If
       If PlotInfo(1, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(1, numfiles% - 1) = 0
          End If
       If PlotInfo(2, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(2, numfiles% - 1) = 0
          End If
       If PlotInfo(3, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(3, numfiles% - 1) = 0
          End If
       If PlotInfo(4, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(4, numfiles% - 1) = 0
          End If
       If PlotInfo(5, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(5, numfiles% - 1) = 0
          End If
       If PlotInfo(6, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(6, numfiles% - 1) = 0
          End If
       If PlotInfo(9, numfiles% - 1) = "" Then 'no plot info ever entered so set default
          PlotInfo(9, numfiles% - 1) = "1"
          End If
       'display file name in flxlstFiles list box
       'find short path name

       MaxDirLen = Int(frmSetCond.flxlstFiles.Width / 70) - 30
       'find path
       FileOut = Files(numfiles% - 1)
       found% = 0
       For J% = Len(FileOut) To 1 Step -1
          If Mid$(FileOut, J%, 1) = "\" Then
             RootName$ = Mid$(FileOut, J% + 1, Len(FileOut) - J%)
             sPath = Mid$(FileOut, 1, Len(FileOut) - Len(RootName$))
             found% = 1
             Exit For
             End If
       Next J%
       If found% = 1 Then
          'shorten this name to fit into plot buffer List Box
          Call ShortPath(sPath, MaxDirLen, sShortPath)
       Else
          sPath = ""
          End If
        
       'add the file to plot buffer
        'frmSetCond.lstFiles.AddItem sShortPath & RootName$
        frmSetCond.flxlstFiles.AddItem sShortPath & RootName$
        flxlstFiles.Refresh
    
    Loop
    Close #filplt%
    
    mnuSave.Enabled = True
    mnuSpline.Enabled = True
    cmdShowEdit.Enabled = True
    cmdWizard.Enabled = True
    cmdAll.Enabled = True
    cmdClear.Enabled = True
    
    chkOrigin.Enabled = True
    chkGridLine.Enabled = True
    txtX0.Enabled = True
    txtX1.Enabled = True
    txtValueY0.Enabled = True
    txtValueY1.Enabled = True
    txtXTitle.Enabled = True
    txtYTitle.Enabled = True
    txtTitle.Enabled = True
    txtValueX0.Enabled = True
    txtValueX1.Enabled = True
    fraLayout.Enabled = True
    frmSetCond.lblEndX1.Enabled = True
    frmSetCond.lblEndY.Enabled = True
    frmSetCond.lblGridLine.Enabled = True
    frmSetCond.lblIndexEnd.Enabled = True
    frmSetCond.lblOrigin.Enabled = True
    frmSetCond.lblIndexStart.Enabled = True
    frmSetCond.lblStartX0.Enabled = True
    frmSetCond.lblStartY.Enabled = True
    frmSetCond.lblXTitle.Enabled = True
    frmSetCond.lblYTitle.Enabled = True
    
    frmSetCond.txtXTitle = XTitle$
    frmSetCond.txtYTitle = YTitle$
    frmSetCond.txtTitle = Title$

   On Error GoTo 0
   Exit Sub

mnuRestoreOther_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuRestoreOther_Click of Form frmSetCond"
End Sub

Private Sub mnuSave_Click()
   'save the plot information for future retrieval
   Dim filplt%, I%
   
   With frmSetCond
      XTitle$ = Trim$(.txtXTitle.Text)
      YTitle$ = Trim$(.txtYTitle.Text)
      Title$ = Trim$(.txtTitle.Text)
   End With
   
   filplt% = FreeFile
   Open App.Path & "\PlotFiles.txt" For Output As #filplt%
   Print #filplt%, "This file is used by Plot. Don't erase it!" & "," & txtAxisLabelSize.Text
   Write #filplt%, XTitle$, Val(txtTitleXfont.Text)
   Write #filplt%, YTitle$, Val(txtTitleYfont.Text)
   Write #filplt%, Title$, Val(txtTitlefont.Text)
   
   'now store xmin,xmax,ymin,ymax,etc.
   Write #filplt%, txtX0.Text
   Write #filplt%, txtX1.Text
   Write #filplt%, txtValueY0.Text
   Write #filplt%, txtValueY1.Text
   Write #filplt%, txtValueX0.Text
   Write #filplt%, txtValueX1.Text
   
   'now store file names and plot info for the plot buffer
   For I% = 0 To numfiles% - 1
      Write #filplt%, PlotInfo(0, I%), PlotInfo(1, I%), _
                     PlotInfo(2, I%), PlotInfo(3, I%), _
                     PlotInfo(4, I%), PlotInfo(5, I%), _
                     PlotInfo(6, I%), Files(I%), PlotInfo(8, I%), PlotInfo(9, I%)
   Next I%
   
   Close #filplt%
   
   mnuRestore.Enabled = True
   mnuRestoreOther.Enabled = True
   
End Sub

Private Sub mnuSpline_Click()
   frmSpline.Visible = True
End Sub

Private Sub mnuStatistics_Click()
    frmStat.Visible = True
End Sub
