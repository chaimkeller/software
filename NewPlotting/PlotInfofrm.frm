VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form PlotInfofrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plot Information"
   ClientHeight    =   4410
   ClientLeft      =   4560
   ClientTop       =   2325
   ClientWidth     =   6300
   Icon            =   "PlotInfofrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6300
   Begin VB.CheckBox chkSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   5400
      TabIndex        =   56
      ToolTipText     =   "Save the format pattern for the next file in the list"
      Top             =   4080
      Width           =   735
   End
   Begin MSComCtl2.UpDown updwnLineWidth 
      Height          =   285
      Left            =   720
      TabIndex        =   52
      Top             =   4040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtLineWidth"
      BuddyDispid     =   196610
      OrigLeft        =   960
      OrigTop         =   4080
      OrigRight       =   1215
      OrigBottom      =   4335
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtLineWidth 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   51
      Text            =   "1"
      ToolTipText     =   "Line width of line, thickness of circle"
      Top             =   4040
      Width           =   480
   End
   Begin VB.Frame frmWrapper 
      Height          =   520
      Left            =   4140
      TabIndex        =   48
      Top             =   2980
      Width           =   2055
      Begin VB.ComboBox cmbfuncY 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1050
         TabIndex        =   50
         Text            =   "func(y)"
         ToolTipText     =   "function(y)"
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cmbfuncX 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         TabIndex        =   49
         Text            =   "func(x)"
         ToolTipText     =   "function(x)"
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Frame frmReverse 
      Height          =   480
      Left            =   4140
      TabIndex        =   46
      Top             =   2550
      Width           =   2055
      Begin VB.CheckBox chkReverse 
         Caption         =   "Reverse X order"
         Height          =   255
         Left            =   300
         TabIndex        =   47
         ToolTipText     =   "Detect reversed x order, reverse, save"
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Height          =   460
      Left            =   4140
      TabIndex        =   37
      ToolTipText     =   "Conversion Engine"
      Top             =   3480
      Width           =   2055
      Begin VB.CheckBox chkJKH 
         Caption         =   "JKH Data?"
         Height          =   285
         Left            =   480
         TabIndex        =   38
         Top             =   140
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tweaks"
      Height          =   2080
      Left            =   4140
      TabIndex        =   32
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtYB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         TabIndex        =   27
         Text            =   "0.0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtYA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         TabIndex        =   26
         Text            =   "1.0"
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtXB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         TabIndex        =   25
         Text            =   "0.0"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtXA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         TabIndex        =   24
         Text            =   "1.0"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Y+B, B="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   1740
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Y/A, A ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   795
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1980
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "X+B, B="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "X/A, A ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   4020
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   315
      Left            =   2340
      TabIndex        =   0
      Top             =   4020
      Width           =   1395
   End
   Begin VB.Frame Frame3 
      Caption         =   "Colors"
      Height          =   3460
      Left            =   2520
      TabIndex        =   31
      Top             =   480
      Width           =   1515
      Begin VB.OptionButton Option9 
         Caption         =   "&Light Blue"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   2940
         Width           =   1155
      End
      Begin VB.OptionButton Option8 
         Caption         =   "&Gray"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   2640
         Width           =   795
      End
      Begin VB.OptionButton Option7 
         Caption         =   "&Yellow"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   2340
         Width           =   795
      End
      Begin VB.OptionButton Option6 
         Caption         =   "&Magneta"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton optRed 
         Caption         =   "&Red"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1740
         Width           =   795
      End
      Begin VB.OptionButton optCyan 
         Caption         =   "&Cyan"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1440
         Width           =   795
      End
      Begin VB.OptionButton optGreen 
         Caption         =   "&Green"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1140
         Width           =   795
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "&Blue"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.OptionButton optBlack 
         Caption         =   "&Black"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   540
         Width           =   795
      End
      Begin VB.OptionButton optAutomatic 
         Caption         =   "&Automatic"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H00808080&
      Height          =   375
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   6195
      TabIndex        =   30
      Top             =   0
      Width           =   6255
      Begin VB.Label lblFileName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   60
         TabIndex        =   39
         Top             =   30
         Width           =   6075
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Plot Marks"
      Height          =   3460
      Left            =   1320
      TabIndex        =   29
      Top             =   480
      Width           =   1095
      Begin VB.OptionButton optFilledCircle 
         Caption         =   "Filled Circle"
         Height          =   360
         Left            =   180
         TabIndex        =   55
         ToolTipText     =   "Filled circles"
         Top             =   3020
         Width           =   735
      End
      Begin VB.OptionButton optCircle 
         Caption         =   "Circle"
         Height          =   375
         Left            =   180
         TabIndex        =   54
         ToolTipText     =   "Circles"
         Top             =   2620
         Width           =   735
      End
      Begin VB.OptionButton optPoint 
         Caption         =   "&Point"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   500
         Width           =   735
      End
      Begin VB.OptionButton optLine 
         Caption         =   "&Line"
         Height          =   315
         Left            =   180
         TabIndex        =   44
         Top             =   180
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optDashDotDot 
         Caption         =   "Dash Dot Dot"
         Height          =   675
         Left            =   180
         TabIndex        =   43
         Top             =   2040
         Width           =   675
      End
      Begin VB.OptionButton optDashDot 
         Caption         =   "Dash Dot"
         Height          =   375
         Left            =   180
         TabIndex        =   42
         Top             =   1680
         Width           =   795
      End
      Begin VB.OptionButton optDot 
         Caption         =   "Dot"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   1400
         Width           =   675
      End
      Begin VB.OptionButton optDash 
         Caption         =   "Dash"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1100
         Width           =   675
      End
      Begin VB.OptionButton optBar 
         Caption         =   "&Bar"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   800
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Formats"
      Height          =   3460
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   1095
      Begin VB.OptionButton optFF11 
         Caption         =   "#11"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   3180
         Width           =   735
      End
      Begin VB.OptionButton optFF10 
         Caption         =   "#10"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   735
      End
      Begin VB.OptionButton optFF9 
         Caption         =   "#9"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2580
         Width           =   735
      End
      Begin VB.OptionButton optFF8 
         Caption         =   "#8"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton optFF7 
         Caption         =   "#7"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1980
         Width           =   735
      End
      Begin VB.OptionButton optFF6 
         Caption         =   "#6"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optFF5 
         Caption         =   "#5"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1380
         Width           =   735
      End
      Begin VB.OptionButton optFF4 
         Caption         =   "#4"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton optFF3 
         Caption         =   "#3"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   735
      End
      Begin VB.OptionButton optFF2 
         Caption         =   "#2"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optFF1 
         Caption         =   "#1"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Label lblLineWidth 
      Caption         =   "Line Width"
      Height          =   255
      Left            =   1080
      TabIndex        =   53
      Top             =   4080
      Width           =   855
   End
End
Attribute VB_Name = "PlotInfofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public chkFF%, chkPlotType%, chkColor%

Private Sub chkJKH_Click()
   'show convert program
   JKHplot.Visible = True
   JKHplot.cmdConvert.Caption = JKHplot.cmdConvert.Caption & " " & PlotInfofrm.lblFileName.Caption
End Sub

Private Sub chkSave_Click()
  If chkSave.Value = vbChecked And Not SaveFormat Then
    SaveFormat = True
    LastSelected% = numSelected%
  ElseIf chkSave.Value = vbUnchecked Then
    SaveFormat = False
    End If
End Sub

Private Sub cmdAccept_Click()
   Dim FileLines() As String, FileReverseLines() As String
   Dim numLines As Long, doclin$
   
   'record inputed values and exit
   On Error GoTo cmdAccept_Click_Error

   PlotInfo(0, numSelected%) = Str$(chkFF%)
   PlotInfo(1, numSelected%) = Str$(chkPlotType%)
   PlotInfo(2, numSelected%) = Str$(chkColor%)
   PlotInfo(3, numSelected%) = txtXA
   If Val(PlotInfo(3, numSelected%)) = 0 Then
      PlotInfo(3, numSelected%) = "1.0"
      End If
   PlotInfo(4, numSelected%) = txtXB
   PlotInfo(5, numSelected%) = txtYA
   If Val(PlotInfo(5, numSelected%)) = 0 Then
      PlotInfo(5, numSelected%) = "1.0"
      End If
   PlotInfo(6, numSelected%) = txtYB
   PlotInfo(7, numSelected%) = PlotInfofrm.lblFileName
   PlotInfo(8, numSelected%) = cmbfuncX.Text & ":" & cmbfuncY.Text
   PlotInfo(9, numSelected%) = txtLineWidth.Text
   PlotInfoCancel = True
   
   '////////////////////additions 11/15/2019//////////////////////
   
   If chkReverse.Value = vbChecked Then
      'reverse the file's x order and save it under the original file number
      'first write backup
      
      tmp$ = PlotInfofrm.Caption
      PlotInfofrm.Caption = PlotInfofrm.Caption & " --> reversing order...please wati"
      Screen.MousePointer = vbHourglass
      
      If Dir(App.Path & "\temp.bak") <> sEmpty Then
         Kill App.Path & "\temp.bak"
         End If
         
      FileCopy Files(numSelected%), App.Path & "\temp.bak"
      backup% = 1
      
      filerev% = FreeFile
      Open Files(numSelected%) For Input As #filerev%
      
      numLines = 0
      Do Until EOF(filerev%)
         Line Input #filerev%, doclin$
         ReDim Preserve FileLines(numLines)
         FileLines(numLines) = doclin$
         numLines = numLines + 1
      Loop
      ReDim Preserve FileReverseLines(numLines - 1)
      For I% = O To numLines - 1
         FileReverseLines(numLines - I% - 1) = FileLines(I%)
      Next I%
      Close #filerev%
      
      'now replace in reversed x order
      filerev% = FreeFile
      Open Files(numSelected%) For Output As #filerev%
      backup% = 2
      
      For I% = 0 To numLines - 1
         Print #filerev%, FileReverseLines(I%)
      Next I%
      Close #filerev%
      
      'erase back up file
       If Dir(App.Path & "\tmp.bat") <> sEmpty Then
          Kill App.Path & "\tmp.bat"
          backup% = 0
          End If
     
      End If
      
      PlotInfofrm.Caption = tmp$
      Screen.MousePointer = vbDefault
      
   Unload Me

   On Error GoTo 0
   Exit Sub

cmdAccept_Click_Error:

    Close
    Screen.MousePointer = vbDefault
    
    If backup% = 2 And Dir(App.Path & "\tmp.bat") <> sEmpty Then
    
        Select Case MsgBox("Reversing failed." _
                           & vbCrLf & "" _
                           & vbCrLf & "Do you want to restore the file?" _
                           , vbYesNoCancel Or vbExclamation Or vbDefaultButton1, "Reverse x order")
        
          Case vbYes
          
             Kill Files(numSelected%)
             FileCopy App.Path & "\tmp.bat", Files(numSelected%)

        
          Case vbNo
        
          Case vbCancel
        
        End Select
        
    ElseIf backup% = 1 Then
    
       'simply erase the backup file
       If Dir(App.Path & "\tmp.bat") <> sEmpty Then
          Kill App.Path & "\tmp.bat"
          backup% = 0
          End If
       
       End If

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAccept_Click of Form PlotInfofrm"
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Dim FuncX As String, FuncY As String, numNewSelected%
   
   Me.Left = Screen.Width * 0.5
   Me.Top = Screen.Height * 0.5
   
   PlotInfoCancel = False 'flag that form is loading
                          'and no entries have been accepted
   'load up old information if it exists
   
   'if save format, then use it
   If SaveFormat Then
      chkSave.Value = vbChecked
      numNewSelected% = LastSelected%
      txtLineWidth.Text = PlotInfo(9, numNewSelected%)
   Else
      numNewSelected% = numSelected%
      End If
      
   PlotInfofrmVis = True
   chkFF% = Val(PlotInfo(0, numNewSelected%))
   Select Case chkFF%
      Case 0
         optFF1.Value = True
      Case 1
         optFF2.Value = True
      Case 2
         optFF3.Value = True
      Case 3
         optFF4.Value = True
      Case 4
         optFF5.Value = True
      Case 5
         optFF6.Value = True
      Case 6
         optFF7.Value = True
      Case 7
         optFF8.Value = True
      Case 8
         optFF9.Value = True
      Case 9
         optFF10.Value = True
      Case 10
         optFF11.Value = True
   End Select
   chkPlotType% = Val(PlotInfo(1, numNewSelected%))
   Select Case chkPlotType%
      Case 1
         optPoint.Value = True
      Case 0
         optLine.Value = True
      Case 2
         optBar.Value = True
      Case 3
         optDash.Value = True
      Case 4
         optDot.Value = True
      Case 5
         optDashDot.Value = True
      Case 6
         optDashDotDot.Value = True
      Case 7
         optCircle.Value = True
      Case 8
         optFilledCircle.Value = True
   End Select
   chkColor% = Val(PlotInfo(2, numNewSelected%))
   Select Case chkColor%
      Case 0
        optAutomatic.Value = True
      Case 1
        optBlack.Value = True
      Case 2
        optBlue.Value = True
      Case 3
        optGreen.Value = True
      Case 4
        optCyan.Value = True
      Case 5
        optRed.Value = True
      Case 6
        Option6.Value = True
      Case 7
        Option7.Value = True
      Case 8
        Option8.Value = True
      Case 9
        Option9.Value = True
   End Select
   If Val(PlotInfo(3, numNewSelected%)) <> 0 Then
      txtXA = PlotInfo(3, numNewSelected%)
   Else
      txtXA = 1#
      End If
   If Val(PlotInfo(4, numNewSelected%)) <> 0 Then
      txtXB = PlotInfo(4, numNewSelected%)
   Else
      txtXB = 0#
      End If
   If Val(PlotInfo(5, numNewSelected%)) <> 0 Then
      txtYA = PlotInfo(5, numNewSelected%)
   Else
      txtYA = 1#
      End If
   If Val(PlotInfo(6, numNewSelected%)) <> 0 Then
      txtYB = PlotInfo(6, numNewSelected%)
   Else
      txtYB = 0#
      End If
   If PlotInfo(7, numSelected%) <> "" Then
      PlotInfofrm.lblFileName = PlotInfo(7, numSelected%)
      End If
   If PlotInfo(9, numNewSelected%) <> "" Then
      txtLineWidth.Text = PlotInfo(9, numNewSelected%)
      End If
      
   
   With cmbfuncX
     .AddItem "none"
     .AddItem "log"
     .AddItem "exp"
     .AddItem "cos"
     .AddItem "sin"
     .AddItem "tan"
     .ListIndex = 0
   End With
   
   With cmbfuncY
     .AddItem "none"
     .AddItem "log"
     .AddItem "exp"
     .AddItem "cos"
     .AddItem "sin"
     .AddItem "tan"
     .ListIndex = 0
   End With
      
   If PlotInfo(8, numNewSelected%) <> "" Then
      pos% = InStr(PlotInfo(8, numNewSelected%), ":")
      If pos% > 0 Then
         FuncX = Mid$(PlotInfo(8, numNewSelected%), 1, pos% - 1)
         FuncY = Mid$(PlotInfo(8, numNewSelected%), pos% + 1, Len(PlotInfo(8, numNewSelected%)) - pos%)
         Select Case FuncX
            Case "none"
                cmbfuncX.ListIndex = 0
            Case "log"
                cmbfuncX.ListIndex = 1
            Case "exp"
                cmbfuncX.ListIndex = 2
            Case "cos"
                cmbfuncX.ListIndex = 3
            Case "sin"
                cmbfuncX.ListIndex = 4
            Case "tan"
                cmbfuncX.ListIndex = 5
            Case Else
                cmbfuncX.ListIndex = 0
         End Select
         Select Case FuncY
            Case "none"
                cmbfuncY.ListIndex = 0
            Case "log"
                cmbfuncY.ListIndex = 1
            Case "exp"
                cmbfuncY.ListIndex = 2
            Case "cos"
                cmbfuncY.ListIndex = 3
            Case "sin"
                cmbfuncY.ListIndex = 4
            Case "tan"
                cmbfuncY.ListIndex = 5
            Case Else
                cmbfuncY.ListIndex = 0
         End Select
         End If
      End If
      
   'shift the focus to the cmdAccept button
   'so that the user can just enter a carriage return
   PlotInfofrm.Show
   waitime = Timer
   Do Until Timer > waitime + 0.01
      DoEvents
   Loop
   Call keybd_event(VK_TAB, 0, 0, 0)
   Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
   Do Until Timer > waitime + 0.01
      DoEvents
   Loop
   Call keybd_event(VK_TAB, 0, 0, 0)
   Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
   Do Until Timer > waitime + 0.01
      DoEvents
   Loop
   Call keybd_event(VK_TAB, 0, 0, 0)
   Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'check if information has been changed
   'and record changes
   Cancel = False
   Set PlotInfofrm = Nothing
   PlotInfofrmVis = False
End Sub

Private Sub optAutomatic_Click()
   chkColor% = 0
End Sub

Private Sub optBar_Click()
   chkPlotType% = 2
End Sub

Private Sub optBlack_Click()
   chkColor% = 1
End Sub

Private Sub optFilledCircle_Click()
   chkPlotType% = 8
End Sub

Private Sub optDash_Click()
   chkPlotType% = 3
End Sub

Private Sub optDashDot_Click()
   chkPlotType% = 5
End Sub

Private Sub optDashDotDot_Click()
   chkPlotType% = 6
End Sub

Private Sub optDot_Click()
   chkPlotType% = 4
End Sub

Private Sub optFF1_Click()
  chkFF% = 0
End Sub

Private Sub optFF10_Click()
  chkFF% = 9
End Sub

Private Sub optFF11_Click()
  chkFF% = 10
End Sub

Private Sub optFF2_Click()
  chkFF% = 1
End Sub

Private Sub optFF3_Click()
  chkFF% = 2
End Sub

Private Sub optFF4_Click()
  chkFF% = 3
End Sub

Private Sub optFF5_Click()
   chkFF% = 4
End Sub

Private Sub optFF6_Click()
   chkFF% = 5
End Sub

Private Sub optFF7_Click()
   chkFF% = 6
End Sub

Private Sub optFF8_Click()
   chkFF% = 7
End Sub

Private Sub optFF9_Click()
   chkFF% = 8
End Sub

Private Sub optCyan_Click()
   chkColor% = 4
End Sub

Private Sub Option5_Click()
   chkColor% = 5
End Sub

Private Sub optCircle_Click()
   chkPlotType% = 7
End Sub

Private Sub Option6_Click()
   chkColor% = 6
End Sub

Private Sub Option7_Click()
   chkColor% = 7
End Sub

Private Sub Option8_Click()
   chkColor% = 8
End Sub

Private Sub Option9_Click()
   chkColor% = 9
End Sub

Private Sub optLine_Click()
   chkPlotType% = 0
End Sub

Private Sub optPoint_Click()
   chkPlotType% = 1
End Sub

Private Sub optBlue_Click()
   chkColor% = 2
End Sub

Private Sub optGreen_Click()
   chkColor% = 3
End Sub

Private Sub optRed_Click()
   chkColor% = 5
End Sub
