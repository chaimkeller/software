VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FileFormatfm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Formats"
   ClientHeight    =   4830
   ClientLeft      =   4845
   ClientTop       =   2790
   ClientWidth     =   6345
   Icon            =   "FileFormatfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6345
   Begin VB.Frame JKHframe 
      Caption         =   "JKH formats only! (Use this option if the FileFormat.txt file was missing or corrupted)"
      Height          =   855
      Left            =   60
      TabIndex        =   22
      Top             =   480
      Width           =   6195
      Begin VB.CheckBox chkJKH 
         Caption         =   "Automatically record JKH sound velocity vs depth and CSV file formats"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   21
      Top             =   4380
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Step #3: Column Format"
      Height          =   1035
      Left            =   60
      TabIndex        =   11
      Top             =   3180
      Width           =   6195
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   285
         Left            =   5640
         TabIndex        =   20
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtYVal"
         BuddyDispid     =   196614
         OrigLeft        =   5640
         OrigTop         =   420
         OrigRight       =   5880
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtYVal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   19
         Text            =   "0"
         Top             =   420
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   285
         Left            =   3660
         TabIndex        =   18
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtXVal"
         BuddyDispid     =   196616
         OrigLeft        =   3480
         OrigTop         =   420
         OrigRight       =   3720
         OrigBottom      =   675
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtXVal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   17
         Text            =   "0"
         Top             =   420
         Width           =   555
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   1620
         TabIndex        =   16
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtNumCol"
         BuddyDispid     =   196618
         OrigLeft        =   1560
         OrigTop         =   480
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNumCol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Text            =   "0"
         Top             =   420
         Width           =   555
      End
      Begin VB.Label lblXVal 
         Caption         =   "For Y values use column #:"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "For X values use column #"
         Height          =   375
         Left            =   2100
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Total Num. of Columns:"
         Height          =   435
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step #2: Row Format"
      Height          =   915
      Left            =   60
      TabIndex        =   8
      Top             =   2220
      Width           =   6195
      Begin VB.OptionButton optCsvNum 
         Caption         =   "Rows are comma or tab or space delimited NUMBERS ONLY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   540
         Width           =   5535
      End
      Begin VB.OptionButton optCsvAll 
         Caption         =   "Rows are comma delimited MIXTURE of strings and numbers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Accept && &Record"
      Height          =   375
      Left            =   2460
      TabIndex        =   6
      Top             =   4380
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   4380
      Width           =   1155
   End
   Begin VB.ComboBox cmbFormat 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2820
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step #1: Header Lines"
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   1440
      Width           =   6195
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   4216
         TabIndex        =   3
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtHeader"
         BuddyDispid     =   196630
         OrigLeft        =   4440
         OrigTop         =   300
         OrigRight       =   4680
         OrigBottom      =   585
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHeader 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3540
         TabIndex        =   2
         Text            =   "0"
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Number of (string) Header  Lines:"
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
         Left            =   300
         TabIndex        =   4
         Top             =   300
         Width           =   3015
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6360
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label2 
      Caption         =   "File Format #:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1560
      TabIndex        =   7
      Top             =   50
      Width           =   1215
   End
End
Attribute VB_Name = "FileFormatfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormChange As Boolean
Dim FormSave As Boolean, optcsv%

Private Sub chkJKH_Click()
  Dim fnum%
  For fnum% = 9 To 10
        Select Case fnum%
           Case 9 'evaulation files, T vs P, SV vs P, SVC vs P
              FilForm(0, fnum%) = 4
              FilForm(1, fnum%) = 1
              FilForm(2, fnum%) = 2
              FilForm(3, fnum%) = 2
              FilForm(4, fnum%) = 1
           Case 10 'CSV file format depth vs temperature
              FilForm(0, fnum%) = 4
              FilForm(1, fnum%) = 0
              FilForm(2, fnum%) = 8
              FilForm(3, fnum%) = 3
              FilForm(4, fnum%) = 2
           Case Else
        End Select
        fil% = FreeFile
        Open direct$ & "\FilFormat.txt" For Output As #fil%
        Dim I%
        Write #fil%, "This file is used by Plot. Don't erase it!"
        For I% = 0 To 10
          Write #fil%, I%, FilForm(0, I%), FilForm(1, I%), FilForm(2, I%), FilForm(3, I%), FilForm(4, I%)
        Next I%
        Close #fil%
        'unload it and reload it
        Unload FileFormatfm
        FileFormatfm.Visible = True
   Next fnum%
End Sub

Private Sub cmbFormat_Change()
  'before going on, record any changes
  Dim fnum%
  fnum% = cmbFormat.List(cmbFormat.ListIndex)
  FilForm(0, fnum%) = FileFormatfm.txtHeader
  If optcsv% = 0 Then
     FilForm(1, fnum%) = 0
  Else
     FilForm(1, fnum%) = 1
     End If
  FilForm(2, fnum%) = FileFormatfm.txtNumCol
  FilForm(3, fnum%) = FileFormatfm.txtXVal
  FilForm(4, fnum%) = FileFormatfm.txtYVal
End Sub

Private Sub cmbFormat_click()
   Dim fnum%
   fnum% = cmbFormat.List(cmbFormat.ListIndex) - 1
   SetFormat (fnum%)
End Sub
Private Sub cmbFormat_DropDown()
  'before going on, record any changes
  'Dim fnum%
  'fnum% = cmbFormat.List(cmbFormat.ListIndex)
  'FilForm(0, fnum%) = FileFormatfm.txtHeader
  'If optcsv% = 0 Then
  '   FilForm(1, fnum%) = 0
  'Else
  '   FilForm(1, fnum%) = 1
  '   End If
  'FilForm(2, fnum%) = FileFormatfm.txtNumCol
  'FilForm(3, fnum%) = FileFormatfm.txtXVal
  'FilForm(4, fnum%) = FileFormatfm.txtYVal
End Sub

Private Sub cmdAccept_Click()
  'record changes
  Dim fnum%
  fnum% = cmbFormat.List(cmbFormat.ListIndex) - 1
  FilForm(0, fnum%) = FileFormatfm.txtHeader
  If optcsv% = 0 Then
     FilForm(1, fnum%) = 0
  Else
     FilForm(1, fnum%) = 1
     End If
  FilForm(2, fnum%) = FileFormatfm.txtNumCol
  FilForm(3, fnum%) = FileFormatfm.txtXVal
  FilForm(4, fnum%) = FileFormatfm.txtYVal
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   Set FileFormatfm = Nothing
End Sub

Private Sub cmdRecord_Click()
   
  'Record changes permanently
  Dim fnum%
  fnum% = cmbFormat.List(cmbFormat.ListIndex) - 1
  FilForm(0, fnum%) = FileFormatfm.txtHeader
  If optcsv% = 0 Then
     FilForm(1, fnum%) = 0
  Else
     FilForm(1, fnum%) = 1
     End If
  FilForm(2, fnum%) = FileFormatfm.txtNumCol
  FilForm(3, fnum%) = FileFormatfm.txtXVal
  FilForm(4, fnum%) = FileFormatfm.txtYVal
      
  fil% = FreeFile
  Open direct$ & "\FilFormat.txt" For Output As #fil%
  Dim I%
  Write #fil%, "This file is used by Plot. Don't erase it!"
  For I% = 0 To 10
    Write #fil%, I%, FilForm(0, I%), FilForm(1, I%), FilForm(2, I%), FilForm(3, I%), FilForm(4, I%)
  Next I%
  Close #fil%
  FormSave = True
  FormChange = True
End Sub

Private Sub Form_Load()
   Dim I%
   
   'load up file format numbers
   For I% = 1 To 11
      cmbFormat.AddItem Str(I%)
   Next I%
   
   'load up saved formats if they exist
   'if they don't exist, create defaults
     If Dir(direct$ & "\FilFormat.txt") = "" Then
      fil% = FreeFile
      Open direct$ & "\FilFormat.txt" For Output As #fil%
      Write #fil%, "This file is used by Plot. Don't erase it!"
      For I% = 0 To 10
         Write #fil%, I%, 0, 0, 0, 0, 0
         FilForm(0, I%) = 0
       FilForm(1, I%) = 0
         FilForm(2, I%) = 0
         FilForm(3, I%) = 0
         FilForm(4, I%) = 0
      Next I%
      Close #fil%
   Else
      fil% = FreeFile
      Open direct$ & "\FilFormat.txt" For Input As #fil%
      Input #fil%, doclin$
      For I% = 0 To 10
         Input #fil%, num%, FilForm(0, I%), FilForm(1, I%), FilForm(2, I%), FilForm(3, I%), FilForm(4, I%)
      Next I%
      Close #fil%
   End If
   
   'now set up controls according to saved formats
   'display first file format
   cmbFormat.ListIndex = 0
   
End Sub

Sub SetFormat(FilNum As Integer)
   FileFormatfm.txtHeader = FilForm(0, FilNum)
   FileFormatfm.txtHeader.Refresh
   If FilForm(1, FilNum) = 0 Then
      FileFormatfm.optCsvAll.Value = True
   Else
      FileFormatfm.optCsvNum = True
      End If
   FileFormatfm.txtNumCol = FilForm(2, FilNum)
   FileFormatfm.txtNumCol.Refresh
   FileFormatfm.txtXVal = FilForm(3, FilNum)
   FileFormatfm.txtXVal.Refresh
   FileFormatfm.txtYVal = FilForm(4, FilNum)
   FileFormatfm.txtYVal.Refresh
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FormSave = False Then CheckFormChange
If FormChange = True And FormSave = False Then
      response = MsgBox("Save the changes?", vbYesNoCancel + vbExclamation, "Plotting")
      If response = vbYes Then
         Call cmdRecord_Click
      ElseIf response = vbCancel Then
         Cancel = True
         Exit Sub
         End If
      End If
End Sub

Private Sub optCsvAll_Click()
  optcsv% = 0
End Sub

Private Sub optCsvNum_Click()
   optcsv% = 1
End Sub

Sub CheckFormChange()

    'check if the inputs changed from last recorded values
    FormChange = False
    
    fil% = FreeFile
    Open direct$ & "\FilFormat.txt" For Input As #fil%
    Input #fil%, doclin$
    For fnum% = 0 To 10
        Input #fil%, num%, FilForm0%, FilForm1%, FilForm2%, FilForm3%, FilForm4%
         
        If FilForm(0, fnum%) <> FilForm0% Then FormChange = True
        
        If FilForm(1, fnum%) <> FilForm1% Then FormChange = True
          
        If FilForm(2, fnum%) <> FilForm2% Then FormChange = True
        If FilForm(3, fnum%) <> FilForm3% Then FormChange = True
        If FilForm(4, fnum%) <> FilForm4% Then FormChange = True
    Next fnum%
    Close #fil%

End Sub
