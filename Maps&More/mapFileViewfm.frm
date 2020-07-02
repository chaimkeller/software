VERSION 5.00
Begin VB.Form mapFileViewfm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File View"
   ClientHeight    =   6015
   ClientLeft      =   5910
   ClientTop       =   1920
   ClientWidth     =   5505
   Icon            =   "mapFileViewfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5505
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   255
      Left            =   4485
      TabIndex        =   15
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear &All"
      Height          =   255
      Left            =   50
      TabIndex        =   14
      Top             =   1140
      Width           =   855
   End
   Begin VB.CommandButton cmdNoEdit 
      Cancel          =   -1  'True
      Caption         =   "Accept without &editing"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Frame frameCities 
      Caption         =   "Click on the Correct Cities Directory (If new city, click on ""d:"")"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5295
      Begin VB.DriveListBox drvFileView 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   280
         Width           =   1335
      End
      Begin VB.DirListBox dirFileView 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblCities 
         Alignment       =   2  'Center
         Caption         =   "d:\cities\acco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   440
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label lblDirectory 
         Caption         =   "Directory Location:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblDrive 
         Caption         =   "Disk Location:"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   280
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5430
      TabIndex        =   3
      Top             =   4920
      Width           =   5490
      Begin VB.Label lblHelp 
         Alignment       =   2  'Center
         Caption         =   "Check the desired lines of the file and press ""Accept"""
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -120
         TabIndex        =   4
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3730
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.ListBox lstFileView 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1440
      Width           =   5500
   End
   Begin VB.TextBox txtFileView 
      Height          =   3450
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   1440
      Width           =   5550
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblFileName 
      Alignment       =   2  'Center
      Caption         =   "Contents of File:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1120
      Width           =   5415
   End
End
Attribute VB_Name = "mapFileViewfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldMode As Boolean

Private Sub cmdAccept_Click()
   'write the checked lines to the file: drivjk$ + "\viewout.tmp"
   FileEdit = True
   
   If InStr(lblFileName.Caption, "scanlist.txt") = 0 Then
      lstFileView.Refresh
      filout% = FreeFile
      Open drivjk_c$ + "viewout.tmp" For Output As #filout%
      For i% = 0 To lstFileView.ListCount - 1
         If lstFileView.Selected(i%) = True Then
            Print #filout%, lstFileView.List(i%)
         End If
      Next
      Close #filout%
   
   ElseIf InStr(lblFileName.Caption, "scanlist.txt") <> 0 Then
      'add the edited version of scanlist.txt to viewout.tmp
      FileViewError = False
      'output the edited version of scanlist.txt to viewout.tmp
      filtmp% = FreeFile
      Open drivjk_c$ & "viewout.tmp" For Output As #filtmp%
      Print #filtmp%, txtFileView.Text
      Close #filtmp%
      Call form_queryunload(0, 0)
      End If
   
  'store name of Eretz Yisroel city.
   If InStr(lblFileName.Caption, "scanlist.txt") = 0 Then
      FileViewDir$ = dirFileView.List(dirFileView.ListIndex)
      If Len(FileViewDir$) < 3 Then FileViewDir$ = sEmpty
      End If
   
   FileViewError = False
   Call form_queryunload(0, 0)
End Sub

Private Sub cmdCancel_Click()
   tblbuttons(26) = 0
   tblbuttons(27) = 0
   Maps.Toolbar1.Buttons(26).value = tbrUnpressed
   Maps.Toolbar1.Buttons(27).value = tbrUnpressed
   FileViewError = True
   FileEdit = False
   Call form_queryunload(0, 0)
End Sub

Private Sub cmdClearAll_Click()
    For i% = 0 To lstFileView.ListCount - 1
       lstFileView.Selected(i%) = False
    Next
End Sub

Private Sub cmdNoEdit_Click()
   
   FileEdit = False
   
   If InStr(lblFileName.Caption, "scanlist.txt") <> 0 Then
      'finished with the analysis
      FileViewError = False
      Call form_queryunload(0, 0)
      End If
   
  'store name of Eretz Yisroel city.
   If InStr(lblFileName.Caption, "scanlist.txt") = 0 Then
      FileViewDir$ = dirFileView.List(dirFileView.ListIndex)
      If Len(FileViewDir$) < 3 Then FileViewDir$ = sEmpty
      End If
   
   FileViewError = False
   
   Call form_queryunload(0, 0)
End Sub

Private Sub cmdSelectAll_Click()
    For i% = 0 To lstFileView.ListCount - 1
       lstFileView.Selected(i%) = True
    Next
End Sub

Private Sub form_load()
   
   AbrevDir$ = sEmpty
   
   oldMode = False 'oldMode = True for listing contents of
                   '                of scanlist in list box
                   '                (old rdhalbat formats)

   lblFileName.Caption = "Contents of File: " + FileViewName
   
   If InStr(lblFileName.Caption, "scanlist.txt") = 0 Then
      'list cities directory
      drvFileView.Drive = Mid$(drivcities$, 1, 1)
      dirFileView.Path = drivcities$
      dirFileView.ListIndex = 1 'put it on the first city directory
   Else 'display scanlist.txt in the edit box
      drvFileView.Visible = False
      dirFileView.Visible = False
      lblDirectory.Visible = False
      lblDrive.Visible = False
      lblCities.Visible = True
      lblCities.Caption = FileViewDir$
      frameCities.Caption = "Cities Directory"
      If Not oldMode Then
         lblHelp.Caption = "Edit the begkmx, endkmx and accept"
         lstFileView.Visible = False
         End If
      End If

   'Copy the desired file for viewing to the temp.
   'file viewin.tmp.  The separate line of this file
   'will be displayed
   FileViewError = False
   lstFileView.Clear

   'Open the file and display each line
   If Dir(drivjk_c$ + "viewin.tmp") = sEmpty Then
      MsgBox "Can't locate viewtmp.tmp file!", vbCritical + vbOKOnly, "FileView"
      cmdCancel_Click
      Exit Sub
      End If
   filview% = FreeFile
   Open drivjk_c$ + "viewin.tmp" For Input As #filview%
   Do Until EOF(filview%)
      Line Input #filview%, doclin$
      If Trim$(doclin$) = sEmpty Then Exit Do
      If InStr(lblFileName.Caption, drivjk_c$ + "scanlist.txt") = 0 Then
         'this is placlist.txt so enable each line
         lstFileView.AddItem doclin$
         lstFileView.Selected(lstFileView.ListCount - 1) = True
      ElseIf InStr(lblFileName.Caption, drivjk_c$ + "scanlist.txt") <> 0 And Not oldMode Then
         'add all the lines to the text box
         If txtFileView.Text = sEmpty Then
            txtFileView.Text = doclin$
         Else
            txtFileView.Text = txtFileView.Text & vbNewLine & doclin$
            End If
      ElseIf InStr(lblFileName.Caption, drivjk_c$ + "scanlist.txt") <> 0 And oldMode Then
        'enable according to sunmode%
         lstFileView.AddItem doclin$
         Select Case sunmode%
            Case 1 'sunrise
               If InStr(doclin$, ".001") <> 0 Or _
                  InStr(doclin$, ".002") <> 0 Or _
                  InStr(doclin$, ".003") <> 0 Then
                  lstFileView.Selected(lstFileView.ListCount - 1) = True
               End If
            Case 0 'sunset
               If InStr(doclin$, ".004") <> 0 Or _
                  InStr(doclin$, ".005") <> 0 Or _
                  InStr(doclin$, ".006") <> 0 Then
                  lstFileView.Selected(lstFileView.ListCount - 1) = True
               End If
         End Select
         End If
   Loop
   Close #filview%
End Sub

Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
   Close
   Unload Me
   Set mapFileViewfm = Nothing
   FileView = False
End Sub
