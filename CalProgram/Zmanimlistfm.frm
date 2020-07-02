VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Zmanimlistfm 
   Caption         =   "Z'manim list"
   ClientHeight    =   3780
   ClientLeft      =   4545
   ClientTop       =   2670
   ClientWidth     =   3525
   Icon            =   "Zmanimlistfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   3525
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "resortbut"
            Object.ToolTipText     =   "sort times from earliest to lattest"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tablebut"
            Object.ToolTipText     =   "Generate table"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "gridbut"
            Object.ToolTipText     =   "Generate flex grid table"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "savebut"
            Object.ToolTipText     =   "save the current display to a file"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printbut"
            Object.ToolTipText     =   "print out the currently displayed listing"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.CommandButton zmanbut 
         Caption         =   "Command1"
         Height          =   195
         Left            =   3360
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.ComboBox cmbFontName 
         Height          =   315
         Left            =   1820
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "Arial"
         Top             =   60
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3060
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1860
         Top             =   420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3315
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5847
      _Version        =   393216
      Rows            =   385
      Cols            =   50
      FixedCols       =   3
      BackColor       =   12648447
      BackColorFixed  =   8438015
      ForeColorFixed  =   0
      GridColor       =   33023
      GridColorFixed  =   16576
      AllowUserResizing=   1
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2220
      Top             =   -180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":0EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":10A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":13BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":16D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Zmanimlistfm.frx":1A88
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Zmanimlistfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tmpzman(49) As String * 9



Private Sub cmbFontName_Click()
'   Call cmbFontName_Change
        On Error GoTo errhand
        List1.Visible = False
        Screen.MousePointer = vbHourglass
        If reorder = True Then
           totnum% = numsort%
        Else
           totnum% = newnum%
           End If
        MSFlexGrid1.Visible = True
        cmbFontName.Visible = True
        If MSFlexGrid1.Visible = False Then
             List1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
             List1.Width = Zmanimlistfm.Width - 255 '315 ' 30
             List1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
        Else
             MSFlexGrid1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
             MSFlexGrid1.Width = Zmanimlistfm.Width - 255 '315 ' 30
             MSFlexGrid1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
             End If
        'populate the flex grid with the zemanim
        MSFlexGrid1.Rows = difdyy% + 1
        MSFlexGrid1.Cols = totnum% + 3
        MSFlexGrid1.Font = cmbFontName.Text
        MSFlexGrid1.Font.Size = 10
        'MSFlexGrid1.Font.Bold = True
        
        'generate header
        outdoc$ = sEmpty
        For m% = 0 To totnum%
           'restore the spaces
           For n% = 1 To Len(zmannames$(m%))
              If Mid$(zmannames$(m%), n%, 1) = "_" Then
                 Mid$(zmannames$(m%), n%, 1) = Chr$(32)
                 End If
           Next n%
           outdoc$ = outdoc$ + "|^" + zmannames$(m%)
        Next m%
        'outdoc$ = "^תאריך_לועזי" + "|^    יום" + "|^תאריך_עברי" + outdoc$
        If optionheb = True Then
           outdoc$ = "^תאריך עברי    " + "|^יום    " + "|^civil date    " + outdoc$
        ElseIf optionheb = False Then
           outdoc$ = "^hebrew date   |^day               |^civil date      " + outdoc$
           End If
        
'        If optionheb = True Then
'           outdoc$ = "^תאריך עברי    " + "|^יום    " + "|^civil date    " + outdoc$
'        ElseIf optionheb = False Then
'           outdoc$ = outdoc$ + "   |^civil date   |^day   |^hebrew date"
'           End If
        MSFlexGrid1.FormatString = outdoc$
        
        numday% = -1
        For i% = 1 To endyr%
           If mmdate%(2, i%) > mmdate%(1, i%) Then
              k% = 0
              For j% = mmdate%(1, i%) To mmdate%(2, i%)
                  numday% = numday% + 1
                  k% = k% + 1
                  outdoc$ = sEmpty
                  For m% = 3 To totnum% + 3
                     If Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00" Then
                        zmantimes(m% - 3 - 3, numday%) = String$(Len(zmantimes(m% - 3, numday%)), "-")
                     ElseIf Mid$(zmantimes(m% - 3, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00"
                        End If
                     MSFlexGrid1.TextArray(skyp2(numday% + 1, m%)) = Trim$(zmantimes(m% - 3, numday%))
                  Next m%
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 0)) = Trim$(stortim$(3, i% - 1, k% - 1))
                  Call InsertHolidays(calday$, i%, k%)
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 1)) = calday$
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 2)) = Trim$(stortim$(2, i% - 1, k% - 1))
              Next j%
           ElseIf mmdate%(2, i%) < mmdate%(1, i%) Then
              k% = 0
              For j% = mmdate%(1, i%) To yrend%(0)
                  numday% = numday% + 1
                  k% = k% + 1
                  outdoc$ = sEmpty
                  For m% = 3 To totnum% + 3
                     If Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00" Then
                        zmantimes(m% - 3, numday%) = String$(Len(zmantimes(m% - 3, numday%)), "-")
                     ElseIf Mid$(zmantimes(m% - 3, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00"
                        End If
                     MSFlexGrid1.TextArray(skyp2(numday% + 1, m%)) = Trim$(zmantimes(m% - 3, numday%))
                  Next m%
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 0)) = Trim$(stortim$(3, i% - 1, k% - 1))
                  Call InsertHolidays(calday$, i%, k%)
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 1)) = calday$
                     MSFlexGrid1.TextArray(skyp2(numday% + 1, 2)) = Trim$(stortim$(2, i% - 1, k% - 1))
             Next j%
              yrn% = yrn% + 1
              For j% = 1 To mmdate%(2, i%)
                  k% = k% + 1
                  numday% = numday% + 1
                  outdoc$ = sEmpty
                  For m% = 3 To totnum% + 3
                     If Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00" Then
                        zmantimes(m% - 3, numday%) = String$(Len(zmantimes(m% - 3, numday%)), "-")
                     ElseIf Mid$(zmantimes(m% - 3, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00"
                        End If
                     MSFlexGrid1.TextArray(skyp2(numday% + 1, m%)) = Trim$(zmantimes(m% - 3, numday%))
                  Next m%
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 0)) = Trim$(stortim$(3, i% - 1, k% - 1))
                  Call InsertHolidays(calday$, i%, k%)
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 1)) = calday$
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 2)) = Trim$(stortim$(2, i% - 1, k% - 1))
              Next j%
              End If
        Next i%
        
errhand:
        Screen.MousePointer = vbDefault
        
        End Sub

Private Sub cmbFontName_DblClick()
   Call cmbFontName_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'version: 04/08/2003
  
   Unload Me
   Set Zmanimlistfm = Nothing
End Sub

Private Sub Form_Resize()
   On Error GoTo 900
   If MSFlexGrid1.Visible = False Then
        List1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
        List1.Width = Zmanimlistfm.Width - 255 '315 ' 30
        List1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
   Else
        MSFlexGrid1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
        MSFlexGrid1.Width = Zmanimlistfm.Width - 255 '315 ' 30
        MSFlexGrid1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
        End If
900
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)

Dim noxsl As Boolean

   Select Case Button.Key
     Case "savebut"
        On Error GoTo c3error
5       CommonDialog1.CancelError = True
        CommonDialog1.Filter = "zmanim as spreadsheet file (csv)|*.csv|output table in htm format (*.htm)|*.htm|output table in zipped htm format (*.zip)|*.zip|output table in xml format (*.xml)|*.xml"
        If Not TufikZman Then
            CommonDialog1.FilterIndex = 0
            CommonDialog1.FileName = drivjk$ + "*.csv"
        Else
            CommonDialog1.FilterIndex = 3
            prefix$ = Replace(eroscity$, " ", "_")
            CommonDialog1.FileName = drivjk$ + prefix$ & yrheb% & ".xml"
            End If
        CommonDialog1.ShowSave
        filnam$ = CommonDialog1.FileName
        myfile = Dir(filnam$)
        If myfile <> sEmpty Then
           response = MsgBox("A file with this name already exists! Overwrite it?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Cal Program")
           If response <> vbYes Then
              GoTo 5
              End If
           End If
        Screen.MousePointer = vbHourglass
        pos% = InStr(filnam$, ".")
        fileroot$ = Mid$(filnam$, 1, pos% - 1)
        'find root without directory
        For jr% = pos% - 1 To 1 Step -1
           cha$ = Mid$(filnam$, jr%, 1)
           If cha$ = "\" Or cha$ = ":" Then Exit For
        Next jr%
        roots$ = Mid$(filnam$, jr% + 1, pos% - jr% - 1)
        ext$ = Mid(filnam$, pos% + 1, Len(filnam$) - pos%)
        
        If ext$ = "zip" Then 'test for 8 characters
           If Len(roots$) > 8 Then
              MsgBox "PKZIP only supports old DOS format names with length <= 8!" & vbCrLf & _
                     "Pick another name...", _
                     vbOKOnly + vbExclamation, "Cal Program"
              GoTo 5
              End If
           End If
        
        Screen.MousePointer = vbHourglass
        
'        Call WriteTables(fileroot$, roots$, ext$)
        
        If ext$ = "xml" And TufikZman Then
           Call WriteTufikTables(fileroot$, roots$, ext$)
        Else
           Call WriteTables(fileroot$, roots$, ext$)
           End If
        
        Screen.MousePointer = vbDefault
        changes = False
c3error:
        Close
        Screen.MousePointer = vbDefault
        Exit Sub
     Case "resortbut"
       If resortbutton = True Then Exit Sub
       'load sort array
       daycheck% = 0 'try using the first day as sorting model
10     For i% = 0 To newnum%
          If Trim$(zmantimes(i%, daycheck%)) = "none" Or Mid$(Trim$(zmantimes(i%, daycheck%)), 1, 2) = "00" Then
             daycheck% = daycheck% + 1 'try the next day as the sorting model
             If daycheck% > numday% Then
                response = MsgBox("Too many ""nones"", can't do resorting", vbExclamation + vbOKOnly, "Cal Program")
                Exit Sub
             Else
                GoTo 10
                End If
             End If
          tmpzman(i%) = zmantimes(i%, daycheck%)
       Next i%
       'resort the times according to earliest to lattest times
       begi% = 0
50     earlyi% = begi%
       tim0$ = Trim$(tmpzman(begi%)) 'use day 1 to determine sorting
       If Len(tim0$) = 7 And InStr(tim0$, ":") Then
          zman0 = Val(Mid$(tim0$, 1, 1)) + Val(Mid$(tim0$, 3, 2)) / 60 + Val(Mid$(tim0$, 6, 2)) / 3600
       ElseIf Len(tim0$) = 8 And InStr(tim0$, ":") Then
          zman0 = Val(Mid$(tim0$, 1, 2)) + Val(Mid$(tim0$, 4, 2)) / 60 + Val(Mid$(tim0$, 7, 2)) / 3600
       Else
          zman0 = Val(tim0$)
          End If
       For i% = begi% + 1 To newnum%
           tim0$ = Trim$(tmpzman(i%))
           If Len(tim0$) = 7 And InStr(tim0$, ":") Then
              zmantim = Val(Mid$(tim0$, 1, 1)) + Val(Mid$(tim0$, 3, 2)) / 60 + Val(Mid$(tim0$, 6, 2)) / 3600
           ElseIf Len(tim0$) = 8 And InStr(tim0$, ":") Then
              zmantim = Val(Mid$(tim0$, 1, 2)) + Val(Mid$(tim0$, 4, 2)) / 60 + Val(Mid$(tim0$, 7, 2)) / 3600
           Else
              zmantim = Val(tim0$)
              End If
           If zmantim < zman0 Then
              zman0 = zmantim
              earlyi% = i%
              End If
        Next i%
        'record this time, and shift
        If begi% <> earlyi% Then  'shift
           tmptim$ = tmpzman(earlyi%)
           tmpzman(earlyi%) = tmpzman(begi%)
           tmpzman(begi%) = tmptim$
           End If
        begi% = begi% + 1
        
        If begi% = newnum% Then GoTo 100 'finished sorting, record values
        GoTo 50 'else, sort next item
        
100    sortnum% = FreeFile
       Open drivjk$ + "zmansort.out" For Output As #sortnum%
       For i% = 0 To newnum%
          For j% = 0 To newnum%
             If zmantimes(j%, daycheck%) = tmpzman(i%) Then
                Write #sortnum%, j%
                Exit For
                End If
          Next j%
       Next i%
       Close #sortnum%
       Unload Zmanimlistfm
       neworder = True
       reorder = True
       Zmanimform.calendarbut.Value = True
       
     Case "printbut"
     Case "tablebut"
        RemoveUnderline = False
        'generate table to temporary file
        MSFlexGrid1.Visible = False
        cmbFontName.Visible = False
        List1.Visible = True
        If MSFlexGrid1.Visible = False Then
             List1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
             List1.Width = Zmanimlistfm.Width - 255 '315 ' 30
             List1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
        Else
             MSFlexGrid1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
             MSFlexGrid1.Width = Zmanimlistfm.Width - 255 '315 ' 30
             MSFlexGrid1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
             End If
        tmpnum% = FreeFile
        Open drivjk$ + "table.new" For Output As #tmpnum%
        Screen.MousePointer = vbHourglass
        
        'generate header
        outdoc$ = sEmpty
        If reorder = True Then
           totnum% = numsort%
        Else
           totnum% = newnum%
           End If
        For m% = totnum% To 0 Step -1
           'fill all the spaces with
           For n% = 1 To Len(zmannames$(m%))
              If Mid$(zmannames$(m%), n%, 1) = Chr$(32) Then
                 Mid$(zmannames$(m%), n%, 1) = "_"
                 End If
           Next n%
           If optionheb = True Then
              outdoc$ = outdoc$ + "   " + zmannames$(m%)
           Else
              outdoc$ = zmannames$(m%) + "   " + outdoc$
              End If
        Next m%
        If optionheb = True Then
           outdoc$ = outdoc$ & String(3, " ") & "תעריך עברי   יום    תעריך לועזי"
        ElseIf optionheb = False Then
           outdoc$ = "hebrew date" & "   " & "day" & "   " & "civil date" & "    " & outdoc$
           End If
        Print #tmpnum%, outdoc$
        
        numday% = -1
        For i% = 1 To endyr%
           If mmdate%(2, i%) > mmdate%(1, i%) Then
              k% = 0
              For j% = mmdate%(1, i%) To mmdate%(2, i%)
                  numday% = numday% + 1
                  k% = k% + 1
                  outdoc$ = sEmpty
                  For m% = totnum% To 0 Step -1
                     If Mid$(zmantimes(m%, numday%), 1, 2) = "00" Then
                        zmantimes(m%, numday%) = String$(6, "-") 'String$(Len(zmantimes(m%, numday%)), "-")
                     ElseIf Mid$(zmantimes(m%, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m%, numday%), 1, 2) = "00"
                        End If
                     If optionheb = True Then
                        outdoc$ = outdoc$ + "   " + Trim$(zmantimes(m%, numday%))
                     Else
                        outdoc$ = Trim$(zmantimes(m%, numday%)) + "   " + outdoc$
                        End If
                  Next m%
                  Call InsertHolidays(calday$, i%, k%)
                  If optionheb = True Then
                     outdoc$ = outdoc$ + "   " + Trim$(stortim$(2, i% - 1, k% - 1)) + "   " + calday$ + "       " + Trim$(stortim$(3, i% - 1, k% - 1))
                  Else
                     outdoc$ = "   " + Trim$(stortim$(3, i% - 1, k% - 1)) + "   " + calday$ + "   " + Trim$(stortim$(2, i% - 1, k% - 1)) + "   " + outdoc$
                     End If
                  Print #tmpnum%, outdoc$
              Next j%
           ElseIf mmdate%(2, i%) < mmdate%(1, i%) Then
              k% = 0
              For j% = mmdate%(1, i%) To yrend%(0)
                  numday% = numday% + 1
                  k% = k% + 1
                  outdoc$ = sEmpty
                  For m% = totnum% To 0 Step -1
                     If Mid$(zmantimes(m%, numday%), 1, 2) = "00" Then
                        zmantimes(m%, numday%) = String$(6, "-") ' String$(Len(zmantimes(m%, numday%)), "-")
                     ElseIf Mid$(zmantimes(m%, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m%, numday%), 1, 2) = "00"
                        End If
                     If optionheb = True Then
                        outdoc$ = outdoc$ + "   " + Trim$(zmantimes(m%, numday%))
                     Else
                        outdoc$ = Trim$(zmantimes(m%, numday%)) + "   " + outdoc$
                        End If
                  Next m%
                  Call InsertHolidays(calday$, i%, k%)
                  If optionheb = True Then
                     outdoc$ = outdoc$ + "   " + Trim$(stortim$(2, i% - 1, k% - 1)) + "   " + calday$ + "       " + Trim$(stortim$(3, i% - 1, k% - 1))
                  Else
                     outdoc$ = "   " + Trim$(stortim$(3, i% - 1, k% - 1)) + "   " + calday$ + "   " + Trim$(stortim$(2, i% - 1, k% - 1)) + "   " + outdoc$
                     End If
                  Print #tmpnum%, outdoc$
              Next j%
              yrn% = yrn% + 1
              For j% = 1 To mmdate%(2, i%)
                  k% = k% + 1
                  numday% = numday% + 1
                  outdoc$ = sEmpty
                  For m% = totnum% To 0 Step -1
                     If Mid$(zmantimes(m%, numday%), 1, 2) = "00" Then
                        zmantimes(m%, numday%) = String$(6, "-") 'String$(Len(zmantimes(m%, numday%)), "-")
                     ElseIf Mid$(zmantimes(m%, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m%, numday%), 1, 2) = "00"
                        End If
                     If optionheb = True Then
                        outdoc$ = outdoc$ + "   " + Trim$(zmantimes(m%, numday%))
                     Else
                        outdoc$ = Trim$(zmantimes(m%, numday%)) + "   " + outdoc$
                        End If
                  Next m%
                  Call InsertHolidays(calday$, i%, k%)
                  If optionheb = True Then
                     outdoc$ = outdoc$ + "   " + Trim$(stortim$(2, i% - 1, k% - 1)) + "   " + calday$ + "       " + Trim$(stortim$(3, i% - 1, k% - 1))
                  Else
                     outdoc$ = "   " + Trim$(stortim$(3, i% - 1, k% - 1)) + "   " + calday$ + "   " + Trim$(stortim$(2, i% - 1, k% - 1)) + "   " + outdoc$
                     End If
                   Print #tmpnum%, outdoc$
              Next j%
              End If
        Next i%
        Close #tmpnum%
        List1.Clear
        tmpnum% = FreeFile
        Open drivjk$ + "table.new" For Input As #tmpnum%
        Do Until EOF(tmpnum%)
           Line Input #tmpnum%, doclin$
           List1.AddItem doclin$
        Loop
        Close #tmpnum%
        Screen.MousePointer = vbDefault
     Case "gridbut" 'Flex-Grid display
        RemoveUnderline = True
        List1.Visible = False
        Screen.MousePointer = vbHourglass
        If reorder = True Then
           totnum% = numsort%
        Else
           totnum% = newnum%
           End If
        MSFlexGrid1.Visible = True
        cmbFontName.Visible = True
        If MSFlexGrid1.Visible = False Then
             List1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
             List1.Width = Zmanimlistfm.Width - 255 '315 ' 30
             List1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
        Else
             MSFlexGrid1.Left = 60 '120 'Zmanimlistfm.Left + 120 '- 200
             MSFlexGrid1.Width = Zmanimlistfm.Width - 255 '315 ' 30
             MSFlexGrid1.Height = Zmanimlistfm.Height - 1005 '1200 '615 '30
             End If
        'populate the flex grid with the zemanim
        MSFlexGrid1.Rows = difdyy% + 1
        MSFlexGrid1.Cols = totnum% + 3
        MSFlexGrid1.Font = cmbFontName.Text
        MSFlexGrid1.Font.Size = 10
        'MSFlexGrid1.Font.Bold = True
        
        'generate header
        outdoc$ = sEmpty
        For m% = 0 To totnum%
           'restore the spaces
           For n% = 1 To Len(zmannames$(m%))
              If Mid$(zmannames$(m%), n%, 1) = "_" Then
                 Mid$(zmannames$(m%), n%, 1) = Chr$(32)
              ElseIf Mid$(zmannames$(m%), n%, 1) = "|" Then 'remove the rest of the string that was used for adding/subtracting minutes from zman
                 zmannames$(m%) = Mid$(zmannames$(m%), 1, n% - 1)
                 Exit For
                 End If
           Next n%
           outdoc$ = outdoc$ + "|^" + zmannames$(m%)
        Next m%
        'outdoc$ = "^תאריך_לועזי" + "|^    יום" + "|^תאריך_עברי" + outdoc$
        If optionheb = True Then
           outdoc$ = "^תאריך עברי    " + "|^יום                                " + "|^civil date    " + outdoc$
        ElseIf optionheb = False Then
           outdoc$ = "^hebrew date   |^day               |^civil date      " + outdoc$
           End If
        
        MSFlexGrid1.FormatString = outdoc$
        
        numday% = -1
        For i% = 1 To endyr%
           If mmdate%(2, i%) > mmdate%(1, i%) Then
              k% = 0
              For j% = mmdate%(1, i%) To mmdate%(2, i%)
                  numday% = numday% + 1
                  k% = k% + 1
                  outdoc$ = sEmpty
                  For m% = 3 To totnum% + 3
                     If Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00" Then
                        zmantimes(m% - 3, numday%) = String$(Len(zmantimes(m% - 3, numday%)), "-")
                     ElseIf Mid$(zmantimes(m% - 3, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00"
                        End If
                     MSFlexGrid1.TextArray(skyp2(numday% + 1, m%)) = Trim$(zmantimes(m% - 3, numday%))
                  Next m%
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 0)) = Trim$(stortim$(3, i% - 1, k% - 1))
                  Call InsertHolidays(calday$, i%, k%)
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 1)) = calday$
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 2)) = Trim$(stortim$(2, i% - 1, k% - 1))
              Next j%
           ElseIf mmdate%(2, i%) < mmdate%(1, i%) Then
              k% = 0
              For j% = mmdate%(1, i%) To yrend%(0)
                  numday% = numday% + 1
                  k% = k% + 1
                  outdoc$ = sEmpty
                  For m% = 3 To totnum% + 3
                     If Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00" Then
                        zmantimes(m% - 3, numday%) = String$(Len(zmantimes(m% - 3, numday%)), "-")
                     ElseIf Mid$(zmantimes(m% - 3, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00"
                        End If
                     MSFlexGrid1.TextArray(skyp2(numday% + 1, m%)) = Trim$(zmantimes(m% - 3, numday%))
                  Next m%
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 0)) = Trim$(stortim$(3, i% - 1, k% - 1))
                  Call InsertHolidays(calday$, i%, k%)
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 1)) = calday$
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 2)) = Trim$(stortim$(2, i% - 1, k% - 1))
             Next j%
              yrn% = yrn% + 1
              For j% = 1 To mmdate%(2, i%)
                  k% = k% + 1
                  numday% = numday% + 1
                  outdoc$ = sEmpty
                  For m% = 3 To totnum% + 3
                     If Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00" Then
                        zmantimes(m% - 3, numday%) = String$(Len(zmantimes(m% - 3, numday%)), "-")
                     ElseIf Mid$(zmantimes(m% - 3, numday%), 1, 2) = "24" Then
                        Mid$(zmantimes(m% - 3, numday%), 1, 2) = "00"
                        End If
                     MSFlexGrid1.TextArray(skyp2(numday% + 1, m%)) = Trim$(zmantimes(m% - 3, numday%))
                  Next m%
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 0)) = Trim$(stortim$(3, i% - 1, k% - 1))
                  Call InsertHolidays(calday$, i%, k%)
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 1)) = calday$
                  MSFlexGrid1.TextArray(skyp2(numday% + 1, 2)) = Trim$(stortim$(2, i% - 1, k% - 1))
              Next j%
              End If
        Next i%
        Screen.MousePointer = vbDefault

        
     Case Else
  End Select
End Sub
Function skyp2(row As Integer, col As Integer) As Long
     skyp2 = row * MSFlexGrid1.Cols + col
End Function

Private Sub zmanbut_Click()
        'write html,xml version of sorted zmanim table
        filroot$ = dirint$ + "\" + Mid(servnam$, 1, 8)
        If zmantype% = 0 Then
           ext$ = "csv"
        ElseIf zmantype% = 1 Then
           ext$ = "zip"
        ElseIf zmantype% = 2 Then
           ext$ = "xml"
           End If
           
        Call WriteTables(filroot$, servnam$, ext$)

End Sub
