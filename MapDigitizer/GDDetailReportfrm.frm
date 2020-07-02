VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form GDDetailReportfrm 
   Caption         =   "Detailed Report"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "GDDetailReportfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   10695
   Begin VB.Frame frmMove 
      DragMode        =   1  'Automatic
      Height          =   915
      Left            =   2640
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   -70
      Width           =   110
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":0442
            Key             =   "cono"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":059C
            Key             =   "diatom"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":06F6
            Key             =   "foram"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":0850
            Key             =   "mega"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":09AA
            Key             =   "multi"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":0B04
            Key             =   "nano"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":0C5E
            Key             =   "ostra"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":0DB8
            Key             =   "paly"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":0F12
            Key             =   "fossil"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":106C
            Key             =   "specimen"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":11C6
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":1320
            Key             =   "categories"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDDetailReportfrm.frx":1774
            Key             =   "Tifview"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDetailReport 
      Height          =   870
      Left            =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1535
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   1508
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "GDDetailReportfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'this form displays database infor about the clicked point
    
    On Error GoTo errhand
    
    With GDDetailReportfrm
    
        .Top = 0   'default starting position and dimensions
        .Left = 0
        .Height = 1290
        .Width = 10815
        
        .TreeView1.Left = 40
        .TreeView1.Width = 2595
        .TreeView1.Top = 0
        
        .frmMove.Left = 2640
        .frmMove.Height = 915
        .frmMove.Top = -70
        .frmMove.Width = 110
        
        .lvwDetailReport.Height = 870
        .lvwDetailReport.Left = 2715
        .lvwDetailReport.Top = 0
        .lvwDetailReport.Width = 7935
    
    End With
    
    LoadDefaultDetailReportInfo
    
    ShowDetails = True 'flag that the form is visible
    
    Exit Sub
    
errhand:
   Screen.MousePointer = vbDefault
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          sEmpty, vbCritical + vbOKOnly, "MapDigitizer"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set GDDetailReportfrm = Nothing
   DetailRecordNum& = 0
   ShowDetails = False
End Sub

Private Sub Form_Resize()

   On Error GoTo errhand
   
   If GDDetailReportfrm.WindowState = vbMinimized Then
      'nothing to do
      Exit Sub
   Else
        If GDDetailReportfrm.Width - GDDetailReportfrm.lvwDetailReport.Left - 120 <= 0 Then
           Exit Sub
           End If
        
        If GDDetailReportfrm.Height <= 435 Then 'too small
           Exit Sub
           End If
   
        With GDDetailReportfrm
        
            .TreeView1.Top = 0
            .TreeView1.Left = 40
            .TreeView1.Height = .Height - 435
            
            .frmMove.Top = -70
            .frmMove.Height = .Height - 375
            .frmMove.Width = 110
            
            .lvwDetailReport.Top = 0
            .lvwDetailReport.Height = .Height - 420
            .lvwDetailReport.Width = .Width - .lvwDetailReport.Left - 120
        
        End With
        
        End If
        
     Exit Sub
     
errhand:

End Sub

Private Sub lvwDetailReport_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   'sort search results according to the column clicked
   lvwDetailReport.Sorted = True
   lvwDetailReport.SortKey = ColumnHeader.Index - 1
   lvwDetailReport.SortOrder = lvwAscending
End Sub

Private Sub lvwDetailReport_DragDrop(Source As Control, X As Single, Y As Single)
  If Source = frmMove Then 'changing widths of treeview and listview
    lvwDetailReport.Left = lvwDetailReport.Left + X
    lvwDetailReport.Width = lvwDetailReport.Width - X
    TreeView1.Width = TreeView1.Width + X
    frmMove.Left = frmMove.Left + X
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lvwDetailReport_ItemClick
' DateTime  : 5/10/2004 09:09
' Author    : Chaim Keller
' Purpose   : load up tree view for chosen item, or view tif file of chosen record
'---------------------------------------------------------------------------------------
'
Private Sub lvwDetailReport_ItemClick(ByVal item As MSComctlLib.ListItem)
  
  Dim RecordNum&
  RecordNum& = item.Index + NearestPnt& - 1
  
   On Error GoTo lvwDetailReport_ItemClick_Error

  If InStr(lvwDetailReport.ColumnHeaders(1), "Place Name") <> 0 Or InStr(lvwDetailReport.ColumnHeaders(1), "Well Name") <> 0 Then
     'load up treeView Control for this record if clicked on record info
      DetailRecordNum& = RecordNum&
      Call LoadTreeView(RecordNum&)
  
  ElseIf InStr(lvwDetailReport.ColumnHeaders(1), "Fossil Names") <> 0 And GDDetailReportfrm.lvwDetailReport.ListItems.item(1) = "Click to view tif's file fossil information" Then
      'clicking on Fossil Name information of scanned database file
      'this means that want to view tif file
      
     'another check for scanned database
      If InStr(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), "*") Then
         'old scanned database record, derive its OKey
         pos1& = InStr(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), "*")
         pos2& = InStr(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), "/")
         numOKey& = val(Mid$(GDReportfrm.lvwReport.ListItems(DetailRecordNum&).SubItems(4), _
                 pos1& + 1, pos2& - pos1& - 1))
       
         'now query for the tif file name and path
         Call FindTifPath(numOKey&, numOFile$)
       
         Select Case numOFile$
            Case "-1" 'error flag
               MsgBox "Tif file not found!" & _
                      vbCrLf & "(Apparently no tif file is associated with this record)", _
                      vbInformation + vbOKOnly, App.Title
            Case Else 'view the file
               If Dir(tifDir$ & "\" & UCase$(numOFile$)) <> sEmpty Then
                  Shell (tifCommandLine$ & " " & tifDir$ & "\" & numOFile$)
               Else
                  Call MsgBox("The path: " & tifDir$ & "\" & UCase$(numOFile$) & " was not found or is not accessible!" & vbLf & vbLf & _
                              "Check the defined path to the tif files in the options menu, and try again", vbExclamation + vbOKOnly, App.Title)
                  End If
         End Select
       
         End If

     
     End If

   On Error GoTo 0
   Exit Sub

lvwDetailReport_ItemClick_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lvwDetailReport_ItemClick of Form GDDetailReportfrm"
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
  If Source = frmMove Then 'changing widths of treeview and listview
    deltaX = frmMove.Left - X
    lvwDetailReport.Left = lvwDetailReport.Left - deltaX
    lvwDetailReport.Width = lvwDetailReport.Width + deltaX
    If TreeView1.Width - deltaX >= 0 Then '(width has to be >=0)
       TreeView1.Width = TreeView1.Width - deltaX
       End If
    frmMove.Left = frmMove.Left - deltaX
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    On Error GoTo errhand
   
    Dim FosStr$, fosstbl$, FossilTag$
    Dim FosTbl As String
    
    If DetailRecordNum& = 0 Then DetailRecordNum& = NearestPnt&

    NoPop& = 0
    Select Case Node.Tag
        'for each unique tag poopulate the listview control
        'first extract the fossil id (res_id) in the appropriate
        'table and then query it for that record
        Case "conod fossil"
            FosStr$ = "conod"
            fosstbl$ = "condores"
            FosIcon$ = "cono"
        Case "conod fossil tag"
            fosstbl$ = "condores"
            FosTbl = "condofos"
            FossilTag$ = "conod"
            FosIcon$ = "cono"
            'query for foram fossil names
            Call DetailedFos(FossilTag$, fosstbl$, FosTbl, sArrConoNames, FosIcon$)
            Exit Sub
        Case "diatom fossil"
            FosStr$ = "diato"
            fosstbl$ = "diatores"
            FosIcon$ = "diatom"
        Case "diatom fossil tag"
            fosstbl$ = "diatores"
            FosTbl = "diatofos"
            FossilTag$ = "diato"
            FosIcon$ = "diatom"
            'query for foram fossil names
            Call DetailedFos(FossilTag$, fosstbl$, FosTbl, sArrDiatomNames, FosIcon$)
            Exit Sub
        Case "foram fossil"
            FosStr$ = "foram"
            fosstbl$ = "foramres"
            FosIcon$ = "foram"
        Case "foram fossil tag"
            fosstbl$ = "foramres"
            FosTbl = "foramfos"
            FossilTag$ = "foram"
            FosIcon$ = "foram"
            'query for foram fossil names
            Call DetailedFos(FossilTag$, fosstbl$, FosTbl, sArrForamNames, FosIcon$)
            Exit Sub
        Case "megaf fossil"
            FosStr$ = "megaf"
            fosstbl$ = "megares"
            FosIcon$ = "mega"
        Case "mega fossil tag"
            fosstbl$ = "megares"
            FosTbl = "megafos"
            FossilTag$ = "megaf"
            FosIcon$ = "mega"
            'query for foram fossil names
            Call DetailedFos(FossilTag$, fosstbl$, FosTbl, sArrMegaNames, FosIcon$)
            Exit Sub
        Case "nano fossil"
            FosStr$ = "nanno"
            fosstbl$ = "nanores"
            FosIcon$ = "nano"
        Case "nano fossil tag"
            fosstbl$ = "nanores"
            FosTbl = "nanofos"
            FossilTag$ = "nanno"
            FosIcon$ = "nano"
            'query for foram fossil names
            Call DetailedFos(FossilTag$, fosstbl$, FosTbl, sArrNanoNames, FosIcon$)
            Exit Sub
        Case "ostra fossil"
            FosStr$ = "ostra"
            fosstbl$ = "ostrares"
            FosIcon$ = "ostra"
        Case "ostra fossil tag"
            fosstbl$ = "ostrares"
            FosTbl = "ostrafos"
            FossilTag$ = "ostra"
            FosIcon$ = "ostra"
            'query for foram fossil names
            Call DetailedFos(FossilTag$, fosstbl$, FosTbl, sArrOstraNames, FosIcon$)
            Exit Sub
        Case "palyn fossil"
            FosStr$ = "palyn"
            fosstbl$ = "palynres"
            FosIcon$ = "paly"
        Case "paly fossil tag"
            fosstbl$ = "palynres"
            FosTbl = "palynfos"
            FossilTag$ = "palyn"
            FosIcon$ = "paly"
            'query for foram fossil names
            Call DetailedFos(FossilTag$, fosstbl$, FosTbl, sArrPalyNames, FosIcon$)
            Exit Sub
        Case "Specimen Name"
            'back to root
            LoadDefaultDetailReportInfo
            ShowDetailedReport
            NoPop& = 1
       Case "categories"
           Exit Sub
       Case Else
            'back to root
            LoadDefaultDetailReportInfo
            ShowDetailedReport
            NoPop& = 1
    End Select
    'query the relevant information from the database
    'and display in the the ListView control
    If NoPop& = 0 Then
       Call PopulateListView(FosStr$, fosstbl$, FosIcon$)
    Else 'returning to main branch, rehighlight current item in listview
       GDDetailReportfrm.lvwDetailReport.ListItems(DetailRecordNum& - NearestPnt& + 1).EnsureVisible
       GDDetailReportfrm.lvwDetailReport.ListItems(DetailRecordNum& - NearestPnt& + 1).Selected = True
       End If
       
    If Not SearchDB Then
      'gdsearchfrm loaded since took genus/species
      'info from it.  So make it invisible
      'but don't unload it case more querying of it will be done
      GDSearchfrm.Visible = False
      End If

    Exit Sub
    
errhand:
    Screen.MousePointer = vbDefault
    MsgBox "Encountered error #: " & Err.Number & vbLf & _
         Err.Description & vbLf & _
         "in module TreeView1:NodeClick" & vbLf & _
         "You probably won't get a complete detailed report.", _
         vbCritical + vbOKOnly, "MapDigitizer"

    
End Sub

