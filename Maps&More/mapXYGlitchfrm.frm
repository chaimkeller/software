VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mapXYGlitchfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XY Glitch Fix"
   ClientHeight    =   3075
   ClientLeft      =   7260
   ClientTop       =   5670
   ClientWidth     =   4560
   Icon            =   "mapXYGlitchfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFix 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fix th'Glitch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Press to run fix"
      Top             =   2160
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar progbarFixGlitch 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame frmOptions 
      Caption         =   "Options"
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   4095
      Begin MSComCtl2.UpDown UpDownThreshhold 
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   20
         BuddyControl    =   "txtThreshhold"
         BuddyDispid     =   196611
         OrigLeft        =   2640
         OrigTop         =   480
         OrigRight       =   2895
         OrigBottom      =   735
         Max             =   1000
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtThreshhold 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "20"
         Top             =   480
         Width           =   600
      End
      Begin VB.OptionButton optVertical 
         Caption         =   "Vertically oriented"
         Height          =   240
         Left            =   2160
         TabIndex        =   10
         Top             =   200
         Width           =   1575
      End
      Begin VB.OptionButton optHorizontal 
         Caption         =   "Horizontally oriented"
         Height          =   220
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Glitch is mostly in horizontal direction (rows)"
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label lblThreshhold 
         Caption         =   "Threshhold"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   520
         Width           =   855
      End
   End
   Begin VB.TextBox txtYEnd 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "txtYEnd"
      ToolTipText     =   "Y end (> Y start)"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtYStart 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "txtYStart"
      ToolTipText     =   "Y Start (l< Y End)"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtXEnd 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "txtXEnd"
      ToolTipText     =   "X End"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtXStart 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "XStart"
      ToolTipText     =   "Starting X Coordinates"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblYEnd 
      Caption         =   "YEnd"
      Height          =   255
      Left            =   2385
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblYStart 
      Caption         =   "YStart"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblXEnd 
      Caption         =   "XEnd"
      Height          =   255
      Left            =   2385
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblXStart 
      Caption         =   "XStart"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "mapXYGlitchfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure : cmdFix_Click
' Author    : chaim
' Date      : 9/13/2022
' Purpose   : Fix XY glitch
'---------------------------------------------------------------------------------------
'
Private Sub cmdFix_Click()
   Dim KmxStart, KmyStart, TotalNum%, StepResolution%, DTMResolution%
   
   On Error GoTo cmdFix_Click_Error

   progbarFixGlitch.Visible = True
   progbarFixGlitch.value = 0
   progbarFixGlitch.Max = 100
   progbarFixGlitch.Min = 0
   
   StepResolution = 1 'step size for search for glitches, should be smaller than the DTM grid size
   DTMResolution% = 25 'DTM grid size (resolution)
   
'   Call ScreenToGeo(drag1x, drag1y, kmxDrag, kmyDrag, 1, ier%)
'   KmxStart = kmxDrag
'   KmyEnd = kmyDrag
'
'   Call ScreenToGeo(drag2x, drag2y, kmxDrag, kmyDrag, 1, ier%)
'   KmxEnd = kmxDrag
'   KmyStart = kmyDrag
'
'   'Now determine end and start with respect to actual DTM
'   'points:
'   KmxStart = CLng(KmxStart * 0.04) * 25#
'   KmxEnd = CLng(KmxEnd * 0.04) * 25#
'   KmyStart = CLng(KmyStart * 0.04) * 25#
'   KmyEnd = CLng(KmyEnd * 0.04) * 25#

    KmxStart = Val(txtXStart)
    KmxEnd = Val(txtXEnd)
    If (KmxStart < 40000 Or KmxEnd < 40000 Or KmxStart > 300000 Or KmxEnd > 300000) Then
        Call MsgBox("X coordinates are beyond the permitted range." _
                    & vbCrLf & "" _
                    & vbCrLf & "Check your X inputs!" _
                    , vbInformation, "Fix Glitch")
        Exit Sub
        End If
    KmyStart = Val(txtYStart)
    KmyEnd = Val(txtYEnd)
    If (KmyStart < 20000 Or KmyEnd < 20000 Or KmyStart > 1200000 Or KmyEnd > 1200000) Then
        Call MsgBox("Y coordinates are beyond the permitted range." _
                    & vbCrLf & "" _
                    & vbCrLf & "Check your Y inputs!" _
                    , vbInformation, "Fix Glitch")
        Exit Sub
        End If
    
   'convert to grid
   KmxStart = CLng(KmxStart * 0.04) * DTMResolution%
   txtXStart.Text = KmxStart
   KmxEnd = CLng(KmxEnd * 0.04) * DTMResolution%
   txtXEnd.Text = KmxEnd
   KmyStart = CLng(KmyStart * 0.04) * DTMResolution%
   txtYStart.Text = KmyStart
   KmyEnd = CLng(KmyEnd * 0.04) * DTMResolution%
   txtYEnd.Text = KmyEnd
       
tp500: '-----------save the points to the DTM tile----------------------
       response = MsgBox("Backup DTM files before saving changes?" & vbLf & _
                     "(The date will be added as a suffix to the backup tiles)", _
                     vbQuestion + vbYesNoCancel, "Maps&More")
       If response = vbYes Then
          backup% = 1
          End If
               
          
        'determine which tile(s) are being used and back them up
        If optVertical Then
           'glitch is mostly column oriented, i.e, vertical
            TotalNum% = CInt((KmyEnd - KmyStart) / StepResolution)
            CHFind$ = sEmpty
            For kmy = KmyStart To KmyEnd Step StepResolution
              progbarFixGlitch.value = CInt(100 * (kmy - KmyStart) / (KmyEnd - KmyStart))
              numKmx = 0
              sgnkmx = 1
              kmx0 = KmxStart
KMXStep:
              numKmx = numKmx + 1
              kmx = kmx0 + (numKmx - 1) * sgnkmx * StepResolution
              If sgnkmx = 1 And kmx > KmxEnd Then
                 'nothing found with this threshhold
                 GoTo NextKmy
              ElseIf sgnkmx = -1 And kmx < kmxGlitch1 Then
                 'nothing found with this threshhold
                 GoTo NextKmy
                 End If
                 
               'old DTM height as this point
               If numKmx = 1 Then
                  kmxDTM = kmx
                  kmyDTM = kmy
                  Call heights(kmxDTM, kmyDTM, hgt0)
                  GoTo KMXStep
               Else
                  kmxDTM = kmx
                  kmyDTM = kmy
                  Call heights(kmxDTM, kmyDTM, hgt2)
                  If hgt2 >= hgt0 + Val(txtThreshhold.Text) And sgnkmx = 1 Then
                     kmxGlitch1 = kmx
                     hgtGlitch1 = hgt0
                     sgnkmx = -1
                     kmx0 = KmxEnd
                     numKmx = 0
                     GoTo KMXStep
                  ElseIf hgt2 >= hgt0 + Val(txtThreshhold.Text) And sgnkmx = -1 Then
                     kmxGlitch2 = kmx
                     hgtGlitch2 = hgt0
                     sgnkmx = 1
                     kmx0 = KmxStart
                     numKmx = 0
                     
                     If kmxGlitch2 < kmxGlitch1 Then
                        tmp = kmxGlitch1
                        hgttmp = hgtGlitch1
                        kmxGlitch1 = kmxGlitch2
                        hgtGlitch1 = hgtGlitch2
                        kmxGlitch2 = tmp
                        hgtGlitch2 = hgttmp
                        End If
                     
                     For kmxFix = kmxGlitch1 To kmxGlitch2 Step StepResolution
                        'fix these glitch points
                          kmxDTM = kmxFix * 0.001
                          kmyDTM = (kmy - 1000000) * 0.001
                          IKMX& = Int((kmxDTM + 20!) * 40!) + 1
                          IKMY& = Int((380! - kmyDTM) * 40!) + 1
                          NROW% = IKMY&: NCOL% = IKMX&
            
                          'FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
                          Jg% = 1 + Int((NROW% - 2) / 800)
                          Ig% = 1 + Int((NCOL% - 2) / 800)
                          
                         'Since roundoff errors in converting from coord to
                         'integer indexes, just count columns and rows assuming
                         'that the first one has no roundoff error
                          If kmxFix = KmxStart And kmy = KmyStart Then
                             IR% = NROW% - (Jg% - 1) * 800
                             IC% = NCOL% - (Ig% - 1) * 800
                             IR0% = IR%
                             IC0% = IC%
                          Else
                             IC% = CInt((kmxFix - KmxStart) * 0.04) + IC0%
                             IR% = IR0% - CInt((kmy - KmyStart) * 0.04)
                             End If
                          
                          IFN& = (IR% - 1) * 801! + IC%
                          
                          CHFindTmp$ = CHMAP(Ig%, Jg%)
                          If CHFindTmp$ <> CHFind$ Then
                             CHFind$ = CHFindTmp$
                             End If
                             
                          If backup% = 1 Then
                             'back it up if not already backed up
                             If Dir(israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)) = sEmpty Then
                                FileCopy israeldtm + ":\dtm\" & CHFind$, israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                                End If
                             End If
                                
                          CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
                          Close 'close any opened files (for example by sub heights)
                          filn% = FreeFile
                          Open israeldtm + ":\dtm\" & CHFind$ For Random As #filn% Len = 2
                     
                          'write the changes to the DTM tile
                          'Since roundoff errors in converting from coord to
                          'integer indexes, just count columns and rows assuming
                          'that the first one has no roundoff error
                          IR% = NROW% - (Jg% - 1) * 800
                          IC% = NCOL% - (Ig% - 1) * 800
                          IR0% = IR%
                          IC0% = IC%
                        
                          IFN& = (IR% - 1) * 801! + IC%
                          
                          If (kmxGlitch1 <> kmxGlitch2) Then
                            hgtNew = hgtGlitch1 + (kmxFix - kmxGlitch1) * (hgtGlitch2 - hgtGlitch1) / (kmxGlitch2 - kmxGlitch1)
                          Else
                            '1 point glitch take average of heights of border
                            kmxDTM = kmxFix - DTMResolution%
                            kmyDTM = kmy
                            Call heights(kmxDTM, kmyDTM, hgt3)
                            kmxDTM = kmxFix + DTMResolution%
                            kmyDTM = kmy
                            Call heights(kmxDTM, kmyDTM, hgt4)
                            hgtNew = 0.5 * (hgt3 + hgt4)
                            End If
                            
                        Put #filn%, IFN&, CInt(hgtNew * 10)
                           'diagnostics////////////////////
'                          Get #filn%, IFN&, ION% 'diagnostics
'                           kmxDTM = kmxFix
'                           kmyDTM = kmy
'                           Call heights(kmxDTM, kmyDTM, hgt5)
                           '//////////////////////////////
                     Next kmxFix
                     'fixed glitches in kmy direction
                     'move to next kmx
                     'goto NextKMy
                Else
                  hgt0 = hgt2
                  GoTo KMXStep
                  End If
            End If
NextKmy:
        Next kmy
        
        ElseIf optHorizontal Then
            'glitch is mainly horizontal, i.e., row oriented
            CHFind$ = sEmpty
            TotalNum% = CInt((KmxEnd - KmxStart) / StepResolution)
            For kmx = KmxStart To KmxEnd Step StepResolution
              progbarFixGlitch.value = CInt(100 * (kmx - KmxStart) / (KmxEnd - KmxStart))
              numKmy = 0
              sgnkmy = 1
              kmy0 = KmyStart
KMYStep:
              numKmy = numKmy + 1
              kmy = kmy0 + (numKmy - 1) * sgnkmy * StepResolution
              If sgnkmy = 1 And kmy > KmyEnd Then
                 'nothing found with this threshhold
                 GoTo NextKmx
              ElseIf sgnkmy = -1 And kmy < kmyGlitch1 Then
                 'nothing found with this threshhold
                 GoTo NextKmx
                 End If
                        
              'old DTM height as this point
              If numKmy = 1 Then
                  kmxDTM = kmx
                  kmyDTM = kmy
                  Call heights(kmxDTM, kmyDTM, hgt0)
                  GoTo KMYStep
              Else
                  kmxDTM = kmx
                  kmyDTM = kmy
                  Call heights(kmxDTM, kmyDTM, hgt2)
                   If hgt2 >= hgt0 + Val(txtThreshhold.Text) And sgnkmy = 1 Then
                      kmyGlitch1 = kmy
                      hgtGlitch1 = hgt0
                      sgnkmy = -1
                      kmy0 = KmyEnd
                      numKmy = 0
                      GoTo KMYStep
                   ElseIf hgt2 >= hgt0 + Val(txtThreshhold.Text) And sgnkmy = -1 Then
                      kmyGlitch2 = kmy
                      hgtGlitch2 = hgt0
                      sgnkmy = 1
                      kmy0 = KmyStart
                      numKmy = 0
                      
                     If kmyGlitch2 < kmyGlitch1 Then
                        tmp = kmyGlitch1
                        hgttmp = hgtGlitch1
                        kmyGlitch1 = kmyGlitch2
                        hgtGlitch1 = hgtGlitch2
                        kmyGlitch2 = tmp
                        hgtGlitch2 = hgttmp
                        End If
                        
                     For kmyFix = kmyGlitch1 To kmyGlitch2 Step StepResolution
                        'fix these glitch points
                          kmxDTM = kmx * 0.001
                          kmyDTM = (kmyFix - 1000000) * 0.001
                          IKMX& = Int((kmxDTM + 20!) * 40!) + 1
                          IKMY& = Int((380! - kmyDTM) * 40!) + 1
                          NROW% = IKMY&: NCOL% = IKMX&
            
                          'FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
                          Jg% = 1 + Int((NROW% - 2) / 800)
                          Ig% = 1 + Int((NCOL% - 2) / 800)
                          
                         'Since roundoff errors in converting from coord to
                         'integer indexes, just count columns and rows assuming
                         'that the first one has no roundoff error
                          If kmx = KmxStart And kmyFix = KmyStart Then
                             IR% = NROW% - (Jg% - 1) * 800
                             IC% = NCOL% - (Ig% - 1) * 800
                             IR0% = IR%
                             IC0% = IC%
                          Else
                             IC% = CInt((kmx - KmxStart) * 0.04) + IC0%
                             IR% = IR0% - CInt((kmyFix - KmyStart) * 0.04)
                             End If
                          
                          IFN& = (IR% - 1) * 801! + IC%
                              
                          CHFindTmp$ = CHMAP(Ig%, Jg%)
                          If CHFindTmp$ <> CHFind$ Then
                            CHFind$ = CHFindTmp$
                            End If
                               
                            If backup% = 1 Then
                               'back it up if not already backed up
                               If Dir(israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)) = sEmpty Then
                                  FileCopy israeldtm + ":\dtm\" & CHFind$, israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                                  End If
                               End If
Retry:
                              CHMNEO = sEmpty 'reinitialize DTM reading (GETZ)
                              Close 'close any opened files (for example by sub heights)
                              filn% = FreeFile
                              Open israeldtm + ":\dtm\" & CHFind$ For Random As #filn% Len = 2
                         
                              'write the changes to the DTM tile
                              'Since roundoff errors in converting from coord to
                              'integer indexes, just count columns and rows assuming
                              'that the first one has no roundoff error
                              IR% = NROW% - (Jg% - 1) * 800
                              IC% = NCOL% - (Ig% - 1) * 800
                              IR0% = IR%
                              IC0% = IC%
                            
                              IFN& = (IR% - 1) * 801! + IC%
                              
                              If kmyGlitch2 <> kmyGlitch1 Then
                                 hgtNew = hgtGlitch1 + (kmyFix - kmyGlitch1) * (hgtGlitch2 - hgtGlitch1) / (kmyGlitch2 - kmyGlitch1)
                              Else
                                '1 point giltch, take average of heights of border
                                kmxDTM = kmx
                                kmyDTM = kmyFix - DTMResolution%
                                Call heights(kmxDTM, kmyDTM, hgt3)
                                kmxDTM = kmx
                                kmyDTM = kmyFix + DTMResolution%
                                Call heights(kmxDTM, kmyDTM, hgt4)
                                hgtNew = 0.5 * (hgt3 + hgt4)
                              
                                 hgtNew = hgtGlitch1
                                 End If
                                 
                              Put #filn%, IFN&, CInt(hgtNew * 10)
                         Next kmyFix
                         'fixed glitches in kmy direction
                         'move to next kmx
                         'goto NextKMy
                    Else
                      hgt0 = hgt2
                      GoTo KMYStep
                      End If
                End If
NextKmx:
            Next kmx
       Else
            Call MsgBox("Choose either horizontal or vertical orientation of glitch!", vbInformation, "Fix Glitch")
            Exit Sub
            End If
            
       progbarFixGlitch.Visible = False
          
       If filn% > 0 Then Close #filn%
       CHMNEO = sEmpty
          
    Exit Sub


   On Error GoTo 0
   Exit Sub

cmdFix_Click_Error:
    cc = Err.Number
    
    Close 'close any open files
    Select Case Err.Number
       Case 55 'can't save file to this directory
          If FirstTry% = 0 Then 'try once again
             FirstTry% = 1
             Resume
             End If
          response = MsgBox("Can't save the tile to the dtm directory!" & vbLf & _
                 "Do you want to save it to a different directory?", _
                 vbExclamation + vbYesNoCancel, "Maps & More")
          If response = vbYes Then
             res = InputBox("Input directory name." & vbLf & _
                 "Add a final backslashes '\'" & vbLf & _
                 "Example: ""c:\windows\temp\""", "Backup Tile", _
                 israeldtm + ":\dtm\", 6450)
             If res = sEmpty Then
                progbarFixGlitch.Visible = False
                Exit Sub
             Else
                FileCopy israeldtm + ":\dtm\" & CHFind$, res & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                Resume Next
                End If
             End If
       Case 75 'can't save to tile
          If Retrial% = 0 Then
             Retrial% = 1
             GoTo Retry
          Else
             Retrial% = 0
             End If
          MsgBox "The dtm tile is read only and can't be edited!", vbCritical + vbOKOnly, "Maps&More"
          progbarFixGlitch.Visible = False
          Exit Sub
       Case Else
            MsgBox "Error Number: " & Str$(Err.Number) & " encountered!" & vbLf & _
                   Err.Description & vbLf & _
                   "Current procedure will be aborted!", vbCritical + vbOKOnly, "Maps&More"
            'ExcelApp.Quit
            
            'Set ExcelApp = Nothing
            'Set ExcelBook = Nothing
            'Set ExcelSheet = Nothing
      End Select
    
    progbarFixGlitch.Visible = False
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFix_Click of Form mapXYGlitchfrm"
   
End Sub

Private Sub Form_Load()

  Dim KmxStart, KmxEnd, KmyStart, KmyEnd

   Call ScreenToGeo(drag1x, drag1y, kmxDrag, kmyDrag, 1, ier%)
   KmxStart = kmxDrag
   KmyEnd = kmyDrag

   Call ScreenToGeo(drag2x, drag2y, kmxDrag, kmyDrag, 1, ier%)
   KmxEnd = kmxDrag
   KmyStart = kmyDrag
   
   'Now determine end and start with respect to actual DTM
   'points:
   KmxStart = CLng(KmxStart * 0.04) * 25#
   KmxEnd = CLng(KmxEnd * 0.04) * 25#
   KmyStart = CLng(KmyStart * 0.04) * 25#
   KmyEnd = CLng(KmyEnd * 0.04) * 25#

   Call sCenterForm(Me)
   txtXStart = KmxStart
   txtXEnd = KmxEnd
   txtYStart = KmyStart
   txtYEnd = KmyEnd
End Sub
